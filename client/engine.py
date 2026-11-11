"""本地推理引擎管理：内嵌 llama-server（CUDA / Vulkan / CPU 三套），按显存自动切分层数，失败自动降级。

策略（RTX 3050 Laptop 4GB 这类小显存卡）：
  1. nvidia-smi 读空闲显存 → 估算能放下的层数（每层 ≈ (模型大小 − 词表/输出层) / 层数，另留计算缓冲与系统余量）
  2. 先试 CUDA；进程退出或日志里出现显存分配失败 → 层数减 4 重试；仍不行换 Vulkan；最后纯 CPU
  3. 纯 CPU 模式：-ngl 0，8 线程（i7-11800H 8 个物理核）
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path

import requests

NO_WINDOW = 0x08000000 if os.name == "nt" else 0
OOM_RX = re.compile(r"out of memory|failed to allocate|cudaMalloc failed|ErrorOutOfDeviceMemory|unable to allocate|"
                    r"failed to load model|error loading model", re.I)


def _kill_on_close_job():
    """Windows Job Object：本程序退出（哪怕崩溃 / 被强杀）时，系统自动结束 llama-server，避免它在后台占着显存。"""
    if os.name != "nt":
        return None
    try:
        import ctypes
        from ctypes import wintypes
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)

        class BASIC(ctypes.Structure):
            _fields_ = [("PerProcessUserTimeLimit", ctypes.c_int64), ("PerJobUserTimeLimit", ctypes.c_int64),
                        ("LimitFlags", wintypes.DWORD), ("MinimumWorkingSetSize", ctypes.c_size_t),
                        ("MaximumWorkingSetSize", ctypes.c_size_t), ("ActiveProcessLimit", wintypes.DWORD),
                        ("Affinity", ctypes.c_size_t), ("PriorityClass", wintypes.DWORD), ("SchedulingClass", wintypes.DWORD)]

        class IOC(ctypes.Structure):
            _fields_ = [(n, ctypes.c_uint64) for n in ("r", "w", "o", "rb", "wb", "ob")]

        class EXT(ctypes.Structure):
            _fields_ = [("Basic", BASIC), ("Io", IOC), ("ProcessMemoryLimit", ctypes.c_size_t),
                        ("JobMemoryLimit", ctypes.c_size_t), ("PeakProcessMemoryUsed", ctypes.c_size_t),
                        ("PeakJobMemoryUsed", ctypes.c_size_t)]
        k32.CreateJobObjectW.restype = wintypes.HANDLE
        job = k32.CreateJobObjectW(None, None)
        info = EXT()
        info.Basic.LimitFlags = 0x2000  # JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        k32.SetInformationJobObject(job, 9, ctypes.byref(info), ctypes.sizeof(info))  # 9 = ExtendedLimitInformation
        return job, k32
    except Exception:  # noqa: BLE001
        return None


_JOB = _kill_on_close_job()


def _bind_to_job(proc):
    if _JOB is None:
        return
    job, k32 = _JOB
    try:
        k32.AssignProcessToJobObject(job, int(proc._handle))  # noqa: SLF001
    except Exception:  # noqa: BLE001
        pass


@dataclass
class EngineState:
    status: str = "stopped"          # stopped / starting / ready / failed
    backend: str = ""
    ngl: int = 0
    n_layers: int = 28
    gpu: str = ""
    vram_free_mb: int = 0
    tok_s: float = 0.0
    detail: str = ""
    tried: list = field(default_factory=list)


def app_dir() -> Path:
    if os.environ.get("TABULAREDGE_HOME"):
        return Path(os.environ["TABULAREDGE_HOME"])
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent.parent


def query_gpu():
    """返回 (名称, 总显存MB, 空闲显存MB)；无 NVIDIA 卡或驱动返回 None。"""
    try:
        out = subprocess.run(["nvidia-smi", "--query-gpu=name,memory.total,memory.free", "--format=csv,noheader,nounits"],
                             capture_output=True, text=True, timeout=10, creationflags=NO_WINDOW).stdout.strip()
        name, tot, free = [x.strip() for x in out.splitlines()[0].split(",")]
        return name, int(float(tot)), int(float(free))
    except Exception:  # noqa: BLE001
        return None


def plan_layers(model_bytes: int, n_layers: int, free_mb: int, ctx: int, reserve_mb: int = 900) -> int:
    """估算可放进显存的层数。每层 = 权重 + 该层 KV 缓存（Qwen2.5-7B：4 个 KV 头 × 128 维，f16）。"""
    head_mb = 760  # token_embd + output（Q4_K_M 7B 约 0.75GB，留在 CPU）
    per_layer = max((model_bytes / 2**20 - head_mb) / n_layers, 1) + ctx * 2 * 4 * 128 * 2 / 2**20
    n = int((free_mb - reserve_mb) // per_layer)
    return max(0, min(n_layers, n))


class Engine:
    def __init__(self, cfg: dict, log_dir: Path, on_change=None):
        self.cfg = cfg
        self.log_dir = log_dir
        self.on_change = on_change or (lambda st: None)
        self.state = EngineState(n_layers=int(cfg.get("n_layers", 28)))
        self.proc: subprocess.Popen | None = None
        self.port = int(cfg.get("port", 18080))
        self.url = f"http://127.0.0.1:{self.port}"
        self._lock = threading.Lock()

    # ---------------------------------------------------------------- helpers
    def _emit(self, **kw):
        for k, v in kw.items():
            setattr(self.state, k, v)
        self.on_change(self.state)

    def _bin(self, backend):
        return app_dir() / "engine" / backend / ("llama-server.exe" if os.name == "nt" else "llama-server")

    def healthy(self) -> bool:
        try:
            return requests.get(self.url + "/health", timeout=2).status_code == 200
        except requests.RequestException:
            return False

    def stop(self):
        p, self.proc = self.proc, None
        if p and p.poll() is None:
            p.terminate()
            try:
                p.wait(8)
            except subprocess.TimeoutExpired:
                p.kill()
        self._emit(status="stopped")

    def model_path(self) -> Path:
        """配置里的相对路径（默认 models\\TabularEdge-7B-Q4_K_M.gguf）按程序目录解析。"""
        m = Path(self.cfg["model"])
        return m if m.is_absolute() else app_dir() / m

    # ---------------------------------------------------------------- start
    def candidates(self, mode: str):
        """(backend, ngl) 依次尝试。mode: auto / cpu / gpu-full。"""
        model = self.model_path()
        size = model.stat().st_size
        ctx = int(self.cfg.get("ctx", 4096))
        L = self.state.n_layers
        out = []
        g = query_gpu() if mode != "cpu" else None
        if g:
            name, tot, free = g
            self._emit(gpu=name, vram_free_mb=free)
            n = L + 1 if mode == "gpu-full" else plan_layers(size, L, free, ctx, int(self.cfg.get("vram_reserve_mb", 900)))
            if self.cfg.get("ngl", "auto") != "auto":
                n = int(self.cfg["ngl"])
            for backend in ("cuda", "vulkan"):
                if not self._bin(backend).exists():
                    continue
                k = n
                while k > 0:
                    out.append((backend, k))
                    k -= 4
        out.append(("cpu", 0))
        return out

    def start(self, mode: str = "auto"):
        with self._lock:
            self.stop()
            model = self.model_path()
            if not model.exists():
                self._emit(status="failed", detail=f"model not found: {model}")
                return False
            self._emit(status="starting", tried=[])
            if self.proc is None and self.healthy():  # 端口上已有本程序的引擎在跑（例如程序开了两次）→ 直接复用
                self._emit(status="ready", backend=self.state.backend or "external", detail="reuse running engine")
                threading.Thread(target=self._bench, daemon=True).start()
                return True
            tried = []
            for backend, ngl in self.candidates(mode):
                exe = self._bin(backend)
                if not exe.exists():
                    continue
                tried.append(f"{backend}:{ngl}")
                self._emit(backend=backend, ngl=ngl, detail=f"{backend} ngl={ngl}", tried=tried)
                ok, why = self._launch(exe, backend, ngl)
                if ok:
                    self._emit(status="ready", detail=f"{backend} {ngl}/{self.state.n_layers}")
                    threading.Thread(target=self._bench, daemon=True).start()
                    return True
                self._emit(detail=f"{backend} ngl={ngl} failed: {why}")
            self._emit(status="failed", detail="all backends failed: " + ", ".join(tried))
            return False

    def _launch(self, exe: Path, backend: str, ngl: int):
        threads = int(self.cfg.get("threads", 0)) or max(1, (os.cpu_count() or 8) // 2)
        args = [str(exe), "-m", str(self.model_path()), "--host", "127.0.0.1", "--port", str(self.port),
                "-c", str(self.cfg.get("ctx", 4096)), "-np", "1", "-t", str(threads), "-ngl", str(ngl),
                "--cache-reuse", "256"]
        if backend != "cpu":
            args += ["-dev", "CUDA0" if backend == "cuda" else "Vulkan0"]
        if self.cfg.get("fit_off", True):
            args += ["--fit", "off"]  # 层数由本程序决定（见 plan_layers），失败再降级
        args += list(self.cfg.get("extra_args", []))
        self.log_dir.mkdir(parents=True, exist_ok=True)
        logf = self.log_dir / f"llama-server-{backend}.log"
        fh = open(logf, "w", encoding="utf-8", errors="replace")
        self.proc = subprocess.Popen(args, stdout=fh, stderr=subprocess.STDOUT, cwd=str(exe.parent), creationflags=NO_WINDOW)
        _bind_to_job(self.proc)
        t0 = time.time()
        while time.time() - t0 < float(self.cfg.get("start_timeout_s", 180)):
            if self.proc.poll() is not None:
                fh.close()
                txt = logf.read_text(encoding="utf-8", errors="replace")
                if "invalid argument: --fit" in txt or "error: invalid argument" in txt and "--fit" in txt:
                    self.cfg["fit_off"] = False  # 老版本没有 --fit
                    return self._launch(exe, backend, ngl)
                m = OOM_RX.search(txt)
                return False, (m.group(0) if m else txt.strip().splitlines()[-1][:160] if txt.strip() else f"exit {self.proc.returncode}")
            if self.healthy():
                return True, ""
            time.sleep(0.5)
        self.stop()
        return False, "timeout"

    def _bench(self):
        """启动后热身一次，顺便测生成速度（tok/s）显示在状态栏。"""
        try:
            r = requests.post(self.url + "/completion", json={"prompt": "TabularEdge warmup:", "n_predict": 24, "temperature": 0,
                                                              "cache_prompt": False}, timeout=120).json()
            self._emit(tok_s=round(float(r.get("timings", {}).get("predicted_per_second", 0.0)), 1))
        except Exception:  # noqa: BLE001
            pass


def load_config(path: Path) -> dict:
    cfg = {"model": r"models\TabularEdge-7B-Q4_K_M.gguf", "mode": "auto", "ngl": "auto",
           "ctx": 4096, "threads": 8, "port": 18080, "n_layers": 28, "vram_reserve_mb": 900, "lang": "zh",
           "workdir": "", "excel": "auto", "extra_args": []}
    try:
        cfg.update(json.loads(path.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        pass
    return cfg


def save_config(path: Path, cfg: dict):
    try:
        path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass
