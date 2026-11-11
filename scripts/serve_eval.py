"""宿主机上：起 llama-server（Vulkan / RTX 3050）→ 等 /health → 跑 eval_plans → 关 server。

    python scripts/serve_eval.py --gguf models/out/sft-1.5b-cf/sft-1.5b-cf-Q4_K_M.gguf --tag sft15 --prompt sft
    python scripts/serve_eval.py --gguf ... --tag base15 --prompt fewshot
其余参数原样传给 synth/eval_plans.py（例如 --limit 20 --grammar）。
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parent.parent
WIN = os.name == "nt"
SERVER = Path(os.environ.get("TABULAREDGE_SERVER", r"tools\llama-vulkan\llama-server.exe" if WIN
                             else "llama.cpp/build/bin/llama-server"))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--gguf", required=True)
    ap.add_argument("--port", type=int, default=8080)
    ap.add_argument("--ngl", default="99")
    ap.add_argument("--ctx", default="4096")
    ap.add_argument("--keep", action="store_true", help="评测完不关 server（给 demo 用）")
    a, rest = ap.parse_known_args()
    (ROOT / "logs").mkdir(exist_ok=True)
    log = open(ROOT / "logs" / f"server_{Path(a.gguf).stem}.log", "w", encoding="utf-8")
    srv = subprocess.Popen([str(SERVER), "-m", a.gguf, "--host", "127.0.0.1", "--port", str(a.port), "-c", a.ctx,
                            "-ngl", a.ngl, *(["-dev", "Vulkan0"] if WIN else []), "--cache-reuse", "256", "-np", "1"],
                           stdout=log, stderr=subprocess.STDOUT)
    url = f"http://127.0.0.1:{a.port}"
    t0 = time.time()
    while True:
        try:
            if requests.get(url + "/health", timeout=2).status_code == 200:
                break
        except requests.RequestException:
            pass
        if srv.poll() is not None or time.time() - t0 > 180:
            print("server failed to start; see", log.name, flush=True)
            sys.exit(2)
        time.sleep(1)
    print(f"server ready in {time.time() - t0:.1f}s", flush=True)
    try:
        rc = subprocess.call([sys.executable, "-u", str(ROOT / "synth" / "eval_plans.py"), "--backend", "server", "--url", url, *rest])
    finally:
        if not a.keep:
            srv.terminate()
            srv.wait(timeout=30)
    sys.exit(rc)


if __name__ == "__main__":
    main()
