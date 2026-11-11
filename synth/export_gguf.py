"""合并 LoRA（可选）→ HF safetensors → GGUF f16 → 量化（Q4_K_M / Q8_0）。在宿主机运行。

    python synth/export_gguf.py --base models/Qwen2.5-Coder-1.5B-Instruct \
        --adapter models/out/lora-1.5b-cf/adapter --name sft-1.5b-cf --quant Q4_K_M Q8_0
    python synth/export_gguf.py --hf models/out/edited-7b --name edited-7b --quant Q4_K_M   # 已是 HF 目录
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

WIN = os.name == "nt"
TOOLS = Path(os.environ.get("TABULAREDGE_TOOLS", r"tools" if WIN else "."))
OUTROOT = Path(os.environ.get("TABULAREDGE_OUT", r"models\out" if WIN else "models/out"))
QUANTIZE = Path(os.environ.get("TABULAREDGE_QUANTIZE", TOOLS / "llama-vulkan" / "llama-quantize.exe" if WIN
                               else TOOLS / "llama.cpp" / "build" / "bin" / "llama-quantize"))


def run(cmd, env=None):
    print(">", " ".join(map(str, cmd)), flush=True)
    t0 = time.time()
    subprocess.run(list(map(str, cmd)), check=True, env=env)
    print(f"  ({time.time() - t0:.0f}s)", flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base")
    ap.add_argument("--adapter")
    ap.add_argument("--hf", help="已合并好的 HF 目录（跳过合并）")
    ap.add_argument("--name", required=True)
    ap.add_argument("--quant", nargs="*", default=["Q4_K_M"])
    ap.add_argument("--merge-only", action="store_true", help="只合并 LoRA 存 HF，不转 GGUF（磁盘紧张时分两步）")
    ap.add_argument("--rm-f16", action="store_true", help="量化后删除 f16 中间 GGUF（省磁盘）")
    ap.add_argument("--final-name", help="量化文件改用这个名字（如 TabularEdge-7B），得到 <final-name>-<quant>.gguf")
    a = ap.parse_args()
    out = OUTROOT / a.name
    out.mkdir(parents=True, exist_ok=True)
    hf = Path(a.hf) if a.hf else (out / "hf" if a.adapter else Path(a.base))
    if a.adapter:
        import torch
        from peft import PeftModel
        from transformers import AutoModelForCausalLM, AutoTokenizer
        print("merging LoRA on CPU ...", flush=True)
        m = AutoModelForCausalLM.from_pretrained(a.base, dtype=torch.bfloat16, low_cpu_mem_usage=True)
        m = PeftModel.from_pretrained(m, a.adapter).merge_and_unload()
        m.save_pretrained(hf, safe_serialization=True)
        AutoTokenizer.from_pretrained(a.base).save_pretrained(hf)
        del m
        if a.merge_only:
            print("merged ->", hf, flush=True)
            return
    f16 = out / f"{a.name}-f16.gguf"
    env = {**os.environ, "PYTHONPATH": str(TOOLS / "llama.cpp" / "gguf-py"), "PYTHONUTF8": "1"}
    run([sys.executable, TOOLS / "llama.cpp" / "convert_hf_to_gguf.py", hf, "--outtype", "f16", "--outfile", f16], env)
    for q in a.quant:
        run([QUANTIZE, f16, out / f"{a.final_name or a.name}-{q}.gguf", q])
    if a.rm_f16 and a.quant:
        f16.unlink()
        print("removed", f16.name, flush=True)
    for p in sorted(out.glob("*.gguf")):
        print(f"{p.name}\t{p.stat().st_size / 2**30:.2f} GiB")


if __name__ == "__main__":
    main()
