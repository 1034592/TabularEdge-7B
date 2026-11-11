"""手术脚本公共件。只在本地电脑（32GB RAM）运行，VPS 上禁止 import torch 加载权重。

硬件事实（暗影精灵 i7-11800H / 32GB / RTX 3050 Laptop 4GB）：
- 7B bf16 ≈ 15.2GB → 只能放 CPU 内存；4GB 显存放不下 7B（NF4 也要 ~4.5GB+激活）。
- 因此：方法开发先用 Qwen2.5-Coder-1.5B（GPU 可跑，分钟级迭代），结论稳定后再上 7B（CPU，小时级）。
"""
from __future__ import annotations

import json
import os
from contextlib import contextmanager
from pathlib import Path

import torch
from transformers import AutoModelForCausalLM, AutoTokenizer

DEFAULT_MODEL = os.environ.get("TABULAREDGE_MODEL", r"models\Qwen2.5-Coder-1.5B-Instruct")


def load(model_path: str = DEFAULT_MODEL, device: str | None = None, dtype=None):
    """1.5B：GPU fp16（约 3.1GB，刚好放进 3050 的 4GB）；CPU 用 fp32（11800H 没有原生 bf16，bf16 在 CPU 上反而慢）。"""
    big_gpu = torch.cuda.is_available() and torch.cuda.get_device_properties(0).total_memory >= 20 * 2**30
    device = device or os.environ.get("TABULAREDGE_DEVICE") or (
        "cuda" if torch.cuda.is_available() and (big_gpu or "7B" not in str(model_path)) else "cpu")
    # 4090 等支持 bf16 的大显存卡用 bf16（与 GGUF 源权重同精度）；3050 4GB 上 1.5B 用 fp16；CPU 用 fp32
    if dtype is None:
        dtype = (torch.bfloat16 if big_gpu else torch.float16) if device == "cuda" else torch.float32
    tok = AutoTokenizer.from_pretrained(model_path)
    model = AutoModelForCausalLM.from_pretrained(model_path, dtype=dtype, low_cpu_mem_usage=True).to(device).eval()
    if os.environ.get("TABULAREDGE_ADAPTER"):  # 底座 + LoRA 在内存里合并（不写盘），用于测未导出的组合
        from peft import PeftModel
        model = PeftModel.from_pretrained(model, os.environ["TABULAREDGE_ADAPTER"]).merge_and_unload().eval()
    for p in model.parameters():
        p.requires_grad_(False)
    return model, tok


def layers(model):
    return model.model.layers


def module(model, layer: int, kind: str):
    """kind: hidden(整层输出) | mlp | attn | down(down_proj；其 *输入* 即 MLP 神经元激活，用 pre-hook 取)"""
    L = layers(model)[layer]
    return {"hidden": L, "mlp": L.mlp, "attn": L.self_attn, "down": L.mlp.down_proj}[kind]


def _main(out):
    """HF 各版本层输出格式不同：tensor 或 (tensor, ...)。"""
    return out[0] if isinstance(out, tuple) else out


def _replace(out, new):
    return (new, *out[1:]) if isinstance(out, tuple) else new


@contextmanager
def hooks(handles):
    try:
        yield
    finally:
        for h in handles:
            h.remove()


def read_jsonl(p):
    return [json.loads(l) for l in Path(p).read_text(encoding="utf-8").splitlines() if l.strip()]


def target_id(tok, target: str) -> int:
    """取目标答案的第一个 token。探针设计时应保证首 token 已有区分度（如 ' XLOOKUP' vs ' VLOOKUP'）。"""
    return tok(target, add_special_tokens=False).input_ids[0]


def subject_span(tok, prompt: str, subject: str) -> list[int]:
    """subject 子串对应的 token 位置（用 offset mapping，中文安全）。"""
    enc = tok(prompt, return_offsets_mapping=True, add_special_tokens=False)
    s = prompt.index(subject)
    e = s + len(subject)
    return [i for i, (a, b) in enumerate(enc.offset_mapping) if a < e and b > s]
