"""LoRA 增量缩放：W = W_merged − (1−α)·ΔW，ΔW = scaling·B@A（不需要底座，直接从合并权重回退）。
用于在“任务能力”与“通用语言能力（困惑度）”之间找折中点。

    python surgery/scale_lora.py sweep --merged <hf> --adapter <dir> --alphas 0.9 0.75 0.6 --out outputs/measure/alpha_sweep.json
    python surgery/scale_lora.py save  --merged <hf> --adapter <dir> --alpha 0.75 --dst <hf_out>
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import torch
from safetensors.torch import load_file

from common import load
from measure import measure

FAMS = ("sumifs", "vlookup", "xlookup", "filter", "unique", "textjoin", "xmatch", "textbefore", "textsplit", "vstack", "sortby", "ifs", "regex")


def deltas(model, adapter_dir):
    cfg = json.loads((Path(adapter_dir) / "adapter_config.json").read_text())
    scale = cfg["lora_alpha"] / cfg["r"]
    sd = load_file(str(Path(adapter_dir) / "adapter_model.safetensors"))
    params = dict(model.named_parameters())
    out = {}
    for k, A in sd.items():
        if ".lora_A." not in k:
            continue
        B = sd[k.replace(".lora_A.", ".lora_B.")]
        name = k.replace("base_model.model.", "", 1).replace(".lora_A.weight", ".weight")
        W = params[name]
        out[name] = (scale * (B.to(W.device, torch.float32) @ A.to(W.device, torch.float32)))
    return out, params


@torch.no_grad()
def apply(params, D, factor):
    """W += factor·ΔW（factor 为负表示回退）。"""
    for n, d in D.items():
        W = params[n]
        W.copy_((W.float() + factor * d).to(W.dtype))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["sweep", "save"])
    ap.add_argument("--merged", required=True)
    ap.add_argument("--adapter", required=True)
    ap.add_argument("--alphas", type=float, nargs="*", default=[0.9, 0.75, 0.6])
    ap.add_argument("--alpha", type=float)
    ap.add_argument("--out")
    ap.add_argument("--dst")
    a = ap.parse_args()
    model, tok = load(a.merged)
    D, params = deltas(model, a.adapter)
    print("lora modules:", len(D), flush=True)
    if a.cmd == "save":
        apply(params, D, -(1 - a.alpha))
        model.save_pretrained(a.dst, safe_serialization=True)
        tok.save_pretrained(a.dst)
        print("saved", a.dst)
        return
    res, cur = {}, 1.0
    for al in sorted(a.alphas, reverse=True):
        apply(params, D, al - cur)
        cur = al
        m = measure(model, tok, "heldout", FAMS)
        res[str(al)] = {"ppl": m["ppl"], **{f: m[f]["pair_acc"] for f in FAMS}}
        print(al, json.dumps(res[str(al)]), flush=True)
    Path(a.out).write_text(json.dumps(res, indent=1))


if __name__ == "__main__":
    main()
