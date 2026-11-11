"""在宿主机 RTX 3050 4GB 上对 Qwen2.5-Coder-1.5B 做 QLoRA（4-bit 基座 + LoRA），零云端成本。

损失 = 标准 SFT 交叉熵（只算 assistant 部分）
     + λ · 反事实对比损失：对反事实对 (x,y) / (x',y')，要求 log p(y|x) − log p(y'|x) ≥ margin。
       y' 是“只改了一个因果变量”的指令 x' 的正确答案。模型如果只看表面词、不看那个变量，
       就会给 y 和 y' 差不多的概率，这一项会惩罚它。

    python synth/train_lora.py --data data/synth/train.jsonl --base models/Qwen2.5-Coder-1.5B-Instruct --out models/out/lora-1.5b-cf
"""
from __future__ import annotations

import argparse
import json
import math
import random
import sys
import time
from pathlib import Path

import torch
import torch.nn.functional as F
from peft import LoraConfig, get_peft_model, prepare_model_for_kbit_training
from transformers import AutoModelForCausalLM, AutoTokenizer, BitsAndBytesConfig

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from tabularedge.prompting import messages, target  # noqa: E402


def encode(tok, row, plan=None, max_len=1024):
    if row.get("raw"):  # 原始文本语言模型正则：不套模板，全部 token 算损失（约束“裸文本”建模能力不漂移）
        ids = tok(row["raw"] + tok.eos_token, add_special_tokens=False).input_ids[:max_len]
        return torch.tensor([ids]), torch.tensor([[-100] + ids[1:]])
    if row.get("chat"):  # 通用对话（自蒸馏正则化数据）：模板默认系统提示 + 原版模型的回答
        p = tok.apply_chat_template(row["messages"], add_generation_prompt=True, tokenize=True)
        if not isinstance(p, list):
            p = p["input_ids"]
        y = tok(row["response"] + "<|im_end|>", add_special_tokens=False).input_ids
        ids, labels = (p + y)[:max_len], ([-100] * len(p) + y)[:max_len]
        return torch.tensor([ids]), torch.tensor([labels])
    msgs = messages(row["file"], row["profile"], row["text"])
    p = tok.apply_chat_template(msgs, add_generation_prompt=True, tokenize=True)
    if not isinstance(p, list):  # 新版 transformers 可能返回 BatchEncoding
        p = p["input_ids"]
    y = tok(target(plan or row["plan"]) + "<|im_end|>", add_special_tokens=False).input_ids
    ids = (p + y)[:max_len]
    labels = ([-100] * len(p) + y)[:max_len]
    return torch.tensor([ids]), torch.tensor([labels])


def seq_logprob(model, ids, labels):
    """只在 assistant 位置算 lm_head：152k 词表 × 全序列的 fp32 logits 会撑爆 4GB 显存。"""
    base = model.get_base_model()
    h = base.model(input_ids=ids).last_hidden_state[0]
    tgt = labels[0, 1:]
    pos = (tgt != -100).nonzero().squeeze(-1)
    logits = base.lm_head(h[pos]).float()
    lp = F.log_softmax(logits, -1).gather(1, tgt[pos].unsqueeze(-1)).squeeze(-1)
    return lp.sum(), torch.tensor(len(pos), device=ids.device)


def batch_logprob(model, tok, seqs):
    """seqs: [(ids_tensor[1,T], labels_tensor[1,T])] → (每条序列 assistant 部分的 log p 之和 [B], token 数 [B])。右侧填充。"""
    base = model.get_base_model()
    L = max(x.shape[1] for x, _ in seqs)
    pad = tok.pad_token_id if tok.pad_token_id is not None else tok.eos_token_id
    ids = torch.full((len(seqs), L), pad, dtype=torch.long)
    lab = torch.full((len(seqs), L), -100, dtype=torch.long)
    att = torch.zeros((len(seqs), L), dtype=torch.long)
    for b, (x, y) in enumerate(seqs):
        ids[b, :x.shape[1]], lab[b, :y.shape[1]], att[b, :x.shape[1]] = x[0], y[0], 1
    ids, lab, att = ids.cuda(), lab.cuda(), att.cuda()
    h = base.model(input_ids=ids, attention_mask=att).last_hidden_state[:, :-1]
    tgt = lab[:, 1:]
    mask = tgt != -100
    lp = F.log_softmax(base.lm_head(h[mask]).float(), -1).gather(1, tgt[mask].unsqueeze(-1)).squeeze(-1)
    rows = mask.nonzero()[:, 0]
    sums = torch.zeros(len(seqs), device=lp.device).index_add(0, rows, lp)
    cnts = torch.zeros(len(seqs), device=lp.device).index_add(0, rows, torch.ones_like(lp))
    return sums, cnts


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default="data/synth/train.jsonl")
    ap.add_argument("--base", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--epochs", type=float, default=1.0)
    ap.add_argument("--lr", type=float, default=2e-4)
    ap.add_argument("--accum", type=int, default=8)
    ap.add_argument("--rank", type=int, default=16)
    ap.add_argument("--cf-weight", type=float, default=0.2, help="0 = 关闭反事实对比损失（消融）")
    ap.add_argument("--cf-margin", type=float, default=2.0)
    ap.add_argument("--max-len", type=int, default=1024)
    ap.add_argument("--no-quant", action="store_true", help="bf16 LoRA（大显存卡），不做 4-bit 量化")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--init-adapter", help="从已有 LoRA 继续训练（收官淬火）")
    ap.add_argument("--no-ckpt", action="store_true", help="关闭梯度检查点（大显存卡上约快 30%）")
    ap.add_argument("--bucket", type=int, default=64, help="批量路径：每 bucket 个批次内按长度排序再组批，减少填充浪费")
    ap.add_argument("--batch", type=int, default=1, help="每个前向的样本数（>1 时走对齐填充的批量路径；反事实兄弟答案同批计算）")
    ap.add_argument("--save-every", type=int, default=250, help="每 N 步存一次 adapter（中途停机最多丢 N 步）")
    a = ap.parse_args()

    rows = [json.loads(l) for l in open(a.data, encoding="utf-8")]
    if a.limit:
        rows = rows[:a.limit]
    pairs = {}
    for r in rows:
        if r.get("pair"):
            pairs.setdefault(r["pair"], []).append(r)
    sib = {}
    for v in pairs.values():
        if len(v) == 2:
            sib[id(v[0])], sib[id(v[1])] = v[1], v[0]

    tok = AutoTokenizer.from_pretrained(a.base)
    if a.no_quant:
        model = AutoModelForCausalLM.from_pretrained(a.base, dtype=torch.bfloat16, device_map={"": 0})
        if not a.no_ckpt:
            model.gradient_checkpointing_enable()
        model.enable_input_require_grads()
    else:
        bnb = BitsAndBytesConfig(load_in_4bit=True, bnb_4bit_quant_type="nf4", bnb_4bit_compute_dtype=torch.bfloat16,
                                 bnb_4bit_use_double_quant=True)
        model = AutoModelForCausalLM.from_pretrained(a.base, quantization_config=bnb, dtype=torch.bfloat16, device_map={"": 0})
        model = prepare_model_for_kbit_training(model, use_gradient_checkpointing=True)
    if a.init_adapter:
        from peft import PeftModel
        model = PeftModel.from_pretrained(model, a.init_adapter, is_trainable=True)
    else:
        model = get_peft_model(model, LoraConfig(r=a.rank, lora_alpha=2 * a.rank, lora_dropout=0.05, task_type="CAUSAL_LM",
                                                 target_modules=["q_proj", "k_proj", "v_proj", "o_proj", "gate_proj", "up_proj", "down_proj"]))
    model.print_trainable_parameters()
    params = [p for p in model.parameters() if p.requires_grad]
    opt = torch.optim.AdamW(params, lr=a.lr, weight_decay=0.0)
    total = math.ceil(len(rows) * a.epochs)
    sched = torch.optim.lr_scheduler.LambdaLR(opt, lambda s: min(1.0, (s + 1) / 20) * 0.5 * (1 + math.cos(math.pi * min(s, total // a.accum) / max(total // a.accum, 1))))
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    log = (out / "train_log.jsonl").open("a", encoding="utf-8")
    rng = random.Random(0)
    order = []
    while len(order) < total:
        o = list(range(len(rows)))
        rng.shuffle(o)
        order += o
    order = order[:total]
    if a.batch > 1 and a.bucket > 1:  # 长度分桶：桶内按长度排序组批，再打乱批次顺序（不改变每条样本被看到的次数）
        L = lambda i: len(rows[i]["raw"]) if rows[i].get("raw") else \
            (len(rows[i]["messages"][0]["content"]) + len(rows[i]["response"])) if rows[i].get("chat") else \
            len(rows[i]["profile"]) + len(rows[i]["text"]) + len(json.dumps(rows[i]["plan"], ensure_ascii=False))  # noqa: E731
        span = a.batch * a.bucket
        batches = []
        for k in range(0, len(order), span):
            seg = sorted(order[k:k + span], key=L)
            batches += [seg[j:j + a.batch] for j in range(0, len(seg), a.batch)]
        rng.shuffle(batches)
        order = [i for b in batches for i in b]
    model.train()
    t0, agg = time.time(), {"ce": 0.0, "cf": 0.0, "cf_gap": 0.0, "n": 0, "ncf": 0}
    B = max(1, a.batch)
    n_batches = math.ceil(len(order) / B)
    if B > 1:  # 调度器按“优化器步”计；批量路径下每 accum 个批次一步
        total_opt = max(n_batches // a.accum, 1)
        sched = torch.optim.lr_scheduler.LambdaLR(opt, lambda s: min(1.0, (s + 1) / 20) * 0.5 * (1 + math.cos(math.pi * min(s, total_opt) / total_opt)))
    for bi in range(n_batches):
        chunk = [rows[i] for i in order[bi * B:(bi + 1) * B]]
        if B == 1:  # 原单样本路径（与之前的实验逐位一致）
            r = chunk[0]
            ids, lab = encode(tok, r, max_len=a.max_len)
            lp, n = seq_logprob(model, ids.cuda(), lab.cuda())
            loss = -lp / n
            agg["ce"] += loss.item()
            s = sib.get(id(r))
            if a.cf_weight > 0 and s is not None:
                ids2, lab2 = encode(tok, r, plan=s["plan"], max_len=a.max_len)
                lp2, _ = seq_logprob(model, ids2.cuda(), lab2.cuda())
                gap = lp - lp2
                cf = F.softplus(a.cf_margin - gap)
                loss = loss + a.cf_weight * cf
                agg["cf"] += cf.item()
                agg["cf_gap"] += gap.item()
                agg["ncf"] += 1
            agg["n"] += 1
        else:
            seqs = [encode(tok, r, max_len=a.max_len) for r in chunk]
            pairs = [(j, sib[id(r)]) for j, r in enumerate(chunk) if a.cf_weight > 0 and id(r) in sib]
            seqs += [encode(tok, chunk[j], plan=sb["plan"], max_len=a.max_len) for j, sb in pairs]
            sums, cnts = batch_logprob(model, tok, seqs)
            k = len(chunk)
            ce = (-sums[:k] / cnts[:k]).mean()
            loss = ce
            agg["ce"] += ce.item() * k
            agg["n"] += k
            if pairs:
                gap = sums[[j for j, _ in pairs]] - sums[k:]
                cf = F.softplus(a.cf_margin - gap)
                loss = loss + a.cf_weight * cf.mean()
                agg["cf"] += cf.sum().item()
                agg["cf_gap"] += gap.sum().item()
                agg["ncf"] += len(pairs)
        (loss / a.accum).backward()
        step = (bi + 1) * B
        if (bi + 1) % a.accum == 0:
            torch.nn.utils.clip_grad_norm_(params, 1.0)
            opt.step()
            sched.step()
            opt.zero_grad(set_to_none=True)
        if (bi + 1) % max(1, 50 // B) == 0:
            rec = {"step": step, "of": total, "ce": round(agg["ce"] / max(agg["n"], 1), 4),
                   "cf": round(agg["cf"] / max(agg["ncf"], 1), 4), "cf_gap": round(agg["cf_gap"] / max(agg["ncf"], 1), 2),
                   "lr": sched.get_last_lr()[0], "sec": round(time.time() - t0), "mem_gb": round(torch.cuda.max_memory_allocated() / 2**30, 2)}
            print(json.dumps(rec), flush=True)
            log.write(json.dumps(rec) + "\n")
            log.flush()
            agg = {kk: 0 if isinstance(v, int) else 0.0 for kk, v in agg.items()}
        if (bi + 1) % max(1, a.save_every // B) == 0:
            model.save_pretrained(out / "adapter")
            (out / "adapter" / "STEP").write_text(str(step))
    model.save_pretrained(out / "adapter")
    print("DONE", out, round(time.time() - t0), "s", flush=True)


if __name__ == "__main__":
    main()
