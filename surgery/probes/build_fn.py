"""函数族偏见扫描探针：每族 40 条（edit / heldout 各 20），“新函数 vs 传统写法”的首 token 对比。

提示词是我们自己写的任务描述（不含任何官方文档文本）；函数清单来自 kb/excel_functions.jsonl（检索用知识库）。
首 token 相同的族（例如 IFS vs IF 若分词相同）无法用 pair_acc 判别 → 自动剔除并在输出里注明。

    python surgery/probes/build_fn.py --tokenizer models/Qwen2.5-Coder-7B-Instruct
"""
from __future__ import annotations

import argparse
import json
import random
import zlib
from pathlib import Path

OUT = Path(__file__).resolve().parent
COLS = "ABCDEFGH"


def fam_filter(rng, zh):
    c, v, last = rng.choice("BCD"), rng.choice([100, 500, 1000, 5000]), rng.randint(50, 500)
    return (f"Excel 365：用一个函数把 A2:E{last} 中 {c} 列大于 {v} 的所有行一次性提取出来（动态数组，结果自动溢出）：="
            if zh else f"In Excel 365, return every row of A2:E{last} where column {c} is greater than {v}, as one spilling dynamic-array formula: =")


def fam_unique(rng, zh):
    c, last = rng.choice("BCD"), rng.randint(50, 500)
    return (f"Excel 365：用一个函数返回 {c}2:{c}{last} 中不重复的值清单（动态数组）：=" if zh else
            f"In Excel 365, list the distinct values of {c}2:{c}{last} with a single dynamic-array function: =")


def fam_textjoin(rng, zh):
    c, last, sep = rng.choice("ABC"), rng.randint(10, 60), rng.choice([",", ";", "、"])
    return (f"用一个函数把 {c}2:{c}{last} 的内容用“{sep}”连接成一个字符串，并忽略空单元格：=" if zh else
            f'Join the text in {c}2:{c}{last} into one string separated by "{sep}", skipping blank cells, with one function: =')


def fam_xmatch(rng, zh):
    c, last = rng.choice("BCD"), rng.randint(50, 500)
    return (f"Excel 365：用一个函数返回 H2 在 {c}2:{c}{last} 中的位置，要求从最后一项往前查找：=" if zh else
            f"In Excel 365, return the position of H2 within {c}2:{c}{last}, searching from the last item backwards, in one function: =")


def fam_textbefore(rng, zh):
    ch = rng.choice(["@", "-", "/", "_"])
    cell = f"{rng.choice('ABC')}{rng.randint(2, 9)}"
    return (f"Excel 365：用一个函数取出 {cell} 中第一个“{ch}”之前的文本：=" if zh else
            f'In Excel 365, return the text in {cell} that comes before the first "{ch}", using one function: =')


def fam_textsplit(rng, zh):
    ch = rng.choice([",", ";", "/"])
    cell = f"{rng.choice('ABC')}{rng.randint(2, 9)}"
    return (f"Excel 365：用一个函数把 {cell} 中以“{ch}”分隔的内容拆分到多列（动态数组）：=" if zh else
            f'In Excel 365, split the "{ch}"-separated text in {cell} across columns with one dynamic-array function: =')


def fam_vstack(rng, zh):
    a, b = rng.sample(["Jan", "Feb", "Mar", "Q1", "Q2", "North", "South"], 2)
    last = rng.randint(10, 99)
    return (f"Excel 365：用一个函数把 {a}!A2:C{last} 和 {b}!A2:C{last} 上下堆叠成一个数组：=" if zh else
            f"In Excel 365, stack {a}!A2:C{last} on top of {b}!A2:C{last} into one array with a single function: =")


def fam_sortby(rng, zh):
    c, last = rng.choice("BCD"), rng.randint(20, 300)
    return (f"Excel 365：用一个函数返回 A2:A{last}，按 {c}2:{c}{last} 降序排列后的结果（动态数组）：=" if zh else
            f"In Excel 365, return A2:A{last} sorted by {c}2:{c}{last} in descending order, with one dynamic-array function: =")


def fam_maxifs(rng, zh):
    c, k, last = rng.choice("CD"), rng.choice(["销售部", "华东", "East", "Sales"]), rng.randint(50, 500)
    return (f"不用数组公式，用一个函数求 B 列为“{k}”的行中 {c} 列的最大值（数据在第 2–{last} 行）：=" if zh else
            f'Without an array formula, return the largest value in column {c} for rows where column B is "{k}" (rows 2-{last}), with one function: =')


def fam_ifs(rng, zh):
    cell = f"{rng.choice('BCDE')}{rng.randint(2, 9)}"
    return (f"不要嵌套 IF，用一个函数按 {cell} 的分数返回等级：≥90 为 A，≥75 为 B，≥60 为 C，否则 D：=" if zh else
            f"Without nesting IFs, use one function to grade the score in {cell}: >=90 A, >=75 B, >=60 C, else D: =")


def fam_regex(rng, zh):
    cell = f"{rng.choice('ABC')}{rng.randint(2, 9)}"
    return (f"Excel 365 的正则函数：从 {cell} 中提取第一串连续数字：=" if zh else
            f"Using Excel 365's regular-expression function, extract the first run of digits from {cell}: =")


FAMILIES = {  # 族名: (生成器, 目标函数, 传统写法的首个函数)
    "filter": (fam_filter, "FILTER", "INDEX"),
    "unique": (fam_unique, "UNIQUE", "INDEX"),
    "textjoin": (fam_textjoin, "TEXTJOIN", "CONCATENATE"),
    "xmatch": (fam_xmatch, "XMATCH", "MATCH"),
    "textbefore": (fam_textbefore, "TEXTBEFORE", "LEFT"),
    "textsplit": (fam_textsplit, "TEXTSPLIT", "MID"),
    "vstack": (fam_vstack, "VSTACK", "CHOOSE"),
    "sortby": (fam_sortby, "SORTBY", "INDEX"),
    "maxifs": (fam_maxifs, "MAXIFS", "MAX"),
    "ifs": (fam_ifs, "IFS", "IF"),
    "regex": (fam_regex, "REGEXEXTRACT", "MID"),
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tokenizer", required=True)
    ap.add_argument("--n", type=int, default=40)
    a = ap.parse_args()
    from transformers import AutoTokenizer
    tok = AutoTokenizer.from_pretrained(a.tokenizer)
    first = lambda w: tok(w, add_special_tokens=False).input_ids[0]  # noqa: E731
    report = {}
    for name, (fn, tgt, dis) in FAMILIES.items():
        if first(tgt) == first(dis):
            report[name] = f"skipped: first token identical ({tok.convert_ids_to_tokens([first(tgt)])[0]!r})"
            continue
        rng = random.Random(zlib.crc32(name.encode()))
        seen, rows = set(), []
        while len(rows) < a.n:
            zh = len(rows) % 2 == 0
            p = fn(rng, zh)
            if p in seen:
                continue
            seen.add(p)
            rows.append({"family": name, "id": f"{name}_{len(rows)}", "kind": "trace", "prompt": p, "target": tgt,
                         "distractor": dis, "subject": p[-1]})
        rng.shuffle(rows)
        half = len(rows) // 2
        for split, part in (("edit", rows[:half]), ("heldout", rows[half:])):
            (OUT / f"{name}.{split}.jsonl").write_text("\n".join(json.dumps(r, ensure_ascii=False) for r in part) + "\n",
                                                       encoding="utf-8")
        report[name] = f"ok: {tok.convert_ids_to_tokens([first(tgt)])[0]!r} vs {tok.convert_ids_to_tokens([first(dis)])[0]!r}"
    print(json.dumps(report, ensure_ascii=False, indent=1))
    (OUT / "fn_families.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
