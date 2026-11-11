"""生成“已知顽疾”探针集（每类按 edit / heldout 切分；手术只看 edit，指标只报 heldout）。

families
  sumifs   SUMIFS 第 1 参数是求和区域；SUMIF 第 1 参数是条件区域（模型常把两者混淆）
  vlookup  VLOOKUP 第 3 参数列号（差一错误）
  xlookup  “找不到时返回默认值、无需 IFERROR”→ XLOOKUP 而非 VLOOKUP
  repeat   长公式生成（嵌套 IF / IFS / SWITCH / 长 SUMIFS 链），测复读死循环
每条 trace 探针含 target 与 distractor（错误答案首 token），用 P(target) > P(distractor) 判对。

    python surgery/probes/build.py
"""
import json
import random
import string
from pathlib import Path

OUT = Path(__file__).resolve().parent
rng = random.Random(42)
L = string.ascii_uppercase[:12]


def sumifs():
    rows = []
    zh_items = [("销售额", "区域", "华东"), ("金额", "部门", "财务部"), ("数量", "产品", "键盘"), ("工资", "岗位", "经理")]
    en_items = [("sales", "region", "East"), ("amount", "department", "Finance"), ("units", "product", "Laptop"), ("cost", "category", "Travel")]
    for i in range(80):
        s, k = rng.sample(L, 2)
        multi = i % 2 == 0
        zh = i % 4 >= 2
        val, crit, cv = rng.choice(zh_items if zh else en_items)
        fn = "SUMIFS" if multi else "SUMIF"
        if zh:
            p = f"{val}在 {s}2:{s}200，{crit}在 {k}2:{k}200。用 {fn} 计算{crit}为“{cv}”的{val}合计：=" + fn + "("
        else:
            p = f"The {val} are in {s}2:{s}200 and the {crit} in {k}2:{k}200. Using {fn}, total the {val} where {crit} is \"{cv}\": =" + fn + "("
        tgt, dis = (s, k) if multi else (k, s)
        rows.append({"family": "sumifs", "id": f"sumifs_{i}", "kind": "trace", "prompt": p, "target": tgt, "distractor": dis,
                     "subject": fn, "variant": fn})
    return rows


def vlookup():
    rows = []
    for i in range(60):
        last = rng.randint(3, 8)
        X = rng.randint(2, last)
        lc = L[last - 1]
        xc = L[X - 1]
        zh = i % 2
        if zh:
            p = f"数据区域是 A:{lc}，查找键在 A 列，需要返回 {xc} 列的值。公式：=VLOOKUP(H2,A:{lc},"
        else:
            p = f"The table is A:{lc}, the lookup key is in column A and I need the value from column {xc}. Formula: =VLOOKUP(H2,A:{lc},"
        d = X + 1 if X + 1 <= last else X - 1
        rows.append({"family": "vlookup", "id": f"vlookup_{i}", "kind": "trace", "prompt": p, "target": str(X), "distractor": str(d),
                     "subject": f"column {xc}" if not zh else f"{xc} 列"})
    return rows


def xlookup():
    rows = []
    zh_t = ["按{k}查{v}，找不到时显示“{m}”，不要嵌套 IFERROR，用一个函数：=",
            "Excel 365 里根据{k}返回{v}，查不到返回“{m}”，单个函数写法：="]
    en_t = ["Look up the {v} for each {k}; if it is not found show \"{m}\", without wrapping in IFERROR, using one function: =",
            "In Excel 365, return the {v} matching a {k}, with \"{m}\" when there is no match, in a single function: ="]
    zh_kv = [("员工编号", "姓名", "无此人"), ("订单号", "金额", "未找到"), ("物料编码", "库存", "缺货")]
    en_kv = [("employee ID", "name", "Not found"), ("order number", "amount", "N/A"), ("SKU", "stock", "missing")]
    for i in range(40):
        zh = i % 2
        k, v, m = rng.choice(zh_kv if zh else en_kv)
        p = rng.choice(zh_t if zh else en_t).format(k=k, v=v, m=m)
        rows.append({"family": "xlookup", "id": f"xlookup_{i}", "kind": "trace", "prompt": p, "target": "XLOOKUP",
                     "distractor": "VLOOKUP", "subject": m})
    return rows


def repeat():
    rows = []
    grades = ["A", "B", "C", "D", "E", "F"]
    for i in range(40):
        t = i % 4
        cell = f"{rng.choice('BCDE')}{rng.randint(2, 9)}"
        if t == 0:
            th = sorted(rng.sample(range(40, 100, 5), 5), reverse=True)
            p = (f"Write one Excel formula that converts the score in {cell} to grades "
                 + ", ".join(f"{g} (>={x})" for g, x in zip(grades, th)) + f", else F:\n=IF({cell}>={th[0]},\"A\",IF(")
        elif t == 1:
            p = f"用 IFS 把 {cell} 中的月份数字(1-12)转成中文月份名：\n=IFS({cell}=1,\"一月\",{cell}=2,\"二月\","
        elif t == 2:
            p = f"Excel formula using SWITCH to turn weekday number in {cell} (1-7) into names:\n=SWITCH({cell},1,\"Monday\",2,\"Tuesday\","
        else:
            reg = rng.choice(["华东", "华南", "华北"])
            p = f"=SUMIFS(销售额,区域,\"{reg}\",月份,1)+SUMIFS(销售额,区域,\"{reg}\",月份,2)+SUMIFS("
        rows.append({"family": "repeat", "id": f"repeat_{i}", "kind": "repeat", "prompt": p, "template": t})
    return rows


def main():
    for fam in (sumifs(), vlookup(), xlookup(), repeat()):
        rng.shuffle(fam)
        half = len(fam) // 2
        name = fam[0]["family"]
        for split, part in (("edit", fam[:half]), ("heldout", fam[half:])):
            (OUT / f"{name}.{split}.jsonl").write_text("\n".join(json.dumps(r, ensure_ascii=False) for r in part) + "\n", encoding="utf-8")
        print(name, len(fam))


if __name__ == "__main__":
    main()
