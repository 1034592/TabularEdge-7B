"""人类扰动发生器：把干净指令变成“真实用户会打出来的样子”。金标 plan 不变，模型必须透过噪声识别意图。

扰动族（每条按概率独立施加，可叠加）：
  homophone   拼音输入法同音/近音错字（会总、删出、不门、被注…）
  initials    拼音首字母缩写（hz=汇总、qc=去重、bh=标红…）
  punct       全半角标点混乱、重复标点、丢标点
  mix         中英混杂（sort一下、delete掉、标red…）
  invert      倒装：动作在前、对象/条件在后
  filler      口语填充词（帮我、麻烦、那个、嗯、pls、asap）
  en_typo     英文拼写错误（换位/丢字/双写），以及缩写（amt、qty、w/）
  keyboard    键盘相邻键误触（只用于留出测试集）

split="train" 只用训练表；split="heldout" 只用留出表 + keyboard 族 —— 用来测“没见过的错法”。
数字与引号内的值永不扰动：阈值、颜色值、填充值是因果变量，改了金标就不对了。
"""
from __future__ import annotations

import random
import re

# (正确, 错误)。训练表与留出表互不相交
HOMO_TRAIN = [("汇总", "会总"), ("删除", "删出"), ("部门", "不门"), ("备注", "被注"), ("重复", "重付"), ("排序", "排续"),
              ("标红", "表红"), ("数量", "数两"), ("单价", "单架"), ("区域", "区预"), ("合计", "和计"), ("新表", "新标"),
              ("空格", "空隔"), ("填充", "填冲"), ("括号", "扩号"), ("库存", "酷存"), ("报销", "报消"), ("降序", "将序"),
              ("负数", "付数"), ("文件", "文建"), ("拆分", "拆份"), ("格式", "各式"), ("记录", "纪录"), ("平均", "平军")]
HOMO_HELD = [("去重", "去从"), ("金额", "金鹅"), ("升序", "生序"), ("标黄", "表黄"), ("清理", "清里"), ("统计", "统计计"),
             ("千分位", "千分为"), ("分组", "分租"), ("保存", "保层"), ("销售", "消售")]
INITIALS_TRAIN = [("汇总", "hz"), ("去重", "qc"), ("删除", "sc"), ("标红", "bh"), ("排序", "px"), ("降序", "jx")]
# v3 补强：财务/行政常见缩写（留出 sx/bf/cf 仍不进训练）
INITIALS_TRAIN_V3 = INITIALS_TRAIN + [("统计", "tj"), ("部门", "bm"), ("金额", "je"), ("数量", "sl"), ("单价", "dj"), ("日期", "rq"),
                                      ("备注", "bz"), ("新表", "xb"), ("保存", "bc"), ("合并", "hb"), ("填充", "tc"), ("空格", "kg"),
                                      ("括号", "kh"), ("删掉", "sc掉"), ("合计", "hj")]
INITIALS_HELD = [("升序", "sx"), ("备份", "bf"), ("拆分", "cf")]
MIX_TRAIN = [("排序", "sort一下"), ("删掉", "delete掉"), ("汇总", "sum一下"), ("标红", "标red"), ("新表", "new sheet"),
             ("去重", "dedup"), ("另存为", "save as"), ("备份", "backup"), ("空格", "space")]
MIX_HELD = [("拆成", "split成"), ("填", "fill"), ("统计", "count一下")]
FILL_ZH = ["帮我", "麻烦", "那个", "嗯，", "老板说", "急！", "顺便", "就是"]
FILL_EN = ["pls ", "can you ", "hey, ", "quickly ", "asap: ", "boss wants: "]
EN_ABBR_TRAIN = [("amount", "amt"), ("quantity", "qty"), ("with", "w/"), ("please", "pls"), ("department", "dept"), ("and", "&")]
EN_ABBR_HELD = [("descending", "desc"), ("duplicates", "dupes"), ("highlight", "hilite")]
KEYS = "qwertyuiop asdfghjkl zxcvbnm".split()
FW = {",": "，", ";": "；", ":": "：", "(": "（", ")": "）", "!": "！", "?": "？"}
HW = {v: k for k, v in FW.items()}

_PROTECT = re.compile(r'\d+(?:\.\d+)?|“[^”]*”|"[^"]*"|\[[^\]]*\]|[A-Za-z0-9_\-/]+\.(?:xlsx|csv)')


def _protected(text):
    return [(m.start(), m.end()) for m in _PROTECT.finditer(text)]


def _safe_replace(text, old, new, rng, p):
    """把 old 替换成 new（概率 p，逐处判断），不碰受保护片段和表头里的词（columns 里的精确列名由调用方传入保护）。"""
    out, i = [], 0
    spans = _protected(text)
    for m in re.finditer(re.escape(old), text):
        if any(a <= m.start() < b for a, b in spans) or rng.random() > p:
            continue
        out.append(text[i:m.start()])
        out.append(new)
        i = m.end()
    out.append(text[i:])
    return "".join(out)


def _en_typo(word, rng):
    if len(word) < 4 or not word.isalpha():
        return word
    k = rng.randrange(1, len(word) - 1)
    op = rng.choice(["swap", "drop", "dup"])
    if op == "swap":
        return word[:k] + word[k + 1] + word[k] + word[k + 2:]
    if op == "drop":
        return word[:k] + word[k + 1:]
    return word[:k] + word[k] + word[k:]


def _keyboard(word, rng):
    if len(word) < 4 or not word.isalpha():
        return word
    k = rng.randrange(len(word))
    ch = word[k].lower()
    for row in KEYS:
        j = row.find(ch)
        if j >= 0:
            nb = [row[x] for x in (j - 1, j + 1) if 0 <= x < len(row)]
            return word[:k] + rng.choice(nb) + word[k + 1:]
    return word


def _num_abbr(t, rng):
    """数字单位缩写：10000→1w、5000→5k（只改独立的整万/整千数字）。"""
    def sub(m):
        v = int(m.group(0))
        if v >= 10000 and v % 10000 == 0 and rng.random() < 0.7:
            return f"{v // 10000}w"
        if 1000 <= v < 10000 and v % 1000 == 0 and rng.random() < 0.5:
            return f"{v // 1000}k"
        return m.group(0)
    return re.sub(r"(?<![\d.])\d{4,6}(?![\d.%])", sub, t)


def perturb(text: str, lang: str, rng: random.Random, split: str = "train", strength: float = 1.0,
            protect: tuple[str, ...] = (), v3: bool = False) -> tuple[str, list[str]]:
    """返回 (扰动后文本, 施加的扰动族列表)。protect：精确列名，homophone/mix 不改它们（列名错字由 homophone 表单独覆盖）。"""
    held = split == "heldout"
    applied = []
    p = lambda base: min(1.0, base * strength)  # noqa: E731
    t = text
    shield = {}
    for i, w in enumerate(sorted(protect, key=len, reverse=True)):
        if w and w in t:
            key = f"\u0001{i}\u0002"
            shield[key] = w
            t = t.replace(w, key)
    if lang == "zh":
        homo = HOMO_HELD if held else HOMO_TRAIN
        ini = INITIALS_HELD if held else (INITIALS_TRAIN_V3 if v3 else INITIALS_TRAIN)
        mix = MIX_HELD if held else MIX_TRAIN
        if rng.random() < p(0.5):
            before = t
            for a, b in rng.sample(homo, k=min(4, len(homo))):
                t = _safe_replace(t, a, b, rng, 0.8)
            if t != before:
                applied.append("homophone")
        if rng.random() < p(0.45 if v3 else 0.25):
            before = t
            for a, b in ini:
                t = _safe_replace(t, a, b, rng, 0.7)
            if t != before:
                applied.append("initials")
        if rng.random() < p(0.3):
            before = t
            for a, b in rng.sample(mix, k=min(3, len(mix))):
                t = _safe_replace(t, a, b, rng, 0.8)
            if t != before:
                applied.append("mix")
        if rng.random() < p(0.25):
            parts = re.split(r"(，|；|然后|再)", t)
            clauses = [x for x in parts[::2] if x.strip()]
            if len(clauses) == 1 and len(t) > 8 and "的" in t:  # 单句倒装：把“…的”条件挪到后面
                a, _, b = t.partition("的")
                if a and b:
                    t = f"{b.strip('，')}，{a}的"
                    applied.append("invert")
        if rng.random() < p(0.3):
            t = rng.choice(FILL_ZH) + t
            applied.append("filler")
    else:
        abbr = EN_ABBR_HELD if held else EN_ABBR_TRAIN
        if rng.random() < p(0.4):
            before = t
            for a, b in abbr:
                t = re.sub(rf"\b{a}\b", b, t) if rng.random() < 0.7 else t
            if t != before:
                applied.append("en_abbr")
        if rng.random() < p(0.5):
            words = t.split(" ")
            fn = _keyboard if held else _en_typo
            idx = [i for i, w in enumerate(words) if w.isalpha() and len(w) >= 4 and "\u0001" not in w]
            for i in rng.sample(idx, k=min(len(idx), rng.randint(1, 2))):
                words[i] = fn(words[i], rng)
            if " ".join(words) != t:
                t = " ".join(words)
                applied.append("keyboard" if held else "en_typo")
        if rng.random() < p(0.3) and t != t.lower():
            t = t.lower()
            applied.append("lower")
        if rng.random() < p(0.25):
            t = rng.choice(FILL_EN) + t
            applied.append("filler")
    if v3 and rng.random() < p(0.3):
        before = t
        t = _num_abbr(t, rng)
        if t != before:
            applied.append("num_abbr")
    if rng.random() < p(0.5):  # 标点：全半角互换 / 重复 / 丢失
        before = t
        mode = rng.choice(["swap", "dup", "drop"])
        if mode == "swap":
            t = "".join(FW.get(ch, HW.get(ch, ch)) if rng.random() < 0.6 else ch for ch in t)
        elif mode == "dup":
            t = re.sub(r"([，,！!？?])", lambda m: m.group(1) * rng.randint(1, 3), t)
        else:
            t = re.sub(r"[，,；;]", lambda m: " " if rng.random() < 0.5 else m.group(0), t)
        if t != before:
            applied.append("punct")
    for k, w in shield.items():
        t = t.replace(k, w)
    return t, applied


if __name__ == "__main__":
    r = random.Random(0)
    for s, lang in [("把备注里括号及括号里的内容去掉，然后按部门汇总金额，结果放新表", "zh"),
                    ("金额是负数的标红", "zh"), ("按金额降序排序，再删除重复行", "zh"),
                    ("delete rows where amount is less than 100, then sort by quantity descending", "en")]:
        for split in ("train", "heldout"):
            print(split, perturb(s, lang, r, split, 1.5))
