"""人工“地狱”测试集 v2：304 条手写口语指令 + 手写金标（不经过任何生成器模板）。

    python data/human/build_v2.py   → data/human/human_v2.jsonl

类别 cat：abbr 拼音缩写 · typo 错别字/同音 · oral 倒装口语 · en 英文/中英混杂 · amb 高危歧义（必须追问） ·
          risky 明确的破坏性操作（不该追问，但执行器硬确认） · complex 多步复合/跨表 · newop 新函数族
steps=None 的 amb 题：故意不说明白、没有唯一合理计划 → 只考“是否追问”，不考执行。
金标按协议的规范阶段排序（清洗→加列→格式→汇总/拆分→保存），与训练时的计划约定一致。
"""
from __future__ import annotations

import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
PAREN = r"[\(（][^\)）]*[\)）]"
RED, YEL, GRN = "FF0000", "FFFF00", "00B050"
HEADS = {"L": ["日期", "部门", "金额", "备注"], "S": ["Order ID", "Sales Rep", "Region", "Product", "Qty", "Unit Price", "Remarks"],
         "I": ["发票号", "客户", "含税金额", "不含税金额", "开票日期", "备注"], "P": ["工号", "姓名", "部门", "基本工资", "绩效奖金", "实发工资", "备注"],
         "K": ["物料编码", "物料名称", "仓库", "库存数量", "单价", "备注"]}
FILES = {"L": "流水账.xlsx", "S": "sales.xlsx", "I": "发票台账.xlsx", "P": "工资表.xlsx", "K": "库存.xlsx"}
DIMS = {"L": "部门预算", "S": "Catalog", "I": "客户信息", "P": "部门信息", "K": "供应商"}


def rx(c): return {"op": "regex_replace", "args": {"col": c, "pattern": PAREN, "repl": ""}}
def trim(*cs): return {"op": "trim", "args": {"cols": list(cs)}}
def fill(c, v): return {"op": "fill_empty", "args": {"col": c, "value": v}}
def fmt(c, f): return {"op": "number_format", "args": {"col": c, "format": f}}
def addc(n, f): return {"op": "add_column", "args": {"name": n, "formula": f}}
def split(by, tpl=None): return {"op": "split_to_files", "args": {"by": by, "path_template": tpl or f"拆分_{by}/{{value}}.xlsx"}}
def save(p): return {"op": "save", "args": {"path": p}}
def bak(b): return {"op": "fs_backup", "args": {"path": FILES[b]}}
def uq(c, cnt=False): return {"op": "unique_to_sheet", "args": {"col": c, "output_sheet": "清单", **({"with_count": True} if cnt else {})}}


def W(c, cmp, v=None):
    w = {"col": c, "cmp": cmp}
    if v is not None:
        w["value"] = v
    return w


def dele(c, cmp, v=None): return {"op": "delete_rows", "args": {"where": W(c, cmp, v)}}
def dall(b): return {"op": "delete_rows", "args": {"where": {"all": [W(h, "is_empty") for h in HEADS[b]]}}}
def dedup(*cs): return {"op": "dedupe", "args": {"cols": list(cs)} if cs else {}}
def srt(c, desc=False): return {"op": "sort", "args": {"by": [{"col": c, "desc": desc}]}}


def cf(c, cmp, v, color=RED, row=False, where=None):
    style = {"fill_color": color} if color == YEL else {"font_color": color}
    return {"op": "conditional_format", "args": {"where": where or W(c, cmp, v), "apply_to": "row" if row else "cell", "style": style}}


def grp(by, col, fn="sum", aggs=None):
    by = [by] if isinstance(by, str) else by
    return {"op": "group_aggregate", "args": {"by": by, "aggs": aggs or [{"col": col, "fn": fn}], "output_sheet": "汇总", "mode": "formula"}}


def flt(c, cmp, v): return {"op": "filter_to_sheet", "args": {"where": W(c, cmp, v), "output_sheet": "筛选"}}


def lk(b, name, key, ret, nf=None, value=False, from_key=None):
    a = {"name": name, "key": key, "from_sheet": DIMS[b], "from_key": from_key or key, "return": ret}
    if nf is not None:
        a["if_not_found"] = nf
    if value:
        a["mode"] = "value"
    return {"op": "lookup_column", "args": a}


def ifs(c, hi, lo, labels=("高", "中", "低")):
    return f'=IFS([{c}]>={hi},"{labels[0]}",[{c}]>={lo},"{labels[1]}",TRUE,"{labels[2]}")'


I = []  # noqa: E741


def H(b, cat, text, steps, amb=False):
    I.append({"book": b, "cat": cat, "text": text, "steps": steps, "expect_amb": amb})


# ================================================================ L 流水账（流水 / 部门预算）
H("L", "abbr", "备注kh里的内容sc掉", [rx("备注")])
H("L", "abbr", "按bm hz je", [grp("部门", "金额")])
H("L", "abbr", "je jx px", [srt("金额", True)])
H("L", "abbr", "je<0的bh", [cf("金额", "<", 0)])
H("L", "abbr", "按bm tj一下je，放xb", [grp("部门", "金额")])
H("L", "abbr", "qc一下", [dedup()])
H("L", "abbr", "je sx排", [srt("金额")])
H("L", "abbr", "bz为空的tc“无”", [fill("备注", "无")])
H("L", "abbr", "先bf一下 再按je jx", [bak("L"), srt("金额", True)])
H("L", "abbr", "按bm cf成多个文件", [split("部门")])
H("L", "abbr", "je>1w的bh", [cf("金额", ">", 10000)])
H("L", "abbr", "按rq sx", [srt("日期")])
H("L", "typo", "把被注里的扩号内容删出", [rx("备注")])
H("L", "typo", "金鹅小于零的表红", [cf("金额", "<", 0)])
H("L", "typo", "按不门会总金额放到新标", [grp("部门", "金额")])
H("L", "typo", "金额从大到小排续", [srt("金额", True)])
H("L", "typo", "删出重付的行", [dedup()])
H("L", "typo", "被注空着的填冲成“待补”", [fill("备注", "待补")])
H("L", "typo", "金额设成千分为两位小数", [fmt("金额", "#,##0.00")])
H("L", "typo", "负数金额整航标红", [cf("金额", "<", 0, row=True)])
H("L", "typo", "被注的空隔清理下", [trim("备注")])
H("L", "oral", "标红，金额是负的那些", [cf("金额", "<", 0)])
H("L", "oral", "汇总，按部门的金额，弄个新表", [grp("部门", "金额")])
H("L", "oral", "那个备注啊，括号里的东西不要了", [rx("备注")])
H("L", "oral", "倒着排，按金额", [srt("金额", True)])
H("L", "oral", "空的备注，全给我填上无", [fill("备注", "无")])
H("L", "oral", "金额过万的，整行黄色底", [cf("金额", ">", 10000, YEL, row=True)])
H("L", "oral", "能不能把研发部的单独弄一张表出来", [flt("部门", "==", "研发部")])
H("L", "oral", "部门都有啥？列出来，每个多少条", [uq("部门", True)])
H("L", "oral", "我要看每个部门花了多少，单独一张表", [grp("部门", "金额")])
H("L", "oral", "金额那列，负的红色，正的不用管", [cf("金额", "<", 0)])
H("L", "oral", "日期从新到旧", [srt("日期", True)])
H("L", "en", "sum 金额 by 部门 pls", [grp("部门", "金额")])
H("L", "en", "金额 < 0 的 rows 标 red", [cf("金额", "<", 0)])
H("L", "en", "sort 金额 descending then 把备注 trim 一下", [trim("备注"), srt("金额", True)])
H("L", "en", "remove brackets in 备注", [rx("备注")])
H("L", "en", "delete rows where 备注 is empty", [dele("备注", "is_empty")])
H("L", "amb", "带括号的备注都不要了", [rx("备注")], True)
H("L", "amb", "把空的都删了", [dall("L")], True)
H("L", "amb", "处理完保存：把负数金额标红", [cf("金额", "<", 0)], True)
H("L", "amb", "把有括号的备注那些给删喽", [rx("备注")], True)
H("L", "amb", "清一下空记录", [dall("L")], True)
H("L", "amb", "改完存一下，金额降序", [srt("金额", True)], True)
H("L", "amb", "把不对的数据删了", None, True)
H("L", "amb", "金额有问题的改一下", None, True)
H("L", "amb", "把没用的那些记录清掉", None, True)
H("L", "amb", "把这个表清理干净", None, True)
H("L", "risky", "金额为负的行全部删除", [dele("金额", "<", 0)])
H("L", "risky", "直接覆盖原表：把备注括号内容去掉", [rx("备注"), save("流水账.xlsx")])
H("L", "risky", "完全一样的重复行删掉，直接存回原文件", [dedup(), save("流水账.xlsx")])
H("L", "risky", "把行政部的删掉", [dele("部门", "==", "行政部")])
H("L", "complex", "备注去括号去空格，金额负数标红，按部门汇总金额到新表，最后另存为 整理后.xlsx",
  [rx("备注"), trim("备注"), cf("金额", "<", 0), grp("部门", "金额"), save("整理后.xlsx")])
H("L", "complex", "先删掉金额小于0的，再按部门统计条数放新表", [dele("金额", "<", 0), grp("部门", "金额", "count")])
H("L", "complex", "按部门从部门预算把预算和负责人都查过来，查不到写“无”",
  [lk("L", "预算", "部门", "预算", "无"), lk("L", "负责人", "部门", "负责人", "无")])
H("L", "complex", "把研发部筛出来放新表，再把金额设成千分位取整", [fmt("金额", "#,##0"), flt("部门", "==", "研发部")])
H("L", "complex", "金额超过5000的标绿，小于0的标红，然后金额降序", [cf("金额", ">", 5000, GRN), cf("金额", "<", 0), srt("金额", True)])
H("L", "complex", "去重，然后按部门把金额求平均，放新表", [dedup(), grp("部门", "金额", "avg")])
H("L", "complex", "把备注为空的行删掉，再按金额升序", [dele("备注", "is_empty"), srt("金额")])
H("L", "complex", "加一列年份取日期的年，再加一列月份取日期的月", [addc("年份", "=YEAR([日期])"), addc("月份", "=MONTH([日期])")])
H("L", "complex", "备注括号去掉，然后另存为 清洗.xlsx", [rx("备注"), save("清洗.xlsx")])
H("L", "complex", "金额在0到1000之间的标黄", [cf(None, None, None, YEL, where={"all": [W("金额", ">=", 0), W("金额", "<=", 1000)]})])
H("L", "complex", "按部门汇总金额，同时也算一下平均值", [grp("部门", None, aggs=[{"col": "金额", "fn": "sum"}, {"col": "金额", "fn": "avg"}])])
H("L", "complex", "部门预算里的预算带过来，再加一列 差额 = 预算 - 金额", [lk("L", "预算", "部门", "预算"), addc("差额", "=[预算]-[金额]")])
H("L", "complex", "备注里有“退款”的标红", [cf(None, None, None, where=W("备注", "contains", "退款"))])
H("L", "complex", "按部门拆分成多个文件，文件夹就叫按部门", [split("部门", "按部门/{value}.xlsx")])
H("L", "newop", "按部门把预算查过来，没有的话写0", [lk("L", "预算", "部门", "预算", 0)])
H("L", "newop", "金额大于等于1000的记录筛到新表", [flt("金额", ">=", 1000)])
H("L", "newop", "加一列金额档位：>=10000 高，>=1000 中，其余 低", [addc("金额档位", ifs("金额", 10000, 1000))])
H("L", "newop", "加一列叫“取整金额”，把金额四舍五入到整数", [addc("取整金额", "=ROUND([金额],0)")])
H("L", "newop", "把部门和备注用“-”连起来放新列，列名叫“说明”", [addc("说明", '=TEXTJOIN("-",TRUE,[部门],[备注])')])
H("L", "newop", "列出所有不重复的部门", [uq("部门")])
H("L", "newop", "部门预算表里的负责人按部门匹配过来，匹配不到显示“未分配”", [lk("L", "负责人", "部门", "负责人", "未分配")])
H("L", "newop", "把金额改成千分位取整", [fmt("金额", "#,##0")])

# ================================================================ S sales（Orders / Catalog）
H("S", "abbr", "按Region hz Qty", [grp("Region", "Qty")])
H("S", "abbr", "Unit Price jx px", [srt("Unit Price", True)])
H("S", "abbr", "按Order ID qc", [dedup("Order ID")])
H("S", "abbr", "Remarks里的kh内容sc", [rx("Remarks")])
H("S", "abbr", "Qty sx", [srt("Qty")])
H("S", "abbr", "按Region cf文件", [split("Region")])
H("S", "abbr", "Unit Price>5000 bh", [cf("Unit Price", ">", 5000)])
H("S", "abbr", "tj每个Sales Rep的Qty，放xb", [grp("Sales Rep", "Qty")])
H("S", "abbr", "qty jx，然后 bh unit price>5k", [srt("Qty", True), cf("Unit Price", ">", 5000)])
H("S", "typo", "按区预汇总数量", [grp("Region", "Qty")])
H("S", "typo", "单架从高到低排序", [srt("Unit Price", True)])
H("S", "typo", "按订单号去从", [dedup("Order ID")])
H("S", "typo", "数两大于10的标红", [cf("Qty", ">", 10)])
H("S", "typo", "remakrs里括号的内容去掉", [rx("Remarks")])
H("S", "typo", "按销售员会总数量", [grp("Sales Rep", "Qty")])
H("S", "oral", "每个区域单独一个文件", [split("Region")])
H("S", "oral", "数量最多的排前面", [srt("Qty", True)])
H("S", "oral", "单价超过五千的，红色", [cf("Unit Price", ">", 5000)])
H("S", "oral", "Remarks 空着的写个 N/A", [fill("Remarks", "N/A")])
H("S", "oral", "订单号重复的，留第一条就行", [dedup("Order ID")])
H("S", "oral", "按产品看看各卖了多少件，新表", [grp("Product", "Qty")])
H("S", "oral", "sales rep 们都有谁", [uq("Sales Rep")])
H("S", "oral", "把 Mouse 的订单单独筛出来", [flt("Product", "==", "Mouse")])
H("S", "en", "add a Revenue column = Qty * Unit Price", [addc("Revenue", "=[Qty]*[Unit Price]")])
H("S", "en", "sort by unit price high to low", [srt("Unit Price", True)])
H("S", "en", "highlight qty under 5 in yellow", [cf("Qty", "<", 5, YEL)])
H("S", "en", "dedupe by order id and total qty per region in a new sheet", [dedup("Order ID"), grp("Region", "Qty")])
H("S", "en", "trim spaces in remarks", [trim("Remarks")])
H("S", "en", 'fill empty remarks with "n/a"', [fill("Remarks", "n/a")])
H("S", "en", "pls split into one file per sales rep", [split("Sales Rep")])
H("S", "en", "avg unit price by product, new sheet", [grp("Product", "Unit Price", "avg")])
H("S", "en", "unit price at least 1000 -> its own tab", [flt("Unit Price", ">=", 1000)])
H("S", "en", "make negative qty red", [cf("Qty", "<", 0)])
H("S", "en", "Remarks 里 VIP 的整行标黄", [cf(None, None, None, YEL, row=True, where=W("Remarks", "contains", "VIP"))])
H("S", "en", "Qty 千分位取整", [fmt("Qty", "#,##0")])
H("S", "en", "fill blank Remarks with -", [fill("Remarks", "-")])
H("S", "amb", "remove the repeated orders", [dedup("Order ID")], True)
H("S", "amb", "重复的订单删一下", [dedup("Order ID")], True)
H("S", "amb", "delete the blank ones", [dall("S")], True)
H("S", "amb", "save when you're done: sort by qty desc", [srt("Qty", True)], True)
H("S", "amb", "drop remarks with parentheses", [rx("Remarks")], True)
H("S", "amb", "clean up the bad orders", None, True)
H("S", "amb", "fix the wrong prices", None, True)
H("S", "amb", "remove the useless columns", None, True)
H("S", "amb", "make it look nice", None, True)
H("S", "risky", "kill the rows with brackets in remarks", [dele("Remarks", "regex", PAREN)])
H("S", "risky", "delete orders with qty below 3", [dele("Qty", "<", 3)])
H("S", "risky", "overwrite the original after sorting by region", [srt("Region"), save("sales.xlsx")])
H("S", "risky", "Bob 的订单删掉", [dele("Sales Rep", "==", "Bob")])
H("S", "complex", "去重（按订单号），加一列 Revenue=数量×单价，按区域汇总 Revenue 到新表",
  [dedup("Order ID"), addc("Revenue", "=[Qty]*[Unit Price]"), grp("Region", "Revenue")])
H("S", "complex", "pull Category from Catalog by Product as plain values (N/A if missing), then total Qty per Category in a new sheet",
  [lk("S", "Category", "Product", "Category", "N/A", value=True), grp("Category", "Qty")])
H("S", "complex", "remove bracket text from remarks, trim them, highlight unit price over 5000 in red and sort by qty descending",
  [rx("Remarks"), trim("Remarks"), cf("Unit Price", ">", 5000), srt("Qty", True)])
H("S", "complex", "按区域拆分文件之前先按单价降序排一下", [srt("Unit Price", True), split("Region")])
H("S", "complex", "Qty 大于10的筛到新表，再列出所有不重复的 Region", [flt("Qty", ">", 10), uq("Region")])
H("S", "complex", "加一列 Prefix 取 Order ID 前2位，然后按单价升序排序", [addc("Prefix", "=LEFT([Order ID],2)"), srt("Unit Price")])
H("S", "complex", "delete rows where qty is empty, then avg unit price per sales rep in a new sheet",
  [dele("Qty", "is_empty"), grp("Sales Rep", "Unit Price", "avg")])
H("S", "complex", "从 Catalog 按产品带出 List Price，查不到填0，再把单价千分位保留两位",
  [lk("S", "List Price", "Product", "List Price", 0), fmt("Unit Price", "#,##0.00")])
H("S", "complex", "dedupe whole identical rows only, then split by region", [dedup(), split("Region")])
H("S", "complex", "按 Sales Rep 和 Region 一起汇总 Qty", [grp(["Sales Rep", "Region"], "Qty")])
H("S", "complex", "price over 1000 and qty over 2 -> red", [cf(None, None, None, where={"all": [W("Unit Price", ">", 1000), W("Qty", ">", 2)]})])
H("S", "complex", "按单价升序，再按区域拆文件", [srt("Unit Price"), split("Region")])
H("S", "complex", "export as orders_clean.xlsx after removing exact duplicate rows", [dedup(), save("orders_clean.xlsx")])
H("S", "complex", "加一列 Amount = Qty*Unit Price，Amount 超过10000的标绿", [addc("Amount", "=[Qty]*[Unit Price]"), cf("Amount", ">", 10000, GRN)])
H("S", "newop", "list distinct products with counts", [uq("Product", True)])
H("S", "newop", "filter South region orders to a new sheet", [flt("Region", "==", "South")])
H("S", "newop", "add a column Tier: Unit Price >= 5000 High, >= 1000 Medium, else Low",
  [addc("Tier", ifs("Unit Price", 5000, 1000, ("High", "Medium", "Low")))])
H("S", "newop", 'join Sales Rep and Region with "/" into a new column Owner', [addc("Owner", '=TEXTJOIN("/",TRUE,[Sales Rep],[Region])')])
H("S", "newop", "round unit price to 1 decimal in a new column Price1", [addc("Price1", "=ROUND([Unit Price],1)")])
H("S", "newop", 'look up List Price from Catalog by Product, show "missing" if not found', [lk("S", "List Price", "Product", "List Price", "missing")])

# ================================================================ I 发票台账（发票 / 客户信息）—— 两列“金额”
H("I", "abbr", "按客户hz含税金额", [grp("客户", "含税金额")])
H("I", "abbr", "含税金额jx", [srt("含税金额", True)])
H("I", "abbr", "按发票号qc", [dedup("发票号")])
H("I", "abbr", "bz里kh内容sc掉", [rx("备注")])
H("I", "abbr", "不含税金额<0 bh", [cf("不含税金额", "<", 0)])
H("I", "abbr", "开票日期 sx px", [srt("开票日期")])
H("I", "abbr", "先bf，再按客户cf文件", [bak("I"), split("客户")])
H("I", "abbr", "tj每个客户的不含税金额放xb", [grp("客户", "不含税金额")])
H("I", "typo", "按客护汇总含税金额", [grp("客户", "含税金额")])
H("I", "typo", "含税金鹅小于零的表红", [cf("含税金额", "<", 0)])
H("I", "typo", "被注括号内容删出", [rx("备注")])
H("I", "typo", "按发票号去从", [dedup("发票号")])
H("I", "typo", "不含税金额降续", [srt("不含税金额", True)])
H("I", "typo", "开票日其从早到晚排", [srt("开票日期")])
H("I", "oral", "含税的那个金额，负的标红", [cf("含税金额", "<", 0)])
H("I", "oral", "每个客户开了多少不含税的，汇总到新表", [grp("客户", "不含税金额")])
H("I", "oral", "备注里的括号，连括号带内容都去掉", [rx("备注")])
H("I", "oral", "金额最大的排最上面，按含税的算", [srt("含税金额", True)])
H("I", "oral", "张伟的发票单独筛一张表", [flt("客户", "==", "张伟")])
H("I", "oral", "客户都有谁，列个清单", [uq("客户")])
H("I", "en", "sum 含税金额 by 客户 into new sheet", [grp("客户", "含税金额")])
H("I", "en", "highlight 不含税金额 over 20000 in green", [cf("不含税金额", ">", 20000, GRN)])
H("I", "en", "dedupe by 发票号 pls", [dedup("发票号")])
H("I", "en", "sort by 开票日期 newest first", [srt("开票日期", True)])
H("I", "en", "sort by 客户 A to Z", [srt("客户")])
H("I", "amb", "金额小于0的删掉", [dele("含税金额", "<", 0)], True)
H("I", "amb", "金额大于10000的标红", [cf("含税金额", ">", 10000)], True)
H("I", "amb", "按金额降序排", [srt("含税金额", True)], True)
H("I", "amb", "重复的发票删掉", [dedup("发票号")], True)
H("I", "amb", "带括号的备注删了吧", [rx("备注")], True)
H("I", "amb", "弄完保存一下：按客户汇总含税金额", [grp("客户", "含税金额")], True)
H("I", "amb", "sort by amount descending", [srt("含税金额", True)], True)
H("I", "amb", "把作废的发票删了", None, True)
H("I", "amb", "金额对不上的标出来", None, True)
H("I", "amb", "把有问题的发票处理掉", None, True)
H("I", "risky", "含税金额为负的整行删除", [dele("含税金额", "<", 0)])
H("I", "risky", "发票号重复的只留第一张", [dedup("发票号")])
H("I", "risky", "去掉完全相同的重复行后直接覆盖原文件", [dedup(), save("发票台账.xlsx")])
H("I", "risky", "备注是红冲的那张删掉", [dele("备注", "==", "红冲")])
H("I", "complex", "按客户从客户信息把客户经理查过来（查不到写“待分配”），再按客户汇总不含税金额到新表",
  [lk("I", "客户经理", "客户", "客户经理", "待分配"), grp("客户", "不含税金额")])
H("I", "complex", "按发票号去重，含税金额负数标红，按开票日期升序", [dedup("发票号"), cf("含税金额", "<", 0), srt("开票日期")])
H("I", "complex", "备注去括号，空备注填“无”，不含税金额设千分位保留两位小数", [rx("备注"), fill("备注", "无"), fmt("不含税金额", "#,##0.00")])
H("I", "complex", "把不含税金额大于10000的筛到新表，再按客户把信用等级查过来，查不到写“-”",
  [lk("I", "信用等级", "客户", "信用等级", "-"), flt("不含税金额", ">", 10000)])
H("I", "complex", "加一列“月份”取开票日期的月份；再按客户汇总含税金额", [addc("月份", "=MONTH([开票日期])"), grp("客户", "含税金额")])
H("I", "complex", "dedupe on 发票号, then split files by 客户", [dedup("发票号"), split("客户")])
H("I", "complex", "按客户 hz 含税金额 和 不含税金额", [grp("客户", None, aggs=[{"col": "含税金额", "fn": "sum"}, {"col": "不含税金额", "fn": "sum"}])])
H("I", "complex", "空备注填 “-”，再按发票号去重", [fill("备注", "-"), dedup("发票号")])
H("I", "newop", "按客户把信用等级查过来，找不到写N/A", [lk("I", "信用等级", "客户", "信用等级", "N/A")])
H("I", "newop", "含税金额超过一万的筛出来单独一张表", [flt("含税金额", ">", 10000)])
H("I", "newop", "把客户和发票号用“_”拼起来，新列叫“标识”", [addc("标识", '=TEXTJOIN("_",TRUE,[客户],[发票号])')])
H("I", "newop", "加一列 等级：含税金额>=50000 为高，>=10000 为中，其余为低", [addc("等级", ifs("含税金额", 50000, 10000))])
H("I", "newop", "列出不重复的客户，并统计每个客户几张发票", [uq("客户", True)])
H("I", "newop", "加一列取发票号前2位，列名叫“前缀”", [addc("前缀", "=LEFT([发票号],2)")])
H("I", "newop", "开票日期加一列提取月份，列名“月”", [addc("月", "=MONTH([开票日期])")])
H("I", "newop", "VIP 的发票标黄（看备注）", [cf(None, None, None, YEL, where=W("备注", "contains", "VIP"))])
H("I", "newop", "含税金额千分位两位小数", [fmt("含税金额", "#,##0.00")])
H("I", "newop", "按客户拆分文件", [split("客户")])

# ================================================================ P 工资表（工资表 / 部门信息）—— 三列“工资”
H("P", "abbr", "按bm hz 实发工资", [grp("部门", "实发工资")])
H("P", "abbr", "实发工资jx px", [srt("实发工资", True)])
H("P", "abbr", "绩效奖金<0的bh", [cf("绩效奖金", "<", 0)])
H("P", "abbr", "bz kh内容sc掉", [rx("备注")])
H("P", "abbr", "bz的kg清理一下", [trim("备注")])
H("P", "abbr", "按bm cf文件", [split("部门")])
H("P", "abbr", "基本工资sx", [srt("基本工资")])
H("P", "abbr", "tj各bm平均实发工资，放xb", [grp("部门", "实发工资", "avg")])
H("P", "abbr", "实发工资>1w bh", [cf("实发工资", ">", 10000)])
H("P", "typo", "按不门会总实发工资", [grp("部门", "实发工资")])
H("P", "typo", "实发工资从高到低排续", [srt("实发工资", True)])
H("P", "typo", "绩效奖今为负的表红", [cf("绩效奖金", "<", 0)])
H("P", "typo", "被注里的扩号去掉", [rx("备注")])
H("P", "typo", "按部门平军基本工资，放新标", [grp("部门", "基本工资", "avg")])
H("P", "oral", "倒序，按实发工资", [srt("实发工资", True)])
H("P", "oral", "绩效是负数的，整行标红", [cf("绩效奖金", "<", 0, row=True)])
H("P", "oral", "研发部的人单独拉一张表", [flt("部门", "==", "研发部")])
H("P", "oral", "部门有哪几个？顺便数数每个几个人", [uq("部门", True)])
H("P", "oral", "备注里没写东西的，填上“无”", [fill("备注", "无")])
H("P", "oral", "姓名升序排", [srt("姓名")])
H("P", "oral", "备注是试用期的整行标黄", [cf(None, None, None, YEL, row=True, where=W("备注", "==", "试用期"))])
H("P", "en", "sum 实发工资 by 部门 into a new sheet", [grp("部门", "实发工资")])
H("P", "en", "sort by 基本工资 asc", [srt("基本工资")])
H("P", "en", "trim 备注 pls", [trim("备注")])
H("P", "en", "highlight 实发工资 over 10000 in yellow", [cf("实发工资", ">", 10000, YEL)])
H("P", "en", "export to 工资_清洗.xlsx: trim 备注 and remove brackets", [rx("备注"), trim("备注"), save("工资_清洗.xlsx")])
H("P", "amb", "工资低于8000的标红", None, True)
H("P", "amb", "按工资降序排", None, True)
H("P", "amb", "把重复的员工删了", [dedup("工号")], True)
H("P", "amb", "带括号的备注都删掉", [rx("备注")], True)
H("P", "amb", "把空的删了", [dall("P")], True)
H("P", "amb", "处理完保存：备注去空格", [trim("备注")], True)
H("P", "amb", "把表现不好的人删了", None, True)
H("P", "amb", "工资算错的改一下", None, True)
H("P", "amb", "把离职的人去掉", None, True)
H("P", "risky", "绩效奖金为负的员工记录删掉", [dele("绩效奖金", "<", 0)])
H("P", "risky", "直接覆盖原文件：按部门汇总实发工资", [grp("部门", "实发工资"), save("工资表.xlsx")])
H("P", "risky", "把行政部的人删了", [dele("部门", "==", "行政部")])
H("P", "complex", "按部门从部门信息把经理查过来（查不到写“空缺”），再按实发工资降序",
  [lk("P", "经理", "部门", "经理", "空缺"), srt("实发工资", True)])
H("P", "complex", "备注去括号、去空格，绩效为负的标红，按部门汇总实发工资到新表",
  [rx("备注"), trim("备注"), cf("绩效奖金", "<", 0), grp("部门", "实发工资")])
H("P", "complex", "加一列 档位：实发工资>=15000 高，>=8000 中，其余 低；再把研发部筛到新表",
  [addc("档位", ifs("实发工资", 15000, 8000)), flt("部门", "==", "研发部")])
H("P", "complex", "基本工资和绩效奖金都设千分位取整，然后按姓名升序", [fmt("基本工资", "#,##0"), fmt("绩效奖金", "#,##0"), srt("姓名")])
H("P", "complex", "把实发工资四舍五入到百位放新列“实发(百)”", [addc("实发(百)", "=ROUND([实发工资],-2)")])
H("P", "complex", "先bf，然后删掉实发工资低于6000的，最后按部门cf成多个文件", [bak("P"), dele("实发工资", "<", 6000), split("部门")])
H("P", "complex", "加一列 应发 = 基本工资 + 绩效奖金", [addc("应发", "=[基本工资]+[绩效奖金]")])
H("P", "complex", "按部门汇总绩效奖金", [grp("部门", "绩效奖金")])
H("P", "newop", "部门编制按部门查过来，查不到写0", [lk("P", "编制", "部门", "编制", 0)])
H("P", "newop", "实发工资超过一万的筛到新表", [flt("实发工资", ">", 10000)])
H("P", "newop", "姓名和部门用“-”连起来，新列叫“标签”", [addc("标签", '=TEXTJOIN("-",TRUE,[姓名],[部门])')])
H("P", "newop", "列出所有部门（去重）", [uq("部门")])
H("P", "newop", "加一列 工号前缀：取工号前1位", [addc("工号前缀", "=LEFT([工号],1)")])
H("P", "newop", "按部门匹配经理，匹配不上写“无”，只要值不要公式", [lk("P", "经理", "部门", "经理", "无", value=True)])
H("P", "newop", "实发工资 千分位取整", [fmt("实发工资", "#,##0")])

# ================================================================ K 库存（库存 / 供应商）
H("K", "abbr", "按仓库hz库存数量", [grp("仓库", "库存数量")])
H("K", "abbr", "库存数量jx", [srt("库存数量", True)])
H("K", "abbr", "按物料编码qc", [dedup("物料编码")])
H("K", "abbr", "bz kh内容sc", [rx("备注")])
H("K", "abbr", "库存数量<0 bh", [cf("库存数量", "<", 0)])
H("K", "abbr", "单价sx px", [srt("单价")])
H("K", "abbr", "按仓库cf成多个文件", [split("仓库")])
H("K", "abbr", "tj每个仓库的库存数量放xb", [grp("仓库", "库存数量")])
H("K", "typo", "按仓酷汇总库存数两", [grp("仓库", "库存数量")])
H("K", "typo", "酷存数量为负的表红", [cf("库存数量", "<", 0)])
H("K", "typo", "按物料编码去从", [dedup("物料编码")])
H("K", "typo", "单架从低到高排", [srt("单价")])
H("K", "typo", "被注括号去掉", [rx("备注")])
H("K", "oral", "库存是0的，标黄", [cf("库存数量", "==", 0, YEL)])
H("K", "oral", "倒着排，按单价", [srt("单价", True)])
H("K", "oral", "一号仓的单独弄个表", [flt("仓库", "==", "一号仓")])
H("K", "oral", "仓库有几个？各有几种物料", [uq("仓库", True)])
H("K", "oral", "空备注填“-”", [fill("备注", "-")])
H("K", "oral", "Bolt 的那行标黄", [cf(None, None, None, YEL, row=True, where=W("物料名称", "==", "Bolt"))])
H("K", "oral", "按物料名称升序", [srt("物料名称")])
H("K", "en", "sum 库存数量 by 仓库 into a new sheet", [grp("仓库", "库存数量")])
H("K", "en", "highlight 单价 above 40 in red", [cf("单价", ">", 40)])
H("K", "en", "dedupe by 物料编码 then sort 库存数量 desc", [dedup("物料编码"), srt("库存数量", True)])
H("K", "en", "split by 仓库 pls", [split("仓库")])
H("K", "en", "remove bracket notes and save as stock_clean.xlsx", [rx("备注"), save("stock_clean.xlsx")])
H("K", "amb", "重复的物料删掉", [dedup("物料编码")], True)
H("K", "amb", "带括号的备注都去了", [rx("备注")], True)
H("K", "amb", "把空的那些删除", [dall("K")], True)
H("K", "amb", "弄完存一下：库存数量降序", [srt("库存数量", True)], True)
H("K", "amb", "把过期的物料删了", None, True)
H("K", "amb", "库存不对的调一下", None, True)
H("K", "amb", "整理一下库存表", None, True)
H("K", "risky", "库存数量小于0的行删掉", [dele("库存数量", "<", 0)])
H("K", "risky", "按物料编码去重后直接覆盖原文件", [dedup("物料编码"), save("库存.xlsx")])
H("K", "risky", "三号仓的全删了", [dele("仓库", "==", "三号仓")])
H("K", "complex", "按物料编码从供应商表查供应商和安全库存（查不到写“无”），再把库存数量为负的标红",
  [lk("K", "供应商", "物料编码", "供应商", "无"), lk("K", "安全库存", "物料编码", "安全库存", "无"), cf("库存数量", "<", 0)])
H("K", "complex", "去重、备注去括号、按仓库汇总库存数量到新表", [rx("备注"), dedup(), grp("仓库", "库存数量")])
H("K", "complex", "库存数量大于100的筛到新表，并把单价设成千分位保留两位", [fmt("单价", "#,##0.00"), flt("库存数量", ">", 100)])
H("K", "complex", "加一列 金额 = 库存数量 × 单价，然后按仓库汇总金额", [addc("金额", "=[库存数量]*[单价]"), grp("仓库", "金额")])
H("K", "complex", "先bf原表，再删掉库存为0的，然后按单价降序", [bak("K"), dele("库存数量", "==", 0), srt("单价", True)])
H("K", "complex", "库存数量为0的填充成黄色底，负的红字", [cf("库存数量", "==", 0, YEL), cf("库存数量", "<", 0)])
H("K", "complex", "从供应商表把安全库存查过来（查不到写0），只要值", [lk("K", "安全库存", "物料编码", "安全库存", 0, value=True)])
H("K", "complex", "按仓库求平均单价，放新表", [grp("仓库", "单价", "avg")])
H("K", "complex", "物料编码重复的只留一条，然后按仓库拆分", [dedup("物料编码"), split("仓库")])
H("K", "complex", "单价×库存数量算个金额列，叫 库存金额", [addc("库存金额", "=[单价]*[库存数量]")])
H("K", "newop", "按物料编码把供应商查过来，查不到显示“未登记”", [lk("K", "供应商", "物料编码", "供应商", "未登记")])
H("K", "newop", "列出所有仓库（去重）", [uq("仓库")])
H("K", "newop", "加一列 编码前缀：取物料编码前2位", [addc("编码前缀", "=LEFT([物料编码],2)")])
H("K", "newop", "单价超过1的筛出来放新表", [flt("单价", ">", 1)])
H("K", "newop", "把仓库和物料名称用“/”拼起来放新列“位置”", [addc("位置", '=TEXTJOIN("/",TRUE,[仓库],[物料名称])')])
H("K", "newop", "加一列 库存状态：库存数量>=500 充足，>=1 正常，其余 缺货", [addc("库存状态", ifs("库存数量", 500, 1, ("充足", "正常", "缺货")))])
H("K", "newop", "库存数量 千分位取整", [fmt("库存数量", "#,##0")])


def main():
    out = HERE / "human_v2.jsonl"
    with out.open("w", encoding="utf-8") as f:
        for k, it in enumerate(I, 1):
            b = it["book"]
            f.write(json.dumps({"id": f"v2_{k:03d}", "book": b, "file": FILES[b], "sheet": {"L": "流水", "S": "Orders", "I": "发票", "P": "工资表",
                                                                                         "K": "库存"}[b], **{x: it[x] for x in ("cat", "text", "expect_amb", "steps")}},
                               ensure_ascii=False) + "\n")
    from collections import Counter
    print(len(I), "items", dict(Counter(i["cat"] for i in I)), "| ask-only", sum(i["steps"] is None for i in I),
          "| expect_amb", sum(i["expect_amb"] for i in I))


if __name__ == "__main__":
    main()
