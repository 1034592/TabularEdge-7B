"""通用能力正则化：自写的通用编程 / 办公 / 问答提示（不取任何外部数据集，规避许可问题）。
回答由原版 Qwen2.5-Coder-7B-Instruct 生成（自蒸馏，Apache-2.0），混入最终微调，把模型往原版分布拉回、压住困惑度漂移。
另留 200 条作“通用漂移 ppl”评测（不参与训练）。8 条 regress_texts 不出现在这里。

    python synth/distill_prompts.py --out data/distill/prompts.jsonl
"""
from __future__ import annotations

import argparse
import itertools
import json
import random
from pathlib import Path

PY = ["读取一个 CSV 文件并统计每列的空值个数", "把列表按指定字段分组求和", "判断一个字符串是否是回文", "合并两个字典，遇到重复键时把值相加",
      "实现二分查找", "把嵌套列表展平", "计算两个日期之间的工作日天数", "用正则从文本里提取所有邮箱地址", "实现一个简单的 LRU 缓存",
      "把 Excel 列号（如 AB）转换成数字", "统计一段英文文本里出现频率最高的 10 个单词", "用 pandas 把两个表按订单号左连接",
      "把金额字符串（如“¥1,234.50”）转成浮点数", "生成指定长度的随机密码", "检查括号是否匹配", "把秒数格式化成 时:分:秒"]
PY_EN = ["parse a JSON file and print all top-level keys", "remove duplicate rows from a pandas DataFrame keeping the last one",
         "compute a moving average over a list", "validate an email address with a regex", "convert a list of dicts to CSV",
         "find the second largest number in a list", "rename all .txt files in a folder to .md", "count words in a text file",
         "merge overlapping intervals", "group a list of transactions by month and sum the amounts", "implement a stack with push and pop",
         "read an Excel sheet with openpyxl and print the header row"]
SQL = ["查询每个部门的平均工资，并按平均工资降序排列", "找出订单金额最高的前 5 个客户", "统计每个月的新增用户数", "删除表中重复的记录只保留 id 最小的一条",
       "查询没有下过订单的客户", "计算每个产品的累计销售额"]
SQL_EN = ["find customers who placed more than three orders", "list employees whose salary is above their department average",
          "get the latest order for each customer", "count orders per status"]
XL = ["解释 SUMIFS 和 SUMIF 的参数顺序有什么不同", "VLOOKUP 和 INDEX+MATCH 各有什么优缺点", "条件格式怎么把整行标红", "数据透视表里如何按月分组",
      "怎么冻结首行和首列", "如何用 COUNTIF 统计重复值", "TEXT 函数怎么把日期显示成“2026年01月”", "怎样保护工作表但允许编辑部分单元格",
      "绝对引用和相对引用有什么区别", "怎么批量删除空行"]
XL_EN = ["explain the difference between COUNT, COUNTA and COUNTBLANK", "how do I highlight duplicates with conditional formatting",
         "what does the $ sign mean in a cell reference", "how to split full names into first and last name columns",
         "how to make a dropdown list with data validation", "explain what a circular reference is"]
VBA = ["写一个 VBA 宏，把当前工作表所有空单元格填成 0", "写一个 VBA 宏，按 A 列的值把数据拆分到不同工作表", "写一个 VBA 宏，遍历文件夹里所有 xlsx 并合并第一张表"]
OFFICE = ["帮我写一封邮件，提醒各部门本周五前提交报销单", "把下面这句话改得更正式：这个月预算超了，下个月省着点用", "给新员工写一段 Excel 使用规范说明，三条即可",
          "总结一下做月度财务报表的一般步骤", "写一段会议纪要模板", "解释一下什么是应收账款账龄分析"]
OFFICE_EN = ["write a short email asking the team to update the sales tracker by Friday", "summarize the steps to reconcile a bank statement",
             "give three tips for keeping spreadsheets tidy", "explain what accounts payable aging means"]
LANG = ["Python", "JavaScript", "Go", "Java", "TypeScript", "C#"]
ZH_LANG = ["Python", "JavaScript", "Java", "Go"]
NOUNS = ["销售明细", "员工花名册", "库存台账", "报销记录", "客户名单", "订单表", "考勤表", "发票清单"]
PD_TASKS = ["删除全空的行", "把日期列转成年月", "按部门统计人数", "找出金额异常（超过均值三倍）的记录", "把两列拼成一列", "去掉文本两端空格",
            "按客户汇总金额并排序", "把宽表转成长表"]
FORMULA = ["计算 {c}2:{c}100 中大于 1000 的个数", "取 {c}2 单元格的前 3 个字符", "把 {c}2 的日期显示成星期几", "判断 {c}2 是否为空，空则显示“缺失”",
           "计算 {c}2:{c}100 的中位数", "对 {c}2:{c}100 中的正数求和", "把 {c}2 四舍五入到百位", "返回 {c}2:{c}100 中第二大的值"]
CONCEPT = ["什么是哈希表", "解释一下递归和迭代的区别", "什么是数据库索引，为什么能加速查询", "解释 HTTP 和 HTTPS 的区别", "什么是正则表达式的贪婪匹配",
           "解释一下事务的 ACID", "什么是浮点数精度误差，金额计算该怎么处理", "Git 的 merge 和 rebase 有什么区别", "什么是 CSV 文件的编码问题，怎么避免乱码",
           "解释一下时间复杂度 O(n log n)"]
CONCEPT_EN = ["What is a race condition?", "Explain the difference between a list and a tuple in Python.", "What is normalization in databases?",
              "Explain what an API rate limit is.", "What is the difference between UTF-8 and UTF-16?", "Explain lazy evaluation."]
STYLE = ["", "请附上简短注释。", "只给代码，不要解释。", "顺便说明时间复杂度。"]
STYLE_EN = ["", " Add brief comments.", " Code only, no explanation.", " Also state the time complexity."]


def build(seed=0):
    rng = random.Random(seed)
    out = set()
    for t, lang, s in itertools.product(PY, ZH_LANG, STYLE):
        out.add(f"用 {lang} {t}。{s}")
    for t, n, s in itertools.product(PD_TASKS, NOUNS, ["", "用 pandas 实现。"]):
        out.add(f"有一张{n}，{t}，怎么用 Python 做？{s}")
    for t, c, s in itertools.product(FORMULA, "BCDE", ["", "并解释公式含义。"]):
        out.add("写一个 Excel 公式：" + t.format(c=c) + "。" + s)
    for t, s in itertools.product(CONCEPT, ["", "用通俗的话。", "举个例子。"]):
        out.add(f"{t}？{s}")
    for t, s in itertools.product(CONCEPT_EN, ["", " Use a simple analogy.", " Give a code example."]):
        out.add(t + s)
    for t, lang, s in itertools.product(PY_EN, LANG, STYLE_EN):
        out.add(f"Write a {lang} function to {t}.{s}")
    for t, s in itertools.product(SQL, ["", "用 MySQL 语法。", "用 PostgreSQL 语法。"]):
        out.add(f"写一条 SQL：{t}。{s}")
    for t, s in itertools.product(SQL_EN, ["", " Use PostgreSQL.", " Explain each clause."]):
        out.add(f"Write a SQL query to {t}.{s}")
    for t, s in itertools.product(XL, ["", "举个例子。", "简短回答。"]):
        out.add(f"{t}？{s}")
    for t, s in itertools.product(XL_EN, ["", " Give an example.", " Keep it short."]):
        out.add(f"In Excel, {t}?{s}")
    for t in VBA:
        out.add(t)
    for t, s in itertools.product(OFFICE, ["", "200 字以内。"]):
        out.add(t + ("，" + s if s else ""))
    for t, s in itertools.product(OFFICE_EN, ["", " Under 120 words."]):
        out.add(t + s)
    out = sorted(out)
    rng.shuffle(out)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--heldout", type=int, default=200)
    a = ap.parse_args()
    ps = build()
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as f:
        for i, p in enumerate(ps):
            f.write(json.dumps({"id": i, "split": "heldout" if i < a.heldout else "train", "prompt": p}, ensure_ascii=False) + "\n")
    print(len(ps), "prompts;", a.heldout, "heldout")


if __name__ == "__main__":
    main()
