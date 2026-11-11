"""生成 10 个工业级中文 Excel 实测场景（真实 .xlsx + scenarios.json 指令脚本）。

数据全部为程序合成（固定随机种子），模拟真实企业表格里的“恶心”之处：
跨文件主数据、中英文括号混杂、文本型日期与金额、重复订单、空行、离职人员、负数、超期等。

    python test_cases/make_cases.py [输出目录]
"""
from __future__ import annotations

import datetime as dt
import json
import random
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

R = random.Random(20261105)
SURN = "王李张刘陈杨黄赵周吴徐孙马朱胡郭何高林罗郑梁谢宋唐许韩冯邓曹彭曾萧田董袁潘蒋蔡余杜叶程苏魏吕丁任沈姚卢姜崔钟谭陆汪范金石廖贾夏韦付方白邹孟熊秦邱江尹薛闫段雷侯龙史陶黎贺顾毛郝龚邵万钱严覃武戴莫孔向汤"
GIVEN = ["伟", "芳", "娜", "敏", "静", "丽", "强", "磊", "军", "洋", "勇", "艳", "杰", "娟", "涛", "明", "超", "秀英", "霞", "平",
         "刚", "桂英", "建华", "晓东", "子涵", "浩然", "雨桐", "梓萱", "思远", "一鸣", "欣怡", "宇轩", "嘉怡", "俊杰", "诗琪"]
HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")


def name():
    return R.choice(SURN) + R.choice(GIVEN)


def save(rows, path: Path, sheet: str, widths=None, extra=None, numfmt=None):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    _fill(ws, rows, widths, numfmt)
    for title, erows in (extra or []):
        _fill(wb.create_sheet(title), erows, None, None)
    wb.save(path)


def _fill(ws, rows, widths, numfmt):
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    for i, h in enumerate(rows[0], 1):
        w = (widths or {}).get(h) or max(8, min(28, max(len(str(r[i - 1])) if i - 1 < len(r) and r[i - 1] is not None else 0
                                                        for r in rows[:60]) * 1.6 + 2))
        ws.column_dimensions[get_column_letter(i)].width = w
    for h, fmt in (numfmt or {}).items():
        c = rows[0].index(h) + 1
        for r in range(2, ws.max_row + 1):
            ws.cell(r, c).number_format = fmt


# ---------------------------------------------------------------- 01 跨文件提成
def case01(d: Path):
    depts = ["华东销售一部", "华东销售二部", "华北大客户部", "华南渠道部", "西南拓展部"]
    levels = ["P3", "P4", "P5", "P6", "M1"]
    roster = [["工号", "姓名", "部门", "职级", "入职日期"]]
    ids = [f"SR{1000 + i}" for i in range(1, 49)]
    names = {}
    for i in ids:
        names[i] = name()
        roster.append([i, names[i], R.choice(depts), R.choice(levels), dt.date(2019 + R.randint(0, 6), R.randint(1, 12), R.randint(1, 28))])
    save(roster, d / "01_员工花名册.xlsx", "花名册", numfmt={"入职日期": "yyyy-mm-dd"})
    tiers = [["净销售额下限", "提成比例", "档位说明"], [0, 0.01, "保底档"], [50000, 0.02, "达标档"], [100000, 0.03, "优秀档"],
             [200000, 0.045, "卓越档"], [400000, 0.06, "封顶档"]]
    save(tiers, d / "01_提成档位.xlsx", "提成档位", numfmt={"提成比例": "0.0%"})
    perf = [["工号", "销售员", "销售额", "退货额", "回款状态", "备注"]]
    sel = R.sample(ids, 40) + ["SR2001", "SR2002"]  # 两个不在花名册（新人未登记）
    for i in sel:
        s = R.choice([R.randint(8000, 60000), R.randint(60000, 180000), R.randint(180000, 520000)])
        perf.append([i, names.get(i, name()), s, R.choice([0, 0, 0, R.randint(500, 20000)]),
                     R.choice(["已回款", "已回款", "已回款", "部分回款", "未回款"]), R.choice([None, None, "大客户", "含跨月单", "补录"])])
    save(perf, d / "01_销售业绩.xlsx", "10月业绩", numfmt={"销售额": "#,##0", "退货额": "#,##0"})
    return {"id": 1, "title": "跨文件工号匹配 + 提成档位动态查找 + 条件抵扣", "file": "01_销售业绩.xlsx", "sheet": "10月业绩",
            "refs": ["01_员工花名册.xlsx", "01_提成档位.xlsx"],
            "steps": [
                {"text": "根据员工花名册按工号把部门和职级匹配过来，花名册里找不到的写“未登记”",
                 "text_en": "Look up department and level from the employee roster (员工花名册) by 工号; write 未登记 when not found",
                 "expect": "新增 部门、职级 两列（XLOOKUP 活公式引用花名册表），SR2001/SR2002 显示“未登记”"},
                {"text": "加一列净销售额=销售额-退货额",
                 "expect": "净销售额活公式列"},
                {"text": "加一列提成比例=VLOOKUP(净销售额,提成档位!A:B,2,TRUE)",
                 "expect": "提成比例列：按净销售额近似匹配档位（找不超过净销售额的最大档位下限），活公式"},
                {"text": "加一列应发提成：回款状态是未回款的记0，否则等于净销售额×提成比例",
                 "expect": "应发提成活公式列；未回款为 0"},
            ]}


# ---------------------------------------------------------------- 02 银行流水清洗
def case02(d: Path):
    rows = [["交易时间", "摘要", "对方户名", "金额", "备注"]]
    cos = ["上海华拓科技有限公司", "北京云杉数据服务有限公司", "杭州青禾贸易有限公司", "深圳市鹏远电子有限公司", "成都蜀信物流有限公司",
           "广州恒泰建材有限公司"]
    t = dt.datetime(2026, 10, 1, 9, 0)
    for i in range(120):
        t += dt.timedelta(minutes=R.randint(20, 600))
        so = f"SO{t:%Y%m%d}-{R.randint(1, 60):02d}"
        kind = R.random()
        if kind < 0.55:
            amt, summ = R.randint(800, 98000), R.choice(["货款", "  货款 ", "货款收入", "预收货款"])
            note = R.choice([f"货款（订单号：{so}）", f"货款({so})", f"{so} 尾款【含税】", f"订单{so}（分批到账）", f"[{so}] 首付款"])
        elif kind < 0.75:
            amt, summ = -R.randint(300, 30000), R.choice(["采购付款", "  付款", "供应商付款 "])
            note = R.choice([f"采购单 PO{R.randint(10000, 99999)}（月结）", "付款(加急)", f"付款 {so}"])
        elif kind < 0.87:
            amt, summ = -R.randint(100, 12000), R.choice(["退款", "退款 ", "客户退款"])
            note = R.choice([f"退款（{so}）", f"退款-{so}(质量问题)", f"Refund {so}"])
        else:
            amt, summ = -round(R.uniform(2, 60), 2), R.choice(["手续费", " 手续费", "跨行手续费"])
            note = R.choice(["【手续费】", "(银行扣费)", None])
        rows.append([t, summ, R.choice(cos) if "手续费" not in summ else "中国工商银行", amt, note])
    save(rows, d / "02_银行流水.xlsx", "10月流水", numfmt={"交易时间": "yyyy-mm-dd hh:mm", "金额": "#,##0.00"})
    return {"id": 2, "title": "杂乱流水清洗：正则提取订单号 + 中英文括号清理 + 剔除退款", "file": "02_银行流水.xlsx", "sheet": "10月流水",
            "steps": [
                {"text": "从备注里把SO开头的订单号（比如 SO20261003-17）提取出来，单独放一列“订单号”，没有的留空",
                 "expect": "新增 订单号 列：正则提取 SO+8位日期-2位序号（Excel 2019 无 REGEXEXTRACT → 写入提取结果）"},
                {"text": "备注里中英文括号连同括号里的内容都去掉，摘要去掉多余空格",
                 "expect": "备注中 （…）(…) 整段清除（订单号已在上一步提取保存）；摘要首尾空格去掉"},
                {"text": "摘要里带“退款”的行全部删掉",
                 "expect": "弹窗确认后删除退款行；结果另存，原表不动"},
            ]}


# ---------------------------------------------------------------- 03 加权汇总
def case03(d: Path):
    rows = [["部门", "姓名", "季度", "目标额", "完成额", "权重系数"]]
    depts = ["销售一部", "销售二部", "渠道部", "大客户部"]
    people = {dp: [name() for _ in range(5)] for dp in depts}
    for q in ["Q1", "Q2", "Q3", "Q4"]:
        for dp in depts:
            for p in people[dp]:
                tgt = R.choice([200000, 300000, 400000, 500000])
                rows.append([dp, p, q, tgt, int(tgt * R.uniform(0.45, 1.35)), R.choice([0.8, 1.0, 1.0, 1.2, 1.5])])
    save(rows, d / "03_季度绩效.xlsx", "绩效明细", numfmt={"目标额": "#,##0", "完成额": "#,##0"})
    return {"id": 3, "title": "多条件加权汇总：达标率阈值 + 分部门分季度活公式汇总", "file": "03_季度绩效.xlsx", "sheet": "绩效明细",
            "steps": [
                {"text": "加一列达标率=完成额/目标额，再加一列核算奖金：达标率大于等于80%的按完成额×权重系数×5%算，不到80%的记0",
                 "expect": "达标率、核算奖金两列活公式；达标率<0.8 的奖金为 0"},
                {"text": "按部门和季度汇总核算奖金和完成额，放新表“部门季度汇总”，用公式，明细改了汇总自动更新",
                 "text_en": "Sum the bonus (核算奖金) and 完成额 by department and quarter into a new sheet 部门季度汇总, as live formulas",
                 "expect": "新表 16 行（4 部门 × 4 季度），SUMIFS 活公式"},
            ]}


# ---------------------------------------------------------------- 04 去重 + 按大区拆文件
def case04(d: Path):
    prov = {"华东": ["上海", "江苏", "浙江", "安徽", "山东"], "华北": ["北京", "天津", "河北", "山西"], "华南": ["广东", "广西", "海南", "福建"],
            "西南": ["四川", "重庆", "云南", "贵州"], "华中": ["湖北", "湖南", "河南"]}
    rows = [["订单号", "客户名称", "省份", "大区", "产品线", "金额", "下单日期"]]
    base = []
    for i in range(160):
        rg = R.choice(list(prov))
        base.append([f"DD{260000 + i}", name() + R.choice(["商贸", "科技", "实业", "电子"]) + "有限公司", R.choice(prov[rg]), rg,
                     R.choice(["工业网关", "边缘服务器", "传感器套件", "运维服务"]), R.randint(3000, 260000),
                     dt.date(2026, R.randint(7, 10), R.randint(1, 28))])
    dup = R.sample(base, 23)  # 系统重复导出的整行重复
    allr = base + [list(x) for x in dup]
    R.shuffle(allr)
    save(rows + allr, d / "04_客户订单.xlsx", "订单汇总", numfmt={"金额": "#,##0", "下单日期": "yyyy-mm-dd"})
    return {"id": 4, "title": "脏数据去重 + 按大区拆分成独立文件", "file": "04_客户订单.xlsx", "sheet": "订单汇总",
            "steps": [
                {"text": "把完全重复的行去掉，然后按大区拆成单独的文件，放到“按大区拆分”文件夹里",
                 "text_en": "Remove fully duplicated rows, then split into one file per 大区 under the folder 按大区拆分",
                 "expect": "弹窗确认去重（23 行重复）；生成 按大区拆分/华东.xlsx 等 5 个文件，每个文件带表头"},
            ]}


# ---------------------------------------------------------------- 05 动态高亮
def case05(d: Path):
    rows = [["项目编号", "项目名称", "负责人", "预算余额", "超期天数", "完成度"]]
    for i in range(60):
        rows.append([f"PRJ-{2026}{i + 1:03d}", R.choice(["智慧园区", "产线改造", "ERP升级", "数据中台", "安防集成", "能耗监测"]) + f"{R.randint(1, 9)}期",
                     name(), R.choice([R.randint(-80000, -500), R.randint(1000, 300000), R.randint(1000, 300000)]),
                     R.choice([0, 0, R.randint(1, 30), R.randint(31, 120)]), round(R.uniform(0.05, 1.0), 3)])
    save(rows, d / "05_项目进度.xlsx", "项目台账", numfmt={"预算余额": "#,##0", "完成度": "0.0%"})
    return {"id": 5, "title": "活规则动态高亮：负数标红 / 超期标黄 / 前10%浅绿加粗", "file": "05_项目进度.xlsx", "sheet": "项目台账",
            "steps": [
                {"text": "预算余额是负数的标红，超期天数大于30天的整行标黄",
                 "expect": "两条条件格式规则（活规则，改数自动变色）"},
                {"text": "加一列完成度百分位=PERCENTRANK.INC(完成度整列,完成度)，完成度百分位大于等于0.9的标浅绿色背景、字体加粗",
                 "expect": "新增百分位活公式列 + 条件格式 ≥0.9 浅绿加粗；改完成度后自动重新排名变色"},
            ]}


# ---------------------------------------------------------------- 06 非标日期金额
def case06(d: Path):
    rows = [["报销单号", "报销人", "费用类别", "报销日期", "金额"]]
    for i in range(80):
        day = dt.date(2026, 9, 1) + dt.timedelta(days=R.randint(0, 60))
        fmt = R.choice([f"{day.year}.{day.month}.{day.day}", f"{day.year}年{day.month}月{day.day}日", f"{day.year}/{day.month}/{day.day}",
                        f"{day:%Y-%m-%d}", f"{day.year}年{day.month:02d}月{day.day:02d}日"])
        amt = round(R.uniform(30, 6800), 2)
        a = R.choice([f"¥{amt:,.2f}", f"{amt:,.2f}元", f"￥{amt:.2f}", f"{amt:,.2f}", f"{amt}元"])
        rows.append([f"BX{2609000 + i}", name(), R.choice(["差旅", "交通", "餐饮", "办公", "招待"]), fmt, a])
    save(rows, d / "06_报销明细.xlsx", "报销明细")
    return {"id": 6, "title": "非标准日期与金额规范化 + 提取年月 + 按周聚合", "file": "06_报销明细.xlsx", "sheet": "报销明细",
            "steps": [
                {"text": "金额列里的¥、￥、逗号和“元”都去掉",
                 "expect": "金额变成纯数字文本"},
                {"text": "报销日期里的“年”“月”“.”“/”都替换成“-”，再把“日”删掉",
                 "expect": "日期统一成 2026-9-5 这种文本格式"},
                {"text": "加一列金额数值=VALUE(金额)，加一列标准日期=DATEVALUE(报销日期)，再加一列年月=TEXT(标准日期,\"yyyy-mm\")，再加一列周次=ISOWEEKNUM(标准日期)",
                 "expect": "四列活公式；周次是 ISO 周号"},
                {"text": "按周次汇总金额数值，放新表“按周汇总”",
                 "expect": "新表按周次分组的 SUMIFS 活公式"},
            ]}


# ---------------------------------------------------------------- 07 重复订单留最新
def case07(d: Path):
    rows = [["订单号", "客户", "金额", "状态", "更新时间"]]
    t0 = dt.datetime(2026, 10, 1, 8, 0)
    recs = []
    for i in range(70):
        oid = f"OD{8800 + i}"
        cust = name()
        n = R.choice([1, 1, 1, 2, 3])
        t = t0 + dt.timedelta(hours=R.randint(0, 600))
        amt = R.randint(500, 50000)
        for k in range(n):
            st = ["已下单", "已发货", "已签收"][min(k, 2)]
            recs.append([oid, cust, amt + (R.choice([0, 0, -100, 200]) if k else 0), st, t])
            t += dt.timedelta(hours=R.randint(2, 72), minutes=R.randint(0, 59))
    R.shuffle(recs)
    save(rows + recs, d / "07_订单变更记录.xlsx", "订单", numfmt={"更新时间": "yyyy-mm-dd hh:mm", "金额": "#,##0"})
    return {"id": 7, "title": "复合去重与逆向查找：重复订单保留最新一条，其余归档", "file": "07_订单变更记录.xlsx", "sheet": "订单",
            "steps": [
                {"text": "加一列“是否最新”：同一订单号里更新时间等于该订单最大更新时间的标“保留”，其他标“归档”（用MAXIFS）",
                 "expect": "活公式 =IF(更新时间=MAXIFS(更新时间列,订单号列,本行订单号),\"保留\",\"归档\")"},
                {"text": "把是否最新等于归档的行筛选到新表“归档记录”",
                 "expect": "新表 归档记录（静态值）"},
                {"text": "然后把订单表里是否最新等于归档的行删掉",
                 "expect": "弹窗确认后删除；每个订单号只剩最新一条"},
            ]}


# ---------------------------------------------------------------- 08 动态供应商清单
def case08(d: Path):
    sup = ["苏州恒立精密", "宁波海天塑机", "东莞立讯配件", "无锡先导智能", "深圳汇川技术", "常州星宇车灯", "合肥国轩材料", "厦门宏发电声"]
    rows = [["采购单号", "供应商", "物料", "数量", "单价", "采购日期"]]
    for i in range(150):
        rows.append([f"PO{70100 + i}", R.choice(sup), R.choice(["轴承", "伺服电机", "继电器", "线束", "外壳", "电芯", "PCB板"]),
                     R.randint(10, 800), round(R.uniform(3, 900), 2), dt.date(2026, R.randint(7, 10), R.randint(1, 28))])
    save(rows, d / "08_采购明细.xlsx", "采购", numfmt={"采购日期": "yyyy-mm-dd", "单价": "#,##0.00"})
    return {"id": 8, "title": "动态数组：一键拉出不重复供应商名单 + 实时统计", "file": "08_采购明细.xlsx", "sheet": "采购",
            "steps": [
                {"text": "用动态公式拉出不重复的供应商名单放到新表“供应商清单”，带上每家的采购次数",
                 "text_en": "Use a dynamic formula to list distinct suppliers (供应商) in a new sheet 供应商清单 with the number of orders each",
                 "expect": "UNIQUE 动态数组 + COUNTIF 活公式，8 家供应商"},
                {"text": "加一列采购金额=数量×单价",
                 "expect": "采购金额活公式列"},
                {"text": "按供应商汇总采购金额，放新表“供应商金额”",
                 "expect": "SUMIFS 汇总表（采购金额是公式列，汇总照样是活的）"},
            ]}


# ---------------------------------------------------------------- 09 交叉汇总
def case09(d: Path):
    stores = {"华东": ["上海徐汇店", "上海浦东店", "杭州西湖店", "南京新街口店"], "华北": ["北京朝阳店", "北京西城店", "天津和平店"],
              "华南": ["广州天河店", "深圳南山店", "深圳福田店"]}
    months = ["1月", "2月", "3月", "4月", "5月", "6月"]
    rows = [["大区", "门店", "品类"] + months]
    for rg, ss in stores.items():
        for s in ss:
            for cat in ["饮品", "烘焙", "轻食"]:
                rows.append([rg, s, cat] + [R.randint(20000, 160000) for _ in months])
    save(rows, d / "09_门店月度销售.xlsx", "上半年", numfmt={m: "#,##0" for m in months})
    return {"id": 9, "title": "混合表格交叉汇总：行合计 + 大区×品类二维汇总 + 大区×月份矩阵", "file": "09_门店月度销售.xlsx", "sheet": "上半年",
            "steps": [
                {"text": "加一列上半年合计=1月到6月求和",
                 "expect": "每行 SUM 活公式"},
                {"text": "按大区和品类汇总上半年合计，放新表“大区品类汇总”",
                 "expect": "大区×品类 9 行 SUMIFS 活公式"},
                {"text": "按大区汇总1月、2月、3月、4月、5月、6月，放新表“大区月度矩阵”",
                 "expect": "大区为行、6 个月为列的交叉矩阵（SUMIFS）；注意：纯行列转置（TRANSPOSE 整表）不在 v0.2 操作集里"},
            ]}


# ---------------------------------------------------------------- 10 防灾压力测试
def case10(d: Path):
    rows = [["工号", "姓名", "部门", "基本工资", "绩效奖金", "实发工资", "备注"]]
    for i in range(45):
        base = R.choice([6500, 8000, 9500, 12000, 15000, 18000])
        bonus = R.randint(-2000, 6000)
        rows.append([f"E{3000 + i}", name(), R.choice(["研发部", "销售部", "财务部", "行政部", "生产部"]), base, bonus,
                     base + bonus - R.randint(300, 1500), R.choice([None, None, None, "已离职", "试用期", "已离职（10月）", "调岗"])])
    for k in R.sample(range(5, 45), 4):  # 整行空白（复制粘贴遗留）
        rows.insert(k, [None] * 7)
    save(rows, d / "10_工资表.xlsx", "10月工资", numfmt={"基本工资": "#,##0", "绩效奖金": "#,##0", "实发工资": "#,##0"})
    return {"id": 10, "title": "防灾压力测试：错别字 / 倒装句 / 高危口语（删除确认 + 安全另存）", "file": "10_工资表.xlsx", "sheet": "10月工资",
            "steps": [
                {"text": "把空的都给我删咯",
                 "expect": "模型应追问：删整行全空的行，还是删某列为空的行？删除前弹窗确认"},
                {"text": "实发工资低于8000的给我表红，顺便按不门汇总下实发工资",
                 "expect": "错别字（表红/不门）也能理解：条件格式标红 + 按部门汇总实发工资"},
                {"text": "备注含“离职”的行都给我删了，弄完直接存原来那个文件",
                 "expect": "删除含“离职”的行 + 覆盖原文件：两次高危确认；覆盖前自动备份到 _backup"},
            ]}


CASES = [case01, case02, case03, case04, case05, case06, case07, case08, case09, case10]


def main(out: Path):
    out.mkdir(parents=True, exist_ok=True)
    scen = [f(out) for f in CASES]
    (out / "scenarios.json").write_text(json.dumps(scen, ensure_ascii=False, indent=1), encoding="utf-8")
    print("wrote", len(scen), "scenarios to", out)


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent)
