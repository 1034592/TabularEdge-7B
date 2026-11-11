import sys, subprocess, shutil, tempfile, datetime as dt
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from openpyxl import load_workbook
from tabularedge.formula_eval import Evaluator, XLError, to_serial
G = Path(sys.argv[1]); TC = Path(__file__).resolve().parent
def vals(p, sheet):
    wb = load_workbook(p); ws = wb[sheet]; ev = Evaluator(wb)
    h = [c.value for c in ws[1]]
    return h, [[ev.cell(ws, r, c) for c in range(1, ws.max_column+1)] for r in range(2, ws.max_row+1)], wb
def last(d, stem):
    fs = [f for f in (G / d).glob(stem + '*.xlsx')]
    return max(fs, key=lambda f: (f.name.count('_处理后'), f.stat().st_mtime))
ok = lambda c, m: print(("PASS " if c else "FAIL ") + m)
# case1
h, rows, _ = vals(last('case01_gold', '01_销售业绩'), '10月业绩')
ro = {r[0]: r for r in load_workbook(TC/'01_员工花名册.xlsx').active.iter_rows(min_row=2, values_only=True)}
tiers = [(0,.01),(50000,.02),(100000,.03),(200000,.045),(400000,.06)]
bad = 0
for r in rows:
    d = dict(zip(h, r)); net = d['销售额'] - d['退货额']
    rate = max(t for t in tiers if t[0] <= net)[1]
    exp = 0 if d['回款状态'] == '未回款' else net * rate
    dept = ro[d['工号']][2] if d['工号'] in ro else '未登记'
    bad += abs(d['应发提成'] - exp) > 1e-6 or d['部门'] != dept or abs(d['提成比例'] - rate) > 1e-9
ok(bad == 0, f'case1 提成/部门/档位 逐行核对 {len(rows)} 行, 错 {bad}')
# case2
h, rows, _ = vals(last('case02_gold', '02_银行流水'), '10月流水')
src = list(load_workbook(TC/'02_银行流水.xlsx').active.iter_rows(min_row=2, values_only=True))
import re
exp_n = sum(1 for r in src if '退款' not in r[1])
keep = [r for r in src if '退款' not in r[1]]
ids_ok = len(keep) == len(rows) and all((dict(zip(h, r))['订单号'] or '') == (re.search(r'SO\d{8}-\d{2}', o[4] or '') or [''])[0] for r, o in zip(rows, keep))
ok(len(rows) == exp_n and ids_ok and not any(re.search(r'[()（）【】\[\]]', str(dict(zip(h,r))['备注'] or '')) for r in rows), f'case2 剩 {len(rows)}/{exp_n} 行, 订单号提取正确={ids_ok}')
# case3
h, rows, wb = vals(last('case03_gold', '03_季度绩效'), '绩效明细')
src = list(load_workbook(TC/'03_季度绩效.xlsx').active.iter_rows(min_row=2, values_only=True))
from collections import defaultdict
agg = defaultdict(float)
for d_, n, q, t, c, w in src: agg[(d_, q)] += c*w*0.05 if c/t >= 0.8 else 0
ws = wb['部门季度汇总']; ev = Evaluator(wb)
got = {(ws.cell(r,1).value, ws.cell(r,2).value): ev.cell(ws, r, 3) for r in range(2, ws.max_row+1)}
ok(len(got) == 16 and all(abs(got[k]-v) < 1e-6 for k, v in agg.items()), f'case3 16 组奖金汇总一致')
# case4
src = list(load_workbook(TC/'04_客户订单.xlsx').active.iter_rows(min_row=2, values_only=True))
uniq = list(dict.fromkeys(src))
files = sorted((G/'case04_gold/按大区拆分').glob('*.xlsx'))
n = sum(load_workbook(f).active.max_row - 1 for f in files)
ok(len(files) == 5 and n == len(uniq) == 160, f'case4 拆出 {len(files)} 个文件, 共 {n} 行 (去重后应 {len(uniq)})')
# case5: top10% marks
h, rows, wb = vals(last('case05_gold', '05_项目进度'), '项目台账')
comp = sorted(r[5] for r in rows)
top = [r for r in rows if r[6] >= 0.9]
ok(len(top) == 6 and sorted(r[5] for r in top) == comp[-6:], f'case5 前10% 标记 {len(top)} 行 / {len(rows)}')
# case6
h, rows, wb = vals(last('case06_gold', '06_报销明细'), '报销明细')
src = list(load_workbook(TC/'06_报销明细.xlsx').active.iter_rows(min_row=2, values_only=True))
wk = defaultdict(float)
for r in src:
    y, m, d_ = map(int, re.findall(r'\d+', r[3])) if not r[3].isdigit() else (0,0,0)
    amt = float(re.sub(r'[¥￥,元]', '', r[4])); wk[dt.date(y, m, d_).isocalendar()[1]] += amt
ws = wb['按周汇总']; ev = Evaluator(wb)
got = {ws.cell(r,1).value: ev.cell(ws, r, 2) for r in range(2, ws.max_row+1)}
ok(all(abs(got[k]-v) < 1e-6 for k, v in wk.items()) and len(got) == len(wk), f'case6 {len(wk)} 个周次汇总一致')
# case7
h, rows, wb = vals(last('case07_gold', '07_订单变更记录'), '订单')
src = list(load_workbook(TC/'07_订单变更记录.xlsx').active.iter_rows(min_row=2, values_only=True))
latest = {}
for r in src:
    if r[0] not in latest or r[4] > latest[r[0]][4]: latest[r[0]] = r
arch = wb['归档记录'].max_row - 1
ok(len(rows) == len(latest) and sorted(r[0] + str(round(r[4], 9)) for r in rows) == sorted(k + str(round(to_serial(v[4]), 9)) for k, v in latest.items()) and arch == len(src) - len(latest),
   f'case7 保留 {len(rows)} 条最新 / 归档 {arch} 条')
# case8/9/10 quick
_, rows, wb = vals(last('case08_gold', '08_采购明细'), '采购')
ok(sum(1 for r in range(2, wb['供应商清单'].max_row + 1) if wb['供应商清单'].cell(r, 1).value is not None) in (8, 28), 'case8 供应商清单（2019 版为 8 家 + 20 行预留公式）')
_, rows, wb = vals(last('case09_gold', '09_门店月度销售'), '上半年')
ws = wb['大区月度矩阵']; ev = Evaluator(wb)
tot = sum(ev.cell(ws, r, c) for r in range(2, ws.max_row+1) for c in range(2, 8)); ok(abs(tot - sum(r[9] for r in rows)) < 1e-6 and ws.max_row == 4, 'case9 矩阵总和=行合计总和')
_, rows, wb = vals(last('case10_gold', '10_工资表'), '10月工资')
bk = list((G/'case10_gold/_backup').glob('*.xlsx'))
ok(len(rows) == 45 - sum(1 for r in load_workbook(TC/'10_工资表.xlsx').active.iter_rows(min_row=2, values_only=True) if r[6] and '离职' in r[6]) and len(bk) == 1,
   f'case10 剩 {len(rows)} 行, 覆盖前备份 {len(bk)} 个')
