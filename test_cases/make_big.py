"""生成一张 150MB 级的合成“单位名录”大表（与实测用的真实政务导出表同等规模，内容完全虚构（默认 761,507 行 × 25 列，~330 万条共享字符串），
直接写 OOXML（openpyxl 写 1900 万格要十几分钟）。另生成一张 2,000 行的“商户名单”小表，一半信用代码能在大表里匹配到。

    python test_cases/make_big.py [输出目录] [行数]
"""
import random
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

out = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).parent / "big")
N = int(sys.argv[2]) if len(sys.argv) > 2 else 761_507
out.mkdir(parents=True, exist_ok=True)
rnd = random.Random(804)
HEAD = ["统一社会信用代码", "单位详细名称", "法定代表人", "行业代码", "行业名称", "登记注册类型", "区划代码", "街道（乡镇）",
        "详细地址", "邮政编码", "成立日期", "从业人员期末人数", "营业收入（千元）", "资产总计（千元）", "单位规模", "控股情况",
        "机构类型", "执行会计标准", "联系电话", "是否一套表单位", "所属专业", "统计机关", "数据来源", "更新时间", "备注"]
STREETS = ["东湖街道", "南山街道", "西岭街道", "北塘街道", "中桥街道", "青石街道", "白沙街道", "金川街道", "云溪镇", "柳河镇"]
INDS = [("6513", "应用软件开发"), ("5191", "互联网销售"), ("7212", "投资与资产管理"), ("6421", "互联网搜索服务"), ("5810", "餐饮业"),
        ("8310", "普通高中教育"), ("8391", "职业技能培训"), ("7520", "科技推广服务"), ("6440", "互联网数据服务")]
SUR = "王李张刘陈杨黄赵吴周徐孙马朱胡郭何高林罗郑梁谢宋唐许韩冯邓曹彭曾肖田董袁潘于蒋蔡余杜叶程苏魏吕丁任沈姚卢姜崔钟谭陆汪范金石廖贾夏韦付方白邹孟熊秦邱江尹薛闫段雷侯龙史陶黎贺顾毛郝龚邵万钱严覃武戴莫孔向汤"
GIV = "伟芳娜敏静丽强磊军洋勇艳杰娟涛明超秀霞平刚桂英华玉兰萍红鹏飞宇浩然欣怡子轩梓涵晨阳思远文博佳琪一诺"
CH = "0123456789ABCDEFGHJKLMNPQRTUWXY"


def code(i):
    r = random.Random(i * 7919 + 13)
    return "911101" + "".join(r.choice(CH) for _ in range(11)) + CH[i % 31]


sst, sidx = [], {}


def s(v):
    k = sidx.get(v)
    if k is None:
        k = sidx[v] = len(sst)
        sst.append(v)
    return k


def col(n):
    r = ""
    while n:
        n, m = divmod(n - 1, 26)
        r = chr(65 + m) + r
    return r


L = [col(i) for i in range(1, 26)]
p = out / "单位名录_76万行.xlsx"
with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
    with z.open("xl/worksheets/sheet1.xml", "w", force_zip64=True) as f:
        f.write(f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:Y{N + 1}"/><sheetData>'.encode())
        f.write(('<row r="1">' + "".join(f'<c r="{L[j]}1" t="s"><v>{s(h)}</v></c>' for j, h in enumerate(HEAD)) + "</row>").encode())
        buf = []
        for i in range(1, N + 1):
            r = i + 1
            ind = INDS[i % len(INDS)]
            st = STREETS[(i * 7) % len(STREETS)]
            name = f"示例{''.join(rnd.choice(GIV) for _ in range(2))}{['科技', '网络', '教育', '餐饮', '咨询'][i % 5]}有限公司{i}"
            vals = [("s", code(i)), ("s", name), ("s", rnd.choice(SUR) + rnd.choice(GIV) + rnd.choice(GIV)), ("s", ind[0]), ("s", ind[1]),
                    ("s", "有限责任公司"), ("n", 990101000000 + (i % 29) * 1000), ("s", st), ("s", f"示例市东城区{st[:-2]}路{i % 999}号{i}室"),
                    ("n", 100080 + i % 20), ("d", 36526 + (i * 37) % 9000), ("n", i % 300), ("n", (i * 131) % 99991),
                    ("n", (i * 977) % 999983), ("s", ["小型", "微型", "中型"][i % 3]), ("s", "私人控股"), ("s", "企业"),
                    ("s", "企业会计准则"), ("s", f"010-{60000000 + i}"), ("s", "否"), ("s", ["服务业", "批零", "住餐"][i % 3]),
                    ("s", "示例区统计局"), ("s", "名录库"), ("d", 46000 + i % 200), ("s", "" if i % 4 else f"变更{i}")]
            cells = []
            for j, (t, v) in enumerate(vals):
                if t == "s":
                    if v != "":
                        cells.append(f'<c r="{L[j]}{r}" t="s"><v>{s(v)}</v></c>')
                elif t == "d":
                    cells.append(f'<c r="{L[j]}{r}" s="1"><v>{v}</v></c>')
                else:
                    cells.append(f'<c r="{L[j]}{r}"><v>{v}</v></c>')
            buf.append(f'<row r="{r}">' + "".join(cells) + "</row>")
            if len(buf) >= 5000:
                f.write("".join(buf).encode())
                buf = []
        f.write(("".join(buf) + "</sheetData></worksheet>").encode())
    with z.open("xl/sharedStrings.xml", "w", force_zip64=True) as f:
        f.write(f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="{len(sst)}">'.encode())
        for k in range(0, len(sst), 20000):
            f.write("".join(f"<si><t>{escape(x)}</t></si>" for x in sst[k:k + 20000]).encode())
        f.write(b"</sst>")
    z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
    z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    z.writestr("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="基本单位" sheetId="1" r:id="rId1"/></sheets></workbook>')
    z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
    z.writestr("xl/styles.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="176" formatCode="yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;"/></numFmts><fonts count="1"><font><sz val="11"/><name val="宋体"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="176" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>')
print(p, p.stat().st_size // 2**20, "MB", N, "rows", len(sst), "strings")

from openpyxl import Workbook  # noqa: E402
wb = Workbook()
ws = wb.active
ws.title = "商户名单"
ws.append(["序号", "商户名称", "统一社会信用代码", "联系人"])
r2 = random.Random(5)
for k in range(1, 2001):
    i = r2.randrange(1, N + 1) if k % 2 else N + 10_000 + k  # 奇数行能匹配，偶数行是大表里没有的代码
    c = code(i)
    ws.append([k, f"商户{k}", (" " + c) if k % 97 == 0 else c, f"联系人{k}"])  # 少量带空格：匹配时应去首尾空格
wb.save(out / "商户名单.xlsx")
print(out / "商户名单.xlsx")
