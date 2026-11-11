"""大表界面响应测试（Windows 宿主机）：工作目录里放 150MB 的大表（例如 test_cases/make_big.py 生成的《单位名录_76万行.xlsx》）+ 一张小名单，
真实创建 TabularEdge 窗口，用 10ms 一次的 after() 心跳测界面主线程最长卡顿（Windows 判“未响应”的阈值是 5 秒）。

    python -m client.gui_bigtest <工作目录> [大表文件名] [小表文件名]
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

import client.app as A

wd = Path(sys.argv[1])
BIG = sys.argv[2] if len(sys.argv) > 2 else "单位名录_76万行.xlsx"
SMALL = sys.argv[3] if len(sys.argv) > 3 else "待匹配名单.xlsx"
A.CONFIG_PATH = wd / "_gui_config.json"
popups, statuses = [], []
A.messagebox.askyesno = lambda title, body, **kw: popups.append(("askyesno", title, body[:80])) or True
A.messagebox.showinfo = lambda *a, **kw: popups.append(("info",) + tuple(str(x)[:160] for x in a))
A.messagebox.showerror = lambda *a, **kw: popups.append(("error",) + tuple(str(x)[:160] for x in a))

HT: dict = {}  # 每类队列消息在界面线程里的处理耗时：[次数, 总计, 最长]
_orig_get = None


def _timed(name, fn):
    def w(*a):
        t = time.perf_counter()
        try:
            return fn(*a)
        finally:
            d = time.perf_counter() - t
            h = HT.setdefault(name, [0, 0.0, 0.0])
            h[0] += 1
            h[1] += d
            h[2] = max(h[2], d)
    return w


root = A.tk.Tk()
app = A.App(root)
for _n in ("delta", "summary", "plan", "ctx", "status", "repair", "engine", "profile", "ran", "error"):
    setattr(app, "_on_" + _n, _timed(_n, getattr(app, "_on_" + _n)))
_orig_status = app._on_status
app._on_status = lambda m: (statuses.append((round(time.time() - T0[0], 2), m)), _orig_status(m))
T0 = [time.time()]
hb = {"last": time.time(), "max": 0.0}


def tick():
    now = time.time()
    hb["max"] = max(hb["max"], now - hb["last"])
    hb["last"] = now
    root.after(10, tick)


def pump(cond, timeout=600):
    t0 = time.time()
    while not cond():
        root.update()
        time.sleep(0.01)
        if time.time() - t0 > timeout:
            raise TimeoutError("ui wait")


def reset():
    root.update()
    hb["last"], hb["max"] = time.time(), 0.0
    T0[0] = time.time()
    statuses.clear()


tick()
pump(lambda: app.engine.state.status in ("ready", "failed"), 300)
root.update()
print("engine:", app.engine.state.status, app.engine.state.detail, app.w_eng.cget("text"), flush=True)
out = {}

# 1) 选中 150MB 大表：表名 + 表结构面板
reset()
t0 = time.time()
app.set_workdir(str(wd), select=BIG)
app.on_file()
call_s = time.time() - t0
pump(lambda: "…" not in app.w_prof.get("1.0", "end"))
out["select_big"] = {"call_returned_s": round(call_s, 3), "profile_shown_s": round(time.time() - t0, 3),
                     "max_ui_stall_s": round(hb["max"], 3), "panel_head": app.w_prof.get("1.0", "3.0").strip()[:120]}
print(json.dumps(out["select_big"], ensure_ascii=False), flush=True)


def parse(text, typing_s=4.0):
    """模拟用户输入：文字进输入框（触发一次 KeyRelease 预填），停顿 typing_s 秒后回车。"""
    app.w_inp.delete("1.0", "end")
    app.w_inp.insert("end", text)
    app._on_typing()
    t_type = time.time()
    pump(lambda: time.time() - t_type > typing_s)
    reset()
    t0 = time.time()
    app.do_parse()
    call_s = time.time() - t0
    pump(lambda: app.v_sum.get().rstrip("…") != app.t("understood") or not app.busy)
    first_text = time.time() - t0
    pump(lambda: not app.busy)
    return {"do_parse_returned_s": round(call_s, 3), "first_understood_text_s": round(first_text, 2),
            "plan_done_s": round(time.time() - t0, 2), "max_ui_stall_s": round(hb["max"], 3), "info": app.v_info.get()[:90], "summary": app.v_sum.get()[:100],
            "steps": [app.w_plan.item(x, "values")[:2] for x in app.w_plan.get_children()], "statuses": statuses[:4]}


# 2) 大表本身作为主表：解析要快；执行会被大表保护拦下并给出建议（不会尝试全量载入）
out["parse_big_main"] = parse("把状态为正常的单位筛出来")
print(json.dumps(out["parse_big_main"], ensure_ascii=False), flush=True)
if app.plan:
    n = len(popups)
    reset()
    app.do_run()
    pump(lambda: not app.busy and len(popups) > n, 60)
    out["run_big_main"] = {"popup": popups[n:][:1], "max_ui_stall_s": round(hb["max"], 3), "s": round(time.time() - T0[0], 2)}
    print(json.dumps(out["run_big_main"], ensure_ascii=False), flush=True)

# 3) 小名单 + 指令里点名大表 → 大表作参考表，流式匹配
app.set_workdir(str(wd), select=SMALL)
app.on_file()
pump(lambda: "…" not in app.w_prof.get("1.0", "end"))
out["parse_lookup"] = parse(f"从{Path(BIG).stem}里按统一社会信用代码匹配单位名称和状态，匹配不到的写未匹配")
print(json.dumps(out["parse_lookup"], ensure_ascii=False), flush=True)
if app.plan and app.plan.get("steps"):
    n = len(popups)
    reset()
    app.do_run()
    pump(lambda: not app.busy, 600)
    out["run_lookup"] = {"s": round(time.time() - T0[0], 1), "max_ui_stall_s": round(hb["max"], 3), "statuses": statuses,
                         "result": app.v_out.get()[:160], "popups": popups[n:],
                         "log_tail": app.w_log.get("end-8l", "end").strip()[-600:]}
    print(json.dumps(out["run_lookup"], ensure_ascii=False), flush=True)
(wd / "_gui_bigtest.json").write_text(json.dumps(out, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
print("popups:", popups, flush=True)
print("handler timing [n, total_s, max_s]:", {k: [v[0], round(v[1], 3), round(v[2], 3)] for k, v in HT.items()}, flush=True)
app.on_close()
print("GUI_BIGTEST_DONE", flush=True)
