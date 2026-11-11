"""界面冒烟测试（Windows 宿主机）：真实创建 TabularEdge 窗口，自动走场景 1 与场景 10 的全部步骤。
弹窗被替换为“自动点是”并记录内容；验证 解析 → 计划预览 → 歧义选项 → 高危确认 → 执行 → 结果列表 → 链式下一步 全链路。

    python -m client.gui_smoketest <工作副本目录>
"""
from __future__ import annotations

import json
import shutil
import sys
import time
from pathlib import Path

import client.app as A

work = Path(sys.argv[1])
src = Path(__file__).resolve().parent.parent / "test_cases"
shutil.rmtree(work, ignore_errors=True)
(work / "test_cases").mkdir(parents=True)
for f in src.glob("*.xlsx"):
    shutil.copy2(f, work / "test_cases" / f.name)
shutil.copy2(src / "scenarios.json", work / "test_cases" / "scenarios.json")
A.BASE = work
A.CONFIG_PATH = work / "config.json"

popups = []
A.messagebox.askyesno = lambda title, body, **kw: popups.append(("askyesno", title, body)) or True
A.messagebox.showinfo = lambda *a, **kw: popups.append(("info",) + a)
A.messagebox.showerror = lambda *a, **kw: popups.append(("error",) + a)

root = A.tk.Tk()
app = A.App(root)


def pump(cond, timeout=180):
    t0 = time.time()
    while not cond():
        root.update()
        time.sleep(0.03)
        if time.time() - t0 > timeout:
            raise TimeoutError("ui wait")


pump(lambda: app.engine.state.status in ("ready", "failed"), 240)
print("engine:", app.w_eng.cget("text"), flush=True)
report = []
for sid in (1, 10):
    i = next(k for k, s in enumerate(app.scenarios) if s["id"] == sid)
    app.pick_scenario(i)
    root.update()
    for k in range(len(app.scn["steps"])):
        text = app.w_inp.get("1.0", "end").strip()
        t0 = time.time()
        app.do_parse()
        pump(lambda: not app.busy)
        summary = app.v_sum.get()
        plan_rows = [app.w_plan.item(x, "values") for x in app.w_plan.get_children()]
        amb = [w.cget("text") for w in app.w_amb.winfo_children()] if app.w_amb.winfo_ismapped() else []
        n_pop = len(popups)
        app.do_run()
        pump(lambda: not app.busy)
        report.append({"scenario": sid, "step": k + 1, "file": app.cur_file, "text": text, "summary": summary,
                       "info": app.v_info.get(), "plan": plan_rows, "ambiguity_ui": amb, "popups": popups[n_pop:],
                       "result": app.v_out.get(), "outputs": [app.w_outs.item(x, "text") for x in app.w_outs.get_children()],
                       "sec": round(time.time() - t0, 1)})
        print(json.dumps(report[-1], ensure_ascii=False), flush=True)
        app.next_scenario_step()
        root.update()
app.set_lang("en")
root.update()
print("lang-en title:", root.title(), "| run button:", app.w_run.cget("text"), flush=True)
app.set_lang("zh")
app.engine.proc = None  # 复用的外部引擎不要关
root.destroy()
print("GUI_SMOKE_OK", flush=True)
