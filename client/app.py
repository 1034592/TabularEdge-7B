"""TabularEdge 桌面客户端（Windows，离线）。

  选择工作目录 → 选表 → 输入自然语言（中/英，回车提交，Shift+回车换行）
  → 本地 7B 模型流式编译：先出“我理解的是：……”，再出 Action Plan 预览（歧义给选项）
  → 一键执行：删除整行 / 去重 / 覆盖原文件 一律弹窗确认；写原文件前自动备份
  → 结果：生成文件列表 + 打开文件 / 打开所在文件夹；可在结果上继续下一条指令
"""
from __future__ import annotations

import copy
import json
import os
import queue
import re
import sys
import threading
import time
import traceback
import webbrowser
from pathlib import Path

if getattr(sys, "frozen", False):
    BASE = Path(sys.executable).resolve().parent
else:
    BASE = Path(__file__).resolve().parent.parent
    sys.path.insert(0, str(BASE))

DATA_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "TabularEdge"
DATA_DIR.mkdir(parents=True, exist_ok=True)
os.environ.setdefault("TABULAREDGE_HARDCASES", str(DATA_DIR / "hard_cases.jsonl"))

import tkinter as tk  # noqa: E402
from tkinter import filedialog, messagebox, ttk  # noqa: E402

from client import describe  # noqa: E402
from client.core import RefExecutor, mark_big_lookups, prompt_context, quick_profile, tidy_save_paths  # noqa: E402
from client.engine import Engine, load_config, save_config  # noqa: E402
from tabularedge import llm_parser  # noqa: E402
from tabularedge.compat import detect_excel  # noqa: E402
from tabularedge.executor import Workspace  # noqa: E402
from tabularedge.protocol import validate  # noqa: E402

CONFIG_PATH = BASE / "config.json"

T = {
    "zh": {
        "title": "TabularEdge · 本地表格智能助手", "lang": "语言", "workdir": "工作目录", "browse": "选择…", "open_dir": "打开",
        "files": "表格文件", "sheet": "工作表", "refresh": "刷新", "columns": "列", "rows": "行",
        "input": "用自然语言描述你要做的事（Enter 提交，Shift+Enter 换行）", "parse": "解析指令", "cancel": "停止",
        "understood": "我理解的是：", "plan": "计划预览", "raw": "模型实时输出", "log": "执行日志",
        "run": "一键执行", "open_out": "打开结果", "open_folder": "打开所在文件夹", "scenario": "示例场景",
        "next_step": "下一步指令", "engine": "引擎", "eng_starting": "正在启动本地模型…", "eng_ready": "就绪",
        "eng_failed": "启动失败", "eng_stopped": "未启动", "mode_auto": "自动（GPU+CPU 混合）", "mode_cpu": "纯 CPU",
        "restart": "重启引擎", "amb": "需要你确认的地方", "amb_default": "按计划：", "amb_alt": "改为：",
        "risky_title": "高危操作确认", "risky_body": "以下操作会删除数据或覆盖原文件：\n\n{0}\n\n结果默认另存为新文件；若覆盖原文件，执行前会自动备份到 _backup 文件夹。\n确定执行吗？",
        "no_file": "请先在左侧选择一个表格文件", "no_engine": "本地模型还没就绪，请稍候", "no_plan": "还没有可执行的计划",
        "done": "完成，用时 {0:.1f} 秒", "backup_note": "原表已自动备份：{0}", "untouched": "原表未改动，结果另存为新文件",
        "outputs": "生成 / 修改的文件", "fixes": "已自动修正模型输出 {0} 处（详见“模型实时输出”）",
        "dropped": "以下外观步骤无法修正，已跳过（数据操作不受影响）：{0}",
        "failed": "这条指令没能得到可执行的计划",
        "st_step": "正在执行步骤 {0}：{1}…", "failed_body": "模型给出的计划里有以下问题，自动修正后仍无法执行：\n\n{0}\n\n建议：按提示改一下说法，或把长指令拆成两三句分步执行。",
        "repair": "第 {0} 次自我修正：{1}", "refs": "已自动载入参考表：{0}", "steps": "步骤", "risk": "风险",
        "thinking": "正在解析…", "first": "首字 {0:.2f}s · 摘要 {1:.2f}s · 总计 {2:.1f}s",
        "declined": "你取消了：{0}", "err": "出错了", "pick_dir": "选择 Excel 所在文件夹", "speed": "{0} tok/s",
        "chain": "已切换到结果文件，可直接输入下一条指令", "help": "使用指南",
        "ask_only": "指令有歧义，请按上面的问题补充说明后重新提交",
        "st_main": "正在分析表格结构…", "st_ref": "正在分析参考表结构：{0}…", "st_prof": "正在读取表结构…",
        "st_load": "正在载入表格…", "st_copy": "正在带入参考表：{0}…", "st_stream": "正在大表《{0}》中流式匹配…",
        "st_idx": "大表匹配 1/3：建立键索引…", "st_scan": "大表匹配 2/3：逐行扫描…", "st_res": "大表匹配 3/3：取回结果…",
        "big_note": "⚠ 大表（{0:,} 行）：可作为参考表被匹配；直接逐格编辑请先筛出需要的部分",
    },
    "en": {
        "title": "TabularEdge · Local Spreadsheet Assistant", "lang": "Language", "workdir": "Folder", "browse": "Browse…",
        "open_dir": "Open", "files": "Workbooks", "sheet": "Sheet", "refresh": "Refresh", "columns": "columns", "rows": "rows",
        "input": "Describe what you want in plain words (Enter = submit, Shift+Enter = new line)", "parse": "Parse",
        "cancel": "Stop", "understood": "My understanding: ", "plan": "Plan preview", "raw": "Live model output", "log": "Run log",
        "run": "Run", "open_out": "Open result", "open_folder": "Open folder", "scenario": "Sample scenario",
        "next_step": "Next instruction", "engine": "Engine", "eng_starting": "Starting local model…", "eng_ready": "Ready",
        "eng_failed": "Failed", "eng_stopped": "Stopped", "mode_auto": "Auto (GPU+CPU split)", "mode_cpu": "CPU only",
        "restart": "Restart engine", "amb": "Please confirm", "amb_default": "As planned: ", "amb_alt": "Instead: ",
        "risky_title": "Confirm risky operation", "risky_body": "This step deletes data or overwrites the original file:\n\n{0}\n\nResults are saved as a new file by default; the original is backed up to _backup before any overwrite.\nProceed?",
        "no_file": "Pick a workbook on the left first", "no_engine": "Local model is not ready yet", "no_plan": "No plan to run yet",
        "done": "Done in {0:.1f}s", "backup_note": "Original backed up: {0}", "untouched": "Original untouched; result saved as a new file",
        "outputs": "Files created / changed", "fixes": "Auto-fixed {0} issue(s) in the model output (see Live model output)",
        "dropped": "These formatting steps could not be fixed and were skipped (data steps unaffected): {0}",
        "failed": "No runnable plan for this instruction",
        "st_step": "Running step {0}: {1}…", "failed_body": "The model's plan still has these problems after automatic repair:\n\n{0}\n\nTry rephrasing as suggested, or split a long instruction into two or three shorter ones.",
        "repair": "Self-repair #{0}: {1}", "refs": "Reference tables loaded: {0}", "steps": "Step", "risk": "Risk",
        "thinking": "Parsing…", "first": "first token {0:.2f}s · summary {1:.2f}s · total {2:.1f}s",
        "declined": "You declined: {0}", "err": "Error", "pick_dir": "Choose the folder with your Excel files", "speed": "{0} tok/s",
        "chain": "Switched to the result file; type the next instruction", "help": "Guide",
        "ask_only": "The instruction is ambiguous; please answer the question above and resubmit",
        "st_main": "Analysing table structure…", "st_ref": "Analysing reference table: {0}…", "st_prof": "Reading table structure…",
        "st_load": "Loading workbook…", "st_copy": "Bringing in reference table: {0}…", "st_stream": "Streaming lookup in large table {0}…",
        "st_idx": "Large lookup 1/3: indexing keys…", "st_scan": "Large lookup 2/3: scanning rows…", "st_res": "Large lookup 3/3: collecting results…",
        "big_note": "⚠ Large table ({0:,} rows): use it as a lookup reference; filter it before editing cell by cell",
    },
}

FONT = ("Microsoft YaHei UI", 10)
ACCENT = "#1F6FEB"


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.cfg = load_config(CONFIG_PATH)
        if "--cpu" in sys.argv:  # 纯CPU模式启动.bat
            self.cfg["mode"] = "cpu"
        self.lang = self.cfg.get("lang", "zh")
        self.q: queue.Queue = queue.Queue()
        self.plan = None
        self.plan_meta = {}
        self.cur_file = None      # 相对工作目录
        self.cur_sheet = None
        self.refs = []
        self.outputs = []
        self.stop_evt = threading.Event()
        self.busy = False
        self.amb_vars = []
        self.scenarios = self._load_scenarios()
        self.scn = None
        self.scn_step = 0
        xv = self.cfg.get("excel", "auto")
        self.excel = detect_excel() if xv == "auto" else str(xv)  # "2019" → 公式改写成 Excel 2019 可用的形式
        self.engine = Engine(self.cfg, DATA_DIR / "logs", on_change=lambda st: self.q.put(("engine", copy.copy(st))))
        llm_parser.ENDPOINT = self.engine.url
        self._build()
        self._apply_lang()
        wd = self.cfg.get("workdir") or (str(BASE / "test_cases") if (BASE / "test_cases").exists() else "")
        if wd:
            self.set_workdir(wd)
        self.root.after(50, self._pump)
        self.start_engine(self.cfg.get("mode", "auto"))
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def t(self, k):
        return T[self.lang][k]

    # ------------------------------------------------------------------ UI
    def _build(self):
        r = self.root
        r.geometry("1280x820")
        r.minsize(1000, 680)
        st = ttk.Style()
        try:
            st.theme_use("vista" if os.name == "nt" else "clam")
        except tk.TclError:
            pass
        st.configure(".", font=FONT)
        st.configure("Title.TLabel", font=("Microsoft YaHei UI", 15, "bold"))
        st.configure("Sum.TLabel", font=("Microsoft YaHei UI", 12, "bold"), foreground="#0B4FA3")
        st.configure("Muted.TLabel", foreground="#666666")
        st.configure("Accent.TButton", font=("Microsoft YaHei UI", 11, "bold"))
        st.configure("Treeview", rowheight=26)

        top = ttk.Frame(r, padding=(12, 8))
        top.pack(fill="x")
        self.w_title = ttk.Label(top, style="Title.TLabel")
        self.w_title.pack(side="left")
        right = ttk.Frame(top)
        right.pack(side="right")
        self.w_eng_dot = tk.Label(right, text="●", fg="#999999", font=("Segoe UI", 12))
        self.w_eng_dot.pack(side="left")
        self.w_eng = ttk.Label(right, width=46)
        self.w_eng.pack(side="left", padx=(2, 10))
        self.v_mode = tk.StringVar(value=self.cfg.get("mode", "auto"))
        self.w_mode = ttk.Combobox(right, state="readonly", width=18)
        self.w_mode.pack(side="left")
        self.w_mode.bind("<<ComboboxSelected>>", lambda e: self.start_engine(["auto", "cpu"][self.w_mode.current()]))
        self.w_restart = ttk.Button(right, command=lambda: self.start_engine(self.v_mode.get()))
        self.w_restart.pack(side="left", padx=6)
        self.w_lang_l = ttk.Label(right)
        self.w_lang_l.pack(side="left", padx=(12, 2))
        self.w_lang = ttk.Combobox(right, state="readonly", width=8, values=["中文", "English"])
        self.w_lang.current(0 if self.lang == "zh" else 1)
        self.w_lang.pack(side="left")
        self.w_lang.bind("<<ComboboxSelected>>", lambda e: self.set_lang("zh" if self.w_lang.current() == 0 else "en"))
        self.w_help = ttk.Button(right, command=self.open_help)
        self.w_help.pack(side="left", padx=6)

        bar = ttk.Frame(r, padding=(12, 0, 12, 6))
        bar.pack(fill="x")
        self.w_wd_l = ttk.Label(bar)
        self.w_wd_l.pack(side="left")
        self.v_wd = tk.StringVar()
        e = ttk.Entry(bar, textvariable=self.v_wd)
        e.pack(side="left", fill="x", expand=True, padx=6)
        e.bind("<Return>", lambda ev: self.set_workdir(self.v_wd.get()))
        self.w_browse = ttk.Button(bar, command=self.pick_dir)
        self.w_browse.pack(side="left")
        self.w_open_dir = ttk.Button(bar, command=lambda: self._open(self.v_wd.get()))
        self.w_open_dir.pack(side="left", padx=4)
        self.w_scn_l = ttk.Label(bar)
        self.w_scn_l.pack(side="left", padx=(16, 4))
        self.w_scn = ttk.Combobox(bar, state="readonly", width=34,
                                  values=[f"{s['id']}. {s['title']}" for s in self.scenarios])
        self.w_scn.pack(side="left")
        self.w_scn.bind("<<ComboboxSelected>>", lambda e: self.pick_scenario(self.w_scn.current()))
        self.w_next = ttk.Button(bar, command=self.next_scenario_step, state="disabled")
        self.w_next.pack(side="left", padx=4)

        pw = ttk.PanedWindow(r, orient="horizontal")
        pw.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        left = ttk.Frame(pw, width=330)
        pw.add(left, weight=1)
        mid = ttk.Frame(pw)
        pw.add(mid, weight=4)

        lh = ttk.Frame(left)
        lh.pack(fill="x")
        self.w_files_l = ttk.Label(lh, font=("Microsoft YaHei UI", 10, "bold"))
        self.w_files_l.pack(side="left")
        self.w_refresh = ttk.Button(lh, command=lambda: self.set_workdir(self.v_wd.get()), width=6)
        self.w_refresh.pack(side="right")
        self.w_files = ttk.Treeview(left, show="tree", selectmode="browse", height=14)
        self.w_files.pack(fill="both", expand=True, pady=4)
        self.w_files.bind("<<TreeviewSelect>>", self.on_file)
        sh = ttk.Frame(left)
        sh.pack(fill="x")
        self.w_sheet_l = ttk.Label(sh)
        self.w_sheet_l.pack(side="left")
        self.w_sheet = ttk.Combobox(sh, state="readonly")
        self.w_sheet.pack(side="left", fill="x", expand=True, padx=4)
        self.w_sheet.bind("<<ComboboxSelected>>", lambda e: self.on_sheet())
        self.w_prof = tk.Text(left, height=12, wrap="word", font=("Consolas", 9), relief="flat", background="#F6F8FA")
        self.w_prof.pack(fill="both", expand=False, pady=4)

        inp = ttk.Frame(mid)
        inp.pack(fill="x")
        self.w_inp_l = ttk.Label(inp, style="Muted.TLabel")
        self.w_inp_l.pack(anchor="w")
        row = ttk.Frame(inp)
        row.pack(fill="x", pady=2)
        self.w_inp = tk.Text(row, height=4, wrap="word", font=("Microsoft YaHei UI", 11), relief="solid", borderwidth=1)
        self.w_inp.pack(side="left", fill="x", expand=True)
        self.w_inp.bind("<Return>", self._on_enter)
        self.w_inp.bind("<KeyRelease>", self._on_typing)  # 打字停顿 0.5 秒 → 后台预填（含指令里点名的参考表）
        btns = ttk.Frame(row)
        btns.pack(side="left", padx=6, fill="y")
        self.w_parse = ttk.Button(btns, command=self.do_parse, style="Accent.TButton")
        self.w_parse.pack(fill="x")
        self.w_cancel = ttk.Button(btns, command=lambda: self.stop_evt.set(), state="disabled")
        self.w_cancel.pack(fill="x", pady=4)

        self.v_sum = tk.StringVar()
        ttk.Label(mid, textvariable=self.v_sum, style="Sum.TLabel", wraplength=880, justify="left").pack(anchor="w", pady=(8, 2))
        self.v_info = tk.StringVar()
        ttk.Label(mid, textvariable=self.v_info, style="Muted.TLabel", wraplength=880, justify="left").pack(anchor="w")

        self.w_amb = ttk.LabelFrame(mid, padding=6)
        self.nb = ttk.Notebook(mid)
        self.nb.pack(fill="both", expand=True, pady=6)
        f1 = ttk.Frame(self.nb)
        self.w_plan = ttk.Treeview(f1, columns=("risk", "desc"), show="headings")
        self.w_plan.column("risk", width=60, anchor="center", stretch=False)
        self.w_plan.column("desc", width=800)
        self.w_plan.tag_configure("risky", foreground="#C00000")
        self.w_plan.pack(fill="both", expand=True)
        f2 = ttk.Frame(self.nb)
        # wrap="char"：模型输出是不带空格的长 JSON 行，wrap="word" 每插入一个 token 都要重排整行（越来越慢，曾把界面卡死几十秒）
        self.w_raw = tk.Text(f2, wrap="char", font=("Consolas", 9), background="#0D1117", foreground="#C9D1D9",
                             insertbackground="white")
        self.w_raw.pack(fill="both", expand=True)
        f3 = ttk.Frame(self.nb)
        self.w_log = tk.Text(f3, wrap="word", font=("Consolas", 9))
        self.w_log.pack(fill="both", expand=True)
        self.nb.add(f1)
        self.nb.add(f2)
        self.nb.add(f3)

        bot = ttk.Frame(mid)
        bot.pack(fill="x")
        self.w_run = ttk.Button(bot, command=self.do_run, style="Accent.TButton", state="disabled")
        self.w_run.pack(side="left")
        self.w_open_out = ttk.Button(bot, command=self.open_output, state="disabled")
        self.w_open_out.pack(side="left", padx=6)
        self.w_open_folder = ttk.Button(bot, command=self.open_output_folder, state="disabled")
        self.w_open_folder.pack(side="left")
        self.v_out = tk.StringVar()
        ttk.Label(bot, textvariable=self.v_out, style="Muted.TLabel", wraplength=640, justify="left").pack(side="left", padx=10)
        self.w_outs = ttk.Treeview(mid, show="tree", height=4)
        self.w_outs.pack(fill="x", pady=(6, 0))
        self.w_outs.bind("<Double-1>", lambda e: self.open_output(self.w_outs.focus()))

    def _apply_lang(self):
        t = self.t
        self.root.title(t("title"))
        self.w_title.config(text="TabularEdge")
        self.w_lang_l.config(text=t("lang"))
        self.w_wd_l.config(text=t("workdir"))
        self.w_browse.config(text=t("browse"))
        self.w_open_dir.config(text=t("open_dir"))
        self.w_files_l.config(text=t("files"))
        self.w_refresh.config(text=t("refresh"))
        self.w_sheet_l.config(text=t("sheet"))
        self.w_inp_l.config(text=t("input"))
        self.w_parse.config(text=t("parse"))
        self.w_cancel.config(text=t("cancel"))
        self.w_run.config(text="▶ " + t("run"))
        self.w_open_out.config(text=t("open_out"))
        self.w_open_folder.config(text=t("open_folder"))
        self.w_scn_l.config(text=t("scenario"))
        self.w_next.config(text=t("next_step"))
        self.w_restart.config(text=t("restart"))
        self.w_help.config(text=t("help"))
        self.w_mode.config(values=[t("mode_auto"), t("mode_cpu")])
        self.w_mode.current(1 if self.v_mode.get() == "cpu" else 0)
        self.w_amb.config(text=t("amb"))
        self.w_plan.heading("risk", text=t("risk"))
        self.w_plan.heading("desc", text=t("steps"))
        for i, k in enumerate(("plan", "raw", "log")):
            self.nb.tab(i, text=t(k))
        self._show_engine(self.engine.state)
        if self.plan:
            self.render_plan()

    def set_lang(self, lang):
        self.lang = lang
        self.cfg["lang"] = lang
        save_config(CONFIG_PATH, self.cfg)
        self._apply_lang()

    # ------------------------------------------------------------------ engine
    def start_engine(self, mode):
        self.v_mode.set(mode)
        self.cfg["mode"] = mode
        save_config(CONFIG_PATH, self.cfg)
        threading.Thread(target=self.engine.start, args=(mode,), daemon=True).start()

    def _show_engine(self, st):
        color = {"ready": "#1A7F37", "starting": "#BF8700", "failed": "#C00000"}.get(st.status, "#999999")
        self.w_eng_dot.config(fg=color)
        txt = {"ready": self.t("eng_ready"), "starting": self.t("eng_starting"), "failed": self.t("eng_failed"),
               "stopped": self.t("eng_stopped")}[st.status]
        if st.status == "ready":
            where = "CPU" if st.backend == "cpu" else f"{st.backend.upper()} {st.ngl}/{st.n_layers}" + (" + CPU" if st.ngl <= st.n_layers else "")
            txt += f" · {where}" + (f" · {self.t('speed').format(st.tok_s)}" if st.tok_s else "")
        elif st.detail:
            txt += f" · {st.detail}"
        self.w_eng.config(text=txt[:70])

    # ------------------------------------------------------------------ files
    def pick_dir(self):
        d = filedialog.askdirectory(title=self.t("pick_dir"), initialdir=self.v_wd.get() or str(Path.home()))
        if d:
            self.set_workdir(d)

    def set_workdir(self, d, select=None):
        p = Path(d)
        if not p.is_dir():
            return
        self.v_wd.set(str(p))
        self.cfg["workdir"] = str(p)
        save_config(CONFIG_PATH, self.cfg)
        self.w_files.delete(*self.w_files.get_children())
        files = sorted(x for x in p.rglob("*.xlsx") if "_backup" not in x.parts and "_original" not in x.parts and not x.name.startswith("~$"))
        for f in files:
            rel = str(f.relative_to(p))
            self.w_files.insert("", "end", iid=rel, text=("📄 " if os.name == "nt" else "") + rel)
        if select and select in self.w_files.get_children():
            self.w_files.selection_set(select)
            self.w_files.see(select)

    def on_file(self, _=None):
        """选中文件：表名和表结构都在后台线程读取（150MB 的大表也不会卡住界面）。"""
        sel = self.w_files.selection()
        if not sel:
            return
        self.cur_file = sel[0]
        self._load_profile(self.cur_file, self.cur_sheet, set_sheets=True)

    def on_sheet(self):
        self.cur_sheet = self.w_sheet.get()
        self._load_profile(self.cur_file, self.cur_sheet)

    def _load_profile(self, file, sheet, set_sheets=False):
        self._prof_token = token = object()
        self.w_prof.delete("1.0", "end")
        self.w_prof.insert("end", self.t("st_prof"))
        path = Path(self.v_wd.get()) / file

        def work():
            try:
                prof, names = quick_profile(path, sheet)
                self.q.put(("profile", token, file, prof, names, set_sheets, None))
            except Exception as e:  # noqa: BLE001
                self.q.put(("profile", token, file, None, [], set_sheets, f"{type(e).__name__}: {e}"))
        threading.Thread(target=work, daemon=True).start()

    def _on_profile(self, token, file, prof, names, set_sheets, err):
        if token is not getattr(self, "_prof_token", None) or file != self.cur_file:
            return  # 用户已经切到别的文件 / 表
        self.w_prof.delete("1.0", "end")
        if err:
            self.w_prof.insert("end", err)
            return
        if set_sheets:
            self.w_sheet.config(values=names)
        self.cur_sheet = prof.sheet
        self.w_sheet.set(prof.sheet)
        txt = f"{prof.sheet} · {prof.n_rows:,} {self.t('rows')} · {len(prof.headers)} {self.t('columns')}\n"
        if prof.n_rows * max(len(prof.headers), 1) > 3_000_000:
            txt += self.t("big_note").format(prof.n_rows) + "\n"
        txt += "\n".join(f"  {h} ({prof.types.get(h, '?')})" for h in prof.headers)
        others = [x for x in names if x != prof.sheet]
        if others:
            txt += "\n+ " + ", ".join(others)
        self.w_prof.insert("end", txt)
        self._prefill()

    def _prefill(self, _=None):
        """后台预填：把“当前表 +（指令里已经点名的）参考表”的画像先送进模型缓存，用户还在打字时就算好。
        回车后只剩指令本身那几十个 token 要处理，首字明显更快。只保留最新一次请求；失败无所谓。"""
        self._type_job = None
        if self.engine.state.status != "ready" or self.busy or not self.cur_file:
            return
        self._pf_req = (Path(self.v_wd.get()), self.cur_file, self.cur_sheet, self.w_inp.get("1.0", "end").strip(), self.engine.url)
        if not getattr(self, "_pf_running", False):
            self._pf_running = True
            threading.Thread(target=self._prefill_worker, daemon=True).start()

    def _prefill_worker(self):
        try:
            while True:
                req, self._pf_req = getattr(self, "_pf_req", None), None
                if req is None:
                    break
                wd, f, sh, text, url = req
                try:
                    prof, extra, _ = prompt_context(wd, f, sh, text)
                    key = (str(wd), f, prof.to_prompt(), extra)
                    if key != getattr(self, "_pf_key", None):
                        llm_parser.prefill(prof, f, extra, url)
                        self._pf_key = key
                except Exception:  # noqa: BLE001
                    pass
        finally:
            self._pf_running = False

    def _on_typing(self, _=None):
        if getattr(self, "_type_job", None):
            self.root.after_cancel(self._type_job)
        self._type_job = self.root.after(500, self._prefill)

    def _on_status(self, msg):
        self.v_info.set(msg)

    # ------------------------------------------------------------------ scenarios
    def _load_scenarios(self):
        for p in (BASE / "test_cases" / "scenarios.json",):
            try:
                return json.loads(p.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                pass
        return []

    def pick_scenario(self, i):
        self.scn = self.scenarios[i]
        self.scn_step = 0
        self.set_workdir(str(BASE / "test_cases"), select=self.scn["file"])
        self.cur_sheet = self.scn.get("sheet")
        self.on_file()
        self._fill_scn_step()

    def _fill_scn_step(self):
        steps = self.scn["steps"]
        s = steps[self.scn_step]
        self.w_inp.delete("1.0", "end")
        self.w_inp.insert("end", s["text"] if self.lang == "zh" or not s.get("text_en") else s["text_en"])
        self.w_next.config(state="normal" if self.scn_step + 1 < len(steps) else "disabled")
        self._on_typing()
        self.v_info.set(f"{self.scn['id']}. {self.scn['title']} — {self.scn_step + 1}/{len(steps)}")

    def next_scenario_step(self):
        if self.scn and self.scn_step + 1 < len(self.scn["steps"]):
            self.scn_step += 1
            self._fill_scn_step()

    # ------------------------------------------------------------------ parse
    def _on_enter(self, ev):
        if ev.state & 0x0001:  # Shift+Enter 换行
            return None
        self.do_parse()
        return "break"

    def do_parse(self):
        if self.busy:
            return
        text = self.w_inp.get("1.0", "end").strip()
        if not text:
            return
        if not self.cur_file:
            messagebox.showinfo("TabularEdge", self.t("no_file"))
            return
        if self.engine.state.status != "ready":
            messagebox.showinfo("TabularEdge", self.t("no_engine"))
            return
        self.plan = None
        self.refs = []
        self.busy = True
        self.stop_evt.clear()
        self.v_sum.set(self.t("understood") + "…")
        self.v_info.set(self.t("st_main"))
        self.w_raw.delete("1.0", "end")
        self._raw_col = 0
        self.w_plan.delete(*self.w_plan.get_children())
        self.w_amb.pack_forget()
        self.w_run.config(state="disabled")
        self.w_parse.config(state="disabled")
        self.w_cancel.config(state="normal")
        self.nb.select(1)
        args = (text, Path(self.v_wd.get()), self.cur_file, self.cur_sheet)
        threading.Thread(target=self._parse_worker, args=args, daemon=True).start()

    def _parse_worker(self, text, wd, file, sheet):
        """后台线程：表格画像（文件读取）+ 模型解析。界面线程只通过队列收消息，始终保持响应。"""
        t0 = time.time()
        try:
            def prog(m):
                self.q.put(("status", self.t("st_ref").format(m[4:]) if m.startswith("ref ") else self.t("st_main")))
            prof, extra, refs = prompt_context(wd, file, sheet, text, progress=prog)
            self.q.put(("ctx", refs, time.time() - t0))
            if self.stop_evt.is_set():
                raise InterruptedError
            plan, info = llm_parser.parse(
                text, prof, file, extra_profile=extra, url=self.engine.url, stop=self.stop_evt,
                on_summary=lambda s: self.q.put(("summary", s, time.time() - t0)),
                on_delta=lambda d: self.q.put(("delta", d)),
                on_repair=lambda n, errs: self.q.put(("repair", n, errs)))
            self.q.put(("plan", plan, info, time.time() - t0, prof))
        except InterruptedError:
            self.q.put(("cancelled",))
        except Exception as e:  # noqa: BLE001
            self.q.put(("error", f"{type(e).__name__}: {e}", traceback.format_exc()))

    # ------------------------------------------------------------------ plan view
    def render_plan(self, plan=None, bad=frozenset()):
        plan = plan if plan is not None else (self.plan or {})
        zh = self.lang == "zh"
        self.w_plan.delete(*self.w_plan.get_children())
        for i, st in enumerate(plan.get("steps") or []):
            sid = str(st.get("id") or f"#{i + 1}")
            try:
                risky = describe.is_risky(st, self.cur_file)
                text = describe.step(st, zh)
            except Exception:  # noqa: BLE001  草稿里格式不对的步骤：原样显示
                risky, text = False, f"{st.get('op')} {json.dumps(st.get('args'), ensure_ascii=False)[:160]}"
            mark = "❌" if (sid in bad or i in bad) else "⚠" if risky else ""
            iid = sid if not self.w_plan.exists(sid) else f"{sid}_{i}"
            self.w_plan.insert("", "end", iid=iid, values=(mark, f"{sid}  " + text), tags=("risky",) if mark else ())
        for w in self.w_amb.winfo_children():
            w.destroy()
        self.amb_vars = []
        ambs = plan.get("ambiguities") or []
        if ambs:
            self.w_amb.pack(fill="x", pady=4, before=self.nb)
            for a in ambs:
                ttk.Label(self.w_amb, text="❓ " + a.get("question", ""), wraplength=860, justify="left",
                          font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w")
                cur = next((s for s in plan["steps"] if s.get("id") == a.get("step_id")), None)
                v = tk.IntVar(value=0)
                self.amb_vars.append((a, v))
                if cur:
                    ttk.Radiobutton(self.w_amb, variable=v, value=0,
                                    text=self.t("amb_default") + describe.step(cur, zh)).pack(anchor="w", padx=16)
                for j, alt in enumerate(a.get("alternatives") or [], 1):
                    if isinstance(alt, dict) and "op" in alt:
                        ttk.Radiobutton(self.w_amb, variable=v, value=j,
                                        text=self.t("amb_alt") + describe.step(alt, zh)).pack(anchor="w", padx=16)
        else:
            self.w_amb.pack_forget()

    def _final_plan(self):
        plan = copy.deepcopy(self.plan)
        for a, v in self.amb_vars:
            j = v.get()
            if j:
                alt = copy.deepcopy(a["alternatives"][j - 1])
                for i, s in enumerate(plan["steps"]):
                    if s["id"] == alt.get("id", a.get("step_id")):
                        alt.setdefault("id", s["id"])
                        if "sheet" not in alt and s.get("sheet"):
                            alt["sheet"] = s["sheet"]
                        plan["steps"][i] = alt
        plan.pop("ambiguities", None)
        return plan

    # ------------------------------------------------------------------ run
    def do_run(self):
        if not self.plan or self.busy:
            return
        plan = self._final_plan()
        errs = validate(plan)
        if errs:
            messagebox.showerror(self.t("err"), "\n".join(errs[:5]))
            return
        self.busy = True
        self.w_run.config(state="disabled")
        self.nb.select(2)
        threading.Thread(target=self._run_worker, args=(plan, Path(self.v_wd.get())), daemon=True).start()

    def _confirm(self, step):
        """执行器在写盘 / 破坏性步骤前回调（工作线程）。高危步骤切回界面线程弹窗；普通写盘视为已在预览里确认。"""
        if not describe.is_risky(step, self.cur_file):
            return True
        ev, box = threading.Event(), {}
        self.q.put(("confirm", step, ev, box))
        ev.wait()
        return box.get("ok", False)

    def _run_worker(self, plan, wd: Path):  # 工作线程里不能碰 Tk 变量，目录由界面线程传入
        t0 = time.time()
        before = {p: p.stat().st_mtime for p in wd.rglob("*") if p.is_file()}
        try:
            def prog(m):
                k = m.split(" ", 1)
                msg = {"step": lambda: self.t("st_step").format(k[1].split(" ")[0], llm_parser._OP_ZH.get(k[1].split(" ")[-1], k[1].split(" ")[-1])),
                       "copy-ref": lambda: self.t("st_copy").format(k[1]), "stream-lookup": lambda: self.t("st_stream").format(k[1]),
                       "index-strings": lambda: self.t("st_idx"), "scan-rows": lambda: self.t("st_scan"),
                       "resolve": lambda: self.t("st_res")}.get(k[0])
                if msg:
                    self.q.put(("status", msg()))
            self.q.put(("status", self.t("st_load")))
            ex = RefExecutor(Workspace(wd), self._confirm, self.refs, compat=self.excel, progress=prog)
            log = ex.run(plan)
            after = [p for p in wd.rglob("*") if p.is_file() and before.get(p) != p.stat().st_mtime]
            self.q.put(("ran", log, after, time.time() - t0))
        except Exception as e:  # noqa: BLE001
            self.q.put(("error", f"{type(e).__name__}: {e}", traceback.format_exc()))

    # ------------------------------------------------------------------ results
    def _open(self, p):
        try:
            os.startfile(str(p))  # noqa: S606  (Windows)
        except AttributeError:
            webbrowser.open(Path(p).as_uri())
        except OSError as e:
            messagebox.showerror(self.t("err"), str(e))

    def open_output(self, which=None):
        if which:
            self._open(Path(self.v_wd.get()) / which)
        elif self.outputs:
            self._open(self.outputs[0])

    def open_output_folder(self):
        if self.outputs:
            p = self.outputs[0]
            if os.name == "nt":
                import subprocess
                subprocess.Popen(["explorer", "/select,", str(p)])
            else:
                self._open(p.parent)

    def open_help(self):
        for name in ("运行指南.html", "运行指南.md", "README.txt"):
            p = BASE / name
            if p.exists():
                self._open(p)
                return

    # ------------------------------------------------------------------ queue pump
    def _pump(self):
        """界面线程每 40ms 处理一次后台消息：每轮最多占用 30ms（之后让出给窗口重绘 / 鼠标键盘），
        连续的流式 token 合并成一次插入。"""
        deadline = time.perf_counter() + 0.03
        buf = []
        try:
            while time.perf_counter() < deadline:
                m = self.q.get_nowait()
                if m[0] == "delta":
                    buf.append(m[1])
                    continue
                if buf:
                    self._on_delta("".join(buf))
                    buf = []
                getattr(self, "_on_" + m[0])(*m[1:])
        except queue.Empty:
            pass
        if buf:
            self._on_delta("".join(buf))
        self.root.after(40, self._pump)

    def _on_engine(self, st):
        self._show_engine(st)
        if st.status == "ready":
            self.root.after(1500, self._prefill)  # 等引擎热身（_bench）结束再预填

    def _on_ctx(self, refs, dt_):
        self.refs = refs
        self.plan_meta["ctx_s"] = dt_
        self.v_info.set(self.t("thinking") + (("  " + self.t("refs").format("、".join(r[0] for r in refs))) if refs else ""))

    def _on_summary(self, s, dt_):
        self.v_sum.set(self.t("understood") + s)
        self.plan_meta["summary_s"] = dt_

    def _on_delta(self, d):
        # 每 160 字软换行：Tk Text 插入的开销随“当前逻辑行长度”增长，不断开的话长计划越到后面越卡
        out = []
        for ch in d:
            out.append(ch)
            self._raw_col = 0 if ch == "\n" else getattr(self, "_raw_col", 0) + 1
            if self._raw_col >= 160:
                out.append("\n")
                self._raw_col = 0
        self.w_raw.insert("end", "".join(out))
        self.w_raw.see("end")

    def _on_repair(self, n, errs):
        self.w_raw.insert("end", f"\n\n--- {self.t('repair').format(n, '; '.join(errs[:2]))} ---\n")

    def _on_plan(self, plan, info, total, prof):
        self.busy = False
        self.w_parse.config(state="normal")
        self.w_cancel.config(state="disabled")
        plan = plan or {"steps": []}
        failed = bool(info.get("failed"))
        self.plan = None if failed else mark_big_lookups(tidy_save_paths(plan, self.cur_file, Path(self.v_wd.get())), self.refs)
        self.v_sum.set(self.t("understood") + (plan.get("summary") or "—"))
        msg = self.t("first").format(info.get("first_token_s") or 0, info.get("summary_s") or 0, total)
        if info.get("fixes"):
            msg += "  ·  " + self.t("fixes").format(len(info["fixes"]))
            self.w_raw.insert("end", "\n\n--- 自动修正 / auto-fixes ---\n" + "\n".join(info["fixes"]) + "\n")
        if info.get("attempts", 1) > 1 and not failed:
            msg += f"  ·  {self.t('repair').format(info['attempts'] - 1, 'ok')}"
        if self.refs:
            msg += "  ·  " + self.t("refs").format("、".join(r[0] for r in self.refs))
        if info.get("dropped"):
            msg = "⚠ " + self.t("dropped").format("、".join(info["dropped"])) + "  ·  " + msg
            self.w_raw.insert("end", "\n--- 已跳过 / skipped ---\n" + "\n".join(info.get("dropped_explain", [])) + "\n")
        self.v_info.set(msg)
        self.render_plan(plan, bad=self._bad_steps(info) if failed else set())
        self.nb.select(0)
        if failed:  # 不再退回规则解析：给出逐条、具体的说明（一次），计划草稿仅供查看，不能执行
            lines = info.get("explain") or info.get("errors") or ["?"]
            self.v_info.set("❌ " + self.t("failed") + "：" + lines[0][:90] + "  ·  " + msg)
            self.w_raw.insert("end", "\n--- 无法执行 / not runnable ---\n" + "\n".join(lines) + "\n")
            self.w_run.config(state="disabled")
            messagebox.showwarning(self.t("failed"), self.t("failed_body").format("\n".join(f"• {x}" for x in lines[:8])))
            return
        if info.get("ask_only") or not plan.get("steps"):  # 只追问没有计划：请用户补充说明后重新提交
            self.v_info.set("❓ " + self.t("ask_only") + "  ·  " + msg)
            self.w_run.config(state="disabled")
            return
        self.w_run.config(state="normal")

    @staticmethod
    def _bad_steps(info):
        out = set()
        for e in info.get("errors") or []:
            m = re.match(r"semantic: (s\d+)\b", e)
            if m:
                out.add(m.group(1))
            m = re.match(r"schema: steps/(\d+)", e)
            if m:
                out.add(int(m.group(1)))
        return out

    def _on_cancelled(self):
        self.busy = False
        self.w_parse.config(state="normal")
        self.w_cancel.config(state="disabled")
        self.v_info.set("—")

    def _on_error(self, msg, tb):
        self.busy = False
        self.w_parse.config(state="normal")
        self.w_cancel.config(state="disabled")
        self.w_run.config(state="normal" if self.plan else "disabled")
        self.w_log.insert("end", tb + "\n")
        messagebox.showerror(self.t("err"), msg)

    def _on_confirm(self, step, ev, box):
        box["ok"] = messagebox.askyesno(self.t("risky_title"),
                                        self.t("risky_body").format(describe.step(step, self.lang == "zh")), icon="warning")
        if not box["ok"]:
            self.w_log.insert("end", self.t("declined").format(describe.step(step, self.lang == "zh")) + "\n")
        ev.set()

    def _on_ran(self, log, changed, dt_):
        self.busy = False
        self.v_info.set(self.t("done").format(dt_))
        self.w_run.config(state="normal")
        wd = Path(self.v_wd.get())
        self.w_log.insert("end", "\n".join(log) + "\n" + "-" * 60 + "\n")
        self.w_log.see("end")
        backups = [ln.split("-> ")[-1] for ln in log if ln.startswith("auto-backup")]
        outs = [p for p in changed if "_backup" not in p.parts]
        outs.sort(key=lambda p: (p.suffix != ".xlsx", -p.stat().st_mtime))
        self.outputs = outs
        self.w_outs.delete(*self.w_outs.get_children())
        for p in outs:
            rel = str(p.relative_to(wd))
            self.w_outs.insert("", "end", iid=rel, text="📄 " + rel)
        note = self.t("backup_note").format(", ".join(backups)) if backups else self.t("untouched")
        self.v_out.set(f"{self.t('done').format(dt_)} · {note} · {self.t('outputs')}: {len(outs)}")
        self.w_open_out.config(state="normal" if outs else "disabled")
        self.w_open_folder.config(state="normal" if outs else "disabled")
        # 在结果上继续：自动选中生成的主文件（与源文件同表名的那个）
        main = next((p for p in outs if p.suffix == ".xlsx" and p.parent == wd), None)
        if main is not None:
            rel = str(main.relative_to(wd))
            self.set_workdir(str(wd), select=rel)
            self.cur_file = rel
            self.on_file()
            self.v_out.set(self.v_out.get() + " · " + self.t("chain"))
        else:
            self.set_workdir(str(wd), select=self.cur_file)

    def on_close(self):
        try:
            self.engine.stop()
        finally:
            self.root.destroy()


def _ver(cp) -> str:
    txt = ((cp.stdout or "") + (cp.stderr or "")).strip()
    v = next((ln for ln in txt.splitlines() if "version" in ln.lower()), txt.splitlines()[0] if txt else "")
    if cp.returncode != 0:
        raise RuntimeError(f"exit {cp.returncode}: {v[:120]}")
    return v[:80]


def selftest() -> int:
    """打包后自检（不开窗口）：协议/Schema、执行器 + 公式求值、Excel 2019 改写、三套引擎可执行、模型文件。结果写到 selftest.txt。"""
    import shutil
    import subprocess
    import tempfile
    from client.engine import NO_WINDOW, app_dir
    lines, ok = [], True

    def chk(name, fn):
        nonlocal ok
        try:
            lines.append(f"PASS {name}: {fn()}")
        except Exception as e:  # noqa: BLE001
            ok = False
            lines.append(f"FAIL {name}: {type(e).__name__}: {e}")
    cfg = load_config(CONFIG_PATH)
    chk("schema", lambda: f"loaded; empty plan rejected with {len(validate({'version': '0.2', 'inputs': [], 'steps': []}))} errors")
    chk("model file", lambda: f"{Engine(cfg, Path(tempfile.gettempdir())).model_path().stat().st_size / 2**30:.2f} GiB")

    def run_case():
        from openpyxl import load_workbook
        d = Path(tempfile.mkdtemp())
        shutil.copy2(BASE / "test_cases" / "03_季度绩效.xlsx", d / "a.xlsx")
        plan = {"version": "0.2", "lang": "zh", "inputs": [{"alias": "src", "path": "a.xlsx", "sheet": "绩效明细"}], "steps": [
            {"id": "s1", "op": "add_column", "sheet": "绩效明细", "args": {"name": "达标率", "formula": "=[完成额]/[目标额]"}},
            {"id": "s2", "op": "add_column", "sheet": "绩效明细", "args": {"name": "x", "formula": "=XLOOKUP([部门],[部门],[季度],\"无\")"}},
            {"id": "s3", "op": "group_aggregate", "sheet": "绩效明细", "args": {"by": ["部门"], "aggs": [{"col": "达标率", "fn": "avg"}], "output_sheet": "汇总"}},
            {"id": "s4", "op": "save", "args": {"path": "b.xlsx"}}]}
        log = RefExecutor(Workspace(d), lambda s: True, [], compat="2019").run(plan)
        wb = load_workbook(d / "b.xlsx")
        return f"{len(log)} steps; G2={wb['绩效明细']['H2'].value[:40]}; groups={wb['汇总'].max_row - 1}"
    chk("executor+2019", run_case)

    def stream_case():
        from client.core import stream_profile
        from tabularedge.xlsx_stream import XlsxStream
        t0 = time.time()
        with XlsxStream(BASE / "test_cases" / "03_季度绩效.xlsx") as xs:
            p = stream_profile(xs, xs.sheet_names[0])
        return f"{p.sheet}: {len(p.headers)} cols, {p.n_rows} rows in {time.time() - t0:.3f}s"
    chk("stream reader", stream_case)
    for b in ("cuda", "vulkan", "cpu"):
        exe = app_dir() / "engine" / b / "llama-server.exe"
        chk(f"engine {b}", lambda exe=exe: _ver(subprocess.run([str(exe), "--version"], capture_output=True, text=True, timeout=60,
                                                          cwd=str(exe.parent), creationflags=NO_WINDOW)))
    chk("excel detect", lambda: __import__("tabularedge.compat", fromlist=["x"]).detect_excel())
    out = DATA_DIR / "selftest.txt"
    out.write_text("\n".join(lines) + f"\nRESULT {'OK' if ok else 'FAIL'}\n", encoding="utf-8")
    return 0 if ok else 1


def main():
    if "--selftest" in sys.argv:
        sys.exit(selftest())
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # noqa: BLE001
            pass
    root = tk.Tk()
    try:
        root.iconbitmap(default=str(Path(getattr(sys, "_MEIPASS", BASE / "client")) / "tabularedge.ico"))
    except tk.TclError:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
