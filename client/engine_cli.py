"""无界面启动本地引擎（自动化测试 / 排障用）：python -m client.engine_cli [auto|cpu] [--hold]"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

from client.engine import Engine, app_dir, load_config, query_gpu


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("-") else "auto"
    cfg = load_config(app_dir() / "config.json")
    print("gpu", query_gpu(), flush=True)
    eng = Engine(cfg, Path(app_dir()) / "logs", on_change=lambda st: print("state", json.dumps(st.__dict__, ensure_ascii=False), flush=True))
    ok = eng.start(mode)
    print("READY" if ok else "FAILED", flush=True)
    time.sleep(8)  # 等热身测速
    print("final", json.dumps(eng.state.__dict__, ensure_ascii=False), flush=True)
    if "--hold" in sys.argv and ok:
        try:
            while eng.proc and eng.proc.poll() is None:
                time.sleep(2)
        finally:
            eng.stop()
    else:
        eng.stop()


if __name__ == "__main__":
    main()
