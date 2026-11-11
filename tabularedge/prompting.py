"""训练与推理共用的提示格式（两边必须逐字一致）。

模型输出只含 summary / lang / steps / ambiguities：
- summary 放第一个键 → 流式解码时约 1 秒内即可把“我理解的是……”先给用户确认；
- version / inputs 由代码补齐，不浪费生成 token。
"""
from __future__ import annotations

import json

# 冻结：下面这句系统提示与 SFT 训练时逐字一致（"StarTorch" 是训练期沿用的代号）。改动任何一个字都会让输入分布偏离训练分布，
# 实测会降低计划正确率，所以开源版保持原样。
SYSTEM = "你是 StarTorch 表格指令编译器。根据表格画像，把用户的中文或英文指令编译成 Action Plan JSON（v0.1），只输出 JSON。"


def user_msg(file: str, profile: str, text: str) -> str:
    return f"文件：{file}\n{profile}\n指令：{text}"


def target(plan: dict) -> str:
    out = {"summary": plan.get("summary", ""), "lang": plan["lang"], "steps": plan["steps"]}
    if plan.get("ambiguities"):
        out["ambiguities"] = plan["ambiguities"]
    return json.dumps(out, ensure_ascii=False, separators=(",", ":"))


def messages(file: str, profile: str, text: str) -> list[dict]:
    return [{"role": "system", "content": SYSTEM}, {"role": "user", "content": user_msg(file, profile, text)}]


def complete(raw: dict, file: str, sheet: str | None, header_row: int = 1) -> dict:
    """模型输出 → 完整 plan（补 version / inputs）。"""
    plan = {"version": "0.1", **raw}
    plan["inputs"] = [{"alias": "src", "path": file, **({"sheet": sheet} if sheet else {}),
                       **({"header_row": header_row} if header_row and header_row != 1 else {})}]
    return plan
