---
license: apache-2.0
base_model: Qwen/Qwen2.5-Coder-7B-Instruct
base_model_relation: finetune
language:
  - zh
  - en
library_name: gguf
pipeline_tag: text-generation
tags:
  - spreadsheet
  - excel
  - table
  - structured-output
  - json
  - on-device
  - edge
  - gguf
  - llama.cpp
  - qwen2
  - rome
  - lora
model-index:
  - name: TabularEdge-7B
    results:
      - task: {type: text-generation, name: Spreadsheet instruction → executable Action Plan}
        dataset: {type: custom, name: TabularEdge human colloquial set v2 (285 gold items)}
        metrics:
          - {type: accuracy, name: execution equivalence, value: 0.923}
---

# TabularEdge-7B (GGUF)

**Privacy-First Edge LLM for Local Spreadsheets.** TabularEdge-7B compiles plain-language spreadsheet requests (Chinese, English or mixed) into a JSON **Action Plan**, and the open-source [TabularEdge engine](https://github.com/1034592/TabularEdge-7B) executes that plan deterministically on local `.xlsx` files. The model runs fully offline on a laptop GPU with **4 GB of VRAM**.

| | |
|---|---|
| **Model repository** | `https://huggingface.co/AngranFei/TabularEdge-7B-GGUF` |
| **Code** | `https://github.com/1034592/TabularEdge-7B` |
| **File** | `TabularEdge-7B-Q4_K_M.gguf`: 4.36 GiB, sha256 `d00bdc02d74aa29a397e0f73c07903d8188022b345c11493893009ec0afd29bb` |
| **Base model** | [Qwen/Qwen2.5-Coder-7B-Instruct](https://huggingface.co/Qwen/Qwen2.5-Coder-7B-Instruct) (7.61B parameters, 28 layers, GQA 28Q/4KV, Apache-2.0) |
| **Adaptation** | 3 gated ROME rank-one weight edits (3 further edits rejected by the gate) → bf16 LoRA SFT (r=32) → LoRA continuation with raw-text regularization → merge → Q4_K_M quantization |
| **Languages** | Chinese, English, code-mixed |
| **License** | Apache-2.0 (derivative of an Apache-2.0 base; see NOTICE) |

## Intended use

- **Primary use:** an on-device planner for spreadsheet automation where data may not leave the machine (finance, public sector, healthcare back office).
- **Strengths:**
  - single- and two-table cleaning;
  - regex extraction;
  - cross-column live formulas;
  - XLOOKUP / INDEX-MATCH lookups;
  - conditional formatting;
  - group-by aggregation;
  - filter / unique / split;
  - file and folder operations;
  - asking for clarification when a request is ambiguous.
- **Input:** a *table profile* (sheet name, header names with types, at most 3 sample rows) plus the instruction. The model never sees the full table.
- **Output:** a JSON plan of typed ops. Run it only through a validating executor; the reference executor validates, confirms destructive steps and backs up before overwriting.

### Out of scope

- General chat or coding assistance. It still works, but it is not the target.
- One-shot planning of long multi-branch business rules that span 4 or more tables. Split them into steps.
- Free-form VBA or Python generation. The model is trained to emit plans, not code.

## How to use

```bash
llama-server -m TabularEdge-7B-Q4_K_M.gguf -c 4096 -ngl 21 --port 8080   # -ngl 21 fits a 4 GB GPU; -ngl 0 for CPU only
```

Prompt format (chat template of the base model). **The system prompt must be used verbatim.** It is the exact string used during training. "StarTorch" is the training-time codename and is kept unchanged on purpose.

```text
system: 你是 StarTorch 表格指令编译器。根据表格画像，把用户的中文或英文指令编译成 Action Plan JSON（v0.1），只输出 JSON。
user:   文件：ledger.xlsx
        [sheet=流水 rows=8] columns: 日期:text, 部门:text, 金额:num, 备注:text
          | 2026-10-01 | 销售部 | 12000 | 客户A（首付）
        指令：金额负数的标红，按部门汇总到新表
```

Output (illustrative; the `summary` key comes first so a UI can show "I understood: …" while the rest is still streaming):

```json
{"summary":"金额为负的单元格标红；按部门汇总金额到新表","lang":"zh","steps":[
 {"id":"s1","op":"conditional_format","sheet":"流水","args":{"where":{"col":"金额","cmp":"<","value":0},"apply_to":"cell","style":{"font_color":"FF0000"}}},
 {"id":"s2","op":"group_aggregate","sheet":"流水","args":{"by":["部门"],"aggs":[{"col":"金额","fn":"sum"}],"output_sheet":"按部门汇总","mode":"formula"}},
 {"id":"s3","op":"save","args":{"path":"ledger_处理后.xlsx"}}]}
```

Recommended settings:
- temperature 0 (repair rounds use 0.4);
- `cache_prompt: true`, and prefill the system prompt and table profile while the user is still typing; this gives a first token in about 0.2 s on a 4 GB GPU;
- `max_tokens` of about 600–2400, depending on prompt length.

Use [`tabularedge.llm_parser.parse`](https://github.com/1034592/TabularEdge-7B/blob/main/tabularedge/llm_parser.py). It handles streaming, sanitizing, validation and repair rounds.

## Training

### Data: 100 % synthetic or self-generated, no customer data

| Set | Size | How it was built |
|---|---|---|
| SFT v2 | 35,526 | Program-generated (instruction, plan) pairs across business themes (ledger, sales, expense, inventory, payroll, receivable, invoice, attendance; "project" is test-only). Every gold plan was **executed on a randomly generated real `.xlsx`** and kept only if execution succeeded. The set contains 8,392 **counterfactual pairs**: a minimal instruction change gives a different execution result, and pairs whose interventions had no causal effect were dropped. It also has 4,799 ambiguity samples covering 5 ambiguity classes, with 3,134 "clarified" siblings, and 7,479 samples with human-noise perturbations: homophone typos, pinyin abbreviations, inversions, code-mixing, punctuation noise. |
| Continuation ("quench") | 12,959 | 7,590 targeted samples (new intents, new phrasings, a new theme), 3,001 replayed v2 samples, and **self-distillation regularization**: answers from the original base model to 885 self-written general prompts, used both in chat format and as **raw text with full-token LM loss**. |
| Held-out families | — | One family of noise (the pinyin abbreviations `sx`/`bf`/`cf`), one theme and a keyboard-typo family were never seen in training. |

No third-party dataset and no user or customer data was used. The hand-written evaluation set was frozen (hash-locked) before the continuation data was generated.

### Procedure

1. **Bias scan:** first-token pair accuracy over 13 Excel function families on the base model. It found strong biases against `XLOOKUP` (0.05), `XMATCH` (0.0), `TEXTBEFORE` (0.0) and `IFS` (0.0).
2. **Causal tracing + ROME:** rank-one MLP edits at the layer each family's trace ranked highest: **XLOOKUP @ L10, TEXTBEFORE @ L25, XMATCH @ L7**. Every edit had to pass a gate:
   - all 13 families held, with no family dropping;
   - perplexity within tolerance (15.78 → 15.82);
   - first-token specificity of at least 17/18.

   Three IFS edits (L18, L20, L21) **failed** the gate and were rejected.

   **Why `IFS` resisted rank-one editing.** A rank-one update rewrites one key→value association inside a single MLP. The three families that landed were first-token *selection* errors — the function was known but ranked below a competitor — which is exactly that shape. `IFS` failed differently: the model emitted structurally degenerate nested branches, a sequence-level generation failure that no single key→value association controls. That is outside the stated scope of the method, not a hyper-parameter problem.

3. **Repetition neurons (explored, not shipped).** As a second route for `IFS`, repetition neurons were located with the activation-difference method of Hiraoka & Inui (NAACL 2025, [arXiv:2410.13497](https://arxiv.org/abs/2410.13497)) — mean MLP activation on "pre-repeat" versus "in-repeat" spans — then damped by scaling the corresponding `down_proj` columns. Sweep in [`results/final/measure/alpha_sweep.json`](results/final/measure/alpha_sweep.json):

   | scale α | `IFS` first-token acc. | other 12 families | chat PPL |
   |---|---|---|---|
   | 0.85 | 0.50 | all 1.00 | 22.68 |
   | 0.70 | 0.45 | all 1.00 | 20.77 |
   | 0.55 | 0.15 | all 1.00 | 19.31 |

   Damping monotonically **hurt** `IFS` while the other twelve families stayed at 1.00, so this path was rejected too. `IFS` was finally solved by fine-tuning. Both negative results are reported here because together they bound where weight-space surgery helps and where it does not.

4. **LoRA SFT:** bf16, r=32, 1 epoch over 35,526 samples, effective batch 16, on 1× RTX 4090 (4.9 h). A contrastive counterfactual loss term (λ=0.2) was used in earlier iterations.
5. **Continuation:** lr 1e-4, 1 epoch over 12,959 samples (119 min). Without the raw-text regularization, raw-text perplexity rose **+56 %**. With it, perplexity ended **2.4 % below** the base (15.78 → 15.40).
6. **Merge → GGUF → Q4_K_M** with llama.cpp. All evaluations below use the quantized file.

## Evaluation

Metric: **execution equivalence**. The predicted plan and the gold plan are both executed, and the resulting workbooks are compared value by value; plan text is not compared. All runs use temperature 0 with the Q4_K_M GGUF on llama-server.

| Test set | Rule baseline | v1 (SFT only) | **TabularEdge-7B** |
|---|---|---|---|
| Synthetic, in-distribution (536) | 0.071 | 0.424 | **0.972** |
| Synthetic, unseen themes (598) | 0.132 | 0.467 | **0.972** |
| Synthetic, held-out noise families (598) | 0.099 | 0.440 | **0.963** |
| Hand-written colloquial set v1 (34) | — | 0.735 | **0.912** |
| **Hand-written colloquial set v2 (285 with gold)** | — | 0.582 | **0.923** |
| Same, excluding 10 items near-duplicate to training text | — | 0.571 | **0.920** |

Hand-written v2 by category:
- pinyin abbreviations 0.804
- typos 1.0
- colloquial / inverted word order 0.949
- English / code-mixed 0.879
- multi-step / cross-table 0.907
- new Excel functions 0.947
- explicit destructive requests 1.0

**Ambiguity handling** (45 ambiguous items in the hand-written set):
- The model asks a clarifying question **86.7 %** of the time.
- Combined with the executor's mandatory confirmation of destructive steps, **91.1 %** of ambiguous items are intercepted.
- Unnecessary questions on non-ambiguous items: 1.9 %.

**General-capability checks** (against the base model):
- raw-text perplexity 15.78 → 15.40;
- chat-drift perplexity 1.1885 → 1.136;
- degenerate-loop rate 0 / 20;
- first-token specificity 0.889.

**Edge performance**, on an i7-11800H with an RTX 3050 Laptop 4 GB running llama.cpp b12119:

| Mode | Speed | Notes |
|---|---|---|
| Auto mode: CUDA, 21/28 layers on GPU | 16.8–17.0 tok/s | VRAM 3.5 / 4.0 GB |
| CPU only, 8 threads | 8.5 tok/s | |

| Latency | Median |
|---|---|
| Plan summary | 1.6 s |
| Full plan | 7.7 s |
| First token, with KV prefill | 0.2–0.5 s |

On an RTX 4090 the same model generates 141.6 tok/s (tg128).

## Limitations and risks

- **Not 100 % reliable.** About 8 % of colloquial instructions still produce a wrong plan. Always review the plan preview; the reference client shows a plain-language description of every step before running.
- English instructions (0.879) are weaker than Chinese ones.
- The weakest ambiguity class is "duplicate orders": whether to dedupe by order ID or by the whole row.
- Regex character classes are often under-specified; for example, "brackets" may cover only `()`/`（）`.
- The hand-written test sets were written by a single author, so author bias is likely. Independent blind evaluation is welcome.
- Long instructions that combine many tables and branches exceed what the model plans reliably in one shot.
- The model can emit destructive operations (`delete_rows`, `dedupe`, overwrite). **Never execute plans without a validating executor that confirms such steps.** The reference executor does this in code, independent of the model.

## Citation

```bibtex
@software{tabularedge7b_2026,
  title  = {TabularEdge-7B: Privacy-First Edge LLM and Native Execution Engine for Local Spreadsheets},
  author = {Angran Fei},
  year   = {2026},
  url    = {https://github.com/1034592/TabularEdge-7B}
}
```

Built on [Qwen2.5-Coder-7B-Instruct](https://huggingface.co/Qwen/Qwen2.5-Coder-7B-Instruct) (Apache-2.0) and [llama.cpp](https://github.com/ggml-org/llama.cpp) (MIT).
