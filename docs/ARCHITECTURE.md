# TabularEdge-7B Architecture

## 1. Design goals

| Goal | Consequence |
|---|---|
| Data never leaves the device | The model sees only a **table profile**: header names, types and ≤ 3 sample rows. The executor is the only component that reads full tables. |
| Run on a 4 GB laptop GPU | 7B model at Q4_K_M (4.36 GiB), split 21/28 layers between GPU and CPU. Context is 4096 tokens, so prompts must stay compact. |
| LLM output is untrusted | The model emits **data** (a typed plan), never code. Validation, safety and execution are deterministic. |
| Real files are large and messy | The streaming reader never materializes a large sheet. Profiling tolerates title rows, formatting residue and mixed key types. |
| Fail loudly | Every rejection names the step, the problem and how to rephrase. There is no silent fallback to guesswork. |

## 2. Request lifecycle

```
select file ──► profile (background thread)
                 ├─ ≤ 2 MB: openpyxl + formula evaluation (exact types for formula columns)
                 └─ > 2 MB: xlsx_stream (dimension + first 200 rows, lazy shared strings)
             ──► KV prefill: system prompt + file name + profile are sent to llama-server with cache_prompt
type instruction (debounced prefill continues)
Enter ──► prompt = frozen system prompt + profile (+ compact reference-table profiles)
      ──► stream decode; the "summary" field is surfaced first ("I understood: …")
      ──► JSON repair (unescaped quotes in formulas) ──► normalize ──► sanitize
      ──► complete (version, inputs, header_row) ──► structural post-normalize (IFERROR around REGEXEXTRACT)
      ──► validate: schema → sheet refs → reference-column refs → sense checks
            errors? → feed the exact errors back to the model (≤ 2 rounds, temperature 0.4), keep the best attempt
            still failing on cosmetic steps only → drop them and report it
            still failing → draft shown with ❌ on the bad steps; Run disabled; per-step explanation
Run ──► executor: load + trim formatting residue ──► each step (destructive / write steps → confirm)
     ──► Excel-2019 rewrite where needed ──► save as a new file (overwrite only after confirm + backup)
```

## 3. Components

### 3.1 Streaming OOXML reader (`tabularedge/xlsx_stream.py`)

An `.xlsx` file is a zip archive of XML parts. Even `openpyxl` in `read_only` mode parses all of `xl/sharedStrings.xml` when the file opens. On a 150 MB registry export with 3.35 M shared strings, that open alone took **67 s**.

`XlsxStream` uses only `zipfile` and byte-level regexes:

- **`dims`** reads `<dimension ref>` from the first kilobytes of the sheet part.
- **`head(n)`** decompresses only the beginning of the sheet stream, collects the shared-string *indices* those rows use, and then resolves only those indices with one streaming pass over `sharedStrings.xml`.
- **`lookup(key_col, ret_cols, keys)`** works in three passes, so memory is O(|keys| + |hits|) and independent of table size:
  1. Stream `sharedStrings.xml` and find the indices whose normalized text equals one of the keys.
  2. Stream the sheet and extract only the key and return columns. A fast path uses `findall` on `r="A1"` coordinates, with a general path for files that have no cell refs.
  3. Resolve the shared strings referenced by hit rows, again in one streaming pass.
- **`norm_key`** makes `123`, `123.0`, `"123"` and `" 123 "` compare equal. This handles credit codes and IDs typed as numbers in one file and as text in another.

### 3.2 Profiling and context (`client/core.py`, `tabularedge/protocol.py`)

- **Header-row detection:** skip title rows that have one filled cell or a merged banner. The header is the first row whose filled-cell count is close to the widest row and whose values are mostly text.
- **Used width:** formatting residue (borders painted across 239 columns) is trimmed by finding the last column that holds a value.
- **Reference tables:** the client finds them through file names mentioned in the instruction (a plain substring match against files in the folder, not a sentence pattern). A single-sheet file with a meaningless sheet name (`"0"`, `"Sheet1"`, a hash) is titled by its file name.
- **Compaction:** when there are more than 3 reference files, or the extra profile text exceeds 3,000 characters, each reference profile is reduced to its sheet name, its key columns and the columns the instruction mentions. This keeps the prompt within a 4k context.

### 3.3 Sanitizer (`tabularedge/sanitize.py`)

LLMs drift from a schema in predictable ways: synonyms, CSS-style keys, strings where booleans belong, flattened nesting. A strict validator rejects such plans even when they are semantically fine. The sanitizer **walks the JSON Schema** and repairs the plan before validation:

| Kind | Examples |
|---|---|
| Key aliases | `column → col`, `operator → cmp`, `font-weight/weight → bold`, `background → fill_color` |
| Enum aliases | `">" / "gt" / "大于"`; `"highlight" → conditional_format`; `"average" → avg` |
| Type coercion | `"true"`, `"1"`, `700` → `true`; numeric strings → numbers where the schema says number |
| Colors | Chinese and English names, `#RGB`, `#RRGGBB`, ARGB, `rgb(…)` → `RRGGBB` |
| Structure | nested style objects flattened; args written onto the step moved back under `args`; unknown non-core fields dropped; step ids renumbered with `depends_on` and `ambiguities[].step_id` remapped; `"F列"` → `"F"` |

**Invariant:** a plan that already validates is returned unchanged. This was checked on 110 recorded plans. Every change is reported as a `fixes` entry, so the UI can say "auto-corrected N issues".

### 3.4 Validation and repair (`tabularedge/protocol.py`, `tabularedge/llm_parser.py`)

1. **JSON Schema** (`schema/action_plan.schema.json`, Draft 2020-12): 20 ops, each with typed args.
2. **Semantic checks:**
   - step ids are unique;
   - each `depends_on` names an earlier step;
   - regexes compile;
   - every column reference exists in the profile or was created earlier in the plan;
   - `[Column]` references inside formulas are checked the same way;
   - cross-table `from_sheet` / `right` names exist;
   - reference-table columns exist.
3. **Sense checks** catch plans that are valid but certainly wrong: a row formula pinned to a fixed cell (`=H2`), or a new column named just a column letter.
4. **Repair loop:** the exact error strings go back to the model in a follow-up turn. If the output was truncated at `max_tokens`, the loop stops and the user is advised to split the instruction.

The runtime pipeline has **no rule-based parser and no instruction-text templates**. `benchmarks/rule_baseline.py` exists only as the baseline row in the accuracy tables and as the seed labeller used in data generation.

### 3.5 Executor (`tabularedge/executor.py`)

- **Ops:**
  - cleaning: `trim`, `regex_replace`, `fill_empty`, `delete_rows`, `dedupe`, `sort`;
  - formulas and formatting: `add_column`, `lookup_column`, `lookup_join`, `conditional_format`, `number_format`;
  - new sheets and files: `group_aggregate`, `filter_to_sheet`, `unique_to_sheet`, `split_to_files`, `save`, `fs_*`.
- **Live formulas:** templates such as `=[单价]*[数量]` are expanded per row. In range-argument positions, `[Column]` becomes a whole-column range (`MAXIFS([更新时间],[订单号],[订单号])`).
- **Reference maintenance:** after row deletes or sorts, every formula in the workbook is re-anchored the way Excel would do it. Without this, formulas silently point at the wrong rows.
- **`_xlfn.` prefixes** are added for about 40 functions introduced after Excel 2007 (`_xlfn._xlws.` for FILTER/SORT); without them Excel shows `#NAME?`.
- **Safety:** a path sandbox under the workspace; `delete_rows`, `dedupe` and overwrites always go through `confirm`, whatever the model said; automatic backup before any overwrite.
- **Big references:** a reference table above 20 MB or 1.5 M cells is never copied into the output. `lookup_column` / `lookup_join` run through `XlsxStream.lookup` and write values. A main table above 3 M cells is refused, with an explanation, before anything loads.

### 3.6 Excel 2019 compatibility (`tabularedge/compat.py`)

`detect_excel()` reads the installed Office version on Windows, and `config.json` can override it. When the target is 2019:

| 365 function | Rewrite |
|---|---|
| `XLOOKUP` exact | `IFERROR(INDEX(ret, MATCH(key, col, 0)), default)` |
| `XLOOKUP` approximate (tier tables) | `INDEX/MATCH` + `MAXIFS`/`MINIFS`, which needs no sorted range |
| `XMATCH` | `MATCH` with mapped match modes |
| `UNIQUE` (to sheet) | classic `INDEX/MATCH/COUNTIF` array formula with 20 spare rows, so it stays live |
| `FILTER`, `REGEXEXTRACT`, `TEXTBEFORE`, … | computed values, with a note in the run log |

### 3.7 Formula evaluator (`tabularedge/formula_eval.py`)

`openpyxl` stores formulas but does not compute them. A pure-Python evaluator for 90 functions covers lookups, conditional aggregation, text, dates, regex and error semantics. With it, the executor can group, filter or split on formula columns, and the test suite can compare *computed* results. 1,407 gold formulas were cross-checked against LibreOffice.

### 3.8 Desktop client (`client/`)

- **Tkinter UI.** Every file access and every model call runs on worker threads, and the UI thread drains a queue with a 30 ms budget per tick. Token deltas are coalesced. The raw-output pane uses `wrap="char"`, because word-wrapping a JSON line with no spaces made each insert O(n²).
- **Engine manager:** reads free VRAM with `nvidia-smi`, estimates how many layers fit, and walks down CUDA → fewer layers → Vulkan → CPU. On Windows, `llama-server` is bound to a Job Object, so a killed client never leaves an orphaned GPU process.
- **Measured UI responsiveness** with a 10 ms heartbeat probe: the longest stall during parse and run on the 150 MB workbook was 0.11 s.

## 4. Model pipeline (`synth/`, `surgery/`)

```
Qwen2.5-Coder-7B-Instruct
  │  surgery/: causal tracing per Excel-function family → ROME rank-one MLP edits
  │            gate = 13-family pair accuracy + perplexity + first-token specificity
  │            accepted: XLOOKUP@L10, TEXTBEFORE@L25, XMATCH@L7   rejected: IFS@L18/L20/L21
  ▼
LoRA SFT (bf16, r=32) on 35,526 executed-and-verified synthetic samples (synth/gen2.py + synth/noise.py)
  ▼
LoRA continuation on 12,959 samples incl. self-distilled raw-text regularization (synth/mix_final.py)
  ▼
merge → GGUF → Q4_K_M (synth/export_gguf.py)  →  evaluation by execution equivalence (synth/eval_plans.py, synth/eval_human.py)
```

The full training corpora are not committed. They are regenerated deterministically by `synth/gen.py` / `synth/gen2.py`, which use fixed seeds. The committed `data/` holds the v1 training set, all synthetic test sets and the two hand-written human test sets.
