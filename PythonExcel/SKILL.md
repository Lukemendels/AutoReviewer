---
name: python-in-excel
description: Use this skill whenever the user is writing or troubleshooting Python inside Microsoft Excel — including any mention of =PY() cells, the xl() ingestion function, Python DataFrames inside Excel, #SPILL/#PYTHON/#CALC/#TIMEOUT/#BUSY errors, the Initialization Pane, "DataFrame card" previews, or toggling between Python Object and Excel Value output states. Always trigger this skill if the user is staging data for ingest into a =PY() cell, even if they don't name "Python in Excel" explicitly. The skill enforces three approved data-ingest strategies (named range/Excel Table, JSONL-in-a-column, Power Query connection-only), the row-major execution model, the cloud-sandbox limits (no internet, ~100MB payload, ephemeral container), and the correct output-state choice for intermediate vs. final cells.
---

# Python in Excel — Authoring & Ingest Skill

This skill helps the user write Python code that runs inside Excel's `=PY()` cells. It assumes the runtime is the standard Microsoft cloud-hosted Anaconda distribution, with no internet access, no local file system, and no `pip install`.

**Before writing any code, think carefully and step-by-step through the user's request.** Reason explicitly about which ingest strategy fits, where the cells should live, and what could go wrong before producing output. Do not skip this even if the request seems simple — the failure modes in Python in Excel are unusual enough that pattern-matching on standard Python advice will produce broken code.

Then work through these three confirmations explicitly with the user:

1. **Where will the data come from?** (Already in the workbook, pasted as JSONL, or external?) → picks the ingest strategy.
2. **Is Sheet 1 already set up as an initialization sheet?** → if not, propose it.
3. **Is this an intermediate calculation or a final output cell?** → drives the output-state choice (DataFrame Object vs. Spill).

Get those answers before producing code.

---

## 1. Architecture in one paragraph (why these rules exist)

`=PY()` cells execute in a hypervisor-isolated Azure container, not on the local machine. The container has **no internet, no local file access, no VBA/macro/PivotTable visibility, and is destroyed at session end.** The only way data crosses into Python is through the `xl()` function, and the only way results come back is by rendering the cell as a Python Object (a "card") or as an Excel Value (a spilled dynamic array). All cells in the workbook share **one global Python namespace**, evaluated in **strict row-major order** — top-left of Sheet 1, sweeping right, then down, then on to Sheet 2, etc. A variable defined in C5 is visible in D5 and on every later sheet, but **not** in A2 of the same sheet. Treat that ordering as a hard constraint when laying out the workbook.

**Concrete example of the row-major trap.** This is the single most common bug in Python in Excel:

```python
# ❌ Cell A2 (Sheet 1) — runs FIRST in row-major order
df_summary = df_raw.groupby("Region")["Throughput"].sum()
# Returns #PYTHON!  →  NameError: name 'df_raw' is not defined

# Cell A4 (Sheet 1) — runs AFTER A2, too late
df_raw = xl("PassengerTraffic[#All]", headers=True)
```

The fix is structural, not syntactic: swap the cell positions so ingestion happens first (A2) and analysis happens second (A4). This is exactly why Section 2b dedicates Sheet 1 to ingest-only — putting all `xl()` calls in the upper rows of Sheet 1 makes this entire failure class impossible.

Hard limits to keep in mind:
- ~100 MB per-cell data payload → exceeding it returns `#CALC!`
- Container timeout → exceeding it returns `#TIMEOUT!` (extendable in Advanced settings)
- No outbound network → `pandas.read_csv("https://...")`, `requests.get(...)`, `SQLAlchemy` etc. **will fail**

---

## 2. Step 1: Set up Sheet 1 as the Initialization Sheet

There are two complementary setup locations. Use both.

### 2a. The Initialization Pane (Formulas ribbon → "Initialization")

This is Microsoft's per-workbook startup script, equivalent to `__init__.py`. It runs once per session and resets the Python runtime when saved. **Do not delete the default `import excel`, `excel.set_xl_scalar_conversion()`, or `excel.set_xl_array_conversion()` lines — they are the bridge between the Excel grid and the Python interpreter.**

Add specialized imports here so they're available everywhere without repeating them:

```python
# In the Initialization Pane (Formulas ribbon → Initialization)
# DO NOT remove the default lines above this block.

import json
import re
from datetime import datetime, timedelta

# Explicit imports (these are NOT pre-loaded by default):
import sklearn
import networkx as nx
# import nltk          # uncomment if needed
```

Pre-loaded by default (do **not** re-import in the pane): `pandas as pd`, `numpy as np`, `matplotlib.pyplot as plt`, `seaborn as sns`, `statsmodels as sm`.

### 2b. Sheet 1 — the "Init" worksheet

Because evaluation is row-major and starts at Sheet 1 cell A1, dedicate Sheet 1 to **data ingestion and variable definition only**. No analytics, no plots. This guarantees that every downstream sheet has the variables it needs in scope.

Suggested Sheet 1 layout:

| Cell | Purpose |
|------|---------|
| A1   | Title / sheet description (plain text, not `=PY()`) |
| A3   | `=PY()` — ingest primary table → assign to `df_raw` |
| A5   | `=PY()` — ingest secondary table or JSONL → assign to `df_lookup` (or similar) |
| A7   | `=PY()` — basic cleaning / type coercion → assign to `df` |
| A9   | `=PY()` — derived constants the rest of the workbook needs |

Render every Sheet 1 cell as a **Python Object** (DataFrame card preview), not as a spilled value. Sheet 1 is for the runtime, not the user.

---

## 3. Step 2: Choose ONE of three ingest strategies

All three rely on the `xl()` function. `xl()` is the only authorized bridge into the container.

### Routing — pick the strategy based on what the user says or has

Match the user's situation to the right column. Do **not** default to Strategy A unless the data is genuinely already in the workbook as a table or range.

| User says or situation | Use Strategy |
|---|---|
| Data is already a Table, named range, or pasted block in the workbook | **A** — Named range / Excel Table |
| "I have a CSV/Excel file I can paste in" | **A** — Named range / Excel Table |
| "I copied this from an API response" / nested JSON / irregular records | **B** — JSONL in a single column |
| "The data has a weird shape" / mixed schemas / I don't want to flatten it | **B** — JSONL in a single column |
| "I'm staging records from a system that exports JSON" | **B** — JSONL in a single column |
| "I need to pull from SQL / SharePoint / OData / Azure / a database" | **C** — Power Query (connection-only) |
| "I want this to refresh when the source updates" | **C** — Power Query (connection-only) |
| "It's a CSV that lives on SharePoint" or any URL the user names | **C** — Power Query (connection-only) |
| "I've been hitting `pandas.read_csv` with a URL and it fails" | **C** — Power Query (connection-only) |

If the situation isn't a clear match, **ask which one applies before producing code.** Picking the wrong strategy wastes the user's time on workbook setup that has to be redone.

### Strategy A — Named range or Excel Table (default, preferred)

**When to use:** the data already lives in the workbook as a structured Table or a stable named range. This is the right answer ~70% of the time.

**Why prefer Tables over raw ranges:** Tables auto-expand. A reference like `xl("PassengerTraffic[#All]", headers=True)` keeps working when rows are added, with no code change.

```python
# Sheet 1, cell A3 — render as Python Object (DataFrame card)
df_traffic = xl("PassengerTraffic[#All]", headers=True)
df_traffic
```

Key parameters:
- `[#All]` after the table name → ingest header row + all data rows.
- `headers=True` → first row becomes the DataFrame's column index. **Always set this** when the source has headers; otherwise pandas will treat the header row as data.

For a named range without a header row, omit `headers=True`:

```python
df_block = xl("EconomicConstants")   # named range, no headers
```

### Strategy B — JSONL in a single column

**When to use:** the user has structured data they can't easily reshape into a wide table — nested objects, irregular records, API response captures, anything they'd normally hand a JSON parser. Also the right call when staging data copy-pasted from another system into a clean, narrow column.

**The pattern:** create an Excel Table with **one column** (e.g., named `json_line`), where each cell holds **one JSON object as a string**. Reference the table by name; parse each row in Python.

Setup in the workbook:
1. Create a new sheet (or use a staging area on Sheet 1).
2. Insert a Table named, say, `JsonlData` with a single column header `json_line`.
3. Paste one JSON object per cell down the column. Each cell's text should be a complete, valid JSON record like `{"id": 1, "region": "NE", "throughput": 12044}`.

Then in `=PY()`:

```python
import json
import pandas as pd

raw = xl("JsonlData[#All]", headers=True)            # one-column DataFrame
records = [json.loads(line) for line in raw["json_line"].dropna()]
df = pd.DataFrame(records)
df
```

This pattern bypasses the wide-vs-long structuring problem entirely — the JSON records carry their own schema, and `pd.DataFrame(records)` produces a clean tidy frame in one step. It is also the most resilient strategy for federal environments where richer data sources are blocked.

If individual records are large or contain newlines, wrap the cell value in quotes when pasting and confirm Excel hasn't truncated at the 32,767-character per-cell limit.

### Strategy C — Power Query connection-only (for external sources)

**When to use:** the data lives in a database, SharePoint file, OData feed, SQL server, or Azure blob — anywhere outside the workbook. `pandas.read_csv("https://...")` will fail because the container has no internet. Power Query is the authorized escape hatch: it runs **inside the Excel desktop process** (which has network access), pulls the data, and exposes it to Python via `xl()` by name.

**Critical:** load as **Connection Only**, not to the visible grid. Loading to the grid forces Excel to render millions of rows and bloats the file.

Setup in the workbook (one-time, no code):
1. **Data ribbon → Get Data** → pick the source (SQL, OData, SharePoint, CSV from a SharePoint URL, etc.).
2. In the Power Query Editor, do upstream filtering and column trimming. Push as much shrinking work to Power Query as possible — every row dropped here is a row Python doesn't have to carry against the 100 MB ceiling.
3. **Close & Load To...** → choose **Only Create Connection**. **Do not** click plain "Load."
4. Note the query name (e.g., `Live_Screening_Feed`).

Then in `=PY()`:

```python
feed_data = xl("Live_Screening_Feed")
feed_data
```

The pipeline is: **External source → Power Query connection → `xl()` → Python DataFrame.** This is the only sanctioned way to get live external data into Python in Excel without violating the container's network isolation.

---

## 4. Step 3: Pick the right output state for each cell

Every `=PY()` cell has two possible output states. Toggle with the formula bar dropdown, the right-click menu, or **Ctrl+Alt+Shift+M**.

| State | What the user sees | When to use |
|-------|-------------------|-------------|
| **Python Object** ("DataFrame card") | A small icon in the cell; hover shows schema preview (columns, dtypes, first/last rows, shape). Data stays in the cloud container's memory. | **All intermediate steps.** Keeps the workbook fast and small. The variable is still available to every downstream cell. |
| **Excel Value** (spill) | The DataFrame spills into adjacent cells as a dynamic array, with a faint blue border. | **Final outputs the user needs to read, format, share, or feed to a chart.** |

**Default rule:** intermediate cells stay as Python Object cards; final summary/result cells switch to Excel Value (Spill). When generating code, tell the user explicitly which state the cell should be in. Example:

> "Set cell A3 to **Python Object** (intermediate). Set cell A20 to **Excel Value / Spill** so the summary table renders on the grid."

**About `#SPILL!`:** this is **not** a Python error. It means the dynamic array can't expand because adjacent cells contain data — even a stray space character will block it. Resolution: locate and clear the obstructing cells. The Python code itself is fine.

---

## 5. Code patterns the user will need most often

### Reshape wide → long (almost always required before plotting/grouping)

Pandas, seaborn, and scikit-learn expect **tidy/long format**: one row per observation, one column per variable. Excel data usually arrives wide (e.g., one column per month). Reshape with `melt`:

```python
# df_wide: columns = ["Region", "Jan", "Feb", "Mar", ..., "Dec"]
df_long = df_wide.melt(
    id_vars=["Region"],
    var_name="Month",
    value_name="Throughput"
)
# df_long: columns = ["Region", "Month", "Throughput"], one row per Region-Month pair
```

### Replace common Excel formulas

| Excel | pandas |
|-------|--------|
| `=AVERAGE(A:A)` | `df["col"].mean()` |
| `=SUMIFS(...)` / PivotTable | `df.groupby("category")["value"].sum()` |
| `=VLOOKUP` / `=XLOOKUP` | `df.merge(lookup_df, on="key", how="left")` |
| `=SUBSTITUTE(A1, "-", "_")` | `df["col"].str.replace("-", "_")` |
| `=FILTER(...)` | `df.query('Status == "Active"')` |
| Summary stats block | `df.describe()` |

Filtering syntax notes: equality is `==`, the whole condition string is in quotes, multi-condition filters use parentheses around each clause:

```python
critical = df.query('(Status in ("Active", "Warning")) and (Throughput > 10000)')
```

### Plot to the grid

```python
import seaborn as sns
sns.heatmap(df.corr())
```

The plot returns as an image inside the cell. To free it from the grid: right-click the cell → **Display Plot over Cells**. This converts it to a free-floating image that can be moved and resized for dashboards.

---

## 6. Error reference (what to tell the user when things break)

| Error | Means | Fix |
|-------|-------|-----|
| `#PYTHON!` | Real Python error (NameError, SyntaxError, type mismatch). | Right-click cell → **Diagnostics for this cell** to read the traceback. Most common cause: referencing a variable defined in a cell that comes *later* in row-major order. |
| `#BUSY!` | Cloud container is computing. | Wait. If it persists past ~60 seconds, reset the runtime: Formulas ribbon, or **Ctrl+Alt+Shift+F9**. |
| `#TIMEOUT!` | Computation exceeded the runtime ceiling. | Optimize (vectorize, push filtering to Power Query). Or extend timeout in Excel Advanced settings → Python Formula Timeout. |
| `#CALC!` | Payload exceeded ~100 MB. | Filter/aggregate in Power Query before ingest. Don't pull whole tables when a slice will do. |
| `#SPILL!` | Dynamic array blocked by adjacent cell content. | Clear the obstructing cells. **Not a code bug.** |
| `#BLOCKED!` | Licensing or "Connected Experiences" disabled in Trust Center. | Enable Connected Experiences in Excel privacy settings; verify license. |
| `#CONNECT!` | Handshake failure with the Azure backend. | Reset the Python runtime; check network. |

---

## 7. Library availability cheat sheet

**Pre-loaded — use directly, no import needed in cells:**
- `pd` (pandas), `np` (numpy), `plt` (matplotlib.pyplot), `sns` (seaborn), `sm` (statsmodels)

**Available but require explicit `import` (do it in the Initialization Pane to avoid repeating):**
- `sklearn` (scikit-learn), `nx` (networkx), `nltk`

**Not available, never will be:** anything that needs network (`requests`, `urllib`, `SQLAlchemy` over a live connection), anything that touches the local filesystem, anything not in the curated Anaconda distribution.

---

## 8. Common failure modes — do not do these

These are mistakes that look reasonable in standard Python but break in Python in Excel. Do not generate code that does any of the following, and flag the user if you see them attempting these patterns:

- **Do not wrap output in `=PY(...)`.** Excel adds the wrapper automatically when the user types `=PY` and presses Tab. Code blocks should start directly with the Python statements.
- **Do not suggest `pip install` or `conda install`.** The runtime is a curated Anaconda distribution; arbitrary packages cannot be installed. If a library isn't in Section 7's cheat sheet, it isn't available — propose an alternative or push that work upstream into Power Query.
- **Do not use `pandas.read_csv("https://...")`, `requests.get(...)`, `urllib.request`, or `SQLAlchemy`** against a live endpoint. The container has no internet. These will fail. Always route external data through Strategy C (Power Query).
- **Do not read or write local files** (`open("C:/...")`, `pd.read_excel("file.xlsx")`, `os.path` operations). The container has no local filesystem access.
- **Do not omit `headers=True`** when ingesting an Excel Table that has a header row. Forgetting this turns the header row into a data row and the column index into `0, 1, 2, ...`, silently breaking every downstream column reference.
- **Do not re-import the pre-loaded libraries** in cells (`pd`, `np`, `plt`, `sns`, `sm`). They are already in the namespace.
- **Do not put intermediate calculations on Spill.** Spilling forces Excel to render every row on the grid, bloating the file and slowing recalculation. Only the final, user-facing cell should spill.
- **Do not load Power Query results to the grid** (the plain "Load" button). Always choose **Only Create Connection** so the data lives in workbook memory, not visible cells.
- **Do not claim a library exists without verifying it against Section 7.** If the user asks for something exotic (e.g., `xgboost`, `tensorflow`, `polars`, `pytorch`), check the cheat sheet first. If it isn't there, say so plainly and offer the closest available alternative (e.g., scikit-learn for xgboost, statsmodels for some torch use cases).
- **Do not interact with VBA, macros, PivotTables, charts, or named ranges programmatically from inside `=PY()`.** The Python runtime cannot see them. The bridge is `xl()` going in and Spill/Object going out — nothing else.
- **Do not assume the variable exists "somewhere in the workbook."** If you reference `df` in a cell, verify it was assigned in an earlier cell in row-major order. If the user hasn't set up Sheet 1 as an init sheet, propose it (Section 2b) before writing code that depends on global state.

---

## 9. Default response template

When the user asks for help with a Python in Excel task, structure the answer as:

1. **Confirm the ingest strategy** (A/B/C from Section 3) and name the Excel Table or Power Query connection it expects.
2. **Confirm the cell location** (which sheet, what row — Sheet 1 if it's setup; later sheets if it's analysis).
3. **State the output state** (Python Object for intermediate, Spill for final).
4. **Provide the `=PY()` code** in a single block, ready to paste.
5. **List any one-time setup the user must do** in the workbook (create the Table, set up the Power Query, add an import to the Initialization Pane).

Keep code blocks pasteable as-is — no surrounding `=PY(` wrapper, since Excel adds that automatically when the user types `=PY` and presses Tab.
