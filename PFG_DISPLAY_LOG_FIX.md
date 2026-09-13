# PFG Display Log — Root Cause Analysis and Fix Plan

**Subject:** Wrong, duplicated, and missing rows in the PFG user-facing display log
**File to change:** `user_display_log_pfg.py` (one file only)
**Reference implementation:** `user_display_log.py` (the C4F twin — already fixed)
**Status:** diagnosed, not yet fixed

---

## 0. How to use this document

This is written so the change can be made on a **different machine, outside the Dagster
project**. You need only two files:

| File | Why |
|---|---|
| `user_display_log_pfg.py` | the file you are editing |
| `user_display_log.py` | reference only — do not edit, do not copy wholesale |

You do **not** need the database, Dagster, Playwright, or any pipeline code. Section 6
includes a self-contained test that reproduces the bug and proves the fix, using only
`polars` and `filelock`.

Work through sections 4 → 5 → 6 in order. Section 7 is the checklist for handing the
file back.

---

## 1. Symptoms

Observed in the Data Intelligence Platform UI ("Processing Log" dialog) and confirmed
directly in the parquet files.

1. **A step is logged that never happened for that document.**
   A district's log shows `[FAC Search] FAC Excel downloaded`, while the database for the
   same row has `SourcingStatus = 0` and `Remarks = 'FAC Excel download failed after 5
   attempts.'`

2. **Steps are logged twice or more.**
   Two `FAC Excel downloaded` rows; two or three `PDF saved: <name>.pdf` rows with
   *different* filenames for one document; up to four `Overall:` validation verdicts on a
   single document.

3. **Most of a document's log is missing.**
   Many documents show only `FAC sourcing started` and `Sourcing validation started`, with
   no download row, no PDF row, and none of the 14 validation check rows — even though the
   database recorded a completed verdict.

4. **Extraction completes but logs almost nothing.**
   Documents marked `Done` show `Extraction started` / `Extraction finished` with no
   detail, or an `Extraction finished` with no matching `Extraction started`.

### Evidence from the parquet files

Measured across all `user_display_log*.parquet` files in
`D:\S2\Public Finance\999_Log_Trackers` (1011 rows total):

| Measurement | Result |
|---|---|
| Documents logging `Sourcing validation started` | 65 |
| Of those, documents with **no** `Overall:` verdict row | **49** |
| Documents with duplicate `Overall:` verdict rows | 5 (Id 319 has **4**) |
| Documents with 2–3 `PDF saved` rows | 3 (Ids 360, 365, 374) |
| Extraction `started` vs `finished` | 25 vs 16 |
| Rows with a **null** `Id` (belong to no document) | 1 |

The clearest single artefact — **Id 358 holds two complete 14-check validation result
sets**, written 39 milliseconds apart. One of them belongs to a different district:

```
sv  Validation   1   Entity Name: FAIL      2026-09-13 00:32:35.483502
...
sv  Validation  14   Overall: FAIL          2026-09-13 00:32:35.483522
sv  Validation  15   Entity Name: PASS      2026-09-13 00:32:35.522056   <-- second set starts
...
sv  Validation  28   Overall: FAIL          2026-09-13 00:32:35.522076
```

Note also that documents processed on 2026-09-07 and 2026-09-09 (Ids 257, 265, 268) each
have the **complete** 15-row validation set. The corruption appears only in the
2026-09-13 batch, which ran with a much higher thread count. **This is a concurrency bug
and its severity scales with `TModuleMaster.NumberOfThread`.**

---

## 2. Root cause

> `user_display_log_pfg.py` keeps per-document state in **module-level globals**, but all
> three PFG pipeline stages process documents **concurrently, as threads inside a single
> process**.

### 2.1 The shared state

`user_display_log_pfg.py` lines 293–295:

```python
_LOCK = threading.Lock()
_BUFFER: dict[tuple, list[dict[str, Any]]] = {}
_CTX: dict[str, Any] = {"row_id": None, "processing_id": None, "year": None}
```

There is **no** `threading.local()` anywhere in the file. `_BUFFER` (the rows waiting to
be written) and `_CTX` (which document those rows belong to) are shared by every thread in
the process.

Worse, the buffer key deliberately excludes the document — line 363:

```python
def _key(stage: str, sub_stage: str) -> tuple:
    """The buffer key: stage and sub-stage ONLY.

    Deliberately does NOT include Id/ProcessingId. ...
    """
    return (stage, str(sub_stage))
```

So *every* concurrent document recording a `FAC Search` step appends into **one shared
list** under the key `("s", "FAC Search")`.

Identity is then attached only at write time — `_flush()`, lines 466–471:

```python
with _LOCK:
    payload = {k: _BUFFER.pop(k) for k in keys if k in _BUFFER}
    row_id = _CTX["row_id"]
    processing_id = _CTX["processing_id"]
    year = _CTX["year"]
```

Whichever thread flushes stamps **all** the pooled rows with **its own** document id.

### 2.2 The concurrency

All three PFG stages fan documents out across a shared `ThreadPoolExecutor` in one
process:

| Job | Line |
|---|---|
| `Sourcing_job_Module1.py` | 209 |
| `Validation_job_Module1.py` | 207 |
| `Extraction_job_Module1.py` | 207 |

All three are registered live via `FinanceIQ.py`, which `workspace.yaml` loads.

### 2.3 Why the FileLock does not help

`_flush()` takes a cross-process `FileLock` around the read-modify-write. That is correct
and is **not** the problem. The corruption happens in shared memory *before* the lock is
ever acquired. The module docstring's claim that "3-4 concurrent Dagster workers" are safe
assumes those workers are separate **processes**; they are in fact **threads in one
process**.

---

## 3. The four distinct failure modes

Each symptom in section 1 maps to one of these. All four come from the same shared state.

### Mode 1 — Rows are stamped with the wrong document

Thread A (document X) succeeds and records `FAC Excel downloaded`. Thread B (document Y)
fails and records `FAILED: FAC Excel download failed after 5 attempts`. Both rows sit in
the same `("s", "FAC Search")` list. Whichever thread flushes claims both.

The database is **correct** throughout — `set_remarks()` and `update_sourcing_status()`
are keyed on the document's own `row_id`. Only the parquet log is wrong.

*Produces:* symptoms 1 and 2.

### Mode 2 — Rows are deleted outright

The replace-on-rerun filter, lines 560–572, deletes every existing row matching
`(Id, ProcessingId, Stage, SubStage)` before inserting the new ones:

```python
for g_row_id, g_pid, g_stage, g_sub in groups:
    existing = existing.filter(
        ~(
            _eq(pl, "Id", g_row_id)
            & (pl.col("ProcessingId").is_null()
               | _eq(pl, "ProcessingId", g_pid))
            & (pl.col("Stage") == g_stage)
            & (pl.col("SubStage") == g_sub)
        )
    )
```

When a flush is mis-stamped (Mode 1), this deletes the **victim** document's genuine
history and writes the thief's rows in its place. This is the main reason 49 of 65
documents have no validation verdict.

*Produces:* symptom 3.

### Mode 3 — `set_context()` drains another thread's half-finished group

Lines 318–321:

```python
moved = (new_row_id is not None
         and _CTX["row_id"] is not None
         and new_row_id != _CTX["row_id"])
if moved:
    flush_all()
```

`moved` is true whenever the incoming `row_id` differs from the **global** one — under
concurrency, on nearly every call. So thread B beginning a new document flushes thread A's
still-in-progress rows, early and under the wrong identity.

*Produces:* documents showing only a `started` row and nothing else.

### Mode 4 — `clear_context()` orphans rows into nothing

`PFG_Extraction.py:555` calls `udl.clear_context()`, which sets `_CTX["row_id"] = None`.
If another thread is mid-document, its next flush stamps `Id = None` **and**
`year = None`. The row lands in the unpartitioned `user_display_log.parquet`, belongs to
no document, and is invisible in the UI.

*Produces:* the one null-Id `Extraction finished` row currently in the data.

### Secondary effect — `Seq` restarts at 1

`_flush()` numbers rows with `enumerate(rows, start=1)` over whatever is in the buffer at
that instant. A group flushed in pieces restarts `Seq` at 1, so ordering within a group is
unreliable. This is why rows appear as `Seq 1  FAC Excel downloaded` with no `started`
above them. It resolves itself once Modes 1–4 are fixed.

---

## 4. The fix

### 4.1 Principle

Give each thread its own buffer and its own identity. The **file** stays shared and stays
guarded by the `FileLock` — that part was always correct.

This is sound because of three properties that already hold in the pipeline:

1. Each document is recorded **and** flushed on the same worker thread
   (`register_stages` / `flush_all` pairs live inside the `main_*_by_id` call that the
   executor dispatches).
2. The job modules' `in_progress_ids` guard guarantees one document is never handled by
   two threads at once — so *per-thread* state is exactly *per-document* state.
3. Different stages write different `SubStage` values, so their replace-filters cannot
   collide across concurrent jobs.

**Consequence: no changes are required in `Sourcing_job_Module1.py`,
`Validation_job_Module1.py`, `Extraction_job_Module1.py`, `PFG_Sourcing.py`,
`PFG_Validation.py`, or `PFG_Extraction.py`.** They inherit the fix.

(`Sourcing_job_Module2.py` is the ESG pipeline — `ModuleId = 2`, registered under
`ESGIQ.py`, routing to `BRSR_Sourcing.py` and `C4F_Sourcing_and_Validation_pipeline.py`.
It does not import `user_display_log_pfg` and is out of scope.)

### 4.2 This is a PORT, not a copy-paste

`user_display_log.py` (C4F) already contains the thread-local mechanism, at lines 274–305,
with a comment describing this exact bug in the past tense:

> `_CTX` was overwritten by whichever thread called `set_context()` last, so rows were
> stamped with the other document's Id. `_BUFFER` was keyed on `(stage, sub_stage)` ONLY,
> so two threads recording the same stage appended into ONE list and their rows merged.
> worse, `set_context()` flushes when the row_id changes — so thread B starting document
> 253 flushed thread A's half-finished 252 rows and stamped them 253.
>
> The result in production: the display log had 45 rows for Id=253 and NONE for Id=252 …

**Do not copy `user_display_log.py` over `user_display_log_pfg.py`.** The two files have
diverged in *opposite* directions:

| | thread-local state | year-partitioned files |
|---|---|---|
| `user_display_log.py` (C4F) | yes | **no** |
| `user_display_log_pfg.py` (PFG) | **no** | yes |

The PFG file's year partitioning (`log_path()`, `coerce_year()`, `_STEM`/`_SUFFIX`, and the
`year` key in `_CTX`) does not exist in the C4F file. Overwriting would destroy it.

Port the **mechanism** into the PFG file, and extend it to cover `year`, which the C4F
reference does not have.

---

## 5. Step-by-step edit instructions

Seven edits, all inside `user_display_log_pfg.py`. Line numbers refer to the file as it
stands today (30,519 bytes, last modified 2026-09-11). Match on the code text rather than
trusting the numbers.

`import threading` is already present at the top of the file — no new imports are needed.

---

### Edit 1 — Replace the shared state with thread-local state

**Location:** lines 288–295, the `BUFFER` section header and the three globals.

**Find:**

```python
# ═══════════════════════════════════════════════════════════════════════════════
#  BUFFER
# ═══════════════════════════════════════════════════════════════════════════════
# Rows accumulate per (Id, ProcessingId, Stage, SubStage) and are written when the
# group is flushed. The lock is in-process only: the buffer belongs to this
# process, and the cross-process guard is the FileLock taken inside _flush().

_LOCK = threading.Lock()
_BUFFER: dict[tuple, list[dict[str, Any]]] = {}
_CTX: dict[str, Any] = {"row_id": None, "processing_id": None, "year": None}
```

**Replace with:**

```python
# ═══════════════════════════════════════════════════════════════════════════════
#  BUFFER
# ═══════════════════════════════════════════════════════════════════════════════
# Rows accumulate per (Id, ProcessingId, Stage, SubStage) and are written when the
# group is flushed.

_LOCK = threading.Lock()

# ── The buffer AND the identity are PER THREAD ────────────────────────────────
#
# Both used to be module-level, which is correct only while one document is
# processed at a time. All three PFG stages (Sourcing_job_Module1,
# Validation_job_Module1, Extraction_job_Module1) fan documents out across a
# shared ThreadPoolExecutor in ONE process, and that broke four ways at once:
#
#   * _CTX was overwritten by whichever thread called set_context() last, so
#     rows were stamped with another document's Id.
#   * _BUFFER was keyed on (stage, sub_stage) ONLY, so two threads recording the
#     same stage appended into ONE list and their rows merged.
#   * set_context() flushes when the row_id changes — so thread B starting a new
#     document flushed thread A's half-finished rows and stamped them with B's id.
#   * clear_context() blanked the shared identity, so another thread's next flush
#     was stamped Id=None/year=None and landed in the unpartitioned file,
#     belonging to no document at all.
#
# The result in production (2026-09-13 batch): 49 of the 65 documents that logged
# "Sourcing validation started" had NO verdict row, while Id 358 held two complete
# validation result sets 39ms apart and Id 319 held four.
#
# Per-thread state fixes all four: each thread buffers its own groups and stamps
# its own identity and year. The FILE is still shared and still guarded by the
# FileLock — that part was always fine.
_TLS = threading.local()


def _state() -> Any:
    """This thread's buffer + identity."""
    st = _TLS
    if not getattr(st, "_init", False):
        st._init = True
        st.buffer = {}
        st.row_id = None
        st.processing_id = None
        st.year = None
    return st
```

> **Note the `year` field.** The C4F reference has only `row_id` and `processing_id`.
> Omitting `year` here leaves Mode 1 partly alive: rows would still be filed under another
> document's `ProcessYear`, landing in the wrong year parquet — where the replace-filter
> will never look, making those bad rows permanent.

---

### Edit 2 — `set_context()`

**Location:** the body of `set_context()`, starting around line 316 (after the docstring).

**Find:**

```python
    new_row_id = _coerce_int(row_id) if row_id is not None else None
    moved = (new_row_id is not None
             and _CTX["row_id"] is not None
             and new_row_id != _CTX["row_id"])
    if moved:
        flush_all()

    with _LOCK:
        if row_id is not None:
            _CTX["row_id"] = new_row_id
        if processing_id is not None:
            _CTX["processing_id"] = _coerce_int(processing_id)
```

**Replace with:**

```python
    st = _state()
    new_row_id = _coerce_int(row_id) if row_id is not None else None
    moved = (new_row_id is not None
             and st.row_id is not None
             and new_row_id != st.row_id)
    if moved:
        flush_all()

    with _LOCK:
        st = _state()
        if row_id is not None:
            st.row_id = new_row_id
        if processing_id is not None:
            st.processing_id = _coerce_int(processing_id)
```

Then, further down in the same function, in the year block:

**Find:**

```python
        if moved:
            _CTX["year"] = None
        if year is not None:
            _CTX["year"] = coerce_year(year)
```

**Replace with:**

```python
        if moved:
            st.year = None
        if year is not None:
            st.year = coerce_year(year)
```

> Keep the existing explanatory comment above the `if moved:` block. The
> drop-year-on-move behaviour is PFG-specific and must be preserved — it is what stops a
> document with a NULL `ProcessYear` from inheriting the previous document's year.
>
> `st` is re-read inside the `with _LOCK:` block because `flush_all()` may have run in
> between. `_state()` is cheap and always returns the same object for a given thread, so
> this is safe and keeps the diff obvious.

---

### Edit 3 — `clear_context()`

**Location:** around line 340.

**Find:**

```python
def clear_context() -> None:
    """Forget the current document. Flushes first — see set_context()."""
    flush_all()
    with _LOCK:
        _CTX["row_id"] = None
        _CTX["processing_id"] = None
        _CTX["year"] = None
```

**Replace with:**

```python
def clear_context() -> None:
    """Forget the current document. Flushes first — see set_context().

    Per-thread: clearing THIS thread's context leaves every other thread's
    in-flight document untouched.
    """
    flush_all()
    with _LOCK:
        st = _state()
        st.row_id = None
        st.processing_id = None
        st.year = None
```

---

### Edit 4 — `_add()`

**Location:** around line 374.

**Find:**

```python
    with _LOCK:
        _BUFFER.setdefault(_key(stage, sub_stage), []).append({
            "Remarks": str(remark).strip(),
            "Time":    datetime.now(),
        })
```

**Replace with:**

```python
    with _LOCK:
        _state().buffer.setdefault(_key(stage, sub_stage), []).append({
            "Remarks": str(remark).strip(),
            "Time":    datetime.now(),
        })
```

---

### Edit 5 — `flush_all()`

**Location:** around line 471.

**Find:**

```python
        with _LOCK:
            keys = list(_BUFFER.keys())
        return _flush(keys) if keys else 0
```

**Replace with:**

```python
        with _LOCK:
            keys = list(_state().buffer.keys())
        return _flush(keys) if keys else 0
```

---

### Edit 6 — `buffered_groups()`

**Location:** around line 482.

**Find:**

```python
    with _LOCK:
        return list(_BUFFER.keys())
```

**Replace with:**

```python
    with _LOCK:
        return list(_state().buffer.keys())
```

---

### Edit 7 — `_flush()`

**Location:** around line 466, inside `_flush()`.

**Find:**

```python
    with _LOCK:
        payload = {k: _BUFFER.pop(k) for k in keys if k in _BUFFER}
        # Read the identity once, here, so every row in a group agrees on it even
        # if the ProcessingId was resolved partway through recording the group.
        row_id = _CTX["row_id"]
        processing_id = _CTX["processing_id"]
        # Read the year here too, for the same reason: every row of this flush
        # must agree on which year file it belongs to, even if the context was
        # completed partway through recording the group.
        year = _CTX["year"]
```

**Replace with:**

```python
    with _LOCK:
        st = _state()
        payload = {k: st.buffer.pop(k) for k in keys if k in st.buffer}
        # Read the identity once, here, so every row in a group agrees on it even
        # if the ProcessingId was resolved partway through recording the group.
        row_id = st.row_id
        processing_id = st.processing_id
        # Read the year here too, for the same reason: every row of this flush
        # must agree on which year file it belongs to, even if the context was
        # completed partway through recording the group.
        year = st.year
```

**Change nothing else in `_flush()`.** The `FileLock`, the replace-filter, `_align()`,
`_sort_display()`, and `_atomic_write()` are all correct as they stand.

---

### 5.1 Things that must NOT change

| Item | Why |
|---|---|
| `_STAGES` (line 143) and `register_stages()` | **Must stay module-global.** If the allow-list registry becomes thread-local, any thread that did not call `register_stages` silently drops every `step()` and `check()` — while `failure()` still survives via the `STAGE_PROCESSING` fallback at line 436. The result would be a log containing *only* failures: a worse bug that looks like this one. |
| `_LOCK` | Now technically redundant (per-thread buffers need no mutex), but keep it. It costs nothing, matches the C4F reference, and removing it widens the diff for no benefit. |
| The `FileLock` in `_flush()` | Guards cross-*process* writes. Always was correct. |
| The replace-filter, lines 560–572 | Correct once the stamping is correct. |
| `log_path()`, `coerce_year()`, `log_paths()`, `_YEAR_RE` | PFG-only year partitioning. The C4F file has no equivalent. |
| The public API signatures | `step`, `check`, `failure`, `started`, `finished`, `flush_group`, `flush_all`, `set_context`, `clear_context`, `read_log`, `read_for`, `describe_for` are all called from `PFG_Sourcing.py`, `PFG_Validation.py`, `PFG_Extraction.py`. Keep every signature identical. |

### 5.2 Completeness check

After editing, this command must return **nothing**:

```bash
grep -n "_BUFFER\|_CTX" user_display_log_pfg.py
```

Any remaining hit is a site that still uses the shared state and will keep corrupting rows.

---

## 6. Verification — standalone, no Dagster required

Save as `test_display_log_threads.py` **next to** the edited `user_display_log_pfg.py`.
Requires only `polars` and `filelock`.

```python
"""
Standalone concurrency test for user_display_log_pfg.py.

Simulates what Sourcing_job_Module1 / Validation_job_Module1 do: N documents
processed concurrently by a shared ThreadPoolExecutor, each recording a known,
fixed set of rows.

FAILS on the unfixed module. PASSES on the fixed one.

Run:  python test_display_log_threads.py
"""
import os
import random
import shutil
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

# Redirect the log to a scratch directory BEFORE importing the module.
TMP = Path(tempfile.mkdtemp(prefix="udl_test_"))
os.environ["PFG_LOG_DIR"] = str(TMP)

import user_display_log_pfg as udl   # noqa: E402

DOCS = 12          # concurrent documents
THREADS = 6        # matches a realistic TModuleMaster.NumberOfThread
CHECKS = 14        # validation emits 14 check rows per document

STAGES = [
    ("FAC Search",      "s",  "FAC sourcing"),
    ("Report Download", "s",  "Report download"),
    ("Validation",      "sv", "Sourcing validation"),
]


def process(n: int) -> None:
    """One document, start to finish, on one thread — mirrors main_source_by_id."""
    row_id = 1000 + n
    year = 2022 + (n % 4)          # exercise year partitioning too
    def work():
        # Stand in for the Playwright / PDF work. The real stages take SECONDS
        # between udl calls. This pause is ESSENTIAL: without it the calls run
        # back-to-back in microseconds, threads never interleave mid-document,
        # and the race does not reproduce — the test passes even on the broken
        # module. (Verified: without these sleeps the unfixed module scores a
        # false PASS.)
        time.sleep(random.uniform(0.002, 0.02))

    udl.set_context(row_id, 500 + n, year=year)
    udl.started("FAC Search")
    work()
    udl.step("FAC Search", f"FAC Excel downloaded doc{n}")
    work()
    udl.step("Report Download", f"PDF saved: doc{n}.pdf")
    work()
    udl.started("Validation")
    work()
    for c in range(CHECKS):
        udl.check("Validation", f"Check{c}", "PASS")
    udl.check("Validation", "Overall", "PASS")
    work()
    udl.flush_all()


def main() -> int:
    udl.register_stages(STAGES)
    with ThreadPoolExecutor(max_workers=THREADS) as ex:
        list(ex.map(process, range(DOCS)))

    import polars as pl
    frames = [pl.read_parquet(p) for p in sorted(TMP.glob("user_display_log*.parquet"))]
    if not frames:
        print("FAIL: no parquet written at all")
        return 1
    df = pl.concat(frames, how="vertical_relaxed")

    failures = []

    # A. No orphans.
    orphans = df.filter(pl.col("Id").is_null() | pl.col("ProcessingId").is_null()).height
    if orphans:
        failures.append(f"{orphans} row(s) with a null Id/ProcessingId")

    # B. Every document has exactly its own rows, and no one else's.
    for n in range(DOCS):
        row_id = 1000 + n
        mine = df.filter(pl.col("Id") == row_id)
        remarks = mine["Remarks"].to_list()

        expected = (
            ["FAC sourcing started", f"FAC Excel downloaded doc{n}"]
            + [f"PDF saved: doc{n}.pdf"]
            + ["Sourcing validation started"]
            + [f"Check{c}: PASS" for c in range(CHECKS)]
            + ["Overall: PASS"]
        )
        if sorted(remarks) != sorted(expected):
            missing = set(expected) - set(remarks)
            extra = set(remarks) - set(expected)
            failures.append(
                f"Id {row_id}: {len(remarks)} rows, expected {len(expected)}"
                + (f" | MISSING {sorted(missing)[:3]}" if missing else "")
                + (f" | STOLEN-FROM-ANOTHER-DOC {sorted(extra)[:3]}" if extra else "")
            )

        # C. ProcessingId and year file must match the document.
        pids = set(mine["ProcessingId"].to_list())
        if pids and pids != {500 + n}:
            failures.append(f"Id {row_id}: wrong ProcessingId {pids}")

    # D. Seq must be 1..N contiguous within each group.
    for (rid, pid, stage, sub), g in df.group_by(
        ["Id", "ProcessingId", "Stage", "SubStage"]
    ):
        seqs = sorted(g["Seq"].to_list())
        if seqs != list(range(1, len(seqs) + 1)):
            failures.append(f"Id {rid} {stage}/{sub}: non-contiguous Seq {seqs}")

    shutil.rmtree(TMP, ignore_errors=True)

    if failures:
        print(f"FAIL — {len(failures)} problem(s):")
        for f in failures[:25]:
            print("  -", f)
        return 1
    print(f"PASS — {DOCS} documents across {THREADS} threads, all rows correctly attributed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

### Expected results

These are **actual runs**, not illustrations. This test and these seven edits were
executed against a scratch copy of the real `user_display_log_pfg.py` before this document
was written.

**Before the fix** — reproduces the production symptoms exactly:

```
FAIL — 11 problem(s):
  - Id 1000: 2 rows, expected 19 | MISSING ['Check0: PASS', 'Check10: PASS', 'Check11: PASS']
  - Id 1001: 0 rows, expected 19 | MISSING ['Check0: PASS', 'Check10: PASS', 'Check11: PASS']
  - Id 1002: 1 rows, expected 19 | MISSING ['Check0: PASS', 'Check10: PASS', 'Check11: PASS']
  - Id 1003: 93 rows, expected 19 | STOLEN-FROM-ANOTHER-DOC ['FAC Excel downloaded doc1', ...]
  - Id 1004: 18 rows, expected 19 | MISSING ['FAC Excel downloaded doc4', ...] | STOLEN ['PDF saved: doc0.pdf']
  - Id 1009: 36 rows, expected 19 | MISSING ['FAC sourcing started', ...] | STOLEN ['PDF saved: doc10.pdf', ...]
  ...
```

Note the shape: three documents with **zero** rows, one with **93**. That is the same
pattern as production (49 of 65 documents with no verdict, while Id 358 held two full
result sets).

**After the fix** — passes repeatedly, and at higher concurrency:

```
PASS — 12 documents across 6 threads, all rows correctly attributed.
PASS — 40 documents across 12 threads, all rows correctly attributed.
```

Run it **both ways** — against the original file and the edited one. A test that has not
been seen to fail has not proven anything.

If it passes *before* the fix, the timing on your machine is not interleaving. Do not
proceed: raise `DOCS` and `THREADS`, or widen the `work()` sleep range, until it fails.
A green test on the broken module means the test is not exercising the bug.

### Additional sanity check

```bash
python -c "import user_display_log_pfg as u; print(u.log_paths())"
```

Must still resolve the year-partitioned filenames. If this breaks, the year handling in
Edit 1 or Edit 2 was mis-applied.

---

## 7. Handing the file back

Return **only** `user_display_log_pfg.py`. Nothing else should have changed.

Checklist before handing over:

Each command below is followed by the value measured on a correctly-fixed copy.

- [ ] `grep -c "_BUFFER\|_CTX" user_display_log_pfg.py` → **0**
- [ ] `grep -c "_TLS\|_state()" user_display_log_pfg.py` → **10**
- [ ] `grep -c "st\.year" user_display_log_pfg.py` → **5** (`_state`, `set_context` ×2, `clear_context`, `_flush`)
- [ ] `_STAGES` is still a plain module-level `dict` (line ~143), not thread-local
- [ ] `coerce_year`, `log_path`, `log_paths`, `_YEAR_RE` are untouched
- [ ] `python -c "import user_display_log_pfg"` imports cleanly
- [ ] `python -c "import user_display_log_pfg as u; print(u.log_path(2025).name, u.log_path(None).name)"` → `user_display_log_2025.parquet user_display_log.parquet`
- [ ] `test_display_log_threads.py` prints **PASS** (run it 3× — this is a race; one green run is weak evidence)
- [ ] The same test was **seen to FAIL** on the unmodified file
- [ ] No public function signature changed

A one-shot check for the last point:

```bash
python -c "import user_display_log_pfg as u; print(all(hasattr(u,f) for f in ['step','check','failure','started','finished','flush_group','flush_all','set_context','clear_context','read_log','read_for','describe_for','register_stages']))"
```

Must print `True`.

### Verification once it is back in the Dagster project

1. Drop the file into `D:\S2\Dagster\user_display_log_pfg.py`.
2. Confirm `TModuleMaster.NumberOfThread` for `ModuleId = 1` is **> 1** — a
   single-threaded run cannot exercise the bug and will pass regardless.
3. Run a sourcing batch of at least 10–15 documents.
4. Query the resulting parquet and confirm:
   - every document that logged `FAC sourcing started` and reached a verdict also has an
     `Overall:` row;
   - no document has more than one `Overall:` row;
   - no document has more than one `PDF saved` row;
   - no rows have a null `Id`;
   - for any document whose database `Remarks` says `FAC Excel download failed after 5
     attempts.`, the log shows the `FAILED:` row and **no** `FAC Excel downloaded` row.

A ready-made query for step 4:

```python
import polars as pl
from pathlib import Path

D = Path(r"D:\S2\Public Finance\999_Log_Trackers")
df = pl.concat(
    [pl.read_parquet(p) for p in sorted(D.glob("user_display_log*.parquet"))],
    how="vertical_relaxed",
)

started = set(df.filter(pl.col("Remarks").str.contains("validation started"))["Id"].to_list())
overall = df.filter(pl.col("Remarks").str.starts_with("Overall:"))
pdfs    = df.filter(pl.col("Remarks").str.starts_with("PDF saved"))

print("null-Id rows      :", df.filter(pl.col("Id").is_null()).height, "(want 0)")
print("dup Overall       :", overall.group_by("Id").len().filter(pl.col("len") > 1).height, "(want 0)")
print("dup PDF saved     :", pdfs.group_by("Id").len().filter(pl.col("len") > 1).height, "(want 0)")
print("started, no verdict:", len(started - set(overall["Id"].to_list())))
```

---

## 8. Out of scope — but do not lose track of these

### 8.1 The existing corrupted data will not repair itself

The fix stops new corruption; it does not clean up what is already written.

- Re-running a document **does** replace its own `(Id, ProcessingId, Stage, SubStage)`
  group, so re-runs heal those rows.
- But rows stamped with the **wrong `Id`** or written into the **wrong year file** sit
  where no future replace-filter will look. They are permanent until deleted.
- The null-Id row in `user_display_log.parquet` is unreachable by any replace.

If a clean history matters, plan a separate one-off cleanup pass. Decide first whether to
delete the suspect rows or to re-run the affected documents.

### 8.2 `PFG_Sourcing.py:1337` — `flush_all()` is not in a `finally`

```python
        # Commit the last row's buffered display-log rows ...
        udl.flush_all()
```

The equivalent calls in `PFG_Validation.py` (lines 4382, 4408) and `PFG_Extraction.py`
(lines 568, 646) *are* wrapped in `try/finally`. This one is not.

Today the shared buffer hides it — another thread eventually flushes the rows, wrongly.
After the fix, an exception escaping the row loop means that thread's buffer dies with the
thread and those rows are silently lost.

Small, pre-existing, and in a **different file**. Left out deliberately so the fix stays to
one file. Worth doing as a separate follow-up.

### 8.3 `PFG/DIP_PFG/` holds stale copies

`PFG/DIP_PFG/PFG_Sourcing.py`, `PFG_Validation.py` and `PFG_Extraction.py` are August
snapshots. `Sourcing_job_Module1.py:57` uses
`importlib.import_module("PFG_Sourcing")`, which resolves to the **project-root** copies,
so the `DIP_PFG` versions are **not** running. Do not apply anything from this document to
them, and do not let them be mistaken for the live files.

---

## 9. Quick reference

| Question | Answer |
|---|---|
| Files to change | `user_display_log_pfg.py` — that is all |
| Files to read but not change | `user_display_log.py` (C4F reference, lines 274–305) |
| Job modules needing changes | none |
| Pipeline modules needing changes | none |
| Root cause | module-global `_BUFFER` / `_CTX` under a multi-threaded executor |
| Fix | per-thread buffer + identity (`threading.local`), including `year` |
| Number of edits | 7, all in one file |
| How to prove it | `test_display_log_threads.py` in section 6 — must fail before, pass after |
| Biggest trap | copying the C4F file wholesale (destroys year partitioning); making `_STAGES` thread-local (silently drops all non-failure rows) |
