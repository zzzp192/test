# Tensile Report V3.14 Compatibility Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Upgrade tensile report processing to support both legacy and newer Excel exports while preserving current behavior.

**Architecture:** Add small worksheet-detection helpers around the existing tensile extraction and plotting paths. Prefer header-based identification, retain a legacy fixed-column fallback, and isolate curve worksheet selection so it can be tested without launching Origin.

**Tech Stack:** Python, `openpyxl`, `pandas`, standard-library `unittest`

---

### Task 1: Add Regression Tests

**Files:**
- Create: `tests/test_tensile_compatibility.py`

- [ ] **Step 1: Add workbook builders and failing tests**

Create in-memory temporary workbooks representing legacy and newer summary layouts, plus pandas workbook fixtures for curve sheet selection.

- [ ] **Step 2: Run tests to verify RED**

Run: `python -m unittest tests.test_tensile_compatibility -v`

Expected: failures because the newer report returns no groups and the curve worksheet selector does not exist.

### Task 2: Add Header-Based Summary Extraction

**Files:**
- Modify: `tensile_processor.py`
- Test: `tests/test_tensile_compatibility.py`

- [ ] **Step 1: Add normalized header matching**

Introduce helpers that locate summary sheets and map sample ID, thickness, Rp, Rm, Ag, A, and At columns.

- [ ] **Step 2: Preserve legacy fallback**

Keep the existing `Sheet1` fixed positions when dynamic detection cannot identify a complete summary layout.

- [ ] **Step 3: Run summary tests to verify GREEN**

Run: `python -m unittest tests.test_tensile_compatibility -v`

Expected: legacy and newer summary extraction tests pass.

### Task 3: Add Curve Worksheet Selection

**Files:**
- Modify: `origin_processor.py`
- Test: `tests/test_tensile_compatibility.py`

- [ ] **Step 1: Add paired stress/strain validation**

Introduce a helper that recognizes columns alternating between stress and strain.

- [ ] **Step 2: Select legacy and newer curve sheets**

Prefer `曲线` names, then `原始数据`, then any valid paired-column worksheet. Raise a clear error if no valid curve sheet exists.

- [ ] **Step 3: Use selector in plotting path**

Replace the current silent fallback to the first sheet.

- [ ] **Step 4: Run tests to verify GREEN**

Run: `python -m unittest tests.test_tensile_compatibility -v`

Expected: all compatibility tests pass.

### Task 4: Update Source Version

**Files:**
- Modify: `main.py`
- Modify: `README.md`

- [ ] **Step 1: Change V3.13 labels to V3.14**

Update the source version constant, window title, README heading, README version field, and changelog.

### Task 5: Verify and Commit

**Files:**
- Verify: `tests/test_tensile_compatibility.py`
- Verify: `样板数据/拉伸旧.xlsx`
- Verify: `样板数据/拉伸新.xlsx`

- [ ] **Step 1: Run automated tests**

Run: `python -m unittest tests.test_tensile_compatibility -v`

Expected: all tests pass.

- [ ] **Step 2: Run local sample verification**

Run extraction and curve sheet detection against both local sample workbooks.

Expected: legacy extracts 6 samples and selects `曲线数据(1)`; newer extracts 36 samples and selects `原始数据`.

- [ ] **Step 3: Stage only source, tests, and docs**

Do not stage `样板数据/`, cache files, existing build output, or unrelated local settings.

- [ ] **Step 4: Commit**

Create one V3.14 source upgrade commit.
