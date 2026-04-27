# Performance Optimization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reduce end-to-end processing time for invoice, stock, declaration template, and register flows without changing outputs, workbook styling, sheet order, or business results.

**Architecture:** Keep the current UI and flow boundaries intact. Optimize by first instrumenting the pipeline, then reducing repeated workbook I/O and repeated dataframe work, and finally tightening register-path workbook access while preserving written values and styles exactly.

**Tech Stack:** Python, pandas, openpyxl, PyQt6, unittest

---

## File Map

- Modify: `invoice_data.py`
  - Add timing instrumentation around major invoice flow stages
  - Reduce repeated workbook/dataframe work in invoice-related register/template paths
- Modify: `stock_preparation.py`
  - Add timing instrumentation around major stock flow stages
  - Reduce repeated workbook/dataframe work in stock-related register/template paths
- Modify: `app/core/page_setup.py`
  - Add timing/log-friendly hooks without changing page setup behavior
- Modify: `app/core/metrics.py`
  - Keep behavior same; only touch if timing/log formatting needs shared helper support
- Test: `tests/test_invoice_data.py`
  - Add regression tests for any extracted helpers or unchanged output-path behavior
- Test: `tests/test_stock_preparation.py`
  - Add regression tests for any extracted helpers or unchanged output-path behavior

### Task 1: Add Stage Timing Instrumentation

**Files:**
- Modify: `invoice_data.py`
- Modify: `stock_preparation.py`
- Modify: `app/flows/invoice_flow.py`
- Modify: `app/flows/stock_flow.py`
- Test: no new automated test required; verify with targeted command and log output inspection

- [ ] **Step 1: Add a tiny timing helper in both generator modules**

```python
def _log_timed_step(self, label: str, start_time: datetime.datetime) -> None:
    duration = (datetime.datetime.now() - start_time).total_seconds()
    self.log_message(f"{label}耗时: {duration:.3f}s")
```

- [ ] **Step 2: Wrap declaration-data preparation in timing logs**

```python
step_started_at = datetime.datetime.now()
self._prepare_declaration_data(...)
self._log_timed_step("准备报关数据", step_started_at)
```

- [ ] **Step 3: Wrap workbook rendering in timing logs**

```python
step_started_at = datetime.datetime.now()
render_detailed_list(...)
render_packing_list(...)
render_bill(...)
render_contract(...)
render_customs_declaration(...)
render_elements_sheet(...)
self._log_timed_step("渲染报关资料工作簿", step_started_at)
```

- [ ] **Step 4: Wrap template fill and register steps in timing logs**

```python
step_started_at = datetime.datetime.now()
declaration_template_path = generator.copy_declaration_template_file(output_dir)
generator.fill_declaration_template_file(declaration_template_path)
generator._log_timed_step("填充申报模板", step_started_at)
```

```python
step_started_at = datetime.datetime.now()
generator.flow.register_all_documents(...)
generator._log_timed_step("登记数据", step_started_at)
```

- [ ] **Step 5: Verify the app still passes existing tests**

Run: `python -m unittest discover -s tests -v`
Expected: `OK`

- [ ] **Step 6: Commit**

```bash
git add invoice_data.py stock_preparation.py app/flows/invoice_flow.py app/flows/stock_flow.py
git commit -m "chore: add performance timing instrumentation"
```

### Task 2: Reduce Repeated Workbook Opens During Page Setup

**Files:**
- Modify: `app/core/page_setup.py`
- Modify: `app/flows/invoice_flow.py`
- Modify: `app/flows/stock_flow.py`
- Test: `tests/test_core_page_setup.py`

- [ ] **Step 1: Write a failing test for optional COM stage skipping**

```python
def test_set_excel_page_view_can_skip_com_stage():
    workbook = Workbook()
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        temp_path = tmp.name
    try:
        workbook.save(temp_path)
        workbook.close()
        set_excel_page_view(temp_path, enable_com=False)
        reloaded = load_workbook(temp_path)
        ws = reloaded.active
        assert ws.sheet_view.view == "pageBreakPreview"
    finally:
        if "reloaded" in locals():
            reloaded.close()
        os.remove(temp_path)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_core_page_setup -v`
Expected: FAIL because `set_excel_page_view()` does not accept `enable_com`

- [ ] **Step 3: Add an optional `enable_com` parameter with default `True`**

```python
def set_excel_page_view(file_path: str, enable_com: bool = True) -> None:
    workbook = load_workbook(file_path)
    for sheet in workbook.worksheets:
        apply_persistent_page_setup(sheet)
    workbook.save(file_path)
    workbook.close()

    if not enable_com or win32 is None:
        return
```

- [ ] **Step 4: Keep current behavior by not changing existing call sites yet**

```python
generator.set_excel_page_view(output_filepath)
```

- [ ] **Step 5: Run focused and full tests**

Run: `python -m unittest tests.test_core_page_setup -v`
Expected: PASS

Run: `python -m unittest discover -s tests -v`
Expected: `OK`

- [ ] **Step 6: Commit**

```bash
git add app/core/page_setup.py tests/test_core_page_setup.py
git commit -m "refactor: make page setup com stage explicit"
```

### Task 3: Reuse Workbook State in Register Paths

**Files:**
- Modify: `invoice_data.py`
- Modify: `stock_preparation.py`
- Test: `tests/test_invoice_data.py`
- Test: `tests/test_stock_preparation.py`

- [ ] **Step 1: Write a failing test for helper-driven register file selection without repeated file scans**

```python
def test_register_flow_reuses_precomputed_row_counts():
    generator = CustomsDeclaration()
    generator.shipping_details = {"invoice_number": "INV-001"}
    generator.result_df = pd.DataFrame(...)
    with patch.object(generator, "find_last_data_row", side_effect=[5, 8]), \
         patch.object(generator, "check_invoice_number", return_value=0), \
         patch.object(generator, "write_invoicing_register") as write_register:
        ...
        assert write_register.called
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_invoice_data tests.test_stock_preparation -v`
Expected: FAIL because helper structure does not exist yet

- [ ] **Step 3: Extract helper(s) that compute row-count and duplicate-check state once per register file**

```python
def _prepare_register_write(self, register_filepath: str, sheet_name: str, invoice_number: str) -> tuple[int, int]:
    original_rows = self.find_last_data_row(register_filepath, sheet_name)
    deleted_rows = self.check_invoice_number(register_filepath, invoice_number, sheet_name, "合同号")
    return original_rows, deleted_rows
```

- [ ] **Step 4: Use the helper in both invoice and stock register flows**

```python
register_original_rows, deleted_rows = self._prepare_register_write(
    register_filepath, invoicing_sheet, self.shipping_details["invoice_number"]
)
expected_rows = register_original_rows - deleted_rows + len(register_df)
```

- [ ] **Step 5: Run focused and full tests**

Run: `python -m unittest discover -s tests -v`
Expected: `OK`

- [ ] **Step 6: Commit**

```bash
git add invoice_data.py stock_preparation.py tests/test_invoice_data.py tests/test_stock_preparation.py
git commit -m "refactor: reduce repeated register workbook scans"
```

### Task 4: Reduce Repeated DataFrame Cleanup in Register Preparation

**Files:**
- Modify: `invoice_data.py`
- Modify: `stock_preparation.py`
- Test: `tests/test_invoice_data.py`
- Test: `tests/test_stock_preparation.py`

- [ ] **Step 1: Write a failing test for normalized SKU/shop cleanup helper behavior**

```python
def test_register_combined_invoice_normalizes_shop_data_once():
    generator = CustomsDeclaration()
    source = pd.DataFrame([{"SKU": "AE123", "店铺": "ezarc-us"}])
    normalized = generator._normalize_register_shop_data(source.copy())
    assert normalized.loc[0, "SKU"] == "123"
    assert normalized.loc[0, "店铺"] == "EZARC"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_invoice_data tests.test_stock_preparation -v`
Expected: FAIL because helper does not exist yet

- [ ] **Step 3: Extract small normalization helpers and replace duplicated inline cleanup**

```python
def _normalize_register_shop_data(self, df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result["SKU"] = self.normalize_stock_sku(result["SKU"])
    shop_series = result["店铺"].astype(str).str.upper()
    result["店铺"] = np.select(
        [
            shop_series.str.contains("EZARC", na=False),
            shop_series.str.contains("TOLESA", na=False),
            shop_series.str.contains("LIBRATON", na=False),
        ],
        ["EZARC", "TOLESA", "LIBRATON"],
        default=shop_series,
    )
    return result
```

- [ ] **Step 4: Replace duplicated inline cleanup with helper calls**

```python
invoice_details_df = self._normalize_register_shop_data(invoice_details_df)
```

- [ ] **Step 5: Run focused and full tests**

Run: `python -m unittest discover -s tests -v`
Expected: `OK`

- [ ] **Step 6: Commit**

```bash
git add invoice_data.py stock_preparation.py tests/test_invoice_data.py tests/test_stock_preparation.py
git commit -m "refactor: share register dataframe normalization"
```

### Task 5: Validate Performance Gains Safely

**Files:**
- Modify: none unless small log formatting follow-up is needed
- Test: manual validation only

- [ ] **Step 1: Run full automated tests**

Run: `python -m unittest discover -s tests -v`
Expected: `OK`

- [ ] **Step 2: Run one representative发货单 and one representative备货单 manually**

Run:

```bash
python ui_pyqt.py
```

Expected:
- Output files generated successfully
- Workbook sheet order unchanged
- Register files written successfully
- Timing logs visible for major stages

- [ ] **Step 3: Compare timing logs before and after**

Check:
- `准备报关数据耗时`
- `渲染报关资料工作簿耗时`
- `填充申报模板耗时`
- `登记数据耗时`
- `设置分页视图耗时`

Expected: no regressions; at least one repeated-I/O hotspot reduced

- [ ] **Step 4: Commit**

```bash
git add .
git commit -m "perf: optimize document generation workflow"
```

## Self-Review

- Spec coverage:
  - Instrumentation covered by Task 1
  - I/O optimization covered by Tasks 2 and 3
  - DataFrame repetition reduction covered by Task 4
  - Full verification covered by Task 5
- Placeholder scan:
  - No `TODO` or `TBD` markers remain
- Type consistency:
  - Shared helper names are consistent across invoice and stock modules in this plan
