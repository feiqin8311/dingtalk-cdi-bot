# Invoice Customs Mode Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Change the invoice-mode customs declaration sheet so `监管方式` renders as `9710(<发票号>)`, with the appended invoice-number suffix colored red when supported.

**Architecture:** Keep the workbook generation flow unchanged and localize the behavior to `invoice_data.CustomsDeclaration.add_customs_declaration_sheet()`. Add a focused test that exercises the target cell directly in an in-memory workbook and verifies both the text content and rich-text styling.

**Tech Stack:** Python, `openpyxl`, `unittest`

---

### Task 1: Cover the new监管方式 formatting with a failing test

**Files:**
- Modify: `tests/test_invoice_data.py`
- Test: `tests/test_invoice_data.py`

- [ ] **Step 1: Write the failing test**

```python
def test_customs_declaration_mode_includes_invoice_number_in_red_suffix(self):
    generator = CustomsDeclaration()
    generator.shipping_details.update(
        {
            "domestic_shipper": "SHIPPER",
            "production_sales_unit": "PROD",
            "invoice_number": "26YP065",
            "country_zh": "美国",
            "TOTAL_PACKED_IN": 1,
            "TOTAL_G.W": 2.5,
            "TOTAL_N.W": 2.0,
            "freight": 12.3,
            "domestic_source_of_goods": "33019 杭州其他",
            "images": ["image0", "image1"],
        }
    )
    generator.result_df = pd.DataFrame(
        [[ "SKU-1", "品名", False, "1234567890", "1PCS", "1", "x", "y", 10.5, 10.5 ]],
        columns=["SKU", "报关中文品名", "是否含金刚石", "商品编号", "数量及单位", "箱数", "c6", "c7", "单价", "总价"],
    )
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_invoice_data.py -k customs_declaration_mode_includes_invoice_number_in_red_suffix -v`
Expected: FAIL because the cell still contains plain `9710` or lacks the red suffix formatting.

- [ ] **Step 3: Write minimal implementation**

```python
customs_mode_value = f"9710({self.shipping_details['invoice_number']})"
```

Add rich-text assignment for the target cell after worksheet population, with a fallback to plain text if rich text APIs are unavailable.

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_invoice_data.py -k customs_declaration_mode_includes_invoice_number_in_red_suffix -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add tests/test_invoice_data.py invoice_data.py docs/superpowers/specs/2026-06-16-invoice-customs-mode-design.md docs/superpowers/plans/2026-06-16-invoice-customs-mode.md
git commit -m "feat: format invoice customs mode with invoice suffix"
```

### Task 2: Verify no regression in focused invoice tests

**Files:**
- Test: `tests/test_invoice_data.py`

- [ ] **Step 1: Run focused invoice tests**

```bash
pytest tests/test_invoice_data.py -v
```

- [ ] **Step 2: Confirm expected result**

Expected: existing invoice tests remain green, including the new customs declaration formatting test.

- [ ] **Step 3: Commit**

```bash
git add tests/test_invoice_data.py invoice_data.py
git commit -m "test: verify invoice customs declaration formatting"
```
