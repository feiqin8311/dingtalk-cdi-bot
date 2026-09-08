# Register Detail Merge Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve each packing row when one shipment-detail SKU record applies to multiple packing rows during invoice registration.

**Architecture:** Use the packing sheet as the row-granularity source and reduce duplicate shipment-detail keys to one record before merging by `SKU` and `发货单号`. This retains a valid one-to-many relationship without a Cartesian product and preserves the existing missing-detail validation.

**Tech Stack:** Python, pandas, unittest.

---

### Task 1: Cover a detail-to-multiple-packing relationship

**Files:**
- Modify: `tests/test_invoice_data.py`
- Test: `tests/test_invoice_data.py`

- [x] **Step 1: Write the failing test**

```python
def test_register_combined_invoice_reuses_one_detail_for_multiple_packing_rows(self):
    invoice_df = pd.DataFrame([
        {"发货单号": "SP260714032", "SKU": "T6HW84", "发货数量": 60},
        {"发货单号": "SP260714032", "SKU": "T6HW84", "发货数量": 18},
    ])
    invoice_details_df = pd.DataFrame([
        {"发货单号": "SP260714032", "SKU": "T6HW84", "物流商单号": "26TL066"},
    ])
```

Complete the fixture with the columns used by `register_combined_invoice()`, patch `_read_excel_cached`, and assert that both result rows retain `物流商单号 == "26TL066"`.

- [x] **Step 2: Run the test to verify it fails**

Run: `python -m unittest discover -s tests -p 'test_invoice_data.py' -v`

Expected: FAIL because the second packing row is assigned `_merge_idx == 1` and lacks a detail record.

- [x] **Step 3: Implement the minimal merge change**

```python
detail_keys = ["SKU", "发货单号"]
invoice_details_df = invoice_details_df.drop_duplicates(subset=detail_keys, keep="first")
result_df = invoice_df.merge(invoice_details_df, on=detail_keys, how="left")
```

Remove the `_merge_idx` columns and merge key. Keep the existing null-metadata validation unchanged.

- [x] **Step 4: Run the focused test and full suite**

Run: `python -m unittest discover -s tests -p 'test_invoice_data.py' -v`

Expected: PASS.

Run: `python -m unittest discover -s tests -v`

Expected: all tests pass.
