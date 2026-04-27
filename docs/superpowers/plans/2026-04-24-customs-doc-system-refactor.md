# Customs Document System Refactor Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reorganize the customs document system so shared logic is implemented once, invoice and stock workflows stay separate, and future rule changes only touch focused modules.

**Architecture:** Introduce `app/core`, `app/services`, `app/renderers`, and `app/flows` as explicit boundaries. Migrate common infrastructure first, then extract shared business services, then move sheet rendering out of the giant classes, and finally reduce `invoice_data.py` and `stock_preparation.py` to compatibility wrappers over the new flows.

**Tech Stack:** Python 3, `openpyxl`, `pandas`, `numpy`, `unittest`, optional `win32com`

---

## File Structure

**Create**
- `app/__init__.py`
- `app/core/__init__.py`
- `app/core/logging.py`
- `app/core/metrics.py`
- `app/core/excel_io.py`
- `app/core/page_setup.py`
- `app/core/template_files.py`
- `app/core/models.py`
- `app/services/__init__.py`
- `app/services/declared_price_service.py`
- `app/services/freight_service.py`
- `app/services/elements_service.py`
- `app/services/register_service.py`
- `app/renderers/__init__.py`
- `app/renderers/common.py`
- `app/renderers/detailed_list.py`
- `app/renderers/packing_list.py`
- `app/renderers/bill.py`
- `app/renderers/contract.py`
- `app/renderers/customs_declaration.py`
- `app/renderers/elements_sheet.py`
- `app/flows/__init__.py`
- `app/flows/invoice_flow.py`
- `app/flows/stock_flow.py`
- `tests/test_core_page_setup.py`
- `tests/test_core_excel_io.py`
- `tests/test_services_declared_price.py`
- `tests/test_services_register.py`

**Modify**
- `invoice_data.py`
- `stock_preparation.py`
- `ui_pyqt.py`
- `tests/test_invoice_data.py`
- `tests/test_stock_preparation.py`

## Task 1: Extract Core Infrastructure

**Files:**
- Create: `app/__init__.py`
- Create: `app/core/__init__.py`
- Create: `app/core/logging.py`
- Create: `app/core/metrics.py`
- Create: `app/core/excel_io.py`
- Create: `app/core/page_setup.py`
- Create: `app/core/template_files.py`
- Create: `app/core/models.py`
- Test: `tests/test_core_page_setup.py`
- Test: `tests/test_core_excel_io.py`
- Modify: `invoice_data.py`
- Modify: `stock_preparation.py`

- [ ] **Step 1: Write the failing tests for page setup persistence and cached Excel reads**

```python
# tests/test_core_page_setup.py
import os
import tempfile
import unittest
from openpyxl import Workbook, load_workbook

from app.core.page_setup import apply_persistent_page_setup, set_excel_page_view


class CorePageSetupTests(unittest.TestCase):
    def test_apply_persistent_page_setup_survives_save(self):
        wb = Workbook()
        ws = wb.active
        apply_persistent_page_setup(ws)
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            path = tmp.name
        try:
            wb.save(path)
            reloaded = load_workbook(path)
            ws2 = reloaded.active
            self.assertEqual(ws2.sheet_view.view, "pageBreakPreview")
            self.assertEqual(ws2.page_setup.fitToWidth, 1)
            self.assertEqual(ws2.page_setup.fitToHeight, 1)
        finally:
            if "reloaded" in locals():
                reloaded.close()
            os.remove(path)

    def test_set_excel_page_view_is_safe_without_com(self):
        wb = Workbook()
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            path = tmp.name
        try:
            wb.save(path)
            wb.close()
            set_excel_page_view(path)
            reloaded = load_workbook(path)
            ws2 = reloaded.active
            self.assertEqual(ws2.sheet_view.view, "pageBreakPreview")
        finally:
            if "reloaded" in locals():
                reloaded.close()
            os.remove(path)
```

```python
# tests/test_core_excel_io.py
import os
import tempfile
import unittest
import pandas as pd

from app.core.excel_io import ExcelCache


class ExcelCacheTests(unittest.TestCase):
    def test_read_excel_cached_returns_copy(self):
        source = pd.DataFrame({"SKU": ["A1"], "数量": [2]})
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            path = tmp.name
        try:
            source.to_excel(path, index=False, sheet_name="数据")
            cache = ExcelCache()
            df1 = cache.read(path, "数据")
            df2 = cache.read(path, "数据")
            df1.loc[0, "数量"] = 99
            self.assertEqual(df2.loc[0, "数量"], 2)
        finally:
            os.remove(path)
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```bash
python -m unittest discover -s tests -p 'test_core_*.py'
```

Expected:
- `ModuleNotFoundError: No module named 'app'`
- or `ImportError` for missing `app.core.page_setup` / `app.core.excel_io`

- [ ] **Step 3: Create the core modules with minimal working implementations**

```python
# app/core/models.py
from dataclasses import dataclass, field
from typing import Any


@dataclass
class ShippingContext:
    data: dict[str, Any] = field(default_factory=dict)


@dataclass
class RegisterTarget:
    filepath: str
    sheet_name: str
```

```python
# app/core/logging.py
import datetime
import os
import sys


def get_log_dir(module_filename: str) -> str:
    base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(module_filename))
    return os.path.join(base_dir, "Logs")


def log_message(log_file_path: str, message: str) -> None:
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_line = f"[{timestamp}] {message}"
    print(log_line)
    try:
        with open(log_file_path, "a", encoding="utf-8") as f:
            f.write(log_line + "\n")
    except Exception as exc:
        print(f"写入日志文件失败: {exc}, 路径: {log_file_path}")
```

```python
# app/core/metrics.py
import datetime
import os
import openpyxl
from openpyxl import load_workbook


def append_call_record(metrics_path: str, sheet_name: str, headers: list[str], module_name: str,
                       action: str, success: bool, duration_seconds: float, invoice_number: str = "",
                       source_file: str = "", error_message: str = "") -> None:
    try:
        if metrics_path and os.path.exists(metrics_path):
            wb = load_workbook(metrics_path)
            ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = sheet_name
        if ws.max_row == 1 and ws.cell(row=1, column=1).value is None:
            for idx, header in enumerate(headers, start=1):
                ws.cell(row=1, column=idx, value=header)
        ws.append([
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            module_name,
            action,
            "成功" if success else "失败",
            round(duration_seconds, 3),
            invoice_number,
            source_file,
            error_message,
        ])
        wb.save(metrics_path)
    except Exception as exc:
        print(f"写入调用统计失败: {exc}, 路径: {metrics_path}")
```

```python
# app/core/excel_io.py
from __future__ import annotations

from typing import Any
import pandas as pd


class ExcelCache:
    def __init__(self) -> None:
        self._cache: dict[tuple[Any, ...], pd.DataFrame] = {}

    @staticmethod
    def _freeze(value: Any) -> Any:
        if isinstance(value, list):
            return tuple(value)
        if isinstance(value, dict):
            return tuple(sorted((k, ExcelCache._freeze(v)) for k, v in value.items()))
        return value

    def read(self, file_path: str, sheet_name: str, **kwargs) -> pd.DataFrame:
        cache_key = (file_path, sheet_name, tuple(sorted((k, self._freeze(v)) for k, v in kwargs.items())))
        cached = self._cache.get(cache_key)
        if cached is None:
            cached = pd.read_excel(file_path, sheet_name=sheet_name, **kwargs)
            self._cache[cache_key] = cached
        return cached.copy()
```

```python
# app/core/page_setup.py
from __future__ import annotations

from openpyxl import load_workbook

try:
    import win32com.client as win32
except Exception:
    win32 = None


def apply_persistent_page_setup(ws) -> None:
    ws.sheet_view.view = "pageBreakPreview"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1


def set_excel_page_view(file_path: str) -> None:
    workbook = load_workbook(file_path)
    for sheet in workbook.worksheets:
        apply_persistent_page_setup(sheet)
    workbook.save(file_path)
    workbook.close()
    if win32 is None:
        return
    app = None
    for prog_id in ("Ket.Application", "WPS.Application", "Excel.Application"):
        try:
            app = win32.Dispatch(prog_id)
            break
        except Exception:
            continue
    if app is None or not hasattr(app, "Workbooks"):
        return
    try:
        workbook = app.Workbooks.Open(file_path)
        for sheet in workbook.Worksheets:
            sheet.Activate()
            page_setup = sheet.PageSetup
            page_setup.Zoom = False
            page_setup.FitToPagesWide = 1
            page_setup.FitToPagesTall = 1
        workbook.Save()
        workbook.Close()
    except Exception:
        return
```

```python
# app/core/template_files.py
import os
import shutil


def copy_template_file(template_path: str, output_dir: str, output_name: str) -> str:
    output_path = os.path.join(output_dir, output_name)
    shutil.copy(template_path, output_path)
    return output_path
```

- [ ] **Step 4: Wire the two legacy modules to the new core helpers**

```python
# invoice_data.py / stock_preparation.py
from app.core.excel_io import ExcelCache
from app.core.page_setup import apply_persistent_page_setup, set_excel_page_view
from app.core.logging import get_log_dir, log_message as core_log_message
```

```python
# inside CustomsDeclaration.__init__
self._excel_cache = ExcelCache()
```

```python
# replace method body
def _read_excel_cached(self, file_path: str, sheet_name: str, **kwargs) -> pd.DataFrame:
    return self._excel_cache.read(file_path, sheet_name, **kwargs)
```

```python
# replace page setup wrappers
@staticmethod
def apply_persistent_page_setup(ws) -> None:
    apply_persistent_page_setup(ws)

@classmethod
def set_excel_page_view(cls, file_path: str) -> None:
    set_excel_page_view(file_path)
```

- [ ] **Step 5: Run the new core tests and existing page setup regression tests**

Run:
```bash
python -m unittest discover -s tests -p 'test_core_*.py'
python -m unittest discover -s tests -p 'test_*data.py'
python -m unittest discover -s tests -p 'test_stock_preparation.py'
```

Expected:
- All tests pass

- [ ] **Step 6: Commit**

```bash
git add app/core app/__init__.py tests/test_core_page_setup.py tests/test_core_excel_io.py invoice_data.py stock_preparation.py
git commit -m "refactor: extract shared core infrastructure"
```

## Task 2: Extract Shared Business Services

**Files:**
- Create: `app/services/__init__.py`
- Create: `app/services/declared_price_service.py`
- Create: `app/services/freight_service.py`
- Create: `app/services/elements_service.py`
- Create: `app/services/register_service.py`
- Test: `tests/test_services_declared_price.py`
- Test: `tests/test_services_register.py`
- Modify: `invoice_data.py`
- Modify: `stock_preparation.py`

- [ ] **Step 1: Write failing tests for shared service behavior**

```python
# tests/test_services_declared_price.py
import unittest
import pandas as pd

from app.services.declared_price_service import select_declared_price_columns


class DeclaredPriceServiceTests(unittest.TestCase):
    def test_easycan_non_awd_uses_cny_column(self):
        source = pd.DataFrame({
            "*SKU": ["A1"],
            "易达-欣耀\n申报单价（CNY）": [12.5],
            "易达-亚马逊\n申报单价(USD)": [3.5],
        })
        result = select_declared_price_columns(
            source,
            sheet_key="易达报关单价",
            company_info_index=0,
            is_awd=False,
        )
        self.assertEqual(result.loc[0, "申报单价"], 12.5)
```

```python
# tests/test_services_register.py
import unittest
from openpyxl import Workbook

from app.services.register_service import check_invoice_number_in_sheet


class RegisterServiceTests(unittest.TestCase):
    def test_check_invoice_number_raises_when_duplicate_exists(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["发票号"])
        ws.append(["INV-001"])
        with self.assertRaises(Exception):
            check_invoice_number_in_sheet(ws, "INV-001", "发票号")
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```bash
python -m unittest discover -s tests -p 'test_services_*.py'
```

Expected:
- Import errors for missing service modules

- [ ] **Step 3: Implement the services**

```python
# app/services/declared_price_service.py
import pandas as pd


def select_declared_price_columns(declared_unit_price_df: pd.DataFrame, sheet_key: str,
                                  company_info_index: int, is_awd: bool) -> pd.DataFrame:
    if company_info_index == 0 and not is_awd:
        result = declared_unit_price_df[["*SKU", "易达-欣耀\n申报单价（CNY）"]].copy()
        result.rename(columns={"易达-欣耀\n申报单价（CNY）": "申报单价"}, inplace=True)
        return result
    if sheet_key == "易达报关单价":
        price_col = next((col for col in declared_unit_price_df.columns if "易达-亚马逊" in str(col) and "申报单价" in str(col)), "易达-亚马逊\n申报单价(USD)")
        result = declared_unit_price_df[["*SKU", price_col]].copy()
        result.rename(columns={price_col: "申报单价"}, inplace=True)
        return result
    if sheet_key == "匠晟报关单价":
        result = declared_unit_price_df[["*SKU", "匠晟-亚马逊\n申报单价(USD)"]].copy()
        result.rename(columns={"匠晟-亚马逊\n申报单价(USD)": "申报单价"}, inplace=True)
        return result
    return declared_unit_price_df[["*SKU", "申报单价"]].copy()
```

```python
# app/services/freight_service.py
def calculate_shipping_costs(transportation_method: str, logistics_provider: str, country_zh: str,
                             currency: str, total_volume: float, total_weight: float) -> tuple[float, str]:
    exchange_rate = 7.2 if currency == "USD" else 1.0
    multiplier_num1 = 1250
    multiplier_num2 = 1
    freight_basis = "volume"
    if "海卡" in transportation_method:
        sea_card_rates = {
            "加拿大": 800,
            "日本": 800,
            "德国": 732 if logistics_provider == "北京世纪卓越快递服务有限公司（龙舟）" else 1200,
            "英国": 634 if logistics_provider == "北京世纪卓越快递服务有限公司（龙舟）" else 1200,
        }
        multiplier_num1 = sea_card_rates.get(country_zh, 1250)
        multiplier_num2 = max(1.0, total_volume)
    elif "海派" in transportation_method:
        multiplier_num1 = {"美国": 12, "加拿大": 12, "澳洲": 10, "日本": 10, "英国": 11}.get(country_zh, 12)
        multiplier_num2 = max(12.0, total_weight)
        freight_basis = "weight"
    elif "空派" in transportation_method:
        multiplier_num1 = {"美国": 45, "德国": 42}.get(country_zh, 45)
        multiplier_num2 = max(12.0, total_weight)
        freight_basis = "weight"
    elif "卡航" in transportation_method:
        multiplier_num1 = {"英国": 15.5, "德国": 16}.get(country_zh, 15.5)
        multiplier_num2 = total_weight
        freight_basis = "weight"
    elif "美国海运" in transportation_method:
        multiplier_num1 = 500
        multiplier_num2 = max(1.0, total_volume)
    return round(multiplier_num1 * multiplier_num2 / exchange_rate, 2), freight_basis
```

```python
# app/services/elements_service.py
import re
import pandas as pd


def find_data_group(elements_file_path: str, result_df: pd.DataFrame, read_excel_fn, cache: dict) -> list[list[str]]:
    name_pattern = re.compile(r"品名[；;：:+\\-=\\s]*([^。.;；]+)")
    result_groups = []
    sheet_col = result_df.columns[-1]
    col_idx = {col: i for i, col in enumerate(result_df.columns)}
    for data in result_df.itertuples(index=False, name=None):
        sheet_name = data[col_idx[sheet_col]]
        cache_key = (elements_file_path, sheet_name)
        if cache_key not in cache:
            elements_df = read_excel_fn(elements_file_path, sheet_name=sheet_name, header=None)
            values = [str(x) if pd.notna(x) else "" for x in elements_df.iloc[:, 0].tolist()]
            groups = []
            current = []
            for value in values:
                if value.strip():
                    current.append(value)
                elif current:
                    groups.append(current)
                    current = []
            if current:
                groups.append(current)
            cache[cache_key] = groups
        groups = cache[cache_key]
        target_name = str(data[col_idx["报关中文品名"]]).strip()
        target_hs_code = str(data[col_idx["HS CODE"]]).strip()
        for group in groups:
            group_text = " ".join(group)
            if target_hs_code not in group_text:
                continue
            for row in group:
                match = name_pattern.search(row)
                if match and match.group(1).strip() == target_name:
                    result_groups.append(list(group) + [data[col_idx["SKU"]]])
                    break
            else:
                continue
            break
    return result_groups
```

```python
# app/services/register_service.py
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def check_invoice_number_in_sheet(ws, invoice_number: str, column_name: str) -> int:
    col_idx = None
    for cell in ws[1]:
        if str(cell.value).strip() == column_name:
            col_idx = cell.column
            break
    if col_idx is None:
        raise ValueError(f"登记表未找到列：{column_name}")
    invoice_number = str(invoice_number).strip()
    if not invoice_number:
        return 0
    for row_idx in range(2, ws.max_row + 1):
        value = ws.cell(row=row_idx, column=col_idx).value
        value_str = str(value) if value is not None else ""
        if value_str == invoice_number or invoice_number in value_str:
            raise Exception("发票号已登记")
    return 0


def append_dataframe_rows(worksheet, df):
    rows = list(dataframe_to_rows(df, index=False, header=False))
    if not rows:
        raise ValueError("登记表生成的数据为空，未写入任何行")
    for row in rows:
        worksheet.append(row)
    return len(rows)
```

- [ ] **Step 4: Replace duplicated service logic in both legacy modules**

```python
# imports
from app.services.declared_price_service import select_declared_price_columns
from app.services.freight_service import calculate_shipping_costs
from app.services.elements_service import find_data_group as service_find_data_group
from app.services.register_service import check_invoice_number_in_sheet as service_check_invoice_number_in_sheet
```

```python
# process_declared_unit_price
declared_unit_price_df = self._read_excel_cached(product_filepath, sheet_name=self.shipping_details["customs_declaration_price"])
declared_unit_price_df = select_declared_price_columns(
    declared_unit_price_df,
    sheet_key=self.shipping_details["customs_declaration_price"],
    company_info_index=self.company_info_index,
    is_awd=self.shipping_details.get("is_awd", False),
).drop_duplicates()
declared_unit_price_df["*SKU"] = declared_unit_price_df["*SKU"].astype(str)
return declared_unit_price_df.copy()
```

```python
# calculate_shipping_costs wrapper
freight, freight_basis = calculate_shipping_costs(
    self.shipping_details["transportation_method"],
    self.shipping_details["logistics_provider"],
    self.shipping_details["country_zh"],
    self.shipping_details.get("currency", "USD"),
    total_volume,
    total_weight,
)
self.shipping_details["freight"] = freight
self.shipping_details["freight_basis"] = freight_basis
```

- [ ] **Step 5: Run service tests and existing regression tests**

Run:
```bash
python -m unittest discover -s tests -p 'test_services_*.py'
python -m unittest discover -s tests -p 'test_*.py'
```

Expected:
- All tests pass

- [ ] **Step 6: Commit**

```bash
git add app/services tests/test_services_declared_price.py tests/test_services_register.py invoice_data.py stock_preparation.py
git commit -m "refactor: extract shared business services"
```

## Task 3: Extract Shared Renderer Utilities and Sheet Renderers

**Files:**
- Create: `app/renderers/__init__.py`
- Create: `app/renderers/common.py`
- Create: `app/renderers/detailed_list.py`
- Create: `app/renderers/packing_list.py`
- Create: `app/renderers/bill.py`
- Create: `app/renderers/contract.py`
- Create: `app/renderers/customs_declaration.py`
- Create: `app/renderers/elements_sheet.py`
- Modify: `invoice_data.py`
- Modify: `stock_preparation.py`

- [ ] **Step 1: Write a failing renderer smoke test**

```python
# add to tests/test_stock_preparation.py
import unittest
from openpyxl import Workbook
from app.renderers.common import ensure_default_sheet_removed


class RendererCommonTests(unittest.TestCase):
    def test_ensure_default_sheet_removed(self):
        wb = Workbook()
        wb.create_sheet("明细单")
        ensure_default_sheet_removed(wb)
        self.assertNotIn("Sheet", wb.sheetnames)
```

- [ ] **Step 2: Run the renderer smoke test to verify it fails**

Run:
```bash
python -m unittest tests.test_stock_preparation.RendererCommonTests
```

Expected:
- Import error for `app.renderers.common`

- [ ] **Step 3: Implement the common renderer helpers**

```python
# app/renderers/common.py
from app.core.page_setup import apply_persistent_page_setup


def ensure_default_sheet_removed(workbook) -> None:
    if "Sheet" in workbook.sheetnames and len(workbook.sheetnames) > 1:
        del workbook["Sheet"]


def finalize_workbook_page_setup(workbook) -> None:
    for sheet in workbook.worksheets:
        apply_persistent_page_setup(sheet)
```

- [ ] **Step 4: Move one renderer at a time, starting with the lowest-risk sheets**

```python
# app/renderers/detailed_list.py
def render_detailed_list(generator, excel_file_path: str, image_path: str, wb=None):
    return generator.add_detailed_list_sheet(excel_file_path, image_path, wb=wb)


# app/renderers/packing_list.py
def render_packing_list(generator, excel_file_path: str, image_path: str, wb=None):
    return generator.add_packing_list_sheet(excel_file_path, image_path, wb=wb)


# app/renderers/bill.py
def render_bill(generator, excel_file_path: str, image_path: str, wb=None, **kwargs):
    return generator.add_bill_sheet(excel_file_path, image_path, wb=wb, **kwargs)
```

```python
# continue same pattern for contract/customs_declaration/elements_sheet
```

- [ ] **Step 5: Replace direct `add_*_sheet` orchestration with renderer module calls**

```python
# invoice_data.py / stock_preparation.py
from app.renderers.common import ensure_default_sheet_removed, finalize_workbook_page_setup
from app.renderers.detailed_list import render_detailed_list
from app.renderers.packing_list import render_packing_list
from app.renderers.bill import render_bill
```

```python
wb = openpyxl.Workbook()
render_detailed_list(self, output_filepath, pictures_path, wb=wb)
ensure_default_sheet_removed(wb)
render_packing_list(self, output_filepath, pictures_path, wb=wb)
render_bill(self, output_filepath, pictures_path, wb=wb)
finalize_workbook_page_setup(wb)
wb.save(output_filepath)
```

- [ ] **Step 6: Run the full regression suite**

Run:
```bash
python -m unittest discover -s tests -p 'test_*.py'
```

Expected:
- All tests pass

- [ ] **Step 7: Commit**

```bash
git add app/renderers invoice_data.py stock_preparation.py tests/test_stock_preparation.py
git commit -m "refactor: extract renderer orchestration"
```

## Task 4: Introduce Invoice and Stock Flows

**Files:**
- Create: `app/flows/__init__.py`
- Create: `app/flows/invoice_flow.py`
- Create: `app/flows/stock_flow.py`
- Modify: `invoice_data.py`
- Modify: `stock_preparation.py`

- [ ] **Step 1: Write a failing flow-level smoke test**

```python
# add to tests/test_invoice_data.py
import unittest
from app.flows.invoice_flow import InvoiceFlow


class InvoiceFlowImportTests(unittest.TestCase):
    def test_invoice_flow_is_importable(self):
        self.assertIsNotNone(InvoiceFlow)
```

- [ ] **Step 2: Run the smoke test to verify it fails**

Run:
```bash
python -m unittest tests.test_invoice_data.InvoiceFlowImportTests
```

Expected:
- Import error for `app.flows.invoice_flow`

- [ ] **Step 3: Create the flow wrappers**

```python
# app/flows/invoice_flow.py
class InvoiceFlow:
    def __init__(self, generator):
        self.generator = generator

    def generate_all_documents(self, invoice_filepath, product_filepath, product_name_filepath,
                               site_address_filepath, elements_filepath, output_filepath,
                               pictures_path, flag, generate_declaration_file=True):
        return self.generator.generate_all_documents(
            invoice_filepath, product_filepath, product_name_filepath,
            site_address_filepath, elements_filepath, output_filepath,
            pictures_path, flag, generate_declaration_file=generate_declaration_file,
        )

    def register_all_documents(self, *args, **kwargs):
        return self.generator.register_all_documents(*args, **kwargs)
```

```python
# app/flows/stock_flow.py
class StockFlow:
    def __init__(self, generator):
        self.generator = generator

    def generate_all_documents(self, invoice_filepath, product_filepath, product_name_filepath,
                               site_address_filepath, elements_filepath, output_filepath,
                               pictures_path, flag, generate_declaration_file=True):
        return self.generator.generate_all_documents(
            invoice_filepath, product_filepath, product_name_filepath,
            site_address_filepath, elements_filepath, output_filepath,
            pictures_path, flag, generate_declaration_file=generate_declaration_file,
        )

    def register_all_documents(self, *args, **kwargs):
        return self.generator.register_all_documents(*args, **kwargs)
```

- [ ] **Step 4: Add compatibility properties inside the legacy classes**

```python
# invoice_data.py / stock_preparation.py __init__
from app.flows.invoice_flow import InvoiceFlow
from app.flows.stock_flow import StockFlow
```

```python
# invoice_data.py __init__
self.flow = InvoiceFlow(self)

# stock_preparation.py __init__
self.flow = StockFlow(self)
```

- [ ] **Step 5: Update the UI entry points to call `.flow` instead of binding directly to giant methods**

```python
# ui_pyqt.py
generator = InvoiceDataClass()
generator.flow.generate_all_documents(...)
generator.flow.register_all_documents(...)
```

- [ ] **Step 6: Run all tests**

Run:
```bash
python -m unittest discover -s tests -p 'test_*.py'
```

Expected:
- All tests pass

- [ ] **Step 7: Commit**

```bash
git add app/flows invoice_data.py stock_preparation.py ui_pyqt.py tests/test_invoice_data.py
git commit -m "refactor: introduce invoice and stock flows"
```

## Task 5: Shrink Legacy Classes into Compatibility Wrappers

**Files:**
- Modify: `invoice_data.py`
- Modify: `stock_preparation.py`

- [ ] **Step 1: Write a failing test that asserts the legacy modules still expose `CustomsDeclaration`**

```python
# add to tests/test_invoice_data.py
import unittest
from invoice_data import CustomsDeclaration as InvoiceDeclaration
from stock_preparation import CustomsDeclaration as StockDeclaration


class LegacyCompatibilityTests(unittest.TestCase):
    def test_legacy_customs_declaration_classes_still_exist(self):
        self.assertIsNotNone(InvoiceDeclaration)
        self.assertIsNotNone(StockDeclaration)
```

- [ ] **Step 2: Run the compatibility test**

Run:
```bash
python -m unittest tests.test_invoice_data.LegacyCompatibilityTests
```

Expected:
- PASS before refactor, used as guard before deleting internals

- [ ] **Step 3: Move duplicated helper code out, keep wrappers and adapter methods only**

```python
# keep in legacy modules
class CustomsDeclaration:
    def __init__(self):
        self.shipping_details = {}
        self.result_df = None
        self._excel_cache = ExcelCache()
        self.flow = InvoiceFlow(self)  # or StockFlow(self)

    def generate_all_documents(self, *args, **kwargs):
        return self.flow.generate_all_documents(*args, **kwargs)

    def register_all_documents(self, *args, **kwargs):
        return self.flow.register_all_documents(*args, **kwargs)
```

```python
# delete from legacy modules after migration
- direct duplicated page setup implementation
- duplicated excel cache implementation
- duplicated shared service implementations
```

- [ ] **Step 4: Run the full regression suite plus one manual smoke run**

Run:
```bash
python -m unittest discover -s tests -p 'test_*.py'
```

Manual smoke:
```bash
python invoice_data.py
python stock_preparation.py
```

Expected:
- Tests pass
- Scripts still start without import errors

- [ ] **Step 5: Commit**

```bash
git add invoice_data.py stock_preparation.py
git commit -m "refactor: reduce legacy declaration classes to wrappers"
```

## Task 6: Final Verification and Documentation Sync

**Files:**
- Modify: `docs/superpowers/specs/2026-04-24-refactor-customs-doc-system-design.md`
- Modify: `docs/superpowers/plans/2026-04-24-customs-doc-system-refactor.md`

- [ ] **Step 1: Verify the refactor matches the spec**

Run:
```bash
rg -n "core|services|renderers|flows" app
```

Expected:
- All four layers exist

- [ ] **Step 2: Verify test evidence**

Run:
```bash
python -m unittest discover -s tests -p 'test_*.py'
```

Expected:
- All tests pass

- [ ] **Step 3: Verify legacy entry points still exist**

Run:
```bash
python - <<'PY'
from invoice_data import CustomsDeclaration as InvoiceDeclaration
from stock_preparation import CustomsDeclaration as StockDeclaration
print(InvoiceDeclaration.__name__)
print(StockDeclaration.__name__)
PY
```

Expected:
```text
CustomsDeclaration
CustomsDeclaration
```

- [ ] **Step 4: Update the spec document with any final path adjustments introduced during implementation**

```markdown
- Replace any outdated file paths in the spec
- Keep the architectural boundaries unchanged
```

- [ ] **Step 5: Commit**

```bash
git add docs/superpowers/specs/2026-04-24-refactor-customs-doc-system-design.md docs/superpowers/plans/2026-04-24-customs-doc-system-refactor.md
git commit -m "docs: finalize refactor plan and verification notes"
```
