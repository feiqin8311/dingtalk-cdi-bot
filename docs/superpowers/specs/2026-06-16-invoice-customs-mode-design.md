# Invoice Customs Mode Cell Design

**Goal:** Update the invoice-mode customs declaration sheet so the `监管方式` value becomes `9710(<发票号>)`, with the appended invoice-number suffix displayed in red when workbook rich text support allows it.

**Scope:** This change applies only to the invoice-mode customs declaration sheet generated from `invoice_data.py`. Stock-mode behavior remains unchanged.

**Design:**
- Keep the existing report layout, row structure, and workbook generation flow unchanged.
- Replace the hard-coded `9710` value in the invoice-mode customs declaration header row with a formatted string derived from `self.shipping_details['invoice_number']`.
- Prefer a single-cell rich text rendering so only the appended `(<发票号>)` portion is red.
- If rich text assignment is unavailable in the current `openpyxl` version or workbook object, fall back to plain text `9710(<发票号>)` in the existing cell styling.

**Data Source:**
- `self.shipping_details['invoice_number']` is already prepared before `add_customs_declaration_sheet()` runs.

**Testing:**
- Add a focused unit test around invoice-mode `add_customs_declaration_sheet()`.
- Verify the target cell contains `9710(INV-XXX)`-style content.
- Verify rich text is emitted when supported, and that the invoice suffix carries a red font.
