import pandas as pd


DEYUN_SUPPLIER_NAME = "宁波德韵工具有限公司"
NINGBO_SOURCE_CODE = "33029 宁波其他"
NINGBO_SOURCE_DISPLAY = "宁波其他"
INTERNAL_INVOICE_SHEET = "2026（内部开票）"
EXTERNAL_INVOICE_SHEET = "2026（外部开票）"
INTERNAL_REGISTER_HELPER_COLUMNS = ["境内货源地", "境内货源地显示", "供应商名称", "采购员"]
EXTERNAL_INVOICE_COLUMNS = [
    "报关渠道", "合同号", "出口日期", "预录单编号", "报关方式（9710/一般贸易）", "品牌", "SKU",
    "品名", "HS编码", "数量", "单位", "单价", "总价", "币别", "供应商名称", "采购员",
    "采购合同号", "发票号", "开票日期", "数量", "单价", "不含税金额", "税额", "申报期",
    "入库时间", "入库数量", "入库金额", "抵扣认证",
]


def source_display(source_code: str) -> str:
    value = str(source_code or "").strip()
    if " " in value:
        return value.split(" ", 1)[1]
    return value


def apply_lumiya_source_columns(df: pd.DataFrame, company_info_index: int,
                                default_source_code: str) -> pd.DataFrame:
    result = df.copy()
    default_display = source_display(default_source_code)
    result["境内货源地"] = default_source_code
    result["境内货源地显示"] = default_display

    if company_info_index != 1 or "供应商名称" not in result.columns:
        return result

    supplier = result["供应商名称"].fillna("").astype(str).str.strip()
    deyun_mask = supplier.eq(DEYUN_SUPPLIER_NAME)
    result.loc[deyun_mask, "境内货源地"] = NINGBO_SOURCE_CODE
    result.loc[deyun_mask, "境内货源地显示"] = NINGBO_SOURCE_DISPLAY
    return result


def update_dominant_source_display(shipping_details: dict, result_df: pd.DataFrame) -> None:
    default_source = shipping_details.get("domestic_source_of_goods", "")
    default_display = source_display(default_source)

    if "境内货源地显示" not in result_df.columns or "总价" not in result_df.columns:
        shipping_details["domestic_source_display"] = default_display
        return

    totals = (
        result_df
        .assign(总价=pd.to_numeric(result_df["总价"], errors="coerce").fillna(0))
        .groupby("境内货源地显示", dropna=False)["总价"]
        .sum()
    )
    if totals.empty:
        shipping_details["domestic_source_display"] = default_display
        return

    max_total = totals.max()
    winners = totals[totals == max_total].index.astype(str).tolist()
    if len(winners) == 1:
        shipping_details["domestic_source_display"] = winners[0]
    else:
        shipping_details["domestic_source_display"] = default_display


def is_lumiya_external_source(df: pd.DataFrame, company_info_index: int) -> pd.Series:
    if company_info_index != 1 or "境内货源地" not in df.columns:
        return pd.Series(False, index=df.index)
    return df["境内货源地"].fillna("").astype(str).str.strip().eq(NINGBO_SOURCE_CODE)


def attach_lumiya_register_source(register_df: pd.DataFrame, result_df: pd.DataFrame,
                                  company_info_index: int) -> pd.DataFrame:
    if company_info_index != 1:
        return register_df.copy()

    result = register_df.reset_index(drop=True).copy()
    source_df = result_df.reset_index(drop=True)
    if all(col in result.columns for col in INTERNAL_REGISTER_HELPER_COLUMNS):
        return result

    if "SKU" in result.columns and "SKU" in source_df.columns:
        source_by_sku = source_df.drop_duplicates("SKU").set_index("SKU")
        for col in INTERNAL_REGISTER_HELPER_COLUMNS:
            if col in result.columns:
                continue
            if col in source_by_sku.columns:
                result[col] = result["SKU"].map(source_by_sku[col]).fillna("")
            else:
                result[col] = ""
        return result

    for col in INTERNAL_REGISTER_HELPER_COLUMNS:
        if col in source_df.columns:
            result[col] = source_df[col].reindex(result.index).fillna("")
        elif col not in result.columns:
            result[col] = ""
    return result


def drop_internal_register_helper_columns(register_df: pd.DataFrame) -> pd.DataFrame:
    return register_df.drop(
        columns=[col for col in INTERNAL_REGISTER_HELPER_COLUMNS if col in register_df.columns]
    )


def format_lumiya_external_register_df(register_df: pd.DataFrame, contract_column: str) -> pd.DataFrame:
    rows = []
    for _, row in register_df.iterrows():
        rows.append([
            row.get("报关渠道", ""),
            row.get(contract_column, ""),
            row.get("出口日期", ""),
            row.get("预录单编号", ""),
            row.get("报关方式（9710/一般贸易）", ""),
            row.get("店铺", ""),
            row.get("SKU", ""),
            row.get("开票/报关品名", row.get("报关中文品名", "")),
            row.get("HS CODE", ""),
            row.get("数量", ""),
            row.get("单位", ""),
            row.get("申报单价", ""),
            row.get("总价", ""),
            row.get("币别", ""),
            row.get("供应商名称", DEYUN_SUPPLIER_NAME) or DEYUN_SUPPLIER_NAME,
            row.get("采购员", ""),
            "", "", "", "", "", "", "", "", "", "", "", "",
        ])
    return pd.DataFrame(rows, columns=EXTERNAL_INVOICE_COLUMNS)


def split_lumiya_register_entries(register_df: pd.DataFrame, company_info_index: int,
                                  register_filepath: str, contract_column: str,
                                  external_df: pd.DataFrame | None = None,
                                  internal_df: pd.DataFrame | None = None) -> list[tuple[pd.DataFrame, str, str]]:
    if company_info_index != 1:
        return [(register_df, register_filepath, INTERNAL_INVOICE_SHEET)]

    external_mask = is_lumiya_external_source(register_df, company_info_index)
    entries = []

    register_internal_df = internal_df.copy() if internal_df is not None else register_df.loc[~external_mask].copy()
    if not register_internal_df.empty:
        entries.append((
            drop_internal_register_helper_columns(register_internal_df).reset_index(drop=True),
            register_filepath,
            INTERNAL_INVOICE_SHEET,
        ))

    register_external_df = external_df.copy() if external_df is not None else register_df.loc[external_mask].copy()
    if not register_external_df.empty:
        entries.append((
            format_lumiya_external_register_df(register_external_df, contract_column).reset_index(drop=True),
            register_filepath,
            EXTERNAL_INVOICE_SHEET,
        ))

    return entries


def build_lumiya_internal_declaration_register(register_df: pd.DataFrame, result_df: pd.DataFrame,
                                               company_info_index: int, quantity_column: str,
                                               currency: str) -> pd.DataFrame:
    if company_info_index != 1:
        return register_df.copy()

    internal_register_df = register_df[
        ~is_lumiya_external_source(register_df, company_info_index)
    ].copy().reset_index(drop=True)
    if result_df is None or result_df.empty or internal_register_df.empty:
        return internal_register_df

    internal_declaration_df = result_df[
        ~is_lumiya_external_source(result_df, company_info_index)
    ].copy().reset_index(drop=True)
    if internal_declaration_df.empty:
        return internal_register_df

    target_columns = ["开票/报关品名", "英文品名", "HS CODE", "数量", "单位", "申报单价", "总价", "币别"]
    for col in target_columns:
        internal_register_df[col] = pd.Series([None] * len(internal_register_df), dtype=object)

    fill_count = min(len(internal_register_df), len(internal_declaration_df))
    declaration_values = pd.DataFrame({
        "开票/报关品名": internal_declaration_df["报关中文品名"],
        "英文品名": internal_declaration_df["英文品名"],
        "HS CODE": internal_declaration_df["HS CODE"],
        "数量": internal_declaration_df[quantity_column],
        "单位": internal_declaration_df["单位"],
        "申报单价": internal_declaration_df["申报单价"],
        "总价": internal_declaration_df["总价"],
        "币别": currency,
    })
    internal_register_df.loc[:fill_count - 1, target_columns] = declaration_values.loc[:fill_count - 1, target_columns].values
    return internal_register_df


def build_lumiya_external_declaration_register(register_df: pd.DataFrame, result_df: pd.DataFrame,
                                               company_info_index: int, contract_column: str,
                                               quantity_column: str, currency: str) -> pd.DataFrame:
    if company_info_index != 1 or result_df is None or result_df.empty:
        return pd.DataFrame()

    external_declaration_df = result_df[
        is_lumiya_external_source(result_df, company_info_index)
    ].copy().reset_index(drop=True)
    if external_declaration_df.empty:
        return external_declaration_df

    contract_value = ""
    if contract_column in register_df.columns:
        contract_values = register_df[contract_column].dropna().astype(str).str.strip()
        if not contract_values.empty:
            contract_value = contract_values.iloc[0]

    external_declaration_df["报关渠道"] = (
        register_df["报关渠道"].dropna().iloc[0] if "报关渠道" in register_df.columns and not register_df["报关渠道"].dropna().empty else ""
    )
    external_declaration_df[contract_column] = contract_value
    if len(external_declaration_df) > 1:
        external_declaration_df.loc[external_declaration_df.index[1:], contract_column] = ""
    external_declaration_df["出口日期"] = ""
    external_declaration_df["预录单编号"] = ""
    external_declaration_df["报关方式（9710/一般贸易）"] = (
        register_df["报关方式（9710/一般贸易）"].dropna().iloc[0]
        if "报关方式（9710/一般贸易）" in register_df.columns and not register_df["报关方式（9710/一般贸易）"].dropna().empty
        else ""
    )
    external_declaration_df["店铺"] = (
        register_df["店铺"].dropna().iloc[0] if "店铺" in register_df.columns and not register_df["店铺"].dropna().empty else ""
    )
    external_declaration_df["开票/报关品名"] = external_declaration_df["报关中文品名"]
    external_declaration_df["数量"] = pd.to_numeric(external_declaration_df[quantity_column], errors="coerce")
    external_declaration_df["总价"] = pd.to_numeric(external_declaration_df["总价"], errors="coerce").round(2)
    external_declaration_df["申报单价"] = (
        external_declaration_df["总价"] / external_declaration_df["数量"]
    ).round(2)
    external_declaration_df["币别"] = currency
    external_declaration_df["供应商名称"] = DEYUN_SUPPLIER_NAME
    external_declaration_df["采购员"] = ""
    return external_declaration_df
