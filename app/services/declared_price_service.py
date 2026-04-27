import pandas as pd


def select_declared_price_columns(declared_unit_price_df: pd.DataFrame, sheet_key: str,
                                  company_info_index: int, is_awd: bool) -> pd.DataFrame:
    if company_info_index == 0 and not is_awd:
        result = declared_unit_price_df[["*SKU", "易达-欣耀\n申报单价（CNY）"]].copy()
        result.rename(columns={"易达-欣耀\n申报单价（CNY）": "申报单价"}, inplace=True)
        return result

    if sheet_key == "易达报关单价":
        price_col = next(
            (col for col in declared_unit_price_df.columns if "易达-亚马逊" in str(col) and "申报单价" in str(col)),
            "易达-亚马逊\n申报单价(USD)",
        )
        result = declared_unit_price_df[["*SKU", price_col]].copy()
        result.rename(columns={price_col: "申报单价"}, inplace=True)
        return result

    if sheet_key == "匠晟报关单价":
        result = declared_unit_price_df[["*SKU", "匠晟-亚马逊\n申报单价(USD)"]].copy()
        result.rename(columns={"匠晟-亚马逊\n申报单价(USD)": "申报单价"}, inplace=True)
        return result

    return declared_unit_price_df[["*SKU", "申报单价"]].copy()
