import re

import pandas as pd


def find_data_group(elements_file_path: str, result_df: pd.DataFrame, cache: dict,
                    sheet_name_transform=None) -> list[list[str]]:
    result_groups = []
    name_pattern = re.compile(r'品名[；;：:+\-=\s]*([^。.;；]+)')
    sheet_col = result_df.columns[-1]
    col_idx = {col: i for i, col in enumerate(result_df.columns)}
    required_cols = ['SKU', '报关中文品名', 'HS CODE', '是否含金刚石']
    missing_cols = [col for col in required_cols if col not in col_idx]
    if missing_cols:
        raise ValueError(f"要素匹配缺少必要字段：{', '.join(missing_cols)}")

    for data in result_df.itertuples(index=False, name=None):
        raw_sheet_name = data[col_idx[sheet_col]]
        if pd.isna(raw_sheet_name):
            raise ValueError("要素表匹配失败：店铺/Sheet 名称为空，请检查数据！")
        sheet_name = sheet_name_transform(raw_sheet_name) if sheet_name_transform else raw_sheet_name
        cache_key = (elements_file_path, sheet_name)
        cache_entry = cache.get(cache_key)
        if cache_entry is None:
            elements_df = pd.read_excel(elements_file_path, sheet_name=sheet_name, header=None)
            elements_data_list = [str(x) if pd.notna(x) else '' for x in elements_df.iloc[:, 0].tolist()]
            groups = []
            current_group = []
            for value in elements_data_list:
                if value.strip() == '':
                    if current_group:
                        groups.append(current_group)
                        current_group = []
                else:
                    current_group.append(value)
            if current_group:
                groups.append(current_group)
            prepared_groups = []
            for group in groups:
                group_text = ' '.join(group)
                prepared_groups.append((group, group_text, "金刚石" in group_text))
            cache_entry = {'prepared_groups': prepared_groups, 'index': {}}
            cache[cache_key] = cache_entry

        prepared_groups = cache_entry['prepared_groups']
        index_cache = cache_entry.setdefault('index', {})
        sku = data[col_idx['SKU']]
        target_name = str(data[col_idx['报关中文品名']]).strip()
        target_hs_code = data[col_idx['HS CODE']]
        contains_diamond = bool(data[col_idx['是否含金刚石']])

        if pd.isna(target_hs_code):
            raise ValueError(f"报关中文品名：{target_name}，HS CODE为空，请去品名表.xlsx，请添加此报关中文品名！")

        target_hs_code = str(target_hs_code).strip()
        cache_key_detail = (target_hs_code, target_name, contains_diamond)
        if cache_key_detail not in index_cache:
            found_group = None
            for group, group_text, group_contains in prepared_groups:
                if target_hs_code not in group_text or group_contains != contains_diamond:
                    continue
                for temp_date in group:
                    match = name_pattern.search(temp_date)
                    if not match:
                        continue
                    if match.group(1).strip() == target_name:
                        found_group = list(group)
                        break
                if found_group:
                    break
            if found_group is None:
                raise ValueError(f"HS CODE：{target_hs_code}，报关中文品名：{target_name}，请去要素表.xlsx的{sheet_name}页添加此要素！")
            index_cache[cache_key_detail] = found_group

        found_group = list(index_cache[cache_key_detail])
        found_group.append(sku)
        if '单位' in col_idx and '总PCS' in col_idx and data[col_idx['单位']] == '套':
            sets = f"{data[col_idx.get('备货数量', col_idx.get('发货数量'))]}{data[col_idx['单位']]}={data[col_idx['总PCS']]}个"
            found_group.append(sets)
        result_groups.append(found_group)

    return result_groups
