"""
File: Invoice_Data.py
Created Time: 2025-09-03
Author: KePengxiang
Comment: 对发货单进行数据处理
"""

import math
import sys
import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import pandas as pd
import numpy as np
import re
from num2words import num2words
from typing import List, Optional
import os
import shutil
from bisect import bisect_left, bisect_right
from openpyxl.drawing.image import Image
import win32com.client as win32
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook

_PC_PATTERN = re.compile(r'(\d+)(?:[Pp][Cc]|件|片)')
_ELEMENTS_GROUP_CACHE = {}
_THIN_SIDE = Side(style='thin')
_NO_SIDE = Side(style='none')
_THIN_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)
_NO_BORDER = Border(left=_NO_SIDE, right=_NO_SIDE, top=_NO_SIDE, bottom=_NO_SIDE)
_PRODUCT_INFO_COLS = [
    '*SKU', '报关中文品名', '单品毛重(g)', '外箱规格长(cm)', '外箱规格宽(cm)', '外箱规格高(cm)',
    '单箱数量(pcs)', '材质', '货源地', '供应商名称', '报关用采购成本(CNY)'
]
_INVOICE_DETAILS_COLS = ['SKU', '物流中心编码', '运输方式', '物流商单号', '店铺', '物流商']
_FONT_CACHE = {}
_ALIGNMENT_CACHE = {}
_FONT_CLS = Font
_ALIGNMENT_CLS = Alignment

# 日志文件路径（打包时落在可执行文件所在目录）
def _get_log_dir() -> str:
    base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "Logs")


_LOG_DIR = _get_log_dir()
_LOG_FILE_PATH = os.path.join(_LOG_DIR, 'invoice_data_log.txt')
_METRICS_FILE_PATH = r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\报关资料-调用统计.xlsx"
_METRICS_SHEET_NAME = '调用统计'
_METRICS_HEADERS = ['记录时间', '模块', '动作', '状态', '耗时(秒)', '发票号', '源文件', '错误信息']
_DECLARATION_TEMPLATE_MAP = {
    "EZARC": r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\申报文件模板\EA B2B订单导入模板.xlsx",
    "LIBRATON": r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\申报文件模板\LT B2B订单导入模板.xlsx",
    "PACMANN": r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\申报文件模板\LT B2B订单导入模板.xlsx",
    "SIXPRAY": r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\申报文件模板\EA B2B订单导入模板-六浪.xlsx",
    "TOLESA": r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\申报文件模板\EA B2B订单导入模板-匠晟.xlsx",
}
_DEFAULT_DECLARATION_TEMPLATE = r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\申报文件模板\YP B2B订单导入模板.xlsx"
_PLATFORM_CODE_MAP = {
    "EZARC": "3301965G22",
    "LIBRATON": "3301969CW7",
    "PACMANN": "3301969CW7",
    "SIXPRAY": "3301969GM3",
    "TOLESA": "3301969GLM",
}
_DEFAULT_PLATFORM_CODE = "3301969FDU"
_PLATFORM_NAME_MAP = {
    "EZARC": "杭州易达工具有限公司",
    "LIBRATON": "杭州鲁米亚科技有限公司",
    "PACMANN": "杭州鲁米亚科技有限公司",
    "SIXPRAY": "杭州六浪电子商务有限公司",
    "TOLESA": "杭州匠晟电子商务有限公司",
}
_DEFAULT_PLATFORM_NAME = "杭州素桐电子商务有限公司"
_TRANSPORT_METHOD_DISPLAY_MAP = {
    "卡航": "BY TRUCK",
    "海卡": "BY SEA",
    "海派": "BY SEA",
    "空派": "BY AIR",
    "美国海运": "BY SEA",
}
_CUSTOMS_TRANSPORT_METHOD_DISPLAY_MAP = {
    "卡航": "公路运输",
    "海卡": "海运",
    "海派": "海运",
    "空派": "空运",
    "美国海运": "海运",
}

# 确保日志目录存在
if not os.path.exists(_LOG_DIR):
    os.makedirs(_LOG_DIR)


def log_message(message: str) -> None:
    """同时输出日志到控制台和文件"""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_line = f"[{timestamp}] {message}"
    print(log_line)
    try:
        with open(_LOG_FILE_PATH, 'a', encoding='utf-8') as f:
            f.write(log_line + '\n')
    except Exception as e:
        print(f"写入日志文件失败: {e}, 路径: {_LOG_FILE_PATH}")


def append_call_record(module_name: str, action: str, success: bool, duration_seconds: float,
                       invoice_number: str = "", source_file: str = "", error_message: str = "") -> None:
    """将调用记录追加到Excel统计表"""
    try:
        metrics_dir = os.path.dirname(_METRICS_FILE_PATH)
        if metrics_dir and not os.path.exists(metrics_dir):
            os.makedirs(metrics_dir, exist_ok=True)

        if os.path.exists(_METRICS_FILE_PATH):
            wb = load_workbook(_METRICS_FILE_PATH)
            if _METRICS_SHEET_NAME in wb.sheetnames:
                ws = wb[_METRICS_SHEET_NAME]
            else:
                ws = wb.create_sheet(_METRICS_SHEET_NAME)
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = _METRICS_SHEET_NAME

        # 空表时补齐表头
        if ws.max_row == 1 and ws.cell(row=1, column=1).value is None:
            for idx, header in enumerate(_METRICS_HEADERS, start=1):
                ws.cell(row=1, column=idx, value=header)
        elif ws.max_row == 1:
            current_headers = [ws.cell(row=1, column=i).value for i in range(1, len(_METRICS_HEADERS) + 1)]
            if current_headers != _METRICS_HEADERS:
                ws.insert_rows(1)
                for idx, header in enumerate(_METRICS_HEADERS, start=1):
                    ws.cell(row=1, column=idx, value=header)

        ws.append([
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            module_name,
            action,
            "成功" if success else "失败",
            round(duration_seconds, 3),
            invoice_number or "",
            source_file or "",
            error_message or "",
        ])
        wb.save(_METRICS_FILE_PATH)
    except Exception as e:
        print(f"写入调用统计失败: {e}, 路径: {_METRICS_FILE_PATH}")



def _font(*args, **kwargs):
    key = (args, tuple(sorted(kwargs.items())))
    font = _FONT_CACHE.get(key)
    if font is None:
        font = _FONT_CLS(*args, **kwargs)
        _FONT_CACHE[key] = font
    return font


def _alignment(*args, **kwargs):
    key = (args, tuple(sorted(kwargs.items())))
    alignment = _ALIGNMENT_CACHE.get(key)
    if alignment is None:
        alignment = _ALIGNMENT_CLS(*args, **kwargs)
        _ALIGNMENT_CACHE[key] = alignment
    return alignment


class CustomsDeclaration:
    """生成报关资料"""

    def __init__(self):
        """存储数据"""
        self.shipping_details = {}
        self.result_df = None
        self._excel_cache = {}
        self._product_name_cache = {}
        self._declared_price_cache = {}
        self._register_combined_cache = {}
        # Excel工作表名称
        self.sheet_name = ['明细单', '装箱单', '发票', '合同', '报关单', '要素']
        # 月份编码
        self.month_code = {"1": "Jan", "2": "Feb", "3": "Mar", "4": "Apr", "5": "May", "6": "Jun", "7": "Jul",
                           "8": "Aug",
                           "9": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"}
        # 国家编码
        self.country_code = {"美国": "USA", "日本": "JP", "加拿大": "CA", "英国": "UK", "德国": "DE", "澳洲": "AU",
                             "阿联酋": "AE"}
        # 读取的发货单Sheet页
        self.invoice_sheet_name = "装箱信息"
        self.invoice_detail_sheet_name = "发货单详情"
        # 产品信息单证专用的Sheet页
        self.product_sheet_name = "产品信息单证专用"
        # 默认公司索引，默认杭州素桐电子商务有限公司
        self.company_info_index = 2
        # 公司抬头和发货人
        self.company_info = [
            {
                "company_name_zh": "杭州易达工具有限公司",
                "company_name_en": "HANGZHOU EASY CAN TOOLS CO., LTD.",
                # 发货人
                "consignor": "HANGZHOU EASY CAN TOOLS CO., LTD.\nADD:Room 403-1, Building 9, Yunkong Chengfeng Zhizhu, Wuchang Street, Yuhang District, Hangzhou, Zhejiang\nTel：0571-85236685",
                # 境内发货人
                "domestic_shipper": "杭州易达工具有限公司3301965G22",
                # 生产销售单位
                "production_sales_unit": "杭州易达工具有限公司91330108092054785Y",
                # 发票Sheet页的bank_informations
                "bank_informations": "5.BENEFICIARY NAME: HANGZHOU EASY CAN TOOLS CO., LTD.\nA/C NO: 71012025000083231\nADDRESS: Room 403-1, Building 9, Yunkong Chengfeng Zhizhu, Wuchang Street, Yuhang District, Hangzhou,Zhejiang,China\nBENEFICIARY BANK: BANK OF NINGBO\nSWIFT BIC:BKNBCN2N\nBANK ADDRESS:  NO. 700 NINGNAN SOUTH ROAD, NINGBO, P.R.CHINA, 315100",
                # 报关单价Sheet页
                "customs_declaration_price": "易达报关单价",
                # 境内货源地
                "domestic_source_of_goods": "33019 杭州其他",
                # 图片名称
                "images": ['image4', 'image5'],
            },
            {
                "company_name_zh": "杭州鲁米亚科技有限公司",
                "company_name_en": "Hangzhou Lumiya Technology Co., Ltd",
                "consignor": "Hangzhou Lumiya Technology Co., Ltd\nADD:Room 403-5, Building 9, Yunkong Chengfeng Zhizhu, Wuchang Street, Yuhang District, Hangzhou, Zhejiang\nTel：0571-85236685",
                "domestic_shipper": "杭州鲁米亚科技有限公司 (3301969CW7)",
                "production_sales_unit": "杭州鲁米亚科技有限公司 91330106MA2GNY10XB",
                "bank_informations": "Bank information:                                                                                                                                                                                                                                                     \nBENEFICIARY NAME: Hangzhou Lumiya Technology Co., Ltd\nA/C NO: 71012025000126517\nADDRESS: Room 403-5, Building 9, Yunkong Chengfeng Zhizhu, Wuchang Street, Yuhang District, Hangzhou, Zhejiang\nBENEFICIARY BANK:  BANK OF NINGBO\nSWIFT BIC:  BKNBCN2N\nBANK ADDRESS: NO. 700 NINGNAN SOUTH ROAD, NINGBO, P.R.CHINA, 315100\nCORRESPONDENT BANK: BANK OF AMERICA N.A. NEW YORK NY US \nSWIFT BIC: BOFAUS3NXXX",
                "customs_declaration_price": "鲁米亚报关单价",
                "domestic_source_of_goods": "33019 杭州其他",
                "images": ['image2', 'image3'],
            },
            {
                "company_name_zh": "杭州素桐电子商务有限公司",
                "company_name_en": "HANGZHOU SUTONG E-COMMERCE CO.,LTD.",
                "consignor": "HANGZHOU SUTONG E-COMMERCE CO.,LTD.\nADD:Room 403-3, Building 9, Yunkong Chengfeng Zhizhu, Wuchang Street, Yuhang District, Hangzhou, Zhejiang\nTel：0571-85236685",
                "domestic_shipper": "杭州素桐电子商务有限公司91330106MA2KC3T74A",
                "production_sales_unit": "杭州素桐电子商务有限公司91330106MA2KC3T74A",
                "bank_informations": "5.BENEFICIARY NAME: Hangzhou Sutong E-Commerce Co.,Ltd.\nA/C NO: 71012025000113945\nADDRESS: Room 403-3, Building 9, YunkongChengFengzhzhu, Wuchang Street, Yuhang District, Hangzhou City, Zhejiang,China\nBENEFICIARY BANK: BANK OF NINGBO\nSWIFT BIC:BKNBCN2N\nBANK ADDRESS:  146 Baoshu Road, Xihu District, Hangzhou City",
                "customs_declaration_price": "素桐报关单价",
                "domestic_source_of_goods": "31222 上海浦东新区",
                "images": ['image6', 'image7'],
            },
            {
                "company_name_zh": "杭州六浪电子商务有限公司",
                "company_name_en": "HANGZHOU LIULANG E-COMMERCE CO., LTD.",
                "consignor": "HANGZHOU LIULANG E-COMMERCE CO., LTD.\nADD:Room W318-1-2, 3rd Floor, Unit 5, Zone E, Building 10, Alibaba Digital Ecosystem Park, No.1,Ai Cheng Street, Wuchang Street, Yuhang District, Hangzhou, Zhejiang Province\nTel：0571-85236685",
                "domestic_shipper": "杭州六浪电子商务有限公司3301969GM3",
                "production_sales_unit": "杭州六浪电子商务有限公司91330106MA2KC5AM13",
                "bank_informations": "5.BENEFICIARY NAME: HANGZHOU LIULANG E-COMMERCE CO., LTD.\nA/C NO: 571915607932002\nADDRESS:  Room W318-1-2, 3rd Floor, Unit 5, Zone E, Building 10, Alibaba Digital Ecosystem Park, No.1,Ai Cheng Street, Wuchang Street, Yuhang District, Hangzhou, Zhejiang Province\nBENEFICIARY BANK: China Merchants Bank, Head Office, Shenzhen, P.R.China\nBANK ADDRESS:  HeadOffice address: China Merchants Bank Tower No.7088, Shennan Boulevard, Shenzhen, China",
                "customs_declaration_price": "易达报关单价",
                "domestic_source_of_goods": "33019 杭州其他",
                "images": ['image8', 'image9'],
            },
            {
                "company_name_zh": "杭州匠晟电子商务有限公司",
                "company_name_en": "HANGZHOU JIANGSHENG E-COMMERCE CO.,LTD",
                "consignor": "HANGZHOU JIANGSHENG E-COMMERCE CO.,LTD\nAdd: Room W318-1-1, 3rd Floor, Unit 5, Zone E, Building 10, Alibaba Digital Ecosystem Park, No.1,Ai Cheng Street, Wuchang Street, Yuhang District, Hangzhou, Zhejiang Province\nTel：0571-85236685",
                "domestic_shipper": "杭州匠晟电子商务有限公司3301969GLM",
                "production_sales_unit": "杭州匠晟电子商务有限公司91330106MA2GNY1R0B",
                "bank_informations": "5.BENEFICIARY NAME：HANGZHOU JIANGSHENG E-COMMERCE CO.,LTD\nA/C NO：571922701032002\nADDRESS：Room W318-1-1, 3rd Floor, Unit 5, Zone E, Building 10, Alibaba Digital Ecosystem Park, No.1,Ai Cheng Street, Wuchang Street, Yuhang District, Hangzhou, Zhejiang Province\nBENEFICIARY BANK：China Merchants Bank, Head Office, Shenzhen, P.R.China；\nBANK ADDRESS：HeadOffice address: China Merchants Bank Tower No.7088, Shennan Boulevard, Shenzhen, China",
                "customs_declaration_price": "匠晟报关单价",
                "domestic_source_of_goods": "33019 杭州其他",
                "images": ['image10', 'image11'],
            }
        ]
        # 是否拼箱
        self.lcl_flag = True
        # 管制物备注
        self.control_info = ""

    @staticmethod
    def open_excel_file(file_path: str, sheet_name: str) -> Optional[openpyxl.Workbook]:
        """打开Excel文件并返回Workbook对象"""
        try:
            wb = openpyxl.load_workbook(file_path)
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
            return wb
        except FileNotFoundError:
            print(f"错误: 文件 '{file_path}' 不存在")
            return None
        except Exception as e:
            print(f"打开Excel文件时出错: {e}")
            return None

    @staticmethod
    def seeting_excel_format(ws, data, mode: str) -> None:
        """设置Excel行高、列宽、合并单元格、边框、保留两位小数"""
        try:
            if mode == "column_widths":
                # 设置列宽
                for col, width in data.items():
                    ws.column_dimensions[col].width = width
            elif mode == "row_heights":
                # 设置行高
                for row, height in data.items():
                    ws.row_dimensions[row].height = height
            elif mode == "merge_cells":
                # 设置合并单元格
                for merge_range in data:
                    ws.merge_cells(merge_range)
            elif mode == "frame":
                # 设置边框
                border = _THIN_BORDER
                for row in ws.iter_rows(min_row=data['min_row'], max_row=data['max_row'], min_col=data['min_col'],
                                        max_col=data['max_col']):
                    for cell in row:
                        cell.border = border
            elif mode == "right_frame":
                # 设置右边框
                cell = ws.cell(row=data['row'], column=data['col'])
                current_border = cell.border
                cell.border = Border(
                    left=current_border.left,  # 保持左边框
                    right=_THIN_SIDE,  # 设置右边框
                    top=current_border.top,  # 保持上边框不变
                    bottom=current_border.bottom  # 保持下边框不变
                )
            elif mode == "left_frame":
                # 设置左边框
                cell = ws.cell(row=data['row'], column=data['col'])
                current_border = cell.border
                cell.border = Border(
                    left=_THIN_SIDE,  # 设置左边框
                    right=current_border.right,  # 保添加右边框
                    top=current_border.top,  # 保持上边框不变
                    bottom=current_border.bottom  # 保持下边框不变
                )
            elif mode == "cancel_internal_frame":
                # 取消内边框
                no_border = _NO_BORDER
                for row in ws.iter_rows(min_row=data['min_row'], max_row=data['max_row'], min_col=data['min_col'],
                                        max_col=data['max_col']):
                    for cell in row:
                        cell.border = no_border
                for row in ws.iter_rows(min_row=data['min_row'], max_row=data['min_row'], min_col=data['min_col'],
                                        max_col=data['max_col']):
                    for cell in row:
                        cell.border = Border(
                            top=_THIN_SIDE
                        )
            elif mode == "bottom_frame":
                # 设置底边框
                cell = ws.cell(row=data['row'], column=data['col'])
                current_border = cell.border
                cell.border = Border(
                    left=current_border.left,
                    right=current_border.right,
                    top=current_border.top,
                    bottom=_THIN_SIDE
                )
            elif mode == "keep_decimals":
                # 对指定区域设置两位小数格式
                for row in ws[data]:
                    for cell in row:
                        cell.number_format = '0.00'
        except Exception as e:
            print(f"设置Excel格式时出错: {e}")

    @staticmethod
    def set_bottom_border_range(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
        """批量设置指定区域的底边框"""
        for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            for cell in row:
                current = cell.border
                cell.border = Border(
                    left=current.left,
                    right=current.right,
                    top=current.top,
                    bottom=_THIN_SIDE
                )

    @staticmethod
    def set_left_border_range(ws, min_row: int, max_row: int, col: int) -> None:
        """批量设置指定列的左边框"""
        for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=col, max_col=col):
            for cell in row:
                current = cell.border
                cell.border = Border(
                    left=_THIN_SIDE,
                    right=current.right,
                    top=current.top,
                    bottom=current.bottom
                )

    @staticmethod
    def set_right_border_cells(ws, cells) -> None:
        """批量设置指定单元格的右边框"""
        for row, col in cells:
            cell = ws.cell(row=row, column=col)
            current = cell.border
            cell.border = Border(
                left=current.left,
                right=_THIN_SIDE,
                top=current.top,
                bottom=current.bottom
            )

    @staticmethod
    def _freeze_cache_value(value):
        if isinstance(value, list):
            return tuple(value)
        if isinstance(value, dict):
            return tuple(sorted((k, CustomsDeclaration._freeze_cache_value(v)) for k, v in value.items()))
        if isinstance(value, set):
            return tuple(sorted(value))
        return value

    def _read_excel_cached(self, file_path: str, sheet_name: str, **kwargs) -> pd.DataFrame:
        key = (
            file_path,
            sheet_name,
            tuple(sorted((k, self._freeze_cache_value(v)) for k, v in kwargs.items())),
        )
        cached = self._excel_cache.get(key)
        if cached is None:
            cached = pd.read_excel(file_path, sheet_name=sheet_name, **kwargs)
            self._excel_cache[key] = cached
        return cached.copy()

    def _get_product_name_df(self, product_name_filepath: str) -> pd.DataFrame:
        cached = self._product_name_cache.get(product_name_filepath)
        if cached is None:
            name_df = self._read_excel_cached(product_name_filepath, sheet_name="品名",
                                              usecols=['报关品名', '英文品名', 'HS CODE', '单位'])
            # 去掉.0
            name_df['HS CODE'] = name_df['HS CODE'].apply(lambda x: str(x).replace('.0', ''))
            cached = name_df
            self._product_name_cache[product_name_filepath] = cached
        return cached.copy()

    @staticmethod
    def determine_pc_count(product_name: str) -> int:
        """
        根据品名确定PC数量
        1. 首先查找是否包含"几PC"或"几pc"或者"几件"或者"几片"，有的话返回数字
        2. 如果没有包含PC或者pc，检查是否包含"套装"或"组套"，有则返回2
        3. 否则返回1
        """
        if pd.isna(product_name):
            return 1

        product_name = str(product_name)

        # 首先查找是否包含数字+PC或pc或者"几件"或者"几片"的模式,匹配"6PC", "12pc", "8PC", "24pc"等格式
        match = _PC_PATTERN.search(product_name)
        if match:
            return int(match.group(1))

        # 如果没包含PC，检查是否包含"套装"或"组套"
        if "套装" in product_name or "组套" in product_name:
            return 2

        return 1

    @staticmethod
    def find_data_group(elements_file_path: str, result_df: pd.DataFrame) -> List[List[str]]:
        """查找要素表的数据组"""

        # 存储最终的数据组
        result_groups = []
        # 预编译匹配"品名"+任意标点符号+后面的所有内容
        name_pattern = re.compile(r'品名[；;：:+\-=\s]*([^。.;；]+)')
        # 使用itertuples避免多余的中间列表
        for data in result_df.itertuples(index=False, name=None):
            # 如果是SIXPRAY品牌，就使用TOLESA Sheet页
            if data[-1] in ['SIXPRAY']:
                sheet_name = 'TOLESA'
            else:
                sheet_name = data[-1]

            cache_key = (elements_file_path, sheet_name)
            if cache_key not in _ELEMENTS_GROUP_CACHE:
                # 读取时没有列名
                elements_df = pd.read_excel(elements_file_path, sheet_name=sheet_name, header=None)
                # 数据转列表
                elements_data_list = elements_df.iloc[:, 0].tolist()
                # 将NaN转换为空字符串
                elements_data_list = [str(x) if pd.notna(x) else '' for x in elements_data_list]
                # 根据空白行分组
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

                # 预计算每组文本与金刚石标记
                prepared_groups = []
                for group in groups:
                    group_text = ' '.join(group)
                    prepared_groups.append((group, group_text, "金刚石" in group_text))
                _ELEMENTS_GROUP_CACHE[cache_key] = prepared_groups

            prepared_groups = _ELEMENTS_GROUP_CACHE[cache_key]

            # 遍历groups,查找是否有包含指定HS编码的组
            # 从result_data中获取HS编码
            target_hs_code = data[3]
            # 如果target_hs_code为空，就报错
            if pd.isna(target_hs_code):
                raise ValueError(
                    f"报关中文品名：{data[1]}，HS CODE为空，请去品名表.xlsx，请添加此报关中文品名！")

            found_group = None
            # 判断是否找到
            elements_falg = False
            for group, group_text, contains_diamond in prepared_groups:
                # 如果包含HS编码并都包含"金刚石"或者都不包含"金刚石"
                if target_hs_code in group_text and contains_diamond == data[-2]:
                    for temp_date in group:
                        match = name_pattern.search(temp_date)
                        if not match:
                            continue
                        if match.group(1).strip() == data[1].strip():
                            found_group = list(group)
                            elements_falg = True
                            break
                    if elements_falg:
                        break

            # 如果found_group为空，就报错
            if found_group is None:
                print(data)
                raise ValueError(
                    f"HS CODE：{data[3]}，报关中文品名：{data[1]}，请去要素表.xlsx的{data[-1]}页添加此要素！")

            # 添加SKU
            found_group.append(data[0])

            # 添加套数
            if data[5] == '套':
                sets = f"{data[4]}{data[5]}={data[-3]}个"
                found_group.append(sets)

            if found_group:
                result_groups.append(found_group)

        return result_groups

    @staticmethod
    def number_to_words(amount: float) -> str:
        """将数字金额转换为英文大写"""
        # 分离整数部分和小数部分
        amount_str = str(amount)
        if '.' in amount_str:
            integer_part, decimal_part = amount_str.split('.')
            integer_num = int(integer_part)
            decimal_num = int(decimal_part)
        else:
            integer_num = int(amount_str)
            decimal_num = 0

        # 转换整数部分
        integer_words = num2words(integer_num, lang='en').upper()

        # 转换小数部分
        if decimal_num > 0:
            decimal_words = num2words(decimal_num, lang='en').upper()
            result = f"SAY TOTAL USD {integer_words} AND CENTS {decimal_words} ONLY"
        else:
            result = f"SAY TOTAL USD {integer_words} AND NO CENTS ONLY"

        # 格式化：除最后一个AND外，其他AND替换为逗号
        parts = result.split(" AND ")

        if len(parts) > 1:
            # 除了最后一部分，其他部分用逗号连接
            result = ",".join(parts[:-1]) + " AND " + parts[-1]
        return result

    @staticmethod
    def calculate_net_weight(gross_series: pd.Series) -> pd.Series:
        """根据每件毛重计算每件净重"""
        return pd.Series(
            np.where(gross_series < 2, gross_series * 0.8, gross_series - 1),
            index=gross_series.index
        )

    @staticmethod
    def get_transport_method_display(transportation_method: str) -> str:
        """将运输方式转换为明细单显示文案"""
        transportation_method = str(transportation_method or "").strip()
        for key, display in _TRANSPORT_METHOD_DISPLAY_MAP.items():
            if key in transportation_method:
                return display
        return transportation_method or "BY SEA"

    @staticmethod
    def get_customs_transport_method_display(transportation_method: str) -> str:
        """将运输方式转换为报关单显示文案"""
        transportation_method = str(transportation_method or "").strip()
        for key, display in _CUSTOMS_TRANSPORT_METHOD_DISPLAY_MAP.items():
            if key in transportation_method:
                return display
        return transportation_method or "海运"

    @staticmethod
    def add_images(ws, image_path: str, cell_location: str, width=None,
                   height=None) -> None:
        """将图片悬浮在Excel工作表的指定位置，并可自定义图片大小"""

        # 检查图片是否存在
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"图片文件不存在: {image_path}")

        # 创建图片对象
        img = Image(image_path)

        # 自定义图片大小
        if width is not None:
            img.width = width
        if height is not None:
            img.height = height

        # 添加图片到指定位置
        ws.add_image(img, cell_location)

    @staticmethod
    def set_excel_page_view(file_path: str) -> None:
        """
        设置Excel为分页预览模式，并将数据放在第一页
        """
        try:
            # 尝试连接 WPS 表格应用程序
            try:
                app = win32.Dispatch("Ket.Application")
            except:
                app = win32.Dispatch("WPS.Application")
        except Exception as e:
            print(f"无法连接 WPS: {e}")
            # 尝试连接 Excel 作为备选
            try:
                app = win32.Dispatch("Excel.Application")
                print("已连接到 Microsoft Excel")
            except:
                print("也无法连接 Excel")
                raise

        try:
            # 打开工作簿
            workbook = app.Workbooks.Open(file_path)

            # 定义分页预览模式
            preview_mode = 2

            # 遍历所有工作表并设置为分页预览模式，同时调整页面设置
            for sheet in workbook.Worksheets:
                sheet.Activate()
                try:
                    # 设置分页预览模式
                    app.ActiveWindow.View = preview_mode

                    # 获取页面设置对象
                    page_setup = sheet.PageSetup

                    # 将所有内容调整为一页宽和一页高
                    page_setup.FitToPagesWide = 1
                    page_setup.FitToPagesTall = 1

                    print(f"已调整工作表 {sheet.Name} 的页面设置")

                except Exception as e:
                    print(f"无法设置工作表 {sheet.Name} 的页面设置: {e}")

            # 保存并关闭
            workbook.Save()
            workbook.Close()

        except Exception as e:
            print(f"处理文件时出错: {e}")
            raise

    def calculate_shipping_costs(self, total_volume: float, total_weight: float) -> None:
        """计算运费"""

        # 运费,总体积（m³）-箱子的总和<1时取1，总重量（kg）-箱子的总和少于12取12
        exchange_rate = 7.2
        # 默认美国
        multiplier_num1 = 1250
        # 默认小于1
        multiplier_num2 = 1
        freight_basis = "volume"
        if '海卡' in self.shipping_details['transportation_method']:
            sea_card_rates = {
                '加拿大': 800,
                '日本': 800,
                '德国': 732 if self.shipping_details[
                                   'logistics_provider'] == '北京世纪卓越快递服务有限公司（龙舟）' else 1200,
                '英国': 634 if self.shipping_details[
                                   'logistics_provider'] == '北京世纪卓越快递服务有限公司（龙舟）' else 1200,
            }
            multiplier_num1 = sea_card_rates.get(self.shipping_details['country_zh'], 1250)
            multiplier_num2 = max(1.0, total_volume)
            freight_basis = "volume"

        elif '海派' in self.shipping_details['transportation_method']:
            sea_delivery_rates = {
                '美国': 12,
                '加拿大': 12,
                '澳洲': 10,
                '日本': 10,
                '英国': 11
            }
            multiplier_num1 = sea_delivery_rates.get(self.shipping_details['country_zh'], 12)
            multiplier_num2 = max(12.0, total_weight)
            freight_basis = "weight"
        elif '空派' in self.shipping_details['transportation_method']:
            air_delivery_rates = {
                '美国': 45,
                '德国': 42,
            }
            multiplier_num1 = air_delivery_rates.get(self.shipping_details['country_zh'], 45)
            multiplier_num2 = max(12.0, total_weight)
            freight_basis = "weight"
        elif '卡航' in self.shipping_details['transportation_method']:
            truck_delivery_rates = {
                '英国': 15.5,
                '德国': 16,
            }
            multiplier_num1 = truck_delivery_rates.get(self.shipping_details['country_zh'], 15.5)
            multiplier_num2 = total_weight
            freight_basis = "weight"

        # 总运费
        self.shipping_details['freight'] = round(multiplier_num1 * multiplier_num2 / exchange_rate, 2)
        self.shipping_details['freight_basis'] = freight_basis

    def get_basic_information(self, invoice_filepath: str, site_address_filepath: str,
                              product_filepath: str):
        """获取基本信息:日期、唛头、目的国、发票号、目的港、收件人、公司抬头、发货人、境内发货人、生产销售单位、运输方式、物流商"""

        # 获取当天日期并加序号后缀
        today_list = datetime.datetime.now().strftime("%Y-%m-%d").split('-')
        day_num = int(today_list[2])
        if day_num in (1, 21):
            day_suffix = "ˢᵀ"
        elif day_num in (2, 22):
            day_suffix = "ᴺᴰ"
        elif day_num in (3, 23):
            day_suffix = "ᴿᴰ"
        else:
            day_suffix = "ᵀᴴ"
        today = "{}.{}{},{}".format(
            self.month_code[str(int(today_list[1]))],
            today_list[2],
            day_suffix,
            today_list[0],
        )
        self.shipping_details['today'] = today

        # 读取发货单的装箱信息
        invoice_df = self._read_excel_cached(invoice_filepath, sheet_name=self.invoice_sheet_name)
        invoice_df['SKU'] = invoice_df['SKU'].astype(str)
        invoice_df = invoice_df.ffill()
        if '总体积（m³）-箱子' in invoice_df.columns:
            volume_series = pd.to_numeric(invoice_df['总体积（m³）-箱子'], errors='coerce')
            invoice_df.loc[volume_series == 0, '总体积（m³）-箱子'] = 0.01

        # 货件单号列去重后转列表,得到唛头
        shipping_number_list = invoice_df["货件单号"].unique().tolist()
        shipping_number_list.append("MADE IN CHINA")
        self.shipping_details['mark'] = "\n".join(shipping_number_list)

        # 目的国
        self.shipping_details['country_zh'] = invoice_df["国家"].unique().tolist()[0]
        self.shipping_details['country'] = self.country_code[self.shipping_details['country_zh']]

        # 发货单详情Sheet页获取店铺
        invoice_details_df = self._read_excel_cached(
            invoice_filepath, sheet_name=self.invoice_detail_sheet_name, usecols=_INVOICE_DETAILS_COLS
        )
        # 去除重复数据
        invoice_details_df = invoice_details_df[
            ['SKU', '物流中心编码', '运输方式', '物流商单号', '店铺', '物流商']].drop_duplicates()
        # 检查物流商单号是否为空，为空就报错
        self.check_data(invoice_details_df, "check_invoice_num")

        # 将*SKU转为字符串
        invoice_details_df['SKU'] = invoice_details_df['SKU'].astype(str)
        # 取第一个' '或者'-'之前的内容，并且大写
        invoice_details_df['店铺'] = invoice_details_df['店铺'].str.split('[ -]').str[0].str.upper()
        # 如果店铺名称等于SIXPRAY，将店铺名称改为TOLESA
        # invoice_details_df.loc[invoice_details_df['店铺'] == "SIXPRAY", '店铺'] = "TOLESA"

        # 发票号
        invoice_number = invoice_details_df["物流商单号"].unique().tolist()[0].strip()

        self.shipping_details['invoice_number'] = invoice_number

        # 物流商
        self.shipping_details['logistics_provider'] = invoice_details_df["物流商"].unique().tolist()[0]

        # 获取物流中心编码列
        logistics_center_list = invoice_details_df["物流中心编码"][
            invoice_details_df["物流中心编码"].notna()].unique().tolist()

        # 目的港,如果国家是美国 并且 物流商是'北京世纪卓越快递服务有限公司（龙舟）'，则物流中心编码为ONT6，否则为物流中心编码列
        if len(logistics_center_list) > 1:
            # 如果物流中心编码长度大于1，就为空
            self.shipping_details['Logistics_center'] = ""
        else:
            if self.shipping_details['logistics_provider'] == '北京世纪卓越快递服务有限公司（龙舟）' and \
                    self.shipping_details['country'] == 'USA':
                self.shipping_details['Logistics_center'] = "ONT6"
            else:
                self.shipping_details['Logistics_center'] = invoice_details_df["物流中心编码"].unique().tolist()[0]

        # 收件人
        site_address_df = self._read_excel_cached(site_address_filepath, sheet_name=self.shipping_details['country_zh'])

        if self.shipping_details['Logistics_center'] == "":
            # 如果目的港为空，收件人就为"Amazon.com Services, Inc."
            self.shipping_details['recipient'] = "Amazon.com Services, Inc."
        else:
            # 查找目的港包含Logistics_center的收件人列
            try:
                self.shipping_details['recipient'] = \
                    site_address_df[site_address_df['目的港'].str.contains(self.shipping_details['Logistics_center'])][
                        '收件人'].values[0]
            except Exception:
                raise Exception(
                    f"目的港：{self.shipping_details['Logistics_center']}，不在全站点地址.xlsx的{self.shipping_details['country_zh']}页，请添加此站点！")

        # 公司抬头、发货人、境内发货人、生产销售单位
        brand = invoice_details_df["店铺"].unique().tolist()[0]
        brand_index_map = {
            'EZARC': 0,
            'LIBRATON': 1,
            'PACMANN': 1,
            'SIXPRAY': 3,
            'TOLESA': 4,
        }
        self.company_info_index = brand_index_map.get(brand, self.company_info_index)
        self.shipping_details.update(self.company_info[self.company_info_index])
        self.shipping_details['brand'] = brand

        # 运输方式
        self.shipping_details['transportation_method'] = invoice_details_df["运输方式"].unique().tolist()[0]
        if pd.isna(self.shipping_details['transportation_method']):
            raise ValueError(f"发货单的运输方式为空！")

        # 读取产品信息单证专用
        product_df = self._read_excel_cached(
            product_filepath, sheet_name=self.product_sheet_name, usecols=_PRODUCT_INFO_COLS
        )
        # 将*SKU转为字符串
        product_df['*SKU'] = product_df['*SKU'].astype(str)

        return invoice_df, invoice_details_df, product_df

    def update_shipment_quantity(self, df: pd.DataFrame) -> pd.DataFrame:
        """如果同一个发货单号里面，如果有重复的SKU，就重新计算发货数量，发货数量 = 箱数 * 单箱数量，否则就不处理"""

        dup_mask = df.duplicated(subset=["发货单号", "SKU"], keep=False)
        df.loc[dup_mask, "发货数量"] = df.loc[dup_mask, "箱数"] * df.loc[
            dup_mask, "单箱数量"]

        return df

    def check_data(self, df: pd.DataFrame, mode: str) -> None:
        """如果merge_df的*SKU列有为NaN的值，就报错，将那一行数据提示输出"""
        if mode == "check_sku":
            nan_sku_rows = df[df['*SKU'].isna()]['SKU']
            # nan_sku_rows去重
            nan_sku_rows = nan_sku_rows.drop_duplicates()
            # 如果有 NaN 值，抛出异常并显示相关行
            if not nan_sku_rows.empty:
                # 将包含 NaN 的行转换为可读字符串
                error_data = nan_sku_rows.to_string(index=False)
                # 抛出异常，包含错误数据
                raise ValueError(f"SKU:{error_data}，SKU的信息在 产品信息单证专用表 里找不到！")
        elif mode == "check_price":
            # 如果申报单价列有为NaN的数据就报错,并且将那一行数据输出
            if df['申报单价'].isna().any():
                # 获取包含NaN的行
                nan_rows = df[df['申报单价'].isna()]
                # 提取SKU列表
                sku_list = nan_rows['SKU'].tolist()
                # 格式化SKU字符串
                sku_str = '，'.join(sku_list)
                raise ValueError(
                    f"SKU：{sku_str}申报单价为空，请去产品信息单证专用.xlsx，添加此申报单价！")
        elif mode == "check_invoice_num":
            # 如果物流商单号全部都为空，就报错
            if df['物流商单号'].isna().all():
                raise ValueError(
                    f"发货单的物流商单号为空，请添加！")
        elif mode == "check_shop":
            # 如果申报单价列有为NaN的数据就报错,并且将那一行数据输出
            if df['店铺'].isna().any():
                # 获取包含NaN的行
                nan_rows = df[df['店铺'].isna()]
                # 提取SKU列表
                sku_list = nan_rows['SKU'].tolist()
                # 格式化SKU字符串
                sku_str = '，'.join(sku_list)
                raise ValueError(
                    f"SKU：{sku_str}店铺为空，请检查发货单的发货单详情和装箱信息Sheet，两边的SKU是否一致！")

    def process_declared_unit_price(self, product_filepath: str) -> pd.DataFrame:
        """处理申报单价"""

        # 申报单价
        cache_key = (product_filepath, self.shipping_details['customs_declaration_price'])
        cached = self._declared_price_cache.get(cache_key)
        if cached is not None:
            return cached.copy()

        declared_unit_price_df = self._read_excel_cached(
            product_filepath, sheet_name=self.shipping_details['customs_declaration_price']
        )
        # 易达报关单价Sheet页取"易达-亚马逊申报单价(USD)"列
        if self.shipping_details['customs_declaration_price'] == "易达报关单价":
            declared_unit_price_df = declared_unit_price_df[['*SKU', '易达-亚马逊\n申报单价(USD)']]
            # 重命名为申报单价
            declared_unit_price_df.rename(columns={'易达-亚马逊\n申报单价(USD)': '申报单价'}, inplace=True)
        elif self.shipping_details['customs_declaration_price'] == "匠晟报关单价":
            declared_unit_price_df = declared_unit_price_df[['*SKU', '匠晟-亚马逊\n申报单价(USD)']]
            # 重命名为申报单价
            declared_unit_price_df.rename(columns={'匠晟-亚马逊\n申报单价(USD)': '申报单价'}, inplace=True)
        else:
            declared_unit_price_df = declared_unit_price_df[['*SKU', '申报单价']]

        # 去除重复数据
        declared_unit_price_df = declared_unit_price_df.drop_duplicates()

        # 将*SKU转为字符串
        declared_unit_price_df['*SKU'] = declared_unit_price_df['*SKU'].astype(str)

        self._declared_price_cache[cache_key] = declared_unit_price_df
        return declared_unit_price_df.copy()

    def lcl_situation(self, invoice_filepath: str, product_filepath: str,
                      product_name_filepath: str, site_address_filepath: str) -> None:
        """处理发货单的装箱信息"""

        invoice_df, invoice_details_df, product_df = self.get_basic_information(invoice_filepath, site_address_filepath,
                                                                                product_filepath)

        # 如果同一个发货单号里面，如果有重复的SKU，就重新计算发货数量，发货数量 = 箱数 * 单箱数量，否则就不处理
        invoice_df = self.update_shipment_quantity(invoice_df)

        product_df = product_df[
            ['*SKU', '报关中文品名', '单品毛重(g)', '外箱规格长(cm)', '外箱规格宽(cm)', '外箱规格高(cm)',
             '单箱数量(pcs)', '材质']]
        # 转换单品毛重(g)
        product_df['单品毛重(g)'] = product_df['单品毛重(g)'].astype(float) / 1000
        # 重命名列
        product_df.rename(columns={'单品毛重(g)': '单品毛重(kg)'}, inplace=True)

        # 合并invoice_df和product_df
        merge_df = invoice_df.merge(product_df, left_on="SKU", right_on="*SKU", how="left")
        self.check_data(merge_df, "check_sku")

        # 删除'*SKU'列
        merge_df.drop(columns=['*SKU'], inplace=True)

        # 外箱规格长(cm) * 外箱规格宽(cm) * 外箱规格高(cm) / 单箱数量(pcs) * 发货数量  / 1000000
        merge_df['总体积'] = merge_df['外箱规格长(cm)'] * merge_df['外箱规格宽(cm)'] * merge_df[
            '外箱规格高(cm)'] / \
                             merge_df['单箱数量(pcs)'] * merge_df['发货数量'] / 1000000

        # 处理申报单价
        declared_unit_price_df = self.process_declared_unit_price(product_filepath)

        # merge_df合并declared_unit_price_df，获取SKU对应的申报单价
        merge_df = merge_df.merge(declared_unit_price_df, left_on="SKU", right_on="*SKU", how="left")

        # 如果申报单价列有为NaN的数据就报错,并且将那一行数据输出
        self.check_data(merge_df, 'check_price')
        # 删除'*SKU'列
        merge_df.drop(columns=['*SKU'], inplace=True)

        # 总PCS，根据品名
        merge_df['总PCS'] = invoice_df['品名'].apply(self.determine_pc_count) * invoice_df['发货数量']

        # 申报总价 = 申报单价 * 发货数量
        merge_df['申报总价'] = merge_df['申报单价'] * merge_df['发货数量']
        # 新增是否含金刚石列
        merge_df['是否含金刚石'] = merge_df['材质'].str.contains('金刚石', na=False)

        # 根据报关中文品名、是否含金刚石进行分组，汇总发货数量、总PCS、申报总价、总体积、箱数
        sum_df = merge_df.groupby(['报关中文品名', '是否含金刚石']).agg({
            '发货数量': 'sum',
            '申报总价': 'sum',
            '总PCS': 'sum',
            '总体积': 'sum',
            '箱数': 'sum',
            'SKU': 'first',
            '箱子长度（cm）': 'first',
            '箱子宽度（cm）': 'first'
        }).reset_index()

        # 存储拼箱数据
        sum_dict = {}
        # 拼箱的品名
        groups_list = []
        for i, row in sum_df.iterrows():
            name = (row['报关中文品名'], row['是否含金刚石'])
            sum_dict[name] = {
                'SKU': row['SKU'],
                '发货数量': row['发货数量'],
                '总PCS': row['总PCS'],
                '申报总价': row['申报总价'],
                '总体积': row['总体积'],
                '箱数': row['箱数'],
                '箱子长度（cm）': row['箱子长度（cm）'],
                '箱子宽度（cm）': row['箱子宽度（cm）'],
            }

        float_columns = ['箱子毛重（kg）', 'CBM（m³）-箱子', '总重量（kg）-箱子', '总体积（m³）-箱子']
        for column in float_columns:
            if column in merge_df.columns:
                merge_df[column] = pd.to_numeric(merge_df[column], errors='coerce').astype(float)

        # 将索引转成新的列
        merge_df.reset_index(inplace=True)
        merge_df['index'] = merge_df['index'].astype(int)
        # 按索引列排序以确保连续性
        merge_df = merge_df.sort_values('index').reset_index(drop=True)

        # 存储总体积
        lcl_total_volume = 0.0
        # 存储总重量
        lcl_total_weight = 0.0
        # 存储总箱数
        lcl_total_boxes = 0

        # 根据发货单号、箱数、报箱子毛重（kg）、箱子长度（cm）、箱子宽度（cm）、箱子高度（cm）、总重量（kg）-箱子、总体积（m³）-箱子进行分组
        for value, group in merge_df.groupby(
                ['发货单号', '箱数', '箱子毛重（kg）', '箱子长度（cm）', '箱子宽度（cm）', '箱子高度（cm）', '总重量（kg）-箱子',
                 '总体积（m³）-箱子']):
            # 如果长度等于1，则说明没有拼箱，先加入到boxes_num_dict中
            if len(group) == 1:
                # 计算总体积和总毛重
                lcl_total_volume += group['总体积（m³）-箱子'].sum()
                lcl_total_weight += group['总重量（kg）-箱子'].sum()
                lcl_total_boxes += group['箱数'].sum()
            else:
                lcl_total_volume += group.iloc[0]['总体积（m³）-箱子']
                lcl_total_weight += group.iloc[0]['总重量（kg）-箱子']
                lcl_total_boxes += group.iloc[0]['箱数']

                # 存储拼箱的数据
                groups_names = group["报关中文品名"].unique().tolist()
                groups_list = list(set(groups_list + groups_names))

                # 单品毛重汇总 = 单品毛重(kg) * 发货数量
                group['单品毛重汇总'] = group['单品毛重(kg)'] * group['发货数量']
                group['实际箱子毛重（kg）'] = group['单品毛重汇总'] / group['单品毛重汇总'].sum() * \
                                            group['箱子毛重（kg）']
                group['实际CBM（m³）-箱子'] = group['单品毛重汇总'] / group['单品毛重汇总'].sum() * \
                                            group['CBM（m³）-箱子']

                # 合并merge_df和group，根据index、SKU、发货单号合并
                merge_df = merge_df.merge(
                    group[['index', 'SKU', '发货单号', '实际箱子毛重（kg）', '实际CBM（m³）-箱子']],
                    on=['index', 'SKU', '发货单号'], how='left')
                # 如果merge_df的实际箱子毛重（kg）列不为空，则将merge_df的箱子毛重（kg）列的值替换为实际箱子毛重（kg）列的值
                condition = merge_df['实际箱子毛重（kg）'].notna()
                merge_df.loc[condition, '箱子毛重（kg）'] = merge_df.loc[condition, '实际箱子毛重（kg）']
                # 同理
                condition = merge_df['实际CBM（m³）-箱子'].notna()
                merge_df.loc[condition, 'CBM（m³）-箱子'] = merge_df.loc[condition, '实际CBM（m³）-箱子']

                # 删除列
                merge_df = merge_df.drop(columns=['实际CBM（m³）-箱子', '实际箱子毛重（kg）'])

        # 更新总重量和总体积
        merge_df['总重量（kg）-箱子'] = merge_df['箱子毛重（kg）'] * merge_df['箱数']
        merge_df['总体积（m³）-箱子'] = merge_df['CBM（m³）-箱子'] * merge_df['箱数']

        # 根据报关中文品名、是否含金刚石进行分组，汇总总重量（kg）-箱子、总体积（m³）-箱子
        weight_df = (merge_df
                     .groupby(['报关中文品名', '是否含金刚石'], as_index=False)[
                         ['总重量（kg）-箱子', '总体积（m³）-箱子']]
                     .sum())

        # 存储报关品名数据
        for i, row in weight_df.iterrows():
            name = (row['报关中文品名'], row['是否含金刚石'])
            sum_dict[name]['总重量（kg）-箱子'] = row['总重量（kg）-箱子']
            sum_dict[name]['总体积（m³）-箱子'] = row['总体积（m³）-箱子']
            # 如果是拼箱数据，就 /15，否则不变
            if row['报关中文品名'] in groups_list:
                sum_dict[name]['箱数'] = 1 if math.floor(row['总重量（kg）-箱子'] / 15.0) == 0 else math.floor(
                    row['总重量（kg）-箱子'] / 15.0)
            lcl_total_boxes -= sum_dict[name]['箱数']

        # 按照箱数从大到小排序
        sorted_data = dict(sorted(sum_dict.items(), key=lambda x: x[1]['箱数'], reverse=True))
        # 分配多余的箱数
        if lcl_total_boxes > 0:
            for key, value in sorted_data.items():
                if key[0] in groups_list:
                    value['箱数'] += lcl_total_boxes
                    break
        elif lcl_total_boxes < 0:
            for key, value in sorted_data.items():
                if key[0] in groups_list:
                    value['箱数'] += lcl_total_boxes
                    break

        # 转换为DataFrame
        result_df = pd.DataFrame.from_dict(sorted_data, orient='index').reset_index()
        result_df = result_df.rename(columns={'level_0': '报关中文品名', 'level_1': '是否含金刚石'})
        result_df['实际总体积（m³）-箱子'] = (result_df['总体积'] / result_df['总体积'].sum() * result_df[
            '总体积（m³）-箱子'].sum() * 1.05)
        # 如果'实际总体积（m³）-箱子'小于0.01，就变成0.01
        result_df.loc[result_df['实际总体积（m³）-箱子'] < 0.01, '实际总体积（m³）-箱子'] = 0.01

        result_df = result_df[
            ['报关中文品名', 'SKU', '发货数量', '总PCS', '申报总价', '箱数', '总重量（kg）-箱子', '实际总体积（m³）-箱子',
             '箱子长度（cm）', '箱子宽度（cm）', '是否含金刚石']]

        invalid_box_df = result_df[result_df['箱数'] <= 0].copy()
        if not invalid_box_df.empty:
            invalid_box_df = invalid_box_df[['SKU', '报关中文品名', '箱数', '总重量（kg）-箱子', '实际总体积（m³）-箱子']]
            raise ValueError(
                "拼箱结果存在箱数<=0的数据，无法继续生成报关资料，请检查拼箱分组或箱数回补逻辑：\n"
                f"{invalid_box_df.to_string(index=False)}"
            )

        result_df['箱子高度（cm）'] = (
                result_df['实际总体积（m³）-箱子'] * 1000000 / result_df['箱子长度（cm）'] / result_df['箱子宽度（cm）'] /
                result_df[
                    '箱数']).round(2)
        # 重命名列
        result_df.rename(columns={'实际总体积（m³）-箱子': '总体积（m³）-箱子'}, inplace=True)

        # 申报单价 = 申报总价 / 发货数量
        result_df['申报单价'] = (result_df['申报总价'] / result_df['发货数量']).round(2)
        # 箱子毛重（kg） = 总重量（kg）-箱子 / 箱数
        result_df['箱子毛重（kg）'] = (result_df['总重量（kg）-箱子'] / result_df['箱数']).round(2)

        # 计算运费
        self.calculate_shipping_costs(lcl_total_volume, lcl_total_weight)

        # 含运费申报单价 = 按运费计费口径分摊(体积/重量)
        freight_basis = self.shipping_details.get('freight_basis', 'volume')
        if freight_basis == 'weight':
            base_series = result_df['总重量（kg）-箱子']
            basis_label = '总重量（kg）-箱子'
        else:
            base_series = result_df['总体积（m³）-箱子']
            basis_label = '总体积（m³）-箱子'
        base_series = pd.to_numeric(base_series, errors='coerce').fillna(0)
        base_total = base_series.sum()
        if not base_total:
            raise ValueError(
                f"运费分摊失败：分摊口径为{basis_label}，但合计为0，请检查报关汇总数据该列。"
            )
        allocation = base_series / base_total * self.shipping_details['freight']
        result_df['含运费申报单价'] = (
                allocation / result_df['发货数量'] + result_df['申报单价']).round(2)
        # 含运费申报总价 = 含运费申报单价 * 发货数量
        result_df['含运费申报总价'] = result_df['含运费申报单价'] * result_df['发货数量']
        # 包装种类
        result_df['包装种类'] = '纸箱'
        # 每件净重KGS
        result_df['每件净重KGS'] = self.calculate_net_weight(result_df['箱子毛重（kg）'])
        # 共重净重KGS
        result_df['共重净重KGS'] = (result_df['每件净重KGS'] * result_df['箱数']).round(2)
        result_df[['总重量（kg）-箱子', '总体积（m³）-箱子']] = result_df[['总重量（kg）-箱子', '总体积（m³）-箱子']].round(2)

        result_df = result_df[
            ['SKU', '报关中文品名', '发货数量', '箱数', '包装种类', '含运费申报单价', '含运费申报总价',
             '箱子毛重（kg）', '总重量（kg）-箱子', '每件净重KGS', '共重净重KGS', '箱子长度（cm）', '箱子宽度（cm）',
             '箱子高度（cm）', '总体积（m³）-箱子', '总PCS', '是否含金刚石']]

        # 重命名列
        result_df.rename(columns={'含运费申报单价': '申报单价', '含运费申报总价': '总价', '箱子毛重（kg）': '每件毛重KGS',
                                  '总重量（kg）-箱子': '箱子总毛重（kg）', '总体积（m³）-箱子': '体积'}, inplace=True)

        # 读取品名表
        name_df = self._get_product_name_df(product_name_filepath)
        # 合并merge_df和name_df
        result_df = result_df.merge(name_df, left_on="报关中文品名", right_on="报关品名", how="left")
        # 删除'报关品名'列
        result_df.drop(columns=['报关品名'], inplace=True)

        result_df = result_df[
            ['SKU', '报关中文品名', '英文品名', 'HS CODE', '发货数量', '单位', '箱数', '包装种类', '申报单价', '总价',
             '每件毛重KGS', '箱子总毛重（kg）', '每件净重KGS', '共重净重KGS', '箱子长度（cm）', '箱子宽度（cm）',
             '箱子高度（cm）', '体积', '总PCS', '是否含金刚石']]

        # 总发货数量
        self.shipping_details['TOTAL_QUANTITY'] = result_df['发货数量'].sum()
        # 总箱数
        self.shipping_details['TOTAL_PACKED_IN'] = result_df['箱数'].sum()
        # 总毛重 KGS
        self.shipping_details['TOTAL_G.W'] = result_df['箱子总毛重（kg）'].sum().round(2)
        # 总净重 KGS
        self.shipping_details['TOTAL_N.W'] = result_df['共重净重KGS'].sum().round(2)
        # 总体积
        self.shipping_details['TOTAL_MEAS'] = result_df['体积'].sum().round(2)
        # 总价
        self.shipping_details['TOTAL_PRICE'] = result_df['总价'].sum().round(2)

        # 筛选发货单详情Sheet页的字段
        invoice_details_df = invoice_details_df[['SKU', '店铺']]
        # 去除重复数据
        invoice_details_df = invoice_details_df.drop_duplicates()

        # 合并result_df和invoice_details_df，根据SKU
        self.result_df = result_df.merge(invoice_details_df, on="SKU", how="left")

        print(self.result_df)



    def process_invoice(self, invoice_filepath: str, product_filepath: str,
                        product_name_filepath: str, site_address_filepath: str) -> None:
        """处理发货单的装箱信息"""

        invoice_df, invoice_details_df, product_df = self.get_basic_information(invoice_filepath, site_address_filepath,
                                                                                product_filepath)
        # 如果SKU列有重复的数据，就汇总合并
        if invoice_df['SKU'].duplicated().any():
            # 汇总数据，根据SKU分组，合并发货数量、箱数、总体积、总毛重，同一SKU的单箱数量、箱子毛重、箱子长度、箱子宽度、品名保持不变
            invoice_df = invoice_df.groupby('SKU').agg({
                '发货数量': 'sum',
                '箱数': 'sum',
                '总体积（m³）-箱子': 'sum',
                '总重量（kg）-箱子': 'sum',
                '单箱数量': 'first',
                '箱子毛重（kg）': 'first',
                '箱子长度（cm）': 'first',
                '箱子宽度（cm）': 'first',
                '品名': 'first'
            }).reset_index()
        else:
            print(f"\n{'=' * 60}")
            print("不是分5仓的情况")
            print(f"{'=' * 60}")

        # 计算运费
        total_volume = invoice_df['总体积（m³）-箱子'].sum()
        total_weight = invoice_df['总重量（kg）-箱子'].sum()
        self.calculate_shipping_costs(total_volume, total_weight)

        # 运费分摊：按运费计费口径分摊(体积/重量)
        freight_basis = self.shipping_details.get('freight_basis', 'volume')
        if freight_basis == 'weight':
            base_series = invoice_df['总重量（kg）-箱子']
            basis_label = '总重量（kg）-箱子'
        else:
            base_series = invoice_df['总体积（m³）-箱子']
            basis_label = '总体积（m³）-箱子'
        base_series = pd.to_numeric(base_series, errors='coerce').fillna(0)
        base_total = base_series.sum()
        if not base_total:
            raise ValueError(
                f"运费分摊失败：分摊口径为{basis_label}，但合计为0，请检查装箱信息该列数据。"
            )
        invoice_df['运费分摊'] = self.shipping_details['freight'] / base_total * base_series

        # 总PCS，根据品名
        invoice_df['总PCS'] = invoice_df['品名'].apply(self.determine_pc_count) * invoice_df['发货数量']

        # 筛选数据列
        invoice_df = invoice_df[
            ["SKU", "发货数量", "总重量（kg）-箱子", "箱数", "总体积（m³）-箱子", '箱子毛重（kg）', '箱子长度（cm）',
             '箱子宽度（cm）', '运费分摊', '总PCS']]
        # 将SKU转为字符串
        invoice_df['SKU'] = invoice_df['SKU'].astype(str)

        product_df = product_df[
            ['*SKU', '报关中文品名', '货源地', '供应商名称', '报关用采购成本(CNY)', '材质']]

        # 读取品名表
        name_df = self._get_product_name_df(product_name_filepath)
        # 合并product_df和name_df
        product_df = product_df.merge(name_df, left_on="报关中文品名", right_on="报关品名", how="left")
        # 删除'报关品名'列
        product_df.drop(columns=['报关品名'], inplace=True)

        # 处理申报单价
        declared_unit_price_df = self.process_declared_unit_price(product_filepath)

        # invoice_df合并product_df，获取SKU对应的货源地、供应商名称、报关中文品名、报关用采购成本(CNY)，
        merge_df = invoice_df.merge(product_df, left_on="SKU", right_on="*SKU", how="left")

        # 如果merge_df的*SKU列有为NaN的值，就报错，将那一行数据提示输出
        self.check_data(merge_df, "check_sku")
        # 删除'*SKU'列
        merge_df.drop(columns=['*SKU'], inplace=True)
        # merge_df合并declared_unit_price_df，获取SKU对应的申报单价
        merge_df = merge_df.merge(declared_unit_price_df, left_on="SKU", right_on="*SKU", how="left")
        # 如果申报单价列有为NaN的数据就报错,并且将那一行数据输出
        self.check_data(merge_df, 'check_price')
        # 删除'*SKU'列
        merge_df.drop(columns=['*SKU'], inplace=True)
        # 申报总价
        merge_df['申报总价'] = merge_df['发货数量'] * merge_df['申报单价'] + merge_df['运费分摊']
        # 箱子总毛重（kg）
        merge_df['箱子总毛重（kg）'] = merge_df['箱子毛重（kg）'] * merge_df['箱数']
        # 新增是否含金刚石列
        merge_df['是否含金刚石'] = merge_df['材质'].str.contains('金刚石', na=False)

        # 对数据进行分组，对同一报关中文品名、是否含金刚石分组，合计申报总价、发货数量、总重量（kg）-箱子、箱数、总体积（m³）-箱子、'箱子总毛重（kg）'、'总PCS'
        sum_df = (merge_df
                  .groupby(['报关中文品名', '是否含金刚石'], as_index=False)[
                      ['申报总价', '发货数量', '总重量（kg）-箱子', '箱数', '总体积（m³）-箱子', '箱子总毛重（kg）', '总PCS']]
                  .sum())

        # 用 apply 单独取每组申报总价最多的 SKU,以及获取SKU对应的箱子长度（cm）、箱子宽度（cm）
        top_sku_sr = (merge_df
                      .sort_values('申报总价', ascending=False)
                      .groupby(['报关中文品名', '是否含金刚石'])[['SKU']]
                      .first())

        # 用 apply 单独取每组箱数、总体积（m³）-箱子最多的长度和宽度
        top_length_sr = (merge_df
                         .sort_values(['箱数', '总体积（m³）-箱子'], ascending=False)
                         .groupby(['报关中文品名', '是否含金刚石'])[['箱子长度（cm）', '箱子宽度（cm）']]
                         .first())

        # 合并结果
        temp_df = sum_df.merge(top_sku_sr, on=['报关中文品名', '是否含金刚石'])
        result_df = temp_df.merge(top_length_sr, on=['报关中文品名', '是否含金刚石'])

        # 申报单价,保留两位小数
        result_df['申报单价'] = (result_df['申报总价'] / result_df['发货数量']).round(2)

        # 单箱重量
        result_df['单箱重量'] = (result_df['总重量（kg）-箱子'] / result_df['箱数'])

        # 获取英文品名、中国报关\nhs编码、单位、供应商名称
        temp_df = product_df[['*SKU', '英文品名', 'HS CODE', '单位', '供应商名称']]
        result_df = result_df.merge(temp_df, left_on="SKU", right_on="*SKU", how="left")
        # 删除'*SKU'列
        result_df.drop(columns=['*SKU'], inplace=True)

        # 包装种类
        result_df['包装种类'] = '纸箱'
        # 总价
        result_df['总价'] = result_df['发货数量'] * result_df['申报单价']
        # 每件毛重KGS
        result_df['每件毛重KGS'] = (result_df['箱子总毛重（kg）'] / result_df['箱数'])
        # 每件净重KGS
        result_df['每件净重KGS'] = self.calculate_net_weight(result_df['每件毛重KGS'])
        # 共重净重KGS
        result_df['共重净重KGS'] = (result_df['每件净重KGS'] * result_df['箱数'])
        # 高
        result_df['箱子高度（cm）'] = (
                result_df['总体积（m³）-箱子'] * 1000000 / result_df['箱子长度（cm）'] / result_df['箱子宽度（cm）'] /
                result_df[
                    '箱数'])
        # 体积
        result_df['体积'] = (
                result_df['箱子高度（cm）'] * result_df['箱子宽度（cm）'] * result_df['箱子长度（cm）'] * result_df[
            '箱数'] / 1000000)

        # 对指定列保留两位小数
        result_df[
            ['申报单价', '总价', '每件毛重KGS', '箱子总毛重（kg）', '每件净重KGS', '共重净重KGS', '箱子长度（cm）',
             '箱子宽度（cm）', '箱子高度（cm）', '体积']] = result_df[
            ['申报单价', '总价', '每件毛重KGS', '箱子总毛重（kg）', '每件净重KGS', '共重净重KGS', '箱子长度（cm）',
             '箱子宽度（cm）', '箱子高度（cm）', '体积']].round(2)

        # 列顺序
        result_df = result_df[
            ['SKU', '报关中文品名', '英文品名', 'HS CODE', '发货数量', '单位', '箱数', '包装种类', '申报单价',
             '总价', '每件毛重KGS', '箱子总毛重（kg）', '每件净重KGS', '共重净重KGS', '箱子长度（cm）', '箱子宽度（cm）',
             '箱子高度（cm）', '体积', '总PCS', '是否含金刚石']]

        # 总发货数量
        self.shipping_details['TOTAL_QUANTITY'] = result_df['发货数量'].sum()
        # 总箱数
        self.shipping_details['TOTAL_PACKED_IN'] = result_df['箱数'].sum()
        # 总毛重 KGS
        self.shipping_details['TOTAL_G.W'] = result_df['箱子总毛重（kg）'].sum().round(2)
        # 总净重 KGS
        self.shipping_details['TOTAL_N.W'] = result_df['共重净重KGS'].sum().round(2)
        # 总体积
        self.shipping_details['TOTAL_MEAS'] = result_df['体积'].sum().round(2)
        # 总价
        self.shipping_details['TOTAL_PRICE'] = result_df['总价'].sum().round(2)

        # 筛选发货单详情Sheet页的字段
        invoice_details_df = invoice_details_df[['SKU', '店铺']]
        # 去除重复数据
        invoice_details_df = invoice_details_df.drop_duplicates()

        # 合并result_df和invoice_details_df，根据SKU
        self.result_df = result_df.merge(invoice_details_df, on="SKU", how="left")

        # 检查物流商单号是否为空，为空就报错
        self.check_data(self.result_df, "check_shop")

    def add_detailed_list_sheet(self, excel_file_path: str, image_path: str,
                                wb: Optional[openpyxl.Workbook] = None) -> None:
        """添加明细单sheet页"""

        save_after = False
        if wb is None:
            # 创建新工作簿
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = self.sheet_name[0]
            save_after = True
        else:
            if self.sheet_name[0] in wb.sheetnames:
                del wb[self.sheet_name[0]]
            ws = wb.create_sheet(self.sheet_name[0])

        # 如果是素桐，就显示上海浦东新区
        domestic_source_display = "上海浦东新区" if self.company_info_index == 2 else "杭州其他"

        # 设置数据
        excel_data = [
            [self.shipping_details['company_name_zh'], "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
             ""],
            [self.shipping_details['company_name_en'], "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
             ""],
            ["出货明细表", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["发货人", self.shipping_details['consignor'], "", "", "", "", "", "", "", "", "", "", "日期", "", "",
             self.shipping_details['today'],
             "",
             ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", "发票号", "", "", self.shipping_details['invoice_number'],
             "", ""],
            ["收件人", self.shipping_details['recipient'], "", "", "", "", "", "", "", "", "", "", "境内货源地", "", "",
             domestic_source_display, "",
             ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", "起运口岸", "", "", "宁波", "",
             ""],
            ["通知人", "SAME AS CONSIGNEE", "", "", "", "", "", "", "", "", "", "", "目的港", "", "",
             self.shipping_details['Logistics_center'], "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", "目的国", "", "", self.shipping_details['country'], "",
             ""],
            ["唛头", self.shipping_details['mark'], "", "", "", "", "", "", "", "", "", "", "价格条款", "", "",
             "CFR",
             "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", "付款方式", "", "", "TT", "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", "运输方式", "", "",
             self.get_transport_method_display(self.shipping_details.get('transportation_method', '')), "",
             ""],
            ["货号", "品名", "品名(英文）", "商品编码", "数量", "单位", "包装\n件数", "包装\n种类", "单价", "总价",
             "毛重 KGS",
             "", "净重 KGS", "", "尺寸 KGS", "", "", "体积"],
            ["", "", "", "", "", "", "", "", "USD", "USD", "每件", "共重", "每件", "共重", "L", "W", "H", "CBM"]
        ]

        # 复制result_df
        detailed_list_df = self.result_df.copy()

        # 将result_data中的列表加入到excel_data
        # 删除'店铺','总PCS','是否含金刚石'列
        detailed_list_df.drop(columns=['店铺', '总PCS', '是否含金刚石'], inplace=True)

        detailed_list__data = detailed_list_df.values.tolist()
        excel_data.extend(detailed_list__data)

        # 增加合计行
        excel_data.append(
            ["合计", "", "", "", self.shipping_details['TOTAL_QUANTITY'], "", self.shipping_details['TOTAL_PACKED_IN'],
             "",
             self.shipping_details['TOTAL_PRICE'], "", self.shipping_details['TOTAL_G.W'], "",
             self.shipping_details['TOTAL_N.W'], "",
             self.shipping_details['TOTAL_MEAS'], "", "", ""])

        # 最大行数
        max_row = len(excel_data)

        # 填充数据
        for row_idx, row_data in enumerate(excel_data, 1):
            # 设置行高
            if row_idx in [6, 7]:
                ws.row_dimensions[row_idx].height = 45
            else:
                ws.row_dimensions[row_idx].height = 25.5

            for col_idx, value in enumerate(row_data, 1):
                # 要求数值保留小数点后两位，如果小数点位数不足两位，要求其补0

                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                # 设置单元格格式
                if row_idx == 1:
                    cell.font = _font(name='宋体-简', size=20)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == 2:
                    cell.font = _font(name='Cambria', size=18)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == 3:
                    cell.font = _font(name='宋体-简', bold=True, size=16)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx <= 12:
                    if col_idx == 13:
                        cell.font = _font(name='宋体', bold=True, size=10)
                    else:
                        cell.font = _font(name='Cambria', bold=True, size=10)
                    if col_idx == 2:
                        cell.alignment = _alignment(horizontal='left', vertical='center', wrap_text=True)
                    else:
                        cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx <= max_row:
                    cell.font = _font(name='Cambria', bold=True, size=10)
                    cell.alignment = _alignment(horizontal='center', vertical='center', wrap_text=True)

        # 对指定区域保留两位小数
        self.seeting_excel_format(ws, f"I15:R{max_row}", "keep_decimals")

        # 设置列宽
        column_widths = {
            'A': 12, 'B': 12, 'C': 19, 'D': 11.96,
            'E': 6.04, 'F': 4.71, 'G': 4.71,
            'H': 4.71, 'I': 7, 'J': 10,
            'K': 7, 'L': 10, 'M': 7,
            'N': 10, 'O': 7, 'P': 7,
            'Q': 7, 'R': 7
        }
        self.seeting_excel_format(ws, column_widths, "column_widths")

        # 设置边框
        frame_data = {
            'min_row': 4, 'max_row': max_row, 'min_col': 1, 'max_col': 18
        }
        self.seeting_excel_format(ws, frame_data, "frame")

        # 合并单元格
        merge_ranges = [
            "A1:R1",  # 杭州
            "A2:R2",  # Hangzhou
            "A3:R3",  # 出货明细表
            "A4:A5",  # 发货人
            "B4:L5",  # 发货人数值
            "M4:O4",  # 日期
            "P4:R4",  # 日期值
            "M5:O5",  # 发票号
            "P5:R5",  # 发票号数值
            "A6:A7",  # 收件人
            "B6:L7",  # 收件人数值
            "M6:O6",  # 境内货源地
            "P6:R6",  # 境内货源地数值
            "M7:O7",  # 起运口岸
            "P7:R7",  # 起运口岸数值
            "A8:A9",  # 通知人
            "B8:L9",  # 通知人数值
            "M8:O8",  # 目的港
            "P8:R8",  # 目的港数值
            "M9:O9",  # 目的国
            "P9:R9",  # 目的国数值
            "A10:A12",  # 唛头
            "B10:L12",  # 唛头数值
            "M10:O10",  # 价格条款
            "P10:R10",  # 价格条款数值
            "M11:O11",  # 付款方式
            "P11:R11",  # 付款方式数值
            "M12:O12",  # 运输方式
            "P12:R12",  # 运输方式数值
            "A13:A14",  # 货号
            "B13:B14",  # 品名
            "C13:C14",  # 品名(英文)
            "D13:D14",  # 商品编码
            "E13:E14",  # 数量
            "F13:F14",  # 单位
            "G13:G14",  # 包装件数
            "H13:H14",  # 包装种类
            "K13:L13",  # 毛重 KGS
            "M13:N13",  # 净重 KGS
            "O13:Q13",  # 尺寸 KGS
            f"I{max_row}:J{max_row}",  # 总价合计
            f"K{max_row}:L{max_row}",  # 毛重 KGS合计
            f"M{max_row}:N{max_row}",  # 净重 KGS合计
            f"O{max_row}:R{max_row}"  # 体积合计
        ]
        self.seeting_excel_format(ws, merge_ranges, "merge_cells")

        # 添加图片
        self.add_images(ws, f"{image_path}\\{self.shipping_details['images'][0]}.png", f"B{max_row}", 300, 150)

        if save_after:
            # 保存文件
            wb.save(excel_file_path)
            print(f"已成功在 '{excel_file_path}' 中添加明细单Sheet页")

    def add_packing_list_sheet(self, excel_file_path: str, image_path: str,
                               wb: Optional[openpyxl.Workbook] = None) -> None:
        """添加装箱单sheet页"""
        save_after = False
        if wb is None:
            # 打开现有的Excel文件
            wb = self.open_excel_file(excel_file_path, self.sheet_name[1])
            save_after = True
        else:
            if self.sheet_name[1] in wb.sheetnames:
                del wb[self.sheet_name[1]]

        # 创建新的"装箱单"工作表
        ws = wb.create_sheet(self.sheet_name[1])

        # 填充数据
        excel_data = [
            [self.shipping_details['company_name_zh'], "", "", "", "", "", "", "", "", "", "", "", ""],
            [self.shipping_details['company_name_en'], "", "", "", "", "", "", "", "", "", "", "", ""],
            ["装箱单 PACKING LIST", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["TO:", "", "", "", "", "", "", "", "", "", "", "", ""],
            [self.shipping_details["recipient"], "", "", "", "", "", "", "日 期:\nDate:", "",
             self.shipping_details['today'], "", "",
             ""],
            ["", "", "", "", "", "", "", "发票号:\nInvoice No.", "", self.shipping_details['invoice_number'], "", "",
             ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["Mark & Numbers", "ITEM NO.:", "DESCRIPTION", "QUANTITY", "", "", "", "G. W.\n(KGS)", "N. W.\n(KGS)",
             "MEASUREMENT\n(CM)", "", "", "CBM"],
            ["", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["", "", "TOTAL QUANTITY:", self.shipping_details['TOTAL_QUANTITY'], "", "SETS", "", "", "", "", "", "",
             ""],
            ["", "", "TOTAL PACKED IN:", self.shipping_details['TOTAL_PACKED_IN'], "", "CTNS", "", "", "", "", "", "",
             ""],
            ["", "", "TOTAL G.W.:", self.shipping_details['TOTAL_G.W'], "", "KGS", "", "", "", "", "", "", ""],
            ["", "", "TOTAL N.W.:", self.shipping_details['TOTAL_N.W'], "", "KGS", "", "", "", "", "", "", ""],
            ["", "", "TOTAL MEAS.:", self.shipping_details['TOTAL_MEAS'], "", "M³", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", ""]
        ]

        # 插入的数据
        result_data = self.result_df.values.tolist()
        insert_data = []
        for temp_data in result_data:
            insert_data.append(
                [self.shipping_details['mark'], temp_data[0], temp_data[2], temp_data[4],
                 'SETS' if temp_data[5] == '套' else 'PCS', temp_data[6], 'CTNS',
                 temp_data[11], temp_data[13], temp_data[14], temp_data[15], temp_data[16], temp_data[17]])
        # 数据行的长度
        data_row = len(insert_data)

        # 将insert_data中的列表加入到excel_data
        excel_data[8:8] = insert_data
        # print(excel_data)

        # 有数据的最大行数
        max_row = len(excel_data) - 4

        # 填充数据
        for row_idx, row_data in enumerate(excel_data, 1):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)

                # 设置单元格格式
                if row_idx == 1:
                    cell.font = _font(name='宋体-简', size=22)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == 2:
                    cell.font = _font(name='Calibri', size=19)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == 3:
                    cell.font = _font(name='Calibri', bold=True, size=16)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == 4:
                    cell.font = _font(name='Calibri', bold=True, size=11)
                    cell.alignment = _alignment(horizontal='left', vertical='center')
                elif row_idx <= 6:
                    cell.font = _font(name='Calibri', bold=True, size=11)
                    if col_idx == 1:
                        cell.alignment = _alignment(horizontal='left', vertical='center', wrap_text=True)
                    elif col_idx == 8:
                        cell.alignment = _alignment(horizontal='right', vertical='center', wrap_text=True)
                    else:
                        cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == 8:
                    cell.font = _font(name='Calibri', bold=True, size=11)
                    cell.alignment = _alignment(horizontal='center', vertical='center', wrap_text=True)
                elif row_idx <= 8 + data_row:
                    if col_idx == 1:
                        cell.alignment = _alignment(horizontal='center', vertical='center', wrap_text=True)
                    else:
                        cell.alignment = _alignment(horizontal='center', vertical='center')
                    cell.font = _font(name='Calibri', bold=True, size=11)
                elif row_idx <= max_row:
                    cell.font = _font(name='Calibri', bold=True, size=11, italic=True)
                    if col_idx in [3, 6]:
                        cell.alignment = _alignment(horizontal='left', vertical='center')
                    else:
                        cell.alignment = _alignment(horizontal='center', vertical='center')

        # 设置列宽
        column_widths = {
            'A': 18, 'B': 11, 'C': 25, 'D': 7,
            'E': 6, 'F': 7, 'G': 6,
            'H': 8.5, 'I': 8.13, 'J': 8.13,
            'K': 8.13, 'L': 8.13, 'M': 8.13
        }
        self.seeting_excel_format(ws, column_widths, "column_widths")

        # 设置行高
        row_heights = {
            1: 30.65, 2: 24, 3: 23.5, 4: 12,
            5: 45, 6: 45, 7: 19.5, 8: 43.5
        }
        for i in range(9, 9 + data_row):
            row_heights[i] = 30
        for i in range(9 + data_row, len(excel_data)):
            row_heights[i] = 15
        self.seeting_excel_format(ws, row_heights, "row_heights")

        # 设置边框
        frame_data = {
            'min_row': 8, 'max_row': 8 + data_row, 'min_col': 1, 'max_col': 13
        }
        self.seeting_excel_format(ws, frame_data, "frame")

        # 对指定区域保留两位小数
        self.seeting_excel_format(ws, f"H9:M{8 + data_row}", "keep_decimals")
        self.seeting_excel_format(ws, f"D{max_row - 2}:D{max_row}", "keep_decimals")

        # 合并单元格
        merge_ranges = [
            "A1:M1",  # 公司标题
            "A2:M2",  # Hangzhou
            "A3:M3",  # 装箱单标题
            "A4:L4",  # TO:
            "A5:G6",  # TO值
            "H5:I5",  # 日期
            "J5:L5",  # 日期数值
            "H6:I6",  # 发票号
            "J6:L6",  # 发票号数值
            "D8:G8",  # QUANTITY
            "J8:L8",  # MEASUREMENT(CM)
            f"A9:A{8 + data_row}",  # Mark & Numbers
            f"D{max_row - 4}:E{max_row - 4}",  # TOTAL QUANTITY
            f"D{max_row - 3}:E{max_row - 3}",  # TOTAL PACKED IN
            f"D{max_row - 2}:E{max_row - 2}",  # TOTAL G.W.
            f"D{max_row - 1}:E{max_row - 1}",  # TOTAL N.W.
            f"D{max_row}:E{max_row}"  # TOTAL MEAS
        ]
        self.seeting_excel_format(ws, merge_ranges, "merge_cells")

        # 添加图片
        self.add_images(ws, f"{image_path}\\{self.shipping_details['images'][0]}.png", f"G{max_row - 5}", 300, 150)

        if save_after:
            # 保存文件
            wb.save(excel_file_path)
            print(f"已成功在 '{excel_file_path}' 中添加装箱单Sheet页")

    def add_bill_sheet(self, excel_file_path: str, image_path: str,
                       wb: Optional[openpyxl.Workbook] = None) -> None:
        """添加发票sheet页"""
        save_after = False
        if wb is None:
            # 打开现有的Excel文件
            wb = self.open_excel_file(excel_file_path, self.sheet_name[2])
            save_after = True
        else:
            if self.sheet_name[2] in wb.sheetnames:
                del wb[self.sheet_name[2]]

        # 创建新的"装箱单"工作表
        ws = wb.create_sheet(self.sheet_name[2])

        # 填充数据
        excel_data = [
            [self.shipping_details['company_name_zh'], "", "", "", "", "", "", "", ""],
            [self.shipping_details['company_name_en'], "", "", "", "", "", "", "", ""],
            ["发票 COMMERCIAL INVOICE", "", "", "", "", "", "", "", ""],
            ["TO:", "", "", "", "", "", "", "", ""],
            [self.shipping_details["recipient"], "", "", "", "", "日 期:\nDate:", "", self.shipping_details['today'],
             ""],
            ["", "", "", "", "", "发票号:\nInvoice No.", "", self.shipping_details['invoice_number'], ""],
            ["Mark & Numbers", "ITEM NO.:", "DESCRIPTION", "QUANTITY", "", "UNIT PRICE", "", "AMOUNT\nCFR NINGBO", ""],
            ["BANK INFORMATIONS:", "", "", "", "", "TOTAL:", "", "USD", self.shipping_details['TOTAL_PRICE']],
            [self.shipping_details['bank_informations'], "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", ""]
        ]

        # 插入的数据
        result_data = self.result_df.values.tolist()
        insert_data = []
        for temp_data in result_data:
            insert_data.append(
                [self.shipping_details['mark'], temp_data[0], temp_data[2], temp_data[4],
                 'SETS' if temp_data[5] == '套' else 'PCS', 'USD', temp_data[8],
                 'USD',
                 temp_data[9]])
        # 数据行的长度
        data_row = len(insert_data)

        # 将insert_data中的列表加入到excel_data
        excel_data[7:7] = insert_data
        # print(excel_data)

        # 有数据的最大行数
        max_row = len(excel_data)

        # 填充数据
        for row_idx, row_data in enumerate(excel_data, 1):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)

                # 设置单元格格式
                if row_idx == 1:
                    cell.font = _font(name='宋体-简', size=22)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == 2:
                    cell.font = _font(name='Calibri', size=19)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == 3:
                    cell.font = _font(name='Calibri', bold=True, size=16)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == 4:
                    cell.font = _font(name='Calibri', bold=True, size=11)
                    cell.alignment = _alignment(horizontal='left', vertical='center')
                elif row_idx <= 6:
                    cell.font = _font(name='Calibri', bold=True, size=11)
                    if col_idx == 1:
                        cell.alignment = _alignment(horizontal='left', vertical='center', wrap_text=True)
                    elif col_idx == 6:
                        cell.alignment = _alignment(horizontal='right', vertical='center', wrap_text=True)
                    else:
                        cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == 7:
                    cell.font = _font(name='Calibri', bold=True, size=11)
                    cell.alignment = _alignment(horizontal='center', vertical='center', wrap_text=True)
                elif row_idx <= 7 + data_row:
                    if col_idx == 1:
                        cell.alignment = _alignment(horizontal='center', vertical='center', wrap_text=True)
                    else:
                        cell.alignment = _alignment(horizontal='center', vertical='center')
                    cell.font = _font(name='Calibri', bold=True, size=11)
                elif row_idx <= max_row:
                    if col_idx == 1:
                        cell.font = _font(name='Calibri', size=11)
                        cell.alignment = _alignment(horizontal='left', vertical='top', wrap_text=True)
                    else:
                        cell.font = _font(name='Calibri', bold=True, size=11)
                        cell.alignment = _alignment(horizontal='center', vertical='center')

        # 设置列宽
        column_widths = {
            'A': 18, 'B': 11, 'C': 22, 'D': 5.21,
            'E': 5.21, 'F': 5.21, 'G': 5.21,
            'H': 5.21, 'I': 12
        }
        self.seeting_excel_format(ws, column_widths, "column_widths")

        # 设置行高
        row_heights = {
            1: 30.65, 2: 23.5, 3: 27, 4: 13,
            5: 45, 6: 45, 7: 33.75, 8 + data_row: 15
        }
        for i in range(8, 8 + data_row):
            row_heights[i] = 30
        for i in range(9 + data_row, len(excel_data) + 1):
            row_heights[i] = 30
        self.seeting_excel_format(ws, row_heights, "row_heights")

        # 设置边框
        frame_data = {
            'min_row': 7, 'max_row': 7 + data_row, 'min_col': 1, 'max_col': 9
        }
        self.seeting_excel_format(ws, frame_data, "frame")

        # 对指定区域保留两位小数
        self.seeting_excel_format(ws, f"G8:G{7 + data_row}", "keep_decimals")
        self.seeting_excel_format(ws, f"I8:I{8 + data_row}", "keep_decimals")

        # 合并单元格
        merge_ranges = [
            "A1:I1",  # 公司标题
            "A2:I2",  # Hangzhou
            "A3:I3",  # 装箱单标题
            "A4:I4",  # TO:
            "A5:E6",  # TO值
            "F5:G5",  # 日期
            "H5:I5",  # 日期数值
            "F6:G6",  # 发票号
            "H6:I6",  # 发票号数值
            "D7:E7",  # QUANTITY
            "F7:G7",  # UNIT PRICE
            "H7:I7",  # AMOUNT FOB NINGBO
            f"A8:A{7 + data_row}",  # Mark & Numbers
            f"A{8 + data_row}:C{8 + data_row}",  # BANK INFORMATIONS:
            f"F{8 + data_row}:G{8 + data_row}",  # TOTAL:
            f"A{9 + data_row}:E{max_row}"
        ]
        self.seeting_excel_format(ws, merge_ranges, "merge_cells")

        # 添加图片
        self.add_images(ws, f"{image_path}\\{self.shipping_details['images'][0]}.png", f"D{max_row - 4}", 300, 150)

        if save_after:
            # 保存文件
            wb.save(excel_file_path)
            print(f"已成功在 '{excel_file_path}' 中添加发票Sheet页")

    def add_contract_sheet(self, excel_file_path: str, image_path: str,
                           wb: Optional[openpyxl.Workbook] = None) -> None:
        """添加合同sheet页"""
        save_after = False
        if wb is None:
            # 打开现有的Excel文件
            wb = self.open_excel_file(excel_file_path, self.sheet_name[3])
            save_after = True
        else:
            if self.sheet_name[3] in wb.sheetnames:
                del wb[self.sheet_name[3]]

        # 创建新的合同工作表
        ws = wb.create_sheet(self.sheet_name[3])

        # 填充数据
        excel_data = [
            [self.shipping_details["company_name_zh"], "", "", "", "", "", "", "", "", "", ""],
            [self.shipping_details["company_name_en"], "", "", "", "", "", "", "", "", "", ""],
            ["销  售  合  同  SALES CONTRACT", "", "", "", "", "", "", "", "", "", ""],
            ["Seller:", self.shipping_details["consignor"], "", "", "", "", "", "", "日 期:\nDate:", "",
             self.shipping_details["today"], ""],
            ["Buyer:", self.shipping_details['recipient'], "", "", "", "", "", "", "发票号:\nInvoice No.", "",
             self.shipping_details["invoice_number"], ""],
            ["The Buyer and the Seller, through friendly negotiation, have ececuted this contract and agreed to the term and conditions stipulated below.",
             "", "", "", "", "", "", "", "", "", "", ""],
            ["ITEM NO.:", "Description & Specification", "", "Quantity", "", "No. of Packs", "", "Meas\n(M3)",
             " Unit Price",
             "", "Amount", ""],
            ["", "", "", "", "", "", "", "", "", "", "", ""],
            [self.number_to_words(self.shipping_details['TOTAL_PRICE']), "", "", "", "", "", "", "",
             "", "Total", "USD", self.shipping_details['TOTAL_PRICE']],
            ["", "", "", "", "", "", "", "", "", "", "", ""],
            ["1.Trade Terms: CFR", "", "", "", "", "", "", "", "", "", "", ""],
            [f"2.Destination Port/Country: Amazon address ,{self.shipping_details['country']}", "", "", "", "", "", "",
             "", "", "", "", ""],
            ["3.Delivery Date : Around 35DAYS AFTER RECEIVED ADVANCE PAYMENT.", "", "", "", "", "", "", "", "", "", "",
             ""],
            ["4.Payment term: 30% DEPOSIT IN ADVANCE ,BALANCE BEFORE SHIPMENT", "", "", "", "", "", "", "", "", "", "",
             ""],
            [self.shipping_details['bank_informations'], "", "", "", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", ""],
            ["BUYER:                                                                                                                          SELLER: ",
             "", "", "", "", "", "", "", "", "", "", ""]
        ]

        # 插入的数据
        result_data = self.result_df.values.tolist()
        insert_data = []
        for temp_data in result_data:
            insert_data.append(
                [temp_data[0], temp_data[2], '', temp_data[4], 'SETS' if temp_data[5] == '套' else 'PCS', temp_data[6],
                 'CTNS', temp_data[17], 'USD',
                 temp_data[8],
                 'USD', temp_data[9]])
        # 数据行的长度
        data_row = len(insert_data)

        # 将insert_data中的列表加入到excel_data
        excel_data[8:8] = insert_data

        # 有数据的最大行数
        max_row = len(excel_data)

        # 填充数据
        for row_idx, row_data in enumerate(excel_data, 1):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)

                # 设置单元格格式
                if row_idx == 1:
                    cell.font = _font(name='宋体-简', size=22)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == 2:
                    cell.font = _font(name='Calibri', size=19)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == 3:
                    cell.font = _font(name='Calibri', bold=True, size=16)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx <= 5:
                    if col_idx in [1, 9]:
                        cell.font = _font(name='Calibri', bold=True, size=11)
                    else:
                        cell.font = _font(name='Calibri', bold=True, size=10)
                    if col_idx <= 8:
                        cell.alignment = _alignment(horizontal='left', vertical='center', wrap_text=True)
                    elif col_idx <= 10:
                        cell.alignment = _alignment(horizontal='right', vertical='center', wrap_text=True)
                    else:
                        cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == 6:
                    cell.font = _font(name='Calibri', size=9)
                    cell.alignment = _alignment(horizontal='left', vertical='center')
                elif row_idx <= 8:
                    cell.font = _font(name='Calibri', bold=True, size=10)
                    cell.alignment = _alignment(horizontal='center', vertical='center', wrap_text=True)
                elif row_idx <= 9 + data_row:
                    cell.font = _font(name='Calibri', bold=True, size=10)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == max_row - 2:
                    cell.font = _font(name='Calibri', size=10)
                    cell.alignment = _alignment(horizontal='left', vertical='top', wrap_text=True)
                elif row_idx <= max_row - 1:
                    cell.font = _font(name='Calibri', size=10)
                    cell.alignment = _alignment(horizontal='left', vertical='center')
                elif row_idx <= max_row:
                    cell.font = _font(name='Calibri', size=11)
                    cell.alignment = _alignment(horizontal='left', vertical='top')

        # 设置列宽
        column_widths = {
            'A': 12, 'B': 12, 'C': 11.96, 'D': 10,
            'E': 7, 'F': 10, 'G': 7,
            'H': 10, 'I': 7, 'J': 10,
            'K': 7, 'L': 12
        }
        self.seeting_excel_format(ws, column_widths, "column_widths")

        # 设置行高
        row_heights = {
            1: 30.65, 2: 24, 3: 23.5, 4: 57,
            5: 90, 6: 14.25, 7: 16, 8: 16,
            9 + data_row: 30,
            max_row - 2: 117,
            max_row - 1: 15,
            max_row: 32.5
        }
        for i in range(9, 9 + data_row):
            row_heights[i] = 18
        for i in range(11 + data_row, max_row - 2):
            row_heights[i] = 15
        self.seeting_excel_format(ws, row_heights, "row_heights")

        # 设置边框
        frame_data = {
            'min_row': 7, 'max_row': 9 + data_row, 'min_col': 1, 'max_col': 12
        }
        self.seeting_excel_format(ws, frame_data, "frame")

        # 对指定区域保留两位小数
        self.seeting_excel_format(ws, f"H9:H{8 + data_row}", "keep_decimals")
        self.seeting_excel_format(ws, f"J9:J{8 + data_row}", "keep_decimals")
        self.seeting_excel_format(ws, f"L9:L{9 + data_row}", "keep_decimals")

        # 合并单元格
        merge_ranges = [
            "A1:L1",  # 公司标题
            "A2:L2",  # Hangzhou
            "A3:L3",  # 装箱单标题
            "B4:H4",  # Seller
            "I4:J4",  # 日期
            "K4:L4",  # 日期数值
            "B5:H5",  # Buyer
            "I5:J5",  # 发票号
            "K5:L5",  # Date数值
            "A6:L6",
            "A7:A8",  # ITEM NO.
            "B7:C8",  # Description & Specification
            "D7:E8",  # Quantity
            "F7:G8",  # No. of Packs
            "H7:H8",  # Meas(M3)
            "I7:J8",  # Unit Price
            "K7:L8",  # Amount
            f"A{9 + data_row}:I{9 + data_row}",
            f"A{11 + data_row}:L{11 + data_row}",
            f"A{12 + data_row}:L{12 + data_row}",
            f"A{13 + data_row}:L{13 + data_row}",
            f"A{14 + data_row}:L{14 + data_row}",
            f"A{15 + data_row}:L{15 + data_row}",
            f"A{max_row}:L{max_row}"
        ]
        for row in range(9, 9 + data_row):
            merge_ranges.append(f"B{row}:C{row}")
        self.seeting_excel_format(ws, merge_ranges, "merge_cells")

        # 添加图片
        self.add_images(ws, f"{image_path}\\{self.shipping_details['images'][1]}.png", f"B{max_row}", 150, 150)

        if save_after:
            # 保存文件
            wb.save(excel_file_path)
            print(f"已成功在 '{excel_file_path}' 中添加合同Sheet页")

    def add_customs_declaration_sheet(self, excel_file_path: str, image_path: str,
                                      wb: Optional[openpyxl.Workbook] = None) -> None:
        """添加报关单sheet页"""

        save_after = False
        if wb is None:
            # 打开现有的Excel文件
            wb = self.open_excel_file(excel_file_path, self.sheet_name[4])
            save_after = True
        else:
            if self.sheet_name[4] in wb.sheetnames:
                del wb[self.sheet_name[4]]

        # 创建新的"发票"工作表
        ws = wb.create_sheet(self.sheet_name[4])

        # 如果含金刚石，则将control_info改为'不属于管制物项'
        if self.result_df['是否含金刚石'].any():
            self.control_info = '不属于管制物项'

        # 填充数据
        excel_data = [
            ["中华人民共和国海关出口货物报关单", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["", "预录入编号：", "", "", "", "", "海关编号：", "", "", "", "", "", "", "", "页码/页数：1/1"],
            ["境内发货人", "", "", "", "出境关别", "", "", "出口日期", "", "", "申报日期", "", "", "备案号", ""],
            [self.shipping_details['domestic_shipper'], "", "", "", "宁波海关", "", "", "", "", "", "", "", "", "", ""],
            ["境外收货人", "", "", "", "运输方式", "", "", "运输工具名称及航次号", "", "", "提运单号", "", "", "", ""],
            ["Amazon.com Services, Inc.", "", "", "",
             self.get_customs_transport_method_display(self.shipping_details.get('transportation_method', '')),
             "", "", "", "", "", "", "", "", "", ""],
            ["生产销售单位", "", "", "", "监管方式", "", "", "征免性质", "", "", "许可证号", "", "", "", ""],
            [self.shipping_details['production_sales_unit'], "", "", "", "9710", "", "", "一般征税", "", "", "", "", "",
             "",
             ""],
            ["合同协议号", "", "", "", "贸易国（地区）", "", "", "运抵国（地区）", "", "", "指运港", "", "离境口岸", "",
             ""],
            [self.shipping_details['invoice_number'], "", "", "", self.shipping_details['country_zh'], "", "",
             self.shipping_details['country_zh'], "", "", "", "", "", "", ""],
            ["包装种类", "", "", "", "件数", "", "毛重（千克）", "净重（千克）", "", "成交方式", "运费", "", "保费", "",
             "杂费"],
            ["纸箱", "", "", "", self.shipping_details['TOTAL_PACKED_IN'], "", self.shipping_details['TOTAL_G.W'],
             self.shipping_details['TOTAL_N.W'], "", "CFR", f"${self.shipping_details['freight']}", "", "", "", ""],
            ["集装箱箱号", "", "商品项号关系", "", "集装箱货重", "", "", "随附单证及编码", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["标记唛码及备注", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
            [self.control_info, "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["项号", "商品编号", "商品名称及规格型号", "", "数量及单位", "", "单价/总价/币制", "", "", "原产国（地区）",
             "",
             "最终目的国（地区）", "", "境内货源地", "征免"],
            ["", "", "", "", "", "", "", "", "支付特许权使用费确认：", "", "", "", "自报自缴：", "", ""],
            ["申报人员", "", "申报人员证号", "", "电话", "", "", "兹申明以上内容承担如实申报、依法纳税之法律责任", "",
             "",
             "", "", "海关批注及签章", "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["申报单位", "", "", "", "", "", "", "", "申报单位（签章）", "", "", "", "", "", ""]
        ]

        # 插入的数据
        result_data = self.result_df.values.tolist()
        insert_data = []
        for i in range(0, len(result_data)):
            temp_data = result_data[i]
            insert_data.append(
                [i + 1, temp_data[3], temp_data[1], '', temp_data[4], temp_data[5], temp_data[8], temp_data[9],
                 '（502）\n美元', '中国', '', self.shipping_details['country_zh'], '',
                 self.shipping_details['domestic_source_of_goods'], '照章征税'])
        # 数据行的长度
        data_row = len(insert_data)

        # 将insert_data中的列表加入到excel_data
        excel_data[17:17] = insert_data

        # 有数据的最大行数
        max_row = len(excel_data)

        # 填充数据
        for row_idx, row_data in enumerate(excel_data, 1):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)

                # 设置单元格格式
                if row_idx == 1:
                    cell.font = _font(name='宋体', bold=True, size=16)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == 2:
                    cell.font = _font(name='宋体', size=10)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx <= 12:
                    if row_idx in [3, 5, 7, 9, 11]:
                        cell.font = _font(name='宋体', size=10)
                    else:
                        cell.font = _font(name='宋体', bold=True, size=10)
                    cell.alignment = _alignment(horizontal='left', vertical='center')
                elif row_idx in [13, 17]:
                    cell.font = _font(name='宋体', size=10)
                    cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx in [15, max_row - 3]:
                    cell.font = _font(name='宋体', size=10)
                    cell.alignment = _alignment(horizontal='left', vertical='center')
                elif row_idx <= 17 + data_row:
                    if col_idx == 14:
                        cell.font = _font(name='宋体', bold=True, size=12, color="FF0000")
                        cell.alignment = _alignment(horizontal='center', vertical='center', wrap_text=True)
                    elif col_idx == 9:
                        cell.font = _font(name='宋体', bold=True, size=12)
                        cell.alignment = _alignment(horizontal='center', vertical='center', wrap_text=True)
                    else:
                        if col_idx == 2:
                            cell.font = _font(name='Calibri', bold=True, size=11)
                        else:
                            if row_idx == 16:
                                cell.font = _font(name='宋体', bold=True, size=12, color="FF0000")
                            else:
                                cell.font = _font(name='宋体', bold=True, size=12)
                        cell.alignment = _alignment(horizontal='center', vertical='center')
                elif row_idx == max_row - 2:
                    cell.font = _font(name='宋体', size=9)
                    if col_idx == 8:
                        cell.alignment = _alignment(horizontal='center', vertical='center')
                    else:
                        cell.alignment = _alignment(horizontal='left', vertical='center')
                elif row_idx == max_row:
                    cell.font = _font(name='宋体', size=9)
                    if col_idx == 9:
                        cell.alignment = _alignment(horizontal='center', vertical='center')
                    else:
                        cell.alignment = _alignment(horizontal='left', vertical='center')

        # 设置列宽
        column_widths = {
            'A': 4, 'B': 13, 'C': 10.38, 'D': 14.25, 'E': 6.71, 'F': 5.46,
            'G': 9.79, 'H': 11.88, 'I': 9, 'J': 7.96, 'K': 2.88, 'L': 9.02,
            'M': 3.46, 'N': 14.75, 'O': 12.46
        }
        self.seeting_excel_format(ws, column_widths, "column_widths")

        # 设置行高
        row_heights = {
            1: 45, 2: 16.5, 3: 18.75, 4: 18, 5: 18, 6: 16.5, 7: 19.5,
            8: 17.25, 9: 20.25, 10: 16.5, 11: 19.5, 12: 15.75, 13: 21.75,
            14: 21.75, 15: 20.25, 16: 25.5, 17: 18, max_row - 3: 19.5,
            max_row - 2: 19.5, max_row - 1: 13.5, max_row: 19.5
        }
        for i in range(18, 18 + data_row):
            row_heights[i] = 40
        self.seeting_excel_format(ws, row_heights, "row_heights")

        # 设置边框
        frame_data = {
            'min_row': 3, 'max_row': max_row - 3, 'min_col': 1, 'max_col': 15
        }
        self.seeting_excel_format(ws, frame_data, "frame")

        # 取消内边框
        temp_frame_data = [
            {
                'min_row': 17, 'max_row': 17, 'min_col': 1, 'max_col': 15
            },
            {
                'min_row': 18, 'max_row': 17 + data_row, 'min_col': 3, 'max_col': 15
            },
            {
                'min_row': 18 + data_row, 'max_row': 18 + data_row, 'min_col': 3, 'max_col': 8
            }
        ]
        for i in range(3, 16, 2):
            temp_frame_data.append({'min_row': i, 'max_row': i + 1, 'min_col': 1, 'max_col': 15})
        for frame_data in temp_frame_data:
            self.seeting_excel_format(ws, frame_data, "cancel_internal_frame")

        # 设置右边框
        right_border_cells = []
        for i in range(3, 5):
            for j in [4, 7, 10, 14, 15]:
                right_border_cells.append((i, j))
        for i in range(5, 9):
            for j in [4, 7, 10, 15]:
                right_border_cells.append((i, j))
        for i in range(9, 11):
            for j in [4, 7, 10, 12, 15]:
                right_border_cells.append((i, j))
        for i in range(11, 13):
            for j in [4, 6, 7, 9, 10, 12, 14, 15]:
                right_border_cells.append((i, j))
        for i in range(13, 15):
            for j in [2, 4, 7, 15]:
                right_border_cells.append((i, j))
        for i in range(15, 18 + data_row):
            right_border_cells.append((i, 15))
        self.set_right_border_cells(ws, right_border_cells)

        # 设置底边框
        self.set_bottom_border_range(ws, 18, 17 + data_row, 3, 15)
        self.set_bottom_border_range(ws, 18 + data_row, 18 + data_row, 3, 8)

        # 设置左边框
        self.set_left_border_range(ws, 3, 17, 1)

        # 对指定区域保留两位小数
        self.seeting_excel_format(ws, f"G18:H{17 + data_row}", "keep_decimals")
        self.seeting_excel_format(ws, f"G12:H12", "keep_decimals")

        # 合并单元格
        merge_ranges = [
            "A1:O1",
            "K3:M3",
            "K4:M4",
            "K5:O5",
            "K6:O6",
            "K7:O7",
            "K8:O8",
            "K9:L9",
            "K10:L10",
            "M9:N9",
            "M10:N10",
            "E11:F11",
            "E12:F12",
            "H11:I11",
            "H12:I12",
            "K11:L11",
            "K12:L12",
            "M11:N11",
            "M12:N12",
            "A13:B13",
            "A14:B14",
            "C13:D13",
            "C14:D14",
            "E13:F13",
            "E14:F14",
            "H13:I13",
            "A15:O15",
            "A16:O16",
            "C17:D17",
            "E17:F17",
            "G17:I17",
            "J17:K17",
            "L17:M17",
            f"A{max_row - 2}:B{max_row - 2}",
            f"H{max_row - 2}:L{max_row - 2}",
            f"M{max_row - 2}:N{max_row - 2}",
            f"M{max_row - 3}:N{max_row - 3}",
            f"I{max_row - 3}:J{max_row - 3}",
            f"A{max_row}:B{max_row}",
            f"I{max_row}:J{max_row}",
        ]
        for row in range(3, 13):
            merge_ranges.append(f"A{row}:D{row}")
        for row in range(3, 11):
            merge_ranges.append(f"E{row}:G{row}")
            merge_ranges.append(f"H{row}:J{row}")
        for row in range(18, 18 + data_row):
            merge_ranges.append(f"C{row}:D{row}")
            merge_ranges.append(f"J{row}:K{row}")
            merge_ranges.append(f"L{row}:M{row}")
        self.seeting_excel_format(ws, merge_ranges, "merge_cells")

        # 添加图片
        self.add_images(ws, f"{image_path}\\{self.shipping_details['images'][1]}.png", f"C{max_row}", 150, 150)
        self.add_images(ws, f"{image_path}\\image1.png", "B1", 50, 50)

        if save_after:
            # 保存文件
            wb.save(excel_file_path)
            print(f"已成功在 '{excel_file_path}' 中添加报关单Sheet页")

    def add_elements_sheet(self, elements_file_path: str, excel_file_path: str, image_path: str,
                           wb: Optional[openpyxl.Workbook] = None) -> None:
        """添加要素sheet页"""
        save_after = False
        if wb is None:
            # 打开现有的Excel文件
            wb = self.open_excel_file(excel_file_path, self.sheet_name[5])
            save_after = True
        else:
            if self.sheet_name[5] in wb.sheetnames:
                del wb[self.sheet_name[5]]

        # 创建新的工作表
        ws = wb.create_sheet(self.sheet_name[5])

        # 查找数据
        result_groups = self.find_data_group(elements_file_path, self.result_df)

        # 设置数据
        excel_data = []
        for group in result_groups:
            for group_data in group:
                excel_data.append([group_data])
            excel_data.append([""])

        # 设置列宽
        column_widths = {
            'A': 40,
            'B': 20
        }
        self.seeting_excel_format(ws, column_widths, "column_widths")

        # 填充数据
        for row_idx, row_data in enumerate(excel_data, 1):
            for col_idx, value in enumerate(row_data, 1):
                ws.cell(row=row_idx, column=col_idx, value=value)

        # 添加图片
        self.add_images(ws, f"{image_path}\\{self.shipping_details['images'][1]}.png", "B5", 150, 150)

        if save_after:
            # 保存文件
            wb.save(excel_file_path)
            print(f"已成功在 '{excel_file_path}' 中添加要素Sheet页")

    def copy_declaration_template_file(self, output_dir: str) -> str:
        """按店铺复制对应申报模板文件，后续用于填充数据"""
        if not output_dir:
            raise ValueError("输出目录为空，无法复制申报模板文件")
        os.makedirs(output_dir, exist_ok=True)

        brand = str(self.shipping_details.get('brand', '')).strip().upper()
        template_path = _DECLARATION_TEMPLATE_MAP.get(brand, _DEFAULT_DECLARATION_TEMPLATE)

        if not os.path.exists(template_path):
            raise FileNotFoundError(f"申报模板文件不存在：{template_path}")

        template_filename = os.path.basename(template_path)
        target_filename = template_filename
        target_path = os.path.join(output_dir, target_filename)

        if os.path.exists(target_path):
            self.shipping_details['declaration_template_output'] = target_path
            log_message(f"复用已有申报文件 -> 店铺: {brand or '默认'}, 输出: {target_path}")
            return target_path

        shutil.copyfile(template_path, target_path)
        self.shipping_details['declaration_template_output'] = target_path
        log_message(f"已复制申报模板文件 -> 店铺: {brand or '默认'}, 模板: {template_path}, 输出: {target_path}")
        return target_path

    def fill_declaration_template_file(self, template_output_path: str) -> None:
        """填充申报模板文件（第一个Sheet页）"""
        if not os.path.exists(template_output_path):
            raise FileNotFoundError(f"申报模板文件不存在：{template_output_path}")
        if self.result_df is None or self.result_df.empty:
            raise ValueError("self.result_df为空，无法填充申报模板文件")

        declaration_df = self.result_df.copy()

        brand = str(self.shipping_details.get('brand', '')).strip().upper()
        invoice_number = str(self.shipping_details.get('invoice_number', '')).strip()
        if not invoice_number:
            raise ValueError("发票号为空，无法填充申报模板文件")

        report_time = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        platform_code = _PLATFORM_CODE_MAP.get(brand, _DEFAULT_PLATFORM_CODE)
        platform_name = _PLATFORM_NAME_MAP.get(brand, _DEFAULT_PLATFORM_NAME)

        required_cols = ['HS CODE', '报关中文品名', '单位', '发货数量', '申报单价', '总价']
        missing_cols = [col for col in required_cols if col not in declaration_df.columns]
        if missing_cols:
            raise ValueError(f"申报模板填充失败：缺少必要列 {missing_cols}")

        unit_values = declaration_df['单位'].dropna().astype(str).str.strip()
        if unit_values.empty:
            raise ValueError("申报模板填充失败：单位列为空，无法匹配计量单位代码")

        try:
            unit_df = pd.read_excel(
                template_output_path,
                sheet_name="计量单位代码",
                header=1,
                usecols=['计量单位代码', '计量单位名称'],
                dtype={'计量单位代码': str, '计量单位名称': str}
            )
        except Exception as e:
            raise ValueError(f"读取模板的计量单位代码Sheet失败：{e}")

        unit_df['计量单位代码'] = unit_df['计量单位代码'].fillna('').astype(str).str.strip()
        unit_df['计量单位名称'] = unit_df['计量单位名称'].astype(str).str.strip()

        # 按行匹配单位代码（避免整票存在多个单位时取错）
        def resolve_unit_code(unit_name: str) -> str:
            exact_match = unit_df[unit_df['计量单位名称'] == unit_name]
            if exact_match.empty:
                exact_match = unit_df[unit_df['计量单位名称'].str.contains(re.escape(unit_name), na=False)]
            if exact_match.empty:
                raise ValueError(f"计量单位代码未匹配到：单位={unit_name}")
            unit_code_value = str(exact_match.iloc[0]['计量单位代码']).strip()
            if re.fullmatch(r"\d+(?:\.0+)?", unit_code_value):
                unit_code_value = unit_code_value.split('.')[0].zfill(3)
            return unit_code_value

        declaration_df['单位'] = declaration_df['单位'].fillna('').astype(str).str.strip()
        if (declaration_df['单位'] == '').any():
            missing_rows = declaration_df.index[declaration_df['单位'] == ''].tolist()
            raise ValueError(f"申报模板填充失败：存在空单位，行索引={missing_rows}")

        unit_code_cache = {}
        declaration_df['计量单位代码'] = declaration_df['单位'].map(
            lambda unit_name: unit_code_cache.setdefault(unit_name, resolve_unit_code(unit_name))
        )

        declaration_df['订单编号'] = invoice_number
        declaration_df['报送时间'] = report_time
        declaration_df['报送类型'] = '1'
        declaration_df['报送状态'] = '2'
        declaration_df['订单类型'] = 'B'
        declaration_df['电商平台代码'] = platform_code
        declaration_df['电商平台名称'] = platform_name
        declaration_df['商品金额'] = self.shipping_details.get('TOTAL_PRICE', 0)
        declaration_df['运杂费'] = self.shipping_details.get('freight', 0)
        declaration_df['币制'] = '502'
        declaration_df['明细币制'] = '502'
        wb = load_workbook(template_output_path)
        ws = wb.worksheets[0]
        data_start_row = 4
        center_alignment = _alignment(horizontal='center', vertical='center')

        def _is_blank(value) -> bool:
            return value is None or str(value).strip() == ''

        # 查找已有明细最后一行（按订单编号/HS CODE判断）
        last_data_row = data_start_row - 1
        for row_idx in range(ws.max_row, data_start_row - 1, -1):
            order_no = ws.cell(row=row_idx, column=5).value
            hs_code = ws.cell(row=row_idx, column=13).value
            if not (_is_blank(order_no) and _is_blank(hs_code)):
                last_data_row = row_idx
                break

        existing_invoice_numbers = set()
        for row_idx in range(data_start_row, last_data_row + 1):
            order_no = ws.cell(row=row_idx, column=5).value
            if not _is_blank(order_no):
                existing_invoice_numbers.add(str(order_no).strip())

        if invoice_number in existing_invoice_numbers:
            raise ValueError(f"申报文件中已存在订单编号：{invoice_number}，禁止重复追加")

        start_row = max(data_start_row, last_data_row + 1)
        # 追加场景下，L列项号按本次批次重新从1开始
        start_seq = 1

        for idx, (_, row) in enumerate(declaration_df.iterrows(), start=0):
            excel_row = start_row + idx
            ws.cell(row=excel_row, column=1, value=row['报送类型'])          # A
            ws.cell(row=excel_row, column=2, value=row['报送时间'])          # B
            ws.cell(row=excel_row, column=3, value=row['报送状态'])          # C
            ws.cell(row=excel_row, column=4, value=row['订单类型'])          # D
            ws.cell(row=excel_row, column=5, value=row['订单编号'])          # E
            ws.cell(row=excel_row, column=6, value=row['电商平台代码'])      # F
            ws.cell(row=excel_row, column=7, value=row['电商平台名称'])      # G
            ws.cell(row=excel_row, column=8, value=row['商品金额'])          # H
            ws.cell(row=excel_row, column=9, value=row['运杂费'])            # I
            ws.cell(row=excel_row, column=10, value=row['币制'])             # J
            ws.cell(row=excel_row, column=12, value=start_seq + idx)         # L
            ws.cell(row=excel_row, column=13, value=row['HS CODE'])          # M
            ws.cell(row=excel_row, column=14, value=row['报关中文品名'])      # N
            o_cell = ws.cell(row=excel_row, column=15, value=str(row['计量单位代码']))  # O
            o_cell.data_type = 's'
            o_cell.number_format = '@'
            ws.cell(row=excel_row, column=16, value=row['明细币制'])          # P
            ws.cell(row=excel_row, column=17, value=row['发货数量'])         # Q
            ws.cell(row=excel_row, column=18, value=row['申报单价'])          # R
            ws.cell(row=excel_row, column=19, value=row['总价'])             # S
            for col_idx in range(1, 20):
                ws.cell(row=excel_row, column=col_idx).alignment = center_alignment

        wb.save(template_output_path)
        log_message(
            f"申报模板填充完成 -> 文件: {template_output_path}, 店铺: {brand or '默认'}, "
            f"起始行: {start_row}, 起始项号: {start_seq}, 单位映射: {unit_code_cache}, 行数: {len(declaration_df)}"
        )

    def generate_all_documents(self, invoice_filepath: str, product_filepath: str,
                               product_name_filepath: str, site_address_filepath: str,
                               elements_filepath: str, output_filepath: str, pictures_path: str,
                               flag: bool, generate_declaration_file: bool = True) -> None:
        """生成所有文档"""

        # 更新，是否拼箱
        self.lcl_flag = flag
        start_time = datetime.datetime.now()
        source_file = os.path.basename(invoice_filepath)

        try:
            log_message(
                f"开始生成报关资料(发货单) | 发货单: {os.path.basename(invoice_filepath)}, "
                f"输出目录: {output_filepath}, 拼箱: {self.lcl_flag}"
            )

            # 处理发货单
            if flag:
                print("拼箱")
                self.lcl_situation(invoice_filepath, product_filepath, product_name_filepath, site_address_filepath)
            else:
                print("不拼箱")
                self.process_invoice(invoice_filepath, product_filepath, product_name_filepath, site_address_filepath)

            # 更新输出路径
            output_dir = output_filepath
            output_filepath = os.path.join(output_dir, f"{self.shipping_details['invoice_number']} 报关资料.xlsx")
            log_message(f"报关资料输出文件: {output_filepath}")

            # 添加各sheet页（统一保存，减少反复读写）
            wb = openpyxl.Workbook()
            self.add_detailed_list_sheet(output_filepath, pictures_path, wb=wb)
            if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
                del wb["Sheet"]
            self.add_packing_list_sheet(output_filepath, pictures_path, wb=wb)
            self.add_bill_sheet(output_filepath, pictures_path, wb=wb)
            self.add_contract_sheet(output_filepath, pictures_path, wb=wb)
            self.add_customs_declaration_sheet(output_filepath, pictures_path, wb=wb)
            self.add_elements_sheet(elements_filepath, output_filepath, pictures_path, wb=wb)
            wb.save(output_filepath)

            # 复制并准备填充申报模板（可选）
            if generate_declaration_file:
                declaration_template_path = self.copy_declaration_template_file(output_dir)
                self.fill_declaration_template_file(declaration_template_path)
            else:
                log_message("申报文件开关为否，跳过申报模板复制与填充")

            # 设置分页预览
            self.set_excel_page_view(output_filepath)
            log_message("报关资料生成完成(发货单)")
            append_call_record(
                module_name="发货单",
                action="生成报关资料",
                success=True,
                duration_seconds=(datetime.datetime.now() - start_time).total_seconds(),
                invoice_number=self.shipping_details.get('invoice_number', ''),
                source_file=source_file,
            )

        except Exception as e:
            print(f"生成文档时出错: {e}")
            log_message(f"报关资料生成失败(发货单): {e}")
            append_call_record(
                module_name="发货单",
                action="生成报关资料",
                success=False,
                duration_seconds=(datetime.datetime.now() - start_time).total_seconds(),
                invoice_number=self.shipping_details.get('invoice_number', ''),
                source_file=source_file,
                error_message=str(e),
            )
            raise

    def check_invoice_number(self, register_filepath: str, invoice_number: str, sheet_name: str,
                             column_name: str) -> int:
        """检查发票号是否已经登记过，如果登记过就报错"""

        register_df = pd.read_excel(register_filepath, sheet_name=sheet_name, usecols=[column_name])
        register_df = register_df.ffill()

        # 查找发货票的数据是否存在
        filtered_df = register_df[register_df[column_name] == invoice_number]

        # 如果filtered_df为空，尝试用包含列名进行匹配
        if filtered_df.empty:
            filtered_df = register_df[register_df[column_name].str.contains(invoice_number)]

        if not filtered_df.empty:
            raise Exception("发票号已登记")

        return 0

    def register_combined_invoice(self, invoice_filepath: str) -> pd.DataFrame:
        """合并发货单数据"""
        cached = self._register_combined_cache.get(invoice_filepath)
        if cached is not None:
            return cached.copy()

        # 读取发货单数据
        invoice_df = self._read_excel_cached(
            invoice_filepath,
            sheet_name=self.invoice_sheet_name,
            usecols=[
                '发货单号', 'SKU', '发货数量', '箱数', '总重量（kg）-箱子', '总体积（m³）-箱子', '国家', '单箱数量',
                '货件单号', '箱子毛重（kg）', '箱子长度（cm）', '箱子宽度（cm）', '箱子高度（cm）'
            ],
        )
        invoice_df = invoice_df.ffill()
        invoice_df['SKU'] = invoice_df['SKU'].astype(str)
        invoice_df = invoice_df[
            ['发货单号', 'SKU', '发货数量', '箱数', '总重量（kg）-箱子', '总体积（m³）-箱子', '国家', '单箱数量',
             '货件单号', '箱子毛重（kg）', '箱子长度（cm）', '箱子宽度（cm）', '箱子高度（cm）']]

        # 读取发货单详情
        invoice_details_df = self._read_excel_cached(
            invoice_filepath,
            sheet_name=self.invoice_detail_sheet_name,
            usecols=[
                'SKU', '运输方式', '物流商单号', '店铺', '物流中心编码', '发货单号', '发货仓库（单据）', '物流商'
            ],
        )
        invoice_details_df['SKU'] = invoice_details_df['SKU'].astype(str)
        invoice_details_df['店铺'] = invoice_details_df['店铺'].str.split('[ -]').str[0].str.upper()
        # invoice_details_df.loc[invoice_details_df['店铺'] == "SIXPRAY", '店铺'] = "TOLESA"
        invoice_details_df = invoice_details_df[
            ['SKU', '运输方式', '物流商单号', '店铺', '物流中心编码', '发货单号', '发货仓库（单据）', '物流商']]
        invoice_details_df = invoice_details_df.ffill()
        # 如果发货仓库（单据）等于杭州虚拟仓，就为"工厂直发"，否则就为空
        invoice_details_df['发货仓库（单据）'] = invoice_details_df['发货仓库（单据）'].apply(
            lambda x: "工厂直发" if x == "杭州虚拟仓" else ""
        )

        # 合并数据
        result_df = invoice_df.merge(invoice_details_df, on=["SKU", "发货单号"], how="left")

        self._register_combined_cache[invoice_filepath] = result_df
        return result_df.copy()

    # 从下往上查找真正有数据的最后一行
    # def find_last_data_row(self, ws) -> int:
    #     for row in range(ws.max_row, 0, -1):
    #         # 检查该行是否有非空单元格
    #         for col in range(1, ws.max_column + 1):
    #             cell_value = ws.cell(row=row, column=col).value
    #             if cell_value is not None and str(cell_value).strip() != "":
    #                 return row
    #
    #     # 如果整个工作表都为空
    #     return 0

    def find_last_data_row(self, file_path, sheet_name):
        """使用pandas查找最后有数据的行"""
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

            # 删除全为空的行
            df_clean = df.dropna(how='all')

            if not df_clean.empty:
                return df_clean.last_valid_index() + 1
            else:
                return 0
        except Exception as e:
            print(f"读取文件时出错: {e}")
            return 0

    def verify_register_written(self, register_filepath: str, sheet_name: str,
                                column_name: str, invoice_number: str) -> None:
        """校验登记是否写入成功（根据发票号/合同号）"""
        try:
            register_df = pd.read_excel(register_filepath, sheet_name=sheet_name, usecols=[column_name])
        except Exception as e:
            raise Exception(f"登记校验失败：无法读取 {register_filepath} 的 {sheet_name} 表：{e}")

        col = register_df[column_name].astype(str).fillna("")
        needle = str(invoice_number).strip()
        pattern = re.escape(needle)
        found = (col == needle) | col.str.contains(pattern, na=False)
        if not found.any():
            raise Exception(
                f"登记校验失败：在 {os.path.basename(register_filepath)} 的 {sheet_name} 表 {column_name} 列未找到 {needle}"
            )

    def verify_register_row_count(self, register_filepath: str, sheet_name: str,
                                  expected_rows: int) -> None:
        """校验登记后的行数是否等于预计行数"""
        new_rows = self.find_last_data_row(register_filepath, sheet_name)
        if new_rows != expected_rows:
            raise Exception(
                f"登记校验失败：出运明细行数不匹配，预计行数:{expected_rows}，实际行数:{new_rows}"
            )

    def ensure_file_writable(self, file_path: str, label: str) -> None:
        """检查文件是否可写（避免被占用导致登记中断）"""
        if not os.path.isfile(file_path):
            raise Exception(f"{label}不存在：{file_path}")
        try:
            with open(file_path, "a"):
                pass
        except PermissionError:
            raise Exception(f"{label}正在被占用，无法写入：{file_path}")
        except Exception as e:
            raise Exception(f"{label}无法写入：{file_path}，原因：{e}")

    def register_shipping_invoice_number(self, invoice_filepath: str, invoice_number_register_filepath: str) -> int:
        """登记出货发票号登记表"""

        sheet_name = '出运明细'
        site = '亚马逊'

        result_df = self.register_combined_invoice(invoice_filepath)
        original_rows = len(result_df)
        result_df['发货日期'] = ""
        result_df['站点'] = site
        result_df['关税'] = ""
        result_df['VAT/消费税'] = ""
        result_df['货运费'] = ""

        # 如果是拼箱，处理result_df,拼箱的同一组数据中，箱数、总重量（kg）-箱子、总体积（m³）-箱子只保留第一行
        if self.lcl_flag:
            # 如果同一个发货单号里面，如果有重复的SKU，就重新计算发货数量，发货数量 = 箱数 * 单箱数量，否则就不处理
            result_df = self.update_shipment_quantity(result_df)

            # 标记需要处理的列
            cols_to_check = ['箱数', '箱子毛重（kg）', '箱子长度（cm）', '箱子宽度（cm）', '箱子高度（cm）', '总重量（kg）-箱子',
                             '总体积（m³）-箱子']

            # 按发货单号分组
            for shipment_id, group in result_df.groupby('发货单号'):
                # 在每组内进一步按箱数、重量、体积分组
                for _, subgroup in group.groupby(cols_to_check):
                    if len(subgroup) > 1:
                        # 保留第一行，其他行的指定列设置为NaN
                        indices_to_update = subgroup.index[1:]
                        result_df.loc[indices_to_update, cols_to_check] = None

        # 删除不需要的列
        result_df = result_df[
            ['发货日期', '物流商单号', '站点', '物流中心编码', '国家', '运输方式', '货件单号', '店铺',
             'SKU',
             '发货数量',
             '箱数',
             '总重量（kg）-箱子', '总体积（m³）-箱子', '关税', 'VAT/消费税', '货运费', '发货仓库（单据）', '物流商']]

        # 更换店铺，将 SIXPRAY替换为TOLESA
        result_df.loc[result_df['店铺'] == "SIXPRAY", '店铺'] = "TOLESA"

        deleted_rows = 0
        try:
            # 检查发票号是否已经存在,如果存在就报错
            column_name = '发票号'
            deleted_rows = self.check_invoice_number(
                invoice_number_register_filepath, self.shipping_details['invoice_number'],
                sheet_name, column_name
            )

            # raise  Exception

            # 找到工作表的最后一行
            start_row = self.find_last_data_row(invoice_number_register_filepath, sheet_name) + 1

            # 打开目标Excel文件
            workbook = load_workbook(invoice_number_register_filepath)
            worksheet = workbook[sheet_name]

            # 将DataFrame转换为行数据
            rows = list(dataframe_to_rows(result_df, index=False, header=False))

            # 如无可写入数据，直接报错便于排查
            if not rows:
                raise ValueError("登记表生成的数据为空，未写入任何行")

            original_rows = start_row - 1 + deleted_rows
            expected_total = original_rows - deleted_rows + len(rows)

            # 逐行写入数据，写完仅加边框
            for r_idx, row in enumerate(rows, start_row):
                for c_idx, value in enumerate(row, 1):
                    # worksheet.cell(row=r_idx, column=c_idx, value=value)
                    cell = worksheet.cell(row=r_idx, column=c_idx, value=value)
                    # 设置边框
                    cell.border = _THIN_BORDER
            # 保存文件
            workbook.save(invoice_number_register_filepath)

            # 从文件实际读取登记后总行数
            actual_total = self.find_last_data_row(invoice_number_register_filepath, sheet_name)
            log_message(
                f"登记出运明细 -> 文件: {os.path.basename(invoice_number_register_filepath)}, "
                f"Sheet: {sheet_name}, 原始行数: {original_rows}, 本次新增行数: {len(rows)}, "
                f"删除行数: {deleted_rows}, 预计登记后总行数: {expected_total}, 登记后总行数: {actual_total}, 发票号: {self.shipping_details['invoice_number']}"
            )

            # 校验登记是否写入成功
            self.verify_register_written(
                invoice_number_register_filepath, sheet_name, column_name, self.shipping_details['invoice_number']
            )
            return deleted_rows

        except Exception as e:
            # 如果是PermissionError，说明有人占用文件
            if isinstance(e, PermissionError):
                raise Exception(
                    f"登记失败：{invoice_number_register_filepath}！！！")
            else:
                raise Exception(
                    f"{e}")

    def build_invoicing_register_df(self, invoice_filepath: str, product_filepath: str,
                                    lumiya_register_filepath: str, easycan_register_filepath: str,
                                    sutong_register_filepath: str, liulang_register_filepath: str,
                                    jiangsheng_register_filepath: str) -> tuple[pd.DataFrame, str, str, str, int]:
        """构建报关开票登记数据与目标表信息"""

        # 设置默认值,报关渠道、报关方式、成交方式、供应商名称、币别
        customs_declaration_channels = "亚马逊"
        customs_declaration_method = 9710
        transaction_method = "CFR"
        supplier_name = "上海友谊卷笔刀有限公司"
        currency = "USD"

        # 读取产品信息单证专用
        product_df = self._read_excel_cached(
            product_filepath, sheet_name=self.product_sheet_name, usecols=['*SKU', '报关中文品名']
        )
        # 将*SKU转为字符串
        product_df['*SKU'] = product_df['*SKU'].astype(str)
        product_df = product_df[['*SKU', '报关中文品名']]
        product_df = product_df.rename(columns={"*SKU": "SKU"})

        # 读取发货单数据
        invoice_df = self.register_combined_invoice(invoice_filepath)
        original_rows = len(invoice_df)

        if self.lcl_flag:
            # 如果同一个发货单号里面，如果有重复的SKU，就重新计算发货数量，发货数量 = 箱数 * 单箱数量，否则就不处理
            invoice_df = self.update_shipment_quantity(invoice_df)

        # 合并数据
        merged_df = invoice_df.merge(product_df, on="SKU", how="left")
        merged_df["报关渠道"] = customs_declaration_channels
        merged_df["出口日期"] = ""
        merged_df["预录单编号"] = ""
        merged_df["报关方式（9710/一般贸易）"] = customs_declaration_method
        merged_df["成交方式(FOB/CIF/CFR)"] = transaction_method
        merged_df = merged_df[
            ['报关渠道', '物流商单号', '出口日期', '预录单编号', '报关方式（9710/一般贸易）', '成交方式(FOB/CIF/CFR)',
             '店铺', '报关中文品名', 'SKU', '发货数量']]

        # 读取报关资料表
        customs_declaration_df = self.result_df.copy()
        customs_declaration_df = customs_declaration_df[
            ['SKU', '报关中文品名', '英文品名', 'HS CODE', '发货数量', '单位', '申报单价', '总价']]
        customs_declaration_df['币别'] = currency
        # 重命名列
        customs_declaration_df.rename(columns={'发货数量': '数量', '报关中文品名': '开票/报关品名'}, inplace=True)

        # 合并数据,如果是素桐，就都用报关资料的数据，否则用发货单 + 报关资料的数据
        if self.company_info_index == 2:
            register_df = customs_declaration_df.merge(merged_df, on="SKU", how="left")
        else:
            # 将索引转成新的列
            merged_df.reset_index(inplace=True)
            merged_df['index'] = merged_df['index'].astype(int)

            customs_declaration_df.reset_index(inplace=True)
            customs_declaration_df['index'] = customs_declaration_df['index'].astype(int)
            customs_declaration_df.drop(columns=['SKU'], inplace=True)

            register_df = merged_df.merge(customs_declaration_df, on="index", how="left")
            # 删除index列
            register_df.drop(columns=['index'], inplace=True)

        # 店铺
        brand = merged_df['店铺'].unique().tolist()[0]
        sheet_name = '2026（内部开票）'

        # 根据店铺判断公司
        if brand in ['EZARC']:
            register_filepath = easycan_register_filepath
        elif brand in ['LIBRATON', 'PACMANN']:
            register_filepath = lumiya_register_filepath
        elif brand in ['SIXPRAY']:
            register_filepath = liulang_register_filepath
        elif brand in ['TOLESA']:
            register_filepath = jiangsheng_register_filepath
        else:
            register_filepath = sutong_register_filepath
            sheet_name = '2026'

            # 添加供应商名称
            register_df['供应商名称'] = supplier_name
            # 添加序号
            register_df['序号'] = ''
            register_df = register_df[
                ['序号', '物流商单号', '出口日期', '预录单编号', '报关方式（9710/一般贸易）', '成交方式(FOB/CIF/CFR)',
                 '店铺', 'SKU', '开票/报关品名', 'HS CODE', '数量', '单位', '申报单价', '总价', '供应商名称']]

        # 替换店铺名称
        register_df.loc[register_df['店铺'] == "SIXPRAY", '店铺'] = "TOLESA"
        return register_df, register_filepath, sheet_name, brand, original_rows

    def write_invoicing_register(self, register_df: pd.DataFrame, register_filepath: str,
                                 sheet_name: str, brand: str, original_rows: int, deleted_rows: int) -> None:
        """写入报关开票登记数据"""
        edited_rows = len(register_df)

        try:
            # 找到工作表的最后一行
            start_row = self.find_last_data_row(register_filepath, sheet_name) + 1

            # 打开目标Excel文件
            workbook = load_workbook(register_filepath)
            worksheet = workbook[sheet_name]

            # 将DataFrame转换为行数据
            rows = list(dataframe_to_rows(register_df, index=False, header=False))

            # 如无可写入数据，直接报错便于排查
            if not rows:
                raise ValueError("登记表生成的数据为空，未写入任何行")

            original_rows = start_row - 1 + deleted_rows
            expected_total = original_rows - deleted_rows + edited_rows

            # 处理写入区域内的合并单元格，避免写入MergedCell报错
            end_row = start_row + len(rows) - 1
            end_col = max(len(r) for r in rows)
            ranges_to_restore = []
            for merged_range in list(worksheet.merged_cells.ranges):
                # 判断合并区域是否与写入区域有交集
                if (merged_range.min_row <= end_row and merged_range.max_row >= start_row and
                        merged_range.min_col <= end_col and merged_range.max_col >= 1):
                    # 记录需要恢复的上半部分
                    if merged_range.min_row < start_row:
                        new_min_row = merged_range.min_row
                        new_max_row = start_row - 1
                        if (new_min_row < new_max_row) or (merged_range.min_col < merged_range.max_col):
                            ranges_to_restore.append(
                                (new_min_row, new_max_row, merged_range.min_col, merged_range.max_col))
                    # 记录需要恢复的下半部分
                    if merged_range.max_row > end_row:
                        new_min_row = end_row + 1
                        new_max_row = merged_range.max_row
                        if (new_min_row < new_max_row) or (merged_range.min_col < merged_range.max_col):
                            ranges_to_restore.append(
                                (new_min_row, new_max_row, merged_range.min_col, merged_range.max_col))
                    worksheet.unmerge_cells(str(merged_range))

            # 逐行写入数据；不复制上一行样式，后续统一加边框和合并
            for r_idx, row in enumerate(rows, start_row):
                for c_idx, value in enumerate(row, 1):
                    worksheet.cell(row=r_idx, column=c_idx, value=value)

            # 对指定行设置边框
            frame_data = {
                'min_row': start_row, 'max_row': worksheet.max_row, 'min_col': 1, 'max_col': 20
            }
            self.seeting_excel_format(worksheet, frame_data, "frame")

            # 合并发货日期和物流商单号列、货代公司
            if start_row <= worksheet.max_row:
                worksheet.merge_cells(start_row=start_row, start_column=2, end_row=worksheet.max_row, end_column=2)
                # 居中对齐
                cell = worksheet.cell(row=start_row, column=2)
                cell.alignment = _alignment(horizontal='center', vertical='center')

            # 恢复写入区域外被拆开的合并单元格
            for min_row, max_row, min_col, max_col in ranges_to_restore:
                worksheet.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)

            # 保存文件
            workbook.save(register_filepath)

            # 从文件实际读取登记后总行数
            actual_total = self.find_last_data_row(register_filepath, sheet_name)
            log_message(
                f"登记报关开票 -> 文件: {os.path.basename(register_filepath)}, "
                f"Sheet: {sheet_name}, 原始行数: {original_rows}, 本次新增行数: {edited_rows}, "
                f"删除行数: {deleted_rows}, 预计登记后总行数: {expected_total}, 登记后总行数: {actual_total}, 合同号: {self.shipping_details['invoice_number']}, 店铺: {brand}"
            )

            # 校验登记是否写入成功
            column_name = '合同号'
            self.verify_register_written(
                register_filepath, sheet_name, column_name, self.shipping_details['invoice_number']
            )

        except Exception as e:
            # 如果是PermissionError，说明有人占用文件
            if isinstance(e, PermissionError):
                raise Exception(
                    f"登记失败：{register_filepath}！！！")
            else:
                raise Exception(
                    f"{e}")

    def register_invoicing(self, invoice_filepath: str, product_filepath: str, lumiya_register_filepath: str,
                           easycan_register_filepath: str, sutong_register_filepath: str,
                           liulang_register_filepath: str, jiangsheng_register_filepath: str) -> None:
        """报关开票登记"""
        register_df, register_filepath, sheet_name, brand, original_rows = self.build_invoicing_register_df(
            invoice_filepath, product_filepath, lumiya_register_filepath, easycan_register_filepath,
            sutong_register_filepath, liulang_register_filepath, jiangsheng_register_filepath
        )
        self.write_invoicing_register(register_df, register_filepath, sheet_name, brand, original_rows)

    def register_all_documents(self, invoice_filepath: str, invoice_number_register_filepath: str,
                               product_filepath: str, lumiya_register_filepath: str, easycan_register_filepath: str,
                               sutong_register_filepath: str, liulang_register_filepath: str, jiangsheng_register_filepath: str) -> None:
        """登记所有数据"""
        print("登记数据")
        start_time = datetime.datetime.now()
        source_file = os.path.basename(invoice_filepath)
        try:
            # 先构建报关开票登记数据，确定目标文件并检查可写
            register_df, register_filepath, invoicing_sheet, brand, invoice_rows = self.build_invoicing_register_df(
                invoice_filepath, product_filepath, lumiya_register_filepath, easycan_register_filepath,
                sutong_register_filepath, liulang_register_filepath, jiangsheng_register_filepath
            )
            self.ensure_file_writable(register_filepath, "报关开票登记表")

            # 登记出货发票号登记表
            sheet_name = '出运明细'
            shipping_original_rows = self.find_last_data_row(invoice_number_register_filepath, sheet_name)
            insert_rows = len(self.register_combined_invoice(invoice_filepath))
            deleted_rows = self.register_shipping_invoice_number(
                invoice_filepath, invoice_number_register_filepath
            )
            expected_rows = shipping_original_rows - deleted_rows + insert_rows

            self.verify_register_row_count(invoice_number_register_filepath, sheet_name, expected_rows)

            # 报关开票登记
            register_original_rows = self.find_last_data_row(register_filepath, invoicing_sheet)
            deleted_rows = self.check_invoice_number(
                register_filepath, self.shipping_details['invoice_number'], invoicing_sheet, '合同号'
            )
            expected_rows = register_original_rows - deleted_rows + len(register_df)
            self.write_invoicing_register(register_df, register_filepath, invoicing_sheet, brand, invoice_rows, deleted_rows)
            self.verify_register_row_count(register_filepath, invoicing_sheet, expected_rows)
            append_call_record(
                module_name="发货单",
                action="登记数据",
                success=True,
                duration_seconds=(datetime.datetime.now() - start_time).total_seconds(),
                invoice_number=self.shipping_details.get('invoice_number', ''),
                source_file=source_file,
            )
        except Exception as e:
            append_call_record(
                module_name="发货单",
                action="登记数据",
                success=False,
                duration_seconds=(datetime.datetime.now() - start_time).total_seconds(),
                invoice_number=self.shipping_details.get('invoice_number', ''),
                source_file=source_file,
                error_message=str(e),
            )
            raise Exception(
                f"登记失败：{e}！！！")


if __name__ == "__main__":
    """主函数"""
    # 设置显示所有列
    pd.set_option('display.max_columns', None)
    # 设置显示所有行
    pd.set_option('display.max_rows', None)
    # 设置列宽
    pd.set_option('display.max_colwidth', None)
    # 设置显示宽度
    pd.set_option('display.width', None)

    """测试和调试代码"""
    generator = CustomsDeclaration()

    default_paths = [
        r"C:\Users\12091\Desktop\26EA065 发货单.xlsx",
        r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\产品信息单证专用.xlsx",
        r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\品名.xlsx",
        r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\全站点地址.xlsx",
        r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\要素.xlsx",
        r"C:/Users/12091/Desktop",
        r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\pictures",
        r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\测试\2026出运明细.xlsx",
        r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\测试\2025.9-鲁米亚报关明细.xlsx",
        r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\测试\2025.9-易达报关明细 修复.xlsx",
        r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\测试\2025-素桐报关开票明细.xlsx",
        r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\测试\2025-六浪报关明细.xlsx",
        r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\测试\2025-匠晟报关明细.xlsx"
    ]

    # 生成所有文档
    generator.generate_all_documents(
        default_paths[0], default_paths[1], default_paths[2],
        default_paths[3], default_paths[4], default_paths[5],
        default_paths[6], False)

    # 登记所有文档
    # generator.register_all_documents(default_paths[0], default_paths[7], default_paths[1], default_paths[8],
    #                                  default_paths[9],default_paths[10],default_paths[11],default_paths[12])

# pyinstaller --windowed --onedir --name "customs_declaration_info" --icon "\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\pictures\tool.ico" D:\Yida_project\物流部-数据处理工具\报关资料生成\ui_interfac
