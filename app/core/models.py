from dataclasses import dataclass, field
from typing import Any


@dataclass
class ShippingContext:
    data: dict[str, Any] = field(default_factory=dict)


@dataclass
class RegisterTarget:
    filepath: str
    sheet_name: str


def select_register_filepath(company_info_index: int, lumiya_register_filepath: str,
                             easycan_register_filepath: str, sutong_register_filepath: str,
                             liulang_register_filepath: str, jiangsheng_register_filepath: str) -> str:
    register_filepaths = {
        0: easycan_register_filepath,
        1: lumiya_register_filepath,
        2: sutong_register_filepath,
        3: liulang_register_filepath,
        4: jiangsheng_register_filepath,
    }
    return register_filepaths.get(company_info_index, sutong_register_filepath)
