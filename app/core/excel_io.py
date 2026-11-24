from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path
from typing import Any

import pandas as pd
from pywintypes import error as WindowsError
import win32con
import win32file


@contextmanager
def workbook_write_lock(path: Path, label: str):
    """Serialize cooperating writers even while the destination is being replaced."""
    lock_path = path.with_name(f'.{path.name}.write.lock')
    try:
        handle = win32file.CreateFile(
            str(lock_path), win32con.GENERIC_WRITE | win32con.DELETE, 0, None,
            win32con.OPEN_ALWAYS, win32con.FILE_ATTRIBUTE_HIDDEN | win32con.FILE_FLAG_DELETE_ON_CLOSE, None,
        )
    except WindowsError as exc:
        raise RuntimeError(f'无法锁定{label}：可能正在写入，或没有共享目录写入权限，请重试') from exc
    try:
        yield
    finally:
        handle.Close()


class ExcelCache:
    def __init__(self) -> None:
        self._cache: dict[tuple[Any, ...], pd.DataFrame] = {}

    @staticmethod
    def _freeze(value: Any) -> Any:
        if isinstance(value, list):
            return tuple(value)
        if isinstance(value, dict):
            return tuple(sorted((k, ExcelCache._freeze(v)) for k, v in value.items()))
        if isinstance(value, set):
            return tuple(sorted(value))
        return value

    def read(self, file_path: str, sheet_name: str, **kwargs) -> pd.DataFrame:
        local_kwargs = dict(kwargs)
        dtype = local_kwargs.get("dtype")
        usecols = local_kwargs.get("usecols")
        if isinstance(dtype, dict) and isinstance(usecols, list):
            filtered_dtype = {k: v for k, v in dtype.items() if k in usecols}
            if filtered_dtype:
                local_kwargs["dtype"] = filtered_dtype
            else:
                local_kwargs.pop("dtype", None)

        cache_key = (
            file_path,
            sheet_name,
            tuple(sorted((k, self._freeze(v)) for k, v in local_kwargs.items())),
        )
        cached = self._cache.get(cache_key)
        if cached is None:
            cached = pd.read_excel(file_path, sheet_name=sheet_name, **local_kwargs)
            self._cache[cache_key] = cached
        return cached.copy()
