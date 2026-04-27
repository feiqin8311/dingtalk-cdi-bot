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
