from __future__ import annotations

import math


def calculate_wrapped_text_height(value: object, minimum_height: float,
                                  characters_per_line: int, line_height: float = 15,
                                  padding: float = 10) -> float:
    """Estimate a row height that leaves room for explicit and wrapped lines."""
    lines = str(value or "").splitlines() or [""]
    visual_line_count = sum(
        max(1, math.ceil(len(line) / characters_per_line)) for line in lines
    )
    return max(minimum_height, visual_line_count * line_height + padding)
