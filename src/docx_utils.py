from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Tuple

from docx import Document  # type: ignore[import-not-found]

TOKEN_REGEX = re.compile(r"(?i)cell_\d+")


@dataclass(frozen=True)
class CellText:
    table_index: int
    row_index: int
    col_index: int
    merged_rect: tuple[int, int, int, int]  # (row_start, col_start, row_end, col_end)
    text: str


def _normalize_whitespace(text: str) -> str:
    # Convert NBSP to space and collapse whitespace
    text = text.replace("\u00A0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _iter_tables(document: Document):
    for idx, table in enumerate(document.tables):
        yield idx, table


def _get_merged_map(table) -> dict[Tuple[int, int], tuple[int, int]]:
    # Map each logical grid position (r,c) to its top-left owner using table.cell(r,c)
    owner: dict[Tuple[int, int], tuple[int, int]] = {}
    rows_count = len(table.rows)
    cols_count = len(table.columns) if rows_count > 0 else 0
    cell_owner_by_id: dict[int, tuple[int, int]] = {}
    for r_idx in range(rows_count):
        for c_idx in range(cols_count):
            cell = table.cell(r_idx, c_idx)
            key = id(cell)
            tl = cell_owner_by_id.get(key)
            if tl is None:
                tl = (r_idx, c_idx)
                cell_owner_by_id[key] = tl
            owner[(r_idx, c_idx)] = tl
    return owner


def _merged_rect(owner_map: dict[Tuple[int, int], tuple[int, int]], target: tuple[int, int]) -> tuple[int, int, int, int]:
    top_left = owner_map[target]
    rs, cs = top_left
    re_max, ce_max = rs, cs
    # Find bounding rect of all cells that share this owner
    for (r, c), tl in owner_map.items():
        if tl == top_left:
            re_max = max(re_max, r)
            ce_max = max(ce_max, c)
    return (rs, cs, re_max, ce_max)


def extract_table_cell_texts(doc_path: Path) -> list[CellText]:
    doc = Document(str(doc_path))
    results: list[CellText] = []
    for t_idx, table in _iter_tables(doc):
        owner_map = _get_merged_map(table)
        seen_owners: set[tuple[int, int]] = set()
        rows_count = len(table.rows)
        cols_count = len(table.columns) if rows_count > 0 else 0
        for r_idx in range(rows_count):
            for c_idx in range(cols_count):
                owner = owner_map[(r_idx, c_idx)]
                if owner in seen_owners:
                    continue
                seen_owners.add(owner)
                rect = _merged_rect(owner_map, (r_idx, c_idx))

                # Build combined text across positions in the merged rectangle
                rs, cs, re_idx, ce_idx = rect
                parts: list[str] = []
                for rr in range(rs, re_idx + 1):
                    for cc in range(cs, ce_idx + 1):
                        parts.append(table.cell(rr, cc).text)
                text = _normalize_whitespace("\n".join(parts))

                results.append(
                    CellText(
                        table_index=t_idx,
                        row_index=owner[0],
                        col_index=owner[1],
                        merged_rect=rect,
                        text=text,
                    )
                )
    return results


def find_tokens(text: str) -> list[tuple[int, int]]:
    # Return list of (start, end) indices for tokens
    matches: list[tuple[int, int]] = []
    for m in TOKEN_REGEX.finditer(text):
        matches.append((m.start(), m.end()))
    return matches


def strip_tokens(text: str) -> tuple[str, list[int]]:
    # Remove tokens and return (base_text, token_starts_in_base_coords)
    starts_in_base: list[int] = []
    out_chars: list[str] = []
    i = 0
    removed_so_far = 0
    for m in TOKEN_REGEX.finditer(text):
        # text[i:m.start()] remains in base
        out_chars.append(text[i : m.start()])
        # The start position of token in base equals current base length
        starts_in_base.append(sum(len(chunk) for chunk in out_chars))
        i = m.end()
    out_chars.append(text[i:])
    base_text = "".join(out_chars)
    return (base_text, starts_in_base)


