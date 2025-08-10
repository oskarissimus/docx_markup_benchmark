from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple

from difflib import SequenceMatcher

from .docx_utils import CellText, extract_table_cell_texts, strip_tokens


@dataclass
class CellEvaluation:
    table_index: int
    row_index: int
    col_index: int
    merged_rect: tuple[int, int, int, int]
    gt_positions: list[int]
    eval_positions: list[int]
    mapped_eval_positions: list[int]
    correct: int
    missed: int
    misplaced: int


def _map_positions(gt_base: str, eval_base: str, gt_positions: list[int]) -> list[int]:
    if gt_base == eval_base:
        return gt_positions.copy()

    matcher = SequenceMatcher(a=gt_base, b=eval_base, autojunk=False)
    opcodes = matcher.get_opcodes()

    def map_index(i: int) -> int:
        # Map base-text index i from GT to Eval using piecewise-linear mapping on equal blocks
        total_a = 0
        total_b = 0
        for tag, i1, i2, j1, j2 in opcodes:
            if tag == "equal":
                if i1 <= i <= i2:
                    # clamp within the equal segment
                    offset = max(0, min(i - i1, i2 - i1))
                    return j1 + offset
        # If not within an equal block, anchor to the closest following mapping
        # Find next block after i
        next_blocks = [(i1, j1) for tag, i1, i2, j1, j2 in opcodes if i < i1]
        if next_blocks:
            next_i1, next_j1 = next_blocks[0]
            return next_j1
        # Otherwise map to end
        return len(eval_base)

    return [map_index(pos) for pos in gt_positions]


def _evaluate_cells(gt_cells: list[CellText], eval_cells: list[CellText], debug: bool) -> tuple[list[CellEvaluation], dict]:
    evaluations: list[CellEvaluation] = []
    totals = {"gt_total": 0, "eval_total": 0, "correct": 0, "misplaced": 0, "missed": 0}

    # Pair cells by table order and merged top-left coordinates
    key = lambda c: (c.table_index, c.row_index, c.col_index)
    gt_index: dict[tuple[int, int, int], CellText] = {key(c): c for c in gt_cells}
    eval_index: dict[tuple[int, int, int], CellText] = {key(c): c for c in eval_cells}

    all_keys = sorted(set(gt_index.keys()) | set(eval_index.keys()))

    for k in all_keys:
        gt_cell = gt_index.get(k)
        ev_cell = eval_index.get(k)
        gt_text = gt_cell.text if gt_cell else ""
        ev_text = ev_cell.text if ev_cell else ""

        gt_base, gt_positions = strip_tokens(gt_text)
        ev_base, ev_positions = strip_tokens(ev_text)

        mapped_positions = _map_positions(gt_base, ev_base, gt_positions)

        # Correct if mapped position is present in eval positions set (already base coords)
        ev_set = set(ev_positions)
        correct = sum(1 for p in mapped_positions if p in ev_set)
        missed = len(gt_positions) - correct
        misplaced = len(ev_positions) - correct

        totals["gt_total"] += len(gt_positions)
        totals["eval_total"] += len(ev_positions)
        totals["correct"] += correct
        totals["missed"] += missed
        totals["misplaced"] += misplaced

        evaluations.append(
            CellEvaluation(
                table_index=k[0],
                row_index=k[1],
                col_index=k[2],
                merged_rect=(gt_cell or ev_cell).merged_rect if (gt_cell or ev_cell) else (0, 0, 0, 0),
                gt_positions=gt_positions,
                eval_positions=ev_positions,
                mapped_eval_positions=mapped_positions,
                correct=correct,
                missed=missed,
                misplaced=misplaced,
            )
        )

    return evaluations, totals


def evaluate_documents(gt_path: Path, eval_path: Path, debug: bool = False) -> dict:
    gt_cells = extract_table_cell_texts(gt_path)
    ev_cells = extract_table_cell_texts(eval_path)

    evaluations, totals = _evaluate_cells(gt_cells, ev_cells, debug)

    result: dict = {
        "gt_total": totals["gt_total"],
        "eval_total": totals["eval_total"],
        "correct": totals["correct"],
        "misplaced": totals["misplaced"],
        "missed": totals["missed"],
    }

    if debug:
        # Include per-cell details
        result["cells"] = [
            {
                "table": e.table_index,
                "row": e.row_index,
                "col": e.col_index,
                "rect": e.merged_rect,
                "gt_positions": e.gt_positions,
                "eval_positions": e.eval_positions,
                "mapped_from_gt": e.mapped_eval_positions,
                "correct": e.correct,
                "missed": e.missed,
                "misplaced": e.misplaced,
            }
            for e in evaluations
        ]

    # Sanity checks
    assert result["correct"] + result["missed"] == result["gt_total"]
    assert result["correct"] + result["misplaced"] == result["eval_total"]

    return result


