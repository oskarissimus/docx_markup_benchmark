from __future__ import annotations

import sys
from pathlib import Path

import pytest

# Allow importing src and helpers
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.evaluator import evaluate_documents  # noqa: E402
from src.report import format_report  # noqa: E402
from tests.helpers import (  # noqa: E402
    add_runs,
    add_table,
    export_artifacts,
    new_doc,
    save,
    set_cell_text,
)

FIXTURES = Path(__file__).parent / "fixtures" / "generated"


def scenario_1(tmp: Path) -> tuple[Path, Path]:
    # Perfect alignment (2×2, one token per cell)
    gt = new_doc()
    ev = new_doc()
    t1 = add_table(gt, 2, 2)
    t2 = add_table(ev, 2, 2)
    tokens = [["CELL_1", "CELL_2"], ["CELL_3", "CELL_4"]]
    for r in range(2):
        for c in range(2):
            set_cell_text(t1.cell(r, c), tokens[r][c])
            set_cell_text(t2.cell(r, c), tokens[r][c])
    p_gt, p_ev = tmp / "s1_gt.docx", tmp / "s1_ev.docx"
    save(gt, p_gt)
    save(ev, p_ev)
    return p_gt, p_ev


def scenario_2(tmp: Path) -> tuple[Path, Path]:
    # Missing one token
    gt = new_doc()
    ev = new_doc()
    t1 = add_table(gt, 2, 2)
    t2 = add_table(ev, 2, 2)
    toks = [["CELL_1", "CELL_2"], ["CELL_3", "CELL_4"]]
    for r in range(2):
        for c in range(2):
            set_cell_text(t1.cell(r, c), toks[r][c])
            set_cell_text(t2.cell(r, c), toks[r][c])
    # Remove one token in eval
    set_cell_text(t2.cell(1, 1), "")
    p_gt, p_ev = tmp / "s2_gt.docx", tmp / "s2_ev.docx"
    save(gt, p_gt)
    save(ev, p_ev)
    return p_gt, p_ev


def scenario_3(tmp: Path) -> tuple[Path, Path]:
    # Extra token in new cell
    gt = new_doc()
    ev = new_doc()
    t1 = add_table(gt, 2, 2)
    t2 = add_table(ev, 2, 2)
    set_cell_text(t1.cell(0, 0), "CELL_1")
    set_cell_text(t1.cell(0, 1), "CELL_2")
    set_cell_text(t1.cell(1, 0), "")
    set_cell_text(t1.cell(1, 1), "")

    set_cell_text(t2.cell(0, 0), "CELL_1")
    set_cell_text(t2.cell(0, 1), "CELL_2")
    set_cell_text(t2.cell(1, 0), "CELL_3")  # extra
    set_cell_text(t2.cell(1, 1), "")

    p_gt, p_ev = tmp / "s3_gt.docx", tmp / "s3_ev.docx"
    save(gt, p_gt)
    save(ev, p_ev)
    return p_gt, p_ev


def scenario_4(tmp: Path) -> tuple[Path, Path]:
    # Wrong position in same cell
    gt = new_doc()
    ev = new_doc()
    t1 = add_table(gt, 1, 1)
    t2 = add_table(ev, 1, 1)
    set_cell_text(t1.cell(0, 0), "foo CELL_1 bar")
    # Move token position
    set_cell_text(t2.cell(0, 0), "foo bar CELL_1")
    p_gt, p_ev = tmp / "s4_gt.docx", tmp / "s4_ev.docx"
    save(gt, p_gt)
    save(ev, p_ev)
    return p_gt, p_ev


def scenario_5(tmp: Path) -> tuple[Path, Path]:
    # Multiple tokens in one cell, some missing
    gt = new_doc()
    ev = new_doc()
    t1 = add_table(gt, 1, 1)
    t2 = add_table(ev, 1, 1)
    set_cell_text(t1.cell(0, 0), "a CELL_1 b CELL_2 c CELL_3")
    set_cell_text(t2.cell(0, 0), "a CELL_1 b c")  # only CELL_1 present
    p_gt, p_ev = tmp / "s5_gt.docx", tmp / "s5_ev.docx"
    save(gt, p_gt)
    save(ev, p_ev)
    return p_gt, p_ev


def scenario_6(tmp: Path) -> tuple[Path, Path]:
    # Split-runs token + a miss
    gt = new_doc()
    ev = new_doc()
    t1 = add_table(gt, 1, 2)
    t2 = add_table(ev, 1, 2)
    # GT: left cell has split token across runs, right has a token
    add_runs(t1.cell(0, 0), ["CE", "LL", "_", "1"])  # CELL_1 split
    set_cell_text(t1.cell(0, 1), "CELL_2")
    # Eval: left token present, right missing
    add_runs(t2.cell(0, 0), ["cE", "Ll", "_", "1"])  # mixed case
    set_cell_text(t2.cell(0, 1), "")
    p_gt, p_ev = tmp / "s6_gt.docx", tmp / "s6_ev.docx"
    save(gt, p_gt)
    save(ev, p_ev)
    return p_gt, p_ev


def scenario_7(tmp: Path) -> tuple[Path, Path]:
    # Case-insensitive match (same position)
    gt = new_doc()
    ev = new_doc()
    t1 = add_table(gt, 1, 1)
    t2 = add_table(ev, 1, 1)
    set_cell_text(t1.cell(0, 0), "CELL_10")
    set_cell_text(t2.cell(0, 0), "cell_999")
    p_gt, p_ev = tmp / "s7_gt.docx", tmp / "s7_ev.docx"
    save(gt, p_gt)
    save(ev, p_ev)
    return p_gt, p_ev


def scenario_8(tmp: Path) -> tuple[Path, Path]:
    # Lorem buried mid-paragraph (same position)
    gt = new_doc()
    ev = new_doc()
    t1 = add_table(gt, 1, 1)
    t2 = add_table(ev, 1, 1)
    lorem = (
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin non lorem. "
    )
    set_cell_text(t1.cell(0, 0), f"{lorem}CELL_1 {lorem}")
    set_cell_text(t2.cell(0, 0), f"{lorem}CeLl_999 {lorem}")
    p_gt, p_ev = tmp / "s8_gt.docx", tmp / "s8_ev.docx"
    save(gt, p_gt)
    save(ev, p_ev)
    return p_gt, p_ev


def scenario_9(tmp: Path) -> tuple[Path, Path]:
    # Lorem in wrong cell
    gt = new_doc()
    ev = new_doc()
    t1 = add_table(gt, 1, 2)
    t2 = add_table(ev, 1, 2)
    set_cell_text(t1.cell(0, 0), "Lorem CELL_1 ipsum")
    set_cell_text(t1.cell(0, 1), "dolor sit amet")
    # Eval puts token in right cell instead of left
    set_cell_text(t2.cell(0, 0), "Lorem ipsum")
    set_cell_text(t2.cell(0, 1), "dolor CELL_999 sit amet")
    p_gt, p_ev = tmp / "s9_gt.docx", tmp / "s9_ev.docx"
    save(gt, p_gt)
    save(ev, p_ev)
    return p_gt, p_ev


def scenario_10(tmp: Path) -> tuple[Path, Path]:
    # NBSP/adjacency noise
    gt = new_doc()
    ev = new_doc()
    t1 = add_table(gt, 1, 2)
    t2 = add_table(ev, 1, 2)
    set_cell_text(t1.cell(0, 0), "\u00A0CELL_1,adjacent")  # NBSP then token, followed by text
    set_cell_text(t1.cell(0, 1), "(CELL_2)")  # punctuation
    set_cell_text(t2.cell(0, 0), " CELL_999,adjacent")  # space normalized
    set_cell_text(t2.cell(0, 1), "(cell_7)")
    p_gt, p_ev = tmp / "s10_gt.docx", tmp / "s10_ev.docx"
    save(gt, p_gt)
    save(ev, p_ev)
    return p_gt, p_ev


SCENARIOS = {
    1: scenario_1,
    2: scenario_2,
    3: scenario_3,
    4: scenario_4,
    5: scenario_5,
    6: scenario_6,
    7: scenario_7,
    8: scenario_8,
    9: scenario_9,
    10: scenario_10,
}


@pytest.mark.parametrize("scenario_id", list(SCENARIOS.keys()))
def test_scenarios(tmp_path: Path, scenario_id: int):
    out_dir = FIXTURES / f"scenario_{scenario_id}"
    out_dir.mkdir(parents=True, exist_ok=True)

    gt_p, ev_p = SCENARIOS[scenario_id](tmp_path)

    # Persist artifacts
    export_artifacts(gt_p, out_dir, "gt")
    export_artifacts(ev_p, out_dir, "eval")

    res = evaluate_documents(gt_p, ev_p, debug=True)

    # Save markdown report alongside artifacts for reference
    (out_dir / "report.md").write_text(format_report(res, "md"), encoding="utf-8")

    if scenario_id == 1:
        assert res["gt_total"] == 4
        assert res["eval_total"] == 4
        assert res["correct"] == 4
        assert res["missed"] == 0
        assert res["misplaced"] == 0

    if scenario_id == 2:
        assert res["gt_total"] == 4
        assert res["eval_total"] == 3
        assert res["correct"] == 3
        assert res["missed"] == 1
        assert res["misplaced"] == 0

    if scenario_id == 3:
        assert res["gt_total"] == 2
        assert res["eval_total"] == 3
        assert res["correct"] == 2
        assert res["missed"] == 0
        assert res["misplaced"] == 1

    if scenario_id == 4:
        assert res["gt_total"] == 1
        assert res["eval_total"] == 1
        # Moved position => not correct
        assert res["correct"] == 0
        assert res["missed"] == 1
        assert res["misplaced"] == 1

    if scenario_id == 5:
        assert res["gt_total"] == 3
        assert res["eval_total"] == 1
        assert res["correct"] == 1
        assert res["missed"] == 2
        assert res["misplaced"] == 0

    if scenario_id == 6:
        assert res["gt_total"] == 2
        assert res["eval_total"] == 1
        assert res["correct"] == 1
        assert res["missed"] == 1
        assert res["misplaced"] == 0

    if scenario_id == 7:
        assert res["gt_total"] == 1
        assert res["eval_total"] == 1
        assert res["correct"] == 1
        assert res["missed"] == 0
        assert res["misplaced"] == 0

    if scenario_id == 8:
        assert res["gt_total"] == 1
        assert res["eval_total"] == 1
        assert res["correct"] == 1
        assert res["missed"] == 0
        assert res["misplaced"] == 0

    if scenario_id == 9:
        assert res["gt_total"] == 1
        assert res["eval_total"] == 1
        assert res["correct"] == 0
        assert res["missed"] == 1
        assert res["misplaced"] == 1

    if scenario_id == 10:
        assert res["gt_total"] == 2
        assert res["eval_total"] == 2
        assert res["correct"] == 2
        assert res["missed"] == 0
        assert res["misplaced"] == 0


