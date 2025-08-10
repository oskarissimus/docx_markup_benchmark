from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.evaluator import evaluate_documents  # noqa: E402
from src.report import format_report  # noqa: E402
from tests.helpers import (  # noqa: E402
    add_table,
    export_artifacts,
    merge,
    new_doc,
    save,
    set_cell_text,
)


FIXTURES = Path(__file__).parent / "fixtures" / "generated"


def test_multiple_tables_perfect(tmp_path: Path):
    gt = new_doc()
    ev = new_doc()

    # Table 1
    t1_gt = add_table(gt, 1, 2)
    t1_ev = add_table(ev, 1, 2)
    set_cell_text(t1_gt.cell(0, 0), "CELL_1")
    set_cell_text(t1_gt.cell(0, 1), "")
    set_cell_text(t1_ev.cell(0, 0), "CELL_1")
    set_cell_text(t1_ev.cell(0, 1), "")

    # Table 2
    t2_gt = add_table(gt, 1, 2)
    t2_ev = add_table(ev, 1, 2)
    set_cell_text(t2_gt.cell(0, 0), "CELL_2")
    set_cell_text(t2_gt.cell(0, 1), "CELL_3")
    set_cell_text(t2_ev.cell(0, 0), "CELL_2")
    set_cell_text(t2_ev.cell(0, 1), "CELL_3")

    p_gt, p_ev = tmp_path / "mt_perf_gt.docx", tmp_path / "mt_perf_ev.docx"
    save(gt, p_gt)
    save(ev, p_ev)

    out_dir = FIXTURES / "multiple_tables_perfect"
    export_artifacts(p_gt, out_dir, "gt")
    export_artifacts(p_ev, out_dir, "eval")

    res = evaluate_documents(p_gt, p_ev, debug=True)
    (out_dir / "report.md").write_text(format_report(res, "md"), encoding="utf-8")

    assert res["gt_total"] == 3
    assert res["eval_total"] == 3
    assert res["correct"] == 3
    assert res["missed"] == 0
    assert res["misplaced"] == 0


def test_multiple_tables_misplaced(tmp_path: Path):
    gt = new_doc()
    ev = new_doc()

    # Table 1
    t1_gt = add_table(gt, 1, 2)
    t1_ev = add_table(ev, 1, 2)
    set_cell_text(t1_gt.cell(0, 0), "CELL_1")
    set_cell_text(t1_gt.cell(0, 1), "")
    set_cell_text(t1_ev.cell(0, 0), "")
    set_cell_text(t1_ev.cell(0, 1), "CELL_999")  # misplaced

    # Table 2
    t2_gt = add_table(gt, 1, 2)
    t2_ev = add_table(ev, 1, 2)
    set_cell_text(t2_gt.cell(0, 0), "CELL_2")
    set_cell_text(t2_gt.cell(0, 1), "")
    set_cell_text(t2_ev.cell(0, 0), "CELL_2")
    set_cell_text(t2_ev.cell(0, 1), "")

    p_gt, p_ev = tmp_path / "mt_mis_gt.docx", tmp_path / "mt_mis_ev.docx"
    save(gt, p_gt)
    save(ev, p_ev)

    out_dir = FIXTURES / "multiple_tables_misplaced"
    export_artifacts(p_gt, out_dir, "gt")
    export_artifacts(p_ev, out_dir, "eval")

    res = evaluate_documents(p_gt, p_ev, debug=True)
    (out_dir / "report.md").write_text(format_report(res, "md"), encoding="utf-8")

    assert res["gt_total"] == 2
    assert res["eval_total"] == 2
    assert res["correct"] == 1
    assert res["missed"] == 1
    assert res["misplaced"] == 1


def test_merged_cells_correct_horizontal(tmp_path: Path):
    gt = new_doc()
    ev = new_doc()
    t1 = add_table(gt, 1, 3)
    t2 = add_table(ev, 1, 3)

    # Set text before merge so merged content is preserved
    set_cell_text(t1.cell(0, 0), "A CELL_1 B")
    set_cell_text(t1.cell(0, 1), "")
    set_cell_text(t1.cell(0, 2), "")

    set_cell_text(t2.cell(0, 0), "A cell_999 B")
    set_cell_text(t2.cell(0, 1), "")
    set_cell_text(t2.cell(0, 2), "")

    # Merge first two cells horizontally in both docs
    merge(t1, 0, 0, 0, 1)
    merge(t2, 0, 0, 0, 1)

    p_gt, p_ev = tmp_path / "mc_h_gt.docx", tmp_path / "mc_h_ev.docx"
    save(gt, p_gt)
    save(ev, p_ev)

    out_dir = FIXTURES / "merged_cells_horizontal_correct"
    export_artifacts(p_gt, out_dir, "gt")
    export_artifacts(p_ev, out_dir, "eval")

    res = evaluate_documents(p_gt, p_ev, debug=True)
    (out_dir / "report.md").write_text(format_report(res, "md"), encoding="utf-8")

    assert res["gt_total"] == 1
    assert res["eval_total"] == 1
    assert res["correct"] == 1
    assert res["missed"] == 0
    assert res["misplaced"] == 0


def test_merged_cells_wrong_position_across_parts(tmp_path: Path):
    gt = new_doc()
    ev = new_doc()
    t1 = add_table(gt, 2, 1)
    t2 = add_table(ev, 2, 1)

    # GT: token in top part; Eval: token in bottom part
    set_cell_text(t1.cell(0, 0), "Top CELL_1")
    set_cell_text(t1.cell(1, 0), "Bottom")

    set_cell_text(t2.cell(0, 0), "Top")
    set_cell_text(t2.cell(1, 0), "Bottom cell_9")

    # Vertical merge full column (after setting text)
    merge(t1, 0, 0, 1, 0)
    merge(t2, 0, 0, 1, 0)

    p_gt, p_ev = tmp_path / "mc_v_gt.docx", tmp_path / "mc_v_ev.docx"
    save(gt, p_gt)
    save(ev, p_ev)

    out_dir = FIXTURES / "merged_cells_wrong_position"
    export_artifacts(p_gt, out_dir, "gt")
    export_artifacts(p_ev, out_dir, "eval")

    res = evaluate_documents(p_gt, p_ev, debug=True)
    (out_dir / "report.md").write_text(format_report(res, "md"), encoding="utf-8")

    assert res["gt_total"] == 1
    assert res["eval_total"] == 1
    assert res["correct"] == 0
    assert res["missed"] == 1
    assert res["misplaced"] == 1


