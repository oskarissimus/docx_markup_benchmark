import argparse
import json
import sys
from pathlib import Path

from .evaluator import evaluate_documents
from .report import format_report


def _validate_paths(gt_path: Path, eval_path: Path, out_path: Path) -> None:
    if not gt_path.exists() or gt_path.suffix.lower() != ".docx":
        raise SystemExit(f"Invalid --gt path: {gt_path}")
    if not eval_path.exists() or eval_path.suffix.lower() != ".docx":
        raise SystemExit(f"Invalid --eval path: {eval_path}")
    if out_path.suffix.lower() not in {".json", ".csv", ".md"}:
        raise SystemExit(f"Invalid --out extension: {out_path.suffix}")
    out_dir = out_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        # touch-like writability check
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("")
    except Exception as exc:  # noqa: BLE001
        raise SystemExit(f"Cannot write to --out path: {out_path} ({exc})") from exc


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docx-markup-eval",
        description="Evaluate DOCX markup placement within tables",
    )
    parser.add_argument("--gt", required=True, help="Path to ground-truth .docx")
    parser.add_argument("--eval", required=True, help="Path to evaluated .docx")
    parser.add_argument(
        "--format",
        required=True,
        choices=["json", "csv", "md"],
        help="Output format",
    )
    parser.add_argument("--out", required=True, help="Output file path")
    parser.add_argument("--debug", action="store_true", help="Enable debug output")
    return parser


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    gt_path = Path(args.gt)
    eval_path = Path(args.eval)
    out_path = Path(args.out)
    _validate_paths(gt_path, eval_path, out_path)

    result = evaluate_documents(gt_path, eval_path, debug=args.debug)

    report_text = format_report(result, args.format)
    out_path.write_text(report_text, encoding="utf-8")

    # Optional debug print to stdout for ease of use
    if args.debug:
        try:
            print(json.dumps(result, indent=2))
        except Exception:  # noqa: BLE001
            print(report_text)


if __name__ == "__main__":
    main()


