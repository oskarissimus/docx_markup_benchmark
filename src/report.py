from __future__ import annotations

import csv
import io
import json


def format_report(result: dict, fmt: str) -> str:
    fields = ["gt_total", "eval_total", "correct", "misplaced", "missed"]
    if fmt == "json":
        return json.dumps({k: result[k] for k in fields}, indent=2)
    if fmt == "csv":
        buf = io.StringIO()
        writer = csv.DictWriter(buf, fieldnames=fields)
        writer.writeheader()
        writer.writerow({k: result[k] for k in fields})
        return buf.getvalue()
    if fmt == "md":
        lines = ["| field | value |", "|---|---|"]
        for k in fields:
            lines.append(f"| {k} | {result[k]} |")
        return "\n".join(lines) + "\n"
    raise ValueError(f"Unsupported format: {fmt}")


