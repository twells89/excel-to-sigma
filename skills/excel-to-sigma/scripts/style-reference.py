#!/usr/bin/env python3
"""Read the look of an existing Sigma workbook ("make it look like Golden Orders") into a small JSON file.

usage: style-reference.py <workbookId | workbook URL> --out look.json [--sigma-profile P]
       style-reference.py --spec saved-get-response.json --out look.json     (offline: a saved GET response)

Read-only: one ``GET /v2/workbooks/{id}/spec`` through ``lib/sigma_auth.api``; nothing is ever written to Sigma.
The look keeps the theme, card styles, KPI type, text type and table preset only - not the arrangement.
Pass it to ``draft-table-report.py --look look.json``; Excel's own number formats, header fill, row order and
titles still win.
"""
import argparse
import json
import sys
from pathlib import Path

_LIB = Path(__file__).resolve().parent / "lib"
if str(_LIB) not in sys.path:
    sys.path.insert(0, str(_LIB))
import sigma_auth  # noqa: E402
from reference_look import extract_look, workbook_id  # noqa: E402


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument("workbook", nargs="?", help="workbook id or URL")
    parser.add_argument("--spec", help="a saved GET /v2/workbooks/{id}/spec response (no network)")
    parser.add_argument("--out", required=True, help="look JSON to write")
    sigma_auth.add_cli_args(parser)
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    if args.spec:
        response = json.loads(Path(args.spec).read_text(encoding="utf-8"))
    elif args.workbook:
        wid = workbook_id(args.workbook)
        sigma_auth.apply_cli_args(args)
        response = sigma_auth.api("GET", f"/v2/workbooks/{wid}/spec")
    else:
        print("give a workbook id / URL, or --spec", file=sys.stderr)
        return 2
    look = extract_look(response)
    look["source"] = {"workbookId": response.get("workbookId") or (workbook_id(args.workbook or "")),
                      "name": response.get("name")}
    out = Path(args.out).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(look, indent=1), encoding="utf-8")
    print(f"look -> {out}")
    theme = look["theme"] or {}
    print(f"  from {look['source']['name']!r} ({look['source']['workbookId']}): theme "
          f"{theme.get('name') or ('overrides' if theme.get('overrides') else 'none')}, "
          f"card styles {[k for k, v in look['styles'].items() if v] or 'none'}, "
          f"KPI type {'yes' if look['kpi']['name'] or look['kpi']['value'] else 'default'}, "
          f"table preset {look['tablePreset'] or 'default'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
