#!/usr/bin/env python3
"""Phase 6 -- Enhance: recommend a Sigma Agent for the workflows the macros became.

  recommend-agent.py --macros macros.json --workflows workflows.json
  recommend-agent.py ... --emit fragment.json [--workbook wb-spec.json] [--name "FP&A forecast agent"]
  recommend-agent.py ... --merge-into wb-spec.json --out merged.json

--macros is the `macro-classify.py --json` output; --workflows is workflows.json.
Prints the recommendation (read it aloud: it is the on-camera moment). Never contacts Sigma and
never sends email. For a forecast build the agent reaches the workbook through
`build-forecast-model.py ... --agent-macros macros.json --post` (one POST); --emit / --merge-into
are for workbook specs that builder did not build. Get the user's yes before either.
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "lib"))
import agent_enhance as ae  # noqa: E402
import code_rep  # noqa: E402

GROUNDING = ae.GROUNDING


def load(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--macros", required=True)
    p.add_argument("--workflows", required=True)
    p.add_argument("--emit", metavar="FRAGMENT_JSON", help="write the agent + chat fragment here")
    p.add_argument("--workbook", metavar="WB_SPEC_JSON",
                   help="built workbook spec: grounds the agent in its real elements and reuses its button effects")
    p.add_argument("--merge-into", metavar="WB_SPEC_JSON", help="add the agent + chat to this workbook spec")
    p.add_argument("--out", help="merged workbook spec path (with --merge-into)")
    p.add_argument("--name", default="FP&A forecast agent")
    a = p.parse_args(argv)
    if a.merge_into and not a.out:
        p.error("--merge-into needs --out")

    wb_path = a.merge_into or a.workbook
    wb = load(wb_path) if wb_path else None
    rec, frag = ae.recommend_for_workbook(load(a.macros), load(a.workflows), wb, a.name,
                                          fragment=bool(a.emit or a.merge_into))
    print(rec["pitch"] if rec["recommend"] else "No agent recommended.")
    for r in rec["reasons"]:
        print("  - " + r)
    if not (a.emit or a.merge_into):
        return 0
    if not rec["recommend"]:
        print("Nothing to emit.", file=sys.stderr)
        return 1

    if a.emit:
        Path(a.emit).write_text(json.dumps(frag, indent=2) + "\n", encoding="utf-8")
        print(f"fragment written: {a.emit}")
    if a.merge_into:
        Path(a.out).write_text(json.dumps(ae.merge_fragment(wb, frag), indent=2) + "\n", encoding="utf-8")
        print(f"merged workbook spec written: {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
