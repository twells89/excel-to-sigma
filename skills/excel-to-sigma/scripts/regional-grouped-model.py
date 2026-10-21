#!/usr/bin/env python3
"""Build the source-bound Regional grouped plan or full assigned-output packet.

Offline only: this script never creates or updates Sigma objects.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import sys

sys.path[:0] = [str(Path(__file__).resolve().parent),
                str(Path(__file__).resolve().parent / "lib")]

from lib.regional_grouped_model import build_grouped_model
from lib.regional_grouped_output import build_full_output_graph, validate_full_assigned_rows
from lib.source_inputs import validate_source_input_plan


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--input-plan", type=Path, required=True)
    parser.add_argument("--native-plan", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--canary", type=Path)
    parser.add_argument("--current", type=Path,
                        help="fresh full-document readback for all five input tables")
    parser.add_argument("--observed-inputs", type=Path,
                        help="fresh five-input query snapshots with physical row IDs")
    args = parser.parse_args()
    sources = [p for p in (args.source, args.input_plan, args.native_plan,
                           args.canary, args.current, args.observed_inputs) if p]
    for source in sources:
        if args.out.resolve() == source.resolve() or (args.out.exists() and source.exists()
                                                      and os.path.samefile(args.out, source)):
            raise ValueError("Regional grouped output aliases source artifact")
    inputs = json.loads(args.input_plan.read_text())
    native = json.loads(args.native_plan.read_text())
    source_hash = hashlib.sha256(args.source.read_bytes()).hexdigest()
    if source_hash != inputs.get("sourceSha256") or source_hash != native.get("sourceSha256"):
        raise ValueError("Regional source bytes/hash changed")
    validate_source_input_plan(inputs, args.source)
    model = build_grouped_model(inputs, native)
    if args.canary or args.current or args.observed_inputs:
        if not (args.canary and args.current and args.observed_inputs):
            raise ValueError("full packet requires canary, current document, and five observed inputs")
        current = json.loads(args.current.read_text())
        observed = json.loads(args.observed_inputs.read_text())
        assigned = validate_full_assigned_rows(inputs, current, observed)
        packet = build_full_output_graph(model, json.loads(args.canary.read_text()), assigned)
    else:
        packet = model
    args.out.parent.mkdir(parents=True, exist_ok=True)
    for source in sources:
        if args.out.resolve() == source.resolve() or (args.out.exists() and source.exists()
                                                      and os.path.samefile(args.out, source)):
            raise ValueError("Regional grouped output aliases source artifact")
    args.out.write_text(json.dumps(packet, indent=2, ensure_ascii=False) + "\n")
    print(json.dumps({"status": packet["status"], "formulaCount": packet.get("sourceFormulaCount", packet.get("formulaCount")),
                      "constructionBudget": packet.get("constructionBudget"), "out": str(args.out)}))


if __name__ == "__main__":
    main()
