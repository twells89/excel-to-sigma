"""Run one shard of the unit-test suite.

Discovers the same tests as `python -m unittest discover -s <dir>`, orders them by
test id, and runs every Nth one so a CI matrix can run the suite in parallel without a
hand-kept module list (a new test file is picked up automatically). Every test lands in
exactly one shard; `--list` prints the ids so that can be checked.

    python .github/scripts/run_shard.py skills/excel-to-sigma/tests 2/4 [--list]
"""
import argparse
import sys
import unittest


def flatten(suite):
    for item in suite:
        if isinstance(item, unittest.TestSuite):
            yield from flatten(item)
        else:
            yield item


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("start_dir")
    parser.add_argument("shard", help="INDEX/TOTAL, 1-based, e.g. 2/4")
    parser.add_argument("--list", action="store_true", help="print this shard's test ids and exit")
    args = parser.parse_args()
    index, total = (int(part) for part in args.shard.split("/"))
    if not 1 <= index <= total:
        parser.error("shard must satisfy 1 <= INDEX <= TOTAL")
    # top_level_dir defaults to start_dir, matching `unittest discover -s`.
    suite = unittest.defaultTestLoader.discover(args.start_dir)
    tests = sorted(flatten(suite), key=lambda test: test.id())
    mine = [test for position, test in enumerate(tests) if position % total == index - 1]
    if args.list:
        print("\n".join(test.id() for test in mine))
        return 0
    print(f"shard {index}/{total}: {len(mine)} of {len(tests)} tests", flush=True)
    result = unittest.TextTestRunner(verbosity=2).run(unittest.TestSuite(mine))
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    sys.exit(main())
