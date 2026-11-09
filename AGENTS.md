# Excel to Sigma agent entry

Use [`skills/excel-to-sigma/SKILL.md`](skills/excel-to-sigma/SKILL.md) when the
user wants to inventory, assess, or migrate an Excel `.xlsx` or `.xlsm` planning
or analytical workbook into Sigma. The skill owns the workflow; read its
applicable `refs/` files before running its `scripts/` from the skill directory.

The same skill folder is distributed through the Claude Code plugin and can be
read directly by Codex, Cursor, Cortex Code, and other coding agents. Its local
`SKILL.md`, `refs/`, `scripts/`, and `requirements.txt` are the canonical files.
The agent's skill loader or the user supplies the Sigma tools and credentials;
this repository does not install a Sigma MCP server.

For installation and prerequisites, see [README.md](README.md#install).

## Cursor Cloud specific instructions

Install from the repository root with `python3 -m pip install --user --disable-pip-version-check -r requirements.txt` (Python 3.12, idempotent, user site-packages). There is no dev server.

Run the offline sample from `skills/excel-to-sigma/scripts` as in [QUICKSTART.md](skills/excel-to-sigma/QUICKSTART.md) (`make-sample-forecast.py`, then `audit-xlsx.py "Sample Forecast.xlsx" --allow-review`). Packaging checks are `python3 -m unittest discover -s tests -v`. The unit suite is `python3 .github/scripts/run_shard.py skills/excel-to-sigma/tests 1/1` (CI splits the same discovery into four shards). `audit-openapi.py` fetches Sigma's public OpenAPI from `assets.sigmacomputing.com` and needs network. Live `--post` needs Sigma credentials.

`skills/excel-to-sigma/tests/test_regional_canary_inputs.py` resolves the repo with `Path(__file__).resolve().parents[5]`. That index does not exist when the checkout is `/workspace`, so the module fails to import and `test_native_exact_plan.py` fails with it. The named `work/` fixtures are not in the tree.
