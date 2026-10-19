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
