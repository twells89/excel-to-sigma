"""Static, source-bound VBA behavior inventory.

The analyzer never executes VBA.  Exact module/procedure hashes are provenance;
the comment-stripped view is used only to identify conservative behavior facts.
"""
from __future__ import annotations

import copy
import hashlib
import re
from pathlib import Path


class MacroBehaviorError(ValueError):
    pass


_PROC = re.compile(
    r"^[ \t]*(?P<signature>(?:(?:Public|Private|Friend)[ \t]+)?"
    r"(?:(?:Static)[ \t]+)?(?P<kind>Sub|Function)[ \t]+"
    r"(?P<name>[A-Za-z_][A-Za-z0-9_]*)[ \t]*\([^\r\n]*\))",
    re.I | re.M,
)
_END_PROC = re.compile(r"^[ \t]*End[ \t]+(?:Sub|Function)[ \t]*$", re.I | re.M)
_MODULE_NAME = re.compile(r'^Attribute[ \t]+VB_Name[ \t]*=[ \t]*"([^"]+)"', re.I | re.M)
_CALL = re.compile(r"\bCall[ \t]+([A-Za-z_][A-Za-z0-9_]*)\b", re.I)
_STATIC_RANGE = re.compile(
    r'(?:Sheets|Worksheets)\([ \t]*"([^"]+)"[ \t]*\)\.Range\([ \t]*"([^"]+)"[ \t]*\)',
    re.I,
)
_DYNAMIC_RANGE = re.compile(r"(?:Sheets|Worksheets)\([^\"\r\n][^\)]*\)\.Range\(", re.I)
_CELL_WRITE = re.compile(
    r"(?P<target>(?:(?:[A-Za-z_][A-Za-z0-9_]*|(?:Sheets|Worksheets)\([^\)]*\))\.)?"
    r"Cells\([ \t]*(?P<row>[^,\)]+)[ \t]*,[ \t]*(?P<column>[^\)]+)\))"
    r"\.Value[ \t]*=[ \t]*(?P<value>[^\r\n]+)",
    re.I,
)
_RANGE_WRITE = re.compile(
    r"(?P<target>(?:Sheets|Worksheets)\([ \t]*\"(?P<sheet>[^\"]+)\"[ \t]*\)\."
    r"Range\([ \t]*\"(?P<address>[^\"]+)\"[ \t]*\))\.Value[ \t]*=[ \t]*(?P<value>[^\r\n]+)",
    re.I,
)


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _strip_comments(source: str) -> str:
    """Remove apostrophe and Rem comments without changing line count."""
    result = []
    for line in source.splitlines(keepends=True):
        content = line.rstrip("\r\n")
        ending = line[len(content):]
        if re.match(r"^[ \t]*Rem(?:[ \t]|$)", content, re.I):
            result.append(ending)
            continue
        quoted = False
        cut = len(content)
        index = 0
        while index < len(content):
            if content[index] == '"':
                if quoted and index + 1 < len(content) and content[index + 1] == '"':
                    index += 2
                    continue
                quoted = not quoted
            elif content[index] == "'" and not quoted:
                cut = index
                break
            index += 1
        result.append(content[:cut] + ending)
    return "".join(result)


def _mask_strings(source):
    return re.sub(r'"(?:""|[^"])*"', lambda m:' '*len(m[0]), source)


def _code_matches(pattern, source, flags=0):
    masked=_mask_strings(source)
    for match in re.finditer(pattern,source,flags):
        if masked[match.start():match.start()+1].strip():
            yield match


def _extract_source(path: Path) -> tuple[bytes, str, dict]:
    raw = path.read_bytes()
    if path.suffix.lower() in {".bas", ".cls", ".frm", ".vba", ".txt"}:
        return raw, raw.decode("utf-8", errors="replace"), {}
    try:
        from oletools.olevba import VBA_Parser
    except ImportError as error:  # pragma: no cover - environment dependent
        raise MacroBehaviorError("olevba is required for an Office VBA container") from error
    parser = VBA_Parser(str(path))
    if not parser.detect_vba_macros():
        return raw, "", {"containerByteSha256": _sha256(raw)}
    modules = [(name, code) for _file, _stream, name, code in parser.extract_macros()]
    extracted = "\n".join(code for _name, code in modules)
    return extracted.encode("utf-8"), extracted, {
        "containerByteSha256": _sha256(raw),
        "extractedModules": [name for name, _code in modules],
    }


def _procedure_slices(source: str) -> list[dict]:
    lines = source.splitlines(keepends=True)
    starts = []
    offset = 0
    offsets = []
    for line_number, line in enumerate(lines, 1):
        offsets.append(offset)
        match = _PROC.match(line)
        if match:
            starts.append((line_number, offset, match))
        offset += len(line)
    procedures = []
    for line_number, start, match in starts:
        tail = source[start:]
        end_match = _END_PROC.search(tail)
        if end_match is None:
            raise MacroBehaviorError(f"unterminated procedure: {match.group('name')}")
        end = start + end_match.end()
        while end < len(source) and source[end] in "\r\n":
            end += 1
        exact = source[start:end]
        end_line = line_number + exact.rstrip("\r\n").count("\n")
        procedures.append({
            "name": match.group("name"),
            "kind": match.group("kind").title(),
            "signature": match.group("signature").strip(),
            "startLine": line_number,
            "endLine": end_line,
            "exactSource": exact,
        })
    names = [item["name"].lower() for item in procedures]
    if len(names) != len(set(names)):
        raise MacroBehaviorError("duplicate procedure name")
    return procedures


def _range_shape(address: str) -> tuple[int, int] | None:
    match = re.fullmatch(r"\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", address, re.I)
    if not match:
        return None
    def number(column: str) -> int:
        total = 0
        for character in column.upper():
            total = total * 26 + ord(character) - 64
        return total
    return abs(int(match.group(4)) - int(match.group(2))) + 1, \
        abs(number(match.group(3)) - number(match.group(1))) + 1


def _first_cell(sheet: str, address: str) -> str:
    return f"{sheet}!{address.split(':', 1)[0].replace('$', '')}"


def _analyze_body(executable: str) -> dict:
    calls = [{"name": match.group(1)} for match in _code_matches(_CALL, executable)]
    # A stand-alone identifier is also a VBA Sub invocation. Member calls,
    # declarations, labels and assignments are deliberately excluded here.
    keywords = {"end", "else", "next", "exit", "option", "return", "wend"}
    for line in executable.splitlines():
        bare = re.fullmatch(r"[ \t]*([A-Za-z_][A-Za-z0-9_]*)[ \t]*", line)
        if bare and bare[1].lower() not in keywords:
            calls.append({"name": bare[1]})
    reads = []
    affected = []
    for match in _code_matches(_STATIC_RANGE, executable):
        sheet, address = match.groups()
        record = {"kind": "range", "sheet": sheet, "address": address,
                  "identity": f"{sheet}!{address}"}
        if record not in reads:
            reads.append(record)
        shape = _range_shape(address)
        if shape:
            affected.append({**record, "rows": shape[0], "columns": shape[1]})
    writes = []
    for match in _code_matches(_CELL_WRITE, executable):
        row_text = match.group("row").strip()
        column_text = match.group("column").strip()
        writes.append({
            "kind": "cell-update",
            "target": match.group("target"),
            "row": int(row_text) if row_text.isdigit() else row_text,
            "column": int(column_text) if column_text.isdigit() else column_text,
            "value": match.group("value").strip(),
            "usesNow": bool(re.fullmatch(r"Now(?:\(\))?", match.group("value").strip(), re.I)),
        })
    for match in _code_matches(_RANGE_WRITE, executable):
        writes.append({
            "kind": "range-update", "target": match.group("target"),
            "sheet": match.group("sheet"), "address": match.group("address"),
            "value": match.group("value").strip(),
            "usesNow": bool(re.fullmatch(r"Now(?:\(\))?", match.group("value").strip(), re.I)),
        })
    masked = _mask_strings(executable)
    append_evidence = []
    if re.search(r"\bListRows\.Add\b", masked, re.I):
        append_evidence.append("ListRows.Add")
        writes.append({"kind": "append", "evidence": "ListRows.Add"})
    paste = re.search(r'([A-Za-z_]\w*)\.Cells\(\s*([A-Za-z_]\w*|\d+)\s*,\s*1\)\.PasteSpecial[ \t]+Paste:=xlPasteValues', masked, re.I)
    if paste:
        destination,row=paste.groups()
        next_row = re.search(r'\b'+re.escape(row)+r'\s*=\s*'+re.escape(destination)+r'\.Cells\([^\n]+\)\.End\(xlUp\)\.Row\s*\+\s*1\b',masked,re.I)
        if next_row:
            append_evidence.append('PasteSpecial xlPasteValues to proven next row')
            writes.append({'kind':'append','evidence':'PasteSpecial xlPasteValues',
                           'destination':destination,'rowVariable':row})
        else:
            writes.append({'kind':'paste-values','destination':destination,'row':row})

    guards = []
    for match in re.finditer(r"^[ \t]*If[ \t]+(.+?)[ \t]+Then[ \t]+Exit[ \t]+Sub", executable, re.I | re.M):
        condition = match.group(1).strip()
        if re.search(r"targetRow[ \t]*<[ \t]*2", condition, re.I):
            kind = "minimum-row"
        elif re.search(r"ActiveSheet\.Name[ \t]*<>", condition, re.I):
            kind = "active-sheet"
        else:
            kind = "early-exit"
        guards.append({"kind": kind, "condition": condition})
    selected = None
    selected_match = re.search(r"([A-Za-z_][A-Za-z0-9_]*)[ \t]*=[ \t]*ActiveCell\.Row", executable, re.I)
    if selected_match:
        selected = {"kind": "active-row", "variable": selected_match.group(1),
                    "source": "ActiveCell.Row", "cardinality": "implicit-single-cell"}

    external = []
    for kind, pattern in (
        ("email", r"CreateObject\([ \t]*\"Outlook\.Application\""),
        ("file-export", r"ExportAsFixedFormat"),
        ("shell", r"\bShell\b"),
        ("file-io", r"\bKill\b|\bMkDir\b|FileSystemObject|\bOpen .+ For (?:Input|Output|Append)"),
        ("dynamic-invocation", r"Application\.Run|CallByName|ExecuteExcel4Macro"),
        ("external-refresh", r"\bRefreshAll\b"),
        ("http", r"MSXML2|WinHttp|ServerXMLHTTP"),
    ):
        if list(_code_matches(pattern, executable, re.I)):
            external.append({"kind": kind})
    unresolved = []
    if _DYNAMIC_RANGE.search(executable) or re.search(r"\.Range\([ \t]*[A-Za-z_]", executable, re.I):
        unresolved.append({"kind": "dynamic-sheet-or-range",
                           "reason": "sheet or range identity is computed at runtime"})

    overwrites = []
    if affected and re.search(r"PasteSpecial[\s\S]*?Cells\([^,]+,[ \t]*1\)\.Value[ \t]*=[ \t]*Now", executable, re.I):
        overwrites.append({"kind": "first-destination-cell", "sourceCell": _first_cell(
            affected[0]["sheet"], affected[0]["address"]), "replacement": "Now"})

    if append_evidence:
        operation, status = "append", "ACTION"
    elif writes:
        operation, status = "update", "ACTION"
    elif re.search(r"(?:CalculateFull|Application\.Calculate)\b", executable, re.I):
        operation, status = "recalculate", "AUTO"
    elif re.search(r"\.Activate\b|\.Select\b", executable, re.I):
        operation, status = "navigate", "AUTO"
    elif external:
        operation, status = "external", "UNRESOLVED"
    else:
        operation, status = "unknown", "UNRESOLVED"
    return {
        "calls": calls, "reads": reads, "writes": writes, "guards": guards,
        "selectedRecord": selected, "affectedRanges": affected,
        "overwrites": overwrites, "externalSideEffects": external,
        "unresolvedDependencies": unresolved,
        "disposition": {"operation": operation, "status": status},
    }


def analyze_macro_source(path: str | Path, *, source_kind: str = "raw-module") -> dict:
    source_path = Path(path).expanduser().resolve()
    if not source_path.is_file():
        raise MacroBehaviorError(f"macro source not found: {source_path}")
    module_bytes, source, extra = _extract_source(source_path)
    module_hash = _sha256(module_bytes)
    name_match = _MODULE_NAME.search(source)
    module_name = name_match.group(1) if name_match else source_path.stem
    slices = _procedure_slices(source)
    records = []
    for item in slices:
        exact = item.pop("exactSource")
        executable = _strip_comments(exact)
        facts = _analyze_body(executable)
        procedure_id = f"{module_name}.{item['name']}"
        body_hash = _sha256(exact.encode("utf-8"))
        procedure_hash = _sha256(
            module_hash.encode("ascii") + b"\0" + procedure_id.encode("utf-8")
            + b"\0" + exact.encode("utf-8")
        )
        records.append({
            "procedureId": procedure_id, **item, "bodySha256": body_hash,
            "procedureSha256": procedure_hash, "sourceBody": exact, "helperDefinitions": [], **facts,
        })
    by_name = {record["name"].lower(): record for record in records}
    resolved = set()
    def expand(record, stack=()):
        key = record['name'].lower()
        if key in resolved:
            return
        for call in record['calls']:
            helper = by_name.get(call['name'].lower())
            if helper is None or helper['name'].lower() in stack + (key,):
                record['unresolvedDependencies'].append({
                    'kind': 'missing-helper' if helper is None else 'recursive-helper',
                    'name': call['name'], 'reason': 'helper absent or recursive in exact module'})
                continue
            expand(helper, stack + (key,))
            definitions = [{'procedureId': helper['procedureId'],
                            'procedureSha256': helper['procedureSha256'],
                            'bodySha256': helper['bodySha256']}] + helper['helperDefinitions']
            for value in definitions:
                if value not in record['helperDefinitions']:
                    record['helperDefinitions'].append(copy.deepcopy(value))
            for field in ('reads', 'writes', 'externalSideEffects', 'affectedRanges',
                          'overwrites', 'unresolvedDependencies', 'guards'):
                for value in helper[field]:
                    if value not in record[field]:
                        record[field].append(copy.deepcopy(value))
        if record['unresolvedDependencies'] or record['externalSideEffects']:
            record['disposition']['status'] = 'UNRESOLVED'
        elif any(write['kind'] == 'append' for write in record['writes']):
            record['disposition'] = {'operation': 'append', 'status': 'ACTION'}
        elif record['writes']:
            record['disposition'] = {'operation': 'update', 'status': 'ACTION'}
        resolved.add(key)
    for record in records:
        expand(record)
    return {
        "analysisVersion": 1, "sourceKind": source_kind,
        "sourcePath": str(source_path), "moduleName": module_name,
        "moduleByteSha256": module_hash, **extra, "procedures": records,
    }


def validate_macro_source(analysis: dict) -> None:
    if not isinstance(analysis, dict) or analysis.get("analysisVersion") != 1:
        raise MacroBehaviorError("unsupported macro analysis")
    current = analyze_macro_source(
        analysis.get("sourcePath", ""), source_kind=analysis.get("sourceKind", "raw-module"))
    if current["moduleByteSha256"] != analysis.get("moduleByteSha256"):
        raise MacroBehaviorError("macro module bytes changed")
    expected = {record.get("procedureId"): record for record in analysis.get("procedures", [])}
    observed = {record["procedureId"]: record for record in current["procedures"]}
    if None in expected or len(expected) != len(analysis.get("procedures", [])):
        raise MacroBehaviorError("duplicate or invalid analyzed procedure")
    if set(expected) != set(observed):
        raise MacroBehaviorError("macro procedure set changed")
    for procedure_id in expected:
        if expected[procedure_id] != observed[procedure_id]:
            raise MacroBehaviorError(f"macro procedure/helper/behavior changed: {procedure_id}")


def procedure_by_identity(analysis: dict, procedure_id: str) -> dict:
    matches = [record for record in analysis.get("procedures", [])
               if record.get("procedureId") == procedure_id]
    if len(matches) != 1:
        raise MacroBehaviorError(f"procedure identity did not resolve exactly once: {procedure_id}")
    return matches[0]
