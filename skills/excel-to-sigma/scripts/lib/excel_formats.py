"""Excel ``number_format`` codes -> Sigma column ``format`` objects (pure functions).

Sigma column formats are ``{"kind": "number", "formatString": <d3-format>, prefix?, suffix?}`` and
``{"kind": "datetime", "formatString": <d3 time / strftime>}`` (OpenAPI ``Table.columns[].format``; the
number ``formatString`` is d3-format syntax, never an Excel mask - ``$#,##0`` is rejected by Sigma).

``translate(code)`` never raises and never drops a code silently: a code it cannot map exactly comes back
with ``approximated`` reasons so the drafter can list it as a ``NUMBER_FORMAT_APPROXIMATED`` review issue.

Negatives
---------
* ``(`` d3 sign (``(,.0f``) shows a negative as ``(1,234)`` - the d3-format sign option, so it needs no
  Sigma-specific field. ``translate(code, native_parens=False)`` falls back to a plain minus sign and says so.
* ``[Red]`` has no number-format equivalent; ``Translation.negative_red`` is set and the builder emits a
  ``conditionalFormats`` rule (``type: single``, ``condition: "<"``, ``value: 0``, ``style.color``) on the
  columns that carry it (OpenAPI ``Table`` / ``pivot-table`` ``conditionalFormats``).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

THOUSANDS = ","


@dataclass
class Translation:
    """Result of translating one Excel format code."""
    code: str
    category: str                        # general | text | number | percent | currency | date | unknown
    format: dict | None = None           # Sigma format object; None for General / text
    negative_red: bool = False
    approximated: list = field(default_factory=list)

    @property
    def is_percent(self):
        return self.category == "percent"

    @property
    def is_date(self):
        return self.category == "date"


# ---------------------------------------------------------------- number format roles (d3-format)
def plain_number(decimals=0, thousands=True):
    """A plain numeric format (thousands separator, fixed decimals): what General / unknown codes become so
    Sigma never inherits a ``%`` (or any other) format from an operand."""
    return {"kind": "number", "formatString": f"{THOUSANDS if thousands else ''}.{int(decimals)}f"}


def number_decimals(values, cap=4):
    """Decimals a plain format needs to show every value as typed (float noise beyond 6 dp ignored)."""
    most = 0
    for value in values:
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            continue
        if value != value or value in (float("inf"), float("-inf")):
            continue
        text = f"{round(float(value), 6):.6f}".rstrip("0")
        places = len(text.split(".")[1]) if "." in text else 0
        most = max(most, places)
    return min(most, cap)


DEFAULT_DATE_FORMAT = {"kind": "datetime", "formatString": "%Y-%m-%d"}


def is_percent_format(fmt):
    """True for a plan/Sigma format that scales by 100 (legacy ``"percent"`` or a ``%`` d3 string)."""
    if fmt == "percent":
        return True
    return isinstance(fmt, dict) and fmt.get("kind") == "number" \
        and str(fmt.get("formatString", "")).endswith("%")


# ---------------------------------------------------------------- tokenising a section
_ESCAPED = re.compile(r'"([^"]*)"|\\(.)|_(.)|\*(.)|\[([^\]]*)\]')


def _split_sections(code):
    sections, current, quoted, escaped = [], [], False, False
    for char in code:
        if escaped:
            current.append(char)
            escaped = False
        elif char == "\\":
            current.append(char)
            escaped = True
        elif char == '"':
            current.append(char)
            quoted = not quoted
        elif char == ";" and not quoted:
            sections.append("".join(current))
            current = []
        else:
            current.append(char)
    sections.append("".join(current))
    return sections


def _clean(section):
    """(text with quoted/escaped/bracketed items blanked, literals, color, currency symbol, conditions)."""
    literals, color, symbol, conditions = [], None, None, []
    out = []
    index = 0
    while index < len(section):
        match = _ESCAPED.match(section, index)
        if match:
            quoted, esc, pad, fill, bracket = match.groups()
            if quoted is not None:
                out.append(chr(0xE000 + len(literals)))
                literals.append(quoted)
            elif esc is not None:
                out.append(chr(0xE000 + len(literals)))
                literals.append(esc)
            elif pad is not None or fill is not None:
                out.append(" " if pad is not None else "")   # _) pads a space, * fills: both display-only
            else:
                low = bracket.lower()
                if low in ("red", "black", "blue", "green", "white", "yellow", "magenta", "cyan") \
                        or low.startswith("color"):
                    color = low
                elif bracket.startswith("$"):
                    symbol = bracket[1:].split("-")[0]       # [$€-407] -> € ; [$-409] -> locale tag only
                elif re.match(r"^(<=|>=|<>|<|>|=)", bracket):
                    conditions.append(bracket)
                elif re.match(r"^[hms]+$", low):
                    conditions.append(bracket)               # elapsed time [h]
            index = match.end()
        else:
            out.append(section[index])
            index += 1
    return "".join(out), literals, color, symbol, conditions


def _literal_text(text, literals):
    """Replace the literal placeholders by their text."""
    return re.sub("[\ue000-\uf8ff]", lambda m: literals[ord(m.group(0)) - 0xE000], text)


# ---------------------------------------------------------------- dates
_DATE_TOKEN = re.compile(r"(yyyy|yy|y|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM/PM|A/P|am/pm|a/p)", re.I)


def _is_date_code(text):
    """True when the (cleaned) first section is a date/time mask."""
    if re.search(r"[0#?]", text):
        return False
    return bool(re.search(r"[ymdhs]|AM/PM|A/P", text, re.I))


def _translate_date(code, text, literals, conditions):
    notes = []
    if conditions:
        return Translation(code, "unknown", None, False,
                           ["elapsed-time / conditional date masks are not supported"])
    tokens = []
    for part in re.split(_DATE_TOKEN, text):
        if part == "":
            continue
        tokens.append((part.lower(), part) if _DATE_TOKEN.fullmatch(part) else (None, part))
    has_ampm = any(t in ("am/pm", "a/p") for t, _ in tokens)
    out = []
    for index, (kind, raw) in enumerate(tokens):
        if kind is None:
            out.append(_date_literal(raw, literals))
            continue
        if kind in ("yyyy",):
            out.append("%Y")
        elif kind in ("yy", "y"):
            out.append("%y")
        elif kind in ("mmmm",):
            out.append("%B")
        elif kind == "mmmmm":
            out.append("%b")
            notes.append("mmmmm (first letter of the month) is shown as the short month name")
        elif kind == "mmm":
            out.append("%b")
        elif kind in ("mm", "m"):
            is_minute = _is_minute(tokens, index)
            if is_minute:
                out.append("%M" if kind == "mm" else "%-M")
            else:
                out.append("%m" if kind == "mm" else "%-m")
        elif kind == "dddd":
            out.append("%A")
        elif kind == "ddd":
            out.append("%a")
        elif kind == "dd":
            out.append("%d")
        elif kind == "d":
            out.append("%-d")
        elif kind in ("hh", "h"):
            if has_ampm:
                out.append("%I" if kind == "hh" else "%-I")
            else:
                out.append("%H" if kind == "hh" else "%-H")
        elif kind == "ss":
            out.append("%S")
        elif kind == "s":
            out.append("%-S")
        else:
            out.append("%p")
    mask = "".join(out)
    if re.search(r"\.0+", _literal_text("".join(r for k, r in tokens if k is None), literals)) and "%S" in mask:
        notes.append("fractional seconds are not shown")
    return Translation(code, "date", {"kind": "datetime", "formatString": mask}, False, notes)


def _date_literal(raw, literals):
    text = _literal_text(raw, literals)
    return text.replace("%", "%%")


def _is_minute(tokens, index):
    """Excel's ``m`` is minutes right after an hour token or right before a seconds token."""
    def neighbour(step):
        position = index + step
        while 0 <= position < len(tokens):
            kind, raw = tokens[position]
            if kind is not None:
                return kind
            if re.search(r"[^\s:.]", re.sub("[\ue000-\uf8ff]", "", raw)):
                return None
            position += step
        return None
    return neighbour(-1) in ("h", "hh") or neighbour(1) in ("s", "ss")


# ---------------------------------------------------------------- numbers
_DIGITS = re.compile(r"[0#?]")


def _parse_number_section(text, literals, symbol):
    """Structure of one numeric section or ``None`` when it has no digit placeholder."""
    if not _DIGITS.search(text):
        return None
    sci = re.search(r"[eE][+-]", text)
    exponent = ""
    if sci:
        text, exponent = text[:sci.start()], text[sci.start():]
    first = _DIGITS.search(text).start()
    last = max(m.end() for m in _DIGITS.finditer(text))
    prefix_raw, body, suffix_raw = text[:first], text[first:last], text[last:]
    percent = "%" in body or "%" in prefix_raw or "%" in suffix_raw
    scale = len(re.match(r",*", suffix_raw).group(0))          # trailing commas divide by 1000 each
    suffix_raw = suffix_raw[scale:]
    integer, _, decimals = body.partition(".")
    thousands = "," in integer
    int_digits = re.sub(r"[^0#?]", "", integer)
    dec_digits = re.sub(r"[^0#?]", "", decimals)
    fixed = dec_digits.count("0")
    optional = len(dec_digits) - fixed
    min_int = int_digits.count("0")
    return {
        "prefix": _literal_text(prefix_raw.replace("%", ""), literals),
        "suffix": _literal_text(suffix_raw.replace("%", ""), literals),
        "percent": percent, "scale": scale, "scientific": bool(sci), "thousands": thousands,
        "decimals": len(dec_digits), "fixed": fixed, "optional": optional, "min_int": min_int,
        "symbol": symbol,
        "fraction": "/" in body,
    }


CURRENCY_CHARS = "$€£¥"


def _take_currency(parsed):
    """(currency symbol | None, prefix text, suffix text) once the currency literal is lifted out."""
    prefix, suffix, symbol = parsed["prefix"], parsed["suffix"], parsed["symbol"]
    if not symbol:
        for char in CURRENCY_CHARS:
            if char in prefix:
                symbol, prefix = char, prefix.replace(char, "", 1)
                break
            if char in suffix:
                symbol, suffix = char, suffix.replace(char, "", 1)
                break
    return symbol or None, prefix.lstrip(), suffix.rstrip()


def translate(code, *, native_parens=True):
    """Translate one Excel ``number_format`` string. See the module docstring for the contract."""
    if code is None or str(code).strip() == "" or str(code).strip().lower() == "general":
        return Translation(code or "General", "general")
    code = str(code)
    if code == "mm-dd-yy":                     # openpyxl's name for built-in 14 (locale short date, m/d/yyyy)
        return Translation(code, "date", {"kind": "datetime", "formatString": "%-m/%-d/%Y"})
    sections = _split_sections(code)
    pos_text, pos_lits, _, pos_symbol, pos_cond = _clean(sections[0])
    if pos_text.strip() == "@" and len(sections) == 1:
        return Translation(code, "text")
    if _is_date_code(pos_text):
        return _translate_date(code, pos_text, pos_lits, pos_cond)
    notes = []
    if pos_cond:
        notes.append("conditional sections are not supported")
    parsed = _parse_number_section(pos_text, pos_lits, pos_symbol)
    if parsed is None:
        return Translation(code, "unknown", None, False, ["no digit placeholder"])
    if parsed["fraction"]:
        return Translation(code, "unknown", None, False, ["fraction formats are not supported"])
    if parsed["scale"]:
        notes.append("trailing comma (divide by 1000) scaling is not supported")
    if parsed["scientific"]:
        notes.append("scientific exponent is shown as e+N, not E+NN")
    if parsed["min_int"] > 1:
        notes.append("leading zeros (minimum integer digits) are not padded")
    if parsed["fixed"] and parsed["optional"]:
        notes.append("optional decimals (#) next to required ones (0) are all shown")
    negative_red, negative_parens = False, False
    if len(sections) > 1:
        neg_text, neg_lits, neg_color, neg_symbol, _ = _clean(sections[1])
        negative_red = bool(neg_color and neg_color.startswith("red"))
        negative_parens = "(" in neg_text and ")" in neg_text
        if _DIGITS.search(neg_text):
            neg = _parse_number_section(neg_text, neg_lits, neg_symbol)
            if neg and (neg["decimals"] != parsed["decimals"] or neg["percent"] != parsed["percent"]):
                notes.append("the negative section uses different digits than the positive one")
        elif neg_text.strip():
            notes.append("the negative section has no digits")
    if len(sections) > 2:
        zero_text = _clean(sections[2])[0]
        if re.sub(r"[^0#?.,]", "", zero_text) != re.sub(r"[^0#?.,]", "", pos_text):
            notes.append("the zero section is shown with the positive format")
    if len(sections) > 3 and re.sub(r"[\s@]", "", _clean(sections[3])[0]):
        notes.append("the text section is ignored")
    if negative_parens and not native_parens:
        notes.append("negatives in parentheses are shown with a minus sign")
        negative_parens = False
    symbol, prefix, suffix = _take_currency(parsed)
    dollar = symbol == "$"
    precision = parsed["decimals"]
    if parsed["scientific"]:
        body = f".{precision}e"
    elif parsed["percent"]:
        body = f".{precision}{'~' if parsed['optional'] and not parsed['fixed'] and precision else ''}%"
    else:
        body = f".{precision}{'~' if parsed['optional'] and not parsed['fixed'] and precision else ''}f"
    group = THOUSANDS if parsed["thousands"] and not parsed["scientific"] else ""
    fmt = {"kind": "number",
           "formatString": f"{'(' if negative_parens else ''}{'$' if dollar else ''}{group}{body}"}
    if symbol and not dollar:
        prefix = symbol + prefix
    if prefix:
        fmt["prefix"] = prefix
    if suffix:
        fmt["suffix"] = suffix
    category = "percent" if parsed["percent"] else "currency" if symbol else "number"
    return Translation(code, category, fmt, negative_red, notes)


def translate_for(code, *, kind, values=(), native_parens=True):
    """Final Sigma format for a column of ``kind`` (number | datetime) whose source cells use ``code``.

    ``General`` and codes that map to nothing become a plain thousands-separated number with the decimals the
    values need (numbers), or ISO dates (datetimes) - an explicit format always, so Sigma never inherits one.
    Returns ``(format | None, negative_red, approximated_reasons, translation)``.
    """
    result = translate(code, native_parens=native_parens)
    if kind == "datetime":
        if result.is_date and result.format:
            return result.format, False, list(result.approximated), result
        return dict(DEFAULT_DATE_FORMAT), False, ([] if result.category in ("general", "text") else
                                                  [f"{code!r} is not a date mask"]), result
    if kind == "number":
        if result.format and not result.is_date:
            return result.format, result.negative_red, list(result.approximated), result
        reasons = [] if result.category in ("general", "text") else (result.approximated or ["unsupported code"])
        return plain_number(number_decimals(values)), False, reasons, result
    return None, False, [], result
