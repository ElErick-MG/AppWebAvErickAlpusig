"""Static sanity checks for the PoliMusic ASP pages.

Checks:
  1. <% %> delimiters balanced
  2. block structure balanced (If/Do/For/Select/Function/Sub/With)
  3. identifiers that look like variables but are never Dim'd (Option Explicit)
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

KEYWORDS = {
    "if", "then", "else", "elseif", "end", "for", "each", "next", "do", "loop",
    "while", "wend", "select", "case", "function", "sub", "exit", "dim", "redim",
    "const", "set", "let", "call", "and", "or", "not", "xor", "mod", "is", "new",
    "nothing", "empty", "null", "true", "false", "on", "error", "resume", "goto",
    "with", "byval", "byref", "option", "explicit", "preserve", "to", "step",
    "private", "public", "class", "property", "get", "errmaybe", "in",
}

BUILTINS = {
    "response", "request", "session", "application", "server", "err",
    "len", "left", "right", "mid", "trim", "ltrim", "rtrim", "ucase", "lcase",
    "replace", "split", "join", "instr", "instrrev", "strcomp", "cstr", "clng",
    "cint", "cdbl", "csng", "cbool", "cdate", "isnull", "isempty", "isarray",
    "isnumeric", "isdate", "isobject", "isnull", "ubound", "lbound", "array",
    "now", "date", "time", "year", "month", "day", "hour", "minute", "second",
    "dateserial", "dateadd", "datediff", "formatnumber", "formatdatetime",
    "monthname", "weekdayname", "rnd", "randomize", "int", "fix", "abs", "sgn",
    "round", "sqr", "exp", "log", "sin", "cos", "tan", "atn", "asc", "chr",
    "string", "space", "strreverse", "vbcrlf", "vblf", "vbtab", "vbcr",
    "vbcritical", "vbquestion", "vbnewline", "typename", "vartype", "eval",
    "execute", "getobject", "createobject", "scriptengine", "rgb", "timer",
    "filter", "datevalue", "timevalue", "cbyte", "ccur", "escape", "unescape",
    "msgbox", "inputbox", "hex", "oct", "lenb", "midb", "leftb", "rightb",
    "ascb", "chrb", "strcomp", "partition", "switch", "me", "varptr", "strptr",
    "objptr", "loadpicture", "calendar", "weekday", "datepart", "dateserial",
    "regexp", "_",
}


def strip_comment(line):
    """Remove a VBScript comment, honouring double-quoted strings."""
    out = []
    i = 0
    in_str = False
    while i < len(line):
        ch = line[i]
        if ch == '"':
            in_str = not in_str
            out.append(ch)
        elif ch == "'" and not in_str:
            break
        else:
            out.append(ch)
        i += 1
    return "".join(out)


def script_blocks(text):
    """Yield (start_line, code) for every <% ... %> block (not <%=)."""
    blocks = []
    for m in re.finditer(r"<%(?!@)(=?)(.*?)%>", text, re.S):
        if m.group(1) == "=":
            continue
        code = m.group(2)
        start = text[: m.start()].count("\n") + 1
        blocks.append((start, code))
    return blocks


def analyse(path):
    text = open(path, encoding="utf-8", errors="replace").read()
    problems = []

    # 1. delimiters
    opens = len(re.findall(r"<%", text))
    closes = len(re.findall(r"%>", text))
    if opens != closes:
        problems.append(f"delimiter mismatch: {opens} '<%' vs {closes} '%>'")

    # 2. block structure
    depth = 0
    declared = set()
    params = set()

    for start, code in script_blocks(text):
        for offset, raw in enumerate(code.split("\n")):
            line = strip_comment(raw).strip()
            if not line:
                continue
            low = line.lower()

            m = re.match(r"^(?:public |private )?(?:function|sub)\s+([a-z_]\w*)\s*\(([^)]*)\)", low)
            if m:
                depth += 1
                declared.add(m.group(1))
                for p in m.group(2).split(","):
                    p = p.strip()
                    p = re.sub(r"^(byval|byref)\s+", "", p)
                    if p:
                        params.add(p)
                        declared.add(p)
                continue
            m = re.match(r"^(?:public |private )?(?:function|sub)\s+([a-z_]\w*)", low)
            if m:
                depth += 1
                declared.add(m.group(1))
                continue

            if low.startswith("end function") or low.startswith("end sub"):
                depth -= 1
                if depth < 0:
                    problems.append(f"line {start + offset}: unbalanced End Function/Sub")
                    depth = 0
                continue

            m = re.match(r"^dim\s+(.+)$", low)
            if m:
                for name in m.group(1).split(","):
                    name = name.strip().split("(")[0].strip()
                    if re.match(r"^[a-z_]\w*$", name):
                        declared.add(name)
                continue
            m = re.match(r"^(?:public |private )?const\s+([a-z_]\w*)", low)
            if m:
                declared.add(m.group(1))
                continue
            m = re.match(r"^const\s+(.+)$", low)
            if m:
                for name in m.group(1).split(","):
                    name = name.split("=")[0].strip()
                    if re.match(r"^[a-z_]\w*$", name):
                        declared.add(name)
                continue

            if re.match(r"^if\b", low):
                # multi-line only when "then" ends the statement
                body = re.sub(r"\s+", " ", low)
                if body.endswith("then"):
                    depth += 1
                continue
            if low.startswith("elseif"):
                body = re.sub(r"\s+", " ", low)
                if body.endswith("then"):
                    pass
                continue
            if low.startswith("end if"):
                depth -= 1
                if depth < 0:
                    problems.append(f"line {start + offset}: unbalanced End If")
                    depth = 0
                continue

            if re.match(r"^do\b", low):
                depth += 1
                continue
            if low.startswith("loop"):
                depth -= 1
                if depth < 0:
                    problems.append(f"line {start + offset}: unbalanced Loop")
                    depth = 0
                continue

            if re.match(r"^for\b", low):
                depth += 1
                m2 = re.match(r"^for\s+(?:each\s+)?([a-z_]\w*)", low)
                if m2:
                    declared.add(m2.group(1))
                continue
            if low.startswith("next"):
                depth -= 1
                if depth < 0:
                    problems.append(f"line {start + offset}: unbalanced Next")
                    depth = 0
                continue

            if low.startswith("select case"):
                depth += 1
                continue
            if low.startswith("end select"):
                depth -= 1
                if depth < 0:
                    problems.append(f"line {start + offset}: unbalanced End Select")
                    depth = 0
                continue

            if re.match(r"^with\b", low):
                depth += 1
                continue
            if low.startswith("end with"):
                depth -= 1
                if depth < 0:
                    problems.append(f"line {start + offset}: unbalanced End With")
                    depth = 0
                continue

            if low.startswith("case ") or low == "case else":
                continue

    if depth != 0:
        problems.append(f"block depth ends at {depth} (expected 0)")

    for start, code in script_blocks(text):
        for offset, raw in enumerate(code.split("\n")):
            if raw.count('"') % 2 == 1:
                problems.append(f"line {start + offset}: odd number of double quotes -> {raw.strip()[:70]}")

    return problems, declared, text


def check_identifiers(path, declared, text):
    """Heuristic Option Explicit check across script blocks."""
    undeclared = {}
    for start, code in script_blocks(text):
        for offset, raw in enumerate(code.split("\n")):
            line = strip_comment(raw)
            # drop string literals
            line = re.sub(r'"[^"]*"', '""', line)
            for m in re.finditer(r"(?<![.\w])([A-Za-z_]\w*)", line):
                name = m.group(1)
                low = name.lower()
                if low in KEYWORDS or low in BUILTINS:
                    continue
                if low in declared:
                    continue
                after = line[m.end():]
                if after.lstrip().startswith("("):
                    # function/sub call
                    continue
                if re.match(r"^\s*=", after) or re.match(r"^\s*=", after):
                    pass
                undeclared.setdefault(low, start + offset)
    return undeclared


def main():
    files = []
    for dirpath, _dirs, names in os.walk(ROOT):
        if "SongFilesAndImages_resources" in dirpath or ".workbuddy-ai" in dirpath:
            continue
        for n in names:
            if n.lower().endswith(".asp"):
                files.append(os.path.join(dirpath, n))
    files.sort()

    global_declared = set()
    for f in files:
        _p, decl, _t = analyse(f)
        global_declared |= decl

    total = 0
    for f in files:
        problems, declared, text = analyse(f)
        undeclared = check_identifiers(f, declared | global_declared, text)
        rel = os.path.relpath(f, ROOT)
        if problems or undeclared:
            total += 1
            print(f"\n=== {rel} ===")
            for p in problems:
                print("  STRUCT:", p)
            if undeclared:
                items = ", ".join(f"{k}@{v}" for k, v in sorted(undeclared.items()))
                print("  MAYBE UNDECLARED:", items)
    print(f"\n{len(files)} files checked, {total} with findings")


if __name__ == "__main__":
    main()
