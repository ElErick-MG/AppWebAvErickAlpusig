"""Checks every Function/Sub call against its declaration.

VBScript raises "Wrong number of arguments" at run time, which is exactly the
class of bug that broke login.asp, so it is worth catching statically.

Run:  python deploy/check_calls.py
"""
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKIP_DIRS = {"SongFilesAndImages_resources", ".workbuddy-ai", "docs", "assets"}

# Calls that belong to the ASP/ADO object model, not to our code.
IGNORE = {
    "if", "for", "while", "select", "with", "do", "elseif", "else", "case",
    "response", "request", "session", "application", "server", "err", "conn",
    "rs", "cmd", "cn", "fso", "re", "obj",
}


def asp_files():
    out = []
    for dirpath, dirs, names in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
        for n in names:
            if n.lower().endswith(".asp"):
                out.append(os.path.join(dirpath, n))
    return sorted(out)


def mask_strings(line):
    """Replace string contents with x, keeping quotes and length so that
    character offsets stay valid for argument counting."""
    return re.sub(r'"([^"]*)"', lambda m: '"' + "x" * len(m.group(1)) + '"', line)


def strip_comment(line):
    out, in_str = [], False
    for ch in line:
        if ch == '"':
            in_str = not in_str
        if ch == "'" and not in_str:
            break
        out.append(ch)
    return "".join(out)


def count_args(text, start):
    """Given text and index of '(', return (argc, end_index)."""
    depth = 0
    argc = 1
    i = start
    empty = True
    while i < len(text):
        ch = text[i]
        if ch == '"':
            if depth == 1:
                empty = False
            i += 1
            while i < len(text) and text[i] != '"':
                i += 1
        elif ch == "(":
            depth += 1
            if depth > 1:
                empty = False
        elif ch == ")":
            depth -= 1
            if depth == 0:
                if empty:
                    return 0, i
                return argc, i
        elif ch == "," and depth == 1:
            argc += 1
        elif not ch.isspace() and depth == 1:
            empty = False
        i += 1
    return None, -1


def main():
    files = asp_files()
    sources = {f: open(f, encoding="utf-8", errors="replace").read() for f in files}

    # 1. collect declarations
    decl = {}
    for f, text in sources.items():
        for m in re.finditer(r"^\s*(?:Public |Private )?(Function|Sub)\s+([A-Za-z_]\w*)\s*\(([^)]*)\)",
                             text, re.I | re.M):
            kind, name, params = m.group(1), m.group(2).lower(), m.group(3).strip()
            if params == "":
                total = 0
                required = 0
            else:
                parts = [p.strip() for p in params.split(",")]
                total = len(parts)
                required = sum(1 for p in parts if not p.lower().startswith("optional"))
            decl[name] = (kind, required, total, f)

    print(f"{len(decl)} procedures declared\n")

    problems = 0
    for f, text in sources.items():
        for lineno, raw in enumerate(text.split("\n"), 1):
            line = strip_comment(raw)
            if not line.strip():
                continue
            code = mask_strings(line)
            for m in re.finditer(r"(?<![.\w])([A-Za-z_]\w*)\s*\(", code):
                name = m.group(1).lower()
                if name not in decl or name in IGNORE:
                    continue
                # skip the declaration itself
                head = code[: m.start()].strip().lower()
                if head.endswith("function") or head.endswith("sub"):
                    continue
                argc, _end = count_args(code, m.end() - 1)
                if argc is None:
                    continue
                kind, req, total, where = decl[name]
                if argc < req or argc > total:
                    problems += 1
                    print(f"{os.path.relpath(f, ROOT)}:{lineno}: {name}({argc} args) "
                          f"-> declared {req}..{total} in {os.path.relpath(where, ROOT)}")
                    print(f"    {raw.strip()[:110]}")

    # 2. a Function that assigns its result to a DIFFERENT function name
    #    (VBScript cannot assign to another procedure's name: runtime error)
    bad_assign = 0
    for f, text in sources.items():
        for m in re.finditer(r"^\s*(?:Public |Private )?Function\s+([A-Za-z_]\w*)\s*\(([^)]*)\)",
                             text, re.I | re.M):
            name = m.group(1).lower()
            end = re.search(r"^\s*End Function", text[m.end():], re.I | re.M)
            if not end:
                continue
            body = text[m.end(): m.end() + end.start()]
            local = set()
            for dl in re.finditer(r"^\s*Dim\s+(.+)$", body, re.I | re.M):
                for nm in dl.group(1).split(","):
                    nm = nm.strip().split("(")[0].strip().lower()
                    if nm:
                        local.add(nm)
            for line in body.split("\n"):
                l = line.strip()
                if not l or l.startswith("'"):
                    continue
                l = re.sub(r"^set\s+", "", l, flags=re.I)
                mm = re.match(r"^([A-Za-z_]\w*)\s*=", l)
                if mm:
                    target = mm.group(1).lower()
                    if target in decl and target != name and target not in local:
                        bad_assign += 1
                        print(f"{os.path.relpath(f, ROOT)}: Function {name}() assigns to '{target}'")
                        print(f"    {l[:100]}")
    print(f"\n{bad_assign} wrong return-variable problem(s) found")
    print(f"{problems} arity problem(s) found")
    return 0 if problems == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
