"""Smoke test for the deployed PoliMusic site.

Logs in as the administrator and requests every page, reporting HTTP status,
response size and whether IIS injected an error message.

    python deploy/smoke_test.py
    python deploy/smoke_test.py http://localhost/webapppolimusicv1_asp
"""
import http.cookiejar
import re
import sys
import urllib.parse
import urllib.request

BASE = (sys.argv[1] if len(sys.argv) > 1 else "http://localhost/webapppolimusicv1_asp").rstrip("/")

ERROR_MARKERS = [
    "An error occurred on the server when processing the URL",
    "Microsoft VBScript runtime error",
    "Microsoft VBScript compilation error",
    "Error del compilador de Microsoft VBScript",
    "error 'ASP",
    "Error de tiempo de ejecución de Microsoft VBScript",
]

jar = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))


def fetch(path, data=None):
    url = BASE + "/" + path.lstrip("/")
    body = urllib.parse.urlencode(data).encode() if data else None
    req = urllib.request.Request(url, data=body, method="POST" if data else "GET")
    if data:
        req.add_header("Content-Type", "application/x-www-form-urlencoded")
    try:
        with opener.open(req, timeout=25) as r:
            return r.status, r.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "replace")
    except Exception as e:  # noqa: BLE001
        return 0, f"{type(e).__name__}: {e}"


def check(label, path, expect=200):
    status, body = fetch(path)
    problems = [m for m in ERROR_MARKERS if m in body]
    ok = status == expect and not problems
    flag = "OK  " if ok else "FAIL"
    print(f"  [{flag}] {status} {len(body):>6}b  {label:<28} {path}")
    if problems:
        idx = body.find(problems[0])
        print("         ...", body[max(0, idx - 220):idx + 90].replace("\n", " ")[-300:])
    return ok


print(f"PoliMusic smoke test -> {BASE}\n")

status, page = fetch("login.asp")
token = re.search(r'name="csrf-token" content="([^"]+)"', page)
if not token:
    print("  [FAIL] login.asp did not render (no CSRF meta tag)")
    print(page[:600])
    raise SystemExit(1)
csrf = token.group(1)
print(f"  [OK  ] login.asp rendered, csrf={csrf[:12]}...\n")

status, body = fetch("login.asp", {
    "_csrf": csrf,
    "next": "",
    "username": "userAdmin",
    "password": "12345678",
})
errors = [m for m in ERROR_MARKERS if m in body]
logged = status == 200 and not errors and ("Log out" in body or "Welcome" in body)
print(f"  [{'OK  ' if logged else 'FAIL'}] POST login.asp -> {status} "
      f"({'session started' if logged else 'login rejected or page error'})\n")

# the token is regenerated on sign-in: refresh it from a protected page
status, page = fetch("default.asp")
m = re.search(r'name="csrf-token" content="([^"]+)"', page)
if m:
    csrf = m.group(1)

pages = [
    ("Home / welcome", "default.asp"),
    ("Songs list", "songs.asp"),
    ("Songs + filter", "songs.asp?q=ad"),
    ("Artists", "artists.asp"),
    ("Albums", "albums.asp"),
    ("Album detail", "album_detail.asp?id=1"),
    ("Playlists", "playlists.asp"),
    ("Playlist detail", "playlist_detail.asp?id=1"),
    ("Users (admin)", "users.asp"),
    ("User form", "user_form.asp?id=1"),
    ("Song form", "song_form.asp?id=1"),
    ("Artist form", "artist_form.asp?id=1"),
    ("Album form", "album_form.asp?id=1"),
    ("Playlist form", "playlist_form.asp"),
    ("New song form", "song_form.asp"),
]

results = [check(label, path) for label, path in pages]

status, body = fetch("api_play.asp", {"_csrf": csrf, "id": "1"})
json_ok = '"ok":true' in body.replace(" ", "")
print(f"  [{'OK  ' if json_ok else 'FAIL'}] {status} api_play.asp -> {body.strip()[:80]}")

check("Logout", "logout.asp")
status, body = fetch("default.asp")
print(f"  [{'OK  ' if 'Welcome back' in body or 'Sign in' in body else 'FAIL'}] "
      f"{status} default.asp after logout -> protected")

print(f"\n{sum(results)}/{len(results)} pages OK")
raise SystemExit(0 if all(results) else 1)
