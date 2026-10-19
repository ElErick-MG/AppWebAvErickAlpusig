"""Functional checks for PoliMusic: roles, rendered content and static assets.

Complements smoke_test.py (which only checks that pages answer 200).
    python deploy/functional_test.py
"""
import http.cookiejar
import re
import sys
import urllib.error
import urllib.parse
import urllib.request

BASE = (sys.argv[1] if len(sys.argv) > 1 else "http://localhost/webapppolimusicv1_asp").rstrip("/")

results = []


def session():
    jar = http.cookiejar.CookieJar()
    return jar, urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))


def get(op, path):
    try:
        with op.open(BASE + "/" + path.lstrip("/"), timeout=25) as r:
            return r.status, r.read().decode("utf-8", "replace"), r.geturl()
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "replace"), ""
    except Exception as e:  # noqa: BLE001
        return 0, str(e), ""


def login(user, password):
    jar, op = session()
    status, page, _ = get(op, "login.asp")
    token = re.search(r'name="csrf-token" content="([^"]+)"', page)
    if not token:
        return None, "login page did not render"
    data = urllib.parse.urlencode({
        "_csrf": token.group(1), "next": "", "username": user, "password": password,
    }).encode()
    req = urllib.request.Request(BASE + "/login.asp", data=data, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    try:
        with op.open(req, timeout=25) as r:
            body = r.read().decode("utf-8", "replace")
            ok = r.status == 200 and "Log out" in body and "Sign in" not in body.split("<title>")[0]
            return op, ("ok" if ok else "rejected")
    except Exception as e:  # noqa: BLE001
        return None, str(e)


def check(label, condition, detail=""):
    results.append((label, bool(condition)))
    print(f"  [{'OK  ' if condition else 'FAIL'}] {label}{('  -> ' + detail) if detail and not condition else ''}")


print(f"PoliMusic functional test -> {BASE}\n")

# ---------------------------------------------------------------- admin
print("Administrator (userAdmin)")
op_admin, state = login("userAdmin", "12345678")
check("login as admin", state == "ok", state)

status, songs, _ = get(op_admin, "songs.asp")
audios = re.findall(r'<audio controls[^>]*src="([^"]+)"', songs)
check("songs.asp lists 5 players", len(audios) == 5, f"found {len(audios)}")
check("player paths resolve to songFiles", all("songFiles/" in a for a in audios), str(audios[:2]))
check("admin sees Actions column", ">Actions<" in songs)
check("admin sees New song button", "song_form.asp" in songs)
check("welcome message present", "Welcome" in songs and "userAdmin" in songs)
check("songs table headers", all(h in songs for h in ("Song Name", "Play", "Reproductions")))
check("delete form posts to delete.asp", 'action="delete.asp"' in songs)
check("no missing-file chip rendered", "audio-missing" not in songs)

status, page, _ = get(op_admin, "users.asp")
check("admin opens users.asp", status == 200 and "userAdmin" in page)
check("users table shows roles", "Administrator" in page and "User" in page)

status, page, _ = get(op_admin, "playlist_detail.asp?id=1")
check("admin opens a playlist", status == 200 and "Playlist 1" in page)

# ---------------------------------------------------------------- normal user
print("\nListener (Erick.Alpusig)")
op_user, state = login("Erick.Alpusig", "12345678")
check("login as listener", state == "ok", state)

status, songs, _ = get(op_user, "songs.asp")
check("listener can open songs.asp", status == 200 and "Song List" in songs)
check("listener does NOT see Actions column", ">Actions<" not in songs)
check("listener does NOT see New song", "song_form.asp" not in songs)
check("listener does NOT see Users nav", "users.asp" not in songs)
check("listener still gets the 5 players", len(re.findall(r"<audio controls", songs)) == 5)

status, page, url = get(op_user, "users.asp")
check("users.asp blocked for listener", "User form" not in page and "New user" not in page)

status, page, _ = get(op_user, "song_form.asp")
check("song_form.asp blocked for listener", "Save song" not in page)

status, page, _ = get(op_user, "user_form.asp?id=2")
check("user_form.asp blocked for listener", "Save user" not in page)

status, page, _ = get(op_user, "playlists.asp")
check("listener opens playlists", status == 200 and "Playlists" in page)

# ---------------------------------------------------------------- assets
print("\nStatic assets")
for asset in ["assets/css/polimusic.css", "assets/js/polimusic.js", "assets/img/favicon.svg",
              "assets/img/art-placeholder.svg", "albumImages/album1.png",
              "ArtistImages/artist1.jpg", "userImages/default.jpg",
              "playlistImages/playlist1.png", "songFiles/bensound-adventure.mp3"]:
    status, body, _ = get(op_user, asset)
    check(f"GET {asset}", status == 200, f"HTTP {status}")

ok = sum(1 for _, r in results if r)
print(f"\n{ok}/{len(results)} checks passed")
raise SystemExit(0 if ok == len(results) else 1)
