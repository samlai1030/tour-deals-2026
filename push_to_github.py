#!/usr/bin/env python3
"""Push this directory's files to GitHub repo samlai1030/tour-deals-2026 via Contents API.
Creates the repo if it does not exist. Uses the stored custom.github credential."""
import base64
import json
import os
import sys
import urllib.parse
import urllib.request

sys.path.insert(0, "/opt/hatch/skills/skill-creator/bin")
from dynamic_credentials import add_surrogate_to_request, read_json_response

API = "https://api.github.com"
ALLOWED = ["api.github.com"]
CRED = "custom.github"
OWNER = "samlai1030"
REPO = "tour-deals-2026"

FILES = [
    "README.md",
    "tours.json",
    "build_xlsx.py",
    "push_to_github.py",
    "tour-deals-2026-0925-1011.xlsx",
]


def api(method, path, body=None):
    req = urllib.request.Request(
        API + path,
        method=method,
        headers={"Accept": "application/vnd.github+json",
                 "User-Agent": "tony-assistant",
                 "Content-Type": "application/json"},
        data=json.dumps(body).encode() if body is not None else None,
    )
    add_surrogate_to_request(req, CRED, entry_name="access_token",
                             allowed_hosts=ALLOWED)
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return read_json_response(resp)
    except urllib.error.HTTPError as e:
        try:
            detail = json.loads(e.read().decode())
        except Exception:
            detail = {}
        return {"__error__": e.code, "detail": detail}


def ensure_repo():
    r = api("GET", f"/repos/{OWNER}/{REPO}")
    if "__error__" not in r:
        print(f"repo exists: {r['full_name']}")
        return
    if r["__error__"] == 404:
        r = api("POST", "/user/repos", {
            "name": REPO,
            "description": "2026-09-25～10-11 團體旅遊比價（8 家台灣旅行社）",
            "private": False,
            "auto_init": False,
        })
        if "__error__" in r:
            sys.exit(f"create repo failed: {r}")
        print(f"repo created: {r['full_name']}")
    else:
        sys.exit(f"repo check failed: {r}")


def push_file(relpath):
    full = os.path.join(os.path.dirname(os.path.abspath(__file__)), relpath)
    with open(full, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    urlpath = urllib.parse.quote(relpath)
    existing = api("GET", f"/repos/{OWNER}/{REPO}/contents/{urlpath}")
    body = {"message": f"update {relpath}", "content": b64, "branch": "main"}
    if "__error__" not in existing:
        body["sha"] = existing["sha"]
    r = api("PUT", f"/repos/{OWNER}/{REPO}/contents/{urlpath}", body)
    if "__error__" in r:
        sys.exit(f"push {relpath} failed: {r}")
    print(f"pushed {relpath} -> {r['content']['path']}")


def main():
    ensure_repo()
    for f in FILES:
        push_file(f)
    print("done")


if __name__ == "__main__":
    main()
