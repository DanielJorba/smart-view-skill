#!/usr/bin/env python3
"""Read-only client for the Oracle EPM Cloud REST APIs (Planning, FCCS, FreeForm).

Part of the smart-view skill: it registers each cloud once, caches its cubes and
hierarchies for fast lookups and gives Smart View the connection to open (sv).

Every request goes through _guard: GET on the EPM read APIs and one POST
(exportdataslice, which only reads). Passwords live in Windows Credential
Manager and are typed by the user in their own terminal (cred-set).
"""
from __future__ import annotations

import argparse
import base64
import csv
import ctypes
import datetime as dt
import functools
import getpass
import hashlib
import hmac
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

VERSION = "1.1.1"
FIGURES_ENV = "SMARTVIEW_SKILL_ALLOW_FIGURES"
CRED_PREFIX = "smart-view-skill:"
LEGACY_CRED_PREFIX = "epm-cloud-skill:"  # credentials saved before the merge
LEGACY_HOME = "epm-cloud-skill"
API = "/HyperionPlanning/rest/v3"
INTEROP = "/interop/rest/11.1.2.3.600"

READ_PREFIXES = ("/HyperionPlanning/rest/", "/interop/rest/", "/aif/rest/")
POST_ALLOWED = re.compile(r"^/HyperionPlanning/rest/v3/applications/[^/]+/plantypes/[^/]+/exportdataslice$")
RAW_GET_DENY = re.compile(r"(?i)(run|execute|import|reset|recreate|restart|copy|delete|refresh|clear|upload|rename)")
NUMBER = re.compile(r"^\s*[-+]?(\d+([.,]\d*)?|[.,]\d+)([eE][-+]?\d+)?\s*$")
JOB_STATUS = {-1: "in progress", 0: "success", 1: "error", 2: "cancel pending", 3: "cancelled", 4: "invalid parameter"}
CUBE_TYPES = {0: "BSO", 136: "ASO"}
LOOPBACK = {"127.0.0.1", "localhost", "::1"}
EMAIL = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
USER_IN_FILE = re.compile(r"[A-Za-z0-9._%+-]+_[A-Za-z0-9-]+\.[A-Za-z]{2,}_")  # Oracle writes user@domain as user_domain_
MAX_TEXT_KB = 512


def redact(text: str) -> str:
    return USER_IN_FILE.sub("<user>_", EMAIL.sub("<user>", text))


class EpmError(Exception):
    pass


def fail(msg: str, code: int = 2):
    print(f"ERROR: {msg}", file=sys.stderr)
    sys.exit(code)


# ---------- local storage ----------

def home() -> Path:
    base = os.environ.get("SMARTVIEW_SKILL_HOME")
    if base:
        p = Path(base)
    else:
        root = Path(os.environ.get("LOCALAPPDATA") or Path.home() / ".local" / "share")
        p = root / "smart-view-skill"
        _migrate(root / LEGACY_HOME, p)
    p.mkdir(parents=True, exist_ok=True)
    return p


def _migrate(old: Path, new: Path):
    """Copy pods, cache and digest salt from the former epm-cloud skill once; the old folder stays."""
    if not (old / "pods.json").exists() or (new / "pods.json").exists():
        return
    import shutil
    new.mkdir(parents=True, exist_ok=True)
    for name in ("pods.json", "digest.key"):
        if (old / name).exists():
            shutil.copy2(old / name, new / name)
    for name in ("cache", "logs"):
        if (old / name).is_dir() and not (new / name).exists():
            shutil.copytree(old / name, new / name)


def get_cred(name: str):
    return cred_read(CRED_PREFIX + name) or cred_read(LEGACY_CRED_PREFIX + name)


def load_pods() -> dict:
    f = home() / "pods.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}


def save_pods(pods: dict):
    (home() / "pods.json").write_text(json.dumps(pods, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def get_pod(name: str) -> dict:
    pods = load_pods()
    if name not in pods:
        fail(f"Unknown pod '{name}'. Known pods: {', '.join(sorted(pods)) or 'none'}. Add it with pod-add.")
    return {**pods[name], "name": name}


def normalize_url(url: str) -> str:
    parts = urllib.parse.urlsplit(url.strip())
    host = parts.hostname or ""
    if parts.scheme != "https" and not (parts.scheme == "http" and host in LOOPBACK):
        raise EpmError(f"Pod URL must be https: {url}")
    if not host:
        raise EpmError(f"Pod URL has no host: {url}")
    return f"{parts.scheme}://{parts.netloc}"


def figures_allowed() -> tuple[bool, str]:
    value = os.environ.get(FIGURES_ENV)
    source = "process"
    if value is None and os.name == "nt":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
                value = str(winreg.QueryValueEx(key, FIGURES_ENV)[0])
                source = "user registry"
        except OSError:
            value = None
    return value == "1", (source if value is not None else "not set")


def log_call(pod: dict, method: str, path: str, status, t0: float, size: int):
    f = home() / "logs" / "calls.log"
    f.parent.mkdir(exist_ok=True)
    stamp = dt.datetime.now().isoformat(timespec="seconds")
    ms = int((time.time() - t0) * 1000)
    with f.open("a", encoding="utf-8") as fh:
        fh.write(f"{stamp}\t{pod['name']}\t{method}\t{path}\t{status}\t{ms}ms\t{size}B\n")


# ---------- Windows Credential Manager ----------

@functools.lru_cache(maxsize=1)
def _cred_api():
    if os.name != "nt":
        raise EpmError("Credentials are stored in Windows Credential Manager; this needs Windows.")
    from ctypes import wintypes

    class Credential(ctypes.Structure):
        _fields_ = [
            ("Flags", wintypes.DWORD),
            ("Type", wintypes.DWORD),
            ("TargetName", wintypes.LPWSTR),
            ("Comment", wintypes.LPWSTR),
            ("LastWritten", wintypes.FILETIME),
            ("CredentialBlobSize", wintypes.DWORD),
            ("CredentialBlob", ctypes.POINTER(ctypes.c_char)),
            ("Persist", wintypes.DWORD),
            ("AttributeCount", wintypes.DWORD),
            ("Attributes", ctypes.c_void_p),
            ("TargetAlias", wintypes.LPWSTR),
            ("UserName", wintypes.LPWSTR),
        ]

    adv = ctypes.WinDLL("advapi32", use_last_error=True)
    adv.CredWriteW.argtypes = [ctypes.POINTER(Credential), wintypes.DWORD]
    adv.CredWriteW.restype = wintypes.BOOL
    adv.CredReadW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                              ctypes.POINTER(ctypes.POINTER(Credential))]
    adv.CredReadW.restype = wintypes.BOOL
    adv.CredDeleteW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD]
    adv.CredDeleteW.restype = wintypes.BOOL
    adv.CredFree.argtypes = [ctypes.c_void_p]
    adv.CredFree.restype = None
    return adv, Credential


CRED_TYPE_GENERIC = 1
CRED_PERSIST_LOCAL_MACHINE = 2
ERROR_NOT_FOUND = 1168


def cred_write(target: str, user: str, password: str):
    adv, Credential = _cred_api()
    blob = password.encode("utf-16-le")
    buf = ctypes.create_string_buffer(blob, len(blob))
    cred = Credential()
    cred.Type = CRED_TYPE_GENERIC
    cred.TargetName = target
    cred.CredentialBlobSize = len(blob)
    cred.CredentialBlob = ctypes.cast(buf, ctypes.POINTER(ctypes.c_char))
    cred.Persist = CRED_PERSIST_LOCAL_MACHINE
    cred.UserName = user
    if not adv.CredWriteW(ctypes.byref(cred), 0):
        raise ctypes.WinError(ctypes.get_last_error())


def cred_read(target: str):
    adv, Credential = _cred_api()
    ptr = ctypes.POINTER(Credential)()
    if not adv.CredReadW(target, CRED_TYPE_GENERIC, 0, ctypes.byref(ptr)):
        err = ctypes.get_last_error()
        if err == ERROR_NOT_FOUND:
            return None
        raise ctypes.WinError(err)
    try:
        c = ptr.contents
        blob = ctypes.string_at(c.CredentialBlob, c.CredentialBlobSize)
        return c.UserName or "", blob.decode("utf-16-le")
    finally:
        adv.CredFree(ctypes.cast(ptr, ctypes.c_void_p))


def cred_delete(target: str) -> bool:
    adv, _ = _cred_api()
    if adv.CredDeleteW(target, CRED_TYPE_GENERIC, 0):
        return True
    if ctypes.get_last_error() == ERROR_NOT_FOUND:
        return False
    raise ctypes.WinError(ctypes.get_last_error())


# ---------- HTTP ----------

def _guard(method: str, path: str):
    p = urllib.parse.urlsplit(path).path
    if not p.startswith(READ_PREFIXES):
        raise EpmError(f"Blocked: {p} is outside the EPM REST APIs this client reads.")
    if method == "GET":
        return
    if method == "POST" and POST_ALLOWED.match(p):
        return
    raise EpmError(f"Blocked: {method} {p}. This client only reads.")


def _http_message(code: int, payload: bytes) -> str:
    text = payload.decode("utf-8", "replace").strip()
    detail = text[:400]
    try:
        body = json.loads(text)
        if isinstance(body, dict):
            detail = body.get("detail") or body.get("message") or body.get("details") or body.get("error") or detail
    except ValueError:
        pass
    hints = {
        401: "authentication failed. OCI pods take the user without identity domain; users with MFA "
             "or SSO-only users cannot use basic authentication. Re-save the password with cred-set.",
        403: "the user's role does not allow this call (Get Member, job definitions, job status and "
             "file listing need Service Administrator).",
        404: "not found: check application, cube, dimension and member names.",
    }
    hint = hints.get(code, "")
    return f"HTTP {code}: {detail}" + (f"\n  Hint: {hint}" if hint else "")


def request(pod: dict, method: str, path: str, body=None, raw: bool = False, timeout: int = 180):
    _guard(method, path)
    cred = get_cred(pod["name"])
    if not cred:
        raise EpmError(f"No credential for pod '{pod['name']}'. The user saves it in their own terminal with: "
                       f"python epm.py cred-set {pod['name']}")
    user, password = cred
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(pod["url"].rstrip("/") + path, data=data, method=method)
    req.add_header("Authorization", "Basic " + base64.b64encode(f"{user}:{password}".encode("utf-8")).decode())
    req.add_header("Accept", "application/json")
    if data is not None:
        req.add_header("Content-Type", "application/json")
    t0 = time.time()
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            status, payload = resp.status, resp.read()
    except urllib.error.HTTPError as e:
        payload = e.read()
        log_call(pod, method, path, e.code, t0, len(payload))
        raise EpmError(_http_message(e.code, payload)) from None
    except urllib.error.URLError as e:
        log_call(pod, method, path, "network", t0, 0)
        raise EpmError(f"Cannot reach {pod['url']}: {e.reason}") from None
    log_call(pod, method, path, status, t0, len(payload))
    if raw:
        return payload
    if not payload.strip():
        return {}
    try:
        return json.loads(payload)
    except ValueError:
        if b"<html" in payload[:2000].lower():
            raise EpmError("The pod answered with an HTML page (probably an SSO sign-in page): "
                           "this user needs native credentials for REST.") from None
        raise EpmError(f"Unexpected non-JSON response: {payload[:300]!r}") from None


def q(s: str) -> str:
    return urllib.parse.quote(s, safe="")


def app_of(pod: dict, args) -> str:
    app = getattr(args, "app", None) or pod.get("app")
    if not app:
        fail(f"Pod '{pod['name']}' has no application; pass --app.")
    return app


def cube_of(pod: dict, args) -> str:
    cube = getattr(args, "cube", None) or (pod.get("cubes") or [None])[0]
    if not cube:
        fail(f"Pod '{pod['name']}' has no default cube; pass --cube.")
    return cube


# ---------- dimension cache ----------

def _safe(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", name)


def dim_cache_file(pod: dict, app: str, cube: str, dim: str) -> Path:
    return home() / "cache" / _safe(pod["name"]) / _safe(app) / _safe(cube) / f"{_safe(dim)}.json"


def load_dimension(pod: dict, app: str, cube: str, dim: str, refresh: bool = False, alias_table=None) -> dict:
    f = dim_cache_file(pod, app, cube, dim)
    if f.exists() and not refresh:
        return json.loads(f.read_text(encoding="utf-8"))
    path = f"{API}/applications/{q(app)}/plantypes/{q(cube)}/dimensions/{q(dim)}"
    if alias_table:
        path += "?" + urllib.parse.urlencode({"aliasTableName": alias_table})
    data = request(pod, "GET", path)
    data["_fetched"] = dt.datetime.now().isoformat(timespec="seconds")
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return data


def walk(node: dict, depth: int = 0):
    yield node, depth
    for child in node.get("children") or []:
        yield from walk(child, depth + 1)


def node_line(node: dict, depth: int) -> str:
    name = node.get("name", "?")
    alias = node.get("alias")
    tags = [t for t in (node.get("dataStorage"), f"L{node['level']}" if "level" in node else None) if t]
    line = "  " * depth + name
    if alias and alias != name:
        line += f" ({alias})"
    if tags:
        line += "  [" + ", ".join(tags) + "]"
    return line


def render_tree(root: dict, max_depth: int, max_lines: int) -> list[str]:
    lines, hidden = [], 0
    for node, depth in walk(root):
        if depth > max_depth:
            continue
        if len(lines) >= max_lines:
            hidden += 1
            continue
        line = node_line(node, depth)
        kids = node.get("children") or []
        if depth == max_depth and kids:
            line += f"  +{len(kids)} children"
        lines.append(line)
    if hidden:
        lines.append(f"... {hidden} more lines hidden (raise --max-lines or start from a lower member)")
    return lines


def find_nodes(root: dict, member: str) -> list[dict]:
    exact = [n for n, _ in walk(root) if n.get("name") == member]
    return exact or [n for n, _ in walk(root) if (n.get("name") or "").lower() == member.lower()]


# ---------- data slice ----------

def split_members(spec: str) -> list[str]:
    out, cur, depth, quoted = [], [], 0, False
    for ch in spec:
        if ch == '"':
            quoted = not quoted
        elif not quoted and ch == "(":
            depth += 1
        elif not quoted and ch == ")":
            depth -= 1
        if ch == "," and depth == 0 and not quoted:
            out.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    out.append("".join(cur))
    members = [m.strip().strip('"').strip() for m in out]
    return [m for m in members if m]


def parse_axis(specs: list[str]) -> tuple[list[str], list[list[str]]]:
    dims, members = [], []
    for spec in specs or []:
        if "=" not in spec:
            raise EpmError(f"Axis needs Dimension=members: {spec}")
        dim, mbrs = spec.split("=", 1)
        parsed = split_members(mbrs)
        if not dim.strip() or not parsed:
            raise EpmError(f"Axis needs Dimension=members: {spec}")
        dims.append(dim.strip())
        members.append(parsed)
    return dims, members


def parse_pov(spec: str) -> tuple[list[str], list[list[str]]]:
    dims, members = [], []
    for part in (p.strip() for p in (spec or "").split(";")):
        if not part:
            continue
        if "=" not in part:
            raise EpmError(f"POV needs Dimension=member pairs separated by ';': {part}")
        dim, mbr = part.split("=", 1)
        dims.append(dim.strip())
        members.append([mbr.strip().strip('"')])
    return dims, members


def build_slice_payload(pov: str, rows: list[str], cols: list[str], suppress: bool = True) -> dict:
    pov_dims, pov_mbrs = parse_pov(pov)
    row_dims, row_mbrs = parse_axis(rows)
    col_dims, col_mbrs = parse_axis(cols)
    if not row_dims or not col_dims:
        raise EpmError("A slice needs at least one --rows and one --cols.")
    used = pov_dims + row_dims + col_dims
    dupes = sorted({d for d in used if used.count(d) > 1})
    if dupes:
        raise EpmError(f"Dimension used twice: {', '.join(dupes)}")
    return {
        "exportPlanningData": False,
        "gridDefinition": {
            "suppressMissingBlocks": suppress,
            "suppressMissingRows": suppress,
            "suppressMissingColumns": suppress,
            "pov": {"dimensions": pov_dims, "members": pov_mbrs},
            "columns": [{"dimensions": col_dims, "members": col_mbrs}],
            "rows": [{"dimensions": row_dims, "members": row_mbrs}],
        },
    }


def slice_table(resp: dict) -> tuple[list[str], list[str], list[tuple[list[str], list[str]]]]:
    cols = resp.get("columns") or []
    ncols = max((len(c) for c in cols), default=0)
    labels = [" / ".join(str(c[i]) for c in cols if i < len(c)) for i in range(ncols)]
    rows = [([str(h) for h in r.get("headers") or []], ["" if v is None else str(v).strip() for v in r.get("data") or []])
            for r in resp.get("rows") or []]
    return [str(p) for p in resp.get("pov") or []], labels, rows


def mask(value: str) -> str:
    return "<n>" if NUMBER.match(value) else value


def digest_key() -> bytes:
    f = home() / "digest.key"
    if not f.exists():
        f.write_bytes(os.urandom(32))
    return f.read_bytes()


def digest(key: bytes, values: list[str]) -> str:
    """Salted fingerprint of a list of cells: same digest means same values, without showing them."""
    normalized = [str(as_number(v)) if v else "" for v in values]
    return hmac.new(key, "\x1f".join(normalized).encode("utf-8"), hashlib.sha256).hexdigest()[:12]


def digest_lines(labels: list[str], rows) -> list[str]:
    key = digest_key()
    lines = ["Digests (salted on this PC; equal digest = equal values; no figures shown)"]
    for hdrs, data in rows:
        lines.append("row\t" + " | ".join(hdrs) + "\t" + digest(key, data))
    for i, label in enumerate(labels):
        lines.append("col\t" + label + "\t" + digest(key, [d[i] if i < len(d) else "" for _, d in rows]))
    lines.append("grid\tall\t" + digest(key, [v for _, d in rows for v in d]))
    return lines


def as_number(value: str):
    if not NUMBER.match(value):
        return value
    try:
        return float(value.replace(",", "."))
    except ValueError:
        return value


def write_slice_file(out: Path, meta: list[tuple[str, str]], header: list[str], rows) -> Path:
    out.parent.mkdir(parents=True, exist_ok=True)
    if out.suffix.lower() == ".xlsx":
        try:
            from openpyxl import Workbook
        except ImportError:
            out = out.with_suffix(".csv")
            print("openpyxl is not installed: writing CSV instead.", file=sys.stderr)
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Data"
            ws.append(header)
            for hdrs, data in rows:
                ws.append(hdrs + [as_number(v) if v else None for v in data])
            ws.freeze_panes = ws.cell(row=2, column=len(rows[0][0]) + 1 if rows else 1)
            info = wb.create_sheet("POV")
            info.append(["Item", "Value"])
            for k, v in meta:
                info.append([k, v])
            wb.save(out)
            return out
    with out.open("w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        for k, v in meta:
            w.writerow([f"# {k}", v])
        w.writerow(header)
        for hdrs, data in rows:
            w.writerow(hdrs + data)
    return out


# ---------- commands ----------

def cmd_status(args):
    allowed, source = figures_allowed()
    pods = load_pods()
    print(f"smart-view epm.py {VERSION}")
    print(f"Data folder: {home()}")
    print(f"Figures: {'SHOWN' if allowed else 'MASKED as <n>'} ({FIGURES_ENV}, {source})")
    if not pods:
        print("Pods: none. Add one with pod-add.")
    for name, pod in sorted(pods.items()):
        try:
            cred_text = "credential set" if get_cred(name) else "NO credential"
        except EpmError as e:
            cred_text = str(e)
        build = "BUILD enabled" if build_status({**pod, "name": name})[0] else "read-only"
        print(f"- {name} [{pod.get('env', '?')}] {pod['url']} app={pod.get('app')} "
              f"cubes={','.join(pod.get('cubes') or [])} · {cred_text} · {build}")


def cmd_pods(args):
    cmd_status(args)


def cmd_pod_add(args):
    url = normalize_url(args.url)
    host = urllib.parse.urlsplit(url).hostname or ""
    env = args.env or ("test" if "-test" in host else "prod")
    pods = load_pods()
    pods[args.pod] = {
        "url": url,
        "app": args.app,
        "cubes": [c.strip() for c in (args.cube or []) if c.strip()],
        "env": env,
        "notes": args.notes or "",
    }
    save_pods(pods)
    print(f"Pod '{args.pod}' saved ({env}): {url}. Next: the user saves the password in their own terminal with "
          f"'python epm.py cred-set {args.pod}'.")


def cmd_pod_remove(args):
    pods = load_pods()
    if pods.pop(args.pod, None) is None:
        fail(f"Unknown pod '{args.pod}'.")
    save_pods(pods)
    print(f"Pod '{args.pod}' removed. Its credential stays until cred-delete.")


def cmd_cred_set(args):
    pod = get_pod(args.pod)
    if not _interactive():
        fail("cred-set is interactive and only the user runs it, in their own terminal. "
             "Claude never types, receives or stores passwords.")
    print(f"Pod {pod['name']} ({pod['url']}). OCI pods: user without identity domain. Nothing you type is shown.")
    user = (args.user or getpass.getpass("User: ")).strip()
    if not user:
        fail("Empty user.")
    password = getpass.getpass("Password (not shown): ")
    if not password:
        fail("Empty password.")
    cred_write(CRED_PREFIX + pod["name"], user, password)
    print(f"Saved in Windows Credential Manager as '{CRED_PREFIX}{pod['name']}'. Testing...")
    cmd_ping(args)


def _interactive() -> bool:
    """True only for a real console. On Windows the NUL device also reports isatty()."""
    if not sys.stdin.isatty():
        return False
    if os.name != "nt":
        return True
    import msvcrt
    from ctypes import wintypes
    try:
        handle = msvcrt.get_osfhandle(sys.stdin.fileno())
    except (OSError, ValueError):
        return False
    mode = wintypes.DWORD()
    return bool(ctypes.windll.kernel32.GetConsoleMode(handle, ctypes.byref(mode)))


def pod_host(pod: dict) -> str:
    return urllib.parse.urlsplit(pod["url"]).hostname or ""


def build_status(pod: dict) -> tuple[bool, str]:
    host = pod_host(pod)
    build = pod.get("build") or {}
    if pod.get("env") != "test" or "-test" not in host:
        return False, f"{pod['name']} ({host}) is not a test pod: read-only."
    if build.get("host") != host:
        return False, (f"Building is not enabled on {pod['name']} ({host}): read-only. "
                       f"Only the user enables it, in their own terminal: epm.py allow-build {pod['name']}")
    return True, f"Building enabled on {pod['name']} ({host}) since {build.get('since')}."


def log_build(pod: dict, text: str):
    f = home() / "logs" / "build.log"
    f.parent.mkdir(exist_ok=True)
    stamp = dt.datetime.now().isoformat(timespec="seconds")
    with f.open("a", encoding="utf-8") as fh:
        fh.write(f"{stamp}\t{pod['name']}\t{pod_host(pod)}\t{redact(text)}\n")


def cmd_allow_build(args):
    pod = get_pod(args.pod)
    host = pod_host(pod)
    if pod.get("env") != "test" or "-test" not in host:
        fail(f"{pod['name']} ({host}) is not a test pod. Building is only possible on hosts with '-test'; "
             "production stays read-only.")
    if not _interactive():
        fail("allow-build is interactive and only the user runs it, in their own terminal.")
    print(f"Claude will be able to BUILD in {pod['name']} ({host}): create and change what you ask, never execute.")
    try:
        confirmed = input(f"Type the pod name ({pod['name']}) to confirm: ").strip() == pod["name"]
    except EOFError:
        confirmed = False
    if not confirmed:
        fail("Not confirmed. Nothing changed.")
    pods = load_pods()
    pods[pod["name"]]["build"] = {"host": host, "since": dt.datetime.now().isoformat(timespec="seconds")}
    save_pods(pods)
    log_build(pod, "allow-build by the user")
    print(f"Enabled. Back to read-only at any time with: epm.py deny-build {pod['name']}")


def cmd_deny_build(args):
    pod = get_pod(args.pod)
    pods = load_pods()
    had = pods[pod["name"]].pop("build", None) is not None
    save_pods(pods)
    if had:
        log_build(pod, "deny-build")
    print(f"{pod['name']} is read-only.")


def cmd_build_check(args):
    pod = get_pod(args.pod)
    ok, msg = build_status(pod)
    print(msg)
    if not ok:
        sys.exit(3)
    print(f"Before every change the browser address bar must start with https://{pod_host(pod)}/")


def cmd_buildlog(args):
    pod = get_pod(args.pod)
    if args.show:
        f = home() / "logs" / "build.log"
        lines = [l for l in (f.read_text(encoding="utf-8").splitlines() if f.exists() else [])
                 if l.split("\t")[1:2] == [pod["name"]]]
        print("\n".join(lines[-args.last:]) or f"No build entries for {pod['name']}.")
        return
    ok, msg = build_status(pod)
    if not ok:
        fail(msg, 3)
    text = " ".join(args.text).strip()
    if not text:
        fail("Nothing to log.")
    log_build(pod, text)
    print("Logged.")


def cmd_cred_show(args):
    pod = get_pod(args.pod)
    cred = get_cred(pod["name"])
    print(f"{pod['name']}: " + ("credential set (user and password are never shown)" if cred else "no credential"))


def cmd_cred_delete(args):
    pod = get_pod(args.pod)
    deleted = [cred_delete(prefix + pod["name"]) for prefix in (CRED_PREFIX, LEGACY_CRED_PREFIX)]
    print("Deleted." if any(deleted) else "There was no credential.")


def sv_connection(pod: dict, app: str, cube: str) -> tuple[str, str]:
    host = pod_host(pod)
    url = f"https://{host}/HyperionPlanning/SmartView"
    return url, f"{url}|{host}|{app}|{cube}"


def cmd_sv(args):
    pod = get_pod(args.pod)
    app, cube = app_of(pod, args), cube_of(pod, args)
    url, inline = sv_connection(pod, app, cube)
    print(f"Smart View URL: {url}")
    print(f"Inline connection: {inline}")
    print(f"Bridge: sv.ps1 call UIConnect \"\" <sheet> \"pod:{pod['name']}" + (f"/{cube}" if args.cube else "") + "\"")


def cmd_ping(args):
    pod = get_pod(args.pod)
    data = request(pod, "GET", f"{API}/applications")
    apps = [i.get("name") for i in data.get("items") or []]
    print(f"OK: signed in to {pod['name']} ({pod['url']}). Applications: {', '.join(map(str, apps)) or 'none visible'}")
    if pod.get("app") and pod["app"] not in apps:
        print(f"Warning: configured application '{pod['app']}' is not in that list.")


def cmd_apps(args):
    pod = get_pod(args.pod)
    for item in request(pod, "GET", f"{API}/applications").get("items") or []:
        extra = {k: v for k, v in item.items() if k not in ("links", "name")}
        print(f"{item.get('name')}\t{json.dumps(extra, ensure_ascii=False)}")


def cmd_cubes(args):
    pod = get_pod(args.pod)
    app = app_of(pod, args)
    for item in request(pod, "GET", f"{API}/applications/{q(app)}/plantypes").get("items") or []:
        kind = CUBE_TYPES.get(item.get("cubeType"), item.get("cubeType"))
        print(f"{item.get('planTypeName')}\tcube={item.get('cubeName')}\t{kind}\tdimensions={item.get('numDimensions')}")


def cmd_dims(args):
    pod = get_pod(args.pod)
    app, cube = app_of(pod, args), cube_of(pod, args)
    path = f"{API}/applications/{q(app)}/plantypes/{q(cube)}/dimensions?limit=-1"
    items = request(pod, "GET", path).get("items") or []
    f = dim_cache_file(pod, app, cube, "_dimensions")
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")
    print(f"{app}.{cube}: {len(items)} dimensions")
    for item in items:
        print(f"{item.get('name')}\t{item.get('dimType')}")


def cmd_tree(args):
    pod = get_pod(args.pod)
    app, cube = app_of(pod, args), cube_of(pod, args)
    root = load_dimension(pod, app, cube, args.dim, args.refresh, args.alias_table)
    total = sum(1 for _ in walk(root)) - 1
    print(f"{args.dim} in {app}.{cube}: {total} members (cached {root.get('_fetched', '?')}; --refresh to reload)")
    starts = [root] if not args.member else find_nodes(root, args.member)
    if not starts:
        fail(f"Member '{args.member}' not found in {args.dim}. Try: find {pod['name']} {args.dim} <text>", 1)
    if len(starts) > 1:
        print(f"'{args.member}' appears {len(starts)} times (shared members): showing each.")
    for node in starts:
        if args.member:
            print(f"-- {node.get('path') or node.get('name')}")
        print("\n".join(render_tree(node, args.depth, args.max_lines)))


def cmd_find(args):
    pod = get_pod(args.pod)
    app, cube = app_of(pod, args), cube_of(pod, args)
    root = load_dimension(pod, app, cube, args.dim, args.refresh, args.alias_table)
    needle = args.text.lower()
    hits = [n for n, d in walk(root) if d > 0 and (needle in (n.get("name") or "").lower()
                                                   or needle in (n.get("alias") or "").lower())]
    print(f"{len(hits)} matches for '{args.text}' in {args.dim}")
    for n in hits[: args.limit]:
        print(node_line(n, 0) + f"   path={n.get('path') or '?'}")
    if len(hits) > args.limit:
        print(f"... {len(hits) - args.limit} more (raise --limit)")


def cmd_member(args):
    pod = get_pod(args.pod)
    app = app_of(pod, args)
    data = request(pod, "GET", f"{API}/applications/{q(app)}/dimensions/{q(args.dim)}/members/{q(args.member)}")
    data.pop("links", None)
    print(json.dumps(data, indent=2, ensure_ascii=False))


def cmd_subvars(args):
    pod = get_pod(args.pod)
    app = app_of(pod, args)
    items = request(pod, "GET", f"{API}/applications/{q(app)}/substitutionvariables").get("items") or []
    print(f"{app} (application): {len(items)} variables")
    for item in items:
        print(f"  {item.get('name')} = {item.get('value')}  [{item.get('planType', '-')}]")
    cubes = [args.cube] if args.cube else (pod.get("cubes") or [])
    for cube in cubes:
        items = request(pod, "GET", f"{API}/applications/{q(app)}/plantypes/{q(cube)}/substitutionvariables").get("items") or []
        print(f"{cube}: {len(items)} variables")
        for item in items:
            print(f"  {item.get('name')} = {item.get('value')}")


RULE_TYPES = ("RULES", "RULESET")


def load_jobdefs(pod: dict, app: str, refresh: bool = False) -> dict:
    """All job definitions of the application, cached: one call answers jobdefs and rules."""
    f = home() / "cache" / _safe(pod["name"]) / _safe(app) / "_jobdefs.json"
    if f.exists() and not refresh:
        return json.loads(f.read_text(encoding="utf-8"))
    items = request(pod, "GET", f"{API}/applications/{q(app)}/jobdefinitions").get("items") or []
    for item in items:
        item.pop("links", None)
    data = {"_fetched": dt.datetime.now().isoformat(timespec="seconds"), "items": items}
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return data


def cmd_jobdefs(args):
    pod = get_pod(args.pod)
    app = app_of(pod, args)
    data = load_jobdefs(pod, app, args.refresh)
    items = [i for i in data["items"] if not args.type or (i.get("jobType") or "").upper() == args.type.upper()]
    print(f"{len(items)} job definitions (cached {data['_fetched']}; --refresh to reload)")
    for item in items:
        print(f"{item.get('jobType')}\t{item.get('jobName')}")


def cmd_rules(args):
    pod = get_pod(args.pod)
    app = app_of(pod, args)
    data = load_jobdefs(pod, app, args.refresh)
    needle = (args.text or "").lower()
    rules = [i for i in data["items"] if (i.get("jobType") or "").upper() in RULE_TYPES
             and needle in (i.get("jobName") or "").lower()]
    rules.sort(key=lambda i: ((i.get("jobType") or ""), (i.get("jobName") or "").lower()))
    what = f" matching '{args.text}'" if args.text else ""
    print(f"{app}: {len(rules)} rules and rulesets{what} (cached {data['_fetched']}; --refresh to reload)")
    for item in rules[: args.limit]:
        print(f"{item.get('jobType')}\t{item.get('jobName')}")
    if len(rules) > args.limit:
        print(f"... {len(rules) - args.limit} more (raise --limit or add text to search)")


def cmd_job(args):
    pod = get_pod(args.pod)
    app = app_of(pod, args)
    data = request(pod, "GET", f"{API}/applications/{q(app)}/jobs/{q(str(args.job_id))}")
    status = data.get("status")
    print(f"Job {data.get('jobId', args.job_id)} {data.get('jobName', '')}: "
          f"{JOB_STATUS.get(status, status)} ({data.get('descriptiveStatus') or data.get('jobStatus') or '-'})")
    if data.get("details"):
        print(f"Details: {data['details']}")


def cmd_files(args):
    pod = get_pod(args.pod)
    items = request(pod, "GET", f"{INTEROP}/applicationsnapshots").get("items") or []
    print(f"{len(items)} files and snapshots")
    for item in items:
        modified = item.get("lastmodifiedtime")
        when = dt.datetime.fromtimestamp(int(modified) / 1000).isoformat(timespec="minutes") if modified else "-"
        print(f"{item.get('type')}\t{when}\t{item.get('size') or '-'}\t{redact(str(item.get('name')))}")


def cmd_cat(args):
    pod = get_pod(args.pod)
    data = request(pod, "GET", f"{INTEROP}/applicationsnapshots/{q(args.file)}/contents", raw=True)
    if data[:2] == b"PK" or b"\x00" in data[:4096]:
        fail(f"{args.file} is binary ({len(data)} bytes): cat only prints text files.", 1)
    text = data[: args.max_kb * 1024].decode("utf-8", "replace")
    print(redact(text))
    if len(data) > args.max_kb * 1024:
        print(f"... truncated at {args.max_kb} KB of {len(data) // 1024} KB (raise --max-kb)")


def cmd_dijob(args):
    pod = get_pod(args.pod)
    data = request(pod, "GET", f"/aif/rest/V1/jobs/{q(str(args.job_id))}")
    status = data.get("status")
    try:
        status = int(status)
    except (TypeError, ValueError):
        pass
    print(f"Process {data.get('jobId', args.job_id)} {data.get('processType') or data.get('jobName') or ''}: "
          f"{JOB_STATUS.get(status, status)} ({data.get('jobStatus') or data.get('descriptiveStatus') or '-'})")
    for key in ("logFileName", "outputFileName", "details"):
        if data.get(key) not in (None, "null"):
            print(f"{key}: {redact(str(data[key]))}")


def cmd_slice(args):
    pod = get_pod(args.pod)
    app, cube = app_of(pod, args), cube_of(pod, args)
    if args.grid:
        payload = json.loads(Path(args.grid).read_text(encoding="utf-8"))
        row_dims = (payload.get("gridDefinition", {}).get("rows") or [{}])[0].get("dimensions") or []
    else:
        payload = build_slice_payload(args.pov, args.rows, args.cols, not args.no_suppress)
        row_dims = payload["gridDefinition"]["rows"][0]["dimensions"]
        dims_file = dim_cache_file(pod, app, cube, "_dimensions")
        if dims_file.exists():
            known = {d.get("name") for d in json.loads(dims_file.read_text(encoding="utf-8"))}
            grid = payload["gridDefinition"]
            used = set(grid["pov"]["dimensions"] + grid["rows"][0]["dimensions"] + grid["columns"][0]["dimensions"])
            missing = sorted(d for d in known - used if d)
            if missing:
                print(f"Note: not in the grid, the cube decides their member: {', '.join(missing)}", file=sys.stderr)
    resp = request(pod, "POST", f"{API}/applications/{q(app)}/plantypes/{q(cube)}/exportdataslice", payload)
    pov, labels, rows = slice_table(resp)
    allowed, _ = figures_allowed()
    show = allowed and not args.hide_figures
    if args.out and not show:
        fail(f"--out writes the figures to a file, so it needs {FIGURES_ENV}=1 (set by the user). "
             "Without it use the masked grid, --digest or the counts of cells with a value.")
    filled = sum(1 for _, data in rows for v in data if v)
    print(f"{app}.{cube} POV: {' | '.join(pov) or '-'}")
    print(f"Rows {len(rows)} · columns {len(labels)} · cells with a value {filled}"
          + ("" if show else " · figures masked as <n>"))
    header = [d for d in row_dims] if row_dims else [f"Row{i + 1}" for i in range(len(rows[0][0]) if rows else 1)]
    if rows and len(header) != len(rows[0][0]):
        header = [f"Row{i + 1}" for i in range(len(rows[0][0]))]
    header += labels
    if args.out:
        meta = [("Pod", pod["name"]), ("Application", app), ("Cube", cube), ("POV", " | ".join(pov)),
                ("Exported", dt.datetime.now().isoformat(timespec="seconds"))]
        out = write_slice_file(Path(args.out).expanduser().resolve(), meta, header, rows)
        print(f"Saved: {out}")
        if args.open and os.name == "nt":
            os.startfile(out)
    if args.digest:
        print("\n".join(digest_lines(labels, rows)))
    if (args.out and not args.print) or (args.digest and not args.print):
        return
    print("\t".join(header))
    for hdrs, data in rows[: args.max_rows]:
        print("\t".join(hdrs + [v if show else mask(v) for v in data]))
    if len(rows) > args.max_rows:
        print(f"... {len(rows) - args.max_rows} more rows (use --out to get them all in a file)")


def unmangle_path(path: str) -> str:
    """Git Bash turns '/HyperionPlanning/...' into 'C:/Program Files/Git/HyperionPlanning/...'."""
    path = path.replace("\\", "/")
    if re.match(r"^/?[A-Za-z]:/", path):
        hits = [i for i in (path.find(p) for p in READ_PREFIXES) if i >= 0]
        if hits:
            return path[min(hits):]
    return path if path.startswith("/") else "/" + path


def cmd_get(args):
    pod = get_pod(args.pod)
    path = unmangle_path(args.path)
    if RAW_GET_DENY.search(urllib.parse.urlsplit(path).path):
        fail("Blocked: that path looks like an action, not a read.")
    data = request(pod, "GET", path, raw=True)
    text = data.decode("utf-8", "replace")
    try:
        text = json.dumps(json.loads(text), indent=2, ensure_ascii=False)
    except ValueError:
        pass
    print(redact(text))


# ---------- CLI ----------

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="epm.py", description="Read-only client for Oracle EPM Cloud REST APIs.")
    p.add_argument("--version", action="version", version=VERSION)
    sub = p.add_subparsers(dest="command", required=True)

    def add(name, func, help_text, pod=True, app=False, cube=False):
        sp = sub.add_parser(name, help=help_text)
        if pod:
            sp.add_argument("pod")
        if app:
            sp.add_argument("--app", help="application (default: the pod's)")
        if cube:
            sp.add_argument("--cube", help="cube / plan type (default: the pod's first cube)")
        sp.set_defaults(func=func)
        return sp

    add("status", cmd_status, "figures switch, pods and credentials", pod=False)
    add("pods", cmd_pods, "same as status", pod=False)
    sp = add("pod-add", cmd_pod_add, "register a pod (no secrets)")
    sp.add_argument("--url", required=True)
    sp.add_argument("--app")
    sp.add_argument("--cube", action="append", help="repeat for several cubes; the first is the default")
    sp.add_argument("--env", choices=["test", "prod"])
    sp.add_argument("--notes")
    add("pod-remove", cmd_pod_remove, "forget a pod")
    sp = add("cred-set", cmd_cred_set, "USER ONLY, interactive: save user and password")
    sp.add_argument("--user")
    add("cred-show", cmd_cred_show, "show whether a credential exists (never the password)")
    add("cred-delete", cmd_cred_delete, "delete the pod's credential")
    add("allow-build", cmd_allow_build, "USER ONLY, interactive: let Claude build (never execute) in a test pod")
    add("deny-build", cmd_deny_build, "put a pod back to read-only")
    add("build-check", cmd_build_check, "exit 0 if building is enabled on the pod, 3 if it is read-only")
    sp = add("buildlog", cmd_buildlog, "record what Claude built in the pod (or --show the log)")
    sp.add_argument("text", nargs="*")
    sp.add_argument("--show", action="store_true")
    sp.add_argument("--last", type=int, default=50)
    add("sv", cmd_sv, "Smart View URL and inline connection of the pod's cube (no sign-in needed)",
        app=True, cube=True)
    add("ping", cmd_ping, "check sign-in and list applications")
    add("apps", cmd_apps, "list applications")
    add("cubes", cmd_cubes, "list cubes (plan types) with BSO/ASO", app=True)
    add("dims", cmd_dims, "list the dimensions of a cube", app=True, cube=True)
    for name, func, text in (("tree", cmd_tree, "show a dimension hierarchy (cached)"),
                             ("find", cmd_find, "search members by name or alias (cached)")):
        sp = add(name, func, text, app=True, cube=True)
        sp.add_argument("dim")
        if name == "tree":
            sp.add_argument("member", nargs="?")
            sp.add_argument("--depth", type=int, default=2)
            sp.add_argument("--max-lines", type=int, default=300)
        else:
            sp.add_argument("text")
            sp.add_argument("--limit", type=int, default=100)
        sp.add_argument("--refresh", action="store_true", help="reload the dimension from the pod")
        sp.add_argument("--alias-table")
    sp = add("member", cmd_member, "member properties (Service Administrator)", app=True)
    sp.add_argument("dim")
    sp.add_argument("member")
    add("subvars", cmd_subvars, "substitution variables of a cube", app=True, cube=True)
    sp = add("jobdefs", cmd_jobdefs, "job definitions, cached (Service Administrator)", app=True)
    sp.add_argument("--type", help="RULES, RULESET, IMPORT_DATA, EXPORT_DATA, CUBE_REFRESH...")
    sp.add_argument("--refresh", action="store_true", help="reload the job definitions from the pod")
    sp = add("rules", cmd_rules, "business rules and rulesets, cached; search by part of the name "
             "(Service Administrator)", app=True)
    sp.add_argument("text", nargs="?")
    sp.add_argument("--limit", type=int, default=300)
    sp.add_argument("--refresh", action="store_true", help="reload the job definitions from the pod")
    sp = add("job", cmd_job, "status of a job by id (Service Administrator)", app=True)
    sp.add_argument("job_id")
    add("files", cmd_files, "files and snapshots in the pod (Service Administrator)")
    sp = add("cat", cmd_cat, "print a text file of the pod, e.g. outbox/logs/251.log (Service Administrator)")
    sp.add_argument("file")
    sp.add_argument("--max-kb", type=int, default=MAX_TEXT_KB)
    sp = add("dijob", cmd_dijob, "status of a Data Integration process (pipeline, integration) by id")
    sp.add_argument("job_id")
    sp = add("slice", cmd_slice, "read a data grid (exportdataslice)", app=True, cube=True)
    sp.add_argument("--pov", default="", help='"Dim=Member;Dim=Member"')
    sp.add_argument("--rows", action="append", help='"Dim=M1,M2,ILvl0Descendants(X)"; repeat to cross dimensions')
    sp.add_argument("--cols", action="append", help="same format as --rows")
    sp.add_argument("--grid", help="JSON file with a full exportdataslice payload instead of --pov/--rows/--cols")
    sp.add_argument("--no-suppress", action="store_true", help="keep empty rows and columns")
    sp.add_argument("--out", help="write all rows to .xlsx (or .csv) with the figures, for the user")
    sp.add_argument("--open", action="store_true", help="open the file after writing it")
    sp.add_argument("--print", action="store_true", help="also print the grid when --out is used")
    sp.add_argument("--max-rows", type=int, default=200)
    sp.add_argument("--hide-figures", action="store_true", help="mask figures in this output even if allowed")
    sp.add_argument("--digest", action="store_true",
                    help="print a salted fingerprint per row, per column and for the grid instead of the grid "
                         "(compare cells or dates without seeing figures; add --print for the grid too)")
    sp = add("get", cmd_get, "raw GET on an EPM REST path (read-only paths only)")
    sp.add_argument("path")
    return p


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except AttributeError:
            pass
    args = build_parser().parse_args(argv)
    try:
        args.func(args)
    except EpmError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
