"""Shared helpers: Entra device-code / refresh-token sign-in and a paced Microsoft Graph client.
Standard library only. Tokens are never printed."""
import base64, json, os, sys, time, urllib.error, urllib.parse, urllib.request

GRAPH = "https://graph.microsoft.com"
BASE = f"{GRAPH}/beta/copilot/costManagement"
GRAPH_CLI = "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Microsoft Graph Command Line Tools (Connect-MgGraph)
SCOPE = lambda name: f"{GRAPH}/{name}"
READ_SCOPES = [SCOPE(s) for s in (
    "CopilotCostManagement.Read.All", "CopilotCostManagement-Policy.Read.All",
    "CopilotCostManagement-UserData.Read.All", "CopilotCostManagement-Assignment.Read.All")]


def _post_form(url, data):
    req = urllib.request.Request(url, urllib.parse.urlencode(data).encode())
    try:
        with urllib.request.urlopen(req) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        return json.load(e)


def claims(token):
    p = token.split(".")[1]
    return json.loads(base64.urlsafe_b64decode(p + "=" * (-len(p) % 4)))


def device_code(tenant, client_id, scopes):
    """Interactive sign-in: prints the device-code message, returns the token response."""
    auth = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0"
    dc = _post_form(f"{auth}/devicecode", {"client_id": client_id, "scope": " ".join(scopes)})
    if "device_code" not in dc:
        sys.exit(f"device code request failed: {dc.get('error')}: {dc.get('error_description')}")
    print(dc["message"], flush=True)
    interval, deadline = dc.get("interval", 5), time.time() + dc.get("expires_in", 900)
    while time.time() < deadline:
        time.sleep(interval)
        t = _post_form(f"{auth}/token", {"grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                                         "client_id": client_id, "device_code": dc["device_code"]})
        if "access_token" in t:
            return t
        if t.get("error") == "slow_down":
            interval += 5
        elif t.get("error") != "authorization_pending":
            sys.exit(f"sign-in failed: {t.get('error')}: {t.get('error_description')}")
    sys.exit("sign-in timed out")


def refresh(tenant, client_id, refresh_token, scopes):
    """Redeem a refresh token. Returns the token response (with a rotated refresh_token) or the error."""
    return _post_form(f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token", {
        "grant_type": "refresh_token", "client_id": client_id,
        "refresh_token": refresh_token, "scope": " ".join(scopes)})


def write_private(path, obj):
    os.makedirs(os.path.dirname(path), mode=0o700, exist_ok=True)
    fd = os.open(path + ".tmp", os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w") as f:
        json.dump(obj, f)
    os.replace(path + ".tmp", path)


def check_token(token, tenant):
    c = claims(token)
    if c.get("tid") != tenant or c.get("idtyp") == "app":
        sys.exit("token is not a delegated token for the intended tenant; stopping")
    return c


class Graph:
    """One request at a time, a pause after each, Retry-After honoured on 429. Every call is logged."""

    def __init__(self, token, pause=2.0):
        self.token, self.pause, self.log = token, pause, []

    def call(self, method, url, body=None, headers=None, label=None):
        url = url if url.startswith("http") else BASE + url
        url = urllib.parse.quote(url, safe=":/?&=$,@'()+%")
        h ={"Authorization": f"Bearer {self.token}", **(headers or {})}
        data = None
        if body is not None:
            data, h["Content-Type"] = json.dumps(body).encode(), "application/json"
        delay = 5
        for _ in range(6):
            req = urllib.request.Request(url, data=data, headers=h, method=method)
            try:
                with urllib.request.urlopen(req) as r:
                    raw = r.read().decode()
                    status, rid = r.status, r.headers.get("request-id")
                    etag = r.headers.get("ETag")
            except urllib.error.HTTPError as e:
                raw, status, rid, etag = e.read().decode(errors="replace"), e.code, e.headers.get("request-id"), None
                if status == 429:
                    wait = int(e.headers.get("Retry-After") or delay)
                    print(f"  429, waiting {wait}s", flush=True)
                    time.sleep(wait)
                    delay *= 2
                    continue
            break
        try:
            parsed = json.loads(raw) if raw else None
        except ValueError:
            parsed = raw
        entry = {"label": label, "method": method, "url": url, "request_body": body,
                 "status": status, "request_id": rid, "etag": etag, "body": parsed}
        self.log.append(entry)
        err = parsed.get("error", {}).get("message", "") if isinstance(parsed, dict) and "error" in parsed else ""
        print(f"{status} {method:6} {url.replace(BASE, '')}  {label or ''} {err[:160]}", flush=True)
        time.sleep(self.pause)
        return entry
