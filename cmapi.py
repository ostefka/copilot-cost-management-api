"""Shared helpers: Entra device-code / refresh-token sign-in and a paced Microsoft Graph client.
Standard library only. Tokens are never printed."""
import base64, email.utils, json, os, sys, tempfile, time, urllib.error, urllib.parse, urllib.request

GRAPH = "https://graph.microsoft.com"
BASE = f"{GRAPH}/beta/copilot/costManagement"
GRAPH_CLI = "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Microsoft Graph Command Line Tools (Connect-MgGraph)
SCOPE = lambda name: f"{GRAPH}/{name}"
READ_SCOPES = [SCOPE(s) for s in (
    "CopilotCostManagement.Read.All", "CopilotCostManagement-Policy.Read.All",
    "CopilotCostManagement-UserData.Read.All", "CopilotCostManagement-Assignment.Read.All")]
TIMEOUT = 30  # seconds per HTTP request


def _post_form(url, data):
    """POST a form to Entra. Network or parse failures come back as an OAuth-style error, never an exception."""
    req = urllib.request.Request(url, urllib.parse.urlencode(data).encode())
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        try:
            return json.load(e)
        except ValueError:
            return {"error": f"http_{e.code}", "error_description": "non-JSON error response"}
    except (urllib.error.URLError, TimeoutError, ValueError) as e:
        return {"error": "network_error", "error_description": str(e)}


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
    """Atomically replace path with obj as JSON, mode 0600, flushed to disk."""
    d = os.path.dirname(path)
    os.makedirs(d, mode=0o700, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=d, prefix=".tmp-")  # unique name, created 0600
    try:
        with os.fdopen(fd, "w") as f:
            json.dump(obj, f)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    except BaseException:
        os.unlink(tmp)
        raise
    dfd = os.open(d, os.O_RDONLY)
    try:
        os.fsync(dfd)
    finally:
        os.close(dfd)


def retry_after(value, default):
    """Seconds to wait from a Retry-After header (delta-seconds or HTTP-date), capped at 5 minutes."""
    try:
        wait = int(value)
    except (TypeError, ValueError):
        try:
            wait = email.utils.parsedate_to_datetime(value).timestamp() - time.time()
        except (TypeError, ValueError):
            wait = default
    return max(1, min(int(wait), 300))


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
        h = {"Authorization": f"Bearer {self.token}", **(headers or {})}
        data = None
        if body is not None:
            data, h["Content-Type"] = json.dumps(body).encode(), "application/json"
        delay = 5
        for _ in range(6):
            req = urllib.request.Request(url, data=data, headers=h, method=method)
            try:
                with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
                    raw = r.read().decode()
                    status, rid = r.status, r.headers.get("request-id")
                    etag = r.headers.get("ETag")
            except urllib.error.HTTPError as e:
                raw, status, rid, etag = e.read().decode(errors="replace"), e.code, e.headers.get("request-id"), None
                if status == 429:
                    wait = retry_after(e.headers.get("Retry-After"), delay)
                    print(f"  429, waiting {wait}s", flush=True)
                    time.sleep(wait)
                    delay *= 2
                    continue
            except (urllib.error.URLError, TimeoutError) as e:
                raw, status, rid, etag = json.dumps({"error": {"message": f"network error: {e}"}}), 0, None, None
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

    def list_all(self, url, label=None):
        """GET a collection and every @odata.nextLink page. Returns (items, last entry); stops at the first failure."""
        items, entry = [], None
        while url:
            entry = self.call("GET", url, label=label)
            if entry["status"] != 200 or not isinstance(entry["body"], dict):
                break
            items += entry["body"].get("value", [])
            url = entry["body"].get("@odata.nextLink")
        return items, entry
