#!/usr/bin/env python3
"""Read-only probe of the Copilot cost-management Graph beta API (early-access guide, 18 Sep 2026).

Delegated device-code sign-in -> checks the token is for the intended tenant and a user ->
calls GET /getTenantCreditBalance() and GET /spendingPolicies one at a time, following
@odata.nextLink unchanged and honouring Retry-After. Writes results/<UTC timestamp>.json.
The access token is never printed or written to disk.

    python3 costmgmt_probe.py --tenant <tenant-id> [--client-id <app-id>]

--client-id defaults to Microsoft Graph Command Line Tools (what Connect-MgGraph uses);
pass the "Copilot cost management test" app registration from the guide to use that instead.
"""
import argparse, base64, json, os, sys, time, urllib.error, urllib.parse, urllib.request
from datetime import datetime, timezone

GRAPH_CLI = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
BASE = "https://graph.microsoft.com/beta/copilot/costManagement"
SCOPES = {
    "balance": "https://graph.microsoft.com/CopilotCostManagement.Read.All",
    "policies": "https://graph.microsoft.com/CopilotCostManagement-Policy.Read.All",
    "userdata": "https://graph.microsoft.com/CopilotCostManagement-UserData.Read.All",
    "assignment": "https://graph.microsoft.com/CopilotCostManagement-Assignment.Read.All",
}
BASIC = ["balance", "policies"]


def service_ids(body):
    """Ids of a service-keyed collection; the field name is not documented yet."""
    items = body.get("value", []) if isinstance(body, dict) else []
    return [i.get("serviceId") or i.get("id") for i in items if i.get("serviceId") or i.get("id")]


def post_form(url, data):
    req = urllib.request.Request(url, urllib.parse.urlencode(data).encode())
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        return json.load(e)


def claims(token):
    p = token.split(".")[1]
    return json.loads(base64.urlsafe_b64decode(p + "=" * (-len(p) % 4)))


def device_code_token(tenant, client_id, scopes):
    auth = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0"
    dc = post_form(f"{auth}/devicecode", {"client_id": client_id, "scope": " ".join(scopes)})
    if "device_code" not in dc:
        sys.exit(f"device code request failed: {dc.get('error')}: {dc.get('error_description')}")
    print(dc["message"], flush=True)
    interval, deadline = dc.get("interval", 5), time.time() + dc.get("expires_in", 900)
    while time.time() < deadline:
        time.sleep(interval)
        t = post_form(f"{auth}/token", {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": client_id, "device_code": dc["device_code"]})
        if "access_token" in t:
            return t["access_token"]
        if t.get("error") == "slow_down":
            interval += 5
        elif t.get("error") != "authorization_pending":
            sys.exit(f"sign-in failed: {t.get('error')}: {t.get('error_description')}")
    sys.exit("sign-in timed out")


def get(url, token, tries=5):
    delay = 5
    for _ in range(tries):
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                return {"status": r.status, "request_id": r.headers.get("request-id"), "body": json.load(r)}
        except urllib.error.HTTPError as e:
            body = e.read().decode(errors="replace")
            try:
                body = json.loads(body)
            except ValueError:
                pass
            if e.code == 429:
                wait = int(e.headers.get("Retry-After") or delay)
                print(f"429 on {url}, waiting {wait}s", flush=True)
                time.sleep(wait)
                delay *= 2
                continue
            return {"status": e.code, "request_id": e.headers.get("request-id"), "body": body}
    return {"status": 429, "body": "gave up after repeated 429"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tenant", required=True)
    ap.add_argument("--client-id", default=GRAPH_CLI)
    ap.add_argument("--only", choices=BASIC, help="request just one permission")
    ap.add_argument("--deep", action="store_true",
                    help="all four read permissions; also read user balances and every policy's sub-resources (GET only)")
    a = ap.parse_args()

    wanted = list(SCOPES) if a.deep else [a.only] if a.only else BASIC
    token = device_code_token(a.tenant, a.client_id, [SCOPES[k] for k in wanted])
    c = claims(token)
    who = {k: c.get(k) for k in ("tid", "upn", "idtyp", "scp", "wids", "app_displayname")}
    print("signed in:", json.dumps(who), flush=True)
    if c.get("tid") != a.tenant or c.get("idtyp") == "app":
        sys.exit("token is not a delegated token for the intended tenant; stopping")

    out = {"run_utc": datetime.now(timezone.utc).isoformat(), "client_id": a.client_id, "token": who, "calls": []}
    urls = []
    if "balance" in wanted:
        urls.append(f"{BASE}/getTenantCreditBalance()")
    if "policies" in wanted:
        urls.append(f"{BASE}/spendingPolicies")
    def call(url):
        """GET url and every @odata.nextLink page; returns the merged value list (or the single body)."""
        first, items = None, []
        while url:
            r = get(url, token)
            print(r["status"], url, r.get("request_id"), flush=True)
            out["calls"].append({"url": url, **r})
            time.sleep(2)  # guide: strict rate limits, one request at a time
            if r["status"] != 200 or not isinstance(r["body"], dict):
                return first
            first = first or r["body"]
            items += r["body"].get("value", [])
            url = r["body"].get("@odata.nextLink")
        return {**first, "value": items} if "value" in first else first

    bodies = {u: call(u) for u in urls}
    if a.deep:
        me = f"{BASE}/userBalances/{c['oid']}"
        call(me)
        for sid in service_ids(call(f"{me}/serviceBalances")):
            call(f"{me}/serviceBalances/{sid}")
        for p in (bodies.get(f"{BASE}/spendingPolicies") or {}).get("value", []):
            pol = f"{BASE}/spendingPolicies/{p['id']}"
            call(pol)
            for sid in service_ids(call(f"{pol}/serviceSettings")):
                call(f"{pol}/serviceSettings/{sid}")
            call(f"{pol}/emailThresholdSettings")
            # the PDF and the scopes sheet disagree on the trailing /$ref; try both
            call(f"{pol}/emailThresholdSettings/recipients")
            call(f"{pol}/emailThresholdSettings/recipients/$ref")
            call(f"{pol}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups")
            call(f"{pol}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups/$ref")

    os.makedirs("results", exist_ok=True)
    path = f"results/{datetime.now(timezone.utc):%Y%m%dT%H%M%SZ}.json"
    with open(path, "w") as f:
        json.dump(out, f, indent=2)
    print("wrote", path, flush=True)


if __name__ == "__main__":
    main()
