#!/usr/bin/env python3
"""Unattended collector proof: delegated access kept alive with a rotating refresh token.

  python3 collector.py bootstrap   # once: device-code sign-in (offline_access) with the collector app
  python3 collector.py run         # unattended: refresh -> snapshot -> store the rotated refresh token

The refresh token lives in ~/.cache/costmgmt-collector/token.json (0600), never in the repo.
Each run appends one line to collector-data/runs.jsonl: whether the refresh worked, whether the
token rotated, how long since the interactive sign-in, and the cost figures read.
"""
import fcntl, json, os, sys, time
from datetime import datetime, timezone
from cmapi import READ_SCOPES, Graph, check_token, device_code, refresh, write_private

HERE = os.path.dirname(os.path.abspath(__file__))
TOKEN = os.path.expanduser("~/.cache/costmgmt-collector/token.json")
LOCK = TOKEN + ".lock"
APP = os.path.expanduser("~/.cache/costmgmt/app.json")
OUT = os.path.join(HERE, "collector-data", "runs.jsonl")
SCOPES = READ_SCOPES + ["offline_access"]


def now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def record(entry):
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "a") as f:
        f.write(json.dumps(entry) + "\n")
    print(json.dumps(entry))


def bootstrap():
    with open(APP) as f:
        app = json.load(f)
    t = device_code(app["tenant"], app["client_id"], SCOPES)
    if "refresh_token" not in t:
        sys.exit("no refresh token returned; offline_access was not granted")
    c = check_token(t["access_token"], app["tenant"])
    write_private(TOKEN, {"tenant": app["tenant"], "client_id": app["client_id"],
                          "refresh_token": t["refresh_token"], "bootstrapped_at": time.time(), "refreshes": 0})
    print("bootstrapped as", c.get("upn"), "| scopes:", t.get("scope"))
    run()


def run():
    # one run at a time: an overlapping run would redeem the same refresh token and race on the file
    os.makedirs(os.path.dirname(LOCK), mode=0o700, exist_ok=True)
    with open(LOCK, "w") as lock:
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            sys.exit("another collector run is in progress")
        _run()


def _run():
    with open(TOKEN) as f:
        st = json.load(f)
    age_h = round((time.time() - st["bootstrapped_at"]) / 3600, 2)
    t = refresh(st["tenant"], st["client_id"], st["refresh_token"], SCOPES)
    if "access_token" not in t:
        # invalid_grant here means a new interactive bootstrap is needed (revoked, expired, CA policy...)
        record({"ts": now(), "ok": False, "hours_since_signin": age_h, "error": t.get("error"),
                "error_codes": t.get("error_codes"), "description": (t.get("error_description") or "")[:300]})
        sys.exit(1)
    new = t.get("refresh_token")
    rotated = isinstance(new, str) and bool(new) and new != st["refresh_token"]
    if rotated:  # keep the previous token unless Entra returned a usable new one
        st["refresh_token"] = new
    st.update(refreshes=st["refreshes"] + 1, last_ok=time.time())
    write_private(TOKEN, st)  # save the rotated token before anything else can fail
    c = check_token(t["access_token"], st["tenant"])
    g = Graph(t["access_token"])
    bal = g.call("GET", "/getTenantCreditBalance()")
    pols = g.call("GET", "/spendingPolicies")
    svc = g.call("GET", f"/userBalances/{c['oid']}/serviceBalances")
    record({
        "ts": now(), "ok": all(x["status"] == 200 for x in (bal, pols, svc)), "hours_since_signin": age_h,
        "refresh_number": st["refreshes"], "refresh_token_rotated": rotated,
        "access_token_minutes": round((c["exp"] - time.time()) / 60), "scopes": t.get("scope"),
        "status": [bal["status"], pols["status"], svc["status"]],
        "tenant_balance": [{k: b.get(k) for k in ("billingMethod", "consumedQuantity", "remainingQuantity")}
                           | {"asOf": b.get("dataStatus", {}).get("lastModifiedDateTime")}
                           for b in (bal["body"] or {}).get("billingMethodBalances", [])] if bal["status"] == 200 else None,
        "policies": len((pols["body"] or {}).get("value", [])) if pols["status"] == 200 else None,
        "my_consumption": {s["serviceId"]: s["consumedQuantity"] for s in (svc["body"] or {}).get("value", [])
                           if s["consumedQuantity"]} if svc["status"] == 200 else None,
    })


if __name__ == "__main__":
    {"bootstrap": bootstrap, "run": run}[sys.argv[1] if len(sys.argv) > 1 else "run"]()
