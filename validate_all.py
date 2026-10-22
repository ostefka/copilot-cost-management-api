#!/usr/bin/env python3
"""Validate every Copilot cost-management endpoint (Graph beta) in a lab tenant.

  python3 validate_all.py login --tenant <id>   # admin device-code sign-in; caches the ACCESS token only (~1 h)
  python3 validate_all.py setup-app             # app registration for the unattended collector (idempotent)
  python3 validate_all.py validate              # every GET, then the full write lifecycle on a throwaway policy
  python3 validate_all.py cleanup               # remove anything a crashed run left behind

Writes touch ONLY objects this script created (a test group and test policies named "zz api test…").
The tenant's existing policies are read, never changed. Results: results/validate-<UTC>.json.
"""
import argparse, json, os, sys, time
from datetime import datetime, timezone
from cmapi import BASE, GRAPH, GRAPH_CLI, READ_SCOPES, SCOPE, Graph, check_token, device_code, write_private

STATE = os.path.expanduser("~/.cache/costmgmt")
SESSION, CREATED, APP = (f"{STATE}/{n}.json" for n in ("session", "created", "app"))
PREFIX = "zz api test"  # names: 10-64 letters, digits and spaces only
WRITE_SCOPES = READ_SCOPES + [SCOPE(s) for s in (
    "CopilotCostManagement-Policy.ReadWrite.All", "CopilotCostManagement-Assignment.ReadWrite.All",
    "Group.ReadWrite.All", "User.ReadBasic.All", "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All")]


def load(path, default=None):
    try:
        with open(path) as f:
            return json.load(f)
    except FileNotFoundError:
        return default


def session():
    s = load(SESSION)
    if not s or s["expires_at"] < time.time() + 120:
        sys.exit("no valid session; run: validate_all.py login --tenant <id>")
    c = check_token(s["access_token"], s["tenant"])
    return Graph(s["access_token"]), c, s["tenant"]


def remember(kind, oid):
    created = load(CREATED, {"groups": [], "policies": []})
    created[kind].append(oid)
    write_private(CREATED, created)


def forget(kind, oid):
    created = load(CREATED, {"groups": [], "policies": []})
    created[kind] = [x for x in created[kind] if x != oid]
    write_private(CREATED, created)


def cmd_login(a):
    t = device_code(a.tenant, GRAPH_CLI, WRITE_SCOPES)
    c = check_token(t["access_token"], a.tenant)
    write_private(SESSION, {"tenant": a.tenant, "access_token": t["access_token"],
                            "expires_at": time.time() + int(t["expires_in"])})
    print("signed in:", c.get("upn"), "| scopes:", c.get("scp"))


def cmd_setup_app(a):
    g, c, tenant = session()
    app, stored = None, load(APP)
    if stored and stored.get("tenant") == tenant:  # reuse only the app this script created, never a name match
        r = g.call("GET", f"{GRAPH}/v1.0/applications(appId='{stored['client_id']}')?$select=id,appId", label="stored app")
        if r["status"] == 200:
            app = r["body"]
        elif r["status"] != 404:  # a timeout or throttle is not "missing": never create a second registration
            sys.exit(f"could not look up the stored app (HTTP {r['status']}); app.json left unchanged")
    if not app:
        r = g.call("POST", f"{GRAPH}/v1.0/applications", {
            "displayName": "Copilot cost management collector (test)", "signInAudience": "AzureADMyOrg",
            "isFallbackPublicClient": True, "publicClient": {"redirectUris": ["http://localhost"]}}, label="create app")
        if r["status"] != 201:
            sys.exit("could not create the app registration")
        app = r["body"]
    sps = g.call("GET", f"{GRAPH}/v1.0/servicePrincipals?$filter=appId eq '{app['appId']}'&$select=id",
                 label="find SP")["body"].get("value", [])
    sp = sps[0] if sps else g.call("POST", f"{GRAPH}/v1.0/servicePrincipals",
                                   {"appId": app["appId"], "appRoleAssignmentRequired": True}, label="create SP")["body"]
    assigned = g.call("GET", f"{GRAPH}/v1.0/servicePrincipals/{sp['id']}/appRoleAssignedTo", label="assignments")
    if not any(x["principalId"] == c["oid"] for x in assigned["body"].get("value", [])):
        g.call("POST", f"{GRAPH}/v1.0/servicePrincipals/{sp['id']}/appRoleAssignedTo", {
            "principalId": c["oid"], "resourceId": sp["id"],
            "appRoleId": "00000000-0000-0000-0000-000000000000"}, label="assign signed-in user")
    write_private(APP, {"tenant": tenant, "client_id": app["appId"], "sp_id": sp["id"]})
    print("collector app client id:", app["appId"])


def first_service(entry):
    v = (entry["body"] or {}).get("value", []) if entry["status"] == 200 else []
    return v[0]["serviceId"] if v else None


def pw(g, method, url, body=None, label=None):
    """A write on a test policy: try as-is; if the service demands If-Match, retry with the policy's current ETag."""
    r = g.call(method, url, body, label=label)
    if r["status"] in (412, 428) or (r["status"] == 400 and "If-Match" in json.dumps(r["body"])):
        policy = "/".join(url.split("/")[:3])  # /spendingPolicies/{id}
        et = g.call("GET", policy, label="read ETag")["body"].get("@odata.etag")
        r = g.call(method, url, body, headers={"If-Match": et}, label=f"{label or ''} + If-Match")
    return r


def cmd_validate(a):
    g, c, tenant = session()
    me = c["oid"]
    print("== READ: tenant and existing policies (never modified)")
    g.call("GET", "", label="costManagement root")
    g.call("GET", "/getTenantCreditBalance()")
    policies, _ = g.list_all("/spendingPolicies")
    g.call("GET", "/spendingPolicies?$top=1", label="known issue: $top")
    g.call("GET", "/spendingPolicies?$select=id,displayName,status", label="$select")
    existing = [p for p in policies if not p["displayName"].startswith(PREFIX)]
    for p in existing:
        pid = f"/spendingPolicies/{p['id']}"
        g.call("GET", pid)
        sid = first_service(g.call("GET", f"{pid}/serviceSettings"))
        if sid:
            g.call("GET", f"{pid}/serviceSettings/{sid}")
        g.call("GET", f"{pid}/emailThresholdSettings")
        g.call("GET", f"{pid}/emailThresholdSettings/recipients")
        g.call("GET", f"{pid}/emailThresholdSettings/recipients/$ref")
        if "selectedGroups" in p.get("@odata.type", ""):
            g.call("GET", f"{pid}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups")

    print("== READ: user balances")
    g.call("GET", "/userBalances", label="list all users' balances (undocumented)")
    g.call("GET", f"/userBalances/{me}", label="signed-in user")
    sid = first_service(g.call("GET", f"/userBalances/{me}/serviceBalances"))
    if sid:
        g.call("GET", f"/userBalances/{me}/serviceBalances/{sid}")
    users = g.call("GET", f"{GRAPH}/v1.0/users?$top=10&$select=id", label="pick another user")["body"].get("value", [])
    other = next((u["id"] for u in users if u["id"] != me), None)
    if other:
        g.call("GET", f"/userBalances/{other}", label="another user")
        g.call("GET", f"/userBalances/{other}/serviceBalances", label="another user")
    g.call("GET", "/userBalances/00000000-0000-0000-0000-000000000001", label="non-existent user")

    started = time.time()
    try:
        writes(g, me)
    finally:
        print("== CLEAN-UP")
        cleanup(g, since=started)
        left, _ = g.list_all("/spendingPolicies", label="final policy list")
        print("policies now:", [(p["displayName"], p["status"], p.get("@odata.etag")) for p in left])
        save(g, c)


def writes(g, me):
    print("== WRITE: lifecycle on throwaway objects")
    gids = []
    for n in (1, 2):  # two empty groups: the last group of a selected-groups policy cannot be removed
        grp = g.call("POST", f"{GRAPH}/v1.0/groups", {
            "displayName": f"{PREFIX} group {n}", "mailEnabled": False, "mailNickname": f"zzapitest{int(time.time())}{n}",
            "securityEnabled": True, "description": "Empty group for Copilot cost API tests. Safe to delete."},
            label="create empty test group")
        if grp["status"] == 201:
            gids.append(grp["body"]["id"])
            remember("groups", grp["body"]["id"])
    time.sleep(10)  # let the new groups replicate before referencing them

    limited = lambda v: {"@odata.type": "#microsoft.graph.limitedCostManagementQuantity", "value": v, "unit": "copilotCredit"}
    # Create rules (found by testing): no 'status', no inline groups (assignedGroups@odata.bind is rejected and an inline
    # assignedGroups array timed out with 504 and left a 'Failed create cleanup' policy behind), serviceSettings required
    # with at least one enabled service, displayName 10-64 letters, digits and spaces. With no groups it is created disabled.
    body = {"@odata.type": "#microsoft.graph.selectedGroupsSpendingPolicy", "displayName": f"{PREFIX} policy",
            "billingMethods": "prepaidCredits", "periodLimit": limited(10), "perUserLimit": limited(1),
            "serviceSettings": [{"serviceId": "cowork", "state": "enabled"}]}
    pol = g.call("POST", "/spendingPolicies", body, label="create policy (no groups)")
    if pol["status"] not in (200, 201):
        print("create failed; nothing else to test on writes")
        return
    pid = pol["body"]["id"]
    remember("policies", pid)
    P = f"/spendingPolicies/{pid}"
    # every write below needs If-Match with the policy's ETag; pw() shows the refusal, then retries with it
    pw(g, "PATCH", P, {"displayName": f"{PREFIX} policy patched"}, label="update: rename")
    pw(g, "PATCH", P, {"perUserLimit": limited(2), "periodLimit": limited(20)}, label="update: limits")
    pw(g, "PATCH", P, {"@odata.type": "#microsoft.graph.selectedGroupsSpendingPolicy", "displayName": f"{PREFIX} policy x"},
       label="update with @odata.type (rejected)")
    pw(g, "PATCH", P, {"status": "disabled"}, label="update status (rejected)")
    g.call("GET", P, label="verify update")

    svc = "workIQ"
    pw(g, "POST", f"{P}/serviceSettings", {"serviceId": svc, "state": "disabled"}, label="add service setting")
    g.call("GET", f"{P}/serviceSettings")
    g.call("GET", f"{P}/serviceSettings/{svc}")
    pw(g, "PATCH", f"{P}/serviceSettings/{svc}", {"state": "enabled"}, label="update service setting")
    g.call("GET", f"{P}/serviceSettings/{svc}", label="verify")

    E = f"{P}/emailThresholdSettings"
    g.call("GET", E, label="none yet (404)")
    pct = lambda v: {"@odata.type": "#microsoft.graph.percentageCostManagementAlertThreshold", "value": v}
    pw(g, "PATCH", E, {"isEnabled": False, "thresholds": [pct(80)]}, label="create alert settings, disabled, 80%")
    pw(g, "PATCH", E, {"thresholds": [pct(50), pct(90)]}, label="two percentage thresholds")
    pw(g, "PATCH", E, {"thresholds": [{"@odata.type": "#microsoft.graph.absoluteCostManagementAlertThreshold",
                                       "value": 5, "unit": "copilotCredit"}]}, label="absolute threshold (5 credits)")
    pw(g, "PATCH", E, {"isEnabled": True}, label="enable before a recipient exists (rejected)")
    pw(g, "POST", f"{E}/recipients/$ref", {"@odata.id": f"{GRAPH}/beta/users/{me}"}, label="add recipient (me)")
    pw(g, "PATCH", E, {"isEnabled": True}, label="enable alerts")
    g.call("GET", E, label="verify")
    g.call("GET", f"{E}/recipients")
    g.call("GET", f"{E}/recipients/$ref")
    pw(g, "DELETE", f"{E}/recipients/{me}/$ref", label="remove last recipient while enabled (rejected)")
    pw(g, "PATCH", E, {"isEnabled": False}, label="disable alerts")
    pw(g, "DELETE", f"{E}/recipients/{me}/$ref", label="remove recipient")
    pw(g, "DELETE", E, label="remove alert settings")
    g.call("GET", E, label="after delete (404)")
    pw(g, "PATCH", P, {"periodLimit": {"@odata.type": "#microsoft.graph.unlimitedCostManagementQuantity",
                                       "unit": "copilotCredit"}}, label="update: unlimited period limit")
    pw(g, "PATCH", E, {"isEnabled": False, "thresholds": [pct(80)]}, label="percentage threshold, unlimited limit (rejected)")
    pw(g, "DELETE", E, label="remove alert settings if they were created")

    A = f"{P}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups"
    for gid in gids:
        pw(g, "POST", f"{A}/$ref", {"@odata.id": f"{GRAPH}/beta/groups/{gid}"}, label="assign empty test group")
    g.call("GET", P, label="status after assignment (becomes enabled)")
    g.call("GET", A)
    g.call("GET", f"{A}/$ref")
    if len(gids) == 2:
        pw(g, "DELETE", f"{A}/{gids[1]}/$ref", label="unassign second group")
        pw(g, "DELETE", f"{A}/{gids[0]}/$ref", label="unassign last group (rejected)")
    g.call("GET", A, label="after unassign")

    if pw(g, "DELETE", P, label="delete policy")["status"] == 204:
        forget("policies", pid)
    g.call("GET", P, label="after delete (404)")
    for gid in gids:
        if g.call("DELETE", f"{GRAPH}/v1.0/groups/{gid}", label="delete test group")["status"] == 204:
            forget("groups", gid)


def cleanup(g, since=None):
    """Delete what the tests created: tracked ids, then anything named with PREFIX. A 'Failed create cleanup: ...'
    policy (left by a failed create) is deleted only if it appeared during this run; otherwise it is reported."""
    created = load(CREATED, {"groups": [], "policies": []})
    for pid in list(created["policies"]):
        if pw(g, "DELETE", f"/spendingPolicies/{pid}", label="clean-up")["status"] in (204, 404):
            forget("policies", pid)
    for p in g.list_all("/spendingPolicies", label="clean-up scan")[0]:
        if p["displayName"].startswith(PREFIX):
            pw(g, "DELETE", f"/spendingPolicies/{p['id']}", label="clean-up by name")
        elif p["displayName"].startswith("Failed create cleanup:"):
            made = datetime.fromisoformat(p["createdDateTime"].replace("Z", "+00:00")).timestamp()
            if since and made >= since - 60:
                pw(g, "DELETE", f"/spendingPolicies/{p['id']}", label="clean-up failed create from this run")
            else:
                print("NOTE: leftover from an earlier failed create, not deleted:", p["id"], p["status"])
    for gid in list(created["groups"]):
        if g.call("DELETE", f"{GRAPH}/v1.0/groups/{gid}", label="clean-up")["status"] in (204, 404):
            forget("groups", gid)
    for grp in g.list_all(f"{GRAPH}/v1.0/groups?$filter=startswith(displayName,'{PREFIX} group')&$select=id",
                          label="clean-up group scan")[0]:
        g.call("DELETE", f"{GRAPH}/v1.0/groups/{grp['id']}", label="clean-up group by name")


def cmd_cleanup(a):
    g, c, _ = session()
    cleanup(g)


def save(g, c):
    os.makedirs("results", exist_ok=True)
    path = f"results/validate-{datetime.now(timezone.utc):%Y%m%dT%H%M%SZ}.json"
    with open(path, "w") as f:
        json.dump({"signed_in": {k: c.get(k) for k in ("upn", "scp", "wids")}, "calls": g.log}, f, indent=2)
    print("wrote", path)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("login").add_argument("--tenant", required=True)
    for n in ("setup-app", "validate", "cleanup"):
        sub.add_parser(n)
    a = ap.parse_args()
    globals()["cmd_" + a.cmd.replace("-", "_")](a)
