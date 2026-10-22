#!/usr/bin/env python3
"""Read-only probe of the Copilot cost-management Graph beta API.

Delegated device-code sign-in -> checks the token is for the intended tenant and a user ->
calls the read endpoints one at a time, following @odata.nextLink and honouring Retry-After.
Writes results/<UTC timestamp>.json. The access token is never printed or written to disk.

    python3 costmgmt_probe.py --tenant <tenant-id> [--client-id <app-id>] [--deep]

--client-id defaults to Microsoft Graph Command Line Tools (what Connect-MgGraph uses).
--deep also reads the signed-in user's balances and every policy's sub-resources.
"""
import argparse, json, os
from datetime import datetime, timezone
from cmapi import GRAPH_CLI, READ_SCOPES, SCOPE, Graph, check_token, device_code

BASIC = [SCOPE("CopilotCostManagement.Read.All"), SCOPE("CopilotCostManagement-Policy.Read.All")]


def service_ids(items):
    return [i["serviceId"] for i in items if i.get("serviceId")]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tenant", required=True)
    ap.add_argument("--client-id", default=GRAPH_CLI)
    ap.add_argument("--deep", action="store_true",
                    help="all four read permissions; also user balances and every policy's sub-resources (GET only)")
    a = ap.parse_args()

    t = device_code(a.tenant, a.client_id, READ_SCOPES if a.deep else BASIC)
    c = check_token(t["access_token"], a.tenant)
    print("signed in:", c.get("upn"), "| scopes:", c.get("scp"), flush=True)
    g = Graph(t["access_token"])

    g.call("GET", "/getTenantCreditBalance()")
    policies, _ = g.list_all("/spendingPolicies")
    if a.deep:
        me = f"/userBalances/{c['oid']}"
        g.call("GET", me)
        for sid in service_ids(g.list_all(f"{me}/serviceBalances")[0]):
            g.call("GET", f"{me}/serviceBalances/{sid}")
        for p in policies:
            pol = f"/spendingPolicies/{p['id']}"
            g.call("GET", pol)
            for sid in service_ids(g.list_all(f"{pol}/serviceSettings")[0]):
                g.call("GET", f"{pol}/serviceSettings/{sid}")
            g.call("GET", f"{pol}/emailThresholdSettings")
            g.call("GET", f"{pol}/emailThresholdSettings/recipients")
            g.call("GET", f"{pol}/emailThresholdSettings/recipients/$ref")
            if "selectedGroups" in p.get("@odata.type", ""):
                g.call("GET", f"{pol}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups")
                g.call("GET", f"{pol}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups/$ref")

    os.makedirs("results", exist_ok=True)
    path = f"results/{datetime.now(timezone.utc):%Y%m%dT%H%M%SZ}.json"
    with open(path, "w") as f:
        json.dump({"run_utc": datetime.now(timezone.utc).isoformat(), "client_id": a.client_id,
                   "token": {k: c.get(k) for k in ("tid", "upn", "idtyp", "scp")}, "calls": g.log}, f, indent=2)
    print("wrote", path)


if __name__ == "__main__":
    main()
