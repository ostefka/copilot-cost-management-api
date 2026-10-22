# Copilot cost management API (Microsoft Graph beta): hands-on findings

Microsoft is previewing a Graph API for **Copilot credits**: the tenant's credit balance, per-user consumption by
Copilot service, and the spending policies that cap it. This repository records a full hands-on test of every
endpoint in a Microsoft 365 lab tenant (October 2026), the request rules the preview documentation does not mention,
and a working pattern for collecting the data **unattended** for your own portal or reporting.

> The endpoints are Graph **beta** and in early access: your tenant has to be enabled by Microsoft first, and the
> contract can change. Everything below was observed on 12–22 October 2026.

Base URL for every path: `https://graph.microsoft.com/beta/copilot/costManagement`

## Results at a glance

All 24 endpoints answered successfully with **delegated** access (a signed-in admin). Writes were tested only on
throwaway policies and empty groups created for the test and deleted afterwards; the tenant's real policy was read but
never changed (its ETag was identical before and after).

| Method | Path | Permission (delegated) | Result |
|---|---|---|---|
| GET | `/getTenantCreditBalance()` | `CopilotCostManagement.Read.All` | ✅ 200 |
| GET | `/userBalances/{userId}` | `CopilotCostManagement-UserData.Read.All` | ✅ 200 (also for other users; 404 for an unknown id) |
| GET | `/userBalances/{userId}/serviceBalances` | `…-UserData.Read.All` | ✅ 200 |
| GET | `/userBalances/{userId}/serviceBalances/{serviceId}` | `…-UserData.Read.All` | ✅ 200 |
| GET | `/spendingPolicies` | `…-Policy.Read.All` | ✅ 200 |
| POST | `/spendingPolicies` | `…-Policy.ReadWrite.All` | ✅ 201 (see rules) |
| GET | `/spendingPolicies/{id}` | `…-Policy.Read.All` | ✅ 200 |
| PATCH | `/spendingPolicies/{id}` | `…-Policy.ReadWrite.All` | ✅ 200 (see rules) |
| DELETE | `/spendingPolicies/{id}` | `…-Policy.ReadWrite.All` | ✅ 204 |
| GET | `/spendingPolicies/{id}/serviceSettings` | `…-Policy.Read.All` | ✅ 200 |
| POST | `/spendingPolicies/{id}/serviceSettings` | `…-Policy.ReadWrite.All` | ✅ 201 |
| GET | `/spendingPolicies/{id}/serviceSettings/{serviceId}` | `…-Policy.Read.All` | ✅ 200 |
| PATCH | `/spendingPolicies/{id}/serviceSettings/{serviceId}` | `…-Policy.ReadWrite.All` | ✅ 200 |
| GET | `/spendingPolicies/{id}/emailThresholdSettings` | `…-Policy.Read.All` | ✅ 200 (404 until created) |
| PATCH | `/spendingPolicies/{id}/emailThresholdSettings` | `…-Policy.ReadWrite.All` | ✅ 200 (creates it too) |
| DELETE | `/spendingPolicies/{id}/emailThresholdSettings` | `…-Policy.ReadWrite.All` | ✅ 204 |
| GET | `/spendingPolicies/{id}/emailThresholdSettings/recipients` | `…-Policy.Read.All` | ✅ 200 |
| GET | `/spendingPolicies/{id}/emailThresholdSettings/recipients/$ref` | `…-Policy.Read.All` | ✅ 200 |
| POST | `/spendingPolicies/{id}/emailThresholdSettings/recipients/$ref` | `…-Policy.ReadWrite.All` | ✅ 204 |
| DELETE | `/spendingPolicies/{id}/emailThresholdSettings/recipients/{userId}/$ref` | `…-Policy.ReadWrite.All` | ✅ 204 |
| GET | `/spendingPolicies/{id}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups` | `…-Assignment.Read.All` | ✅ 200 |
| GET | `/spendingPolicies/{id}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups/$ref` | `…-Assignment.Read.All` | ✅ 200 |
| POST | `/spendingPolicies/{id}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups/$ref` | `…-Assignment.ReadWrite.All` | ✅ 204 |
| DELETE | `/spendingPolicies/{id}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups/{groupId}/$ref` | `…-Assignment.ReadWrite.All` | ✅ 204 |

Also observed:

- **App-only tokens are refused** (403 on `getTenantCreditBalance()` and `spendingPolicies` with a client-credentials
  token). Only delegated access works today — see [Running it unattended](#running-it-unattended).
- **There is no tenant-wide per-user list:** `GET /userBalances` returns 404. Per-user data is one call per user.
- **`serviceBalances` is paged at 10 items.** The lab user had 11 services, so the 11th came only by following
  `@odata.nextLink`. Always follow it, unchanged.
- `$select` works on `/spendingPolicies`; `$top=1` was accepted (the preview guide lists `$top` as a known issue, and
  the lab had only one policy, so paging could not be verified).
- No 429s in roughly 300 requests sent one at a time, two seconds apart.
- The standard **Microsoft Graph Command Line Tools** app (the one `Connect-MgGraph` uses) can request these
  permissions by dynamic consent, so a first test needs no app registration.

## Rules the documentation does not tell you

Found by testing; the error messages quoted are the service's own.

**Every write needs `If-Match`** with the policy's current `@odata.etag` — including writes to sub-resources
(service settings, alert settings, recipients, groups), where the ETag of the *parent policy* is expected. Without it:
`400 "The If-Match header is required"`, or `428`/`412` on some routes.

**Creating a policy** (`POST /spendingPolicies`):

```json
{
  "@odata.type": "#microsoft.graph.selectedGroupsSpendingPolicy",
  "displayName": "Pilot group policy",
  "billingMethods": "prepaidCredits",
  "periodLimit":  { "@odata.type": "#microsoft.graph.limitedCostManagementQuantity", "value": 500, "unit": "copilotCredit" },
  "perUserLimit": { "@odata.type": "#microsoft.graph.limitedCostManagementQuantity", "value": 20,  "unit": "copilotCredit" },
  "serviceSettings": [ { "serviceId": "cowork", "state": "enabled" } ]
}
```

- `serviceSettings` is **required** and must contain **at least one enabled** service.
- `displayName` must be **10–64 characters, letters, digits and spaces only** (no hyphens).
- `status` is **rejected** on create, and groups cannot be given inline: `assignedGroups@odata.bind` is rejected, and
  an inline `assignedGroups` array timed out with **504 and left a broken policy behind** (an all-users policy named
  `Failed create cleanup: <id>` in status `provisioningFailed`, which had to be deleted by hand). Assign groups
  afterwards with `POST …/assignedGroups/$ref`.
- A selected-groups policy is created **disabled**, and becomes **enabled automatically when the first group is
  assigned**.

**Updating a policy** (`PATCH`): send plain properties only (`displayName`, `periodLimit`, `perUserLimit`). Including
`@odata.type`, or trying to set `status`, fails with `400 "The request body does not match the spending policy update
contract"`. Switching `billingMethods` to `prepaidCredits,payAsYouGo` returned `409` ("conflicts with its current state
or funding configuration") in a tenant without pay-as-you-go set up.

**Groups:** the **last group cannot be removed** from a selected-groups policy (`409`); delete the policy instead.
Assign with `{"@odata.id": "https://graph.microsoft.com/beta/groups/{id}"}` (a `directoryObjects/{id}` URL is rejected).

**Email alerts** must be built in this order:

1. `PATCH …/emailThresholdSettings` with `isEnabled: false` and the thresholds — this creates the settings.
2. `POST …/emailThresholdSettings/recipients/$ref` with `{"@odata.id": "https://graph.microsoft.com/beta/users/{id}"}`.
3. `PATCH …/emailThresholdSettings` with `isEnabled: true`.

Enabling with no recipient fails (`"At least one alert recipient must be linked before alerting can be enabled"`), and
the last recipient cannot be removed while alerts are on — disable first. Thresholds are
`#microsoft.graph.percentageCostManagementAlertThreshold` (`value` in %) or
`#microsoft.graph.absoluteCostManagementAlertThreshold` (`value` + `unit`). **Percentage thresholds need a finite
`periodLimit`**: accepted while the limit was 20 credits, rejected (`"The supplied alert thresholds are invalid"`) once
it was set to unlimited. Absolute thresholds were accepted.

## What data you get

**Tenant balance** — one entry per billing method; `remainingQuantity` and `purchasedQuantity` are absent for
pay-as-you-go. Lab values: 57,000 prepaid credits purchased, 2,060.57 consumed, 54,939.43 remaining.

**Per-user, per-service balance** (`serviceBalances`) — `consumedQuantity`, `isConsumptionAllowed`,
`accessDecisionReason` (`allowed`, `notEntitled`, `limitExhausted`), the governing `spendingPolicyId`, configured and
effective limits, `remainingQuantity` and `dataStatus`. `/userBalances/{id}` adds `selectionDataStatus`
(`evaluatedDateTime`, `usageAsOfDateTime`) and `unevaluatedServices`. The lab user had 11 services (two pages), among them `cowork`, `workIQ`,
`copilotCode`, `m365CopilotChatAdvancedWork`, `researcherCritiqueAndCouncil`, `outlook`, `opal`, `scout`,
`teamsInterpreter` and `teamsPhoneAgent`.

**Policies** — limits, billing method, status (`enabled`, `disabled`, `provisioning`, `provisioningFailed`) and per-service
settings; the lab's all-users policy listed 32 services, including Copilot Studio, Microsoft Fabric, GitHub Copilot,
SharePoint and OneDrive Copilot, and Windows 365 for Agents.

**Store the freshness timestamps.** In the lab, per-user usage was reported as of 2 October while the tenant balance
was updated on 12 October, and the two "remaining" figures differed (54,595.52 vs 54,939.43).

The full schema (types, enums, properties) is in Graph's public `$metadata`: `https://graph.microsoft.com/beta/$metadata`
— search for `costManagement`.

## Running it unattended

A normal daemon signs in as itself (client credentials) and gets an app-only token. **This API refuses app-only
tokens today** (403) and checks the signed-in user's Entra role, so a daemon has to act *as a user*. That works with
the standard OAuth `offline_access` refresh token, and we proved it end to end:

1. **A dedicated app registration** (`validate_all.py setup-app`): single tenant, public client
   (`isFallbackPublicClient: true`, redirect `http://localhost`), **assignment required**, and only the signing-in
   account assigned. No API permissions are pre-declared; the hidden early-access scopes are granted by consent at
   sign-in.
2. **One interactive sign-in** (`collector.py bootstrap`): device code, requesting the four read scopes plus
   `offline_access`. The admin consents once; Entra returns a refresh token, which is stored outside the repository
   with mode 0600.
3. **Unattended runs** (`collector.py run`): each run redeems the refresh token for a fresh access token, saves the
   **rotated** refresh token immediately, then reads the tenant balance, the policies and per-user service balances.

What we observed:

| Check | Result |
|---|---|
| Refresh token returned for the hidden early-access scopes | ✅ yes |
| Redeemed from a separate process, no user present | ✅ 200 on all three reads |
| Refresh token rotated on every redemption | ✅ yes (new token each time) |
| Access-token lifetime | 60–90 minutes |
| Scopes on the unattended token | the four `CopilotCostManagement…Read.All` scopes plus the standard `openid profile email` — nothing else, even though the same admin had consented broader scopes (chat, meetings, write) to the Microsoft Graph Command Line Tools app — refresh tokens are per app, so a **dedicated app keeps the stored credential narrow** |
| Data freshness seen by the collector | tenant balance `lastModifiedDateTime` about 20 hours behind the time of the call |

Not proven yet: how long the chain survives in practice. By Entra defaults a refresh token stays valid for 90 days
while it is used, and stops working on password reset, session revocation, account disable, or a Conditional Access
sign-in-frequency policy. `deploy/` contains an hourly systemd timer for a long-running proof.

**Production pattern** — the same three steps, with:

- a **dedicated service account** (cloud-only, the **AI Reader** directory role only, MFA at the one-time sign-in)
  instead of an administrator;
- the refresh token in **Azure Key Vault**, read and rewritten by the collector's **managed identity** (Azure
  Container Apps job or Functions timer);
- results stored as **snapshots** (raw JSON plus a flat table keyed by time, user and service), because the API
  returns current cumulative values, not history; the portal reads the store, never Graph;
- an **alert** when a refresh fails with `invalid_grant`, and a short runbook to redo the one-time sign-in.

Do not use the username/password (ROPC) flow, exclude the account from MFA, or store its password. If Microsoft adds
application permissions to this API later, the collector can switch to its managed identity and drop the service
account.

## Scripts

Python 3, standard library only. Tokens are never printed; the refresh token is kept outside the repository with
mode 0600.

| File | What it does |
|---|---|
| `cmapi.py` | Device-code and refresh-token sign-in, and a paced Graph client (one request at a time, `Retry-After` honoured). |
| `costmgmt_probe.py` | Quick read-only check: `--deep` reads every GET endpoint for the signed-in user and every policy. |
| `validate_all.py` | The full validation: every GET, then the complete write lifecycle on throwaway objects. Clean-up always runs, by tracked id and by name. Also creates the collector's app registration (`setup-app`). |
| `collector.py` | The unattended collector proof: `bootstrap` once, then `run` from a timer. |
| `deploy/` | systemd user timer that runs the collector hourly. |

```bash
python3 costmgmt_probe.py --tenant <tenant-id> --deep          # read-only, admin or AI Reader
python3 validate_all.py login --tenant <tenant-id>             # lab tenants only: it creates and deletes objects
python3 validate_all.py setup-app && python3 validate_all.py validate
python3 collector.py bootstrap                                 # once, interactive
python3 collector.py run                                       # unattended
```

`docs/` has a customer-facing write-up of the same findings (Word).
