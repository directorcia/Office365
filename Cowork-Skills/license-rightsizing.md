---
name: license-rightsizing
description: |
  Audits a client's Microsoft 365 license assignments to find unused, oversized,
  or duplicated licenses, and produces a rightsizing recommendation with dollar
  savings. Use when user says "license review for [client]", "are we over-
  licensed", "rightsize M365", "find unused licenses", or "save money on M365".
---

# License Rightsizing

License waste is the easiest win in any MSP's commercial conversation. Most clients are paying for 10–25% more than they use. This skill finds it, quantifies it, and frames the recommendation.

## When to use

- "Run a license review for [client]"
- "Are we over-licensed on M365"
- "Find unused licenses"
- "Rightsize before renewal"
- "How much could [client] save on M365"

## Data to pull

For the target tenant:
1. **License inventory** — all SKUs, quantity purchased, quantity assigned
2. **User list** — including last sign-in, account state (enabled/disabled), user type (member/guest)
3. **Mailbox state** — active, shared, inactive, on litigation hold
4. **Service usage** — Teams, Exchange, OneDrive, SharePoint last activity dates per user
5. **Renewal dates** and current pricing per SKU

## Analysis

For each user with a paid license, ask:

### 1. Are they still here?
- Disabled account → license should be removed
- Last sign-in > 30 days and account enabled → confirm with client before removing
- Account never signed in (created > 14 days ago) → confirm whether onboarding stalled

### 2. Are they using what they're paying for?
- E5 user with no Teams Phone, no Power BI Pro usage, no Defender for O365 P2 features used → candidate to drop to E3
- E3 user with no Office desktop install in 90 days → candidate for Business Basic / F3
- Frontline / shift worker on E3 → candidate for F1 or F3

### 3. Should this even be a user license?
- Shared mailbox under 50 GB → no license needed
- Mailbox for a former staff member kept for retention → InPlace / Litigation Hold + remove license
- Resource mailbox (room, equipment) → no license needed

### 4. Are there duplicate or overlapping SKUs?
- Standalone Exchange Online Plan + E3 (E3 already includes Exchange) — drop the standalone
- Defender add-ons stacked on top of an SKU that includes them
- Visio / Project assigned to people who haven't opened the app in 90 days

## Output

### Executive summary
- Total licenses purchased vs assigned vs actively used
- Estimated monthly waste in $
- Estimated annualized savings if recommendations adopted
- One-line headline: *"42 licenses out of 187 are reclaimable, ~$X/month"*

### Detailed table
| User / Mailbox | Current SKU | Last sign-in | Recommendation | Monthly $ impact |
|----------------|-------------|--------------|----------------|------------------|
| j.smith@... | E5 | 92 days ago | Disable + remove license | -$57 |
| accounts@... | E3 (shared) | n/a (shared) | Convert to shared mailbox | -$22 |
| ... | ... | ... | ... | ... |

### Recommendations grouped by action
1. **Remove immediately** — disabled or departed users
2. **Confirm with client, then remove** — long-inactive users
3. **Downgrade SKU** — overlicensed users
4. **Consolidate** — duplicate / stacked SKUs
5. **Restructure at renewal** — quantity reductions on the next anniversary

## Workflow

1. **Pull data** from the tenant. Compute every total and percentage with a code tool — never eyeball.
2. **Apply the analysis rules** above to each user.
3. **Group recommendations** by risk level (immediate / confirm / next renewal).
4. **Cost the impact** at the client's actual contracted rate, not list price.
5. **Produce the deliverable** — for the MSP team, an Excel workbook (`xlsx` skill); for the client, a 1-page Word summary (`docx` skill) plus the workbook as an appendix.

## Guardrails

- **Never auto-remove a license.** Every removal is a client-approved action.
- **Sign-in data lag** — Azure AD sign-in logs default to 30 days. Don't claim "never used" without checking the longer audit window.
- **Some inactive accounts are intentional** — service accounts, break-glass, seasonal staff, executives on leave. Always confirm before recommending removal.
- **Compliance holds override cost optimization** — a mailbox under legal hold stays as it is, even if it looks idle.
- **Renewal terms matter** — you may not be able to reduce quantity mid-term on an annual commitment. Flag the next true-down window.
- **Don't quote vendor list prices** as savings. Use the client's actual price.

## Done when

The client has a numbered list of decisions to make, each with a dollar value and a recommended action date.
