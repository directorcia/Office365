---
name: client-onboarding
description: |
  Runs a standardized new-client onboarding workflow for an MSP — gathers tenant
  information, builds a kickoff pack, drafts welcome communications, and creates
  the initial documentation set. Use when a new client signs an MSA or when the
  user says "onboard new client", "kickoff [client name]", "set up new tenant",
  or "welcome pack for [company]".
---

# Client Onboarding

Standardize the first 30 days of every new MSP engagement so nothing slips between sales and service delivery.

## When to use

Trigger when the user says:
- "onboard [client name]"
- "new client kickoff"
- "set up [company] in our system"
- "create a welcome pack for [client]"
- "start onboarding for the new MSA"

## Inputs to gather

Before drafting anything, ask the user (via `AskUserQuestion`) for any missing items:

1. **Client legal name + trading name**
2. **Primary contact** (name, email, mobile, role)
3. **Billing contact** (if different)
4. **Technical contact / champion** at the client
5. **Tenant domain(s)** — primary M365 domain, vanity domains
6. **Headcount** — current staff, expected 12-month growth
7. **Service tier** — e.g. Bronze / Silver / Gold, or hours-block
8. **Go-live date** for managed services
9. **Existing systems** — current MSP (if switching), LOB apps, line-of-business software, existing backup/AV/RMM
10. **Compliance obligations** — ISO 27001, Essential Eight, HIPAA, PCI, etc.

## Workflow

### Step 1 — Build the client record
Create a structured client profile (markdown or docx) covering:
- Company snapshot
- Contacts table
- Tenant + domain inventory
- Service tier and contracted hours
- Renewal date and notice period
- Escalation matrix (Tier 1 → Tier 2 → Account Manager → vCIO)

### Step 2 — Draft the kickoff communications
Produce 3 deliverables:
- **Welcome email** to the primary contact — warm, clear, lists the next 3 things that will happen this week
- **Internal Teams announcement** to the MSP delivery team — who the client is, what they bought, who owns the relationship
- **Kickoff meeting agenda** — 45 minutes, covers introductions, support process, ticket portal walkthrough, escalation, and Q&A

### Step 3 — Create the technical onboarding checklist
Generate a checklist the engineer will work through:
- [ ] GDAP / delegated admin relationship requested
- [ ] Tenant added to RMM
- [ ] Backup configured (M365 + endpoints + servers)
- [ ] EDR / AV deployed to all endpoints
- [ ] Conditional Access baseline applied
- [ ] MFA verified on all admin accounts
- [ ] Break-glass account created and stored
- [ ] DNS records audited (SPF, DKIM, DMARC, MX)
- [ ] License inventory captured + rightsizing review scheduled
- [ ] Documentation site (IT Glue / Hudu) populated
- [ ] PSA: client, contracts, and contacts entered
- [ ] Standard alerting and reporting enabled

### Step 4 — Schedule the recurring rhythm
Suggest calendar entries the user can confirm:
- **Day 7** — health-check call with client champion
- **Day 30** — 30-day review (engineer + account manager)
- **Day 90** — first formal QBR
- **Quarterly** — recurring QBR cadence

### Step 5 — Deliver
Output the welcome pack as a single Word document (use the `docx` skill) **only if the user asks for a file**. Otherwise return the content inline so the user can copy what they need.

## Guardrails

- Never invent client details — if a field is missing, ask.
- Do not send the welcome email automatically. Always present the draft for the user to review and send.
- Treat the break-glass credential and any tenant secrets as sensitive — never include them in emails or shared documents.
- If the client is switching from another MSP, flag that the offboarding-from-incumbent checklist (credential rotation, GDAP cleanup, backup handoff) is a separate, higher-risk workflow.

## Done when

The user has:
1. A populated client profile
2. Three reviewed draft communications
3. A technical checklist assigned to an engineer
4. The 7 / 30 / 90-day touchpoints in the calendar
