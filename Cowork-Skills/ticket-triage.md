---
name: ticket-triage
description: |
  Triages incoming MSP support tickets — classifies by urgency and impact,
  assigns priority, suggests the right queue and engineer tier, and drafts the
  first client response. Use when user says "triage these tickets", "what
  should I work on next", "prioritize the queue", or "is this a P1".
---

# Ticket Triage

Every minute spent on the wrong ticket is a minute stolen from a P1. This skill applies a consistent, defensible triage framework so prioritization survives shift handover.

## When to use

- "Triage my queue"
- "What should I pick up next"
- "Is this a P1 or P2"
- "Help me prioritize these tickets"
- A bulk dump of unread tickets pasted into chat

## Triage framework

Use Impact × Urgency to derive priority — never priority by gut feel.

### Impact
| Level | Definition |
|-------|------------|
| High | Whole site, whole client, or business-critical system down |
| Medium | A team or department affected, or a critical user (CEO, head of finance during close) |
| Low | One user, with a workaround |

### Urgency
| Level | Definition |
|-------|------------|
| High | Work has stopped right now |
| Medium | Work is degraded but proceeding |
| Low | Future-dated or scheduled |

### Priority matrix
|  | Urgency H | Urgency M | Urgency L |
|---|---|---|---|
| **Impact H** | P1 | P2 | P3 |
| **Impact M** | P2 | P3 | P4 |
| **Impact L** | P3 | P4 | P4 |

### SLA defaults (override per contract)
- **P1** — respond 15 min, restore 2 hours, all-hands
- **P2** — respond 30 min, restore 4 hours
- **P3** — respond 2 hours, resolve next business day
- **P4** — respond 1 business day, resolve within 5

## Workflow

For each ticket:

1. **Read the subject + body**, ignore the user's self-assigned priority.
2. **Classify** Impact and Urgency using the table above.
3. **Derive priority** and compare to SLA. Flag any ticket already at risk.
4. **Identify category**: account/access, M365, network, hardware, application, security, request (new starter, new app).
5. **Route**:
   - Tier 1 — password resets, mailbox permissions, simple how-to
   - Tier 2 — Intune, Conditional Access, AV alerts, complex M365
   - Tier 3 — networking, server, identity architecture, security incidents
   - Project queue — anything > 4 hours of work
6. **Detect security signals** — phishing report, suspicious sign-in, MFA fatigue, ransomware indicators. These bypass the matrix and become P1 regardless of stated impact.
7. **Draft a first response** to the client — acknowledges the issue, sets the right expectation against SLA, and asks for the one missing piece of information that will unblock the engineer.

## Special cases

- **VIP user** — bump impact one level (CEO, partners, named contacts in the contract).
- **Multiple tickets, same root cause** — merge and treat as one P1/P2 incident, not N separate P3s.
- **Repeat ticket from same user this week** — flag for an engineer-led root-cause review, not another quick fix.
- **Ticket from a new starter on day 1** — high urgency regardless of stated impact; first-day experience drives client satisfaction.

## Output format

For a queue, return a table:

| # | Client | Subject | Cat | Impact | Urg | Priority | SLA risk | Suggested owner | First-response draft |
|---|--------|---------|-----|--------|-----|----------|----------|-----------------|----------------------|

For a single ticket, return:
- **Priority + reasoning** (1–2 sentences)
- **Suggested queue / tier**
- **Draft first response** (the user can paste straight into the PSA)
- **What you need from the client** (the one question)

## Guardrails

- Never auto-respond. Always present drafts for the engineer to send.
- Do not invent client SLAs. If the SLA is unknown, say so and use defaults with a flag.
- Security signals are never downgraded — even if the user pushes back, escalate.
- If a ticket mentions credentials, secrets, or session tokens in the body, flag it for cleanup.

## Done when

Every ticket in the input set has a priority, an owner, and a drafted response.
