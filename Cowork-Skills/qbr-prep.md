---
name: qbr-prep
description: |
  Prepares a Quarterly Business Review (QBR) pack for an MSP client — pulls
  ticket trends, security posture, license utilization, project status, and
  strategic recommendations into a presentation-ready brief. Use when user says
  "QBR for [client]", "prep the quarterly review", "build the business review
  deck", or "what should I tell [client] this quarter".
---

# QBR Prep

A QBR is the single most important commercial conversation an MSP has with a client each quarter. This skill produces a defensible, data-led pack that justifies the spend, surfaces risk, and opens upsell.

## When to use

- "Build a QBR for [client]"
- "Prep the quarterly review"
- "What's our story for [client] this quarter"
- "QBR deck"

## Inputs

Ask the user for:
1. **Client name** and tenant
2. **Reporting period** (e.g., Q1 2026: Jan–Mar)
3. **Attendees** on the client side (and their roles — CEO needs business outcomes, IT manager needs operational detail)
4. **Strategic themes** from the last QBR, if any
5. **Known sensitivities** (e.g., recent outage, budget pressure, M&A activity)

## Sections to produce

### 1. Executive summary (1 slide)
3 bullets the CEO can repeat:
- What we delivered
- What we protected them from
- What we recommend next

### 2. Service delivery scorecard
- Tickets opened / closed / still open
- Average response and resolution time vs SLA
- Top 3 ticket categories (and what they say about user experience or system health)
- After-hours / P1 incidents

### 3. Security posture
- MFA coverage %
- Conditional Access policy state
- Patch compliance (endpoints, servers)
- EDR detections and response actions
- Phishing simulation results (if applicable)
- Secure Score trend

### 4. Backup and resilience
- Backup success rate
- Last successful test restore date
- RPO/RTO posture vs documented standard

### 5. License and cost optimization
- Current M365 license mix and utilization
- Inactive accounts still licensed
- Recommended rightsizing — the dollar number matters
- Upcoming renewals and expected price impact

### 6. Projects
- Completed this quarter
- In flight
- Recommended for next quarter (with rough effort and value)

### 7. Strategic roadmap (next 90 / 180 / 365 days)
Tie each item to a business outcome the client cares about:
- Reduce risk
- Reduce cost
- Enable growth
- Improve productivity

### 8. Risks and asks
- Top 3 risks the client should know about
- What you need from them (decisions, budget, access, sponsorship)

## Workflow

1. **Gather data** — pull from PSA (tickets), RMM (patch/backup), security tooling (Secure Score, EDR), and M365 admin (licenses). Prefer tool calls over recall.
2. **Spot the story** — 2–3 themes connect every section. Lead with the theme, support with data.
3. **Draft the narrative** as bullets first, in plain English. Numbers without narrative is a report, not a QBR.
4. **Build the deck** — invoke the `pptx` skill. Use the client's brand colours if known, otherwise the MSP's house template.
5. **Produce a leave-behind** — a 2-page Word summary (`docx` skill) for attendees who want detail.

## Guardrails

- **Numerical accuracy is non-negotiable.** Compute every percentage and total with a code tool — do not eyeball arithmetic. A wrong MFA % undermines the whole pack.
- Do not paint over bad news. A missed SLA discussed openly builds more trust than a chart that hides it.
- Never quote a competitor's pricing or product unless the user provides the source.
- Recommendations must be specific and costed at a high level. "Look at Intune" is not a recommendation — "Roll out Intune to the 42 unmanaged laptops, ~12 hours engineering, ~$X licensing" is.
- If data is missing for a section, say so on the slide. Empty sections are better than fabricated numbers.

## Done when

The user has a deck, a leave-behind, and a one-paragraph talk-track for the opening of the meeting.
