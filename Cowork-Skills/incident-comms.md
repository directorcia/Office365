---
name: incident-comms
description: |
  Drafts client-facing incident communications during an outage or major event —
  initial notification, hourly updates, all-clear, and post-incident review.
  Use when user says "we have an outage", "draft an update for clients",
  "M365 is down — what do we say", "send the all-clear", or "write the PIR".
---

# Incident Comms

During an outage, what you say is as important as what you do. Customers tolerate downtime; they don't tolerate silence. This skill produces calm, accurate, on-brand updates at every stage of an incident.

## When to use

- "We have an outage / incident / P1"
- "Draft an update for affected clients"
- "M365 / Azure / Exchange is down — what do we tell people"
- "Send the all-clear"
- "Write the post-incident review"

## Principles

1. **Acknowledge fast, even with little info.** Silence breeds calls.
2. **Say what you know, what you don't, and when you'll update next.**
3. **One voice, one channel.** Pick the channel (email, status page, Teams) and stick with it for the incident.
4. **Plain English. No vendor jargon, no blame, no speculation about root cause until confirmed.**
5. **Customers care about their service, not your tickets.** Lead with impact to them.

## Inputs to gather

Before drafting:
1. **What's affected** — service, region, clients, user count
2. **When it started** (timestamp + timezone)
3. **What we know** — symptoms, vendor status page, internal investigation findings
4. **What we're doing** — actions in flight
5. **ETA for next update** — even if "no ETA, next update in 30 minutes"
6. **Audience** — all clients, affected clients only, internal staff, leadership

## Templates

### 1. Initial notification (within 15 min of declaring incident)

> **Subject:** [Service] disruption — we're investigating
>
> We're aware of an issue affecting **[service]** that started at **[time, timezone]**. Symptoms include **[1–2 user-visible symptoms]**.
>
> Our team is engaged and investigating. We'll send the next update by **[time]** even if we don't yet have a fix.
>
> If you're affected and need to log a related ticket, please reply to this email and reference **incident [ID]** so we can keep your case linked.

### 2. Progress update (every 30–60 min during active incident)

> **Subject:** [Service] disruption — update [N]
>
> **Status:** Investigating / Identified / Mitigating / Monitoring
>
> **What we know now:** [1–3 sentences, factual only]
>
> **What we're doing:** [Current action]
>
> **Impact:** [Who, what, how many]
>
> **Next update:** [Time]

### 3. Resolution / all-clear

> **Subject:** [Service] disruption — resolved
>
> The issue affecting **[service]** was resolved at **[time]**. Total duration: **[X hours Y minutes]**.
>
> Services are now operating normally. We're monitoring closely for recurrence.
>
> A post-incident review with root cause and preventive actions will follow within **5 business days**.
>
> Thank you for your patience. If you continue to experience issues, please log a ticket referencing **incident [ID]**.

### 4. Post-incident review (within 5 business days)

Sections:
- **Summary** — 3 sentences: what happened, who was affected, how long
- **Timeline** — detection, declaration, mitigation, resolution
- **Root cause** — what actually caused it (technical and contributing factors)
- **What worked** — detection time, escalation, communication
- **What didn't** — be honest, blameless, specific
- **Preventive actions** — each with an owner and a date
- **Customer impact** — including any SLA credits owed

## Workflow

1. **Confirm facts** before drafting. Pull from vendor status pages (Microsoft 365 Service Health, Azure Status), monitoring tools, and the engineering bridge — not assumptions.
2. **Match the template** to the incident stage.
3. **Tailor the audience** — leadership at the client wants impact and ETA; the IT contact wants symptoms and workarounds.
4. **Send via the right channel** — confirmed status page first, email second, Teams third. Don't fragment.
5. **Log the comms** in the incident ticket so the PIR can reference them.

## Guardrails

- **Never speculate about root cause** in customer comms until confirmed by engineering. "Investigating" is honest; a wrong root cause statement creates trust debt.
- **Never blame a vendor by name** in initial comms unless the vendor has publicly acknowledged. Say "an upstream provider" until confirmed.
- **Never promise an ETA you can't hit.** Promise an *update time* instead.
- **Do not auto-send.** Every customer-facing update must be reviewed by the incident commander before going out.
- **Treat affected client lists as confidential** — don't CC clients onto a single email; use BCC or per-client sends.
- For security incidents involving potential data loss, escalate to legal/compliance before any external communication.

## Done when

The incident has a clean comms log, a final all-clear, and a scheduled PIR.
