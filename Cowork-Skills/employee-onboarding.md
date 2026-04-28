---
name: employee-onboarding
description: |
  Runs the new-starter provisioning workflow for an MSP client — creates the
  user, assigns licenses and groups, prepares hardware, schedules induction,
  and produces a welcome pack. Use when user says "new starter for [client]",
  "set up [name] at [company]", "onboard [name] starting [date]", or
  "provision a new user".
---

# Employee Onboarding

The first day at a new job is the loudest test of the MSP relationship. A missed mailbox or a broken laptop on day 1 is what clients remember at renewal. This skill standardizes the provisioning flow.

## When to use

- "New starter at [client] — [name] starting [date]"
- "Onboard [name]"
- "Set up [name]'s account"
- "Provision a new user"

## Inputs to gather

Ask the user via `AskUserQuestion` for any missing fields:

1. **Client tenant**
2. **Full legal name** (for the account)
3. **Preferred display name** (what they go by)
4. **Job title and department**
5. **Manager** (so groups, access requests, and the welcome email can be routed)
6. **Office / location**
7. **Start date and time**
8. **Employment type** — full-time, part-time, contractor (affects licensing)
9. **Role template** — match to an existing role profile if available (e.g., "Sales rep", "Site engineer", "Frontline / shift worker")
10. **Hardware needs** — laptop model, peripherals, mobile, headset
11. **Special access** — finance system, CRM, line-of-business apps
12. **Sensitive flags** — executive (extra protection), board member, contractor (date-bounded access)

## Workflow

### 1. Identity provisioning
- Create the user account in Entra ID with a standard naming convention
- Set initial password and force change on first login
- Enable MFA registration requirement
- Assign manager attribute
- Add to all-staff and department dynamic groups (if not membership-rule driven)

### 2. License assignment
- Match the role template to an SKU (E3 / Business Premium / F3 / etc.)
- Assign via group-based licensing where possible
- Add any role-specific add-ons (Teams Phone, Power BI Pro, Project)

### 3. Mailbox and Teams
- Confirm mailbox provisioned
- Add to shared mailboxes the role requires
- Add to Teams the role requires
- Add to distribution lists per the role template

### 4. Files and apps
- SharePoint site memberships
- OneDrive provisioned (light-warm if the role uses Office)
- LOB app access requests submitted (CRM, ERP, line-of-business apps) — capture the access form for client sign-off
- VPN / SSE profile if applicable

### 5. Devices
- Laptop imaged (Autopilot profile assigned to device serial)
- Compliance policy applied
- BitLocker enrolled, recovery key escrowed
- Defender for Endpoint reporting
- Mobile device enrolled if BYOD policy permits

### 6. Day-1 experience
- Welcome email drafted to the new starter (sent from their manager, not the MSP)
- Day-1 sign-in walk-through (laptop, MFA setup, password manager, Teams)
- IT induction scheduled (15–30 minutes, video call) on day 1 or day 2
- Quick-reference card: how to log a ticket, hours of support, who to call for what

### 7. Manager briefing
- Confirm what's been set up
- Flag any access still pending client approval
- Confirm the Day-1 contact for the new starter

## Outputs

- A populated checklist (markdown or Excel) tracking each step
- The welcome email draft (for the manager to send)
- An IT induction agenda
- A quick-reference card (PDF if requested)
- A pending-approvals list — anything the client needs to sign off

## Guardrails

- **Never set "must change password at next sign-in" off.** Always force a change.
- **Never email the temporary password and the username together.** Send via two channels (email username, SMS / phone the temp password to the manager).
- **Don't enrol the user in groups that grant elevated privileges** unless explicitly requested and signed off by the client.
- **Contractors get end-dated accounts.** Set the access expiry up front; don't rely on memory.
- **Executives** — apply the executive protection profile (impersonation protection, Conditional Access scope, restricted external sharing) automatically.
- **Avoid "copy-from-existing-user"** as a shortcut — it copies stale group memberships and creates entitlement drift. Use the role template instead.
- **Pause before sending** — surface every drafted communication for review. Don't auto-send.

## Done when

- All checklist items are green or explicitly waived
- The manager has confirmed the Day-1 plan
- The new starter has a working sign-in, mailbox, Teams, laptop, and induction booked
- Any pending client approvals are tracked with a follow-up date
