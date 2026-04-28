---
name: security-baseline-check
description: |
  Runs a Microsoft 365 security baseline review against a defined standard
  (e.g., Essential Eight, CIS, MSP house standard) and produces a gap report
  with prioritized remediation. Use when user says "security review for
  [client]", "are they Essential Eight compliant", "check the baseline",
  "what's [client]'s Secure Score story", or "security gap analysis".
---

# Security Baseline Check

Most MSP security incidents trace back to a control that was supposed to be on, and wasn't. This skill checks the baseline, produces a gap report, and stages the remediation work so nothing slips.

## When to use

- "Security review for [client]"
- "Run the baseline check on [tenant]"
- "Are we Essential Eight aligned"
- "Pre-renewal security audit"
- "What's our Secure Score story for the QBR"

## Choose the standard

Confirm with the user which standard applies:
- **MSP house baseline** (default — most clients)
- **Essential Eight Maturity Level 1 / 2 / 3** (Australia)
- **CIS Microsoft 365 Foundations Benchmark**
- **NIST CSF mapped controls**
- **A specific client framework** (HIPAA, ISO 27001, SOC 2)

## Controls to check (default MSP house baseline)

### Identity
- [ ] MFA enforced for all users (not just admins)
- [ ] Number-matching MFA enabled (no SMS for admins)
- [ ] No legacy authentication protocols permitted
- [ ] Conditional Access: block from outside expected geography (or risk-based)
- [ ] Conditional Access: require compliant or hybrid-joined device for admins
- [ ] Privileged Identity Management (PIM) used for Global Admin and Privileged Role Admin
- [ ] At least 2 break-glass accounts, excluded from CA, credentials in offline vault
- [ ] Self-service password reset configured
- [ ] Sign-in risk and user risk policies enabled (P2 only)

### Email
- [ ] SPF, DKIM, DMARC published and DMARC at p=quarantine or p=reject
- [ ] Anti-phishing policy with impersonation protection for execs
- [ ] Safe Links and Safe Attachments enabled
- [ ] External sender warning banner enabled
- [ ] Outbound spam policy with auto-forwarding disabled
- [ ] Mailbox auditing enabled tenant-wide

### Devices
- [ ] All Windows endpoints enrolled in Intune (or equivalent MDM)
- [ ] Compliance policies enforce: encryption, password, OS version
- [ ] Defender for Endpoint deployed and reporting
- [ ] BitLocker enabled with recovery keys escrowed
- [ ] Application control / WDAC where the tier supports it
- [ ] USB / removable media policy applied
- [ ] Patch SLA: critical < 14 days, high < 30 days

### Data
- [ ] Sensitivity labels published and applied to top sites
- [ ] DLP policies for at-minimum: credit card, government ID, health info
- [ ] External sharing scoped (no anonymous links for sensitive sites)
- [ ] OneDrive / SharePoint retention configured
- [ ] Teams external access and guest access scoped to known partners

### Backup and recovery
- [ ] Third-party M365 backup running (Exchange, OneDrive, SharePoint, Teams)
- [ ] Last successful test restore < 90 days ago
- [ ] Documented RPO and RTO per workload
- [ ] Server / endpoint backup tested

### Visibility
- [ ] Defender / Sentinel alerts routed to MSP SOC or PSA
- [ ] Audit log retention configured (1 year minimum)
- [ ] Secure Score reviewed monthly with a target

## Workflow

1. **Confirm the standard** (default house baseline if unspecified).
2. **Pull the data** — Conditional Access policies, MFA stats, DMARC records (via DNS lookup), Intune compliance reports, Defender deployment, Secure Score, DLP policy state.
3. **Score each control** — Pass / Partial / Fail / N/A with one-line evidence.
4. **Rank gaps** by risk × ease-of-fix:
   - **Quick wins** — high risk, low effort (e.g., enable number matching, enforce DMARC quarantine)
   - **Strategic** — high risk, high effort (e.g., roll out Intune to BYOD)
   - **Hygiene** — medium / low risk (e.g., publish sensitivity labels)
5. **Produce the report**:
   - Executive summary (1 page) — overall posture, top 5 risks, headline recommendations
   - Control-by-control table — Pass / Partial / Fail with evidence
   - Remediation roadmap — 30 / 60 / 90-day plan with owner and effort estimate
6. **Output format** — `docx` for the executive readout, `xlsx` for the control-by-control evidence appendix.

## Guardrails

- **A "Pass" must cite evidence** — the policy name, the screenshot, the DNS record. No "looks good" passes.
- **Don't change policies during the audit.** This is a read-only assessment. Remediation is a separate, scheduled change.
- **Flag "compensating controls"** carefully — claiming a control is met "in spirit" is how breaches start.
- **Some controls require client decisions** (e.g., blocking countries, restricting external sharing). Do not assume; raise as a recommendation.
- **Do not export sensitive policy detail** to a public location. The report is confidential.
- For regulated clients, results may be subject to client confidentiality and external audit — make sure the report is timestamped and the methodology is documented.

## Done when

The client has a posture rating, an evidence-backed gap list, and a 90-day remediation plan with named owners and dates.
