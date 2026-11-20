# Entra ID & M365 Identity Lifecycle Pipeline

Declarative Joiner-Mover-Leaver (JML) identity automation pipeline built with Microsoft Graph REST API, PowerShell, Azure RBAC, and passwordless GitHub Actions OIDC.

## Capabilities

* **Automated Onboarding (Joiner):** Provisions Entra ID accounts with enforced first-login password rotation, adds users to cohort Security Groups, assigns available M365 license SKUs, and grants least-privilege Azure Reader RBAC.
* **Instant Offboarding (Leaver):** Disables Entra ID sign-in, revokes all active refresh/session tokens via revokeSignInSessions, removes group memberships, reclaims M365 licenses, and strips Azure RBAC permissions.
* **Zero-Secret CI/CD & Audit Artifacts:** Authenticates via Workload Identity Federation (OIDC) and outputs a timestamped JSON compliance audit trail on every run.
