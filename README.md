# ATK VPE Automation

Automation system for Advanced Toastmasters of Kerala (ATK) VPE operations.

## Deployment model

- GitHub is the source-control layer.
- GitHub Actions deploys source changes automatically to the **ATK DEV Apps Script project** using `clasp`.
- DEV is the only automatic deployment target.
- LIVE/production must never be changed without explicit approval.
- Secrets are not stored in this repository.

## DEV targets

- Apps Script Script ID: `1LS3HabXjXOainAMiNLVVlRX5LEaA-o5Y7Kf31Fg-JQJZsR4ZKHI5zHUN`
- DEV Spreadsheet ID: `1qZnZov1l9DfvhTzruqlLubGhydcOOFKLf2kG8xplNLM`

The initial workflow safely bootstraps the current Apps Script source into this repository before future code changes are deployed from GitHub.

<!-- deployment refresh 2026-10-02 agenda -->
