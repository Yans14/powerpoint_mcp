# Security Policy

## Reporting a vulnerability

Please do not open public issues for security vulnerabilities.
Contact the maintainer directly with reproduction details and impact.

## Secret management requirements

- Never commit API keys, tokens, credentials, or customer confidential data.
- Use `.env` locally and keep only `.env.example` in git.
- Pre-commit and CI include secret scanning (`detect-secrets` and `gitleaks`).

## Hardening checklist

- Enable GitHub branch protection on `main`.
- Require passing `ci` and `security` workflows before merge.
- Enable GitHub secret scanning and push protection for the repository.
