# Security policy

## Threat model

`office-cli` runs **fully offline**. It does not initiate any network connection at
runtime, has no remote authentication, and stores nothing in the cloud. Inputs and
outputs are local files only.

The realistic risks are therefore:

1. A maliciously crafted document (xlsx / pdf / docx / pptx) that triggers a bug
   in one of the underlying parsing libraries (`excelize`, `pdfcpu`,
   `ledongthuc/pdf`, or our own zip+xml code).
2. Audit logs accidentally capturing sensitive flag values.

We address (1) by tracking upstream security advisories and bumping dependencies
promptly. We address (2) by stripping `--token`, `--password`, `--secret`,
`--user-password`, and `--owner-password` flag values in the audit logger.

## Reporting a vulnerability

**Do not open a public issue for security problems.** Email
`guosong6886@gmail.com` with:

- A description of the issue and the impact you observed.
- A minimal repro: the exact command line and a sample file when applicable.
- Any suggested mitigation.

Expect an acknowledgement within 5 business days and a fix or mitigation timeline
within 30 days for confirmed vulnerabilities.

## Supported versions

The latest minor release receives security fixes. Older versions may be patched
on a best-effort basis if downstream packagers request it.

## Local hardening

- Set `OFFICE_CLI_NO_AUDIT=1` if you cannot store any operational metadata under
  `~/.office-cli/audit/`.
- Set `OFFICE_CLI_HOME` to redirect all on-disk state to a custom directory.
- The audit logger writes files with mode `0600` and the home directory with mode
  `0700`; verify these on multi-user machines.
- Set `OFFICE_CLI_PERMISSIONS=read-only` to enforce the most restrictive permission
  level, preventing any write or irreversible operations.
