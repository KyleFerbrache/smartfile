<div id="banner" style="width:100%; height:140px; background:#fff; display:flex; align-items:center; justify-content:center;">
  <img src="../assets/Precision Driven - Full Banner.png" height="120" style="padding:10px" />
</div>

# Security Policy

## Scope

This repository contains PowerShell utilities that perform direct file-system operations, including in-place updates to file metadata. Because these utilities can modify files at scale, security in this context is primarily about **safe execution, controlled access, and responsible change management**.

This document defines how to use SmartFile responsibly and how to report security concerns related to the repository.

---

## Supported Use

SmartFile utilities are intended to be used in environments where:

- the executing user or automation identity has legitimate access to the target files,
- the files being modified are understood and within scope,
- and changes can be validated by the team responsible for the workflow.

These utilities should not be run against unknown, untrusted, or poorly scoped file sets.

---

## Security Expectations

When using or maintaining SmartFile, the following expectations apply:

### 1. Principle of Least Privilege

Run utilities with only the file-system permissions required for the task at hand. Avoid elevated execution unless it is explicitly necessary and approved.

### 2. Controlled Targeting

Before running any utility that modifies files:

- confirm the root path,
- confirm the file types in scope,
- validate any filters or patterns,
- and, where practical, test against a small subset first.

### 3. Change Awareness

Utilities in this repository may update files in place. Users and maintainers should ensure that intended changes are understood before broad execution in operational repositories, shared drives, or synchronized folders.

### 4. Trusted Execution Context

Run SmartFile only from trusted workstations, approved administrative jump hosts, or sanctioned automation environments. Do not execute modified versions of repository scripts from unverified sources.

---

## Safe Usage Guidance

To minimize operational risk:

- validate path inputs before execution,
- avoid running against files that are open or actively being edited,
- review utility-specific guidance in the `docs/` directory,
- and use version control for the repository so changes to source are reviewable.

If a utility is being introduced into a new environment, the recommended practice is to:

1. test with a representative non-production sample,
2. confirm expected file changes,
3. review any summaries or logs produced,
4. then expand the scope deliberately.

---

## What to Report

Please report issues such as:

- behavior that causes unintended file modifications,
- path handling or pattern handling that could expand scope unexpectedly,
- unsafe defaults that increase the chance of accidental bulk changes,
- weaknesses in input validation,
- or opportunities for a malicious or mistaken operator to cause broader impact than intended.

Operational bugs are important as well, but issues with **security impact** should be reported privately.

---

## Reporting a Vulnerability

Please **do not** open a public issue for a suspected security problem.

Instead, report it privately to the repository maintainers through your organization's approved security reporting channel or internal contact path.

When reporting a concern, include:

- the affected utility,
- a concise description of the issue,
- the conditions required to reproduce it,
- the potential impact,
- and any suggested mitigation if available.

A clear report helps maintainers assess severity, confirm scope, and respond appropriately.

---

## Maintainer Responsibilities

Maintainers should:

- review PowerShell changes for file-scope safety,
- prefer explicit parameters and clear defaults,
- document breaking behavior changes,
- and avoid merging file-modifying logic that has not been tested with realistic examples.

Security is not only about hostile misuse; it is also about reducing the chance of accidental high-impact changes in everyday operations.

---

## Repository Hygiene

To keep the repository safer to operate and maintain:

- avoid committing environment-specific secrets or credentials,
- avoid embedding organization-sensitive paths in source files,
- document assumptions in code and in `docs/`,
- and prefer small, reviewable pull requests for behavioral changes.

---

## Disclosure and Remediation

When a security issue is confirmed, maintainers should:

1. assess the practical impact,
2. identify affected utilities and usage conditions,
3. prepare a fix and supporting documentation updates,
4. communicate remediation guidance to users of the repository,
5. and, where necessary, advise teams to validate any prior runs.

The goal is to resolve issues responsibly while preserving trust in the utilities and their operational use.
