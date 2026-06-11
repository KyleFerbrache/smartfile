<div id="banner" style="width:100%; height:140px; background:#fff; display:flex; align-items:center; justify-content:center;">
  <img src="../assets/Precision Driven - Full Banner.png" height="120" style="padding:10px" />
</div>

# Contributing to SmartFile

Thank you for contributing to **SmartFile**.

This repository is intended to be a maintainable PowerShell module for file-based automation. Contributions should improve the quality, clarity, and operational reliability of the module rather than add one-off scripts without a clear fit.

---

## Contribution Goals

A good contribution to SmartFile generally does one or more of the following:

- improves an existing utility,
- adds a new utility that clearly fits the module's file-centric scope,
- strengthens documentation,
- improves safety or maintainability,
- or corrects behavior that could create operational confusion or inconsistent results.

Before beginning substantial work, make sure the change aligns with the current direction of the module.

---

## What Belongs in This Repository

SmartFile is for **PowerShell utilities that perform file-based operations**.

Examples of good fit include:

- file metadata normalization,
- controlled bulk renaming,
- file inspection or validation utilities,
- and supporting helpers that make those operations safer and easier to maintain.

Contributions are less likely to be accepted if they:

- depend on unrelated frameworks without a strong reason,
- introduce unnecessary complexity,
- are tightly coupled to a one-off local environment,
- or do not include sufficient documentation for future maintainers.

---

## Development Expectations

### Write for Maintainability

Contributed PowerShell should be readable by someone other than the original author. Favor straightforward logic over overly clever implementations.

### Use Clear Naming

Use conventional PowerShell verb-noun naming for command-style scripts and functions wherever practical. Names should describe intent, not implementation detail.

### Comment Intentionally

Use comments to explain important decisions, assumptions, or non-obvious behavior. Avoid comments that simply restate the code.

### Preserve Operational Safety

Because these utilities often modify files directly, changes should be scoped carefully and implemented defensively.

---

## Repository Conventions

### Source Layout

- place PowerShell source under `src/`,
- place script-specific documentation under `docs/`,
- and use `assets/` only for repository media and supporting images.

### Documentation Expectations

If you add or change a utility, update documentation accordingly.

At minimum, contributors should ensure:

- the top-level README still accurately describes the module,
- utility-specific behavior is documented in `docs/`,
- parameters and assumptions are explained clearly,
- and examples are realistic and easy to follow.

### Backward Awareness

Avoid unnecessary breaking changes. If behavior must change, document it clearly and explain why.

---

## Recommended Workflow

1. Fork the repository or create a working branch.
2. Make focused, reviewable changes.
3. Update documentation alongside the code.
4. Test against representative sample files.
5. Open a pull request with a clear description of the change.

---

## Pull Request Guidance

A strong pull request should include:

- a concise summary of the problem being solved,
- an explanation of the chosen approach,
- any assumptions or limitations,
- the files or utilities affected,
- and any documentation updates required for reviewers and future maintainers.

If the change modifies file-writing behavior, explain:

- what is being modified,
- how scope is controlled,
- and how the behavior was validated.

---

## Testing Expectations

Contributors are expected to validate their changes before submitting them.

Validation should be appropriate to the type of change and may include:

- running the utility against representative test files,
- checking edge cases such as missing metadata or unexpected file contents,
- confirming that unchanged files remain unchanged,
- and reviewing output for clarity and operational usefulness.

If a change cannot be tested easily, explain the reason in the pull request.

---

## Style Guidance

When contributing PowerShell:

- prefer explicit parameters over hidden assumptions,
- keep functions and scripts focused on one responsibility,
- handle errors deliberately,
- and keep output meaningful and restrained.

The goal is professional utility code that operators can trust.

---

## Documentation Tone

Write documentation for maintainers and operators, not just for authors.

Good documentation in this repository should be:

- specific,
- implementation-aware,
- operationally useful,
- and free of unnecessary marketing language.

---

## Issues and Discussion

If you are unsure whether a proposed change fits the module, raise the question early through the repository's normal review or discussion process before building a large implementation.

---

## Final Note

SmartFile should remain a cohesive, dependable module. The best contributions strengthen that identity by improving clarity, safety, and maintainability across the repository.
