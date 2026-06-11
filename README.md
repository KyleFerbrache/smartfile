<div id="banner" style="width:100%; height:140px; background:#fff; display:flex; align-items:center; justify-content:center;">
  <img src="assets/Precision Driven - Full Banner.png" height="120" style="padding:10px" />
</div>

## Overview

**SmartFile** is a PowerShell module for deterministic, file-based automation in enterprise environments.

The module is organized around a simple goal: provide reliable utilities for bulk file operations without introducing unnecessary runtime dependencies or opaque execution behavior. Each utility is designed to operate directly against the file system, favor predictable outcomes, and remain readable and maintainable for teams that support operational scripts over time.

Today, the module includes utilities focused on **metadata normalization for Excel workbooks** used in downstream SharePoint indexing scenarios.

---

## Design Principles

SmartFile is built around the following principles:

- **File-first operations**  
  The module operates directly on files and directories, with behavior that is easy to reason about and validate.

- **Deterministic execution**  
  Given the same inputs, utilities should produce the same outputs with minimal side effects.

- **Operational clarity**  
  Scripts are written to be understandable by analysts, engineers, and maintainers who may need to review or extend them later.

- **Minimal dependencies**  
  Utilities should avoid heavy runtime dependencies when native PowerShell and .NET capabilities are sufficient.

- **Enterprise suitability**  
  Documentation, structure, and contribution patterns are written with internal teams and long-term support in mind.

---

## Current Scope

SmartFile currently provides PowerShell utilities for file-based metadata operations.

### Included Utility

- **`Set-SharePointIndex`**  
  Updates the internal **Title** metadata of Excel (`.xlsx`) files so that the document title aligns with the workbook filename. This supports downstream indexing and retrieval workflows in SharePoint document libraries where metadata consistency improves searchability and reduces unnecessary processing overhead.

---

## Module Structure

The repository is organized to keep source, documentation, and supporting assets separate and maintainable.

```text
smartfile/
├── assets/
├── docs/
├── src/
├── README.md
├── CONTRIBUTING.md
└── SECURITY.md
```

- **`assets/`** contains shared images and other repository media.
- **`docs/`** contains script-specific guides and usage references.
- **`src/`** contains the PowerShell source files for the module utilities.

---

## Documentation

The top-level README is intentionally focused on the **module itself**: its purpose, scope, conventions, and contribution model.

Detailed implementation notes, usage examples, and utility-specific guidance should live in the **`docs/`** directory so they can evolve independently without turning the repository landing page into a script manual.

---

## Installation

Clone the repository to a local working directory:

```bash
git clone <repository-url>
```

Move into the repository root:

```bash
cd smartfile
```

If your execution policy requires it, unblock local scripts before use:

```powershell
Get-ChildItem .\src\*.ps1 | Unblock-File
```

---

## Execution Model

SmartFile utilities are intended to be run by users with appropriate file-system access to the target directories they manage.

In practice, that means:

- utilities should be executed from a trusted workstation or approved automation host,
- the caller should have read/write access to the target files,
- files should not be locked by another process during execution,
- and any environment-specific usage details should be documented in the corresponding guide under **`docs/`**.


### Set-SharePointIndex

Command to preview changes for all files modified since midnight this morning
```powershell
. "C:/Path/To/SmartFile/Scripts/Set-SharePointIndex.ps1" -rootPath "Freeman Mathis and Gary/FMG - Financial Systems - Documents/01 - Data Sources" -Days 0 -DoNotApply
```

Command to apply changes for all files modified since midnight yesterday morning
```powershell
. "C:/Path/To/SmartFile/Scripts/Set-SharePointIndex.ps1" -rootPath "Freeman Mathis and Gary/FMG - Financial Systems - Documents/01 - Data Sources" -Days 1
```

---

## Who This Repository Is For

SmartFile is intended for teams that need to manage large collections of files in a controlled and repeatable way, including:

- data and reporting analysts,
- operations engineers,
- document governance teams,
- and automation developers supporting file-centric workflows.


#### Repository Standards

This repository favors:

- explicit, well-commented PowerShell,
- descriptive naming,
- careful handling of file modifications,
- and documentation that explains not only **what** a utility does, but also **why** it exists.

Contributors should treat SmartFile as a maintainable internal toolset rather than a collection of one-off scripts.

---

## Contributing & Security

- Guidance for proposing changes, preparing pull requests, naming utilities, and maintaining code quality is available in [CONTRIBUTING.md](CONTRIBUTING.md).


- Security expectations, reporting guidance, and safe usage practices for file-modifying utilities are documented in [SECURITY.md](SECURITY.md).

---

## License

This repository is not currently licensed.
