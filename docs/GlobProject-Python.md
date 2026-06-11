<div id="banner" style="width:100%; height:140px; background:#fff; display:flex; align-items:center; justify-content:center;">
  <img src="../assets/SmartFile Banner.png" height="120" style="padding:10px" />
</div>

# GlobProject-Python

### Purpose

`GlobProject-Python` is a PowerShell utility within the **SmartFile** module that recursively collects Python source files from a project directory and consolidates them into a single output file named `glob.py`.

The utility exists to support file-based review, archival, handoff, and analysis workflows where a complete project-wide representation of Python source code is easier to consume as a single artifact than as a directory tree.

In practical terms, the command is used to flatten a Python project into a deterministic, readable output format that preserves file boundaries while reducing the operational overhead of working across many small source files.

This is especially useful when teams need to:

- provide a consolidated code artifact for internal review,
- prepare a project snapshot for downstream analysis tooling,
- create a portable source bundle without packaging the repository itself,
- or inspect a project quickly without browsing its full directory structure.


#### When to Use It

Use `GlobProject-Python` when all of the following are true:

1. A project directory contains Python source files distributed across multiple subdirectories.
2. A single-file representation of that project is more useful for the current workflow than the original directory structure.
3. Virtual environment and cache directories should be excluded from the collected source set.
4. The output should be deterministic and easy to regenerate as the underlying project changes.

This utility is especially useful for inspection and export workflows where preserving every repository feature is not necessary, but preserving source visibility and file boundaries is.

---

## Operating Requirements

`GlobProject-Python` operates entirely against the local file system. It does **not** call package managers, source control systems, or remote APIs.

The command scans a local root path, identifies `.py` files, excludes certain directory names, and writes a consolidated output file back into the same root directory.

#### Required Conditions

The provided root path must:

- exist locally,
- be accessible to the executing user,
- and contain Python source files beneath it for the command to produce meaningful output.

The command writes its output file to:

```text
<RootPath>\glob.py
```

Because the output file itself is a Python file located in the scanned directory, the command explicitly excludes `glob.py` from future runs so that repeated executions do not recursively ingest prior output.

#### Default Directory Exclusions

The command excludes the following directory names by default:

- `.venv`
- `__pycache__`

These defaults are intentionally built in because both directories typically contain generated or environment-specific artifacts rather than source code that should appear in a project-level consolidated output.

#### Additional Exclusions

Callers may provide extra directory names through the cmdlet so the working set can be narrowed further.

Typical examples include:

- `.git`
- `build`
- `dist`
- `.mypy_cache`
- `.pytest_cache`

Additional exclusions are matched by **directory name** as a path segment, not by wildcard pattern.

---

## Functional Behavior

At execution time, `GlobProject-Python` performs the following high-level operations:

1. Resolves the provided root directory.
2. Recursively scans that root for `.py` files.
3. Excludes `.venv` and `__pycache__` directories by default.
4. Applies any additional excluded directory names provided by the caller.
5. Excludes the generated `glob.py` file itself if it already exists.
6. Computes a relative path for each eligible file.
7. Sorts all eligible files in deterministic relative-path order.
8. Reads the contents of each file.
9. Writes those contents into `glob.py` using comment-delimited file headers.
10. Emits progress during scanning and concatenation, followed by a concise execution summary.

The command intentionally avoids packaging, parsing, or interpreting Python code. It performs a deterministic file collation task only.


### Output Format

Each included file is written into the output using a section header that identifies its relative path within the project.

The output structure follows this format:

```python
# ======== relative/path/to/file.py ======== #
<file contents>
```

This preserves the boundary between source files while keeping the output easy to read, search, and navigate.

The command trims trailing whitespace at the end of each file before writing so that repeated execution does not accumulate excess blank lines in the final output.

##### Not Supported

- non-Python file types,
- wildcard-based exclusion patterns,
- dependency capture,
- repository history export,
- or preservation of original per-file metadata inside the consolidated output.


### Parameters

The documented command behavior assumes the following inputs.

##### `-RootPath`

Specifies the local root directory to scan recursively.

The resulting output file is always written to that same directory as `glob.py`.

##### `-AdditionalExcludeDirectories`

Optional string array of additional directory names to exclude.

These values are appended to the built-in exclusions for `.venv` and `__pycache__`.

This parameter is intended to support projects with additional generated, cached, or distribution-related directories that should not appear in the consolidated output.

---

### Typical Execution Pattern

A typical use case looks like this:

1. A Python project exists locally with source distributed across multiple subfolders.
2. The project also contains environment-specific or generated directories that should not be included.
3. `GlobProject-Python` is run against the project root.
4. The command writes a deterministic `glob.py` file into that root.
5. The resulting single-file artifact is used for review, handoff, inspection, or downstream processing.

This pattern is particularly effective when the goal is to preserve source visibility while simplifying how the project is consumed for a specific operational task.

---

## Detailed Processing Logic

#### Step 1: Recursive Discovery

The command walks the specified root path recursively and evaluates candidate `.py` files.

All discovered Python files are scanned first so the command can report progress and determine the eligible file set before output generation begins.

#### Step 2: Directory Exclusion Evaluation

For each discovered file, the command computes its relative path from the root and splits that path into directory segments.

If any segment matches one of the excluded directory names, the file is removed from the working set.

This ensures that excluded directories are omitted consistently regardless of how deeply they are nested.

#### Step 3: Output File Self-Exclusion

If a previous run already produced `glob.py` in the root directory, that file is excluded from discovery.

This is an essential safeguard because the output file is itself a `.py` file and would otherwise be eligible for re-ingestion during subsequent runs.

#### Step 4: Deterministic Sort Order

Eligible files are sorted by relative path before concatenation begins.

This makes the output stable across repeated runs provided the file set and file contents remain unchanged.

Deterministic ordering is important for:

- reviewability,
- reproducibility,
- diff-based comparisons,
- and predictable downstream processing.

#### Step 5: Content Read and Write

Each eligible file is read as text and written to the output file under a comment-delimited relative-path header.

The command uses UTF-8 output encoding without BOM and appends a blank line between file blocks for readability.

#### Step 6: Summary Output

The command is designed to keep console output minimal while still exposing progress during scanning and concatenation phases. A compact summary is emitted at the end of execution so operators can validate the scope and result of the run.

---

### Why This Utility Uses Deterministic File Concatenation

There are several reasons the implementation favors direct, deterministic file concatenation rather than packaging or repository-level export.

**Operational Simplicity**

The task is intentionally narrow: collect source files, exclude known non-source directories, and produce a readable single-file output. Introducing packaging semantics would make the command heavier than necessary for this purpose.

**Deterministic Review Artifacts**

A stable file ordering and repeatable output structure are useful when the generated file is meant to be inspected, compared, or passed through downstream tools.

**Minimal Runtime Dependencies**

The command relies only on local file-system access and standard PowerShell capabilities. It does not require Git, Python packaging tools, or repository-specific tooling to perform its job.

---
## Operational Risks and Considerations

### Limitations

`GlobProject-Python` is intentionally narrow in scope.

**It Does Not**
- package the project,
- validate or lint Python code,
- interpret imports or dependencies,
- preserve file permissions or metadata in the output,
- or filter files based on contents.

**It Assumes**

- the project exists locally,
- `.py` files are the desired source scope,
- excluded directory names are sufficient to remove unwanted generated content,
- and a single-file source representation is acceptable for the downstream workflow.


### Large Project Strategy

For large repositories, the resulting `glob.py` file may become substantial in size. This is expected and should be considered before using the output in tools or environments with file-size or token limitations.

### Encoding Expectations

The command reads files as UTF-8 and writes the output as UTF-8 without BOM. Projects containing files encoded differently may require review if decoding issues are encountered.

### Partial Read Failures

If an individual file cannot be read, the command increments an error counter and continues processing the remaining files. This allows the run to complete while still surfacing that not all candidate files were successfully written.

### Repeated Runs

Repeated execution is safe from a file-discovery perspective because the output `glob.py` is explicitly excluded. As a result, subsequent runs regenerate the artifact from the original project files rather than recursively compounding prior output.

---

## Example Usage Patterns

### Basic Project Consolidation

```powershell
. "C:\Path\To\SmartFile\scripts\GlobProject-Python.ps1" -RootPath "C:\Path\To\ProjectRoot"
```

### Consolidation with Additional Exclusions

```powershell
. "C:\Path\To\SmartFile\scripts\GlobProject-Python.ps1" -RootPath "C:\Path\To\ProjectRoot" -AdditionalExcludeDirectories ".git", "build", "dist"
```

### Validation Checklist

After running the command, validate the following:

- The expected number of files were discovered.
- The expected number of files were eligible after exclusions.
- The output `glob.py` file exists in the project root.
- File section headers reflect the expected relative paths.
- A sample of source sections appear in the expected deterministic order.
- The summary counts for files, lines, and errors align with expectations.

---

## Troubleshooting

#### Symptom: The Output Includes Files That Should Have Been Excluded

Check the exclusion input first:

- Was the unwanted content inside a directory named `.venv` or `__pycache__`?
- If not, was the additional directory name supplied correctly through `-AdditionalExcludeDirectories`?
- Does the excluded value match the actual directory segment name rather than a broader wildcard concept?

#### Symptom: `glob.py` Is Reappearing in the Working Set

Confirm that the command is running against the intended root path and that the generated output file is being written to that same root as `glob.py`.

The command is designed to exclude that file explicitly during discovery.

#### Symptom: Fewer Files Than Expected Are Included

Common causes include:

- a root path that is too narrow,
- source files stored under excluded directories,
- or file read failures counted during execution.

#### Symptom: The Output File Is Very Large

This is common for larger repositories or deeply nested projects. Consider narrowing the project root or adding additional excluded directories if the resulting artifact is too broad for the intended downstream workflow.

---


## Summary

`GlobProject-Python` is a deterministic project-collation utility for Python source code within the SmartFile module.

Its job is narrow but operationally valuable:

- discover Python source recursively,
- exclude known environment and cache directories,
- preserve readable file boundaries,
- and produce a stable single-file artifact that is easy to review, transport, and process.

When paired with a well-scoped project root and appropriate exclusion settings, the utility provides a clean and repeatable way to flatten a Python codebase into a single generated source artifact.
