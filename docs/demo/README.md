# Thesis format demo runbook

> **DRAFT/DEMO only.** The rules and samples below are synthetic and are not an official NKU policy check or a compliance conclusion. Do not use real theses or personal data for this demo.

## Requirements

- .NET 10 SDK
- A checkout of this repository

Run the following commands from the repository root. The CLI reads the input and does not modify it.

## Build and test

```sh
dotnet restore NkuThesisTool.sln
dotnet build NkuThesisTool.sln --configuration Release --no-restore
dotnet test NkuThesisTool.sln --configuration Release --no-build
```

## Run the demo

The clean synthetic sample should produce **0 findings**:

```sh
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/clean_sample.docx
```

The known-issues sample should produce **exactly 3 findings**:

```sh
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/known_issues_sample.docx
```

Both reports identify the mode as `DRAFT/DEMO` and show the rule version, coverage, unsupported scopes, and a disclaimer. The three expected findings are:

| Rule | Location | Expected | Actual |
| --- | --- | --- | --- |
| `demo/east-asian-font` | `word/document.xml#paragraph[0]/run[0]` | 宋体 | 黑体 |
| `demo/western-font` | `word/document.xml#paragraph[1]/run[0]` | Times New Roman | Arial |
| `demo/font-size` | `word/document.xml#paragraph[2]/run[0]` | 12pt | 10.5pt |

The expected values and sample contents are recorded in [the synthetic fixture manifest](../../tests/fixtures/synthetic/demo/manifest.json). All fixture text is manually constructed; no thesis or personal information is included.

## What this demo checks

The current reader examines ordinary paragraphs and runs in the main document part. It resolves the direct run formatting, paragraph and character styles with their `basedOn` chains, and document defaults for the East Asian font, ASCII/HAnsi font, and point size. Locations identify the OOXML part and paragraph/run indexes; they are not page numbers.

This is a deliberately small demo. It does not establish compliance with any current university rule set, and it does not automatically repair documents. The report lists unsupported scopes where the reader detected content outside its coverage, including tables, headers/footers, footnotes/endnotes, hyperlinks, tracked changes, content controls, and unresolved theme fonts. Even a clean report means only that these three demo rules did not find a mismatch in the covered runs.

## Troubleshooting

- If the sample path is not found, change to the repository root and retry the command.
- If the CLI project is missing from the build output, run the Release build command above before using `--no-build`.
- If restore fails because the SDK is missing, install the .NET 10 SDK and rerun restore.
- If the findings differ from the three listed above, regenerate the synthetic fixtures using the steps in [the fixture README](../../tests/fixtures/synthetic/demo/README.md), then rerun the tests.

## Demonstration image

[Open the synthetic CLI output PNG](demo-output.png). It is a static visual rendering of excerpts verified by the process-level end-to-end tests, not a live screen capture. Intermediate report sections are omitted; run the commands above for full reports. The image contains only the synthetic sample paths and demo output.
