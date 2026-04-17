# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Run

```bash
# Debug run
dotnet run --project WordAutoTool

# Publish single-file exe
dotnet publish WordAutoTool -c Release -r win-x64
# Output: WordAutoTool/bin/Release/net9.0-windows/win-x64/publish/WordAutoTool.exe
```

Requirements: .NET 9 SDK, Windows, Microsoft Word installed (for COM automation).

## Architecture

This is a **WinForms + ASP.NET Core hybrid** — both run in the same process:

- `Program.cs` starts Kestrel on a random free port (background MTA thread), waits for it to be ready, then runs WinForms on the STA thread.
- `MainForm.cs` hosts a WebView2 control that navigates to `http://localhost:{port}/`. File System Access API permission is auto-granted so the browser picker dialogs don't show trust warnings.
- `wwwroot/` is embedded as resources in the exe (not copied to publish output). The embedded file provider serves them at runtime.

### Request pipeline (`POST /api/process`)

`ProcessController` orchestrates the full pipeline in order:

1. Detect `.doc` vs `.docx` by magic bytes (not file extension)
2. If `.doc` → convert to `.docx` via Word COM (`WordComService.ConvertDocToDocx`)
3. Sort uploaded images using **natural sort** (regex zero-pad all digit groups), rotate portrait images to landscape
4. Replace inline images via Word COM (`WordComService.ReplaceImages`) — preserves original dimensions
5. Replace dates via OpenXML (`WordProcessingService.Process`) — replaces ALL text box content with ROC date; replaces `日期：XXX` patterns in body paragraphs
6. Convert `.docx` → `.doc` via Word COM
7. Optionally convert `.doc` → PDF and zip both; or return `.doc` directly
8. In "overwrite" mode, return with original filename instead of the default `8_查驗照片MMDD` name

### Service responsibilities

- **`WordComService`** (static): All Word COM operations. Uses `Type.GetTypeFromProgID("Word.Application")` + `dynamic` — no Office PIAs needed. Each operation spawns and quits a Word instance via `WithWord()`. All temp files use `wt_{guid}` prefix in `%TEMP%`.
- **`WordProcessingService`** (singleton): OpenXML-only text replacement. Operates entirely in memory on a `MemoryStream`.
- **`InspectService`** + **`InspectController`**: Document scanning — returns info about inline shapes, text boxes, and paragraph text for the UI preview before processing.

### Key constraints

- Word COM requires STA; Kestrel runs MTA. The `WithWord` helper creates/destroys a Word COM instance per call — not pooled.
- Image sort order matters: images are mapped 1-to-1 to inline shapes in document order. Natural sort (zero-padded digit groups) is used to ensure `1_` < `10_` < `2_` doesn't occur.
- `.doc` files must round-trip through `.docx` for OpenXML processing, then back to `.doc`.
