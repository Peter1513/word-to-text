# word-to-text

[![tests](https://github.com/Peter1513/word-to-text/actions/workflows/test.yml/badge.svg)](https://github.com/Peter1513/word-to-text/actions/workflows/test.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A single PowerShell script that batch-converts Word documents to **Markdown** (default), plain text, or JSON.

- **No installs.** `.docx` is read directly from its ZIP/XML. No Word, no Pandoc, no Python, no modules.
- **Runs anywhere PowerShell does.** Windows PowerShell 5.1 (built into Windows 10/11) and PowerShell 7+ on Windows, Linux, and macOS.
- **Legacy formats.** `.doc` / `.rtf` / `.odt` are converted through Microsoft Word (if installed) or LibreOffice, then go through the same parser.
- **Safe for batch jobs.** A broken file never stops the run, existing output is never overwritten unless you ask, and exit codes tell scripts what happened.

[中文說明在下方](#中文說明)

## Quick start

1. Download `Export-WordText.ps1` from this repo (or `git clone` it).
2. Open PowerShell in that folder and unblock the downloaded script:
   ```powershell
   Unblock-File .\Export-WordText.ps1
   ```
3. Convert a folder of Word files:
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\Export-WordText.ps1 -Path .\input -OutDir .\output
   ```
   One `.md` file is written per document. The last line reports `done: N ok, 0 failed, N total`.

`-ExecutionPolicy Bypass` applies only to this one run; it doesn't change your system settings. If `Get-ExecutionPolicy -List` shows a value for `MachinePolicy` or `UserPolicy`, your organization's group policy blocks unsigned scripts; ask your IT department.

## Usage

```powershell
# one .md per Word file
.\Export-WordText.ps1 -Path .\input -OutDir .\output

# include subfolders (output mirrors the folder structure)
.\Export-WordText.ps1 -Path .\input -Recurse -OutDir .\output

# everything in one Markdown file: each document under "# <file name>",
# its own headings shifted down one level
.\Export-WordText.ps1 -Path .\input -Combine .\all.md

# plain text (table cells tab-separated) or JSON
.\Export-WordText.ps1 -Path .\input -OutDir .\output -Format Txt
.\Export-WordText.ps1 -Path .\input -Combine .\all.json -Format Json

# pipeline input
Get-ChildItem .\input\*.docx | .\Export-WordText.ps1 -OutDir .\output
```

| Parameter | Description |
|---|---|
| `-Path` | Files and/or folders. Accepts several values and pipeline input |
| `-Recurse` | Search folders recursively |
| `-OutDir` | One output file per input (same name, `.md` / `.txt`) |
| `-Combine` | Write everything into one file. With neither `-OutDir` nor `-Combine`, prints to the console |
| `-Format` | `Md` (default), `Txt`, `Json`. `Json` needs `-Combine` or console output |
| `-NoExtras` | Body only: skip headers, footers, footnotes, endnotes |
| `-Bom` | Write UTF-8 *with* BOM, which old Notepad and Excel need to display CJK text |
| `-Overwrite` | Replace existing output files (refused by default) |

**Exit codes:** `0` all succeeded · `1` some files failed (the rest were still written) · `2` no Word files found or invalid arguments.

> Through `powershell -File`, a comma list such as `-Path a.docx, b.docx` is passed as one literal string. Pass a folder or use the pipeline instead. Inside an interactive PowerShell session the comma form works.

## What gets converted

| Word | Markdown |
|---|---|
| Heading styles: Heading 1–6, Title, outline levels, and custom styles based on them | `#` … `######` |
| Bullet / numbered lists, including ones defined by list styles | `- ` / `1. `, nested 3 spaces per level |
| Tables | Pipe table: first row is the header, multi-paragraph cells joined with `<br>`, `\|` escaped, short rows padded |
| Paragraphs | Separated by a blank line; empty paragraphs dropped |
| Manual line breaks (Shift+Enter) | Hard line break |
| Text that would be read as Markdown syntax (`# `, `1. `, `- `, `>` …) | Escaped with `\` |
| Text boxes, content controls | Included (each text box once) |
| Tracked changes | Insertions kept, deletions dropped |
| Headers, footers, footnotes, endnotes | Appended after a `---` rule under **Headers** / **Footers** / **Notes** |

**Not converted:** bold/italic/underline, hyperlink targets (the link text is kept), images, charts, comments, SmartArt. Numbered lists always use `1.`; Markdown renderers renumber them.

**Tip:** heading detection relies on Word *styles*. A "heading" that is just a bigger, bold normal paragraph comes out as a normal paragraph.

## Limitations

- Password-protected or DRM-encrypted documents fail with `corrupt or not a .docx file`. Save an unencrypted copy from Word first.
- Without Word or LibreOffice, `.doc` / `.rtf` / `.odt` files are reported as failed. `.docx` always works.

## How it works

1. **Collect files.** Expands folders, filters by extension, skips Word lock files (`~$…`), de-duplicates. All path checks use `-LiteralPath`, so `[ ]` in file names isn't treated as a wildcard.
2. **Parse each file.**
   - `.docx` → `Read-Docx`: `Import-Styles` (styles.xml: heading level and list, following `basedOn`) and `Import-Numbering` (numbering.xml: bullet or ordered) → `Get-PartText` for the document, headers, footers and notes → recursive `Get-BlockLines` (`ConvertTo-MdParagraph`, `Get-MdTableLines`) → `Join-Blocks` (keeps list items and table rows on adjacent lines, blank line elsewhere).
   - Legacy formats → `Read-Legacy`: Word COM `SaveAs2`, or headless LibreOffice with a private profile and a 120 s timeout, converts to a temporary `.docx`, which is then parsed by `Read-Docx`.
3. **Write.** Every parser returns the same object `{File, Body, Headers, Footers, Notes}`; `Format-Result` renders it as per-file output, combined output, or JSON.

To support a new input format, add one `Read-Xxx` function that returns that object and one dispatch branch. Nothing on the output side changes.

## Tests

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-ExportWordText.ps1
```

The tests are self-contained (no Pester): they build `.docx` fixtures in a temp folder. Coverage includes CJK text, tables, text boxes, tracked changes, headers/footers/footnotes, `[ ]` in file names, overwrite protection, JSON, corrupt input, exit codes, and Markdown headings, lists, tables, escaping and line breaks. The `.doc` test runs only when Word or LibreOffice is available. CI runs the suite on Windows PowerShell 5.1, PowerShell 7 on Windows, and PowerShell 7 on Ubuntu.

## License

[MIT](LICENSE)

---

## 中文說明

把一整批 Word 檔轉成 **Markdown**（預設）、純文字或 JSON 的 PowerShell 腳本。`.docx` 不需要安裝 Word 或任何套件，Windows 內建的 PowerShell 5.1 就能執行。

```powershell
Unblock-File .\Export-WordText.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\Export-WordText.ps1 -Path .\input -OutDir .\output
```

- Word 的「標題 1／標題 2」等樣式會轉成 `#`／`##`，清單轉成 `-`／`1.`，表格轉成 Markdown 表格。
- 合併成一份：把 `-OutDir .\output` 換成 `-Combine .\all.md`。
- 用舊版記事本或 Excel 開啟時中文變亂碼：加上 `-Bom`。
- 有密碼保護或公司 DRM 加密的檔案無法讀取，請先用 Word 另存一份未加密的版本。
- 參數、轉換規則與限制請見上方英文說明。
