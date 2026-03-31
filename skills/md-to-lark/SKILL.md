---
name: md-to-lark
description: Use when user wants to convert a Markdown file for pasting into Lark/Feishu documents with proper formatting - tables as document tables (not Sheet), correct heading hierarchy, numbered lists, and bullet lists. Triggers on "轉成Lark格式", "粘貼到飛書", "md to lark", "lark format", "飛書文檔格式".
---

# Markdown to Lark Document Converter

## Overview

Converts Markdown files into Lark-compatible format and copies to clipboard. The core technique is **generating Microsoft Word-format HTML** (`MsoTableGrid` class + Word XML namespaces) so Lark interprets pasted tables as **document tables** (Block Type 31), not Sheet/spreadsheet components (Block Type 30).

## Why This Exists

Lark's paste behavior:
- Plain markdown pipe tables → not recognized, renders as broken text
- Standard HTML `<table>` → **Sheet** (spreadsheet component)
- HTML with Word metadata (`MsoTableGrid`, `xmlns:w`) → **Document table** ✓
- Lark markdown import → does NOT support table syntax at all

## Format Rules (Apply to Markdown BEFORE Conversion)

Before converting to HTML, the source `.md` file must follow these rules for correct Lark rendering:

### Heading Hierarchy (for Lark TOC)

Lark renders heading levels as:
- `#` → Document title (only ONE per document)
- `##` → Main sections (large headings, appear as top-level in TOC sidebar)
- `###` → Sub-sections (blue/teal colored headings in Lark, nested in TOC)

Rules:
1. Use `#` ONLY for the document title
2. Use `##` for major chapters with sequential numbering: `## 1. Section Name`, `## 2. Section Name`
3. Use `###` for sub-sections: `### 2.1 Sub-section Name`, `### 2.2 Sub-section Name`
4. NEVER have multiple `#` headings — this flattens the TOC

### Table Format

Standard markdown pipe table with left-alignment separator:

```
| Header1 | Header2 | Header3 |
| :--- | :--- | :--- |
| data | data | data |
```

- First column header should be descriptive (e.g. `功能編號` not `#`)
- Use `:---` (with colon) for left-alignment
- No line breaks within cells — use `;` to separate items within a cell
- Wrap code/parameters in backticks: \`API_PARAM\`

### Numbered Lists

Use `1.` prefix with trailing period for auto-formatting:
```
1. First item
2. Second item
```

For hierarchical numbering in headings, use the heading level:
```
## 1. Main Section
### 1.1 Sub-section
```

### Bullet Lists

Use `- ` (dash + space) which Lark recognizes:
```
- Item one
- Item two
```

### Special Characters

- Wrap API parameters in backticks: \`30min\`, \`AF=0.02\`
- Wrap formulas in backticks: \`(O+H+L+C)/4\`
- Use `→` with spaces around arrows: `A → B`
- Use `≈` directly (Lark supports Unicode)
- Avoid `~` prefix for numbers (can trigger strikethrough) — use `約` instead

## Conversion Process

### Step 1: Parse Markdown to Word-Format HTML

Generate HTML with these critical attributes:

**Document wrapper:**
```html
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta name="ProgId" content="Word.Document">
  <meta name="Generator" content="Microsoft Word 15">
  <meta charset="utf-8">
</head>
```

**Tables — MUST use Word classes:**
```html
<table class="MsoTableGrid" border="1" cellspacing="0" cellpadding="0"
       style="border-collapse:collapse;border:none">
  <tr>
    <td style="border:solid windowtext 1.0pt;padding:3.0pt 5.0pt;background:#F2F2F2">
      <p class="MsoNormal"><b>Header</b></p>
    </td>
  </tr>
  <tr>
    <td style="border:solid windowtext 1.0pt;padding:3.0pt 5.0pt">
      <p class="MsoNormal">Data</p>
    </td>
  </tr>
</table>
```

Key attributes that trigger Lark's "Word content" detection:
- `class="MsoTableGrid"` on `<table>`
- `class="MsoNormal"` on `<p>` inside cells
- `border:solid windowtext 1.0pt` on `<td>`
- `background:#F2F2F2` on header `<td>`

**Other elements:**
- Headings: `<h1>`, `<h2>`, `<h3>` standard tags
- Blockquotes: `<p>` with left border style
- Lists: `<ul><li>` standard tags
- Code spans: `<span style="font-family:Consolas">code</span>`
- Bold: `<b>text</b>`

### Step 2: Copy to Clipboard (macOS)

Use Swift to set ONLY `public.html` type — no RTF, no plain text:

```swift
import Cocoa
let html = try! String(contentsOfFile: "/tmp/output.html", encoding: .utf8)
let pb = NSPasteboard.general
pb.clearContents()
if let data = html.data(using: .utf8) {
    pb.setData(data, forType: .html)
}
```

**Critical:** Setting only `public.html` prevents Lark from receiving extra MIME types that could trigger Sheet detection.

### Step 3: User pastes into Lark with Cmd+V

## Quick Reference Script

A shell script `md-to-lark.sh` should exist in the project directory. If not, create it with the full conversion pipeline:

1. Read `.md` file
2. Parse markdown (headings, tables, lists, blockquotes, paragraphs)
3. Generate Word-format HTML with `MsoTableGrid` tables
4. Write to temp file
5. Use Swift to copy HTML to clipboard
6. Print success message

Usage: `./md-to-lark.sh filename.md`
Then user does `Cmd+V` in Lark.

## What NOT to Do

| Approach | Result in Lark |
| :--- | :--- |
| Copy-paste raw markdown text | Tables show as broken pipe-text |
| Copy from HTML file in browser | Tables become **Sheet** (spreadsheet) |
| Copy RTF to clipboard | Tables become **Sheet** |
| Lark "Import → Markdown" | Does NOT support table syntax |
| HTML without Word metadata | Tables become **Sheet** |
| pandoc HTML with `<colgroup>` | Tables become **Sheet** |

## Troubleshooting

**Tables still become Sheet:**
- Verify HTML has `xmlns:w="urn:schemas-microsoft-com:office:word"` namespace
- Verify `<table>` has `class="MsoTableGrid"`
- Verify clipboard only has `public.html` type (check with `pbpaste -Prefer html | head`)
- Verify no `data-sheets-*` attributes in HTML

**Heading hierarchy wrong in Lark TOC:**
- Ensure only ONE `#` heading (document title)
- Main sections must be `##`, sub-sections `###`

**Bullet/numbered lists not formatted:**
- Ensure `- ` (dash space) for bullets
- Ensure blank line before list starts
