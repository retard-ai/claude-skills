---
name: md-to-lark
description: Use when user wants to convert a Markdown file for pasting into Lark/Feishu documents with proper formatting - tables as document tables (not Sheet), correct heading hierarchy, numbered lists, and bullet lists. Triggers on "轉成Lark格式", "粘貼到飛書", "md to lark", "lark format", "飛書文檔格式".
---

# Markdown to Lark Document Converter

## Overview

Converts Markdown files into Lark-compatible format and copies to clipboard. The core technique is generating **Microsoft Word-format HTML** (`MsoTableGrid` class + Word XML namespaces) so Lark interprets pasted tables as **document tables** (Block Type 31), not Sheet/spreadsheet components (Block Type 30).

## Conversion Rules

### Heading Promotion + Auto-Numbering

ALL headings promote up one level EXCEPT `#` which stays as `#`. Auto-numbering is applied after promotion:

| Source Markdown | Output HTML | Numbering |
| :--- | :--- | :--- |
| `#` Title | `<h1>` | No numbering (document title) |
| `## Section` | `<h1>` | `1.` `2.` `3.` (trailing period) |
| `### Sub-section` | `<h2>` | `1.1.` `1.2.` `2.1.` (trailing period) |
| `#### Sub-sub` | `<h3>` | `1.1.1.` `1.1.2.` (trailing period) |
| `##### Deep` | `<h4>` | `1.1.1.1.` (trailing period) |
| `######` and beyond | `<h5>+` | Continue same pattern with trailing period |

ALL numbered headings MUST end with a trailing period (e.g. `1.1.` not `1.1`) because Lark only auto-formats numbered lists when the number ends with `.` followed by a space.

If the source heading already has numbering (e.g. `### 1.1 Name`), strip the existing number before applying auto-numbering. Use `strip_existing_number()` regex: `^\d+(\.\d+)*\.?\s+`

### Tables → Word MsoTableGrid

Every table MUST use Word-compatible HTML:

```html
<table class="MsoTableGrid" border="1" cellspacing="0" cellpadding="0"
       style="border-collapse:collapse;border:none">
<tr>
  <td style="border:solid windowtext 1.0pt;padding:3.0pt 5.0pt;background:#F2F2F2">
    <p class="MsoNormal"><b>Header text</b></p>
  </td>
</tr>
<tr>
  <td style="border:solid windowtext 1.0pt;padding:3.0pt 5.0pt">
    <p class="MsoNormal">Data text</p>
  </td>
</tr>
</table>
<p class="MsoNormal">&nbsp;</p>
```

Key attributes for Lark's "Word content" detection:
- `class="MsoTableGrid"` on `<table>`
- `class="MsoNormal"` on `<p>` inside cells
- `border:solid windowtext 1.0pt` on `<td>`
- `background:#F2F2F2` on header row `<td>` (visual styling only; Lark native header row CANNOT be set via paste)
- Add `<p class="MsoNormal">&nbsp;</p>` after each table for spacing

### Lists

- **Bullet list** (`- item`): `<ul><li>text</li></ul>`
- **Numbered list** (`1. item`): `<ol><li>text</li></ol>`
- **Todo list** (`- [ ] item` / `- [x] item`): Treat as regular bullet list `<ul><li>text</li></ul>` — strip the `[ ]`/`[x]` checkbox marker. Lark ignores `list-style:none` CSS so using ☐/☑ with `<ul>` creates double markers.

### Other Elements

- **Blockquote** (`> text`): `<blockquote style="border-left:4px solid #ccc;padding:8px 16px;color:#555;background:#f9f9f9"><p class="MsoNormal">text</p></blockquote>`
- **Code block** (` ``` `): `<pre style="background:#f5f5f5;padding:12px;font-family:Consolas;font-size:10pt;white-space:pre-wrap;border:1px solid #e0e0e0"><code>text</code></pre>`
- **Inline code** (`` `code` ``): `<span style="font-family:Consolas;background:#f0f0f0;padding:1px 3px">code</span>`
- **Bold** (`**text**`): `<b>text</b>`
- **Italic** (`*text*`): `<i>text</i>`
- **Horizontal rule** (`---`): `<hr>`
- **Paragraph**: `<p class="MsoNormal">text</p>`

### Document Wrapper

Every generated HTML MUST be wrapped in:

```html
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta name="ProgId" content="Word.Document">
  <meta name="Generator" content="Microsoft Word 15">
  <meta charset="utf-8">
  <style>
    p.MsoNormal { margin:0; font-family:Calibri; font-size:11pt; }
    table.MsoTableGrid { border-collapse:collapse; }
    ul, ol { margin:4px 0; padding-left:2em; }
  </style>
</head>
<body lang="ZH-TW">
  ... content ...
</body>
</html>
```

## Clipboard (macOS only)

Use Swift to set ONLY `public.html` — no RTF, no plain text:

```swift
import Cocoa
let html = try! String(contentsOfFile: "/tmp/tf_word_clipboard.html", encoding: .utf8)
let pb = NSPasteboard.general
pb.clearContents()
if let data = html.data(using: .utf8) {
    pb.setData(data, forType: .html)
}
```

Setting only `public.html` prevents Lark from receiving extra MIME types that trigger Sheet detection.

After clipboard is set, tell user: `已複製到剪貼簿，去 Lark 文檔 Cmd+V 粘貼即可。`

## Known Limitations

These Lark-native properties CANNOT be set via HTML paste:
- **Header row** (Lark's "Set as header row"): Only settable via Lark UI or Open API. Visual workaround: bold + background color on first row.
- **Auto-fit column width**: Lark ignores all CSS/HTML width hints. All columns paste as equal width.
- Both can be set via **Lark Open API** (`docx/v1` endpoint) if the user has API credentials.

## What NOT to Do

| Approach | Result in Lark |
| :--- | :--- |
| Copy-paste raw markdown | Broken pipe-text |
| Copy from HTML file in browser | **Sheet** (spreadsheet) |
| Copy RTF to clipboard | **Sheet** |
| Lark "Import → Markdown" | Does NOT support tables |
| HTML without Word namespace | **Sheet** |
| pandoc HTML with `<colgroup>` | **Sheet** |
| `<thead>`/`<th>` tags | Ignored by Lark |
| `<col width>` or CSS width | Ignored by Lark |
