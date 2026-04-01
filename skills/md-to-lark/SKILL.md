---
name: md-to-lark
description: Use when user wants to convert a Markdown file for pasting into Lark/Feishu documents with proper formatting - tables as document tables (not Sheet), correct heading hierarchy, numbered lists, and bullet lists. Triggers on "轉成Lark格式", "粘貼到飛書", "md to lark", "lark format", "飛書文檔格式".
---

# Markdown to Lark Document Converter

## Execution

Run the converter script (same directory as this SKILL.md):

```bash
python3 ~/.claude/plugins/marketplaces/claude-skills/skills/md-to-lark/md_to_lark.py "<input_file.md>"
```

- Output: `/tmp/tf_word_clipboard.html` → auto-copied to macOS clipboard (`public.html` only)
- Prints summary with char count and execution time (<2s typical)
- After success, tell user: `已複製到剪貼簿，去 Lark 文檔 Cmd+V 粘貼即可。`

If path not found, locate: `find ~/.claude -name "md_to_lark.py" -path "*/md-to-lark/*" 2>/dev/null | head -1`

Optional second argument for custom output path: `python3 md_to_lark.py input.md /tmp/custom.html`

## What the Script Does

Generates **Microsoft Word-format HTML** (`MsoTableGrid` + Word XML namespaces) so Lark interprets pasted tables as **document tables** (Block Type 31), not Sheet/spreadsheet.

### Conversion Rules (automated by script)

- **Heading promotion**: `##` → `<h1>`, `###` → `<h2>`, `####` → `<h3>` (except `#` stays `<h1>`)
- **Auto-numbering**: `1.` `1.1.` `1.1.1.` with mandatory trailing period
- **Number stripping**: removes Western (`^\d+(\.\d+)*\.?\s+`) and Chinese ordinal (`^[一二三四五六七八九十百千]+、\s*`)
- **Tables**: MsoTableGrid class, windowtext borders, header row bold + grey `#F2F2F2` background, `&nbsp;` spacing after
- **Lists**: `<ul>/<ol>` with nesting; todo `[ ]`/`[x]` → bullet list
- **Inline**: `**bold**` → `<b>`, `*italic*` → `<i>`, `` `code` `` → `<span>` with Consolas
- **Blockquote/code block/hr**: styled HTML elements
- **Wrapper**: Word XML namespaces (`ProgId`, `Generator`), `lang="ZH-TW"`

## Known Limitations

- **Header row**: Only settable via Lark UI or Open API. Workaround: bold + grey background.
- **Auto-fit column width**: Lark ignores all width hints. Columns paste as equal width.
- **Multi-level numbering** (`N.N.`, `N.N.N.`): Only single-level `N.` on `<h2>` gets blue numbered list.
- Header row and auto-fit settable via **Lark Open API** (`docx/v1`) if user has credentials.

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
