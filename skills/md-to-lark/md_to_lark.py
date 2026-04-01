#!/usr/bin/env python3
"""md_to_lark.py — Markdown to Lark-compatible Word HTML converter.

Converts Markdown into Microsoft Word-format HTML that Lark/Feishu
interprets as document tables (Block Type 31), not spreadsheets.

Usage: python3 md_to_lark.py <input.md> [output.html]
"""

import re
import sys
import subprocess
import time


class MdToLarkConverter:
    """Converts Markdown to Lark-compatible Word HTML."""

    def __init__(self):
        self.h1 = 0  # ## counter
        self.h2 = 0  # ### counter
        self.h3 = 0  # #### counter
        self.h4 = 0  # ##### counter

    # ── inline formatting ──────────────────────────────────────────

    @staticmethod
    def _inline(text):
        """Process bold, italic, inline code within text."""
        # inline code first (prevent markdown inside code spans)
        text = re.sub(r'`([^`]+)`',
                       r'<span style="font-family:Consolas;background:#f0f0f0;padding:1px 3px">\1</span>', text)
        # bold (non-greedy, handles multiple segments per line)
        text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
        # italic (single * only, after bold is gone)
        text = re.sub(r'\*(.+?)\*', r'<i>\1</i>', text)
        return text

    # ── heading ────────────────────────────────────────────────────

    @staticmethod
    def _strip_numbering(text):
        """Strip existing Western numeric and Chinese ordinal numbering."""
        text = re.sub(r'^[一二三四五六七八九十百千]+、\s*', '', text)
        text = re.sub(r'^\d+(\.\d+)*\.?\s+', '', text)
        return text

    def _heading(self, level, raw_text):
        text = self._strip_numbering(raw_text)
        text = self._inline(text)

        if level == 1:
            return f'<h1>{text}</h1>\n'
        if level == 2:
            self.h1 += 1; self.h2 = self.h3 = self.h4 = 0
            return f'<h1>{self.h1}. {text}</h1>\n'
        if level == 3:
            self.h2 += 1; self.h3 = self.h4 = 0
            return f'<h2>{self.h1}.{self.h2}. {text}</h2>\n'
        if level == 4:
            self.h3 += 1; self.h4 = 0
            return f'<h3>{self.h1}.{self.h2}.{self.h3}. {text}</h3>\n'
        if level == 5:
            self.h4 += 1
            return f'<h4>{self.h1}.{self.h2}.{self.h3}.{self.h4}. {text}</h4>\n'
        return f'<h5>{text}</h5>\n'

    # ── table ──────────────────────────────────────────────────────

    def _table(self, rows):
        """Convert parsed table rows → MsoTableGrid HTML."""
        h = ['<table class="MsoTableGrid" border="1" cellspacing="0" cellpadding="0"',
             '       style="border-collapse:collapse;border:none">']
        for ri, row in enumerate(rows):
            cells = [c.strip() for c in row.strip('|').split('|')]
            h.append('<tr>')
            for cell in cells:
                ct = self._inline(cell)
                if ri == 0:
                    h.append('  <td style="border:solid windowtext 1.0pt;padding:3.0pt 5.0pt;background:#F2F2F2">')
                    h.append(f'    <p class="MsoNormal"><b>{ct}</b></p>')
                else:
                    h.append('  <td style="border:solid windowtext 1.0pt;padding:3.0pt 5.0pt">')
                    h.append(f'    <p class="MsoNormal">{ct}</p>')
                h.append('  </td>')
            h.append('</tr>')
        h.append('</table>')
        h.append('<p class="MsoNormal">&nbsp;</p>\n')
        return '\n'.join(h)

    @staticmethod
    def _is_table_sep(line):
        """True if line is a table separator like |---|---|."""
        return bool(re.match(r'^\|[\s\-:|]+\|$', line.strip()))

    # ── lists ──────────────────────────────────────────────────────

    _LIST_RE = re.compile(r'^(\s*)([-*]|\d+\.)\s+(.+)$')

    def _collect_list(self, lines, start):
        """Collect list items from start. Returns (items, next_index).
        Each item: (indent, html_text, is_ordered)."""
        items = []
        i = start
        while i < len(lines):
            m = self._LIST_RE.match(lines[i])
            if not m:
                break
            indent = len(m.group(1))
            marker = m.group(2)
            text = m.group(3)
            is_ordered = marker[-1] == '.'
            text = re.sub(r'^\[[ x]\]\s*', '', text)  # strip todo checkbox
            items.append((indent, self._inline(text), is_ordered))
            i += 1
        return items, i

    def _render_list(self, items, idx, base):
        """Recursively render nested list. Returns (html, next_idx)."""
        if idx >= len(items):
            return '', idx
        tag = 'ol' if items[idx][2] else 'ul'
        parts = [f'<{tag}>']
        while idx < len(items):
            indent, text, _ = items[idx]
            if indent < base:
                break
            if indent > base:
                nested, idx = self._render_list(items, idx, indent)
                if parts[-1].endswith('</li>'):
                    parts[-1] = parts[-1][:-5]
                    parts.append(nested)
                    parts.append('</li>')
                else:
                    parts.append(nested)
                continue
            parts.append(f'<li>{text}</li>')
            idx += 1
        parts.append(f'</{tag}>')
        return '\n'.join(parts), idx

    # ── main converter ─────────────────────────────────────────────

    def convert(self, markdown):
        lines = markdown.split('\n')
        out = []
        i = 0
        n = len(lines)

        while i < n:
            line = lines[i]

            # ── code block ──
            if line.startswith('```'):
                code = []
                i += 1
                while i < n and not lines[i].startswith('```'):
                    code.append(lines[i])
                    i += 1
                if i < n:
                    i += 1  # skip closing ```
                raw = '\n'.join(code)
                raw = raw.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                out.append(
                    '<pre style="background:#f5f5f5;padding:12px;font-family:Consolas;'
                    'font-size:10pt;white-space:pre-wrap;border:1px solid #e0e0e0">'
                    f'<code>{raw}</code></pre>\n'
                    '<p class="MsoNormal">&nbsp;</p>\n')
                continue

            # ── heading ──
            m = re.match(r'^(#{1,6})\s+(.+)$', line)
            if m:
                out.append(self._heading(len(m.group(1)), m.group(2)))
                i += 1
                continue

            # ── table ──
            if line.startswith('|'):
                rows = []
                while i < n and lines[i].startswith('|'):
                    if not self._is_table_sep(lines[i]):
                        rows.append(lines[i])
                    i += 1
                if rows:
                    out.append(self._table(rows))
                continue

            # ── horizontal rule ──
            if re.match(r'^---+\s*$', line):
                out.append('<hr>\n')
                i += 1
                continue

            # ── blockquote ──
            if line.startswith('> '):
                text = self._inline(line[2:])
                out.append(
                    '<blockquote style="border-left:4px solid #ccc;padding:8px 16px;'
                    f'color:#555;background:#f9f9f9"><p class="MsoNormal">{text}'
                    '</p></blockquote>\n')
                i += 1
                continue

            # ── list (bullet / ordered / todo) ──
            if self._LIST_RE.match(line):
                items, i = self._collect_list(lines, i)
                html, _ = self._render_list(items, 0, 0)
                out.append(html + '\n')
                continue

            # ── empty line ──
            if not line.strip():
                i += 1
                continue

            # ── paragraph ──
            text = self._inline(line)
            out.append(f'<p class="MsoNormal">{text}</p>\n')
            i += 1

        return self._wrap('\n'.join(out))

    # ── document wrapper ───────────────────────────────────────────

    @staticmethod
    def _wrap(content):
        return (
            '<html xmlns:o="urn:schemas-microsoft-com:office:office"\n'
            '      xmlns:w="urn:schemas-microsoft-com:office:word"\n'
            '      xmlns="http://www.w3.org/TR/REC-html40">\n'
            '<head>\n'
            '  <meta name="ProgId" content="Word.Document">\n'
            '  <meta name="Generator" content="Microsoft Word 15">\n'
            '  <meta charset="utf-8">\n'
            '  <style>\n'
            '    p.MsoNormal { margin:0; font-family:Calibri; font-size:11pt; }\n'
            '    table.MsoTableGrid { border-collapse:collapse; }\n'
            '    ul, ol { margin:4px 0; padding-left:2em; }\n'
            '  </style>\n'
            '</head>\n'
            f'<body lang="ZH-TW">\n{content}\n</body>\n</html>'
        )


# ── clipboard (macOS) ──────────────────────────────────────────────

def copy_to_clipboard(path):
    subprocess.run(['swift', '-e', (
        'import Cocoa\n'
        f'let html = try! String(contentsOfFile: "{path}", encoding: .utf8)\n'
        'let pb = NSPasteboard.general\n'
        'pb.clearContents()\n'
        'if let data = html.data(using: .utf8) {\n'
        '    pb.setData(data, forType: .html)\n'
        '}'
    )], check=True)


# ── main ───────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage: python3 md_to_lark.py <input.md> [output.html]")
        sys.exit(1)

    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else '/tmp/tf_word_clipboard.html'

    t0 = time.time()

    try:
        with open(src, encoding='utf-8') as f:
            md = f.read()
    except FileNotFoundError:
        print(f"Error: {src} not found"); sys.exit(1)

    html = MdToLarkConverter().convert(md)

    with open(dst, 'w', encoding='utf-8') as f:
        f.write(html)

    copy_to_clipboard(dst)

    elapsed = time.time() - t0
    print(f"✓ {len(html):,} chars → {dst} → clipboard ({elapsed:.1f}s)")


if __name__ == '__main__':
    main()
