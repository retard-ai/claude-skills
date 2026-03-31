# md-to-lark

A Claude Code skill that converts Markdown files into Lark/Feishu-compatible format and copies to clipboard. Tables paste as **document tables** (not Sheet/spreadsheet).

## The Problem

Pasting Markdown into Lark doesn't work:
- Pipe tables (`| col | col |`) → broken text
- Standard HTML `<table>` → becomes Sheet (spreadsheet component)
- Lark "Import → Markdown" → doesn't support tables at all

## The Solution

This skill generates **Microsoft Word-format HTML** (`MsoTableGrid` + Word XML namespaces). Lark recognizes this as Word content and creates proper document tables.

## Install

```bash
/plugin marketplace add retard-ai/claude-skills
/plugin install md-to-lark@claude-skills
/reload-plugins
```

## Usage

```bash
/md-to-lark:md-to-lark /path/to/your-file.md
```

Or say in conversation: `幫我將 XXX.md 轉成 Lark 格式`

Then **Cmd+V** in your Lark document.

## What It Does

- **Tables** → Word `MsoTableGrid` format → Lark document tables
- **Headings** → Promoted by one level (`##`→`h1`, `###`→`h2`, `####`→`h3`) with auto-numbering (trailing period: `1.` `1.1.` `1.1.1.`)
- **Bullet lists** → `<ul><li>`
- **Numbered lists** → `<ol><li>`
- **Todo lists** (`- [ ]`) → bullet list (checkbox stripped)
- **Code blocks** → `<pre><code>` with background
- **Inline code** → `<span>` with Consolas font
- **Blockquotes** → styled `<blockquote>`
- **Bold / Italic** → `<b>` / `<i>`

## Known Limitations

| Feature | Status |
| --- | --- |
| Document tables | ✅ Works |
| Heading hierarchy + numbering | ✅ Works |
| Lists (bullet/numbered/todo) | ✅ Works |
| Code blocks | ✅ Works |
| Blockquotes | ✅ Works |
| Header row ("Set as header row") | ❌ Lark UI only |
| Auto-fit column width | ❌ Lark ignores CSS width |
| Multi-level numbered heading color | ❌ Only `N.` gets blue; `N.N.` stays black |

## Requirements

- macOS (uses Swift + NSPasteboard for clipboard)
- Claude Code

## License

MIT

---

# md-to-lark（繁體中文）

一個 Claude Code skill，將 Markdown 文件轉換為 Lark/飛書 相容格式並複製到剪貼簿。表格會粘貼為**文檔表格**（非 Sheet 電子表格）。

## 問題

直接將 Markdown 粘貼到 Lark 唔得：
- Pipe 表格（`| col | col |`）→ 變成亂碼文字
- 標準 HTML `<table>` → 變成 Sheet（電子表格組件）
- Lark「導入 → Markdown」→ 完全唔支持表格

## 解決方案

呢個 skill 生成 **Microsoft Word 格式 HTML**（`MsoTableGrid` + Word XML 命名空間），令 Lark 將內容識別為 Word 文檔，自動創建文檔表格。

## 安裝

```bash
/plugin marketplace add retard-ai/claude-skills
/plugin install md-to-lark@claude-skills
/reload-plugins
```

## 使用方式

```bash
/md-to-lark:md-to-lark /path/to/your-file.md
```

或者喺對話中講：`幫我將 XXX.md 轉成 Lark 格式`

然後喺 Lark 文檔 **Cmd+V** 粘貼。

## 功能

- **表格** → Word `MsoTableGrid` 格式 → Lark 文檔表格
- **標題** → 向上提升一級（`##`→`h1`、`###`→`h2`、`####`→`h3`），自動編號帶尾隨句號（`1.` `1.1.` `1.1.1.`）
- **點列式** → `<ul><li>`
- **數字列表** → `<ol><li>`
- **待辦事項**（`- [ ]`）→ 純 bullet list（去除 checkbox）
- **代碼塊** → `<pre><code>` 帶背景色
- **行內代碼** → `<span>` Consolas 字體
- **引用區塊** → 帶左邊線嘅 `<blockquote>`
- **粗體 / 斜體** → `<b>` / `<i>`

## 已知限制

| 功能 | 狀態 |
| --- | --- |
| 文檔表格 | ✅ 正常 |
| 標題層級 + 編號 | ✅ 正常 |
| 列表（bullet/數字/待辦） | ✅ 正常 |
| 代碼塊 | ✅ 正常 |
| 引用區塊 | ✅ 正常 |
| Header row（設為標題行） | ❌ 只能喺 Lark UI 手動設置 |
| 自動調整欄寬 | ❌ Lark 忽略 CSS 寬度 |
| 多層編號藍色格式 | ❌ 只有 `N.` 會變藍色；`N.N.` 維持黑色 |

## 系統需求

- macOS（使用 Swift + NSPasteboard 操作剪貼簿）
- Claude Code

---

# md-to-lark（简体中文）

一个 Claude Code skill，将 Markdown 文件转换为 Lark/飞书兼容格式并复制到剪贴板。表格会粘贴为**文档表格**（非 Sheet 电子表格）。

## 问题

直接将 Markdown 粘贴到飞书不行：
- Pipe 表格（`| col | col |`）→ 变成乱码文字
- 标准 HTML `<table>` → 变成 Sheet（电子表格组件）
- 飞书「导入 → Markdown」→ 完全不支持表格

## 解决方案

这个 skill 生成 **Microsoft Word 格式 HTML**（`MsoTableGrid` + Word XML 命名空间），让飞书将内容识别为 Word 文档，自动创建文档表格。

## 安装

```bash
/plugin marketplace add retard-ai/claude-skills
/plugin install md-to-lark@claude-skills
/reload-plugins
```

## 使用方式

```bash
/md-to-lark:md-to-lark /path/to/your-file.md
```

或者在对话中说：`帮我将 XXX.md 转成飞书格式`

然后在飞书文档 **Cmd+V** 粘贴。

## 功能

- **表格** → Word `MsoTableGrid` 格式 → 飞书文档表格
- **标题** → 向上提升一级（`##`→`h1`、`###`→`h2`、`####`→`h3`），自动编号带尾随句号（`1.` `1.1.` `1.1.1.`）
- **无序列表** → `<ul><li>`
- **有序列表** → `<ol><li>`
- **待办事项**（`- [ ]`）→ 纯 bullet list（去除 checkbox）
- **代码块** → `<pre><code>` 带背景色
- **行内代码** → `<span>` Consolas 字体
- **引用块** → 带左边线的 `<blockquote>`
- **粗体 / 斜体** → `<b>` / `<i>`

## 已知限制

| 功能 | 状态 |
| --- | --- |
| 文档表格 | ✅ 正常 |
| 标题层级 + 编号 | ✅ 正常 |
| 列表（bullet/数字/待办） | ✅ 正常 |
| 代码块 | ✅ 正常 |
| 引用块 | ✅ 正常 |
| Header row（设为标题行） | ❌ 只能在飞书 UI 手动设置 |
| 自动调整列宽 | ❌ 飞书忽略 CSS 宽度 |
| 多层编号蓝色格式 | ❌ 只有 `N.` 会变蓝色；`N.N.` 保持黑色 |

## 系统要求

- macOS（使用 Swift + NSPasteboard 操作剪贴板）
- Claude Code

## 许可证

MIT
