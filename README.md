# WordCreate

A Claude Code skill for structured .docx document modification — replace text, rewrite sections, and insert content while preserving all images, formatting, and table structures intact.

## What it does

- Replace specific text fields (names, dates, IDs, company info) across tables
- Insert or rewrite multi-paragraph content while matching the original paragraph count
- Add text between existing images without touching the images themselves
- Automate batch document template filling

## Why this exists

Built from real-world experience modifying complex Word documents programmatically. The core challenge isn't editing text — it's doing so without corrupting images, breaking table layouts, or destroying paragraph structure. This skill encodes the workflow that reliably avoids those pitfalls.

## How it works

1. **Reconnaissance** — Deep-scan every table cell: paragraph count, image presence, text length
2. **Isolate** — Mark image-containing paragraphs as read-only
3. **Execute** — Modify only text paragraphs, preserving paragraph count and structure
4. **Verify** — Re-scan output to confirm image count, paragraph structure, and field values

## Key design decisions

| Decision | Rationale |
|---|---|
| UTF-8 text files for non-ASCII content | Avoids encoding chain corruption across CLI → Python boundaries |
| PowerShell `Out-File -Encoding utf8` for scripts | Prevents encoding artifacts from intermediate tooling |
| Temp file → verify → rename workflow | Guards against file-lock errors when output is open in Word |
| Paragraph-level replacement only | Prevents accidental destruction of inline images and drawings |

## Requirements

- Python 3.x with `python-docx`
- PowerShell (Windows)
- Claude Code

## Usage

Trigger the skill in Claude Code by requesting any .docx text modification:

```
> 帮我把这份报告里的公司名从"A公司"改成"B公司"，其他的不要动
> 把这份文档的心得部分重写一下，图片和排版不要动
> 把这批文档的日期字段从2025改成2026
```

The skill handles the document inspection, modification, and verification automatically.

## License

MIT
