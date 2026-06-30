# word-doc-mcp

MCP Server for creating and reading Word (.docx) documents — with rich formatting support.

## Tools

| Tool | Description |
|------|-------------|
| `create_document` | Create `.docx` with headings, rich-text paragraphs, lists, tables, dividers |
| `read_document` | Read `.docx` as plain text |
| `append_document` | Append content to existing `.docx` |

## Content Types

```jsonc
{ "type": "heading", "level": 1, "text": "Title" }
{ "type": "paragraph", "text": "Plain text" }
{ "type": "paragraph", "runs": [{"text":"Bold", "bold":true}, {"text":" normal"}] }
{ "type": "list", "ordered": false, "items": ["Item 1", "Item 2"] }
{ "type": "table", "headers": ["Col A", "Col B"], "rows": [["a","b"]] }
{ "type": "divider" }
{ "type": "pageBreak" }
```

## Install

### VS Code / Copilot

```json
{
  "mcpServers": {
    "word-doc": {
      "command": "npx",
      "args": ["-y", "word-doc-mcp"]
    }
  }
}
```

### Claude Desktop

```json
{
  "mcpServers": {
    "word-doc": {
      "command": "npx",
      "args": ["-y", "word-doc-mcp"]
    }
  }
}
```

## Usage

Ask the AI:

> Create a weekly report at E:/docs/report.docx with my 3 completed features and next week's plan.

Or:

> Read the document at E:/docs/report.docx and summarize it.

## License

MIT
