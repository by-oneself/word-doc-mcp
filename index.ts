#!/usr/bin/env node
/**
 * Word Document MCP Server — Professional .docx generation
 *
 * Tools: create_document | read_document | append_document
 *
 * Supports: headings, rich-text paragraphs (inline bold/italic/color),
 *           bullet lists, numbered lists, tables, dividers, page breaks.
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js'
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  BorderStyle,
} from 'docx'
import { readFile, writeFile, mkdir, access } from 'fs/promises'
import { dirname } from 'path'
import AdmZip from 'adm-zip'

// ── Helpers ──

async function fileExists(path: string): Promise<boolean> {
  try { await access(path); return true } catch { return false }
}

async function readDocxText(filePath: string): Promise<string> {
  const buffer = await readFile(filePath)
  const zip = new AdmZip(buffer) as any
  const docXml = zip.readAsText('word/document.xml')
  const lines: string[] = []
  const paras = docXml.match(/<w:p[^>]*>[\s\S]*?<\/w:p>/g) || []
  for (const p of paras) {
    const texts = p.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || []
    const line = texts.map((t: string) => t.replace(/<[^>]+>/g, '')).join('')
    if (line.trim()) lines.push(line)
  }
  return lines.join('\n')
}

// ────────────────────────────────────────────────────────────
//  Schema: rich content blocks
// ────────────────────────────────────────────────────────────

interface TextSpan {
  text: string
  bold?: boolean
  italic?: boolean
  fontSize?: number       // pt
  color?: string          // hex, e.g. "333333"
  font?: string           // "Microsoft YaHei" etc.
}

interface ContentBlock {
  type: 'heading' | 'paragraph' | 'list' | 'table' | 'pageBreak' | 'divider'
  level?: number                           // heading 1-3
  text?: string                            // plain text (single-style)
  runs?: TextSpan[]                        // rich-text (multi-style)
  alignment?: 'left' | 'center' | 'right'
  spacing?: number                         // extra spacing after, in pt
  // list
  ordered?: boolean
  items?: (string | TextSpan[])[]
  // table
  headers?: string[]
  rows?: string[][]
}

interface CreateDocArgs { filePath: string; title?: string; content: ContentBlock[] }
interface ReadDocArgs { filePath: string }
interface AppendDocArgs { filePath: string; content: ContentBlock[] }

// ────────────────────────────────────────────────────────────
//  Renderer: ContentBlock[] → (Paragraph | Table)[]
// ────────────────────────────────────────────────────────────

const FONT_FAMILY = 'Microsoft YaHei'
const BODY_SIZE = 22  // 11pt in half-points
const HEADING_COLORS: Record<number, string> = { 1: '1a1a2e', 2: '16213e', 3: '0f3460' }

function makeRun(s: TextSpan): TextRun {
  return new TextRun({
    text: s.text,
    bold: s.bold,
    italics: s.italic,
    size: (s.fontSize || 11) * 2,
    color: s.color ?? undefined,
    font: { name: s.font || FONT_FAMILY },
  })
}

function alignment(a?: string): (typeof AlignmentType)[keyof typeof AlignmentType] {
  if (a === 'center') return AlignmentType.CENTER
  if (a === 'right') return AlignmentType.RIGHT
  return AlignmentType.LEFT
}

type DocNode = Paragraph | Table

function renderContent(blocks: ContentBlock[]): DocNode[] {
  const out: DocNode[] = []

  for (const b of blocks) {
    switch (b.type) {
      // ── heading ──
      case 'heading': {
        const lvl = Math.min(Math.max(b.level || 1, 1), 3)
        const headingMap: Record<number, typeof HeadingLevel[keyof typeof HeadingLevel]> = {
          1: HeadingLevel.HEADING_1,
          2: HeadingLevel.HEADING_2,
          3: HeadingLevel.HEADING_3,
        }
        out.push(new Paragraph({
          children: [
            new TextRun({
              text: b.text || '',
              bold: true,
              size: [36, 28, 24][lvl - 1],
              color: HEADING_COLORS[lvl] || '1a1a2e',
              font: { name: FONT_FAMILY },
            }),
          ],
          heading: headingMap[lvl],
          spacing: { before: lvl === 1 ? 480 : 300, after: 160 },
        }))
        break
      }

      // ── divider (light horizontal rule) ──
      case 'divider':
        out.push(new Paragraph({
          children: [],
          border: { bottom: { style: BorderStyle.SINGLE, color: 'cccccc', size: 6, space: 8 } },
          spacing: { before: 160, after: 160 },
        }))
        break

      // ── page break ──
      case 'pageBreak':
        out.push(new Paragraph({ children: [], pageBreakBefore: true }))
        break

      // ── paragraph (plain text or rich runs) ──
      case 'paragraph': {
        const runs = b.runs?.length
          ? b.runs.map(makeRun)
          : b.text
            ? [new TextRun({ text: b.text, font: { name: FONT_FAMILY }, size: BODY_SIZE })]
            : []
        out.push(new Paragraph({
          children: runs,
          alignment: alignment(b.alignment),
          spacing: { after: b.spacing ? b.spacing * 20 : 120, line: 320 },
        }))
        break
      }

      // ── bullet / numbered list ──
      case 'list': {
        if (!b.items?.length) break
        for (let i = 0; i < b.items.length; i++) {
          const item = b.items[i]
          const runs: TextRun[] = typeof item === 'string'
            ? [new TextRun({ text: item, font: { name: FONT_FAMILY }, size: BODY_SIZE })]
            : item.map(makeRun)

          const prefix = b.ordered ? `${i + 1}.` : '•'
          out.push(new Paragraph({
            children: [
              new TextRun({
                text: `${prefix}\t`,
                bold: true,
                font: { name: FONT_FAMILY },
                size: BODY_SIZE,
                color: b.ordered ? '3b82f6' : '6b7280',
              }),
              ...runs,
            ],
            spacing: { after: 80, line: 320 },
            indent: { left: 480, hanging: 240 },
          }))
        }
        break
      }

      // ── table ──
      case 'table': {
        if (!b.headers || !b.rows) break
        const allRows = [b.headers, ...b.rows]
        const COLORS = ['f0f4ff', 'ffffff']
        out.push(new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: allRows.map((row, ri) =>
            new TableRow({
              children: row.map((cell) =>
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [new TextRun({
                        text: cell,
                        bold: ri === 0,
                        size: ri === 0 ? 22 : 20,
                        font: { name: FONT_FAMILY },
                        color: ri === 0 ? 'ffffff' : '333333',
                      })],
                      spacing: { before: 60, after: 60 },
                    }),
                  ],
                  borders: {
                    top: { style: BorderStyle.SINGLE, size: 1, color: 'd1d5db' },
                    bottom: { style: BorderStyle.SINGLE, size: 1, color: 'd1d5db' },
                    left: { style: BorderStyle.SINGLE, size: 1, color: 'd1d5db' },
                    right: { style: BorderStyle.SINGLE, size: 1, color: 'd1d5db' },
                  },
                  shading: ri === 0
                    ? { fill: '3b82f6', type: 'solid' as const }
                    : { fill: COLORS[ri % 2], type: 'solid' as const },
                })
              ),
            })
          ),
        }))
        break
      }
    }
  }
  return out
}

// ────────────────────────────────────────────────────────────
//  MCP Server
// ────────────────────────────────────────────────────────────

const server = new Server(
  { name: 'word-doc-mcp', version: '1.1.0' },
  { capabilities: { tools: {} } }
)

const contentItemSchema = {
  type: 'object',
  properties: {
    type: { type: 'string', enum: ['heading', 'paragraph', 'list', 'table', 'pageBreak', 'divider'] },
    text: { type: 'string' },
    runs: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          text: { type: 'string' },
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          fontSize: { type: 'number' },
          color: { type: 'string' },
          font: { type: 'string' },
        },
        required: ['text'],
      },
    },
    level: { type: 'number' },
    alignment: { type: 'string', enum: ['left', 'center', 'right'] },
    spacing: { type: 'number' },
    ordered: { type: 'boolean' },
    items: { type: 'array' },
    headers: { type: 'array', items: { type: 'string' } },
    rows: { type: 'array', items: { type: 'array', items: { type: 'string' } } },
  },
  required: ['type'],
}

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    {
      name: 'create_document',
      description:
        '创建 Word 文档。content 为结构化数组，每项 type: heading|paragraph|list|table|pageBreak|divider。\n' +
        'paragraph 支持 runs 富文本（数组 {text,bold?,italic?,color?,fontSize?}）；list 支持 ordered|items；table 支持 headers|rows；divider 为水平分隔线。',
      inputSchema: {
        type: 'object',
        properties: {
          filePath: { type: 'string', description: '.docx 保存路径（绝对路径）' },
          title: { type: 'string' },
          content: { type: 'array', items: contentItemSchema as any },
        },
        required: ['filePath', 'content'],
      },
    },
    {
      name: 'read_document',
      description: '读取 Word 文档的纯文本内容',
      inputSchema: {
        type: 'object',
        properties: { filePath: { type: 'string' } },
        required: ['filePath'],
      },
    },
    {
      name: 'append_document',
      description: '在已有 Word 文档末尾追加内容',
      inputSchema: {
        type: 'object',
        properties: {
          filePath: { type: 'string' },
          content: { type: 'array', items: contentItemSchema as any },
        },
        required: ['filePath', 'content'],
      },
    },
  ],
}))

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params

  try {
    switch (name) {
      case 'create_document': {
        const { filePath, content } = args as unknown as CreateDocArgs
        await mkdir(dirname(filePath), { recursive: true })

        const children = renderContent(content)
        const doc = new Document({
          styles: {
            default: {
              document: {
                run: { font: { name: FONT_FAMILY }, size: BODY_SIZE },
              },
            },
          },
          sections: [{ children }],
        })
        const buffer = await Packer.toBuffer(doc)
        await writeFile(filePath, buffer)

        return {
          content: [{ type: 'text', text: `✅ Word 文档已保存至 ${filePath}（${content.length} 个内容块）` }],
        }
      }

      case 'read_document': {
        const { filePath } = args as unknown as ReadDocArgs
        if (!(await fileExists(filePath))) {
          return { content: [{ type: 'text', text: `❌ 文件不存在: ${filePath}` }] }
        }
        const text = await readDocxText(filePath)
        return { content: [{ type: 'text', text: text || '(空文档)' }] }
      }

      case 'append_document': {
        const { filePath, content } = args as unknown as AppendDocArgs
        if (!(await fileExists(filePath))) {
          return { content: [{ type: 'text', text: `❌ 文件不存在: ${filePath}` }] }
        }
        const existingBuffer = await readFile(filePath)
        const zip = new AdmZip(existingBuffer) as any
        const oldXml = zip.readAsText('word/document.xml')
        return {
          content: [{ type: 'text', text: `⚠️ append 暂不支持复杂操作。已读取现有文档(${oldXml.length}字符)。建议将合并后的完整内容用 create_document 重新生成。` }],
        }
      }

      default:
        throw new Error(`未知工具: ${name}`)
    }
  } catch (err: any) {
    return {
      content: [{ type: 'text', text: `❌ 错误: ${err.message || String(err)}` }],
      isError: true,
    }
  }
})

async function main() {
  const transport = new StdioServerTransport()
  await server.connect(transport)
  console.error('📄 Word Document MCP Server v1.1 running (stdio)')
}

main().catch((err) => {
  console.error('Fatal:', err)
  process.exit(1)
})
