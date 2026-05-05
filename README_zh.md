# office-cli

[English](./README.md) | 简体中文

面向人类与 AI Agent 的本地 Office 文档（Word、Excel、PPT、PDF）工具集。

`office-cli` 通过单一、JSON 友好的命令行，让你**读取、编辑、搜索、合并、拆分、加水印、相互转换**各类 Office 文档。**完全离线，无需登录、无遥测**——所有操作都在本机运行，使用纯 Go 库实现。

CLI 设计上有意配合一个 Skill（`skills/office-cli/SKILL.md`）一同使用，方便 AI Agent 安全可控地驱动办公文档工作流。

## 亮点

| 格式 | 命令数 | 读取 | 写入 | 核心能力 |
|------|--------|------|------|----------|
| **Excel (.xlsx)** | 45 | sheets, read, cell, search, info, cell-style, default-font | write, append, create, to-csv, from-csv, to-json, from-json, style, sort, freeze, merge/unmerge, chart, multi-chart, data-bar, icon-set, color-scale, cond-format, auto-filter, add-image, insert/delete 行列, 列宽行高, hide/show sheet, rename/copy/delete sheet, formula, column-formula, copy, copy-range, fill-range, validation, hyperlink | 全面掌控电子表格 |
| **Word (.docx)** | 25 | read, meta, stats, headings, images, search | create, replace, add-paragraph, add-heading, add-table, add-image, add-page-break, delete, insert-before/after, update-table-cell, update-table, add/delete table rows/cols, merge, style, style-table | 创建与编辑文档 |
| **PowerPoint (.pptx)** | 17 | read, count, outline, meta, images, layout | create, replace, add-slide, set-content, set-notes, delete-slide, reorder, add-image, build, set-style, add-shape | 创建与管理演示文稿 |
| **PDF** | 22 | read, pages, info, bookmarks, extract-images, search | merge, split, trim, watermark, stamp-image, rotate, optimize, reorder, insert-blank, add-bookmarks, set-meta, encrypt, decrypt, create, replace, add-text | PDF 完整生命周期 |
| **通用** | 3 | info, meta, extract-text | — | 自动识别格式 |

**总计：115 条原子命令**——不做编排，每条命令只做一件事，AI Agent 自由组合。

### AI 友好设计

| 特性 | 说明 |
|------|------|
| `--json` | 机器可读扁平 JSON 输出（stdout），错误输出到 stderr |
| `--quiet` | 抑制非 JSON 输出，适合管道组合 |
| `--dry-run` | 预览写操作，不触及文件系统 |
| `--force` | 跳过交互确认，用于 CI/Agent 自动化 |
| `errorCode` + `hint` | 结构化错误响应，稳定可解析 |
| 扁平 Schema | 节省 Token：`FlatSheet`、`FlatPage`、`FlatSlide`、`FlatParagraph` |
| `office-cli reference` | 全部命令 + flag 输出为 Markdown，Agent 自主发现 |
| 审计日志 | 每条写命令记录到 `~/.office-cli/audit/audit-YYYY-MM.jsonl` |

## 安装

### 通过 npm 安装（推荐）

```bash
# 安装 CLI 二进制
npm install -g @fatecannotbealtered-/office-cli

# 安装 AI Agent Skill（适用于 Cursor / Claude Code / OpenClaw 等）
npx skills add fatecannotbealtered/office-cli -y -g
```

post-install 脚本会下载与你 OS/架构匹配的预编译二进制，并校验 SHA-256。

### 从源码构建

```bash
git clone https://github.com/fatecannotbealtered/office-cli.git
cd office-cli
make build
./bin/office-cli doctor
```

## 快速上手

```bash
# 首次使用（创建 ~/.office-cli/config.json）
office-cli setup

# 通用 —— 根据扩展名自动分发
office-cli info  any.docx --json          # 一行概要
office-cli meta  any.pdf  --json          # 完整元数据
office-cli extract-text any.pptx --json   # 提取任意格式的纯文本

# Excel —— 45 条命令，全面掌控电子表格
office-cli excel sheets sales.xlsx --json
office-cli excel read sales.xlsx --range A1:D10 --json
office-cli excel cell sales.xlsx B2 --sheet Sales --json
office-cli excel append sales.xlsx --rows '[["张三", 100], ["李四", 200]]'
office-cli excel from-csv data.csv --output data.xlsx
office-cli excel to-csv sales.xlsx --sheet Sales --output sales.csv --bom
office-cli excel rename-sheet workbook.xlsx --from 旧表 --to 新表
office-cli excel copy-sheet workbook.xlsx --from 模板 --to 备份
office-cli excel style sales.xlsx --sheet Sales --range A1:D1 --bold --bg-color "4472C4" --text-color "FFFFFF"
office-cli excel sort sales.xlsx --sheet Sales --range A1:D100 --by-col 2 --ascending
office-cli excel freeze sales.xlsx --sheet Sales --cell A2
office-cli excel chart sales.xlsx --sheet Sales --cell E2 --type bar --data-range "A1:B10" --title "收入"

# Word —— 创建、写入、读取
office-cli word create 报告.docx --title "季度报告" --author "张三"
office-cli word add-heading 报告.docx --text "引言" --level 1 --output 报告.docx
office-cli word add-paragraph 报告.docx --text "正文内容..." --output 报告.docx
office-cli word add-table 报告.docx --rows '[["姓名","分数"],["张三",95]]' --output 报告.docx
office-cli word read 报告.docx --format markdown
office-cli word replace 模板.docx --find "{{姓名}}" --replace "张三"

# PowerPoint —— 创建、写入、读取
office-cli ppt create 演示.pptx --title "季度汇报" --author "张三"
office-cli ppt add-slide 演示.pptx --title "路线图" --bullets '["Q1: 完成","Q2: 进行中"]' --output 演示.pptx
office-cli ppt set-notes 演示.pptx --slide 1 --notes "演讲者备注" --output 演示.pptx
office-cli ppt delete-slide 演示.pptx --slide 3 --output 演示.pptx
office-cli ppt reorder 演示.pptx --order "3,1,2" --output 演示.pptx
office-cli ppt outline 演示.pptx --json
office-cli ppt build 演示.pptx --spec '{"title":"Q2 Review","author":"Alice","slides":[{"title":"Overview","bullets":["Revenue up 20%"]}]}'
office-cli ppt build 演示.pptx --template company.pptx --spec '{"title":"Q2 Report","slides":[{"title":"Overview"},{"title":"Details"}]}'
office-cli ppt build 演示.pptx --spec '{"slides":[{"title":"Chart","image":"chart.png","width":6000000,"height":4000000},{"title":"Summary"}]}'

# PDF —— 22 条命令，完整生命周期
office-cli pdf info 报告.pdf --json
office-cli pdf merge a.pdf b.pdf --output combined.pdf
office-cli pdf trim 大文件.pdf --pages "1,3,5-7" --output excerpt.pdf
office-cli pdf rotate 报告.pdf --degrees 90 --pages 1 --output rotated.pdf
office-cli pdf watermark 报告.pdf --text "草稿" --output 报告.draft.pdf
office-cli pdf encrypt 报告.pdf --user-password 1234 --output secret.pdf
office-cli pdf optimize 报告.pdf --output smaller.pdf
office-cli pdf reorder 报告.pdf --order "3,1,2,4" --output reordered.pdf
office-cli pdf insert-blank 报告.pdf --after 2 --count 1 --output 报告.with-blanks.pdf
office-cli pdf bookmarks 报告.pdf --json
office-cli pdf set-meta 报告.pdf --title "年度报告" --author "张三" --output 报告.meta.pdf
```

## 中文与跨平台

- **路径**：所有命令都对路径调用 `filepath.Clean / Abs` 标准化，支持 `~/`、相对路径、含空格、含 CJK 字符的路径。
- **CSV 编码**：`excel to-csv` 默认带 UTF-8 BOM（`--bom=true`），Windows Excel 双击打开就不会变乱码；导入时使用 `excel from-csv`，会自动剥离 BOM。
- **JSON 输出**：始终为 UTF-8，没有 BOM；中文字符直接写入而非转义为 `\uXXXX`，便于阅读。
- **测试**：`scripts/gen-fixtures` 会生成包含中文文件名（`样本.xlsx`）和中文内容的真实文档；`scripts/e2e-all-commands.{ps1,sh}` 会跑遍所有命令做端到端验证。

## 项目结构

```
office-cli/
├── cmd/
│   └── office-cli/
│       └── main.go               # 入口
│   ├── root.go                   # 全局 flag、退出码、审计日志钩子
│   ├── setup.go                  # 配置文件创建
│   ├── doctor.go                 # 环境/工具检测
│   ├── reference.go              # 全部命令输出为结构化 Markdown
│   ├── install_skill.go          # 复制 skill 文件到 ~/.agent/skills/
│   ├── universal.go              # info / meta / extract-text（自动识别格式）
│   ├── helpers.go                # resolveInput / resolveOutput / 路径校验
│   ├── excel.go                  # 45 条命令：读/写/样式/排序/图表/...
│   ├── word.go                   # 25 条命令：读/替换/创建/添加内容/样式
│   ├── ppt.go                    # 17 条命令：读/替换/创建/添加内容/build/layout
│   ├── ppt_write.go              # PPT 写入命令实现
│   ├── pdf.go                    # 22 条命令：读/合并/拆分/书签/...
│   ├── e2e_test.go               # E2E 测试（原有 44 条命令）
│   └── e2e_new_test.go           # E2E 测试（新增 36 条命令）
├── internal/
│   ├── output/
│   │   ├── output.go             # 终端表格、JSON 输出
│   │   ├── errors.go             # 结构化错误码（errorCode + hint）
│   │   └── flat.go               # FlatSheet / FlatPage / FlatSlide / FlatParagraph
│   ├── audit/
│   │   └── audit.go              # JSONL 写命令审计日志
│   ├── config/
│   │   └── config.go             # ~/.office-cli 路径解析
│   └── engine/
│       ├── common/
│       │   └── zipxml.go         # RewriteEntries / WriteNewZip / ReadEntry / CoreProps / AppProps
│       ├── excel/
│       │   └── excel.go          # excelize/v2 —— 30 个函数 + 4 个类型
│       ├── word/
│       │   └── word.go           # 原生 ZIP+XML docx —— 8 个写入函数
│       ├── ppt/
│       │   └── ppt.go            # 原生 ZIP+XML pptx —— 8 个写入函数
│       └── pdf/
│           └── pdf.go            # pdfcpu + ledongthuc/pdf —— 6 个新函数
├── skills/
│   └── office-cli/
│       ├── SKILL.md              # AI Agent 入口（路由 + 约定）
│       └── docs/
│           ├── excel.md          # 45 条 Excel 命令参考
│           ├── word.md           # 25 条 Word 命令参考
│           ├── ppt.md            # 17 条 PPT 命令参考
│           └── pdf.md            # 22 条 PDF 命令参考
├── scripts/
│   ├── install.js                # npm postinstall —— 从 GitHub Release 下载二进制
│   ├── run.js                    # npm bin 包装器 —— 执行二进制
│   └── gen-fixtures/             # 测试用文档生成器（xlsx/docx/pptx/pdf）
├── .github/
│   └── workflows/
│       ├── ci.yml                # 测试矩阵：3 OS × Go 版本
│       └── release.yml           # goreleaser + v* tag 触发 npm 发布
├── package.json                  # npm 包（@fatecannotbealtered-/office-cli）
├── go.mod / go.sum               # Go 模块依赖
├── Makefile                      # build / test / vet / fmt / snapshot
├── CHANGELOG.md
├── CONTRIBUTING.md
├── SECURITY.md
├── LICENSE                       # MIT
└── .goreleaser.yml               # 交叉编译 + 归档 + 校验和配置
```

### 为什么选纯 Go（不依赖 pandoc / libreoffice）

- 跨机器行为可预测，没有外部工具版本错配的问题。
- 依赖面更小，更易审计。
- 四类核心格式都有成熟 Go 库支撑：
  - **Excel**：`excelize/v2`，事实标准。
  - **PDF**：`pdfcpu` 负责结构操作；`ledongthuc/pdf` 负责文本提取。
  - **Word/PPT**：docx 和 pptx 本质上是 ZIP + XML，我们直接读写。

`office-cli doctor` 会报告系统是否安装了 `pandoc` 或 `libreoffice`——它们用于解锁更复杂的格式转换，但**所有上文文档化的命令都不需要它们**。

## 设计原则

1. **AI 优先**：每条命令都支持 `--json`；错误带稳定的 `errorCode` 和 `hint`；读类命令支持 `--range`/`--keyword`/`--limit` 过滤，便于控制 token 消耗。
2. **第一参数永远是文件路径**；输出参数显式（`--output`）。
3. **默认非破坏**：PDF 写入命令（`trim/merge/split/watermark/reorder/insert-blank` 等）必须指定 `--output`，绝不修改源文件。Word/PPT 写入命令默认原地修改，便于迭代编辑；传 `--output` 可写入新文件。`excel write` / `excel append` 及 Excel 格式化命令也是原地修改（Excel 的自然形态）。
4. **写操作都可审计**：日志写到 `~/.office-cli/audit/`。
5. **依赖隔离**：每种格式的 engine 都封装在 `internal/engine/<format>` 下，将来要换底层库是一天的工作量。

## 权限系统

office-cli 默认为 `read-only`。写入和破坏性命令需要在 `~/.office-cli/config.json` 中提升权限：

```json
{ "permissions": { "mode": "write" } }
```

| 级别 | 允许的操作 |
|---|---|
| `read-only`（默认） | 所有读取命令 |
| `write` | + 写入单元格、追加行、替换文本、样式/排序/冻结/合并、添加图表/图片、创建 Word/PPT 文档、添加幻灯片/段落/表格、PDF 合并/拆分/裁剪/重排/水印/印章/元数据/书签 |
| `full` | + 删除工作表、加密/解密 PDF |

环境变量覆盖：`OFFICE_CLI_PERMISSIONS=write`。超出配置级别的命令会以退出码 5 终止。

## 环境变量

| 变量 | 默认值 | 作用 |
|---|---|---|
| `OFFICE_CLI_HOME` | `~/.office-cli` | 自定义家目录 |
| `OFFICE_CLI_PERMISSIONS` | `read-only` | 覆盖权限级别（`read-only`、`write`、`full`） |
| `OFFICE_CLI_NO_AUDIT` | unset | 设为 `1` 关闭审计日志 |
| `OFFICE_CLI_AUDIT_RETENTION_MONTHS` | `3` | 自动清理超过 N 个月的审计日志（`0`=保留全部） |
| `NO_COLOR` | unset | 设任何值即可关闭彩色输出 |

## 开发

```bash
make build       # 编译到 ./bin/office-cli
make test        # go test ./...
make vet         # go vet
make fmt         # 检查 gofmt
make snapshot    # 本地 goreleaser 试跑

# 端到端测试（生成 fixtures 后跑遍所有命令）
go run ./scripts/gen-fixtures --out tmp/fixtures
pwsh ./scripts/e2e-all-commands.ps1   # Windows
bash  ./scripts/e2e-all-commands.sh   # Linux / macOS
```

## 许可证

MIT，详见 `LICENSE`。
