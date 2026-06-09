---
name: pptx-translator
description: "Translate Office documents (.pptx, .docx) between languages using Amazon Translate, Bedrock LLM, Google Translate, or Cerebras LLM. Use when: (1) translating presentations or Word documents to other languages, (2) localizing content for international audiences, (3) converting Chinese/Japanese/Korean documents to English or vice versa. Supports four engines: Amazon Translate (fast, cost-effective), Bedrock LLM (high quality), Google Translate (widely supported), and Cerebras LLM (open-source models). Preserves formatting, styles, and layout."
---

# Office Document Translator

Translate PowerPoint (.pptx) and Word (.docx) files using AWS services and LLMs.

## Quick Start

```bash
# Unified entry (auto-detects format)
python scripts/translate.py input.pptx -s zh -t en
python scripts/translate.py input.docx -s zh -t en -e cerebras

# Or use format-specific scripts directly:

# PowerPoint — Amazon Translate (fast, cheap)
python scripts/translate_pptx.py input.pptx -s zh -t en

# PowerPoint — Bedrock LLM (high quality)
python scripts/translate_pptx.py input.pptx -s zh -t en -e bedrock

# Word — Cerebras LLM (open-source models, key via api_keys.txt or --cerebras-api-key)
python scripts/translate_docx.py input.docx -s zh -t en -e cerebras

# Word — Google Translate (widely supported)
python scripts/translate_docx.py input.docx -s zh -t en -e google --google-api-key YOUR_API_KEY

# Word — Bedrock LLM (high quality)
python scripts/translate_docx.py input.docx -s zh -t en -e bedrock
```

## Engine Selection

| Use Case | Engine | Flag |
|----------|--------|------|
| Technical docs, high volume | Amazon Translate | `-e translate` (default) |
| Marketing, nuanced content | Bedrock LLM | `-e bedrock` |
| General purpose, wide language support | Google Translate | `-e google` |
| Open-source models, cost-effective | Cerebras LLM | `-e cerebras` |
| Cost-sensitive | Amazon Translate | |
| Quality-sensitive | Bedrock LLM | |
| Google Cloud users | Google Translate | |
| Open-source enthusiasts | Cerebras LLM | |

## Commands

All commands work for both `.pptx` and `.docx` — just change the script name and file extension:

### Translate with Amazon Translate
```bash
# PowerPoint
python scripts/translate_pptx.py input.pptx -s en -t zh --engine translate
# Word
python scripts/translate_docx.py input.docx -s en -t zh --engine translate
```

### Translate with Bedrock LLM
```bash
# PowerPoint
python scripts/translate_pptx.py input.pptx -s en -t zh -e bedrock \
  --model anthropic.claude-3-5-sonnet-20241022-v2:0 \
  --glossary glossary.json --style professional
# Word
python scripts/translate_docx.py input.docx -s en -t zh -e bedrock \
  --model anthropic.claude-3-5-sonnet-20241022-v2:0 \
  --glossary glossary.json --style professional
```

### Translate with Google Translate
```bash
python scripts/translate_pptx.py input.pptx -s en -t zh -e google --google-api-key KEY
python scripts/translate_docx.py input.docx -s en -t zh -e google --google-api-key KEY
```

### Translate with Cerebras LLM
```bash
# 将 API key 放入 scripts/api_keys.txt（复制自 api_keys.example.txt），或通过参数传入
python scripts/translate_pptx.py input.pptx -s en -t zh -e cerebras --style professional
python scripts/translate_docx.py input.docx -s en -t zh -e cerebras --style professional
# 也可直接传 key:
python scripts/translate_pptx.py input.pptx -s en -t zh -e cerebras --cerebras-api-key YOUR_KEY
```

### Extract texts for review
```bash
python scripts/extract_texts.py input.pptx -o texts.json
```

## Options

| Option | Description |
|--------|-------------|
| `-s, --source` | Source language code (required) |
| `-t, --target` | Target language code (required) |
| `-o, --output` | Output file path |
| `-e, --engine` | `translate` (Amazon), `bedrock` (LLM), `google`, or `cerebras` |
| `-m, --model` | Bedrock/Cerebras model ID |
| `--terminology` | CSV file for Amazon Translate |
| `-g, --glossary` | JSON/text glossary for LLM |
| `--style` | Translation style (professional, casual, technical) |
| `--batch-size` | Texts per LLM batch (default: 5) |
| `--no-batch` | Disable batch mode for LLM |
| `--region` | AWS region |
| `--google-api-key` | Google Cloud API key |
| `--google-project-id` | Google Cloud project ID (optional) |
| `--cerebras-api-key` | Cerebras API key |
| `--cerebras-model` | Cerebras model ID (default: gpt-oss-120b) |
| `--cerebras-base-url` | Cerebras API base URL (default: https://api.cerebras.ai/v1) |

## Language Codes

Common codes: `en`, `zh`, `zh-TW`, `ja`, `ko`, `de`, `fr`, `es`, `pt`, `it`, `ru`

See [references/language-codes.md](references/language-codes.md) for full list.

## Terminology/Glossary

- **Amazon Translate**: Use `--terminology` with CSV file
- **Bedrock LLM**: Use `--glossary` with JSON or key=value file

See [references/glossary-format.md](references/glossary-format.md) for format details.

## Dependencies

```bash
# Core (always required)
pip install boto3 python-pptx

# For Word document support
pip install python-docx

# For Google Translate support (optional)
pip install google-cloud-translate

# For Cerebras support (optional)
pip install openai httpx
```

AWS credentials must be configured for Amazon Translate and Bedrock (`~/.aws/credentials` or environment variables).  

**API Key 配置**: 将你的 key 写入 `scripts/api_keys.txt` 文件即可，一行一个：

```bash
cp scripts/api_keys.example.txt scripts/api_keys.txt
# 编辑 api_keys.txt，填入真实的 key
```

更换 key 时只需编辑该文件，无需改代码。

## Examples

### PowerPoint: Chinese → English
```bash
python scripts/translate_pptx.py deck-cn.pptx -s zh -t en -e bedrock
# Output: deck-cn-en.pptx
```

### Word: Chinese → English
```bash
python scripts/translate_docx.py report-cn.docx -s zh -t en -e cerebras --style professional
# Output: report-cn-en.docx
```

### Word: English → Japanese
```bash
python scripts/translate_docx.py document.docx -s en -t ja -e google --google-api-key KEY
# Output: document-ja.docx
```

### PowerPoint: English → Japanese with terminology
```bash
python scripts/translate_pptx.py slides.pptx -s en -t ja --terminology aws-terms.csv
```

> **网络**: 脚本会自动检测 Cerebras API 连通性——优先直连，直连不通时自动尝试 HTTP_PROXY/HTTPS_PROXY 环境变量中的代理。如需强制使用代理，设置 `HTTPS_PROXY` 环境变量即可。

### Batch translate with custom style
```bash
python scripts/translate_pptx.py marketing.pptx -s en -t zh \
  -e bedrock --style "friendly and engaging" --glossary brand-terms.json
```

## Word Document Support

The skill also translates Word (.docx) documents, covering:
- Body paragraphs (preserving formatting: font, size, color, bold, italic)
- Tables (all cells)
- Headers and footers
- Preserves document styles, page layout, and non-text content

Unlike PPTX, Word documents are **flow-based** (no fixed text boxes), so the auto-fit font algorithm is not needed — the document naturally reflows text.

### Usage
```bash
# Unified entry (auto-detects format)
python scripts/translate.py input.docx -s zh -t en -e cerebras

# Or use the docx-specific script
python scripts/translate_docx.py input.docx -s zh -t en -e cerebras
```

All engines and options (glossary, style, batch-size, etc.) work identically for both formats.

# 这是测试项目同步的一条信息