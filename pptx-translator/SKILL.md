---
name: pptx-translator
description: "Translate PowerPoint (.pptx) files between languages using Amazon Translate, Bedrock LLM, Google Translate, or Cerebras LLM. Use when: (1) translating presentations to other languages, (2) localizing slide decks for international audiences, (3) converting Chinese/Japanese/Korean presentations to English or vice versa. Supports four engines: Amazon Translate (fast, cost-effective), Bedrock LLM (high quality), Google Translate (widely supported), and Cerebras LLM (open-source models). Preserves formatting, styles, and layout."
---

# PPTX Translator

Translate PowerPoint files using AWS services.

## Quick Start

```bash
# Amazon Translate (fast, cheap)
python scripts/translate_pptx.py input.pptx -s zh -t en

# Bedrock LLM (high quality)
python scripts/translate_pptx.py input.pptx -s zh -t en -e bedrock

# Google Translate (widely supported)
python scripts/translate_pptx.py input.pptx -s zh -t en -e google --google-api-key YOUR_API_KEY

# Cerebras LLM (open-source models, API key built-in)
python scripts/translate_pptx.py input.pptx -s zh -t en -e cerebras
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

### Translate with Amazon Translate
```bash
python scripts/translate_pptx.py input.pptx \
  --source en --target zh \
  --engine translate \
  --terminology terms.csv  # optional
```

### Translate with Bedrock LLM
```bash
python scripts/translate_pptx.py input.pptx \
  --source en --target zh \
  --engine bedrock \
  --model anthropic.claude-3-5-sonnet-20241022-v2:0 \
  --glossary glossary.json \
  --style professional \
  --batch-size 20
```

### Translate with Google Translate
```bash
python scripts/translate_pptx.py input.pptx \
  --source en --target zh \
  --engine google \
  --google-api-key YOUR_API_KEY \
  --google-project-id YOUR_PROJECT_ID  # optional
```

### Translate with Cerebras LLM
```bash
python scripts/translate_pptx.py input.pptx \
  --source en --target zh \
  --engine cerebras \
  --style professional
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
pip install boto3 python-pptx

# For Google Translate support (optional)
pip install google-cloud-translate

# For Cerebras support (optional)
pip install openai httpx
```

AWS credentials must be configured for Amazon Translate and Bedrock (`~/.aws/credentials` or environment variables).  
Google Translate API key must be provided via `--google-api-key` argument or environment variable.  
**Cerebras API key 已内置于脚本中，无需额外配置。** 脚本会自动检测网络连通性（直连 → 代理回退），无需手动设置代理。

## Examples

### Chinese presentation → English
```bash
python scripts/translate_pptx.py deck-cn.pptx -s zh -t en -e bedrock
# Output: deck-cn-en.pptx
```

### English → Japanese with terminology
```bash
python scripts/translate_pptx.py slides.pptx -s en -t ja \
  --terminology aws-terms.csv
```

### Google Translate: Chinese → English
```bash
python scripts/translate_pptx.py deck-cn.pptx -s zh -t en -e google \
  --google-api-key YOUR_API_KEY
# Output: deck-cn-en.pptx
```

### Cerebras LLM: Chinese → English
```bash
python scripts/translate_pptx.py deck-cn.pptx -s zh -t en -e cerebras --style professional
# Output: deck-cn-en.pptx
```

> **网络**: 脚本会自动检测 Cerebras API 连通性——优先直连，直连不通时自动尝试 HTTP_PROXY/HTTPS_PROXY 环境变量中的代理。如需强制使用代理，设置 `HTTPS_PROXY` 环境变量即可。

### Batch translate with custom style
```bash
python scripts/translate_pptx.py marketing.pptx -s en -t zh \
  -e bedrock --style "friendly and engaging" --glossary brand-terms.json
```
# 这是测试项目同步的一条信息