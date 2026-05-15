#!/usr/bin/env python3
"""
PPTX Translator - Translate PowerPoint files using Amazon Translate, Bedrock LLM, or Google Translate.

Usage:
    python translate_pptx.py input.pptx --source zh --target en
    python translate_pptx.py input.pptx --source en --target zh --engine bedrock --model anthropic.claude-3-5-sonnet-20241022-v2:0
    python translate_pptx.py input.pptx --source zh --target en --engine google --google-api-key YOUR_API_KEY
    python translate_pptx.py input.pptx --source ja --target en --engine translate --terminology terms.csv
"""

import argparse
import json
import os
import re
import sys
from pathlib import Path

import boto3
from botocore.exceptions import ClientError
from pptx import Presentation
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.shapes import MSO_SHAPE_TYPE

# Language code to MSO_LANGUAGE_ID mapping
LANGUAGE_CODE_TO_LANGUAGE_ID = {
    'af': MSO_LANGUAGE_ID.AFRIKAANS,
    'am': MSO_LANGUAGE_ID.AMHARIC,
    'ar': MSO_LANGUAGE_ID.ARABIC,
    'bg': MSO_LANGUAGE_ID.BULGARIAN,
    'bn': MSO_LANGUAGE_ID.BENGALI,
    'bs': MSO_LANGUAGE_ID.BOSNIAN,
    'cs': MSO_LANGUAGE_ID.CZECH,
    'da': MSO_LANGUAGE_ID.DANISH,
    'de': MSO_LANGUAGE_ID.GERMAN,
    'el': MSO_LANGUAGE_ID.GREEK,
    'en': MSO_LANGUAGE_ID.ENGLISH_US,
    'es': MSO_LANGUAGE_ID.SPANISH,
    'et': MSO_LANGUAGE_ID.ESTONIAN,
    'fi': MSO_LANGUAGE_ID.FINNISH,
    'fr': MSO_LANGUAGE_ID.FRENCH,
    'fr-CA': MSO_LANGUAGE_ID.FRENCH_CANADIAN,
    'ha': MSO_LANGUAGE_ID.HAUSA,
    'he': MSO_LANGUAGE_ID.HEBREW,
    'hi': MSO_LANGUAGE_ID.HINDI,
    'hr': MSO_LANGUAGE_ID.CROATIAN,
    'hu': MSO_LANGUAGE_ID.HUNGARIAN,
    'id': MSO_LANGUAGE_ID.INDONESIAN,
    'it': MSO_LANGUAGE_ID.ITALIAN,
    'ja': MSO_LANGUAGE_ID.JAPANESE,
    'ka': MSO_LANGUAGE_ID.GEORGIAN,
    'ko': MSO_LANGUAGE_ID.KOREAN,
    'lv': MSO_LANGUAGE_ID.LATVIAN,
    'ms': MSO_LANGUAGE_ID.MALAYSIAN,
    'nl': MSO_LANGUAGE_ID.DUTCH,
    'no': MSO_LANGUAGE_ID.NORWEGIAN_BOKMOL,
    'pl': MSO_LANGUAGE_ID.POLISH,
    'ps': MSO_LANGUAGE_ID.PASHTO,
    'pt': MSO_LANGUAGE_ID.BRAZILIAN_PORTUGUESE,
    'ro': MSO_LANGUAGE_ID.ROMANIAN,
    'ru': MSO_LANGUAGE_ID.RUSSIAN,
    'sk': MSO_LANGUAGE_ID.SLOVAK,
    'sl': MSO_LANGUAGE_ID.SLOVENIAN,
    'so': MSO_LANGUAGE_ID.SOMALI,
    'sq': MSO_LANGUAGE_ID.ALBANIAN,
    'sr': MSO_LANGUAGE_ID.SERBIAN_LATIN,
    'sv': MSO_LANGUAGE_ID.SWEDISH,
    'sw': MSO_LANGUAGE_ID.SWAHILI,
    'ta': MSO_LANGUAGE_ID.TAMIL,
    'th': MSO_LANGUAGE_ID.THAI,
    'tr': MSO_LANGUAGE_ID.TURKISH,
    'uk': MSO_LANGUAGE_ID.UKRAINIAN,
    'ur': MSO_LANGUAGE_ID.URDU,
    'vi': MSO_LANGUAGE_ID.VIETNAMESE,
    'zh': MSO_LANGUAGE_ID.CHINESE_SINGAPORE,
    'zh-TW': MSO_LANGUAGE_ID.CHINESE_HONG_KONG_SAR,
}

LANGUAGE_NAMES = {
    'en': 'English', 'zh': 'Chinese (Simplified)', 'zh-TW': 'Chinese (Traditional)',
    'ja': 'Japanese', 'ko': 'Korean', 'de': 'German', 'fr': 'French',
    'es': 'Spanish', 'pt': 'Portuguese', 'it': 'Italian', 'ru': 'Russian',
    'ar': 'Arabic', 'hi': 'Hindi', 'th': 'Thai', 'vi': 'Vietnamese',
}


class TranslationEngine:
    """Base class for translation engines."""
    
    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        raise NotImplementedError


class AmazonTranslateEngine(TranslationEngine):
    """Amazon Translate engine for traditional machine translation."""
    
    def __init__(self, terminology_names: list = None, region: str = None):
        self.client = boto3.client('translate', region_name=region)
        self.terminology_names = terminology_names or []
    
    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        if not text.strip():
            return text
        try:
            response = self.client.translate_text(
                Text=text,
                SourceLanguageCode=source_lang,
                TargetLanguageCode=target_lang,
                TerminologyNames=self.terminology_names
            )
            return response.get('TranslatedText', text)
        except ClientError as e:
            if e.response['Error']['Code'] == 'ValidationException':
                print(f"  [WARN] Invalid text, skipping: {text[:50]}...")
                return text
            raise


class BedrockLLMEngine(TranslationEngine):
    """Bedrock LLM engine for context-aware translation."""
    
    def __init__(self, model_id: str = None, region: str = None, glossary: dict = None,
                 style: str = None, batch_size: int = 20):
        self.client = boto3.client('bedrock-runtime', region_name=region)
        self.model_id = model_id or 'anthropic.claude-3-5-sonnet-20241022-v2:0'
        self.glossary = glossary or {}
        self.style = style or 'professional'
        self.batch_size = batch_size
        self._cache = {}
    
    def _build_prompt(self, texts: list, source_lang: str, target_lang: str) -> str:
        source_name = LANGUAGE_NAMES.get(source_lang, source_lang)
        target_name = LANGUAGE_NAMES.get(target_lang, target_lang)
        
        glossary_section = ""
        if self.glossary:
            glossary_items = [f"  - {k} → {v}" for k, v in self.glossary.items()]
            glossary_section = f"\n\nTerminology (use these translations consistently):\n" + "\n".join(glossary_items)
        
        texts_json = json.dumps(texts, ensure_ascii=False)
        
        return f"""Translate the following texts from {source_name} to {target_name}.

Style: {self.style}{glossary_section}

Rules:
1. Preserve any formatting markers, placeholders, or special characters
2. Maintain consistent terminology across all texts
3. Keep translations natural and fluent in the target language
4. Return ONLY a JSON array of translated strings in the same order

Input texts:
{texts_json}

Output (JSON array only):"""
    
    def _call_bedrock(self, prompt: str) -> str:
        body = json.dumps({
            "anthropic_version": "bedrock-2023-05-31",
            "max_tokens": 8192,
            "messages": [{"role": "user", "content": prompt}]
        })
        
        response = self.client.invoke_model(
            modelId=self.model_id,
            body=body,
            contentType="application/json",
            accept="application/json"
        )
        
        result = json.loads(response['body'].read())
        return result['content'][0]['text']
    
    def translate_batch(self, texts: list, source_lang: str, target_lang: str) -> list:
        """Translate a batch of texts together for context consistency."""
        if not texts:
            return []
        
        # Filter out empty texts but track indices
        non_empty = [(i, t) for i, t in enumerate(texts) if t.strip()]
        if not non_empty:
            return texts
        
        indices, to_translate = zip(*non_empty)
        
        prompt = self._build_prompt(list(to_translate), source_lang, target_lang)
        response = self._call_bedrock(prompt)
        
        # Parse JSON response
        try:
            # Extract JSON array from response
            json_match = re.search(r'\[.*\]', response, re.DOTALL)
            if json_match:
                translated = json.loads(json_match.group())
            else:
                raise ValueError("No JSON array found in response")
        except (json.JSONDecodeError, ValueError) as e:
            print(f"  [WARN] Failed to parse LLM response, falling back to original: {e}")
            return texts
        
        # Reconstruct full list with translations
        result = list(texts)
        for idx, trans in zip(indices, translated):
            result[idx] = trans
        
        return result
    
    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        """Single text translation (uses batch internally with size 1)."""
        if not text.strip():
            return text
        
        cache_key = (text, source_lang, target_lang)
        if cache_key in self._cache:
            return self._cache[cache_key]
        
        result = self.translate_batch([text], source_lang, target_lang)[0]
        self._cache[cache_key] = result
        return result


class GoogleTranslateEngine(TranslationEngine):
    """Google Translate API engine."""

    def __init__(self, api_key: str = None, project_id: str = None):
        try:
            from google.cloud import translate_v2 as translate
        except ImportError:
            raise ImportError("google-cloud-translate is required. Install with: pip install google-cloud-translate")

        self.client = translate.Client(
            api_key=api_key,
            project=project_id
        )

    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        if not text.strip():
            return text

        try:
            # Google Translate automatically detects source if source_lang='auto'
            result = self.client.translate(
                text,
                source_language=source_lang if source_lang and source_lang != 'auto' else None,
                target_language=target_lang
            )
            return result['translatedText']
        except Exception as e:
            print(f"  [WARN] Google Translate error: {e}")
            return text


class CerebrasEngine(TranslationEngine):
    """Cerebras Inference API engine for LLM-based translation."""

    def __init__(self, api_key: str = None, model_id: str = None,
                 base_url: str = "https://api.cerebras.ai/v1", glossary: dict = None,
                 style: str = None, batch_size: int = 5):
        try:
            import openai
        except ImportError:
            raise ImportError("openai is required. Install with: pip install openai")

        self.client = openai.OpenAI(
            api_key=api_key or os.getenv("CEREBRAS_API_KEY") or "csk-pcmdrrkk43k5pt9ykhyjep5jj2528f686evhwh8ce3xm35d9",
            base_url=base_url
        )
        self.model_id = model_id or "gpt-oss-120b"
        self.glossary = glossary or {}
        self.style = style or "professional"
        self.batch_size = batch_size
        self._cache = {}

    def _build_prompt(self, texts: list, source_lang: str, target_lang: str) -> str:
        source_name = LANGUAGE_NAMES.get(source_lang, source_lang)
        target_name = LANGUAGE_NAMES.get(target_lang, target_lang)

        glossary_section = ""
        if self.glossary:
            glossary_items = [f"  - {k} → {v}" for k, v in self.glossary.items()]
            glossary_section = f"\n\nTerminology (use these translations consistently):\n" + "\n".join(glossary_items)

        texts_json = json.dumps(texts, ensure_ascii=False)

        return f"""Translate the following texts from {source_name} to {target_name}.

Style: {self.style}{glossary_section}

Rules:
1. Preserve any formatting markers, placeholders, or special characters
2. Maintain consistent terminology across all texts
3. Keep translations natural and fluent in the target language
4. Return ONLY a valid JSON array of translated strings in the same order
5. Do NOT include any explanatory text, keys, values, or metadata
6. Do NOT wrap translations in JSON key-value format
7. Do NOT create dictionaries like key-value pairs
8. Return ONLY the plain translated text strings
9. Ensure the output is clean, valid JSON that can be parsed by json.loads()

Texts to translate:
{texts_json}

Example output format:
["translated text 1", "translated text 2", "translated text 3"]

Output (JSON array only, NO additional text or formatting):"""

    def translate_batch(self, texts: list, source_lang: str, target_lang: str) -> list:
        """Batch translation using Cerebras API."""
        if not texts:
            return []

        # Split into batches
        batches = [texts[i:i + self.batch_size] for i in range(0, len(texts), self.batch_size)]

        results = []

        for batch_idx, batch in enumerate(batches):
            if not batch:
                continue

            print(f"  Translating batch {batch_idx + 1}/{len(batches)} with {len(batch)} texts...")

            prompt = self._build_prompt(batch, source_lang, target_lang)

            try:
                response = self.client.chat.completions.create(
                    model=self.model_id,
                    messages=[{"role": "user", "content": prompt}],
                    max_tokens=max(2000, len(batch) * 100),
                    temperature=0.3,
                    n=1
                )

                # Extract JSON from response
                content = response.choices[0].message.content
                if content is None:
                    # 模型拒绝响应或内容被过滤，降级为原始文本
                    refusal = getattr(response.choices[0].message, 'refusal', None)
                    print(f"  [WARN] Empty content in Cerebras response (refusal: {refusal})")
                    results.extend([str(text) for text in batch])
                    continue

                response_text = content.strip()
                # Clean the response text
                response_text = re.sub(r'^```json\s*', '', response_text)
                response_text = re.sub(r'\s*```$', '', response_text)
                response_text = response_text.strip()

                # Try to parse as JSON first, with better error handling
                try:
                    # First try to parse the entire response as JSON
                    translated_batch = json.loads(response_text)
                    if not isinstance(translated_batch, list):
                        translated_batch = [translated_batch]

                    # Ensure all elements are strings
                    for item in translated_batch:
                        if isinstance(item, str):
                            results.append(item)
                        else:
                            results.append(str(item))

                    # If we successfully parsed, continue to next batch
                    continue

                except json.JSONDecodeError:
                    print(f"  [WARN] JSON parsing failed, trying fallback...")
                    pass

                # If JSON parsing failed, try to extract array manually
                json_match = re.search(r'\[(.*?)\]', response_text, re.DOTALL)
                if json_match:
                    try:
                        translated_batch = json.loads(f"[{json_match.group(1)}]")
                        if not isinstance(translated_batch, list):
                            translated_batch = [translated_batch]

                        # Ensure all elements are strings
                        for item in translated_batch:
                            if isinstance(item, str):
                                results.append(item)
                            else:
                                results.append(str(item))
                    except json.JSONDecodeError:
                        # If that still fails, try to extract individual strings
                        translated_batch = []
                        # Extract quoted strings using regex
                        string_matches = re.findall(r'"([^"]*)"', response_text)
                        if string_matches:
                            translated_batch = string_matches
                        else:
                            # Fallback: split by lines and take non-empty lines
                            for line in response_text.split('\n'):
                                line = line.strip()
                                if line and line not in ['[', ']']:
                                    translated_batch.append(line)

                        # Ensure all elements are strings
                        for item in translated_batch:
                            if isinstance(item, str):
                                results.append(item)
                            else:
                                results.append(str(item))
                else:
                    # If no array found, extract individual strings
                    string_matches = re.findall(r'"([^"]*)"', response_text)
                    if string_matches:
                        translated_batch = string_matches
                    else:
                        # Fallback: split by lines and take non-empty lines
                        translated_batch = []
                        for line in response_text.split('\n'):
                            line = line.strip()
                            if line and line not in ['[', ']']:
                                translated_batch.append(line)

                    # Ensure all elements are strings
                    for item in translated_batch:
                        if isinstance(item, str):
                            results.append(item)
                        else:
                            results.append(str(item))

            except Exception as e:
                print(f"  [WARN] Cerebras error in batch {batch_idx + 1}: {e}")
                # Return original texts for this batch, maintaining exact count
                results.extend([str(text) for text in batch])

        # Fill results with original texts for any failed batches
        while len(results) < len(texts):
            original_text = texts[len(results)]
            results.append(str(original_text) if original_text else "")

        # Trim results if we have too many (shouldn't happen, but safety check)
        if len(results) > len(texts):
            print(f"  [WARN] Translation count mismatch: {len(texts)} original vs {len(results)} translated")
            results = results[:len(texts)]

        return results

    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        """Single text translation (uses translate_batch internally)."""
        if not text.strip():
            return text

        # Use translate_batch with single text
        results = self.translate_batch([text], source_lang, target_lang)
        result = results[0] if results else text

        # Ensure result is a string
        if isinstance(result, list):
            return ' '.join(str(item) for item in result)
        elif isinstance(result, dict):
            return str(result)
        return str(result)


def import_terminology(client, terminology_file: str, name: str = 'pptx-translator-terminology'):
    """Import terminology file to Amazon Translate."""
    print(f"Importing terminology from {terminology_file}...")
    with open(terminology_file, 'rb') as f:
        client.import_terminology(
            Name=name,
            MergeStrategy='OVERWRITE',
            TerminologyData={'File': bytearray(f.read()), 'Format': 'CSV'}
        )
    return name


def load_glossary(glossary_file: str) -> dict:
    """Load glossary file for LLM translation (JSON or simple key=value format)."""
    glossary = {}
    path = Path(glossary_file)
    
    if path.suffix == '.json':
        with open(path) as f:
            glossary = json.load(f)
    else:
        # Simple format: source=target per line
        with open(path) as f:
            for line in f:
                line = line.strip()
                if line and '=' in line:
                    src, tgt = line.split('=', 1)
                    glossary[src.strip()] = tgt.strip()
    
    return glossary


def extract_texts_from_presentation(presentation) -> list:
    """Extract all translatable texts with their locations."""
    texts = []

    for slide_idx, slide in enumerate(presentation.slides):
        # Shapes with text frames
        for shape_idx, shape in enumerate(iter_shapes(slide.shapes)):
            if shape.has_table:
                for row_idx, row in enumerate(shape.table.rows):
                    for cell_idx, cell in enumerate(row.cells):
                        for para_idx, para in enumerate(cell.text_frame.paragraphs):
                            for run_idx, run in enumerate(para.runs):
                                if run.text.strip():
                                    texts.append({
                                        'text': run.text,
                                        'location': ('table', slide_idx, shape_idx, row_idx, cell_idx, para_idx, run_idx),
                                        'original_text': run.text
                                    })
            elif shape.has_text_frame:
                for para_idx, para in enumerate(shape.text_frame.paragraphs):
                    # 按段落合并所有runs的文本
                    para_text = ' '.join(run.text.strip() for run in para.runs if run.text.strip())
                    if para_text.strip():
                        texts.append({
                            'text': para_text,
                            'location': ('shape', slide_idx, shape_idx, para_idx, 'paragraph'),
                            'original_text': para_text,
                            'paragraph': para  # 保存段落引用
                        })

        # Notes
        if slide.has_notes_slide:
            for para_idx, para in enumerate(slide.notes_slide.notes_text_frame.paragraphs):
                # 按段落合并所有runs的文本
                para_text = ' '.join(run.text.strip() for run in para.runs if run.text.strip())
                if para_text.strip():
                    texts.append({
                        'text': para_text,
                        'location': ('notes', slide_idx, para_idx, 'paragraph'),
                        'original_text': para_text,
                        'paragraph': para  # 保存段落引用
                    })

    return texts


def find_run_by_location(location, presentation):
    """Find the specific run object based on location information."""
    loc_type, slide_idx, *rest = location

    slide = presentation.slides[slide_idx]

    shapes = list(iter_shapes(slide.shapes))

    if loc_type == 'shape':
        shape_idx, para_idx, run_type = rest
        if shape_idx >= len(shapes):
            return None
        shape = shapes[shape_idx]
        if shape.has_text_frame:
            para = shape.text_frame.paragraphs[para_idx]
            if run_type == 'paragraph':
                # 返回段落中的第一个run作为示例
                return para.runs[0] if para.runs else None
    elif loc_type == 'table':
        shape_idx, row_idx, cell_idx, para_idx, run_idx = rest
        if shape_idx >= len(shapes):
            return None
        shape = shapes[shape_idx]
        if shape.has_table:
            row = shape.table.rows[row_idx]
            cell = row.cells[cell_idx]
            para = cell.text_frame.paragraphs[para_idx]
            if run_idx < len(para.runs):
                return para.runs[run_idx]
    elif loc_type == 'notes':
        para_idx, run_type = rest
        if slide.has_notes_slide:
            para = slide.notes_slide.notes_text_frame.paragraphs[para_idx]
            if run_type == 'paragraph':
                # 返回段落中的第一个run作为示例
                return para.runs[0] if para.runs else None

    return None


def apply_paragraph_translation(para, translated_text):
    """将翻译后的文本应用到段落的所有runs中"""
    if not para or not translated_text:
        return

    runs = para.runs
    if runs:
        # 把完整翻译文本放入第一个run，清空其余run
        # 不要用空格分割——对中文等无空格语言会出错
        runs[0].text = translated_text
        for run in runs[1:]:
            run.text = ""
    else:
        run = para.add_run(translated_text)


def create_engine(engine_type: str, **kwargs):
    """Factory function to create translation engines."""
    if engine_type == 'translate':
        return AmazonTranslateEngine(**kwargs)
    elif engine_type == 'bedrock':
        return BedrockLLMEngine(**kwargs)
    elif engine_type == 'google':
        return GoogleTranslateEngine(**kwargs)
    elif engine_type == 'cerebras':
        return CerebrasEngine(**kwargs)
    else:
        raise ValueError(f"Unknown engine: {engine_type}")


def iter_shapes(shapes):
    """递归遍历形状，展开组合形状（GroupShape）"""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)
        else:
            yield shape


def translate_presentation(presentation, engine: TranslationEngine, source_lang: str,
                          target_lang: str, batch_mode: bool = False):
    """Translate all text in a presentation."""
    
    if batch_mode and (isinstance(engine, BedrockLLMEngine) or isinstance(engine, CerebrasEngine)):
        # Batch mode: extract all texts, translate in batches, apply back
        print("Extracting texts...")
        text_items = extract_texts_from_presentation(presentation)
        
        if not text_items:
            print("No text found to translate.")
            return
        
        print(f"Found {len(text_items)} text segments")
        
        # Translate in batches
        texts = [item['text'] for item in text_items]
        batch_size = engine.batch_size
        translated_texts = []
        
        for i in range(0, len(texts), batch_size):
            batch = texts[i:i + batch_size]
            print(f"Translating batch {i//batch_size + 1}/{(len(texts) + batch_size - 1)//batch_size}...")
            translated_batch = engine.translate_batch(batch, source_lang, target_lang)
            translated_texts.extend(translated_batch)
        
        # Apply translations
        print("Applying translations...")
        # Validate that we have the same number of translations as original texts
        if len(text_items) != len(translated_texts):
            print(f"  [WARN] Translation count mismatch: {len(text_items)} original vs {len(translated_texts)} translated")
            # Fill missing translations with original texts
            while len(translated_texts) < len(text_items):
                original_item = text_items[len(translated_texts)]
                translated_texts.append(str(original_item['text']) if original_item['text'] else "")

        for idx, (item, translated) in enumerate(zip(text_items, translated_texts)):
            # Ensure we have a valid translation
            if not translated or len(translated.strip()) == 0:
                print(f"  [WARN] Empty translation at position {idx}, using original text")
                translated = str(item['text']) if item['text'] else ""

            # Ensure translated result is a string and apply it
            if isinstance(translated, dict):
                print(f"  [WARN] Dictionary format at position {idx}, converting to string")
                translated_text = str(translated)
            elif isinstance(translated, list):
                print(f"  [WARN] List format at position {idx}, joining to string")
                translated_text = ' '.join(str(t) for t in translated)
            else:
                translated_text = str(translated)

            # 检查是否是段落级别的翻译
            if 'paragraph' in item:
                # 使用段落级别的翻译
                para = item['paragraph']
                try:
                    apply_paragraph_translation(para, translated_text)
                    print(f"  Applied paragraph translation to {item['location'][0]} on slide {item['location'][1] + 1}")
                except Exception as e:
                    print(f"  [ERROR] Failed to apply paragraph translation at position {idx}: {e}")
                    # 使用原始文本作为后备
                    apply_paragraph_translation(para, item['original_text'] if item['original_text'] else "")
            else:
                # 使用原有的run级别翻译
                run = find_run_by_location(item['location'], presentation)
                if run:
                    try:
                        run.text = translated_text
                        print(f"  Applied translation to {item['location'][0]} on slide {item['location'][1] + 1}")
                    except Exception as e:
                        print(f"  [ERROR] Failed to apply translation at position {idx}: {e}")
                        # 使用原始文本作为后备
                        run.text = item['original_text'] if item['original_text'] else ""
                else:
                    print(f"  [ERROR] Could not find run at position {idx} with location {item['location']}")
                    # 尝试查找备选或跳过
                    pass
    else:
        # Sequential mode: translate one by one
        total_slides = len(presentation.slides)
        for slide_idx, slide in enumerate(presentation.slides, start=1):
            print(f"Slide {slide_idx}/{total_slides}")
            
            for shape in iter_shapes(slide.shapes):
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            translate_text_frame(cell.text_frame, engine, source_lang, target_lang)
                elif shape.has_text_frame:
                    translate_text_frame(shape.text_frame, engine, source_lang, target_lang)
            
            if slide.has_notes_slide:
                translate_text_frame(slide.notes_slide.notes_text_frame, engine, source_lang, target_lang)


def translate_text_frame(text_frame, engine: TranslationEngine, source_lang: str, target_lang: str):
    """Translate all text in a text frame."""
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            if run.text.strip():
                run.text = engine.translate(run.text, source_lang, target_lang)


def main():
    parser = argparse.ArgumentParser(
        description='Translate PPTX files using Amazon Translate or Bedrock LLM'
    )
    parser.add_argument('input_file', help='Input PPTX file path')
    parser.add_argument('--source', '-s', required=True, help='Source language code (e.g., en, zh, ja)')
    parser.add_argument('--target', '-t', required=True, help='Target language code (e.g., en, zh, ja)')
    parser.add_argument('--output', '-o', help='Output file path (default: input-{target}.pptx)')
    parser.add_argument('--engine', '-e', choices=['translate', 'bedrock', 'google', 'cerebras'], default='translate',
                       help='Translation engine: translate (Amazon Translate), bedrock (LLM), google (Google Translate), or cerebras (Cerebras LLM)')
    parser.add_argument('--model', '-m', help='Bedrock model ID (default: claude-3.5-sonnet)')
    parser.add_argument('--terminology', help='Terminology CSV file for Amazon Translate')
    parser.add_argument('--glossary', '-g', help='Glossary file for Bedrock LLM (JSON or key=value format)')
    parser.add_argument('--style', help='Translation style for LLM (default: professional)')
    parser.add_argument('--batch-size', type=int, default=5, help='Batch size for LLM translation')
    parser.add_argument('--region', help='AWS region')
    parser.add_argument('--google-api-key', help='Google Cloud API key')
    parser.add_argument('--google-project-id', help='Google Cloud project ID')
    parser.add_argument('--cerebras-api-key', help='Cerebras API key')
    parser.add_argument('--cerebras-model', help='Cerebras model ID (default: llama3.1-8b)')
    parser.add_argument('--cerebras-base-url', help='Cerebras API base URL (default: https://api.cerebras.ai/v1)')
    parser.add_argument('--no-batch', action='store_true', help='Disable batch mode for LLM')
    
    args = parser.parse_args()
    
    # Validate input file
    input_path = Path(args.input_file)
    if not input_path.exists():
        print(f"Error: Input file not found: {args.input_file}")
        sys.exit(1)
    
    # Determine output path
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_name(f"{input_path.stem}-{args.target}.pptx")
    
    # Initialize engine
    if args.engine == 'bedrock':
        glossary = load_glossary(args.glossary) if args.glossary else {}
        engine = create_engine(
            'bedrock',
            model_id=args.model,
            region=args.region,
            glossary=glossary,
            style=args.style or 'professional',
            batch_size=args.batch_size
        )
        batch_mode = not args.no_batch
        print(f"Using Bedrock LLM engine: {engine.model_id}")
    elif args.engine == 'google':
        engine = create_engine(
            'google',
            api_key=args.google_api_key,
            project_id=args.google_project_id
        )
        batch_mode = False
        print("Using Google Translate engine")
    elif args.engine == 'cerebras':
        glossary = load_glossary(args.glossary) if args.glossary else {}
        engine = create_engine(
            'cerebras',
            api_key=args.cerebras_api_key,
            model_id=args.cerebras_model,
            glossary=glossary,
            style=args.style or 'professional',
            batch_size=args.batch_size
        )
        batch_mode = not args.no_batch
        print(f"Using Cerebras LLM engine: {engine.model_id}")
    else:
        terminology_names = []
        if args.terminology:
            client = boto3.client('translate', region_name=args.region)
            name = import_terminology(client, args.terminology)
            terminology_names = [name]
        engine = create_engine(
            'translate',
            terminology_names=terminology_names,
            region=args.region
        )
        batch_mode = False
        print("Using Amazon Translate engine")
    
    # Load and translate presentation
    print(f"Loading {args.input_file}...")
    presentation = Presentation(args.input_file)
    
    print(f"Translating from {args.source} to {args.target}...")
    translate_presentation(presentation, engine, args.source, args.target, batch_mode=batch_mode)
    
    # Save output
    print(f"Saving {output_path}...")
    presentation.save(str(output_path))
    print("Done!")


if __name__ == '__main__':
    main()
