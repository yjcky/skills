#!/usr/bin/env python3
"""
Shared translation engine module — used by both PPTX and DOCX translators.

Provides:
  - 4 translation engines: Amazon Translate, Bedrock LLM, Google Translate, Cerebras LLM
  - CJK post-processing: punctuation conversion, spacing fix, word concatenation fix
  - Utility functions: safe path resolution, glossary loading
"""

import json
import os
import re
import sys
from pathlib import Path

import boto3
from botocore.exceptions import ClientError


# ---------------------------------------------------------------------------
# API key loading from external file
# ---------------------------------------------------------------------------

def load_api_keys(key_file: str = None) -> dict:
    """Load API keys from a simple key=value text file.

    Searches in order:
      1. Explicit key_file path (if provided)
      2. api_keys.txt in the same directory as this script
      3. api_keys.txt in the current working directory

    File format (one per line):
      # comments start with #
      CEREBRAS_API_KEY=csk-xxx
      GOOGLE_API_KEY=xxx

    Returns dict of key→value, empty dict if no file found.
    """
    search_paths = []
    if key_file:
        search_paths.append(Path(key_file))

    # Same directory as this script
    script_dir = Path(__file__).resolve().parent
    search_paths.append(script_dir / 'api_keys.txt')

    # Current working directory
    search_paths.append(Path.cwd() / 'api_keys.txt')

    for path in search_paths:
        try:
            if path.exists() and path.is_file():
                keys = {}
                with open(path, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        # Skip empty lines and comments
                        if not line or line.startswith('#'):
                            continue
                        if '=' in line:
                            key, value = line.split('=', 1)
                            key = key.strip()
                            value = value.strip()
                            if value:  # Only store non-empty values
                                keys[key] = value
                if keys:
                    print(f"  [INFO] Loaded API keys from: {path}")
                return keys
        except OSError:
            continue

    return {}

# ---------------------------------------------------------------------------
# Language metadata
# ---------------------------------------------------------------------------

LANGUAGE_NAMES = {
    'en': 'English', 'zh': 'Chinese (Simplified)', 'zh-TW': 'Chinese (Traditional)',
    'ja': 'Japanese', 'ko': 'Korean', 'de': 'German', 'fr': 'French',
    'es': 'Spanish', 'pt': 'Portuguese', 'it': 'Italian', 'ru': 'Russian',
    'ar': 'Arabic', 'hi': 'Hindi', 'th': 'Thai', 'vi': 'Vietnamese',
}


# ---------------------------------------------------------------------------
# CJK post-processing
# ---------------------------------------------------------------------------

def post_process_translation(text: str, target_lang: str) -> str:
    """Clean up common translation artifacts when translating to English."""
    # Remove stray backslash escapes (LLM JSON artifact: \" → ", \; → ;, etc.)
    # Handles cases where LLM outputs escaped characters in JSON that survive parsing
    for ch in ('"', "'", ';', ':', '.', ','):
        text = text.replace('\\' + ch, ch)
    # Fix double-escaped newlines from LLM JSON output
    text = text.replace('\\n', '\n').replace('\\t', '\t')
    # Replace non-breaking hyphen (U+2011, LLM artifact) with regular hyphen
    text = text.replace('‑', '-')

    if target_lang != 'en':
        return text

    # CJK punctuation → English equivalents
    punctuation_map = {
        '，': ',', '。': '.', '；': ';', '：': ':', '？': '?', '！': '!',
        '（': '(', '）': ')', '"': '"', '"': '"', ''': "'", ''': "'",
        '～': '~', '…': '...', '、': ',',
        '《': '<', '》': '>', '【': '[', '】': ']',
    }
    for cjk, eng in punctuation_map.items():
        text = text.replace(cjk, eng)

    # Fix concatenation of CJK char + Latin word
    text = re.sub(r'([一-鿿㐀-䶿豈-﫿])([A-Za-z])', r'\1 \2', text)
    text = re.sub(r'([A-Za-z])([一-鿿㐀-䶿豈-﫿])', r'\1 \2', text)

    # Fix Latin word concatenation: camelCase / word boundary splitting
    # Split lowercase→Uppercase+lowercase (true camelCase): myVariable → my Variable
    # Does NOT split acronym endings: PaaS stays intact (a→S where S has no following lowercase)
    text = re.sub(r'([a-z])([A-Z][a-z])', r'\1 \2', text)
    # Split lowercase→ACRONYM: mySUTPC → my SUTPC
    text = re.sub(r'([a-z])([A-Z]{2,})', r'\1 \2', text)
    # Split ACRONYM→Uppercase+lowercase: SUTPCWill → SUTPC Will
    text = re.sub(r'([A-Z]{2,})([A-Z][a-z])', r'\1 \2', text)
    # Split ACRONYM→lowercase: SUTPCwill → SUTPC will
    text = re.sub(r'([A-Z]{2,})([a-z])', r'\1 \2', text)

    # Fix missing space between letter and digit, or digit and letter
    # over15000 → over 15000, 15000intersections → 15000 intersections
    text = re.sub(r'([A-Za-z])(\d)', r'\1 \2', text)
    text = re.sub(r'(\d)([A-Za-z])', r'\1 \2', text)

    # Fix missing space after punctuation when followed by a letter (not digit —
    # digit means it's a number like "15,000" where no space is needed)
    text = re.sub(r'([.,;:!?])([A-Za-z])', r'\1 \2', text)

    # Remove spaces before English punctuation (common LLM artifact)
    text = re.sub(r'\s+([.,;:!?)])', r'\1', text)

    return text.strip()


def is_cjk_char(c):
    """Return True if character is CJK (Chinese, Japanese, Korean) or fullwidth form."""
    cp = ord(c)
    return (
        (0x4E00 <= cp <= 0x9FFF) or
        (0x3400 <= cp <= 0x4DBF) or
        (0xF900 <= cp <= 0xFAFF) or
        (0x3040 <= cp <= 0x309F) or
        (0x30A0 <= cp <= 0x30FF) or
        (0xAC00 <= cp <= 0xD7AF) or
        (0xFF00 <= cp <= 0xFFEF)
    )


# ---------------------------------------------------------------------------
# Translation engines
# ---------------------------------------------------------------------------

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

        cjk_rules = _build_cjk_rules(source_lang, target_lang)

        return f"""Translate the following texts from {source_name} to {target_name}.

Style: {self.style}{glossary_section}

Rules:
1. Maintain consistent terminology across all texts
2. Keep translations natural and fluent in the target language
3. Return ONLY a JSON array of translated strings in the same order
4. Do NOT include any explanatory text, keys, values, or metadata{cjk_rules}
6. Do NOT wrap translations in JSON key-value format
7. Do NOT create dictionaries like key-value pairs
8. Return ONLY the plain translated text strings
9. Ensure the output is clean, valid JSON that can be parsed by json.loads()

Texts to translate:
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

        non_empty = [(i, t) for i, t in enumerate(texts) if t.strip()]
        if not non_empty:
            return texts

        indices, to_translate = zip(*non_empty)

        prompt = self._build_prompt(list(to_translate), source_lang, target_lang)
        response = self._call_bedrock(prompt)

        try:
            json_match = re.search(r'\[.*\]', response, re.DOTALL)
            if json_match:
                translated = json.loads(json_match.group())
            else:
                raise ValueError("No JSON array found in response")
        except (json.JSONDecodeError, ValueError) as e:
            print(f"  [WARN] Failed to parse LLM response, falling back to original: {e}")
            return texts

        result = list(texts)
        for idx, trans in zip(indices, translated):
            result[idx] = trans

        return result

    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
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

        api_key = api_key or load_api_keys().get("GOOGLE_API_KEY")
        project_id = project_id or load_api_keys().get("GOOGLE_PROJECT_ID")

        self.client = translate.Client(
            api_key=api_key,
            project=project_id
        )

    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        if not text.strip():
            return text

        try:
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
            import httpx
        except ImportError:
            raise ImportError("openai and httpx are required. Install with: pip install openai httpx")

        api_key = api_key or load_api_keys().get("CEREBRAS_API_KEY")

        http_client = self._create_http_client(base_url, api_key)
        connect_timeout = 10
        total_timeout = 120

        self.client = openai.OpenAI(
            api_key=api_key,
            base_url=base_url,
            http_client=http_client,
            timeout=httpx.Timeout(total_timeout, connect=connect_timeout)
        )
        self.model_id = model_id or "gpt-oss-120b"
        self.glossary = glossary or {}
        self.style = style or "professional"
        self.batch_size = batch_size
        self._cache = {}

    def _create_http_client(self, base_url: str, api_key: str):
        """Create httpx client with auto-detected proxy settings."""
        import httpx

        proxies = {}
        for var in ('HTTPS_PROXY', 'HTTP_PROXY', 'https_proxy', 'http_proxy', 'ALL_PROXY', 'all_proxy'):
            if os.getenv(var):
                proxies['http://'] = os.getenv(var)
                proxies['https://'] = os.getenv(var)
                break

        connect_timeout = 5
        direct_client = httpx.Client(timeout=httpx.Timeout(connect_timeout, connect=connect_timeout))
        direct_ok = self._test_connectivity(direct_client, base_url, api_key)
        direct_client.close()

        if direct_ok:
            print("  [INFO] Direct connection to Cerebras API OK")
            return httpx.Client(timeout=httpx.Timeout(120, connect=10))

        if proxies:
            print("  [INFO] Direct connection failed, trying proxy...")
            proxy_client = httpx.Client(proxy=proxies['https://'],
                                        timeout=httpx.Timeout(connect_timeout, connect=connect_timeout))
            proxy_ok = self._test_connectivity(proxy_client, base_url, api_key)
            proxy_client.close()
            if proxy_ok:
                print(f"  [INFO] Proxy connection to Cerebras API OK ({proxies['https://']})")
                return httpx.Client(proxy=proxies['https://'], timeout=httpx.Timeout(120, connect=10))

        print("  [WARN] Cannot reach Cerebras API (direct and proxy both failed)")
        return httpx.Client(timeout=httpx.Timeout(120, connect=10))

    @staticmethod
    def _test_connectivity(client, base_url: str, api_key: str) -> bool:
        """Quick test to check if Cerebras API is reachable."""
        try:
            response = client.get(
                f"{base_url}/models",
                headers={"Authorization": f"Bearer {api_key}"},
            )
            return response.status_code in (200, 401, 403)
        except Exception:
            return False

    def _build_prompt(self, texts: list, source_lang: str, target_lang: str) -> str:
        source_name = LANGUAGE_NAMES.get(source_lang, source_lang)
        target_name = LANGUAGE_NAMES.get(target_lang, target_lang)

        glossary_section = ""
        if self.glossary:
            glossary_items = [f"  - {k} → {v}" for k, v in self.glossary.items()]
            glossary_section = f"\n\nTerminology (use these translations consistently):\n" + "\n".join(glossary_items)

        texts_json = json.dumps(texts, ensure_ascii=False)

        cjk_rules = _build_cjk_rules(source_lang, target_lang)

        return f"""Translate the following texts from {source_name} to {target_name}.

Style: {self.style}{glossary_section}

Rules:
1. Maintain consistent terminology across all texts
2. Keep translations natural and fluent in the target language
3. Return ONLY a valid JSON array of translated strings in the same order
4. Do NOT include any explanatory text, keys, values, or metadata{cjk_rules}
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
        """Translate a single batch of texts using Cerebras API, with retry on rate limit."""
        result = self._translate_batch_inner(texts, source_lang, target_lang)
        return _clean_cjk_spacing(result, target_lang)

    def _translate_batch_inner(self, texts: list, source_lang: str, target_lang: str) -> list:
        if not texts:
            return []

        prompt = self._build_prompt(texts, source_lang, target_lang)
        import time

        # Estimate tokens needed: prompt + room for output
        prompt_chars = len(prompt)
        estimated_input_tokens = prompt_chars // 3  # rough: 3 chars per token for EN, 1.5 for CJK
        self._last_prompt_size = estimated_input_tokens

        max_retries = 3
        for attempt in range(max_retries + 1):
            try:
                response = self.client.chat.completions.create(
                    model=self.model_id,
                    messages=[{"role": "user", "content": prompt}],
                    max_tokens=max(4096, len(texts) * 500),
                    temperature=0.1,
                    n=1
                )

                content = response.choices[0].message.content
                if content is None:
                    refusal = getattr(response.choices[0].message, 'refusal', None)
                    if refusal:
                        print(f"  [WARN] Cerebras refused: {refusal}")
                        return [str(text) for text in texts]
                    # Empty content without refusal — likely timeout/overload, retry
                    if attempt < max_retries:
                        delay = (attempt + 1) * 4
                        print(f"  [RETRY] Cerebras returned empty, waiting {delay}s "
                              f"(attempt {attempt + 1}/{max_retries}, "
                              f"prompt ~{estimated_input_tokens} tok)...")
                        time.sleep(delay)
                        continue
                    print(f"  [WARN] Empty content after {max_retries + 1} attempts")
                    return [str(text) for text in texts]

                response_text = content.strip()
                response_text = re.sub(r'^```json\s*', '', response_text)
                response_text = re.sub(r'\s*```$', '', response_text)
                response_text = response_text.strip()

                # Try to parse as JSON
                try:
                    translated_batch = json.loads(response_text)
                    if not isinstance(translated_batch, list):
                        translated_batch = [translated_batch]
                    return [str(item) for item in translated_batch]
                except json.JSONDecodeError:
                    pass

                # Try to extract array manually
                json_match = re.search(r'\[(.*)\]', response_text, re.DOTALL)
                if json_match:
                    try:
                        translated_batch = json.loads(f"[{json_match.group(1)}]")
                        if not isinstance(translated_batch, list):
                            translated_batch = [translated_batch]
                        return [str(item) for item in translated_batch]
                    except json.JSONDecodeError:
                        pass

                # Extract quoted strings as last resort
                # Use a regex that handles escaped quotes inside strings
                string_matches = re.findall(r'"((?:[^"\\]|\\.)*)"', response_text)
                if string_matches:
                    # Unescape common escapes in extracted strings
                    unescaped = []
                    for s in string_matches:
                        for ch in ('"', "'", ';', ':', '.', ','):
                            s = s.replace('\\' + ch, ch)
                        unescaped.append(s)
                    return unescaped

                # Ultimate fallback: split by lines
                lines = [line.strip() for line in response_text.split('\n')
                         if line.strip() and line.strip() not in ('[', ']')]
                return lines if lines else [str(text) for text in texts]

            except Exception as e:
                err = str(e)
                is_rate_limit = '429' in err or 'rate' in err.lower() or 'too_many_requests' in err.lower()
                if is_rate_limit and attempt < max_retries:
                    delay = (attempt + 1) * 3
                    print(f"  [RETRY] Rate limited, waiting {delay}s (attempt {attempt + 1}/{max_retries})...")
                    time.sleep(delay)
                    continue
                print(f"  [WARN] Cerebras error: {e}")
                return [str(text) for text in texts]

        return [str(text) for text in texts]

    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        """Single text translation (uses translate_batch internally)."""
        if not text.strip():
            return text

        results = self.translate_batch([text], source_lang, target_lang)
        result = results[0] if results else text

        if isinstance(result, list):
            return ' '.join(str(item) for item in result)
        elif isinstance(result, dict):
            return str(result)
        return str(result)


# ---------------------------------------------------------------------------
# Prompt helper
# ---------------------------------------------------------------------------

def _build_cjk_rules(source_lang: str, target_lang: str) -> str:
    """Build CJK-specific translation rules for the LLM prompt."""
    if source_lang in ('zh', 'zh-TW', 'ja', 'ko') and target_lang == 'en':
        return """
5. Convert ALL Chinese/Japanese/Korean punctuation to English equivalents:
   - ，→ ,  。→ .  ；→ ;  ：→ :  ？→ ?  ！→ !
   - （→ (  ）→ )  " → "  " → "  ' → '  ' → '
   - ～→ ~  …→ ...  、→ ,
6. CRITICAL — Word spacing: The source text has NO spaces (Chinese doesn't use them). Your English output MUST have proper spaces between EVERY word. Examples:
   - "SUTPC将配合LTA建立" → "SUTPC will cooperate with LTA to establish" (NOT "SUTPCwill cooperateLTAEstablish")
   - "SUTPC会配合" → "SUTPC will cooperate" (NOT "SUTPCWill cooperate" or "SUTPCwill cooperate")
   - Treat embedded Latin acronyms (SUTPC, LTA, CI/CD, VPN, AD, WOG, SEED) as separate words requiring surrounding spaces
7. Follow English capitalization rules strictly:
   - Only capitalize proper nouns, acronyms, and the first word of a sentence
   - Do NOT capitalize common words in the middle of a sentence (will → will, NOT Will)
"""
    elif target_lang in ('zh', 'ja', 'ko'):
        return """
5. CRITICAL — CJK spacing: Chinese, Japanese, and Korean text uses NO spaces between characters. Do NOT insert spaces between CJK characters. Examples:
   - CORRECT: "标准制定" / INCORRECT: "标准 制定"
   - CORRECT: "交通信号系统" / INCORRECT: "交通 信号 系统"
   - When mixing CJK with Latin/numbers, a thin space is acceptable but not required
6. Preserve formatting tags (<b>...</b>, <i>...</i>, <bi>...</bi>) ONLY where they appear in the source text. Keep them wrapped around the corresponding translated text. Do NOT add NEW tags, do NOT drop existing tags.
"""
    else:
        return """
5. Preserve any formatting markers, placeholders, or special characters
"""


# ---------------------------------------------------------------------------
# CJK spacing cleaner (for LLM outputs targeting CJK languages)
# ---------------------------------------------------------------------------

def _clean_cjk_spacing(texts: list, target_lang: str) -> list:
    """Remove stray spaces inserted between CJK characters by the LLM."""
    if target_lang not in ('zh', 'ja', 'ko'):
        return texts
    cjk_range = r'[一-鿿㐀-䶿぀-ゟ゠-ヿ가-힯豈-﫿　-〿＀-￯]'
    pattern = re.compile(f'({cjk_range})\\s+({cjk_range})')
    return [pattern.sub(r'\1\2', str(t)) for t in texts]


# ---------------------------------------------------------------------------
# Token-aware batch sizing (for LLM engines with TPM limits like Cerebras)
# ---------------------------------------------------------------------------

# Tunable constants for token estimation
_TOKEN_DIVISOR_CJK = 1.5      # CJK characters per token (Chinese, Japanese, Korean)
_TOKEN_DIVISOR_LATIN = 3.5    # Latin characters per token (English, etc.)
_BATCH_FIXED_OVERHEAD = 250   # Prompt template overhead per batch (instructions, rules)
_BATCH_GLOSSARY_OVERHEAD = 50 # Extra overhead when glossary is present
_PER_TEXT_WRAPPER = 8         # JSON array wrapping per text (quotes, comma)
_MAX_ITEMS_PER_BATCH = 40    # Max items per batch to prevent LLM output ordering issues


def estimate_tokens(text: str, target_lang: str = None) -> int:
    """Estimate token count for a text based on character types.

    - CJK characters: ~{_TOKEN_DIVISOR_CJK} chars per token
    - Latin characters: ~{_TOKEN_DIVISOR_LATIN} chars per token
    - Mixed text: counted proportionally by character type
    """
    if not text:
        return 0

    cjk_chars = 0
    latin_chars = 0

    for ch in text:
        cp = ord(ch)
        # CJK Unified Ideographs, CJK Extension A, Hiragana, Katakana, Hangul,
        # CJK Symbols/Punctuation, Fullwidth Forms
        if (0x4E00 <= cp <= 0x9FFF or 0x3400 <= cp <= 0x4DBF or
            0x3040 <= cp <= 0x30FF or 0xAC00 <= cp <= 0xD7AF or
            0x3000 <= cp <= 0x303F or 0xFF00 <= cp <= 0xFFEF):
            cjk_chars += 1
        elif ch.isalpha() or ch.isdigit() or ch in ' \t\n\r':
            latin_chars += 1
        else:
            # Punctuation, symbols — closer to Latin token density
            latin_chars += 0.5

    return int(cjk_chars / _TOKEN_DIVISOR_CJK + latin_chars / _TOKEN_DIVISOR_LATIN)


def estimate_prompt_overhead(num_texts: int, has_glossary: bool = False) -> int:
    """Estimate token overhead from the prompt template (NOT including text content).

    Covers: instruction text, JSON formatting wrapper, optional glossary, CJK rules.
    """
    overhead = _BATCH_FIXED_OVERHEAD
    if has_glossary:
        overhead += _BATCH_GLOSSARY_OVERHEAD
    overhead += num_texts * _PER_TEXT_WRAPPER
    return overhead


def token_aware_batch(texts: list, max_tokens: int = 25000,
                      max_items: int = _MAX_ITEMS_PER_BATCH,
                      target_lang: str = None, has_glossary: bool = False) -> list:
    """Group texts into batches based on estimated token counts (greedy algorithm).

    Each text's token count is estimated; texts are added to the current batch
    until adding the next would exceed ``max_tokens`` or ``max_items``, then a
    new batch starts.

    Args:
        texts: List of text strings to group.
        max_tokens: Maximum estimated tokens per batch (input + output).
        max_items: Maximum number of texts per batch (prevents LLM output misalignment).
        target_lang: Target language code (for future output-size estimation).
        has_glossary: Whether a glossary is included (adds prompt overhead).

    Returns:
        List of batches, where each batch is a list of text strings.
    """
    if not texts:
        return []

    batch_fixed = _BATCH_FIXED_OVERHEAD + (_BATCH_GLOSSARY_OVERHEAD if has_glossary else 0)

    batches = []
    current_batch = []
    current_tokens = 0

    for text in texts:
        text_tokens = estimate_tokens(text, target_lang) + _PER_TEXT_WRAPPER
        # Add fixed overhead only for the first text in a batch
        is_first = len(current_batch) == 0
        extra = batch_fixed if is_first else 0

        tokens_exceeded = current_tokens + text_tokens + extra > max_tokens
        items_exceeded = len(current_batch) >= max_items

        if (tokens_exceeded or items_exceeded) and current_batch:
            batches.append(current_batch)
            current_batch = [text]
            current_tokens = text_tokens + batch_fixed
        else:
            current_batch.append(text)
            current_tokens += text_tokens + extra

    if current_batch:
        batches.append(current_batch)

    # Warn for single-text batches that still exceed the limit
    for i, batch in enumerate(batches):
        if len(batch) == 1:
            single_tokens = (estimate_tokens(batch[0], target_lang) +
                             _PER_TEXT_WRAPPER + batch_fixed)
            if single_tokens > max_tokens:
                print(f"  [WARN] Batch {i + 1}: single text segment "
                      f"({len(batch[0])} chars, ~{single_tokens} est. tokens) "
                      f"exceeds max_batch_tokens ({max_tokens})")

    return batches


# ---------------------------------------------------------------------------
# Engine factory
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Glossary & terminology helpers
# ---------------------------------------------------------------------------

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
        with open(path, encoding='utf-8') as f:
            glossary = json.load(f)
    else:
        with open(path, encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and '=' in line:
                    src, tgt = line.split('=', 1)
                    glossary[src.strip()] = tgt.strip()

    return glossary


# ---------------------------------------------------------------------------
# Safe path resolution (cross-platform, CJK-friendly)
# ---------------------------------------------------------------------------

def safe_resolve_path(path_str: str) -> Path:
    """Resolve input path with encoding-safe handling.

    - expanduser: handle ~ in paths
    - resolve: get absolute canonical path (handles Chinese chars via Unicode API)
    - normalize whitespace: \\xa0 (non-breaking space) often corrupts paths
      passed through MSYS2/Git Bash on Chinese Windows
    """
    p = Path(path_str).expanduser()
    if p.exists():
        return p.resolve()

    # Recovery: try substituting special whitespace with regular spaces
    fixed = path_str
    for ws in ('\xa0', ' ', ' ', ' ', ' ', ' ',
               ' ', ' ', ' ', ' ', ' ', ' ',
               ' ', ' ', '　'):
        fixed = fixed.replace(ws, ' ')
    fixed_p = Path(fixed).expanduser()
    if fixed_p.exists():
        return fixed_p.resolve()
    return p.resolve()


# ---------------------------------------------------------------------------
# Shared CLI argument setup
# ---------------------------------------------------------------------------

def add_common_arguments(parser: argparse.ArgumentParser):
    """Add shared CLI arguments for translation scripts."""
    parser.add_argument('--source', '-s', required=True, help='Source language code (e.g., en, zh, ja)')
    parser.add_argument('--target', '-t', required=True, help='Target language code (e.g., en, zh, ja)')
    parser.add_argument('--output', '-o', help='Output file path')
    parser.add_argument('--engine', '-e', choices=['translate', 'bedrock', 'google', 'cerebras'], default='translate',
                       help='Translation engine')
    parser.add_argument('--model', '-m', help='Model ID')
    parser.add_argument('--terminology', help='Terminology CSV file for Amazon Translate')
    parser.add_argument('--glossary', '-g', help='Glossary file for LLM (JSON or key=value format)')
    parser.add_argument('--style', help='Translation style for LLM (default: professional)')
    parser.add_argument('--batch-size', type=int, default=5, help='Batch size for LLM translation')
    parser.add_argument('--auto-batch', action='store_true',
                       help='Enable token-aware intelligent batching (respects TPM limits)')
    parser.add_argument('--max-batch-tokens', type=int, default=25000,
                       help='Max tokens per batch when --auto-batch is active (default: 25000)')
    parser.add_argument('--max-batch-items', type=int, default=40,
                       help='Max texts per batch when --auto-batch is active (default: 40)')
    parser.add_argument('--region', help='AWS region')
    parser.add_argument('--google-api-key', help='Google Cloud API key')
    parser.add_argument('--google-project-id', help='Google Cloud project ID')
    parser.add_argument('--cerebras-api-key', help='Cerebras API key')
    parser.add_argument('--cerebras-model', help='Cerebras model ID')
    parser.add_argument('--cerebras-base-url', help='Cerebras API base URL')
    parser.add_argument('--no-batch', action='store_true', help='Disable batch mode for LLM')


def setup_windows_encoding():
    """Fix Windows stdio encoding for Chinese character display."""
    if sys.platform == 'win32':
        for stream in [sys.stdout, sys.stderr]:
            if hasattr(stream, 'reconfigure'):
                try:
                    stream.reconfigure(encoding='utf-8', errors='replace')
                except Exception:
                    pass


def build_engine_from_args(args):
    """Create translation engine from CLI args. Returns (engine, batch_mode)."""
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

    return engine, batch_mode


import argparse  # noqa: E402 (import at top, repeated here for self-contained helpers)
