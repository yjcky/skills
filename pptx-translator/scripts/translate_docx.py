#!/usr/bin/env python3
"""
DOCX Translator - Translate Word documents using Amazon Translate, Bedrock LLM,
Google Translate, or Cerebras LLM.

Usage:
    python translate_docx.py input.docx --source zh --target en
    python translate_docx.py input.docx --source en --target zh --engine bedrock
    python translate_docx.py input.docx --source zh --target en --engine cerebras
"""

import argparse
import sys
from pathlib import Path

from docx import Document

# Import shared translation infrastructure
from translate_common import (
    TranslationEngine, AmazonTranslateEngine, BedrockLLMEngine,
    GoogleTranslateEngine, CerebrasEngine,
    create_engine, load_glossary, import_terminology,
    post_process_translation,
    LANGUAGE_NAMES,
    safe_resolve_path, setup_windows_encoding, add_common_arguments,
    build_engine_from_args,
)


# ---------------------------------------------------------------------------
# DOCX-specific: text extraction
# ---------------------------------------------------------------------------

def _merge_paragraph_text(para):
    """Merge all runs in a paragraph into a single string for full-context translation."""
    return ''.join(run.text for run in para.runs if run.text.strip())


def extract_texts_from_document(doc) -> list:
    """Extract all translatable texts from a Word document with location metadata.

    Covers: body paragraphs, tables, headers, and footers.
    """
    texts = []

    # --- Body paragraphs ---
    for para_idx, para in enumerate(doc.paragraphs):
        text = _merge_paragraph_text(para)
        if text.strip():
            texts.append({
                'text': text,
                'location': ('paragraph', para_idx),
                'original_text': text,
                'paragraph': para,
            })

    # --- Tables ---
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                for para_idx, para in enumerate(cell.paragraphs):
                    text = _merge_paragraph_text(para)
                    if text.strip():
                        texts.append({
                            'text': text,
                            'location': ('table', table_idx, row_idx, cell_idx, para_idx),
                            'original_text': text,
                            'paragraph': para,
                        })

    # --- Headers & Footers ---
    for section_idx, section in enumerate(doc.sections):
        # Header
        if section.header:
            for para_idx, para in enumerate(section.header.paragraphs):
                text = _merge_paragraph_text(para)
                if text.strip():
                    texts.append({
                        'text': text,
                        'location': ('header', section_idx, para_idx),
                        'original_text': text,
                        'paragraph': para,
                    })
            # Header tables
            for table_idx, table in enumerate(section.header.tables):
                for row_idx, row in enumerate(table.rows):
                    for cell_idx, cell in enumerate(row.cells):
                        for para_idx, para in enumerate(cell.paragraphs):
                            text = _merge_paragraph_text(para)
                            if text.strip():
                                texts.append({
                                    'text': text,
                                    'location': ('header_table', section_idx, table_idx, row_idx, cell_idx, para_idx),
                                    'original_text': text,
                                    'paragraph': para,
                                })

        # Footer
        if section.footer:
            for para_idx, para in enumerate(section.footer.paragraphs):
                text = _merge_paragraph_text(para)
                if text.strip():
                    texts.append({
                        'text': text,
                        'location': ('footer', section_idx, para_idx),
                        'original_text': text,
                        'paragraph': para,
                    })
            # Footer tables
            for table_idx, table in enumerate(section.footer.tables):
                for row_idx, row in enumerate(table.rows):
                    for cell_idx, cell in enumerate(row.cells):
                        for para_idx, para in enumerate(cell.paragraphs):
                            text = _merge_paragraph_text(para)
                            if text.strip():
                                texts.append({
                                    'text': text,
                                    'location': ('footer_table', section_idx, table_idx, row_idx, cell_idx, para_idx),
                                    'original_text': text,
                                    'paragraph': para,
                                })

    return texts


# ---------------------------------------------------------------------------
# DOCX-specific: text application
# ---------------------------------------------------------------------------

def find_paragraph_by_location(location, doc):
    """Resolve a location tuple back to the corresponding paragraph object."""
    loc_type = location[0]

    if loc_type == 'paragraph':
        _, para_idx = location
        if para_idx < len(doc.paragraphs):
            return doc.paragraphs[para_idx]

    elif loc_type == 'table':
        _, table_idx, row_idx, cell_idx, para_idx = location
        if table_idx < len(doc.tables):
            table = doc.tables[table_idx]
            if row_idx < len(table.rows):
                row = table.rows[row_idx]
                if cell_idx < len(row.cells):
                    cell = row.cells[cell_idx]
                    if para_idx < len(cell.paragraphs):
                        return cell.paragraphs[para_idx]

    elif loc_type == 'header':
        _, section_idx, para_idx = location
        if section_idx < len(doc.sections):
            header = doc.sections[section_idx].header
            if header and para_idx < len(header.paragraphs):
                return header.paragraphs[para_idx]

    elif loc_type == 'header_table':
        _, section_idx, table_idx, row_idx, cell_idx, para_idx = location
        if section_idx < len(doc.sections):
            header = doc.sections[section_idx].header
            if header and table_idx < len(header.tables):
                table = header.tables[table_idx]
                if row_idx < len(table.rows):
                    row = table.rows[row_idx]
                    if cell_idx < len(row.cells):
                        cell = row.cells[cell_idx]
                        if para_idx < len(cell.paragraphs):
                            return cell.paragraphs[para_idx]

    elif loc_type == 'footer':
        _, section_idx, para_idx = location
        if section_idx < len(doc.sections):
            footer = doc.sections[section_idx].footer
            if footer and para_idx < len(footer.paragraphs):
                return footer.paragraphs[para_idx]

    elif loc_type == 'footer_table':
        _, section_idx, table_idx, row_idx, cell_idx, para_idx = location
        if section_idx < len(doc.sections):
            footer = doc.sections[section_idx].footer
            if footer and table_idx < len(footer.tables):
                table = footer.tables[table_idx]
                if row_idx < len(table.rows):
                    row = table.rows[row_idx]
                    if cell_idx < len(row.cells):
                        cell = row.cells[cell_idx]
                        if para_idx < len(cell.paragraphs):
                            return cell.paragraphs[para_idx]

    return None


def apply_paragraph_translation(para, translated_text):
    """Apply translated text to a paragraph — first run gets all text, rest cleared.

    Preserves the first run's formatting (font, size, color, bold, italic, etc.)
    and applies it to the translated text.
    """
    if not para or not translated_text:
        return

    runs = para.runs
    if runs:
        runs[0].text = translated_text
        for run in runs[1:]:
            run.text = ""
    else:
        para.add_run(translated_text)


# ---------------------------------------------------------------------------
# Sequential translation fallback (for non-batch engines)
# ---------------------------------------------------------------------------

def translate_document_sequential(doc, engine: TranslationEngine, source_lang: str, target_lang: str):
    """Translate the entire document one paragraph at a time (for non-LLM engines)."""
    todo = []

    # Collect body paragraphs
    for para in doc.paragraphs:
        if _merge_paragraph_text(para).strip():
            todo.append(para)

    # Collect table paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if _merge_paragraph_text(para).strip():
                        todo.append(para)

    # Collect headers/footers
    for section in doc.sections:
        for header in (section.header,):
            if header:
                for para in header.paragraphs:
                    if _merge_paragraph_text(para).strip():
                        todo.append(para)
                for table in header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                if _merge_paragraph_text(para).strip():
                                    todo.append(para)
        for footer in (section.footer,):
            if footer:
                for para in footer.paragraphs:
                    if _merge_paragraph_text(para).strip():
                        todo.append(para)
                for table in footer.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                if _merge_paragraph_text(para).strip():
                                    todo.append(para)

    total = len(todo)
    for idx, para in enumerate(todo, 1):
        print(f"  Paragraph {idx}/{total}")
        translated = False
        for run in para.runs:
            if run.text.strip():
                translated_text = engine.translate(run.text, source_lang, target_lang)
                run.text = post_process_translation(translated_text, target_lang)
                translated = True
        # If no runs were translated but paragraph has text, translate the first run
        if not translated and para.runs:
            para_text = _merge_paragraph_text(para)
            if para_text.strip():
                translated_text = engine.translate(para_text, source_lang, target_lang)
                para.runs[0].text = post_process_translation(translated_text, target_lang)


# ---------------------------------------------------------------------------
# Main translation logic
# ---------------------------------------------------------------------------

def translate_document(doc, engine: TranslationEngine, source_lang: str,
                       target_lang: str, batch_mode: bool = False):
    """Translate all text in a Word document."""

    if batch_mode and (isinstance(engine, (BedrockLLMEngine, CerebrasEngine))):
        print("Extracting texts...")
        text_items = extract_texts_from_document(doc)

        if not text_items:
            print("No text found to translate.")
            return

        print(f"Found {len(text_items)} text segments")

        texts = [item['text'] for item in text_items]
        batch_size = engine.batch_size

        batches = []
        for i in range(0, len(texts), batch_size):
            batches.append(texts[i:i + batch_size])
        batch_count = len(batches)

        import concurrent.futures
        translated_texts = []

        max_workers = min(3, batch_count)
        print(f"Translating {batch_count} batches with {max_workers} concurrent workers...")

        results_by_index = {}
        failed_count = 0

        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_idx = {}
            for idx, batch in enumerate(batches):
                future = executor.submit(engine.translate_batch, batch, source_lang, target_lang)
                future_to_idx[future] = idx

            for future in concurrent.futures.as_completed(future_to_idx):
                idx = future_to_idx[future]
                try:
                    results_by_index[idx] = future.result()
                except Exception as e:
                    failed_count += 1
                    print(f"  [ERROR] Batch {idx + 1}/{batch_count} failed: {e}")
                    results_by_index[idx] = [str(t) for t in batches[idx]]

            for i in range(batch_count):
                batch_result = results_by_index.get(i, [str(t) for t in batches[i]])
                translated_texts.extend(batch_result)

        if batch_count > 0 and failed_count == batch_count:
            raise RuntimeError(
                f"All {batch_count} translation batches failed. "
                "Check network connectivity to Cerebras API."
            )

        print("Applying translations...")
        if len(text_items) != len(translated_texts):
            print(f"  [WARN] Translation count mismatch: {len(text_items)} original vs {len(translated_texts)} translated")
            while len(translated_texts) < len(text_items):
                original_item = text_items[len(translated_texts)]
                translated_texts.append(str(original_item['text']) if original_item['text'] else "")

        for idx, (item, translated) in enumerate(zip(text_items, translated_texts)):
            if not translated or len(str(translated).strip()) == 0:
                print(f"  [WARN] Empty translation at position {idx}, using original text")
                translated = str(item['text']) if item['text'] else ""

            if isinstance(translated, dict):
                print(f"  [WARN] Dictionary format at position {idx}, converting to string")
                translated_text = str(translated)
            elif isinstance(translated, list):
                print(f"  [WARN] List format at position {idx}, joining to string")
                translated_text = ' '.join(str(t) for t in translated)
            else:
                translated_text = str(translated)

            translated_text = post_process_translation(translated_text, target_lang)

            para = item.get('paragraph')
            if para is not None:
                try:
                    apply_paragraph_translation(para, translated_text)
                    loc_type = item['location'][0]
                    print(f"  Applied translation to {loc_type} #{idx + 1}")
                except Exception as e:
                    print(f"  [ERROR] Failed to apply translation at position {idx}: {e}")
                    apply_paragraph_translation(para, item['original_text'] if item['original_text'] else "")
            else:
                # Fallback: resolve by location
                para = find_paragraph_by_location(item['location'], doc)
                if para:
                    try:
                        apply_paragraph_translation(para, translated_text)
                        print(f"  Applied translation via location lookup #{idx + 1}")
                    except Exception as e:
                        print(f"  [ERROR] Failed to apply translation at position {idx}: {e}")
                else:
                    print(f"  [ERROR] Could not find paragraph at position {idx} with location {item['location']}")
    else:
        # Sequential mode
        translate_document_sequential(doc, engine, source_lang, target_lang)


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description='Translate DOCX files using Amazon Translate, Bedrock LLM, Google Translate, or Cerebras LLM'
    )
    parser.add_argument('input_file', help='Input DOCX file path')
    add_common_arguments(parser)

    args = parser.parse_args()

    setup_windows_encoding()

    input_path = safe_resolve_path(args.input_file)
    if not input_path.exists():
        print(f"Error: Input file not found: {args.input_file}")
        sys.exit(1)

    if args.output:
        output_path = safe_resolve_path(args.output).parent / Path(args.output).name
        output_path.parent.mkdir(parents=True, exist_ok=True)
    else:
        output_path = input_path.with_name(f"{input_path.stem}-{args.target}.docx")

    engine, batch_mode = build_engine_from_args(args)

    print(f"Loading {input_path}...")
    doc = Document(str(input_path))

    print(f"Translating from {args.source} to {args.target}...")
    try:
        translate_document(doc, engine, args.source, args.target, batch_mode=batch_mode)
    except RuntimeError as e:
        print(f"Error: {e}")
        sys.exit(1)

    print(f"Saving {output_path}...")
    doc.save(str(output_path.resolve()))
    print("Done!")


if __name__ == '__main__':
    main()
