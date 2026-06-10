"""
Ingestion functions for Agent SDK version.

This is a standalone version of ingestion functions that doesn't depend
on the old config.py to avoid circular imports and Unicode issues.
"""

import os
import re
import hashlib
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple


def classify_file(filename: str, content: str) -> str:
    """
    Classify a file based on its name and content.

    Args:
        filename: Name of the file
        content: File content (for content-based classification)

    Returns:
        File type classification
    """
    filename_lower = filename.lower()
    content_lower = content.lower()[:2000] if content else ""

    # TML files
    if filename_lower.endswith('.tml') or 'testsuite' in content_lower or 'testflow' in content_lower:
        return 'tml'

    # TDC files
    if 'tdc' in filename_lower or 'cookbook' in content_lower:
        return 'tdc'

    # Datasheets
    if 'datasheet' in filename_lower or 'absolute max' in content_lower or 'electrical characteristics' in content_lower:
        return 'datasheet'

    # Schematics
    if 'hib' in filename_lower or 'schematic' in content_lower:
        return 'schematic'

    # Pin maps
    if 'pin' in filename_lower or 'pinmap' in content_lower:
        return 'pinmap'

    # RTL files
    if filename_lower.endswith(('.sv', '.v')) or 'module' in content_lower:
        return 'rtl'

    # Register maps
    if (filename_lower.endswith(('.xml', '.json')) and
        ('register' in content_lower or 'regmap' in content_lower)):
        return 'regmap'

    # Tester config
    if 'config' in filename_lower or 'instrument' in content_lower:
        return 'tester_config'

    return 'general'


def semantic_chunk(
    text: str,
    max_tokens: int = 400,
    token_chars: int = 4,
    overlap_sentences: int = 2,
    min_chunk_chars: int = 50
) -> List[str]:
    """
    Split text into semantic chunks with sentence overlap.

    Args:
        text: Text to chunk
        max_tokens: Maximum tokens per chunk
        token_chars: Approximate characters per token
        overlap_sentences: Number of sentences to overlap
        min_chunk_chars: Minimum chunk size in characters

    Returns:
        List of text chunks
    """
    if not text or len(text.strip()) < min_chunk_chars:
        return []

    # Split into sentences
    sentence_pattern = r'(?<=[.!?])\s+'
    sentences = re.split(sentence_pattern, text.strip())

    if not sentences:
        return []

    chunks = []
    current_chunk = []
    current_size = 0
    max_chars = max_tokens * token_chars

    for sentence in sentences:
        sentence = sentence.strip()
        if not sentence:
            continue

        sentence_size = len(sentence)

        # If single sentence exceeds max, split it
        if sentence_size > max_chars:
            # Save current chunk if exists
            if current_chunk:
                chunks.append(' '.join(current_chunk))
                # Keep last sentences for overlap
                if overlap_sentences > 0 and len(current_chunk) >= overlap_sentences:
                    current_chunk = current_chunk[-overlap_sentences:]
                    current_size = sum(len(s) for s in current_chunk) + len(current_chunk) - 1
                else:
                    current_chunk = []
                    current_size = 0

            # Split long sentence by words
            words = sentence.split()
            word_chunk = []
            word_size = 0

            for word in words:
                word_len = len(word) + 1  # +1 for space
                if word_size + word_len > max_chars and word_chunk:
                    chunks.append(' '.join(word_chunk))
                    # Keep last 20 words for overlap
                    word_chunk = word_chunk[-20:] if len(word_chunk) > 20 else []
                    word_size = sum(len(w) for w in word_chunk) + len(word_chunk) - 1
                word_chunk.append(word)
                word_size += word_len

            if word_chunk:
                current_chunk = word_chunk[-20:] if len(word_chunk) > 20 else word_chunk
                current_size = sum(len(w) for w in current_chunk) + len(current_chunk) - 1
            continue

        # Check if adding sentence exceeds limit
        if current_size + sentence_size + 1 > max_chars and current_chunk:
            # Save current chunk
            chunks.append(' '.join(current_chunk))

            # Keep last sentences for overlap
            if overlap_sentences > 0 and len(current_chunk) >= overlap_sentences:
                current_chunk = current_chunk[-overlap_sentences:]
                current_size = sum(len(s) for s in current_chunk) + len(current_chunk) - 1
            else:
                current_chunk = []
                current_size = 0

        # Add sentence to current chunk
        current_chunk.append(sentence)
        current_size += sentence_size + 1

    # Don't forget last chunk
    if current_chunk:
        chunks.append(' '.join(current_chunk))

    # Filter out chunks that are too small
    return [chunk for chunk in chunks if len(chunk.strip()) >= min_chunk_chars]


def extract_pdf_pages(pdf_path: str) -> List[Dict[str, Any]]:
    """
    Extract text and metadata from PDF pages.

    Args:
        pdf_path: Path to PDF file

    Returns:
        List of page data dictionaries
    """
    pages_data = []

    # Try pdfplumber first (best for tables)
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"Opening PDF with pdfplumber: {total_pages} pages")

            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ""

                # Check for tables
                tables = page.extract_tables() or []
                has_tables = len(tables) > 0

                # Convert tables to text
                if has_tables:
                    for table in tables:
                        if table:
                            table_text = "\n".join([
                                " | ".join([str(cell or "") for cell in row])
                                for row in table if row
                            ])
                            text += f"\n\nTable:\n{table_text}\n"

                if text.strip():  # Only add pages with content
                    pages_data.append({
                        'page_number': i + 1,
                        'total_pages': total_pages,
                        'text': text,
                        'has_tables': has_tables
                    })
                else:
                    print(f"Page {i+1} has no extractable text")

        if pages_data:
            print(f"Extracted {len(pages_data)} pages with pdfplumber")
            return pages_data

    except Exception as e:
        print(f"pdfplumber failed: {e}")

    # Fallback to pypdf
    try:
        import pypdf
        reader = pypdf.PdfReader(pdf_path)
        total_pages = len(reader.pages)
        print(f"Fallback to pypdf: {total_pages} pages")

        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ""

            if text.strip():  # Only add pages with content
                pages_data.append({
                    'page_number': i + 1,
                    'total_pages': total_pages,
                    'text': text,
                    'has_tables': False
                })

        if pages_data:
            print(f"Extracted {len(pages_data)} pages with pypdf")

    except Exception as e:
        print(f"pypdf also failed: {e}")

    return pages_data


def extract_html_document(html_path: str, source_name: str) -> str:
    """
    Extract text from HTML document.

    Args:
        html_path: Path to HTML file
        source_name: Source name for the document

    Returns:
        Extracted text content
    """
    from html import unescape
    import re

    try:
        # Try different encodings
        for encoding in ['utf-8', 'utf-16', 'latin-1']:
            try:
                with open(html_path, 'r', encoding=encoding) as f:
                    html_content = f.read()
                break
            except UnicodeDecodeError:
                continue
        else:
            # If all encodings fail, read as binary and decode with errors='ignore'
            with open(html_path, 'rb') as f:
                html_content = f.read().decode('utf-8', errors='ignore')

        # Remove script and style tags
        html_content = re.sub(r'<script[^>]*>.*?</script>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
        html_content = re.sub(r'<style[^>]*>.*?</style>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
        html_content = re.sub(r'<noscript[^>]*>.*?</noscript>', '', html_content, flags=re.DOTALL | re.IGNORECASE)

        # Remove HTML tags but keep content
        text = re.sub(r'<[^>]+>', ' ', html_content)

        # Unescape HTML entities
        text = unescape(text)

        # Clean up whitespace
        text = re.sub(r'\s+', ' ', text)
        text = text.strip()

        return text

    except Exception as e:
        print(f"Error extracting HTML from {html_path}: {e}")
        return ""


def extract_chm_documents(chm_path: str, source_name: str) -> List[Dict[str, str]]:
    """
    Extract documents from CHM (Windows Help) file.

    Args:
        chm_path: Path to CHM file
        source_name: Source name for the documents

    Returns:
        List of document dictionaries with 'source' and 'text' keys
    """
    import subprocess
    import tempfile
    import shutil
    from pathlib import Path

    documents = []

    # Try to decompile CHM using hh.exe (Windows only)
    if os.name == 'nt':  # Windows
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                # Use hh.exe to decompile
                subprocess.run(
                    ['hh.exe', '-decompile', temp_dir, chm_path],
                    capture_output=True,
                    text=True,
                    timeout=30
                )

                # Extract HTML files
                temp_path = Path(temp_dir)
                for html_file in temp_path.glob('**/*.html'):
                    text = extract_html_document(str(html_file), source_name)
                    if text and len(text) > 50:
                        doc_name = f"{source_name}/{html_file.relative_to(temp_path)}"
                        documents.append({
                            'source': doc_name,
                            'text': text
                        })

        except Exception as e:
            print(f"CHM decompilation failed: {e}")

    # If no documents extracted or not Windows, try binary extraction
    if not documents:
        try:
            with open(chm_path, 'rb') as f:
                content = f.read()
                # Try to extract readable strings
                text = content.decode('utf-8', errors='ignore')
                # Clean up binary artifacts
                text = re.sub(r'[\x00-\x1f\x7f-\x9f]+', ' ', text)
                text = re.sub(r'\s+', ' ', text)

                # Split into chunks if too large
                if len(text) > 10000:
                    chunks = [text[i:i+10000] for i in range(0, len(text), 8000)]
                    for i, chunk in enumerate(chunks):
                        documents.append({
                            'source': f"{source_name}_part{i+1}",
                            'text': chunk
                        })
                else:
                    documents.append({
                        'source': source_name,
                        'text': text
                    })
        except Exception as e:
            print(f"Binary extraction failed: {e}")

    return documents