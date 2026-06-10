"""Document ingestion and vector-store operations."""

from __future__ import annotations

import hashlib
import logging
import os
import re
import shutil
import subprocess
import tempfile
import time
from html import unescape
from pathlib import Path

import chromadb
import pdfplumber
import streamlit as st
from pypdf import PdfReader

from .config import CHROMA_PATH, HTML_DATA_DIR, PDF_DATA_DIR, get_embeddings

_lc_embeddings = get_embeddings()

_logger = logging.getLogger(__name__)
_HTML_EXTENSIONS = (".html", ".htm", ".xhtml")
_CHM_EXTENSION = ".chm"
_IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".svg")
_PROGRAM_EXTENSIONS = (
    ".txt",
    ".bas",
    ".cls",
    ".vb",
    ".vbs",
    ".tml",
    ".py",
    ".c",
    ".cpp",
    ".h",
    ".hpp",
    ".cs",
    ".js",
    ".ts",
    ".java",
    ".json",
    ".xml",
    ".yaml",
    ".yml",
    ".ini",
    ".cfg",
    ".log",
    ".md",
    ".csv",
)


def classify_file(filename: str, content: str) -> str:
    ext = filename.lower().rsplit(".", 1)[-1]
    cl = content.lower()
    if f".{ext}" in _IMAGE_EXTENSIONS:
        return "image"
    if ext == "pdf":
        if "tdc" in filename.lower() or "cookbook" in cl:
            return "tdc"
        if "datasheet" in filename.lower() or "absolute max" in cl:
            return "datasheet"
        if "hib" in filename.lower() or "schematic" in cl:
            return "schematic"
        if "pin" in filename.lower() or "pinmap" in cl:
            return "pinmap"
        return "datasheet"
    if ext in ("html", "htm", "xhtml", "chm"):
        if "datasheet" in filename.lower() or "absolute max" in cl or "electrical characteristics" in cl:
            return "datasheet"
        if "tdc" in filename.lower() or "cookbook" in cl:
            return "tdc"
        if "hib" in filename.lower() or "schematic" in cl:
            return "schematic"
        return "general"
    # Test program files (IG-XL for UltraFlex, TML for V93000)
    if ext in ("tml", "bas", "cls") or "testsuite" in cl or "testflow" in cl or "testmethod" in cl or "ig-xl" in cl:
        return "test_program"
    if ext in ("sv", "v") or "module " in cl:
        return "rtl"
    if ext in ("xml", "json") and ("register" in cl or "regmap" in cl):
        return "regmap"
    if "pin" in filename.lower() or "pinmap" in cl:
        return "pinmap"
    if "config" in filename.lower() or "instrument" in cl:
        return "tester_config"
    if "hib" in filename.lower() or "schematic" in cl:
        return "schematic"
    if "datasheet" in filename.lower() or "specification" in cl:
        return "datasheet"
    if "tdc" in filename.lower() or "cookbook" in cl:
        return "tdc"
    return "general"


def _sentence_split(text: str) -> list[str]:
    """Split text on sentence boundaries using simple regex."""
    sentences = re.split(r"(?<=[.!?])\s+", text.strip())
    return [s.strip() for s in sentences if s.strip()]


def semantic_chunk(
    text: str,
    max_tokens: int = 400,
    overlap_sentences: int = 2,
) -> list[str]:
    """
    Build chunks that:
    - respect sentence boundaries
    - stay under max_tokens (approx 4 chars ~= 1 token)
    - carry sentence overlap between neighboring chunks
    """
    sentences = _sentence_split(text)
    chunks: list[str] = []
    current: list[str] = []
    current_len = 0
    max_chars = max_tokens * 4

    def _split_long_sentence(sentence: str, max_chars_per_chunk: int) -> list[str]:
        words = sentence.split()
        if not words:
            return [sentence[:max_chars_per_chunk]] if sentence else []

        out: list[str] = []
        bucket: list[str] = []
        bucket_len = 0
        overlap_words = 20

        for word in words:
            wlen = len(word) + 1
            if bucket and bucket_len + wlen > max_chars_per_chunk:
                out.append(" ".join(bucket))
                bucket = bucket[-overlap_words:]
                bucket_len = sum(len(w) + 1 for w in bucket)
            bucket.append(word)
            bucket_len += wlen

        if bucket:
            out.append(" ".join(bucket))
        return out

    for sent in sentences:
        sent_len = len(sent)
        if sent_len > max_chars:
            if current:
                chunks.append(" ".join(current))
                current = current[-overlap_sentences:]
                current_len = sum(len(s) for s in current)

            chunks.extend(_split_long_sentence(sent, max_chars))
            current = []
            current_len = 0
            continue

        if current_len + sent_len > max_chars and current:
            chunks.append(" ".join(current))
            current = current[-overlap_sentences:]
            current_len = sum(len(s) for s in current)
        current.append(sent)
        current_len += sent_len

    if current:
        chunks.append(" ".join(current))

    return [c for c in chunks if len(c.strip()) > 50]


def _decode_text_bytes(raw_bytes: bytes) -> str:
    for enc in ("utf-8", "utf-16", "utf-16-le", "utf-16-be", "latin-1"):
        try:
            decoded = raw_bytes.decode(enc)
            if decoded.strip():
                return decoded
        except UnicodeDecodeError:
            continue
    return raw_bytes.decode("utf-8", errors="replace")


def _clean_html_text(raw_html: str) -> str:
    text = re.sub(r"<script[^>]*>.*?</script>", " ", raw_html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<style[^>]*>.*?</style>", " ", text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<noscript[^>]*>.*?</noscript>", " ", text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)
    text = unescape(text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _extract_printable_runs(raw: bytes, min_run: int = 40) -> str:
    runs: list[str] = []
    current: list[str] = []

    for byte in raw:
        if 32 <= byte <= 126 or byte in (9, 10, 13):
            current.append(chr(byte))
            continue
        if len(current) >= min_run:
            runs.append("".join(current))
        current = []

    if len(current) >= min_run:
        runs.append("".join(current))

    return "\n".join(runs)


def _extract_utf16le_runs(raw: bytes, min_run: int = 8) -> list[str]:
    runs: list[str] = []
    current: list[str] = []
    i = 0
    n = len(raw)

    while i + 1 < n:
        b0 = raw[i]
        b1 = raw[i + 1]
        if b1 == 0 and (32 <= b0 <= 126 or b0 in (9, 10, 13)):
            current.append(chr(b0))
            i += 2
            continue
        if len(current) >= min_run:
            runs.append("".join(current))
        current = []
        i += 1

    if len(current) >= min_run:
        runs.append("".join(current))
    return runs


def _normalize_string_runs(runs: list[str], min_len: int = 8) -> list[str]:
    normalized: list[str] = []
    seen: set[str] = set()

    for run in runs:
        cleaned = re.sub(r"\s+", " ", run).strip()
        if len(cleaned) < min_len:
            continue
        alpha_count = sum(1 for c in cleaned if c.isalpha())
        if alpha_count < 3:
            continue
        key = cleaned.lower()
        if key in seen:
            continue
        seen.add(key)
        normalized.append(cleaned)

    return normalized


def extract_html_document(html_path: str) -> str:
    try:
        raw = Path(html_path).read_bytes()
    except Exception as exc:  # pragma: no cover - streamlit runtime path
        st.warning(f"Could not read {html_path}: {exc}")
        return ""

    decoded = _decode_text_bytes(raw)
    return _clean_html_text(decoded)


def extract_chm_documents(chm_path: str) -> list[dict]:
    docs: list[dict] = []
    chm_name = os.path.basename(chm_path)

    temp_root = os.path.join(os.path.dirname(chm_path), ".chm_tmp")
    os.makedirs(temp_root, exist_ok=True)
    tmpdir = tempfile.mkdtemp(prefix="chm_extract_", dir=temp_root)
    try:
        try:
            create_no_window = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            subprocess.run(
                ["hh.exe", "-decompile", tmpdir, chm_path],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
                timeout=45,
                creationflags=create_no_window,
            )
        except Exception as exc:
            _logger.warning("CHM decompile failed for %s: %s", chm_path, exc)

        html_files: list[Path] = []
        deadline = time.time() + 8.0
        while time.time() < deadline:
            html_files = sorted(
                p
                for p in Path(tmpdir).rglob("*")
                if p.is_file() and p.suffix.lower() in _HTML_EXTENSIONS
            )
            if html_files:
                break
            time.sleep(0.3)

        for html_file in html_files:
            text = extract_html_document(str(html_file))
            if not text.strip():
                continue
            rel = os.path.relpath(str(html_file), tmpdir).replace("\\", "/")
            docs.append({"source": f"{chm_name}::{rel}", "text": text})
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    if docs:
        return docs

    # Fallback for environments where hh.exe decompile fails.
    # This keeps at least some searchable signal from CHM binaries.
    try:
        raw = Path(chm_path).read_bytes()
        ascii_runs = _extract_printable_runs(raw, min_run=8).splitlines()
        utf16_runs = _extract_utf16le_runs(raw, min_run=8)
        merged_runs = _normalize_string_runs([*ascii_runs, *utf16_runs], min_len=8)
        fallback_text = "\n".join(merged_runs)
        if fallback_text.strip():
            return [{"source": f"{chm_name}::binary_strings", "text": fallback_text}]
    except Exception as exc:
        _logger.warning("CHM binary fallback failed for %s: %s", chm_path, exc)
    return []


def extract_pdf_pages(pdf_path: str) -> list[dict]:
    """
    Returns page dictionaries:
      {page_number, total_pages, text, has_tables}
    """
    pages: list[dict] = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total = len(pdf.pages)
            for i, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                tables = page.extract_tables() or []

                table_text = ""
                for table in tables:
                    for row in table:
                        cleaned = [cell or "" for cell in row]
                        table_text += " | ".join(cleaned) + "\n"

                combined = (text + "\n" + table_text).strip()

                if not combined:
                    reader = PdfReader(pdf_path)
                    combined = reader.pages[i - 1].extract_text() or ""

                pages.append(
                    {
                        "page_number": i,
                        "total_pages": total,
                        "text": combined,
                        "has_tables": len(tables) > 0,
                    }
                )
    except Exception as exc:  # pragma: no cover - streamlit runtime path
        st.warning(f"Could not extract {pdf_path}: {exc}")
    return pages


def extract_pdf_image_entries(pdf_path: str, source_name: str, max_images_per_page: int = 12) -> list[dict]:
    """
    Extract embedded images from a PDF and build searchable text entries.

    Uses OCR when optional dependencies are available, and always stores
    structured image metadata plus local page text context.
    """
    entries: list[dict] = []
    try:
        import fitz  # PyMuPDF
    except Exception as exc:
        _logger.debug("PyMuPDF not available for image extraction: %s", exc)
        return entries

    try:
        doc = fitz.open(pdf_path)
    except Exception as exc:
        _logger.warning("Could not open PDF for image extraction %s: %s", pdf_path, exc)
        return entries

    try:
        for page_idx in range(len(doc)):
            page = doc[page_idx]
            page_number = page_idx + 1
            page_text = (page.get_text("text") or "").strip()
            page_text_snippet = page_text[:1200] if page_text else ""

            images = page.get_images(full=True) or []
            if max_images_per_page > 0:
                images = images[:max_images_per_page]

            for image_idx, img in enumerate(images, start=1):
                xref = img[0]
                try:
                    img_data = doc.extract_image(xref)
                except Exception:
                    continue

                image_bytes = img_data.get("image", b"")
                img_ext = (img_data.get("ext") or "unknown").lower()
                width = int(img_data.get("width") or 0)
                height = int(img_data.get("height") or 0)

                ocr_text = ""
                try:
                    import io
                    from PIL import Image  # type: ignore
                    import pytesseract  # type: ignore

                    pil_image = Image.open(io.BytesIO(image_bytes))
                    ocr_text = (pytesseract.image_to_string(pil_image) or "").strip()
                except Exception:
                    ocr_text = ""

                parts = [
                    f"Embedded PDF image from: {source_name}",
                    f"Page: {page_number}",
                    f"Image index on page: {image_idx}",
                    f"Image format: {img_ext}",
                    f"Dimensions: {width}x{height}",
                ]

                if ocr_text:
                    parts.append("OCR text:")
                    parts.append(ocr_text[:4000])
                elif page_text_snippet:
                    parts.append("Nearby page text context:")
                    parts.append(page_text_snippet)
                else:
                    parts.append("No OCR text available for this image.")

                entries.append(
                    {
                        "page_number": page_number,
                        "image_index": image_idx,
                        "text": "\n".join(parts),
                    }
                )
    finally:
        doc.close()

    return entries


def get_openai_embedding(text: str) -> list[float]:
    """Embed a single text string using LangChain OpenAIEmbeddings."""
    return _lc_embeddings.embed_query(text)


@st.cache_resource
def get_chroma_collection():
    try:
        chroma_client = chromadb.PersistentClient(path=CHROMA_PATH)
    except Exception as exc:
        runtime_path = os.path.join(os.path.dirname(CHROMA_PATH), "chroma_runtime_storage")
        os.makedirs(runtime_path, exist_ok=True)
        try:
            chroma_client = chromadb.PersistentClient(path=runtime_path)
            _logger.warning(
                "Persistent Chroma init failed at %s (%s). Using runtime storage at %s.",
                CHROMA_PATH,
                exc,
                runtime_path,
            )
        except Exception as runtime_exc:
            # Keep application behavior intact even when local SQLite-backed
            # storage is unavailable in the current runtime environment.
            _logger.warning(
                "Persistent Chroma init failed at %s (%s). Runtime fallback at %s also failed (%s). "
                "Falling back to EphemeralClient.",
                CHROMA_PATH,
                exc,
                runtime_path,
                runtime_exc,
            )
            chroma_client = chromadb.EphemeralClient()
    return chroma_client.get_or_create_collection(name="tml_copilot_v2", metadata={"hnsw:space": "cosine"})


collection = get_chroma_collection()


def _file_fingerprint(file_path: str) -> str:
    """SHA-256 of the first 64 KB for duplicate-skip checks."""
    h = hashlib.sha256()
    with open(file_path, "rb") as f:
        h.update(f.read(65536))
    return h.hexdigest()[:16]


def is_file_ingested(file_path: str) -> bool:
    fp = _file_fingerprint(file_path)
    results = collection.get(where={"fingerprint": fp}, limit=1, include=[])
    return len(results["ids"]) > 0


def is_pdf_ingested(pdf_path: str) -> bool:
    return is_file_ingested(pdf_path)


def _source_token(source_name: str) -> str:
    return hashlib.sha1(source_name.encode("utf-8", errors="ignore")).hexdigest()[:16]


def ingest_pdf(pdf_path: str, filename: str) -> tuple[int, int]:
    """
    Ingest a single PDF.
    Returns: (pages_processed, chunks_inserted)
    """
    if is_pdf_ingested(pdf_path):
        return 0, 0

    fingerprint = _file_fingerprint(pdf_path)
    pages = extract_pdf_pages(pdf_path)
    file_type = classify_file(filename, " ".join(p["text"] for p in pages[:3]))

    chunks_inserted = 0
    source_token = _source_token(f"pdf:{filename}")
    for page in pages:
        if not page["text"].strip():
            continue

        sub_chunks = semantic_chunk(page["text"])
        for ci, chunk_text in enumerate(sub_chunks):
            chunk_id = f"pdf_{source_token}_p{page['page_number']}_c{ci}"
            embedding = get_openai_embedding(chunk_text)

            collection.upsert(
                ids=[chunk_id],
                documents=[chunk_text],
                embeddings=[embedding],
                metadatas=[
                    {
                        "source": filename,
                        "file_type": file_type,
                        "page_number": page["page_number"],
                        "total_pages": page["total_pages"],
                        "has_tables": str(page["has_tables"]),
                        "chunk_index": ci,
                        "fingerprint": fingerprint,
                    }
                ],
            )
            chunks_inserted += 1

    # Also ingest embedded PDF images as searchable image-context entries.
    image_entries = extract_pdf_image_entries(pdf_path, filename)
    for entry in image_entries:
        image_text = entry.get("text", "").strip()
        if not image_text:
            continue

        embedding = get_openai_embedding(image_text)
        chunk_id = f"pdfimg_{source_token}_p{entry['page_number']}_i{entry['image_index']}"
        collection.upsert(
            ids=[chunk_id],
            documents=[image_text],
            embeddings=[embedding],
            metadatas=[
                {
                    "source": filename,
                    "file_type": "image_from_pdf",
                    "page_number": entry["page_number"],
                    "total_pages": len(pages),
                    "has_tables": "False",
                    "chunk_index": entry["image_index"],
                    "fingerprint": fingerprint,
                }
            ],
        )
        chunks_inserted += 1

    return len(pages), chunks_inserted


def ingest_html_file(html_path: str, source_name: str | None = None) -> tuple[int, int]:
    """Ingest a single HTML/HTM/XHTML document."""
    if is_file_ingested(html_path):
        return 0, 0

    source = source_name or os.path.basename(html_path)
    fingerprint = _file_fingerprint(html_path)
    text = extract_html_document(html_path)
    if not text.strip():
        return 1, 0

    file_type = classify_file(source, text[:4000])
    chunks = semantic_chunk(text)
    if not chunks and text.strip():
        chunks = [text[:2000]]

    source_token = _source_token(f"html:{source}")
    inserted = 0
    for ci, chunk_text in enumerate(chunks):
        chunk_id = f"html_{source_token}_c{ci}"
        embedding = get_openai_embedding(chunk_text)
        collection.upsert(
            ids=[chunk_id],
            documents=[chunk_text],
            embeddings=[embedding],
            metadatas=[
                {
                    "source": source,
                    "file_type": file_type,
                    "page_number": 1,
                    "total_pages": 1,
                    "has_tables": "False",
                    "chunk_index": ci,
                    "fingerprint": fingerprint,
                }
            ],
        )
        inserted += 1

    return 1, inserted


def ingest_chm_file(chm_path: str, source_name: str | None = None) -> tuple[int, int]:
    """Ingest a CHM by decompiling to HTML when possible."""
    source = source_name or os.path.basename(chm_path)
    fingerprint = _file_fingerprint(chm_path)
    existing = collection.get(where={"fingerprint": fingerprint}, include=["metadatas"])
    if existing.get("ids"):
        metas = existing.get("metadatas") or []
        sources = [str(m.get("source", "")) for m in metas if isinstance(m, dict)]
        only_binary_fallback = bool(sources) and all("::binary_strings" in s for s in sources)
        if only_binary_fallback:
            collection.delete(where={"fingerprint": fingerprint})
        else:
            return 0, 0

    docs = extract_chm_documents(chm_path)
    if not docs:
        return 0, 0

    file_type = classify_file(source, " ".join(d["text"][:800] for d in docs[:3]))
    inserted = 0
    total_docs = len(docs)

    for page_idx, doc in enumerate(docs, start=1):
        doc_source = doc.get("source") or f"{source}::section_{page_idx}"
        chunks = semantic_chunk(doc["text"])
        if not chunks and doc["text"].strip():
            chunks = [doc["text"][:2000]]

        source_token = _source_token(f"chm:{doc_source}")
        for ci, chunk_text in enumerate(chunks):
            chunk_id = f"chm_{source_token}_p{page_idx}_c{ci}"
            embedding = get_openai_embedding(chunk_text)
            collection.upsert(
                ids=[chunk_id],
                documents=[chunk_text],
                embeddings=[embedding],
                metadatas=[
                    {
                        "source": doc_source,
                        "file_type": file_type,
                        "page_number": page_idx,
                        "total_pages": total_docs,
                        "has_tables": "False",
                        "chunk_index": ci,
                        "fingerprint": fingerprint,
                    }
                ],
            )
            inserted += 1

    return total_docs, inserted


def get_pdf_folders() -> dict:
    """Get all folders containing PDF files under PDF_DATA_DIR."""
    if not os.path.exists(PDF_DATA_DIR):
        return {}

    folders = {}

    # Check root directory for PDFs
    root_pdfs = [f for f in os.listdir(PDF_DATA_DIR) if f.lower().endswith(".pdf")]
    if root_pdfs:
        folders["(root)"] = {"path": PDF_DATA_DIR, "count": len(root_pdfs)}

    # Check all subdirectories
    for root, dirs, files in os.walk(PDF_DATA_DIR):
        if root == PDF_DATA_DIR:
            continue  # Skip root, already checked

        pdf_files = [f for f in files if f.lower().endswith(".pdf")]
        if pdf_files:
            rel_path = os.path.relpath(root, PDF_DATA_DIR)
            folders[rel_path] = {"path": root, "count": len(pdf_files)}

    return folders


def ingest_pdfs_from_folder(folder_path: str, progress_cb=None) -> dict:
    """Ingest all PDFs from a specific folder."""
    if not os.path.exists(folder_path):
        return {"error": f"Directory '{folder_path}' not found"}

    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    stats = {"total": len(pdf_files), "new": 0, "skipped": 0, "total_chunks": 0}

    for i, fname in enumerate(pdf_files):
        path = os.path.join(folder_path, fname)
        if progress_cb and pdf_files:
            progress_cb(i / len(pdf_files), fname)

        # Include folder name in the filename for better identification
        rel_folder = os.path.relpath(folder_path, PDF_DATA_DIR)
        if rel_folder != ".":
            display_name = f"{rel_folder}/{fname}"
        else:
            display_name = fname

        pages, chunks = ingest_pdf(path, display_name)
        if chunks == 0:
            stats["skipped"] += 1
        else:
            stats["new"] += 1
            stats["total_chunks"] += chunks

    return stats


def ingest_all_pdfs(progress_cb=None) -> dict:
    """Scan PDF data directory and ingest all PDFs including subdirectories."""
    if not os.path.exists(PDF_DATA_DIR):
        return {"error": f"Directory '{PDF_DATA_DIR}' not found"}

    # Collect all PDF files from root and subdirectories
    all_pdf_files = []
    for root, dirs, files in os.walk(PDF_DATA_DIR):
        for fname in files:
            if fname.lower().endswith(".pdf"):
                full_path = os.path.join(root, fname)
                rel_path = os.path.relpath(full_path, PDF_DATA_DIR)
                all_pdf_files.append((full_path, rel_path))

    stats = {"total": len(all_pdf_files), "new": 0, "skipped": 0, "total_chunks": 0}

    for i, (full_path, rel_path) in enumerate(all_pdf_files):
        if progress_cb and all_pdf_files:
            progress_cb(i / len(all_pdf_files), rel_path)

        pages, chunks = ingest_pdf(full_path, rel_path)
        if chunks == 0:
            stats["skipped"] += 1
        else:
            stats["new"] += 1
            stats["total_chunks"] += chunks

    return stats


def ingest_all_html_documents(progress_cb=None) -> dict:
    """Scan HTML data directory and ingest HTML/CHM documents recursively."""
    if not os.path.exists(HTML_DATA_DIR):
        return {"error": f"Directory '{HTML_DATA_DIR}' not found"}

    files: list[str] = []
    for root, _dirs, filenames in os.walk(HTML_DATA_DIR):
        for fname in filenames:
            ext = os.path.splitext(fname)[1].lower()
            if ext in _HTML_EXTENSIONS or ext == _CHM_EXTENSION:
                files.append(os.path.join(root, fname))

    files.sort()
    stats = {"total": len(files), "new": 0, "skipped": 0, "failed": 0, "total_chunks": 0}

    for i, path in enumerate(files):
        rel_source = os.path.relpath(path, HTML_DATA_DIR).replace("\\", "/")
        if progress_cb and files:
            progress_cb(i / len(files), rel_source)

        try:
            ext = os.path.splitext(path)[1].lower()
            if ext == _CHM_EXTENSION:
                _sections, chunks = ingest_chm_file(path, rel_source)
            else:
                _sections, chunks = ingest_html_file(path, rel_source)
        except Exception as exc:
            _logger.warning("Failed to ingest %s: %s", rel_source, exc)
            stats["failed"] += 1
            continue

        if chunks == 0:
            stats["skipped"] += 1
        else:
            stats["new"] += 1
            stats["total_chunks"] += chunks

    return stats


def ingest_uploaded_text_file(filename: str, content: str) -> tuple[str, int]:
    file_type = classify_file(filename, content)
    chunks = semantic_chunk(content)
    inserted = 0

    for i, chunk_text in enumerate(chunks):
        chunk_id = f"{filename}_chunk{i}"
        embedding = get_openai_embedding(chunk_text)
        collection.upsert(
            ids=[chunk_id],
            documents=[chunk_text],
            embeddings=[embedding],
            metadatas=[
                {
                    "source": filename,
                    "file_type": file_type,
                    "page_number": 0,
                    "total_pages": 0,
                    "has_tables": "False",
                    "chunk_index": i,
                    "fingerprint": "",
                }
            ],
        )
        inserted += 1

    return file_type, inserted


def ingest_program_file(file_path: str, source_name: str | None = None) -> tuple[int, int]:
    """Ingest a source/program/text file from disk."""
    if is_file_ingested(file_path):
        return 0, 0

    source = source_name or os.path.basename(file_path)
    fingerprint = _file_fingerprint(file_path)
    source_token = _source_token(f"program:{source}")

    raw = Path(file_path).read_bytes()
    content = _decode_text_bytes(raw)
    if not content.strip():
        placeholder_text = (
            f"Program file: {source}\n"
            "This file is empty or contains no extractable text content."
        )
        embedding = get_openai_embedding(placeholder_text)
        collection.upsert(
            ids=[f"prog_{source_token}_c0"],
            documents=[placeholder_text],
            embeddings=[embedding],
            metadatas=[
                {
                    "source": source,
                    "file_type": "general",
                    "page_number": 1,
                    "total_pages": 1,
                    "has_tables": "False",
                    "chunk_index": 0,
                    "fingerprint": fingerprint,
                }
            ],
        )
        return 1, 1

    file_type = classify_file(source, content[:4000])
    chunks = semantic_chunk(content)
    if not chunks and content.strip():
        chunks = [content[:2000]]

    inserted = 0
    for ci, chunk_text in enumerate(chunks):
        chunk_id = f"prog_{source_token}_c{ci}"
        embedding = get_openai_embedding(chunk_text)
        collection.upsert(
            ids=[chunk_id],
            documents=[chunk_text],
            embeddings=[embedding],
            metadatas=[
                {
                    "source": source,
                    "file_type": file_type,
                    "page_number": 1,
                    "total_pages": 1,
                    "has_tables": "False",
                    "chunk_index": ci,
                    "fingerprint": fingerprint,
                }
            ],
        )
        inserted += 1

    return 1, inserted


def _build_image_text(file_path: str, source_name: str) -> str:
    """Build searchable text for image assets using available metadata."""
    stat = os.stat(file_path)
    size_kb = max(stat.st_size / 1024.0, 0.0)
    stem_tokens = re.sub(r"[_\-.]+", " ", Path(source_name).stem)

    parts = [
        f"Image asset: {source_name}",
        f"Inferred labels: {stem_tokens}",
        f"File extension: {Path(source_name).suffix.lower()}",
        f"File size KB: {size_kb:.1f}",
    ]

    sidecar_base = os.path.splitext(file_path)[0]
    for ext in (".txt", ".md"):
        sidecar = sidecar_base + ext
        if os.path.exists(sidecar):
            try:
                sidecar_text = _decode_text_bytes(Path(sidecar).read_bytes())
                if sidecar_text.strip():
                    parts.append("Image notes: " + sidecar_text[:4000])
            except Exception:
                pass

    return "\n".join(parts)


def ingest_image_file(file_path: str, source_name: str | None = None) -> tuple[int, int]:
    """Ingest an image asset by embedding descriptive metadata text."""
    if is_file_ingested(file_path):
        return 0, 0

    source = source_name or os.path.basename(file_path)
    fingerprint = _file_fingerprint(file_path)
    image_text = _build_image_text(file_path, source)
    embedding = get_openai_embedding(image_text)
    source_token = _source_token(f"image:{source}")

    collection.upsert(
        ids=[f"img_{source_token}_c0"],
        documents=[image_text],
        embeddings=[embedding],
        metadatas=[
            {
                "source": source,
                "file_type": "image",
                "page_number": 1,
                "total_pages": 1,
                "has_tables": "False",
                "chunk_index": 0,
                "fingerprint": fingerprint,
            }
        ],
    )
    return 1, 1


def ingest_assets_from_directory(root_dir: str, progress_cb=None) -> dict:
    """Ingest program/text files and image assets recursively from a directory."""
    if not os.path.exists(root_dir):
        return {"error": f"Directory '{root_dir}' not found"}

    all_files: list[tuple[str, str]] = []
    for root, _dirs, files in os.walk(root_dir):
        for fname in files:
            ext = os.path.splitext(fname)[1].lower()
            if ext in _PROGRAM_EXTENSIONS or ext in _IMAGE_EXTENSIONS:
                full_path = os.path.join(root, fname)
                rel_path = os.path.relpath(full_path, root_dir).replace("\\", "/")
                all_files.append((full_path, rel_path))

    all_files.sort(key=lambda t: t[1].lower())
    stats = {"total": len(all_files), "new": 0, "skipped": 0, "failed": 0, "total_chunks": 0}

    for i, (full_path, rel_path) in enumerate(all_files):
        if progress_cb and all_files:
            progress_cb(i / len(all_files), rel_path)

        ext = os.path.splitext(full_path)[1].lower()
        try:
            if ext in _IMAGE_EXTENSIONS:
                _items, chunks = ingest_image_file(full_path, rel_path)
            else:
                _items, chunks = ingest_program_file(full_path, rel_path)
        except Exception as exc:
            _logger.warning("Failed to ingest asset %s: %s", rel_path, exc)
            stats["failed"] += 1
            continue

        if chunks == 0:
            stats["skipped"] += 1
        else:
            stats["new"] += 1
            stats["total_chunks"] += chunks

    return stats
