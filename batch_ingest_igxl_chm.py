#!/usr/bin/env python3
"""
Enhanced batch ingestion script for IGXL CHM documentation.
Integrates with the existing page_indexing_RAG system for optimal CHM handling.
"""

import os
import sys
import subprocess
import tempfile
import shutil
import hashlib
import time
import logging
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any
from datetime import datetime
import concurrent.futures
from tqdm import tqdm

# Add the src directory to path to import RAG modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

# Import existing RAG system modules
from page_indexing_rag.config import CHROMA_PATH
from page_indexing_rag.ingestion import (
    get_chroma_collection,
    get_openai_embedding,
    classify_file,
    semantic_chunk,
    extract_html_document
)

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('igxl_ingestion.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configuration
IGXL_DOCS_PATH = r"C:\Users\cpatle2\Desktop\DOC from IFX\UltraFlex Learning\IGXL_Help\IGXL_Help\IGXL_Help"
BATCH_SIZE = 3  # Process 3 CHM files concurrently
MAX_WORKERS = 2  # Number of parallel extraction workers
CHUNK_SIZE = 500  # Tokens per chunk (optimized for technical docs)
CHUNK_OVERLAP = 2  # Sentence overlap

# HTML extensions to extract from CHM
HTML_EXTENSIONS = {'.html', '.htm', '.xhtml'}

# Initialize ChromaDB collection
collection = get_chroma_collection()

def extract_chm_enhanced(chm_path: Path) -> List[Dict[str, Any]]:
    """
    Enhanced CHM extraction with better error handling and metadata extraction.

    Args:
        chm_path: Path to CHM file

    Returns:
        List of document dictionaries with content and metadata
    """
    documents = []
    chm_name = chm_path.name

    # Create temporary directory for extraction
    temp_root = Path(tempfile.gettempdir()) / "chm_extract"
    temp_root.mkdir(exist_ok=True)
    tmpdir = temp_root / f"chm_{hashlib.md5(str(chm_path).encode()).hexdigest()[:8]}"

    try:
        # Clean up any previous extraction
        if tmpdir.exists():
            shutil.rmtree(tmpdir, ignore_errors=True)
        tmpdir.mkdir(parents=True, exist_ok=True)

        # Use hh.exe to decompile CHM file (Windows only)
        if os.name == 'nt':
            try:
                # Run hh.exe with proper flags
                create_no_window = getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000)
                result = subprocess.run(
                    ['hh.exe', '-decompile', str(tmpdir), str(chm_path)],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=60,
                    creationflags=create_no_window
                )

                # Wait for extraction to complete
                extraction_complete = False
                for _ in range(20):  # Wait up to 10 seconds
                    html_files = list(tmpdir.rglob("*.htm*"))
                    if html_files:
                        extraction_complete = True
                        break
                    time.sleep(0.5)

                if not extraction_complete:
                    logger.warning(f"No HTML files extracted from {chm_name} after timeout")

            except subprocess.TimeoutExpired:
                logger.error(f"CHM extraction timed out for {chm_name}")
            except Exception as e:
                logger.error(f"CHM extraction failed for {chm_name}: {e}")

        # Process extracted HTML files
        html_files = sorted(
            p for p in tmpdir.rglob("*")
            if p.is_file() and p.suffix.lower() in HTML_EXTENSIONS
        )

        if html_files:
            logger.info(f"Found {len(html_files)} HTML files in {chm_name}")

            for html_file in html_files:
                try:
                    # Extract text content from HTML
                    text = extract_html_document(str(html_file))

                    if text and len(text.strip()) > 50:  # Skip very short files
                        # Get relative path for source tracking
                        rel_path = html_file.relative_to(tmpdir)

                        # Determine topic from file path
                        topic = extract_topic_from_path(rel_path)

                        documents.append({
                            'text': text,
                            'source': f"{chm_name}::{rel_path}",
                            'topic': topic,
                            'file_name': html_file.name,
                            'chm_file': chm_name,
                            'extraction_method': 'hh.exe'
                        })

                except Exception as e:
                    logger.warning(f"Failed to process {html_file}: {e}")
                    continue

        # If no documents extracted, try binary fallback
        if not documents:
            logger.info(f"Attempting binary extraction for {chm_name}")
            documents = extract_chm_binary_fallback(chm_path)

    finally:
        # Clean up temporary directory
        if tmpdir.exists():
            shutil.rmtree(tmpdir, ignore_errors=True)

    return documents

def extract_chm_binary_fallback(chm_path: Path) -> List[Dict[str, Any]]:
    """
    Fallback extraction method for CHM files when hh.exe fails.
    Extracts readable strings from binary content.
    """
    documents = []

    try:
        with open(chm_path, 'rb') as f:
            content = f.read()

        # Extract ASCII strings
        import re
        ascii_pattern = re.compile(b'[\x20-\x7E]{20,}')
        ascii_strings = ascii_pattern.findall(content)

        # Extract UTF-16 strings (common in Windows files)
        utf16_strings = []
        try:
            # Try UTF-16 LE decoding in chunks
            for i in range(0, len(content) - 1, 2):
                chunk = content[i:i+1000]
                try:
                    decoded = chunk.decode('utf-16-le', errors='ignore')
                    if len(decoded) > 20 and decoded.isprintable():
                        utf16_strings.append(decoded)
                except:
                    continue
        except:
            pass

        # Combine and clean extracted strings
        all_text = []
        for s in ascii_strings:
            text = s.decode('ascii', errors='ignore').strip()
            if len(text) > 50:  # Keep meaningful strings
                all_text.append(text)

        all_text.extend(utf16_strings)

        if all_text:
            # Join and clean up
            combined_text = '\n'.join(all_text)
            combined_text = re.sub(r'\s+', ' ', combined_text)

            # Split into manageable chunks
            chunks = [combined_text[i:i+5000] for i in range(0, len(combined_text), 4000)]

            for i, chunk in enumerate(chunks[:10]):  # Limit to 10 chunks per file
                documents.append({
                    'text': chunk,
                    'source': f"{chm_path.name}::binary_part{i+1}",
                    'topic': 'extracted_content',
                    'file_name': chm_path.name,
                    'chm_file': chm_path.name,
                    'extraction_method': 'binary_fallback'
                })

    except Exception as e:
        logger.error(f"Binary extraction failed for {chm_path.name}: {e}")

    return documents

def extract_topic_from_path(rel_path: Path) -> str:
    """
    Extract topic/category from file path structure.
    """
    parts = rel_path.parts
    if len(parts) > 1:
        # Use parent directory as topic
        return parts[0]
    else:
        # Use filename stem as topic
        return rel_path.stem

def process_chm_file(chm_path: Path) -> Tuple[str, int, int]:
    """
    Process a single CHM file and ingest into ChromaDB.

    Returns:
        Tuple of (filename, documents_extracted, chunks_added)
    """
    chm_name = chm_path.name
    logger.info(f"Processing: {chm_name}")

    # Check if already ingested
    fingerprint = compute_file_fingerprint(str(chm_path))
    existing = collection.get(where={"fingerprint": fingerprint}, limit=1, include=[])
    if existing['ids']:
        logger.info(f"Skipping {chm_name} - already ingested")
        return chm_name, 0, 0

    # Extract documents from CHM
    documents = extract_chm_enhanced(chm_path)

    if not documents:
        logger.warning(f"No content extracted from {chm_name}")
        return chm_name, 0, 0

    logger.info(f"Extracted {len(documents)} documents from {chm_name}")

    # Process each document
    chunks_added = 0
    for doc_idx, doc in enumerate(documents):
        # Classify document type
        file_type = classify_file(doc['source'], doc['text'][:2000])

        # Create semantic chunks
        chunks = semantic_chunk(doc['text'], max_tokens=CHUNK_SIZE, overlap_sentences=CHUNK_OVERLAP)

        for chunk_idx, chunk_text in enumerate(chunks):
            # Generate unique ID for chunk
            chunk_id = f"{chm_name}_d{doc_idx}_c{chunk_idx}_{hashlib.md5(chunk_text.encode()).hexdigest()[:8]}"

            # Get embedding
            try:
                embedding = get_openai_embedding(chunk_text)
            except Exception as e:
                logger.error(f"Failed to get embedding for chunk {chunk_id}: {e}")
                continue

            # Prepare metadata
            metadata = {
                "source": doc['source'],
                "file_type": file_type,
                "chm_file": doc['chm_file'],
                "topic": doc.get('topic', 'unknown'),
                "extraction_method": doc.get('extraction_method', 'unknown'),
                "chunk_index": chunk_idx,
                "total_chunks": len(chunks),
                "doc_index": doc_idx,
                "total_docs": len(documents),
                "fingerprint": fingerprint,
                "ingestion_date": datetime.now().isoformat(),
                "platform": "V93000",
                "doc_category": "IGXL_Documentation"
            }

            # Add to collection
            try:
                collection.upsert(
                    ids=[chunk_id],
                    documents=[chunk_text],
                    embeddings=[embedding],
                    metadatas=[metadata]
                )
                chunks_added += 1
            except Exception as e:
                logger.error(f"Failed to add chunk {chunk_id} to collection: {e}")

    logger.info(f"Added {chunks_added} chunks from {chm_name}")
    return chm_name, len(documents), chunks_added

def compute_file_fingerprint(file_path: str) -> str:
    """
    Compute SHA-256 fingerprint of file for deduplication.
    """
    h = hashlib.sha256()
    with open(file_path, 'rb') as f:
        # Read first 64KB for fingerprint
        h.update(f.read(65536))
    return h.hexdigest()[:16]

def process_batch(chm_files: List[Path]) -> List[Tuple[str, int, int]]:
    """
    Process a batch of CHM files in parallel.
    """
    results = []

    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(process_chm_file, chm_file): chm_file
                  for chm_file in chm_files}

        for future in concurrent.futures.as_completed(futures):
            try:
                result = future.result(timeout=300)  # 5 minute timeout per file
                results.append(result)
            except Exception as e:
                chm_file = futures[future]
                logger.error(f"Failed to process {chm_file.name}: {e}")
                results.append((chm_file.name, 0, 0))

    return results

def main():
    """
    Main batch ingestion process.
    """
    print("=" * 80)
    print("IGXL CHM Documentation Batch Ingestion")
    print("=" * 80)

    # Verify source directory
    igxl_path = Path(IGXL_DOCS_PATH)
    if not igxl_path.exists():
        logger.error(f"Source directory not found: {IGXL_DOCS_PATH}")
        sys.exit(1)

    # Get all CHM files
    chm_files = list(igxl_path.glob("*.chm"))
    total_files = len(chm_files)

    print(f"\nFound {total_files} CHM files in: {IGXL_DOCS_PATH}")
    print(f"ChromaDB location: {CHROMA_PATH}")
    print(f"Collection: tml_copilot_v2")
    print(f"Batch size: {BATCH_SIZE} files")
    print(f"Max workers: {MAX_WORKERS}")
    print("-" * 80)

    if not chm_files:
        logger.error("No CHM files found!")
        sys.exit(1)

    # Track overall progress
    total_docs_extracted = 0
    total_chunks_added = 0
    failed_files = []
    successful_files = []

    # Process files in batches with progress bar
    with tqdm(total=total_files, desc="Overall Progress", unit="file") as pbar:
        for i in range(0, total_files, BATCH_SIZE):
            batch = chm_files[i:i + BATCH_SIZE]
            batch_num = (i // BATCH_SIZE) + 1
            total_batches = (total_files + BATCH_SIZE - 1) // BATCH_SIZE

            print(f"\n\nProcessing Batch {batch_num}/{total_batches} ({len(batch)} files)")
            print("-" * 40)

            # Process batch
            batch_results = process_batch(batch)

            # Update statistics
            for filename, docs_extracted, chunks_added in batch_results:
                if chunks_added > 0:
                    successful_files.append(filename)
                    total_docs_extracted += docs_extracted
                    total_chunks_added += chunks_added
                else:
                    failed_files.append(filename)

                pbar.update(1)

            # Small delay between batches
            if i + BATCH_SIZE < total_files:
                time.sleep(2)

    # Print summary
    print("\n" + "=" * 80)
    print("INGESTION COMPLETE")
    print("=" * 80)
    print(f"✓ Successfully processed: {len(successful_files)}/{total_files} files")
    print(f"✓ Total documents extracted: {total_docs_extracted}")
    print(f"✓ Total chunks added to ChromaDB: {total_chunks_added}")
    print(f"✓ Final collection size: {collection.count()} documents")

    if failed_files:
        print(f"\n⚠ Failed or skipped files ({len(failed_files)}):")
        for f in failed_files[:10]:  # Show first 10
            print(f"  - {f}")
        if len(failed_files) > 10:
            print(f"  ... and {len(failed_files) - 10} more")

    # Write detailed log
    log_file = Path("igxl_ingestion_summary.txt")
    with open(log_file, 'w') as f:
        f.write("IGXL CHM Ingestion Summary\n")
        f.write("=" * 50 + "\n")
        f.write(f"Date: {datetime.now().isoformat()}\n")
        f.write(f"Source: {IGXL_DOCS_PATH}\n")
        f.write(f"Total files: {total_files}\n")
        f.write(f"Successful: {len(successful_files)}\n")
        f.write(f"Failed: {len(failed_files)}\n")
        f.write(f"Documents extracted: {total_docs_extracted}\n")
        f.write(f"Chunks added: {total_chunks_added}\n")
        f.write("\n\nSuccessful Files:\n")
        for f in successful_files:
            f.write(f"  - {f}\n")
        f.write("\n\nFailed Files:\n")
        for f in failed_files:
            f.write(f"  - {f}\n")

    print(f"\n📄 Detailed summary written to: {log_file}")
    print("\n🎉 Your IGXL documentation is now searchable in the RAG system!")
    print("You can query it using the Streamlit app: streamlit run app/streamlit_app.py")

if __name__ == "__main__":
    main()