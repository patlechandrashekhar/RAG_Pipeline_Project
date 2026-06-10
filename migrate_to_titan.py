"""
Migration script to re-index ChromaDB with Amazon Titan embeddings.

This script migrates from OpenAI embeddings to Amazon Titan Embed V2,
creating a new collection while preserving all document metadata.
"""

import os
import sys
import asyncio
import json
from pathlib import Path
from typing import List, Dict, Any

import chromadb
import boto3
from dotenv import load_dotenv
from tqdm import tqdm

# Add source to path
PROJECT_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from page_indexing_rag.config_agentsdk import (
    AWS_ACCESS_KEY_ID,
    AWS_SECRET_ACCESS_KEY,
    AWS_REGION,
    CHROMA_PATH
)

load_dotenv()


class ChromaDBMigrator:
    """Migrate ChromaDB from OpenAI to Titan embeddings."""

    def __init__(self):
        """Initialize migrator with both old and new collections."""
        # Initialize AWS Bedrock client
        self.bedrock = boto3.client(
            'bedrock-runtime',
            region_name=AWS_REGION,
            aws_access_key_id=AWS_ACCESS_KEY_ID,
            aws_secret_access_key=AWS_SECRET_ACCESS_KEY
        )

        # Initialize ChromaDB client
        self.chroma_client = chromadb.PersistentClient(path=str(CHROMA_PATH))

        # Old collection (OpenAI embeddings)
        try:
            self.old_collection = self.chroma_client.get_collection("tml_copilot_v2")
            print(f"✓ Found old collection 'tml_copilot_v2' with {self.old_collection.count()} chunks")
        except:
            print("✗ Old collection 'tml_copilot_v2' not found")
            self.old_collection = None

        # New collection (Titan embeddings)
        self.new_collection = self.chroma_client.get_or_create_collection(
            name="tml_copilot_v3_titan",
            metadata={"hnsw:space": "cosine"}
        )
        print(f"✓ Created/opened new collection 'tml_copilot_v3_titan'")

    async def get_titan_embedding(self, text: str) -> List[float]:
        """Get embedding from Amazon Titan Embed V2."""
        body = json.dumps({
            "inputText": text[:8000],  # Titan limit
            "dimensions": 1024,
            "normalize": True
        })

        try:
            response = self.bedrock.invoke_model(
                modelId="amazon.titan-embed-text-v2:0",
                body=body,
                contentType="application/json",
                accept="application/json"
            )
            result = json.loads(response['body'].read())
            return result['embedding']
        except Exception as e:
            print(f"Error getting Titan embedding: {e}")
            return [0.0] * 1024

    async def migrate_batch(self, ids: List[str], documents: List[str], metadatas: List[Dict]) -> int:
        """Migrate a batch of documents."""
        embeddings = []

        # Get Titan embeddings for each document
        for doc in documents:
            embedding = await self.get_titan_embedding(doc)
            embeddings.append(embedding)

        # Add to new collection
        try:
            self.new_collection.add(
                ids=ids,
                embeddings=embeddings,
                documents=documents,
                metadatas=metadatas
            )
            return len(ids)
        except Exception as e:
            print(f"Error adding batch: {e}")
            return 0

    async def migrate(self, batch_size: int = 10):
        """
        Migrate all documents from old to new collection.

        Args:
            batch_size: Number of documents to process at once
        """
        if not self.old_collection:
            print("No old collection to migrate from")
            return

        total = self.old_collection.count()
        print(f"\nStarting migration of {total} chunks...")
        print(f"Batch size: {batch_size}")

        # Get all documents from old collection
        offset = 0
        migrated = 0
        errors = 0

        with tqdm(total=total, desc="Migrating") as pbar:
            while offset < total:
                # Get batch from old collection
                batch = self.old_collection.get(
                    limit=batch_size,
                    offset=offset,
                    include=["documents", "metadatas"]
                )

                if not batch['ids']:
                    break

                # Migrate batch
                batch_migrated = await self.migrate_batch(
                    batch['ids'],
                    batch['documents'],
                    batch['metadatas']
                )

                migrated += batch_migrated
                errors += (len(batch['ids']) - batch_migrated)
                offset += batch_size
                pbar.update(len(batch['ids']))

        print(f"\n✓ Migration complete!")
        print(f"  - Migrated: {migrated} chunks")
        print(f"  - Errors: {errors} chunks")
        print(f"  - New collection size: {self.new_collection.count()} chunks")

    def verify_migration(self, sample_size: int = 5):
        """Verify migration by comparing sample documents."""
        if not self.old_collection:
            return

        print(f"\nVerifying migration (sample size: {sample_size})...")

        # Get sample from old collection
        old_sample = self.old_collection.get(limit=sample_size)

        # Check if documents exist in new collection
        for doc_id in old_sample['ids']:
            try:
                new_doc = self.new_collection.get(ids=[doc_id])
                if new_doc['ids']:
                    print(f"  ✓ Document {doc_id[:8]}... found in new collection")
                else:
                    print(f"  ✗ Document {doc_id[:8]}... NOT found in new collection")
            except:
                print(f"  ✗ Error checking document {doc_id[:8]}...")

    def cleanup_old_collection(self, confirm: bool = False):
        """
        Remove old collection after successful migration.

        Args:
            confirm: Set to True to actually delete
        """
        if not self.old_collection:
            return

        if not confirm:
            print("\n⚠ To delete old collection, run with confirm=True")
            return

        try:
            self.chroma_client.delete_collection("tml_copilot_v2")
            print("✓ Old collection 'tml_copilot_v2' deleted")
        except Exception as e:
            print(f"✗ Error deleting old collection: {e}")


async def main():
    """Main migration process."""
    print("=" * 60)
    print("ChromaDB Migration: OpenAI → Amazon Titan Embeddings")
    print("=" * 60)

    # Check for AWS credentials
    if not AWS_ACCESS_KEY_ID or not AWS_SECRET_ACCESS_KEY:
        print("\n✗ Error: AWS credentials not found in environment")
        print("  Please set AWS_ACCESS_KEY_ID and AWS_SECRET_ACCESS_KEY")
        return

    # Initialize migrator
    migrator = ChromaDBMigrator()

    if not migrator.old_collection:
        print("\n⚠ No existing collection to migrate")
        print("  The new collection 'tml_copilot_v3_titan' is ready for use")
        return

    # Ask for confirmation
    print("\n⚠ This will create a new collection with Titan embeddings")
    print("  Old collection: tml_copilot_v2 (OpenAI)")
    print("  New collection: tml_copilot_v3_titan (Titan)")

    response = input("\nProceed with migration? (y/n): ")
    if response.lower() != 'y':
        print("Migration cancelled")
        return

    # Run migration
    await migrator.migrate(batch_size=10)

    # Verify
    migrator.verify_migration()

    # Option to cleanup
    print("\n" + "=" * 60)
    print("Migration complete! The old collection is still intact.")
    print("To delete the old collection, uncomment the cleanup line in the script")
    # migrator.cleanup_old_collection(confirm=True)


if __name__ == "__main__":
    asyncio.run(main())