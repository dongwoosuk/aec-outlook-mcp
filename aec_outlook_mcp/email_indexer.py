"""
Email Indexer - Embedding and ChromaDB storage for emails
"""

import json
import hashlib
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional
from pathlib import Path

import chromadb
from chromadb.config import Settings

from .config import get_config
from .outlook_reader import EmailMessage, get_outlook_reader

# Lazy load sentence-transformers (slow import)
_embedding_model = None


def get_embedding_model():
    """Lazy load embedding model"""
    global _embedding_model
    if _embedding_model is None:
        from sentence_transformers import SentenceTransformer
        config = get_config()
        _embedding_model = SentenceTransformer(config.embedding_model)
    return _embedding_model


class EmailIndexer:
    """Index and search emails using embeddings and ChromaDB"""

    COLLECTION_EMAILS = "emails"
    COLLECTION_ATTACHMENTS = "attachments"
    METADATA_COLLECTION = "_metadata"

    def __init__(self):
        self.config = get_config()
        self.config.ensure_directories()

        # Initialize ChromaDB
        self._client = chromadb.PersistentClient(
            path=self.config.db_path,
            settings=Settings(anonymized_telemetry=False),
        )

        # Get or create collections
        self._emails_collection = self._client.get_or_create_collection(
            name=self.COLLECTION_EMAILS,
            metadata={"description": "Email content embeddings"},
        )

        self._attachments_collection = self._client.get_or_create_collection(
            name=self.COLLECTION_ATTACHMENTS,
            metadata={"description": "Attachment content embeddings"},
        )

    def _generate_email_id(self, email: EmailMessage) -> str:
        """Generate a unique ID for an email"""
        # Use Outlook's EntryID (already unique)
        return hashlib.sha256(email.entry_id.encode()).hexdigest()[:32]

    def _chunk_text(self, text: str) -> List[str]:
        """Split text into chunks for embedding"""
        if not text:
            return []

        chunk_size = self.config.chunk_size
        overlap = self.config.chunk_overlap

        chunks = []
        start = 0

        while start < len(text):
            end = start + chunk_size
            chunk = text[start:end]

            # Try to break at sentence boundary
            if end < len(text):
                # Look for sentence end in last 100 chars
                for sep in [". ", ".\n", "! ", "!\n", "? ", "?\n", "\n\n"]:
                    last_sep = chunk.rfind(sep)
                    if last_sep > chunk_size - 100:
                        chunk = chunk[: last_sep + len(sep)]
                        break

            chunks.append(chunk.strip())
            start = start + len(chunk) - overlap

        return [c for c in chunks if c]  # Filter empty

    def _prepare_email_document(self, email: EmailMessage) -> str:
        """Prepare email content for embedding"""
        parts = []

        # Subject (weighted by repetition)
        if email.subject:
            parts.append(f"Subject: {email.subject}")
            parts.append(email.subject)  # Add again for emphasis

        # Sender
        if email.sender_name:
            parts.append(f"From: {email.sender_name}")

        # Recipients
        if email.recipients:
            parts.append(f"To: {', '.join(email.recipients[:5])}")

        # Date
        if email.received_time:
            parts.append(f"Date: {email.received_time.strftime('%Y-%m-%d')}")

        # Body
        if email.body:
            parts.append(email.body)

        return "\n".join(parts)

    def index_email(self, email: EmailMessage) -> Dict[str, Any]:
        """Index a single email"""
        email_id = self._generate_email_id(email)

        # Check if already indexed
        try:
            existing = self._emails_collection.get(ids=[email_id])
            if existing and existing["ids"]:
                return {"indexed": False, "reason": "already_exists", "id": email_id}
        except:
            pass

        # Prepare document
        document = self._prepare_email_document(email)
        if not document:
            return {"indexed": False, "reason": "empty_content", "id": email_id}

        # Chunk if necessary
        chunks = self._chunk_text(document)
        if not chunks:
            chunks = [document]

        # Generate embeddings
        model = get_embedding_model()
        embeddings = model.encode(chunks).tolist()

        # Prepare metadata
        metadata = {
            "entry_id": email.entry_id,
            "subject": email.subject[:500] if email.subject else "",
            "sender_name": email.sender_name[:200] if email.sender_name else "",
            "sender_email": email.sender_email[:200] if email.sender_email else "",
            "folder_path": email.folder_path[:200] if email.folder_path else "",
            "received_time": email.received_time.isoformat() if email.received_time else "",
            "has_attachments": email.has_attachments,
            "conversation_id": email.conversation_id[:100] if email.conversation_id else "",
            "importance": email.importance,
        }

        # Add to collection
        if len(chunks) == 1:
            self._emails_collection.add(
                ids=[email_id],
                embeddings=[embeddings[0]],
                documents=[chunks[0]],
                metadatas=[metadata],
            )
        else:
            # Multiple chunks - add with chunk index
            ids = [f"{email_id}_chunk_{i}" for i in range(len(chunks))]
            metadatas = [{**metadata, "chunk_index": i, "parent_id": email_id} for i in range(len(chunks))]

            self._emails_collection.add(
                ids=ids,
                embeddings=embeddings,
                documents=chunks,
                metadatas=metadatas,
            )

        return {"indexed": True, "id": email_id, "chunks": len(chunks)}

    def index_attachment_text(
        self,
        email_id: str,
        attachment_filename: str,
        text_content: str,
    ) -> Dict[str, Any]:
        """Index extracted text from an attachment"""
        att_id = hashlib.sha256(f"{email_id}_{attachment_filename}".encode()).hexdigest()[:32]

        # Chunk text
        chunks = self._chunk_text(text_content)
        if not chunks:
            return {"indexed": False, "reason": "empty_content"}

        # Generate embeddings
        model = get_embedding_model()
        embeddings = model.encode(chunks).tolist()

        # Metadata
        metadata = {
            "email_id": email_id,
            "filename": attachment_filename[:200],
            "type": "attachment",
        }

        # Add to collection
        if len(chunks) == 1:
            self._attachments_collection.add(
                ids=[att_id],
                embeddings=[embeddings[0]],
                documents=[chunks[0]],
                metadatas=[metadata],
            )
        else:
            ids = [f"{att_id}_chunk_{i}" for i in range(len(chunks))]
            metadatas = [{**metadata, "chunk_index": i, "parent_id": att_id} for i in range(len(chunks))]

            self._attachments_collection.add(
                ids=ids,
                embeddings=embeddings,
                documents=chunks,
                metadatas=metadatas,
            )

        return {"indexed": True, "id": att_id, "chunks": len(chunks)}

    def search_emails(
        self,
        query: str,
        limit: int = 10,
        sender_filter: Optional[str] = None,
        date_from: Optional[datetime] = None,
        date_to: Optional[datetime] = None,
        folder_filter: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """
        Search emails by semantic similarity

        Args:
            query: Search query text
            limit: Maximum results
            sender_filter: Filter by sender name (partial match)
            date_from: Filter emails after this date
            date_to: Filter emails before this date
            folder_filter: Filter by folder path (partial match)

        Returns:
            List of search results with email metadata and relevance score
        """
        # Generate query embedding
        model = get_embedding_model()
        query_embedding = model.encode(query).tolist()

        # Build where clause for filtering
        where = None
        where_conditions = []

        if sender_filter:
            where_conditions.append({"sender_name": {"$contains": sender_filter}})

        if folder_filter:
            where_conditions.append({"folder_path": {"$contains": folder_filter}})

        # ChromaDB doesn't support date range directly, we'll filter post-query
        # if date_from or date_to:
        #     # Would need custom filtering

        if where_conditions:
            if len(where_conditions) == 1:
                where = where_conditions[0]
            else:
                where = {"$and": where_conditions}

        # Query collection
        results = self._emails_collection.query(
            query_embeddings=[query_embedding],
            n_results=limit * 2,  # Get more to account for chunked emails
            where=where,
            include=["documents", "metadatas", "distances"],
        )

        # Process results
        seen_emails = set()
        processed_results = []

        for i, doc_id in enumerate(results["ids"][0]):
            metadata = results["metadatas"][0][i]
            distance = results["distances"][0][i]
            document = results["documents"][0][i]

            # Get parent email ID for chunked documents
            email_id = metadata.get("parent_id", doc_id.split("_chunk_")[0])

            # Deduplicate by email
            if email_id in seen_emails:
                continue
            seen_emails.add(email_id)

            # Date filtering (post-query)
            if date_from or date_to:
                received_str = metadata.get("received_time", "")
                if received_str:
                    try:
                        received = datetime.fromisoformat(received_str)
                        # Remove timezone info for comparison (compare as naive datetime)
                        if received.tzinfo is not None:
                            received = received.replace(tzinfo=None)
                        if date_from and received < date_from:
                            continue
                        if date_to and received > date_to:
                            continue
                    except:
                        pass

            # Convert distance to similarity score (0-1, higher is better)
            similarity = 1 / (1 + distance)

            processed_results.append({
                "email_id": email_id,
                "entry_id": metadata.get("entry_id", ""),
                "subject": metadata.get("subject", ""),
                "sender_name": metadata.get("sender_name", ""),
                "sender_email": metadata.get("sender_email", ""),
                "folder_path": metadata.get("folder_path", ""),
                "received_time": metadata.get("received_time", ""),
                "has_attachments": metadata.get("has_attachments", False),
                "relevance_score": round(similarity, 4),
                "matched_text": document[:300] if document else "",
            })

            if len(processed_results) >= limit:
                break

        return processed_results

    def search_attachments(
        self,
        query: str,
        limit: int = 10,
    ) -> List[Dict[str, Any]]:
        """Search attachment content"""
        model = get_embedding_model()
        query_embedding = model.encode(query).tolist()

        results = self._attachments_collection.query(
            query_embeddings=[query_embedding],
            n_results=limit * 2,
            include=["documents", "metadatas", "distances"],
        )

        seen_attachments = set()
        processed_results = []

        for i, doc_id in enumerate(results["ids"][0]):
            metadata = results["metadatas"][0][i]
            distance = results["distances"][0][i]
            document = results["documents"][0][i]

            att_id = metadata.get("parent_id", doc_id.split("_chunk_")[0])

            if att_id in seen_attachments:
                continue
            seen_attachments.add(att_id)

            similarity = 1 / (1 + distance)

            processed_results.append({
                "attachment_id": att_id,
                "email_id": metadata.get("email_id", ""),
                "filename": metadata.get("filename", ""),
                "relevance_score": round(similarity, 4),
                "matched_text": document[:300] if document else "",
            })

            if len(processed_results) >= limit:
                break

        return processed_results

    def get_index_stats(self) -> Dict[str, Any]:
        """Get indexing statistics"""
        return {
            "emails_indexed": self._emails_collection.count(),
            "attachments_indexed": self._attachments_collection.count(),
            "db_path": self.config.db_path,
        }

    def clear_index(self):
        """Clear all indexed data"""
        self._client.delete_collection(self.COLLECTION_EMAILS)
        self._client.delete_collection(self.COLLECTION_ATTACHMENTS)

        # Recreate collections
        self._emails_collection = self._client.get_or_create_collection(
            name=self.COLLECTION_EMAILS,
        )
        self._attachments_collection = self._client.get_or_create_collection(
            name=self.COLLECTION_ATTACHMENTS,
        )


# Singleton
_indexer: Optional[EmailIndexer] = None


def get_indexer() -> EmailIndexer:
    """Get or create singleton EmailIndexer"""
    global _indexer
    if _indexer is None:
        _indexer = EmailIndexer()
    return _indexer


def run_full_index(
    since_days: Optional[int] = None,
    folders: Optional[List[str]] = None,
    progress_callback=None,
    clear_first: bool = False,
) -> Dict[str, Any]:
    """
    Run full email indexing

    Args:
        since_days: Only index emails from last N days (default: config value)
        folders: Folders to index (default: config value)
        progress_callback: Optional callback(current, total, message)
        clear_first: If True, clear the index before rebuilding

    Returns:
        Indexing statistics
    """
    # Import here to avoid circular imports and ensure fresh COM connection
    from .outlook_reader import OutlookReader
    
    config = get_config()
    
    # Indexing disabled - causes hanging in MCP server context
    # Run indexing separately using: python -m outlook_mcp.cli index
    indexer = None
    
    # Create a new reader instance for this thread (COM threading requirement)
    reader = OutlookReader()

    if since_days is None:
        since_days = config.index_period_days

    since_date = datetime.now() - timedelta(days=since_days)

    if folders is None:
        folders = config.folders_to_index

    stats = {
        "total_processed": 0,
        "total_indexed": 0,
        "total_skipped": 0,
        "errors": [],
        "start_time": datetime.now().isoformat(),
    }

    try:
        # Debug: check folders
        if folders == ["*"]:
            available_folders = reader.list_folders()
            stats["available_folders"] = [f["path"] for f in available_folders]
            folders = [f["path"] for f in available_folders]
        
        # Try only Inbox with very small limit for testing
        stats["folder_results"] = {}
        MAX_EMAILS = 100  # Increased limit
        
        folder_path = "Inbox"  # Only test Inbox
        folder_count = 0
        folder_errors = []
        
        try:
            stats["debug"] = ["Starting to get emails..."]
            email_count = 0
            for email in reader.get_emails(folder_path=folder_path, since_date=since_date, max_count=MAX_EMAILS):
                email_count += 1
                stats["debug"].append(f"Got email {email_count}: {email.subject[:30] if email else 'None'}")
                
                if email_count >= MAX_EMAILS:
                    break
                    
                stats["total_processed"] += 1
                folder_count += 1

                # Try indexing if indexer is available
                if indexer is not None:
                    try:
                        result = indexer.index_email(email)
                        if result.get("indexed"):
                            stats["total_indexed"] += 1
                        else:
                            stats["total_skipped"] += 1
                    except Exception as e:
                        stats["errors"].append({"error": str(e)[:100]})
                        stats["total_skipped"] += 1
                else:
                    stats["total_skipped"] += 1
                
        except Exception as e:
            folder_errors.append(str(e))
            stats["debug"].append(f"Error: {str(e)}")
        
        stats["folder_results"][folder_path] = {
            "count": folder_count,
            "errors": folder_errors
        }

    except Exception as e:
        stats["errors"].append({"error": str(e)})

    stats["end_time"] = datetime.now().isoformat()
    return stats
