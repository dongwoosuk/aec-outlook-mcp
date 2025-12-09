"""
Outlook MCP Server - Main MCP server implementation
"""

import asyncio
import json
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

# Load .env file from project root
from dotenv import load_dotenv
env_path = Path(__file__).parent.parent.parent.parent / ".env"
if env_path.exists():
    load_dotenv(env_path)

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from .config import get_config
from .outlook_reader import get_outlook_reader, OutlookReader
from .email_indexer import get_indexer, run_full_index, get_embedding_model
from .attachment_parser import (
    get_attachment_parser, extract_attachment_text,
    PYMUPDF_AVAILABLE, TESSERACT_AVAILABLE, ANTHROPIC_AVAILABLE, GOOGLE_GENAI_AVAILABLE
)


# Pre-load embedding model at startup (takes ~1 minute)
print("Loading embedding model... (this may take about 1 minute)", file=sys.stderr)
_embedding_model = get_embedding_model()
print("Embedding model loaded!", file=sys.stderr)

# Create MCP server
app = Server("outlook-mcp")


def _process_attachment_batch(collection, model, attachments_data, config):
    """Process a batch of attachments with batch embedding

    Args:
        collection: ChromaDB collection for attachments
        model: Embedding model
        attachments_data: List of (email_entry_id, filename, text_content) tuples
        config: Config object
    """
    import hashlib

    indexed = 0
    skipped = 0
    errors = []

    # Prepare documents
    documents = []
    valid_attachments = []

    for email_entry_id, filename, text_content in attachments_data:
        # Generate attachment ID
        att_id = hashlib.sha256(f"{email_entry_id}_{filename}".encode()).hexdigest()[:32]

        # Check if already indexed
        try:
            existing = collection.get(ids=[att_id])
            if existing and existing["ids"]:
                skipped += 1
                continue
        except:
            pass

        # Limit text length
        text = text_content[:10000] if text_content else ""
        if not text.strip():
            skipped += 1
            continue

        documents.append(f"Filename: {filename}\n\n{text}")
        valid_attachments.append((att_id, email_entry_id, filename))

    if not documents:
        return indexed, skipped, errors

    # Batch encode
    try:
        embeddings = model.encode(documents).tolist()

        for i, (att_id, email_entry_id, filename) in enumerate(valid_attachments):
            try:
                metadata = {
                    "email_entry_id": email_entry_id,
                    "filename": filename[:200],
                    "type": "attachment",
                }

                collection.add(
                    ids=[att_id],
                    embeddings=[embeddings[i]],
                    documents=[documents[i][:5000]],
                    metadatas=[metadata],
                )
                indexed += 1
            except Exception as e:
                errors.append(str(e)[:100])
                skipped += 1

    except Exception as e:
        errors.append(f"Attachment batch encoding error: {str(e)[:100]}")
        skipped += len(valid_attachments)

    return indexed, skipped, errors


def _process_email_batch(collection, model, emails):
    """Process a batch of emails with batch embedding"""
    import hashlib

    indexed = 0
    skipped = 0
    errors = []

    # Prepare documents for batch encoding
    documents = []
    valid_emails = []

    for email in emails:
        # Generate email ID
        email_id = hashlib.sha256(email.entry_id.encode()).hexdigest()[:32]

        # Check if already indexed
        try:
            existing = collection.get(ids=[email_id])
            if existing and existing["ids"]:
                skipped += 1
                continue
        except:
            pass

        # Prepare document
        parts = []
        if email.subject:
            parts.append(f"Subject: {email.subject}")
            parts.append(email.subject)
        if email.sender_name:
            parts.append(f"From: {email.sender_name}")
        if email.recipients:
            parts.append(f"To: {', '.join(email.recipients[:5])}")
        if email.received_time:
            try:
                parts.append(f"Date: {email.received_time.strftime('%Y-%m-%d')}")
            except:
                pass
        if email.body:
            parts.append(email.body[:2000])  # Limit body length

        document = "\n".join(parts)
        if document:
            documents.append(document)
            valid_emails.append((email, email_id))

    if not documents:
        return indexed, skipped, errors

    # Batch encode all documents
    try:
        embeddings = model.encode(documents).tolist()

        # Add to collection
        for i, (email, email_id) in enumerate(valid_emails):
            try:
                metadata = {
                    "entry_id": email.entry_id,
                    "subject": email.subject[:500] if email.subject else "",
                    "sender_name": email.sender_name[:200] if email.sender_name else "",
                    "sender_email": email.sender_email[:200] if email.sender_email else "",
                    "folder_path": email.folder_path[:200] if email.folder_path else "",
                    "received_time": email.received_time.isoformat() if email.received_time else "",
                    "has_attachments": email.has_attachments,
                }

                collection.add(
                    ids=[email_id],
                    embeddings=[embeddings[i]],
                    documents=[documents[i][:5000]],  # Limit document size
                    metadatas=[metadata],
                )
                indexed += 1
            except Exception as e:
                errors.append(str(e)[:100])
                skipped += 1

    except Exception as e:
        errors.append(f"Batch encoding error: {str(e)[:100]}")
        skipped += len(valid_emails)

    return indexed, skipped, errors


@app.list_tools()
async def list_tools() -> list[Tool]:
    """List available tools"""
    return [
        # Search Tools
        Tool(
            name="email_search",
            description="Search emails using natural language query. Returns relevant emails based on semantic similarity to your query.",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "Natural language search query (e.g., 'BIM coordination meeting', 'project deadline discussion')"
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum number of results (default: 10)",
                        "default": 10
                    },
                    "sender": {
                        "type": "string",
                        "description": "Optional: filter by sender name"
                    },
                    "days_back": {
                        "type": "integer",
                        "description": "Optional: only search emails from the last N days"
                    },
                    "folder": {
                        "type": "string",
                        "description": "Optional: filter by folder path"
                    }
                },
                "required": ["query"]
            }
        ),
        Tool(
            name="email_search_by_sender",
            description="Find all emails from a specific sender",
            inputSchema={
                "type": "object",
                "properties": {
                    "sender": {
                        "type": "string",
                        "description": "Sender name or email to search for"
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum results (default: 20)",
                        "default": 20
                    },
                    "query": {
                        "type": "string",
                        "description": "Optional: additional search terms to narrow results"
                    }
                },
                "required": ["sender"]
            }
        ),
        Tool(
            name="email_search_by_date",
            description="Search emails within a date range",
            inputSchema={
                "type": "object",
                "properties": {
                    "start_date": {
                        "type": "string",
                        "description": "Start date (YYYY-MM-DD format)"
                    },
                    "end_date": {
                        "type": "string",
                        "description": "End date (YYYY-MM-DD format)"
                    },
                    "query": {
                        "type": "string",
                        "description": "Optional: search query to filter results"
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum results (default: 20)",
                        "default": 20
                    }
                },
                "required": ["start_date"]
            }
        ),
        Tool(
            name="email_search_attachments",
            description="Search within email attachment contents (PDF, Word, Excel files)",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "Search query for attachment content"
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum results (default: 10)",
                        "default": 10
                    }
                },
                "required": ["query"]
            }
        ),

        # Management Tools
        Tool(
            name="email_index_status",
            description="Check the current email indexing status and statistics",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        Tool(
            name="email_index_refresh",
            description="Index new emails that haven't been indexed yet. This is incremental - only adds new emails.",
            inputSchema={
                "type": "object",
                "properties": {
                    "days_back": {
                        "type": "integer",
                        "description": "Number of days to look back for new emails (default: 7)",
                        "default": 7
                    }
                }
            }
        ),
        Tool(
            name="email_index_rebuild",
            description="Rebuild the entire email index. Warning: This can take a long time for large mailboxes.",
            inputSchema={
                "type": "object",
                "properties": {
                    "period": {
                        "type": "string",
                        "description": "Time period to index: '1month', '3months', '6months', '1year', '2years', '3years', 'all'",
                        "enum": ["1month", "3months", "6months", "1year", "2years", "3years", "all"],
                        "default": "1year"
                    },
                    "confirm": {
                        "type": "boolean",
                        "description": "Must be true to confirm rebuild"
                    },
                    "vision_provider": {
                        "type": "string",
                        "description": "Vision AI for image analysis: 'ocr' (default, free), 'gemini' (free, rate-limited), 'claude' (paid, best quality)",
                        "enum": ["ocr", "gemini", "claude"],
                        "default": "ocr"
                    }
                },
                "required": ["confirm"]
            }
        ),

        # Retrieval Tools
        Tool(
            name="email_get_detail",
            description="Get full details of a specific email by its ID",
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "description": "The email's entry ID (from search results)"
                    }
                },
                "required": ["entry_id"]
            }
        ),
        Tool(
            name="email_list_folders",
            description="List all email folders in Outlook",
            inputSchema={
                "type": "object",
                "properties": {
                    "include_system": {
                        "type": "boolean",
                        "description": "Include system/hidden folders (default: false)",
                        "default": False
                    }
                }
            }
        ),

        # Connection Tools
        Tool(
            name="email_connection_status",
            description="Check Outlook connection status",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
    ]


@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    """Handle tool calls"""
    try:
        # Search Tools
        if name == "email_search":
            query = arguments.get("query", "")
            limit = arguments.get("limit", 10)
            sender = arguments.get("sender")
            days_back = arguments.get("days_back")
            folder = arguments.get("folder")

            if not query:
                return [TextContent(type="text", text=json.dumps({
                    "success": False,
                    "error": "Query is required"
                }, indent=2))]

            # Semantic search using pre-loaded embedding model
            try:
                # Use pre-loaded model
                query_embedding = _embedding_model.encode(query).tolist()

                # Get ChromaDB collection
                import chromadb
                from chromadb.config import Settings

                config = get_config()
                client = chromadb.PersistentClient(
                    path=config.db_path,
                    settings=Settings(anonymized_telemetry=False),
                )
                collection = client.get_or_create_collection(name="emails")

                # Query with embedding
                search_results = collection.query(
                    query_embeddings=[query_embedding],
                    n_results=limit * 2,  # Get more to account for filtering
                    include=["documents", "metadatas", "distances"],
                )

                if not search_results or not search_results.get("ids") or not search_results["ids"][0]:
                    return [TextContent(type="text", text=json.dumps({
                        "success": True,
                        "query": query,
                        "result_count": 0,
                        "results": [],
                        "search_type": "semantic",
                        "note": "No indexed emails. Run indexing first."
                    }, indent=2))]

                # Process results
                results = []
                seen_entries = set()

                for i, doc_id in enumerate(search_results["ids"][0]):
                    metadata = search_results["metadatas"][0][i] if search_results.get("metadatas") else {}
                    document = search_results["documents"][0][i] if search_results.get("documents") else ""
                    distance = search_results["distances"][0][i] if search_results.get("distances") else 0

                    entry_id = metadata.get("entry_id", "")

                    # Skip duplicates (chunked emails)
                    if entry_id in seen_entries:
                        continue

                    subject = metadata.get("subject", "")
                    sender_name = metadata.get("sender_name", "")

                    # Apply sender filter
                    if sender and sender.lower() not in sender_name.lower():
                        continue

                    # Apply days filter
                    if days_back:
                        received_str = metadata.get("received_time", "")
                        if received_str:
                            try:
                                received = datetime.fromisoformat(received_str.replace("+00:00", ""))
                                cutoff = datetime.now() - timedelta(days=days_back)
                                if received < cutoff:
                                    continue
                            except:
                                pass

                    # Convert distance to similarity score (0-1, higher is better)
                    similarity = 1 / (1 + distance)

                    seen_entries.add(entry_id)
                    results.append({
                        "entry_id": entry_id,
                        "subject": subject,
                        "sender_name": sender_name,
                        "sender_email": metadata.get("sender_email", ""),
                        "received_time": metadata.get("received_time", ""),
                        "has_attachments": metadata.get("has_attachments", False),
                        "relevance_score": round(similarity, 4),
                        "preview": document[:200] if document else "",
                    })

                    if len(results) >= limit:
                        break

            except Exception as e:
                return [TextContent(type="text", text=json.dumps({
                    "success": False,
                    "error": str(e)
                }, indent=2))]

            return [TextContent(type="text", text=json.dumps({
                "success": True,
                "query": query,
                "search_type": "semantic",
                "result_count": len(results),
                "results": results
            }, indent=2, ensure_ascii=False))]

        elif name == "email_search_by_sender":
            sender = arguments.get("sender", "")
            limit = arguments.get("limit", 20)
            query = arguments.get("query", "")

            if not sender:
                return [TextContent(type="text", text=json.dumps({
                    "success": False,
                    "error": "Sender is required"
                }, indent=2))]

            # Use sender as part of query if no specific query given
            search_query = f"{query} from {sender}" if query else f"emails from {sender}"

            indexer = get_indexer()
            results = indexer.search_emails(
                query=search_query,
                limit=limit,
                sender_filter=sender,
            )

            return [TextContent(type="text", text=json.dumps({
                "success": True,
                "sender_filter": sender,
                "result_count": len(results),
                "results": results
            }, indent=2, ensure_ascii=False))]

        elif name == "email_search_by_date":
            start_date_str = arguments.get("start_date", "")
            end_date_str = arguments.get("end_date")
            query = arguments.get("query", "")
            limit = arguments.get("limit", 20)

            if not start_date_str:
                return [TextContent(type="text", text=json.dumps({
                    "success": False,
                    "error": "start_date is required (YYYY-MM-DD format)"
                }, indent=2))]

            try:
                date_from = datetime.strptime(start_date_str, "%Y-%m-%d")
                date_to = datetime.strptime(end_date_str, "%Y-%m-%d") if end_date_str else datetime.now()
            except ValueError:
                return [TextContent(type="text", text=json.dumps({
                    "success": False,
                    "error": "Invalid date format. Use YYYY-MM-DD"
                }, indent=2))]

            # If no query, use a generic one
            search_query = query if query else "email"

            indexer = get_indexer()
            results = indexer.search_emails(
                query=search_query,
                limit=limit,
                date_from=date_from,
                date_to=date_to,
            )

            return [TextContent(type="text", text=json.dumps({
                "success": True,
                "date_range": {
                    "from": start_date_str,
                    "to": end_date_str or "now"
                },
                "result_count": len(results),
                "results": results
            }, indent=2, ensure_ascii=False))]

        elif name == "email_search_attachments":
            query = arguments.get("query", "")
            limit = arguments.get("limit", 10)

            if not query:
                return [TextContent(type="text", text=json.dumps({
                    "success": False,
                    "error": "Query is required"
                }, indent=2))]

            # Use pre-loaded embedding model for semantic search
            try:
                import chromadb
                from chromadb.config import Settings

                config = get_config()
                client = chromadb.PersistentClient(
                    path=config.db_path,
                    settings=Settings(anonymized_telemetry=False),
                )

                try:
                    attachments_collection = client.get_collection(name="attachments")
                except:
                    return [TextContent(type="text", text=json.dumps({
                        "success": True,
                        "query": query,
                        "result_count": 0,
                        "results": [],
                        "note": "No attachments indexed. Run index rebuild first."
                    }, indent=2))]

                # Generate query embedding using pre-loaded model
                query_embedding = _embedding_model.encode(query).tolist()

                # Query attachments collection
                search_results = attachments_collection.query(
                    query_embeddings=[query_embedding],
                    n_results=limit * 2,
                    include=["documents", "metadatas", "distances"],
                )

                if not search_results or not search_results.get("ids") or not search_results["ids"][0]:
                    return [TextContent(type="text", text=json.dumps({
                        "success": True,
                        "query": query,
                        "result_count": 0,
                        "results": [],
                        "note": "No matching attachments found."
                    }, indent=2))]

                # Process results
                results = []
                seen_attachments = set()

                for i, att_id in enumerate(search_results["ids"][0]):
                    metadata = search_results["metadatas"][0][i] if search_results.get("metadatas") else {}
                    document = search_results["documents"][0][i] if search_results.get("documents") else ""
                    distance = search_results["distances"][0][i] if search_results.get("distances") else 0

                    # Deduplicate
                    if att_id in seen_attachments:
                        continue
                    seen_attachments.add(att_id)

                    # Convert distance to similarity score
                    similarity = 1 / (1 + distance)

                    results.append({
                        "attachment_id": att_id,
                        "email_entry_id": metadata.get("email_entry_id", ""),
                        "filename": metadata.get("filename", ""),
                        "relevance_score": round(similarity, 4),
                        "matched_text": document[:300] if document else "",
                    })

                    if len(results) >= limit:
                        break

            except Exception as e:
                return [TextContent(type="text", text=json.dumps({
                    "success": False,
                    "error": str(e)
                }, indent=2))]

            return [TextContent(type="text", text=json.dumps({
                "success": True,
                "query": query,
                "search_type": "semantic",
                "result_count": len(results),
                "results": results
            }, indent=2, ensure_ascii=False))]

        # Management Tools
        elif name == "email_index_status":
            indexer = get_indexer()
            stats = indexer.get_index_stats()

            # Add Outlook status
            try:
                reader = get_outlook_reader()
                outlook_status = reader.get_connection_status()
            except Exception as e:
                outlook_status = {"connected": False, "error": str(e)}

            return [TextContent(type="text", text=json.dumps({
                "success": True,
                "index_stats": stats,
                "outlook_status": outlook_status,
                "config": {
                    "db_path": get_config().db_path,
                    "embedding_model": get_config().embedding_model,
                }
            }, indent=2))]

        elif name == "email_index_refresh":
            days_back = arguments.get("days_back", 7)

            # Run indexing synchronously (to_thread causes issues with COM/embedding)
            result = run_full_index(since_days=days_back)

            return [TextContent(type="text", text=json.dumps({
                "success": True,
                "message": f"Indexed {result.get('total_indexed', 0)} new emails",
                "details": result
            }, indent=2))]

        elif name == "email_index_rebuild":
            confirm = arguments.get("confirm", False)
            period = arguments.get("period", "1year")
            vision_provider = arguments.get("vision_provider", "ocr")

            # Convert period to days
            period_map = {
                "1month": 30,
                "3months": 90,
                "6months": 180,
                "1year": 365,
                "2years": 730,
                "3years": 1095,
                "all": None  # None means no limit
            }
            days_back = period_map.get(period, 365)

            if not confirm:
                return [TextContent(type="text", text=json.dumps({
                    "success": False,
                    "error": "Must set confirm=true to rebuild index",
                    "available_periods": list(period_map.keys())
                }, indent=2))]

            # Run indexing with pre-loaded model and progress tracking
            import hashlib
            import chromadb
            from chromadb.config import Settings
            from .outlook_reader import OutlookReader

            config = get_config()

            # Initialize ChromaDB
            client = chromadb.PersistentClient(
                path=config.db_path,
                settings=Settings(anonymized_telemetry=False),
            )

            # Clear existing index
            try:
                client.delete_collection("emails")
            except:
                pass
            try:
                client.delete_collection("attachments")
            except:
                pass
            collection = client.get_or_create_collection(name="emails")
            attachments_collection = client.get_or_create_collection(name="attachments")

            # Connect to Outlook
            reader = OutlookReader()
            status = reader.get_connection_status()
            if not status.get("connected"):
                return [TextContent(type="text", text=json.dumps({
                    "success": False,
                    "error": f"Could not connect to Outlook: {status.get('error')}"
                }, indent=2))]

            since_date = datetime.now() - timedelta(days=days_back) if days_back else None
            folders = ["Inbox", "Sent Items"]

            # Progress tracking
            batch_size = 50
            total_processed = 0
            total_indexed = 0
            total_skipped = 0
            errors = []
            progress_log = []

            # Attachment tracking
            att_total = 0
            att_indexed = 0
            att_skipped = 0
            attachment_batch = []  # (email_entry_id, filename, text_content)

            # Create temp directory for attachment extraction
            import tempfile
            import os
            temp_dir = tempfile.mkdtemp(prefix="outlook_mcp_att_")
            parser = get_attachment_parser()

            # Supported attachment extensions
            supported_exts = {".pdf", ".docx", ".xlsx", ".txt", ".png", ".jpg", ".jpeg"}

            progress_log.append(f"Attachment parsing: PDF={PYMUPDF_AVAILABLE}, OCR={TESSERACT_AVAILABLE}, Claude={ANTHROPIC_AVAILABLE}, Gemini={GOOGLE_GENAI_AVAILABLE}")
            progress_log.append(f"Vision provider: {vision_provider}")

            for folder_path in folders:
                progress_log.append(f"Processing folder: {folder_path}")
                email_batch = []

                try:
                    for email in reader.get_emails(folder_path=folder_path, since_date=since_date):
                        total_processed += 1
                        email_batch.append(email)

                        # Process attachments for this email
                        if email.has_attachments and email.attachments:
                            for att_idx, att_info in enumerate(email.attachments):
                                filename = att_info.get("filename", "")
                                ext = os.path.splitext(filename)[1].lower()

                                if ext not in supported_exts:
                                    continue

                                att_total += 1

                                try:
                                    # Save attachment to temp directory
                                    saved_path = reader.save_attachment(
                                        email.entry_id, att_idx, temp_dir
                                    )

                                    if saved_path:
                                        # Extract text (use vision_provider for images)
                                        result = parser.parse_file(saved_path, vision_provider=vision_provider)
                                        if result.get("success") and result.get("text"):
                                            attachment_batch.append((
                                                email.entry_id,
                                                filename,
                                                result["text"]
                                            ))

                                        # Clean up temp file
                                        try:
                                            os.remove(saved_path)
                                        except:
                                            pass
                                except Exception as e:
                                    errors.append(f"Attachment {filename}: {str(e)[:50]}")

                        # Process email batch when full
                        if len(email_batch) >= batch_size:
                            indexed, skipped, batch_errors = _process_email_batch(
                                collection, _embedding_model, email_batch
                            )
                            total_indexed += indexed
                            total_skipped += skipped
                            errors.extend(batch_errors)
                            progress_log.append(
                                f"  Batch complete: {total_indexed} emails indexed (total: {total_processed})"
                            )
                            email_batch = []

                        # Process attachment batch when full
                        if len(attachment_batch) >= 20:
                            a_indexed, a_skipped, a_errors = _process_attachment_batch(
                                attachments_collection, _embedding_model, attachment_batch, config
                            )
                            att_indexed += a_indexed
                            att_skipped += a_skipped
                            errors.extend(a_errors)
                            progress_log.append(
                                f"  Attachments batch: {att_indexed} indexed"
                            )
                            attachment_batch = []

                    # Process remaining emails
                    if email_batch:
                        indexed, skipped, batch_errors = _process_email_batch(
                            collection, _embedding_model, email_batch
                        )
                        total_indexed += indexed
                        total_skipped += skipped
                        errors.extend(batch_errors)
                        progress_log.append(
                            f"  Final batch: {total_indexed} emails indexed (total: {total_processed})"
                        )

                except Exception as e:
                    errors.append(f"Folder {folder_path}: {str(e)[:100]}")
                    progress_log.append(f"  Error: {str(e)[:100]}")

            # Process remaining attachments
            if attachment_batch:
                a_indexed, a_skipped, a_errors = _process_attachment_batch(
                    attachments_collection, _embedding_model, attachment_batch, config
                )
                att_indexed += a_indexed
                att_skipped += a_skipped
                errors.extend(a_errors)

            # Clean up temp directory
            try:
                import shutil
                shutil.rmtree(temp_dir, ignore_errors=True)
            except:
                pass

            progress_log.append(f"=== Indexing Complete ===")
            progress_log.append(f"Emails: {total_indexed} indexed, {total_skipped} skipped")
            progress_log.append(f"Attachments: {att_indexed} indexed out of {att_total} found")

            period_label = "all time" if period == "all" else period
            return [TextContent(type="text", text=json.dumps({
                "success": True,
                "message": f"Rebuilt index with {total_indexed} emails and {att_indexed} attachments ({period_label})",
                "period": period,
                "stats": {
                    "emails_processed": total_processed,
                    "emails_indexed": total_indexed,
                    "emails_skipped": total_skipped,
                    "attachments_found": att_total,
                    "attachments_indexed": att_indexed,
                    "attachments_skipped": att_skipped,
                    "errors_count": len(errors),
                },
                "progress_log": progress_log,
                "errors": errors[:10] if errors else []
            }, indent=2, ensure_ascii=False))]

        # Retrieval Tools
        elif name == "email_get_detail":
            entry_id = arguments.get("entry_id", "")

            if not entry_id:
                return [TextContent(type="text", text=json.dumps({
                    "success": False,
                    "error": "entry_id is required"
                }, indent=2))]

            reader = get_outlook_reader()
            email = reader.get_email_by_id(entry_id)

            if email is None:
                return [TextContent(type="text", text=json.dumps({
                    "success": False,
                    "error": "Email not found"
                }, indent=2))]

            return [TextContent(type="text", text=json.dumps({
                "success": True,
                "email": email.to_dict()
            }, indent=2, ensure_ascii=False))]

        elif name == "email_list_folders":
            include_system = arguments.get("include_system", False)

            reader = get_outlook_reader()
            folders = reader.list_folders(include_system=include_system)

            return [TextContent(type="text", text=json.dumps({
                "success": True,
                "folder_count": len(folders),
                "folders": folders
            }, indent=2, ensure_ascii=False))]

        elif name == "email_connection_status":
            reader = get_outlook_reader()
            status = reader.get_connection_status()

            return [TextContent(type="text", text=json.dumps({
                "success": True,
                "status": status
            }, indent=2))]

        else:
            return [TextContent(type="text", text=json.dumps({
                "success": False,
                "error": f"Unknown tool: {name}"
            }, indent=2))]

    except Exception as e:
        return [TextContent(type="text", text=json.dumps({
            "success": False,
            "error": str(e)
        }, indent=2))]


async def main():
    """Run the server"""
    async with stdio_server() as (read_stream, write_stream):
        await app.run(
            read_stream,
            write_stream,
            app.create_initialization_options()
        )


if __name__ == "__main__":
    asyncio.run(main())
