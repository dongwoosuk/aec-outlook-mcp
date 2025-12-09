"""
CLI for Outlook MCP - Run indexing separately from MCP server
"""

import sys
import argparse
from datetime import datetime, timedelta

def run_index(days: int = 365, clear: bool = False, batch_size: int = 50):
    """Run email indexing with batch processing"""
    print(f"Starting email indexing (last {days} days, batch size: {batch_size})...")
    
    # Import here to avoid slow startup
    from .outlook_reader import OutlookReader
    from .email_indexer import get_indexer, get_embedding_model
    from .config import get_config
    import hashlib
    
    config = get_config()
    
    print("Loading embedding model (this may take a minute)...")
    model = get_embedding_model()
    print("Model loaded!")
    
    print("Initializing indexer...")
    indexer = get_indexer()
    
    if clear:
        print("Clearing existing index...")
        indexer.clear_index()
    
    print("Connecting to Outlook...")
    reader = OutlookReader()
    
    status = reader.get_connection_status()
    if not status.get("connected"):
        print(f"Error: Could not connect to Outlook: {status.get('error')}")
        return
    
    print(f"Connected to: {status.get('account')}")
    print(f"Inbox count: {status.get('inbox_count')}")
    
    since_date = datetime.now() - timedelta(days=days)
    folders = ["Inbox", "Sent Items"]
    
    total_processed = 0
    total_indexed = 0
    total_skipped = 0
    errors = []
    
    for folder_path in folders:
        print(f"\nProcessing folder: {folder_path}")
        
        # Collect emails in batches
        email_batch = []
        
        try:
            for email in reader.get_emails(folder_path=folder_path, since_date=since_date):
                total_processed += 1
                email_batch.append(email)
                
                # Process batch when full
                if len(email_batch) >= batch_size:
                    indexed, skipped, batch_errors = process_batch(indexer, model, email_batch, config)
                    total_indexed += indexed
                    total_skipped += skipped
                    errors.extend(batch_errors)
                    print(f"  Processed batch: {total_indexed} indexed, {total_skipped} skipped")
                    email_batch = []
            
            # Process remaining emails
            if email_batch:
                indexed, skipped, batch_errors = process_batch(indexer, model, email_batch, config)
                total_indexed += indexed
                total_skipped += skipped
                errors.extend(batch_errors)
                print(f"  Processed final batch: {total_indexed} indexed, {total_skipped} skipped")
                    
        except Exception as e:
            print(f"  Error accessing folder: {e}")
    
    print(f"\n=== Indexing Complete ===")
    print(f"Processed: {total_processed}")
    print(f"Indexed: {total_indexed}")
    print(f"Skipped: {total_skipped}")
    if errors:
        print(f"Errors: {len(errors)}")
        for err in errors[:5]:
            print(f"  - {err}")


def process_batch(indexer, model, emails, config):
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
            existing = indexer._emails_collection.get(ids=[email_id])
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
            parts.append(f"Date: {email.received_time.strftime('%Y-%m-%d')}")
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
                
                indexer._emails_collection.add(
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


def main():
    parser = argparse.ArgumentParser(description="Outlook MCP CLI")
    subparsers = parser.add_subparsers(dest="command")
    
    # Index command
    index_parser = subparsers.add_parser("index", help="Run email indexing")
    index_parser.add_argument("--days", type=int, default=365, help="Days to index (default: 365)")
    index_parser.add_argument("--clear", action="store_true", help="Clear existing index first")
    
    args = parser.parse_args()
    
    if args.command == "index":
        run_index(days=args.days, clear=args.clear)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
