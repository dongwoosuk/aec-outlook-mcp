"""
Outlook Reader - Access Outlook emails via win32com
"""

import os
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional, Generator
from dataclasses import dataclass, field
from pathlib import Path

# win32com imports
try:
    import win32com.client
    import pythoncom
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False


@dataclass
class EmailMessage:
    """Represents an email message"""
    entry_id: str
    subject: str
    sender_name: str
    sender_email: str
    recipients: List[str]
    cc: List[str]
    received_time: datetime
    sent_time: Optional[datetime]
    body: str
    html_body: str
    folder_path: str
    has_attachments: bool
    attachments: List[Dict[str, Any]] = field(default_factory=list)
    conversation_id: str = ""
    importance: int = 1  # 0=Low, 1=Normal, 2=High
    is_read: bool = True

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization"""
        return {
            "entry_id": self.entry_id,
            "subject": self.subject,
            "sender_name": self.sender_name,
            "sender_email": self.sender_email,
            "recipients": self.recipients,
            "cc": self.cc,
            "received_time": self.received_time.isoformat() if self.received_time else None,
            "sent_time": self.sent_time.isoformat() if self.sent_time else None,
            "body": self.body,
            "folder_path": self.folder_path,
            "has_attachments": self.has_attachments,
            "attachments": self.attachments,
            "conversation_id": self.conversation_id,
            "importance": self.importance,
            "is_read": self.is_read,
        }


class OutlookReader:
    """Read emails from Outlook via COM API"""

    # Outlook folder type constants
    FOLDER_INBOX = 6
    FOLDER_SENT = 5
    FOLDER_DRAFTS = 16
    FOLDER_DELETED = 3
    FOLDER_OUTBOX = 4
    FOLDER_CALENDAR = 9
    FOLDER_CONTACTS = 10

    def __init__(self):
        if not WIN32COM_AVAILABLE:
            raise ImportError(
                "pywin32 is not installed. Install with: pip install pywin32"
            )

        self._outlook = None
        self._namespace = None

    def _ensure_connection(self):
        """Ensure Outlook connection is established"""
        # Only initialize if not already connected
        if self._outlook is not None and self._namespace is not None:
            return
            
        # Initialize COM for the current thread
        try:
            pythoncom.CoInitialize()
        except:
            pass  # Already initialized for this thread
        
        # Create connection
        try:
            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")
        except Exception as e:
            raise ConnectionError(
                f"Could not connect to Outlook. Is Outlook running? Error: {e}"
            )

    def get_connection_status(self) -> Dict[str, Any]:
        """Check Outlook connection status"""
        try:
            self._ensure_connection()
            # Try to access inbox to verify connection
            inbox = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
            return {
                "connected": True,
                "inbox_count": inbox.Items.Count,
                "account": self._namespace.CurrentUser.Name,
            }
        except Exception as e:
            return {
                "connected": False,
                "error": str(e),
            }

    def list_folders(self, include_system: bool = False) -> List[Dict[str, Any]]:
        """List all mail folders using GetDefaultFolder (avoids Stores access issues)"""
        self._ensure_connection()
        folders = []

        # Use GetDefaultFolder to avoid Stores permission issues
        default_folders = [
            (self.FOLDER_INBOX, "Inbox"),
            (self.FOLDER_SENT, "Sent Items"),
            (self.FOLDER_DRAFTS, "Drafts"),
            (self.FOLDER_DELETED, "Deleted Items"),
        ]

        for folder_type, folder_name in default_folders:
            try:
                folder = self._namespace.GetDefaultFolder(folder_type)
                folders.append({
                    "name": folder.Name,
                    "path": folder.Name,
                    "count": folder.Items.Count,
                    "unread": folder.UnReadItemCount if hasattr(folder, "UnReadItemCount") else 0,
                })
                
                # Also traverse subfolders of default folders
                if not include_system:
                    try:
                        for subfolder in folder.Folders:
                            if not subfolder.Name.startswith("$"):
                                folders.append({
                                    "name": subfolder.Name,
                                    "path": f"{folder.Name}/{subfolder.Name}",
                                    "count": subfolder.Items.Count,
                                    "unread": subfolder.UnReadItemCount if hasattr(subfolder, "UnReadItemCount") else 0,
                                })
                    except:
                        pass
            except Exception as e:
                continue

        return folders

    def _get_folder_by_path(self, folder_path: str):
        """Get a folder by its path"""
        self._ensure_connection()

        # Handle default folders
        if folder_path.lower() in ["inbox", "받은편지함"]:
            return self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
        elif folder_path.lower() in ["sent", "sent items", "보낸편지함"]:
            return self._namespace.GetDefaultFolder(self.FOLDER_SENT)
        elif folder_path.lower() in ["drafts", "임시보관함"]:
            return self._namespace.GetDefaultFolder(self.FOLDER_DRAFTS)
        elif folder_path.lower() in ["deleted", "deleted items", "지운편지함"]:
            return self._namespace.GetDefaultFolder(self.FOLDER_DELETED)

        # Navigate folder path
        parts = folder_path.split("/")

        # Find the store/root folder
        current_folder = None
        for store in self._namespace.Stores:
            try:
                root = store.GetRootFolder()
                if root.Name == parts[0]:
                    current_folder = root
                    break
            except:
                continue

        if current_folder is None:
            raise ValueError(f"Could not find folder: {folder_path}")

        # Navigate to the target folder
        for part in parts[1:]:
            found = False
            for subfolder in current_folder.Folders:
                if subfolder.Name == part:
                    current_folder = subfolder
                    found = True
                    break
            if not found:
                raise ValueError(f"Could not find folder: {folder_path}")

        return current_folder

    def _parse_email(self, item, folder_path: str) -> Optional[EmailMessage]:
        """Parse an Outlook mail item into EmailMessage"""
        try:
            # Get basic properties
            entry_id = item.EntryID
            subject = item.Subject or "(No Subject)"

            # Sender info
            sender_name = ""
            sender_email = ""
            try:
                sender_name = item.SenderName or ""
                sender_email = item.SenderEmailAddress or ""
                # Clean up Exchange addresses
                if sender_email.startswith("/O="):
                    try:
                        sender_email = item.Sender.GetExchangeUser().PrimarySmtpAddress
                    except:
                        pass
            except:
                pass

            # Recipients
            recipients = []
            try:
                for recipient in item.Recipients:
                    recipients.append(recipient.Name)
            except:
                pass

            # CC
            cc = []
            try:
                cc_str = item.CC or ""
                if cc_str:
                    cc = [c.strip() for c in cc_str.split(";")]
            except:
                pass

            # Times
            received_time = None
            sent_time = None
            try:
                received_time = item.ReceivedTime
                if hasattr(received_time, "replace"):  # datetime object
                    pass
                else:
                    received_time = datetime.now()
            except:
                received_time = datetime.now()

            try:
                sent_time = item.SentOn
            except:
                pass

            # Body
            body = ""
            html_body = ""
            try:
                body = item.Body or ""
                html_body = item.HTMLBody or ""
            except:
                pass

            # Attachments
            has_attachments = False
            attachments = []
            try:
                if item.Attachments.Count > 0:
                    has_attachments = True
                    for att in item.Attachments:
                        att_info = {
                            "filename": att.FileName,
                            "size": att.Size if hasattr(att, "Size") else 0,
                            "type": att.Type,  # 1=file, 5=embedded, 6=OLE
                        }
                        attachments.append(att_info)
            except:
                pass

            # Conversation
            conversation_id = ""
            try:
                conversation_id = item.ConversationID or ""
            except:
                pass

            # Importance (0=Low, 1=Normal, 2=High)
            importance = 1
            try:
                importance = item.Importance
            except:
                pass

            # Read status
            is_read = True
            try:
                is_read = not item.UnRead
            except:
                pass

            return EmailMessage(
                entry_id=entry_id,
                subject=subject,
                sender_name=sender_name,
                sender_email=sender_email,
                recipients=recipients,
                cc=cc,
                received_time=received_time,
                sent_time=sent_time,
                body=body,
                html_body=html_body,
                folder_path=folder_path,
                has_attachments=has_attachments,
                attachments=attachments,
                conversation_id=conversation_id,
                importance=importance,
                is_read=is_read,
            )

        except Exception as e:
            # Skip problematic emails
            return None

    def get_emails(
        self,
        folder_path: str = "inbox",
        since_date: Optional[datetime] = None,
        max_count: Optional[int] = None,
        include_body: bool = True,
    ) -> Generator[EmailMessage, None, None]:
        """
        Get emails from a folder

        Args:
            folder_path: Folder path or name (e.g., "inbox", "Sent Items")
            since_date: Only get emails after this date
            max_count: Maximum number of emails to retrieve
            include_body: Whether to include email body (slower if True)

        Yields:
            EmailMessage objects
        """
        self._ensure_connection()

        folder = self._get_folder_by_path(folder_path)
        items = folder.Items
        items.Sort("[ReceivedTime]", True)  # Sort by newest first

        count = 0
        for item in items:
            # Check max count
            if max_count and count >= max_count:
                break

            # Check message class (only mail items)
            try:
                if item.Class != 43:  # 43 = olMail
                    continue
            except:
                continue

            # Check date filter
            if since_date:
                try:
                    received = item.ReceivedTime
                    # Convert COM datetime to Python datetime for comparison
                    if hasattr(received, 'year'):
                        received_dt = datetime(received.year, received.month, received.day,
                                              received.hour, received.minute, received.second)
                        # Make since_date timezone-naive for comparison
                        since_date_naive = since_date.replace(tzinfo=None) if hasattr(since_date, 'tzinfo') and since_date.tzinfo else since_date
                        if received_dt < since_date_naive:
                            break  # Since sorted by date, we can stop here
                except Exception:
                    continue

            # Parse email
            email = self._parse_email(item, folder_path)
            if email:
                count += 1
                yield email

    def get_all_emails(
        self,
        folders: Optional[List[str]] = None,
        since_date: Optional[datetime] = None,
        max_per_folder: Optional[int] = None,
    ) -> Generator[EmailMessage, None, None]:
        """
        Get emails from multiple folders

        Args:
            folders: List of folder paths, or None for all folders
            since_date: Only get emails after this date
            max_per_folder: Maximum emails per folder

        Yields:
            EmailMessage objects
        """
        if folders is None or folders == ["*"]:
            # Get all mail folders
            all_folders = self.list_folders()
            folders = [f["path"] for f in all_folders]

        for folder_path in folders:
            try:
                for email in self.get_emails(
                    folder_path=folder_path,
                    since_date=since_date,
                    max_count=max_per_folder,
                ):
                    yield email
            except Exception as e:
                # Skip folders that can't be accessed
                continue

    def get_email_by_id(self, entry_id: str) -> Optional[EmailMessage]:
        """Get a specific email by its Entry ID"""
        self._ensure_connection()

        try:
            item = self._namespace.GetItemFromID(entry_id)
            return self._parse_email(item, "")
        except Exception as e:
            return None

    def save_attachment(
        self,
        entry_id: str,
        attachment_index: int,
        save_path: str,
    ) -> Optional[str]:
        """
        Save an attachment to disk

        Args:
            entry_id: Email entry ID
            attachment_index: Index of attachment (0-based)
            save_path: Directory to save to

        Returns:
            Full path to saved file, or None if failed
        """
        self._ensure_connection()

        try:
            item = self._namespace.GetItemFromID(entry_id)

            if attachment_index >= item.Attachments.Count:
                return None

            attachment = item.Attachments.Item(attachment_index + 1)  # 1-based
            filename = attachment.FileName

            # Ensure save directory exists
            Path(save_path).mkdir(parents=True, exist_ok=True)

            full_path = os.path.join(save_path, filename)

            # Handle duplicate filenames
            base, ext = os.path.splitext(filename)
            counter = 1
            while os.path.exists(full_path):
                full_path = os.path.join(save_path, f"{base}_{counter}{ext}")
                counter += 1

            attachment.SaveAsFile(full_path)
            return full_path

        except Exception as e:
            return None

    def close(self):
        """Clean up COM resources"""
        self._outlook = None
        self._namespace = None
        try:
            pythoncom.CoUninitialize()
        except:
            pass


# Singleton instance
_reader: Optional[OutlookReader] = None


def get_outlook_reader() -> OutlookReader:
    """Get or create singleton OutlookReader instance"""
    global _reader
    if _reader is None:
        _reader = OutlookReader()
    return _reader
