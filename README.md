# 📧 Outlook MCP - AEC-Specialized AI Email Intelligence

An MCP (Model Context Protocol) server that transforms Outlook email management with **RAG-based semantic search** and **Vision AI for architectural document analysis** - built specifically for Architecture, Engineering & Construction workflows.

![Version](https://img.shields.io/badge/version-0.1.0--beta-orange)
![Status](https://img.shields.io/badge/status-beta-yellow)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)

## 🎯 Built for AEC Professionals

Unlike generic email tools, this is designed for the unique challenges of architectural practice:

### 🔍 Semantic Search with RAG
- **ChromaDB vector database** for intelligent email indexing
- **Natural language queries**: "Find emails about the BIM coordination meeting"
- **Attachment content search**: Search inside PDFs, Word docs, Excel files

### 👁️ Multi-Modal Vision AI
- **Claude Vision API** for architectural drawing analysis
- **Gemini Vision API** for technical document processing
- **Specialized prompts** for floor plans, elevations, sections, details
- **OCR fallback** with Tesseract (Korean + English)

### 🏗️ AEC-Specific Features
- **Bilingual support**: Korean + English for international projects
- **Technical drawing analysis**: Understands architectural documents
- **RFI/Submittal parsing**: Extract key information from construction docs
- **100% Local processing**: Email data never leaves your machine

## 📋 Features Overview

| Feature | Description |
|---------|-------------|
| **Semantic Search** | Natural language email queries via RAG |
| **Vision AI** | Claude + Gemini for image/PDF analysis |
| **Attachment Parsing** | PDF, Word, Excel, Images |
| **Date/Sender Filters** | Metadata-based search |
| **Folder Navigation** | Access all Outlook folders |
| **Local Processing** | Complete privacy, no cloud upload |

## 🛠️ Installation

```bash
git clone https://github.com/dongwoosuk/outlook-mcp.git
cd outlook-mcp
python -m venv .venv
.venv\Scripts\activate  # Windows
pip install -e .
```

For OCR support:
```bash
pip install -e ".[ocr]"
```

For Vision AI:
```bash
pip install anthropic google-generativeai
```

## ⚙️ Configuration

### Environment Variables (Optional)

```bash
# For Vision AI features
set ANTHROPIC_API_KEY=your_claude_api_key
set GOOGLE_API_KEY=your_gemini_api_key
```

### Claude Desktop Configuration

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "path/to/outlook_mcp/.venv/Scripts/python.exe",
      "args": ["-m", "outlook_mcp"],
      "cwd": "path/to/outlook_mcp"
    }
  }
}
```

## 📚 Available Tools

| Tool | Description |
|------|-------------|
| `email_search` | Natural language semantic search |
| `email_search_by_sender` | Filter by sender |
| `email_search_by_date` | Filter by date range |
| `email_search_attachments` | Search attachment contents |
| `email_index_status` | Check indexing progress |
| `email_index_refresh` | Index new emails |
| `email_get_detail` | Get full email details |
| `email_list_folders` | List Outlook folders |

## 💡 Usage Examples

### In Claude Desktop:

```
"Find emails from the structural engineer about beam calculations"

"Search for RFI responses from last week"

"Find all emails with PDF attachments about the facade design"

"What did the contractor say about the schedule delay?"
```

### Vision AI for Attachments:

```
"Analyze the floor plan attached in John's email"

"What dimensions are shown in this elevation drawing?"

"Extract the specification details from the attached PDF"
```

## 🏗️ Architecture

```
outlook_mcp/
├── outlook_mcp/
│   ├── server.py           # Main MCP server
│   ├── outlook_reader.py   # Win32com Outlook integration
│   ├── email_indexer.py    # ChromaDB RAG indexing
│   ├── attachment_parser.py # PDF/Word/Excel/Vision AI
│   ├── config.py           # Configuration management
│   └── cli.py              # Command-line interface
└── pyproject.toml
```

## 🔧 Technical Stack

| Component | Technology |
|-----------|------------|
| **Email Access** | win32com (Outlook COM API) |
| **Vector DB** | ChromaDB |
| **Embeddings** | sentence-transformers (all-MiniLM-L6-v2) |
| **PDF Parsing** | PyMuPDF |
| **Word Parsing** | python-docx |
| **Excel Parsing** | openpyxl |
| **OCR** | Tesseract (kor+eng) |
| **Vision AI** | Anthropic Claude, Google Gemini |

## 🎯 Use Cases

### For Project Architects (PA)
- Quickly find client feedback across hundreds of emails
- Search RFI responses and submittal approvals
- Track consultant coordination threads

### For Project Managers (PM)
- Search contract and schedule discussions
- Find meeting notes and action items
- Track change order communications

### For Designers
- Find design review comments
- Search for reference images and inspiration
- Locate specification discussions

### Future: Office-Wide Deployment
- Centralized email intelligence for entire teams
- Shared knowledge base across projects
- AI-powered email analytics and reporting

## ⚠️ Requirements

- **Windows** with Outlook Desktop app installed and logged in
- **Python 3.10+**
- Outlook must be running for email access

## 🔒 Privacy & Security

- **100% local processing**: Emails are indexed locally in ChromaDB
- **No cloud upload**: Email content never leaves your machine
- **API keys optional**: Vision AI features require API keys, but core search works without them
- **Data location**: `C:\Users\{USERNAME}\Documents\OutlookMCP\`

## 📄 License

MIT License - see [LICENSE](LICENSE) file.

## 🙏 Acknowledgments

- Built on the [Model Context Protocol](https://github.com/anthropics/anthropic-cookbook/tree/main/misc/model_context_protocol) by Anthropic
- ChromaDB for vector storage
- sentence-transformers for embeddings

## 📬 Contact

Dongwoo Suk - Computational Design Specialist

- GitHub: [dongwoosuk](https://github.com/dongwoosuk)
- LinkedIn: [dongwoosuk](https://www.linkedin.com/in/dongwoosuk/)
