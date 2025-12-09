# Outlook MCP Server

로컬 Outlook 이메일을 벡터 검색하고 Claude Desktop에서 사용할 수 있는 MCP 서버

## 기능

- **자연어 이메일 검색**: "BIM 미팅 관련 이메일 찾아줘"
- **첨부파일 내용 검색**: PDF, Word 문서 내용까지 검색
- **발신자/날짜 필터링**: 메타데이터 기반 검색
- **완전 로컬 처리**: 이메일 데이터가 외부로 나가지 않음

## 아키텍처

```
[Outlook Desktop] ← win32com
       ↓
[Email Parser] → 텍스트 + 첨부파일 추출
       ↓
[sentence-transformers] ← all-MiniLM-L6-v2 (로컬 임베딩)
       ↓
[ChromaDB] ← 로컬 벡터 저장
       ↓
[MCP Server] → Claude Desktop
```

## 설치

```bash
cd mcp_servers/outlook_mcp
python -m venv .venv
.venv\Scripts\activate
pip install -e .
```

OCR 기능 (이미지 텍스트 추출)을 원하면:
```bash
pip install -e ".[ocr]"
```

## 설정

데이터 저장 위치: `C:\Users\{USERNAME}\Documents\OutlookMCP\`

## 사용

Claude Desktop에서 자동으로 사용 가능:
- "지난주 김팀장이 보낸 이메일 찾아줘"
- "BIM 코디네이션 관련 PDF 첨부파일 검색해줘"
- "이메일 인덱싱 상태 확인해줘"

## MCP 도구

| 도구명 | 설명 |
|--------|------|
| `email_search` | 자연어로 이메일 검색 |
| `email_search_by_sender` | 발신자로 검색 |
| `email_search_by_date` | 날짜 범위로 검색 |
| `email_search_attachments` | 첨부파일 내용 검색 |
| `email_index_status` | 인덱싱 상태 확인 |
| `email_index_refresh` | 새 이메일 인덱싱 |
| `email_get_detail` | 특정 이메일 상세 조회 |
| `email_list_folders` | Outlook 폴더 목록 |

## 요구사항

- Windows + Outlook Desktop 앱 (로그인 상태)
- Python 3.10+
