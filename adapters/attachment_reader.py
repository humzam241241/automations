"""
Attachment Reader - Extracts text from various file formats.
Supports: PDF, Word (.docx), Text files, and more.
"""
import os
import re
from pathlib import Path
from typing import Optional


def extract_text_from_file(file_path: str) -> str:
    """
    Extract text from a file based on its extension.
    
    Supports:
    - PDF files (.pdf)
    - Word documents (.docx)
    - Text files (.txt)
    - CSV files (.csv)
    
    Args:
        file_path: Path to the file
    
    Returns:
        Extracted text content
    """
    if not file_path or not os.path.exists(file_path):
        return ""
    
    ext = Path(file_path).suffix.lower()
    
    try:
        if ext == '.pdf':
            return extract_pdf_text(file_path)
        elif ext == '.docx':
            return extract_docx_text(file_path)
        elif ext == '.doc':
            return extract_doc_text(file_path)
        elif ext in ['.txt', '.csv', '.log', '.md']:
            return extract_plain_text(file_path)
        else:
            return ""
    except Exception as e:
        print(f"Warning: Could not extract text from {file_path}: {e}")
        return ""


def extract_pdf_text(file_path: str) -> str:
    """Extract text from a PDF file."""
    try:
        from PyPDF2 import PdfReader
        
        reader = PdfReader(file_path)
        text_parts = []
        
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text_parts.append(page_text)
        
        return "\n\n".join(text_parts)
    except ImportError:
        print("PyPDF2 not installed. Install with: pip install PyPDF2")
        return ""
    except Exception as e:
        print(f"PDF extraction error: {e}")
        return ""


def extract_docx_text(file_path: str) -> str:
    """Extract text from a Word document (.docx)."""
    try:
        from docx import Document
        
        doc = Document(file_path)
        text_parts = []
        
        for para in doc.paragraphs:
            if para.text.strip():
                text_parts.append(para.text)
        
        # Also extract from tables
        for table in doc.tables:
            for row in table.rows:
                row_text = " | ".join(cell.text for cell in row.cells)
                if row_text.strip():
                    text_parts.append(row_text)
        
        return "\n".join(text_parts)
    except ImportError:
        print("python-docx not installed. Install with: pip install python-docx")
        return ""
    except Exception as e:
        print(f"DOCX extraction error: {e}")
        return ""


def extract_doc_text(file_path: str) -> str:
    """Extract text from old Word format (.doc) - limited support."""
    # .doc files require additional tools (antiword, etc.)
    # For now, return empty and log warning
    print(f"Warning: .doc format not fully supported: {file_path}")
    return ""


def extract_plain_text(file_path: str) -> str:
    """Extract text from a plain text file."""
    try:
        encodings = ['utf-8', 'latin-1', 'cp1252', 'ascii']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    return f.read()
            except UnicodeDecodeError:
                continue
        
        # Fallback: binary read and decode what we can
        with open(file_path, 'rb') as f:
            return f.read().decode('utf-8', errors='ignore')
    except Exception as e:
        print(f"Text file read error: {e}")
        return ""


def extract_attachment_from_bytes(data: bytes, filename: str) -> str:
    """
    Extract text from attachment data in memory.
    
    Args:
        data: Raw bytes of the attachment
        filename: Original filename (for extension detection)
    
    Returns:
        Extracted text
    """
    import tempfile
    
    ext = Path(filename).suffix.lower()
    
    if ext not in ['.pdf', '.docx', '.txt', '.csv']:
        return ""
    
    try:
        # Write to temp file and extract
        with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        
        text = extract_text_from_file(tmp_path)
        
        # Clean up
        try:
            os.unlink(tmp_path)
        except:
            pass
        
        return text
    except Exception as e:
        print(f"Attachment extraction error: {e}")
        return ""


def extract_numbers_from_text(text: str) -> list:
    """
    Extract all reference numbers and IDs from text.
    
    Returns list of potential reference numbers found.
    """
    patterns = [
        r'[A-Z]{2,5}[-#]?\d{3,10}',  # MI-12345, CAPA123
        r'\d{5,10}',  # Long numbers
        r'[A-Z]{2,3}\d{2}[-/]\d{4,6}',  # PO22-123456
    ]
    
    numbers = []
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        numbers.extend(matches)
    
    return list(set(numbers))
