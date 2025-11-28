"""文件解析工具"""
import os
from pathlib import Path
from typing import Optional, Dict

# 文件大小和文本长度限制
MAX_FILE_SIZE = 5 * 1024 * 1024  # 5 MB
MAX_TEXT_LENGTH = 50000  # 5 万字（约 10 万 tokens）


def parse_file(file_path: str) -> str:
    """解析文件内容
    
    Args:
        file_path: 文件路径
        
    Returns:
        解析后的文本内容
        
    Raises:
        ValueError: 文件过大或格式不支持
    """
    # 检查文件大小
    file_size = os.path.getsize(file_path)
    if file_size > MAX_FILE_SIZE:
        raise ValueError(f"文件过大，最大支持 {MAX_FILE_SIZE / 1024 / 1024:.1f} MB")
    
    ext = Path(file_path).suffix.lower()
    
    if ext in ['.txt', '.md']:
        return parse_text(file_path)
    elif ext == '.docx':
        return parse_docx(file_path)
    elif ext == '.pdf':
        return parse_pdf(file_path)
    else:
        raise ValueError(f"不支持的文件格式: {ext}。支持的格式：TXT, MD, DOCX, PDF")


def parse_text(file_path: str) -> str:
    """解析文本文件（TXT, MD）"""
    try:
        # 尝试多种编码
        encodings = ['utf-8', 'gbk', 'gb2312', 'utf-16']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    content = f.read()
                return truncate_text(content)
            except UnicodeDecodeError:
                continue
        
        raise ValueError("无法识别文件编码")
    except Exception as e:
        raise ValueError(f"解析文本文件失败: {e}")


def parse_docx(file_path: str) -> str:
    """解析 Word 文档"""
    try:
        import docx
    except ImportError:
        raise ValueError("需要安装 python-docx 库：pip install python-docx")
    
    try:
        doc = docx.Document(file_path)
        paragraphs = []
        
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                paragraphs.append(text)
        
        content = '\n'.join(paragraphs)
        return truncate_text(content)
    except Exception as e:
        raise ValueError(f"解析 Word 文档失败: {e}")


def parse_pdf(file_path: str) -> str:
    """解析 PDF 文档"""
    try:
        import PyPDF2
    except ImportError:
        raise ValueError("需要安装 PyPDF2 库：pip install PyPDF2")
    
    try:
        content = []
        with open(file_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            
            for page in reader.pages:
                text = page.extract_text()
                if text.strip():
                    content.append(text)
        
        full_text = '\n'.join(content)
        return truncate_text(full_text)
    except Exception as e:
        raise ValueError(f"解析 PDF 文档失败: {e}")


def truncate_text(text: str, max_length: int = MAX_TEXT_LENGTH) -> str:
    """截断过长的文本
    
    Args:
        text: 原始文本
        max_length: 最大长度
        
    Returns:
        截断后的文本
    """
    if len(text) <= max_length:
        return text
    
    # 尝试在段落边界截断
    truncated = text[:max_length]
    last_newline = truncated.rfind('\n\n')
    
    if last_newline > max_length * 0.8:  # 如果在后 20% 找到段落边界
        truncated = truncated[:last_newline]
    
    return truncated + f"\n\n[文本过长，已截断到 {len(truncated)} 字，原文 {len(text)} 字]"


def get_text_summary(text: str) -> Dict:
    """获取文本摘要信息
    
    Args:
        text: 文本内容
        
    Returns:
        包含统计信息的字典
    """
    return {
        'length': len(text),
        'lines': text.count('\n') + 1,
        'words': len(text.split()),
        'estimated_tokens': len(text) * 2,  # 粗略估计：1 中文字 ≈ 2 tokens
        'is_truncated': '[文本过长，已截断' in text
    }


def validate_file(filename: str) -> bool:
    """验证文件名和扩展名
    
    Args:
        filename: 文件名
        
    Returns:
        是否有效
    """
    if not filename:
        return False
    
    ext = Path(filename).suffix.lower()
    allowed_extensions = ['.txt', '.md', '.docx', '.pdf']
    
    return ext in allowed_extensions
