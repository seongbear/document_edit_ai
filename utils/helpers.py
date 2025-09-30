import datetime
from typing import Any, Optional

def format_file_size(size_bytes: int) -> str:
    """Format file size in human readable format"""
    if size_bytes == 0:
        return "0 B"
    
    size_names = ["B", "KB", "MB", "GB"]
    i = 0
    size = float(size_bytes)
    
    while size >= 1024.0 and i < len(size_names) - 1:
        size /= 1024.0
        i += 1
    
    return f"{size:.1f} {size_names[i]}"

def truncate_text(text: str, max_length: int) -> str:
    """Truncate text to specified length with ellipsis"""
    if not text:
        return ""
    
    if len(text) <= max_length:
        return text
    
    return text[:max_length - 3] + "..."

def format_timestamp(timestamp: Optional[str] = None) -> str:
    """Format timestamp for display"""
    if timestamp is None:
        timestamp = datetime.datetime.now().isoformat()
    
    try:
        if isinstance(timestamp, str):
            dt = datetime.datetime.fromisoformat(timestamp.replace('Z', '+00:00'))
        else:
            dt = timestamp
        
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except:
        return str(timestamp)

def clean_text(text: str) -> str:
    """Clean and normalize text content"""
    if not text:
        return ""
    
    # Remove excessive whitespace
    text = ' '.join(text.split())
    
    # Remove multiple consecutive newlines
    while '\n\n\n' in text:
        text = text.replace('\n\n\n', '\n\n')
    
    return text.strip()

def extract_key_phrases(text: str, max_phrases: int = 5) -> list:
    """Extract key phrases from text (simple implementation)"""
    if not text:
        return []
    
    # Simple keyword extraction based on word frequency
    words = text.lower().split()
    
    # Remove common stop words
    stop_words = {
        'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of', 'with', 'by',
        'is', 'are', 'was', 'were', 'be', 'been', 'being', 'have', 'has', 'had', 'do', 'does', 'did',
        'will', 'would', 'could', 'should', 'may', 'might', 'must', 'can', 'this', 'that', 'these', 'those'
    }
    
    filtered_words = [word for word in words if word not in stop_words and len(word) > 3]
    
    # Count frequency
    word_freq = {}
    for word in filtered_words:
        word_freq[word] = word_freq.get(word, 0) + 1
    
    # Sort by frequency and return top phrases
    sorted_words = sorted(word_freq.items(), key=lambda x: x[1], reverse=True)
    
    return [word for word, freq in sorted_words[:max_phrases]]

def validate_text_content(content: str) -> dict:
    """Validate text content and return analysis"""
    if not content:
        return {
            'valid': False,
            'errors': ['Content is empty'],
            'warnings': [],
            'stats': {'words': 0, 'characters': 0, 'paragraphs': 0}
        }
    
    words = len(content.split())
    characters = len(content)
    paragraphs = len([p for p in content.split('\n\n') if p.strip()])
    
    errors = []
    warnings = []
    
    if words < 10:
        warnings.append('Content is very short (less than 10 words)')
    
    if words > 10000:
        warnings.append('Content is very long (more than 10,000 words)')
    
    if characters > 100000:
        errors.append('Content exceeds maximum character limit (100,000)')
    
    return {
        'valid': len(errors) == 0,
        'errors': errors,
        'warnings': warnings,
        'stats': {
            'words': words,
            'characters': characters,
            'paragraphs': paragraphs
        }
    }

def parse_table_content(text: str) -> list:
    """Parse table-like content from text"""
    tables = []
    lines = text.split('\n')
    
    current_table = []
    for line in lines:
        if '|' in line and line.count('|') >= 2:
            # This looks like a table row
            cells = [cell.strip() for cell in line.split('|')]
            current_table.append(cells)
        else:
            # End of table
            if current_table:
                tables.append(current_table)
                current_table = []
    
    # Don't forget the last table
    if current_table:
        tables.append(current_table)
    
    return tables

def format_conversation_history(history: list) -> str:
    """Format conversation history for display or processing"""
    if not history:
        return "No conversation history"
    
    formatted = []
    for i, message in enumerate(history):
        role = message.get('role', 'unknown')
        content = message.get('content', '')
        timestamp = message.get('timestamp', '')
        
        formatted.append(f"{i+1}. [{role.upper()}] {content}")
        if timestamp:
            formatted.append(f"   Time: {format_timestamp(timestamp)}")
        formatted.append("")  # Empty line for separation
    
    return '\n'.join(formatted)
