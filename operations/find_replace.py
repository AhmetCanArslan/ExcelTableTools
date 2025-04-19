# operations/find_replace.py

def find_replace(data, find_text, replace_text):
    """Replaces occurrences of find_text with replace_text."""
    s_data = str(data)
    # Basic replacement, consider adding options like case sensitivity later
    return s_data.replace(find_text, replace_text)
