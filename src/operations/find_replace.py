# operations/find_replace.py

def find_replace(data, find_text, replace_text, column_name=None):
    """Replaces occurrences of find_text with replace_text."""
    if column_name is not None and str(data) == str(column_name):
        return data
        
    s_data = str(data)
    # Basic replacement, consider adding options like case sensitivity later
    return s_data.replace(find_text, replace_text)
