# operations/find_replace.py

def find_replace(data, find_text, replace_text, column_name=None):
    """Replaces occurrences of find_text with replace_text. Returns tuple (new_value, was_changed)."""
    if column_name is not None and str(data) == str(column_name):
        return data, False
        
    s_data = str(data)
    new_data = s_data.replace(find_text, replace_text)
    was_changed = (s_data != new_data)
    return new_data, was_changed
