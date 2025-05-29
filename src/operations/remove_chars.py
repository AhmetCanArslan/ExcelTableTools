# operations/remove_chars.py
import re

def remove_chars(data, mode='specific', chars_to_remove='', column_name=None):
    """Removes characters based on the mode. Returns tuple (new_value, was_changed) for tracking."""
    if column_name is not None and str(data) == str(column_name):
        return data, False
        
    s_data = str(data)
    original_data = s_data
    
    if mode == 'specific':
        for char in chars_to_remove:
            s_data = s_data.replace(char, '')
        result = s_data
        was_changed = (original_data != result)
        return result, was_changed
    elif mode == 'non_numeric':
        # Keeps digits, decimal points, and minus signs (basic)
        s_data = re.sub(r'[^0-9.-]', '', s_data)
        result = s_data
        was_changed = (original_data != result)
        return result, was_changed
    elif mode == 'non_alphabetic':
        # Use isalpha() for better Unicode support across languages
        # This keeps only letters and spaces
        s_data = ''.join(c for c in s_data if c.isalpha() or c.isspace())
        result = s_data
        was_changed = (original_data != result)
        return result, was_changed
    
    was_changed = (original_data != s_data)
    return s_data, was_changed
