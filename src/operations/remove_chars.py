# operations/remove_chars.py
import re

def remove_chars(data, mode='specific', chars_to_remove='', column_name=None):
    """Removes characters based on the mode."""
    if column_name is not None and str(data) == str(column_name):
        return data
        
    s_data = str(data)
    if mode == 'specific':
        for char in chars_to_remove:
            s_data = s_data.replace(char, '')
        return s_data
    elif mode == 'non_numeric':
        # Keeps digits, decimal points, and minus signs (basic)
        return re.sub(r'[^0-9.-]', '', s_data)
    elif mode == 'non_alphabetic':
        # Use isalpha() for better Unicode support across languages
        # This keeps only letters and spaces
        return ''.join(c for c in s_data if c.isalpha() or c.isspace())
    return s_data
