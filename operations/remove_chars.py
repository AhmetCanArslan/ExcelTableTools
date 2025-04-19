# operations/remove_chars.py
import re

def remove_chars(data, mode='specific', chars_to_remove=''):
    """Removes characters based on the mode."""
    s_data = str(data)
    if mode == 'specific':
        for char in chars_to_remove:
            s_data = s_data.replace(char, '')
        return s_data
    elif mode == 'non_numeric':
        # Keeps digits, decimal points, and minus signs (basic)
        return re.sub(r'[^0-9.-]', '', s_data)
    elif mode == 'non_alphabetic':
        # Keeps only letters (unicode aware)
        return re.sub(r'[^\w\s]', '', s_data, flags=re.UNICODE) # Keeps letters and spaces
        # Or stricter: return re.sub(r'[^a-zA-Z]', '', s_data) # Only basic ASCII letters
    return s_data
