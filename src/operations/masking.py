# operations/masking.py
import re
import pandas as pd

def mask_data(data, mode='default', column_name=None, track_invalid=False):
    """Masks data based on the specified mode.
    'default': Keeps the first 2 and last 2 characters (e.g., 'ab****yz').
    'email': Masks email addresses like 'us***@domain.com'.
    
    If track_invalid=True, returns a tuple (masked_value, is_valid) for email mode.
    """
    if column_name is not None and str(data) == str(column_name):
        return (data, True) if track_invalid and mode == 'email' else data

    s_data = str(data)

    if mode == 'email':
        # Basic email validation
        email_pattern = r'^([^@]+)(@.+)$'
        match = re.match(email_pattern, s_data)
        
        # If input isn't a valid email format
        if not match or '@' not in s_data or len(s_data.split('@')) != 2:
            if track_invalid:
                return (s_data, False)  # Mark as invalid
            return s_data  # Return original if not a valid email format
            
        user, domain = match.groups()
        masked_value = None
        if len(user) <= 2:
            masked_value = f"{user}***{domain}"  # Mask short usernames
        else:
            masked_value = f"{user[:2]}***{domain}"
            
        if track_invalid:
            return (masked_value, True)  # Mark as valid
        return masked_value

    elif mode == 'default':
        if len(s_data) <= 4:
            masked = s_data[:1] + '*' * (len(s_data) - 2) + s_data[-1:]
        else:
            masked = s_data[:2] + '*' * (len(s_data) - 4) + s_data[-2:]
        # Convert to safe string
        try:
            f = float(masked)
            if f.is_integer():
                return str(int(f))
        except Exception:
            pass
        return masked

    return s_data  # Fallback

def mask_email(data, column_name=None, track_invalid=False):
    """Masks email addresses like fi***@domain.com. Updated to track invalid emails."""
    return mask_data(data, mode='email', column_name=column_name, track_invalid=track_invalid)

def mask_words(data, column_name=None):
    """Masks each word except the first 2 letters. E.g., 'Ahmet Can' -> 'Ah*** Ca'."""
    if column_name is not None and str(data) == str(column_name):
        return data

    s_data = str(data)
    def mask_word(word):
        if len(word) <= 2:
            return word
        else:
            return word[:2] + '*' * max(1, len(word) - 2)
    return ' '.join(mask_word(w) for w in s_data.split())
