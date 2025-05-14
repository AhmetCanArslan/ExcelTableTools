# operations/masking.py
import re

def mask_data(data, mode='default', column_name=None):
    """Masks data based on the specified mode.
    'default': Keeps the first 2 and last 2 characters (e.g., 'ab****yz').
    'email': Masks email addresses like 'us***@domain.com'.
    """
    if column_name is not None and str(data) == str(column_name):
        return data

    s_data = str(data)

    if mode == 'email':
        match = re.match(r'^([^@]+)(@.+)$', s_data)
        if match:
            user, domain = match.groups()
            if len(user) <= 2:
                return f"{user}***{domain}" # Mask short usernames
            else:
                return f"{user[:2]}***{domain}"
        else:
            return s_data # Return original if not a valid email format

    elif mode == 'default':
        if len(s_data) <= 4:
            return s_data[:1] + '*' * (len(s_data) - 2) + s_data[-1:]
        else:
            return s_data[:2] + '*' * (len(s_data) - 4) + s_data[-2:]

    return s_data # Fallback

def mask_email(data):
    """Masks email addresses like fi***@domain.com."""
    # This function seems to be superseded by mask_data(mode='email').
    # If it were to be used directly with .apply(), it would also need column_name.
    # For now, assuming it's not the primary path for column operations.
    s_data = str(data)
    if '@' in s_data:
        parts = s_data.split('@')
        local_part = parts[0]
        domain_part = parts[1]
        if len(local_part) <= 2:
            # Mask short local parts completely or differently if needed
            masked_local = '***' # Or local_part[0] + '*' * (len(local_part) -1) if len > 0
        else:
            masked_local = local_part[:2] + '***'
        return f"{masked_local}@{domain_part}"
    else:
        # Not an email, return original or apply generic mask? Returning original for now.
        return s_data

def mask_words(data, column_name=None):
    """Masks each word except the first 2 letters. E.g., 'Ahmet Can' -> 'Ah*** Ca*'."""
    if column_name is not None and str(data) == str(column_name):
        return data

    s_data = str(data)
    def mask_word(word):
        if len(word) <= 2:
            return word
        else:
            return word[:2] + '*' * max(1, len(word) - 2)
    return ' '.join(mask_word(w) for w in s_data.split())
