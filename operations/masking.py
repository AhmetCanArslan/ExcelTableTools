# operations/masking.py
import re

def mask_data(data, mode='default'):
    """Masks data based on the specified mode.
    'default': Keeps the first 2 and last 2 characters (e.g., 'ab****yz').
    'email': Masks email addresses like 'us***@domain.com'.
    """
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
            return s_data # Or return "****" if you want to mask short strings too
        else:
            return s_data[:2] + '*' * (len(s_data) - 4) + s_data[-2:]

    return s_data # Fallback

def mask_email(data):
    """Masks email addresses like fi***@domain.com."""
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
