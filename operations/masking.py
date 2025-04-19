# operations/masking.py

def mask_data(data):
    """Masks data by keeping the first 2 and last 2 characters."""
    s_data = str(data)
    if len(s_data) <= 4:
        return s_data # Or return "****" if you want to mask short strings too
    else:
        return s_data[:2] + '*' * (len(s_data) - 4) + s_data[-2:]
