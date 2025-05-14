# operations/trimming.py

def trim_spaces(data, column_name=None):
    """Removes leading/trailing spaces from data."""
    if column_name is not None and str(data) == str(column_name):
        return data
    return str(data).strip()
