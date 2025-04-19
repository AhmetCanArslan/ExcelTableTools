# operations/case_change.py

def change_case(data, case_type):
    """Changes the case of the string data."""
    s_data = str(data)
    if case_type == 'upper':
        return s_data.upper()
    elif case_type == 'lower':
        return s_data.lower()
    elif case_type == 'title':
        return s_data.title()
    return s_data # Default return original if type unknown
