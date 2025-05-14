# operations/case_change.py

def change_case(data, case_type, column_name=None):
    """Changes the case of the string data."""
    if column_name is not None and str(data) == str(column_name):
        return data
        
    s_data = str(data)
    if case_type == 'upper':
        return s_data.upper()
    elif case_type == 'lower':
        return s_data.lower()
    elif case_type == 'title':
        return s_data.title()
    return s_data # Default return original if type unknown
