import re
import pandas as pd
from datetime import datetime
from dateutil.parser import parse #for validating dates


COMMON_EMAIL_DOMAINS = {
    "gmail.com", "yahoo.com", "hotmail.com", "outlook.com",
    "icloud.com", "protonmail.com", "yandex.com", "mail.com",
    "gmx.com", "zoho.com", "atauni.edu.tr", "ogr.atauni.edu.tr", 'edu.tr',
}

def validate_email(value, column_name=None):
    """Validates if the value is a likely real email address."""
    if column_name is not None and str(value) == str(column_name):
        return False, "Column Header"

    if pd.isna(value) or value.strip() == "":
        return False, "Empty"
    
    value = str(value).strip()

    # Regex with stricter RFC-like rules
    email_pattern = r"^(?!.*\.\.)(?!.*\.$)[^\W][\w.%+-]{0,63}@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    if not re.match(email_pattern, value):
        return False, "Invalid Format"
    
    # Extract domain and check for typos or suspicious domains
    domain = value.split('@')[-1].lower()

    if domain not in COMMON_EMAIL_DOMAINS:
        return True, "Suspicious Domain"
    
    return True, "Valid"



def validate_phone(value, column_name=None):
    """Validates if the value is a likely valid phone number."""
    if column_name is not None and str(value) == str(column_name):
        return False, "Column Header"

    if pd.isna(value) or value.strip() == "":
        return False, "Empty"

    value = str(value).strip()

    # Remove allowed formatting characters to count digits
    digits_only = re.sub(r"[^\d]", "", value)
    if len(digits_only) < 7 or len(digits_only) > 15:
        return False, "Invalid Length"

    # Strict pattern: optional + at start, digits, spaces, dashes, parentheses
    pattern = r"^\+?[0-9\s\-\(\)]{7,20}$"
    if not re.match(pattern, value):
        return False, "Invalid Format"

    return True, "Valid"



def validate_date(value, column_name=None):
    if column_name is not None and str(value) == str(column_name):
        return False, "Column Header"

    if pd.isna(value) or str(value).strip() == "":
        return False, "Empty"

    try:
        parse(str(value), fuzzy=False)
        return True, "Valid"
    except (ValueError, TypeError):
        return False, "Invalid Format"




def validate_numeric(value, column_name=None):
    """Validates if the value is numeric."""
    if column_name is not None and str(value) == str(column_name):
        return False, "Column Header"
    
    if pd.isna(value) or value == "":
        return False, "Empty"
    
    if isinstance(value, (int, float)):
        return True, "Valid"
    
    value = str(value)
    # Allow numbers with decimal points and negative signs
    numeric_pattern = r'^-?\d*\.?\d+$'
    is_valid = bool(re.match(numeric_pattern, value))
    return is_valid, "Valid" if is_valid else "Invalid Format"

def validate_alphanumeric(value, column_name=None):
    """Validates if the value contains only alphanumeric characters and spaces."""
    if column_name is not None and str(value) == str(column_name):
        return False, "Column Header"
    
    if pd.isna(value) or value == "":
        return False, "Empty"
    
    value = str(value)
    # Allow alphanumeric and spaces
    alphanumeric_pattern = r'^[a-zA-Z0-9\s]+$'
    is_valid = bool(re.match(alphanumeric_pattern, value))
    return is_valid, "Valid" if is_valid else "Invalid Format"

def validate_url(value, column_name=None):
    """Validates if the value is a valid URL."""
    if column_name is not None and str(value) == str(column_name):
        return False, "Column Header"
    
    if pd.isna(value) or value == "":
        return False, "Empty"
    
    value = str(value)
    # Simple URL validation
    url_pattern = r'^(https?:\/\/)?(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.~#?&//=]*)$'
    is_valid = bool(re.match(url_pattern, value))
    return is_valid, "Valid" if is_valid else "Invalid Format"

def apply_validation(dataframe, col, validation_type, texts):
    """Applies validation to a column based on the selected type and colors invalid cells red."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))
    
    # Make a copy to avoid modifying the original
    new_df = dataframe.copy()
    
    # Select validation function based on type
    validation_functions = {
        'email': validate_email,
        'phone': validate_phone,
        'date': validate_date,
        'numeric': validate_numeric,
        'alphanumeric': validate_alphanumeric,
        'url': validate_url
    }
    
    if validation_type not in validation_functions:
        return dataframe, ('error', f"Unknown validation type: {validation_type}")
    
    validation_function = validation_functions[validation_type]
    
    # Apply validation and get results
    validation_results = new_df[col].apply(lambda x: validation_function(x, col))
    
    # Extract boolean validation results (is_valid)
    is_valid_series = validation_results.apply(lambda x: x[0])
    
    # Calculate validation statistics
    valid_count = is_valid_series.sum()
    total_count = len(new_df)
    validation_rate = (valid_count / total_count) * 100 if total_count > 0 else 0
    
    # Create the _styled_columns attribute in a safer way
    if not hasattr(new_df, '_styled_columns'):
        # Use object.__setattr__ to avoid pandas warning
        object.__setattr__(new_df, '_styled_columns', {})
    
    # Save which cells should be highlighted
    new_df._styled_columns[col] = ~is_valid_series
    
    # Return success message with statistics
    success_message = texts['check_valid_inputs_success'].format(
        col=col, 
        type=texts.get(f'validation_{validation_type}', validation_type)
    )
    
    # Add validation stats to message
    success_message += f" ({valid_count}/{total_count} valid, {validation_rate:.1f}%)"
    success_message += " " + texts['validation_color_applied'].format(col=col)
    
    # Return the dataframe and status message
    return new_df, ('success', success_message)
