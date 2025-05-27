import re
import pandas as pd
from datetime import datetime
from dateutil.parser import parse #for validating dates
from urllib.parse import urlparse
from .domain_validation import DomainValidator

# Initialize domain validator as a module-level singleton
domain_validator = DomainValidator()


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
    
    # Extract domain and validate using DomainValidator
    domain = value.split('@')[-1].lower()
    is_valid, reason = domain_validator.is_valid_domain(domain)
    
    return is_valid, reason



def validate_phone(value, column_name=None):
    """Validates if the value is a likely valid phone number."""
    if column_name is not None and str(value) == str(column_name):
        return False, "Column Header"

    if pd.isna(value) or str(value).strip() == "":
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
    """Validates if the value is a number (int, float, or scientific notation)."""
    if column_name is not None and str(value) == str(column_name):
        return False, "Column Header"

    if pd.isna(value) or str(value).strip() == "":
        return False, "Empty"

    if isinstance(value, (int, float)):
        return True, "Valid"

    try:
        float(str(value))
        return True, "Valid"
    except ValueError:
        return False, "Invalid Format"


def validate_alphanumeric(value, column_name=None):
    """Validates if the value contains only letters (Unicode) and spaces."""
    if column_name is not None and str(value) == str(column_name):
        return False, "Column Header"
    
    if pd.isna(value) or str(value).strip() == "":
        return False, "Empty"

    value = str(value).strip()
    for ch in value:
        if not (ch.isalpha() or ch.isspace()):
            return False, "Invalid Character"
    
    return True, "Valid"





def validate_url(value, column_name=None):
    """Validates if the value is a valid URL using urlparse."""
    if column_name is not None and str(value) == str(column_name):
        return False, "Column Header"

    if pd.isna(value) or str(value).strip() == "":
        return False, "Empty"

    value = str(value).strip()
    parsed = urlparse(value)

    if parsed.scheme in ('http', 'https') and parsed.netloc:
        return True, "Valid"
    else:
        return False, "Invalid Format"


def apply_validation(dataframe, col, validation_type, texts):
    """Applies validation to a column based on the selected type and colors invalid cells red."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))
    
    new_df = dataframe.copy()
    
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
    
    validation_results = new_df[col].apply(lambda x: validation_function(x, col))
    is_valid_series = validation_results.apply(lambda x: x[0])
    
    valid_count = is_valid_series.sum()
    total_count = len(new_df)
    validation_rate = (valid_count / total_count) * 100 if total_count > 0 else 0
    
    if not hasattr(new_df, '_styled_columns'):
        object.__setattr__(new_df, '_styled_columns', {})
    new_df._styled_columns[col] = ~is_valid_series

    # Ensure column stays as string (prevents .0 for numbers)
    new_df[col] = new_df[col].apply(lambda v: str(v) if not pd.isna(v) else "")

    success_message = texts['check_valid_inputs_success'].format(
        col=col, 
        type=texts.get(f'validation_{validation_type}', validation_type)
    )
    success_message += f" ({valid_count}/{total_count} valid, {validation_rate:.1f}%)"
    success_message += " " + texts['validation_color_applied'].format(col=col)
    
    return new_df, ('success', success_message)
