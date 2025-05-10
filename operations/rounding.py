import pandas as pd

def apply_round_numbers(df, col, decimals, texts):
    """
    Rounds the numbers in the specified column of the DataFrame to the given number of decimal places.

    Args:
        df (pd.DataFrame): The DataFrame to modify.
        col (str): The name of the column to round.
        decimals (int): The number of decimal places to round to.
        texts (dict): Dictionary containing UI texts for translations.

    Returns:
        tuple: A tuple containing the modified DataFrame and a status message.
    """
    try:
        df[col] = pd.to_numeric(df[col], errors='coerce')
        df[col] = df[col].round(decimals)
        return df, ('success', texts['round_success'].format(col=col, decimals=decimals))
    except Exception as e:
        return df, ('error', texts['operation_error'].format(error=str(e)))
