import pandas as pd

def apply_rename_column(dataframe, old_col, new_col_name, texts):
    """Renames a column."""
    if old_col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=old_col))
    if new_col_name in dataframe.columns:
        return dataframe, ('error', texts['column_already_exists'].format(name=new_col_name))
    new_df = dataframe.copy()
    new_df.rename(columns={old_col: new_col_name}, inplace=True)
    return new_df, ('success', texts['rename_success'].format(old=old_col, new=new_col_name))
