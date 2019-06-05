import os
from glob import glob
import pandas as pd

def make_file_list(path, f_type="csv"):
    """
    Returns:
        f_list (list): ディレクトリ内のファイルリスト
    """
    f_list = glob(f'{path}{os.path.sep}**{os.path.sep}*.{f_type}', recursive=True)
    return f_list

def check_datetime_column(df):
    """
    Returns:
        df (pandas.DataFrame): object型の特定のカラムをdatetime型のカラムに変換したdataframe
        datetime_columns (list): datetime型に変換したカラム名のリスト
    """
    for column in df.select_dtypes(include='object').columns:
        try:
            df[column] = pd.to_datetime(df[column])
            print(column, "をdatetime型に変換しました")
        except:
            continue
    datetime_columns = df.select_dtypes(include='datetime64[ns]').columns.tolist()
    return df, datetime_columns