

def describe(df, percentiles=[0.01, 0.1, 0.9, 0.99],
             include='all', exclude=None, std_outlier=3):
    """
    各種統計量を算出.

    Args:
        df (pandas.dfFrame): 統計量を算出するデータ.
        percentiles (list-like of numbers): 分位点.
        include ('all', list-like of dtypes or None): 統計量を算出する変数.
        exclude (list-like of dtypes or None (default)): 統計量を算出しない変数.
        std_outlier (int): 外れ値を判定する際の係数.
            「平均 +- 標準偏差 × std_outlier」の範囲外のレコードを外れ値と判定する.

    Returns:
        desc (pandas.dfFrame): 各種統計量・外れ値の件数・欠損率・データ型.

    """
    desc = df.select_dtypes(exclude=["datetime"]).describe(percentiles=percentiles, include=include, exclude=exclude).T
    null_rate = (df.isnull().sum() / len(df)).to_frame('欠損率')
    dtypes = df.dtypes.to_frame('dtypes')
    df_int = df.select_dtypes(include=['int64', 'object'], exclude=[])
    unique = df_int.nunique().to_frame('unique')
    df_num = df.select_dtypes(['float', 'int'])
    desc_num = desc.loc[df_num.columns]
    lower = df_num < desc_num['mean'] - desc_num['std'] * std_outlier
    upper = df_num > desc_num['mean'] + desc_num['std'] * std_outlier
    lower = lower.sum().to_frame(f'mean-{std_outlier}std')
    upper = upper.sum().to_frame(f'mean+{std_outlier}std')
    #skew = df_num.skew().to_frame(f'歪度')
    #kurt = df_num.kurt().to_frame(f'尖度')

    args_merge = dict(left_index=True, right_index=True, how='left')
    desc = (
        desc
        .merge(null_rate, **args_merge)
        .merge(dtypes, **args_merge)
        # .merge(skew, **args_merge)
        # .merge(kurt, **args_merge)
    )
    desc["unique"] = unique
    return desc[['unique', 'count', '欠損率', 'mean', 'std', 'min', '1%', '10%', '50%', '90%', '99%', 'max', 'dtypes']]
