import chardet
import os
import pandas as pd
import collections
import random
from pandas_2_pptx.pandas_2_pptx import pandas_2_pptx
from utils.utils import make_file_list
from utils.utils import check_datetime_column
from describe.describe import describe

class Acpc():

    def __init__(self, dir_path, pptx_title, save_dir="."):
        desc_list = []
        detail_list = []
        self.f_list = make_file_list(dir_path)
        for n, fname in enumerate(self.f_list):
            try:
                df = pd.read_csv(fname)
            except UnicodeDecodeError:
                df = pd.read_csv(fname,encoding='cp932')
            df, datetime_list = check_datetime_column(df)
            detail_list.append(self.__data_details(fname=fname, parse_dates=datetime_list))
            desc = describe(df)    
            desc[["mean",	"std",	"min",	"1%",	"10%",	"50%",	"90%",	"99%",	"max"]] = desc[["mean",	"std",	"min",	"1%",	"10%",	"50%",	"90%",	"99%",	"max"]].astype(float).round(2)
            desc[["count"]] = desc[["count"]].astype(int)
            desc = desc.round(2)
            desc_list.append(desc)
        pptx = pandas_2_pptx()
        title_slide=pptx.add_title_slide(pptx_title)
        details = pd.concat(detail_list, axis=1, sort=False)
        details.columns = self.f_list
        slide1 = pptx.add_slide("データ概要")
        slide1.add_table(details)
        slide1.generate()
        for n, i in enumerate(desc_list):
            slide2 = pptx.add_slide(f"基礎集計{self.f_list[n]}") #　変数名を使いまわしている
            slide2.add_table(i)                                 # 改善すべし
            slide2.generate()
        self.save_dir_path = f"{save_dir}{os.path.sep}"
        if not os.path.exists(self.save_dir_path):
            os.mkdir(self.save_dir_path)
        pptx.save(f"{self.save_dir_path}{pptx_title}.pptx")
        self.save_dir_path = f"{self.save_dir_path}{pptx_title}.pptx"
        
    def __repr__(self):
        return f"save {self.save_dir_path}"

    def __delimiter_check(self, fname):
        """区切り文字をチェックする関数
        """
        data = open(fname, "r").read()
        try:
            df = pd.read_csv(fname)
        except UnicodeDecodeError:
            df = pd.read_csv(fname, encoding='cp932')
        splitter = 0
        split_list = []
        for n, rand in enumerate([random.randint(0, len(df)) for num in range(5)]):
            li = []
            for i in data.split("\n")[rand]:
                li.append(i)
            split_list.append([key for key, value in collections.Counter(li).items() if value == df.shape[1] - 1])
            try:
                delimiter = set(split_list[n-1]) & set(split_list[n]) & delimiter
            except:
                delimiter = set(split_list[n])
        if(len(delimiter) != 1):
            delimiter = None
        else:
            delimiter = list(delimiter)[0]
        return delimiter

    def __encoding_check(self, fname):
        """encodingをチェックする関数
        """
        data = open(fname, "rb")
        enc = chardet.detect(data.read())["encoding"]
        return enc

    def __data_details(self, fname, parse_dates = None):
        """csvデータを読み込んで、詳細情報を集計している関数
        """
        try:
            df = pd.read_csv(fname, parse_dates = parse_dates)
        except UnicodeDecodeError:
            df = pd.read_csv(fname,encoding='cp932', parse_dates = parse_dates)
        delimiter = self.__delimiter_check(fname)
        encoding = self.__encoding_check(fname)
        details_list = [
            delimiter,
            encoding,
            df.shape[0],
            df.shape[1],
        ]
        index_list = ["区切り文字","文字コード","行数","カラム数"]
        datetime_col = df.columns[(df.dtypes == '<M8[ns]').values]
        period_list = []
        for n, i in enumerate(datetime_col):
            index_list.append(f"データ期間[{i}]")
            period_list.append(f"{df[i].min()} ~ {df[i].max()}")
            details_list.append(period_list[n])
        return pd.Series(details_list, index=index_list)
