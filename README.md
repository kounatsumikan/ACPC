# Automatically_create_PPTX_from_CSV
csvデータが保管されているディレクトリを指定することで

指定されたディレクトリ内のcsvの詳細データ、基礎集計結果をpptxとして出力する。

資料作成の時間を大幅に削減できる。

# version情報
- α版:0.2.0


# 実施予定

- カラム数が多い場合の対処法
    - ページを分けるなど

- リファクタリング

- pathの整理

- コード内のコメントに追加

- READMEの充実

# instration

1. pandas_2_pptxをインストールしてください
    - https://github.com/kounatsumikan/pandas_2_pptx

2. acpcのsetup.pyを実行してインストールしてください(管理者権限で実行)
    - ```python setup.py install```

## フォルダ構成

<!-- ```
.
├── README.md                      README（本ファイル）
├── setup.py                       パッケージセットアップ
├── notebook                       開発用のスクリプト
└── src                            pandas_2_pptxのパッケージに格納されるスクリプト群
    ├── __init__.py
    ├── pandas_2_pptx.py           pptxファイルを生成するスクリプト
    └── slide.py                   スライドを追加するスクリプト

``` -->

## Documentation

<!-- - pandas_2_pptx
    - add_title_slide    function
    - add_slide          function
    - save               function
    - slide              object
- slide
    - add_chart          function
    - add_table          function
    - generate           function -->
