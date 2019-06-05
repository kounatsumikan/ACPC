# pandas_2_pptx
pandasのDataFrameをパワーポイントのグラフやテーブルに変換して出力する

エクセルやグラフの画像ファイルを経由せず直接pptx化し効率化を求める

# version情報
- α版:0.5.0

- 表中の文字の大きさを自動調整する機能を追加。
    - 動作がまだ不安定

## フォルダ構成

```
.
├── README.md                      README（本ファイル）
├── setup.py                       パッケージセットアップ
├── notebook                       開発用のスクリプト
└── src                            pandas_2_pptxのパッケージに格納されるスクリプト群
    ├── __init__.py
    ├── pandas_2_pptx.py           pptxファイルを生成するスクリプト
    └── slide.py                   スライドを追加するスクリプト

```

## Documentation

- pandas_2_pptx
    - add_title_slide    function
    - add_slide          function
    - save               function
    - slide              object
- slide
    - add_chart          function
    - add_table          function
    - generate           function
