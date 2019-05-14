# pandas_2_pptx
pandasのDataFrameをパワーポイントのグラフやテーブルに変換して出力する
エクセルやグラフの画像ファイルを経由せず直接pptx化し効率化を求める

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