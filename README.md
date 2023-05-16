# photo-album-generator
このプログラムは、指定されたディレクトリの写真をランダムに選択し、それらをグリッド状に配置してパワーポイントのプレゼンテーションを生成します。各スライドは指定された行数と列数で構成され、各写真は指定されたアスペクト比に合わせてトリミングされます。

# 必要な環境
- Python 3.7以上（ただし、3.10は除く）
- ライブラリ: requirements.txtに記載されています。以下のコマンドでインストールできます。

```
pip install -r requirements.txt
```

- HEICフォーマットの画像を処理する場合、追加でlibheifとlibexiv2が必要となります。これらのライブラリはシステムレベルでインストールする必要があります。

  - Ubuntuの場合：
    ```
    sudo apt-get install libheif-dev libexiv2-dev
    ```
  - macOSの場合（Homebrewを使用）：
    ```
    brew install libheif
    brew install exiv2
    ```
  また、HEICフォーマットの画像を処理する場合は、Pythonのパッケージpyheifとpyexiv2も必要になります。これらは以下のコマンドでインストールできます。
  
  ```
  pip install pyheif pyexiv2
  ```

  これらのライブラリとパッケージはHEICフォーマットの画像からExifデータを読み込むために必要です。HEIC以外の画像フォーマットを扱う場合は、これらのインストールは不要です。

# 使い方

```sh
python photo_album_generator.py --input_dir [入力ディレクトリ] --rows [行数] --columns [列数] --num_pages [ページ数] --crop_aspect_ratio [トリミングするアスペクト比] --output_file [出力ファイル名]
```

- input_dir：写真が格納されているディレクトリのパスを指定します。
- rows：各ページに表示する写真の行数を指定します。デフォルトは2です。
- columns：各ページに表示する写真の列数を指定します。デフォルトは3です。
- num_pages：生成するページ数を指定します。デフォルトは5です。
- crop_aspect_ratio：写真をトリミングする際のアスペクト比を指定します。例えば、「16:9」のように指定します。デフォルトは1.0（正方形）です。
- output_file：生成されたパワーポイントのファイル名を指定します。デフォルトは「photo_album.pptx」です。