# pptx-to-html
This script converts a ".pptx" file into a clickable map.

## 概要
PowerPointから簡単なクリッカブルマップ(イメージマップ)のhtmlを生成するスクリプトです。

## 使い方
### 動作環境
Python 3.9.6 にて動作確認済みです。

### 必要なライブラリ
PowerPointライブラリ  
`pip install python-pptx`  
画像処理ライブラリ  
`pip install pillow`  

### PowerPointの準備
ハイパーリンク付のPowerPointファイルを作成しプロジェクト直下に格納します。  
※スライド間のリンクのみ対応可能です。  
※グループ化したオブジェクトのリンクは非対応です。  

### クリッカブルマップ用画像の準備
作成したPowerPointを以下の手順にて.png形式でエクスポートします。  
「ファイル」>「名前を付けて保存」>「PNG」ですべてのスライドを保存  
pngをすべてimageフォルダ直下に格納します。  

### PptxToHtmlを実行
PowerPointと画像の準備が完了したら、以下コマンドで実行します。  
`py PptxToHtml.py`    
プロジェクト直下に.htmlファイルが生成されます。

### 実行オプション


## テンプレートHTML/CSSのカスタマイズ
デザインを変更する場合はtemplate.html、style.cssを修正してください。
