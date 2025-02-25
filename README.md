# 概要
国立国会図書館 (National Diet Library) サーチのAPIを使用して、ISBNから書誌情報を取得するExcelファイルです。[NDLサーチAPIの利用規約](https://ndlsearch.ndl.go.jp/help/api)を確認した上で使用してください。

# 使用方法
["ndl-search-api-in-excel.xlsx"](https://github.com/yoichi-kawamoto/ndl-search-api-in-excel/raw/refs/heads/main/ndl-search-api-in-excel.xlsx)をダウンロードして開くと、保護ビューとセキュリティの警告が表示されるのでいずれも有効化します。A2のセルにISBNを入力すると、NDLサーチの[検索用APIと書影API](https://ndlsearch.ndl.go.jp/help/api/specifications)を利用して書誌情報と書影を取得してB2からH2のセルに表示します。ISBNの読み取りは[バーコードリーダー](https://www.iodata.jp/product/interface/barcodereader/)を利用すると便利です。A列に複数のISBNが入力されている場合、B列からH列をコピーアンドペーストすることで複数の書誌情報をまとめて取得できます。

# 使用上の注意
- [WEBSERVICE関数](https://support.microsoft.com/ja-jp/office/webservice-%E9%96%A2%E6%95%B0-0546a35a-ecc6-4739-aed7-c0b7ce1562c4)と[FILTERXML関数](https://support.microsoft.com/ja-jp/office/filterxml-%E9%96%A2%E6%95%B0-4df72efc-11ec-4951-86f5-c1374812f5b7)を使用するため、Windows版のExcel 2016以上のデスクトップアプリでのみ動作します。
- [サーバへの過負荷を避けるため](https://ndlsearch.ndl.go.jp/help/api#sec4)、書誌情報のリストを作成する場合にはこのExcelファイルに追加するのではなく、他のExcelファイルにAPIで取得した値のみを貼り付けるようにしてください（必要な項目をコピーした後、他のExcelファイルで「形式を選択して貼り付け」→「値」を選択して貼り付け）。