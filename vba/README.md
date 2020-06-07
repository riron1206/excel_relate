# エクセルVBA（マクロ）置き場

## マクロを有効化するために必要な設定・操作
- https://www.excelspeedup.com/macrosettei/
```bash
メニューから「ファイル」を選択
→「オプション」を選択
→「トラストセンター」or「セキュリティセンター」を選択
→「トラストセンターの設定」or「セキュリティセンターの設定」を選択
→「メッセージバーの設定」の確認（「メッセージバーを表示する」が選択されている）
→「マクロの設定」を選択
→「警告を表示してすべてのマクロを無効にする」を選択
→「リボンのユーザ設定」で「開発」を追加する
→最後に、いったんエクセルを閉じて、開きなおしたらマクロ使える
```

## マクロ作り方
- https://www.excelspeedup.com/hajimetenovba/
```bash
.xlsm ファイル作成
→メニューから「開発」選び「Visual Basic」を選択 or 「Alt＋F11」
→画面上のメニューから「挿入」選び「標準モジュール」を選択
→マクロとして実行したい作業をSub中に書く
例:
Sub テスト()
    Range("a1") = 1
End Sub
※作ったマクロを実行するショートカットキーは、メニューから「開発」選び「マクロ」を選択して「オプション」から指定できる
```

<!-- 
## License
This software is released under the MIT License, see LICENSE.
-->

## Author
- Github: [riron1206](https://github.com/riron1206)