############################### よく使いそうな関数 ###############################
Sub A1セルを選択して保存() 'Subはmain関数に相当するもの
    Dim WB As Workbook                   'ワークブックを格納する変数宣言
    Set WB = ActiveWorkbook              'ActiveWorkBook：現在表示しているワークブック
    Dim WS As Worksheet                  'ワークシートを格納する変数宣言
    For Each WS In WB.Worksheets         'ブック内のワークシートの集合（Worksheets)からワークシートを1つずつWSに格納'
        WS.Activate                      'ワークシートを選択（アクティブ化）
        WS.Cells(1, 1).Select            'A1セル選択
    Next
    WB.Worksheets(1).Activate            '最初のシート選択
    WB.Save                              ' 上書き保存
End Sub

Sub シート全体のフォント変更()
    Cells.Select                        'Cellsでシート全体指定
    With Selection.Font                 'With Selection は選択範囲の意味。Withは特定の対象に対して一括して操作を行いたい場合に用いる
        .Name = "メイリオ"                'フォント名
        .Size = 11                        'フォントのサイズ
        .Strikethrough = False            '取り消し線削除
        .Superscript = False              '上付き文字にしない
        .Subscript = False                '下付き文字にしない
        .OutlineFont = False              'アウトラインフォントにしない
        .Shadow = False                   '影付きフォントにしない
        .Underline = xlUnderlineStyleNone '下線なし
        .ThemeColor = xlThemeColorLight1  'セルの配色のテーマカラー
        .TintAndShade = 0                 '色を明るく、または暗く しない
        .ThemeFont = xlThemeFontNone      'テーマのフォント
    End With
End Sub

Sub シート全体について折り返して全体を表示を解除する()
    Cells.WrapText = False
End Sub

Sub シート全体の文字色黒にする()
    Cells.Font.ColorIndex = 1
End Sub

Sub シート全体について結合セルを解除して値を埋める
  For Each rng In ActiveSheet.UsedRange
    If rng.MergeCells Then
      With rng.MergeArea
        .UnMerge
        .Value = .Resize(1, 1).Value
      End With
    End If
  Next
End Sub


Sub ブックにあるすべてのシートに同じ処理を行う()
    Dim Ws As Worksheet                      'ワークシート(Worksheetオブジェクト)を格納する変数宣言
    For Each Ws In Worksheets                'ブック内のワークシートの集合（Worksheets)からワークシートを1つずつWSに格納
        Ws.Activate                           'ワークシートを選択（アクティブ化）
        Call シート全体の文字色黒にする       '別のSub 関数を呼び出す
        Next Ws
End Sub

Sub ブックにある特定のシートのみ処理する()
    Dim Ws As Worksheet
    For Each Ws In Worksheets
        Ws.Activate
        If Ws.Name Like "*S*" Then   '「S」を含むシート名なら処理
            Call シート全体について折り返して全体を表示を解除する
        End If
    Next Ws
End Sub


Function 文字列の前と後ろセルの先頭と末尾の改行を削除(文字列 As String) As String
    Dim strTmp As String
    strTmp = 文字列

    Do Until Left(strTmp, 1) <> vbLf
        strTmp = Mid(strTmp, 2)
    Loop
    Do Until Right(strTmp, 1) <> vbLf
        strTmp = Left(strTmp, Len(strTmp) - 1)
    Loop

    TrimLF = strTmp
End Function


############################### 取り消し線についての関数 ###############################
' https://stabucky.com/wp/archives/3209
Sub 選択範囲の取り消し線の付いた文字を削除()
    For Each myCell In Selection
        textBefore = myCell.Value
        textAfter = ""
        For i = 1 To Len(textBefore)
            If myCell.Characters(Start:=i, Length:=1).Font.Strikethrough = False Then
                textAfter = textAfter & Mid(textBefore, i, 1)
            End If
        Next i
        myCell.Value = textAfter
    Next myCell
End Sub

' http://www.excel.studio-kazu.jp/kw/20150711074705.html
 Sub 同フォルダにあるエクセルの取り消し線持つセル一覧取得()
    Dim w As Variant
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim c As Range
    Dim r As Range
    Dim z As Variant
    Dim flag As Boolean
    Dim shT As Worksheet
    Dim fPath As String
    Dim fName As String
    Application.ScreenUpdating = False
    Set shT = ThisWorkbook.Sheets("取り消し線一覧")
    fPath = ThisWorkbook.Path & "\"
    fName = Dir(fPath & "*.xlsx")
    Do While fName <> ""
        Set wb = Workbooks.Open(fPath & fName)
        For Each sh In wb.Worksheets
            Set r = Nothing
            On Error Resume Next
            Set r = sh.UsedRange.SpecialCells(xlCellTypeConstants)
            On Error GoTo 0
            If Not r Is Nothing Then
                For Each c In r
                    flag = False
                    z = c.Characters.Font.Strikethrough
                    If IsNull(z) Then
                        flag = True
                    ElseIf z = True Then
                        flag = True
                    End If
                    If flag Then
                        If IsArray(w) Then
                            ReDim Preserve w(1 To 3, 1 To UBound(w, 2) + 1)
                        Else
                            ReDim w(1 To 3, 1 To 1)
                        End If
                        w(1, UBound(w, 2)) = sh.Parent.Name
                        w(2, UBound(w, 2)) = sh.Name
                        w(3, UBound(w, 2)) = c.Address(False, False)
                    End If
                Next
            End If
        Next
        wb.Close False
        fName = Dir()
    Loop
    shT.Cells.ClearContents
    shT.Range("A1:C1").Value = Array("ブック名", "シート名", "セル")
    shT.Range("A2").Value = "取り消し線付セルはありません"
    If IsArray(w) Then
        shT.Range("A2").Resize(UBound(w, 2), 3).Value = WorksheetFunction.Transpose(w)
        shT.Hyperlinks.Delete
        For Each c In shT.Range("A2", shT.Range("A" & Rows.Count).End(xlUp))
            shT.Hyperlinks.Add Anchor:=c, Address:=c.Value
            shT.Hyperlinks.Add Anchor:=c.Offset(, 1), Address:=c.Value, _
                SubAddress:="'" & c.Offset(, 1).Value & "'!A1"
            shT.Hyperlinks.Add Anchor:=c.Offset(, 2), Address:=c.Value, _
                SubAddress:="'" & c.Offset(, 1).Value & "'!" & c.Offset(, 2).Value
        Next
    End If
    shT.Select
 End Sub

' http://hikaridisk.blogspot.com/2012/03/blog-post_19.html
Sub 複数エクセルファイルの取消文字列を一括削除する()

    Dim objBook As Variant
    Dim objSheet As Variant 'シート
    Dim strFile As String

    Dim intRow As Integer
    Dim intLastRow As Integer

    Dim intCount As Integer
    Dim celCell As Variant

    Dim chaChar As Characters
    Dim strResult As String
    Dim strDelWord As String
    Dim strRange As String

    If ThisWorkbook.Sheets("取り消し線削除するファイル一覧").Range("A2").Value = "" Then Exit Sub

    '最終行を求める
    intLastRow = ThisWorkbook.Sheets("取り消し線削除するファイル一覧").Range("A1").End(xlDown).Row

    '画面更新を無効にする
    Application.ScreenUpdating = False

    '削除処理の範囲
    strRange = ThisWorkbook.Sheets("取り消し線削除するファイル一覧").Range("F1").Value

    intRow = 1
    ThisWorkbook.Sheets("取り消し線削除した結果").Cells.Clear
    ThisWorkbook.Sheets("取り消し線削除した結果").Cells(intRow, 1).Value = "処理結果一覧"

    intRow = 2
    ThisWorkbook.Sheets("取り消し線削除した結果").Cells(intRow, 1).Value = "削除前"
    ThisWorkbook.Sheets("取り消し線削除した結果").Cells(intRow, 2).Value = "削除文字"
    ThisWorkbook.Sheets("取り消し線削除した結果").Cells(intRow, 3).Value = "削除後"
    ThisWorkbook.Sheets("取り消し線削除した結果").Cells(intRow, 4).Value = "セル"
    ThisWorkbook.Sheets("取り消し線削除した結果").Cells(intRow, 5).Value = "シート"
    ThisWorkbook.Sheets("取り消し線削除した結果").Cells(intRow, 6).Value = "ファイル"


    'メッセージ表示を無効にする
    Application.DisplayAlerts = False

    'ファイル毎の処理
    For i = 2 To intLastRow

        strFile = ThisWorkbook.Sheets("取り消し線削除するファイル一覧").Cells(i, 1).Value

        If Right(strFile, 3) = "xls" Or Right(strFile, 4) = "xlsx" Then

            'ファイルをセットする
            Set objBook = Application.Workbooks.Open(strFile)

            'シート毎の処理
            For Each objSheet In objBook.Sheets

                'セル毎の処理
                For Each celCell In objSheet.Range(strRange)

                    '数字だと、エラーになるので、文字列だけ処理する
                    If VarType(celCell.Value) = vbString Then

                      intCount = Len(celCell)
                      strResult = ""
                      strDelWord = ""

                      For j = 1 To intCount

                          Set chaChar = celCell.Characters(j, 1)

                          If chaChar.Font.Strikethrough Then

                             strDelWord = strDelWord + chaChar.Text

                          Else

                            strResult = strResult + chaChar.Text

                          End If

                      Next j

                      If Len(strDelWord) > 0 Then

                          '削除対象文字の一覧を作成する
                          intRow = intRow + 1
                          ThisWorkbook.Sheets("取り消し線削除した結果").Cells(intRow, 1).Value = celCell.Value
                          ThisWorkbook.Sheets("取り消し線削除した結果").Cells(intRow, 2).Value = Trim(strDelWord)
                          ThisWorkbook.Sheets("取り消し線削除した結果").Cells(intRow, 3).Value = Trim(strResult)
                          ThisWorkbook.Sheets("取り消し線削除した結果").Cells(intRow, 4).Value = celCell.Row _
                                                               & "行" & celCell.Column & "列"
                          ThisWorkbook.Sheets("取り消し線削除した結果").Cells(intRow, 5).Value = objSheet.Name
                          ThisWorkbook.Sheets("取り消し線削除した結果").Cells(intRow, 6).Value = objBook.Name

                          '削除後の文字をセットする
                          celCell.Value = Trim(strResult)

                      End If

                    End If

                Next celCell

            Next objSheet

            'ファイルクローズ
            objBook.Close savechanges:=True

        End If

    Next i

    'メッセージ表示を有効にする
    Application.DisplayAlerts = True

    '画面更新を有効に戻す
    Application.ScreenUpdating = True

    'シート処理結果をアクティブにする
    ThisWorkbook.Sheets("取り消し線削除した結果").Activate

    ' 処理完了(結果表示)
    MsgBox "処理が完了しました。"
End Sub


############################### 関数ではないがよく使いそうな処理メモ ###############################
'RangeとCellsは同じもの。書き方違うだけ
Range("B3").Value = 1
Cells(3, 2).Value = 1

'For~Nextループで「縦方向」に選択セルを変化させる
For Row = 1 To 10
    Range("B" & Row).Value = 1
Next

For Row = 1 To 10
  Cells(Row, 2).Value = 1
Next

'For~Nextループで「横方向」に選択セルを変化させる
For Column = 1 To 10
    Cells(3, Column).Value = 1
Next

'セル範囲指定
Range("B3:D6").Value = 1

‘ 書き込むシート指定
rowMin = 3
columnMin = 2
rowMax = 6
columnMax = 4
With Worksheets("書き込みたいシート")
  .Range(.Cells(rowMin, columnMin), .Cells(rowMax, columnMax)).Value = 1
End With

'2-4行削除
Range("2:4").Delete

'B-D列削除
Range("B:D").Delete

'ブックを編集
wb.Sheets(1).Range("A1").Value = "編集"

'IF文
If a<= 10 Then
  処理A
Else If a <=20 Then
  処理B
Else
  処理C
End If