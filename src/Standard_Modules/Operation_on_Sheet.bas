Attribute VB_Name = "Operation_on_Sheet"
Option Explicit


Sub フォーム表示()
    Node_Search.Show vbModeless
    'vbModelessをつけることでフォームを消さなくてもシートの切り替えができるようになる
End Sub


Sub ClearSearchResult()

    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets("検索結果")

    ' 3列のうち最も下までデータが入っている行を取得
    lastRow = Application.Max( _
        ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "C").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "D").End(xlUp).Row)

    If lastRow >= 2 Then
        ws.Range("B2:B" & lastRow).ClearContents
        ws.Range("C2:C" & lastRow).ClearContents
        ws.Range("D2:D" & lastRow).ClearContents
    End If

End Sub



