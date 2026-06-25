VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Node_Search 
   Caption         =   "Node_Search"
   ClientHeight    =   10440
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13788
   OleObjectBlob   =   "Node_Search.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Node_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'=========================================
' フォルダ変更時
'=========================================


Private Sub 検索対象フォルダ_Change()

    Dim folderPath As String

    folderPath = Trim(Me.検索対象フォルダ.Text)

    Me.検索候補ファイル.Clear

    If Dir(folderPath, vbDirectory) = "" Then Exit Sub

    GetTargetFiles _
        folderPath, _
        Me.検索候補ファイル

End Sub

'=========================================
' 対象ファイル取得（サブフォルダ含む）
'=========================================

Public Sub GetTargetFiles( _
        ByVal folderPath As String, _
        ByVal lst As Object, _
        Optional ByVal relativePath As String = "")

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim subFolder As Object
    Dim ext As String

    folderPath = NormalizeFolderPath(folderPath)

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then Exit Sub

    Set folder = fso.GetFolder(folderPath)

    ' ファイル取得
    For Each file In folder.Files

        ext = LCase$(fso.GetExtensionName(file.Name))

        Select Case ext

            Case "xml", "html", "htm"

                lst.AddItem relativePath & file.Name

        End Select

    Next file

    ' サブフォルダ再帰
    For Each subFolder In folder.SubFolders

        GetTargetFiles _
            subFolder.path, _
            lst, _
            relativePath & subFolder.Name & "\"

    Next subFolder

End Sub

'フォルダパスの末尾に \ を付与する
Private Function NormalizeFolderPath(folderPath As String) As String

    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    NormalizeFolderPath = folderPath

End Function


'=========================================
' ボタンクリックイベント
'=========================================

Private Sub Filter_button_Click()
    '検索候補 → 検索対象へ
    MoveSelectedItems _
        Me.検索候補ファイル, _
        Me.検索実行ファイル
End Sub

Private Sub FilterClear_button_Click()
    '検索対象 → 検索候補へ
    MoveSelectedItems _
        Me.検索実行ファイル, _
        Me.検索候補ファイル
End Sub

Private Sub AllFilter_button_Click()
    '検索候補 → 検索対象へ全移動
    MoveAllItems _
        Me.検索候補ファイル, _
        Me.検索実行ファイル
End Sub

Private Sub AllFilterClear_button_Click()
    '検索対象 → 検索候補へ全移動
    MoveAllItems _
        Me.検索実行ファイル, _
        Me.検索候補ファイル
End Sub

Private Sub Close_button_Click()
    'フォームを閉じる
    Unload Me
End Sub





'=========================================
' XML検索実行
'=========================================

Private Sub XMLSearch_button_Click()

    Dim folderPath As String
    Dim xpath As String
    Dim fileName As String

    Dim selector As String

    Dim ws As Worksheet
    Dim rowNo As Long

    Dim i As Long

    folderPath = Trim(Me.検索対象フォルダ.Text)
    xpath = Trim(Me.検索XPath.Text)

    If folderPath = "" Then
        MsgBox "XMLフォルダを入力してください。"
        Exit Sub
    End If

    If xpath = "" Then
        MsgBox "XPathを入力してください。"
        Exit Sub
    End If
    
    
    Dim testDoc As Object

    Set testDoc = CreateObject("MSXML2.DOMDocument.6.0")
    
    On Error GoTo XPathError
    
    testDoc.SelectNodes xpath
    
    On Error GoTo 0

    
    
    Set ws = ThisWorkbook.Worksheets("検索結果")

    rowNo = 2

    folderPath = NormalizeFolderPath(folderPath)

    For i = 0 To Me.検索実行ファイル.ListCount - 1

        Me.XML_LabelProgress.Caption = _
          (i + 1) & _
          " / " & _
          Me.検索実行ファイル.ListCount

        DoEvents


        fileName = Me.検索実行ファイル.List(i)

        SearchXml _
            folderPath & fileName, _
            xpath, _
            ws, _
            rowNo

    Next i

    MsgBox "検索完了"
    
    Exit Sub

XPathError:
    
        MsgBox _
            "入力したXPathに構文ミス、もしくは誤記があります", vbExclamation, "XPath構文ミスまたは誤記"
    
        Exit Sub

End Sub

'=========================================
' HTML検索実行
'=========================================

Private Sub CssSearch_button_Click()

    Dim folderPath As String
    Dim selector As String
    Dim ws As Worksheet
    Dim rowNo As Long
    Dim i As Long
    Dim fileName As String

    folderPath = Trim(Me.検索対象フォルダ.Text)
    selector = Trim(Me.検索CSSセレクタ.Text)

    If folderPath = "" Then
        MsgBox "フォルダを入力してください。"
        Exit Sub
    End If

    If selector = "" Then
        MsgBox "CSSセレクタを入力してください。"
        Exit Sub
    End If


    Dim testHtml As Object
    
    Set testHtml = CreateObject("HTMLFILE")
    
    testHtml.Open
    testHtml.Write "<html><body></body></html>"
    testHtml.Close
    
    On Error GoTo CssError
    
    testHtml.querySelectorAll selector
    
    On Error GoTo 0


    Set ws = ThisWorkbook.Worksheets("検索結果")

    rowNo = 2

    folderPath = NormalizeFolderPath(folderPath)

    For i = 0 To Me.検索実行ファイル.ListCount - 1

        Me.HTML_LabelProgress.Caption = _
            (i + 1) & _
            " / " & _
            Me.検索実行ファイル.ListCount

        DoEvents

        fileName = Me.検索実行ファイル.List(i)

        SearchHtml _
            folderPath & fileName, _
            folderPath, _
            selector, _
            ws, _
            rowNo
    
    Next i

    MsgBox "HTML検索完了"
    
    Exit Sub

CssError:
    
        MsgBox _
            "入力したCSSセレクタに構文ミス、もしくは誤記があります", vbExclamation, "CSSセレクタ構文ミス、または誤記"
    
        Exit Sub

End Sub



'=========================================
' 検索対象ファイルを選ぶときの共通処理
'=========================================

'選択されている項目のみ移動する
Private Sub MoveSelectedItems( _
    sourceList As MSForms.ListBox, _
    targetList As MSForms.ListBox)

    Dim remainFiles As Collection
    Dim item As Variant
    Dim i As Long

    Set remainFiles = New Collection

    For i = 0 To sourceList.ListCount - 1

        If sourceList.Selected(i) Then

            If Not ExistsInListBox(targetList, sourceList.List(i)) Then
                targetList.AddItem sourceList.List(i)
            End If

        Else

            remainFiles.Add sourceList.List(i)

        End If

    Next i

    sourceList.Clear

    For Each item In remainFiles
        sourceList.AddItem item
    Next item

End Sub


'すべての項目を移動する
Private Sub MoveAllItems( _
    sourceList As MSForms.ListBox, _
    targetList As MSForms.ListBox)

    Dim i As Long

    For i = 0 To sourceList.ListCount - 1

        If Not ExistsInListBox(targetList, sourceList.List(i)) Then
            targetList.AddItem sourceList.List(i)
        End If

    Next i

    sourceList.Clear

End Sub





'リストボックス内に値が存在するか確認する
Private Function ExistsInListBox( _
    lst As MSForms.ListBox, _
    value As String) As Boolean

    Dim i As Long

    For i = 0 To lst.ListCount - 1

        If lst.List(i) = value Then

            ExistsInListBox = True
            Exit Function

        End If

    Next i

End Function

