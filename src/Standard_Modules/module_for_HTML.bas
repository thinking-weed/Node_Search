Attribute VB_Name = "module_for_html"
Option Explicit

'=========================================
' HTML検索関数
'=========================================

Public Sub SearchHtml( _
        ByVal htmlPath As String, _
        ByVal baseFolderPath As String, _
        ByVal selector As String, _
        ByVal ws As Worksheet, _
        ByRef rowNo As Long)

    Dim htmlDoc As Object
    Dim nodeList As Object
    Dim node As Object
    Dim idx As Long

    Dim fileText As String
    Dim ff As Integer

    Dim stm As Object
    Dim displayPath As String



    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2
    stm.Charset = "UTF-8"

    stm.Open
    stm.LoadFromFile htmlPath
    
    fileText = stm.ReadText
    
    stm.Close

    Set stm = Nothing
    
    Set htmlDoc = CreateObject("HTMLFILE")

    If htmlDoc Is Nothing Then
        MsgBox "HTMLFILE作成失敗"
        Exit Sub
    End If


    htmlDoc.Open
    
    
    htmlDoc.Write fileText
    
    
    htmlDoc.Close


    Set nodeList = htmlDoc.querySelectorAll(selector)

    
    For idx = 0 To nodeList.Length - 1

    displayPath = Mid$(htmlPath, Len(baseFolderPath) + 1)
    
    Set node = nodeList.item(idx)
        ws.Cells(rowNo, 2).value = displayPath
        'Mid$(htmlPath, InStrRev(htmlPath, "\") + 1)
        ws.Cells(rowNo, 3).value = GetHtmlNodePath(node)
        
        ' タグ込みで出力
        ws.Cells(rowNo, 4).value = node.outerHTML

        
        rowNo = rowNo + 1

    Next idx

End Sub



'=========================================
'HTMLのノードパス取得関数
'=========================================


Public Function GetHtmlNodePath(ByVal node As Object) As String

    Dim current As Object
    Dim path As String
    Dim tag As String

    Set current = node

    Do While Not current Is Nothing

        tag = ""

        On Error Resume Next
        tag = current.tagName
        If Err.Number <> 0 Then Exit Do
        On Error GoTo 0

        If LCase$(tag) <> "html" Then
            path = "/" & LCase$(tag) & path
        End If

        On Error Resume Next
        Set current = current.parentNode
        If Err.Number <> 0 Then Exit Do
        On Error GoTo 0

    Loop

    GetHtmlNodePath = "/html" & path

End Function
