Attribute VB_Name = "module_for_XML"
Option Explicit



'抽出結果を「検索結果」シートに出力する関数

Public Sub SearchXml( _
        ByVal xmlPath As String, _
        ByVal xpath As String, _
        ByVal ws As Worksheet, _
        ByRef rowNo As Long)

    Dim xmlDoc As Object
    Dim nodeList As Object
    Dim node As Object

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")

    xmlDoc.async = False
    xmlDoc.validateOnParse = False


'読み込んだ時になんかあかんかったらエラー理由を吐き出させる
    If xmlDoc.Load(xmlPath) = False Then

    Debug.Print xmlPath
    Debug.Print xmlDoc.parseError.reason

    Exit Sub

    End If

    Set nodeList = xmlDoc.SelectNodes(xpath)

    For Each node In nodeList

        ws.Cells(rowNo, 2).value = Dir(xmlPath)
        ws.Cells(rowNo, 3).value = GetNodePath(node)
        ws.Cells(rowNo, 4).value = GetOutputValue(node)

        rowNo = rowNo + 1

    Next node

End Sub



'ノードパス取得関数

Public Function GetNodePath(ByVal node As Object) As String

    Dim path As String
    Dim current As Object

    path = ""

    If node.NodeType = 2 Then

        path = "/@" & node.NodeName

        Set current = node.OwnerElement

    Else

        Set current = node

    End If

    Do While Not current Is Nothing

        If current.NodeType = 1 Then
            path = "/" & current.NodeName & path
        End If

        Set current = current.parentNode

    Loop

    GetNodePath = path

End Function


'抽出結果の子要素以下を取得する関数

Public Function GetOutputValue(ByVal node As Object) As String

    Dim child As Object
    Dim result As String

    Select Case node.NodeType

        Case 1 ' ELEMENT

            For Each child In node.ChildNodes
                result = result & child.XML & vbCrLf
            Next child

            GetOutputValue = Trim(result)

        Case Else

            GetOutputValue = node.Text

    End Select

End Function
