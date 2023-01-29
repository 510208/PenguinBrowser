Attribute VB_Name = "CheckForUpdate"
' 引用 Microsoft XML, v6.0 和 Microsoft Internet Controls
Public Sub CheckForUpdates()
    Dim xmldoc As New MSXML2.DOMDocument
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim http As New WinHttp.WinHttpRequest
    Dim updateAvailable As Boolean
    Dim response As Integer

    ' 將 Github 頁面上的 XML 資料載入到 xmlDoc 中
    http.Open "GET", "https://github.com/username/repo/releases.atom", False
    http.Send
    xmldoc.LoadXML http.ResponseText

    ' 檢查是否有新版本
    For Each xmlNode In xmldoc.SelectNodes("//entry")
        If xmlNode.SelectSingleNode("title").Text = "v1.0.1" Then ' 更新版本
            updateAvailable = True
            Exit For
        End If
    Next

    ' 如果有新版本，詢問是否要更新
    If updateAvailable Then
        response = MsgBox("有新版本可用，是否要更新？", vbYesNo)
        If response = vbYes Then
            ' 打開更新頁面
            Shell "explorer https://github.com/username/repo/releases/latest", vbNormalFocus
        Else
            ' 繼續執行軟體
            Exit Sub
        End If
    Else
        MsgBox "目前沒有新版本可用。"
    End If
End Sub

