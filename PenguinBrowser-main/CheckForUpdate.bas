Attribute VB_Name = "CheckForUpdate"
' �ޥ� Microsoft XML, v6.0 �M Microsoft Internet Controls
Public Sub CheckForUpdates()
    Dim xmldoc As New MSXML2.DOMDocument
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim http As New WinHttp.WinHttpRequest
    Dim updateAvailable As Boolean
    Dim response As Integer

    ' �N Github �����W�� XML ��Ƹ��J�� xmlDoc ��
    http.Open "GET", "https://github.com/username/repo/releases.atom", False
    http.Send
    xmldoc.LoadXML http.ResponseText

    ' �ˬd�O�_���s����
    For Each xmlNode In xmldoc.SelectNodes("//entry")
        If xmlNode.SelectSingleNode("title").Text = "v1.0.1" Then ' ��s����
            updateAvailable = True
            Exit For
        End If
    Next

    ' �p�G���s�����A�߰ݬO�_�n��s
    If updateAvailable Then
        response = MsgBox("���s�����i�ΡA�O�_�n��s�H", vbYesNo)
        If response = vbYes Then
            ' ���}��s����
            Shell "explorer https://github.com/username/repo/releases/latest", vbNormalFocus
        Else
            ' �~�����n��
            Exit Sub
        End If
    Else
        MsgBox "�ثe�S���s�����i�ΡC"
    End If
End Sub

