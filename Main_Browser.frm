VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main_Browser 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   Caption         =   "Penguin Browser v1.0"
   ClientHeight    =   7815
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   13680
   Icon            =   "Main_Browser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   13680
   StartUpPosition =   2  '螢幕中央
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10200
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "複製網址(&C)"
      Height          =   375
      Left            =   10680
      TabIndex        =   7
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '平面
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   8415
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6375
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   10215
      ExtentX         =   18018
      ExtentY         =   11245
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1440
      Top             =   0
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  '平面
      Caption         =   "×"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      ToolTipText     =   "停止載入"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  '平面
      Caption         =   "⊙"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      ToolTipText     =   "重新載入"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  '平面
      Caption         =   "搜尋(&S)"
      Height          =   375
      Left            =   9840
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  '平面
      Caption         =   "→"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  '平面
      Caption         =   "←"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtTempForLog 
      Appearance      =   0  '平面
      Height          =   495
      Left            =   9000
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   8
      Text            =   "txtTempForLog"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu site_menu 
      Caption         =   "網站(&S)"
      Begin VB.Menu backpage 
         Caption         =   "上一頁(&B)"
         Shortcut        =   +^{F2}
      End
      Begin VB.Menu forwardpage 
         Caption         =   "下一頁(&N)"
         Shortcut        =   +^{F3}
      End
      Begin VB.Menu dashtext 
         Caption         =   "-"
      End
      Begin VB.Menu load 
         Caption         =   "重新載入(&R)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu cancelload 
         Caption         =   "取消載入(&C)"
      End
      Begin VB.Menu dashtext1 
         Caption         =   "-"
      End
      Begin VB.Menu openhtmlfile 
         Caption         =   "開啟HTML文件並載入(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu dashtext2 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "關於(&A)"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu manage 
      Caption         =   "管理(&M)"
      Begin VB.Menu exit 
         Caption         =   "關閉(&E)"
      End
   End
   Begin VB.Menu changuyong_site 
      Caption         =   "常用網站(&S)"
      Begin VB.Menu wpgblpc_homepage 
         Caption         =   "跟著企鵝哥學電腦：官網(&P)"
      End
      Begin VB.Menu searchengine 
         Caption         =   "搜尋引擎(&S)"
         Begin VB.Menu google 
            Caption         =   "Google(&G)"
         End
         Begin VB.Menu yahoo 
            Caption         =   "Yahoo(&Y)"
         End
         Begin VB.Menu bing 
            Caption         =   "Microsoft Bing(&B)"
         End
         Begin VB.Menu pgbsearch 
            Caption         =   "企鵝哥搜尋引擎(&P)"
         End
      End
   End
   Begin VB.Menu visit 
      Caption         =   "檢視(&V)"
      Begin VB.Menu FullScreen 
         Caption         =   "全螢幕(&F)"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu Developer 
      Caption         =   "開發者(&D)"
      Begin VB.Menu code 
         Caption         =   "...檢視網頁原始碼"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "Main_Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub DebugLog(message As String)
    Dim OldContent As String
    Dim filepath As String
    Dim FileNum As Integer
  
    '要先有這個檔案，而且要有內容喔
    filepath = App.Path & "\log\log.txt"
  
    '在畫面上放一個TextBox，命名為txtTempForLog，將Visible設定為False
    '-----↓把Log檔的舊內容讀出來，暫存在畫面上的txtTempForLog裡-----------
    FileNum = FreeFile
    txtTempForLog.Text = ""
    Open filepath For Input As #FileNum ' 開啟文字檔,開始讀出記錄

    ' 若不是空檔案,一行一行把txt讀出來放在txtTempForLog
    If EOF(FileNum) = False Then ' 判斷 Test.txt 是不是空的檔案
      Do ' TextBox容量只有32KB大檔案請用RichTextBox
        Line Input #FileNum, OldContent
        txtTempForLog.SelText = OldContent
      Loop Until EOF(FileNum)
      Close #FileNum
    End If
    '-----↑把Log檔的舊內容讀出來，暫存在畫面上的txtTempForLog裡-----------
  
  
    FileNum = FreeFile ' 先把新now寫進去再把剛讀出來的txt從txtTempForLog寫進去
    'Open filepath For Append As #FileNum  '用Append會把新內容加在後面。我要把新內容加在最前面，所以需要把舊內容暫存在畫面上的txtTempForLog裡再貼進來
    'Print #FileNum, Now & "：" & message
    Open filepath For Output As #FileNum ' 開啟文字檔,準備寫入檔案
    Print #FileNum, Now & "：" & vbCrLf & message & vbCrLf & txtTempForLog.Text  '把舊內容貼在新內容後面，寫入檔案
    Close #FileNum
End Sub
Private Sub about_Click()
Call DebugLog("Private Sub about_Click()")
frmAbout.Show
End Sub

Private Sub backpage_Click()
Call DebugLog("Private Sub backpage_Click()")
Call Command1_Click
End Sub

Private Sub bing_Click()
WebBrowser1.GoSearch
End Sub

Private Sub cancelload_Click()
Call Command5_Click
End Sub

Private Sub code_Click()
Debug.Print WebBrowser1.Document.body.innerHTML
Dev_SiteCode.code.Text = WebBrowser1.Document.body.innerHTML
Dev_SiteCode.Show
End Sub

Private Sub Command1_Click()
On Error GoTo errortext
WebBrowser1.GoBack
Exit Sub
errortext:
DebugLog ("Error:Not website can back")
End Sub

Private Sub Command2_Click()
On Error GoTo errortext
On Error Resume Next
WebBrowser1.GoForward
errortext:
DebugLog ("Error:Not website can back")
End Sub

Private Sub Command3_Click()
If Left(Text1.Text, 4) = "http" Then
    WebBrowser1.Navigate (Text1.Text)
Else
    WebBrowser1.GoSearch
End If
End Sub

Private Sub Command4_Click()
WebBrowser1.Refresh
End Sub

Private Sub Command5_Click()
WebBrowser1.Stop
End Sub

Private Sub Command6_Click()
Clipboard.SetText (WebBrowser1.LocationName)
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
Call DebugLog("8787")
WebBrowser1.Silent = True
WebBrowser1.GoSearch
'WebBrowser1.Navigate ("https://sam0616.pixnet.net")
End Sub

Private Sub Form_Resize()
WebBrowser1.Width = Main_Browser.Width
WebBrowser1.Height = Main_Browser.Width - 375
End Sub

Private Sub forwardpage_Click()
Call Command2_Click
End Sub

Private Sub FullScreen_Click()
If WebBrowser1.FullScreen Then
    WebBrowser1.FullScreen = True
Else
    WebBrowser1.FullScreen = False
End If
End Sub

Private Sub google_Click()
WebBrowser1.Navigate ("https://google.com")
End Sub

Private Sub load_Click()
Call Command4_Click
End Sub



Private Sub openhtmlfile_Click()
'CancelError 為 True。
On Error GoTo ErrHandler
'設置過濾器。
CommonDialog1.Filter = "All Files (*.*)|*.*|HTML Files (*.html)|*.html|HTML HTM Files (*.htm)|*.htm"
'指定缺省過濾器。
CommonDialog1.FilterIndex = 2
'顯示“打開”對話框。
CommonDialog1.ShowOpen
'調用打開文件的過程。
Open CommonDialog1.FileName For Output As #1
Print #1, HtmlTxt
WebBrowser1.Navigate CommonDialog1.FileName
Close
ErrHandler:
'用戶按“取消”按鈕。
Exit Sub
End Sub

Private Sub pgbsearch_Click()
WebBrowser1.Navigate ("https://cse.google.com/cse?cx=d13201a0e272750fd#gsc.tab=0")
End Sub

Private Sub Timer1_Timer()
Main_Browser.Caption = "Penguin Browser v" & App.Major & "." & App.Minor & "." & App.Revision & Space(4) & WebBrowser1.LocationName
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Debug.Print WebBrowser1.Document.body.innerHTML
End Sub

Private Sub wpgblpc_homepage_Click()
WebBrowser1.Navigate ("http://sam0616.pixnet.net/")
End Sub

Private Sub yahoo_Click()
WebBrowser1.Navigate ("https://tw.yahoo.com/")
End Sub
