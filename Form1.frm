VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   13680
   StartUpPosition =   3  '系統預設值
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
      Left            =   2040
      TabIndex        =   2
      Text            =   "https://sam0616.pixnet.net"
      Top             =   0
      Width           =   8415
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2295
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   4575
      ExtentX         =   8070
      ExtentY         =   4048
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
      Location        =   ""
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
      Left            =   10440
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
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
WebBrowser1.GoBack
End Sub

Private Sub Command2_Click()
WebBrowser1.GoForward
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

Private Sub Form_Load()
WebBrowser1.Silent = True
WebBrowser1.GoSearch
'WebBrowser1.Navigate ("https://sam0616.pixnet.net")
End Sub

Private Sub Form_Resize()
WebBrowser1.Width = Form1.Width
WebBrowser1.Height = Form1.Width - 375
End Sub

Private Sub Timer1_Timer()
'If WebBrowser1.LocationName <> Text1.Text Then
'    Text1.Text = WebBrowser1.LocationName
'End If
End Sub
