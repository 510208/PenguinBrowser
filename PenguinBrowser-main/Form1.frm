VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Penguin Browser v1.0"
   ClientHeight    =   7815
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   13680
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command6 
      Caption         =   "�ƻs���}(&C)"
      Height          =   375
      Left            =   11280
      TabIndex        =   7
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '����
      BeginProperty Font 
         Name            =   "�L�n������"
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
      Appearance      =   0  '����
      Caption         =   "��"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      ToolTipText     =   "������J"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  '����
      Caption         =   "��"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      ToolTipText     =   "���s���J"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  '����
      Caption         =   "�j�M(&S)"
      Height          =   375
      Left            =   10440
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  '����
      Caption         =   "��"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  '����
      Caption         =   "��"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Menu site_menu 
      Caption         =   "����(&S)"
      Begin VB.Menu backpage 
         Caption         =   "�W�@��(&B)"
         Shortcut        =   +^{F2}
      End
      Begin VB.Menu forwardpage 
         Caption         =   "�U�@��(&N)"
         Shortcut        =   +^{F3}
      End
      Begin VB.Menu dashtext 
         Caption         =   "-"
      End
      Begin VB.Menu load 
         Caption         =   "���s���J(&R)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu cancelload 
         Caption         =   "�������J(&C)"
      End
      Begin VB.Menu dashtext1 
         Caption         =   "-"
      End
      Begin VB.Menu openhtmlfile 
         Caption         =   "�}��HTML���ø��J(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu dashtext2 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "����(&A)"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu manage 
      Caption         =   "�޲z(&M)"
      Begin VB.Menu exit 
         Caption         =   "����(&E)"
      End
   End
   Begin VB.Menu changuyong_site 
      Caption         =   "�`�κ���(&S)"
      Begin VB.Menu wpgblpc_homepage 
         Caption         =   "��ۥ��Z���ǹq���G�x��(&P)"
      End
      Begin VB.Menu searchengine 
         Caption         =   "�j�M����(&S)"
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
            Caption         =   "���Z���j�M����(&P)"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub backpage_Click()
Call Command1_Click
End Sub

Private Sub bing_Click()
WebBrowser1.GoSearch
End Sub

Private Sub cancelload_Click()
Call Command5_Click
End Sub

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

Private Sub Command6_Click()
Clipboard.SetText (WebBrowser1.LocationName)
End Sub

Private Sub exit_Click()
End
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

Private Sub forwardpage_Click()
Call Command2_Click
End Sub

Private Sub google_Click()
WebBrowser1.Navigate ("https://google.com")
End Sub

Private Sub load_Click()
Call Command4_Click
End Sub

Private Sub pgbsearch_Click()
WebBrowser1.Navigate ("https://cse.google.com/cse?cx=d13201a0e272750fd#gsc.tab=0")
End Sub

Private Sub Timer1_Timer()
Form1.Caption = "Penguin Browser v" & App.Major & "." & App.Minor & "." & App.Revision & Space(4) & WebBrowser1.LocationName
End Sub

Private Sub wpgblpc_homepage_Click()
WebBrowser1.Navigate ("http://sam0616.pixnet.net/")
End Sub

Private Sub yahoo_Click()
WebBrowser1.Navigate ("https://tw.yahoo.com/")
End Sub
