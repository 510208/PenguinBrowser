VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main_Browser 
   Appearance      =   0  '����
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
   StartUpPosition =   2  '�ù�����
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10200
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�ƻs���}(&C)"
      Height          =   375
      Left            =   10680
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
      Left            =   9840
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
   Begin VB.TextBox txtTempForLog 
      Appearance      =   0  '����
      Height          =   495
      Left            =   9000
      ScrollBars      =   3  '��̬Ҧ�
      TabIndex        =   8
      Text            =   "txtTempForLog"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
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
   Begin VB.Menu visit 
      Caption         =   "�˵�(&V)"
      Begin VB.Menu FullScreen 
         Caption         =   "���ù�(&F)"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu Developer 
      Caption         =   "�}�o��(&D)"
      Begin VB.Menu code 
         Caption         =   "...�˵�������l�X"
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
  
    '�n�����o���ɮסA�ӥB�n�����e��
    filepath = App.Path & "\log\log.txt"
  
    '�b�e���W��@��TextBox�A�R�W��txtTempForLog�A�NVisible�]�w��False
    '-----����Log�ɪ��¤��eŪ�X�ӡA�Ȧs�b�e���W��txtTempForLog��-----------
    FileNum = FreeFile
    txtTempForLog.Text = ""
    Open filepath For Input As #FileNum ' �}�Ҥ�r��,�}�lŪ�X�O��

    ' �Y���O���ɮ�,�@��@���txtŪ�X�ө�btxtTempForLog
    If EOF(FileNum) = False Then ' �P�_ Test.txt �O���O�Ū��ɮ�
      Do ' TextBox�e�q�u��32KB�j�ɮ׽Х�RichTextBox
        Line Input #FileNum, OldContent
        txtTempForLog.SelText = OldContent
      Loop Until EOF(FileNum)
      Close #FileNum
    End If
    '-----����Log�ɪ��¤��eŪ�X�ӡA�Ȧs�b�e���W��txtTempForLog��-----------
  
  
    FileNum = FreeFile ' ����snow�g�i�h�A���Ū�X�Ӫ�txt�qtxtTempForLog�g�i�h
    'Open filepath For Append As #FileNum  '��Append�|��s���e�[�b�᭱�C�ڭn��s���e�[�b�̫e���A�ҥH�ݭn���¤��e�Ȧs�b�e���W��txtTempForLog�̦A�K�i��
    'Print #FileNum, Now & "�G" & message
    Open filepath For Output As #FileNum ' �}�Ҥ�r��,�ǳƼg�J�ɮ�
    Print #FileNum, Now & "�G" & vbCrLf & message & vbCrLf & txtTempForLog.Text  '���¤��e�K�b�s���e�᭱�A�g�J�ɮ�
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
'CancelError �� True�C
On Error GoTo ErrHandler
'�]�m�L�o���C
CommonDialog1.Filter = "All Files (*.*)|*.*|HTML Files (*.html)|*.html|HTML HTM Files (*.htm)|*.htm"
'���w�ʬٹL�o���C
CommonDialog1.FilterIndex = 2
'��ܡ����}����ܮءC
CommonDialog1.ShowOpen
'�եΥ��}��󪺹L�{�C
Open CommonDialog1.FileName For Output As #1
Print #1, HtmlTxt
WebBrowser1.Navigate CommonDialog1.FileName
Close
ErrHandler:
'�Τ�������������s�C
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
