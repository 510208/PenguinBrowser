VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BorderStyle     =   3  '���u�T�w��ܤ��
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   3120
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":10CA
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         Caption         =   "���v"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         Caption         =   "���q"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   3390
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         Caption         =   "ĵ�i�I���n�鬰�ǳN��s�ϥΡA�T��c�榹�n��"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  '�a�k���
         Appearance      =   0  '����
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6330
         TabIndex        =   5
         Top             =   2700
         Width           =   525
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  '�a�k���
         Appearance      =   0  '����
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "MS Windows"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5055
         TabIndex        =   6
         Top             =   2340
         Width           =   1800
      End
      Begin VB.Label lblProductName 
         Appearance      =   0  '����
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Penguin Browser"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   27.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   2520
         TabIndex        =   8
         Top             =   1140
         Width           =   3825
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  '�a�k���
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         Caption         =   "���v��"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         Appearance      =   0  '����
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "���Z��510208 �X�~"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   18
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   2355
         TabIndex        =   7
         Top             =   705
         Width           =   3150
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Me.Hide
    Form1.Show
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "���� " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Me.Hide
End Sub

Private Sub Timer1_Timer()
    Call Form_KeyPress(8)
    Timer1.Interval = 0
    Timer1.Enabled = False
End Sub
