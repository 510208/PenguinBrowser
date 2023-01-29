VERSION 5.00
Begin VB.Form Dev_SiteCode 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   Caption         =   "PBrowser v1.0 _ [開發工具] - 程式碼檢視"
   ClientHeight    =   3015
   ClientLeft      =   2235
   ClientTop       =   7290
   ClientWidth     =   5160
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   5160
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.TextBox code 
      Appearance      =   0  '平面
      Height          =   3015
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   0
      Text            =   "Dev_SiteCode.frx":0000
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "Dev_SiteCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
code.Text = Main_Browser.WebBrowser1.Document.body.innerHTML
End Sub

Private Sub Form_Resize()
code.Width = Dev_SiteCode.Width
code.Height = Dev_SiteCode.Width
End Sub
