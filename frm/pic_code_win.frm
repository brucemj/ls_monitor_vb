VERSION 5.00
Begin VB.Form pic_code_win 
   Caption         =   "pic_code_win"
   ClientHeight    =   5040
   ClientLeft      =   4725
   ClientTop       =   4950
   ClientWidth     =   7305
   Icon            =   "pic_code_win.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7305
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton copy_btn 
      Caption         =   "复制代码"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox code_txt 
      Height          =   4455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "pic_code_win"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub copy_btn_Click()
    Dim str As String
    str = code_txt.text
    Clipboard.SetText str
End Sub

Private Sub Form_Activate()
    copy_btn.SetFocus
End Sub


