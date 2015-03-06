VERSION 5.00
Begin VB.Form pic_code_win 
   Caption         =   "pic_code_win"
   ClientHeight    =   5040
   ClientLeft      =   4725
   ClientTop       =   4950
   ClientWidth     =   7800
   Icon            =   "pic_code_win.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7800
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton exm_code 
      Caption         =   "示例代码1"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton run_code 
      Caption         =   "运行代码"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton copy_btn 
      Caption         =   "复制代码"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox code_txt 
      BackColor       =   &H00FFC0C0&
      Height          =   4455
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "pic_code_win"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As Long, ByVal Unknownn1 As Long, ByVal Unknownn2 As Long, ByVal fCheckOnly As Long) As Long

Private Sub copy_btn_Click()
    Dim str As String
    str = code_txt.text
    Clipboard.SetText str
End Sub

Private Sub exm_code_Click()
    code_txt.text = "code_win = dm.findwindow(" & Chr(34) & Chr(34) & "," & Chr(34) & "代码窗口ls" & Chr(34) & ")" & " '//# 示例：仅支持顺序代码,不支持if，for等"
    code_txt.text = code_txt.text & vbCrLf & "MsgBox code_win"
End Sub

Private Sub Form_Activate()
    copy_btn.SetFocus
    pic_code_win.Caption = "代码窗口ls"
    code_txt.text = "code_win = dm.findwindow(" & Chr(34) & Chr(34) & "," & Chr(34) & "代码窗口ls" & Chr(34) & ")" & " '//# 示例：仅支持顺序代码,不支持if，for等"
    code_txt.text = code_txt.text & vbCrLf & "MsgBox code_win"
End Sub

Private Sub run_code_Click()
    Dim code_str() As String
    code_str = Split(code_txt.text, vbCrLf)
    code_str_length = UBound(code_str) - LBound(code_str)
    For i = 0 To code_str_length
        StepLine code_str(i)
    Next i
End Sub

  
Sub StepLine(ByVal cmd As String)
    Call EbExecuteLine(StrPtr(ByVal cmd), 0, 0, 0)
End Sub
