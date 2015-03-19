VERSION 5.00
Begin VB.Form ls_game 
   Caption         =   "ls_login"
   ClientHeight    =   5250
   ClientLeft      =   6090
   ClientTop       =   4380
   ClientWidth     =   7185
   ForeColor       =   &H000000FF&
   Icon            =   "ls_game.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "ls_game.frx":0E42
   ScaleHeight     =   5250
   ScaleWidth      =   7185
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox login_arg 
      Caption         =   "g"
      Height          =   180
      Index           =   5
      Left            =   720
      TabIndex        =   43
      Top             =   840
      Width           =   375
   End
   Begin VB.CheckBox login_arg 
      Caption         =   "10秒后登陆"
      Height          =   300
      Index           =   4
      Left            =   5760
      TabIndex        =   42
      Top             =   4920
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox login_arg 
      Caption         =   "exit"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   41
      Top             =   1080
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox login_arg 
      Caption         =   "pic"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   40
      Top             =   720
      Width           =   975
   End
   Begin VB.CheckBox login_arg 
      Caption         =   "hb"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   39
      Top             =   480
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox login_arg 
      Caption         =   "ls"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   38
      Top             =   240
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton points 
      Caption         =   "x"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   37
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton points 
      Caption         =   "x"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   36
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton points 
      Caption         =   "x"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   35
      Top             =   4200
      Width           =   255
   End
   Begin VB.TextBox color1 
      Height          =   270
      Index           =   8
      Left            =   4800
      TabIndex        =   34
      Text            =   "0.9"
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox color1 
      Height          =   270
      Index           =   7
      Left            =   4800
      TabIndex        =   33
      Text            =   "0.9"
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox color1 
      Height          =   270
      Index           =   6
      Left            =   4800
      TabIndex        =   32
      Text            =   "0.9"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox color1 
      Height          =   270
      Index           =   5
      Left            =   1800
      TabIndex        =   28
      Text            =   "ffffff-000000|ffffff-000000"
      Top             =   4920
      Width           =   2895
   End
   Begin VB.TextBox color1 
      Height          =   270
      Index           =   4
      Left            =   0
      TabIndex        =   27
      Text            =   "0,0,0,0"
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox color1 
      Height          =   270
      Index           =   3
      Left            =   1800
      TabIndex        =   26
      Text            =   "ffffff-000000|ffffff-000000"
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox color1 
      Height          =   270
      Index           =   2
      Left            =   0
      TabIndex        =   25
      Text            =   "0,0,0,0"
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox color1 
      Height          =   270
      Index           =   1
      Left            =   1800
      TabIndex        =   24
      Text            =   "ffffff-000000|ffffff-000000"
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox color1 
      Height          =   270
      Index           =   0
      Left            =   0
      TabIndex        =   23
      Text            =   "0,0,0,0"
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton show_user 
      Caption         =   "user"
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton findwin 
      Caption         =   "移动窗口"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton exit 
      Caption         =   "exit"
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton opendir 
      Caption         =   "open"
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton clear 
      BackColor       =   &H000040C0&
      Caption         =   "清除"
      Height          =   255
      Left            =   6240
      MaskColor       =   &H0000FFFF&
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox logfilebox 
      BackColor       =   &H8000000B&
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Text            =   "logfile: "
      Top             =   1320
      Width           =   5175
   End
   Begin VB.TextBox log 
      BackColor       =   &H80000001&
      ForeColor       =   &H0080FF80&
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "ls_game.frx":1C84
      Top             =   1680
      Width           =   6975
   End
   Begin VB.CommandButton restart 
      Caption         =   "login"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   1200
      TabIndex        =   7
      Top             =   0
      Width           =   5655
      Begin VB.TextBox args 
         Height          =   270
         Index           =   1
         Left            =   2040
         TabIndex        =   31
         Text            =   "画面名1"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox args 
         Height          =   270
         Index           =   0
         Left            =   1080
         TabIndex        =   30
         Text            =   "pic_test"
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton code 
         Caption         =   "code"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton pic_find 
         Caption         =   "pic"
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   22
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton pic_find 
         Caption         =   "pic"
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   21
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton pic_find 
         Caption         =   "pic"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   20
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton pic_find 
         Caption         =   "pic"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   19
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton pic_find 
         BackColor       =   &H000000FF&
         Caption         =   "pic"
         Height          =   255
         Index           =   1
         Left            =   1080
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   18
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton pic_find 
         Caption         =   "pic"
         Height          =   255
         Index           =   0
         Left            =   120
         MaskColor       =   &H008080FF&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton site1 
         Caption         =   "4"
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton site1 
         Caption         =   "3"
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton site1 
         Caption         =   "2"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton site1 
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox win_game 
         Height          =   270
         Index           =   3
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   11
         Text            =   "384"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox win_game 
         Height          =   270
         Index           =   2
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "512"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox win_game 
         Height          =   270
         Index           =   1
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   9
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox win_game 
         Height          =   270
         Index           =   0
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "ls_game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub form_init()
    logfilebox.text = logfile
    logfilebox.Locked = True
    log.Locked = True
    
    For i = 0 To pic_find.Count - 1
        If i = 0 Then
            pic_find(i).Caption = "自定义"
        Else
            pic_find(i).Caption = "pic" & i
        End If
        pic_find(i).ToolTipText = pic_arr(i)
    Next i
End Sub


Private Sub Form_Activate()
    '------- Auto login -------
    delay 10
    If login_arg(4).Value = 1 Then
        'MsgBox "5"
        setsystime
        intogame log, ls_game
    End If
End Sub

Private Sub Form_Load()
    'Me.Move 600, Screen.Height - Me.Height
    ls_game.Move 7000, 3500
    public_init
    reg_dll
    login_init
    pic_lsinit
    
    form_init
End Sub

Private Sub Form_Unload(Cancel As Integer)
    run_cmd "taskkill /f /im ls_login.exe "
End Sub


Private Sub log_Change()
    log.SelStart = Len(log.text)
End Sub

Private Sub opendir_Click()
    run_cmd ("start " + logdir)
End Sub

Private Sub clear_Click()
    log.text = "log info:"
End Sub

Private Sub restart_Click()
    log.text = log.text & vbCrLf
    setsystime
    intogame log, ls_game
End Sub

Private Sub exit_Click()
    run_cmd "taskkill /f /im ls_login.exe "
    Unload Me
End Sub

Private Sub findwin_Click()
    btn_findwin log, _
        win_game(0).text, win_game(1).text, win_game(2).text, win_game(3).text
End Sub

Private Sub site1_Click(Index As Integer)
    Select Case Index
        Case 0
            win_game(0).text = 0: win_game(1).text = 0: win_game(2).text = 512: win_game(3).text = 384
        Case 1
            win_game(0).text = 0: win_game(1).text = 0: win_game(2).text = 768: win_game(3).text = 576
        Case 2
            win_game(0).text = -820: win_game(1).text = -740: win_game(2).text = 1024: win_game(3).text = 768
        Case 3
            win_game(0).text = 0: win_game(1).text = 0: win_game(2).text = 512: win_game(3).text = 384
    End Select
End Sub

Private Sub pic_find_Click(Index As Integer)
    btn_findpic log, Index, color1
End Sub

Private Sub code_Click()
    pic_createCode args, color1, log, pic_code_win
End Sub

Private Sub points_Click(Index As Integer)
    i = Index * 2
    color1(i).text = ""
    color1(i + 1).text = ""
    color1(Index + 6).text = ""
End Sub

