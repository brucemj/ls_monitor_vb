VERSION 5.00
Begin VB.Form ls_monitor 
   Caption         =   "ls_monitor"
   ClientHeight    =   3930
   ClientLeft      =   6090
   ClientTop       =   4380
   ClientWidth     =   7185
   ForeColor       =   &H000000FF&
   Icon            =   "ls_monitor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   7185
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton r_pc_btn 
      Caption         =   "r_pc"
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton freeze_btn 
      Caption         =   "freeze"
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton killall_btn 
      Caption         =   "killall"
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton debug_test 
      Caption         =   "调试函数"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox monitor_st 
      BackColor       =   &H00FFC0C0&
      Height          =   270
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CheckBox monitor_arg 
      Caption         =   "10秒后启动"
      Height          =   300
      Index           =   3
      Left            =   4680
      TabIndex        =   10
      Top             =   3600
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox monitor_arg 
      Caption         =   "stop"
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   9
      Top             =   3600
      Width           =   735
   End
   Begin VB.CheckBox monitor_arg 
      Caption         =   "quests"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.CheckBox monitor_arg 
      Caption         =   "windows"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton ls_start_btn 
      Caption         =   "ls_start"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton monitor 
      Caption         =   "monitor"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox log 
      BackColor       =   &H80000001&
      ForeColor       =   &H0080FF80&
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "ls_monitor.frx":1272
      Top             =   960
      Width           =   6975
   End
   Begin VB.CommandButton exit 
      Caption         =   "exit"
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton opendir 
      Caption         =   "open"
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton clear 
      BackColor       =   &H000040C0&
      Caption         =   "清除"
      Height          =   255
      Left            =   6240
      MaskColor       =   &H0000FFFF&
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox logfilebox 
      BackColor       =   &H8000000B&
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "logfile: "
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "ls_monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub form_init()
    logfilebox.text = monitor_file
    logfilebox.Locked = True
    log.Locked = True
    monitor_st.text = "monitor is stop xxxxxx"
End Sub

Private Sub debug_test_Click()
    a_debug_test log
End Sub

Private Sub Form_Activate()
    setsystime
    Me.Caption = Me.Caption & " -- " & Now & "[release 1.1]"
    '------- Auto login -------
    delay 12
    If monitor_arg(3).Value = 1 Then
        'MsgBox "5"
        If Not monitor_running Then
            monitor_st.text = "monitor is stop xxxxxx"
            monitor_start log, 50, ls_monitor
        End If
    End If
End Sub

Private Sub Form_Load()
    'Me.Move 600, Screen.Height - Me.Height
    ls_monitor.Move 7000, 3500
    public_init
    reg_dll
    login_init
    monitor_init log

    form_init
End Sub

Private Sub Form_Unload(Cancel As Integer)
    run_cmd "taskkill /f /im ls_monitor.exe "
End Sub

Private Sub freeze_btn_Click()
    debug_check = True
    hwin = dm.FindWindow("", "炉石传说")
    Call checkfreezewin(hwin, "炉石传说", log)
    Call checkfreezewin2(hwin, "炉石传说", log)
End Sub

Private Sub killall_btn_Click()
    killall log
End Sub

Private Sub log_Change()
    'log.SetFocus
    log.SelStart = Len(log.text)
End Sub

Private Sub ls_start_btn_Click()
    ls_game.Show
End Sub

Private Sub monitor_Click()
    If Not monitor_running Then
        monitor_st.text = "monitor is stop xxxxxx"
        monitor_start log, 50, ls_monitor
    End If
End Sub

Private Sub opendir_Click()
    run_cmd ("start " + logdir)
End Sub

Private Sub clear_Click()
    log.text = "monitor info:"
End Sub

Private Sub restart_Click()
    setsystime
    intogame log, ls_game

End Sub

Private Sub exit_Click()
    run_cmd "taskkill /f /im ls_monitor.exe "
    Unload Me
End Sub


Private Sub r_pc_btn_Click()
    delay 1
    pc_restart
End Sub
