Attribute VB_Name = "public"
Public dm, logdir
Public hwin
Public Const title_login = "ս����¼"
Public Const title_protocol = "ս��Э��"
Public Const title_nickname = "ս���ǳƴ���"
Public Const title_cannotuse = "ս���޷�ʹ��"
Public Const title_Battle = "ս��"
Public Const title_game = "¯ʯ��˵"
Public Const title_hbcnf = "Configuration Window"
Public Const title_hbupdate = "Hearthbuddy Update Available"
'Public Const title_hb = "HearthbuddyBETA (0.3.799.88) [0.3.799.88]"
Public Const title_hb = "Hearthbuddy (0.3.857.132) [0.3.857.132]"
'Const class_hb = HwndWrapper
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '�����ӳ٣���λ������
Public Declare Function IsHungAppWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Sub public_init()
    Currentdate = Year(Now) & "-" & Month(Now) & "-" & Day(Now)
    logdir = "D:\wg\log\" & Currentdate
End Sub
Public Function delay(PauseTime As Integer) '�������ӳ٣���λ ��,���ǹرմ��ڿ��ܻ�����
    Dim WshShell
    Set WshShell = CreateObject("wscript.Shell")
    WshShell.Run "cmd.exe /c ping 127.0.0.1 -n " & PauseTime, 0, 1
End Function
Public Sub reg_dll()
    'Timer1.Interval = 1000 '����1��ʱ��
    'Dim dm As Object      '���� DM Ϊ�ؼ�����
    Shell "regsvr32 dll\dm.dll /s" 'ע���Į�����ϵͳ
    Set dm = CreateObject("dm.dmsoft") '������Į�������
    'Shell "regsvr32 Window.dll /s" 'ע���Į�����ϵͳ
    'Set dm = CreateObject("Window") '������Į�������
End Sub

Public Function btn_findwin(log, X, Y, sizeW, sizeH)
    lswin = dm.FindWindow("", title_game)
    If lswin = 0 Then
        showlogbox log, title_game & " ����û�ҵ�"
        but_findwin = 0
    Else
        dm.SetWindowState lswin, 1: Sleep 200
        dm.SetWindowSize lswin, sizeW, sizeH: Sleep 200
        dm.MoveWindow lswin, X, Y: Sleep 200
        'log_info = title_game & " �����Ѽ�����ƶ���ָ��λ��:" & X & "-" & Y & "����СΪ:" & sizeW & "-" & sizeH
        showlogbox log, title_game & " �����Ѽ�����ƶ���ָ��λ��:" & X & "-" & Y & "����СΪ:" & sizeW & "-" & sizeH
        but_findwin = 1
    End If
End Function

Public Sub showlog(logbox, text As String)
    t1 = logbox.text
    logbox.text = t1 & vbCrLf & text
    run_cmd "md " & logdir
    run_cmd "echo " & Chr(34) & text & Chr(34) & ">>" & logfile
End Sub

Public Sub showlogbox(logbox, text As String)
    t1 = logbox.text
    logbox.text = t1 + vbCrLf + text
End Sub

Public Sub show_monitorlog(logbox, text As String, Optional sync_file As Boolean = True)
    t1 = logbox.text
    
    logbox.text = t1 & vbCrLf & text
    If sync_file Then
        run_cmd "md " & logdir
        run_cmd "echo " & Chr(34) & text & Chr(34) & ">>" & monitor_file
    End If
End Sub
Public Sub show_monitorlogbox(logbox, text As String)
    t1 = logbox.text
    logbox.text = t1 & vbCrLf & text
End Sub

Public Function run_cmd(cmd)
    Dim WshShell
    Set WshShell = CreateObject("wscript.Shell")
    WshShell.Run "cmd.exe /c " & cmd, 0
End Function
Public Sub monitor_task(log, looptime As Integer)
    showlog log, "monitor_task is running ------------------------  " & Now
    If gamerun = 0 Then
        showlog log, "��Ϸ��û������������ʱ 30 ��" & Now
        delay 30
        showlog log, Now
    End If
    
    windows_st = game_windows(log)
    If windows_st = 1 Then
        showlog log, "windows_st=1��game_restart  " & Now
        game_restart
    End If
    If windows_st = 2 Then
        showlog log, "windows_st=2��game_stop  " & Now
        game_stop
    End If
    
    showlog log, looptime & " �������´�monitor"
    delay looptime
End Sub
Public Function game_windows(log)
    oopwin = dm.FindWindow("", "Oops!")
    If oopwin <> 0 Then
        showlog log, "��Ϸ����Oops!�쳣 !!!!!!!!!!!!!!!"
        oopwin = dm.FindWindow("", "Oops!")
        delay 2: dm.SetWindowState oopwin, 13
        game_windows = 1: Exit Function
    End If
    hb_license = dm.FindWindow("#32770", "Error")
    If hb_license <> 0 Then
        showlog log, "hb_license �����쳣�����������!!!!!!!!!!!!!!!"
        delay 2
        pc_restart
        game_windows = 2: Exit Function
    End If
    If gamerun = 1 And hbrun = 1 Then
        gamewin = dm.FindWindow("", "¯ʯ��˵")
        'checkfreezewin gamewin
        hbwin = dm.FindWindow("", title_hb)
        If hbwin = 0 Then
            showlog log, "hb�����쳣����������!!!!!!!!!!!!!!!"
            game_windows = 1: Exit Function
        End If
    End If
    
    'gamefrozen = dm.FindWindow("", "¯ʯ��˵ (δ��Ӧ)")
    'If gamefrozen > 0 Then
    '    checkfreezewin gamefrozen
    'End If
    
    game_windows = 0 ' 0 is OK , 1 is restart , 2 is stop
End Function
Public Function game_stop()
    
End Function
Public Function game_restart()
End Function
Public Function pc_restart()
    run_cmd "shurdown -r -f -t 3"
End Function
Public Function hbrun()
End Function
Public Function gamerun()
End Function



Public Function hb_LeftClick3()
    dm.LeftClick: Sleep 300: dm.LeftClick: Sleep 300: dm.LeftClick: Sleep 300
End Function

Public Sub setsystime()
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    'objHTTP.Open "PUT", "http://hk.kkapks.com/hb/time.php", False
    objHTTP.Open "PUT", "http://172.21.12.59/hb/time.php", False
    objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    result = objHTTP.Send()
    GetDataFromURL = objHTTP.ResponseText
    strarr = Split(GetDataFromURL, ",", 2)
    Dim WshShell
    Set WshShell = CreateObject("wscript.Shell")
    WshShell.Run "cmd.exe /c date " + strarr(0), 0
    WshShell.Run "cmd.exe /c time " + strarr(1), 0
End Sub

Public Function sql_getplayuser(act, st, glod, log)
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    'objHTTP.Open "POST", "http://172.21.12.59/hb/test.php", False
    objHTTP.Open "POST", "http://172.21.12.59/hb/play.php", False
    objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
 
    result = objHTTP.Send("act=" & act & "&st=" & st & "&glod=" & glod & "&localip=" & getIp)
    GetDataFromURL = objHTTP.ResponseText
    strarr = Split(GetDataFromURL, ",", 2)
    If st = 122 Then
        showlog log, "sql_getplayuser:" & "ս���˺Ű�ȫ����ʧ��,122" & " �û���: " & strarr(0) & "   ; ����: " & strarr(1)
    Else
        showlog log, "sql_getplayuser:" & st & " �û���: " & strarr(0) & "   ; ����: " & strarr(1)
    End If
    
    sql_getplayuser = strarr
End Function

Public Function killall(log)
    showlog log, "�ر����д��ڣ� " & Now
    win = dm.FindWindow("", title_login): Sleep 200: dm.SetWindowState win, 0   '�ر� ս����¼
    win = dm.FindWindow("", title_protocol): Sleep 200: dm.SetWindowState win, 0  '�ر� ս��Э��
    win = dm.FindWindow("", title_nickname): Sleep 200: dm.SetWindowState win, 0  '�ر� ս���ǳƴ���
    win = dm.FindWindow("", title_cannotuse): Sleep 200: dm.SetWindowState win, 0  '�ر� ս���޷�ʹ��
    win = dm.FindWindow("", title_Battle): Sleep 200: dm.SetWindowState win, 0  '�ر� ս��
    win = dm.FindWindow("", title_game): Sleep 200: dm.SetWindowState win, 13  '�ر� ¯ʯ��˵
    win = dm.FindWindow("", title_game): Sleep 200: dm.SetWindowState win, 13  '�ر� ¯ʯ��˵
    win = dm.FindWindow("", "ս���˺Ű�ȫ"): Sleep 200: dm.SetWindowState win, 13  '�ر� ս����ȫ����
    oopwin = dm.FindWindow("", "Oops!"): Sleep 200: dm.SetWindowState oopwin, 0
    Sleep 200: dm.SetWindowState oopwin, 13
    'win = dm.FindWindow("", title_hkj) : Delay ss * 10 : dm.SetWindowState win, 0'�ر� HKJ
    win = dm.FindWindow("", title_hbcnf): Sleep 200: dm.SetWindowState win, 13  '�ر� hbcnf
    win = dm.FindWindow("", title_hbupdate): Sleep 200: dm.SetWindowState win, 0  '�ر� hbupdate
    win = dm.FindWindow("", title_hb): Sleep 200: dm.SetWindowState win, 13  '�ر� hb
    win = dm.FindWindow("", "Hearthbuddy"): Sleep 200: dm.SetWindowState win, 13  '�ر� Hearthbuddy exception
    run_cmd "taskkill /f /im ls_login.exe": Sleep 200
    run_cmd "taskkill /f /im ls_login.exe": Sleep 200
    'Call CreateObject("WScript.Shell").run("D:\wg\flush.cmd") //ˢ��������
    delay 2
    If dm.FindWindow("", title_hb) > 0 Then ' killall ʧ�ܣ�����pc
        delay 5
        If dm.FindWindow("", title_hb) > 0 Then
            delay 5
            pc_restart
        End If
    End If
End Function

Public Function CheckExeIsRun(exeName As String)
    Dim WMI
    Dim Obj
    Dim Objs
    CheckExeIsRun = False
    Set WMI = GetObject("WinMgmts:")
    Set Objs = WMI.InstancesOf("Win32_Process")
    For Each Obj In Objs
      If (InStr(UCase(exeName), UCase(Obj.Description)) <> 0) Then
            CheckExeIsRun = True
            If Not Objs Is Nothing Then Set Objs = Nothing
            If Not WMI Is Nothing Then Set WMI = Nothing
            Exit Function
      End If
    Next
    If Not Objs Is Nothing Then Set Objs = Nothing
    If Not WMI Is Nothing Then Set WMI = Nothing
    Exit Function

End Function