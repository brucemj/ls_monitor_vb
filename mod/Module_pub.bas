Attribute VB_Name = "public"
Public dm, logdir, pc_tag
Public hwin
Public Const title_login = "战网登录"
Public Const title_protocol = "战网协议"
Public Const title_nickname = "战网昵称创建"
Public Const title_cannotuse = "战网无法使用"
Public Const title_Battle = "战网"
Public Const title_game = "炉石传说"
Public Const title_hbcnf = "Configuration Window"
Public Const title_hbupdate = "Hearthbuddy Update Available"
'Public Const title_hb = "HearthbuddyBETA (0.3.799.88) [0.3.799.88]"
'Public Const title_hb = "Hearthbuddy (0.3.857.132) [0.3.857.132]"
Public Const title_hb = "Hearthbuddy (0.3.859.133) [0.3.859.133]"
'Const class_hb = HwndWrapper
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '阻塞延迟，单位：毫秒
Public Declare Function IsHungAppWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Sub public_init()
    Currentdate = Year(Now) & "-" & Month(Now) & "-" & Day(Now)
    logdir = "D:\wg\log\" & Currentdate
    pc_tag = 0
End Sub
Public Function delay(PauseTime As Integer) '非阻塞延迟，单位 秒,但是关闭窗口可能会阻塞
    Dim WshShell
    Set WshShell = CreateObject("wscript.Shell")
    WshShell.Run "cmd.exe /c ping 127.0.0.1 -n " & PauseTime, 0, 1
End Function
Public Sub reg_dll()
    'Timer1.Interval = 1000 '设置1秒时钟
    'Dim dm As Object      '声明 DM 为控件类型
    Shell "regsvr32 dll\dm.dll /s" '注册大漠插件到系统
    Set dm = CreateObject("dm.dmsoft") '创建大漠插件对象
    'Shell "regsvr32 Window.dll /s" '注册大漠插件到系统
    'Set dm = CreateObject("Window") '创建大漠插件对象
End Sub

Public Function btn_findwin(log, X, Y, sizeW, sizeH)
    lswin = dm.FindWindow("", title_game)
    If lswin = 0 Then
        showlogbox log, title_game & " 窗口没找到"
        but_findwin = 0
    Else
        dm.SetWindowState lswin, 1: Sleep 200
        dm.SetWindowSize lswin, sizeW, sizeH: Sleep 200
        dm.MoveWindow lswin, X, Y: Sleep 200
        'log_info = title_game & " 窗口已激活，且移动到指定位置:" & X & "-" & Y & "，大小为:" & sizeW & "-" & sizeH
        showlogbox log, title_game & " 窗口已激活，且移动到指定位置:" & X & "-" & Y & "，大小为:" & sizeW & "-" & sizeH
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

Public Function pc_restart()
    If pc_tag >= 3 Then
        run_cmd "shutdown -r -f -t 3"
    Else
        pc_tag = pc_tag + 1
    End If
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
    showlog log, "sql_getplayuser: GetDataFromURL = " & GetDataFromURL
    strarr = Split(GetDataFromURL, ",", 2)
    If st = 122 Then
        showlog log, "sql_getplayuser:" & "战网账号安全解锁失败,122" & " 用户名: " & strarr(0) & "   ; 密码: " & strarr(1)
    Else
        showlog log, "sql_getplayuser:" & st & " 用户名: " & strarr(0) & "   ; 密码: " & strarr(1)
    End If
    
    sql_getplayuser = strarr
End Function

Public Function sql_getgolduser(act, st, glod, log)
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    'objHTTP.Open "POST", "http://172.21.12.59/hb/test.php", False
    objHTTP.Open "POST", "http://172.21.12.59/hb/gold.php", False
    objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    result = objHTTP.Send("act=" & act & "&st=" & st & "&glod=" & glod & "&localip=" & getIp)
    GetDataFromURL = objHTTP.ResponseText
    strarr = Split(GetDataFromURL, ",", 2)
    
    showlog log, "金币账户 用户名: " & strarr(0) & "   ； 密码: " & strarr(1)
    sql_getgolduser = strarr
End Function

Public Function sql_setgolduser(u_name, st, glod, log)
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    'objHTTP.Open "POST", "http://172.21.12.59/hb/test.php", False
    objHTTP.Open "POST", "http://172.21.12.59/hb/gold.php", False
    objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    result = objHTTP.Send("u_name=" & u_name & "&act=" & "setgold" & "&st=" & st & "&glod=" & glod)
    'GetDataFromURL = objHTTP.ResponseText
    'strarr = Split(GetDataFromURL, ",", 2)
    
    'showlog log, "金币账户 用户名: " & strarr(0) & "   ； 密码: " & strarr(1)
    'sql_getgolduser = strarr
End Function

Public Function killall(log)
    showlog log, "关闭所有窗口： " & Now
    win = dm.FindWindow("", title_login): Sleep 200: dm.SetWindowState win, 0   '关闭 战网登录
    win = dm.FindWindow("", title_protocol): Sleep 200: dm.SetWindowState win, 0  '关闭 战网协议
    win = dm.FindWindow("", title_nickname): Sleep 200: dm.SetWindowState win, 0  '关闭 战网昵称创建
    win = dm.FindWindow("", title_cannotuse): Sleep 200: dm.SetWindowState win, 0  '关闭 战网无法使用
    win = dm.FindWindow("", title_Battle): Sleep 200: dm.SetWindowState win, 0  '关闭 战网
    win = dm.FindWindow("", title_game): Sleep 200: dm.SetWindowState win, 13  '关闭 炉石传说
    win = dm.FindWindow("", title_game): Sleep 200: dm.SetWindowState win, 13  '关闭 炉石传说
    win = dm.FindWindow("", "战网账号安全"): Sleep 200: dm.SetWindowState win, 13  '关闭 战网安全窗口
    oopwin = dm.FindWindow("", "Oops!"): Sleep 200: dm.SetWindowState oopwin, 0
    Sleep 200: dm.SetWindowState oopwin, 13
    'win = dm.FindWindow("", title_hkj) : Delay ss * 10 : dm.SetWindowState win, 0'关闭 HKJ
    run_cmd "taskkill /f /im dirjfjp.exe"
    win = dm.FindWindow("", title_hbcnf): Sleep 200: dm.SetWindowState win, 13  '关闭 hbcnf
    win = dm.FindWindow("", title_hbupdate): Sleep 200: dm.SetWindowState win, 0  '关闭 hbupdate
    win = dm.FindWindow("", title_hb): Sleep 200: dm.SetWindowState win, 13  '关闭 hb
    win = dm.FindWindow("", "Hearthbuddy"): Sleep 200: dm.SetWindowState win, 13  '关闭 Hearthbuddy exception
    
    run_cmd "taskkill /f /im dirjfjp.exe": Sleep 200
    run_cmd "taskkill /f /im dirjfjp.exe": Sleep 200
    run_cmd "taskkill /f /im Hearthstone.exe": Sleep 200
    run_cmd "taskkill /f /im Hearthstone.exe": Sleep 200

    
    'Call CreateObject("WScript.Shell").run("D:\wg\flush.cmd") //刷新任务栏
    delay 2
    If dm.FindWindow("", title_hb) > 0 Then ' killall 失败，重启pc
        delay 5
        If dm.FindWindow("", title_hb) > 0 Then
            delay 5
            showlog log, "killall 失败，重启pc ： " & Now
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

Function getIp()
    Set objWMIService = GetObject("winmgmts:")
    Set colNicConfig = objWMIService.ExecQuery("SELECT * FROM " & _
    "Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    For Each objNicConfig In colNicConfig
        If Not IsNull(objNicConfig.IPAddress) Then
            For Each strIPAddress In objNicConfig.IPAddress
                getIp = strIPAddress
                Exit Function
            Next
        End If
    Next
End Function

Sub get_screen(log, uname_s, step)
    pic_name = Split(uname_s, "@")
    pic22 = Mid(pic_name(0), 1, 2)
    pic_dir = "D:\wg\pic\" & pic22
    Dim WshShell
    Set WshShell = CreateObject("wscript.Shell")
    WshShell.Run "cmd.exe /c md " + pic_dir, 0
    pic_full = pic_dir & "\" & pic_name(0) & "_" & step & ".png"
    showlog log, "png: " & pic_name(0) & "-" & pic_dir & "-" & pic_full
    delay 2
    dm.CapturePng 0, 0, 1024, 768, pic_full
End Sub
