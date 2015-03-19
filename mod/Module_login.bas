Attribute VB_Name = "login"
Public Const zw_loginexe = "D:\game\st_battle.cmd"
Public Const hb_hkj = "D:\wg\key.cmd"
Public Const hb_wg = "D:\wg\hbr.cmd "

Public Const g_resolution_w = 512
Public Const g_resolution_h = 384
Public Const g_hswindow_x = 0
Public Const g_hswindow_y = 0
'userinfo = hb_Fgetplayuser("play", "81", "0")
Public uname, passwd, lsrun, hbrun, logfile, login_tagfile, is_gold
Public Sub login_init()
    logfile = logdir & "\" & "game_work.txt"
    login_tagfile = "D:\wg\log\" & "tag.txt"
    lsrun = 0
    hbrun = 0
    is_gold = 0
End Sub
Public Sub login_0tag(ls, hb)
    UnixTime = DateDiff("s", "01/01/1970 08:00:00", Now())
    array_name = Split(uname, "@")
    'run_cmd "echo " & array_name(0) & ";" & ls & ";" & hb & ";" & UnixTime & "> " & login_tagfile
    run_cmd ">" & login_tagfile & " set /p=" & array_name(0) & ";" & ls & ";" & hb & ";" & UnixTime & "<nul"
End Sub
Function intogame(log, ls_game)
    showlog log, "TAG: hbrun=" & hbrun & " , lsrun=" & lsrun
    If ls_game.login_arg(5).Value = 1 Then ' 是否 找金币用户
        showlog log, "开始 登录金币用户： ------------------------"
        is_gold = 1
        userinfo = sql_getgolduser("gold", 80, 0, log)
    Else
        is_gold = 0
        userinfo = sql_getplayuser("play", 81, 0, log)
    End If
    
    'userinfo = Split("brucemj44@163.com,aaaa2222", ",", 2)
    uname = userinfo(0)
    passwd = userinfo(1)
    'uname = "mjflash111@163.com"
    'passwd = "ffff8888"
    login_0tag lsrun, hbrun
    
    If ls_game.login_arg(0).Value = 1 Then
        showlog log, "开始启动游戏： " & Now & " ------------------------"
        killall log
        login_1start
        login_2input uname, passwd, log
        login_3winds log
        login_4battle log
    End If
    
    If ls_game.login_arg(2).Value = 1 Then ' 是否找 ls的画面
        showlog log, "开始 ls_pic： ------------------------"
    End If
    
    If ls_game.login_arg(1).Value = 1 Then
        showlog log, "开始 hb ：  ------------------------"
        login_5hb log
    End If
    login_0tag hbrun, lsrun
    showlog log, "TAG: hbrun=" & hbrun & " , lsrun=" & lsrun
    showlog log, "游戏登录完成： " & Now & " ------------------------"
    delay 3
    If ls_game.login_arg(3).Value = 1 Then
        run_cmd "taskkill /f /im ls_login.exe "
        End
    End If
End Function
Function login_1start() ' Battle.net.config 统一配置文件
    battle = dm.FindWindow("", title_Battle) '关闭战网界面
    If battle <> 0 Then
        state = dm.SetWindowState(battle, 13): delay 2
    End If
    login_win = dm.FindWindow("", title_login)
    If login_win = 0 Then  ' 打开 战网登录
        loginr = CreateObject("WScript.Shell").Run(zw_loginexe)
        delay 2
    End If
    Do
        delay 1
        login_win = dm.FindWindow("", title_login)
    Loop Until login_win
End Function
Function login_2input(u, p, log)  ' 输入帐号
    login_win = dm.FindWindow("", title_login)
    If login_win <> 0 Then
        dm.SetWindowState login_win, 1
        dm.MoveWindow login_win, g_hswindow_x, g_hswindow_y: delay 1  '//移动 战网登录 窗口到 0，0
        pic_maintain log
        dm.SetWindowState login_win, 1
        dm.Moveto 100, 155: delay 1: hb_LeftClick3: Sleep 500: dm.KeyPress 46: Sleep 500 '// 移动到用户名，点击3次
        dm.SetClipboard u: delay 1   '// 用户名放置到系统剪切板
        dm.RightClick: delay 2
        '//dm.Moveto 139,260 : hb_LeftClick3 : Delay ss * 15
        '//dm.Moveto 153,227 : dm.LeftClick : Delay ss * 15
        dm.KeyPress 80: delay 1  '// 右键，按P,复制
        showlog log, "右键P复制用户名": dm.SetWindowState login_win, 1
        dm.KeyPress 9: delay 2  '// tab键移动到密码，获取密码焦点
        dm.Moveto 273, 210: hb_LeftClick3: delay 1    '// 移动到用户名，点击3次，获取用户名焦点
        dm.SetWindowState login_win, 1
        input_str p
        '//      dm.KeyPress 9 :  Delay ss * 10 : dm.KeyPress 9 : Delay ss*10 // tab键移动到登录，获取登录焦点
        '//      dm.KeyPress 13 : Delay ss * 30 // 按enter
        'inputtime = BeginThread(count_to(10, "登录 按键超时"))
        Do
            delay 1: dm.SetWindowState login_win, 1
            Point = dm.GetResultCount(dm.FindColorEx(134, 281, 145, 292, _
             "0888c3-000000|0b90cd-000000", 0.9, 0))
            'CmdPrint "登录按键, 找到的点数: " & Point
        Loop Until Point > 150
        'CmdPrint "登录 按键超时10秒关闭 ": StopThread inputtime
        dm.Moveto 103, 285: hb_LeftClick3: delay 1   '// 点击登录
        showlog log, "用户帐户输入完成"
    End If
End Function
Sub input_str(str)
    lng = Len(str)
    For n = 1 To lng
        ch = Mid(str, n, 1)
        dm.KeyPressChar ch
        Sleep 150
    Next n
End Sub
Function pic_maintain(log)
    win = dm.FindWindow("", "战网登录") '获取登录窗口
    If win = 0 Then
        pic_maintain = 88
        Exit Function
    End If
    dm.SetWindowState win, 1: delay 1  '激活登录窗口
    dm.MoveWindow win, g_hswindow_x, g_hswindow_y: delay 1  '游戏登录调整
    Point = dm.GetResultCount(dm.FindColorEx(31, 34, 43, 43, _
         "ffd700-000000|f7cb00-000000", 0.9, 0))
    If Point > 60 Then
        showlog log, " --- 画面: 维护提醒 ，找到的点:  " & Point & "继续登录"
        dm.MoveWindow win, g_hswindow_x - 362, g_hswindow_y '游戏登录调整
        pic_maintain = 1
        Exit Function
    End If
    pic_maintain = 0
End Function
Function login_3winds(log)  '
    i = 0
    Do
        i = i + 1
        delay 3
        login_win = dm.FindWindow("", title_login)
        If i > 20 Then
            showlog log, title_login & "超时 60 秒，登录程序关闭" & Now
            End
        End If
    Loop While login_win '等待 战网登录 的下一个窗口
    login_3wsec log
    delay 2: protocol = dm.FindWindow("", title_protocol)  '战网协议 窗口
    If protocol <> 0 Then ' 处理 战网协议 窗口
        showlog log, title_protocol & " 处理"
        delay 1: state = dm.SetWindowState(protocol, 1): delay 1
        dm.MoveWindow protocol, g_hswindow_x, g_hswindow_y: delay 1 ' 移动 战网协议 窗口到 0，0
        dm.Moveto 100, 415: hb_LeftClick3: delay 1    ' 点击接受 战网协议
    End If
    Do
        delay 1: protocol = dm.FindWindow("", title_protocol)
    Loop While protocol '等待 战网协议 的下一个窗口
    delay 2: nickname = dm.FindWindow("", title_nickname)   '战网昵称创建 窗口
    If nickname <> 0 Then ' 处理 战网昵称创建 窗口
        showlog log, title_nickname & " 处理"
        Do
            delay 1: state = dm.SetWindowState(nickname, 1): delay 1
        Loop Until state  ' 等待激活 战网昵称创建 窗口
        dm.MoveWindow nickname, g_hswindow_x, g_hswindow_y: delay 6 ' 移动 战网昵称创建 窗口到 0，0
        '//nicktime = BeginThread(count_to(30, "战网昵称 按键超时"))
        Do
            delay 1: dm.SetWindowState nickname, 1
            dm.Moveto 77, 252: dm.LeftClick: delay 2   ' 点击 随机生成
            Point = dm.GetResultCount(dm.FindColorEx(56, 410, 72, 423, _
                 "0889c5-000000|098ac7-000000", 0.9, 0))
            showlog log, "战网昵称 按键, 找到的点数: " & Point
        Loop Until Point > 150
        '//TracePrint "战网昵称 按键超时30秒关闭 " : isout = 0
        dm.Moveto 119, 423: hb_LeftClick3: delay 1    ' 完成
    End If
    Do
        delay 1: nickname = dm.FindWindow("", title_nickname)
    Loop While nickname '等待 战网昵称创建 的下一个窗口
    delay 2: cannotuse = dm.FindWindow("", title_cannotuse)   '战网无法使用 窗口
    If cannotuse <> 0 Then ' 处理 战网无法使用 窗口
        showlog log, title_cannotuse & " 处理"
        Do
            delay 1: state = dm.SetWindowState(cannotuse, 1): delay 1
        Loop Until state  ' 等待激活 战网无法使用 窗口
        dm.MoveWindow cannotuse, g_hswindow_x, g_hswindow_y: delay 1 ' 移动 战网无法使用 窗口到 0，0
        dm.Moveto 100, 425: dm.LeftClick: delay 2   ' 点击 继续离线
    End If
    Do
        delay 1: cannotuse = dm.FindWindow("", title_cannotuse)
    Loop While cannotuse '等待 战网无法使用 窗口被点击
End Function
Function login_3wsec(log) '//战网安全问题回答 窗口
    delay 1: winsec = dm.FindWindow("", "战网账号安全")
    If winsec <> 0 Then
        showlog log, "战网账号安全处理"
        
        dm.SetWindowState winsec, 1
        dm.MoveWindow winsec, g_hswindow_x, g_hswindow_y: delay 1  '//移动 战网账号安全 窗口到 0，0
        dm.Moveto 81, 285: hb_LeftClick3: delay 1   '// 点击3次，获取安全问题焦点
        dm.SetWindowState winsec, 1: input_str "thisiskk"
        get_screen log, uname, "0"
        delay 3
        dm.Moveto 147, 332: hb_LeftClick3: delay 2   '// 点击3次，获取安全问题焦点
        get_screen log, uname, "1"
        i = 0
        Do
            i = i + 1
            delay 2: winsec2 = dm.FindWindow("", "战网账号安全")
            If i > 8 Then
                get_screen log, uname, "2"
                showlog log, "战网账号安全窗口超时，可能帐号被锁，st设置为 122 "
                If is_gold = 0 Then
                    sql_getplayuser "play", "85", "0", log '战网账号安全解锁失败,122
                End If
                showlog log, "战网账号安全处理失败， 重启pc " & Now
                delay 5
                pc_restart
            End If
        Loop While winsec2 '等待 战网登录 的下一个窗口
    End If
End Function
Function login_4battle(log)  '战网 窗口处理
    showlog log, title_Battle
    Do
        delay 1: battle = dm.FindWindow("", title_Battle)
    Loop Until battle '等待 战网 窗口出现
    delay 1: battle = dm.FindWindow("", title_Battle)  '战网 窗口
    If battle <> 0 Then ' 处理 战网 窗口
        showlog log, title_Battle & " 出现，进行登录处理 ----"
        dm.SetWindowState battle, 1: delay 1
        dm.SetWindowSize battle, 800, 600: delay 1 ' 战网 窗口尺寸到 800,600
        dm.MoveWindow battle, g_hswindow_x, g_hswindow_y: delay 1 ' 移动 战网 窗口到 0，0
        Do
            'position=findonlinebtn
            'dm.Moveto position(0), position(1) : dm.LeftClick :  Delay ss*20 : dm.LeftClick :  Delay ss * 30' 点击上线
            dm.Moveto 644, 58: dm.LeftClick: delay 1: dm.LeftClick: delay 2 ' 点击上线
            login_win = dm.FindWindow("", title_login) ' 如果出现 战网登录
            If login_win > 0 Then
                showlog log, title_login & " 被弹出，在点击上线后。输入账号密码 ----"
                login_2input uname, passwd, log ' 输入账号密码
            ElseIf login_win = 0 Then
                dm.SetWindowState battle, 1
                dm.Moveto 60, 330: hb_LeftClick3: delay 2  ' 在战网窗口，点击 炉石传说
                dm.Moveto 290, 530: hb_LeftClick3: delay 2   ' 在战网窗口，点击 进入游戏
            End If
            lscs = dm.FindWindow("", title_game): delay 2
        Loop Until lscs
        dm.SetWindowState lscs, 1: delay 1: dm.SetWindowSize lscs, 512, 384: delay 2
        dm.MoveWindow lscs, g_hswindow_x, g_hswindow_y: delay 2  '游戏窗口调整
        showlog log, title_game & " 已经启动，登录完成 ................."
    End If
    login_4battle = dm.FindWindow("", title_game)
    lsrun = 1
End Function

Function login_5hb(log)  'hb 窗口处理
    lscs = dm.FindWindow("", title_game)
    If lscs = 0 Then
        showlog log, title_game & " 窗口不存在"
        Exit Function
    End If
    dm.MoveWindow lscs, 970, 680: delay 1  '// 游戏窗口调整
    dm.SetWindowState game_hwnd, 1: delay 4  '//最小化游戏
    
        showlog log, "cdk 窗口没开启，则打开"
        Do
            run_cmd "taskkill /f /im dirjfjp.exe"
            delay 2
            run_cmd "taskkill /f /im dirjfjp.exe"
            delay 2
            cdkr = CreateObject("WScript.Shell").Run(hb_hkj)
            delay 6
            cdk = dm.FindWindow("", "Using VIP auth server")
        Loop Until cdk
        
        showlog log, " cdk 窗口，打开: " & cdk
        dm.MoveWindow cdk, g_hswindow_x, g_hswindow_y: delay 2  ' cdk窗口调整
        dm.SetWindowState cdk, 1: dm.Moveto 109, 61: dm.LeftClick: delay 1: dm.LeftClick       ' cdk窗口点击
        dm.MoveWindow cdk, 1010, -50: delay 1    ' cdk窗口调整

    
    
    showlog log, "准备开启hb -----"
    hb_win = dm.FindWindow("", title_hb)
    If hb_win = 0 Then  '//如果hb窗口不存在，则打开 hb
        username_name = Split(uname, "@")
        hb_wgr = hb_wg & username_name(0)
        hbr = CreateObject("WScript.Shell").Run(hb_wgr)
        Do
            delay 3
            hbcnf = dm.FindWindow("", title_hbcnf)
        Loop Until hbcnf
    
        login_5hb_set_profile uname, hbcnf
    Else
        delay 1
    End If
    
    i = 0
    Do
        i = i + 1
        delay 3
        hb_win = dm.FindWindow("", title_hb)
        If i > 12 Then
            login_5hb = 0
            Exit Function
        End If
    Loop Until hb_win

    dm.SetWindowState hb_win, 1: delay 1: dm.SetWindowSize hb_win, 415, 480: delay 2   '激活 hb 窗口
    dm.MoveWindow hb_win, g_hswindow_x, g_hswindow_y: delay 1 '游戏窗口调整
    hbrun = 1
    login_5hb = 1
    showlog log, "hb 窗口开启 -----"
End Function


Sub login_5hb_set_profile(name, hWnd)
    dm.MoveWindow hWnd, g_hswindow_x, g_hswindow_y: delay 2 '游戏窗口调整
    
    profile = Split(name, "@")
    dm.SetClipboard profile(0): delay 2  ' 用户名放置到系统剪切板
    dm.Moveto 22, 44: delay 1
    dm.LeftClick: delay 2
    dm.RightClick: delay 2
    'dm.Moveto 74,87 : hb_LeftClick3 : Delay ss * 5
    dm.KeyPress 80: delay 2  ' 右键，按P,复制
    dm.SetWindowState hWnd, 1: delay 2: dm.KeyPress 13   '激活 hbcnf 窗口,并回车
End Sub




