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
    If ls_game.login_arg(5).Value = 1 Then ' �Ƿ� �ҽ���û�
        showlog log, "��ʼ ��¼����û��� ------------------------"
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
        showlog log, "��ʼ������Ϸ�� " & Now & " ------------------------"
        killall log
        login_1start
        login_2input uname, passwd, log
        login_3winds log
        login_4battle log
    End If
    
    If ls_game.login_arg(2).Value = 1 Then ' �Ƿ��� ls�Ļ���
        showlog log, "��ʼ ls_pic�� ------------------------"
    End If
    
    If ls_game.login_arg(1).Value = 1 Then
        showlog log, "��ʼ hb ��  ------------------------"
        login_5hb log
    End If
    login_0tag hbrun, lsrun
    showlog log, "TAG: hbrun=" & hbrun & " , lsrun=" & lsrun
    showlog log, "��Ϸ��¼��ɣ� " & Now & " ------------------------"
    delay 3
    If ls_game.login_arg(3).Value = 1 Then
        run_cmd "taskkill /f /im ls_login.exe "
        End
    End If
End Function
Function login_1start() ' Battle.net.config ͳһ�����ļ�
    battle = dm.FindWindow("", title_Battle) '�ر�ս������
    If battle <> 0 Then
        state = dm.SetWindowState(battle, 13): delay 2
    End If
    login_win = dm.FindWindow("", title_login)
    If login_win = 0 Then  ' �� ս����¼
        loginr = CreateObject("WScript.Shell").Run(zw_loginexe)
        delay 2
    End If
    Do
        delay 1
        login_win = dm.FindWindow("", title_login)
    Loop Until login_win
End Function
Function login_2input(u, p, log)  ' �����ʺ�
    login_win = dm.FindWindow("", title_login)
    If login_win <> 0 Then
        dm.SetWindowState login_win, 1
        dm.MoveWindow login_win, g_hswindow_x, g_hswindow_y: delay 1  '//�ƶ� ս����¼ ���ڵ� 0��0
        pic_maintain log
        dm.SetWindowState login_win, 1
        dm.Moveto 100, 155: delay 1: hb_LeftClick3: Sleep 500: dm.KeyPress 46: Sleep 500 '// �ƶ����û��������3��
        dm.SetClipboard u: delay 1   '// �û������õ�ϵͳ���а�
        dm.RightClick: delay 2
        '//dm.Moveto 139,260 : hb_LeftClick3 : Delay ss * 15
        '//dm.Moveto 153,227 : dm.LeftClick : Delay ss * 15
        dm.KeyPress 80: delay 1  '// �Ҽ�����P,����
        showlog log, "�Ҽ�P�����û���": dm.SetWindowState login_win, 1
        dm.KeyPress 9: delay 2  '// tab���ƶ������룬��ȡ���뽹��
        dm.Moveto 273, 210: hb_LeftClick3: delay 1    '// �ƶ����û��������3�Σ���ȡ�û�������
        dm.SetWindowState login_win, 1
        input_str p
        '//      dm.KeyPress 9 :  Delay ss * 10 : dm.KeyPress 9 : Delay ss*10 // tab���ƶ�����¼����ȡ��¼����
        '//      dm.KeyPress 13 : Delay ss * 30 // ��enter
        'inputtime = BeginThread(count_to(10, "��¼ ������ʱ"))
        Do
            delay 1: dm.SetWindowState login_win, 1
            Point = dm.GetResultCount(dm.FindColorEx(134, 281, 145, 292, _
             "0888c3-000000|0b90cd-000000", 0.9, 0))
            'CmdPrint "��¼����, �ҵ��ĵ���: " & Point
        Loop Until Point > 150
        'CmdPrint "��¼ ������ʱ10��ر� ": StopThread inputtime
        dm.Moveto 103, 285: hb_LeftClick3: delay 1   '// �����¼
        showlog log, "�û��ʻ��������"
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
    win = dm.FindWindow("", "ս����¼") '��ȡ��¼����
    If win = 0 Then
        pic_maintain = 88
        Exit Function
    End If
    dm.SetWindowState win, 1: delay 1  '�����¼����
    dm.MoveWindow win, g_hswindow_x, g_hswindow_y: delay 1  '��Ϸ��¼����
    Point = dm.GetResultCount(dm.FindColorEx(31, 34, 43, 43, _
         "ffd700-000000|f7cb00-000000", 0.9, 0))
    If Point > 60 Then
        showlog log, " --- ����: ά������ ���ҵ��ĵ�:  " & Point & "������¼"
        dm.MoveWindow win, g_hswindow_x - 362, g_hswindow_y '��Ϸ��¼����
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
            showlog log, title_login & "��ʱ 60 �룬��¼����ر�" & Now
            End
        End If
    Loop While login_win '�ȴ� ս����¼ ����һ������
    login_3wsec log
    delay 2: protocol = dm.FindWindow("", title_protocol)  'ս��Э�� ����
    If protocol <> 0 Then ' ���� ս��Э�� ����
        showlog log, title_protocol & " ����"
        delay 1: state = dm.SetWindowState(protocol, 1): delay 1
        dm.MoveWindow protocol, g_hswindow_x, g_hswindow_y: delay 1 ' �ƶ� ս��Э�� ���ڵ� 0��0
        dm.Moveto 100, 415: hb_LeftClick3: delay 1    ' ������� ս��Э��
    End If
    Do
        delay 1: protocol = dm.FindWindow("", title_protocol)
    Loop While protocol '�ȴ� ս��Э�� ����һ������
    delay 2: nickname = dm.FindWindow("", title_nickname)   'ս���ǳƴ��� ����
    If nickname <> 0 Then ' ���� ս���ǳƴ��� ����
        showlog log, title_nickname & " ����"
        Do
            delay 1: state = dm.SetWindowState(nickname, 1): delay 1
        Loop Until state  ' �ȴ����� ս���ǳƴ��� ����
        dm.MoveWindow nickname, g_hswindow_x, g_hswindow_y: delay 6 ' �ƶ� ս���ǳƴ��� ���ڵ� 0��0
        '//nicktime = BeginThread(count_to(30, "ս���ǳ� ������ʱ"))
        Do
            delay 1: dm.SetWindowState nickname, 1
            dm.Moveto 77, 252: dm.LeftClick: delay 2   ' ��� �������
            Point = dm.GetResultCount(dm.FindColorEx(56, 410, 72, 423, _
                 "0889c5-000000|098ac7-000000", 0.9, 0))
            showlog log, "ս���ǳ� ����, �ҵ��ĵ���: " & Point
        Loop Until Point > 150
        '//TracePrint "ս���ǳ� ������ʱ30��ر� " : isout = 0
        dm.Moveto 119, 423: hb_LeftClick3: delay 1    ' ���
    End If
    Do
        delay 1: nickname = dm.FindWindow("", title_nickname)
    Loop While nickname '�ȴ� ս���ǳƴ��� ����һ������
    delay 2: cannotuse = dm.FindWindow("", title_cannotuse)   'ս���޷�ʹ�� ����
    If cannotuse <> 0 Then ' ���� ս���޷�ʹ�� ����
        showlog log, title_cannotuse & " ����"
        Do
            delay 1: state = dm.SetWindowState(cannotuse, 1): delay 1
        Loop Until state  ' �ȴ����� ս���޷�ʹ�� ����
        dm.MoveWindow cannotuse, g_hswindow_x, g_hswindow_y: delay 1 ' �ƶ� ս���޷�ʹ�� ���ڵ� 0��0
        dm.Moveto 100, 425: dm.LeftClick: delay 2   ' ��� ��������
    End If
    Do
        delay 1: cannotuse = dm.FindWindow("", title_cannotuse)
    Loop While cannotuse '�ȴ� ս���޷�ʹ�� ���ڱ����
End Function
Function login_3wsec(log) '//ս����ȫ����ش� ����
    delay 1: winsec = dm.FindWindow("", "ս���˺Ű�ȫ")
    If winsec <> 0 Then
        showlog log, "ս���˺Ű�ȫ����"
        
        dm.SetWindowState winsec, 1
        dm.MoveWindow winsec, g_hswindow_x, g_hswindow_y: delay 1  '//�ƶ� ս���˺Ű�ȫ ���ڵ� 0��0
        dm.Moveto 81, 285: hb_LeftClick3: delay 1   '// ���3�Σ���ȡ��ȫ���⽹��
        dm.SetWindowState winsec, 1: input_str "thisiskk"
        get_screen log, uname, "0"
        delay 3
        dm.Moveto 147, 332: hb_LeftClick3: delay 2   '// ���3�Σ���ȡ��ȫ���⽹��
        get_screen log, uname, "1"
        i = 0
        Do
            i = i + 1
            delay 2: winsec2 = dm.FindWindow("", "ս���˺Ű�ȫ")
            If i > 8 Then
                get_screen log, uname, "2"
                showlog log, "ս���˺Ű�ȫ���ڳ�ʱ�������ʺű�����st����Ϊ 122 "
                If is_gold = 0 Then
                    sql_getplayuser "play", "85", "0", log 'ս���˺Ű�ȫ����ʧ��,122
                End If
                showlog log, "ս���˺Ű�ȫ����ʧ�ܣ� ����pc " & Now
                delay 5
                pc_restart
            End If
        Loop While winsec2 '�ȴ� ս����¼ ����һ������
    End If
End Function
Function login_4battle(log)  'ս�� ���ڴ���
    showlog log, title_Battle
    Do
        delay 1: battle = dm.FindWindow("", title_Battle)
    Loop Until battle '�ȴ� ս�� ���ڳ���
    delay 1: battle = dm.FindWindow("", title_Battle)  'ս�� ����
    If battle <> 0 Then ' ���� ս�� ����
        showlog log, title_Battle & " ���֣����е�¼���� ----"
        dm.SetWindowState battle, 1: delay 1
        dm.SetWindowSize battle, 800, 600: delay 1 ' ս�� ���ڳߴ絽 800,600
        dm.MoveWindow battle, g_hswindow_x, g_hswindow_y: delay 1 ' �ƶ� ս�� ���ڵ� 0��0
        Do
            'position=findonlinebtn
            'dm.Moveto position(0), position(1) : dm.LeftClick :  Delay ss*20 : dm.LeftClick :  Delay ss * 30' �������
            dm.Moveto 644, 58: dm.LeftClick: delay 1: dm.LeftClick: delay 2 ' �������
            login_win = dm.FindWindow("", title_login) ' ������� ս����¼
            If login_win > 0 Then
                showlog log, title_login & " ���������ڵ�����ߺ������˺����� ----"
                login_2input uname, passwd, log ' �����˺�����
            ElseIf login_win = 0 Then
                dm.SetWindowState battle, 1
                dm.Moveto 60, 330: hb_LeftClick3: delay 2  ' ��ս�����ڣ���� ¯ʯ��˵
                dm.Moveto 290, 530: hb_LeftClick3: delay 2   ' ��ս�����ڣ���� ������Ϸ
            End If
            lscs = dm.FindWindow("", title_game): delay 2
        Loop Until lscs
        dm.SetWindowState lscs, 1: delay 1: dm.SetWindowSize lscs, 512, 384: delay 2
        dm.MoveWindow lscs, g_hswindow_x, g_hswindow_y: delay 2  '��Ϸ���ڵ���
        showlog log, title_game & " �Ѿ���������¼��� ................."
    End If
    login_4battle = dm.FindWindow("", title_game)
    lsrun = 1
End Function

Function login_5hb(log)  'hb ���ڴ���
    lscs = dm.FindWindow("", title_game)
    If lscs = 0 Then
        showlog log, title_game & " ���ڲ�����"
        Exit Function
    End If
    dm.MoveWindow lscs, 970, 680: delay 1  '// ��Ϸ���ڵ���
    dm.SetWindowState game_hwnd, 1: delay 4  '//��С����Ϸ
    
        showlog log, "cdk ����û���������"
        Do
            run_cmd "taskkill /f /im dirjfjp.exe"
            delay 2
            run_cmd "taskkill /f /im dirjfjp.exe"
            delay 2
            cdkr = CreateObject("WScript.Shell").Run(hb_hkj)
            delay 6
            cdk = dm.FindWindow("", "Using VIP auth server")
        Loop Until cdk
        
        showlog log, " cdk ���ڣ���: " & cdk
        dm.MoveWindow cdk, g_hswindow_x, g_hswindow_y: delay 2  ' cdk���ڵ���
        dm.SetWindowState cdk, 1: dm.Moveto 109, 61: dm.LeftClick: delay 1: dm.LeftClick       ' cdk���ڵ��
        dm.MoveWindow cdk, 1010, -50: delay 1    ' cdk���ڵ���

    
    
    showlog log, "׼������hb -----"
    hb_win = dm.FindWindow("", title_hb)
    If hb_win = 0 Then  '//���hb���ڲ����ڣ���� hb
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

    dm.SetWindowState hb_win, 1: delay 1: dm.SetWindowSize hb_win, 415, 480: delay 2   '���� hb ����
    dm.MoveWindow hb_win, g_hswindow_x, g_hswindow_y: delay 1 '��Ϸ���ڵ���
    hbrun = 1
    login_5hb = 1
    showlog log, "hb ���ڿ��� -----"
End Function


Sub login_5hb_set_profile(name, hWnd)
    dm.MoveWindow hWnd, g_hswindow_x, g_hswindow_y: delay 2 '��Ϸ���ڵ���
    
    profile = Split(name, "@")
    dm.SetClipboard profile(0): delay 2  ' �û������õ�ϵͳ���а�
    dm.Moveto 22, 44: delay 1
    dm.LeftClick: delay 2
    dm.RightClick: delay 2
    'dm.Moveto 74,87 : hb_LeftClick3 : Delay ss * 5
    dm.KeyPress 80: delay 2  ' �Ҽ�����P,����
    dm.SetWindowState hWnd, 1: delay 2: dm.KeyPress 13   '���� hbcnf ����,���س�
End Sub




