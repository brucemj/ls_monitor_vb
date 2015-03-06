Attribute VB_Name = "monitor"
'Public Const zw_loginexe = "D:\game\st_battle.cmd"
'userinfo = hb_Fgetplayuser("play", "81", "0")
Public monitor_file, user_tag, ls_tag, hb_tag, unixtime_tag, tag_tick_time, hb_tick, hb_new
Public login_exe As String, st_keys, st_keys_length, debug_check As Boolean, monitor_running As Boolean

Public Sub a_debug_test(log)
    st_file = "D:\wg\hb_new\Hearthbuddy 0.3.828.115\Settings\" & user_tag & "\Stats.json"
    If dm.IsFileExist(st_file) = 0 Then
        show_monitorlog log, "    �ļ� " & st_file & " ������!!!", False
    End If
    oopwin = dm.FindWindow("", "Oops!")
    gamewin = dm.FindWindow("", "¯ʯ��˵")
    show_monitorlog log, "---- debug_test info: " & oopwin & "," & gamewin, False
End Sub

Public Sub monitor_init(log)
    monitor_file = logdir & "\" & "game_monitor.txt"
    monitor_running = False
    hb_tick = 0
    hb_new = 0
    login_exe = "D:\wg\ls_login.exe"
    user_tag = "default"
    ls_tag = 0
    hb_tag = 0
    unixtime_tag = 0
    tag_tick_time = 0
    monitor_readTag log
    debug_check = False
    st_keys = Array("Wins", "Losses", "Concedes", "Quests", "Newtime", "Ticktime", "DWins", "DLosses")
    st_keys_length = UBound(st_keys) - LBound(st_keys) + 1
End Sub
Public Sub monitor_start(log, looptime As Integer, ls_monitor)
    Dim m_delay_seonde As Integer
    monitor_running = True
    ls_monitor.monitor_st.text = "monitor is running vvvvvv"
    Dim T() As String
    Do
        m_delay_seonde = looptime
    T() = Split(log.text, vbCrLf)
    If UBound(T) > 100 Then
        log.text = "monitor info:"
    End If
    
    debug_check = False
    If ls_monitor.monitor_arg(0).Value = 1 Then ' windows
        debug_check = True
        If monitor_windows(log) = 1 Then ' 0 Ϊ����������1Ϊ�����쳣
            show_monitorlog log, "    monitor_windows �����쳣 !!! [windows����ģʽ,����������¼] ", Not debug_check
            game_restart
        Else
            show_monitorlog log, "    monitor_windows �������� vvv [windows����ģʽ,����������¼] ", Not debug_check
        End If
    End If
    
    If ls_monitor.monitor_arg(1).Value = 1 Then ' quests
        debug_check = True
        If monitor_quests(log) = 0 Then ' 0 is questing , 1 is need change user
            show_monitorlog log, "    monitor_quests ����û��� !!! [quests����ģʽ,����������¼] ", Not debug_check
        Else
            show_monitorlog log, "    monitor_quests ��������� vvv [quests����ģʽ,����������¼] ", Not debug_check
            change_user log
        End If
    End If
    
    If debug_check Then
        Exit Do
    End If
    ' ---------------------- monitor_start run -------------------
    show_monitorlog log, "monitor_start run, looptime=" & looptime & " --- " & Now & "-----------------", Not debug_check

    If CheckExeIsRun("ls_login.exe") = 0 Then 'û�е�¼���� ls_login.exe
        If monitor_readTag(log) Then 'û�е�¼����,tag_file���ڣ���� ls+hb �Ƿ���������
            show_monitorlog log, "  û�е�¼����,tag_file����,��� ls+hb �Ƿ���������" & Now, Not debug_check
            'If tag_tick_time > 240 Then ' ��¼�����Ѿ������� 240 ��
                If monitor_windows(log) = 1 Then ' 0 Ϊ����������1Ϊ�����쳣
                    show_monitorlog log, "    monitor_windows �����쳣 !!! ,����ls_login.exe; " & "tag_tick_time=" & tag_tick_time, Not debug_check
                    killall log
                    game_restart
                Else
                    show_monitorlog log, "    monitor_windows ��������" & "tag_tick_time=" & tag_tick_time & "; " & Now, Not debug_check ' һ�д�������
                    If monitor_quests(log) = 0 Then ' 0 is questing , 1 is need change user
                        show_monitorlog log, "    monitor_quests ����û��� !!!  ", Not debug_check
                        If monitor_hbtime(log) = 0 Then ' 0 is ok , 1 is timeout
                            show_monitorlog log, "    monitor_hbtime ����   ", Not debug_check
                        Else
                            show_monitorlog log, "    monitor_hbtime �쳣 !!!,����ls_login.exe; ", Not debug_check
                            killall log
                            game_restart
                        End If
                    Else
                        show_monitorlog log, "    monitor_quests ��������� vvv  ", Not debug_check
                        change_user log
                    End If
                    
                End If
            'End If
        Else 'û�е�¼���̣����� tag_file ������
            show_monitorlog log, "  û�е�¼���̣����� tag_file �����ڣ������� ls_login.exe " & Now, Not debug_check
            killall log
            game_restart
        End If
    Else '���ڵ�¼���� ls_login.exe
        show_monitorlog log, "  ���ڵ�¼���� ls_login.exe ------ " & Now, Not debug_check
        If monitor_readTag(log) Then '���ڵ�¼���̣�tag_file����
            If tag_tick_time > 240 Then ' ��¼�����Ѿ������� 240 ��,��ʱ��
                show_monitorlog log, "    ��¼�����Ѿ������� 240 ��,��ʱ��,���� ls_login.exe" & Now, Not debug_check
                killall log
                game_restart
            Else
                Dim w_t As Integer, i As Integer
                w_t = 240 - tag_tick_time
                show_monitorlog log, "    ���ڵ�¼����,tag_tick_time=" & tag_tick_time & "�ȴ�������" & w_t & "; " & Now, Not debug_check
                i = 0
                Do
                    i = i + 20
                    delay 20
                    is_login = CheckExeIsRun("ls_login.exe") ' 1 is running
                Loop While is_login And i < w_t
                m_delay_seonde = 10
            End If '
        Else ' ���ڵ�¼���̣���tag_file�ļ�������
            show_monitorlog log, "    ���ڵ�¼���� ls_login.exe����tag_file�ļ������ڣ�����ԭ��" & Now, Not debug_check
            killall log
            game_restart
        End If
    End If
    
    delay m_delay_seonde
    Loop Until ls_monitor.monitor_arg(2).Value
    show_monitorlog log, "--------   monitor stop -------- " & Now, Not debug_check
    monitor_running = False
    ls_monitor.monitor_st.text = "monitor is stop xxxxxx"
End Sub
Public Function monitor_logining()
    monitor_logining = 0
End Function
Public Function monitor_readTag(log)
    If dm.IsFileExist(login_tagfile) Then
        show_monitorlog log, "  " & login_tagfile & "��ȡ ---- " & Now, Not debug_check
        Dim tag_str As String
        'Dim WshShell
        'Set WshShell = CreateObject("wscript.Shell")
        'Set tag_info = WshShell.Exec("cmd.exe /c type " & login_tagfile)
        'tag_str = tag_info.StdOut.ReadLine()
        tag_str = dm.readfile(login_tagfile)
        tag_arr = Split(tag_str, ";")
        user_tag = tag_arr(0)
        ls_tag = tag_arr(1)
        hb_tag = tag_arr(2)
        unixtime_tag = tag_arr(3)
        tag_tick_time = DateDiff("s", "01/01/1970 08:00:00", Now()) - unixtime_tag
        show_monitorlog log, "  Tag��Ϣ�� " & user_tag & ";" & ls_tag & ";" & hb_tag & ";" & unixtime_tag & " ;tag_tick_time=" & tag_tick_time, Not debug_check
        monitor_readTag = 1
        Exit Function
    Else
        show_monitorlog log, "  " & login_tagfile & " �����ڣ�������¼ " & Now, Not debug_check
        monitor_readTag = 0
        Exit Function
    End If
End Function
Public Function monitor_windows(log) ' 0 is OK , 1 is restart
    oopwin = dm.FindWindow("", "Oops!")
    If oopwin <> 0 Then
        show_monitorlog log, "      ��Ϸ����Oops!�쳣 !!!!!!!!!!!!!!!", Not debug_check
        oopwin = dm.FindWindow("", "Oops!")
        delay 2
        dm.SetWindowState oopwin, 13
        monitor_windows = 1
        Exit Function
    End If
    
    hb_license = dm.FindWindow("#32770", "Error")
    If hb_license <> 0 Then
        show_monitorlog log, "      hb_license �����쳣�����������!!!!!!!!!!!!!!!", Not debug_check
        delay 4
        pc_restart
        monitor_windows = 1
        Exit Function
    End If
    
    If ls_tag = 1 And hb_tag = 1 Then '��¼�����Ѿ������� ls+hb
        show_monitorlog log, "      Tag: ��¼�����Ѿ������� ls=1 , hb=1", Not debug_check
        gamewin = dm.FindWindow("", "¯ʯ��˵")
        If gamewin = 0 Then
            show_monitorlog log, "      ls���ڲ����ڣ���������!!!!!!!!!!!!!!!", Not debug_check
            monitor_windows = 1
            Exit Function
        Else
            If checkfreezewin(gamewin, "¯ʯ��˵", log) = 1 Then
                monitor_windows = 1
                Exit Function
            End If
        End If
        
        hbwin = dm.FindWindow("", title_hb)
        If hbwin = 0 Then
            show_monitorlog log, "      hb���ڲ����ڣ���������!!!!!!!!!!!!!!!", Not debug_check
            monitor_windows = 1
            Exit Function
        Else
            If checkfreezewin(hbwin, "hbuddy", log) = 1 Then
                monitor_windows = 1
                Exit Function
            End If
        End If
    Else
        If CheckExeIsRun("ls_login.exe") = 0 Then
            show_monitorlog log, "    ls_tag=" & ls_tag & ",hb_tag=" & hb_tag & " ;���� ls_login.exe ����û�м�鵽 ~!!!!!~~", Not debug_check
            monitor_windows = 1
            Exit Function
        End If
    End If

    monitor_windows = 0 ' 0 is OK , 1 is restart
End Function

Public Function monitor_quests(log) ' 0 is questing , 1 is need change user
    Dim json, unix_Ntime, st_file_txt As String
    'st_keys = Array("Wins", "Losses", "Concedes", "Quests", "Newtime", "Ticktime", "DWins", "DLosses")
    Dim wins, losses, concedes, dwins, dlosses, quests, newtime, ticktime
    st_file = "D:\wg\hb_new\Hearthbuddy 0.3.828.115\Settings\" & user_tag & "\Stats.json"
    

    If dm.IsFileExist(st_file) = 0 Then
        show_monitorlog log, "    �ļ� " & st_file & " ������!!!", Not debug_check
        Exit Function
    End If
    
    Set json = New VbsJson
    st_file_txt = dm.readfile(st_file)
    
    wins = json.ParseJson(st_file_txt, "Wins")
    losses = json.ParseJson(st_file_txt, "Losses")
    concedes = json.ParseJson(st_file_txt, "Concedes")
    dwins = json.ParseJson(st_file_txt, "DWins")
    dlosses = json.ParseJson(st_file_txt, "DLosses")
    quests = json.ParseJson(st_file_txt, "Quests")
    newtime = json.ParseJson(st_file_txt, "Newtime")
    ticktime = json.ParseJson(st_file_txt, "Ticktime")
        
    'For i = 0 To st_keys_length - 1
    '    k = st_keys(i)
    '    v = json.ParseJson(st_file_txt, k)
    '    show_monitorlog log, "    " & k & "=" & v, Not debug_check
    'Next i
    unix_Ntime = DateDiff("s", "01/01/1970 08:00:00", Now())
    hb_tick = unix_Ntime - ticktime
    hb_new = unix_Ntime - newtime
    show_monitorlog log, "  �ۼƳ���Wins/Losses:" & wins & "/" & losses & ";" & "���ճ���DWins/DLosses:" & _
        dwins & "/" & dlosses & ",q=" & quests & ",hb_t=" & hb_tick & ",hb_n=" & hb_new, Not debug_check
    
    
    games = dwins + dlosses
    If dwins >= 1 And quests = 0 Then
        show_monitorlog log, "    ������������ 1+0�����¿�ʼ vvvvvv", Not debug_check
        monitor_quests = 1
        Exit Function
    ElseIf dwins >= 9 Then
        show_monitorlog log, "    ������������ 9 dwins�����¿�ʼ vvvvvv", Not debug_check
        monitor_quests = 1
        Exit Function
    ElseIf games >= 30 Then
        show_monitorlog log, "    ������������ 30 games�����¿�ʼ vvvvvv", Not debug_check
        monitor_quests = 1
        Exit Function
    End If
    
    monitor_quests = 0
End Function

Public Function monitor_hbtime(log) ' 0 is ok , 1 is timeout
    If Not debug_check Then
        If hb_tick > 200 Or hb_new > 1500 Then
            monitor_hbtime = 1
            Exit Function
        End If
    Else
        show_monitorlog log, "    ����ģʽ ,user_tag=" & user_tag & "; hb_tick,hb_new=" & hb_tick & "," & hb_new, Not debug_check
    End If
    monitor_hbtime = 0
End Function
Public Function change_user(log)
    If Not debug_check Then
        sql_getplayuser "play", "82", "0", log ' ����
        killall log
        game_restart
    Else
        show_monitorlog log, "    ����ģʽ ,user_tag=" & user_tag, Not debug_check
    End If
End Function

Public Function game_stop()
End Function

Public Function game_restart()
    If Not debug_check Then
        delay 3
        run_cmd "start " & login_exe
    End If
End Function

Public Function checkfreezewin2(hwin, title, log) ' 0 is OK , 1 is restart
    'Dim x1, y1, x2, y2
    'dm.SetWindowState hwin, 1
    i = 0
    Do
        If IsHungAppWindow(hwin) = 0 Then
            i = 0
            checkfreezewin2 = 0
            show_monitorlog log, "    " & title & " ��������  ----", Not debug_check
            Exit Do
        Else
            i = i + 1
            If i = 20 Then '�������δ��Ӧ20�Σ�
                checkfreezewin2 = 1
                show_monitorlog log, "    " & title & " ����δ��Ӧ20��,60��  !!!! ", Not debug_check
                Exit Do
            End If
        End If
        delay 3
    Loop
End Function

 Public Function checkfreezewin(hwin, title, log) ' 0 is OK , 1 is restart
    'Dim x1, y1, x2, y2
    'dm.SetWindowState hwin, 1
    i = 0
    Do
        If dm.GetWindowState(hwin, 6) = 0 Then
            i = 0
            checkfreezewin = 0
            show_monitorlog log, "    " & title & " 2��������  ----", Not debug_check
            Exit Do
        Else
            i = i + 1
            If i = 20 Then '�������δ��Ӧ20�Σ�
                checkfreezewin = 1
                show_monitorlog log, "    " & title & " 2����δ��Ӧ20��,60��  !!!! ", Not debug_check
                Exit Do
            End If
        End If
        delay 3
    Loop
End Function
 
 
 
 

