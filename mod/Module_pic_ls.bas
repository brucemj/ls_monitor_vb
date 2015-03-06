Attribute VB_Name = "pic_ls"
Public pic_arr, point_num_arr

Public Sub pic_lsinit()
    pic_arr = Array("自定义找图", "正常", "战斗异常", "任务", "地精", "关闭")
End Sub

Public Function btn_findpic(log, Index, textbox)
    ' pic_arr = Array("自定义找图", "正常", "战斗异常", "任务", "地精", "关闭")
     Select Case Index
        Case 0 '"自定义找图"
            'showlogbox log, "area:参数错误，请重新输入"
            If textbox(0).text <> "" And textbox(0).text <> "0,0,0,0" Then
                p1 = pic_pointNum(textbox(0).text, textbox(1).text, 0.9)
            End If
            If textbox(2).text <> "" And textbox(2).text <> "0,0,0,0" Then
                p2 = pic_pointNum(textbox(2).text, textbox(3).text, 0.9)
            End If
            If textbox(4).text <> "" And textbox(4).text <> "0,0,0,0" Then
                p3 = pic_pointNum(textbox(4).text, textbox(5).text, 0.9)
            End If
            showlogbox log, "自定义找到的点： " & p1 & ", " & p2 & ", " & p3
        Case 1 '"正常"
            pic_ls_zc1 log
        Case 2 '"战斗异常"
            pic_ls_zd1 log
        Case 3 '"任务"
            pic_ls_qt1 log
        Case 4 '"地精"
            pic_ls_dj1 log
        Case 5 '"关闭"
            pic_ls_gb1 log
    End Select
End Function

Public Function pic_createCode(args, textbox, log, formcode)
    fun_name = args(0).text
    pic_name = args(1).text
    
    Dim fun_code As String
    fun_code = vbCrLf
    fun_code = fun_code & "Public Function " & fun_name & "(log)" & " " & Chr(39) & "自动生成的代码"
    fun_code = fun_code & vbCrLf & "    " & "pic_name=" & Chr(34) & pic_name & Chr(34)
    fun_code = fun_code & vbCrLf & "    " & "p1=" & "0"
    fun_code = fun_code & vbCrLf & "    " & "p2=" & "0"
    fun_code = fun_code & vbCrLf & "    " & "p3=" & "0"
    fun_code = fun_code & vbCrLf & "    " & "isfind=" & "0"
    
    If textbox(0).text <> "" And textbox(0).text <> "0,0,0,0" Then
        fun_code = fun_code & vbCrLf & "    " & "p1 = pic_pointNum(" & textbox(0).text & "," & Chr(34) & textbox(1).text & Chr(34) & "," & textbox(6).text & ")"
    End If
    If textbox(2).text <> "" And textbox(2).text <> "0,0,0,0" Then
        fun_code = fun_code & vbCrLf & "    " & "p2 = pic_pointNum(" & textbox(2).text & "," & Chr(34) & textbox(3).text & Chr(34) & "," & textbox(7).text & ")"
    End If
    If textbox(4).text <> "" And textbox(4).text <> "0,0,0,0" Then
        fun_code = fun_code & vbCrLf & "    " & "p3 = pic_pointNum(" & textbox(4).text & "," & Chr(34) & textbox(5).text & Chr(34) & "," & textbox(8).text & ")"
    End If
    
    fun_code = fun_code & vbCrLf & "    " & "showlogbox log, " & "pic_name " & Chr(38) & " " & Chr(34) & " 找到的点： " & Chr(34) & " " & Chr(38) & " p1 " & Chr(38) & " " & Chr(34) & "," & Chr(34) & " " & Chr(38) & " p2 " & Chr(38) & " " & Chr(34) & "," & Chr(34) & " " & Chr(38) & " p3 "
    
    fun_code = fun_code & vbCrLf & "    " & "If " & "p1 > 0" & " And " & "p2 > 0" & " And " & "p3 > 0"
    fun_code = fun_code & vbCrLf & "        " & "isfind = 1"
    fun_code = fun_code & vbCrLf & "    " & "End If "
    fun_code = fun_code & vbCrLf & "    " & fun_name & " = " & "isfind"
    
    fun_code = fun_code & vbCrLf & "End Function" & vbCrLf
    
    'showlogbox log, fun_code
    formcode.code_txt.text = fun_code
    formcode.Show
End Function

Public Function pic_pointNum(area, colors, sim)
    area_arr = Split(area, ",")
    area_size = UBound(area_arr) - LBound(area_arr) + 1
    If area_size > 3 Then
        point_num = dm.GetResultCount(dm.FindColorEx(area_arr(0), area_arr(1), area_arr(2), area_arr(3), _
                colors, sim, 0))
        pic_pointNum = point_num
    Else
        pic_pointNum = -1
    End If
End Function

Public Function pic_ls_zc1(log) 'Array("自定义找图", "正常", "战斗异常", "任务", "地精", "关闭")
    pic_name = pic_arr(1)
    p1 = 0: p2 = 0: p3 = 0: p4 = 0: isfind = 0
    
    p1 = pic_pointNum("366, 143, 374, 150", "9c4d10-000000|944910-000000|a55118-000000", 0.9)
    p2 = pic_pointNum("58, 204, 65, 213", "081021-000000", 0.9)
    
    showlogbox log, " --- 画面: [" & pic_name & "] 找到的点:  " & p1 & "," & p2 & "," & p3 & "," & p4
    If p1 = -1 Or p2 = -1 Or p3 = -1 Or p4 = -1 Then
        showlogbox log, "area:参数错误，请重新输入"
        isfind = -1
    ElseIf p1 > 150 Then
        showlogbox log, " --- 画面: [" & pic_name & "] 满足条件,已找到:)"
        isfind = 1
    End If
    pic_ls_zc1 = isfind
End Function

Public Function pic_ls_zd1(log) 'Array("自定义找图", "正常", "战斗异常", "任务", "地精", "关闭")
    pic_name = pic_arr(2)
    p1 = 0: p2 = 0: p3 = 0: p4 = 0: isfind = 0
    
    p1 = pic_pointNum("337, 216, 351, 227", "393431-000000", 0.9)
    p2 = pic_pointNum("152, 207, 159, 212", "efcb4a-000000", 0.9)
    
    showlogbox log, " --- 画面: [" & pic_name & "] 找到的点:  " & p1 & "," & p2 & "," & p3 & "," & p4
    If p1 = -1 Or p2 = -1 Or p3 = -1 Or p4 = -1 Then
        showlogbox log, "area:参数错误，请重新输入"
        isfind = -1
    ElseIf p1 > 120 And p2 > 20 Then
        showlogbox log, " --- 画面: [" & pic_name & "] 满足条件,已找到:)"
        isfind = 1
    End If
    pic_ls_zd1 = isfind
End Function

Public Function pic_ls_qt1(log) 'Array("自定义找图", "正常", "战斗异常", "任务", "地精", "关闭")
    pic_name = pic_arr(3)
    p1 = 0: p2 = 0: p3 = 0: p4 = 0: isfind = 0
    
    p1 = pic_pointNum("209, 102, 221, 109", "c69652-000000|d6ae63-000000", 0.9)
    p2 = pic_pointNum("15, 146, 20, 155", "080818-000000", 0.9)
    
    showlogbox log, " --- 画面: [" & pic_name & "] 找到的点:  " & p1 & "," & p2 & "," & p3 & "," & p4
    If p1 = -1 Or p2 = -1 Or p3 = -1 Or p4 = -1 Then
        showlogbox log, "area:参数错误，请重新输入"
        isfind = -1
    ElseIf p1 > 110 And p2 > 30 Then
        showlogbox log, " --- 画面: [" & pic_name & "] 满足条件,已找到:)"
        isfind = 1
    End If
    pic_ls_qt1 = isfind
End Function

Public Function pic_ls_dj1(log) 'Array("自定义找图", "正常", "战斗异常", "任务", "地精", "关闭")
    pic_name = pic_arr(4)
    p1 = 0: p2 = 0: p3 = 0: p4 = 0: isfind = 0
    
    p1 = pic_pointNum("295, 97, 301, 105", "de7184-000000|de718c-000000|de718c-000000", 0.9)
    p2 = pic_pointNum("125, 233, 131, 244", "ffd794-000000|f7cf8c-000000|ffd38c-000000", 0.9)
    
    showlogbox log, " --- 画面: [" & pic_name & "] 找到的点:  " & p1 & "," & p2 & "," & p3 & "," & p4
    If p1 = -1 Or p2 = -1 Or p3 = -1 Or p4 = -1 Then
        showlogbox log, "area:参数错误，请重新输入"
        isfind = -1
    ElseIf p1 > 50 And p2 > 40 Then
        showlogbox log, " --- 画面: [" & pic_name & "] 满足条件,已找到:)"
        isfind = 1
    End If
    pic_ls_dj1 = isfind
End Function

Public Function pic_ls_gb1(log) 'Array("自定义找图", "正常", "战斗异常", "任务", "地精", "关闭")
    pic_name = pic_arr(5)
    p1 = 0: p2 = 0: p3 = 0: p4 = 0: isfind = 0

    p1 = pic_pointNum("133, 200, 140, 208", "5a3c31-000000", 0.9)
    p2 = pic_pointNum("397, 358, 409, 370", "c65918-000000|bd5918-000000|bd5910-000000", 0.9)
    
    showlogbox log, " --- 画面: [" & pic_name & "] 找到的点:  " & p1 & "," & p2 & "," & p3 & "," & p4
    If p1 = -1 Or p2 = -1 Or p3 = -1 Or p4 = -1 Then
        showlogbox log, "area:参数错误，请重新输入"
        isfind = -1
    ElseIf p1 > 40 And p2 > 60 Then
        showlogbox log, " --- 画面: [" & pic_name & "] 满足条件,已找到:)"
        isfind = 1
    End If
    pic_ls_gb1 = isfind
End Function


