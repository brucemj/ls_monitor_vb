VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VbsJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function ParseJson(json_str, key)
    Dim sc
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    sc.AddCode "var o = " & json_str & ";"
    ParseJson = sc.Eval("o." & key)
End Function
