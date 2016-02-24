VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XHR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private uc As SAClient

Sub ReadyStateChange()
Attribute ReadyStateChange.VB_UserMemId = 0
    Dim state As Integer
    state = uc.XHRequest.readyState
    If state = 4 Then
        uc.hasResponse uc.XHRequest.responseText
    End If
    If uc.Enabled Then uc.setIndicator state + 1
End Sub

Sub setParent(ByRef obj As SAClient)
    Set uc = obj
End Sub

