VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Uni As Integer

Private IsDefu As Boolean

Private Pro_Log As Variant

Private oldWinProc As Long

Private Play_Uni As Integer

Private Play_ProLog As Variant

'//cprin added next line to support BackUpAll feature
Private BuckupAll_Type As String

Private Recording_IsOver As Boolean

'//cprin 27Jul2006: added next line to fix the bug of "can't stop play back when press alt+f10 if repeatcount is more than 1
Private Stop_PlayBack_Loop As Boolean

Public Property Get oldWindowProc() As Long

oldWindowProc = oldWinProc

End Property

Public Property Let oldWindowProc(ByVal vNewValue As Long)

oldWinProc = vNewValue

End Property

Public Property Get IsDefault() As Boolean

IsDefault = IsDefu

End Property

Public Property Let IsDefault(ByVal vNewValue As Boolean)

IsDefu = vNewValue

End Property

Public Property Get RecordingIsOver() As Boolean

RecordingIsOver = Recording_IsOver

End Property

Public Property Let RecordingIsOver(ByVal vNewValue As Boolean)

Recording_IsOver = vNewValue

End Property

Public Property Get ProLog() As Variant

ProLog = Pro_Log

End Property


Public Property Let ProLog(ByVal vNewValue As Variant)

Pro_Log = vNewValue

End Property

Public Property Get UnqiueId() As Variant

UnqiueId = Uni

End Property

Public Property Let UnqiueId(ByVal vNewValue As Variant)

Uni = vNewValue

End Property

Public Property Get PlayPrelog() As Variant

PlayPrelog = Play_ProLog

End Property

Public Property Let PlayPrelog(ByVal vNewValue As Variant)

Play_ProLog = vNewValue

End Property

Public Property Get PlayUniqueTitId() As Variant

PlayUniqueTitId = Play_Uni

End Property

Public Property Let PlayUniqueTitId(ByVal vNewValue As Variant)

Play_Uni = vNewValue

End Property
'//cprin added next Property to support BackUpAll feature
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get BuckupAllType() As Variant
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    BuckupAllType = BuckupAll_Type
End Property
'//cprin added next Property to support BackUpAll feature
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Let BuckupAllType(ByVal vNewValue As Variant)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    BuckupAll_Type = vNewValue
End Property
'//cprin 27Jul2006: added next Property to fix the bug of "can't stop play back when press alt+f10 if repeatcount is more than 1
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get StopPlayBackLoop() As Boolean
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    StopPlayBackLoop = Stop_PlayBack_Loop
End Property
'//cprin 27Jul2006: added next Property to fix the bug of "can't stop play back when press alt+f10 if repeatcount is more than 1
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Let StopPlayBackLoop(ByVal vNewValue As Boolean)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Stop_PlayBack_Loop = vNewValue
End Property
