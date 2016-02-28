VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFixBalloonTooltipError"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function FixTheSillyError(phwnnd As Long, stylist As Long) As Long

FixTheSillyError = CreateWindowEx(0&, _
            TOOLTIPS_CLASSA, _
            vbNullString, _
            stylist, _
            0&, _
            0&, _
            0&, _
            0&, _
            phwnnd, _
            0&, _
            App.hInstance, _
            0&)
            
End Function
