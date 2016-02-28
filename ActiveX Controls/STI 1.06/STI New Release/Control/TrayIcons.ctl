VERSION 5.00
Begin VB.UserControl TrayIcons 
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   525
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   510
   ScaleWidth      =   525
   ToolboxBitmap   =   "TrayIcons.ctx":0000
   Begin VB.PictureBox Picture1 
      Height          =   514
      Left            =   0
      Picture         =   "TrayIcons.ctx":0312
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   0
      Top             =   0
      Width           =   514
   End
End
Attribute VB_Name = "TrayIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type TrayIcon
    hwnd As Long
    uId As Long
    ucallbackMessage As Long
    Param1(1) As Long
    hIcon As Long
    Param2(2) As Long
    APath As String
    ToolTip As String
    Bitmap As Long
    Command As Long
    State As Byte
    Style As Byte
    Data As Long
    str As Long
End Type


Dim MSFlexGrid1 As MSFlexGrid

Private Sub UserControl_Initialize()
 Set MSFlexGrid1 = Controls.Add("MSFlexGridLib.MSFlexGrid", "ASdf")
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 540
UserControl.Height = 540
End Sub

Public Function GetSysTrayIcons() As TrayIcon()

Dim ticons() As TrayIcon

    If IsWindowsNT Then
       
       ticons = GetTrayIcons
       
       GetSysTrayIcons = ticons
       
    Else
       
       ticons = GetTrayIcons_9x
      
       GetSysTrayIcons = ticons
       
    End If


End Function

Friend Function ExtenderName() As String
  ExtenderName = Extender.Name
End Function

Public Property Get IconCount() As Variant
IconCount = IconNumber + 1
End Property

Public Function Hide(index As Integer)
Dim TrayList() As TrayIcon
Dim TrayI As NOTIFYICONDATA
        
    bybassError = True
    TrayList = GetSysTrayIcons
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = TrayList(index - 1).hwnd
    TrayI.uId = TrayList(index - 1).uId
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.hIcon = TrayList(index - 1).hIcon
    Shell_NotifyIcon NIM_DELETE, TrayI
 
End Function

Private Sub initFlexGrid(FlexGrid As Object)
Dim str As String
  FlexGrid.Cols = 7
  FlexGrid.Rows = 100
  str = Len("Process identifier")
  FlexGrid.ColWidth(0) = 300
  FlexGrid.TextMatrix(0, 0) = "I"
  FlexGrid.ColWidth(1) = 3000
  FlexGrid.TextMatrix(0, 1) = "Application Path"
  FlexGrid.ColWidth(2) = 1300
  FlexGrid.TextMatrix(0, 2) = "uID"
  FlexGrid.ColWidth(3) = 1300
  FlexGrid.TextMatrix(0, 3) = "hWnd"
  FlexGrid.ColWidth(4) = 1300
  FlexGrid.TextMatrix(0, 4) = "hIcon"
  FlexGrid.ColWidth(5) = 1700
  FlexGrid.TextMatrix(0, 5) = "ToolTip"
  FlexGrid.ColWidth(6) = 1500
  FlexGrid.TextMatrix(0, 6) = "uCallbackMessage"
  FlexGrid.ForeColorFixed = vbBlue
  FlexGrid.Width = 10455
  FlexGrid.Height = 3615
  FlexGrid.SelectionMode = 1
  FlexGrid.AllowUserResizing = 1
End Sub

Public Function FillMSFlexGrid(FlexGrid As Object)

Dim i As Integer
Dim TrayList() As TrayIcon

bybassError = False

TrayList = GetSysTrayIcons

FlexGrid.Clear

initFlexGrid FlexGrid

For i = 0 To IconCount - 1
    
    FlexGrid.TextMatrix(i + 1, 0) = i + 1
    If TrayList(i).APath = "" Then
       FlexGrid.TextMatrix(i + 1, 1) = "N/A"
    Else
       FlexGrid.TextMatrix(i + 1, 1) = TrayList(i).APath
    End If
    FlexGrid.TextMatrix(i + 1, 2) = TrayList(i).uId
    FlexGrid.TextMatrix(i + 1, 3) = TrayList(i).hwnd
    FlexGrid.TextMatrix(i + 1, 4) = TrayList(i).hIcon
    FlexGrid.TextMatrix(i + 1, 5) = TrayList(i).ToolTip
    FlexGrid.TextMatrix(i + 1, 6) = "&H" & Hex$(TrayList(i).ucallbackMessage)
   
   
   
Next i

End Function

Public Sub sendreview()

ShellExecute vbNull, vbNullString, "mailto:support@cpringold.com", vbNullString, vbNullString, 1

End Sub
