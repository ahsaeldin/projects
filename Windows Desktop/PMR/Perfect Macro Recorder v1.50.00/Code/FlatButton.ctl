VERSION 5.00
Begin VB.UserControl FlatButton 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
End
Attribute VB_Name = "FlatButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type LOGPEN
  lopnStyle As Long
  lopnWidth As POINTAPI
  lopnColor As Long
End Type

Private Enum ColConst
  COLOR_ACTIVEBORDER = 10
  COLOR_ACTIVECAPTION = 2
  COLOR_ADJ_MAX = 100
  COLOR_ADJ_MIN = -100
  COLOR_APPWORKSPACE = 12
  COLOR_BACKGROUND = 1
  COLOR_BTNFACE = 15
  COLOR_BTNHIGHLIGHT = 20
  COLOR_BTNSHADOW = 16
  COLOR_BTNTEXT = 18
  COLOR_CAPTIONTEXT = 9
  COLOR_GRAYTEXT = 17
  COLOR_HIGHLIGHT = 13
  COLOR_HIGHLIGHTTEXT = 14
  COLOR_INACTIVEBORDER = 11
  COLOR_INACTIVECAPTION = 3
  COLOR_INACTIVECAPTIONTEXT = 19
  COLOR_MENU = 4
  COLOR_MENUTEXT = 7
  COLOR_SCROLLBAR = 0
  COLOR_WINDOW = 5
  COLOR_WINDOWFRAME = 6
  COLOR_WINDOWTEXT = 8
End Enum

Private Const DT_CENTER = &H1

Private Declare Function CreatePen Lib "GDI32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As ColConst) As Long
Private Declare Function MoveToEx Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Dim hHBr As Long
Dim hFRgn As Long
Dim hRRgn As Long
Dim hRPen As Long
Dim LP As LOGPEN
Dim hFFBrush As Long
Dim mIcon As Long
Dim Cnt As Long
Dim CurPen As Long
Dim OldPen As Long
Dim dm As POINTAPI
Dim R As RECT
Dim HasFocus As Boolean
Dim strCaption As String
Dim ShowFocusRect As Boolean

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Show()
Public Event Hide()
Public Event Resize()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Function GetPen(ByVal nWidth As Long, ByVal Clr As Long) As Long
  GetPen = CreatePen(0, nWidth, Clr)
End Function

Public Sub ThreedBox(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional Sunken As Boolean = False)
  If Sunken = False Then
     CurPen = GetPen(1, GetSysColor(COLOR_BTNHIGHLIGHT))
  Else
     CurPen = GetPen(1, GetSysColor(COLOR_BTNSHADOW))
  End If

  OldPen = SelectObject(UserControl.hdc, CurPen)
  MoveToEx UserControl.hdc, X1, Y2, dm
  LineTo UserControl.hdc, X1, Y1
  LineTo UserControl.hdc, X2, Y1
  SelectObject UserControl.hdc, OldPen
  DeleteObject CurPen

  If Sunken = False Then
     CurPen = GetPen(1, GetSysColor(COLOR_BTNSHADOW))
  Else
     CurPen = GetPen(1, GetSysColor(COLOR_BTNHIGHLIGHT))
  End If

  OldPen = SelectObject(UserControl.hdc, CurPen)
  MoveToEx UserControl.hdc, X2, Y1, dm
  LineTo UserControl.hdc, X2, Y2
  LineTo UserControl.hdc, X1 - 1, Y2
  SelectObject UserControl.hdc, OldPen
  DeleteObject CurPen
End Sub

Private Sub UserControl_Click()
  '// Raise the event
  RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  '// Raise the event
  RaiseEvent DblClick
End Sub

Private Sub UserControl_ExitFocus()
  '// Focus Rectangle
  UserControl_LostFocus
End Sub

Private Sub UserControl_GotFocus()
  '// Focus Rectangle
  If ShowFocusRect = True Then
     SetRect R, 3, 3, ScaleWidth - 3, ScaleHeight - 3
     DrawFocusRect UserControl.hdc, R
     HasFocus = True
  End If
End Sub

Private Sub UserControl_Hide()
  '// Raise the event
  RaiseEvent Hide
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  '// Raise the event
  RaiseEvent KeyDown(KeyCode, Shift)

  '// Push button down
  If KeyCode = vbKeySpace Then UserControl_MouseDown 1, 0, 1, 1
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  '// Raise the event
  RaiseEvent KeyUp(KeyCode, Shift)

  '// Push button down
  If KeyCode = vbKeySpace Then
     UserControl_MouseUp 1, 0, 1, 1
     RaiseEvent Click
  End If
  If KeyCode = vbKeyReturn Then RaiseEvent Click
End Sub

Private Sub UserControl_LostFocus()
  '// Kill the Focus rectangle
  HasFocus = False
  UserControl_Paint
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
     UserControl.Cls

     '// Draw border
     ThreedBox 0, 0, ScaleWidth - 1, ScaleHeight - 1, True

     '// Draw the caption
     SetRect R, 0, ScaleHeight / 2 - 5, ScaleWidth + 2, ScaleHeight
     DrawText UserControl.hdc, strCaption, Len(strCaption), R, DT_CENTER

     '// Focus Rectangle
     If ShowFocusRect = True Then
        HasFocus = True
        SetRect R, 3, 3, ScaleWidth - 3, ScaleHeight - 3
        DrawFocusRect UserControl.hdc, R
     End If

     '// Raise the event
     RaiseEvent MouseDown(Button, Shift, X, Y)
  End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  '// Raise the event
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
     UserControl.Cls

     '// Draw border
     ThreedBox 0, 0, ScaleWidth - 1, ScaleHeight - 1, False

     '// Draw the caption
     SetRect R, 0, ScaleHeight / 2 - 6, ScaleWidth, ScaleHeight
     DrawText UserControl.hdc, strCaption, Len(strCaption), R, DT_CENTER

     '// Focus Rectangle
     If ShowFocusRect = True Then
        SetRect R, 3, 3, ScaleWidth - 3, ScaleHeight - 3
        DrawFocusRect UserControl.hdc, R
     End If

  End If

  '// Raise the event
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
  UserControl.Cls

  '// Draw border
  ThreedBox 0, 0, ScaleWidth - 1, ScaleHeight - 1, False

  '// Draw the caption
  SetRect R, 0, ScaleHeight / 2 - 6, ScaleWidth, ScaleHeight
  DrawText UserControl.hdc, strCaption, Len(strCaption), R, DT_CENTER

  '// Focus Rectangle
  If HasFocus = True And ShowFocusRect = True Then
     SetRect R, 3, 3, ScaleWidth - 3, ScaleHeight - 3
     DrawFocusRect UserControl.hdc, R
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  strCaption = PropBag.ReadProperty("Caption", "")
  ShowFocusRect = PropBag.ReadProperty("FocusRect", True)
End Sub

Private Sub UserControl_Resize()
  '// Raise the event
  RaiseEvent Resize
End Sub

Private Sub UserControl_Show()
  '// Raise the event
  RaiseEvent Show
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text shown on the control"
  Caption = strCaption
End Property

Public Property Let Caption(ByVal vNewValue As String)
  strCaption = vNewValue
  PropertyChanged "Caption"
  UserControl.Cls
  UserControl_Paint
End Property

Public Property Get FocusRect() As Boolean
Attribute FocusRect.VB_Description = "Return/set if the control shows a focus rect"
  FocusRect = ShowFocusRect
End Property

Public Property Let FocusRect(ByVal vNewValue As Boolean)
  ShowFocusRect = vNewValue
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Caption", strCaption, ""
  PropBag.WriteProperty "FocusRect", ShowFocusRect, True
End Sub
