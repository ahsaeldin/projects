Attribute VB_Name = "SubClass"
Option Explicit



Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_USER = &H400
Public Const GWL_WNDPROC = (-4)

Public Tags As CTags
Dim MsgNames As New Collection

Public Function NewWindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const WM_NCDESTROY = &H82


Rem Load the message names.
    LoadMsgNames
    
    Debug.Print Hex(msg) & ": " & MsgName(msg)
   


'don't forget mousemove
      
  '  End If

Rem If we're being destroyed,
Rem restore the original WindowProc.
    If msg = WM_NCDESTROY Then
    
        SetWindowLong frmJMacro.hwnd, GWL_WNDPROC, Tags.oldWindowProc
    
    End If
    
Rem Continue normal processing. VERY IMPORTANT!
Rem If it isn't taking special action, it should
Rem pass the message on to the original WindowProc
Rem for normal processing. This is extremely important!
Rem If you don't pass the message along, the object will
Rem not be able to update its appearance when it is uncovered,
Rem correctly draw itself, and perform other standard Windows functions.
Rem The new WindowProc should return whatever code the original WindowProc returns.

    NewWindowProc = CallWindowProc(Tags.oldWindowProc, frmJMacro.hwnd, msg, wParam, lParam)
    
End Function

' *********************************************
' Fill the MsgNames collection with windows
' message names.
' *********************************************
Public Sub LoadMsgNames()
Static done_before As Boolean

    If done_before Then Exit Sub
    done_before = True

    MsgNames.Add "WM_NULL", "0"
    MsgNames.Add "WM_CREATE", "1"
    MsgNames.Add "WM_DESTROY", "2"
    MsgNames.Add "WM_MOVE", "3"
    MsgNames.Add "WM_SIZE", "5"
    MsgNames.Add "WM_ACTIVATE", "6"
    MsgNames.Add "WM_SETFOCUS", "7"
    MsgNames.Add "WM_KILLFOCUS", "8"
    MsgNames.Add "WM_ENABLE", "A"
    MsgNames.Add "WM_SETREDRAW", "B"
    MsgNames.Add "WM_SETTEXT", "C"
    MsgNames.Add "WM_GETTEXT", "D"
    MsgNames.Add "WM_GETTEXTLENGTH", "E"
    MsgNames.Add "WM_PAINT", "F"
    MsgNames.Add "WM_CLOSE", "10"
    MsgNames.Add "WM_QUERYENDSESSION", "11"
    MsgNames.Add "WM_QUIT", "12"
    MsgNames.Add "WM_QUERYOPEN", "13"
    MsgNames.Add "WM_ERASEBKGND", "14"
    MsgNames.Add "WM_SYSCOLORCHANGE", "15"
    MsgNames.Add "WM_ENDSESSION", "16"
    MsgNames.Add "WM_SHOWWINDOW", "18"
    MsgNames.Add "WM_WININICHANGE", "1A"
    MsgNames.Add "WM_DEVMODECHANGE", "1B"
    MsgNames.Add "WM_ACTIVATEAPP", "1C"
    MsgNames.Add "WM_FONTCHANGE", "1D"
    MsgNames.Add "WM_TIMECHANGE", "1E"
    MsgNames.Add "WM_CANCELMODE", "1F"
    MsgNames.Add "WM_SETCURSOR", "20"
    MsgNames.Add "WM_MOUSEACTIVATE", "21"
    MsgNames.Add "WM_CHILDACTIVATE", "22"
    MsgNames.Add "WM_QUEUESYNC", "23"
    MsgNames.Add "WM_GETMINMAXINFO", "24"
    MsgNames.Add "WM_PAINTICON", "26"
    MsgNames.Add "WM_ICONERASEBKGND", "27"
    MsgNames.Add "WM_NEXTDLGCTL", "28"
    MsgNames.Add "WM_SPOOLERSTATUS", "2A"
    MsgNames.Add "WM_DRAWITEM", "2B"
    MsgNames.Add "WM_MEASUREITEM", "2C"
    MsgNames.Add "WM_DELETEITEM", "2D"
    MsgNames.Add "WM_VKEYTOITEM", "2E"
    MsgNames.Add "WM_CHARTOITEM", "2F"
    MsgNames.Add "WM_SETFONT", "30"
    MsgNames.Add "WM_GETFONT", "31"
    MsgNames.Add "WM_SETHOTKEY", "32"
    MsgNames.Add "WM_GETHOTKEY", "33"
    MsgNames.Add "WM_QUERYDRAGICON", "37"
    MsgNames.Add "WM_COMPAREITEM", "39"
    MsgNames.Add "WM_COMPACTING", "41"
    MsgNames.Add "WM_OTHERWINDOWCREATED (obsolete)", "42"
    MsgNames.Add "WM_OTHERWINDOWDESTROYED (obsolete)", "43"
    MsgNames.Add "WM_COMMNOTIFY (obsolete)", "44"
    MsgNames.Add "WM_WINDOWPOSCHANGING", "46"
    MsgNames.Add "WM_WINDOWPOSCHANGED", "47"
    MsgNames.Add "WM_POWER", "48"
    MsgNames.Add "WM_COPYDATA", "4A"
    MsgNames.Add "WM_CANCELJOURNAL", "4B"
    MsgNames.Add "WM_NCCREATE", "81"
    MsgNames.Add "WM_NCDESTROY", "82"
    MsgNames.Add "WM_NCCALCSIZE", "83"
    MsgNames.Add "WM_NCHITTEST", "84"
    MsgNames.Add "WM_NCPAINT", "85"
    MsgNames.Add "WM_NCACTIVATE", "86"
    MsgNames.Add "WM_GETDLGCODE", "87"
    MsgNames.Add "WM_NCMOUSEMOVE", "A0"
    MsgNames.Add "WM_NCLBUTTONDOWN", "A1"
    MsgNames.Add "WM_NCLBUTTONUP", "A2"
    MsgNames.Add "WM_NCLBUTTONDBLCLK", "A3"
    MsgNames.Add "WM_NCRBUTTONDOWN", "A4"
    MsgNames.Add "WM_NCRBUTTONUP", "A5"
    MsgNames.Add "WM_NCRBUTTONDBLCLK", "A6"
    MsgNames.Add "WM_NCMBUTTONDOWN", "A7"
    MsgNames.Add "WM_NCMBUTTONUP", "A8"
    MsgNames.Add "WM_NCMBUTTONDBLCLK", "A9"
    MsgNames.Add "WM_KEYFIRST/WM_KEYDOWN", "100"
'    MsgNames.Add "WM_KEYDOWN", "100"
    MsgNames.Add "WM_KEYUP", "101"
    MsgNames.Add "WM_CHAR", "102"
    MsgNames.Add "WM_DEADCHAR", "103"
    MsgNames.Add "WM_SYSKEYDOWN", "104"
    MsgNames.Add "WM_SYSKEYUP", "105"
    MsgNames.Add "WM_SYSCHAR", "106"
    MsgNames.Add "WM_SYSDEADCHAR", "107"
    MsgNames.Add "WM_KEYLAST", "108"
    MsgNames.Add "WM_INITDIALOG", "110"
    MsgNames.Add "WM_COMMAND", "111"
    MsgNames.Add "WM_SYSCOMMAND", "112"
    MsgNames.Add "WM_TIMER", "113"
    MsgNames.Add "WM_HSCROLL", "114"
    MsgNames.Add "WM_VSCROLL", "115"
    MsgNames.Add "WM_INITMENU", "116"
    MsgNames.Add "WM_INITMENUPOPUP", "117"
    MsgNames.Add "WM_MENUSELECT", "11F"
    MsgNames.Add "WM_MENUCHAR", "120"
    MsgNames.Add "WM_ENTERIDLE", "121"
    MsgNames.Add "WM_CTLCOLORMSGBOX", "132"
    MsgNames.Add "WM_CTLCOLOREDIT", "133"
    MsgNames.Add "WM_CTLCOLORLISTBOX", "134"
    MsgNames.Add "WM_CTLCOLORBTN", "135"
    MsgNames.Add "WM_CTLCOLORDLG", "136"
    MsgNames.Add "WM_CTLCOLORSCROLLBAR", "137"
    MsgNames.Add "WM_CTLCOLORSTATIC", "138"
    MsgNames.Add "WM_MOUSEFIRST/WM_MOUSEMOVE", "200"
'    MsgNames.Add "WM_MOUSEMOVE", "200"
    MsgNames.Add "WM_LBUTTONDOWN", "201"
    MsgNames.Add "WM_LBUTTONUP", "202"
    MsgNames.Add "WM_LBUTTONDBLCLK", "203"
    MsgNames.Add "WM_RBUTTONDOWN", "204"
    MsgNames.Add "WM_RBUTTONUP", "205"
    MsgNames.Add "WM_RBUTTONDBLCLK", "206"
    MsgNames.Add "WM_MBUTTONDOWN", "207"
    MsgNames.Add "WM_MBUTTONUP", "208"
    MsgNames.Add "WM_MBUTTONDBLCLK/WM_MOUSELAST", "209"
'    MsgNames.Add "WM_MOUSELAST", "209"
    MsgNames.Add "WM_PARENTNOTIFY", "210"
    MsgNames.Add "WM_ENTERMENULOOP", "211"
    MsgNames.Add "WM_EXITMENULOOP", "212"
    MsgNames.Add "WM_MDICREATE", "220"
    MsgNames.Add "WM_MDIDESTROY", "221"
    MsgNames.Add "WM_MDIACTIVATE", "222"
    MsgNames.Add "WM_MDIRESTORE", "223"
    MsgNames.Add "WM_MDINEXT", "224"
    MsgNames.Add "WM_MDIMAXIMIZE", "225"
    MsgNames.Add "WM_MDITILE", "226"
    MsgNames.Add "WM_MDICASCADE", "227"
    MsgNames.Add "WM_MDIICONARRANGE", "228"
    MsgNames.Add "WM_MDIGETACTIVE", "229"
    MsgNames.Add "WM_MDISETMENU", "230"
    MsgNames.Add "WM_DROPFILES", "233"
    MsgNames.Add "WM_MDIREFRESHMENU", "234"
    MsgNames.Add "WM_CUT", "300"
    MsgNames.Add "WM_COPY", "301"
    MsgNames.Add "WM_PASTE", "302"
    MsgNames.Add "WM_CLEAR", "303"
    MsgNames.Add "WM_UNDO", "304"
    MsgNames.Add "WM_RENDERFORMAT", "305"
    MsgNames.Add "WM_RENDERALLFORMATS", "306"
    MsgNames.Add "WM_DESTROYCLIPBOARD", "307"
    MsgNames.Add "WM_DRAWCLIPBOARD", "308"
    MsgNames.Add "WM_PAINTCLIPBOARD", "309"
    MsgNames.Add "WM_VSCROLLCLIPBOARD", "30A"
    MsgNames.Add "WM_SIZECLIPBOARD", "30B"
    MsgNames.Add "WM_ASKCBFORMATNAME", "30C"
    MsgNames.Add "WM_CHANGECBCHAIN", "30D"
    MsgNames.Add "WM_HSCROLLCLIPBOARD", "30E"
    MsgNames.Add "WM_QUERYNEWPALETTE", "30F"
    MsgNames.Add "WM_PALETTEISCHANGING", "310"
    MsgNames.Add "WM_PALETTECHANGED", "311"
    MsgNames.Add "WM_HOTKEY", "312"
    MsgNames.Add "WM_PENWINFIRST", "380"
    MsgNames.Add "WM_PENWINLAST", "38F"
    MsgNames.Add "WM_USER", "400"
    
    ' Edit control messages.
    MsgNames.Add "EM_GETSEL", "B0"
    MsgNames.Add "EM_SETSEL", "B1"
    MsgNames.Add "EM_GETRECT", "B2"
    MsgNames.Add "EM_SETRECT", "B3"
    MsgNames.Add "EM_SETRECTNP", "B4"
    MsgNames.Add "EM_SCROLL", "B5"
    MsgNames.Add "EM_LINESCROLL", "B6"
    MsgNames.Add "EM_SCROLLCARET", "B7"
    MsgNames.Add "EM_GETMODIFY", "B8"
    MsgNames.Add "EM_SETMODIFY", "B9"
    MsgNames.Add "EM_GETLINECOUNT", "BA"
    MsgNames.Add "EM_LINEINDEX", "BB"
    MsgNames.Add "EM_SETHANDLE", "BC"
    MsgNames.Add "EM_GETHANDLE", "BD"
    MsgNames.Add "EM_GETTHUMB", "BE"
    MsgNames.Add "EM_LINELENGTH", "C1"
    MsgNames.Add "EM_REPLACESEL", "C2"
    MsgNames.Add "EM_GETLINE", "C4"
    MsgNames.Add "EM_LIMITTEXT", "C5"
    MsgNames.Add "EM_CANUNDO", "C6"
    MsgNames.Add "EM_UNDO", "C7"
    MsgNames.Add "EM_FMTLINES", "C8"
    MsgNames.Add "EM_LINEFROMCHAR", "C9"
    MsgNames.Add "EM_SETTABSTOPS", "CB"
    MsgNames.Add "EM_SETPASSWORDCHAR", "CC"
    MsgNames.Add "EM_EMPTYUNDOBUFFER", "CD"
    MsgNames.Add "EM_GETFIRSTVISIBLELINE", "CE"
    MsgNames.Add "EM_SETREADONLY", "CF"
    MsgNames.Add "EM_SETWORDBREAKPROC", "D0"
    MsgNames.Add "EM_GETWORDBREAKPROC", "D1"
    MsgNames.Add "EM_GETPASSWORDCHAR", "D2"
    
    ' Button Control Messages
    MsgNames.Add "BM_GETCHECK", "HF0"
    MsgNames.Add "BM_SETCHECK", "HF1"
    MsgNames.Add "BM_GETSTATE", "HF2"
    MsgNames.Add "BM_SETSTATE", "HF3"
    MsgNames.Add "BM_SETSTYLE", "HF4"
    
    ' Listbox messages.
    MsgNames.Add "LB_ADDSTRING", "H180"
    MsgNames.Add "LB_INSERTSTRING", "H181"
    MsgNames.Add "LB_DELETESTRING", "H182"
    MsgNames.Add "LB_SELITEMRANGEEX", "H183"
    MsgNames.Add "LB_RESETCONTENT", "H184"
    MsgNames.Add "LB_SETSEL", "H185"
    MsgNames.Add "LB_SETCURSEL", "H186"
    MsgNames.Add "LB_GETSEL", "H187"
    MsgNames.Add "LB_GETCURSEL", "H188"
    MsgNames.Add "LB_GETTEXT", "H189"
    MsgNames.Add "LB_GETTEXTLEN", "H18A"
    MsgNames.Add "LB_GETCOUNT", "H18B"
    MsgNames.Add "LB_SELECTSTRING", "H18C"
    MsgNames.Add "LB_DIR", "H18D"
    MsgNames.Add "LB_GETTOPINDEX", "H18E"
    MsgNames.Add "LB_FINDSTRING", "H18F"
    MsgNames.Add "LB_GETSELCOUNT", "H190"
    MsgNames.Add "LB_GETSELITEMS", "H191"
    MsgNames.Add "LB_SETTABSTOPS", "H192"
    MsgNames.Add "LB_GETHORIZONTALEXTENT", "H193"
    MsgNames.Add "LB_SETHORIZONTALEXTENT", "H194"
    MsgNames.Add "LB_SETCOLUMNWIDTH", "H195"
    MsgNames.Add "LB_ADDFILE", "H196"
    MsgNames.Add "LB_SETTOPINDEX", "H197"
    MsgNames.Add "LB_GETITEMRECT", "H198"
    MsgNames.Add "LB_GETITEMDATA", "H199"
    MsgNames.Add "LB_SETITEMDATA", "H19A"
    MsgNames.Add "LB_SELITEMRANGE", "H19B"
    MsgNames.Add "LB_SETANCHORINDEX", "H19C"
    MsgNames.Add "LB_GETANCHORINDEX", "H19D"
    MsgNames.Add "LB_SETCARETINDEX", "H19E"
    MsgNames.Add "LB_GETCARETINDEX", "H19F"
    MsgNames.Add "LB_SETITEMHEIGHT", "H1A0"
    MsgNames.Add "LB_GETITEMHEIGHT", "H1A1"
    MsgNames.Add "LB_FINDSTRINGEXACT", "H1A2"
    MsgNames.Add "LB_SETLOCALE", "H1A5"
    MsgNames.Add "LB_GETLOCALE", "H1A6"
    MsgNames.Add "LB_SETCOUNT", "H1A7"
    MsgNames.Add "LB_MSGMAX", "H1A8"
    
    ' Combo Box messages.
    MsgNames.Add "CB_GETEDITSEL", "140"
    MsgNames.Add "CB_LIMITTEXT", "141"
    MsgNames.Add "CB_SETEDITSEL", "142"
    MsgNames.Add "CB_ADDSTRING", "143"
    MsgNames.Add "CB_DELETESTRING", "144"
    MsgNames.Add "CB_DIR", "145"
    MsgNames.Add "CB_GETCOUNT", "146"
    MsgNames.Add "CB_GETCURSEL", "147"
    MsgNames.Add "CB_GETLBTEXT", "148"
    MsgNames.Add "CB_GETLBTEXTLEN", "149"
    MsgNames.Add "CB_INSERTSTRING", "14A"
    MsgNames.Add "CB_RESETCONTENT", "14B"
    MsgNames.Add "CB_FINDSTRING", "14C"
    MsgNames.Add "CB_SELECTSTRING", "14D"
    MsgNames.Add "CB_SETCURSEL", "14E"
    MsgNames.Add "CB_SHOWDROPDOWN", "14F"
    MsgNames.Add "CB_GETITEMDATA", "150"
    MsgNames.Add "CB_SETITEMDATA", "151"
    MsgNames.Add "CB_GETDROPPEDCONTROLRECT", "152"
    MsgNames.Add "CB_SETITEMHEIGHT", "153"
    MsgNames.Add "CB_GETITEMHEIGHT", "154"
    MsgNames.Add "CB_SETEXTENDEDUI", "155"
    MsgNames.Add "CB_GETEXTENDEDUI", "156"
    MsgNames.Add "CB_GETDROPPEDSTATE", "157"
    MsgNames.Add "CB_FINDSTRINGEXACT", "158"
    MsgNames.Add "CB_SETLOCALE", "159"
    MsgNames.Add "CB_GETLOCALE", "15A"
    MsgNames.Add "CB_MSGMAX", "15B"

    '  Scroll bar messages.
    MsgNames.Add "SBM_SETPOS", "E0"
    MsgNames.Add "SBM_GETPOS", "E1"
    MsgNames.Add "SBM_SETRANGE", "E2"
    MsgNames.Add "SBM_SETRANGEREDRAW", "E6"
    MsgNames.Add "SBM_GETRANGE", "E3"
    MsgNames.Add "SBM_ENABLE_ARROWS", "E4"

    '  DDE window messages
    MsgNames.Add "WM_DDE_FIRST", "3E0"
    MsgNames.Add "WM_DDE_INITIATE = (WM_DDE_FIRST)"
    MsgNames.Add "WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)"
    MsgNames.Add "WM_DDE_ADVISE = (WM_DDE_FIRST + 2)"
    MsgNames.Add "WM_DDE_UNADVISE = (WM_DDE_FIRST + 3)"
    MsgNames.Add "WM_DDE_ACK = (WM_DDE_FIRST + 4)"
    MsgNames.Add "WM_DDE_DATA = (WM_DDE_FIRST + 5)"
    MsgNames.Add "WM_DDE_REQUEST = (WM_DDE_FIRST + 6)"
    MsgNames.Add "WM_DDE_POKE = (WM_DDE_FIRST + 7)"
    MsgNames.Add "WM_DDE_EXECUTE = (WM_DDE_FIRST + 8)"
    MsgNames.Add "WM_DDE_LAST = (WM_DDE_FIRST + 8)"

    ' Multimedia Window Messages"
    MsgNames.Add "MM_JOY1MOVE", "3A0"
    MsgNames.Add "MM_JOY2MOVE", "3A1"
    MsgNames.Add "MM_JOY1ZMOVE", "3A2"
    MsgNames.Add "MM_JOY2ZMOVE", "3A3"
    MsgNames.Add "MM_JOY1BUTTONDOWN", "3B5"
    MsgNames.Add "MM_JOY2BUTTONDOWN", "3B6"
    MsgNames.Add "MM_JOY1BUTTONUP", "3B7"
    MsgNames.Add "MM_JOY2BUTTONUP", "3B8"
    MsgNames.Add "MM_MCINOTIFY", "3B9"
    MsgNames.Add "MM_MCISYSTEM_STRING", "3CA"
    MsgNames.Add "MM_WOM_OPEN", "3BB"
    MsgNames.Add "MM_WOM_CLOSE", "3BC"
    MsgNames.Add "MM_WOM_DONE", "3BD"
    MsgNames.Add "MM_WIM_OPEN", "3BE"
    MsgNames.Add "MM_WIM_CLOSE", "3BF"
    MsgNames.Add "MM_WIM_DATA", "3C0"
    MsgNames.Add "MM_MIM_OPEN", "3C1"
    MsgNames.Add "MM_MIM_CLOSE", "3C2"
    MsgNames.Add "MM_MIM_DATA", "3C3"
    MsgNames.Add "MM_MIM_LONGDATA", "3C4"
    MsgNames.Add "MM_MIM_ERROR", "3C5"
    MsgNames.Add "MM_MIM_LONGERROR", "3C6"
    MsgNames.Add "MM_MOM_OPEN", "3C7"
    MsgNames.Add "MM_MOM_CLOSE", "3C8"
    MsgNames.Add "MM_MOM_DONE", "3C9"
End Sub
Public Function MsgName(ByVal num As Long) As String
    On Error Resume Next
    MsgName = MsgNames(Hex$(num))
    If Err.Number = 0 Then Exit Function

    If num >= 0 And num < WM_USER Then
        MsgName = "Range 1 message reserved for Windows"
    ElseIf num >= WM_USER And num <= &H7FFF Then
        MsgName = "Reserved for private window classes"
    ElseIf num >= &H8000 And num <= &HBFFF Then
        MsgName = "Range 3 message reserved for Windows"
    ElseIf num >= &H8000 And num <= &HBFFF Then
        MsgName = "String message for use by applications"
    Else
        MsgName = "Unknown message" & Str$(num)
    End If
End Function
