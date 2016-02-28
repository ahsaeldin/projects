VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   750
   EditAtDesignTime=   -1  'True
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   750
   ScaleWidth      =   750
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Function CreateToolTip2(ByVal lParentHwnd As Long, sty As Long) As Long
    Dim ptTip As Long
    Dim ret As Long
    Dim lWinStyle As Long
     Dim ToolTipOldProc As Long
     
 
    
       
    CreateToolTip2 = CreateWindowEx(0&, _
            TOOLTIPS_CLASSA, _
            vbNullString, _
            sty, _
            0&, _
            0&, _
            0&, _
            0&, _
            lParentHwnd, _
            0&, _
            App.hInstance, _
            0&)
  
    
End Function



