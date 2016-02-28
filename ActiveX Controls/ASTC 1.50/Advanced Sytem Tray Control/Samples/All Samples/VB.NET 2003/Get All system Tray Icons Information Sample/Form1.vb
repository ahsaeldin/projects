'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
'Description:  A sample code demonstrate how to use ASTC control to get all system tray icons
'                Information as well as how to hide a specified icon from system tray area.
'==========================================================================================

'First of all don't afraid from all these lines of codes, the main line which return the system
'Tray information is only one line which located in FillFlexGrid method [CoreLine] at line 10
'All the remaining lines is to fill the returned data from GetSysTrayIcons method into FlexGrid
'To display the data in a good manner.
'How i return this information?
'Well
'GetSysTrayIcons method returns an array of TrayIconInfo type
'Wait what does it means?
'We have a Type Called TrayIconInfo; you can use Object Browser to see this type in ASTC
'When you call GetSysTrayIcons method, it returns an array of this type contains the information
'I don't understand
'Well take this example
'If the system tray area contains three icons say [sound control icon,
'MacAfee icon and network connection icon].
'Then the returned data will be organized in an Array and the array will have three elements
'And every element will be a type called TrayIconInfo and TrayIconInfo will contain the information.
'
'So what is this information?
'
'Public Type TrayIconInfo
'    Hwnd As Long                 Handle to the window that will receive notification messages associated with an icon in the taskbar status area.
'    uId As Long                  Application-defined identifier of the taskbar icon.
'    ucallbackMessage As Long     Application-defined message identifier. The system uses this identifier for notification messages that it sends to the window identified in hWnd. These notifications are sent when a mouse event occurs in the bounding rectangle of the icon.
'    Param1(1) As Long            unknown interpretation
'    hIcon As Long                Handle to the icon
'    Param2(2) As Long            unknown interpretation
'    APath As String              The path of the application which adds the icon to system tray area
'    ToolTip As String            The tooltip associated with the icon
'    Bitmap as Long               index of the posted icon
'    Command as Long              command id
'    State as Byte                icon state
'    Style as Byte                icon style
'    Data as Long                 pointer towards the data of the tray
'    str As Long                  pointer to tooltip
'End Type

'I think that the only problem comes from you must know that the variable which will receive the
'Returned data type must be an array of TrayIconInfo
'See this example
'
'Dim TrayList() As TrayIconInfo
'TrayList = TrayIcons1.GetSysTrayIcons
'
'As you see the variable TrayList is an array of TrayIconInfo type

'How to hide icon from system tray area?
'Because every icon has an element in the returned array, all you need to do is to pass
'The element index to HideIcon method like this
'
'TrayList1.HideIcon 1
'Which 1 is the index of the icon element (TrayIconInfo) in the array?

'That's it
'If you have any further questions or need more sample code don't hesitate to contact us
'At support@cpringold.com

Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ButHide As System.Windows.Forms.Button
    Friend WithEvents ButRestore As System.Windows.Forms.Button
    Friend WithEvents ButRefresh As System.Windows.Forms.Button
    Friend WithEvents MSFlexGrid1 As AxMSFlexGridLib.AxMSFlexGrid
    Friend WithEvents AxTrayList1 As AxASTC.AxTrayList
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.ButHide = New System.Windows.Forms.Button
        Me.ButRestore = New System.Windows.Forms.Button
        Me.ButRefresh = New System.Windows.Forms.Button
        Me.MSFlexGrid1 = New AxMSFlexGridLib.AxMSFlexGrid
        Me.AxTrayList1 = New AxASTC.AxTrayList
        CType(Me.MSFlexGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxTrayList1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ButHide
        '
        Me.ButHide.Location = New System.Drawing.Point(8, 208)
        Me.ButHide.Name = "ButHide"
        Me.ButHide.Size = New System.Drawing.Size(80, 23)
        Me.ButHide.TabIndex = 0
        Me.ButHide.Text = "Hide"
        '
        'ButRestore
        '
        Me.ButRestore.Location = New System.Drawing.Point(96, 208)
        Me.ButRestore.Name = "ButRestore"
        Me.ButRestore.Size = New System.Drawing.Size(80, 23)
        Me.ButRestore.TabIndex = 1
        Me.ButRestore.Text = "Restore Last"
        '
        'ButRefresh
        '
        Me.ButRefresh.Location = New System.Drawing.Point(184, 208)
        Me.ButRefresh.Name = "ButRefresh"
        Me.ButRefresh.Size = New System.Drawing.Size(80, 23)
        Me.ButRefresh.TabIndex = 2
        Me.ButRefresh.Text = "Refresh"
        '
        'MSFlexGrid1
        '
        Me.MSFlexGrid1.Location = New System.Drawing.Point(8, 8)
        Me.MSFlexGrid1.Name = "MSFlexGrid1"
        Me.MSFlexGrid1.OcxState = CType(resources.GetObject("MSFlexGrid1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MSFlexGrid1.Size = New System.Drawing.Size(664, 192)
        Me.MSFlexGrid1.TabIndex = 3
        '
        'AxTrayList1
        '
        Me.AxTrayList1.Enabled = True
        Me.AxTrayList1.Location = New System.Drawing.Point(640, 200)
        Me.AxTrayList1.Name = "AxTrayList1"
        Me.AxTrayList1.OcxState = CType(resources.GetObject("AxTrayList1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxTrayList1.Size = New System.Drawing.Size(36, 36)
        Me.AxTrayList1.TabIndex = 4
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(680, 238)
        Me.Controls.Add(Me.AxTrayList1)
        Me.Controls.Add(Me.MSFlexGrid1)
        Me.Controls.Add(Me.ButRefresh)
        Me.Controls.Add(Me.ButRestore)
        Me.Controls.Add(Me.ButHide)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VB.Net Tray List"
        Me.TopMost = True
        CType(Me.MSFlexGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxTrayList1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    '==============================================================================
    ' Method:       Form1_Load
    '
    ' Description:  Call FillFlexGrid to display the data
    '==============================================================================
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Fill the FlexGrid control with System Tray Data   
        FillFlexGrid()
    End Sub
    '==============================================================================
    ' Method:       ButHide_Click
    '
    ' Description:  Cool function used to hide any icon from system tray icon
    '               then refresh the FlexGrid control.
    '==============================================================================
    Private Sub ButHide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButHide.Click

        Dim TrayList() As ASTC.TrayIconInfo

        TrayList = AxTrayList1.GetSysTrayIcons
        'Save the Icon Info before Remove it in order to use
        'when we restore the icon
        OldTrayIcon = TrayList(MSFlexGrid1.RowSel - 1)
        'Hide the selected Item from the System Tray area
        AxTrayList1.HideIcon(MSFlexGrid1.RowSel())
        'Clear the FlexGrid
        MSFlexGrid1.Clear()
        'fill the FlexGrid 
        FillFlexGrid()

    End Sub
    '==============================================================================
    ' Method:       ButRefresh_Click
    '
    ' Description:  Refresh FillFlexGrid to display any new data
    '==============================================================================
    Private Sub ButRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButRefresh.Click
        FillFlexGrid()
    End Sub
    '==============================================================================
    ' Method:       ButRestore_Click
    '
    ' Description:  Restore the last removed icon To system tray.
    '==============================================================================
    Private Sub ButRestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButRestore.Click
        RestoreIcon(OldTrayIcon)
    End Sub

    '==============================================================================
    ' Method:       initFlexGrid
    '
    ' Description:  setup the flexgrid control to receive the data
    '==============================================================================
    Private Sub initFlexGrid()
        'Setup the FlexGrid 

        MSFlexGrid1.set_ColWidth(0, 300)
        MSFlexGrid1.set_TextMatrix(0, 0, "I")
        MSFlexGrid1.set_ColWidth(1, 3000)
        MSFlexGrid1.set_TextMatrix(0, 1, "Application Path")
        MSFlexGrid1.set_ColWidth(2, 1300)
        MSFlexGrid1.set_TextMatrix(0, 2, "uID")
        MSFlexGrid1.set_ColWidth(3, 1300)
        MSFlexGrid1.set_TextMatrix(0, 3, "hWnd")
        MSFlexGrid1.set_ColWidth(4, 1300)
        MSFlexGrid1.set_TextMatrix(0, 4, "hIcon")
        MSFlexGrid1.set_ColWidth(5, 1700)
        MSFlexGrid1.set_TextMatrix(0, 5, "ToolTip")
        MSFlexGrid1.set_ColWidth(6, 1500)
        MSFlexGrid1.set_TextMatrix(0, 6, "uCallbackMessage")
        MSFlexGrid1.ForeColorFixed = Color.Blue

    End Sub

    '==============================================================================
    ' Method:       FillFlexGrid
    '
    ' Description:  The core function here we get the information from GetSysTrayIcons
    '               then fill the FlexGrid control with these info.
    '==============================================================================
    Private Sub FillFlexGrid()

        Dim i As Integer
        Dim TrayList() As ASTC.TrayIconInfo

        initFlexGrid()

        'Get the System Tray Data
        TrayList = AxTrayList1.GetSysTrayIcons()

        'Fill the FlexGrid with Data
        For i = 0 To AxTrayList1.IconCount - 1

            MSFlexGrid1.set_TextMatrix(i + 1, 0, i + 1)
            If TrayList(i).APath = "" Then
                MSFlexGrid1.set_TextMatrix(i + 1, 1, "N/A")
            Else
                MSFlexGrid1.set_TextMatrix(i + 1, 1, TrayList(i).APath)
            End If
            MSFlexGrid1.set_TextMatrix(i + 1, 2, TrayList(i).uId)
            MSFlexGrid1.set_TextMatrix(i + 1, 3, TrayList(i).hwnd)
            MSFlexGrid1.set_TextMatrix(i + 1, 4, TrayList(i).hIcon)
            MSFlexGrid1.set_TextMatrix(i + 1, 5, TrayList(i).ToolTip)
            MSFlexGrid1.set_TextMatrix(i + 1, 6, "&H" & Hex$(TrayList(i).ucallbackMessage))

        Next i

    End Sub

    '==============================================================================
    ' Method:       RestoreIcon
    '
    ' Description:  Restore any icon you removed with Hide Method.
    '==============================================================================
    Private Sub RestoreIcon(ByVal TrayIconInfo As ASTC.TrayIconInfo)

        Const NIF_TIP = &H4
        Const NIF_ICON = &H2
        Const NIM_ADD = &H0
        Const NIF_MESSAGE = &H1

        Dim Res As Boolean
        Dim TrayI As NOTIFYICONDATA

        On Error Resume Next

        If TrayIconInfo.hwnd = 0 Then Exit Sub

        TrayI.cbSize = Len(TrayI)
        TrayI.hwnd = TrayIconInfo.hwnd
        TrayI.uId = TrayIconInfo.uId
        TrayI.hIcon = TrayIconInfo.hIcon
        TrayI.szTip = TrayIconInfo.ToolTip
        TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        TrayI.ucallbackMessage = TrayIconInfo.ucallbackMessage

        'Restore the icon
        Res = Shell_NotifyIcon(NIM_ADD, TrayI)

        MSFlexGrid1.Clear()

        FillFlexGrid()

    End Sub
End Class
