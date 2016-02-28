'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:  A sample code demonstrate how to add your software icon to system tray area
'               As well as how to use the advanced features related to system tray area.
'==========================================================================================
'ASTC provide you with a very simple object, TrayIcon object which is the responsible to add the most of system
'Tray functions to your applications.
'
'With a few lines i can demonstrate how TrayIcon object can help you
'
'Use AxTrayIcon1.Show method to add an icon to system tray area
'Use AxTrayIcon1.Hide method to hide the icon you had add by TrayIcon.Show method from system tray
'Use AxTrayIcon1. ChangeIcon method to change the icon you had added by another icon
'Use AxTrayIcon1.Animate method to animate the icon in the system tray
'Use AxTrayIcon1.StopAnimateing method to stop animating the icon
'Use AxTrayIcon1.ShowBalloon to display a tooltip balloon
'Use AxTrayIcon1.HideBalloon to hide the balloon
'Use AxTrayIcon1.Popup to display a popup menu
'Use AxTrayIcon1.TrackPopupmenu to track the popup menu whenever it lost focus in order to close it.
'Use AxTrayIcon1.IsDisplayed to check if the icon is displayed in the system tray area
'
'Note that you have a custom types of events associated with TrayIcon Object like
'BalloonShow, BalloonHide, MedMouseUp, LeftMouseUp, RightMouseUp, MedMouseDown, LeftMouseDown
'RightMouseDown, MedMouseDBLCLK, LeftMouseDBLCLK, BalloonLeftClick, BalloonRightClick, RightMouseDBLCLK ()
'
'Note that TrayIcon Object doesn't subclass the window in which it contains TrayIcon control
'because TrayIcon Object dynamically creates an internal window related to it's instance,
'and destroy this window whenever you call Hide method or close your application,
'so don 't worry about subclassing
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
    Friend WithEvents groupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents AxImageList1 As AxMSComctlLib.AxImageList
    Friend WithEvents butSTrack As System.Windows.Forms.Button
    Friend WithEvents butTrack As System.Windows.Forms.Button
    Friend WithEvents butCBalloon As System.Windows.Forms.Button
    Friend WithEvents butShowBalloon As System.Windows.Forms.Button
    Friend WithEvents butStAni As System.Windows.Forms.Button
    Friend WithEvents butAni As System.Windows.Forms.Button
    Friend WithEvents butHide As System.Windows.Forms.Button
    Friend WithEvents butChgIcon As System.Windows.Forms.Button
    Friend WithEvents butShow As System.Windows.Forms.Button

    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents AxTrayIcon1 As AxASTC.AxTrayIcon
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.groupBox1 = New System.Windows.Forms.GroupBox
        Me.AxImageList1 = New AxMSComctlLib.AxImageList
        Me.butSTrack = New System.Windows.Forms.Button
        Me.butTrack = New System.Windows.Forms.Button
        Me.butCBalloon = New System.Windows.Forms.Button
        Me.butShowBalloon = New System.Windows.Forms.Button
        Me.butStAni = New System.Windows.Forms.Button
        Me.butAni = New System.Windows.Forms.Button
        Me.butHide = New System.Windows.Forms.Button
        Me.butChgIcon = New System.Windows.Forms.Button
        Me.butShow = New System.Windows.Forms.Button
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.AxTrayIcon1 = New AxASTC.AxTrayIcon
        Me.groupBox1.SuspendLayout()
        CType(Me.AxImageList1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxTrayIcon1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'groupBox1
        '
        Me.groupBox1.Controls.Add(Me.AxTrayIcon1)
        Me.groupBox1.Controls.Add(Me.AxImageList1)
        Me.groupBox1.Controls.Add(Me.butSTrack)
        Me.groupBox1.Controls.Add(Me.butTrack)
        Me.groupBox1.Controls.Add(Me.butCBalloon)
        Me.groupBox1.Controls.Add(Me.butShowBalloon)
        Me.groupBox1.Controls.Add(Me.butStAni)
        Me.groupBox1.Controls.Add(Me.butAni)
        Me.groupBox1.Controls.Add(Me.butHide)
        Me.groupBox1.Controls.Add(Me.butChgIcon)
        Me.groupBox1.Controls.Add(Me.butShow)
        Me.groupBox1.Location = New System.Drawing.Point(8, 8)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(216, 176)
        Me.groupBox1.TabIndex = 6
        Me.groupBox1.TabStop = False
        '
        'AxImageList1
        '
        Me.AxImageList1.ContainingControl = Me
        Me.AxImageList1.Enabled = True
        Me.AxImageList1.Location = New System.Drawing.Point(88, 112)
        Me.AxImageList1.Name = "AxImageList1"
        Me.AxImageList1.OcxState = CType(resources.GetObject("AxImageList1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImageList1.Size = New System.Drawing.Size(38, 38)
        Me.AxImageList1.TabIndex = 18
        '
        'butSTrack
        '
        Me.butSTrack.Enabled = False
        Me.butSTrack.Location = New System.Drawing.Point(112, 112)
        Me.butSTrack.Name = "butSTrack"
        Me.butSTrack.Size = New System.Drawing.Size(96, 24)
        Me.butSTrack.TabIndex = 14
        Me.butSTrack.Text = "Stop Track"
        '
        'butTrack
        '
        Me.butTrack.Location = New System.Drawing.Point(112, 80)
        Me.butTrack.Name = "butTrack"
        Me.butTrack.Size = New System.Drawing.Size(96, 24)
        Me.butTrack.TabIndex = 13
        Me.butTrack.Text = "TrackPopUp"
        '
        'butCBalloon
        '
        Me.butCBalloon.Location = New System.Drawing.Point(112, 48)
        Me.butCBalloon.Name = "butCBalloon"
        Me.butCBalloon.Size = New System.Drawing.Size(96, 24)
        Me.butCBalloon.TabIndex = 12
        Me.butCBalloon.Text = "Close Balloon"
        '
        'butShowBalloon
        '
        Me.butShowBalloon.Location = New System.Drawing.Point(112, 16)
        Me.butShowBalloon.Name = "butShowBalloon"
        Me.butShowBalloon.Size = New System.Drawing.Size(96, 24)
        Me.butShowBalloon.TabIndex = 11
        Me.butShowBalloon.Text = "Show Balloon"
        '
        'butStAni
        '
        Me.butStAni.Location = New System.Drawing.Point(56, 144)
        Me.butStAni.Name = "butStAni"
        Me.butStAni.Size = New System.Drawing.Size(104, 24)
        Me.butStAni.TabIndex = 9
        Me.butStAni.Text = "Stop Animation"
        '
        'butAni
        '
        Me.butAni.Location = New System.Drawing.Point(8, 112)
        Me.butAni.Name = "butAni"
        Me.butAni.Size = New System.Drawing.Size(96, 24)
        Me.butAni.TabIndex = 7
        Me.butAni.Text = "Animate"
        '
        'butHide
        '
        Me.butHide.Location = New System.Drawing.Point(8, 80)
        Me.butHide.Name = "butHide"
        Me.butHide.Size = New System.Drawing.Size(96, 24)
        Me.butHide.TabIndex = 6
        Me.butHide.Text = "Hide Icon"
        '
        'butChgIcon
        '
        Me.butChgIcon.Location = New System.Drawing.Point(8, 48)
        Me.butChgIcon.Name = "butChgIcon"
        Me.butChgIcon.Size = New System.Drawing.Size(96, 24)
        Me.butChgIcon.TabIndex = 4
        Me.butChgIcon.Text = "Change Icon"
        '
        'butShow
        '
        Me.butShow.Location = New System.Drawing.Point(8, 16)
        Me.butShow.Name = "butShow"
        Me.butShow.Size = New System.Drawing.Size(96, 24)
        Me.butShow.TabIndex = 2
        Me.butShow.Text = "&Show Icon"
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "main"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "Exit"
        '
        'AxTrayIcon1
        '
        Me.AxTrayIcon1.ContainingControl = Me
        Me.AxTrayIcon1.Enabled = True
        Me.AxTrayIcon1.Location = New System.Drawing.Point(88, 72)
        Me.AxTrayIcon1.Name = "AxTrayIcon1"
        Me.AxTrayIcon1.OcxState = CType(resources.GetObject("AxTrayIcon1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxTrayIcon1.Size = New System.Drawing.Size(36, 36)
        Me.AxTrayIcon1.TabIndex = 19
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(232, 190)
        Me.Controls.Add(Me.groupBox1)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VB.NET TrayIcon"
        Me.groupBox1.ResumeLayout(False)
        CType(Me.AxImageList1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxTrayIcon1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    '==============================================================================
    ' Method:        butShow_Click
    '
    ' Description:  Shows the icon in system tray.
    '==============================================================================
    Private Sub butShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butShow.Click
        'Unlike the previous version of ASTC, you don't need to pass a window handle to Show function because
        'ASTC dynamically creates an internal window related to ASTC instance, and destroy this window whenever
        'you call Hide method or close your application.
        AxTrayIcon1().CtlShow(Me.Icon.Handle.ToInt32(), "VB.netTrayIcon")
    End Sub
    '==============================================================================
    ' Method:        butChgIcon_Click
    '
    ' Description:  Changes the icon in system tray area.
    '==============================================================================
    Private Sub butChgIcon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butChgIcon.Click
        'change the icon in system tray area
        'by another one in ImageList    
        AxTrayIcon1.ChangeIcon(AxImageList1.ListImages.Item(2).ExtractIcon().Handle)
    End Sub
    '==============================================================================
    ' Method:       butHide_Click
    '
    ' Description:  Removes the icon from system tray.
    '==============================================================================
    Private Sub butHide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butHide.Click
        'Hide the icon Only if it was displayed
        If AxTrayIcon1.IsDisplayed Then AxTrayIcon1.CtlHide()
    End Sub
    '==================================================================================
    ' Method:        butAni_Click
    '
    ' Description:  Animate the icon in system tray using in icons in imagelist control.
    '==================================================================================
    Private Sub butAni_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butAni.Click
        'only if the icon is displayed
        If AxTrayIcon1.IsDisplayed Then
            'Only if the icon is not Animated   
            If AxTrayIcon1.AnimateState = False Then
                'Start animating the icons in system tray
                'we use the icons in the imagelist control
                'Then
                'We pass the ImageList1 control to Animate method
                '''you can replace the icons with your own'''
                AxTrayIcon1.Animate(AxImageList1, 1)
            End If
        End If
    End Sub
    '==============================================================================
    ' Method:       butStAni_Click
    '
    ' Description:  Stop animating the icons in system tray.
    '==============================================================================
    Private Sub butStAni_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butStAni.Click
        'Only if the icon is animated 
        If AxTrayIcon1.AnimateState = True Then
            AxTrayIcon1.StopAnimateing(Me.Icon.Handle.ToInt32())
        End If
    End Sub
    '==============================================================================
    ' Method:       butShowBalloon_Click
    '
    ' Description:  Displays a balloon tooltip points to the icon.
    '==============================================================================
    Private Sub butShowBalloon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butShowBalloon.Click
        If AxTrayIcon1.AnimateState = True Then MsgBox("You can't Display a balloon while animating the icon.")
        AxTrayIcon1.ShowBalloon("Click me", "VB.net TrayIcon", ASTC.BIcoType.Info, 5000)
    End Sub
    '==============================================================================
    ' Method:       butCBalloon_Click
    '
    ' Description:  Close the balloon tooltip.
    '==============================================================================
    Private Sub butCBalloon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butCBalloon.Click
        AxTrayIcon1.CloseBalloon()
    End Sub
    '==============================================================================
    ' Method:       butTrack_Click
    '
    ' Description:  Start Tracking the popup menu in system tray area.
    '==============================================================================
    Private Sub butTrack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butTrack.Click
        MsgBox("when you right click the icon a popupmenu will appear,ASTC will Track the popupmenu and close it when you left click again")
        'Track the popup menu so ASTC can destroy it whenever it lost focus   
        AxTrayIcon1.TrackPopMenu = True
        butSTrack.Enabled = True
    End Sub
    '==============================================================================
    ' Method:       butSTrack_Click
    '
    ' Description:  Stop track the popup menu in system tray area.
    '==============================================================================
    Private Sub butSTrack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butSTrack.Click
        MsgBox("ASTC stop tracking the popupmenu,Now right click the icon again and you will see that popupmenu willn't close unless you click an item.")
        'stop tracking the popup menu
        AxTrayIcon1.TrackPopMenu = False
        butSTrack.Enabled = False
    End Sub
    '==============================================================================
    ' Method:       Form1_Closing
    '
    ' Description:  Removes the icon from system tray whenever the
    '               terminates the program.
    '==============================================================================
    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If AxTrayIcon1.IsDisplayed Then AxTrayIcon1.CtlHide()
    End Sub
    '==============================================================================
    ' Method:        MenuItem2_Click
    '
    ' Description:  Removes the icon from system tray.
    '==============================================================================
    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        If AxTrayIcon1.IsDisplayed Then AxTrayIcon1.CtlHide()
        End
    End Sub
    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        MsgBox("main menu Item had been clicked.")
    End Sub
    '====================================================================================
    ' Method:       AxTrayIcon1_RightMouseUp
    '
    ' Description:  Display a popup menu whenever you right click the icon in system tray.
    '====================================================================================
    Private Sub AxTrayIcon1_RightMouseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles AxTrayIcon1.RightMouseUp
        AxTrayIcon1.PopUp(ContextMenu1.Handle.ToInt32(), Handle.ToInt32())
    End Sub
    '==============================================================================
    ' Method:       AxTrayIcon1_BalloonLeftClick
    '
    ' Description:  Display a message box whenever you left click the balloon tooltip.
    '==============================================================================
    Private Sub AxTrayIcon1_BalloonLeftClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles AxTrayIcon1.BalloonLeftClick
        MsgBox("The Balloon Tooltip had been left clicked.")
    End Sub
End Class
