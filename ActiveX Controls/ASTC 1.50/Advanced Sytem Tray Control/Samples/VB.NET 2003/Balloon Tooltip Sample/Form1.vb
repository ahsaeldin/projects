'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:  A sample code demonstrates how to use Balloon Object.
'==========================================================================================

'Use the Optional parameter PHwnd to insure that the balloon tooltip will not appear outside a specified
'Window or a child window region like a command button or a text box.
'Note that:
'
'1.you can't set this parameter to any external window i.e. you can't set to an external Apps window
'
'2.you can set this parameter to the window which contains the current instance of the balloon control
'
'i.e.  For the following code
'
'balloon1.ShowBalloon From1.hwnd
'
'The hwnd is a handle to the window in which contains the Balloon control and that's means that the balloon
'Will only appear if the mouse is over the Form itself not any of its child controls [child windows] in which
'It contains.
'
'3.you can set this parameter to any control of the window that contains the current instance of the balloon
'Control i.e if you have a Form called Form1 which has a command button called Command1 and a text box called
'text1, you can set the parameter as the following
'
'balloon1.ShowBalloon command1.hwnd
'
'balloon1.ShowBalloon text1.hwnd
'
'And that's means the balloon will only appear if the mouse is over the control itself
'
'4. If you didn't pass this optional parameter, Balloon1 object will use its parent window automatically
'I.e. if you have a form called Form1, Balloon object will Form1.Hwnd for PHwnd parameter silently but in
'This case the balloon tooltip may appear at any part of the Form or its Controls
'
'Use DelayTime parameter to set the Delay Time before the balloon appear, the default value for this
'Parameter is 2000 milliseconds and the maximum, value for DelayTime Parameter is 65,535 milliseconds,
'Is equivalent to just over 1 minute?
'The maximum, value for Timeout Parameter is 65,535 milliseconds, is equivalent to just over 1 minute.
'
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ButBalloon1 As System.Windows.Forms.Button
    Friend WithEvents ButBalloon2 As System.Windows.Forms.Button
    Friend WithEvents ButBalloon3 As System.Windows.Forms.Button
    Friend WithEvents Balloon2 As AxASTC.AxBalloon
    Friend WithEvents Balloon1 As AxASTC.AxBalloon
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ButBalloon3 = New System.Windows.Forms.Button
        Me.ButBalloon2 = New System.Windows.Forms.Button
        Me.ButBalloon1 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Balloon1 = New AxASTC.AxBalloon
        Me.Balloon2 = New AxASTC.AxBalloon
        Me.GroupBox1.SuspendLayout()
        CType(Me.Balloon1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Balloon2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Balloon2)
        Me.GroupBox1.Controls.Add(Me.Balloon1)
        Me.GroupBox1.Controls.Add(Me.ButBalloon3)
        Me.GroupBox1.Controls.Add(Me.ButBalloon2)
        Me.GroupBox1.Controls.Add(Me.ButBalloon1)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(272, 112)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'ButBalloon3
        '
        Me.ButBalloon3.Location = New System.Drawing.Point(192, 56)
        Me.ButBalloon3.Name = "ButBalloon3"
        Me.ButBalloon3.TabIndex = 7
        Me.ButBalloon3.Text = "Balloon3"
        '
        'ButBalloon2
        '
        Me.ButBalloon2.Location = New System.Drawing.Point(96, 80)
        Me.ButBalloon2.Name = "ButBalloon2"
        Me.ButBalloon2.TabIndex = 6
        Me.ButBalloon2.Text = "Balloon2"
        '
        'ButBalloon1
        '
        Me.ButBalloon1.Location = New System.Drawing.Point(8, 56)
        Me.ButBalloon1.Name = "ButBalloon1"
        Me.ButBalloon1.TabIndex = 5
        Me.ButBalloon1.Text = "Balloon1"
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(32, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(208, 32)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Move the Cursor to every command button to see the balloon tooltip."
        '
        'Balloon1
        '
        Me.Balloon1.ContainingControl = Me
        Me.Balloon1.Enabled = True
        Me.Balloon1.Location = New System.Drawing.Point(8, 80)
        Me.Balloon1.Name = "Balloon1"
        Me.Balloon1.OcxState = CType(resources.GetObject("Balloon1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Balloon1.Size = New System.Drawing.Size(27, 27)
        Me.Balloon1.TabIndex = 8
        '
        'AxBalloon2
        '
        Me.Balloon2.ContainingControl = Me
        Me.Balloon2.Enabled = True
        Me.Balloon2.Location = New System.Drawing.Point(240, 16)
        Me.Balloon2.Name = "AxBalloon2"
        Me.Balloon2.OcxState = CType(resources.GetObject("AxBalloon2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Balloon2.Size = New System.Drawing.Size(27, 27)
        Me.Balloon2.TabIndex = 9
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(288, 126)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VB.Net Balloon Tooltip Sample"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.Balloon1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Balloon2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    '==============================================================================
    ' Method:        Form1_MouseMove
    '
    ' Description:   Destroy the Balloon Tooltip.
    '==============================================================================
    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
        Balloon1.Destroy()
        Balloon2.Destroy()
    End Sub
    '==============================================================================
    ' Method:        GroupBox1_MouseMove
    '
    ' Description:   Destroy the Balloon Tooltip.
    '==============================================================================
    Private Sub GroupBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GroupBox1.MouseMove
        Balloon1.Destroy()
        Balloon2.Destroy()
    End Sub
    '==============================================================================
    ' Method:        Label1_MouseMove
    '
    ' Description:   Destroy the Balloon Tooltip.
    '==============================================================================
    Private Sub Label1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label1.MouseMove
        Balloon1.Destroy()
        Balloon2.Destroy()
    End Sub
    '==============================================================================
    ' Method:        ButBalloon1_MouseMove
    '
    ' Description:  Display a balloon tooltip for ButBalloon1.
    '==============================================================================
    Private Sub ButBalloon1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ButBalloon1.MouseMove
        Balloon1.Style = ASTC.BalloonStyle.BalloonType
        Balloon1.ShowBalloon(ButBalloon1.Handle.ToInt32(), "Click Me", "Balloon1", ASTC.BIcoType.Info, 1000)
    End Sub
    '==============================================================================
    ' Method:        ButBalloon2_MouseMove
    '
    ' Description:   Display a balloon tooltip for ButBalloon2.
    '==============================================================================
    Private Sub ButBalloon2_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ButBalloon2.MouseMove
        Balloon1.Style = ASTC.BalloonStyle.RectangleType
        Balloon1.ShowBalloon(ButBalloon2.Handle.ToInt32(), "Rectangle Type", "Balloon2", ASTC.BIcoType.Info, 1000)
    End Sub
    '==============================================================================
    ' Method:        ButBalloon3_MouseMove
    '
    ' Description:   Display a balloon tooltip for ButBalloon3.
    '==============================================================================
    Private Sub ButBalloon3_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ButBalloon3.MouseMove
        Balloon2.CtlBackColor = RGB(255, 255, 255)
        Balloon2.CtlForeColor = RGB(255, 50, 150)
        Balloon2.ShowBalloon(ButBalloon3.Handle.ToInt32(), "Customizable balloon tooltip", "Balloon3", ASTC.BIcoType.Info, 1000)
    End Sub
    '==============================================================================
    ' Method:        Balloon1_BalloonLeftClick
    '
    ' Description:   Display a message Box when the left click the balloon.
    '==============================================================================
    Private Sub Balloon1_BalloonLeftClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Balloon1.BalloonLeftClick
        MsgBox("The Balloon Tooltip had been Clicked")
    End Sub

End Class
