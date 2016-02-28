
'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
'Description:  A sample code demonstrates how to use ASTC control to subclass any Form.
'==========================================================================================
'First of all what is the Subclass?
'Subclassing is the processing of intercepting Windows messages that your program normally wouldn't
'Receive, so it extends your ability to process more windows messages and add new features to your program.
''
'With a few lines i can demonstrate how Subclass object can help you
'
'1.use SubClass.BeginSubClass to start subclassing a specified window.
'2.use SubClass.EndSubClass to end subclassing.
'3.use SubClass_MessageReceived event to process the message you want.
'
'About this sample code
'This sample code show you the benefits of subclassing by demonstrate
'How to disable the context menu in a text box?
'How to extend a listbox functions?

'That's it
'If you have any further questions or need more sample code don't hesitate to contact us
'At support@cpringold.com
Imports System.Runtime.InteropServices

Public Class Form1

    Inherits System.Windows.Forms.Form

#Region " Declares"

    Private Structure POINTAPI
        Dim X As Long
        Dim Y As Long
    End Structure

    Private Structure MINMAXINFO
        Dim ptReserved As POINTAPI
        Dim ptMaxSize As POINTAPI
        Dim ptMaxPosition As POINTAPI
        Dim ptMinTrackSize As POINTAPI
        Dim ptMaxTrackSize As POINTAPI
    End Structure
    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
                                             ByVal hwnd As Int32, _
                                             ByVal wMsg As Int32, _
                                             ByVal wParam As Int32, _
                                             ByVal lParam As Int32) As Int32

    Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Int32, ByVal _
    wFlags As Int32, ByVal x As Int32, ByVal y As Int32, ByVal nReserved As Int32, ByVal hwnd _
    As Int32, ByVal lprc As Int32) As Int32

    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Int32, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters _
    As String, ByVal lpDirectory As String, ByVal nShowCmd As Int32) As Int32

    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDst As MINMAXINFO, _
    ByVal pSrc As Integer, ByVal ByteLen As Integer)

    Private Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDst As Integer, _
    ByRef pSrc As MINMAXINFO, ByVal ByteLen As Integer)

#End Region


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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label

    Friend WithEvents ButLSubClass As System.Windows.Forms.Button
    Friend WithEvents ButLUnSubClass As System.Windows.Forms.Button
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents SubClass1 As AxASTC.AxSubClass
    Friend WithEvents Balloon1 As AxASTC.AxBalloon
    Friend WithEvents Balloon2 As AxASTC.AxBalloon
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.SubClass1 = New AxASTC.AxSubClass
        Me.Label3 = New System.Windows.Forms.Label
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.ButLUnSubClass = New System.Windows.Forms.Button
        Me.ButLSubClass = New System.Windows.Forms.Button
        Me.Balloon2 = New AxASTC.AxBalloon
        Me.Balloon1 = New AxASTC.AxBalloon
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.GroupBox1.SuspendLayout()
        CType(Me.SubClass1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Balloon2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Balloon1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.SubClass1)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.ListBox1)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.ButLUnSubClass)
        Me.GroupBox1.Controls.Add(Me.ButLSubClass)
        Me.GroupBox1.Controls.Add(Me.Balloon2)
        Me.GroupBox1.Controls.Add(Me.Balloon1)
        Me.GroupBox1.ForeColor = System.Drawing.Color.Blue
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(408, 232)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'SubClass1
        '
        Me.SubClass1.ContainingControl = Me
        Me.SubClass1.Enabled = True
        Me.SubClass1.Location = New System.Drawing.Point(8, 192)
        Me.SubClass1.Name = "SubClass1"
        Me.SubClass1.OcxState = CType(resources.GetObject("SubClass1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.SubClass1.Size = New System.Drawing.Size(36, 36)
        Me.SubClass1.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Red
        Me.Label3.Location = New System.Drawing.Point(40, 160)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(256, 32)
        Me.Label3.TabIndex = 5
        '
        'ListBox1
        '
        Me.ListBox1.Items.AddRange(New Object() {"item1", "item2", "item3", "item4"})
        Me.ListBox1.Location = New System.Drawing.Point(40, 96)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(320, 56)
        Me.ListBox1.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Red
        Me.Label2.Location = New System.Drawing.Point(40, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(320, 32)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Move the Cursor between listBox items before and after you click SubClass command" & _
        " Button and watch the difference."
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(392, 40)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "ListBox control doesn't provides a mousemove event in which you can use for displ" & _
        "aying a balloon tooltip for the list Box, but if you we subclass the listbox we " & _
        "can assign a Balloon tooltip for every item in the listbox."
        '
        'ButLUnSubClass
        '
        Me.ButLUnSubClass.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.ButLUnSubClass.Location = New System.Drawing.Point(208, 200)
        Me.ButLUnSubClass.Name = "ButLUnSubClass"
        Me.ButLUnSubClass.TabIndex = 2
        Me.ButLUnSubClass.Text = "UnSubClass"
        '
        'ButLSubClass
        '
        Me.ButLSubClass.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.ButLSubClass.Location = New System.Drawing.Point(128, 200)
        Me.ButLSubClass.Name = "ButLSubClass"
        Me.ButLSubClass.TabIndex = 1
        Me.ButLSubClass.Text = "SubClass"
        '
        'Balloon2
        '
        Me.Balloon2.ContainingControl = Me
        Me.Balloon2.Enabled = True
        Me.Balloon2.Location = New System.Drawing.Point(376, 168)
        Me.Balloon2.Name = "Balloon2"
        Me.Balloon2.OcxState = CType(resources.GetObject("Balloon2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Balloon2.Size = New System.Drawing.Size(27, 27)
        Me.Balloon2.TabIndex = 2
        '
        'Balloon1
        '
        Me.Balloon1.ContainingControl = Me
        Me.Balloon1.Enabled = True
        Me.Balloon1.Location = New System.Drawing.Point(376, 200)
        Me.Balloon1.Name = "Balloon1"
        Me.Balloon1.OcxState = CType(resources.GetObject("Balloon1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Balloon1.Size = New System.Drawing.Size(27, 27)
        Me.Balloon1.TabIndex = 1
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Click Me"
        '
        'Form1
        '
        Me.AutoScale = False
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(424, 246)
        Me.Controls.Add(Me.GroupBox1)
        Me.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VB.Net SubClass Sample"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.SubClass1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Balloon2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Balloon1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    '==============================================================================
    ' Method:        ButLSubClass_Click
    '
    ' Description:  Start subclassing List1.
    '==============================================================================
    Private Sub ButLSubClass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButLSubClass.Click
        SubClass1.BeginSubClass(ListBox1.Handle.ToInt32())
    End Sub
    '==============================================================================
    ' Method:        SubClass1_MessageReceived
    '
    ' Description:   Extends the listbox functionality.
    '==============================================================================
    Private Sub SubClass1_MessageReceived(ByVal sender As System.Object, ByVal e As AxASTC.__SubClass_MessageReceivedEvent) Handles SubClass1.MessageReceived

        'e.hwnd()        Handle of the subclassed window or control.
        'e.msg()         The ID of the intercepted message.
        'e.wParam()      The wParam value of the intercepted message.
        'e.lParam()      The lParam value of the intercepted message.
        'e.cancel        Further processing state.

        Dim ret As Long
        Dim Pos As Point
        Dim i As Integer

        Static OldItem As Integer

        Const WM_MOUSEMOVE = 512
        Const LB_ITEMFROMPOINT = 425

        If e.msg = WM_MOUSEMOVE Then
            'Retrieve the zero-based index of the item nearest a specified point in the list box.
            ret = SendMessage(e.hwnd, LB_ITEMFROMPOINT, 0&, e.lParam)
            For i = 0 To ListBox1.Items.Count - 1
                If ret = i Then
                    If OldItem <> i Then
                        Balloon1.Destroy()
                        Balloon2.Destroy()
                    End If
                    ListBox1.SetSelected(i, True)
                    Label3.Text = ListBox1.Items.Item(i) & " had been hovered."
                    'Extends Listbox functionality
                    If ret = 3 Then
                        Balloon2.ShowBalloon(ListBox1.Handle.ToInt32(), "Click Me.", "SubClass", ASTC.BIcoType.Info, 500, 5000)
                    Else
                        Balloon1.ShowBalloon(ListBox1.Handle.ToInt32(), ListBox1.Items(i) & " had been hovered.", "SubClass", ASTC.BIcoType.Info, 500, 5000)
                    End If
                    OldItem = i
                End If
            Next i
        End If

    End Sub
    '==============================================================================
    ' Method:        ButLUnSubClass_Click
    '
    ' Description:  Stop Subclassing List1.
    '==============================================================================
    Private Sub ButLUnSubClass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButLUnSubClass.Click
        SubClass1.EndSubClass()
        Label3.Text = ""
    End Sub

    Private Sub Balloon2_BalloonLeftClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Balloon2.BalloonLeftClick
        MsgBox(ListBox1.Items(ListBox1.SelectedIndex()) & " had been hovered.")
    End Sub
End Class
