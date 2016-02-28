
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
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Label3 As System.Windows.Forms.Label

    Friend WithEvents ButLSubClass As System.Windows.Forms.Button
    Friend WithEvents ButLUnSubClass As System.Windows.Forms.Button
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ButUnSubCont As System.Windows.Forms.Button
    Friend WithEvents ButSubContext As System.Windows.Forms.Button
    Friend WithEvents SubClass2 As AxASTC.AxSubClass
    Friend WithEvents SubClass1 As AxASTC.AxSubClass

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.ButLUnSubClass = New System.Windows.Forms.Button
        Me.ButLSubClass = New System.Windows.Forms.Button
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.ButUnSubCont = New System.Windows.Forms.Button
        Me.ButSubContext = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.SubClass2 = New AxASTC.AxSubClass
        Me.SubClass1 = New AxASTC.AxSubClass
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.SubClass2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SubClass1, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.GroupBox1.ForeColor = System.Drawing.Color.Blue
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(272, 248)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "SubClass Sample 1"
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Red
        Me.Label3.Location = New System.Drawing.Point(8, 184)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(256, 32)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Then click SubClass command and Right Click any item and you will see the differe" & _
        "nce."
        '
        'ListBox1
        '
        Me.ListBox1.Items.AddRange(New Object() {"item1", "item2", "item3", "item4"})
        Me.ListBox1.Location = New System.Drawing.Point(8, 128)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(256, 56)
        Me.ListBox1.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Red
        Me.Label2.Location = New System.Drawing.Point(8, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(256, 32)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Right Click any item in listbox before you click the Subclass command button."
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(256, 72)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "In the Normal ListBox you can't assign a popup menu to any item so when the user " & _
        "right click the item it will appears                                            " & _
        "        But if you we use Subclassing we can achieve this."
        '
        'ButLUnSubClass
        '
        Me.ButLUnSubClass.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.ButLUnSubClass.Location = New System.Drawing.Point(144, 216)
        Me.ButLUnSubClass.Name = "ButLUnSubClass"
        Me.ButLUnSubClass.TabIndex = 2
        Me.ButLUnSubClass.Text = "UnSubClass"
        '
        'ButLSubClass
        '
        Me.ButLSubClass.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.ButLSubClass.Location = New System.Drawing.Point(56, 216)
        Me.ButLSubClass.Name = "ButLSubClass"
        Me.ButLSubClass.TabIndex = 1
        Me.ButLSubClass.Text = "SubClass"
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
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.SubClass2)
        Me.GroupBox2.Controls.Add(Me.ButUnSubCont)
        Me.GroupBox2.Controls.Add(Me.ButSubContext)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.TextBox1)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.ForeColor = System.Drawing.Color.Blue
        Me.GroupBox2.Location = New System.Drawing.Point(8, 264)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(272, 208)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "SubClass Sample 2"
        '
        'ButUnSubCont
        '
        Me.ButUnSubCont.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.ButUnSubCont.Location = New System.Drawing.Point(144, 176)
        Me.ButUnSubCont.Name = "ButUnSubCont"
        Me.ButUnSubCont.TabIndex = 11
        Me.ButUnSubCont.Text = "UnSubClass"
        '
        'ButSubContext
        '
        Me.ButSubContext.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.ButSubContext.Location = New System.Drawing.Point(56, 176)
        Me.ButSubContext.Name = "ButSubContext"
        Me.ButSubContext.TabIndex = 10
        Me.ButSubContext.Text = "SubClass"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.White
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.Location = New System.Drawing.Point(56, 136)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(160, 32)
        Me.Label5.TabIndex = 2
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(56, 64)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(160, 64)
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Text = "Right Click me before and after you click Subclass command  button and watch the " & _
        "difference."
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.Label4.Location = New System.Drawing.Point(8, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(256, 40)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "You can use Subclassing to disable the Textbox context menu in which it appears w" & _
        "hen the user right click the Text box."
        '
        'SubClass2
        '
        Me.SubClass2.ContainingControl = Me
        Me.SubClass2.Enabled = True
        Me.SubClass2.Location = New System.Drawing.Point(232, 168)
        Me.SubClass2.Name = "SubClass2"
        Me.SubClass2.OcxState = CType(resources.GetObject("SubClass2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.SubClass2.Size = New System.Drawing.Size(36, 36)
        Me.SubClass2.TabIndex = 12
        '
        'SubClass1
        '
        Me.SubClass1.ContainingControl = Me
        Me.SubClass1.Enabled = True
        Me.SubClass1.Location = New System.Drawing.Point(232, 208)
        Me.SubClass1.Name = "SubClass1"
        Me.SubClass1.OcxState = CType(resources.GetObject("SubClass1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.SubClass1.Size = New System.Drawing.Size(36, 36)
        Me.SubClass1.TabIndex = 6
        '
        'Form1
        '
        Me.AutoScale = False
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(288, 478)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VB.Net SubClass Sample"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.SubClass2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SubClass1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

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
    'How to stop the form from being resized below or above a user-defined amount?

    'That's it
    'If you have any further questions or need more sample code don't hesitate to contact us
    'At support@cpringold.com


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
        Const WM_RBUTTONUP = 517
        Const LB_ITEMFROMPOINT = 425

        If e.msg = WM_RBUTTONUP Then
            'Retrieve the zero-based index of the item nearest a specified point in the list box.
            ret = SendMessage(e.hwnd, LB_ITEMFROMPOINT, 0&, e.lParam)
            For i = 0 To ListBox1.Items.Count - 1
                If ret = i Then
                    ListBox1.SetSelected(i, True)
                    Label3.Text = ListBox1.Items.Item(i) & " had been clicked"
                    Pos.X = Cursor.Position.X
                    Pos.Y = Cursor.Position.Y
                    'Extends Listbox functionality by Display a popup menu related to the clicked item.
                    TrackPopupMenu(ContextMenu1.Handle.ToInt32(), &H0&, Pos.X, Pos.Y, 0, Handle.ToInt32(), 0&)

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

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        MsgBox(ListBox1.Items.Item(ListBox1.SelectedIndex) & " had been clicked")
    End Sub
    '==============================================================================
    ' Method:        ButSubContext_Click
    '
    ' Description:  Starts subclassing TextBox1.
    '==============================================================================
    Private Sub ButSubContext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButSubContext.Click
        SubClass2.BeginSubClass(TextBox1.Handle.ToInt32())
    End Sub
    '==============================================================================
    ' Method:        SubClass2_MessageReceived
    '
    ' Description:   Disable a context menu in a text box.
    '==============================================================================
    Private Sub SubClass2_MessageReceived(ByVal sender As System.Object, ByVal e As AxASTC.__SubClass_MessageReceivedEvent) Handles SubClass2.MessageReceived

        'e.hwnd()        Handle of the subclassed window or control.
        'e.msg()         The ID of the intercepted message.
        'e.wParam()      The wParam value of the intercepted message.
        'e.lParam()      The lParam value of the intercepted message.
        'e.cancel        Further processing state.

        Const WM_CONTEXTMENU = 123

        If e.msg = WM_CONTEXTMENU Then
            Label5.Text = "Context Menu had been disabled."
            'Cancel any further processing.
            e.cancel = True
        End If

    End Sub
    '==============================================================================
    ' Method:        ButUnSubCont_Click
    '
    ' Description:  Stop Subclassing TextBox1.
    '==============================================================================
    Private Sub ButUnSubCont_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButUnSubCont.Click
        SubClass2.EndSubClass()
        Label5.Text = ""
    End Sub

End Class
