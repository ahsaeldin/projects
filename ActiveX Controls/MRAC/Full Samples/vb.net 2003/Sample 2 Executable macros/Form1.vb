'====================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:  A sample code demonstrates how to use MacroRecorder
'               ActiveX control v1.50 to record and replay mouse clicks, keystrokes
'               and bundle them into an executble file [.exe] in order to replay later.
'=====================================================================================
Public Class Form1
    Inherits System.Windows.Forms.Form

    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

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
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents RecordBut As System.Windows.Forms.Button
    Friend WithEvents butSaveExe As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents AxMacroRecorder1 As AxMacroRecorderActX.AxMacroRecorder
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.butSaveExe = New System.Windows.Forms.Button
        Me.RecordBut = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.AxMacroRecorder1 = New AxMacroRecorderActX.AxMacroRecorder
        Me.GroupBox1.SuspendLayout()
        CType(Me.AxMacroRecorder1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.butSaveExe)
        Me.GroupBox1.Controls.Add(Me.RecordBut)
        Me.GroupBox1.Controls.Add(Me.AxMacroRecorder1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(488, 128)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(456, 24)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "2.Click Save as Exe button to stop recording and save the macro as an exectutable" & _
        " file."
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(208, 16)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "1. Click Record button to start recording."
        '
        'butSaveExe
        '
        Me.butSaveExe.Location = New System.Drawing.Point(8, 96)
        Me.butSaveExe.Name = "butSaveExe"
        Me.butSaveExe.Size = New System.Drawing.Size(88, 24)
        Me.butSaveExe.TabIndex = 2
        Me.butSaveExe.Text = "&Save As Exe"
        '
        'RecordBut
        '
        Me.RecordBut.Location = New System.Drawing.Point(8, 40)
        Me.RecordBut.Name = "RecordBut"
        Me.RecordBut.Size = New System.Drawing.Size(88, 24)
        Me.RecordBut.TabIndex = 0
        Me.RecordBut.Text = "&Record"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        '
        'AxMacroRecorder1
        '
        Me.AxMacroRecorder1.ContainingControl = Me
        Me.AxMacroRecorder1.Enabled = True
        Me.AxMacroRecorder1.Location = New System.Drawing.Point(448, 16)
        Me.AxMacroRecorder1.Name = "AxMacroRecorder1"
        Me.AxMacroRecorder1.OcxState = CType(resources.GetObject("AxMacroRecorder1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMacroRecorder1.Size = New System.Drawing.Size(35, 35)
        Me.AxMacroRecorder1.TabIndex = 1
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(504, 134)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VB.Net Macro Recorder"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.AxMacroRecorder1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    '==============================================================================
    ' Method:        RecordBut_Click
    '
    ' Description:  Start recording a new macro.
    '==============================================================================
    Private Sub RecordBut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RecordBut.Click
        If Me.WindowState <> 1 Then Me.WindowState = 1
        AxMacroRecorder1.Record()
    End Sub

    Private Sub butSaveExe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butSaveExe.Click
        StopRecord()
    End Sub

    Private Sub StopRecord()
        'Only if State is recording.
        If Not AxMacroRecorder1.IsRecord Then
            Exit Sub
        End If

        Me.WindowState = 0

        AxMacroRecorder1.StopRecord()

        With SaveFileDialog1
            .ShowDialog()
        End With

        'Note that
        'Macro Recorder ActiveX Control will add the ".exe" extension
        'to the file path automatically
        'For example suppose that you want to save the macro to c:\
        'you can use one of the following
        'MacroRecorder1.SaveAsExE ("c:\macro.exe")
        'MacroRecorder1.SaveAsExE ("c:\macro")     'Macro Recorder will add the .exe extension to the file

        AxMacroRecorder1.SaveAsExE(SaveFileDialog1.FileName)

        'You can use command line parameters for the executable macro file to change the replay speed
        'For example if you save an executable macro to macro1.exe
        'macro1 /h    for high replay speed.
        'macro1 /n    for normal replay speed "the default one"
        'macro1 /l    for low replay speed.

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Stop Recording when you Click Alt+F10
        If GetAsyncKeyState(121) And GetAsyncKeyState(18) Then
            StopRecord()
        End If
    End Sub

    Private Sub AxMacroRecorder1_RecordStart(ByVal sender As Object, ByVal e As System.EventArgs) Handles AxMacroRecorder1.RecordStart
        MsgBox("Press Alt+F10 to Stop Recording.")
    End Sub

    Private Sub AxMacroRecorder1_SaveComplete(ByVal sender As Object, ByVal e As AxMacroRecorderActX.__MacroRecorder_SaveCompleteEvent) Handles AxMacroRecorder1.SaveComplete
        If e.macFilePath <> "" Then MsgBox("Macro had been saved To " & e.macFilePath)
    End Sub
End Class
