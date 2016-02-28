'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:  A sample code demonstrates how to use MacroRecorder
'               ActiveX control v1.50 to record and replay mouse clicks, keystrokes
'               and bundle them into a file [macro] in order to replay later.
'==========================================================================================
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
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ComSpeed As System.Windows.Forms.ComboBox
    Friend WithEvents PlaybackBut As System.Windows.Forms.Button
    Friend WithEvents StopRecordBut As System.Windows.Forms.Button
    Friend WithEvents RecordBut As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents AxMacroRecorder1 As AxMacroRecorderActX.AxMacroRecorder
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.ComSpeed = New System.Windows.Forms.ComboBox
        Me.PlaybackBut = New System.Windows.Forms.Button
        Me.StopRecordBut = New System.Windows.Forms.Button
        Me.RecordBut = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.AxMacroRecorder1 = New AxMacroRecorderActX.AxMacroRecorder
        Me.GroupBox1.SuspendLayout()
        CType(Me.AxMacroRecorder1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.ComSpeed)
        Me.GroupBox1.Controls.Add(Me.PlaybackBut)
        Me.GroupBox1.Controls.Add(Me.StopRecordBut)
        Me.GroupBox1.Controls.Add(Me.RecordBut)
        Me.GroupBox1.Controls.Add(Me.AxMacroRecorder1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(408, 192)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(4, 136)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(392, 16)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Click Replay button to Replay the macro you had saved"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(392, 24)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "2. Click Stop Record button to stop recording and save the macro to the disk."
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(208, 16)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "1. Click Record button to start recording."
        '
        'ComSpeed
        '
        Me.ComSpeed.Items.AddRange(New Object() {"High Speed", "Normal Speed", "Low Speed"})
        Me.ComSpeed.Location = New System.Drawing.Point(104, 160)
        Me.ComSpeed.Name = "ComSpeed"
        Me.ComSpeed.Size = New System.Drawing.Size(104, 21)
        Me.ComSpeed.TabIndex = 4
        Me.ComSpeed.Text = "Replay Speed"
        '
        'PlaybackBut
        '
        Me.PlaybackBut.Location = New System.Drawing.Point(8, 160)
        Me.PlaybackBut.Name = "PlaybackBut"
        Me.PlaybackBut.Size = New System.Drawing.Size(88, 24)
        Me.PlaybackBut.TabIndex = 2
        Me.PlaybackBut.Text = "&Replay"
        '
        'StopRecordBut
        '
        Me.StopRecordBut.Location = New System.Drawing.Point(8, 96)
        Me.StopRecordBut.Name = "StopRecordBut"
        Me.StopRecordBut.Size = New System.Drawing.Size(88, 24)
        Me.StopRecordBut.TabIndex = 1
        Me.StopRecordBut.Text = "&Stop Record"
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
        Me.AxMacroRecorder1.Location = New System.Drawing.Point(368, 16)
        Me.AxMacroRecorder1.Name = "AxMacroRecorder1"
        Me.AxMacroRecorder1.OcxState = CType(resources.GetObject("AxMacroRecorder1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMacroRecorder1.Size = New System.Drawing.Size(35, 35)
        Me.AxMacroRecorder1.TabIndex = 1
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(424, 206)
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
    '==============================================================================
    ' Method:        StopRecordBut_Click
    '
    ' Description:  Stop recording and save the macro to the disk.
    '============================================================================== 
    Private Sub StopRecordBut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StopRecordBut.Click
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
            'You can use any extension for the macro file name, or you can use no extension.
            'hence it will be easy to integrate the MacroRecorder ActiveX Control in your
            'Applications by using any file extension you want for the macro file generated by
            'MacroRecorder ActiveX Control.      
            .FileName = "macro1"
            .ShowDialog()
        End With

        'You can use any extension for the macro file name, or you can use no extension.
        'hence it will be easy to integrate the MacroRecorder ActiveX Control in your
        'Applications by using any file extension you want for the macro file generated by
        'MacroRecorder ActiveX Control.

        'For example suppose that you want to save the macro to c:\
        'If you want to assign (*.xxx) as an extension for the generated macro file,
        'then all you have to do is
        'MacroRecorder1.Save ("c:\mymacro.xxx")
        'If you want to use no extension then use the following:
        'MacroRecorder1.Save ("c:\mymacro")

        AxMacroRecorder1.Save(SaveFileDialog1.FileName)
  
    End Sub
    '==============================================================================
    ' Method:        PlaybackBut_Click
    '
    ' Description:  Start playback a macro.
    '==============================================================================
    Private Sub PlaybackBut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlaybackBut.Click

        If AxMacroRecorder1.IsReplay Then
            Exit Sub
        End If
        With OpenFileDialog1
            .ShowDialog()
        End With

        'Pass to MacroPath parameter any file you had saved by Save Method,
        'passing invalid macro file has no effect and no Replaying will happen.
        If OpenFileDialog1.FileName <> "" Then
            AxMacroRecorder1.StartReplay(OpenFileDialog1.FileName)
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComSpeed.SelectedIndex = 1
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        'Stop Recording when you Click Alt+F10
        If GetAsyncKeyState(121) And GetAsyncKeyState(18) Then
            StopRecord()
        End If
        'Stop Playbacking when you Click Alt+F9
        If GetAsyncKeyState(120) And GetAsyncKeyState(18) Then
            AxMacroRecorder1.StopReplay()
        End If

    End Sub

    Private Sub ComSpeed_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComSpeed.SelectedIndexChanged
        'To Replay with high speed set Speed parameter to 0
        'To Replay with Normal speed set Speed parameter to 1
        'To Replay with Low speed set Speed parameter to 2
        AxMacroRecorder1.SetReplaySpeed(ComSpeed.SelectedIndex)
    End Sub

    Private Sub AxMacroRecorder1_SaveComplete(ByVal sender As Object, ByVal e As AxMacroRecorderActX.__MacroRecorder_SaveCompleteEvent) Handles AxMacroRecorder1.SaveComplete
        If e.macFilePath <> "" Then MsgBox("Macro Saved To " & e.macFilePath)
    End Sub

    Private Sub AxMacroRecorder1_RecordStart(ByVal sender As Object, ByVal e As System.EventArgs) Handles AxMacroRecorder1.RecordStart
        MsgBox("Press Alt+F10 to Stop Recording.")
    End Sub

    Private Sub AxMacroRecorder1_ReplayStart(ByVal sender As Object, ByVal e As AxMacroRecorderActX.__MacroRecorder_ReplayStartEvent) Handles AxMacroRecorder1.ReplayStart
        MsgBox("Press Alt+F9 to terminate the replay.")
    End Sub

    Private Sub AxMacroRecorder1_ReplayFinish(ByVal sender As Object, ByVal e As AxMacroRecorderActX.__MacroRecorder_ReplayFinishEvent) Handles AxMacroRecorder1.ReplayFinish
        MsgBox("Playback " & e.macFilePath & " Macro Finish")
    End Sub
End Class
