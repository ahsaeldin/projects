'=============================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:  A sample code demonstrates how to use MacroRecorder
'               ActiveX control v1.50 to convert normal macro file to
'               an executable macro file.
'=============================================================================

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
    Friend WithEvents StopRecordBut As System.Windows.Forms.Button
    Friend WithEvents RecordBut As System.Windows.Forms.Button
    Friend WithEvents ButConvert As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
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
        Me.ButConvert = New System.Windows.Forms.Button
        Me.StopRecordBut = New System.Windows.Forms.Button
        Me.RecordBut = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
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
        Me.GroupBox1.Controls.Add(Me.ButConvert)
        Me.GroupBox1.Controls.Add(Me.StopRecordBut)
        Me.GroupBox1.Controls.Add(Me.RecordBut)
        Me.GroupBox1.Controls.Add(Me.AxMacroRecorder1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(432, 192)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 136)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(416, 24)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "3. Click Convert button to convert a normal macro file to an executable macro fil" & _
        "e."
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(392, 24)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "2. Click Stop Record button to stop recording and save the macro to the disk."
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(288, 24)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "1. Click Record button to start recording."
        '
        'ButConvert
        '
        Me.ButConvert.Location = New System.Drawing.Point(8, 160)
        Me.ButConvert.Name = "ButConvert"
        Me.ButConvert.Size = New System.Drawing.Size(88, 24)
        Me.ButConvert.TabIndex = 2
        Me.ButConvert.Text = "&Convert"
        '
        'StopRecordBut
        '
        Me.StopRecordBut.Location = New System.Drawing.Point(8, 104)
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
        Me.AxMacroRecorder1.Location = New System.Drawing.Point(392, 16)
        Me.AxMacroRecorder1.Name = "AxMacroRecorder1"
        Me.AxMacroRecorder1.OcxState = CType(resources.GetObject("AxMacroRecorder1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMacroRecorder1.Size = New System.Drawing.Size(35, 35)
        Me.AxMacroRecorder1.TabIndex = 1
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(448, 198)
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
    ' Method:        ButConvert_Click
    '
    ' Description: convert normal macro file to an executable macro file.
    '==============================================================================
    Private Sub ButConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButConvert.Click
        If Not AxMacroRecorder1.IsRecord And Not AxMacroRecorder1.IsReplay Then
            OpenFileDialog1.Title = "Select Macro to Convert to exe file."
            OpenFileDialog1.ShowDialog()
            AxMacroRecorder1.ConvertToExE(OpenFileDialog1.FileName, "c:\macro1.exe")
        End If
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

    Private Sub AxMacroRecorder1_ConvertComplete(ByVal sender As Object, ByVal e As AxMacroRecorderActX.__MacroRecorder_ConvertCompleteEvent) Handles AxMacroRecorder1.ConvertComplete
        If e.macFilePath <> "" Then MsgBox("Macro Converted and saved To " & e.macFilePath)
    End Sub
End Class
