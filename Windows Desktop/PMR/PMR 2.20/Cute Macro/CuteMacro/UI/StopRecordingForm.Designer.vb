<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class StopRecordingForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(StopRecordingForm))
        Me.ButStopRecording = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ButStopRecording
        '
        Me.ButStopRecording.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButStopRecording.Image = CType(resources.GetObject("ButStopRecording.Image"), System.Drawing.Image)
        Me.ButStopRecording.Location = New System.Drawing.Point(-1, -1)
        Me.ButStopRecording.Margin = New System.Windows.Forms.Padding(18, 3, 18, 3)
        Me.ButStopRecording.Name = "ButStopRecording"
        Me.ButStopRecording.Size = New System.Drawing.Size(39, 38)
        Me.ButStopRecording.TabIndex = 0
        Me.ButStopRecording.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.ButStopRecording.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.ButStopRecording.UseVisualStyleBackColor = True
        '
        'StopRecordingForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.AliceBlue
        Me.ClientSize = New System.Drawing.Size(37, 36)
        Me.ControlBox = False
        Me.Controls.Add(Me.ButStopRecording)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(18, 3, 18, 3)
        Me.Name = "StopRecordingForm"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ButStopRecording As System.Windows.Forms.Button
End Class
