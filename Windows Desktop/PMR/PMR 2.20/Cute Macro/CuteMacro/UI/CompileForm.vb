Imports CuteMacro.Balora

Friend Class CompileForm

    Private Sub txtRepeatTimes_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtRepeatTimes.Validating

        IsOnlyNumeric(txtRepeatTimes.Text, txtRepeatTimes)

        If Val(txtRepeatTimes.Text) <= 0 Then ' <= zero isn't allowed
            txtRepeatTimes.Text = 1
        End If

    End Sub

    Private Sub txtRepeatTimes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRepeatTimes.TextChanged

        If Len(txtRepeatTimes.Text) >= 10 Then

            txtRepeatTimes.Text = Strings.Left(txtRepeatTimes.Text, 9)
            txtRepeatTimes.SelectionStart = Len(txtRepeatTimes.Text)

        End If

    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click

        If is_x2342 = "False" Then
            MsgBox("You can compile macro to exe files in the registered version only.", MsgBoxStyle.Information, "Perfect Macro Recorder")
            Exit Sub
        End If

        Dim RepeatTime As String = txtRepeatTimes.Text
        Dim PlaybackSpeed As String = PlaSpdSlider.Value

        '*RepeatTime#PlaybackSpeed#chkNotify.Checked#chkCompileSelected.Checked#chkEnableEditing.Checked*

        Dim ExecutableSettingsAddedToKey As String = "*" _
        & RepeatTime & "#" & PlaybackSpeed & "#" & chkNotify.Checked _
        & "#" & chkCompileSelected.Checked & "#" & chkEnableEditing.Checked & "*"

        Dim CompiledFilePath As String = MainForm.ShowSaveDialog("Executable " & "Perfect Macro Recorder" & " Files (*.exe)|*.exe")

        If CompiledFilePath = "" Then Exit Sub
        'Ahmed Saad 11/11/2010: commented next line as I will use...
        ' CodeDom to generate executable macro files to avoid stuipd...
        'virus alert by the hell called avira anti-vir.
        'Dim CompilingResult As Result = MainForm.CompileTheList(CompiledFilePath, ExecutableSettingsAddedToKey)
        Dim exeFileSettings As New ArrayList
        exeFileSettings.AddRange(New String() {RepeatTime, PlaybackSpeed})
        Dim CompilingResult As New Result

        If is_x2342 = "True" Then
            CompilingResult = CompileMacroToExE(CompiledFilePath, exeFileSettings)
        End If

        If CompilingResult.Successed = False Then
            If CompilingResult.Message = "" Then
                MsgBox("The compiling process of the macro failed.", MsgBoxStyle.Critical, Me.Text)
            Else
                MsgBox(CompilingResult.Message, MsgBoxStyle.Critical, Me.Text)
            End If
        End If

        Me.Close()

    End Sub

    Private Sub CompileForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MainForm.lstviwMacEdit.SelectedItems.Count > 0 Then
            chkCompileSelected.Enabled = True
        End If
    End Sub

End Class