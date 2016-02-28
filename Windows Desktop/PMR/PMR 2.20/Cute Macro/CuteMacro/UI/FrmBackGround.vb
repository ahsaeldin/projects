Imports System.Windows.Forms

Friend Class FrmBackGround

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        Dim ItemIndex As Integer
        Dim lstItem As ListViewItem = Nothing

        With MainForm
            If TxtTitle.Text <> "" Then
                If Me.Tag = "edit" Then

                    Me.Tag = "" '//إلغاءه وتصفيته وإلا فلافائدة ترجي منه ولن يمكنتي  التحقق 

                    If .lstviwMacEdit.Items.Count > 1 Then 'لو كان الليست فيها أكتر من عنصر واحد
                        'هنضيف الجديد بعد اخر واحد
                        ItemIndex = .lstviwMacEdit.SelectedIndices.Item(.lstviwMacEdit.SelectedIndices.Count - 1)
                    ElseIf .lstviwMacEdit.Items.Count = 1 Then
                        ItemIndex = .lstviwMacEdit.SelectedIndices.Item(0)
                    End If

                    'بس هنلغي العنصر الحالي وهاضيفه من جديد بدل ما أدخل في صداع تعديل أمه
                    .lstviwMacEdit.Items.RemoveAt(ItemIndex)
                    lstItem = .lstviwMacEdit.Items.Insert(ItemIndex, "Background Window")

                Else '//add a new item ()
                    If .lstviwMacEdit.Items.Count > 0 Then 'لو كان فيه ماكرو معروض في القائمة

                        Select Case .lstviwMacEdit.SelectedIndices.Count
                            Case 1
                                ItemIndex = .lstviwMacEdit.SelectedIndices.Item(0)
                            Case Is > 1
                                ItemIndex = .lstviwMacEdit.SelectedIndices.Item(.lstviwMacEdit.SelectedIndices.Count - 1)
                        End Select

                        lstItem = .lstviwMacEdit.Items.Insert(ItemIndex + 1, "Background Window")

                    ElseIf .lstviwMacEdit.Items.Count = 0 Then 'لو مافيش ماكرو معروض
                        lstItem = .lstviwMacEdit.Items.Add("Background Window")
                    End If
                End If
                'Add Subitems
                lstItem.SubItems.Add(TxtTitle.Text)
                lstItem.SubItems.Add(mstLeft.Text)
                lstItem.SubItems.Add(mstTop.Text)
                lstItem.SubItems.Add(CInt(mstWidth.Text) + CInt(mstLeft.Text))
                lstItem.SubItems.Add(CInt(mstHeight.Text) + CInt(mstTop.Text))
                '//add wait to window value, and the tag contains if ture or not
                lstItem.SubItems.Add(TxtWait.Text)
                lstItem.SubItems.Add(ChkWaitForWindow.Checked) '//contains if ture or not
                '28/09/2010
                'اضافة مسار واسم برنامج النافذة للليست كي استخدمها في حال البحث عنها لفتحها
                'lstItem.SubItems.Add(LabelProcessName.Text) 'FoxSmr
                'lstItem.SubItems.Add(LabelProcessPath.Text) 'FoxSmr
                lstItem.ForeColor = MainForm.GetItemColor(lstItem)
                MainForm.FocusThisItem(lstItem)
            End If

        End With
        TxtTitle.Focus() '//return focus to that textbox coz we only hide the form not closing and this causes that textbox to lost focus
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        TxtTitle.Focus() '//return focus to that textbox coz we only hide the form not closing and this causes that textbox to lost focus
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub tmrGetTitle_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrGetTitle.Tick


        If TxtTitle.Text <> "" And mstHeight.Text <> "" And mstLeft.Text <> "" And mstWidth.Text <> "" And mstTop.Text <> "" Then
            OK_Button.Enabled = True
        Else
            OK_Button.Enabled = False
        End If

        If GetAsyncKeyState(VK_LBUTTON) And GetAsyncKeyState(VK_MENU) Then

            Dim Res As Integer
            Dim ForeWinRect As RECT
            Dim WindowText As String
            Dim WindowTextLength As Integer

            Dim ClickedWindowHandle As Integer = _
            Win32API.WindowFromPoint(Windows.Forms.Cursor.Position.X, Windows.Forms.Cursor.Position.Y)

            'لو كانت ويندو رئيسية وليست طفلة عشان لو معملتش كدة أي 
            'child window
            'هتدخل في الزحمة

            WindowTextLength = GetWindowTextLength(ClickedWindowHandle)
            WindowText = Space(WindowTextLength + 1)
            Res = GetWindowText(ClickedWindowHandle, WindowText, WindowTextLength + 1)

            'لماذا
            'Or WindowText.Contains("FolderView")
            '??
            'لأنني وجدت أنه في بعض الأحيان لاتستطيع هذه الوظيفة تعقب نافذة سطح المكتب
            'ومن ثم فأنا اتحسس وجود نافذة ال
            'FolderView
            ' ومن ثم اقتنص الديسكتوب أو
            'Program Manager
            If Not CBool(IsChildWindow(ClickedWindowHandle)) Or WindowText.Contains("FolderView") Then

                If WindowText.Contains("FolderView") Then WindowText = "Program Manager"

                TxtTitle.Text = WindowText '//Press Alt+click any window to copy it's title
                '//and the dimensions of the window 
                GetWindowRect(ClickedWindowHandle, ForeWinRect)
                mstLeft.Text = ForeWinRect.Left
                mstTop.Text = ForeWinRect.Top
                mstHeight.Text = ForeWinRect.Bottom - ForeWinRect.Top
                mstWidth.Text = ForeWinRect.Right - ForeWinRect.Left

                '28/09/2010 'FoxSmr
                'Dim processName As String = Nothing 'FoxSmr
                'Dim processPath As String = Nothing 'FoxSmr 
                'اضافة مسار واسم برنامج النافذة للليست كي استخدمها في حال البحث عنها لفتحها
                'Dim foreProcess As Process = GetProcessPathByHandle(ClickedWindowHandle, WindowText, processName, processPath) 'FoxSmr
                'LabelProcessName.Text = processName 'FoxSmr
                'LabelProcessPath.Text = processPath 'FoxSmr

            End If

        End If

    End Sub

    Private Sub selectMyContent(ByVal box As Object)
        box.HideSelection = False '// to make the selections appears to the user
        box.Focus() '// to allow the user to edit txtRepeatTimes directly without focusing it first 
        box.SelectAll()
    End Sub

    Private Sub mstWidth_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles mstWidth.MouseClick
        selectMyContent(mstWidth)
    End Sub

    Private Sub mstHeight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mstHeight.Click
        selectMyContent(mstHeight)
    End Sub

    Private Sub mstTop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mstTop.Click
        selectMyContent(mstTop)
    End Sub

    Private Sub mstLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mstLeft.Click
        selectMyContent(mstLeft)
    End Sub

    Private Sub TxtWait_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TxtWait.Validating
        validateMyTxtBoxes(TxtWait)
    End Sub

    Private Sub TxtWait_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtWait.TextChanged
        If Len(TxtWait.Text) >= 4 Then '//not more than 10 seconds == 9999
            TxtWait.Text = Strings.Left(TxtWait.Text, 4)
            TxtWait.SelectionStart = Len(TxtWait.Text)
        End If
    End Sub

    Private Sub TxtTitle_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTitle.KeyDown
        If e.KeyCode = Keys.A And e.Control Then
            TxtTitle.SelectAll()
        End If
    End Sub
End Class
