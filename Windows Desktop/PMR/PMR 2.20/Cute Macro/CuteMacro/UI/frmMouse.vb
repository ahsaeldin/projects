Imports System.Windows.Forms

Friend Class frmMouse

    Public CaptureStarted As Boolean = False

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        addMouseAction()

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub addMouseAction()

        Dim addMouseMessage As String = "Use this window to add a mouse action to the list. Select the mouse action type and the coordinates of the cursor associated with the action."

        Dim ItemIndex As Integer
        Dim lstItem As ListViewItem = Nothing

        With MainForm

            If lblDesc.Text = "Use this window to edit a mouse action from the list. Select the mouse action type and the coordinates of the cursor associated with the action." Then 'لو كان تحرير عنصر ماوس

                If .lstviwMacEdit.Items.Count > 1 Then 'لو كان الليست فيها أكتر من عنصر واحد
                    'هنضيف الجديد بعد اخر واحد
                    ItemIndex = .lstviwMacEdit.SelectedIndices.Item(.lstviwMacEdit.SelectedIndices.Count - 1)
                ElseIf .lstviwMacEdit.Items.Count = 1 Then
                    ItemIndex = .lstviwMacEdit.SelectedIndices.Item(0)
                End If

                'بس هنلغي العنصر الحالي وهاضيفه من جديد بدل ما أدخل في صداع تعديل أمه
                .lstviwMacEdit.Items.RemoveAt(ItemIndex)

                lstItem = .lstviwMacEdit.Items.Insert(ItemIndex, comMouseEvents.SelectedItem)

            ElseIf lblDesc.Text = addMouseMessage Then 'لو كان إضافة لعنصر ماوس جديد

                If .lstviwMacEdit.Items.Count > 0 Then 'لو كان فيه ماكرو معروض في القائمة

                    Select Case .lstviwMacEdit.SelectedIndices.Count
                        Case 1
                            ItemIndex = .lstviwMacEdit.SelectedIndices.Item(0)
                        Case Is > 1
                            ItemIndex = .lstviwMacEdit.SelectedIndices.Item(.lstviwMacEdit.SelectedIndices.Count - 1)
                    End Select

                    lstItem = .lstviwMacEdit.Items.Insert(ItemIndex + 1, comMouseEvents.SelectedItem)

                ElseIf .lstviwMacEdit.Items.Count = 0 Then 'لو مافيش ماكرو معروض
                    lstItem = .lstviwMacEdit.Items.Add(comMouseEvents.SelectedItem)
                End If
            End If

            'تحرير أو إضافة العنصر وهذا هو ذكاء الكود الخاص بي حيث لم أعمد إلي التكرار
            Select Case comMouseEvents.SelectedItem
                Case "LeftMouseDown", "LeftMouseUp"
                    '"LeftMouseDown","X = 68","Y = 24","32","1048576","1","0"
                    '"LeftMouseUp","X = 68","Y = 24","93","1048576","1","0"
                    lstItem.SubItems.Add("X = " & mstX.Text).Tag = mstX.Text
                    lstItem.SubItems.Add("Y = " & mstY.Text).Tag = mstY.Text
                    lstItem.SubItems.Add(0)
                    lstItem.SubItems.Add(1048576)
                    lstItem.SubItems.Add(1)
                    lstItem.SubItems.Add(0)
                Case "MouseWheel"
                    '"MouseWheel","65416","X = 651","Y = 406","141","0","1"
                    lstItem.SubItems.Add(mstWheelValue.Text)
                    lstItem.SubItems.Add("X = " & mstX.Text).Tag = mstX.Text
                    lstItem.SubItems.Add("Y = " & mstY.Text).Tag = mstY.Text
                    lstItem.SubItems.Add(0)
                    lstItem.SubItems.Add(0)
                    lstItem.SubItems.Add(1)
                Case "MiddleMouseDown", "MiddleMouseUp"
                    '"MiddleMouseDown","X = 651","Y = 390","172","4194304","1","0"
                    '"MiddleMouseUp","X = 651","Y = 390","218","4194304","1","0"
                    lstItem.SubItems.Add("X = " & mstX.Text).Tag = mstX.Text
                    lstItem.SubItems.Add("Y = " & mstY.Text).Tag = mstY.Text
                    lstItem.SubItems.Add(0)
                    lstItem.SubItems.Add(4194304)
                    lstItem.SubItems.Add(1)
                    lstItem.SubItems.Add(0)
                Case "MouseMove"
                    '"MouseMove","X = 309","Y = 329","16","0","1","0"
                    lstItem.SubItems.Add("X = " & mstX.Text).Tag = mstX.Text
                    lstItem.SubItems.Add("Y = " & mstY.Text).Tag = mstY.Text
                    lstItem.SubItems.Add(0)
                    lstItem.SubItems.Add(0)
                    lstItem.SubItems.Add(1)
                    lstItem.SubItems.Add(0)
                Case "RightMouseDown", "RightMouseUp"
                    '"RightMouseDown","X = 494","Y = 279","733","2097152","1","0"
                    '"RightMouseUp","X = 494","Y = 279","172","2097152","1","0"
                    lstItem.SubItems.Add("X = " & mstX.Text).Tag = mstX.Text
                    lstItem.SubItems.Add("Y = " & mstY.Text).Tag = mstY.Text
                    lstItem.SubItems.Add(0)
                    lstItem.SubItems.Add(2097152)
                    lstItem.SubItems.Add(1)
                    lstItem.SubItems.Add(0)
            End Select

        End With

        MainForm.FocusThisItem(lstItem)
        lstItem.ForeColor = MainForm.GetItemColor(lstItem)

    End Sub

    Private Sub tmrCoordinates_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrCoordinates.Tick

        If GetAsyncKeyState(VK_LBUTTON) And GetAsyncKeyState(VK_MENU) _
        And Not CaptureStarted Then
            CaptureStarted = True
        ElseIf GetAsyncKeyState(VK_LBUTTON) And GetAsyncKeyState(VK_MENU) _
        And CaptureStarted Then
            CaptureStarted = False
        End If

        If CaptureStarted Then
            mstX.Text = Windows.Forms.Cursor.Position.X()
            mstY.Text = Windows.Forms.Cursor.Position.Y()
        End If

    End Sub

    Private Sub comMouseEvents_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles comMouseEvents.SelectedValueChanged

        If comMouseEvents.SelectedIndex = 5 Then '// visible/invisible MouseWheel due to the comMouseEvents
            mstWheelValue.Text = "00000"
            mstWheelValue.Visible = True
            lblMouseWhlVal.Visible = True
        Else
            mstWheelValue.Visible = False
            lblMouseWhlVal.Visible = False
        End If

    End Sub

    Private Sub frmMouse_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CaptureStarted = False
        If mstX.Text = "" Then
            mstX.Text = Windows.Forms.Cursor.Position.X()
            mstY.Text = Windows.Forms.Cursor.Position.Y()
        End If
    End Sub

    Private Sub mstWheelValue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mstWheelValue.Click
        mstWheelValue.HideSelection = False '// to make the selections appears to the user
        mstWheelValue.Focus() '// to allow the user to edit txtRepeatTimes directly without focusing it first 
        mstWheelValue.SelectAll()
    End Sub


End Class