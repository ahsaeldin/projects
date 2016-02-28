Imports System.Windows.Forms

Friend Class FrmKeyStroke

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        Dim lstItem As ListViewItem = Nothing
        Dim SelectKeyboardAction As String = ComKeyAction.SelectedItem.ToString

        With MainForm
            If Me.Tag = "edit" Then
                editSelectedItems(lstItem, SelectKeyboardAction)
            Else '//add a new item ()
                addNewItems(lstItem, SelectKeyboardAction)
            End If
            ProcessSubitems(lstItem)
        End With

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()

    End Sub

    Private Sub addNewItems(ByRef lstItem As ListViewItem, ByVal SelectKeyboardAction As String)
        Dim ItemIndex As Integer

        'add New Items
        With MainForm.lstviwMacEdit
            If .Items.Count > 0 Then 'لو كان فيه ماكرو معروض في القائمة

                Select Case .SelectedIndices.Count
                    Case 1
                        ItemIndex = .SelectedIndices.Item(0)
                    Case Is > 1
                        ItemIndex = .SelectedIndices.Item(.SelectedIndices.Count - 1)
                End Select

                lstItem = .Items.Insert(ItemIndex + 1, SelectKeyboardAction)

            ElseIf .Items.Count = 0 Then 'لو مافيش ماكرو معروض
                lstItem = .Items.Add(SelectKeyboardAction)
            End If
        End With
    End Sub

    Private Sub editSelectedItems(ByRef lstItem As ListViewItem, ByVal SelectKeyboardAction As String)

        With MainForm
            Dim ItemIndex As Integer

            If Me.Tag = "edit" Then

                Me.Tag = "" '//إلغاءه وتصفيته وإلا فلافائدة ترجي منه ولن يمكنتي  التحقق
                For Each selItem As ListViewItem In .lstviwMacEdit.SelectedItems

                    ItemIndex = selItem.Index
                    'بس هنلغي العنصر الحالي وهاضيفه من جديد بدل ما أدخل في صداع تعديل أمه
                    .lstviwMacEdit.Items.RemoveAt(ItemIndex)
                    lstItem = .lstviwMacEdit.Items.Insert(ItemIndex, SelectKeyboardAction)

                    ProcessSubitems(lstItem)

                Next

            End If

        End With
    End Sub

    Private Sub ProcessSubitems(ByRef lstItem As ListViewItem)
        'Add Subitems
        lstItem.SubItems.Add(ComKeys.SelectedItem) 'هذا هو أول عنصر
        'وهؤلاء هم البقية .. يتعين عليا الآن أن أعرف كيف سأضيفهم إلي الأعمدة
        'وهذه هي القيم المتوقعة لكل عنصر منهم , أتيت بها من ملف الماك
        '"KeyDown","A","15","False","False","False","65","65","0","False","False"
        '   0       1    2     3       4       5    ''6''  7   8     9       10
        'ولكي أعرف ما هي العناصر ذات الأهمية .. سأتجه إلي البلاي باك
        'KeyboardSimulator.KeyDown(lstviwMacEdit.Items.Item(lstItem.Index).SubItems(7).Text)
        'اذن العنصر 7 هو الوحيد ذو الأهمية .. حسنأ .. العنصر 7 هو السادس من هنا .. امممم
        'KeyData انه

        lstItem.SubItems.Add(0)
        lstItem.SubItems.Add("False")
        lstItem.SubItems.Add("False")
        lstItem.SubItems.Add("False")
        lstItem.SubItems.Add(GetKeyValue(ComKeys.SelectedItem))
        lstItem.SubItems.Add(GetKeyValue(ComKeys.SelectedItem))
        lstItem.SubItems.Add(0)
        lstItem.SubItems.Add("False")
        lstItem.SubItems.Add("False")
        'keyItm.SubItems.Add(MacroEvent.Time_SinceLastEvent) 'ليس ذو أهمية
        'keyItm.SubItems.Add(MacroEvent.ClonedKeyEventArgs.Alt)
        'keyItm.SubItems.Add(MacroEvent.ClonedKeyEventArgs.Control)
        'keyItm.SubItems.Add(MacroEvent.ClonedKeyEventArgs.Handled)
        'keyItm.SubItems.Add(MacroEvent.ClonedKeyEventArgs.KeyData)ذو أهمية
        'keyItm.SubItems.Add(MacroEvent.ClonedKeyEventArgs.KeyValue)
        'keyItm.SubItems.Add(MacroEvent.ClonedKeyEventArgs.Modifiers)
        'keyItm.SubItems.Add(MacroEvent.ClonedKeyEventArgs.Shift)
        'keyItm.SubItems.Add(MacroEvent.ClonedKeyEventArgs.SuppressKeyPress)

        lstItem.ForeColor = MainForm.GetItemColor(lstItem)
        MainForm.FocusThisItem(lstItem)
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Friend Function GetKeyValue(ByVal KeyName As String) As Integer
        Dim Value As Integer = Keys.Parse(GetType(Keys), ComKeys.SelectedItem)
        Return Value
    End Function

    Friend Sub LoadComKeysByKeys()
        If ComKeys.Items.Count = 0 Then
            Dim KeyEnumNames As Array = System.Enum.GetNames(GetType(System.Windows.Forms.Keys))
            Dim KeyEnumValues As Array = System.Enum.GetValues(GetType(System.Windows.Forms.Keys))

            For Each Name As String In KeyEnumNames
                If Not ProhibtedKey(Name) Then
                    ComKeys.Items.Add(Name)
                End If
                'My.Computer.FileSystem.WriteAllText("C:\Test.txt", Name & vbNewLine, True)
            Next
            'تحديد أول عنصر من هنا , لأنها أول مرة تُحمل
            ComKeys.SelectedIndex = ComKeys.Items.IndexOf("A")
        End If
    End Sub

    Private Function ProhibtedKey(ByVal KeyName As String) As Boolean
        Select Case KeyName 'مش عارف ايه الأزرار دي ومش هضيفها
            Case "None", "LButton", "RButton", "Cancel", "MButton", "XButton1", "XButton2"
                ProhibtedKey = True
        End Select
    End Function

    Private Sub FrmKeyStroke_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadComKeysByKeys()

        If Me.Tag <> "edit" Then
            ComKeyAction.SelectedIndex = 0
            Select Case ComKeys.Items.Count
                Case 0
                    ComKeys.SelectedIndex = 0
                Case Else 'لو حالة أخري , إذن
                    'سيبها زي ما هي,أنا حددتها بالفعل في الإستدعاء السابق لي .. أنا التي تحدثك 
            End Select
        End If

    End Sub

End Class
