Imports System.Windows.Forms

'استخدم هذه الفورم لإظهار شريط تقدم في حالة الغاء عناصر كثيرة في الليست
Friend Class frmProgress

    Private Sub frmProgress_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Me.Tag = "Deleting" Then
            'فقط أظهر هذه الفورم لو كان عدد العناصر الملغاة أكبر من 10
            If MainForm.lstviwMacEdit.SelectedItems.Count <= 10 Then
                Me.Visible = False
                Me.Opacity = 0
            End If
        End If
    End Sub

    Private Sub frmProgress_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If Me.Tag = "Deleting" Then
            Me.Tag = ""
            DeleteItems()
        End If
    End Sub

    Private Sub DeleteItems()

        Dim lstItem As ListViewItem
        Dim LastDeletedItemIndex As Integer

        With MainForm

            .lstviwMacEdit.BeginUpdate()

            ProgBar.Maximum = .lstviwMacEdit.SelectedItems.Count

            For Each lstItem In .lstviwMacEdit.SelectedItems

                Application.DoEvents()
                Dim c As Integer : c = c + 1
                ProgBar.Value = c
                LastDeletedItemIndex = lstItem.Index
                lstItem.Remove()

            Next

            .lstviwMacEdit.EndUpdate()

            If .lstviwMacEdit.Items.Count <> 0 Then

                If LastDeletedItemIndex = .lstviwMacEdit.Items.Count Then
                    '// Story : point to the last item in the last and delete it,
                    '// then an exception will raised .. because "LastDeletedItemIndex"
                    '// is the number of the last item that already deleted .. that already
                    '// has no existence and this is the reason of the thrown exception

                    '// The Sol : just select the previous item which will be "LastDeletedItemIndex - 1" 
                    .lstviwMacEdit.Items(LastDeletedItemIndex - 1).Selected = True
                Else
                    '//this is the normal case where the deleted item is any item but the last item
                    .lstviwMacEdit.Items(LastDeletedItemIndex).Selected = True
                End If

            End If

        End With

        Me.Dispose()

    End Sub

End Class
