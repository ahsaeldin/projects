Imports System.Text.RegularExpressions

Public Class SettingsForm
    Private Sub SettingsForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        If IsValidDropBoxPublicFolder(My.Settings.DropBoxPath) Then DropBoxPublicFolderTextBox.Text = My.Settings.DropBoxPath

        If My.Settings.AutoImportEnabled Then
            AutoScrapeCheckBox.Checked = True
        Else
            AutoScrapeCheckBox.Checked = False
        End If
        AutoImportNumericUpDown.Value = If(My.Settings.AutoImportValue > 5, My.Settings.AutoImportValue, 5)

        If IsValidDropBoxPublicUrl(My.Settings.DropBoxUrl) Then DropBoxTextBoxUrl.Text = My.Settings.DropBoxUrl

        If IsNumeric(My.Settings.PriceChangeByPercentage) AndAlso My.Settings.PriceChangeByPercentage >= 0 Then
            PercentageNumericUpDown.Value = My.Settings.PriceChangeByPercentage
        End If

        If IsNumeric(My.Settings.PriceChangingType) AndAlso My.Settings.PriceChangingType = 0 Then
            RadioButton1.Checked = True
        ElseIf IsNumeric(My.Settings.PriceChangingType) AndAlso My.Settings.PriceChangingType = 1 Then
            RadioButton2.Checked = True
        End If

        If IsNumeric(My.Settings.PriceChangeByValue) AndAlso My.Settings.PriceChangeByValue >= 0 Then
            ConstantChangeNumericUpDown.Value = My.Settings.PriceChangeByValue
        End If

        Select Case My.Settings.PriceType
            Case "US"
                UnitedStatesWarehouseRadioButton.Checked = True
            Case "UK"
                UnitedKingdomWarehouseRadioButton.Checked = True
            Case "HK"
                HongKongWarehouseRadioButton.Checked = True
        End Select

        If My.Settings.ShippingCostIsSet = True Then
            TextBox1.Text = My.Settings.ShippingCost
            ShippingCheckBox.Checked = True
        Else
            ShippingCheckBox.Checked = False
        End If

        If My.Settings.EnablePriceRange Then
            RangeCheckBox.Checked = True
        Else
            RangeCheckBox.Checked = False
        End If
        MinPriceNumericUpDown.Value = My.Settings.MinPrice
        MaxPriceNumericUpDown2.Value = My.Settings.MaxPrice

    End Sub

    Private Sub SaveButton_Click(sender As System.Object, e As System.EventArgs) Handles SaveButton.Click
        My.Settings.AutoImportEnabled = AutoScrapeCheckBox.Checked
        My.Settings.AutoImportValue = AutoImportNumericUpDown.Value
        If DropBoxPublicFolderTextBox.Text <> "" Then My.Settings.DropBoxPath = DropBoxPublicFolderTextBox.Text
        If IsValidDropBoxPublicUrl(DropBoxTextBoxUrl.Text) Then
            My.Settings.DropBoxUrl = DropBoxTextBoxUrl.Text
        Else
            LogState("DropBox Public Url not valid::" + DropBoxTextBoxUrl.Text)
        End If

        If IsNumeric(PercentageNumericUpDown.Value) AndAlso PercentageNumericUpDown.Value >= 0 Then
            My.Settings.PriceChangeByPercentage = PercentageNumericUpDown.Value
        End If

        If RadioButton1.Checked Then
            My.Settings.PriceChangingType = 0
        ElseIf RadioButton2.Checked Then
            My.Settings.PriceChangingType = 1
        End If

        If ShippingCheckBox.Checked Then
            My.Settings.ShippingCostIsSet = True
            My.Settings.ShippingCost = TextBox1.Text
        Else
            My.Settings.ShippingCostIsSet = False
            '   My.Settings.ShippingCost = TextBox1.Text
        End If

        If IsNumeric(ConstantChangeNumericUpDown.Value) AndAlso ConstantChangeNumericUpDown.Value >= 0 Then
            My.Settings.PriceChangeByValue = ConstantChangeNumericUpDown.Value
        End If

        If UnitedStatesWarehouseRadioButton.Checked Then
            My.Settings.PriceType = "US"
        ElseIf UnitedKingdomWarehouseRadioButton.Checked Then
            My.Settings.PriceType = "UK"
        ElseIf HongKongWarehouseRadioButton.Checked Then
            My.Settings.PriceType = "HK"
        End If

        If RangeCheckBox.Checked Then
            My.Settings.EnablePriceRange = True
        Else
            My.Settings.EnablePriceRange = False
        End If
        My.Settings.MinPrice = MinPriceNumericUpDown.Value
        My.Settings.MaxPrice = MaxPriceNumericUpDown2.Value

        My.Settings.Save()

    End Sub

    Private Sub BrowseButton_Click(sender As System.Object, e As System.EventArgs) Handles BrowseButton.Click
        Dim diaRes As DialogResult = FolderBrowserDia.ShowDialog()
        If diaRes = Windows.Forms.DialogResult.OK Then
            Dim dropBoxFolderPath As String = FolderBrowserDia.SelectedPath
            If IsValidDropBoxPublicFolder(dropBoxFolderPath) Then
                DropBoxPublicFolderTextBox.Text = dropBoxFolderPath
                LogState("DropBox folder is valid::" + dropBoxFolderPath)
            Else
                DropBoxPublicFolderTextBox.Text = ""
                LogState("DropBox folder is not valid::" + dropBoxFolderPath)
                MsgBox("DropBox folder path not valid", MsgBoxStyle.Critical, "Tmart Scraper")
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        SaveButton_Click(Me, e)
        Me.Dispose()
    End Sub
End Class