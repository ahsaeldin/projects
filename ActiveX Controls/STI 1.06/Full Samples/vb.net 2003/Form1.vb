Public Class Form1
    Inherits System.Windows.Forms.Form

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

    Friend WithEvents MSFlexGrid1 As AxMSFlexGridLib.AxMSFlexGrid
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents HideButton As System.Windows.Forms.Button
    Friend WithEvents FeedBackButton As System.Windows.Forms.Button
    Friend WithEvents AxTrayIcons1 As AxSTI.AxTrayIcons
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.MSFlexGrid1 = New AxMSFlexGridLib.AxMSFlexGrid
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.FeedBackButton = New System.Windows.Forms.Button
        Me.HideButton = New System.Windows.Forms.Button
        Me.AxTrayIcons1 = New AxSTI.AxTrayIcons
        CType(Me.MSFlexGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.AxTrayIcons1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MSFlexGrid1
        '
        Me.MSFlexGrid1.Location = New System.Drawing.Point(8, 8)
        Me.MSFlexGrid1.Name = "MSFlexGrid1"
        Me.MSFlexGrid1.OcxState = CType(resources.GetObject("MSFlexGrid1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MSFlexGrid1.Size = New System.Drawing.Size(664, 192)
        Me.MSFlexGrid1.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.FeedBackButton)
        Me.GroupBox1.Controls.Add(Me.HideButton)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 200)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(168, 48)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        '
        'FeedBackButton
        '
        Me.FeedBackButton.Location = New System.Drawing.Point(88, 16)
        Me.FeedBackButton.Name = "FeedBackButton"
        Me.FeedBackButton.Size = New System.Drawing.Size(72, 24)
        Me.FeedBackButton.TabIndex = 1
        Me.FeedBackButton.Text = "&FeedBack"
        '
        'HideButton
        '
        Me.HideButton.Location = New System.Drawing.Point(8, 16)
        Me.HideButton.Name = "HideButton"
        Me.HideButton.Size = New System.Drawing.Size(72, 24)
        Me.HideButton.TabIndex = 0
        Me.HideButton.Text = "&Hide"
        '
        'AxTrayIcons1
        '
        Me.AxTrayIcons1.Enabled = True
        Me.AxTrayIcons1.Location = New System.Drawing.Point(184, 208)
        Me.AxTrayIcons1.Name = "AxTrayIcons1"
        Me.AxTrayIcons1.OcxState = CType(resources.GetObject("AxTrayIcons1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxTrayIcons1.Size = New System.Drawing.Size(36, 36)
        Me.AxTrayIcons1.TabIndex = 4
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(680, 254)
        Me.Controls.Add(Me.AxTrayIcons1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.MSFlexGrid1)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VB.Net TrayIcons"
        CType(Me.MSFlexGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.AxTrayIcons1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Fill the FlexGrid control with System Tray Data   
        FillFlexGrid()
    End Sub

    Private Sub initFlexGrid()
        'Set the FlexGrid 

        MSFlexGrid1.set_ColWidth(0, 300)
        MSFlexGrid1.set_TextMatrix(0, 0, "I")
        MSFlexGrid1.set_ColWidth(1, 3000)
        MSFlexGrid1.set_TextMatrix(0, 1, "Application Path")
        MSFlexGrid1.set_ColWidth(2, 1300)
        MSFlexGrid1.set_TextMatrix(0, 2, "uID")
        MSFlexGrid1.set_ColWidth(3, 1300)
        MSFlexGrid1.set_TextMatrix(0, 3, "hWnd")
        MSFlexGrid1.set_ColWidth(4, 1300)
        MSFlexGrid1.set_TextMatrix(0, 4, "hIcon")
        MSFlexGrid1.set_ColWidth(5, 1700)
        MSFlexGrid1.set_TextMatrix(0, 5, "ToolTip")
        MSFlexGrid1.set_ColWidth(6, 1500)
        MSFlexGrid1.set_TextMatrix(0, 6, "uCallbackMessage")
        MSFlexGrid1.ForeColorFixed = Color.Blue

    End Sub

    Private Sub FillFlexGrid()

        Dim i As Integer
        Dim TrayList() As STI.TrayIcon

        initFlexGrid()

        'Get the System Tray Data
        TrayList = AxTrayIcons1.GetSysTrayIcons()

        'Fill the FlexGrid with Data
        For i = 0 To AxTrayIcons1.IconCount - 1

            MSFlexGrid1.set_TextMatrix(i + 1, 0, i + 1)
            If TrayList(i).APath = "" Then
                MSFlexGrid1.set_TextMatrix(i + 1, 1, "N/A")
            Else
                MSFlexGrid1.set_TextMatrix(i + 1, 1, TrayList(i).APath)
            End If
            MSFlexGrid1.set_TextMatrix(i + 1, 2, TrayList(i).uId)
            MSFlexGrid1.set_TextMatrix(i + 1, 3, TrayList(i).hwnd)
            MSFlexGrid1.set_TextMatrix(i + 1, 4, TrayList(i).hIcon)
            MSFlexGrid1.set_TextMatrix(i + 1, 5, TrayList(i).ToolTip)
            MSFlexGrid1.set_TextMatrix(i + 1, 6, "&H" & Hex$(TrayList(i).ucallbackMessage))

        Next i

    End Sub

    Private Sub HideButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HideButton.Click
        If MSFlexGrid1.Visible = True Then
            'Hide the selected Item from the System Tray area
            AxTrayIcons1.CtlHide(MSFlexGrid1.RowSel())
            'Clear the FlexGrid
            MSFlexGrid1.Clear()
            'fill the FlexGrid 
            FillFlexGrid()
        End If
    End Sub

    Private Sub FeedBackButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FeedBackButton.Click
        AxTrayIcons1.sendreview()
    End Sub
End Class
