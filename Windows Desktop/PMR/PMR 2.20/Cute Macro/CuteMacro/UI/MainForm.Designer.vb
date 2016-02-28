<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.ShowSave = New System.Windows.Forms.SaveFileDialog()
        Me.ShowOpen = New System.Windows.Forms.OpenFileDialog()
        Me.TrayIcon = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.SystemTrayMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ShowCuteMacroToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator23 = New System.Windows.Forms.ToolStripSeparator()
        Me.RecordSystemTrayMenu = New System.Windows.Forms.ToolStripMenuItem()
        Me.StopToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PlayBackSystemTrayMenu = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator22 = New System.Windows.Forms.ToolStripSeparator()
        Me.OptionsSystemTrayMenu = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.WebPageToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutSystemTrayMenu = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator21 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitSystemTrayMenu = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpButtonContextMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.WebPageToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ButMacEdit = New System.Windows.Forms.Button()
        Me.btnGetNow = New System.Windows.Forms.Button()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.btnSveMac = New System.Windows.Forms.Button()
        Me.btnNewMac = New System.Windows.Forms.Button()
        Me.BtnStop = New System.Windows.Forms.Button()
        Me.btnRecord = New System.Windows.Forms.Button()
        Me.BtnPlayback = New System.Windows.Forms.Button()
        Me.ButOptions = New System.Windows.Forms.Button()
        Me.TmrStop = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.EditRightClickList = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.lstviewStripRun = New System.Windows.Forms.ToolStripMenuItem()
        Me.RunTheMacroToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RunSelectionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CompileToExeFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripEditCommand = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator14 = New System.Windows.Forms.ToolStripSeparator()
        Me.lstviewStripCut = New System.Windows.Forms.ToolStripMenuItem()
        Me.lstviewStripCopy = New System.Windows.Forms.ToolStripMenuItem()
        Me.lstviewStripPaste = New System.Windows.Forms.ToolStripMenuItem()
        Me.PasteAfterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PasteBeforeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReplaceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator13 = New System.Windows.Forms.ToolStripSeparator()
        Me.InsertToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BackgroundWindowTitleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DelayToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MouseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.KeyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MoreToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.lstviewStripDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator15 = New System.Windows.Forms.ToolStripSeparator()
        Me.lstviewStripMoveUp = New System.Windows.Forms.ToolStripMenuItem()
        Me.lstviewStripMoveDown = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator16 = New System.Windows.Forms.ToolStripSeparator()
        Me.SelectAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DeselectAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TmrGenLoop = New System.Windows.Forms.Timer(Me.components)
        Me.toolstripRecord = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.toolstripStopRecord = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.toolstripOptions = New System.Windows.Forms.ToolStripButton()
        Me.toolstripPlyBack = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.toolstripHelp = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripHorzButtons = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButtonRegister = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripNew = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripOpenMacro = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.CutToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.CopyToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.PasteToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripDelItems = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSveMac = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripAddKeys = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripAddDelay = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripAddMouseEve = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripRun = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator8 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripLeftVertButtons = New System.Windows.Forms.ToolStrip()
        Me.ToolStripCompile = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripBackGround = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripAddMore = New System.Windows.Forms.ToolStripButton()
        Me.StatusBar = New System.Windows.Forms.StatusStrip()
        Me.Status_LoadingLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lstviwMacEdit = New System.Windows.Forms.ListView()
        Me.ColCommand = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Col2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader8 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader9 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader10 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader11 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SystemTrayMenu.SuspendLayout()
        Me.HelpButtonContextMenu.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.EditRightClickList.SuspendLayout()
        Me.ToolStripHorzButtons.SuspendLayout()
        Me.ToolStripLeftVertButtons.SuspendLayout()
        Me.StatusBar.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ShowOpen
        '
        Me.ShowOpen.FileName = "OpenFileDialog1"
        '
        'TrayIcon
        '
        Me.TrayIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.TrayIcon.BalloonTipText = "Balloon Text Here"
        Me.TrayIcon.BalloonTipTitle = "Title Here"
        Me.TrayIcon.ContextMenuStrip = Me.SystemTrayMenu
        Me.TrayIcon.Icon = CType(resources.GetObject("TrayIcon.Icon"), System.Drawing.Icon)
        Me.TrayIcon.Text = "Perfect Macro Recorder"
        '
        'SystemTrayMenu
        '
        Me.SystemTrayMenu.BackColor = System.Drawing.SystemColors.Control
        Me.SystemTrayMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ShowCuteMacroToolStripMenuItem, Me.ToolStripSeparator23, Me.RecordSystemTrayMenu, Me.StopToolStripMenuItem, Me.PlayBackSystemTrayMenu, Me.ToolStripSeparator22, Me.OptionsSystemTrayMenu, Me.ToolStripSeparator1, Me.WebPageToolStripMenuItem1, Me.AboutSystemTrayMenu, Me.ToolStripSeparator21, Me.ExitSystemTrayMenu})
        Me.SystemTrayMenu.Name = "SystemTrayMenu"
        Me.SystemTrayMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional
        Me.SystemTrayMenu.Size = New System.Drawing.Size(231, 204)
        '
        'ShowCuteMacroToolStripMenuItem
        '
        Me.ShowCuteMacroToolStripMenuItem.Name = "ShowCuteMacroToolStripMenuItem"
        Me.ShowCuteMacroToolStripMenuItem.Size = New System.Drawing.Size(230, 22)
        Me.ShowCuteMacroToolStripMenuItem.Text = "&Show Perfect Macro Recorder"
        '
        'ToolStripSeparator23
        '
        Me.ToolStripSeparator23.Name = "ToolStripSeparator23"
        Me.ToolStripSeparator23.Size = New System.Drawing.Size(227, 6)
        '
        'RecordSystemTrayMenu
        '
        Me.RecordSystemTrayMenu.Image = CType(resources.GetObject("RecordSystemTrayMenu.Image"), System.Drawing.Image)
        Me.RecordSystemTrayMenu.Name = "RecordSystemTrayMenu"
        Me.RecordSystemTrayMenu.Size = New System.Drawing.Size(230, 22)
        Me.RecordSystemTrayMenu.Text = "&Record"
        '
        'StopToolStripMenuItem
        '
        Me.StopToolStripMenuItem.Image = CType(resources.GetObject("StopToolStripMenuItem.Image"), System.Drawing.Image)
        Me.StopToolStripMenuItem.Name = "StopToolStripMenuItem"
        Me.StopToolStripMenuItem.Size = New System.Drawing.Size(230, 22)
        Me.StopToolStripMenuItem.Text = "S&top Recording"
        '
        'PlayBackSystemTrayMenu
        '
        Me.PlayBackSystemTrayMenu.Image = CType(resources.GetObject("PlayBackSystemTrayMenu.Image"), System.Drawing.Image)
        Me.PlayBackSystemTrayMenu.Name = "PlayBackSystemTrayMenu"
        Me.PlayBackSystemTrayMenu.Size = New System.Drawing.Size(230, 22)
        Me.PlayBackSystemTrayMenu.Text = "&Plaback"
        '
        'ToolStripSeparator22
        '
        Me.ToolStripSeparator22.Name = "ToolStripSeparator22"
        Me.ToolStripSeparator22.Size = New System.Drawing.Size(227, 6)
        '
        'OptionsSystemTrayMenu
        '
        Me.OptionsSystemTrayMenu.Image = CType(resources.GetObject("OptionsSystemTrayMenu.Image"), System.Drawing.Image)
        Me.OptionsSystemTrayMenu.Name = "OptionsSystemTrayMenu"
        Me.OptionsSystemTrayMenu.Size = New System.Drawing.Size(230, 22)
        Me.OptionsSystemTrayMenu.Text = "&Options"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(227, 6)
        '
        'WebPageToolStripMenuItem1
        '
        Me.WebPageToolStripMenuItem1.Name = "WebPageToolStripMenuItem1"
        Me.WebPageToolStripMenuItem1.Size = New System.Drawing.Size(230, 22)
        Me.WebPageToolStripMenuItem1.Text = "&Web Page"
        '
        'AboutSystemTrayMenu
        '
        Me.AboutSystemTrayMenu.Name = "AboutSystemTrayMenu"
        Me.AboutSystemTrayMenu.Size = New System.Drawing.Size(230, 22)
        Me.AboutSystemTrayMenu.Text = "&About"
        '
        'ToolStripSeparator21
        '
        Me.ToolStripSeparator21.Name = "ToolStripSeparator21"
        Me.ToolStripSeparator21.Size = New System.Drawing.Size(227, 6)
        '
        'ExitSystemTrayMenu
        '
        Me.ExitSystemTrayMenu.Name = "ExitSystemTrayMenu"
        Me.ExitSystemTrayMenu.Size = New System.Drawing.Size(230, 22)
        Me.ExitSystemTrayMenu.Text = "&Exit"
        '
        'HelpButtonContextMenu
        '
        Me.HelpButtonContextMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HelpToolStripMenuItem, Me.WebPageToolStripMenuItem, Me.AboutToolStripMenuItem})
        Me.HelpButtonContextMenu.Name = "HelpButtonContextMenu"
        Me.HelpButtonContextMenu.Size = New System.Drawing.Size(128, 70)
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.HelpToolStripMenuItem.Text = "&Help"
        '
        'WebPageToolStripMenuItem
        '
        Me.WebPageToolStripMenuItem.Name = "WebPageToolStripMenuItem"
        Me.WebPageToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.WebPageToolStripMenuItem.Text = "&Web Page"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.AboutToolStripMenuItem.Text = "&About"
        '
        'ButMacEdit
        '
        Me.ButMacEdit.Location = New System.Drawing.Point(764, 298)
        Me.ButMacEdit.Name = "ButMacEdit"
        Me.ButMacEdit.Size = New System.Drawing.Size(85, 18)
        Me.ButMacEdit.TabIndex = 5
        Me.ButMacEdit.Text = "&Edit Macro"
        Me.ButMacEdit.UseVisualStyleBackColor = True
        '
        'btnGetNow
        '
        Me.btnGetNow.Location = New System.Drawing.Point(764, 422)
        Me.btnGetNow.Name = "btnGetNow"
        Me.btnGetNow.Size = New System.Drawing.Size(85, 18)
        Me.btnGetNow.TabIndex = 10
        Me.btnGetNow.Text = "&GetNow!"
        Me.btnGetNow.UseVisualStyleBackColor = True
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(764, 391)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(85, 18)
        Me.btnRun.TabIndex = 8
        Me.btnRun.Text = "R&un"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'btnSveMac
        '
        Me.btnSveMac.Location = New System.Drawing.Point(764, 360)
        Me.btnSveMac.Name = "btnSveMac"
        Me.btnSveMac.Size = New System.Drawing.Size(85, 18)
        Me.btnSveMac.TabIndex = 7
        Me.btnSveMac.Text = "S&ave Macro"
        Me.btnSveMac.UseVisualStyleBackColor = True
        '
        'btnNewMac
        '
        Me.btnNewMac.Location = New System.Drawing.Point(764, 329)
        Me.btnNewMac.Name = "btnNewMac"
        Me.btnNewMac.Size = New System.Drawing.Size(85, 18)
        Me.btnNewMac.TabIndex = 6
        Me.btnNewMac.Text = "&New Macro"
        Me.btnNewMac.UseVisualStyleBackColor = True
        '
        'BtnStop
        '
        Me.BtnStop.Location = New System.Drawing.Point(97, 19)
        Me.BtnStop.Name = "BtnStop"
        Me.BtnStop.Size = New System.Drawing.Size(85, 25)
        Me.BtnStop.TabIndex = 2
        Me.BtnStop.Text = "&Stop"
        Me.BtnStop.UseVisualStyleBackColor = True
        '
        'btnRecord
        '
        Me.btnRecord.Location = New System.Drawing.Point(6, 20)
        Me.btnRecord.Name = "btnRecord"
        Me.btnRecord.Size = New System.Drawing.Size(85, 24)
        Me.btnRecord.TabIndex = 1
        Me.btnRecord.Text = "&Record"
        Me.btnRecord.UseVisualStyleBackColor = True
        '
        'BtnPlayback
        '
        Me.BtnPlayback.Location = New System.Drawing.Point(188, 19)
        Me.BtnPlayback.Name = "BtnPlayback"
        Me.BtnPlayback.Size = New System.Drawing.Size(85, 25)
        Me.BtnPlayback.TabIndex = 3
        Me.BtnPlayback.Text = "&Play Back"
        Me.BtnPlayback.UseVisualStyleBackColor = True
        '
        'ButOptions
        '
        Me.ButOptions.Location = New System.Drawing.Point(279, 20)
        Me.ButOptions.Name = "ButOptions"
        Me.ButOptions.Size = New System.Drawing.Size(85, 24)
        Me.ButOptions.TabIndex = 4
        Me.ButOptions.Text = "&Options"
        Me.ButOptions.UseVisualStyleBackColor = True
        '
        'TmrStop
        '
        Me.TmrStop.Interval = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.GroupBox1.Controls.Add(Me.btnRecord)
        Me.GroupBox1.Controls.Add(Me.BtnStop)
        Me.GroupBox1.Controls.Add(Me.BtnPlayback)
        Me.GroupBox1.Controls.Add(Me.ButOptions)
        Me.GroupBox1.Location = New System.Drawing.Point(764, 233)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(372, 52)
        Me.GroupBox1.TabIndex = 24
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Macro Recorder"
        Me.GroupBox1.Visible = False
        '
        'EditRightClickList
        '
        Me.EditRightClickList.AccessibleName = ""
        Me.EditRightClickList.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lstviewStripRun, Me.CompileToExeFileToolStripMenuItem, Me.ToolStripEditCommand, Me.ToolStripSeparator14, Me.lstviewStripCut, Me.lstviewStripCopy, Me.lstviewStripPaste, Me.ToolStripSeparator13, Me.InsertToolStripMenuItem, Me.lstviewStripDelete, Me.ToolStripSeparator15, Me.lstviewStripMoveUp, Me.lstviewStripMoveDown, Me.ToolStripSeparator16, Me.SelectAllToolStripMenuItem, Me.DeselectAllToolStripMenuItem})
        Me.EditRightClickList.Name = "EditRightClickList"
        Me.EditRightClickList.Size = New System.Drawing.Size(216, 292)
        '
        'lstviewStripRun
        '
        Me.lstviewStripRun.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RunTheMacroToolStripMenuItem, Me.RunSelectionToolStripMenuItem})
        Me.lstviewStripRun.Image = CType(resources.GetObject("lstviewStripRun.Image"), System.Drawing.Image)
        Me.lstviewStripRun.Name = "lstviewStripRun"
        Me.lstviewStripRun.ShortcutKeyDisplayString = ""
        Me.lstviewStripRun.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
        Me.lstviewStripRun.Size = New System.Drawing.Size(215, 22)
        Me.lstviewStripRun.Text = "Playback"
        '
        'RunTheMacroToolStripMenuItem
        '
        Me.RunTheMacroToolStripMenuItem.Name = "RunTheMacroToolStripMenuItem"
        Me.RunTheMacroToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.RunTheMacroToolStripMenuItem.Text = "Playback The Macro"
        Me.RunTheMacroToolStripMenuItem.Visible = False
        '
        'RunSelectionToolStripMenuItem
        '
        Me.RunSelectionToolStripMenuItem.Name = "RunSelectionToolStripMenuItem"
        Me.RunSelectionToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.RunSelectionToolStripMenuItem.Text = "Playback Selection"
        Me.RunSelectionToolStripMenuItem.Visible = False
        '
        'CompileToExeFileToolStripMenuItem
        '
        Me.CompileToExeFileToolStripMenuItem.Name = "CompileToExeFileToolStripMenuItem"
        Me.CompileToExeFileToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.CompileToExeFileToolStripMenuItem.Size = New System.Drawing.Size(215, 22)
        Me.CompileToExeFileToolStripMenuItem.Text = "Compile to exe file"
        '
        'ToolStripEditCommand
        '
        Me.ToolStripEditCommand.Image = CType(resources.GetObject("ToolStripEditCommand.Image"), System.Drawing.Image)
        Me.ToolStripEditCommand.Name = "ToolStripEditCommand"
        Me.ToolStripEditCommand.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.D), System.Windows.Forms.Keys)
        Me.ToolStripEditCommand.Size = New System.Drawing.Size(215, 22)
        Me.ToolStripEditCommand.Text = "Edit "
        '
        'ToolStripSeparator14
        '
        Me.ToolStripSeparator14.Name = "ToolStripSeparator14"
        Me.ToolStripSeparator14.Size = New System.Drawing.Size(212, 6)
        '
        'lstviewStripCut
        '
        Me.lstviewStripCut.Image = CType(resources.GetObject("lstviewStripCut.Image"), System.Drawing.Image)
        Me.lstviewStripCut.Name = "lstviewStripCut"
        Me.lstviewStripCut.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.lstviewStripCut.Size = New System.Drawing.Size(215, 22)
        Me.lstviewStripCut.Text = "Cut"
        '
        'lstviewStripCopy
        '
        Me.lstviewStripCopy.Image = CType(resources.GetObject("lstviewStripCopy.Image"), System.Drawing.Image)
        Me.lstviewStripCopy.Name = "lstviewStripCopy"
        Me.lstviewStripCopy.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.lstviewStripCopy.Size = New System.Drawing.Size(215, 22)
        Me.lstviewStripCopy.Text = "Copy"
        '
        'lstviewStripPaste
        '
        Me.lstviewStripPaste.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PasteAfterToolStripMenuItem, Me.PasteBeforeToolStripMenuItem, Me.ReplaceToolStripMenuItem})
        Me.lstviewStripPaste.Image = CType(resources.GetObject("lstviewStripPaste.Image"), System.Drawing.Image)
        Me.lstviewStripPaste.Name = "lstviewStripPaste"
        Me.lstviewStripPaste.Size = New System.Drawing.Size(215, 22)
        Me.lstviewStripPaste.Text = "Paste"
        '
        'PasteAfterToolStripMenuItem
        '
        Me.PasteAfterToolStripMenuItem.Name = "PasteAfterToolStripMenuItem"
        Me.PasteAfterToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.PasteAfterToolStripMenuItem.Size = New System.Drawing.Size(212, 22)
        Me.PasteAfterToolStripMenuItem.Text = "Paste After"
        '
        'PasteBeforeToolStripMenuItem
        '
        Me.PasteBeforeToolStripMenuItem.Name = "PasteBeforeToolStripMenuItem"
        Me.PasteBeforeToolStripMenuItem.ShortcutKeys = CType(((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Shift) _
            Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.PasteBeforeToolStripMenuItem.Size = New System.Drawing.Size(212, 22)
        Me.PasteBeforeToolStripMenuItem.Text = "Paste Before"
        '
        'ReplaceToolStripMenuItem
        '
        Me.ReplaceToolStripMenuItem.Name = "ReplaceToolStripMenuItem"
        Me.ReplaceToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
        Me.ReplaceToolStripMenuItem.Size = New System.Drawing.Size(212, 22)
        Me.ReplaceToolStripMenuItem.Text = "Replace"
        '
        'ToolStripSeparator13
        '
        Me.ToolStripSeparator13.Name = "ToolStripSeparator13"
        Me.ToolStripSeparator13.Size = New System.Drawing.Size(212, 6)
        '
        'InsertToolStripMenuItem
        '
        Me.InsertToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BackgroundWindowTitleToolStripMenuItem, Me.DelayToolStripMenuItem, Me.MouseToolStripMenuItem, Me.KeyToolStripMenuItem, Me.MoreToolStripMenuItem})
        Me.InsertToolStripMenuItem.Image = CType(resources.GetObject("InsertToolStripMenuItem.Image"), System.Drawing.Image)
        Me.InsertToolStripMenuItem.Name = "InsertToolStripMenuItem"
        Me.InsertToolStripMenuItem.Size = New System.Drawing.Size(215, 22)
        Me.InsertToolStripMenuItem.Text = "Insert"
        '
        'BackgroundWindowTitleToolStripMenuItem
        '
        Me.BackgroundWindowTitleToolStripMenuItem.Image = CType(resources.GetObject("BackgroundWindowTitleToolStripMenuItem.Image"), System.Drawing.Image)
        Me.BackgroundWindowTitleToolStripMenuItem.Name = "BackgroundWindowTitleToolStripMenuItem"
        Me.BackgroundWindowTitleToolStripMenuItem.Size = New System.Drawing.Size(185, 22)
        Me.BackgroundWindowTitleToolStripMenuItem.Text = "Background Window"
        '
        'DelayToolStripMenuItem
        '
        Me.DelayToolStripMenuItem.Image = CType(resources.GetObject("DelayToolStripMenuItem.Image"), System.Drawing.Image)
        Me.DelayToolStripMenuItem.Name = "DelayToolStripMenuItem"
        Me.DelayToolStripMenuItem.Size = New System.Drawing.Size(185, 22)
        Me.DelayToolStripMenuItem.Text = "Wait"
        '
        'MouseToolStripMenuItem
        '
        Me.MouseToolStripMenuItem.Image = CType(resources.GetObject("MouseToolStripMenuItem.Image"), System.Drawing.Image)
        Me.MouseToolStripMenuItem.Name = "MouseToolStripMenuItem"
        Me.MouseToolStripMenuItem.Size = New System.Drawing.Size(185, 22)
        Me.MouseToolStripMenuItem.Text = "Mouse Action"
        '
        'KeyToolStripMenuItem
        '
        Me.KeyToolStripMenuItem.Image = CType(resources.GetObject("KeyToolStripMenuItem.Image"), System.Drawing.Image)
        Me.KeyToolStripMenuItem.Name = "KeyToolStripMenuItem"
        Me.KeyToolStripMenuItem.Size = New System.Drawing.Size(185, 22)
        Me.KeyToolStripMenuItem.Text = "Keystroke"
        '
        'MoreToolStripMenuItem
        '
        Me.MoreToolStripMenuItem.Name = "MoreToolStripMenuItem"
        Me.MoreToolStripMenuItem.Size = New System.Drawing.Size(185, 22)
        Me.MoreToolStripMenuItem.Text = "More..."
        '
        'lstviewStripDelete
        '
        Me.lstviewStripDelete.Image = CType(resources.GetObject("lstviewStripDelete.Image"), System.Drawing.Image)
        Me.lstviewStripDelete.Name = "lstviewStripDelete"
        Me.lstviewStripDelete.ShortcutKeys = System.Windows.Forms.Keys.Delete
        Me.lstviewStripDelete.Size = New System.Drawing.Size(215, 22)
        Me.lstviewStripDelete.Text = "Delete"
        '
        'ToolStripSeparator15
        '
        Me.ToolStripSeparator15.Name = "ToolStripSeparator15"
        Me.ToolStripSeparator15.Size = New System.Drawing.Size(212, 6)
        '
        'lstviewStripMoveUp
        '
        Me.lstviewStripMoveUp.Image = CType(resources.GetObject("lstviewStripMoveUp.Image"), System.Drawing.Image)
        Me.lstviewStripMoveUp.Name = "lstviewStripMoveUp"
        Me.lstviewStripMoveUp.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Up), System.Windows.Forms.Keys)
        Me.lstviewStripMoveUp.Size = New System.Drawing.Size(215, 22)
        Me.lstviewStripMoveUp.Text = "Move Up"
        '
        'lstviewStripMoveDown
        '
        Me.lstviewStripMoveDown.Image = CType(resources.GetObject("lstviewStripMoveDown.Image"), System.Drawing.Image)
        Me.lstviewStripMoveDown.Name = "lstviewStripMoveDown"
        Me.lstviewStripMoveDown.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Down), System.Windows.Forms.Keys)
        Me.lstviewStripMoveDown.Size = New System.Drawing.Size(215, 22)
        Me.lstviewStripMoveDown.Text = "Move Down"
        '
        'ToolStripSeparator16
        '
        Me.ToolStripSeparator16.Name = "ToolStripSeparator16"
        Me.ToolStripSeparator16.Size = New System.Drawing.Size(212, 6)
        '
        'SelectAllToolStripMenuItem
        '
        Me.SelectAllToolStripMenuItem.Name = "SelectAllToolStripMenuItem"
        Me.SelectAllToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.SelectAllToolStripMenuItem.Size = New System.Drawing.Size(215, 22)
        Me.SelectAllToolStripMenuItem.Text = "Select All"
        '
        'DeselectAllToolStripMenuItem
        '
        Me.DeselectAllToolStripMenuItem.Name = "DeselectAllToolStripMenuItem"
        Me.DeselectAllToolStripMenuItem.ShortcutKeys = CType(((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Shift) _
            Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.DeselectAllToolStripMenuItem.Size = New System.Drawing.Size(215, 22)
        Me.DeselectAllToolStripMenuItem.Text = "Deselect All"
        '
        'TmrGenLoop
        '
        Me.TmrGenLoop.Enabled = True
        Me.TmrGenLoop.Interval = 500
        '
        'toolstripRecord
        '
        Me.toolstripRecord.Image = CType(resources.GetObject("toolstripRecord.Image"), System.Drawing.Image)
        Me.toolstripRecord.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.toolstripRecord.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.toolstripRecord.Name = "toolstripRecord"
        Me.toolstripRecord.Size = New System.Drawing.Size(72, 52)
        Me.toolstripRecord.Text = "&Record"
        Me.toolstripRecord.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.toolstripRecord.ToolTipText = "Start recording a new macro"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 55)
        '
        'toolstripStopRecord
        '
        Me.toolstripStopRecord.Image = CType(resources.GetObject("toolstripStopRecord.Image"), System.Drawing.Image)
        Me.toolstripStopRecord.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.toolstripStopRecord.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.toolstripStopRecord.Name = "toolstripStopRecord"
        Me.toolstripStopRecord.Size = New System.Drawing.Size(107, 52)
        Me.toolstripStopRecord.Text = "&Stop Record"
        Me.toolstripStopRecord.ToolTipText = "Stop recording and save macro"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 55)
        '
        'toolstripOptions
        '
        Me.toolstripOptions.Image = CType(resources.GetObject("toolstripOptions.Image"), System.Drawing.Image)
        Me.toolstripOptions.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.toolstripOptions.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.toolstripOptions.Name = "toolstripOptions"
        Me.toolstripOptions.Size = New System.Drawing.Size(85, 52)
        Me.toolstripOptions.Text = "S&ettings"
        Me.toolstripOptions.ToolTipText = "Perfect Macro Recorder Settings"
        '
        'toolstripPlyBack
        '
        Me.toolstripPlyBack.Image = CType(resources.GetObject("toolstripPlyBack.Image"), System.Drawing.Image)
        Me.toolstripPlyBack.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.toolstripPlyBack.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.toolstripPlyBack.Name = "toolstripPlyBack"
        Me.toolstripPlyBack.Size = New System.Drawing.Size(77, 52)
        Me.toolstripPlyBack.Text = "&Play Back"
        Me.toolstripPlyBack.Visible = False
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(6, 55)
        Me.ToolStripSeparator5.Visible = False
        '
        'toolstripHelp
        '
        Me.toolstripHelp.Image = CType(resources.GetObject("toolstripHelp.Image"), System.Drawing.Image)
        Me.toolstripHelp.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.toolstripHelp.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.toolstripHelp.Name = "toolstripHelp"
        Me.toolstripHelp.Size = New System.Drawing.Size(68, 52)
        Me.toolstripHelp.Text = "&Help"
        '
        'ToolStripHorzButtons
        '
        Me.ToolStripHorzButtons.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.toolstripRecord, Me.ToolStripSeparator2, Me.toolstripStopRecord, Me.ToolStripSeparator3, Me.toolstripOptions, Me.ToolStripSeparator5, Me.toolstripHelp, Me.toolstripPlyBack, Me.ToolStripButtonRegister})
        Me.ToolStripHorzButtons.Location = New System.Drawing.Point(0, 0)
        Me.ToolStripHorzButtons.Name = "ToolStripHorzButtons"
        Me.ToolStripHorzButtons.Size = New System.Drawing.Size(476, 55)
        Me.ToolStripHorzButtons.TabIndex = 25
        Me.ToolStripHorzButtons.Text = "ToolStrip1"
        '
        'ToolStripButtonRegister
        '
        Me.ToolStripButtonRegister.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonRegister.Image = Global.CuteMacro.My.Resources.Resources.label_blue_buy5
        Me.ToolStripButtonRegister.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ToolStripButtonRegister.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripButtonRegister.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButtonRegister.Name = "ToolStripButtonRegister"
        Me.ToolStripButtonRegister.Size = New System.Drawing.Size(52, 52)
        Me.ToolStripButtonRegister.Text = "Register"
        Me.ToolStripButtonRegister.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'ToolStripNew
        '
        Me.ToolStripNew.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripNew.Image = CType(resources.GetObject("ToolStripNew.Image"), System.Drawing.Image)
        Me.ToolStripNew.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripNew.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripNew.Name = "ToolStripNew"
        Me.ToolStripNew.Size = New System.Drawing.Size(50, 36)
        Me.ToolStripNew.Text = "&New"
        Me.ToolStripNew.ToolTipText = "Clear macro script list"
        '
        'ToolStripOpenMacro
        '
        Me.ToolStripOpenMacro.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripOpenMacro.Image = CType(resources.GetObject("ToolStripOpenMacro.Image"), System.Drawing.Image)
        Me.ToolStripOpenMacro.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripOpenMacro.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripOpenMacro.Name = "ToolStripOpenMacro"
        Me.ToolStripOpenMacro.Size = New System.Drawing.Size(50, 36)
        Me.ToolStripOpenMacro.Text = "&Open"
        Me.ToolStripOpenMacro.ToolTipText = "Load a macro file (*.pmac or *.exe) to script list"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(50, 6)
        '
        'CutToolStripButton
        '
        Me.CutToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.CutToolStripButton.Image = CType(resources.GetObject("CutToolStripButton.Image"), System.Drawing.Image)
        Me.CutToolStripButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.CutToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.CutToolStripButton.Name = "CutToolStripButton"
        Me.CutToolStripButton.Size = New System.Drawing.Size(50, 52)
        Me.CutToolStripButton.Text = "C&ut"
        Me.CutToolStripButton.ToolTipText = "Cut selected command(s) in the script list"
        '
        'CopyToolStripButton
        '
        Me.CopyToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.CopyToolStripButton.Image = CType(resources.GetObject("CopyToolStripButton.Image"), System.Drawing.Image)
        Me.CopyToolStripButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.CopyToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.CopyToolStripButton.Name = "CopyToolStripButton"
        Me.CopyToolStripButton.Size = New System.Drawing.Size(50, 36)
        Me.CopyToolStripButton.Text = "&Copy"
        Me.CopyToolStripButton.ToolTipText = "Copy selected command(s) in the script list"
        '
        'PasteToolStripButton
        '
        Me.PasteToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.PasteToolStripButton.Image = CType(resources.GetObject("PasteToolStripButton.Image"), System.Drawing.Image)
        Me.PasteToolStripButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.PasteToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PasteToolStripButton.Name = "PasteToolStripButton"
        Me.PasteToolStripButton.Size = New System.Drawing.Size(50, 36)
        Me.PasteToolStripButton.Text = "&Paste"
        Me.PasteToolStripButton.ToolTipText = "Paste command(s) in the script list"
        '
        'ToolStripDelItems
        '
        Me.ToolStripDelItems.Image = CType(resources.GetObject("ToolStripDelItems.Image"), System.Drawing.Image)
        Me.ToolStripDelItems.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripDelItems.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripDelItems.Name = "ToolStripDelItems"
        Me.ToolStripDelItems.Size = New System.Drawing.Size(50, 36)
        Me.ToolStripDelItems.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolStripDelItems.ToolTipText = "Delete selected command(s) in the list"
        '
        'ToolStripSveMac
        '
        Me.ToolStripSveMac.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripSveMac.Image = CType(resources.GetObject("ToolStripSveMac.Image"), System.Drawing.Image)
        Me.ToolStripSveMac.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripSveMac.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripSveMac.Name = "ToolStripSveMac"
        Me.ToolStripSveMac.Size = New System.Drawing.Size(50, 36)
        Me.ToolStripSveMac.Text = "&Save"
        Me.ToolStripSveMac.ToolTipText = "Save script list to a macro file"
        '
        'ToolStripAddKeys
        '
        Me.ToolStripAddKeys.Image = CType(resources.GetObject("ToolStripAddKeys.Image"), System.Drawing.Image)
        Me.ToolStripAddKeys.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripAddKeys.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripAddKeys.Name = "ToolStripAddKeys"
        Me.ToolStripAddKeys.Size = New System.Drawing.Size(50, 36)
        Me.ToolStripAddKeys.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolStripAddKeys.ToolTipText = "Add Keystroke action"
        '
        'ToolStripAddDelay
        '
        Me.ToolStripAddDelay.Image = CType(resources.GetObject("ToolStripAddDelay.Image"), System.Drawing.Image)
        Me.ToolStripAddDelay.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripAddDelay.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripAddDelay.Name = "ToolStripAddDelay"
        Me.ToolStripAddDelay.Size = New System.Drawing.Size(50, 36)
        Me.ToolStripAddDelay.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolStripAddDelay.ToolTipText = "Add delay"
        '
        'ToolStripAddMouseEve
        '
        Me.ToolStripAddMouseEve.Image = CType(resources.GetObject("ToolStripAddMouseEve.Image"), System.Drawing.Image)
        Me.ToolStripAddMouseEve.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripAddMouseEve.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripAddMouseEve.Name = "ToolStripAddMouseEve"
        Me.ToolStripAddMouseEve.Size = New System.Drawing.Size(50, 36)
        Me.ToolStripAddMouseEve.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolStripAddMouseEve.ToolTipText = "Add mouse action"
        '
        'ToolStripRun
        '
        Me.ToolStripRun.Image = Global.CuteMacro.My.Resources.Resources.icon_48
        Me.ToolStripRun.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripRun.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripRun.Name = "ToolStripRun"
        Me.ToolStripRun.Size = New System.Drawing.Size(50, 52)
        Me.ToolStripRun.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolStripRun.ToolTipText = "Playback all/selected commands in the script list"
        '
        'ToolStripSeparator8
        '
        Me.ToolStripSeparator8.Name = "ToolStripSeparator8"
        Me.ToolStripSeparator8.Size = New System.Drawing.Size(50, 6)
        '
        'ToolStripLeftVertButtons
        '
        Me.ToolStripLeftVertButtons.Dock = System.Windows.Forms.DockStyle.Left
        Me.ToolStripLeftVertButtons.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripNew, Me.ToolStripOpenMacro, Me.ToolStripSveMac, Me.ToolStripCompile, Me.ToolStripSeparator4, Me.ToolStripRun, Me.ToolStripSeparator8, Me.CutToolStripButton, Me.CopyToolStripButton, Me.PasteToolStripButton, Me.ToolStripDelItems, Me.ToolStripSeparator6, Me.ToolStripBackGround, Me.ToolStripAddDelay, Me.ToolStripAddMouseEve, Me.ToolStripAddKeys, Me.ToolStripAddMore})
        Me.ToolStripLeftVertButtons.Location = New System.Drawing.Point(0, 55)
        Me.ToolStripLeftVertButtons.Name = "ToolStripLeftVertButtons"
        Me.ToolStripLeftVertButtons.Size = New System.Drawing.Size(53, 553)
        Me.ToolStripLeftVertButtons.TabIndex = 27
        Me.ToolStripLeftVertButtons.Text = "ToolStrip2"
        '
        'ToolStripCompile
        '
        Me.ToolStripCompile.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripCompile.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripCompile.ForeColor = System.Drawing.Color.Black
        Me.ToolStripCompile.Image = CType(resources.GetObject("ToolStripCompile.Image"), System.Drawing.Image)
        Me.ToolStripCompile.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripCompile.Name = "ToolStripCompile"
        Me.ToolStripCompile.Size = New System.Drawing.Size(50, 23)
        Me.ToolStripCompile.Text = "Exe"
        Me.ToolStripCompile.ToolTipText = "Compile script list to exe file"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(50, 6)
        '
        'ToolStripBackGround
        '
        Me.ToolStripBackGround.Image = CType(resources.GetObject("ToolStripBackGround.Image"), System.Drawing.Image)
        Me.ToolStripBackGround.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripBackGround.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripBackGround.Name = "ToolStripBackGround"
        Me.ToolStripBackGround.Size = New System.Drawing.Size(50, 36)
        Me.ToolStripBackGround.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolStripBackGround.ToolTipText = "Add background window to script list"
        '
        'ToolStripAddMore
        '
        Me.ToolStripAddMore.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripAddMore.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripAddMore.ForeColor = System.Drawing.Color.Black
        Me.ToolStripAddMore.Image = CType(resources.GetObject("ToolStripAddMore.Image"), System.Drawing.Image)
        Me.ToolStripAddMore.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripAddMore.Name = "ToolStripAddMore"
        Me.ToolStripAddMore.Size = New System.Drawing.Size(54, 18)
        Me.ToolStripAddMore.Text = "More..."
        Me.ToolStripAddMore.ToolTipText = "Compile the macro to an exe file"
        Me.ToolStripAddMore.Visible = False
        '
        'StatusBar
        '
        Me.StatusBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Status_LoadingLabel})
        Me.StatusBar.Location = New System.Drawing.Point(0, 608)
        Me.StatusBar.Name = "StatusBar"
        Me.StatusBar.Size = New System.Drawing.Size(476, 22)
        Me.StatusBar.TabIndex = 28
        Me.StatusBar.Text = "StatusStrip1"
        '
        'Status_LoadingLabel
        '
        Me.Status_LoadingLabel.Name = "Status_LoadingLabel"
        Me.Status_LoadingLabel.Size = New System.Drawing.Size(0, 17)
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.lstviwMacEdit)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(53, 55)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(423, 553)
        Me.Panel1.TabIndex = 30
        '
        'lstviwMacEdit
        '
        Me.lstviwMacEdit.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColCommand, Me.Col2, Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7, Me.ColumnHeader8, Me.ColumnHeader9, Me.ColumnHeader10, Me.ColumnHeader11})
        Me.lstviwMacEdit.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstviwMacEdit.FullRowSelect = True
        Me.lstviwMacEdit.GridLines = True
        Me.lstviwMacEdit.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.lstviwMacEdit.Location = New System.Drawing.Point(0, 0)
        Me.lstviwMacEdit.Name = "lstviwMacEdit"
        Me.lstviwMacEdit.Size = New System.Drawing.Size(423, 553)
        Me.lstviwMacEdit.TabIndex = 31
        Me.lstviwMacEdit.UseCompatibleStateImageBehavior = False
        Me.lstviwMacEdit.View = System.Windows.Forms.View.Details
        '
        'ColCommand
        '
        Me.ColCommand.Text = ""
        Me.ColCommand.Width = 130
        '
        'Col2
        '
        Me.Col2.Text = ""
        Me.Col2.Width = 140
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = ""
        Me.ColumnHeader1.Width = 141
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = ""
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = ""
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = ""
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = ""
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = ""
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = ""
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = ""
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = ""
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = ""
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = ""
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(476, 630)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ToolStripLeftVertButtons)
        Me.Controls.Add(Me.ToolStripHorzButtons)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnGetNow)
        Me.Controls.Add(Me.btnSveMac)
        Me.Controls.Add(Me.ButMacEdit)
        Me.Controls.Add(Me.btnNewMac)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.StatusBar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Perfect Macro Recorder"
        Me.SystemTrayMenu.ResumeLayout(False)
        Me.HelpButtonContextMenu.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.EditRightClickList.ResumeLayout(False)
        Me.ToolStripHorzButtons.ResumeLayout(False)
        Me.ToolStripHorzButtons.PerformLayout()
        Me.ToolStripLeftVertButtons.ResumeLayout(False)
        Me.ToolStripLeftVertButtons.PerformLayout()
        Me.StatusBar.ResumeLayout(False)
        Me.StatusBar.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ShowSave As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ShowOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TrayIcon As System.Windows.Forms.NotifyIcon
    Friend WithEvents SystemTrayMenu As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents RecordSystemTrayMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PlayBackSystemTrayMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OptionsSystemTrayMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents AboutSystemTrayMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitSystemTrayMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpButtonContextMenu As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents WebPageToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ButMacEdit As System.Windows.Forms.Button
    Friend WithEvents BtnStop As System.Windows.Forms.Button
    Friend WithEvents btnRecord As System.Windows.Forms.Button
    Friend WithEvents BtnPlayback As System.Windows.Forms.Button
    Friend WithEvents ButOptions As System.Windows.Forms.Button
    Friend WithEvents WebPageToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TmrStop As System.Windows.Forms.Timer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnRun As System.Windows.Forms.Button
    Friend WithEvents btnSveMac As System.Windows.Forms.Button
    Friend WithEvents btnNewMac As System.Windows.Forms.Button
    Friend WithEvents btnGetNow As System.Windows.Forms.Button
    Friend WithEvents ShowCuteMacroToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditRightClickList As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents lstviewStripCut As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lstviewStripCopy As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lstviewStripPaste As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lstviewStripDelete As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lstviewStripMoveUp As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lstviewStripMoveDown As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lstviewStripRun As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PasteAfterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PasteBeforeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReplaceToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator14 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator13 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator15 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator16 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents InsertToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MouseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents KeyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DelayToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SelectAllToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeselectAllToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RunTheMacroToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RunSelectionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents TmrGenLoop As System.Windows.Forms.Timer
    Friend WithEvents toolstripRecord As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents toolstripStopRecord As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents toolstripOptions As System.Windows.Forms.ToolStripButton
    Friend WithEvents toolstripPlyBack As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents toolstripHelp As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripHorzButtons As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripNew As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripOpenMacro As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents CutToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents CopyToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents PasteToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripDelItems As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSveMac As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripAddKeys As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripAddDelay As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripAddMouseEve As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripRun As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator8 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripLeftVertButtons As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripSeparator23 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator22 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator21 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripEditCommand As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripCompile As System.Windows.Forms.ToolStripButton
    Friend WithEvents CompileToExeFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator6 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripAddMore As System.Windows.Forms.ToolStripButton
    Friend WithEvents StopToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripBackGround As System.Windows.Forms.ToolStripButton
    Friend WithEvents BackgroundWindowTitleToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MoreToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusBar As System.Windows.Forms.StatusStrip
	Friend WithEvents Status_LoadingLabel As System.Windows.Forms.ToolStripStatusLabel
	Friend WithEvents ToolStripButtonRegister As System.Windows.Forms.ToolStripButton
	Friend WithEvents Panel1 As System.Windows.Forms.Panel
	Friend WithEvents lstviwMacEdit As System.Windows.Forms.ListView
	Friend WithEvents ColCommand As System.Windows.Forms.ColumnHeader
	Friend WithEvents Col2 As System.Windows.Forms.ColumnHeader
	Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
	Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
	Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
	Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
	Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
	Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
	Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
	Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
	Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
	Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
	Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader

End Class
