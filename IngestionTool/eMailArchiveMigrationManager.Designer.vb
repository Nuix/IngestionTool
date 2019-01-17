<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class eMailArchiveMigrationManager
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(eMailArchiveMigrationManager))
        Me.btnExtractFromLegacy = New System.Windows.Forms.Button()
        Me.btnExtractFromO365 = New System.Windows.Forms.Button()
        Me.btnUploadtoO365 = New System.Windows.Forms.Button()
        Me.btnNuixSettings = New System.Windows.Forms.Button()
        Me.btnSourceDataConversion = New System.Windows.Forms.Button()
        Me.btnCustodianListGenerator = New System.Windows.Forms.Button()
        Me.SettingsToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.NuixLogoPicture = New System.Windows.Forms.PictureBox()
        Me.menuNEAMM = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StartEWSExtraction = New System.Windows.Forms.ToolStripMenuItem()
        Me.GlobalSettingsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenLogDirectoryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutNEAMMToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StartEWSIngestion = New System.Windows.Forms.ToolStripMenuItem()
        Me.StartArchiveExtraction = New System.Windows.Forms.ToolStripMenuItem()
        Me.StartEmailConversion = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        CType(Me.NuixLogoPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuNEAMM.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExtractFromLegacy
        '
        Me.btnExtractFromLegacy.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnExtractFromLegacy.BackColor = System.Drawing.Color.White
        Me.btnExtractFromLegacy.Image = CType(resources.GetObject("btnExtractFromLegacy.Image"), System.Drawing.Image)
        Me.btnExtractFromLegacy.Location = New System.Drawing.Point(98, 77)
        Me.btnExtractFromLegacy.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExtractFromLegacy.Name = "btnExtractFromLegacy"
        Me.btnExtractFromLegacy.Size = New System.Drawing.Size(167, 167)
        Me.btnExtractFromLegacy.TabIndex = 0
        Me.btnExtractFromLegacy.UseVisualStyleBackColor = False
        '
        'btnExtractFromO365
        '
        Me.btnExtractFromO365.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnExtractFromO365.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnExtractFromO365.BackColor = System.Drawing.Color.White
        Me.btnExtractFromO365.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.btnExtractFromO365.FlatAppearance.BorderSize = 5
        Me.btnExtractFromO365.Image = CType(resources.GetObject("btnExtractFromO365.Image"), System.Drawing.Image)
        Me.btnExtractFromO365.Location = New System.Drawing.Point(269, 246)
        Me.btnExtractFromO365.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExtractFromO365.Name = "btnExtractFromO365"
        Me.btnExtractFromO365.Size = New System.Drawing.Size(167, 167)
        Me.btnExtractFromO365.TabIndex = 1
        Me.btnExtractFromO365.UseVisualStyleBackColor = False
        '
        'btnUploadtoO365
        '
        Me.btnUploadtoO365.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnUploadtoO365.BackColor = System.Drawing.Color.White
        Me.btnUploadtoO365.Image = CType(resources.GetObject("btnUploadtoO365.Image"), System.Drawing.Image)
        Me.btnUploadtoO365.Location = New System.Drawing.Point(98, 246)
        Me.btnUploadtoO365.Margin = New System.Windows.Forms.Padding(2)
        Me.btnUploadtoO365.Name = "btnUploadtoO365"
        Me.btnUploadtoO365.Size = New System.Drawing.Size(167, 167)
        Me.btnUploadtoO365.TabIndex = 2
        Me.btnUploadtoO365.UseVisualStyleBackColor = False
        '
        'btnNuixSettings
        '
        Me.btnNuixSettings.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnNuixSettings.BackColor = System.Drawing.Color.White
        Me.btnNuixSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnNuixSettings.Image = CType(resources.GetObject("btnNuixSettings.Image"), System.Drawing.Image)
        Me.btnNuixSettings.Location = New System.Drawing.Point(356, 434)
        Me.btnNuixSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.btnNuixSettings.Name = "btnNuixSettings"
        Me.btnNuixSettings.Size = New System.Drawing.Size(80, 55)
        Me.btnNuixSettings.TabIndex = 3
        Me.btnNuixSettings.UseVisualStyleBackColor = False
        '
        'btnSourceDataConversion
        '
        Me.btnSourceDataConversion.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnSourceDataConversion.BackColor = System.Drawing.Color.White
        Me.btnSourceDataConversion.Image = CType(resources.GetObject("btnSourceDataConversion.Image"), System.Drawing.Image)
        Me.btnSourceDataConversion.Location = New System.Drawing.Point(269, 77)
        Me.btnSourceDataConversion.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSourceDataConversion.Name = "btnSourceDataConversion"
        Me.btnSourceDataConversion.Size = New System.Drawing.Size(167, 167)
        Me.btnSourceDataConversion.TabIndex = 4
        Me.btnSourceDataConversion.UseVisualStyleBackColor = False
        '
        'btnCustodianListGenerator
        '
        Me.btnCustodianListGenerator.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnCustodianListGenerator.BackColor = System.Drawing.Color.White
        Me.btnCustodianListGenerator.Location = New System.Drawing.Point(88, 420)
        Me.btnCustodianListGenerator.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCustodianListGenerator.Name = "btnCustodianListGenerator"
        Me.btnCustodianListGenerator.Size = New System.Drawing.Size(88, 82)
        Me.btnCustodianListGenerator.TabIndex = 5
        Me.btnCustodianListGenerator.Text = "Custodian List Generator (Prototype)"
        Me.btnCustodianListGenerator.UseVisualStyleBackColor = False
        '
        'SettingsToolTip
        '
        Me.SettingsToolTip.IsBalloon = True
        '
        'NuixLogoPicture
        '
        Me.NuixLogoPicture.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.NuixLogoPicture.ErrorImage = CType(resources.GetObject("NuixLogoPicture.ErrorImage"), System.Drawing.Image)
        Me.NuixLogoPicture.Image = CType(resources.GetObject("NuixLogoPicture.Image"), System.Drawing.Image)
        Me.NuixLogoPicture.InitialImage = CType(resources.GetObject("NuixLogoPicture.InitialImage"), System.Drawing.Image)
        Me.NuixLogoPicture.Location = New System.Drawing.Point(25, 441)
        Me.NuixLogoPicture.Name = "NuixLogoPicture"
        Me.NuixLogoPicture.Size = New System.Drawing.Size(75, 100)
        Me.NuixLogoPicture.TabIndex = 6
        Me.NuixLogoPicture.TabStop = False
        '
        'menuNEAMM
        '
        Me.menuNEAMM.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.menuNEAMM.Location = New System.Drawing.Point(0, 0)
        Me.menuNEAMM.Name = "menuNEAMM"
        Me.menuNEAMM.Size = New System.Drawing.Size(534, 24)
        Me.menuNEAMM.TabIndex = 7
        Me.menuNEAMM.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StartArchiveExtraction, Me.StartEWSExtraction, Me.StartEWSIngestion, Me.StartEmailConversion, Me.ToolStripSeparator1, Me.GlobalSettingsToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'StartEWSExtraction
        '
        Me.StartEWSExtraction.Name = "StartEWSExtraction"
        Me.StartEWSExtraction.Size = New System.Drawing.Size(196, 22)
        Me.StartEWSExtraction.Text = "Start EWS Extraction"
        '
        'GlobalSettingsToolStripMenuItem
        '
        Me.GlobalSettingsToolStripMenuItem.Name = "GlobalSettingsToolStripMenuItem"
        Me.GlobalSettingsToolStripMenuItem.Size = New System.Drawing.Size(196, 22)
        Me.GlobalSettingsToolStripMenuItem.Text = "Global Settings"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(196, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenLogDirectoryToolStripMenuItem, Me.AboutNEAMMToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'OpenLogDirectoryToolStripMenuItem
        '
        Me.OpenLogDirectoryToolStripMenuItem.Name = "OpenLogDirectoryToolStripMenuItem"
        Me.OpenLogDirectoryToolStripMenuItem.Size = New System.Drawing.Size(177, 22)
        Me.OpenLogDirectoryToolStripMenuItem.Text = "Open Log Directory"
        '
        'AboutNEAMMToolStripMenuItem
        '
        Me.AboutNEAMMToolStripMenuItem.Name = "AboutNEAMMToolStripMenuItem"
        Me.AboutNEAMMToolStripMenuItem.Size = New System.Drawing.Size(177, 22)
        Me.AboutNEAMMToolStripMenuItem.Text = "About NEAMM"
        '
        'StartEWSIngestion
        '
        Me.StartEWSIngestion.Name = "StartEWSIngestion"
        Me.StartEWSIngestion.Size = New System.Drawing.Size(196, 22)
        Me.StartEWSIngestion.Text = "Start EWS Ingestion"
        '
        'StartArchiveExtraction
        '
        Me.StartArchiveExtraction.Name = "StartArchiveExtraction"
        Me.StartArchiveExtraction.Size = New System.Drawing.Size(196, 22)
        Me.StartArchiveExtraction.Text = "Start Archive Extraction"
        '
        'StartEmailConversion
        '
        Me.StartEmailConversion.Name = "StartEmailConversion"
        Me.StartEmailConversion.Size = New System.Drawing.Size(196, 22)
        Me.StartEmailConversion.Text = "Start Email Conversion"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(193, 6)
        '
        'eMailArchiveMigrationManager
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(534, 561)
        Me.Controls.Add(Me.NuixLogoPicture)
        Me.Controls.Add(Me.btnCustodianListGenerator)
        Me.Controls.Add(Me.btnSourceDataConversion)
        Me.Controls.Add(Me.btnNuixSettings)
        Me.Controls.Add(Me.btnUploadtoO365)
        Me.Controls.Add(Me.btnExtractFromO365)
        Me.Controls.Add(Me.btnExtractFromLegacy)
        Me.Controls.Add(Me.menuNEAMM)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.menuNEAMM
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(550, 600)
        Me.Name = "eMailArchiveMigrationManager"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Nuix Email Archive Migration Manager"
        CType(Me.NuixLogoPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuNEAMM.ResumeLayout(False)
        Me.menuNEAMM.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnExtractFromLegacy As System.Windows.Forms.Button
    Friend WithEvents btnExtractFromO365 As System.Windows.Forms.Button
    Friend WithEvents btnUploadtoO365 As System.Windows.Forms.Button
    Friend WithEvents btnNuixSettings As System.Windows.Forms.Button
    Friend WithEvents btnSourceDataConversion As System.Windows.Forms.Button
    Friend WithEvents btnCustodianListGenerator As System.Windows.Forms.Button
    Friend WithEvents SettingsToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents NuixLogoPicture As System.Windows.Forms.PictureBox
    Friend WithEvents menuNEAMM As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StartEWSExtraction As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GlobalSettingsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutNEAMMToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenLogDirectoryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StartArchiveExtraction As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StartEWSIngestion As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StartEmailConversion As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
End Class
