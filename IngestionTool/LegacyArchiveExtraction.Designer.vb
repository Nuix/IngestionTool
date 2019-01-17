<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LegacyArchiveExtraction
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(LegacyArchiveExtraction))
        Me.cboArchiveType = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.grpExtractionType = New System.Windows.Forms.GroupBox()
        Me.radLoose = New System.Windows.Forms.RadioButton()
        Me.radZipped = New System.Windows.Forms.RadioButton()
        Me.cboExtractionOutputType = New System.Windows.Forms.ComboBox()
        Me.grpAXSOneSISSettings = New System.Windows.Forms.GroupBox()
        Me.treeAXSOneSis = New System.Windows.Forms.TreeView()
        Me.ImgIcons = New System.Windows.Forms.ImageList(Me.components)
        Me.grpSQLArchiveSettings = New System.Windows.Forms.GroupBox()
        Me.lblAuthentication = New System.Windows.Forms.Label()
        Me.cboSecurityType = New System.Windows.Forms.ComboBox()
        Me.lblDomainCheck = New System.Windows.Forms.Label()
        Me.lblSQLPWCheck = New System.Windows.Forms.Label()
        Me.lblSQLUserNameCheck = New System.Windows.Forms.Label()
        Me.lblDBNameCheck = New System.Windows.Forms.Label()
        Me.lblHostCheck = New System.Windows.Forms.Label()
        Me.btnTestSQLConnection = New System.Windows.Forms.Button()
        Me.txtDomain = New System.Windows.Forms.TextBox()
        Me.lblDomain = New System.Windows.Forms.Label()
        Me.txtSQLInfo = New System.Windows.Forms.TextBox()
        Me.lblSQLInfo = New System.Windows.Forms.Label()
        Me.txtSQLUserName = New System.Windows.Forms.TextBox()
        Me.lblSQLUserName = New System.Windows.Forms.Label()
        Me.txtSQLDBName = New System.Windows.Forms.TextBox()
        Me.lblSQLDBName = New System.Windows.Forms.Label()
        Me.txtSQLPortNumber = New System.Windows.Forms.TextBox()
        Me.lblPortNumber = New System.Windows.Forms.Label()
        Me.txtSQLHostName = New System.Windows.Forms.TextBox()
        Me.lblServerName = New System.Windows.Forms.Label()
        Me.grpEMCSettings = New System.Windows.Forms.GroupBox()
        Me.chkAFEX = New System.Windows.Forms.CheckBox()
        Me.chkAFPST = New System.Windows.Forms.CheckBox()
        Me.chkAFSYS = New System.Windows.Forms.CheckBox()
        Me.lblExpandDLto = New System.Windows.Forms.Label()
        Me.cboExpandDLLocation = New System.Windows.Forms.ComboBox()
        Me.lblAddressFiltering = New System.Windows.Forms.Label()
        Me.grpEVSettings = New System.Windows.Forms.GroupBox()
        Me.btnEVUserListFileChooser = New System.Windows.Forms.Button()
        Me.txtEVUserList = New System.Windows.Forms.TextBox()
        Me.lblEVUserList = New System.Windows.Forms.Label()
        Me.chkSkipVaultStorePartitionErrors = New System.Windows.Forms.CheckBox()
        Me.chkSkipAdditionalSQLLookup = New System.Windows.Forms.CheckBox()
        Me.btnClearDates = New System.Windows.Forms.Button()
        Me.grpAXSOneSettings = New System.Windows.Forms.GroupBox()
        Me.chkAXSOneSkipSISLookups = New System.Windows.Forms.CheckBox()
        Me.grpEASSettings = New System.Windows.Forms.GroupBox()
        Me.txtDocServerID = New System.Windows.Forms.TextBox()
        Me.lblDocServerID = New System.Windows.Forms.Label()
        Me.txtCustodianListCSV = New System.Windows.Forms.TextBox()
        Me.lblCustodianListCSV = New System.Windows.Forms.Label()
        Me.btnCustodianListCSVChooser = New System.Windows.Forms.Button()
        Me.grpOutputFormat = New System.Windows.Forms.GroupBox()
        Me.radFlat = New System.Windows.Forms.RadioButton()
        Me.radUser = New System.Windows.Forms.RadioButton()
        Me.btnShowSettings = New System.Windows.Forms.Button()
        Me.grpSourceInformation = New System.Windows.Forms.GroupBox()
        Me.btnIPFile = New System.Windows.Forms.Button()
        Me.txtIPFile = New System.Windows.Forms.TextBox()
        Me.lblIPFile = New System.Windows.Forms.Label()
        Me.btnPEAFile = New System.Windows.Forms.Button()
        Me.txtPEAFile = New System.Windows.Forms.TextBox()
        Me.lblPEAFile = New System.Windows.Forms.Label()
        Me.treeViewFolders = New System.Windows.Forms.TreeView()
        Me.grpSourceLocation = New System.Windows.Forms.GroupBox()
        Me.radCentera = New System.Windows.Forms.RadioButton()
        Me.radAddFolder = New System.Windows.Forms.RadioButton()
        Me.radAddFile = New System.Windows.Forms.RadioButton()
        Me.chkComputeBatchSize = New System.Windows.Forms.CheckBox()
        Me.lblBatchSize = New System.Windows.Forms.Label()
        Me.lblTotalBatchSize = New System.Windows.Forms.Label()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.grdArchiveExtractionBatch = New System.Windows.Forms.DataGridView()
        Me.StatusImage = New System.Windows.Forms.DataGridViewImageColumn()
        Me.SelectBatch = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.BatchName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PercentCompleted = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BytesProcessed = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TotalBytes = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ExtractionStatus = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.ItemsProcessed = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ArchiveName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ArchiveSettings = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ArchiveType = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.LightspeedOutputType = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SQLConnectionInfo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SourceInformation = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.WSSSettings = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ItemsExported = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ItemsFailed = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ItemsSkipped = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProcessingFilesDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CaseDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OutputDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LogDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProcessingStartDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProcessingEndDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SummaryReportLocation = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnAddBatchToGrid = New System.Windows.Forms.Button()
        Me.btnLaunchBatches = New System.Windows.Forms.Button()
        Me.btnExpandCollapse = New System.Windows.Forms.Button()
        Me.grpArchiveType = New System.Windows.Forms.GroupBox()
        Me.radJournalArchive = New System.Windows.Forms.RadioButton()
        Me.radMailboxArchive = New System.Windows.Forms.RadioButton()
        Me.grpWSSControl = New System.Windows.Forms.GroupBox()
        Me.chkContact = New System.Windows.Forms.CheckBox()
        Me.chkRSSFeed = New System.Windows.Forms.CheckBox()
        Me.chkCalendar = New System.Windows.Forms.CheckBox()
        Me.chkEmail = New System.Windows.Forms.CheckBox()
        Me.lblContentFiltering = New System.Windows.Forms.Label()
        Me.btnMappingCSV = New System.Windows.Forms.Button()
        Me.txtMappingCSV = New System.Windows.Forms.TextBox()
        Me.lblMappingCSV = New System.Windows.Forms.Label()
        Me.grpVerbose = New System.Windows.Forms.GroupBox()
        Me.radVerboseFalse = New System.Windows.Forms.RadioButton()
        Me.radVerboseTrue = New System.Windows.Forms.RadioButton()
        Me.grpExcludeItems = New System.Windows.Forms.GroupBox()
        Me.radExcludeItemsFalse = New System.Windows.Forms.RadioButton()
        Me.radExcludeItemsTrue = New System.Windows.Forms.RadioButton()
        Me.btnCommCSVSelector = New System.Windows.Forms.Button()
        Me.lblSearchTermsCSV = New System.Windows.Forms.Label()
        Me.txtSearchTermCSV = New System.Windows.Forms.TextBox()
        Me.lblVerbose = New System.Windows.Forms.Label()
        Me.lblExcludeItems = New System.Windows.Forms.Label()
        Me.dateFromDate = New System.Windows.Forms.DateTimePicker()
        Me.lblFromDate = New System.Windows.Forms.Label()
        Me.lblToDate = New System.Windows.Forms.Label()
        Me.dateToDate = New System.Windows.Forms.DateTimePicker()
        Me.btnLoadPreviousBatches = New System.Windows.Forms.Button()
        Me.NuixLogoPicture = New System.Windows.Forms.PictureBox()
        Me.ArchiveExtractToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnExportProcessingDetails = New System.Windows.Forms.Button()
        Me.btnConsolidateExporterFiles = New System.Windows.Forms.Button()
        Me.ConsolidateExporterFiles = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ConsolidateExporterMetricsFilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ConolidateExporterErrorsFilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.grpExtractionType.SuspendLayout()
        Me.grpAXSOneSISSettings.SuspendLayout()
        Me.grpSQLArchiveSettings.SuspendLayout()
        Me.grpEMCSettings.SuspendLayout()
        Me.grpEVSettings.SuspendLayout()
        Me.grpAXSOneSettings.SuspendLayout()
        Me.grpEASSettings.SuspendLayout()
        Me.grpOutputFormat.SuspendLayout()
        Me.grpSourceInformation.SuspendLayout()
        Me.grpSourceLocation.SuspendLayout()
        CType(Me.grdArchiveExtractionBatch, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpArchiveType.SuspendLayout()
        Me.grpWSSControl.SuspendLayout()
        Me.grpVerbose.SuspendLayout()
        Me.grpExcludeItems.SuspendLayout()
        CType(Me.NuixLogoPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ConsolidateExporterFiles.SuspendLayout()
        Me.SuspendLayout()
        '
        'cboArchiveType
        '
        Me.cboArchiveType.FormattingEnabled = True
        Me.cboArchiveType.Items.AddRange(New Object() {"", "Veritas Enterprise Vault", "EMC EmailXtender", "EMC SourceOne", "HP/Autonomy EAS", "Daegis AXS-One"})
        Me.cboArchiveType.Location = New System.Drawing.Point(95, 25)
        Me.cboArchiveType.Margin = New System.Windows.Forms.Padding(2)
        Me.cboArchiveType.Name = "cboArchiveType"
        Me.cboArchiveType.Size = New System.Drawing.Size(161, 21)
        Me.cboArchiveType.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 27)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Legacy Archive:"
        '
        'grpExtractionType
        '
        Me.grpExtractionType.Controls.Add(Me.radLoose)
        Me.grpExtractionType.Controls.Add(Me.radZipped)
        Me.grpExtractionType.Controls.Add(Me.cboExtractionOutputType)
        Me.grpExtractionType.Location = New System.Drawing.Point(555, 11)
        Me.grpExtractionType.Margin = New System.Windows.Forms.Padding(2)
        Me.grpExtractionType.Name = "grpExtractionType"
        Me.grpExtractionType.Padding = New System.Windows.Forms.Padding(2)
        Me.grpExtractionType.Size = New System.Drawing.Size(291, 42)
        Me.grpExtractionType.TabIndex = 2
        Me.grpExtractionType.TabStop = False
        Me.grpExtractionType.Text = "Lightspeed Extraction Output"
        '
        'radLoose
        '
        Me.radLoose.AutoSize = True
        Me.radLoose.Location = New System.Drawing.Point(167, 15)
        Me.radLoose.Name = "radLoose"
        Me.radLoose.Size = New System.Drawing.Size(54, 17)
        Me.radLoose.TabIndex = 75
        Me.radLoose.TabStop = True
        Me.radLoose.Text = "Loose"
        Me.radLoose.UseVisualStyleBackColor = True
        '
        'radZipped
        '
        Me.radZipped.AutoSize = True
        Me.radZipped.Location = New System.Drawing.Point(108, 15)
        Me.radZipped.Name = "radZipped"
        Me.radZipped.Size = New System.Drawing.Size(58, 17)
        Me.radZipped.TabIndex = 74
        Me.radZipped.TabStop = True
        Me.radZipped.Text = "Zipped"
        Me.radZipped.UseVisualStyleBackColor = True
        '
        'cboExtractionOutputType
        '
        Me.cboExtractionOutputType.FormattingEnabled = True
        Me.cboExtractionOutputType.Items.AddRange(New Object() {"pst", "msg", "eml"})
        Me.cboExtractionOutputType.Location = New System.Drawing.Point(12, 16)
        Me.cboExtractionOutputType.Margin = New System.Windows.Forms.Padding(2)
        Me.cboExtractionOutputType.Name = "cboExtractionOutputType"
        Me.cboExtractionOutputType.Size = New System.Drawing.Size(82, 21)
        Me.cboExtractionOutputType.TabIndex = 73
        '
        'grpAXSOneSISSettings
        '
        Me.grpAXSOneSISSettings.Controls.Add(Me.treeAXSOneSis)
        Me.grpAXSOneSISSettings.Location = New System.Drawing.Point(294, 91)
        Me.grpAXSOneSISSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpAXSOneSISSettings.Name = "grpAXSOneSISSettings"
        Me.grpAXSOneSISSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpAXSOneSISSettings.Size = New System.Drawing.Size(291, 333)
        Me.grpAXSOneSISSettings.TabIndex = 4
        Me.grpAXSOneSISSettings.TabStop = False
        Me.grpAXSOneSISSettings.Text = "AXS-One SIS Settings"
        '
        'treeAXSOneSis
        '
        Me.treeAXSOneSis.CheckBoxes = True
        Me.treeAXSOneSis.ImageIndex = 0
        Me.treeAXSOneSis.ImageList = Me.ImgIcons
        Me.treeAXSOneSis.Location = New System.Drawing.Point(4, 23)
        Me.treeAXSOneSis.Margin = New System.Windows.Forms.Padding(2)
        Me.treeAXSOneSis.Name = "treeAXSOneSis"
        Me.treeAXSOneSis.SelectedImageIndex = 0
        Me.treeAXSOneSis.Size = New System.Drawing.Size(283, 303)
        Me.treeAXSOneSis.TabIndex = 24
        '
        'ImgIcons
        '
        Me.ImgIcons.ImageStream = CType(resources.GetObject("ImgIcons.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImgIcons.TransparentColor = System.Drawing.Color.Transparent
        Me.ImgIcons.Images.SetKeyName(0, "folder_Closed_32xMD.png")
        Me.ImgIcons.Images.SetKeyName(1, "folder_Open_32xMD.png")
        Me.ImgIcons.Images.SetKeyName(2, "document_32xMD.png")
        '
        'grpSQLArchiveSettings
        '
        Me.grpSQLArchiveSettings.Controls.Add(Me.lblAuthentication)
        Me.grpSQLArchiveSettings.Controls.Add(Me.cboSecurityType)
        Me.grpSQLArchiveSettings.Controls.Add(Me.lblDomainCheck)
        Me.grpSQLArchiveSettings.Controls.Add(Me.lblSQLPWCheck)
        Me.grpSQLArchiveSettings.Controls.Add(Me.lblSQLUserNameCheck)
        Me.grpSQLArchiveSettings.Controls.Add(Me.lblDBNameCheck)
        Me.grpSQLArchiveSettings.Controls.Add(Me.lblHostCheck)
        Me.grpSQLArchiveSettings.Controls.Add(Me.btnTestSQLConnection)
        Me.grpSQLArchiveSettings.Controls.Add(Me.txtDomain)
        Me.grpSQLArchiveSettings.Controls.Add(Me.lblDomain)
        Me.grpSQLArchiveSettings.Controls.Add(Me.txtSQLInfo)
        Me.grpSQLArchiveSettings.Controls.Add(Me.lblSQLInfo)
        Me.grpSQLArchiveSettings.Controls.Add(Me.txtSQLUserName)
        Me.grpSQLArchiveSettings.Controls.Add(Me.lblSQLUserName)
        Me.grpSQLArchiveSettings.Controls.Add(Me.txtSQLDBName)
        Me.grpSQLArchiveSettings.Controls.Add(Me.lblSQLDBName)
        Me.grpSQLArchiveSettings.Controls.Add(Me.txtSQLPortNumber)
        Me.grpSQLArchiveSettings.Controls.Add(Me.lblPortNumber)
        Me.grpSQLArchiveSettings.Controls.Add(Me.txtSQLHostName)
        Me.grpSQLArchiveSettings.Controls.Add(Me.lblServerName)
        Me.grpSQLArchiveSettings.Location = New System.Drawing.Point(294, 91)
        Me.grpSQLArchiveSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpSQLArchiveSettings.Name = "grpSQLArchiveSettings"
        Me.grpSQLArchiveSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpSQLArchiveSettings.Size = New System.Drawing.Size(287, 333)
        Me.grpSQLArchiveSettings.TabIndex = 5
        Me.grpSQLArchiveSettings.TabStop = False
        Me.grpSQLArchiveSettings.Text = "SQL Connection Info"
        '
        'lblAuthentication
        '
        Me.lblAuthentication.AutoSize = True
        Me.lblAuthentication.Location = New System.Drawing.Point(4, 27)
        Me.lblAuthentication.Name = "lblAuthentication"
        Me.lblAuthentication.Size = New System.Drawing.Size(78, 13)
        Me.lblAuthentication.TabIndex = 19
        Me.lblAuthentication.Text = "Authentication:"
        '
        'cboSecurityType
        '
        Me.cboSecurityType.FormattingEnabled = True
        Me.cboSecurityType.Items.AddRange(New Object() {"", "Windows Authentication", "SQLServer Authentication"})
        Me.cboSecurityType.Location = New System.Drawing.Point(103, 22)
        Me.cboSecurityType.Name = "cboSecurityType"
        Me.cboSecurityType.Size = New System.Drawing.Size(178, 21)
        Me.cboSecurityType.TabIndex = 18
        '
        'lblDomainCheck
        '
        Me.lblDomainCheck.AutoSize = True
        Me.lblDomainCheck.ForeColor = System.Drawing.Color.Red
        Me.lblDomainCheck.Location = New System.Drawing.Point(84, 100)
        Me.lblDomainCheck.Name = "lblDomainCheck"
        Me.lblDomainCheck.Size = New System.Drawing.Size(14, 13)
        Me.lblDomainCheck.TabIndex = 17
        Me.lblDomainCheck.Text = "X"
        '
        'lblSQLPWCheck
        '
        Me.lblSQLPWCheck.AutoSize = True
        Me.lblSQLPWCheck.ForeColor = System.Drawing.Color.Red
        Me.lblSQLPWCheck.Location = New System.Drawing.Point(84, 151)
        Me.lblSQLPWCheck.Name = "lblSQLPWCheck"
        Me.lblSQLPWCheck.Size = New System.Drawing.Size(14, 13)
        Me.lblSQLPWCheck.TabIndex = 16
        Me.lblSQLPWCheck.Text = "X"
        '
        'lblSQLUserNameCheck
        '
        Me.lblSQLUserNameCheck.AutoSize = True
        Me.lblSQLUserNameCheck.ForeColor = System.Drawing.Color.Red
        Me.lblSQLUserNameCheck.Location = New System.Drawing.Point(84, 126)
        Me.lblSQLUserNameCheck.Name = "lblSQLUserNameCheck"
        Me.lblSQLUserNameCheck.Size = New System.Drawing.Size(14, 13)
        Me.lblSQLUserNameCheck.TabIndex = 15
        Me.lblSQLUserNameCheck.Text = "X"
        '
        'lblDBNameCheck
        '
        Me.lblDBNameCheck.AutoSize = True
        Me.lblDBNameCheck.ForeColor = System.Drawing.Color.Red
        Me.lblDBNameCheck.Location = New System.Drawing.Point(84, 76)
        Me.lblDBNameCheck.Name = "lblDBNameCheck"
        Me.lblDBNameCheck.Size = New System.Drawing.Size(14, 13)
        Me.lblDBNameCheck.TabIndex = 14
        Me.lblDBNameCheck.Text = "X"
        '
        'lblHostCheck
        '
        Me.lblHostCheck.AutoSize = True
        Me.lblHostCheck.ForeColor = System.Drawing.Color.Red
        Me.lblHostCheck.Location = New System.Drawing.Point(84, 51)
        Me.lblHostCheck.Name = "lblHostCheck"
        Me.lblHostCheck.Size = New System.Drawing.Size(14, 13)
        Me.lblHostCheck.TabIndex = 13
        Me.lblHostCheck.Text = "X"
        '
        'btnTestSQLConnection
        '
        Me.btnTestSQLConnection.Location = New System.Drawing.Point(103, 173)
        Me.btnTestSQLConnection.Name = "btnTestSQLConnection"
        Me.btnTestSQLConnection.Size = New System.Drawing.Size(120, 23)
        Me.btnTestSQLConnection.TabIndex = 12
        Me.btnTestSQLConnection.Text = "Test SQL Connection"
        Me.btnTestSQLConnection.UseVisualStyleBackColor = True
        '
        'txtDomain
        '
        Me.txtDomain.Location = New System.Drawing.Point(103, 97)
        Me.txtDomain.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDomain.Name = "txtDomain"
        Me.txtDomain.Size = New System.Drawing.Size(178, 20)
        Me.txtDomain.TabIndex = 11
        '
        'lblDomain
        '
        Me.lblDomain.AutoSize = True
        Me.lblDomain.Location = New System.Drawing.Point(4, 98)
        Me.lblDomain.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblDomain.Name = "lblDomain"
        Me.lblDomain.Size = New System.Drawing.Size(46, 13)
        Me.lblDomain.TabIndex = 10
        Me.lblDomain.Text = "Domain:"
        '
        'txtSQLInfo
        '
        Me.txtSQLInfo.Location = New System.Drawing.Point(103, 148)
        Me.txtSQLInfo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSQLInfo.Name = "txtSQLInfo"
        Me.txtSQLInfo.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSQLInfo.Size = New System.Drawing.Size(178, 20)
        Me.txtSQLInfo.TabIndex = 9
        '
        'lblSQLInfo
        '
        Me.lblSQLInfo.AutoSize = True
        Me.lblSQLInfo.Location = New System.Drawing.Point(4, 150)
        Me.lblSQLInfo.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblSQLInfo.Name = "lblSQLInfo"
        Me.lblSQLInfo.Size = New System.Drawing.Size(56, 13)
        Me.lblSQLInfo.TabIndex = 8
        Me.lblSQLInfo.Text = "Password:"
        '
        'txtSQLUserName
        '
        Me.txtSQLUserName.Location = New System.Drawing.Point(103, 123)
        Me.txtSQLUserName.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSQLUserName.Name = "txtSQLUserName"
        Me.txtSQLUserName.Size = New System.Drawing.Size(178, 20)
        Me.txtSQLUserName.TabIndex = 7
        '
        'lblSQLUserName
        '
        Me.lblSQLUserName.AutoSize = True
        Me.lblSQLUserName.Location = New System.Drawing.Point(4, 124)
        Me.lblSQLUserName.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblSQLUserName.Name = "lblSQLUserName"
        Me.lblSQLUserName.Size = New System.Drawing.Size(58, 13)
        Me.lblSQLUserName.TabIndex = 6
        Me.lblSQLUserName.Text = "Username:"
        '
        'txtSQLDBName
        '
        Me.txtSQLDBName.Location = New System.Drawing.Point(103, 72)
        Me.txtSQLDBName.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSQLDBName.Name = "txtSQLDBName"
        Me.txtSQLDBName.Size = New System.Drawing.Size(178, 20)
        Me.txtSQLDBName.TabIndex = 5
        '
        'lblSQLDBName
        '
        Me.lblSQLDBName.AutoSize = True
        Me.lblSQLDBName.Location = New System.Drawing.Point(4, 73)
        Me.lblSQLDBName.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblSQLDBName.Name = "lblSQLDBName"
        Me.lblSQLDBName.Size = New System.Drawing.Size(56, 13)
        Me.lblSQLDBName.TabIndex = 4
        Me.lblSQLDBName.Text = "DB Name:"
        '
        'txtSQLPortNumber
        '
        Me.txtSQLPortNumber.Location = New System.Drawing.Point(225, 48)
        Me.txtSQLPortNumber.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSQLPortNumber.Name = "txtSQLPortNumber"
        Me.txtSQLPortNumber.Size = New System.Drawing.Size(56, 20)
        Me.txtSQLPortNumber.TabIndex = 3
        '
        'lblPortNumber
        '
        Me.lblPortNumber.AutoSize = True
        Me.lblPortNumber.Location = New System.Drawing.Point(192, 49)
        Me.lblPortNumber.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblPortNumber.Name = "lblPortNumber"
        Me.lblPortNumber.Size = New System.Drawing.Size(29, 13)
        Me.lblPortNumber.TabIndex = 2
        Me.lblPortNumber.Text = "Port:"
        '
        'txtSQLHostName
        '
        Me.txtSQLHostName.Location = New System.Drawing.Point(103, 48)
        Me.txtSQLHostName.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSQLHostName.Name = "txtSQLHostName"
        Me.txtSQLHostName.Size = New System.Drawing.Size(85, 20)
        Me.txtSQLHostName.TabIndex = 1
        '
        'lblServerName
        '
        Me.lblServerName.AutoSize = True
        Me.lblServerName.Location = New System.Drawing.Point(4, 49)
        Me.lblServerName.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblServerName.Name = "lblServerName"
        Me.lblServerName.Size = New System.Drawing.Size(41, 13)
        Me.lblServerName.TabIndex = 0
        Me.lblServerName.Text = "Server:"
        '
        'grpEMCSettings
        '
        Me.grpEMCSettings.Controls.Add(Me.chkAFEX)
        Me.grpEMCSettings.Controls.Add(Me.chkAFPST)
        Me.grpEMCSettings.Controls.Add(Me.chkAFSYS)
        Me.grpEMCSettings.Controls.Add(Me.lblExpandDLto)
        Me.grpEMCSettings.Controls.Add(Me.cboExpandDLLocation)
        Me.grpEMCSettings.Controls.Add(Me.lblAddressFiltering)
        Me.grpEMCSettings.Location = New System.Drawing.Point(589, 91)
        Me.grpEMCSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpEMCSettings.Name = "grpEMCSettings"
        Me.grpEMCSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpEMCSettings.Size = New System.Drawing.Size(295, 333)
        Me.grpEMCSettings.TabIndex = 6
        Me.grpEMCSettings.TabStop = False
        Me.grpEMCSettings.Text = "EMC Settings"
        '
        'chkAFEX
        '
        Me.chkAFEX.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkAFEX.AutoSize = True
        Me.chkAFEX.BackColor = System.Drawing.Color.Green
        Me.chkAFEX.Checked = True
        Me.chkAFEX.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAFEX.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.chkAFEX.Location = New System.Drawing.Point(133, 23)
        Me.chkAFEX.Name = "chkAFEX"
        Me.chkAFEX.Size = New System.Drawing.Size(31, 23)
        Me.chkAFEX.TabIndex = 12
        Me.chkAFEX.Text = "EX"
        Me.chkAFEX.UseVisualStyleBackColor = False
        '
        'chkAFPST
        '
        Me.chkAFPST.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkAFPST.AutoSize = True
        Me.chkAFPST.BackColor = System.Drawing.Color.Green
        Me.chkAFPST.Checked = True
        Me.chkAFPST.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAFPST.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.chkAFPST.Location = New System.Drawing.Point(162, 23)
        Me.chkAFPST.Name = "chkAFPST"
        Me.chkAFPST.Size = New System.Drawing.Size(38, 23)
        Me.chkAFPST.TabIndex = 11
        Me.chkAFPST.Text = "PST"
        Me.chkAFPST.UseVisualStyleBackColor = False
        '
        'chkAFSYS
        '
        Me.chkAFSYS.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkAFSYS.AutoSize = True
        Me.chkAFSYS.BackColor = System.Drawing.Color.Green
        Me.chkAFSYS.Checked = True
        Me.chkAFSYS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAFSYS.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.chkAFSYS.Location = New System.Drawing.Point(97, 23)
        Me.chkAFSYS.Name = "chkAFSYS"
        Me.chkAFSYS.Size = New System.Drawing.Size(38, 23)
        Me.chkAFSYS.TabIndex = 10
        Me.chkAFSYS.Text = "SYS"
        Me.chkAFSYS.UseVisualStyleBackColor = False
        '
        'lblExpandDLto
        '
        Me.lblExpandDLto.AutoSize = True
        Me.lblExpandDLto.Location = New System.Drawing.Point(4, 57)
        Me.lblExpandDLto.Name = "lblExpandDLto"
        Me.lblExpandDLto.Size = New System.Drawing.Size(75, 13)
        Me.lblExpandDLto.TabIndex = 9
        Me.lblExpandDLto.Text = "Expand DL to:"
        '
        'cboExpandDLLocation
        '
        Me.cboExpandDLLocation.FormattingEnabled = True
        Me.cboExpandDLLocation.Items.AddRange(New Object() {"To: + ""Expanded-DL""", "Bcc: + ""Expanded-DL""", """Expanded-DL"""})
        Me.cboExpandDLLocation.Location = New System.Drawing.Point(96, 52)
        Me.cboExpandDLLocation.Name = "cboExpandDLLocation"
        Me.cboExpandDLLocation.Size = New System.Drawing.Size(135, 21)
        Me.cboExpandDLLocation.TabIndex = 8
        '
        'lblAddressFiltering
        '
        Me.lblAddressFiltering.AutoSize = True
        Me.lblAddressFiltering.Location = New System.Drawing.Point(4, 25)
        Me.lblAddressFiltering.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblAddressFiltering.Name = "lblAddressFiltering"
        Me.lblAddressFiltering.Size = New System.Drawing.Size(87, 13)
        Me.lblAddressFiltering.TabIndex = 7
        Me.lblAddressFiltering.Text = "Address Filtering:"
        '
        'grpEVSettings
        '
        Me.grpEVSettings.Controls.Add(Me.btnEVUserListFileChooser)
        Me.grpEVSettings.Controls.Add(Me.txtEVUserList)
        Me.grpEVSettings.Controls.Add(Me.lblEVUserList)
        Me.grpEVSettings.Controls.Add(Me.chkSkipVaultStorePartitionErrors)
        Me.grpEVSettings.Controls.Add(Me.chkSkipAdditionalSQLLookup)
        Me.grpEVSettings.Location = New System.Drawing.Point(589, 71)
        Me.grpEVSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpEVSettings.Name = "grpEVSettings"
        Me.grpEVSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpEVSettings.Size = New System.Drawing.Size(295, 333)
        Me.grpEVSettings.TabIndex = 7
        Me.grpEVSettings.TabStop = False
        Me.grpEVSettings.Text = "Enterprise Vault Settings"
        '
        'btnEVUserListFileChooser
        '
        Me.btnEVUserListFileChooser.Location = New System.Drawing.Point(263, 67)
        Me.btnEVUserListFileChooser.Name = "btnEVUserListFileChooser"
        Me.btnEVUserListFileChooser.Size = New System.Drawing.Size(27, 23)
        Me.btnEVUserListFileChooser.TabIndex = 11
        Me.btnEVUserListFileChooser.Text = "..."
        Me.btnEVUserListFileChooser.UseVisualStyleBackColor = True
        '
        'txtEVUserList
        '
        Me.txtEVUserList.Location = New System.Drawing.Point(54, 68)
        Me.txtEVUserList.Name = "txtEVUserList"
        Me.txtEVUserList.Size = New System.Drawing.Size(203, 20)
        Me.txtEVUserList.TabIndex = 10
        '
        'lblEVUserList
        '
        Me.lblEVUserList.AutoSize = True
        Me.lblEVUserList.Location = New System.Drawing.Point(5, 73)
        Me.lblEVUserList.Name = "lblEVUserList"
        Me.lblEVUserList.Size = New System.Drawing.Size(51, 13)
        Me.lblEVUserList.TabIndex = 9
        Me.lblEVUserList.Text = "User List:"
        '
        'chkSkipVaultStorePartitionErrors
        '
        Me.chkSkipVaultStorePartitionErrors.AutoSize = True
        Me.chkSkipVaultStorePartitionErrors.Location = New System.Drawing.Point(5, 46)
        Me.chkSkipVaultStorePartitionErrors.Margin = New System.Windows.Forms.Padding(2)
        Me.chkSkipVaultStorePartitionErrors.Name = "chkSkipVaultStorePartitionErrors"
        Me.chkSkipVaultStorePartitionErrors.Size = New System.Drawing.Size(256, 17)
        Me.chkSkipVaultStorePartitionErrors.TabIndex = 8
        Me.chkSkipVaultStorePartitionErrors.Text = "Use FileTransactionID over ParentTransactionID"
        Me.chkSkipVaultStorePartitionErrors.UseVisualStyleBackColor = True
        '
        'chkSkipAdditionalSQLLookup
        '
        Me.chkSkipAdditionalSQLLookup.AutoSize = True
        Me.chkSkipAdditionalSQLLookup.Location = New System.Drawing.Point(5, 23)
        Me.chkSkipAdditionalSQLLookup.Margin = New System.Windows.Forms.Padding(2)
        Me.chkSkipAdditionalSQLLookup.Name = "chkSkipAdditionalSQLLookup"
        Me.chkSkipAdditionalSQLLookup.Size = New System.Drawing.Size(164, 17)
        Me.chkSkipAdditionalSQLLookup.TabIndex = 8
        Me.chkSkipAdditionalSQLLookup.Text = "Skip Additional SQL Lookups"
        Me.chkSkipAdditionalSQLLookup.UseVisualStyleBackColor = True
        '
        'btnClearDates
        '
        Me.btnClearDates.Location = New System.Drawing.Point(538, 60)
        Me.btnClearDates.Name = "btnClearDates"
        Me.btnClearDates.Size = New System.Drawing.Size(75, 23)
        Me.btnClearDates.TabIndex = 94
        Me.btnClearDates.Text = "Clear"
        Me.btnClearDates.UseVisualStyleBackColor = True
        '
        'grpAXSOneSettings
        '
        Me.grpAXSOneSettings.Controls.Add(Me.chkAXSOneSkipSISLookups)
        Me.grpAXSOneSettings.Location = New System.Drawing.Point(589, 42)
        Me.grpAXSOneSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpAXSOneSettings.Name = "grpAXSOneSettings"
        Me.grpAXSOneSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpAXSOneSettings.Size = New System.Drawing.Size(295, 333)
        Me.grpAXSOneSettings.TabIndex = 10
        Me.grpAXSOneSettings.TabStop = False
        Me.grpAXSOneSettings.Text = "AXS-One Settings"
        '
        'chkAXSOneSkipSISLookups
        '
        Me.chkAXSOneSkipSISLookups.AutoSize = True
        Me.chkAXSOneSkipSISLookups.Location = New System.Drawing.Point(7, 26)
        Me.chkAXSOneSkipSISLookups.Margin = New System.Windows.Forms.Padding(2)
        Me.chkAXSOneSkipSISLookups.Name = "chkAXSOneSkipSISLookups"
        Me.chkAXSOneSkipSISLookups.Size = New System.Drawing.Size(111, 17)
        Me.chkAXSOneSkipSISLookups.TabIndex = 9
        Me.chkAXSOneSkipSISLookups.Text = "Skip SIS Lookups"
        Me.chkAXSOneSkipSISLookups.UseVisualStyleBackColor = True
        '
        'grpEASSettings
        '
        Me.grpEASSettings.Controls.Add(Me.txtDocServerID)
        Me.grpEASSettings.Controls.Add(Me.lblDocServerID)
        Me.grpEASSettings.Location = New System.Drawing.Point(589, 91)
        Me.grpEASSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpEASSettings.Name = "grpEASSettings"
        Me.grpEASSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpEASSettings.Size = New System.Drawing.Size(295, 300)
        Me.grpEASSettings.TabIndex = 8
        Me.grpEASSettings.TabStop = False
        Me.grpEASSettings.Text = "EAS Settings"
        '
        'txtDocServerID
        '
        Me.txtDocServerID.Location = New System.Drawing.Point(89, 23)
        Me.txtDocServerID.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDocServerID.Name = "txtDocServerID"
        Me.txtDocServerID.Size = New System.Drawing.Size(103, 20)
        Me.txtDocServerID.TabIndex = 1
        '
        'lblDocServerID
        '
        Me.lblDocServerID.AutoSize = True
        Me.lblDocServerID.Location = New System.Drawing.Point(10, 25)
        Me.lblDocServerID.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblDocServerID.Name = "lblDocServerID"
        Me.lblDocServerID.Size = New System.Drawing.Size(78, 13)
        Me.lblDocServerID.TabIndex = 0
        Me.lblDocServerID.Text = "Doc Server ID:"
        '
        'txtCustodianListCSV
        '
        Me.txtCustodianListCSV.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtCustodianListCSV.Location = New System.Drawing.Point(1002, -47)
        Me.txtCustodianListCSV.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCustodianListCSV.Name = "txtCustodianListCSV"
        Me.txtCustodianListCSV.Size = New System.Drawing.Size(237, 20)
        Me.txtCustodianListCSV.TabIndex = 10
        '
        'lblCustodianListCSV
        '
        Me.lblCustodianListCSV.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblCustodianListCSV.AutoSize = True
        Me.lblCustodianListCSV.Location = New System.Drawing.Point(897, -47)
        Me.lblCustodianListCSV.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblCustodianListCSV.Name = "lblCustodianListCSV"
        Me.lblCustodianListCSV.Size = New System.Drawing.Size(100, 13)
        Me.lblCustodianListCSV.TabIndex = 11
        Me.lblCustodianListCSV.Text = "Custodian List CSV:"
        '
        'btnCustodianListCSVChooser
        '
        Me.btnCustodianListCSVChooser.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCustodianListCSVChooser.Location = New System.Drawing.Point(1241, -47)
        Me.btnCustodianListCSVChooser.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCustodianListCSVChooser.Name = "btnCustodianListCSVChooser"
        Me.btnCustodianListCSVChooser.Size = New System.Drawing.Size(21, 21)
        Me.btnCustodianListCSVChooser.TabIndex = 12
        Me.btnCustodianListCSVChooser.Text = "..."
        Me.btnCustodianListCSVChooser.UseVisualStyleBackColor = True
        '
        'grpOutputFormat
        '
        Me.grpOutputFormat.Controls.Add(Me.radFlat)
        Me.grpOutputFormat.Controls.Add(Me.radUser)
        Me.grpOutputFormat.Location = New System.Drawing.Point(426, 11)
        Me.grpOutputFormat.Margin = New System.Windows.Forms.Padding(2)
        Me.grpOutputFormat.Name = "grpOutputFormat"
        Me.grpOutputFormat.Padding = New System.Windows.Forms.Padding(2)
        Me.grpOutputFormat.Size = New System.Drawing.Size(125, 42)
        Me.grpOutputFormat.TabIndex = 78
        Me.grpOutputFormat.TabStop = False
        Me.grpOutputFormat.Text = "Output Type"
        '
        'radFlat
        '
        Me.radFlat.AutoSize = True
        Me.radFlat.Location = New System.Drawing.Point(68, 17)
        Me.radFlat.Margin = New System.Windows.Forms.Padding(2)
        Me.radFlat.Name = "radFlat"
        Me.radFlat.Size = New System.Drawing.Size(42, 17)
        Me.radFlat.TabIndex = 79
        Me.radFlat.TabStop = True
        Me.radFlat.Text = "Flat"
        Me.radFlat.UseVisualStyleBackColor = True
        '
        'radUser
        '
        Me.radUser.AutoSize = True
        Me.radUser.Location = New System.Drawing.Point(9, 17)
        Me.radUser.Margin = New System.Windows.Forms.Padding(2)
        Me.radUser.Name = "radUser"
        Me.radUser.Size = New System.Drawing.Size(47, 17)
        Me.radUser.TabIndex = 78
        Me.radUser.TabStop = True
        Me.radUser.Text = "User"
        Me.radUser.UseVisualStyleBackColor = True
        '
        'btnShowSettings
        '
        Me.btnShowSettings.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnShowSettings.BackColor = System.Drawing.Color.White
        Me.btnShowSettings.Image = CType(resources.GetObject("btnShowSettings.Image"), System.Drawing.Image)
        Me.btnShowSettings.Location = New System.Drawing.Point(9, 533)
        Me.btnShowSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.btnShowSettings.Name = "btnShowSettings"
        Me.btnShowSettings.Size = New System.Drawing.Size(80, 60)
        Me.btnShowSettings.TabIndex = 79
        Me.btnShowSettings.UseVisualStyleBackColor = False
        '
        'grpSourceInformation
        '
        Me.grpSourceInformation.Controls.Add(Me.btnIPFile)
        Me.grpSourceInformation.Controls.Add(Me.txtIPFile)
        Me.grpSourceInformation.Controls.Add(Me.lblIPFile)
        Me.grpSourceInformation.Controls.Add(Me.btnPEAFile)
        Me.grpSourceInformation.Controls.Add(Me.txtPEAFile)
        Me.grpSourceInformation.Controls.Add(Me.lblPEAFile)
        Me.grpSourceInformation.Controls.Add(Me.treeViewFolders)
        Me.grpSourceInformation.Controls.Add(Me.grpSourceLocation)
        Me.grpSourceInformation.Controls.Add(Me.chkComputeBatchSize)
        Me.grpSourceInformation.Controls.Add(Me.lblBatchSize)
        Me.grpSourceInformation.Controls.Add(Me.lblTotalBatchSize)
        Me.grpSourceInformation.Location = New System.Drawing.Point(11, 91)
        Me.grpSourceInformation.Margin = New System.Windows.Forms.Padding(2)
        Me.grpSourceInformation.Name = "grpSourceInformation"
        Me.grpSourceInformation.Padding = New System.Windows.Forms.Padding(2)
        Me.grpSourceInformation.Size = New System.Drawing.Size(279, 333)
        Me.grpSourceInformation.TabIndex = 82
        Me.grpSourceInformation.TabStop = False
        Me.grpSourceInformation.Text = "Source Information"
        '
        'btnIPFile
        '
        Me.btnIPFile.Location = New System.Drawing.Point(251, 294)
        Me.btnIPFile.Name = "btnIPFile"
        Me.btnIPFile.Size = New System.Drawing.Size(24, 23)
        Me.btnIPFile.TabIndex = 96
        Me.btnIPFile.Text = "..."
        Me.btnIPFile.UseVisualStyleBackColor = True
        '
        'txtIPFile
        '
        Me.txtIPFile.Location = New System.Drawing.Point(67, 294)
        Me.txtIPFile.Name = "txtIPFile"
        Me.txtIPFile.Size = New System.Drawing.Size(176, 20)
        Me.txtIPFile.TabIndex = 95
        '
        'lblIPFile
        '
        Me.lblIPFile.AutoSize = True
        Me.lblIPFile.Location = New System.Drawing.Point(4, 300)
        Me.lblIPFile.Name = "lblIPFile"
        Me.lblIPFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblIPFile.Size = New System.Drawing.Size(39, 13)
        Me.lblIPFile.TabIndex = 94
        Me.lblIPFile.Text = "IP File:"
        '
        'btnPEAFile
        '
        Me.btnPEAFile.Location = New System.Drawing.Point(251, 268)
        Me.btnPEAFile.Name = "btnPEAFile"
        Me.btnPEAFile.Size = New System.Drawing.Size(24, 23)
        Me.btnPEAFile.TabIndex = 93
        Me.btnPEAFile.Text = "..."
        Me.btnPEAFile.UseVisualStyleBackColor = True
        '
        'txtPEAFile
        '
        Me.txtPEAFile.Location = New System.Drawing.Point(67, 268)
        Me.txtPEAFile.Name = "txtPEAFile"
        Me.txtPEAFile.Size = New System.Drawing.Size(176, 20)
        Me.txtPEAFile.TabIndex = 92
        '
        'lblPEAFile
        '
        Me.lblPEAFile.AutoSize = True
        Me.lblPEAFile.Location = New System.Drawing.Point(4, 271)
        Me.lblPEAFile.Name = "lblPEAFile"
        Me.lblPEAFile.Size = New System.Drawing.Size(50, 13)
        Me.lblPEAFile.TabIndex = 91
        Me.lblPEAFile.Text = "PEA File:"
        '
        'treeViewFolders
        '
        Me.treeViewFolders.CheckBoxes = True
        Me.treeViewFolders.ImageIndex = 0
        Me.treeViewFolders.ImageList = Me.ImgIcons
        Me.treeViewFolders.Location = New System.Drawing.Point(4, 93)
        Me.treeViewFolders.Margin = New System.Windows.Forms.Padding(2)
        Me.treeViewFolders.Name = "treeViewFolders"
        Me.treeViewFolders.SelectedImageIndex = 0
        Me.treeViewFolders.Size = New System.Drawing.Size(266, 170)
        Me.treeViewFolders.TabIndex = 23
        '
        'grpSourceLocation
        '
        Me.grpSourceLocation.Controls.Add(Me.radCentera)
        Me.grpSourceLocation.Controls.Add(Me.radAddFolder)
        Me.grpSourceLocation.Controls.Add(Me.radAddFile)
        Me.grpSourceLocation.Location = New System.Drawing.Point(8, 18)
        Me.grpSourceLocation.Margin = New System.Windows.Forms.Padding(2)
        Me.grpSourceLocation.Name = "grpSourceLocation"
        Me.grpSourceLocation.Padding = New System.Windows.Forms.Padding(2)
        Me.grpSourceLocation.Size = New System.Drawing.Size(262, 42)
        Me.grpSourceLocation.TabIndex = 0
        Me.grpSourceLocation.TabStop = False
        Me.grpSourceLocation.Text = "Source Location"
        '
        'radCentera
        '
        Me.radCentera.AutoSize = True
        Me.radCentera.Location = New System.Drawing.Point(127, 19)
        Me.radCentera.Name = "radCentera"
        Me.radCentera.Size = New System.Drawing.Size(62, 17)
        Me.radCentera.TabIndex = 2
        Me.radCentera.TabStop = True
        Me.radCentera.Text = "Centera"
        Me.radCentera.UseVisualStyleBackColor = True
        '
        'radAddFolder
        '
        Me.radAddFolder.AutoSize = True
        Me.radAddFolder.Location = New System.Drawing.Point(7, 19)
        Me.radAddFolder.Margin = New System.Windows.Forms.Padding(2)
        Me.radAddFolder.Name = "radAddFolder"
        Me.radAddFolder.Size = New System.Drawing.Size(59, 17)
        Me.radAddFolder.TabIndex = 1
        Me.radAddFolder.TabStop = True
        Me.radAddFolder.Text = "Folders"
        Me.radAddFolder.UseVisualStyleBackColor = True
        '
        'radAddFile
        '
        Me.radAddFile.AutoSize = True
        Me.radAddFile.Location = New System.Drawing.Point(76, 19)
        Me.radAddFile.Margin = New System.Windows.Forms.Padding(2)
        Me.radAddFile.Name = "radAddFile"
        Me.radAddFile.Size = New System.Drawing.Size(46, 17)
        Me.radAddFile.TabIndex = 0
        Me.radAddFile.TabStop = True
        Me.radAddFile.Text = "Files"
        Me.radAddFile.UseVisualStyleBackColor = True
        '
        'chkComputeBatchSize
        '
        Me.chkComputeBatchSize.AutoSize = True
        Me.chkComputeBatchSize.Location = New System.Drawing.Point(15, 70)
        Me.chkComputeBatchSize.Margin = New System.Windows.Forms.Padding(2)
        Me.chkComputeBatchSize.Name = "chkComputeBatchSize"
        Me.chkComputeBatchSize.Size = New System.Drawing.Size(122, 17)
        Me.chkComputeBatchSize.TabIndex = 88
        Me.chkComputeBatchSize.Text = "Compute Batch Size"
        Me.chkComputeBatchSize.UseVisualStyleBackColor = True
        '
        'lblBatchSize
        '
        Me.lblBatchSize.AutoSize = True
        Me.lblBatchSize.Location = New System.Drawing.Point(138, 70)
        Me.lblBatchSize.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblBatchSize.Name = "lblBatchSize"
        Me.lblBatchSize.Size = New System.Drawing.Size(30, 13)
        Me.lblBatchSize.TabIndex = 89
        Me.lblBatchSize.Text = "Size:"
        '
        'lblTotalBatchSize
        '
        Me.lblTotalBatchSize.AutoSize = True
        Me.lblTotalBatchSize.Location = New System.Drawing.Point(176, 70)
        Me.lblTotalBatchSize.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTotalBatchSize.Name = "lblTotalBatchSize"
        Me.lblTotalBatchSize.Size = New System.Drawing.Size(13, 13)
        Me.lblTotalBatchSize.TabIndex = 90
        Me.lblTotalBatchSize.Text = "0"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(61, 4)
        '
        'grdArchiveExtractionBatch
        '
        Me.grdArchiveExtractionBatch.AllowUserToAddRows = False
        Me.grdArchiveExtractionBatch.AllowUserToDeleteRows = False
        Me.grdArchiveExtractionBatch.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grdArchiveExtractionBatch.BackgroundColor = System.Drawing.SystemColors.Control
        Me.grdArchiveExtractionBatch.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdArchiveExtractionBatch.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.StatusImage, Me.SelectBatch, Me.BatchName, Me.PercentCompleted, Me.BytesProcessed, Me.TotalBytes, Me.ExtractionStatus, Me.ItemsProcessed, Me.ArchiveName, Me.ArchiveSettings, Me.ArchiveType, Me.LightspeedOutputType, Me.SQLConnectionInfo, Me.SourceInformation, Me.WSSSettings, Me.ItemsExported, Me.ItemsFailed, Me.ItemsSkipped, Me.ProcessingFilesDirectory, Me.CaseDirectory, Me.OutputDirectory, Me.LogDirectory, Me.ProcessingStartDate, Me.ProcessingEndDate, Me.SummaryReportLocation})
        Me.grdArchiveExtractionBatch.Location = New System.Drawing.Point(10, 437)
        Me.grdArchiveExtractionBatch.Margin = New System.Windows.Forms.Padding(2)
        Me.grdArchiveExtractionBatch.Name = "grdArchiveExtractionBatch"
        Me.grdArchiveExtractionBatch.RowTemplate.Height = 28
        Me.grdArchiveExtractionBatch.Size = New System.Drawing.Size(1468, 91)
        Me.grdArchiveExtractionBatch.TabIndex = 84
        '
        'StatusImage
        '
        Me.StatusImage.Frozen = True
        Me.StatusImage.HeaderText = ""
        Me.StatusImage.Name = "StatusImage"
        Me.StatusImage.Width = 30
        '
        'SelectBatch
        '
        Me.SelectBatch.Frozen = True
        Me.SelectBatch.HeaderText = "Select Batch"
        Me.SelectBatch.Name = "SelectBatch"
        Me.SelectBatch.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.SelectBatch.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.SelectBatch.Width = 50
        '
        'BatchName
        '
        Me.BatchName.Frozen = True
        Me.BatchName.HeaderText = "Batch Name"
        Me.BatchName.Name = "BatchName"
        Me.BatchName.Width = 125
        '
        'PercentCompleted
        '
        Me.PercentCompleted.Frozen = True
        Me.PercentCompleted.HeaderText = "Percent Completed"
        Me.PercentCompleted.Name = "PercentCompleted"
        Me.PercentCompleted.Width = 60
        '
        'BytesProcessed
        '
        Me.BytesProcessed.Frozen = True
        Me.BytesProcessed.HeaderText = "Bytes Processed"
        Me.BytesProcessed.Name = "BytesProcessed"
        '
        'TotalBytes
        '
        Me.TotalBytes.Frozen = True
        Me.TotalBytes.HeaderText = "Total Bytes"
        Me.TotalBytes.Name = "TotalBytes"
        '
        'ExtractionStatus
        '
        Me.ExtractionStatus.Frozen = True
        Me.ExtractionStatus.HeaderText = "Extraction Status"
        Me.ExtractionStatus.Items.AddRange(New Object() {"Not Started", "In Progress", "Completed", "Failed", "Cancelled by User", "SQL Connection Lost", "Restart Extraction"})
        Me.ExtractionStatus.Name = "ExtractionStatus"
        Me.ExtractionStatus.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ExtractionStatus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ExtractionStatus.Width = 150
        '
        'ItemsProcessed
        '
        Me.ItemsProcessed.HeaderText = "Items Processed"
        Me.ItemsProcessed.Name = "ItemsProcessed"
        '
        'ArchiveName
        '
        Me.ArchiveName.HeaderText = "Archive Name"
        Me.ArchiveName.Name = "ArchiveName"
        Me.ArchiveName.Width = 125
        '
        'ArchiveSettings
        '
        Me.ArchiveSettings.HeaderText = "Archive Settings"
        Me.ArchiveSettings.Name = "ArchiveSettings"
        Me.ArchiveSettings.Width = 150
        '
        'ArchiveType
        '
        Me.ArchiveType.HeaderText = "Archive Type"
        Me.ArchiveType.Items.AddRange(New Object() {"Mailbox:User", "Mailbox:Flat", "Journal:User", "Journal:Flat"})
        Me.ArchiveType.Name = "ArchiveType"
        Me.ArchiveType.Width = 125
        '
        'LightspeedOutputType
        '
        Me.LightspeedOutputType.HeaderText = "Extraction Output"
        Me.LightspeedOutputType.Name = "LightspeedOutputType"
        Me.LightspeedOutputType.Width = 60
        '
        'SQLConnectionInfo
        '
        Me.SQLConnectionInfo.HeaderText = "SQL Connection Info"
        Me.SQLConnectionInfo.Name = "SQLConnectionInfo"
        Me.SQLConnectionInfo.Width = 150
        '
        'SourceInformation
        '
        Me.SourceInformation.HeaderText = "Source Information"
        Me.SourceInformation.Name = "SourceInformation"
        Me.SourceInformation.Width = 150
        '
        'WSSSettings
        '
        Me.WSSSettings.HeaderText = "WSS Settings"
        Me.WSSSettings.Name = "WSSSettings"
        Me.WSSSettings.Width = 150
        '
        'ItemsExported
        '
        Me.ItemsExported.HeaderText = "Items Exported"
        Me.ItemsExported.Name = "ItemsExported"
        '
        'ItemsFailed
        '
        Me.ItemsFailed.HeaderText = "Items Failed"
        Me.ItemsFailed.Name = "ItemsFailed"
        '
        'ItemsSkipped
        '
        Me.ItemsSkipped.HeaderText = "Items Skipped"
        Me.ItemsSkipped.Name = "ItemsSkipped"
        '
        'ProcessingFilesDirectory
        '
        Me.ProcessingFilesDirectory.HeaderText = "Processing Files Directory"
        Me.ProcessingFilesDirectory.Name = "ProcessingFilesDirectory"
        '
        'CaseDirectory
        '
        Me.CaseDirectory.HeaderText = "Case Directory"
        Me.CaseDirectory.Name = "CaseDirectory"
        '
        'OutputDirectory
        '
        Me.OutputDirectory.HeaderText = "Export Directory"
        Me.OutputDirectory.Name = "OutputDirectory"
        '
        'LogDirectory
        '
        Me.LogDirectory.HeaderText = "Log Directory"
        Me.LogDirectory.Name = "LogDirectory"
        '
        'ProcessingStartDate
        '
        Me.ProcessingStartDate.HeaderText = "Processing Start Date"
        Me.ProcessingStartDate.Name = "ProcessingStartDate"
        '
        'ProcessingEndDate
        '
        Me.ProcessingEndDate.HeaderText = "Processing End Date"
        Me.ProcessingEndDate.Name = "ProcessingEndDate"
        '
        'SummaryReportLocation
        '
        Me.SummaryReportLocation.HeaderText = "Summary Report"
        Me.SummaryReportLocation.Name = "SummaryReportLocation"
        '
        'btnAddBatchToGrid
        '
        Me.btnAddBatchToGrid.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAddBatchToGrid.BackColor = System.Drawing.Color.White
        Me.btnAddBatchToGrid.Image = CType(resources.GetObject("btnAddBatchToGrid.Image"), System.Drawing.Image)
        Me.btnAddBatchToGrid.Location = New System.Drawing.Point(1398, 364)
        Me.btnAddBatchToGrid.Margin = New System.Windows.Forms.Padding(2)
        Me.btnAddBatchToGrid.Name = "btnAddBatchToGrid"
        Me.btnAddBatchToGrid.Size = New System.Drawing.Size(80, 60)
        Me.btnAddBatchToGrid.TabIndex = 85
        Me.btnAddBatchToGrid.UseVisualStyleBackColor = False
        '
        'btnLaunchBatches
        '
        Me.btnLaunchBatches.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLaunchBatches.BackColor = System.Drawing.Color.White
        Me.btnLaunchBatches.Image = CType(resources.GetObject("btnLaunchBatches.Image"), System.Drawing.Image)
        Me.btnLaunchBatches.Location = New System.Drawing.Point(1398, 532)
        Me.btnLaunchBatches.Margin = New System.Windows.Forms.Padding(2)
        Me.btnLaunchBatches.Name = "btnLaunchBatches"
        Me.btnLaunchBatches.Size = New System.Drawing.Size(80, 60)
        Me.btnLaunchBatches.TabIndex = 86
        Me.btnLaunchBatches.UseVisualStyleBackColor = False
        '
        'btnExpandCollapse
        '
        Me.btnExpandCollapse.AutoSize = True
        Me.btnExpandCollapse.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExpandCollapse.Location = New System.Drawing.Point(851, 17)
        Me.btnExpandCollapse.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExpandCollapse.Name = "btnExpandCollapse"
        Me.btnExpandCollapse.Size = New System.Drawing.Size(28, 30)
        Me.btnExpandCollapse.TabIndex = 87
        Me.btnExpandCollapse.Text = "-"
        Me.btnExpandCollapse.UseVisualStyleBackColor = True
        '
        'grpArchiveType
        '
        Me.grpArchiveType.Controls.Add(Me.radJournalArchive)
        Me.grpArchiveType.Controls.Add(Me.radMailboxArchive)
        Me.grpArchiveType.Location = New System.Drawing.Point(265, 11)
        Me.grpArchiveType.Margin = New System.Windows.Forms.Padding(2)
        Me.grpArchiveType.Name = "grpArchiveType"
        Me.grpArchiveType.Padding = New System.Windows.Forms.Padding(2)
        Me.grpArchiveType.Size = New System.Drawing.Size(157, 42)
        Me.grpArchiveType.TabIndex = 88
        Me.grpArchiveType.TabStop = False
        Me.grpArchiveType.Text = "Archive Type"
        '
        'radJournalArchive
        '
        Me.radJournalArchive.AutoSize = True
        Me.radJournalArchive.Location = New System.Drawing.Point(91, 13)
        Me.radJournalArchive.Margin = New System.Windows.Forms.Padding(2)
        Me.radJournalArchive.Name = "radJournalArchive"
        Me.radJournalArchive.Size = New System.Drawing.Size(59, 17)
        Me.radJournalArchive.TabIndex = 1
        Me.radJournalArchive.TabStop = True
        Me.radJournalArchive.Text = "Journal"
        Me.radJournalArchive.UseVisualStyleBackColor = True
        '
        'radMailboxArchive
        '
        Me.radMailboxArchive.AutoSize = True
        Me.radMailboxArchive.Location = New System.Drawing.Point(9, 13)
        Me.radMailboxArchive.Margin = New System.Windows.Forms.Padding(2)
        Me.radMailboxArchive.Name = "radMailboxArchive"
        Me.radMailboxArchive.Size = New System.Drawing.Size(61, 17)
        Me.radMailboxArchive.TabIndex = 0
        Me.radMailboxArchive.TabStop = True
        Me.radMailboxArchive.Text = "Mailbox"
        Me.radMailboxArchive.UseVisualStyleBackColor = True
        '
        'grpWSSControl
        '
        Me.grpWSSControl.Controls.Add(Me.chkContact)
        Me.grpWSSControl.Controls.Add(Me.chkRSSFeed)
        Me.grpWSSControl.Controls.Add(Me.chkCalendar)
        Me.grpWSSControl.Controls.Add(Me.chkEmail)
        Me.grpWSSControl.Controls.Add(Me.lblContentFiltering)
        Me.grpWSSControl.Controls.Add(Me.btnMappingCSV)
        Me.grpWSSControl.Controls.Add(Me.txtMappingCSV)
        Me.grpWSSControl.Controls.Add(Me.lblMappingCSV)
        Me.grpWSSControl.Controls.Add(Me.grpVerbose)
        Me.grpWSSControl.Controls.Add(Me.grpExcludeItems)
        Me.grpWSSControl.Controls.Add(Me.btnCommCSVSelector)
        Me.grpWSSControl.Controls.Add(Me.lblSearchTermsCSV)
        Me.grpWSSControl.Controls.Add(Me.txtSearchTermCSV)
        Me.grpWSSControl.Controls.Add(Me.lblVerbose)
        Me.grpWSSControl.Controls.Add(Me.lblExcludeItems)
        Me.grpWSSControl.Location = New System.Drawing.Point(889, 91)
        Me.grpWSSControl.Name = "grpWSSControl"
        Me.grpWSSControl.Size = New System.Drawing.Size(439, 333)
        Me.grpWSSControl.TabIndex = 89
        Me.grpWSSControl.TabStop = False
        Me.grpWSSControl.Text = "WSS"
        '
        'chkContact
        '
        Me.chkContact.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkContact.BackColor = System.Drawing.Color.Green
        Me.chkContact.Checked = True
        Me.chkContact.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkContact.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.chkContact.Location = New System.Drawing.Point(359, 92)
        Me.chkContact.Name = "chkContact"
        Me.chkContact.Size = New System.Drawing.Size(70, 23)
        Me.chkContact.TabIndex = 57
        Me.chkContact.Text = "Contact"
        Me.chkContact.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.chkContact.UseVisualStyleBackColor = False
        '
        'chkRSSFeed
        '
        Me.chkRSSFeed.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkRSSFeed.BackColor = System.Drawing.Color.Green
        Me.chkRSSFeed.Checked = True
        Me.chkRSSFeed.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkRSSFeed.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.chkRSSFeed.Location = New System.Drawing.Point(223, 92)
        Me.chkRSSFeed.Name = "chkRSSFeed"
        Me.chkRSSFeed.Size = New System.Drawing.Size(70, 23)
        Me.chkRSSFeed.TabIndex = 56
        Me.chkRSSFeed.Text = "RSS Feed"
        Me.chkRSSFeed.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.chkRSSFeed.UseVisualStyleBackColor = False
        '
        'chkCalendar
        '
        Me.chkCalendar.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkCalendar.BackColor = System.Drawing.Color.Green
        Me.chkCalendar.Checked = True
        Me.chkCalendar.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCalendar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.chkCalendar.Location = New System.Drawing.Point(291, 92)
        Me.chkCalendar.Name = "chkCalendar"
        Me.chkCalendar.Size = New System.Drawing.Size(70, 23)
        Me.chkCalendar.TabIndex = 55
        Me.chkCalendar.Text = "Calendar"
        Me.chkCalendar.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.chkCalendar.UseVisualStyleBackColor = False
        '
        'chkEmail
        '
        Me.chkEmail.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkEmail.BackColor = System.Drawing.Color.Green
        Me.chkEmail.Checked = True
        Me.chkEmail.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkEmail.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.chkEmail.Location = New System.Drawing.Point(154, 92)
        Me.chkEmail.Name = "chkEmail"
        Me.chkEmail.Size = New System.Drawing.Size(70, 23)
        Me.chkEmail.TabIndex = 54
        Me.chkEmail.Text = "Email"
        Me.chkEmail.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.chkEmail.UseVisualStyleBackColor = False
        '
        'lblContentFiltering
        '
        Me.lblContentFiltering.AutoSize = True
        Me.lblContentFiltering.Location = New System.Drawing.Point(3, 93)
        Me.lblContentFiltering.Name = "lblContentFiltering"
        Me.lblContentFiltering.Size = New System.Drawing.Size(86, 13)
        Me.lblContentFiltering.TabIndex = 48
        Me.lblContentFiltering.Text = "Content Filtering:"
        '
        'btnMappingCSV
        '
        Me.btnMappingCSV.Location = New System.Drawing.Point(402, 155)
        Me.btnMappingCSV.Name = "btnMappingCSV"
        Me.btnMappingCSV.Size = New System.Drawing.Size(25, 21)
        Me.btnMappingCSV.TabIndex = 47
        Me.btnMappingCSV.Text = "..."
        Me.btnMappingCSV.UseVisualStyleBackColor = True
        '
        'txtMappingCSV
        '
        Me.txtMappingCSV.Location = New System.Drawing.Point(154, 156)
        Me.txtMappingCSV.Name = "txtMappingCSV"
        Me.txtMappingCSV.Size = New System.Drawing.Size(242, 20)
        Me.txtMappingCSV.TabIndex = 46
        '
        'lblMappingCSV
        '
        Me.lblMappingCSV.AutoSize = True
        Me.lblMappingCSV.Location = New System.Drawing.Point(4, 159)
        Me.lblMappingCSV.Name = "lblMappingCSV"
        Me.lblMappingCSV.Size = New System.Drawing.Size(75, 13)
        Me.lblMappingCSV.TabIndex = 45
        Me.lblMappingCSV.Text = "Mapping CSV:"
        '
        'grpVerbose
        '
        Me.grpVerbose.Controls.Add(Me.radVerboseFalse)
        Me.grpVerbose.Controls.Add(Me.radVerboseTrue)
        Me.grpVerbose.Location = New System.Drawing.Point(155, 51)
        Me.grpVerbose.Name = "grpVerbose"
        Me.grpVerbose.Size = New System.Drawing.Size(83, 26)
        Me.grpVerbose.TabIndex = 37
        Me.grpVerbose.TabStop = False
        '
        'radVerboseFalse
        '
        Me.radVerboseFalse.Appearance = System.Windows.Forms.Appearance.Button
        Me.radVerboseFalse.AutoSize = True
        Me.radVerboseFalse.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.radVerboseFalse.Location = New System.Drawing.Point(37, 0)
        Me.radVerboseFalse.Name = "radVerboseFalse"
        Me.radVerboseFalse.Size = New System.Drawing.Size(44, 25)
        Me.radVerboseFalse.TabIndex = 8
        Me.radVerboseFalse.TabStop = True
        Me.radVerboseFalse.Text = "False"
        Me.radVerboseFalse.UseVisualStyleBackColor = True
        '
        'radVerboseTrue
        '
        Me.radVerboseTrue.Appearance = System.Windows.Forms.Appearance.Button
        Me.radVerboseTrue.AutoSize = True
        Me.radVerboseTrue.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.radVerboseTrue.Location = New System.Drawing.Point(0, 0)
        Me.radVerboseTrue.Name = "radVerboseTrue"
        Me.radVerboseTrue.Size = New System.Drawing.Size(41, 25)
        Me.radVerboseTrue.TabIndex = 7
        Me.radVerboseTrue.TabStop = True
        Me.radVerboseTrue.Text = "True"
        Me.radVerboseTrue.UseVisualStyleBackColor = True
        '
        'grpExcludeItems
        '
        Me.grpExcludeItems.Controls.Add(Me.radExcludeItemsFalse)
        Me.grpExcludeItems.Controls.Add(Me.radExcludeItemsTrue)
        Me.grpExcludeItems.Location = New System.Drawing.Point(155, 23)
        Me.grpExcludeItems.Name = "grpExcludeItems"
        Me.grpExcludeItems.Size = New System.Drawing.Size(81, 28)
        Me.grpExcludeItems.TabIndex = 35
        Me.grpExcludeItems.TabStop = False
        '
        'radExcludeItemsFalse
        '
        Me.radExcludeItemsFalse.Appearance = System.Windows.Forms.Appearance.Button
        Me.radExcludeItemsFalse.AutoSize = True
        Me.radExcludeItemsFalse.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.radExcludeItemsFalse.Location = New System.Drawing.Point(37, 0)
        Me.radExcludeItemsFalse.Name = "radExcludeItemsFalse"
        Me.radExcludeItemsFalse.Size = New System.Drawing.Size(44, 25)
        Me.radExcludeItemsFalse.TabIndex = 6
        Me.radExcludeItemsFalse.TabStop = True
        Me.radExcludeItemsFalse.Text = "False"
        Me.radExcludeItemsFalse.UseVisualStyleBackColor = True
        '
        'radExcludeItemsTrue
        '
        Me.radExcludeItemsTrue.Appearance = System.Windows.Forms.Appearance.Button
        Me.radExcludeItemsTrue.AutoSize = True
        Me.radExcludeItemsTrue.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.radExcludeItemsTrue.Location = New System.Drawing.Point(0, 0)
        Me.radExcludeItemsTrue.Name = "radExcludeItemsTrue"
        Me.radExcludeItemsTrue.Size = New System.Drawing.Size(41, 25)
        Me.radExcludeItemsTrue.TabIndex = 5
        Me.radExcludeItemsTrue.TabStop = True
        Me.radExcludeItemsTrue.Text = "True"
        Me.radExcludeItemsTrue.UseVisualStyleBackColor = True
        '
        'btnCommCSVSelector
        '
        Me.btnCommCSVSelector.Location = New System.Drawing.Point(402, 129)
        Me.btnCommCSVSelector.Name = "btnCommCSVSelector"
        Me.btnCommCSVSelector.Size = New System.Drawing.Size(25, 20)
        Me.btnCommCSVSelector.TabIndex = 33
        Me.btnCommCSVSelector.Text = "..."
        Me.btnCommCSVSelector.UseVisualStyleBackColor = True
        '
        'lblSearchTermsCSV
        '
        Me.lblSearchTermsCSV.AutoSize = True
        Me.lblSearchTermsCSV.Location = New System.Drawing.Point(5, 132)
        Me.lblSearchTermsCSV.Name = "lblSearchTermsCSV"
        Me.lblSearchTermsCSV.Size = New System.Drawing.Size(100, 13)
        Me.lblSearchTermsCSV.TabIndex = 30
        Me.lblSearchTermsCSV.Text = "Search Terms CSV:"
        '
        'txtSearchTermCSV
        '
        Me.txtSearchTermCSV.Location = New System.Drawing.Point(154, 125)
        Me.txtSearchTermCSV.Name = "txtSearchTermCSV"
        Me.txtSearchTermCSV.Size = New System.Drawing.Size(242, 20)
        Me.txtSearchTermCSV.TabIndex = 29
        '
        'lblVerbose
        '
        Me.lblVerbose.AutoSize = True
        Me.lblVerbose.Location = New System.Drawing.Point(5, 52)
        Me.lblVerbose.Name = "lblVerbose"
        Me.lblVerbose.Size = New System.Drawing.Size(90, 13)
        Me.lblVerbose.TabIndex = 21
        Me.lblVerbose.Text = "Verbose Logging:"
        '
        'lblExcludeItems
        '
        Me.lblExcludeItems.AutoSize = True
        Me.lblExcludeItems.Location = New System.Drawing.Point(5, 27)
        Me.lblExcludeItems.Name = "lblExcludeItems"
        Me.lblExcludeItems.Size = New System.Drawing.Size(144, 13)
        Me.lblExcludeItems.TabIndex = 19
        Me.lblExcludeItems.Text = "Exclude Unresponsive Items:"
        '
        'dateFromDate
        '
        Me.dateFromDate.Location = New System.Drawing.Point(68, 63)
        Me.dateFromDate.Name = "dateFromDate"
        Me.dateFromDate.Size = New System.Drawing.Size(200, 20)
        Me.dateFromDate.TabIndex = 90
        '
        'lblFromDate
        '
        Me.lblFromDate.AutoSize = True
        Me.lblFromDate.Location = New System.Drawing.Point(8, 68)
        Me.lblFromDate.Name = "lblFromDate"
        Me.lblFromDate.Size = New System.Drawing.Size(59, 13)
        Me.lblFromDate.TabIndex = 91
        Me.lblFromDate.Text = "From Date:"
        '
        'lblToDate
        '
        Me.lblToDate.AutoSize = True
        Me.lblToDate.Location = New System.Drawing.Point(277, 68)
        Me.lblToDate.Name = "lblToDate"
        Me.lblToDate.Size = New System.Drawing.Size(49, 13)
        Me.lblToDate.TabIndex = 92
        Me.lblToDate.Text = "To Date:"
        '
        'dateToDate
        '
        Me.dateToDate.Location = New System.Drawing.Point(332, 62)
        Me.dateToDate.Name = "dateToDate"
        Me.dateToDate.Size = New System.Drawing.Size(200, 20)
        Me.dateToDate.TabIndex = 93
        '
        'btnLoadPreviousBatches
        '
        Me.btnLoadPreviousBatches.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnLoadPreviousBatches.BackColor = System.Drawing.Color.White
        Me.btnLoadPreviousBatches.Image = CType(resources.GetObject("btnLoadPreviousBatches.Image"), System.Drawing.Image)
        Me.btnLoadPreviousBatches.Location = New System.Drawing.Point(90, 532)
        Me.btnLoadPreviousBatches.Name = "btnLoadPreviousBatches"
        Me.btnLoadPreviousBatches.Size = New System.Drawing.Size(80, 60)
        Me.btnLoadPreviousBatches.TabIndex = 95
        Me.btnLoadPreviousBatches.UseVisualStyleBackColor = False
        '
        'NuixLogoPicture
        '
        Me.NuixLogoPicture.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.NuixLogoPicture.Image = CType(resources.GetObject("NuixLogoPicture.Image"), System.Drawing.Image)
        Me.NuixLogoPicture.Location = New System.Drawing.Point(1401, 16)
        Me.NuixLogoPicture.Name = "NuixLogoPicture"
        Me.NuixLogoPicture.Size = New System.Drawing.Size(75, 100)
        Me.NuixLogoPicture.TabIndex = 96
        Me.NuixLogoPicture.TabStop = False
        '
        'ArchiveExtractToolTip
        '
        Me.ArchiveExtractToolTip.IsBalloon = True
        '
        'btnExportProcessingDetails
        '
        Me.btnExportProcessingDetails.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportProcessingDetails.BackColor = System.Drawing.Color.White
        Me.btnExportProcessingDetails.Image = CType(resources.GetObject("btnExportProcessingDetails.Image"), System.Drawing.Image)
        Me.btnExportProcessingDetails.Location = New System.Drawing.Point(1314, 532)
        Me.btnExportProcessingDetails.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExportProcessingDetails.Name = "btnExportProcessingDetails"
        Me.btnExportProcessingDetails.Size = New System.Drawing.Size(80, 60)
        Me.btnExportProcessingDetails.TabIndex = 97
        Me.btnExportProcessingDetails.UseVisualStyleBackColor = False
        '
        'btnConsolidateExporterFiles
        '
        Me.btnConsolidateExporterFiles.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnConsolidateExporterFiles.Image = CType(resources.GetObject("btnConsolidateExporterFiles.Image"), System.Drawing.Image)
        Me.btnConsolidateExporterFiles.Location = New System.Drawing.Point(1229, 532)
        Me.btnConsolidateExporterFiles.Name = "btnConsolidateExporterFiles"
        Me.btnConsolidateExporterFiles.Size = New System.Drawing.Size(80, 60)
        Me.btnConsolidateExporterFiles.TabIndex = 98
        Me.btnConsolidateExporterFiles.UseVisualStyleBackColor = True
        '
        'ConsolidateExporterFiles
        '
        Me.ConsolidateExporterFiles.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ConsolidateExporterMetricsFilesToolStripMenuItem, Me.ConolidateExporterErrorsFilesToolStripMenuItem})
        Me.ConsolidateExporterFiles.Name = "ContextMenuStrip2"
        Me.ConsolidateExporterFiles.Size = New System.Drawing.Size(254, 48)
        '
        'ConsolidateExporterMetricsFilesToolStripMenuItem
        '
        Me.ConsolidateExporterMetricsFilesToolStripMenuItem.Name = "ConsolidateExporterMetricsFilesToolStripMenuItem"
        Me.ConsolidateExporterMetricsFilesToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.ConsolidateExporterMetricsFilesToolStripMenuItem.Text = "Consolidate Exporter-Metrics Files"
        '
        'ConolidateExporterErrorsFilesToolStripMenuItem
        '
        Me.ConolidateExporterErrorsFilesToolStripMenuItem.Name = "ConolidateExporterErrorsFilesToolStripMenuItem"
        Me.ConolidateExporterErrorsFilesToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.ConolidateExporterErrorsFilesToolStripMenuItem.Text = "Conolidate Exporter-Errors Files"
        '
        'LegacyArchiveExtraction
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(1484, 661)
        Me.Controls.Add(Me.btnConsolidateExporterFiles)
        Me.Controls.Add(Me.btnExportProcessingDetails)
        Me.Controls.Add(Me.btnClearDates)
        Me.Controls.Add(Me.grpEVSettings)
        Me.Controls.Add(Me.NuixLogoPicture)
        Me.Controls.Add(Me.btnLoadPreviousBatches)
        Me.Controls.Add(Me.dateToDate)
        Me.Controls.Add(Me.lblToDate)
        Me.Controls.Add(Me.lblFromDate)
        Me.Controls.Add(Me.dateFromDate)
        Me.Controls.Add(Me.grpSQLArchiveSettings)
        Me.Controls.Add(Me.grpEMCSettings)
        Me.Controls.Add(Me.grpEASSettings)
        Me.Controls.Add(Me.grpAXSOneSettings)
        Me.Controls.Add(Me.grpAXSOneSISSettings)
        Me.Controls.Add(Me.grpWSSControl)
        Me.Controls.Add(Me.grpArchiveType)
        Me.Controls.Add(Me.btnExpandCollapse)
        Me.Controls.Add(Me.btnLaunchBatches)
        Me.Controls.Add(Me.btnAddBatchToGrid)
        Me.Controls.Add(Me.grdArchiveExtractionBatch)
        Me.Controls.Add(Me.grpSourceInformation)
        Me.Controls.Add(Me.btnShowSettings)
        Me.Controls.Add(Me.grpOutputFormat)
        Me.Controls.Add(Me.btnCustodianListCSVChooser)
        Me.Controls.Add(Me.lblCustodianListCSV)
        Me.Controls.Add(Me.txtCustodianListCSV)
        Me.Controls.Add(Me.grpExtractionType)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboArchiveType)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(1500, 700)
        Me.Name = "LegacyArchiveExtraction"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Legacy Email Archive Extraction"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.grpExtractionType.ResumeLayout(False)
        Me.grpExtractionType.PerformLayout()
        Me.grpAXSOneSISSettings.ResumeLayout(False)
        Me.grpSQLArchiveSettings.ResumeLayout(False)
        Me.grpSQLArchiveSettings.PerformLayout()
        Me.grpEMCSettings.ResumeLayout(False)
        Me.grpEMCSettings.PerformLayout()
        Me.grpEVSettings.ResumeLayout(False)
        Me.grpEVSettings.PerformLayout()
        Me.grpAXSOneSettings.ResumeLayout(False)
        Me.grpAXSOneSettings.PerformLayout()
        Me.grpEASSettings.ResumeLayout(False)
        Me.grpEASSettings.PerformLayout()
        Me.grpOutputFormat.ResumeLayout(False)
        Me.grpOutputFormat.PerformLayout()
        Me.grpSourceInformation.ResumeLayout(False)
        Me.grpSourceInformation.PerformLayout()
        Me.grpSourceLocation.ResumeLayout(False)
        Me.grpSourceLocation.PerformLayout()
        CType(Me.grdArchiveExtractionBatch, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpArchiveType.ResumeLayout(False)
        Me.grpArchiveType.PerformLayout()
        Me.grpWSSControl.ResumeLayout(False)
        Me.grpWSSControl.PerformLayout()
        Me.grpVerbose.ResumeLayout(False)
        Me.grpVerbose.PerformLayout()
        Me.grpExcludeItems.ResumeLayout(False)
        Me.grpExcludeItems.PerformLayout()
        CType(Me.NuixLogoPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ConsolidateExporterFiles.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboArchiveType As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents grpExtractionType As System.Windows.Forms.GroupBox
    Friend WithEvents cboExtractionOutputType As System.Windows.Forms.ComboBox
    Friend WithEvents grpAXSOneSISSettings As System.Windows.Forms.GroupBox
    Friend WithEvents grpSQLArchiveSettings As System.Windows.Forms.GroupBox
    Friend WithEvents txtDomain As System.Windows.Forms.TextBox
    Friend WithEvents lblDomain As System.Windows.Forms.Label
    Friend WithEvents txtSQLInfo As System.Windows.Forms.TextBox
    Friend WithEvents lblSQLInfo As System.Windows.Forms.Label
    Friend WithEvents txtSQLUserName As System.Windows.Forms.TextBox
    Friend WithEvents lblSQLUserName As System.Windows.Forms.Label
    Friend WithEvents txtSQLDBName As System.Windows.Forms.TextBox
    Friend WithEvents lblSQLDBName As System.Windows.Forms.Label
    Friend WithEvents txtSQLPortNumber As System.Windows.Forms.TextBox
    Friend WithEvents lblPortNumber As System.Windows.Forms.Label
    Friend WithEvents txtSQLHostName As System.Windows.Forms.TextBox
    Friend WithEvents lblServerName As System.Windows.Forms.Label
    Friend WithEvents grpEMCSettings As System.Windows.Forms.GroupBox
    Friend WithEvents lblAddressFiltering As System.Windows.Forms.Label
    Friend WithEvents grpEVSettings As System.Windows.Forms.GroupBox
    Friend WithEvents chkSkipAdditionalSQLLookup As System.Windows.Forms.CheckBox
    Friend WithEvents chkSkipVaultStorePartitionErrors As System.Windows.Forms.CheckBox
    Friend WithEvents grpEASSettings As System.Windows.Forms.GroupBox
    Friend WithEvents txtDocServerID As System.Windows.Forms.TextBox
    Friend WithEvents lblDocServerID As System.Windows.Forms.Label
    Friend WithEvents txtCustodianListCSV As System.Windows.Forms.TextBox
    Friend WithEvents lblCustodianListCSV As System.Windows.Forms.Label
    Friend WithEvents btnCustodianListCSVChooser As System.Windows.Forms.Button
    Friend WithEvents grpOutputFormat As System.Windows.Forms.GroupBox
    Friend WithEvents radFlat As System.Windows.Forms.RadioButton
    Friend WithEvents radUser As System.Windows.Forms.RadioButton
    Friend WithEvents btnShowSettings As System.Windows.Forms.Button
    Friend WithEvents grpSourceInformation As System.Windows.Forms.GroupBox
    Friend WithEvents treeViewFolders As System.Windows.Forms.TreeView
    Friend WithEvents grpSourceLocation As System.Windows.Forms.GroupBox
    Friend WithEvents radAddFolder As System.Windows.Forms.RadioButton
    Friend WithEvents radAddFile As System.Windows.Forms.RadioButton
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents treeAXSOneSis As System.Windows.Forms.TreeView
    Friend WithEvents grdArchiveExtractionBatch As System.Windows.Forms.DataGridView
    Friend WithEvents btnAddBatchToGrid As System.Windows.Forms.Button
    Friend WithEvents btnLaunchBatches As System.Windows.Forms.Button
    Friend WithEvents btnExpandCollapse As System.Windows.Forms.Button
    Friend WithEvents ImgIcons As System.Windows.Forms.ImageList
    Friend WithEvents chkComputeBatchSize As System.Windows.Forms.CheckBox
    Friend WithEvents lblBatchSize As System.Windows.Forms.Label
    Friend WithEvents lblTotalBatchSize As System.Windows.Forms.Label
    Friend WithEvents grpArchiveType As System.Windows.Forms.GroupBox
    Friend WithEvents radJournalArchive As System.Windows.Forms.RadioButton
    Friend WithEvents radMailboxArchive As System.Windows.Forms.RadioButton
    Friend WithEvents grpWSSControl As System.Windows.Forms.GroupBox
    Friend WithEvents radVerboseFalse As System.Windows.Forms.RadioButton
    Friend WithEvents radVerboseTrue As System.Windows.Forms.RadioButton
    Friend WithEvents radExcludeItemsFalse As System.Windows.Forms.RadioButton
    Friend WithEvents radExcludeItemsTrue As System.Windows.Forms.RadioButton
    Friend WithEvents btnCommCSVSelector As System.Windows.Forms.Button
    Friend WithEvents lblSearchTermsCSV As System.Windows.Forms.Label
    Friend WithEvents txtSearchTermCSV As System.Windows.Forms.TextBox
    Friend WithEvents lblVerbose As System.Windows.Forms.Label
    Friend WithEvents lblExcludeItems As System.Windows.Forms.Label
    Friend WithEvents grpExcludeItems As System.Windows.Forms.GroupBox
    Friend WithEvents grpVerbose As System.Windows.Forms.GroupBox
    Friend WithEvents lblExpandDLto As System.Windows.Forms.Label
    Friend WithEvents cboExpandDLLocation As System.Windows.Forms.ComboBox
    Friend WithEvents grpAXSOneSettings As System.Windows.Forms.GroupBox
    Friend WithEvents chkAXSOneSkipSISLookups As System.Windows.Forms.CheckBox
    Friend WithEvents btnMappingCSV As System.Windows.Forms.Button
    Friend WithEvents txtMappingCSV As System.Windows.Forms.TextBox
    Friend WithEvents lblMappingCSV As System.Windows.Forms.Label
    Friend WithEvents dateFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblFromDate As System.Windows.Forms.Label
    Friend WithEvents lblToDate As System.Windows.Forms.Label
    Friend WithEvents dateToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnClearDates As System.Windows.Forms.Button
    Friend WithEvents btnLoadPreviousBatches As System.Windows.Forms.Button
    Friend WithEvents lblDomainCheck As System.Windows.Forms.Label
    Friend WithEvents lblSQLPWCheck As System.Windows.Forms.Label
    Friend WithEvents lblSQLUserNameCheck As System.Windows.Forms.Label
    Friend WithEvents lblDBNameCheck As System.Windows.Forms.Label
    Friend WithEvents lblHostCheck As System.Windows.Forms.Label
    Friend WithEvents btnTestSQLConnection As System.Windows.Forms.Button
    Friend WithEvents chkAFEX As System.Windows.Forms.CheckBox
    Friend WithEvents chkAFPST As System.Windows.Forms.CheckBox
    Friend WithEvents chkAFSYS As System.Windows.Forms.CheckBox
    Friend WithEvents lblContentFiltering As System.Windows.Forms.Label
    Friend WithEvents chkContact As System.Windows.Forms.CheckBox
    Friend WithEvents chkRSSFeed As System.Windows.Forms.CheckBox
    Friend WithEvents chkCalendar As System.Windows.Forms.CheckBox
    Friend WithEvents chkEmail As System.Windows.Forms.CheckBox
    Friend WithEvents NuixLogoPicture As System.Windows.Forms.PictureBox
    Friend WithEvents radCentera As System.Windows.Forms.RadioButton
    Friend WithEvents btnIPFile As System.Windows.Forms.Button
    Friend WithEvents txtIPFile As System.Windows.Forms.TextBox
    Friend WithEvents lblIPFile As System.Windows.Forms.Label
    Friend WithEvents btnPEAFile As System.Windows.Forms.Button
    Friend WithEvents txtPEAFile As System.Windows.Forms.TextBox
    Friend WithEvents lblPEAFile As System.Windows.Forms.Label
    Friend WithEvents ArchiveExtractToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents btnEVUserListFileChooser As System.Windows.Forms.Button
    Friend WithEvents txtEVUserList As System.Windows.Forms.TextBox
    Friend WithEvents lblEVUserList As System.Windows.Forms.Label
    Friend WithEvents btnExportProcessingDetails As System.Windows.Forms.Button
    Friend WithEvents radLoose As System.Windows.Forms.RadioButton
    Friend WithEvents radZipped As System.Windows.Forms.RadioButton
    Friend WithEvents StatusImage As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents SelectBatch As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents BatchName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PercentCompleted As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BytesProcessed As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TotalBytes As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ExtractionStatus As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ItemsProcessed As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ArchiveName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ArchiveSettings As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ArchiveType As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents LightspeedOutputType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SQLConnectionInfo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SourceInformation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents WSSSettings As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ItemsExported As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ItemsFailed As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ItemsSkipped As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProcessingFilesDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CaseDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OutputDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LogDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProcessingStartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProcessingEndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SummaryReportLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnConsolidateExporterFiles As System.Windows.Forms.Button
    Friend WithEvents ConsolidateExporterFiles As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ConsolidateExporterMetricsFilesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ConolidateExporterErrorsFilesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblAuthentication As System.Windows.Forms.Label
    Friend WithEvents cboSecurityType As System.Windows.Forms.ComboBox
End Class
