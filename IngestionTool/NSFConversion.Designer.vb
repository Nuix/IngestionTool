<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NSFConversion
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(NSFConversion))
        Me.txtNSFLocation = New System.Windows.Forms.TextBox()
        Me.btnLoadPSTInfo = New System.Windows.Forms.Button()
        Me.btnFileSystemChooser = New System.Windows.Forms.Button()
        Me.lblPSTLocation = New System.Windows.Forms.Label()
        Me.grdConversionDataInfo = New System.Windows.Forms.DataGridView()
        Me.StatusImage = New System.Windows.Forms.DataGridViewImageColumn()
        Me.SelectCustodian = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.CustodianName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DateRangeFilters = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PercentComplete = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BytesProcessed = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ConversionStatus = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.GroupID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NumberOfSourceFiles = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SizeOfSourceFiles = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SourcePath = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SourceDataType = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SourceFormat = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OutputDataType = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OutputFormat = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RedisSettings = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ConversionStartTime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ConversionEndTime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProcessID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProcessingFilesDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CaseDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OutputDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LogDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NumberSuccess = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NumberFailed = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SummaryReportLocation = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnLaunchSourceConversionBatches = New System.Windows.Forms.Button()
        Me.btnShowSettings = New System.Windows.Forms.Button()
        Me.cboSourceDataType = New System.Windows.Forms.ComboBox()
        Me.lblSourceDataType = New System.Windows.Forms.Label()
        Me.radSourceUser = New System.Windows.Forms.RadioButton()
        Me.grpSourceFormat = New System.Windows.Forms.GroupBox()
        Me.radSourceFlat = New System.Windows.Forms.RadioButton()
        Me.grpOutputFormat = New System.Windows.Forms.GroupBox()
        Me.radOutputFlat = New System.Windows.Forms.RadioButton()
        Me.radOutputUser = New System.Windows.Forms.RadioButton()
        Me.lblOutputDateType = New System.Windows.Forms.Label()
        Me.cboOutputDataType = New System.Windows.Forms.ComboBox()
        Me.btnLoadCustodianMappingData = New System.Windows.Forms.Button()
        Me.btnCustomerMappingFile = New System.Windows.Forms.Button()
        Me.txtCustodianMappingFile = New System.Windows.Forms.TextBox()
        Me.lblCustomerMappingFile = New System.Windows.Forms.Label()
        Me.cboGroupIDs = New System.Windows.Forms.ComboBox()
        Me.lblSelectByGroupID = New System.Windows.Forms.Label()
        Me.btnSelectAll = New System.Windows.Forms.Button()
        Me.btnClearDates = New System.Windows.Forms.Button()
        Me.dateToDate = New System.Windows.Forms.DateTimePicker()
        Me.lblToDate = New System.Windows.Forms.Label()
        Me.lblFromDate = New System.Windows.Forms.Label()
        Me.dateFromDate = New System.Windows.Forms.DateTimePicker()
        Me.chkDedupData = New System.Windows.Forms.CheckBox()
        Me.FileSystemWatcher1 = New System.IO.FileSystemWatcher()
        Me.NuixLogoPicture = New System.Windows.Forms.PictureBox()
        Me.EmailConversionToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnReloadEmailConversionData = New System.Windows.Forms.Button()
        Me.btnExportProcessingDetails = New System.Windows.Forms.Button()
        Me.btnAddDateRangeFilter = New System.Windows.Forms.Button()
        Me.lblFilters = New System.Windows.Forms.Label()
        Me.btnAddToSelected = New System.Windows.Forms.Button()
        Me.lstDateRangeFilters = New System.Windows.Forms.ListBox()
        Me.ConsolidateExporterFiles = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ConsolidateExporterMetricsFilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ConsolidateExporterErrorsFilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.btnConsolidateExporterFiles = New System.Windows.Forms.Button()
        CType(Me.grdConversionDataInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpSourceFormat.SuspendLayout()
        Me.grpOutputFormat.SuspendLayout()
        CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NuixLogoPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ConsolidateExporterFiles.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtNSFLocation
        '
        Me.txtNSFLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNSFLocation.Location = New System.Drawing.Point(731, 25)
        Me.txtNSFLocation.Name = "txtNSFLocation"
        Me.txtNSFLocation.Size = New System.Drawing.Size(240, 20)
        Me.txtNSFLocation.TabIndex = 75
        '
        'btnLoadPSTInfo
        '
        Me.btnLoadPSTInfo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLoadPSTInfo.Location = New System.Drawing.Point(1013, 20)
        Me.btnLoadPSTInfo.Name = "btnLoadPSTInfo"
        Me.btnLoadPSTInfo.Size = New System.Drawing.Size(121, 30)
        Me.btnLoadPSTInfo.TabIndex = 74
        Me.btnLoadPSTInfo.Text = "Load Source Info"
        Me.btnLoadPSTInfo.UseVisualStyleBackColor = True
        '
        'btnFileSystemChooser
        '
        Me.btnFileSystemChooser.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFileSystemChooser.Location = New System.Drawing.Point(977, 24)
        Me.btnFileSystemChooser.Name = "btnFileSystemChooser"
        Me.btnFileSystemChooser.Size = New System.Drawing.Size(30, 22)
        Me.btnFileSystemChooser.TabIndex = 73
        Me.btnFileSystemChooser.Text = "..."
        Me.btnFileSystemChooser.UseVisualStyleBackColor = True
        '
        'lblPSTLocation
        '
        Me.lblPSTLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblPSTLocation.AutoSize = True
        Me.lblPSTLocation.Location = New System.Drawing.Point(587, 24)
        Me.lblPSTLocation.Name = "lblPSTLocation"
        Me.lblPSTLocation.Size = New System.Drawing.Size(138, 13)
        Me.lblPSTLocation.TabIndex = 72
        Me.lblPSTLocation.Text = "Custodian Source Location:"
        '
        'grdConversionDataInfo
        '
        Me.grdConversionDataInfo.AllowUserToAddRows = False
        Me.grdConversionDataInfo.AllowUserToDeleteRows = False
        Me.grdConversionDataInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grdConversionDataInfo.BackgroundColor = System.Drawing.SystemColors.Control
        Me.grdConversionDataInfo.ColumnHeadersHeight = 60
        Me.grdConversionDataInfo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.StatusImage, Me.SelectCustodian, Me.CustodianName, Me.DateRangeFilters, Me.PercentComplete, Me.BytesProcessed, Me.ConversionStatus, Me.GroupID, Me.NumberOfSourceFiles, Me.SizeOfSourceFiles, Me.SourcePath, Me.SourceDataType, Me.SourceFormat, Me.OutputDataType, Me.OutputFormat, Me.RedisSettings, Me.ConversionStartTime, Me.ConversionEndTime, Me.ProcessID, Me.ProcessingFilesDirectory, Me.CaseDirectory, Me.OutputDirectory, Me.LogDirectory, Me.NumberSuccess, Me.NumberFailed, Me.SummaryReportLocation})
        Me.grdConversionDataInfo.GridColor = System.Drawing.SystemColors.ControlLight
        Me.grdConversionDataInfo.Location = New System.Drawing.Point(12, 214)
        Me.grdConversionDataInfo.Name = "grdConversionDataInfo"
        Me.grdConversionDataInfo.RowTemplate.Height = 28
        Me.grdConversionDataInfo.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.grdConversionDataInfo.Size = New System.Drawing.Size(1215, 341)
        Me.grdConversionDataInfo.TabIndex = 78
        '
        'StatusImage
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.StatusImage.DefaultCellStyle = DataGridViewCellStyle1
        Me.StatusImage.Frozen = True
        Me.StatusImage.HeaderText = ""
        Me.StatusImage.Image = Global.NEAMM.My.Resources.Resources.not_selected_small
        Me.StatusImage.Name = "StatusImage"
        Me.StatusImage.ReadOnly = True
        Me.StatusImage.Width = 30
        '
        'SelectCustodian
        '
        Me.SelectCustodian.Frozen = True
        Me.SelectCustodian.HeaderText = "Select Custodian"
        Me.SelectCustodian.Name = "SelectCustodian"
        Me.SelectCustodian.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.SelectCustodian.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.SelectCustodian.Width = 55
        '
        'CustodianName
        '
        Me.CustodianName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.CustodianName.Frozen = True
        Me.CustodianName.HeaderText = "Custodian Name"
        Me.CustodianName.Name = "CustodianName"
        Me.CustodianName.ReadOnly = True
        Me.CustodianName.Width = 101
        '
        'DateRangeFilters
        '
        Me.DateRangeFilters.Frozen = True
        Me.DateRangeFilters.HeaderText = "Date Range Filters"
        Me.DateRangeFilters.Name = "DateRangeFilters"
        '
        'PercentComplete
        '
        Me.PercentComplete.Frozen = True
        Me.PercentComplete.HeaderText = "Percent Completed"
        Me.PercentComplete.Name = "PercentComplete"
        Me.PercentComplete.Width = 60
        '
        'BytesProcessed
        '
        Me.BytesProcessed.Frozen = True
        Me.BytesProcessed.HeaderText = "Bytes Processed"
        Me.BytesProcessed.Name = "BytesProcessed"
        Me.BytesProcessed.ReadOnly = True
        Me.BytesProcessed.Width = 75
        '
        'ConversionStatus
        '
        Me.ConversionStatus.Frozen = True
        Me.ConversionStatus.HeaderText = "Conversion Status"
        Me.ConversionStatus.Items.AddRange(New Object() {"Conversion Not Started", "Conversion In Progress", "Conversion Completed", "Conversion Failed", "Conversion Cancelled by User", "Restart Conversion"})
        Me.ConversionStatus.Name = "ConversionStatus"
        Me.ConversionStatus.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ConversionStatus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'GroupID
        '
        Me.GroupID.Frozen = True
        Me.GroupID.HeaderText = "Group ID"
        Me.GroupID.Name = "GroupID"
        Me.GroupID.Width = 50
        '
        'NumberOfSourceFiles
        '
        Me.NumberOfSourceFiles.HeaderText = "Number Of Source Files"
        Me.NumberOfSourceFiles.Name = "NumberOfSourceFiles"
        Me.NumberOfSourceFiles.ReadOnly = True
        Me.NumberOfSourceFiles.Width = 50
        '
        'SizeOfSourceFiles
        '
        Me.SizeOfSourceFiles.HeaderText = "Total Size Of Source Files"
        Me.SizeOfSourceFiles.Name = "SizeOfSourceFiles"
        Me.SizeOfSourceFiles.ReadOnly = True
        Me.SizeOfSourceFiles.Width = 75
        '
        'SourcePath
        '
        Me.SourcePath.HeaderText = "Source Path"
        Me.SourcePath.Name = "SourcePath"
        '
        'SourceDataType
        '
        Me.SourceDataType.HeaderText = "Source Data Type"
        Me.SourceDataType.Name = "SourceDataType"
        Me.SourceDataType.Width = 50
        '
        'SourceFormat
        '
        Me.SourceFormat.HeaderText = "Source Format"
        Me.SourceFormat.Name = "SourceFormat"
        Me.SourceFormat.Width = 50
        '
        'OutputDataType
        '
        Me.OutputDataType.HeaderText = "Output Data Type"
        Me.OutputDataType.Name = "OutputDataType"
        Me.OutputDataType.Width = 50
        '
        'OutputFormat
        '
        Me.OutputFormat.HeaderText = "Output Format"
        Me.OutputFormat.Name = "OutputFormat"
        Me.OutputFormat.Width = 50
        '
        'RedisSettings
        '
        Me.RedisSettings.HeaderText = "Redis Settings"
        Me.RedisSettings.Name = "RedisSettings"
        '
        'ConversionStartTime
        '
        Me.ConversionStartTime.HeaderText = "Conversion Start Time"
        Me.ConversionStartTime.Name = "ConversionStartTime"
        '
        'ConversionEndTime
        '
        Me.ConversionEndTime.HeaderText = "Conversion End Time"
        Me.ConversionEndTime.Name = "ConversionEndTime"
        '
        'ProcessID
        '
        Me.ProcessID.HeaderText = "Process ID"
        Me.ProcessID.Name = "ProcessID"
        Me.ProcessID.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.ProcessID.Width = 50
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
        'NumberSuccess
        '
        Me.NumberSuccess.HeaderText = "Success"
        Me.NumberSuccess.Name = "NumberSuccess"
        Me.NumberSuccess.ReadOnly = True
        Me.NumberSuccess.Width = 50
        '
        'NumberFailed
        '
        Me.NumberFailed.HeaderText = "Failed"
        Me.NumberFailed.Name = "NumberFailed"
        Me.NumberFailed.ReadOnly = True
        Me.NumberFailed.Width = 40
        '
        'SummaryReportLocation
        '
        Me.SummaryReportLocation.HeaderText = "Summary Report Location"
        Me.SummaryReportLocation.Name = "SummaryReportLocation"
        Me.SummaryReportLocation.ReadOnly = True
        Me.SummaryReportLocation.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'btnLaunchSourceConversionBatches
        '
        Me.btnLaunchSourceConversionBatches.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLaunchSourceConversionBatches.BackColor = System.Drawing.Color.White
        Me.btnLaunchSourceConversionBatches.Image = CType(resources.GetObject("btnLaunchSourceConversionBatches.Image"), System.Drawing.Image)
        Me.btnLaunchSourceConversionBatches.Location = New System.Drawing.Point(1147, 560)
        Me.btnLaunchSourceConversionBatches.Name = "btnLaunchSourceConversionBatches"
        Me.btnLaunchSourceConversionBatches.Size = New System.Drawing.Size(80, 55)
        Me.btnLaunchSourceConversionBatches.TabIndex = 79
        Me.btnLaunchSourceConversionBatches.UseVisualStyleBackColor = False
        '
        'btnShowSettings
        '
        Me.btnShowSettings.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnShowSettings.BackColor = System.Drawing.Color.White
        Me.btnShowSettings.Image = CType(resources.GetObject("btnShowSettings.Image"), System.Drawing.Image)
        Me.btnShowSettings.Location = New System.Drawing.Point(12, 560)
        Me.btnShowSettings.Name = "btnShowSettings"
        Me.btnShowSettings.Size = New System.Drawing.Size(80, 55)
        Me.btnShowSettings.TabIndex = 80
        Me.btnShowSettings.UseVisualStyleBackColor = False
        '
        'cboSourceDataType
        '
        Me.cboSourceDataType.FormattingEnabled = True
        Me.cboSourceDataType.Items.AddRange(New Object() {"nsf"})
        Me.cboSourceDataType.Location = New System.Drawing.Point(112, 25)
        Me.cboSourceDataType.Name = "cboSourceDataType"
        Me.cboSourceDataType.Size = New System.Drawing.Size(40, 21)
        Me.cboSourceDataType.TabIndex = 81
        '
        'lblSourceDataType
        '
        Me.lblSourceDataType.AutoSize = True
        Me.lblSourceDataType.Location = New System.Drawing.Point(10, 25)
        Me.lblSourceDataType.Name = "lblSourceDataType"
        Me.lblSourceDataType.Size = New System.Drawing.Size(97, 13)
        Me.lblSourceDataType.TabIndex = 82
        Me.lblSourceDataType.Text = "Source Data Type:"
        '
        'radSourceUser
        '
        Me.radSourceUser.AutoSize = True
        Me.radSourceUser.Location = New System.Drawing.Point(6, 15)
        Me.radSourceUser.Name = "radSourceUser"
        Me.radSourceUser.Size = New System.Drawing.Size(47, 17)
        Me.radSourceUser.TabIndex = 83
        Me.radSourceUser.TabStop = True
        Me.radSourceUser.Text = "User"
        Me.radSourceUser.UseVisualStyleBackColor = True
        '
        'grpSourceFormat
        '
        Me.grpSourceFormat.Controls.Add(Me.radSourceFlat)
        Me.grpSourceFormat.Controls.Add(Me.radSourceUser)
        Me.grpSourceFormat.Location = New System.Drawing.Point(615, 125)
        Me.grpSourceFormat.Name = "grpSourceFormat"
        Me.grpSourceFormat.Size = New System.Drawing.Size(110, 40)
        Me.grpSourceFormat.TabIndex = 84
        Me.grpSourceFormat.TabStop = False
        Me.grpSourceFormat.Text = "Source Format"
        '
        'radSourceFlat
        '
        Me.radSourceFlat.AutoSize = True
        Me.radSourceFlat.Location = New System.Drawing.Point(59, 15)
        Me.radSourceFlat.Name = "radSourceFlat"
        Me.radSourceFlat.Size = New System.Drawing.Size(42, 17)
        Me.radSourceFlat.TabIndex = 84
        Me.radSourceFlat.TabStop = True
        Me.radSourceFlat.Text = "Flat"
        Me.radSourceFlat.UseVisualStyleBackColor = True
        '
        'grpOutputFormat
        '
        Me.grpOutputFormat.Controls.Add(Me.radOutputFlat)
        Me.grpOutputFormat.Controls.Add(Me.radOutputUser)
        Me.grpOutputFormat.Location = New System.Drawing.Point(438, 10)
        Me.grpOutputFormat.Name = "grpOutputFormat"
        Me.grpOutputFormat.Size = New System.Drawing.Size(110, 40)
        Me.grpOutputFormat.TabIndex = 85
        Me.grpOutputFormat.TabStop = False
        Me.grpOutputFormat.Text = "Output Format"
        '
        'radOutputFlat
        '
        Me.radOutputFlat.AutoSize = True
        Me.radOutputFlat.Location = New System.Drawing.Point(55, 15)
        Me.radOutputFlat.Name = "radOutputFlat"
        Me.radOutputFlat.Size = New System.Drawing.Size(42, 17)
        Me.radOutputFlat.TabIndex = 1
        Me.radOutputFlat.TabStop = True
        Me.radOutputFlat.Text = "Flat"
        Me.radOutputFlat.UseVisualStyleBackColor = True
        '
        'radOutputUser
        '
        Me.radOutputUser.AutoSize = True
        Me.radOutputUser.Location = New System.Drawing.Point(6, 15)
        Me.radOutputUser.Name = "radOutputUser"
        Me.radOutputUser.Size = New System.Drawing.Size(47, 17)
        Me.radOutputUser.TabIndex = 0
        Me.radOutputUser.TabStop = True
        Me.radOutputUser.Text = "User"
        Me.radOutputUser.UseVisualStyleBackColor = True
        '
        'lblOutputDateType
        '
        Me.lblOutputDateType.AutoSize = True
        Me.lblOutputDateType.Location = New System.Drawing.Point(158, 25)
        Me.lblOutputDateType.Name = "lblOutputDateType"
        Me.lblOutputDateType.Size = New System.Drawing.Size(95, 13)
        Me.lblOutputDateType.TabIndex = 86
        Me.lblOutputDateType.Text = "Output Data Type:"
        '
        'cboOutputDataType
        '
        Me.cboOutputDataType.FormattingEnabled = True
        Me.cboOutputDataType.Items.AddRange(New Object() {"eml", "msg", "pst"})
        Me.cboOutputDataType.Location = New System.Drawing.Point(259, 25)
        Me.cboOutputDataType.Name = "cboOutputDataType"
        Me.cboOutputDataType.Size = New System.Drawing.Size(40, 21)
        Me.cboOutputDataType.TabIndex = 87
        '
        'btnLoadCustodianMappingData
        '
        Me.btnLoadCustodianMappingData.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLoadCustodianMappingData.Location = New System.Drawing.Point(1013, 56)
        Me.btnLoadCustodianMappingData.Name = "btnLoadCustodianMappingData"
        Me.btnLoadCustodianMappingData.Size = New System.Drawing.Size(121, 30)
        Me.btnLoadCustodianMappingData.TabIndex = 91
        Me.btnLoadCustodianMappingData.Text = "Load Mappping Data"
        Me.btnLoadCustodianMappingData.UseVisualStyleBackColor = True
        '
        'btnCustomerMappingFile
        '
        Me.btnCustomerMappingFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCustomerMappingFile.Location = New System.Drawing.Point(977, 59)
        Me.btnCustomerMappingFile.Name = "btnCustomerMappingFile"
        Me.btnCustomerMappingFile.Size = New System.Drawing.Size(30, 22)
        Me.btnCustomerMappingFile.TabIndex = 90
        Me.btnCustomerMappingFile.Text = "..."
        Me.btnCustomerMappingFile.UseVisualStyleBackColor = True
        '
        'txtCustodianMappingFile
        '
        Me.txtCustodianMappingFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCustodianMappingFile.Location = New System.Drawing.Point(731, 62)
        Me.txtCustodianMappingFile.Name = "txtCustodianMappingFile"
        Me.txtCustodianMappingFile.Size = New System.Drawing.Size(240, 20)
        Me.txtCustodianMappingFile.TabIndex = 89
        '
        'lblCustomerMappingFile
        '
        Me.lblCustomerMappingFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblCustomerMappingFile.AutoSize = True
        Me.lblCustomerMappingFile.Location = New System.Drawing.Point(655, 64)
        Me.lblCustomerMappingFile.Name = "lblCustomerMappingFile"
        Me.lblCustomerMappingFile.Size = New System.Drawing.Size(70, 13)
        Me.lblCustomerMappingFile.TabIndex = 88
        Me.lblCustomerMappingFile.Text = "Mapping File:"
        '
        'cboGroupIDs
        '
        Me.cboGroupIDs.FormattingEnabled = True
        Me.cboGroupIDs.Location = New System.Drawing.Point(196, 187)
        Me.cboGroupIDs.Name = "cboGroupIDs"
        Me.cboGroupIDs.Size = New System.Drawing.Size(94, 21)
        Me.cboGroupIDs.Sorted = True
        Me.cboGroupIDs.TabIndex = 94
        '
        'lblSelectByGroupID
        '
        Me.lblSelectByGroupID.AutoSize = True
        Me.lblSelectByGroupID.Location = New System.Drawing.Point(86, 190)
        Me.lblSelectByGroupID.Name = "lblSelectByGroupID"
        Me.lblSelectByGroupID.Size = New System.Drawing.Size(104, 13)
        Me.lblSelectByGroupID.TabIndex = 93
        Me.lblSelectByGroupID.Text = "Select By Group ID: "
        '
        'btnSelectAll
        '
        Me.btnSelectAll.Location = New System.Drawing.Point(12, 181)
        Me.btnSelectAll.Name = "btnSelectAll"
        Me.btnSelectAll.Size = New System.Drawing.Size(59, 30)
        Me.btnSelectAll.TabIndex = 92
        Me.btnSelectAll.Text = "Select All"
        Me.btnSelectAll.UseVisualStyleBackColor = True
        '
        'btnClearDates
        '
        Me.btnClearDates.Location = New System.Drawing.Point(460, 91)
        Me.btnClearDates.Name = "btnClearDates"
        Me.btnClearDates.Size = New System.Drawing.Size(75, 23)
        Me.btnClearDates.TabIndex = 99
        Me.btnClearDates.Text = "Clear"
        Me.btnClearDates.UseVisualStyleBackColor = True
        '
        'dateToDate
        '
        Me.dateToDate.Location = New System.Drawing.Point(348, 65)
        Me.dateToDate.Name = "dateToDate"
        Me.dateToDate.Size = New System.Drawing.Size(187, 20)
        Me.dateToDate.TabIndex = 98
        '
        'lblToDate
        '
        Me.lblToDate.AutoSize = True
        Me.lblToDate.Location = New System.Drawing.Point(287, 65)
        Me.lblToDate.Name = "lblToDate"
        Me.lblToDate.Size = New System.Drawing.Size(49, 13)
        Me.lblToDate.TabIndex = 97
        Me.lblToDate.Text = "To Date:"
        '
        'lblFromDate
        '
        Me.lblFromDate.AutoSize = True
        Me.lblFromDate.Location = New System.Drawing.Point(12, 65)
        Me.lblFromDate.Name = "lblFromDate"
        Me.lblFromDate.Size = New System.Drawing.Size(59, 13)
        Me.lblFromDate.TabIndex = 96
        Me.lblFromDate.Text = "From Date:"
        '
        'dateFromDate
        '
        Me.dateFromDate.Location = New System.Drawing.Point(89, 65)
        Me.dateFromDate.Name = "dateFromDate"
        Me.dateFromDate.Size = New System.Drawing.Size(192, 20)
        Me.dateFromDate.TabIndex = 95
        '
        'chkDedupData
        '
        Me.chkDedupData.AutoSize = True
        Me.chkDedupData.Location = New System.Drawing.Point(319, 25)
        Me.chkDedupData.Name = "chkDedupData"
        Me.chkDedupData.Size = New System.Drawing.Size(200, 17)
        Me.chkDedupData.TabIndex = 100
        Me.chkDedupData.Text = "Perform Top-level Item Deduplication"
        Me.chkDedupData.UseVisualStyleBackColor = True
        '
        'FileSystemWatcher1
        '
        Me.FileSystemWatcher1.EnableRaisingEvents = True
        Me.FileSystemWatcher1.Filter = "summary-report.txt"
        Me.FileSystemWatcher1.SynchronizingObject = Me
        '
        'NuixLogoPicture
        '
        Me.NuixLogoPicture.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.NuixLogoPicture.Image = CType(resources.GetObject("NuixLogoPicture.Image"), System.Drawing.Image)
        Me.NuixLogoPicture.Location = New System.Drawing.Point(1152, 20)
        Me.NuixLogoPicture.Name = "NuixLogoPicture"
        Me.NuixLogoPicture.Size = New System.Drawing.Size(75, 100)
        Me.NuixLogoPicture.TabIndex = 101
        Me.NuixLogoPicture.TabStop = False
        '
        'EmailConversionToolTip
        '
        Me.EmailConversionToolTip.IsBalloon = True
        '
        'btnReloadEmailConversionData
        '
        Me.btnReloadEmailConversionData.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnReloadEmailConversionData.BackColor = System.Drawing.Color.White
        Me.btnReloadEmailConversionData.Image = CType(resources.GetObject("btnReloadEmailConversionData.Image"), System.Drawing.Image)
        Me.btnReloadEmailConversionData.Location = New System.Drawing.Point(98, 560)
        Me.btnReloadEmailConversionData.Name = "btnReloadEmailConversionData"
        Me.btnReloadEmailConversionData.Size = New System.Drawing.Size(80, 55)
        Me.btnReloadEmailConversionData.TabIndex = 102
        Me.btnReloadEmailConversionData.UseVisualStyleBackColor = False
        '
        'btnExportProcessingDetails
        '
        Me.btnExportProcessingDetails.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportProcessingDetails.BackColor = System.Drawing.Color.White
        Me.btnExportProcessingDetails.Image = CType(resources.GetObject("btnExportProcessingDetails.Image"), System.Drawing.Image)
        Me.btnExportProcessingDetails.Location = New System.Drawing.Point(1061, 560)
        Me.btnExportProcessingDetails.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExportProcessingDetails.Name = "btnExportProcessingDetails"
        Me.btnExportProcessingDetails.Size = New System.Drawing.Size(80, 55)
        Me.btnExportProcessingDetails.TabIndex = 103
        Me.btnExportProcessingDetails.UseVisualStyleBackColor = False
        '
        'btnAddDateRangeFilter
        '
        Me.btnAddDateRangeFilter.Location = New System.Drawing.Point(554, 63)
        Me.btnAddDateRangeFilter.Name = "btnAddDateRangeFilter"
        Me.btnAddDateRangeFilter.Size = New System.Drawing.Size(28, 23)
        Me.btnAddDateRangeFilter.TabIndex = 104
        Me.btnAddDateRangeFilter.Text = "+"
        Me.btnAddDateRangeFilter.UseVisualStyleBackColor = True
        '
        'lblFilters
        '
        Me.lblFilters.AutoSize = True
        Me.lblFilters.Location = New System.Drawing.Point(12, 89)
        Me.lblFilters.Name = "lblFilters"
        Me.lblFilters.Size = New System.Drawing.Size(73, 13)
        Me.lblFilters.TabIndex = 106
        Me.lblFilters.Text = "Date Ranges:"
        '
        'btnAddToSelected
        '
        Me.btnAddToSelected.Location = New System.Drawing.Point(348, 91)
        Me.btnAddToSelected.Name = "btnAddToSelected"
        Me.btnAddToSelected.Size = New System.Drawing.Size(103, 23)
        Me.btnAddToSelected.TabIndex = 107
        Me.btnAddToSelected.Text = "Add To Selected"
        Me.btnAddToSelected.UseVisualStyleBackColor = True
        '
        'lstDateRangeFilters
        '
        Me.lstDateRangeFilters.FormattingEnabled = True
        Me.lstDateRangeFilters.Location = New System.Drawing.Point(89, 89)
        Me.lstDateRangeFilters.Name = "lstDateRangeFilters"
        Me.lstDateRangeFilters.Size = New System.Drawing.Size(192, 82)
        Me.lstDateRangeFilters.TabIndex = 108
        '
        'ConsolidateExporterFiles
        '
        Me.ConsolidateExporterFiles.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ConsolidateExporterMetricsFilesToolStripMenuItem, Me.ConsolidateExporterErrorsFilesToolStripMenuItem})
        Me.ConsolidateExporterFiles.Name = "ConsolidateExporterFiles"
        Me.ConsolidateExporterFiles.Size = New System.Drawing.Size(254, 48)
        '
        'ConsolidateExporterMetricsFilesToolStripMenuItem
        '
        Me.ConsolidateExporterMetricsFilesToolStripMenuItem.Name = "ConsolidateExporterMetricsFilesToolStripMenuItem"
        Me.ConsolidateExporterMetricsFilesToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.ConsolidateExporterMetricsFilesToolStripMenuItem.Text = "Consolidate Exporter-Metrics Files"
        '
        'ConsolidateExporterErrorsFilesToolStripMenuItem
        '
        Me.ConsolidateExporterErrorsFilesToolStripMenuItem.Name = "ConsolidateExporterErrorsFilesToolStripMenuItem"
        Me.ConsolidateExporterErrorsFilesToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.ConsolidateExporterErrorsFilesToolStripMenuItem.Text = "Consolidate Exporter-Errors Files"
        '
        'btnConsolidateExporterFiles
        '
        Me.btnConsolidateExporterFiles.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnConsolidateExporterFiles.Image = CType(resources.GetObject("btnConsolidateExporterFiles.Image"), System.Drawing.Image)
        Me.btnConsolidateExporterFiles.Location = New System.Drawing.Point(974, 560)
        Me.btnConsolidateExporterFiles.Name = "btnConsolidateExporterFiles"
        Me.btnConsolidateExporterFiles.Size = New System.Drawing.Size(80, 55)
        Me.btnConsolidateExporterFiles.TabIndex = 110
        Me.btnConsolidateExporterFiles.UseVisualStyleBackColor = True
        '
        'NSFConversion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1239, 621)
        Me.Controls.Add(Me.btnConsolidateExporterFiles)
        Me.Controls.Add(Me.lstDateRangeFilters)
        Me.Controls.Add(Me.btnAddToSelected)
        Me.Controls.Add(Me.lblFilters)
        Me.Controls.Add(Me.btnAddDateRangeFilter)
        Me.Controls.Add(Me.btnExportProcessingDetails)
        Me.Controls.Add(Me.btnReloadEmailConversionData)
        Me.Controls.Add(Me.NuixLogoPicture)
        Me.Controls.Add(Me.chkDedupData)
        Me.Controls.Add(Me.btnClearDates)
        Me.Controls.Add(Me.dateToDate)
        Me.Controls.Add(Me.lblToDate)
        Me.Controls.Add(Me.lblFromDate)
        Me.Controls.Add(Me.dateFromDate)
        Me.Controls.Add(Me.cboGroupIDs)
        Me.Controls.Add(Me.lblSelectByGroupID)
        Me.Controls.Add(Me.btnSelectAll)
        Me.Controls.Add(Me.btnLoadCustodianMappingData)
        Me.Controls.Add(Me.btnCustomerMappingFile)
        Me.Controls.Add(Me.txtCustodianMappingFile)
        Me.Controls.Add(Me.lblCustomerMappingFile)
        Me.Controls.Add(Me.cboOutputDataType)
        Me.Controls.Add(Me.lblOutputDateType)
        Me.Controls.Add(Me.grpOutputFormat)
        Me.Controls.Add(Me.grpSourceFormat)
        Me.Controls.Add(Me.lblSourceDataType)
        Me.Controls.Add(Me.cboSourceDataType)
        Me.Controls.Add(Me.btnShowSettings)
        Me.Controls.Add(Me.btnLaunchSourceConversionBatches)
        Me.Controls.Add(Me.grdConversionDataInfo)
        Me.Controls.Add(Me.txtNSFLocation)
        Me.Controls.Add(Me.btnLoadPSTInfo)
        Me.Controls.Add(Me.btnFileSystemChooser)
        Me.Controls.Add(Me.lblPSTLocation)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(1255, 660)
        Me.Name = "NSFConversion"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Legacy Email Conversion"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.grdConversionDataInfo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpSourceFormat.ResumeLayout(False)
        Me.grpSourceFormat.PerformLayout()
        Me.grpOutputFormat.ResumeLayout(False)
        Me.grpOutputFormat.PerformLayout()
        CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NuixLogoPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ConsolidateExporterFiles.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtNSFLocation As System.Windows.Forms.TextBox
    Friend WithEvents btnLoadPSTInfo As System.Windows.Forms.Button
    Friend WithEvents btnFileSystemChooser As System.Windows.Forms.Button
    Friend WithEvents lblPSTLocation As System.Windows.Forms.Label
    Friend WithEvents grdConversionDataInfo As System.Windows.Forms.DataGridView
    Friend WithEvents btnLaunchSourceConversionBatches As System.Windows.Forms.Button
    Friend WithEvents btnShowSettings As System.Windows.Forms.Button
    Friend WithEvents cboSourceDataType As System.Windows.Forms.ComboBox
    Friend WithEvents lblSourceDataType As System.Windows.Forms.Label
    Friend WithEvents radSourceUser As System.Windows.Forms.RadioButton
    Friend WithEvents grpSourceFormat As System.Windows.Forms.GroupBox
    Friend WithEvents radSourceFlat As System.Windows.Forms.RadioButton
    Friend WithEvents grpOutputFormat As System.Windows.Forms.GroupBox
    Friend WithEvents radOutputFlat As System.Windows.Forms.RadioButton
    Friend WithEvents radOutputUser As System.Windows.Forms.RadioButton
    Friend WithEvents lblOutputDateType As System.Windows.Forms.Label
    Friend WithEvents cboOutputDataType As System.Windows.Forms.ComboBox
    Friend WithEvents btnLoadCustodianMappingData As System.Windows.Forms.Button
    Friend WithEvents btnCustomerMappingFile As System.Windows.Forms.Button
    Friend WithEvents txtCustodianMappingFile As System.Windows.Forms.TextBox
    Friend WithEvents lblCustomerMappingFile As System.Windows.Forms.Label
    Friend WithEvents cboGroupIDs As System.Windows.Forms.ComboBox
    Friend WithEvents lblSelectByGroupID As System.Windows.Forms.Label
    Friend WithEvents btnSelectAll As System.Windows.Forms.Button
    Friend WithEvents btnClearDates As System.Windows.Forms.Button
    Friend WithEvents dateToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblToDate As System.Windows.Forms.Label
    Friend WithEvents lblFromDate As System.Windows.Forms.Label
    Friend WithEvents dateFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkDedupData As System.Windows.Forms.CheckBox
    Friend WithEvents FileSystemWatcher1 As System.IO.FileSystemWatcher
    Friend WithEvents NuixLogoPicture As System.Windows.Forms.PictureBox
    Friend WithEvents EmailConversionToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents btnReloadEmailConversionData As System.Windows.Forms.Button
    Friend WithEvents btnExportProcessingDetails As System.Windows.Forms.Button
    Friend WithEvents StatusImage As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents SelectCustodian As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CustodianName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateRangeFilters As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PercentComplete As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BytesProcessed As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ConversionStatus As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents GroupID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NumberOfSourceFiles As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SizeOfSourceFiles As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SourcePath As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SourceDataType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SourceFormat As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OutputDataType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OutputFormat As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RedisSettings As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ConversionStartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ConversionEndTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProcessID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProcessingFilesDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CaseDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OutputDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LogDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NumberSuccess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NumberFailed As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SummaryReportLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnAddDateRangeFilter As System.Windows.Forms.Button
    Friend WithEvents lblFilters As System.Windows.Forms.Label
    Friend WithEvents btnAddToSelected As System.Windows.Forms.Button
    Friend WithEvents lstDateRangeFilters As System.Windows.Forms.ListBox
    Friend WithEvents btnConsolidateExporterFiles As System.Windows.Forms.Button
    Friend WithEvents ConsolidateExporterFiles As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ConsolidateExporterMetricsFilesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ConsolidateExporterErrorsFilesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
