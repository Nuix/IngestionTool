<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNuixIngestion
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNuixIngestion))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.btnNuixSystemSettings = New System.Windows.Forms.Button()
        Me.grdPSTInfo = New System.Windows.Forms.DataGridView()
        Me.StatusImage = New System.Windows.Forms.DataGridViewImageColumn()
        Me.SelectCustodian = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.CustodianName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PercentageCompleted = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BytesUploaded = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProgressStatus = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.NumberOfPSTs = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SizeOfPSTs = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PSTPath = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DestinationFolder = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DestinationRoot = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DestinationSMTP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProcessID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NumberSuccess = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NumberFailed = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProcessingFilesDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CaseDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OutputDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ReprocessingDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LogDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IngestionStartTime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IngestionEndTime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SummaryReportLocation = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txtPSTLocation = New System.Windows.Forms.TextBox()
        Me.lblPSTLocation = New System.Windows.Forms.Label()
        Me.btnFileSystemChooser = New System.Windows.Forms.Button()
        Me.btnLoadPSTInfo = New System.Windows.Forms.Button()
        Me.FileSystemWatcher1 = New System.IO.FileSystemWatcher()
        Me.btnProcessandRunSelectedUsers = New System.Windows.Forms.Button()
        Me.btnSelectAll = New System.Windows.Forms.Button()
        Me.lblSelectByGroupID = New System.Windows.Forms.Label()
        Me.cboGroupIDs = New System.Windows.Forms.ComboBox()
        Me.btnLoadPreviousConfig = New System.Windows.Forms.Button()
        Me.btnClearGrid = New System.Windows.Forms.Button()
        Me.btnRefreshGridData = New System.Windows.Forms.Button()
        Me.btnCustomerMappingFile = New System.Windows.Forms.Button()
        Me.txtCustodianMappingFile = New System.Windows.Forms.TextBox()
        Me.lblCustomerMappingFile = New System.Windows.Forms.Label()
        Me.btnLoadCustodianMappingData = New System.Windows.Forms.Button()
        Me.btnProcessingDetails = New System.Windows.Forms.Button()
        Me.txtFilter = New System.Windows.Forms.TextBox()
        Me.DataGridViewImageColumn1 = New System.Windows.Forms.DataGridViewImageColumn()
        Me.DataGridViewImageColumn2 = New System.Windows.Forms.DataGridViewImageColumn()
        Me.CaseLockWatcher = New System.IO.FileSystemWatcher()
        Me.NuixLogoPicture = New System.Windows.Forms.PictureBox()
        Me.EWSIngestToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnExportProcessingDetails = New System.Windows.Forms.Button()
        Me.lblFilter = New System.Windows.Forms.Label()
        Me.btnExportNotUploaded = New System.Windows.Forms.Button()
        Me.SelectAllItemsNotProcessed = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExportItemsNotUploadedToPST = New System.Windows.Forms.ToolStripMenuItem()
        Me.ItemsNotUploadedMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ExportFTSDataTooLargeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExportOtherNotUploadedToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.btnConsolidateExporterFiles = New System.Windows.Forms.Button()
        Me.ConsolidateExporterFiles = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ConsolidateExporterMetricsFilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ConsolidateExporterErrorsFilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.grdPSTInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CaseLockWatcher, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NuixLogoPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ItemsNotUploadedMenuStrip.SuspendLayout()
        Me.ConsolidateExporterFiles.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnNuixSystemSettings
        '
        Me.btnNuixSystemSettings.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnNuixSystemSettings.BackColor = System.Drawing.Color.White
        Me.btnNuixSystemSettings.FlatAppearance.BorderSize = 3
        Me.btnNuixSystemSettings.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnNuixSystemSettings.Image = CType(resources.GetObject("btnNuixSystemSettings.Image"), System.Drawing.Image)
        Me.btnNuixSystemSettings.Location = New System.Drawing.Point(20, 617)
        Me.btnNuixSystemSettings.Name = "btnNuixSystemSettings"
        Me.btnNuixSystemSettings.Size = New System.Drawing.Size(80, 55)
        Me.btnNuixSystemSettings.TabIndex = 8
        Me.btnNuixSystemSettings.UseVisualStyleBackColor = False
        '
        'grdPSTInfo
        '
        Me.grdPSTInfo.AllowUserToAddRows = False
        Me.grdPSTInfo.AllowUserToDeleteRows = False
        Me.grdPSTInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grdPSTInfo.BackgroundColor = System.Drawing.SystemColors.Control
        Me.grdPSTInfo.ColumnHeadersHeight = 60
        Me.grdPSTInfo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.StatusImage, Me.SelectCustodian, Me.CustodianName, Me.PercentageCompleted, Me.BytesUploaded, Me.ProgressStatus, Me.NumberOfPSTs, Me.SizeOfPSTs, Me.PSTPath, Me.DestinationFolder, Me.DestinationRoot, Me.DestinationSMTP, Me.GroupID, Me.ProcessID, Me.NumberSuccess, Me.NumberFailed, Me.ProcessingFilesDirectory, Me.CaseDirectory, Me.OutputDirectory, Me.ReprocessingDirectory, Me.LogDirectory, Me.IngestionStartTime, Me.IngestionEndTime, Me.SummaryReportLocation})
        Me.grdPSTInfo.GridColor = System.Drawing.SystemColors.ControlLight
        Me.grdPSTInfo.Location = New System.Drawing.Point(20, 113)
        Me.grdPSTInfo.Name = "grdPSTInfo"
        Me.grdPSTInfo.RowTemplate.Height = 28
        Me.grdPSTInfo.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.grdPSTInfo.Size = New System.Drawing.Size(1407, 498)
        Me.grdPSTInfo.TabIndex = 1
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
        'PercentageCompleted
        '
        Me.PercentageCompleted.Frozen = True
        Me.PercentageCompleted.HeaderText = "Percent Completed"
        Me.PercentageCompleted.Name = "PercentageCompleted"
        Me.PercentageCompleted.Width = 60
        '
        'BytesUploaded
        '
        Me.BytesUploaded.Frozen = True
        Me.BytesUploaded.HeaderText = "Bytes Uploaded"
        Me.BytesUploaded.Name = "BytesUploaded"
        Me.BytesUploaded.ReadOnly = True
        Me.BytesUploaded.Width = 75
        '
        'ProgressStatus
        '
        Me.ProgressStatus.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ProgressStatus.Frozen = True
        Me.ProgressStatus.HeaderText = "Progress Status"
        Me.ProgressStatus.Items.AddRange(New Object() {"Not Started", "Not Started Items Not Uploaded", "In Progress", "In Progress Items Not Uploaded", "Completed", "Completed Processing Items Not Uploaded", "Completed with Items Not Uploaded", "Failed", "Failed (Case Already Exists)", "Failed (Export Already Exists)", "Failed (Case and Export Already Exists)", "Process Terminated by User", "Staged", "Restart Ingestion"})
        Me.ProgressStatus.Name = "ProgressStatus"
        Me.ProgressStatus.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ProgressStatus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ProgressStatus.Width = 97
        '
        'NumberOfPSTs
        '
        Me.NumberOfPSTs.Frozen = True
        Me.NumberOfPSTs.HeaderText = "Number Of PSTs"
        Me.NumberOfPSTs.Name = "NumberOfPSTs"
        Me.NumberOfPSTs.ReadOnly = True
        Me.NumberOfPSTs.Width = 50
        '
        'SizeOfPSTs
        '
        Me.SizeOfPSTs.Frozen = True
        Me.SizeOfPSTs.HeaderText = "Total Size Of PSTs"
        Me.SizeOfPSTs.Name = "SizeOfPSTs"
        Me.SizeOfPSTs.ReadOnly = True
        Me.SizeOfPSTs.Width = 75
        '
        'PSTPath
        '
        Me.PSTPath.Frozen = True
        Me.PSTPath.HeaderText = "PST Path"
        Me.PSTPath.Name = "PSTPath"
        '
        'DestinationFolder
        '
        Me.DestinationFolder.HeaderText = "Destination Folder"
        Me.DestinationFolder.Name = "DestinationFolder"
        Me.DestinationFolder.Width = 75
        '
        'DestinationRoot
        '
        Me.DestinationRoot.HeaderText = "Destination Root"
        Me.DestinationRoot.Name = "DestinationRoot"
        Me.DestinationRoot.Width = 65
        '
        'DestinationSMTP
        '
        Me.DestinationSMTP.HeaderText = "Destination SMTP"
        Me.DestinationSMTP.Name = "DestinationSMTP"
        '
        'GroupID
        '
        Me.GroupID.HeaderText = "Group ID"
        Me.GroupID.Name = "GroupID"
        Me.GroupID.Width = 50
        '
        'ProcessID
        '
        Me.ProcessID.HeaderText = "Process ID"
        Me.ProcessID.Name = "ProcessID"
        Me.ProcessID.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.ProcessID.Width = 50
        '
        'NumberSuccess
        '
        Me.NumberSuccess.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.NumberSuccess.HeaderText = "Items Uploaded"
        Me.NumberSuccess.Name = "NumberSuccess"
        Me.NumberSuccess.ReadOnly = True
        Me.NumberSuccess.Width = 97
        '
        'NumberFailed
        '
        Me.NumberFailed.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.NumberFailed.HeaderText = "Items Not Uploaded"
        Me.NumberFailed.Name = "NumberFailed"
        Me.NumberFailed.ReadOnly = True
        Me.NumberFailed.Width = 115
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
        'ReprocessingDirectory
        '
        Me.ReprocessingDirectory.HeaderText = "Reprocessing Directory"
        Me.ReprocessingDirectory.Name = "ReprocessingDirectory"
        '
        'LogDirectory
        '
        Me.LogDirectory.HeaderText = "Log Directory"
        Me.LogDirectory.Name = "LogDirectory"
        '
        'IngestionStartTime
        '
        Me.IngestionStartTime.HeaderText = "Ingestion Start Time"
        Me.IngestionStartTime.Name = "IngestionStartTime"
        '
        'IngestionEndTime
        '
        Me.IngestionEndTime.HeaderText = "Ingestion End Time"
        Me.IngestionEndTime.Name = "IngestionEndTime"
        '
        'SummaryReportLocation
        '
        Me.SummaryReportLocation.HeaderText = "Summary Report Location"
        Me.SummaryReportLocation.Name = "SummaryReportLocation"
        Me.SummaryReportLocation.ReadOnly = True
        Me.SummaryReportLocation.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'txtPSTLocation
        '
        Me.txtPSTLocation.Location = New System.Drawing.Point(155, 30)
        Me.txtPSTLocation.Name = "txtPSTLocation"
        Me.txtPSTLocation.Size = New System.Drawing.Size(370, 20)
        Me.txtPSTLocation.TabIndex = 5
        '
        'lblPSTLocation
        '
        Me.lblPSTLocation.AutoSize = True
        Me.lblPSTLocation.Location = New System.Drawing.Point(25, 30)
        Me.lblPSTLocation.Name = "lblPSTLocation"
        Me.lblPSTLocation.Size = New System.Drawing.Size(125, 13)
        Me.lblPSTLocation.TabIndex = 4
        Me.lblPSTLocation.Text = "Custodian PST Location:"
        '
        'btnFileSystemChooser
        '
        Me.btnFileSystemChooser.Location = New System.Drawing.Point(535, 30)
        Me.btnFileSystemChooser.Name = "btnFileSystemChooser"
        Me.btnFileSystemChooser.Size = New System.Drawing.Size(30, 22)
        Me.btnFileSystemChooser.TabIndex = 5
        Me.btnFileSystemChooser.Text = "..."
        Me.btnFileSystemChooser.UseVisualStyleBackColor = True
        '
        'btnLoadPSTInfo
        '
        Me.btnLoadPSTInfo.Location = New System.Drawing.Point(575, 30)
        Me.btnLoadPSTInfo.Name = "btnLoadPSTInfo"
        Me.btnLoadPSTInfo.Size = New System.Drawing.Size(100, 31)
        Me.btnLoadPSTInfo.TabIndex = 7
        Me.btnLoadPSTInfo.Text = "Load PST Info"
        Me.btnLoadPSTInfo.UseVisualStyleBackColor = True
        '
        'FileSystemWatcher1
        '
        Me.FileSystemWatcher1.EnableRaisingEvents = True
        Me.FileSystemWatcher1.Filter = "summary-report.txt"
        Me.FileSystemWatcher1.IncludeSubdirectories = True
        Me.FileSystemWatcher1.SynchronizingObject = Me
        '
        'btnProcessandRunSelectedUsers
        '
        Me.btnProcessandRunSelectedUsers.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnProcessandRunSelectedUsers.BackColor = System.Drawing.Color.White
        Me.btnProcessandRunSelectedUsers.Image = CType(resources.GetObject("btnProcessandRunSelectedUsers.Image"), System.Drawing.Image)
        Me.btnProcessandRunSelectedUsers.Location = New System.Drawing.Point(1347, 618)
        Me.btnProcessandRunSelectedUsers.Name = "btnProcessandRunSelectedUsers"
        Me.btnProcessandRunSelectedUsers.Size = New System.Drawing.Size(80, 55)
        Me.btnProcessandRunSelectedUsers.TabIndex = 32
        Me.btnProcessandRunSelectedUsers.UseVisualStyleBackColor = False
        '
        'btnSelectAll
        '
        Me.btnSelectAll.Location = New System.Drawing.Point(20, 77)
        Me.btnSelectAll.Name = "btnSelectAll"
        Me.btnSelectAll.Size = New System.Drawing.Size(63, 30)
        Me.btnSelectAll.TabIndex = 58
        Me.btnSelectAll.Text = "Select All"
        Me.btnSelectAll.UseVisualStyleBackColor = True
        '
        'lblSelectByGroupID
        '
        Me.lblSelectByGroupID.AutoSize = True
        Me.lblSelectByGroupID.Location = New System.Drawing.Point(89, 82)
        Me.lblSelectByGroupID.Name = "lblSelectByGroupID"
        Me.lblSelectByGroupID.Size = New System.Drawing.Size(104, 13)
        Me.lblSelectByGroupID.TabIndex = 59
        Me.lblSelectByGroupID.Text = "Select By Group ID: "
        '
        'cboGroupIDs
        '
        Me.cboGroupIDs.FormattingEnabled = True
        Me.cboGroupIDs.Location = New System.Drawing.Point(192, 79)
        Me.cboGroupIDs.Name = "cboGroupIDs"
        Me.cboGroupIDs.Size = New System.Drawing.Size(76, 21)
        Me.cboGroupIDs.Sorted = True
        Me.cboGroupIDs.TabIndex = 60
        '
        'btnLoadPreviousConfig
        '
        Me.btnLoadPreviousConfig.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnLoadPreviousConfig.BackColor = System.Drawing.Color.White
        Me.btnLoadPreviousConfig.Image = CType(resources.GetObject("btnLoadPreviousConfig.Image"), System.Drawing.Image)
        Me.btnLoadPreviousConfig.Location = New System.Drawing.Point(106, 617)
        Me.btnLoadPreviousConfig.Name = "btnLoadPreviousConfig"
        Me.btnLoadPreviousConfig.Size = New System.Drawing.Size(80, 55)
        Me.btnLoadPreviousConfig.TabIndex = 62
        Me.btnLoadPreviousConfig.UseVisualStyleBackColor = False
        '
        'btnClearGrid
        '
        Me.btnClearGrid.Location = New System.Drawing.Point(538, 76)
        Me.btnClearGrid.Name = "btnClearGrid"
        Me.btnClearGrid.Size = New System.Drawing.Size(94, 33)
        Me.btnClearGrid.TabIndex = 65
        Me.btnClearGrid.Text = "Clear Grid"
        Me.btnClearGrid.UseVisualStyleBackColor = True
        '
        'btnRefreshGridData
        '
        Me.btnRefreshGridData.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnRefreshGridData.BackColor = System.Drawing.Color.White
        Me.btnRefreshGridData.Image = CType(resources.GetObject("btnRefreshGridData.Image"), System.Drawing.Image)
        Me.btnRefreshGridData.Location = New System.Drawing.Point(192, 617)
        Me.btnRefreshGridData.Name = "btnRefreshGridData"
        Me.btnRefreshGridData.Size = New System.Drawing.Size(80, 55)
        Me.btnRefreshGridData.TabIndex = 66
        Me.btnRefreshGridData.UseVisualStyleBackColor = False
        '
        'btnCustomerMappingFile
        '
        Me.btnCustomerMappingFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCustomerMappingFile.Location = New System.Drawing.Point(1153, 30)
        Me.btnCustomerMappingFile.Name = "btnCustomerMappingFile"
        Me.btnCustomerMappingFile.Size = New System.Drawing.Size(30, 22)
        Me.btnCustomerMappingFile.TabIndex = 69
        Me.btnCustomerMappingFile.Text = "..."
        Me.btnCustomerMappingFile.UseVisualStyleBackColor = True
        '
        'txtCustodianMappingFile
        '
        Me.txtCustodianMappingFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCustodianMappingFile.Location = New System.Drawing.Point(776, 30)
        Me.txtCustodianMappingFile.Name = "txtCustodianMappingFile"
        Me.txtCustodianMappingFile.Size = New System.Drawing.Size(370, 20)
        Me.txtCustodianMappingFile.TabIndex = 68
        '
        'lblCustomerMappingFile
        '
        Me.lblCustomerMappingFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblCustomerMappingFile.AutoSize = True
        Me.lblCustomerMappingFile.Location = New System.Drawing.Point(700, 30)
        Me.lblCustomerMappingFile.Name = "lblCustomerMappingFile"
        Me.lblCustomerMappingFile.Size = New System.Drawing.Size(70, 13)
        Me.lblCustomerMappingFile.TabIndex = 67
        Me.lblCustomerMappingFile.Text = "Mapping File:"
        '
        'btnLoadCustodianMappingData
        '
        Me.btnLoadCustodianMappingData.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLoadCustodianMappingData.Location = New System.Drawing.Point(1190, 28)
        Me.btnLoadCustodianMappingData.Name = "btnLoadCustodianMappingData"
        Me.btnLoadCustodianMappingData.Size = New System.Drawing.Size(151, 31)
        Me.btnLoadCustodianMappingData.TabIndex = 70
        Me.btnLoadCustodianMappingData.Text = "Load Custodian Mappping"
        Me.btnLoadCustodianMappingData.UseVisualStyleBackColor = True
        '
        'btnProcessingDetails
        '
        Me.btnProcessingDetails.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnProcessingDetails.BackColor = System.Drawing.Color.White
        Me.btnProcessingDetails.Image = CType(resources.GetObject("btnProcessingDetails.Image"), System.Drawing.Image)
        Me.btnProcessingDetails.Location = New System.Drawing.Point(1091, 618)
        Me.btnProcessingDetails.Name = "btnProcessingDetails"
        Me.btnProcessingDetails.Size = New System.Drawing.Size(80, 55)
        Me.btnProcessingDetails.TabIndex = 75
        Me.btnProcessingDetails.UseVisualStyleBackColor = False
        '
        'txtFilter
        '
        Me.txtFilter.Location = New System.Drawing.Point(353, 79)
        Me.txtFilter.Name = "txtFilter"
        Me.txtFilter.Size = New System.Drawing.Size(179, 20)
        Me.txtFilter.TabIndex = 76
        '
        'DataGridViewImageColumn1
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.NullValue = "Nothing"
        Me.DataGridViewImageColumn1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridViewImageColumn1.Frozen = True
        Me.DataGridViewImageColumn1.HeaderText = ""
        Me.DataGridViewImageColumn1.Name = "DataGridViewImageColumn1"
        '
        'DataGridViewImageColumn2
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.DataGridViewImageColumn2.DefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridViewImageColumn2.Frozen = True
        Me.DataGridViewImageColumn2.HeaderText = ""
        Me.DataGridViewImageColumn2.Image = Global.NEAMM.My.Resources.Resources.not_selected_small
        Me.DataGridViewImageColumn2.Name = "DataGridViewImageColumn2"
        Me.DataGridViewImageColumn2.ReadOnly = True
        Me.DataGridViewImageColumn2.Width = 50
        '
        'CaseLockWatcher
        '
        Me.CaseLockWatcher.EnableRaisingEvents = True
        Me.CaseLockWatcher.Filter = "case.lock"
        Me.CaseLockWatcher.IncludeSubdirectories = True
        Me.CaseLockWatcher.SynchronizingObject = Me
        '
        'NuixLogoPicture
        '
        Me.NuixLogoPicture.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.NuixLogoPicture.Image = CType(resources.GetObject("NuixLogoPicture.Image"), System.Drawing.Image)
        Me.NuixLogoPicture.Location = New System.Drawing.Point(1352, 7)
        Me.NuixLogoPicture.Name = "NuixLogoPicture"
        Me.NuixLogoPicture.Size = New System.Drawing.Size(75, 100)
        Me.NuixLogoPicture.TabIndex = 77
        Me.NuixLogoPicture.TabStop = False
        '
        'EWSIngestToolTip
        '
        Me.EWSIngestToolTip.IsBalloon = True
        '
        'btnExportProcessingDetails
        '
        Me.btnExportProcessingDetails.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportProcessingDetails.BackColor = System.Drawing.Color.White
        Me.btnExportProcessingDetails.Image = CType(resources.GetObject("btnExportProcessingDetails.Image"), System.Drawing.Image)
        Me.btnExportProcessingDetails.Location = New System.Drawing.Point(1176, 618)
        Me.btnExportProcessingDetails.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExportProcessingDetails.Name = "btnExportProcessingDetails"
        Me.btnExportProcessingDetails.Size = New System.Drawing.Size(80, 55)
        Me.btnExportProcessingDetails.TabIndex = 78
        Me.btnExportProcessingDetails.UseVisualStyleBackColor = False
        '
        'lblFilter
        '
        Me.lblFilter.AutoSize = True
        Me.lblFilter.Location = New System.Drawing.Point(284, 82)
        Me.lblFilter.Name = "lblFilter"
        Me.lblFilter.Size = New System.Drawing.Size(63, 13)
        Me.lblFilter.TabIndex = 79
        Me.lblFilter.Text = "Quick Filter:"
        '
        'btnExportNotUploaded
        '
        Me.btnExportNotUploaded.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportNotUploaded.BackColor = System.Drawing.Color.White
        Me.btnExportNotUploaded.Image = CType(resources.GetObject("btnExportNotUploaded.Image"), System.Drawing.Image)
        Me.btnExportNotUploaded.Location = New System.Drawing.Point(1261, 618)
        Me.btnExportNotUploaded.Name = "btnExportNotUploaded"
        Me.btnExportNotUploaded.Size = New System.Drawing.Size(80, 55)
        Me.btnExportNotUploaded.TabIndex = 80
        Me.btnExportNotUploaded.UseVisualStyleBackColor = False
        '
        'SelectAllItemsNotProcessed
        '
        Me.SelectAllItemsNotProcessed.Name = "SelectAllItemsNotProcessed"
        Me.SelectAllItemsNotProcessed.Size = New System.Drawing.Size(288, 22)
        Me.SelectAllItemsNotProcessed.Text = "Select All Items Not Uploaded"
        '
        'ExportItemsNotUploadedToPST
        '
        Me.ExportItemsNotUploadedToPST.Name = "ExportItemsNotUploadedToPST"
        Me.ExportItemsNotUploadedToPST.Size = New System.Drawing.Size(288, 22)
        Me.ExportItemsNotUploadedToPST.Text = "Export All Items Not Uploaded to PST"
        '
        'ItemsNotUploadedMenuStrip
        '
        Me.ItemsNotUploadedMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SelectAllItemsNotProcessed, Me.ExportItemsNotUploadedToPST, Me.ExportFTSDataTooLargeToolStripMenuItem, Me.ExportOtherNotUploadedToolStripMenuItem})
        Me.ItemsNotUploadedMenuStrip.Name = "ItemsNotUploadedMenuStrip"
        Me.ItemsNotUploadedMenuStrip.Size = New System.Drawing.Size(289, 92)
        '
        'ExportFTSDataTooLargeToolStripMenuItem
        '
        Me.ExportFTSDataTooLargeToolStripMenuItem.Name = "ExportFTSDataTooLargeToolStripMenuItem"
        Me.ExportFTSDataTooLargeToolStripMenuItem.Size = New System.Drawing.Size(288, 22)
        Me.ExportFTSDataTooLargeToolStripMenuItem.Text = "Export FTS Data Too Large Not Uploaded"
        '
        'ExportOtherNotUploadedToolStripMenuItem
        '
        Me.ExportOtherNotUploadedToolStripMenuItem.Name = "ExportOtherNotUploadedToolStripMenuItem"
        Me.ExportOtherNotUploadedToolStripMenuItem.Size = New System.Drawing.Size(288, 22)
        Me.ExportOtherNotUploadedToolStripMenuItem.Text = "Export Other Not Uploaded"
        '
        'btnConsolidateExporterFiles
        '
        Me.btnConsolidateExporterFiles.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnConsolidateExporterFiles.ForeColor = System.Drawing.Color.White
        Me.btnConsolidateExporterFiles.Image = CType(resources.GetObject("btnConsolidateExporterFiles.Image"), System.Drawing.Image)
        Me.btnConsolidateExporterFiles.Location = New System.Drawing.Point(1005, 618)
        Me.btnConsolidateExporterFiles.Name = "btnConsolidateExporterFiles"
        Me.btnConsolidateExporterFiles.Size = New System.Drawing.Size(80, 55)
        Me.btnConsolidateExporterFiles.TabIndex = 81
        Me.btnConsolidateExporterFiles.UseVisualStyleBackColor = True
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
        'frmNuixIngestion
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1442, 698)
        Me.Controls.Add(Me.btnConsolidateExporterFiles)
        Me.Controls.Add(Me.btnExportNotUploaded)
        Me.Controls.Add(Me.lblFilter)
        Me.Controls.Add(Me.btnExportProcessingDetails)
        Me.Controls.Add(Me.NuixLogoPicture)
        Me.Controls.Add(Me.txtFilter)
        Me.Controls.Add(Me.btnProcessingDetails)
        Me.Controls.Add(Me.btnLoadCustodianMappingData)
        Me.Controls.Add(Me.btnCustomerMappingFile)
        Me.Controls.Add(Me.txtCustodianMappingFile)
        Me.Controls.Add(Me.lblCustomerMappingFile)
        Me.Controls.Add(Me.btnRefreshGridData)
        Me.Controls.Add(Me.btnClearGrid)
        Me.Controls.Add(Me.btnLoadPreviousConfig)
        Me.Controls.Add(Me.cboGroupIDs)
        Me.Controls.Add(Me.lblSelectByGroupID)
        Me.Controls.Add(Me.btnSelectAll)
        Me.Controls.Add(Me.btnProcessandRunSelectedUsers)
        Me.Controls.Add(Me.btnLoadPSTInfo)
        Me.Controls.Add(Me.btnFileSystemChooser)
        Me.Controls.Add(Me.lblPSTLocation)
        Me.Controls.Add(Me.txtPSTLocation)
        Me.Controls.Add(Me.grdPSTInfo)
        Me.Controls.Add(Me.btnNuixSystemSettings)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(1400, 700)
        Me.Name = "frmNuixIngestion"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "EWS Ingestion"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.grdPSTInfo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CaseLockWatcher, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NuixLogoPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ItemsNotUploadedMenuStrip.ResumeLayout(False)
        Me.ConsolidateExporterFiles.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnNuixSystemSettings As System.Windows.Forms.Button
    Friend WithEvents grdPSTInfo As System.Windows.Forms.DataGridView
    Friend WithEvents txtPSTLocation As System.Windows.Forms.TextBox
    Friend WithEvents lblPSTLocation As System.Windows.Forms.Label
    Friend WithEvents btnFileSystemChooser As System.Windows.Forms.Button
    Friend WithEvents btnLoadPSTInfo As System.Windows.Forms.Button
    Friend WithEvents FileSystemWatcher1 As System.IO.FileSystemWatcher
    Friend WithEvents btnProcessandRunSelectedUsers As System.Windows.Forms.Button
    Friend WithEvents btnSelectAll As System.Windows.Forms.Button
    Friend WithEvents cboGroupIDs As System.Windows.Forms.ComboBox
    Friend WithEvents lblSelectByGroupID As System.Windows.Forms.Label
    Friend WithEvents btnLoadPreviousConfig As System.Windows.Forms.Button
    Friend WithEvents btnClearGrid As System.Windows.Forms.Button
    Friend WithEvents btnRefreshGridData As System.Windows.Forms.Button
    Friend WithEvents btnLoadCustodianMappingData As System.Windows.Forms.Button
    Friend WithEvents btnCustomerMappingFile As System.Windows.Forms.Button
    Friend WithEvents txtCustodianMappingFile As System.Windows.Forms.TextBox
    Friend WithEvents lblCustomerMappingFile As System.Windows.Forms.Label
    Friend WithEvents btnProcessingDetails As System.Windows.Forms.Button
    Friend WithEvents txtFilter As System.Windows.Forms.TextBox
    Friend WithEvents DataGridViewImageColumn1 As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents DataGridViewImageColumn2 As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents CaseLockWatcher As System.IO.FileSystemWatcher
    Friend WithEvents NuixLogoPicture As System.Windows.Forms.PictureBox
    Friend WithEvents EWSIngestToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents btnExportProcessingDetails As System.Windows.Forms.Button
    Friend WithEvents lblFilter As System.Windows.Forms.Label
    Friend WithEvents btnExportNotUploaded As System.Windows.Forms.Button
    Friend WithEvents SelectAllItemsNotProcessed As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExportItemsNotUploadedToPST As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ItemsNotUploadedMenuStrip As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents btnConsolidateExporterFiles As System.Windows.Forms.Button
    Friend WithEvents ConsolidateExporterFiles As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ConsolidateExporterMetricsFilesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ConsolidateExporterErrorsFilesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExportFTSDataTooLargeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExportOtherNotUploadedToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusImage As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents SelectCustodian As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CustodianName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PercentageCompleted As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BytesUploaded As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProgressStatus As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents NumberOfPSTs As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SizeOfPSTs As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PSTPath As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DestinationFolder As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DestinationRoot As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DestinationSMTP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProcessID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NumberSuccess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NumberFailed As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProcessingFilesDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CaseDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OutputDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ReprocessingDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LogDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IngestionStartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IngestionEndTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SummaryReportLocation As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
