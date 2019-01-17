<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class O365Puller
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(O365Puller))
        Me.grpNuixOutput = New System.Windows.Forms.GroupBox()
        Me.cboOutputType = New System.Windows.Forms.ComboBox()
        Me.lblCustodianSMTPCSVFile = New System.Windows.Forms.Label()
        Me.txtPullerSMTPCSVFile = New System.Windows.Forms.TextBox()
        Me.btnPullerCSVChooser = New System.Windows.Forms.Button()
        Me.grdCustodainSMTPInfo = New System.Windows.Forms.DataGridView()
        Me.StatusImage = New System.Windows.Forms.DataGridViewImageColumn()
        Me.SelectCustodian = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.SMTPAddress = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ExtractionRoot = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.GroupID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ExtractionStatus = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.OutputFormat = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.ExtractedItems = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ExtractedSize = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProcessID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProcessingFilesDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CaseDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OutputDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LogDirectory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ExtractionStartTime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ExtractionEndTime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SummaryReportLocation = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnLoadCustodianData = New System.Windows.Forms.Button()
        Me.cboPullerGroupIDs = New System.Windows.Forms.ComboBox()
        Me.lblSelectByGroupID = New System.Windows.Forms.Label()
        Me.btnSelectAll = New System.Windows.Forms.Button()
        Me.btnExtractShowSettings = New System.Windows.Forms.Button()
        Me.FromDateTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.lblFromDateRange = New System.Windows.Forms.Label()
        Me.lblToDateRange = New System.Windows.Forms.Label()
        Me.ToDateTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.btnLaunchO365Extractions = New System.Windows.Forms.Button()
        Me.O365ExtractionWatcher = New System.IO.FileSystemWatcher()
        Me.CaseLockWatcher = New System.IO.FileSystemWatcher()
        Me.btnLoadPreviousBatches = New System.Windows.Forms.Button()
        Me.NuixLogoPicture = New System.Windows.Forms.PictureBox()
        Me.EWSExtractToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnClearDates = New System.Windows.Forms.Button()
        Me.btnExportProcessingDetails = New System.Windows.Forms.Button()
        Me.grpNuixOutput.SuspendLayout()
        CType(Me.grdCustodainSMTPInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.O365ExtractionWatcher, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CaseLockWatcher, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NuixLogoPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpNuixOutput
        '
        Me.grpNuixOutput.Controls.Add(Me.cboOutputType)
        Me.grpNuixOutput.Location = New System.Drawing.Point(8, 8)
        Me.grpNuixOutput.Margin = New System.Windows.Forms.Padding(2)
        Me.grpNuixOutput.Name = "grpNuixOutput"
        Me.grpNuixOutput.Padding = New System.Windows.Forms.Padding(2)
        Me.grpNuixOutput.Size = New System.Drawing.Size(167, 43)
        Me.grpNuixOutput.TabIndex = 57
        Me.grpNuixOutput.TabStop = False
        Me.grpNuixOutput.Text = "Lightspeed Extraction Output"
        '
        'cboOutputType
        '
        Me.cboOutputType.FormattingEnabled = True
        Me.cboOutputType.Items.AddRange(New Object() {"pst", "msg", "eml"})
        Me.cboOutputType.Location = New System.Drawing.Point(11, 18)
        Me.cboOutputType.Margin = New System.Windows.Forms.Padding(2)
        Me.cboOutputType.Name = "cboOutputType"
        Me.cboOutputType.Size = New System.Drawing.Size(59, 21)
        Me.cboOutputType.TabIndex = 72
        '
        'lblCustodianSMTPCSVFile
        '
        Me.lblCustodianSMTPCSVFile.AutoSize = True
        Me.lblCustodianSMTPCSVFile.Location = New System.Drawing.Point(192, 26)
        Me.lblCustodianSMTPCSVFile.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblCustodianSMTPCSVFile.Name = "lblCustodianSMTPCSVFile"
        Me.lblCustodianSMTPCSVFile.Size = New System.Drawing.Size(114, 13)
        Me.lblCustodianSMTPCSVFile.TabIndex = 59
        Me.lblCustodianSMTPCSVFile.Text = "Custodian SMTP CSV:"
        '
        'txtPullerSMTPCSVFile
        '
        Me.txtPullerSMTPCSVFile.Location = New System.Drawing.Point(310, 24)
        Me.txtPullerSMTPCSVFile.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPullerSMTPCSVFile.Name = "txtPullerSMTPCSVFile"
        Me.txtPullerSMTPCSVFile.Size = New System.Drawing.Size(441, 20)
        Me.txtPullerSMTPCSVFile.TabIndex = 60
        '
        'btnPullerCSVChooser
        '
        Me.btnPullerCSVChooser.Location = New System.Drawing.Point(755, 23)
        Me.btnPullerCSVChooser.Margin = New System.Windows.Forms.Padding(2)
        Me.btnPullerCSVChooser.Name = "btnPullerCSVChooser"
        Me.btnPullerCSVChooser.Size = New System.Drawing.Size(27, 19)
        Me.btnPullerCSVChooser.TabIndex = 61
        Me.btnPullerCSVChooser.Text = "..."
        Me.btnPullerCSVChooser.UseVisualStyleBackColor = True
        '
        'grdCustodainSMTPInfo
        '
        Me.grdCustodainSMTPInfo.AllowUserToAddRows = False
        Me.grdCustodainSMTPInfo.AllowUserToDeleteRows = False
        Me.grdCustodainSMTPInfo.AllowUserToOrderColumns = True
        Me.grdCustodainSMTPInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grdCustodainSMTPInfo.BackgroundColor = System.Drawing.SystemColors.Control
        Me.grdCustodainSMTPInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdCustodainSMTPInfo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.StatusImage, Me.SelectCustodian, Me.SMTPAddress, Me.ExtractionRoot, Me.GroupID, Me.ExtractionStatus, Me.OutputFormat, Me.ExtractedItems, Me.ExtractedSize, Me.ProcessID, Me.ProcessingFilesDirectory, Me.CaseDirectory, Me.OutputDirectory, Me.LogDirectory, Me.ExtractionStartTime, Me.ExtractionEndTime, Me.SummaryReportLocation})
        Me.grdCustodainSMTPInfo.Location = New System.Drawing.Point(19, 120)
        Me.grdCustodainSMTPInfo.Margin = New System.Windows.Forms.Padding(2)
        Me.grdCustodainSMTPInfo.Name = "grdCustodainSMTPInfo"
        Me.grdCustodainSMTPInfo.RowTemplate.Height = 28
        Me.grdCustodainSMTPInfo.Size = New System.Drawing.Size(1383, 418)
        Me.grdCustodainSMTPInfo.TabIndex = 62
        '
        'StatusImage
        '
        Me.StatusImage.Frozen = True
        Me.StatusImage.HeaderText = ""
        Me.StatusImage.Name = "StatusImage"
        Me.StatusImage.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.StatusImage.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.StatusImage.Width = 30
        '
        'SelectCustodian
        '
        Me.SelectCustodian.Frozen = True
        Me.SelectCustodian.HeaderText = "Select Custodian"
        Me.SelectCustodian.Name = "SelectCustodian"
        Me.SelectCustodian.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.SelectCustodian.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.SelectCustodian.Width = 75
        '
        'SMTPAddress
        '
        Me.SMTPAddress.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.SMTPAddress.Frozen = True
        Me.SMTPAddress.HeaderText = "SMTPAddress"
        Me.SMTPAddress.Name = "SMTPAddress"
        '
        'ExtractionRoot
        '
        Me.ExtractionRoot.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ExtractionRoot.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox
        Me.ExtractionRoot.Frozen = True
        Me.ExtractionRoot.HeaderText = "Extraction Root"
        Me.ExtractionRoot.Items.AddRange(New Object() {"root_mailbox", "root_archive", "mailbox_purges", "archive_purges", "mailbox_recoverable", "archive_recoverable", "public folders", "mailbox/archive", "mailbox/mailbox_recoverable", "archive/archive_recoverable", "mailbox/mailbox_recoverable/archive/archive_recoverable"})
        Me.ExtractionRoot.Name = "ExtractionRoot"
        Me.ExtractionRoot.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ExtractionRoot.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ExtractionRoot.Width = 96
        '
        'GroupID
        '
        Me.GroupID.Frozen = True
        Me.GroupID.HeaderText = "Group ID"
        Me.GroupID.Name = "GroupID"
        Me.GroupID.Width = 50
        '
        'ExtractionStatus
        '
        Me.ExtractionStatus.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ExtractionStatus.Frozen = True
        Me.ExtractionStatus.HeaderText = "Extraction Status"
        Me.ExtractionStatus.Items.AddRange(New Object() {"", "Not Started", "In Progress", "Completed", "Cancelled by User", "Failed (review logs)", "Failed (No License Available)", "Failed (Case Exists)", "Failed (Export Exists)", "Failed (Case and Export Exists)", "Failed (Case Path Too Long)", "Restart Batch"})
        Me.ExtractionStatus.Name = "ExtractionStatus"
        Me.ExtractionStatus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ExtractionStatus.Width = 103
        '
        'OutputFormat
        '
        Me.OutputFormat.HeaderText = "Output Format"
        Me.OutputFormat.Items.AddRange(New Object() {"pst", "msg", "eml"})
        Me.OutputFormat.Name = "OutputFormat"
        Me.OutputFormat.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.OutputFormat.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'ExtractedItems
        '
        Me.ExtractedItems.HeaderText = "Extracted Item Count"
        Me.ExtractedItems.Name = "ExtractedItems"
        Me.ExtractedItems.ReadOnly = True
        Me.ExtractedItems.Width = 75
        '
        'ExtractedSize
        '
        Me.ExtractedSize.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ExtractedSize.HeaderText = "Extract Size"
        Me.ExtractedSize.Name = "ExtractedSize"
        Me.ExtractedSize.ReadOnly = True
        Me.ExtractedSize.Width = 81
        '
        'ProcessID
        '
        Me.ProcessID.HeaderText = "Process ID"
        Me.ProcessID.Name = "ProcessID"
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
        'ExtractionStartTime
        '
        Me.ExtractionStartTime.HeaderText = "Extraction Start Time"
        Me.ExtractionStartTime.Name = "ExtractionStartTime"
        '
        'ExtractionEndTime
        '
        Me.ExtractionEndTime.HeaderText = "Extraction End Time"
        Me.ExtractionEndTime.Name = "ExtractionEndTime"
        '
        'SummaryReportLocation
        '
        Me.SummaryReportLocation.HeaderText = "Summary Report Location"
        Me.SummaryReportLocation.Name = "SummaryReportLocation"
        Me.SummaryReportLocation.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'btnLoadCustodianData
        '
        Me.btnLoadCustodianData.Location = New System.Drawing.Point(785, 23)
        Me.btnLoadCustodianData.Margin = New System.Windows.Forms.Padding(2)
        Me.btnLoadCustodianData.Name = "btnLoadCustodianData"
        Me.btnLoadCustodianData.Size = New System.Drawing.Size(99, 19)
        Me.btnLoadCustodianData.TabIndex = 63
        Me.btnLoadCustodianData.Text = "Load SMTP Info"
        Me.btnLoadCustodianData.UseVisualStyleBackColor = True
        '
        'cboPullerGroupIDs
        '
        Me.cboPullerGroupIDs.FormattingEnabled = True
        Me.cboPullerGroupIDs.Location = New System.Drawing.Point(195, 88)
        Me.cboPullerGroupIDs.Margin = New System.Windows.Forms.Padding(2)
        Me.cboPullerGroupIDs.Name = "cboPullerGroupIDs"
        Me.cboPullerGroupIDs.Size = New System.Drawing.Size(67, 21)
        Me.cboPullerGroupIDs.TabIndex = 65
        '
        'lblSelectByGroupID
        '
        Me.lblSelectByGroupID.AutoSize = True
        Me.lblSelectByGroupID.Location = New System.Drawing.Point(82, 90)
        Me.lblSelectByGroupID.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblSelectByGroupID.Name = "lblSelectByGroupID"
        Me.lblSelectByGroupID.Size = New System.Drawing.Size(104, 13)
        Me.lblSelectByGroupID.TabIndex = 64
        Me.lblSelectByGroupID.Text = "Select By Group ID: "
        '
        'btnSelectAll
        '
        Me.btnSelectAll.Location = New System.Drawing.Point(8, 88)
        Me.btnSelectAll.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSelectAll.Name = "btnSelectAll"
        Me.btnSelectAll.Size = New System.Drawing.Size(70, 19)
        Me.btnSelectAll.TabIndex = 66
        Me.btnSelectAll.Text = "Select All"
        Me.btnSelectAll.UseVisualStyleBackColor = True
        '
        'btnExtractShowSettings
        '
        Me.btnExtractShowSettings.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnExtractShowSettings.BackColor = System.Drawing.Color.White
        Me.btnExtractShowSettings.Image = CType(resources.GetObject("btnExtractShowSettings.Image"), System.Drawing.Image)
        Me.btnExtractShowSettings.Location = New System.Drawing.Point(8, 542)
        Me.btnExtractShowSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExtractShowSettings.Name = "btnExtractShowSettings"
        Me.btnExtractShowSettings.Size = New System.Drawing.Size(80, 60)
        Me.btnExtractShowSettings.TabIndex = 67
        Me.btnExtractShowSettings.UseVisualStyleBackColor = False
        '
        'FromDateTimePicker
        '
        Me.FromDateTimePicker.Location = New System.Drawing.Point(79, 60)
        Me.FromDateTimePicker.Margin = New System.Windows.Forms.Padding(2)
        Me.FromDateTimePicker.Name = "FromDateTimePicker"
        Me.FromDateTimePicker.Size = New System.Drawing.Size(200, 20)
        Me.FromDateTimePicker.TabIndex = 77
        '
        'lblFromDateRange
        '
        Me.lblFromDateRange.AutoSize = True
        Me.lblFromDateRange.Location = New System.Drawing.Point(16, 60)
        Me.lblFromDateRange.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblFromDateRange.Name = "lblFromDateRange"
        Me.lblFromDateRange.Size = New System.Drawing.Size(59, 13)
        Me.lblFromDateRange.TabIndex = 78
        Me.lblFromDateRange.Text = "From Date:"
        '
        'lblToDateRange
        '
        Me.lblToDateRange.AutoSize = True
        Me.lblToDateRange.Location = New System.Drawing.Point(309, 60)
        Me.lblToDateRange.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblToDateRange.Name = "lblToDateRange"
        Me.lblToDateRange.Size = New System.Drawing.Size(49, 13)
        Me.lblToDateRange.TabIndex = 79
        Me.lblToDateRange.Text = "To Date:"
        '
        'ToDateTimePicker
        '
        Me.ToDateTimePicker.Location = New System.Drawing.Point(362, 60)
        Me.ToDateTimePicker.Margin = New System.Windows.Forms.Padding(2)
        Me.ToDateTimePicker.Name = "ToDateTimePicker"
        Me.ToDateTimePicker.Size = New System.Drawing.Size(200, 20)
        Me.ToDateTimePicker.TabIndex = 80
        '
        'btnLaunchO365Extractions
        '
        Me.btnLaunchO365Extractions.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLaunchO365Extractions.BackColor = System.Drawing.Color.White
        Me.btnLaunchO365Extractions.Image = CType(resources.GetObject("btnLaunchO365Extractions.Image"), System.Drawing.Image)
        Me.btnLaunchO365Extractions.Location = New System.Drawing.Point(1311, 542)
        Me.btnLaunchO365Extractions.Margin = New System.Windows.Forms.Padding(2)
        Me.btnLaunchO365Extractions.Name = "btnLaunchO365Extractions"
        Me.btnLaunchO365Extractions.Size = New System.Drawing.Size(80, 60)
        Me.btnLaunchO365Extractions.TabIndex = 81
        Me.btnLaunchO365Extractions.UseVisualStyleBackColor = False
        '
        'O365ExtractionWatcher
        '
        Me.O365ExtractionWatcher.EnableRaisingEvents = True
        Me.O365ExtractionWatcher.Filter = "summary-report.txt"
        Me.O365ExtractionWatcher.IncludeSubdirectories = True
        Me.O365ExtractionWatcher.SynchronizingObject = Me
        '
        'CaseLockWatcher
        '
        Me.CaseLockWatcher.EnableRaisingEvents = True
        Me.CaseLockWatcher.Filter = "case.lock"
        Me.CaseLockWatcher.IncludeSubdirectories = True
        Me.CaseLockWatcher.SynchronizingObject = Me
        '
        'btnLoadPreviousBatches
        '
        Me.btnLoadPreviousBatches.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnLoadPreviousBatches.BackColor = System.Drawing.Color.White
        Me.btnLoadPreviousBatches.Image = CType(resources.GetObject("btnLoadPreviousBatches.Image"), System.Drawing.Image)
        Me.btnLoadPreviousBatches.Location = New System.Drawing.Point(95, 542)
        Me.btnLoadPreviousBatches.Name = "btnLoadPreviousBatches"
        Me.btnLoadPreviousBatches.Size = New System.Drawing.Size(80, 60)
        Me.btnLoadPreviousBatches.TabIndex = 82
        Me.btnLoadPreviousBatches.UseVisualStyleBackColor = False
        '
        'NuixLogoPicture
        '
        Me.NuixLogoPicture.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.NuixLogoPicture.Image = CType(resources.GetObject("NuixLogoPicture.Image"), System.Drawing.Image)
        Me.NuixLogoPicture.Location = New System.Drawing.Point(1316, 15)
        Me.NuixLogoPicture.Name = "NuixLogoPicture"
        Me.NuixLogoPicture.Size = New System.Drawing.Size(75, 100)
        Me.NuixLogoPicture.TabIndex = 83
        Me.NuixLogoPicture.TabStop = False
        '
        'EWSExtractToolTip
        '
        Me.EWSExtractToolTip.IsBalloon = True
        '
        'btnClearDates
        '
        Me.btnClearDates.Location = New System.Drawing.Point(567, 60)
        Me.btnClearDates.Name = "btnClearDates"
        Me.btnClearDates.Size = New System.Drawing.Size(75, 23)
        Me.btnClearDates.TabIndex = 84
        Me.btnClearDates.Text = "Clear"
        Me.btnClearDates.UseVisualStyleBackColor = True
        '
        'btnExportProcessingDetails
        '
        Me.btnExportProcessingDetails.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportProcessingDetails.BackColor = System.Drawing.Color.White
        Me.btnExportProcessingDetails.Image = CType(resources.GetObject("btnExportProcessingDetails.Image"), System.Drawing.Image)
        Me.btnExportProcessingDetails.Location = New System.Drawing.Point(1225, 542)
        Me.btnExportProcessingDetails.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExportProcessingDetails.Name = "btnExportProcessingDetails"
        Me.btnExportProcessingDetails.Size = New System.Drawing.Size(80, 60)
        Me.btnExportProcessingDetails.TabIndex = 85
        Me.btnExportProcessingDetails.UseVisualStyleBackColor = False
        '
        'O365Puller
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(1411, 628)
        Me.Controls.Add(Me.btnExportProcessingDetails)
        Me.Controls.Add(Me.btnClearDates)
        Me.Controls.Add(Me.NuixLogoPicture)
        Me.Controls.Add(Me.btnLoadPreviousBatches)
        Me.Controls.Add(Me.btnLaunchO365Extractions)
        Me.Controls.Add(Me.ToDateTimePicker)
        Me.Controls.Add(Me.lblToDateRange)
        Me.Controls.Add(Me.lblFromDateRange)
        Me.Controls.Add(Me.FromDateTimePicker)
        Me.Controls.Add(Me.btnExtractShowSettings)
        Me.Controls.Add(Me.btnSelectAll)
        Me.Controls.Add(Me.cboPullerGroupIDs)
        Me.Controls.Add(Me.lblSelectByGroupID)
        Me.Controls.Add(Me.btnLoadCustodianData)
        Me.Controls.Add(Me.grdCustodainSMTPInfo)
        Me.Controls.Add(Me.btnPullerCSVChooser)
        Me.Controls.Add(Me.txtPullerSMTPCSVFile)
        Me.Controls.Add(Me.lblCustodianSMTPCSVFile)
        Me.Controls.Add(Me.grpNuixOutput)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(1150, 600)
        Me.Name = "O365Puller"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "EWS Extraction"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.grpNuixOutput.ResumeLayout(False)
        CType(Me.grdCustodainSMTPInfo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.O365ExtractionWatcher, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CaseLockWatcher, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NuixLogoPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grpNuixOutput As System.Windows.Forms.GroupBox
    Friend WithEvents cboOutputType As System.Windows.Forms.ComboBox
    Friend WithEvents lblCustodianSMTPCSVFile As System.Windows.Forms.Label
    Friend WithEvents txtPullerSMTPCSVFile As System.Windows.Forms.TextBox
    Friend WithEvents btnPullerCSVChooser As System.Windows.Forms.Button
    Friend WithEvents grdCustodainSMTPInfo As System.Windows.Forms.DataGridView
    Friend WithEvents btnLoadCustodianData As System.Windows.Forms.Button
    Friend WithEvents cboPullerGroupIDs As System.Windows.Forms.ComboBox
    Friend WithEvents lblSelectByGroupID As System.Windows.Forms.Label
    Friend WithEvents btnSelectAll As System.Windows.Forms.Button
    Friend WithEvents btnExtractShowSettings As System.Windows.Forms.Button
    Friend WithEvents FromDateTimePicker As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblFromDateRange As System.Windows.Forms.Label
    Friend WithEvents lblToDateRange As System.Windows.Forms.Label
    Friend WithEvents ToDateTimePicker As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnLaunchO365Extractions As System.Windows.Forms.Button
    Friend WithEvents O365ExtractionWatcher As System.IO.FileSystemWatcher
    Friend WithEvents CaseLockWatcher As System.IO.FileSystemWatcher
    Friend WithEvents btnLoadPreviousBatches As System.Windows.Forms.Button
    Friend WithEvents NuixLogoPicture As System.Windows.Forms.PictureBox
    Friend WithEvents EWSExtractToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents btnClearDates As System.Windows.Forms.Button
    Friend WithEvents btnExportProcessingDetails As System.Windows.Forms.Button
    Friend WithEvents StatusImage As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents SelectCustodian As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents SMTPAddress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ExtractionRoot As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents GroupID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ExtractionStatus As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents OutputFormat As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ExtractedItems As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ExtractedSize As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProcessID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProcessingFilesDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CaseDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OutputDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LogDirectory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ExtractionStartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ExtractionEndTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SummaryReportLocation As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
