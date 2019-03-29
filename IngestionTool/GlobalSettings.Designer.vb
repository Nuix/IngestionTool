<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class O365ExtractionSettings
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(O365ExtractionSettings))
        Me.tabParallelProcessingSettings = New System.Windows.Forms.TabControl()
        Me.tabLicensingSettings = New System.Windows.Forms.TabPage()
        Me.grpNMSSettings = New System.Windows.Forms.GroupBox()
        Me.txtNMSPort = New System.Windows.Forms.TextBox()
        Me.lblNMSPort = New System.Windows.Forms.Label()
        Me.lblSourceType = New System.Windows.Forms.Label()
        Me.cboSourceType = New System.Windows.Forms.ComboBox()
        Me.txtNMSPassword = New System.Windows.Forms.TextBox()
        Me.txtNMSUsername = New System.Windows.Forms.TextBox()
        Me.txtNMSAddress = New System.Windows.Forms.TextBox()
        Me.lblNMSPassword = New System.Windows.Forms.Label()
        Me.lblNMSUsername = New System.Windows.Forms.Label()
        Me.lblNuixManagementServerAddress = New System.Windows.Forms.Label()
        Me.tabNuixParallelProcessingSettings = New System.Windows.Forms.TabPage()
        Me.grpNuixDataProcessingSettings = New System.Windows.Forms.GroupBox()
        Me.btnNuixFilesDirSel = New System.Windows.Forms.Button()
        Me.txtNuixFilesDirectory = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnExportDir = New System.Windows.Forms.Button()
        Me.txtExportDir = New System.Windows.Forms.TextBox()
        Me.lblExportDir = New System.Windows.Forms.Label()
        Me.btnWorkerTempDir = New System.Windows.Forms.Button()
        Me.txtWorkerTempDir = New System.Windows.Forms.TextBox()
        Me.lblWorkerTempDir = New System.Windows.Forms.Label()
        Me.btnNuixAppLocation = New System.Windows.Forms.Button()
        Me.txtNuixAppLocation = New System.Windows.Forms.TextBox()
        Me.lblNuixAppLocation = New System.Windows.Forms.Label()
        Me.btnJavaTempDir = New System.Windows.Forms.Button()
        Me.btnLogDir = New System.Windows.Forms.Button()
        Me.btnCaseDirSel = New System.Windows.Forms.Button()
        Me.txtNuixCaseDir = New System.Windows.Forms.TextBox()
        Me.lblCaseDirectory = New System.Windows.Forms.Label()
        Me.txtJavaTempDir = New System.Windows.Forms.TextBox()
        Me.lblJavaTempDir = New System.Windows.Forms.Label()
        Me.lblLogDirectory = New System.Windows.Forms.Label()
        Me.txtLogDirectory = New System.Windows.Forms.TextBox()
        Me.tabNuixDataSettings = New System.Windows.Forms.TabPage()
        Me.grpDataProcessingSettings = New System.Windows.Forms.GroupBox()
        Me.grpDigestSettings = New System.Windows.Forms.GroupBox()
        Me.lblDigestToCompute = New System.Windows.Forms.Label()
        Me.grpEmailDigestSettings = New System.Windows.Forms.GroupBox()
        Me.chkIncludeItemDate = New System.Windows.Forms.CheckBox()
        Me.chkIncludeBCC = New System.Windows.Forms.CheckBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.MaxDigestSize = New System.Windows.Forms.NumericUpDown()
        Me.chkSSDeep = New System.Windows.Forms.CheckBox()
        Me.chkSHA256 = New System.Windows.Forms.CheckBox()
        Me.chkSHA1 = New System.Windows.Forms.CheckBox()
        Me.chkMD5Digest = New System.Windows.Forms.CheckBox()
        Me.grpImageSettings = New System.Windows.Forms.GroupBox()
        Me.chkPerformimagecolouranalysis = New System.Windows.Forms.CheckBox()
        Me.chkGeneratethumbnailsforimagedata = New System.Windows.Forms.CheckBox()
        Me.grpItemContentsSettings = New System.Windows.Forms.GroupBox()
        Me.grpNamedEntitySettings = New System.Windows.Forms.GroupBox()
        Me.chkExtractNamedEntitiesfromProperties = New System.Windows.Forms.CheckBox()
        Me.chkIncludeTextStrippedItems = New System.Windows.Forms.CheckBox()
        Me.chkExtractNamedEntities = New System.Windows.Forms.CheckBox()
        Me.chkEnableTextSummarisation = New System.Windows.Forms.CheckBox()
        Me.chkEnableNearDuplicates = New System.Windows.Forms.CheckBox()
        Me.chkProcesstext = New System.Windows.Forms.CheckBox()
        Me.grpTextIndexingSettings = New System.Windows.Forms.GroupBox()
        Me.chkEnableExactQueries = New System.Windows.Forms.CheckBox()
        Me.chkUsestemming = New System.Windows.Forms.CheckBox()
        Me.chkUseStopWords = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboAnalysisLang = New System.Windows.Forms.ComboBox()
        Me.grpFamilyTextSettings = New System.Windows.Forms.GroupBox()
        Me.chkHideimmaterialItems = New System.Windows.Forms.CheckBox()
        Me.chkCreatefamilySearch = New System.Windows.Forms.CheckBox()
        Me.grpDeletefileSettings = New System.Windows.Forms.GroupBox()
        Me.chkCarveFSunallocated = New System.Windows.Forms.CheckBox()
        Me.chkExtractmailboxslackspace = New System.Windows.Forms.CheckBox()
        Me.chkSmartprocessMSRegistry = New System.Windows.Forms.CheckBox()
        Me.chkExtractEndOfFileSpace = New System.Windows.Forms.CheckBox()
        Me.chkRecoverDeleteFiles = New System.Windows.Forms.CheckBox()
        Me.grpEvidenceSettings = New System.Windows.Forms.GroupBox()
        Me.Maxbinarysize = New System.Windows.Forms.NumericUpDown()
        Me.lblMaxbinarysize = New System.Windows.Forms.Label()
        Me.chkStorebinary = New System.Windows.Forms.CheckBox()
        Me.chkCalculateAuditedSize = New System.Windows.Forms.CheckBox()
        Me.chkReuseEvidenceStores = New System.Windows.Forms.CheckBox()
        Me.lblTraversal = New System.Windows.Forms.Label()
        Me.cboTraversal = New System.Windows.Forms.ComboBox()
        Me.chkCalculateProcessingSize = New System.Windows.Forms.CheckBox()
        Me.chkPerformItemIdentification = New System.Windows.Forms.CheckBox()
        Me.tabArchiveExtractionSettings = New System.Windows.Forms.TabPage()
        Me.grpArchiveExtractionSettings = New System.Windows.Forms.GroupBox()
        Me.numExtractWorkerTimeout = New System.Windows.Forms.NumericUpDown()
        Me.lblExtractWorkerTimeout = New System.Windows.Forms.Label()
        Me.grpEMLExportOptions = New System.Windows.Forms.GroupBox()
        Me.chkEMLAddDistributionListMetadata = New System.Windows.Forms.CheckBox()
        Me.grpPSTExportOptions = New System.Windows.Forms.GroupBox()
        Me.chkPSTAddDistributionListMetadata = New System.Windows.Forms.CheckBox()
        Me.lblPSTExportSize = New System.Windows.Forms.Label()
        Me.numPSTExportSize = New System.Windows.Forms.NumericUpDown()
        Me.txtMemoryPerNuixInstance = New System.Windows.Forms.TextBox()
        Me.txtMemoryPerWorker = New System.Windows.Forms.TextBox()
        Me.lbNuixAppMemory = New System.Windows.Forms.Label()
        Me.lblAvailableRAMAfterNuix = New System.Windows.Forms.Label()
        Me.lblAvailableAfterNuix = New System.Windows.Forms.Label()
        Me.lblNuixInstances = New System.Windows.Forms.Label()
        Me.cboNumberOfNuixInstances = New System.Windows.Forms.ComboBox()
        Me.lblMemoryPerWorker = New System.Windows.Forms.Label()
        Me.cboNuixWorkers = New System.Windows.Forms.ComboBox()
        Me.lblNumberOfNuixWorkers = New System.Windows.Forms.Label()
        Me.lblAvailableRAM = New System.Windows.Forms.Label()
        Me.lblRAM = New System.Windows.Forms.Label()
        Me.lblAvailableSystemMemory = New System.Windows.Forms.Label()
        Me.lblTotalSystemMemory = New System.Windows.Forms.Label()
        Me.tabEWSSettings = New System.Windows.Forms.TabPage()
        Me.grpOffice365Settings = New System.Windows.Forms.GroupBox()
        Me.lblPSTConsolidation = New System.Windows.Forms.Label()
        Me.cboPSTConsolidation = New System.Windows.Forms.ComboBox()
        Me.grpEWSDownloadControl = New System.Windows.Forms.GroupBox()
        Me.chkEnableMailboxSlackspace = New System.Windows.Forms.CheckBox()
        Me.chkCollaborativePrefetching = New System.Windows.Forms.CheckBox()
        Me.chkEnablePrefetch = New System.Windows.Forms.CheckBox()
        Me.numDownloadCount = New System.Windows.Forms.NumericUpDown()
        Me.lblMaxDownloadCount = New System.Windows.Forms.Label()
        Me.numDownloadSize = New System.Windows.Forms.NumericUpDown()
        Me.lblMaxMessageDownloadSize = New System.Windows.Forms.Label()
        Me.grpEWSUploadControl = New System.Windows.Forms.GroupBox()
        Me.lblBulkUploadsze = New System.Windows.Forms.Label()
        Me.numBulkUploadSize = New System.Windows.Forms.NumericUpDown()
        Me.chkEnableBulkUpload = New System.Windows.Forms.CheckBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtO365RemovePathPrefix = New System.Windows.Forms.TextBox()
        Me.lblRemovePathPrefix = New System.Windows.Forms.Label()
        Me.lblMaxItemUploadSize = New System.Windows.Forms.Label()
        Me.numEWSMaxUploadSize = New System.Windows.Forms.NumericUpDown()
        Me.grpEWSThrottling = New System.Windows.Forms.GroupBox()
        Me.numEWSRetryIncrement = New System.Windows.Forms.NumericUpDown()
        Me.numEWSRetryDelay = New System.Windows.Forms.NumericUpDown()
        Me.numEWSRetryCount = New System.Windows.Forms.NumericUpDown()
        Me.lblRetryIncrement = New System.Windows.Forms.Label()
        Me.lblRetryDelay = New System.Windows.Forms.Label()
        Me.lblRetryCount = New System.Windows.Forms.Label()
        Me.grpNuixInstanceSettings = New System.Windows.Forms.GroupBox()
        Me.numWorkerTimeout = New System.Windows.Forms.NumericUpDown()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtO365NuixAppMemory = New System.Windows.Forms.TextBox()
        Me.txtO365MemoryPerWorker = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblO365AvailableMemoryAfterNuix = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.cboO365NumberOfNuixInstances = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cboO365NumberOfNuixWorkers = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lblO365AvailMemory = New System.Windows.Forms.Label()
        Me.lblO365SystemMemory = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.chkO365ApplicationImpersonation = New System.Windows.Forms.CheckBox()
        Me.txtO365AdminInfo = New System.Windows.Forms.TextBox()
        Me.txtO365AdminUserName = New System.Windows.Forms.TextBox()
        Me.txtO365Domain = New System.Windows.Forms.TextBox()
        Me.txtO365ExchangeServer = New System.Windows.Forms.TextBox()
        Me.lblPSTInfoAdminInfo = New System.Windows.Forms.Label()
        Me.lblPSTInfoAdminUserName = New System.Windows.Forms.Label()
        Me.lblPSTInfoDomain = New System.Windows.Forms.Label()
        Me.lblPSTInfoExchangeServer = New System.Windows.Forms.Label()
        Me.tabDB = New System.Windows.Forms.TabPage()
        Me.grpRedisConfig = New System.Windows.Forms.GroupBox()
        Me.txtRedisAuth = New System.Windows.Forms.TextBox()
        Me.lblRedis = New System.Windows.Forms.Label()
        Me.txtRedisPort = New System.Windows.Forms.TextBox()
        Me.lblRedisPort = New System.Windows.Forms.Label()
        Me.txtRedisHostName = New System.Windows.Forms.TextBox()
        Me.lblRedisHostName = New System.Windows.Forms.Label()
        Me.btnSQLiteDBLocationSelection = New System.Windows.Forms.Button()
        Me.txtSQLiteDBLocation = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnSaveSettings = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.txtSaveSettingLocation = New System.Windows.Forms.TextBox()
        Me.lblSaveSettingsLocation = New System.Windows.Forms.Label()
        Me.btnSettingsLocation = New System.Windows.Forms.Button()
        Me.btnLoadSettingXML = New System.Windows.Forms.Button()
        Me.btnSettingsOK = New System.Windows.Forms.Button()
        Me.GlobalSettingsToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.tabParallelProcessingSettings.SuspendLayout()
        Me.tabLicensingSettings.SuspendLayout()
        Me.grpNMSSettings.SuspendLayout()
        Me.tabNuixParallelProcessingSettings.SuspendLayout()
        Me.grpNuixDataProcessingSettings.SuspendLayout()
        Me.tabNuixDataSettings.SuspendLayout()
        Me.grpDataProcessingSettings.SuspendLayout()
        Me.grpDigestSettings.SuspendLayout()
        Me.grpEmailDigestSettings.SuspendLayout()
        CType(Me.MaxDigestSize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpImageSettings.SuspendLayout()
        Me.grpItemContentsSettings.SuspendLayout()
        Me.grpNamedEntitySettings.SuspendLayout()
        Me.grpTextIndexingSettings.SuspendLayout()
        Me.grpFamilyTextSettings.SuspendLayout()
        Me.grpDeletefileSettings.SuspendLayout()
        Me.grpEvidenceSettings.SuspendLayout()
        CType(Me.Maxbinarysize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabArchiveExtractionSettings.SuspendLayout()
        Me.grpArchiveExtractionSettings.SuspendLayout()
        CType(Me.numExtractWorkerTimeout, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpEMLExportOptions.SuspendLayout()
        Me.grpPSTExportOptions.SuspendLayout()
        CType(Me.numPSTExportSize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabEWSSettings.SuspendLayout()
        Me.grpOffice365Settings.SuspendLayout()
        Me.grpEWSDownloadControl.SuspendLayout()
        CType(Me.numDownloadCount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numDownloadSize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpEWSUploadControl.SuspendLayout()
        CType(Me.numBulkUploadSize, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numEWSMaxUploadSize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpEWSThrottling.SuspendLayout()
        CType(Me.numEWSRetryIncrement, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numEWSRetryDelay, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numEWSRetryCount, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpNuixInstanceSettings.SuspendLayout()
        CType(Me.numWorkerTimeout, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabDB.SuspendLayout()
        Me.grpRedisConfig.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabParallelProcessingSettings
        '
        Me.tabParallelProcessingSettings.Controls.Add(Me.tabLicensingSettings)
        Me.tabParallelProcessingSettings.Controls.Add(Me.tabNuixParallelProcessingSettings)
        Me.tabParallelProcessingSettings.Controls.Add(Me.tabNuixDataSettings)
        Me.tabParallelProcessingSettings.Controls.Add(Me.tabArchiveExtractionSettings)
        Me.tabParallelProcessingSettings.Controls.Add(Me.tabEWSSettings)
        Me.tabParallelProcessingSettings.Controls.Add(Me.tabDB)
        Me.tabParallelProcessingSettings.Location = New System.Drawing.Point(6, 8)
        Me.tabParallelProcessingSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.tabParallelProcessingSettings.Name = "tabParallelProcessingSettings"
        Me.tabParallelProcessingSettings.SelectedIndex = 0
        Me.tabParallelProcessingSettings.Size = New System.Drawing.Size(700, 448)
        Me.tabParallelProcessingSettings.TabIndex = 0
        '
        'tabLicensingSettings
        '
        Me.tabLicensingSettings.Controls.Add(Me.grpNMSSettings)
        Me.tabLicensingSettings.Location = New System.Drawing.Point(4, 22)
        Me.tabLicensingSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.tabLicensingSettings.Name = "tabLicensingSettings"
        Me.tabLicensingSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.tabLicensingSettings.Size = New System.Drawing.Size(692, 422)
        Me.tabLicensingSettings.TabIndex = 3
        Me.tabLicensingSettings.Text = "Nuix License"
        Me.tabLicensingSettings.UseVisualStyleBackColor = True
        '
        'grpNMSSettings
        '
        Me.grpNMSSettings.Controls.Add(Me.txtNMSPort)
        Me.grpNMSSettings.Controls.Add(Me.lblNMSPort)
        Me.grpNMSSettings.Controls.Add(Me.lblSourceType)
        Me.grpNMSSettings.Controls.Add(Me.cboSourceType)
        Me.grpNMSSettings.Controls.Add(Me.txtNMSPassword)
        Me.grpNMSSettings.Controls.Add(Me.txtNMSUsername)
        Me.grpNMSSettings.Controls.Add(Me.txtNMSAddress)
        Me.grpNMSSettings.Controls.Add(Me.lblNMSPassword)
        Me.grpNMSSettings.Controls.Add(Me.lblNMSUsername)
        Me.grpNMSSettings.Controls.Add(Me.lblNuixManagementServerAddress)
        Me.grpNMSSettings.Location = New System.Drawing.Point(7, 7)
        Me.grpNMSSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpNMSSettings.Name = "grpNMSSettings"
        Me.grpNMSSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpNMSSettings.Size = New System.Drawing.Size(683, 416)
        Me.grpNMSSettings.TabIndex = 0
        Me.grpNMSSettings.TabStop = False
        Me.grpNMSSettings.Text = "Nuix License"
        '
        'txtNMSPort
        '
        Me.txtNMSPort.Location = New System.Drawing.Point(282, 52)
        Me.txtNMSPort.Name = "txtNMSPort"
        Me.txtNMSPort.Size = New System.Drawing.Size(59, 20)
        Me.txtNMSPort.TabIndex = 73
        '
        'lblNMSPort
        '
        Me.lblNMSPort.AutoSize = True
        Me.lblNMSPort.Location = New System.Drawing.Point(220, 52)
        Me.lblNMSPort.Name = "lblNMSPort"
        Me.lblNMSPort.Size = New System.Drawing.Size(56, 13)
        Me.lblNMSPort.TabIndex = 72
        Me.lblNMSPort.Text = "NMS Port:"
        '
        'lblSourceType
        '
        Me.lblSourceType.AutoSize = True
        Me.lblSourceType.Location = New System.Drawing.Point(4, 24)
        Me.lblSourceType.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblSourceType.Name = "lblSourceType"
        Me.lblSourceType.Size = New System.Drawing.Size(68, 13)
        Me.lblSourceType.TabIndex = 71
        Me.lblSourceType.Text = "Source Type"
        '
        'cboSourceType
        '
        Me.cboSourceType.FormattingEnabled = True
        Me.cboSourceType.Items.AddRange(New Object() {"", "Desktop", "Server", "Cloud Server"})
        Me.cboSourceType.Location = New System.Drawing.Point(100, 22)
        Me.cboSourceType.Margin = New System.Windows.Forms.Padding(2)
        Me.cboSourceType.Name = "cboSourceType"
        Me.cboSourceType.Size = New System.Drawing.Size(82, 21)
        Me.cboSourceType.TabIndex = 70
        '
        'txtNMSPassword
        '
        Me.txtNMSPassword.Location = New System.Drawing.Point(100, 104)
        Me.txtNMSPassword.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNMSPassword.Name = "txtNMSPassword"
        Me.txtNMSPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtNMSPassword.Size = New System.Drawing.Size(101, 20)
        Me.txtNMSPassword.TabIndex = 69
        '
        'txtNMSUsername
        '
        Me.txtNMSUsername.Location = New System.Drawing.Point(100, 78)
        Me.txtNMSUsername.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNMSUsername.Name = "txtNMSUsername"
        Me.txtNMSUsername.Size = New System.Drawing.Size(101, 20)
        Me.txtNMSUsername.TabIndex = 68
        '
        'txtNMSAddress
        '
        Me.txtNMSAddress.Location = New System.Drawing.Point(100, 52)
        Me.txtNMSAddress.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNMSAddress.Name = "txtNMSAddress"
        Me.txtNMSAddress.Size = New System.Drawing.Size(101, 20)
        Me.txtNMSAddress.TabIndex = 64
        '
        'lblNMSPassword
        '
        Me.lblNMSPassword.AutoSize = True
        Me.lblNMSPassword.Location = New System.Drawing.Point(3, 104)
        Me.lblNMSPassword.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNMSPassword.Name = "lblNMSPassword"
        Me.lblNMSPassword.Size = New System.Drawing.Size(83, 13)
        Me.lblNMSPassword.TabIndex = 67
        Me.lblNMSPassword.Text = "NMS Password:"
        '
        'lblNMSUsername
        '
        Me.lblNMSUsername.AutoSize = True
        Me.lblNMSUsername.Location = New System.Drawing.Point(3, 78)
        Me.lblNMSUsername.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNMSUsername.Name = "lblNMSUsername"
        Me.lblNMSUsername.Size = New System.Drawing.Size(85, 13)
        Me.lblNMSUsername.TabIndex = 66
        Me.lblNMSUsername.Text = "NMS Username:"
        '
        'lblNuixManagementServerAddress
        '
        Me.lblNuixManagementServerAddress.AutoSize = True
        Me.lblNuixManagementServerAddress.Location = New System.Drawing.Point(3, 52)
        Me.lblNuixManagementServerAddress.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNuixManagementServerAddress.Name = "lblNuixManagementServerAddress"
        Me.lblNuixManagementServerAddress.Size = New System.Drawing.Size(85, 13)
        Me.lblNuixManagementServerAddress.TabIndex = 65
        Me.lblNuixManagementServerAddress.Text = "NMS Hostname:"
        '
        'tabNuixParallelProcessingSettings
        '
        Me.tabNuixParallelProcessingSettings.Controls.Add(Me.grpNuixDataProcessingSettings)
        Me.tabNuixParallelProcessingSettings.Location = New System.Drawing.Point(4, 22)
        Me.tabNuixParallelProcessingSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.tabNuixParallelProcessingSettings.Name = "tabNuixParallelProcessingSettings"
        Me.tabNuixParallelProcessingSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.tabNuixParallelProcessingSettings.Size = New System.Drawing.Size(692, 422)
        Me.tabNuixParallelProcessingSettings.TabIndex = 0
        Me.tabNuixParallelProcessingSettings.Text = "Nuix Directories"
        Me.tabNuixParallelProcessingSettings.UseVisualStyleBackColor = True
        '
        'grpNuixDataProcessingSettings
        '
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.btnNuixFilesDirSel)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.txtNuixFilesDirectory)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.Label2)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.btnExportDir)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.txtExportDir)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.lblExportDir)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.btnWorkerTempDir)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.txtWorkerTempDir)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.lblWorkerTempDir)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.btnNuixAppLocation)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.txtNuixAppLocation)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.lblNuixAppLocation)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.btnJavaTempDir)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.btnLogDir)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.btnCaseDirSel)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.txtNuixCaseDir)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.lblCaseDirectory)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.txtJavaTempDir)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.lblJavaTempDir)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.lblLogDirectory)
        Me.grpNuixDataProcessingSettings.Controls.Add(Me.txtLogDirectory)
        Me.grpNuixDataProcessingSettings.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.grpNuixDataProcessingSettings.Location = New System.Drawing.Point(2, 4)
        Me.grpNuixDataProcessingSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpNuixDataProcessingSettings.Name = "grpNuixDataProcessingSettings"
        Me.grpNuixDataProcessingSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpNuixDataProcessingSettings.Size = New System.Drawing.Size(689, 419)
        Me.grpNuixDataProcessingSettings.TabIndex = 56
        Me.grpNuixDataProcessingSettings.TabStop = False
        Me.grpNuixDataProcessingSettings.Text = "Nuix Directories"
        '
        'btnNuixFilesDirSel
        '
        Me.btnNuixFilesDirSel.Location = New System.Drawing.Point(443, 43)
        Me.btnNuixFilesDirSel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnNuixFilesDirSel.Name = "btnNuixFilesDirSel"
        Me.btnNuixFilesDirSel.Size = New System.Drawing.Size(27, 19)
        Me.btnNuixFilesDirSel.TabIndex = 61
        Me.btnNuixFilesDirSel.Text = "..."
        Me.btnNuixFilesDirSel.UseVisualStyleBackColor = True
        '
        'txtNuixFilesDirectory
        '
        Me.txtNuixFilesDirectory.Location = New System.Drawing.Point(178, 44)
        Me.txtNuixFilesDirectory.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNuixFilesDirectory.Name = "txtNuixFilesDirectory"
        Me.txtNuixFilesDirectory.Size = New System.Drawing.Size(261, 20)
        Me.txtNuixFilesDirectory.TabIndex = 60
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 44)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(155, 13)
        Me.Label2.TabIndex = 62
        Me.Label2.Text = "Nuix Processing Files Directory:"
        '
        'btnExportDir
        '
        Me.btnExportDir.Location = New System.Drawing.Point(443, 146)
        Me.btnExportDir.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExportDir.Name = "btnExportDir"
        Me.btnExportDir.Size = New System.Drawing.Size(27, 19)
        Me.btnExportDir.TabIndex = 22
        Me.btnExportDir.Text = "..."
        Me.btnExportDir.UseVisualStyleBackColor = True
        '
        'txtExportDir
        '
        Me.txtExportDir.Location = New System.Drawing.Point(177, 148)
        Me.txtExportDir.Margin = New System.Windows.Forms.Padding(2)
        Me.txtExportDir.Name = "txtExportDir"
        Me.txtExportDir.Size = New System.Drawing.Size(262, 20)
        Me.txtExportDir.TabIndex = 21
        '
        'lblExportDir
        '
        Me.lblExportDir.AutoSize = True
        Me.lblExportDir.Location = New System.Drawing.Point(11, 148)
        Me.lblExportDir.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblExportDir.Name = "lblExportDir"
        Me.lblExportDir.Size = New System.Drawing.Size(82, 13)
        Me.lblExportDir.TabIndex = 59
        Me.lblExportDir.Text = "Export Directory"
        '
        'btnWorkerTempDir
        '
        Me.btnWorkerTempDir.Location = New System.Drawing.Point(443, 122)
        Me.btnWorkerTempDir.Margin = New System.Windows.Forms.Padding(2)
        Me.btnWorkerTempDir.Name = "btnWorkerTempDir"
        Me.btnWorkerTempDir.Size = New System.Drawing.Size(27, 19)
        Me.btnWorkerTempDir.TabIndex = 20
        Me.btnWorkerTempDir.Text = "..."
        Me.btnWorkerTempDir.UseVisualStyleBackColor = True
        '
        'txtWorkerTempDir
        '
        Me.txtWorkerTempDir.Location = New System.Drawing.Point(177, 122)
        Me.txtWorkerTempDir.Margin = New System.Windows.Forms.Padding(2)
        Me.txtWorkerTempDir.Name = "txtWorkerTempDir"
        Me.txtWorkerTempDir.Size = New System.Drawing.Size(262, 20)
        Me.txtWorkerTempDir.TabIndex = 19
        '
        'lblWorkerTempDir
        '
        Me.lblWorkerTempDir.AutoSize = True
        Me.lblWorkerTempDir.Location = New System.Drawing.Point(11, 122)
        Me.lblWorkerTempDir.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblWorkerTempDir.Name = "lblWorkerTempDir"
        Me.lblWorkerTempDir.Size = New System.Drawing.Size(91, 13)
        Me.lblWorkerTempDir.TabIndex = 56
        Me.lblWorkerTempDir.Text = "Worker Temp Dir:"
        '
        'btnNuixAppLocation
        '
        Me.btnNuixAppLocation.Location = New System.Drawing.Point(443, 172)
        Me.btnNuixAppLocation.Margin = New System.Windows.Forms.Padding(2)
        Me.btnNuixAppLocation.Name = "btnNuixAppLocation"
        Me.btnNuixAppLocation.Size = New System.Drawing.Size(27, 19)
        Me.btnNuixAppLocation.TabIndex = 24
        Me.btnNuixAppLocation.Text = "..."
        Me.btnNuixAppLocation.UseVisualStyleBackColor = True
        '
        'txtNuixAppLocation
        '
        Me.txtNuixAppLocation.Location = New System.Drawing.Point(177, 174)
        Me.txtNuixAppLocation.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNuixAppLocation.Name = "txtNuixAppLocation"
        Me.txtNuixAppLocation.Size = New System.Drawing.Size(262, 20)
        Me.txtNuixAppLocation.TabIndex = 23
        '
        'lblNuixAppLocation
        '
        Me.lblNuixAppLocation.AutoSize = True
        Me.lblNuixAppLocation.Location = New System.Drawing.Point(11, 174)
        Me.lblNuixAppLocation.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNuixAppLocation.Name = "lblNuixAppLocation"
        Me.lblNuixAppLocation.Size = New System.Drawing.Size(97, 13)
        Me.lblNuixAppLocation.TabIndex = 55
        Me.lblNuixAppLocation.Text = "Nuix App Location:"
        '
        'btnJavaTempDir
        '
        Me.btnJavaTempDir.Location = New System.Drawing.Point(443, 95)
        Me.btnJavaTempDir.Margin = New System.Windows.Forms.Padding(2)
        Me.btnJavaTempDir.Name = "btnJavaTempDir"
        Me.btnJavaTempDir.Size = New System.Drawing.Size(27, 19)
        Me.btnJavaTempDir.TabIndex = 18
        Me.btnJavaTempDir.Text = "..."
        Me.btnJavaTempDir.UseVisualStyleBackColor = True
        '
        'btnLogDir
        '
        Me.btnLogDir.Location = New System.Drawing.Point(443, 69)
        Me.btnLogDir.Margin = New System.Windows.Forms.Padding(2)
        Me.btnLogDir.Name = "btnLogDir"
        Me.btnLogDir.Size = New System.Drawing.Size(27, 19)
        Me.btnLogDir.TabIndex = 16
        Me.btnLogDir.Text = "..."
        Me.btnLogDir.UseVisualStyleBackColor = True
        '
        'btnCaseDirSel
        '
        Me.btnCaseDirSel.Location = New System.Drawing.Point(443, 19)
        Me.btnCaseDirSel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCaseDirSel.Name = "btnCaseDirSel"
        Me.btnCaseDirSel.Size = New System.Drawing.Size(27, 19)
        Me.btnCaseDirSel.TabIndex = 14
        Me.btnCaseDirSel.Text = "..."
        Me.btnCaseDirSel.UseVisualStyleBackColor = True
        '
        'txtNuixCaseDir
        '
        Me.txtNuixCaseDir.Location = New System.Drawing.Point(178, 21)
        Me.txtNuixCaseDir.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNuixCaseDir.Name = "txtNuixCaseDir"
        Me.txtNuixCaseDir.Size = New System.Drawing.Size(261, 20)
        Me.txtNuixCaseDir.TabIndex = 13
        '
        'lblCaseDirectory
        '
        Me.lblCaseDirectory.AutoSize = True
        Me.lblCaseDirectory.Location = New System.Drawing.Point(11, 21)
        Me.lblCaseDirectory.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblCaseDirectory.Name = "lblCaseDirectory"
        Me.lblCaseDirectory.Size = New System.Drawing.Size(79, 13)
        Me.lblCaseDirectory.TabIndex = 50
        Me.lblCaseDirectory.Text = "Case Directory:"
        '
        'txtJavaTempDir
        '
        Me.txtJavaTempDir.Location = New System.Drawing.Point(177, 96)
        Me.txtJavaTempDir.Margin = New System.Windows.Forms.Padding(2)
        Me.txtJavaTempDir.Name = "txtJavaTempDir"
        Me.txtJavaTempDir.Size = New System.Drawing.Size(262, 20)
        Me.txtJavaTempDir.TabIndex = 17
        '
        'lblJavaTempDir
        '
        Me.lblJavaTempDir.AutoSize = True
        Me.lblJavaTempDir.Location = New System.Drawing.Point(11, 96)
        Me.lblJavaTempDir.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblJavaTempDir.Name = "lblJavaTempDir"
        Me.lblJavaTempDir.Size = New System.Drawing.Size(108, 13)
        Me.lblJavaTempDir.TabIndex = 42
        Me.lblJavaTempDir.Text = "Java Temp Directory:"
        '
        'lblLogDirectory
        '
        Me.lblLogDirectory.AutoSize = True
        Me.lblLogDirectory.Location = New System.Drawing.Point(11, 70)
        Me.lblLogDirectory.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblLogDirectory.Name = "lblLogDirectory"
        Me.lblLogDirectory.Size = New System.Drawing.Size(73, 13)
        Me.lblLogDirectory.TabIndex = 41
        Me.lblLogDirectory.Text = "Log Directory:"
        '
        'txtLogDirectory
        '
        Me.txtLogDirectory.Location = New System.Drawing.Point(177, 70)
        Me.txtLogDirectory.Margin = New System.Windows.Forms.Padding(2)
        Me.txtLogDirectory.Name = "txtLogDirectory"
        Me.txtLogDirectory.Size = New System.Drawing.Size(262, 20)
        Me.txtLogDirectory.TabIndex = 15
        '
        'tabNuixDataSettings
        '
        Me.tabNuixDataSettings.Controls.Add(Me.grpDataProcessingSettings)
        Me.tabNuixDataSettings.Location = New System.Drawing.Point(4, 22)
        Me.tabNuixDataSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.tabNuixDataSettings.Name = "tabNuixDataSettings"
        Me.tabNuixDataSettings.Size = New System.Drawing.Size(692, 422)
        Me.tabNuixDataSettings.TabIndex = 2
        Me.tabNuixDataSettings.Text = "Nuix Data Processings"
        Me.tabNuixDataSettings.UseVisualStyleBackColor = True
        '
        'grpDataProcessingSettings
        '
        Me.grpDataProcessingSettings.Controls.Add(Me.grpDigestSettings)
        Me.grpDataProcessingSettings.Controls.Add(Me.grpImageSettings)
        Me.grpDataProcessingSettings.Controls.Add(Me.grpItemContentsSettings)
        Me.grpDataProcessingSettings.Controls.Add(Me.grpTextIndexingSettings)
        Me.grpDataProcessingSettings.Controls.Add(Me.grpFamilyTextSettings)
        Me.grpDataProcessingSettings.Controls.Add(Me.grpDeletefileSettings)
        Me.grpDataProcessingSettings.Controls.Add(Me.grpEvidenceSettings)
        Me.grpDataProcessingSettings.Controls.Add(Me.lblTraversal)
        Me.grpDataProcessingSettings.Controls.Add(Me.cboTraversal)
        Me.grpDataProcessingSettings.Controls.Add(Me.chkCalculateProcessingSize)
        Me.grpDataProcessingSettings.Controls.Add(Me.chkPerformItemIdentification)
        Me.grpDataProcessingSettings.Location = New System.Drawing.Point(2, 0)
        Me.grpDataProcessingSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpDataProcessingSettings.Name = "grpDataProcessingSettings"
        Me.grpDataProcessingSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpDataProcessingSettings.Size = New System.Drawing.Size(691, 425)
        Me.grpDataProcessingSettings.TabIndex = 0
        Me.grpDataProcessingSettings.TabStop = False
        Me.grpDataProcessingSettings.Text = "Data Processing Settings"
        '
        'grpDigestSettings
        '
        Me.grpDigestSettings.Controls.Add(Me.lblDigestToCompute)
        Me.grpDigestSettings.Controls.Add(Me.grpEmailDigestSettings)
        Me.grpDigestSettings.Controls.Add(Me.Label4)
        Me.grpDigestSettings.Controls.Add(Me.MaxDigestSize)
        Me.grpDigestSettings.Controls.Add(Me.chkSSDeep)
        Me.grpDigestSettings.Controls.Add(Me.chkSHA256)
        Me.grpDigestSettings.Controls.Add(Me.chkSHA1)
        Me.grpDigestSettings.Controls.Add(Me.chkMD5Digest)
        Me.grpDigestSettings.Location = New System.Drawing.Point(296, 281)
        Me.grpDigestSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpDigestSettings.Name = "grpDigestSettings"
        Me.grpDigestSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpDigestSettings.Size = New System.Drawing.Size(261, 132)
        Me.grpDigestSettings.TabIndex = 10
        Me.grpDigestSettings.TabStop = False
        Me.grpDigestSettings.Text = "Digest Settings"
        '
        'lblDigestToCompute
        '
        Me.lblDigestToCompute.AutoSize = True
        Me.lblDigestToCompute.Location = New System.Drawing.Point(3, 15)
        Me.lblDigestToCompute.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblDigestToCompute.Name = "lblDigestToCompute"
        Me.lblDigestToCompute.Size = New System.Drawing.Size(96, 13)
        Me.lblDigestToCompute.TabIndex = 8
        Me.lblDigestToCompute.Text = "Digest to compute:"
        '
        'grpEmailDigestSettings
        '
        Me.grpEmailDigestSettings.Controls.Add(Me.chkIncludeItemDate)
        Me.grpEmailDigestSettings.Controls.Add(Me.chkIncludeBCC)
        Me.grpEmailDigestSettings.Location = New System.Drawing.Point(116, 42)
        Me.grpEmailDigestSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpEmailDigestSettings.Name = "grpEmailDigestSettings"
        Me.grpEmailDigestSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpEmailDigestSettings.Size = New System.Drawing.Size(141, 62)
        Me.grpEmailDigestSettings.TabIndex = 7
        Me.grpEmailDigestSettings.TabStop = False
        Me.grpEmailDigestSettings.Text = "Email Digest Settings"
        '
        'chkIncludeItemDate
        '
        Me.chkIncludeItemDate.AutoSize = True
        Me.chkIncludeItemDate.Location = New System.Drawing.Point(4, 40)
        Me.chkIncludeItemDate.Margin = New System.Windows.Forms.Padding(2)
        Me.chkIncludeItemDate.Name = "chkIncludeItemDate"
        Me.chkIncludeItemDate.Size = New System.Drawing.Size(110, 17)
        Me.chkIncludeItemDate.TabIndex = 1
        Me.chkIncludeItemDate.Text = "Include Item Date"
        Me.chkIncludeItemDate.UseVisualStyleBackColor = True
        '
        'chkIncludeBCC
        '
        Me.chkIncludeBCC.AutoSize = True
        Me.chkIncludeBCC.Location = New System.Drawing.Point(4, 19)
        Me.chkIncludeBCC.Margin = New System.Windows.Forms.Padding(2)
        Me.chkIncludeBCC.Name = "chkIncludeBCC"
        Me.chkIncludeBCC.Size = New System.Drawing.Size(83, 17)
        Me.chkIncludeBCC.TabIndex = 0
        Me.chkIncludeBCC.Text = "Include Bcc"
        Me.chkIncludeBCC.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(2, 111)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(109, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Maximize Digest Size:"
        '
        'MaxDigestSize
        '
        Me.MaxDigestSize.Location = New System.Drawing.Point(115, 110)
        Me.MaxDigestSize.Margin = New System.Windows.Forms.Padding(2)
        Me.MaxDigestSize.Maximum = New Decimal(New Integer() {256, 0, 0, 0})
        Me.MaxDigestSize.Minimum = New Decimal(New Integer() {256, 0, 0, 0})
        Me.MaxDigestSize.Name = "MaxDigestSize"
        Me.MaxDigestSize.Size = New System.Drawing.Size(53, 20)
        Me.MaxDigestSize.TabIndex = 5
        Me.MaxDigestSize.Value = New Decimal(New Integer() {256, 0, 0, 0})
        '
        'chkSSDeep
        '
        Me.chkSSDeep.AutoSize = True
        Me.chkSSDeep.Location = New System.Drawing.Point(4, 89)
        Me.chkSSDeep.Margin = New System.Windows.Forms.Padding(2)
        Me.chkSSDeep.Name = "chkSSDeep"
        Me.chkSSDeep.Size = New System.Drawing.Size(66, 17)
        Me.chkSSDeep.TabIndex = 3
        Me.chkSSDeep.Text = "SSDeep"
        Me.chkSSDeep.UseVisualStyleBackColor = True
        '
        'chkSHA256
        '
        Me.chkSHA256.AutoSize = True
        Me.chkSHA256.Location = New System.Drawing.Point(4, 70)
        Me.chkSHA256.Margin = New System.Windows.Forms.Padding(2)
        Me.chkSHA256.Name = "chkSHA256"
        Me.chkSHA256.Size = New System.Drawing.Size(69, 17)
        Me.chkSHA256.TabIndex = 2
        Me.chkSHA256.Text = "SHA-256"
        Me.chkSHA256.UseVisualStyleBackColor = True
        '
        'chkSHA1
        '
        Me.chkSHA1.AutoSize = True
        Me.chkSHA1.Location = New System.Drawing.Point(4, 49)
        Me.chkSHA1.Margin = New System.Windows.Forms.Padding(2)
        Me.chkSHA1.Name = "chkSHA1"
        Me.chkSHA1.Size = New System.Drawing.Size(57, 17)
        Me.chkSHA1.TabIndex = 1
        Me.chkSHA1.Text = "SHA-1"
        Me.chkSHA1.UseVisualStyleBackColor = True
        '
        'chkMD5Digest
        '
        Me.chkMD5Digest.AutoSize = True
        Me.chkMD5Digest.Location = New System.Drawing.Point(4, 30)
        Me.chkMD5Digest.Margin = New System.Windows.Forms.Padding(2)
        Me.chkMD5Digest.Name = "chkMD5Digest"
        Me.chkMD5Digest.Size = New System.Drawing.Size(49, 17)
        Me.chkMD5Digest.TabIndex = 0
        Me.chkMD5Digest.Text = "MD5"
        Me.chkMD5Digest.UseVisualStyleBackColor = True
        '
        'grpImageSettings
        '
        Me.grpImageSettings.Controls.Add(Me.chkPerformimagecolouranalysis)
        Me.grpImageSettings.Controls.Add(Me.chkGeneratethumbnailsforimagedata)
        Me.grpImageSettings.Location = New System.Drawing.Point(296, 222)
        Me.grpImageSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpImageSettings.Name = "grpImageSettings"
        Me.grpImageSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpImageSettings.Size = New System.Drawing.Size(261, 54)
        Me.grpImageSettings.TabIndex = 9
        Me.grpImageSettings.TabStop = False
        Me.grpImageSettings.Text = "Image Settings"
        '
        'chkPerformimagecolouranalysis
        '
        Me.chkPerformimagecolouranalysis.AutoSize = True
        Me.chkPerformimagecolouranalysis.Location = New System.Drawing.Point(5, 34)
        Me.chkPerformimagecolouranalysis.Margin = New System.Windows.Forms.Padding(2)
        Me.chkPerformimagecolouranalysis.Name = "chkPerformimagecolouranalysis"
        Me.chkPerformimagecolouranalysis.Size = New System.Drawing.Size(232, 17)
        Me.chkPerformimagecolouranalysis.TabIndex = 1
        Me.chkPerformimagecolouranalysis.Text = "Perform image colour and skin-tone analysis"
        Me.chkPerformimagecolouranalysis.UseVisualStyleBackColor = True
        '
        'chkGeneratethumbnailsforimagedata
        '
        Me.chkGeneratethumbnailsforimagedata.AutoSize = True
        Me.chkGeneratethumbnailsforimagedata.Location = New System.Drawing.Point(5, 17)
        Me.chkGeneratethumbnailsforimagedata.Margin = New System.Windows.Forms.Padding(2)
        Me.chkGeneratethumbnailsforimagedata.Name = "chkGeneratethumbnailsforimagedata"
        Me.chkGeneratethumbnailsforimagedata.Size = New System.Drawing.Size(193, 17)
        Me.chkGeneratethumbnailsforimagedata.TabIndex = 0
        Me.chkGeneratethumbnailsforimagedata.Text = "Generate thumbnails for image data"
        Me.chkGeneratethumbnailsforimagedata.UseVisualStyleBackColor = True
        '
        'grpItemContentsSettings
        '
        Me.grpItemContentsSettings.Controls.Add(Me.grpNamedEntitySettings)
        Me.grpItemContentsSettings.Controls.Add(Me.chkEnableTextSummarisation)
        Me.grpItemContentsSettings.Controls.Add(Me.chkEnableNearDuplicates)
        Me.grpItemContentsSettings.Controls.Add(Me.chkProcesstext)
        Me.grpItemContentsSettings.Location = New System.Drawing.Point(297, 58)
        Me.grpItemContentsSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpItemContentsSettings.Name = "grpItemContentsSettings"
        Me.grpItemContentsSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpItemContentsSettings.Size = New System.Drawing.Size(266, 161)
        Me.grpItemContentsSettings.TabIndex = 8
        Me.grpItemContentsSettings.TabStop = False
        Me.grpItemContentsSettings.Text = "Item Contents Settings"
        '
        'grpNamedEntitySettings
        '
        Me.grpNamedEntitySettings.Controls.Add(Me.chkExtractNamedEntitiesfromProperties)
        Me.grpNamedEntitySettings.Controls.Add(Me.chkIncludeTextStrippedItems)
        Me.grpNamedEntitySettings.Controls.Add(Me.chkExtractNamedEntities)
        Me.grpNamedEntitySettings.Location = New System.Drawing.Point(5, 83)
        Me.grpNamedEntitySettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpNamedEntitySettings.Name = "grpNamedEntitySettings"
        Me.grpNamedEntitySettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpNamedEntitySettings.Size = New System.Drawing.Size(257, 73)
        Me.grpNamedEntitySettings.TabIndex = 3
        Me.grpNamedEntitySettings.TabStop = False
        Me.grpNamedEntitySettings.Text = "Named Entitiy Settings"
        '
        'chkExtractNamedEntitiesfromProperties
        '
        Me.chkExtractNamedEntitiesfromProperties.AutoSize = True
        Me.chkExtractNamedEntitiesfromProperties.Location = New System.Drawing.Point(4, 55)
        Me.chkExtractNamedEntitiesfromProperties.Margin = New System.Windows.Forms.Padding(2)
        Me.chkExtractNamedEntitiesfromProperties.Name = "chkExtractNamedEntitiesfromProperties"
        Me.chkExtractNamedEntitiesfromProperties.Size = New System.Drawing.Size(202, 17)
        Me.chkExtractNamedEntitiesfromProperties.TabIndex = 2
        Me.chkExtractNamedEntitiesfromProperties.Text = "Extract named entities from properties"
        Me.chkExtractNamedEntitiesfromProperties.UseVisualStyleBackColor = True
        '
        'chkIncludeTextStrippedItems
        '
        Me.chkIncludeTextStrippedItems.AutoSize = True
        Me.chkIncludeTextStrippedItems.Location = New System.Drawing.Point(23, 36)
        Me.chkIncludeTextStrippedItems.Margin = New System.Windows.Forms.Padding(2)
        Me.chkIncludeTextStrippedItems.Name = "chkIncludeTextStrippedItems"
        Me.chkIncludeTextStrippedItems.Size = New System.Drawing.Size(148, 17)
        Me.chkIncludeTextStrippedItems.TabIndex = 1
        Me.chkIncludeTextStrippedItems.Text = "Include text stripped items"
        Me.chkIncludeTextStrippedItems.UseVisualStyleBackColor = True
        '
        'chkExtractNamedEntities
        '
        Me.chkExtractNamedEntities.AutoSize = True
        Me.chkExtractNamedEntities.Location = New System.Drawing.Point(4, 16)
        Me.chkExtractNamedEntities.Margin = New System.Windows.Forms.Padding(2)
        Me.chkExtractNamedEntities.Name = "chkExtractNamedEntities"
        Me.chkExtractNamedEntities.Size = New System.Drawing.Size(173, 17)
        Me.chkExtractNamedEntities.TabIndex = 0
        Me.chkExtractNamedEntities.Text = "Extract named entities from text"
        Me.chkExtractNamedEntities.UseVisualStyleBackColor = True
        '
        'chkEnableTextSummarisation
        '
        Me.chkEnableTextSummarisation.AutoSize = True
        Me.chkEnableTextSummarisation.Location = New System.Drawing.Point(4, 56)
        Me.chkEnableTextSummarisation.Margin = New System.Windows.Forms.Padding(2)
        Me.chkEnableTextSummarisation.Name = "chkEnableTextSummarisation"
        Me.chkEnableTextSummarisation.Size = New System.Drawing.Size(148, 17)
        Me.chkEnableTextSummarisation.TabIndex = 2
        Me.chkEnableTextSummarisation.Text = "Enable text summarisation"
        Me.chkEnableTextSummarisation.UseVisualStyleBackColor = True
        '
        'chkEnableNearDuplicates
        '
        Me.chkEnableNearDuplicates.AutoSize = True
        Me.chkEnableNearDuplicates.Location = New System.Drawing.Point(4, 38)
        Me.chkEnableNearDuplicates.Margin = New System.Windows.Forms.Padding(2)
        Me.chkEnableNearDuplicates.Name = "chkEnableNearDuplicates"
        Me.chkEnableNearDuplicates.Size = New System.Drawing.Size(134, 17)
        Me.chkEnableNearDuplicates.TabIndex = 1
        Me.chkEnableNearDuplicates.Text = "Enable near-duplicates"
        Me.chkEnableNearDuplicates.UseVisualStyleBackColor = True
        '
        'chkProcesstext
        '
        Me.chkProcesstext.AutoSize = True
        Me.chkProcesstext.Location = New System.Drawing.Point(4, 19)
        Me.chkProcesstext.Margin = New System.Windows.Forms.Padding(2)
        Me.chkProcesstext.Name = "chkProcesstext"
        Me.chkProcesstext.Size = New System.Drawing.Size(84, 17)
        Me.chkProcesstext.TabIndex = 0
        Me.chkProcesstext.Text = "Process text"
        Me.chkProcesstext.UseVisualStyleBackColor = True
        '
        'grpTextIndexingSettings
        '
        Me.grpTextIndexingSettings.Controls.Add(Me.chkEnableExactQueries)
        Me.grpTextIndexingSettings.Controls.Add(Me.chkUsestemming)
        Me.grpTextIndexingSettings.Controls.Add(Me.chkUseStopWords)
        Me.grpTextIndexingSettings.Controls.Add(Me.Label3)
        Me.grpTextIndexingSettings.Controls.Add(Me.cboAnalysisLang)
        Me.grpTextIndexingSettings.Location = New System.Drawing.Point(11, 323)
        Me.grpTextIndexingSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpTextIndexingSettings.Name = "grpTextIndexingSettings"
        Me.grpTextIndexingSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpTextIndexingSettings.Size = New System.Drawing.Size(265, 90)
        Me.grpTextIndexingSettings.TabIndex = 7
        Me.grpTextIndexingSettings.TabStop = False
        Me.grpTextIndexingSettings.Text = "Text Indexing Settings"
        '
        'chkEnableExactQueries
        '
        Me.chkEnableExactQueries.AutoSize = True
        Me.chkEnableExactQueries.Location = New System.Drawing.Point(7, 62)
        Me.chkEnableExactQueries.Margin = New System.Windows.Forms.Padding(2)
        Me.chkEnableExactQueries.Name = "chkEnableExactQueries"
        Me.chkEnableExactQueries.Size = New System.Drawing.Size(128, 17)
        Me.chkEnableExactQueries.TabIndex = 10
        Me.chkEnableExactQueries.Text = "Enable Exact Queries"
        Me.chkEnableExactQueries.UseVisualStyleBackColor = True
        '
        'chkUsestemming
        '
        Me.chkUsestemming.AutoSize = True
        Me.chkUsestemming.Location = New System.Drawing.Point(120, 40)
        Me.chkUsestemming.Margin = New System.Windows.Forms.Padding(2)
        Me.chkUsestemming.Name = "chkUsestemming"
        Me.chkUsestemming.Size = New System.Drawing.Size(92, 17)
        Me.chkUsestemming.TabIndex = 9
        Me.chkUsestemming.Text = "Use stemming"
        Me.chkUsestemming.UseVisualStyleBackColor = True
        '
        'chkUseStopWords
        '
        Me.chkUseStopWords.AutoSize = True
        Me.chkUseStopWords.Location = New System.Drawing.Point(7, 40)
        Me.chkUseStopWords.Margin = New System.Windows.Forms.Padding(2)
        Me.chkUseStopWords.Name = "chkUseStopWords"
        Me.chkUseStopWords.Size = New System.Drawing.Size(99, 17)
        Me.chkUseStopWords.TabIndex = 8
        Me.chkUseStopWords.Text = "Use stop words"
        Me.chkUseStopWords.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(5, 18)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(99, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Analysis Language:"
        '
        'cboAnalysisLang
        '
        Me.cboAnalysisLang.FormattingEnabled = True
        Me.cboAnalysisLang.Items.AddRange(New Object() {"English"})
        Me.cboAnalysisLang.Location = New System.Drawing.Point(120, 16)
        Me.cboAnalysisLang.Margin = New System.Windows.Forms.Padding(2)
        Me.cboAnalysisLang.Name = "cboAnalysisLang"
        Me.cboAnalysisLang.Size = New System.Drawing.Size(82, 21)
        Me.cboAnalysisLang.TabIndex = 8
        '
        'grpFamilyTextSettings
        '
        Me.grpFamilyTextSettings.Controls.Add(Me.chkHideimmaterialItems)
        Me.grpFamilyTextSettings.Controls.Add(Me.chkCreatefamilySearch)
        Me.grpFamilyTextSettings.Location = New System.Drawing.Point(11, 263)
        Me.grpFamilyTextSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpFamilyTextSettings.Name = "grpFamilyTextSettings"
        Me.grpFamilyTextSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpFamilyTextSettings.Size = New System.Drawing.Size(265, 57)
        Me.grpFamilyTextSettings.TabIndex = 6
        Me.grpFamilyTextSettings.TabStop = False
        Me.grpFamilyTextSettings.Text = "Family Text Settings"
        '
        'chkHideimmaterialItems
        '
        Me.chkHideimmaterialItems.AutoSize = True
        Me.chkHideimmaterialItems.Location = New System.Drawing.Point(4, 37)
        Me.chkHideimmaterialItems.Margin = New System.Windows.Forms.Padding(2)
        Me.chkHideimmaterialItems.Name = "chkHideimmaterialItems"
        Me.chkHideimmaterialItems.Size = New System.Drawing.Size(238, 17)
        Me.chkHideimmaterialItems.TabIndex = 8
        Me.chkHideimmaterialItems.Text = "Hide immaterial items (text rolled up to parent)"
        Me.chkHideimmaterialItems.UseVisualStyleBackColor = True
        '
        'chkCreatefamilySearch
        '
        Me.chkCreatefamilySearch.AutoSize = True
        Me.chkCreatefamilySearch.Location = New System.Drawing.Point(4, 18)
        Me.chkCreatefamilySearch.Margin = New System.Windows.Forms.Padding(2)
        Me.chkCreatefamilySearch.Name = "chkCreatefamilySearch"
        Me.chkCreatefamilySearch.Size = New System.Drawing.Size(233, 17)
        Me.chkCreatefamilySearch.TabIndex = 7
        Me.chkCreatefamilySearch.Text = "Create family search fields for top level items"
        Me.chkCreatefamilySearch.UseVisualStyleBackColor = True
        '
        'grpDeletefileSettings
        '
        Me.grpDeletefileSettings.Controls.Add(Me.chkCarveFSunallocated)
        Me.grpDeletefileSettings.Controls.Add(Me.chkExtractmailboxslackspace)
        Me.grpDeletefileSettings.Controls.Add(Me.chkSmartprocessMSRegistry)
        Me.grpDeletefileSettings.Controls.Add(Me.chkExtractEndOfFileSpace)
        Me.grpDeletefileSettings.Controls.Add(Me.chkRecoverDeleteFiles)
        Me.grpDeletefileSettings.Location = New System.Drawing.Point(11, 151)
        Me.grpDeletefileSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpDeletefileSettings.Name = "grpDeletefileSettings"
        Me.grpDeletefileSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpDeletefileSettings.Size = New System.Drawing.Size(265, 109)
        Me.grpDeletefileSettings.TabIndex = 5
        Me.grpDeletefileSettings.TabStop = False
        Me.grpDeletefileSettings.Text = "Delete File Recovery/Forensic Settings"
        '
        'chkCarveFSunallocated
        '
        Me.chkCarveFSunallocated.AutoSize = True
        Me.chkCarveFSunallocated.Location = New System.Drawing.Point(4, 89)
        Me.chkCarveFSunallocated.Margin = New System.Windows.Forms.Padding(2)
        Me.chkCarveFSunallocated.Name = "chkCarveFSunallocated"
        Me.chkCarveFSunallocated.Size = New System.Drawing.Size(195, 17)
        Me.chkCarveFSunallocated.TabIndex = 10
        Me.chkCarveFSunallocated.Text = "Carve file system unallocated space"
        Me.chkCarveFSunallocated.UseVisualStyleBackColor = True
        '
        'chkExtractmailboxslackspace
        '
        Me.chkExtractmailboxslackspace.AutoSize = True
        Me.chkExtractmailboxslackspace.Location = New System.Drawing.Point(4, 71)
        Me.chkExtractmailboxslackspace.Margin = New System.Windows.Forms.Padding(2)
        Me.chkExtractmailboxslackspace.Name = "chkExtractmailboxslackspace"
        Me.chkExtractmailboxslackspace.Size = New System.Drawing.Size(180, 17)
        Me.chkExtractmailboxslackspace.TabIndex = 9
        Me.chkExtractmailboxslackspace.Text = "Extract from mailbox slack space"
        Me.chkExtractmailboxslackspace.UseVisualStyleBackColor = True
        '
        'chkSmartprocessMSRegistry
        '
        Me.chkSmartprocessMSRegistry.AutoSize = True
        Me.chkSmartprocessMSRegistry.Location = New System.Drawing.Point(4, 53)
        Me.chkSmartprocessMSRegistry.Margin = New System.Windows.Forms.Padding(2)
        Me.chkSmartprocessMSRegistry.Name = "chkSmartprocessMSRegistry"
        Me.chkSmartprocessMSRegistry.Size = New System.Drawing.Size(201, 17)
        Me.chkSmartprocessMSRegistry.TabIndex = 8
        Me.chkSmartprocessMSRegistry.Text = "Smart process Microsoft Registry files"
        Me.chkSmartprocessMSRegistry.UseVisualStyleBackColor = True
        '
        'chkExtractEndOfFileSpace
        '
        Me.chkExtractEndOfFileSpace.AutoSize = True
        Me.chkExtractEndOfFileSpace.Location = New System.Drawing.Point(4, 34)
        Me.chkExtractEndOfFileSpace.Margin = New System.Windows.Forms.Padding(2)
        Me.chkExtractEndOfFileSpace.Name = "chkExtractEndOfFileSpace"
        Me.chkExtractEndOfFileSpace.Size = New System.Drawing.Size(249, 17)
        Me.chkExtractEndOfFileSpace.TabIndex = 7
        Me.chkExtractEndOfFileSpace.Text = "Extract end-of-file slack space from disk images"
        Me.chkExtractEndOfFileSpace.UseVisualStyleBackColor = True
        '
        'chkRecoverDeleteFiles
        '
        Me.chkRecoverDeleteFiles.AutoSize = True
        Me.chkRecoverDeleteFiles.Location = New System.Drawing.Point(4, 14)
        Me.chkRecoverDeleteFiles.Margin = New System.Windows.Forms.Padding(2)
        Me.chkRecoverDeleteFiles.Name = "chkRecoverDeleteFiles"
        Me.chkRecoverDeleteFiles.Size = New System.Drawing.Size(207, 17)
        Me.chkRecoverDeleteFiles.TabIndex = 6
        Me.chkRecoverDeleteFiles.Text = "Recover deleted files from disk images"
        Me.chkRecoverDeleteFiles.UseVisualStyleBackColor = True
        '
        'grpEvidenceSettings
        '
        Me.grpEvidenceSettings.Controls.Add(Me.Maxbinarysize)
        Me.grpEvidenceSettings.Controls.Add(Me.lblMaxbinarysize)
        Me.grpEvidenceSettings.Controls.Add(Me.chkStorebinary)
        Me.grpEvidenceSettings.Controls.Add(Me.chkCalculateAuditedSize)
        Me.grpEvidenceSettings.Controls.Add(Me.chkReuseEvidenceStores)
        Me.grpEvidenceSettings.Location = New System.Drawing.Point(11, 58)
        Me.grpEvidenceSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpEvidenceSettings.Name = "grpEvidenceSettings"
        Me.grpEvidenceSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpEvidenceSettings.Size = New System.Drawing.Size(265, 90)
        Me.grpEvidenceSettings.TabIndex = 4
        Me.grpEvidenceSettings.TabStop = False
        Me.grpEvidenceSettings.Text = "Evidence Settings"
        '
        'Maxbinarysize
        '
        Me.Maxbinarysize.Location = New System.Drawing.Point(127, 65)
        Me.Maxbinarysize.Margin = New System.Windows.Forms.Padding(2)
        Me.Maxbinarysize.Maximum = New Decimal(New Integer() {256, 0, 0, 0})
        Me.Maxbinarysize.Minimum = New Decimal(New Integer() {256, 0, 0, 0})
        Me.Maxbinarysize.Name = "Maxbinarysize"
        Me.Maxbinarysize.Size = New System.Drawing.Size(80, 20)
        Me.Maxbinarysize.TabIndex = 5
        Me.Maxbinarysize.Value = New Decimal(New Integer() {256, 0, 0, 0})
        '
        'lblMaxbinarysize
        '
        Me.lblMaxbinarysize.AutoSize = True
        Me.lblMaxbinarysize.Location = New System.Drawing.Point(20, 66)
        Me.lblMaxbinarysize.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblMaxbinarysize.Name = "lblMaxbinarysize"
        Me.lblMaxbinarysize.Size = New System.Drawing.Size(103, 13)
        Me.lblMaxbinarysize.TabIndex = 5
        Me.lblMaxbinarysize.Text = "Maximum binary size"
        '
        'chkStorebinary
        '
        Me.chkStorebinary.AutoSize = True
        Me.chkStorebinary.Location = New System.Drawing.Point(4, 49)
        Me.chkStorebinary.Margin = New System.Windows.Forms.Padding(2)
        Me.chkStorebinary.Name = "chkStorebinary"
        Me.chkStorebinary.Size = New System.Drawing.Size(145, 17)
        Me.chkStorebinary.TabIndex = 5
        Me.chkStorebinary.Text = "Store binary of data items"
        Me.chkStorebinary.UseVisualStyleBackColor = True
        '
        'chkCalculateAuditedSize
        '
        Me.chkCalculateAuditedSize.AutoSize = True
        Me.chkCalculateAuditedSize.Location = New System.Drawing.Point(4, 32)
        Me.chkCalculateAuditedSize.Margin = New System.Windows.Forms.Padding(2)
        Me.chkCalculateAuditedSize.Name = "chkCalculateAuditedSize"
        Me.chkCalculateAuditedSize.Size = New System.Drawing.Size(132, 17)
        Me.chkCalculateAuditedSize.TabIndex = 5
        Me.chkCalculateAuditedSize.Text = "Calculate Audited Size"
        Me.chkCalculateAuditedSize.UseVisualStyleBackColor = True
        '
        'chkReuseEvidenceStores
        '
        Me.chkReuseEvidenceStores.AutoSize = True
        Me.chkReuseEvidenceStores.Location = New System.Drawing.Point(4, 16)
        Me.chkReuseEvidenceStores.Margin = New System.Windows.Forms.Padding(2)
        Me.chkReuseEvidenceStores.Name = "chkReuseEvidenceStores"
        Me.chkReuseEvidenceStores.Size = New System.Drawing.Size(138, 17)
        Me.chkReuseEvidenceStores.TabIndex = 5
        Me.chkReuseEvidenceStores.Text = "Reuse Evidence Stores"
        Me.chkReuseEvidenceStores.UseVisualStyleBackColor = True
        '
        'lblTraversal
        '
        Me.lblTraversal.AutoSize = True
        Me.lblTraversal.Location = New System.Drawing.Point(8, 38)
        Me.lblTraversal.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTraversal.Name = "lblTraversal"
        Me.lblTraversal.Size = New System.Drawing.Size(54, 13)
        Me.lblTraversal.TabIndex = 3
        Me.lblTraversal.Text = "Traversal:"
        '
        'cboTraversal
        '
        Me.cboTraversal.FormattingEnabled = True
        Me.cboTraversal.Items.AddRange(New Object() {"full_traversal", "limited_traversal"})
        Me.cboTraversal.Location = New System.Drawing.Point(65, 36)
        Me.cboTraversal.Margin = New System.Windows.Forms.Padding(2)
        Me.cboTraversal.Name = "cboTraversal"
        Me.cboTraversal.Size = New System.Drawing.Size(99, 21)
        Me.cboTraversal.TabIndex = 2
        '
        'chkCalculateProcessingSize
        '
        Me.chkCalculateProcessingSize.AutoSize = True
        Me.chkCalculateProcessingSize.Location = New System.Drawing.Point(163, 16)
        Me.chkCalculateProcessingSize.Margin = New System.Windows.Forms.Padding(2)
        Me.chkCalculateProcessingSize.Name = "chkCalculateProcessingSize"
        Me.chkCalculateProcessingSize.Size = New System.Drawing.Size(192, 17)
        Me.chkCalculateProcessingSize.TabIndex = 1
        Me.chkCalculateProcessingSize.Text = "Calculate Processing Size Up Front"
        Me.chkCalculateProcessingSize.UseVisualStyleBackColor = True
        '
        'chkPerformItemIdentification
        '
        Me.chkPerformItemIdentification.AutoSize = True
        Me.chkPerformItemIdentification.Location = New System.Drawing.Point(11, 16)
        Me.chkPerformItemIdentification.Margin = New System.Windows.Forms.Padding(2)
        Me.chkPerformItemIdentification.Name = "chkPerformItemIdentification"
        Me.chkPerformItemIdentification.Size = New System.Drawing.Size(148, 17)
        Me.chkPerformItemIdentification.TabIndex = 0
        Me.chkPerformItemIdentification.Text = "Perform Item Identification"
        Me.chkPerformItemIdentification.UseVisualStyleBackColor = True
        '
        'tabArchiveExtractionSettings
        '
        Me.tabArchiveExtractionSettings.Controls.Add(Me.grpArchiveExtractionSettings)
        Me.tabArchiveExtractionSettings.Location = New System.Drawing.Point(4, 22)
        Me.tabArchiveExtractionSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.tabArchiveExtractionSettings.Name = "tabArchiveExtractionSettings"
        Me.tabArchiveExtractionSettings.Size = New System.Drawing.Size(692, 422)
        Me.tabArchiveExtractionSettings.TabIndex = 5
        Me.tabArchiveExtractionSettings.Text = "Lightspeed Settings"
        Me.tabArchiveExtractionSettings.UseVisualStyleBackColor = True
        '
        'grpArchiveExtractionSettings
        '
        Me.grpArchiveExtractionSettings.Controls.Add(Me.numExtractWorkerTimeout)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.lblExtractWorkerTimeout)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.grpEMLExportOptions)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.grpPSTExportOptions)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.txtMemoryPerNuixInstance)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.txtMemoryPerWorker)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.lbNuixAppMemory)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.lblAvailableRAMAfterNuix)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.lblAvailableAfterNuix)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.lblNuixInstances)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.cboNumberOfNuixInstances)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.lblMemoryPerWorker)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.cboNuixWorkers)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.lblNumberOfNuixWorkers)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.lblAvailableRAM)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.lblRAM)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.lblAvailableSystemMemory)
        Me.grpArchiveExtractionSettings.Controls.Add(Me.lblTotalSystemMemory)
        Me.grpArchiveExtractionSettings.Location = New System.Drawing.Point(2, 2)
        Me.grpArchiveExtractionSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpArchiveExtractionSettings.Name = "grpArchiveExtractionSettings"
        Me.grpArchiveExtractionSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpArchiveExtractionSettings.Size = New System.Drawing.Size(691, 419)
        Me.grpArchiveExtractionSettings.TabIndex = 0
        Me.grpArchiveExtractionSettings.TabStop = False
        Me.grpArchiveExtractionSettings.Text = "Lightspeed Settings"
        '
        'numExtractWorkerTimeout
        '
        Me.numExtractWorkerTimeout.Increment = New Decimal(New Integer() {3600, 0, 0, 0})
        Me.numExtractWorkerTimeout.Location = New System.Drawing.Point(181, 205)
        Me.numExtractWorkerTimeout.Maximum = New Decimal(New Integer() {86400, 0, 0, 0})
        Me.numExtractWorkerTimeout.Minimum = New Decimal(New Integer() {3600, 0, 0, 0})
        Me.numExtractWorkerTimeout.Name = "numExtractWorkerTimeout"
        Me.numExtractWorkerTimeout.Size = New System.Drawing.Size(55, 20)
        Me.numExtractWorkerTimeout.TabIndex = 56
        Me.numExtractWorkerTimeout.Value = New Decimal(New Integer() {3600, 0, 0, 0})
        '
        'lblExtractWorkerTimeout
        '
        Me.lblExtractWorkerTimeout.AutoSize = True
        Me.lblExtractWorkerTimeout.Location = New System.Drawing.Point(14, 211)
        Me.lblExtractWorkerTimeout.Name = "lblExtractWorkerTimeout"
        Me.lblExtractWorkerTimeout.Size = New System.Drawing.Size(86, 13)
        Me.lblExtractWorkerTimeout.TabIndex = 55
        Me.lblExtractWorkerTimeout.Text = "Worker Timeout:"
        '
        'grpEMLExportOptions
        '
        Me.grpEMLExportOptions.Controls.Add(Me.chkEMLAddDistributionListMetadata)
        Me.grpEMLExportOptions.Location = New System.Drawing.Point(281, 110)
        Me.grpEMLExportOptions.Name = "grpEMLExportOptions"
        Me.grpEMLExportOptions.Size = New System.Drawing.Size(376, 100)
        Me.grpEMLExportOptions.TabIndex = 54
        Me.grpEMLExportOptions.TabStop = False
        Me.grpEMLExportOptions.Text = "EML Export Options"
        '
        'chkEMLAddDistributionListMetadata
        '
        Me.chkEMLAddDistributionListMetadata.AutoSize = True
        Me.chkEMLAddDistributionListMetadata.Location = New System.Drawing.Point(9, 19)
        Me.chkEMLAddDistributionListMetadata.Name = "chkEMLAddDistributionListMetadata"
        Me.chkEMLAddDistributionListMetadata.Size = New System.Drawing.Size(333, 17)
        Me.chkEMLAddDistributionListMetadata.TabIndex = 0
        Me.chkEMLAddDistributionListMetadata.Text = "Add Distribution List Recipients to Individual EXPANDED-DL field"
        Me.chkEMLAddDistributionListMetadata.UseVisualStyleBackColor = True
        '
        'grpPSTExportOptions
        '
        Me.grpPSTExportOptions.Controls.Add(Me.chkPSTAddDistributionListMetadata)
        Me.grpPSTExportOptions.Controls.Add(Me.lblPSTExportSize)
        Me.grpPSTExportOptions.Controls.Add(Me.numPSTExportSize)
        Me.grpPSTExportOptions.Location = New System.Drawing.Point(281, 24)
        Me.grpPSTExportOptions.Name = "grpPSTExportOptions"
        Me.grpPSTExportOptions.Size = New System.Drawing.Size(376, 80)
        Me.grpPSTExportOptions.TabIndex = 53
        Me.grpPSTExportOptions.TabStop = False
        Me.grpPSTExportOptions.Text = "MAPI Export Options"
        '
        'chkPSTAddDistributionListMetadata
        '
        Me.chkPSTAddDistributionListMetadata.AutoSize = True
        Me.chkPSTAddDistributionListMetadata.Location = New System.Drawing.Point(9, 54)
        Me.chkPSTAddDistributionListMetadata.Name = "chkPSTAddDistributionListMetadata"
        Me.chkPSTAddDistributionListMetadata.Size = New System.Drawing.Size(314, 17)
        Me.chkPSTAddDistributionListMetadata.TabIndex = 2
        Me.chkPSTAddDistributionListMetadata.Text = "Add Distribution List Recipients to MAPI-EXPANDED-DL field"
        Me.chkPSTAddDistributionListMetadata.UseVisualStyleBackColor = True
        '
        'lblPSTExportSize
        '
        Me.lblPSTExportSize.AutoSize = True
        Me.lblPSTExportSize.Location = New System.Drawing.Point(6, 28)
        Me.lblPSTExportSize.Name = "lblPSTExportSize"
        Me.lblPSTExportSize.Size = New System.Drawing.Size(111, 13)
        Me.lblPSTExportSize.TabIndex = 1
        Me.lblPSTExportSize.Text = "PST Export Size (GB):"
        '
        'numPSTExportSize
        '
        Me.numPSTExportSize.Location = New System.Drawing.Point(118, 26)
        Me.numPSTExportSize.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.numPSTExportSize.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numPSTExportSize.Name = "numPSTExportSize"
        Me.numPSTExportSize.Size = New System.Drawing.Size(40, 20)
        Me.numPSTExportSize.TabIndex = 0
        Me.numPSTExportSize.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'txtMemoryPerNuixInstance
        '
        Me.txtMemoryPerNuixInstance.Location = New System.Drawing.Point(181, 76)
        Me.txtMemoryPerNuixInstance.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMemoryPerNuixInstance.Name = "txtMemoryPerNuixInstance"
        Me.txtMemoryPerNuixInstance.Size = New System.Drawing.Size(55, 20)
        Me.txtMemoryPerNuixInstance.TabIndex = 40
        '
        'txtMemoryPerWorker
        '
        Me.txtMemoryPerWorker.Location = New System.Drawing.Point(181, 180)
        Me.txtMemoryPerWorker.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMemoryPerWorker.Name = "txtMemoryPerWorker"
        Me.txtMemoryPerWorker.Size = New System.Drawing.Size(55, 20)
        Me.txtMemoryPerWorker.TabIndex = 42
        '
        'lbNuixAppMemory
        '
        Me.lbNuixAppMemory.AutoSize = True
        Me.lbNuixAppMemory.Location = New System.Drawing.Point(14, 76)
        Me.lbNuixAppMemory.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbNuixAppMemory.Name = "lbNuixAppMemory"
        Me.lbNuixAppMemory.Size = New System.Drawing.Size(93, 13)
        Me.lbNuixAppMemory.TabIndex = 52
        Me.lbNuixAppMemory.Text = "Nuix App Memory:"
        '
        'lblAvailableRAMAfterNuix
        '
        Me.lblAvailableRAMAfterNuix.AutoSize = True
        Me.lblAvailableRAMAfterNuix.Location = New System.Drawing.Point(181, 128)
        Me.lblAvailableRAMAfterNuix.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblAvailableRAMAfterNuix.Name = "lblAvailableRAMAfterNuix"
        Me.lblAvailableRAMAfterNuix.Size = New System.Drawing.Size(37, 13)
        Me.lblAvailableRAMAfterNuix.TabIndex = 51
        Me.lblAvailableRAMAfterNuix.Text = "00000"
        '
        'lblAvailableAfterNuix
        '
        Me.lblAvailableAfterNuix.AutoSize = True
        Me.lblAvailableAfterNuix.Location = New System.Drawing.Point(14, 128)
        Me.lblAvailableAfterNuix.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblAvailableAfterNuix.Name = "lblAvailableAfterNuix"
        Me.lblAvailableAfterNuix.Size = New System.Drawing.Size(151, 13)
        Me.lblAvailableAfterNuix.TabIndex = 50
        Me.lblAvailableAfterNuix.Text = "Available After Nuix Instances:"
        '
        'lblNuixInstances
        '
        Me.lblNuixInstances.AutoSize = True
        Me.lblNuixInstances.Location = New System.Drawing.Point(14, 50)
        Me.lblNuixInstances.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNuixInstances.Name = "lblNuixInstances"
        Me.lblNuixInstances.Size = New System.Drawing.Size(132, 13)
        Me.lblNuixInstances.TabIndex = 49
        Me.lblNuixInstances.Text = "Number of Nuix Instances:"
        '
        'cboNumberOfNuixInstances
        '
        Me.cboNumberOfNuixInstances.FormattingEnabled = True
        Me.cboNumberOfNuixInstances.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"})
        Me.cboNumberOfNuixInstances.Location = New System.Drawing.Point(181, 50)
        Me.cboNumberOfNuixInstances.Margin = New System.Windows.Forms.Padding(2)
        Me.cboNumberOfNuixInstances.Name = "cboNumberOfNuixInstances"
        Me.cboNumberOfNuixInstances.Size = New System.Drawing.Size(55, 21)
        Me.cboNumberOfNuixInstances.TabIndex = 39
        '
        'lblMemoryPerWorker
        '
        Me.lblMemoryPerWorker.AutoSize = True
        Me.lblMemoryPerWorker.Location = New System.Drawing.Point(14, 180)
        Me.lblMemoryPerWorker.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblMemoryPerWorker.Name = "lblMemoryPerWorker"
        Me.lblMemoryPerWorker.Size = New System.Drawing.Size(129, 13)
        Me.lblMemoryPerWorker.TabIndex = 48
        Me.lblMemoryPerWorker.Text = "Memory Per Worker (MB):"
        '
        'cboNuixWorkers
        '
        Me.cboNuixWorkers.FormattingEnabled = True
        Me.cboNuixWorkers.Items.AddRange(New Object() {"1", "2", "4", "6", "8", "10", "12", "16", "20", "24"})
        Me.cboNuixWorkers.Location = New System.Drawing.Point(181, 154)
        Me.cboNuixWorkers.Margin = New System.Windows.Forms.Padding(2)
        Me.cboNuixWorkers.Name = "cboNuixWorkers"
        Me.cboNuixWorkers.Size = New System.Drawing.Size(55, 21)
        Me.cboNuixWorkers.TabIndex = 41
        '
        'lblNumberOfNuixWorkers
        '
        Me.lblNumberOfNuixWorkers.AutoSize = True
        Me.lblNumberOfNuixWorkers.Location = New System.Drawing.Point(14, 154)
        Me.lblNumberOfNuixWorkers.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNumberOfNuixWorkers.Name = "lblNumberOfNuixWorkers"
        Me.lblNumberOfNuixWorkers.Size = New System.Drawing.Size(128, 13)
        Me.lblNumberOfNuixWorkers.TabIndex = 47
        Me.lblNumberOfNuixWorkers.Text = "Number Of Nuix Workers:"
        '
        'lblAvailableRAM
        '
        Me.lblAvailableRAM.AutoSize = True
        Me.lblAvailableRAM.Location = New System.Drawing.Point(181, 102)
        Me.lblAvailableRAM.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblAvailableRAM.Name = "lblAvailableRAM"
        Me.lblAvailableRAM.Size = New System.Drawing.Size(84, 13)
        Me.lblAvailableRAM.TabIndex = 46
        Me.lblAvailableRAM.Text = "lblAvailableRAM"
        '
        'lblRAM
        '
        Me.lblRAM.AutoSize = True
        Me.lblRAM.Location = New System.Drawing.Point(181, 24)
        Me.lblRAM.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRAM.Name = "lblRAM"
        Me.lblRAM.Size = New System.Drawing.Size(41, 13)
        Me.lblRAM.TabIndex = 45
        Me.lblRAM.Text = "lblRAM"
        '
        'lblAvailableSystemMemory
        '
        Me.lblAvailableSystemMemory.AutoSize = True
        Me.lblAvailableSystemMemory.Location = New System.Drawing.Point(14, 102)
        Me.lblAvailableSystemMemory.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblAvailableSystemMemory.Name = "lblAvailableSystemMemory"
        Me.lblAvailableSystemMemory.Size = New System.Drawing.Size(129, 13)
        Me.lblAvailableSystemMemory.TabIndex = 44
        Me.lblAvailableSystemMemory.Text = "Available Memory (RAM): "
        '
        'lblTotalSystemMemory
        '
        Me.lblTotalSystemMemory.AutoSize = True
        Me.lblTotalSystemMemory.Location = New System.Drawing.Point(14, 24)
        Me.lblTotalSystemMemory.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTotalSystemMemory.Name = "lblTotalSystemMemory"
        Me.lblTotalSystemMemory.Size = New System.Drawing.Size(120, 13)
        Me.lblTotalSystemMemory.TabIndex = 43
        Me.lblTotalSystemMemory.Text = "System Memory (RAM): "
        '
        'tabEWSSettings
        '
        Me.tabEWSSettings.Controls.Add(Me.grpOffice365Settings)
        Me.tabEWSSettings.Location = New System.Drawing.Point(4, 22)
        Me.tabEWSSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.tabEWSSettings.Name = "tabEWSSettings"
        Me.tabEWSSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.tabEWSSettings.Size = New System.Drawing.Size(692, 422)
        Me.tabEWSSettings.TabIndex = 1
        Me.tabEWSSettings.Text = "Exchange Web Services"
        Me.tabEWSSettings.UseVisualStyleBackColor = True
        '
        'grpOffice365Settings
        '
        Me.grpOffice365Settings.Controls.Add(Me.lblPSTConsolidation)
        Me.grpOffice365Settings.Controls.Add(Me.cboPSTConsolidation)
        Me.grpOffice365Settings.Controls.Add(Me.grpEWSDownloadControl)
        Me.grpOffice365Settings.Controls.Add(Me.grpEWSUploadControl)
        Me.grpOffice365Settings.Controls.Add(Me.grpEWSThrottling)
        Me.grpOffice365Settings.Controls.Add(Me.grpNuixInstanceSettings)
        Me.grpOffice365Settings.Controls.Add(Me.chkO365ApplicationImpersonation)
        Me.grpOffice365Settings.Controls.Add(Me.txtO365AdminInfo)
        Me.grpOffice365Settings.Controls.Add(Me.txtO365AdminUserName)
        Me.grpOffice365Settings.Controls.Add(Me.txtO365Domain)
        Me.grpOffice365Settings.Controls.Add(Me.txtO365ExchangeServer)
        Me.grpOffice365Settings.Controls.Add(Me.lblPSTInfoAdminInfo)
        Me.grpOffice365Settings.Controls.Add(Me.lblPSTInfoAdminUserName)
        Me.grpOffice365Settings.Controls.Add(Me.lblPSTInfoDomain)
        Me.grpOffice365Settings.Controls.Add(Me.lblPSTInfoExchangeServer)
        Me.grpOffice365Settings.Location = New System.Drawing.Point(2, 4)
        Me.grpOffice365Settings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpOffice365Settings.Name = "grpOffice365Settings"
        Me.grpOffice365Settings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpOffice365Settings.Size = New System.Drawing.Size(689, 414)
        Me.grpOffice365Settings.TabIndex = 57
        Me.grpOffice365Settings.TabStop = False
        Me.grpOffice365Settings.Text = "EWS Connection"
        '
        'lblPSTConsolidation
        '
        Me.lblPSTConsolidation.AutoSize = True
        Me.lblPSTConsolidation.Location = New System.Drawing.Point(196, 132)
        Me.lblPSTConsolidation.Name = "lblPSTConsolidation"
        Me.lblPSTConsolidation.Size = New System.Drawing.Size(97, 13)
        Me.lblPSTConsolidation.TabIndex = 111
        Me.lblPSTConsolidation.Text = "PST Consolidation:"
        '
        'cboPSTConsolidation
        '
        Me.cboPSTConsolidation.FormattingEnabled = True
        Me.cboPSTConsolidation.Items.AddRange(New Object() {"Copy", "Move"})
        Me.cboPSTConsolidation.Location = New System.Drawing.Point(301, 128)
        Me.cboPSTConsolidation.Name = "cboPSTConsolidation"
        Me.cboPSTConsolidation.Size = New System.Drawing.Size(86, 21)
        Me.cboPSTConsolidation.TabIndex = 110
        '
        'grpEWSDownloadControl
        '
        Me.grpEWSDownloadControl.Controls.Add(Me.chkEnableMailboxSlackspace)
        Me.grpEWSDownloadControl.Controls.Add(Me.chkCollaborativePrefetching)
        Me.grpEWSDownloadControl.Controls.Add(Me.chkEnablePrefetch)
        Me.grpEWSDownloadControl.Controls.Add(Me.numDownloadCount)
        Me.grpEWSDownloadControl.Controls.Add(Me.lblMaxDownloadCount)
        Me.grpEWSDownloadControl.Controls.Add(Me.numDownloadSize)
        Me.grpEWSDownloadControl.Controls.Add(Me.lblMaxMessageDownloadSize)
        Me.grpEWSDownloadControl.Location = New System.Drawing.Point(420, 160)
        Me.grpEWSDownloadControl.Name = "grpEWSDownloadControl"
        Me.grpEWSDownloadControl.Size = New System.Drawing.Size(248, 140)
        Me.grpEWSDownloadControl.TabIndex = 109
        Me.grpEWSDownloadControl.TabStop = False
        Me.grpEWSDownloadControl.Text = "EWS Download Control"
        '
        'chkEnableMailboxSlackspace
        '
        Me.chkEnableMailboxSlackspace.AutoSize = True
        Me.chkEnableMailboxSlackspace.Location = New System.Drawing.Point(9, 117)
        Me.chkEnableMailboxSlackspace.Name = "chkEnableMailboxSlackspace"
        Me.chkEnableMailboxSlackspace.Size = New System.Drawing.Size(162, 17)
        Me.chkEnableMailboxSlackspace.TabIndex = 6
        Me.chkEnableMailboxSlackspace.Text = "Enable Mailbox Slack Space"
        Me.chkEnableMailboxSlackspace.UseVisualStyleBackColor = True
        '
        'chkCollaborativePrefetching
        '
        Me.chkCollaborativePrefetching.AutoSize = True
        Me.chkCollaborativePrefetching.Location = New System.Drawing.Point(10, 96)
        Me.chkCollaborativePrefetching.Name = "chkCollaborativePrefetching"
        Me.chkCollaborativePrefetching.Size = New System.Drawing.Size(167, 17)
        Me.chkCollaborativePrefetching.TabIndex = 5
        Me.chkCollaborativePrefetching.Text = "Enable Collaborative Fetching"
        Me.chkCollaborativePrefetching.UseVisualStyleBackColor = True
        '
        'chkEnablePrefetch
        '
        Me.chkEnablePrefetch.AutoSize = True
        Me.chkEnablePrefetch.Location = New System.Drawing.Point(10, 23)
        Me.chkEnablePrefetch.Name = "chkEnablePrefetch"
        Me.chkEnablePrefetch.Size = New System.Drawing.Size(134, 17)
        Me.chkEnablePrefetch.TabIndex = 4
        Me.chkEnablePrefetch.Text = "Enable Bulk Download"
        Me.chkEnablePrefetch.UseVisualStyleBackColor = True
        '
        'numDownloadCount
        '
        Me.numDownloadCount.Increment = New Decimal(New Integer() {100, 0, 0, 0})
        Me.numDownloadCount.Location = New System.Drawing.Point(178, 67)
        Me.numDownloadCount.Maximum = New Decimal(New Integer() {500, 0, 0, 0})
        Me.numDownloadCount.Minimum = New Decimal(New Integer() {100, 0, 0, 0})
        Me.numDownloadCount.Name = "numDownloadCount"
        Me.numDownloadCount.Size = New System.Drawing.Size(60, 20)
        Me.numDownloadCount.TabIndex = 3
        Me.numDownloadCount.Value = New Decimal(New Integer() {500, 0, 0, 0})
        '
        'lblMaxDownloadCount
        '
        Me.lblMaxDownloadCount.AutoSize = True
        Me.lblMaxDownloadCount.Location = New System.Drawing.Point(28, 69)
        Me.lblMaxDownloadCount.Name = "lblMaxDownloadCount"
        Me.lblMaxDownloadCount.Size = New System.Drawing.Size(136, 13)
        Me.lblMaxDownloadCount.TabIndex = 2
        Me.lblMaxDownloadCount.Text = "Maximum Download Count:"
        '
        'numDownloadSize
        '
        Me.numDownloadSize.Increment = New Decimal(New Integer() {5, 0, 0, 0})
        Me.numDownloadSize.Location = New System.Drawing.Point(178, 42)
        Me.numDownloadSize.Maximum = New Decimal(New Integer() {50, 0, 0, 0})
        Me.numDownloadSize.Minimum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.numDownloadSize.Name = "numDownloadSize"
        Me.numDownloadSize.Size = New System.Drawing.Size(60, 20)
        Me.numDownloadSize.TabIndex = 1
        Me.numDownloadSize.Value = New Decimal(New Integer() {25, 0, 0, 0})
        '
        'lblMaxMessageDownloadSize
        '
        Me.lblMaxMessageDownloadSize.AutoSize = True
        Me.lblMaxMessageDownloadSize.Location = New System.Drawing.Point(28, 44)
        Me.lblMaxMessageDownloadSize.Name = "lblMaxMessageDownloadSize"
        Me.lblMaxMessageDownloadSize.Size = New System.Drawing.Size(148, 13)
        Me.lblMaxMessageDownloadSize.TabIndex = 0
        Me.lblMaxMessageDownloadSize.Text = "Maximum Message Size (MB):"
        '
        'grpEWSUploadControl
        '
        Me.grpEWSUploadControl.Controls.Add(Me.lblBulkUploadsze)
        Me.grpEWSUploadControl.Controls.Add(Me.numBulkUploadSize)
        Me.grpEWSUploadControl.Controls.Add(Me.chkEnableBulkUpload)
        Me.grpEWSUploadControl.Controls.Add(Me.Label11)
        Me.grpEWSUploadControl.Controls.Add(Me.txtO365RemovePathPrefix)
        Me.grpEWSUploadControl.Controls.Add(Me.lblRemovePathPrefix)
        Me.grpEWSUploadControl.Controls.Add(Me.lblMaxItemUploadSize)
        Me.grpEWSUploadControl.Controls.Add(Me.numEWSMaxUploadSize)
        Me.grpEWSUploadControl.Location = New System.Drawing.Point(419, 10)
        Me.grpEWSUploadControl.Name = "grpEWSUploadControl"
        Me.grpEWSUploadControl.Size = New System.Drawing.Size(248, 144)
        Me.grpEWSUploadControl.TabIndex = 108
        Me.grpEWSUploadControl.TabStop = False
        Me.grpEWSUploadControl.Text = "EWS Upload Control"
        '
        'lblBulkUploadsze
        '
        Me.lblBulkUploadsze.AutoSize = True
        Me.lblBulkUploadsze.Location = New System.Drawing.Point(25, 78)
        Me.lblBulkUploadsze.Name = "lblBulkUploadsze"
        Me.lblBulkUploadsze.Size = New System.Drawing.Size(116, 13)
        Me.lblBulkUploadsze.TabIndex = 114
        Me.lblBulkUploadsze.Text = "Bulk Upload Size (MB):"
        '
        'numBulkUploadSize
        '
        Me.numBulkUploadSize.Increment = New Decimal(New Integer() {5, 0, 0, 0})
        Me.numBulkUploadSize.Location = New System.Drawing.Point(160, 76)
        Me.numBulkUploadSize.Maximum = New Decimal(New Integer() {50, 0, 0, 0})
        Me.numBulkUploadSize.Minimum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.numBulkUploadSize.Name = "numBulkUploadSize"
        Me.numBulkUploadSize.Size = New System.Drawing.Size(55, 20)
        Me.numBulkUploadSize.TabIndex = 113
        Me.numBulkUploadSize.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'chkEnableBulkUpload
        '
        Me.chkEnableBulkUpload.AutoSize = True
        Me.chkEnableBulkUpload.Location = New System.Drawing.Point(10, 52)
        Me.chkEnableBulkUpload.Name = "chkEnableBulkUpload"
        Me.chkEnableBulkUpload.Size = New System.Drawing.Size(120, 17)
        Me.chkEnableBulkUpload.TabIndex = 112
        Me.chkEnableBulkUpload.Text = "Enable Bulk Upload"
        Me.chkEnableBulkUpload.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(76, 107)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(113, 13)
        Me.Label11.TabIndex = 111
        Me.Label11.Text = " levels from folder path"
        '
        'txtO365RemovePathPrefix
        '
        Me.txtO365RemovePathPrefix.Location = New System.Drawing.Point(58, 107)
        Me.txtO365RemovePathPrefix.Margin = New System.Windows.Forms.Padding(2)
        Me.txtO365RemovePathPrefix.Name = "txtO365RemovePathPrefix"
        Me.txtO365RemovePathPrefix.Size = New System.Drawing.Size(18, 20)
        Me.txtO365RemovePathPrefix.TabIndex = 110
        '
        'lblRemovePathPrefix
        '
        Me.lblRemovePathPrefix.AutoSize = True
        Me.lblRemovePathPrefix.Location = New System.Drawing.Point(10, 107)
        Me.lblRemovePathPrefix.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRemovePathPrefix.Name = "lblRemovePathPrefix"
        Me.lblRemovePathPrefix.Size = New System.Drawing.Size(50, 13)
        Me.lblRemovePathPrefix.TabIndex = 109
        Me.lblRemovePathPrefix.Text = "Remove "
        '
        'lblMaxItemUploadSize
        '
        Me.lblMaxItemUploadSize.AutoSize = True
        Me.lblMaxItemUploadSize.Location = New System.Drawing.Point(10, 26)
        Me.lblMaxItemUploadSize.Name = "lblMaxItemUploadSize"
        Me.lblMaxItemUploadSize.Size = New System.Drawing.Size(148, 13)
        Me.lblMaxItemUploadSize.TabIndex = 108
        Me.lblMaxItemUploadSize.Text = "Maximum Message Size (MB):"
        '
        'numEWSMaxUploadSize
        '
        Me.numEWSMaxUploadSize.Increment = New Decimal(New Integer() {5, 0, 0, 0})
        Me.numEWSMaxUploadSize.Location = New System.Drawing.Point(160, 24)
        Me.numEWSMaxUploadSize.Maximum = New Decimal(New Integer() {50, 0, 0, 0})
        Me.numEWSMaxUploadSize.Minimum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.numEWSMaxUploadSize.Name = "numEWSMaxUploadSize"
        Me.numEWSMaxUploadSize.Size = New System.Drawing.Size(55, 20)
        Me.numEWSMaxUploadSize.TabIndex = 107
        Me.numEWSMaxUploadSize.Value = New Decimal(New Integer() {25, 0, 0, 0})
        '
        'grpEWSThrottling
        '
        Me.grpEWSThrottling.Controls.Add(Me.numEWSRetryIncrement)
        Me.grpEWSThrottling.Controls.Add(Me.numEWSRetryDelay)
        Me.grpEWSThrottling.Controls.Add(Me.numEWSRetryCount)
        Me.grpEWSThrottling.Controls.Add(Me.lblRetryIncrement)
        Me.grpEWSThrottling.Controls.Add(Me.lblRetryDelay)
        Me.grpEWSThrottling.Controls.Add(Me.lblRetryCount)
        Me.grpEWSThrottling.Location = New System.Drawing.Point(419, 303)
        Me.grpEWSThrottling.Margin = New System.Windows.Forms.Padding(2)
        Me.grpEWSThrottling.Name = "grpEWSThrottling"
        Me.grpEWSThrottling.Padding = New System.Windows.Forms.Padding(2)
        Me.grpEWSThrottling.Size = New System.Drawing.Size(248, 104)
        Me.grpEWSThrottling.TabIndex = 104
        Me.grpEWSThrottling.TabStop = False
        Me.grpEWSThrottling.Text = "EWS Throttle Control"
        '
        'numEWSRetryIncrement
        '
        Me.numEWSRetryIncrement.Location = New System.Drawing.Point(98, 75)
        Me.numEWSRetryIncrement.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.numEWSRetryIncrement.Name = "numEWSRetryIncrement"
        Me.numEWSRetryIncrement.Size = New System.Drawing.Size(35, 20)
        Me.numEWSRetryIncrement.TabIndex = 110
        Me.numEWSRetryIncrement.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'numEWSRetryDelay
        '
        Me.numEWSRetryDelay.Location = New System.Drawing.Point(98, 48)
        Me.numEWSRetryDelay.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.numEWSRetryDelay.Name = "numEWSRetryDelay"
        Me.numEWSRetryDelay.Size = New System.Drawing.Size(35, 20)
        Me.numEWSRetryDelay.TabIndex = 109
        Me.numEWSRetryDelay.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'numEWSRetryCount
        '
        Me.numEWSRetryCount.Location = New System.Drawing.Point(98, 19)
        Me.numEWSRetryCount.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.numEWSRetryCount.Name = "numEWSRetryCount"
        Me.numEWSRetryCount.Size = New System.Drawing.Size(35, 20)
        Me.numEWSRetryCount.TabIndex = 108
        Me.numEWSRetryCount.Value = New Decimal(New Integer() {3, 0, 0, 0})
        '
        'lblRetryIncrement
        '
        Me.lblRetryIncrement.AutoSize = True
        Me.lblRetryIncrement.Location = New System.Drawing.Point(11, 73)
        Me.lblRetryIncrement.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRetryIncrement.Name = "lblRetryIncrement"
        Me.lblRetryIncrement.Size = New System.Drawing.Size(85, 13)
        Me.lblRetryIncrement.TabIndex = 107
        Me.lblRetryIncrement.Text = "Retry Increment:"
        '
        'lblRetryDelay
        '
        Me.lblRetryDelay.AutoSize = True
        Me.lblRetryDelay.Location = New System.Drawing.Point(12, 47)
        Me.lblRetryDelay.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRetryDelay.Name = "lblRetryDelay"
        Me.lblRetryDelay.Size = New System.Drawing.Size(65, 13)
        Me.lblRetryDelay.TabIndex = 104
        Me.lblRetryDelay.Text = "Retry Delay:"
        '
        'lblRetryCount
        '
        Me.lblRetryCount.AutoSize = True
        Me.lblRetryCount.Location = New System.Drawing.Point(12, 21)
        Me.lblRetryCount.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRetryCount.Name = "lblRetryCount"
        Me.lblRetryCount.Size = New System.Drawing.Size(66, 13)
        Me.lblRetryCount.TabIndex = 102
        Me.lblRetryCount.Text = "Retry Count:"
        '
        'grpNuixInstanceSettings
        '
        Me.grpNuixInstanceSettings.Controls.Add(Me.numWorkerTimeout)
        Me.grpNuixInstanceSettings.Controls.Add(Me.Label6)
        Me.grpNuixInstanceSettings.Controls.Add(Me.txtO365NuixAppMemory)
        Me.grpNuixInstanceSettings.Controls.Add(Me.txtO365MemoryPerWorker)
        Me.grpNuixInstanceSettings.Controls.Add(Me.Label5)
        Me.grpNuixInstanceSettings.Controls.Add(Me.lblO365AvailableMemoryAfterNuix)
        Me.grpNuixInstanceSettings.Controls.Add(Me.Label7)
        Me.grpNuixInstanceSettings.Controls.Add(Me.Label8)
        Me.grpNuixInstanceSettings.Controls.Add(Me.cboO365NumberOfNuixInstances)
        Me.grpNuixInstanceSettings.Controls.Add(Me.Label9)
        Me.grpNuixInstanceSettings.Controls.Add(Me.cboO365NumberOfNuixWorkers)
        Me.grpNuixInstanceSettings.Controls.Add(Me.Label10)
        Me.grpNuixInstanceSettings.Controls.Add(Me.lblO365AvailMemory)
        Me.grpNuixInstanceSettings.Controls.Add(Me.lblO365SystemMemory)
        Me.grpNuixInstanceSettings.Controls.Add(Me.Label13)
        Me.grpNuixInstanceSettings.Controls.Add(Me.Label14)
        Me.grpNuixInstanceSettings.Location = New System.Drawing.Point(10, 160)
        Me.grpNuixInstanceSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.grpNuixInstanceSettings.Name = "grpNuixInstanceSettings"
        Me.grpNuixInstanceSettings.Padding = New System.Windows.Forms.Padding(2)
        Me.grpNuixInstanceSettings.Size = New System.Drawing.Size(377, 247)
        Me.grpNuixInstanceSettings.TabIndex = 95
        Me.grpNuixInstanceSettings.TabStop = False
        Me.grpNuixInstanceSettings.Text = "Lightspeed Settings"
        '
        'numWorkerTimeout
        '
        Me.numWorkerTimeout.Increment = New Decimal(New Integer() {3600, 0, 0, 0})
        Me.numWorkerTimeout.Location = New System.Drawing.Point(180, 209)
        Me.numWorkerTimeout.Maximum = New Decimal(New Integer() {86400, 0, 0, 0})
        Me.numWorkerTimeout.Minimum = New Decimal(New Integer() {3600, 0, 0, 0})
        Me.numWorkerTimeout.Name = "numWorkerTimeout"
        Me.numWorkerTimeout.Size = New System.Drawing.Size(77, 20)
        Me.numWorkerTimeout.TabIndex = 110
        Me.numWorkerTimeout.Value = New Decimal(New Integer() {3600, 0, 0, 0})
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(15, 209)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(86, 13)
        Me.Label6.TabIndex = 109
        Me.Label6.Text = "Worker Timeout:"
        '
        'txtO365NuixAppMemory
        '
        Me.txtO365NuixAppMemory.Location = New System.Drawing.Point(180, 77)
        Me.txtO365NuixAppMemory.Margin = New System.Windows.Forms.Padding(2)
        Me.txtO365NuixAppMemory.Name = "txtO365NuixAppMemory"
        Me.txtO365NuixAppMemory.Size = New System.Drawing.Size(55, 20)
        Me.txtO365NuixAppMemory.TabIndex = 96
        '
        'txtO365MemoryPerWorker
        '
        Me.txtO365MemoryPerWorker.Location = New System.Drawing.Point(180, 181)
        Me.txtO365MemoryPerWorker.Margin = New System.Windows.Forms.Padding(2)
        Me.txtO365MemoryPerWorker.Name = "txtO365MemoryPerWorker"
        Me.txtO365MemoryPerWorker.Size = New System.Drawing.Size(55, 20)
        Me.txtO365MemoryPerWorker.TabIndex = 98
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(15, 77)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(93, 13)
        Me.Label5.TabIndex = 108
        Me.Label5.Text = "Nuix App Memory:"
        '
        'lblO365AvailableMemoryAfterNuix
        '
        Me.lblO365AvailableMemoryAfterNuix.AutoSize = True
        Me.lblO365AvailableMemoryAfterNuix.Location = New System.Drawing.Point(180, 129)
        Me.lblO365AvailableMemoryAfterNuix.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblO365AvailableMemoryAfterNuix.Name = "lblO365AvailableMemoryAfterNuix"
        Me.lblO365AvailableMemoryAfterNuix.Size = New System.Drawing.Size(37, 13)
        Me.lblO365AvailableMemoryAfterNuix.TabIndex = 107
        Me.lblO365AvailableMemoryAfterNuix.Text = "00000"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(15, 129)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(151, 13)
        Me.Label7.TabIndex = 106
        Me.Label7.Text = "Available After Nuix Instances:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(15, 51)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(132, 13)
        Me.Label8.TabIndex = 105
        Me.Label8.Text = "Number of Nuix Instances:"
        '
        'cboO365NumberOfNuixInstances
        '
        Me.cboO365NumberOfNuixInstances.FormattingEnabled = True
        Me.cboO365NumberOfNuixInstances.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"})
        Me.cboO365NumberOfNuixInstances.Location = New System.Drawing.Point(180, 51)
        Me.cboO365NumberOfNuixInstances.Margin = New System.Windows.Forms.Padding(2)
        Me.cboO365NumberOfNuixInstances.Name = "cboO365NumberOfNuixInstances"
        Me.cboO365NumberOfNuixInstances.Size = New System.Drawing.Size(55, 21)
        Me.cboO365NumberOfNuixInstances.TabIndex = 95
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(15, 181)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(129, 13)
        Me.Label9.TabIndex = 104
        Me.Label9.Text = "Memory Per Worker (MB):"
        '
        'cboO365NumberOfNuixWorkers
        '
        Me.cboO365NumberOfNuixWorkers.FormattingEnabled = True
        Me.cboO365NumberOfNuixWorkers.Items.AddRange(New Object() {"1", "2", "4"})
        Me.cboO365NumberOfNuixWorkers.Location = New System.Drawing.Point(180, 155)
        Me.cboO365NumberOfNuixWorkers.Margin = New System.Windows.Forms.Padding(2)
        Me.cboO365NumberOfNuixWorkers.Name = "cboO365NumberOfNuixWorkers"
        Me.cboO365NumberOfNuixWorkers.Size = New System.Drawing.Size(55, 21)
        Me.cboO365NumberOfNuixWorkers.TabIndex = 97
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(15, 155)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(128, 13)
        Me.Label10.TabIndex = 103
        Me.Label10.Text = "Number Of Nuix Workers:"
        '
        'lblO365AvailMemory
        '
        Me.lblO365AvailMemory.AutoSize = True
        Me.lblO365AvailMemory.Location = New System.Drawing.Point(180, 103)
        Me.lblO365AvailMemory.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblO365AvailMemory.Name = "lblO365AvailMemory"
        Me.lblO365AvailMemory.Size = New System.Drawing.Size(103, 13)
        Me.lblO365AvailMemory.TabIndex = 102
        Me.lblO365AvailMemory.Text = "lblO365AvailMemory"
        '
        'lblO365SystemMemory
        '
        Me.lblO365SystemMemory.AutoSize = True
        Me.lblO365SystemMemory.Location = New System.Drawing.Point(180, 25)
        Me.lblO365SystemMemory.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblO365SystemMemory.Name = "lblO365SystemMemory"
        Me.lblO365SystemMemory.Size = New System.Drawing.Size(114, 13)
        Me.lblO365SystemMemory.TabIndex = 101
        Me.lblO365SystemMemory.Text = "lblO365SystemMemory"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(15, 103)
        Me.Label13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(129, 13)
        Me.Label13.TabIndex = 100
        Me.Label13.Text = "Available Memory (RAM): "
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(15, 25)
        Me.Label14.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(120, 13)
        Me.Label14.TabIndex = 99
        Me.Label14.Text = "System Memory (RAM): "
        '
        'chkO365ApplicationImpersonation
        '
        Me.chkO365ApplicationImpersonation.AutoSize = True
        Me.chkO365ApplicationImpersonation.Location = New System.Drawing.Point(10, 128)
        Me.chkO365ApplicationImpersonation.Margin = New System.Windows.Forms.Padding(2)
        Me.chkO365ApplicationImpersonation.Name = "chkO365ApplicationImpersonation"
        Me.chkO365ApplicationImpersonation.Size = New System.Drawing.Size(128, 17)
        Me.chkO365ApplicationImpersonation.TabIndex = 72
        Me.chkO365ApplicationImpersonation.Text = "Enable Impersonation"
        Me.chkO365ApplicationImpersonation.UseVisualStyleBackColor = True
        '
        'txtO365AdminInfo
        '
        Me.txtO365AdminInfo.Location = New System.Drawing.Point(119, 97)
        Me.txtO365AdminInfo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtO365AdminInfo.Name = "txtO365AdminInfo"
        Me.txtO365AdminInfo.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtO365AdminInfo.Size = New System.Drawing.Size(268, 20)
        Me.txtO365AdminInfo.TabIndex = 29
        '
        'txtO365AdminUserName
        '
        Me.txtO365AdminUserName.Location = New System.Drawing.Point(119, 71)
        Me.txtO365AdminUserName.Margin = New System.Windows.Forms.Padding(2)
        Me.txtO365AdminUserName.Name = "txtO365AdminUserName"
        Me.txtO365AdminUserName.Size = New System.Drawing.Size(268, 20)
        Me.txtO365AdminUserName.TabIndex = 28
        '
        'txtO365Domain
        '
        Me.txtO365Domain.Location = New System.Drawing.Point(119, 45)
        Me.txtO365Domain.Margin = New System.Windows.Forms.Padding(2)
        Me.txtO365Domain.Name = "txtO365Domain"
        Me.txtO365Domain.Size = New System.Drawing.Size(268, 20)
        Me.txtO365Domain.TabIndex = 27
        '
        'txtO365ExchangeServer
        '
        Me.txtO365ExchangeServer.Location = New System.Drawing.Point(119, 19)
        Me.txtO365ExchangeServer.Margin = New System.Windows.Forms.Padding(2)
        Me.txtO365ExchangeServer.Name = "txtO365ExchangeServer"
        Me.txtO365ExchangeServer.Size = New System.Drawing.Size(268, 20)
        Me.txtO365ExchangeServer.TabIndex = 26
        '
        'lblPSTInfoAdminInfo
        '
        Me.lblPSTInfoAdminInfo.AutoSize = True
        Me.lblPSTInfoAdminInfo.Location = New System.Drawing.Point(10, 97)
        Me.lblPSTInfoAdminInfo.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblPSTInfoAdminInfo.Name = "lblPSTInfoAdminInfo"
        Me.lblPSTInfoAdminInfo.Size = New System.Drawing.Size(56, 13)
        Me.lblPSTInfoAdminInfo.TabIndex = 15
        Me.lblPSTInfoAdminInfo.Text = "Password:"
        '
        'lblPSTInfoAdminUserName
        '
        Me.lblPSTInfoAdminUserName.AutoSize = True
        Me.lblPSTInfoAdminUserName.Location = New System.Drawing.Point(10, 71)
        Me.lblPSTInfoAdminUserName.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblPSTInfoAdminUserName.Name = "lblPSTInfoAdminUserName"
        Me.lblPSTInfoAdminUserName.Size = New System.Drawing.Size(63, 13)
        Me.lblPSTInfoAdminUserName.TabIndex = 14
        Me.lblPSTInfoAdminUserName.Text = "User Name:"
        '
        'lblPSTInfoDomain
        '
        Me.lblPSTInfoDomain.AutoSize = True
        Me.lblPSTInfoDomain.Location = New System.Drawing.Point(10, 45)
        Me.lblPSTInfoDomain.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblPSTInfoDomain.Name = "lblPSTInfoDomain"
        Me.lblPSTInfoDomain.Size = New System.Drawing.Size(46, 13)
        Me.lblPSTInfoDomain.TabIndex = 13
        Me.lblPSTInfoDomain.Text = "Domain:"
        '
        'lblPSTInfoExchangeServer
        '
        Me.lblPSTInfoExchangeServer.AutoSize = True
        Me.lblPSTInfoExchangeServer.Location = New System.Drawing.Point(10, 20)
        Me.lblPSTInfoExchangeServer.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblPSTInfoExchangeServer.Name = "lblPSTInfoExchangeServer"
        Me.lblPSTInfoExchangeServer.Size = New System.Drawing.Size(92, 13)
        Me.lblPSTInfoExchangeServer.TabIndex = 12
        Me.lblPSTInfoExchangeServer.Text = "Exchange Server:"
        '
        'tabDB
        '
        Me.tabDB.Controls.Add(Me.grpRedisConfig)
        Me.tabDB.Controls.Add(Me.btnSQLiteDBLocationSelection)
        Me.tabDB.Controls.Add(Me.txtSQLiteDBLocation)
        Me.tabDB.Controls.Add(Me.Label1)
        Me.tabDB.Location = New System.Drawing.Point(4, 22)
        Me.tabDB.Margin = New System.Windows.Forms.Padding(2)
        Me.tabDB.Name = "tabDB"
        Me.tabDB.Size = New System.Drawing.Size(692, 422)
        Me.tabDB.TabIndex = 4
        Me.tabDB.Text = "Database"
        Me.tabDB.UseVisualStyleBackColor = True
        '
        'grpRedisConfig
        '
        Me.grpRedisConfig.Controls.Add(Me.txtRedisAuth)
        Me.grpRedisConfig.Controls.Add(Me.lblRedis)
        Me.grpRedisConfig.Controls.Add(Me.txtRedisPort)
        Me.grpRedisConfig.Controls.Add(Me.lblRedisPort)
        Me.grpRedisConfig.Controls.Add(Me.txtRedisHostName)
        Me.grpRedisConfig.Controls.Add(Me.lblRedisHostName)
        Me.grpRedisConfig.Location = New System.Drawing.Point(18, 51)
        Me.grpRedisConfig.Name = "grpRedisConfig"
        Me.grpRedisConfig.Size = New System.Drawing.Size(530, 122)
        Me.grpRedisConfig.TabIndex = 55
        Me.grpRedisConfig.TabStop = False
        Me.grpRedisConfig.Text = "Redis Configuration"
        '
        'txtRedisAuth
        '
        Me.txtRedisAuth.Location = New System.Drawing.Point(115, 60)
        Me.txtRedisAuth.Name = "txtRedisAuth"
        Me.txtRedisAuth.Size = New System.Drawing.Size(100, 20)
        Me.txtRedisAuth.TabIndex = 5
        '
        'lblRedis
        '
        Me.lblRedis.AutoSize = True
        Me.lblRedis.Location = New System.Drawing.Point(16, 60)
        Me.lblRedis.Name = "lblRedis"
        Me.lblRedis.Size = New System.Drawing.Size(32, 13)
        Me.lblRedis.TabIndex = 4
        Me.lblRedis.Text = "Auth:"
        '
        'txtRedisPort
        '
        Me.txtRedisPort.Location = New System.Drawing.Point(267, 25)
        Me.txtRedisPort.Name = "txtRedisPort"
        Me.txtRedisPort.Size = New System.Drawing.Size(100, 20)
        Me.txtRedisPort.TabIndex = 3
        '
        'lblRedisPort
        '
        Me.lblRedisPort.AutoSize = True
        Me.lblRedisPort.Location = New System.Drawing.Point(232, 28)
        Me.lblRedisPort.Name = "lblRedisPort"
        Me.lblRedisPort.Size = New System.Drawing.Size(29, 13)
        Me.lblRedisPort.TabIndex = 2
        Me.lblRedisPort.Text = "Port:"
        '
        'txtRedisHostName
        '
        Me.txtRedisHostName.Location = New System.Drawing.Point(115, 25)
        Me.txtRedisHostName.Name = "txtRedisHostName"
        Me.txtRedisHostName.Size = New System.Drawing.Size(100, 20)
        Me.txtRedisHostName.TabIndex = 1
        '
        'lblRedisHostName
        '
        Me.lblRedisHostName.AutoSize = True
        Me.lblRedisHostName.Location = New System.Drawing.Point(16, 28)
        Me.lblRedisHostName.Name = "lblRedisHostName"
        Me.lblRedisHostName.Size = New System.Drawing.Size(93, 13)
        Me.lblRedisHostName.TabIndex = 0
        Me.lblRedisHostName.Text = "Redis Host Name:"
        '
        'btnSQLiteDBLocationSelection
        '
        Me.btnSQLiteDBLocationSelection.Location = New System.Drawing.Point(521, 13)
        Me.btnSQLiteDBLocationSelection.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSQLiteDBLocationSelection.Name = "btnSQLiteDBLocationSelection"
        Me.btnSQLiteDBLocationSelection.Size = New System.Drawing.Size(27, 19)
        Me.btnSQLiteDBLocationSelection.TabIndex = 52
        Me.btnSQLiteDBLocationSelection.Text = "..."
        Me.btnSQLiteDBLocationSelection.UseVisualStyleBackColor = True
        '
        'txtSQLiteDBLocation
        '
        Me.txtSQLiteDBLocation.Location = New System.Drawing.Point(154, 14)
        Me.txtSQLiteDBLocation.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSQLiteDBLocation.Name = "txtSQLiteDBLocation"
        Me.txtSQLiteDBLocation.Size = New System.Drawing.Size(364, 20)
        Me.txtSQLiteDBLocation.TabIndex = 51
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 16)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(135, 13)
        Me.Label1.TabIndex = 53
        Me.Label1.Text = "SQLite Database Location:"
        '
        'btnSaveSettings
        '
        Me.btnSaveSettings.BackColor = System.Drawing.Color.White
        Me.btnSaveSettings.Image = CType(resources.GetObject("btnSaveSettings.Image"), System.Drawing.Image)
        Me.btnSaveSettings.Location = New System.Drawing.Point(105, 484)
        Me.btnSaveSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSaveSettings.Name = "btnSaveSettings"
        Me.btnSaveSettings.Size = New System.Drawing.Size(80, 60)
        Me.btnSaveSettings.TabIndex = 1
        Me.btnSaveSettings.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.White
        Me.btnCancel.Image = CType(resources.GetObject("btnCancel.Image"), System.Drawing.Image)
        Me.btnCancel.Location = New System.Drawing.Point(626, 484)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 60)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'txtSaveSettingLocation
        '
        Me.txtSaveSettingLocation.Location = New System.Drawing.Point(98, 456)
        Me.txtSaveSettingLocation.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSaveSettingLocation.Name = "txtSaveSettingLocation"
        Me.txtSaveSettingLocation.Size = New System.Drawing.Size(571, 20)
        Me.txtSaveSettingLocation.TabIndex = 3
        '
        'lblSaveSettingsLocation
        '
        Me.lblSaveSettingsLocation.AutoSize = True
        Me.lblSaveSettingsLocation.Location = New System.Drawing.Point(8, 458)
        Me.lblSaveSettingsLocation.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblSaveSettingsLocation.Name = "lblSaveSettingsLocation"
        Me.lblSaveSettingsLocation.Size = New System.Drawing.Size(87, 13)
        Me.lblSaveSettingsLocation.TabIndex = 4
        Me.lblSaveSettingsLocation.Text = "Setting Location:"
        '
        'btnSettingsLocation
        '
        Me.btnSettingsLocation.Location = New System.Drawing.Point(673, 456)
        Me.btnSettingsLocation.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSettingsLocation.Name = "btnSettingsLocation"
        Me.btnSettingsLocation.Size = New System.Drawing.Size(32, 24)
        Me.btnSettingsLocation.TabIndex = 5
        Me.btnSettingsLocation.Text = "..."
        Me.btnSettingsLocation.UseVisualStyleBackColor = True
        '
        'btnLoadSettingXML
        '
        Me.btnLoadSettingXML.BackColor = System.Drawing.Color.White
        Me.btnLoadSettingXML.Image = CType(resources.GetObject("btnLoadSettingXML.Image"), System.Drawing.Image)
        Me.btnLoadSettingXML.Location = New System.Drawing.Point(11, 484)
        Me.btnLoadSettingXML.Margin = New System.Windows.Forms.Padding(2)
        Me.btnLoadSettingXML.Name = "btnLoadSettingXML"
        Me.btnLoadSettingXML.Size = New System.Drawing.Size(80, 60)
        Me.btnLoadSettingXML.TabIndex = 6
        Me.btnLoadSettingXML.UseVisualStyleBackColor = False
        '
        'btnSettingsOK
        '
        Me.btnSettingsOK.BackColor = System.Drawing.Color.White
        Me.btnSettingsOK.Image = CType(resources.GetObject("btnSettingsOK.Image"), System.Drawing.Image)
        Me.btnSettingsOK.Location = New System.Drawing.Point(542, 484)
        Me.btnSettingsOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSettingsOK.Name = "btnSettingsOK"
        Me.btnSettingsOK.Size = New System.Drawing.Size(80, 60)
        Me.btnSettingsOK.TabIndex = 7
        Me.btnSettingsOK.UseVisualStyleBackColor = False
        '
        'GlobalSettingsToolTip
        '
        Me.GlobalSettingsToolTip.IsBalloon = True
        '
        'O365ExtractionSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(730, 557)
        Me.Controls.Add(Me.btnSettingsOK)
        Me.Controls.Add(Me.btnLoadSettingXML)
        Me.Controls.Add(Me.btnSettingsLocation)
        Me.Controls.Add(Me.lblSaveSettingsLocation)
        Me.Controls.Add(Me.txtSaveSettingLocation)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSaveSettings)
        Me.Controls.Add(Me.tabParallelProcessingSettings)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(750, 600)
        Me.Name = "O365ExtractionSettings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Nuix Email Archive Migration Manager - Global Settings"
        Me.tabParallelProcessingSettings.ResumeLayout(False)
        Me.tabLicensingSettings.ResumeLayout(False)
        Me.grpNMSSettings.ResumeLayout(False)
        Me.grpNMSSettings.PerformLayout()
        Me.tabNuixParallelProcessingSettings.ResumeLayout(False)
        Me.grpNuixDataProcessingSettings.ResumeLayout(False)
        Me.grpNuixDataProcessingSettings.PerformLayout()
        Me.tabNuixDataSettings.ResumeLayout(False)
        Me.grpDataProcessingSettings.ResumeLayout(False)
        Me.grpDataProcessingSettings.PerformLayout()
        Me.grpDigestSettings.ResumeLayout(False)
        Me.grpDigestSettings.PerformLayout()
        Me.grpEmailDigestSettings.ResumeLayout(False)
        Me.grpEmailDigestSettings.PerformLayout()
        CType(Me.MaxDigestSize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpImageSettings.ResumeLayout(False)
        Me.grpImageSettings.PerformLayout()
        Me.grpItemContentsSettings.ResumeLayout(False)
        Me.grpItemContentsSettings.PerformLayout()
        Me.grpNamedEntitySettings.ResumeLayout(False)
        Me.grpNamedEntitySettings.PerformLayout()
        Me.grpTextIndexingSettings.ResumeLayout(False)
        Me.grpTextIndexingSettings.PerformLayout()
        Me.grpFamilyTextSettings.ResumeLayout(False)
        Me.grpFamilyTextSettings.PerformLayout()
        Me.grpDeletefileSettings.ResumeLayout(False)
        Me.grpDeletefileSettings.PerformLayout()
        Me.grpEvidenceSettings.ResumeLayout(False)
        Me.grpEvidenceSettings.PerformLayout()
        CType(Me.Maxbinarysize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabArchiveExtractionSettings.ResumeLayout(False)
        Me.grpArchiveExtractionSettings.ResumeLayout(False)
        Me.grpArchiveExtractionSettings.PerformLayout()
        CType(Me.numExtractWorkerTimeout, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpEMLExportOptions.ResumeLayout(False)
        Me.grpEMLExportOptions.PerformLayout()
        Me.grpPSTExportOptions.ResumeLayout(False)
        Me.grpPSTExportOptions.PerformLayout()
        CType(Me.numPSTExportSize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabEWSSettings.ResumeLayout(False)
        Me.grpOffice365Settings.ResumeLayout(False)
        Me.grpOffice365Settings.PerformLayout()
        Me.grpEWSDownloadControl.ResumeLayout(False)
        Me.grpEWSDownloadControl.PerformLayout()
        CType(Me.numDownloadCount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numDownloadSize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpEWSUploadControl.ResumeLayout(False)
        Me.grpEWSUploadControl.PerformLayout()
        CType(Me.numBulkUploadSize, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numEWSMaxUploadSize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpEWSThrottling.ResumeLayout(False)
        Me.grpEWSThrottling.PerformLayout()
        CType(Me.numEWSRetryIncrement, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numEWSRetryDelay, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numEWSRetryCount, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpNuixInstanceSettings.ResumeLayout(False)
        Me.grpNuixInstanceSettings.PerformLayout()
        CType(Me.numWorkerTimeout, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabDB.ResumeLayout(False)
        Me.tabDB.PerformLayout()
        Me.grpRedisConfig.ResumeLayout(False)
        Me.grpRedisConfig.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tabParallelProcessingSettings As System.Windows.Forms.TabControl
    Friend WithEvents tabNuixParallelProcessingSettings As System.Windows.Forms.TabPage
    Friend WithEvents tabEWSSettings As System.Windows.Forms.TabPage
    Friend WithEvents tabNuixDataSettings As System.Windows.Forms.TabPage
    Friend WithEvents grpNuixDataProcessingSettings As System.Windows.Forms.GroupBox
    Friend WithEvents btnExportDir As System.Windows.Forms.Button
    Friend WithEvents txtExportDir As System.Windows.Forms.TextBox
    Friend WithEvents lblExportDir As System.Windows.Forms.Label
    Friend WithEvents btnWorkerTempDir As System.Windows.Forms.Button
    Friend WithEvents txtWorkerTempDir As System.Windows.Forms.TextBox
    Friend WithEvents lblWorkerTempDir As System.Windows.Forms.Label
    Friend WithEvents btnNuixAppLocation As System.Windows.Forms.Button
    Friend WithEvents txtNuixAppLocation As System.Windows.Forms.TextBox
    Friend WithEvents lblNuixAppLocation As System.Windows.Forms.Label
    Friend WithEvents btnJavaTempDir As System.Windows.Forms.Button
    Friend WithEvents btnLogDir As System.Windows.Forms.Button
    Friend WithEvents btnCaseDirSel As System.Windows.Forms.Button
    Friend WithEvents txtNuixCaseDir As System.Windows.Forms.TextBox
    Friend WithEvents lblCaseDirectory As System.Windows.Forms.Label
    Friend WithEvents txtJavaTempDir As System.Windows.Forms.TextBox
    Friend WithEvents lblJavaTempDir As System.Windows.Forms.Label
    Friend WithEvents lblLogDirectory As System.Windows.Forms.Label
    Friend WithEvents txtLogDirectory As System.Windows.Forms.TextBox
    Friend WithEvents btnSaveSettings As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents txtSaveSettingLocation As System.Windows.Forms.TextBox
    Friend WithEvents lblSaveSettingsLocation As System.Windows.Forms.Label
    Friend WithEvents btnSettingsLocation As System.Windows.Forms.Button
    Friend WithEvents tabLicensingSettings As System.Windows.Forms.TabPage
    Friend WithEvents grpNMSSettings As System.Windows.Forms.GroupBox
    Friend WithEvents txtNMSPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtNMSUsername As System.Windows.Forms.TextBox
    Friend WithEvents txtNMSAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblNMSPassword As System.Windows.Forms.Label
    Friend WithEvents lblNMSUsername As System.Windows.Forms.Label
    Friend WithEvents lblNuixManagementServerAddress As System.Windows.Forms.Label
    Friend WithEvents grpOffice365Settings As System.Windows.Forms.GroupBox
    Friend WithEvents chkO365ApplicationImpersonation As System.Windows.Forms.CheckBox
    Friend WithEvents txtO365AdminInfo As System.Windows.Forms.TextBox
    Friend WithEvents txtO365AdminUserName As System.Windows.Forms.TextBox
    Friend WithEvents txtO365Domain As System.Windows.Forms.TextBox
    Friend WithEvents txtO365ExchangeServer As System.Windows.Forms.TextBox
    Friend WithEvents lblPSTInfoAdminInfo As System.Windows.Forms.Label
    Friend WithEvents lblPSTInfoAdminUserName As System.Windows.Forms.Label
    Friend WithEvents lblPSTInfoDomain As System.Windows.Forms.Label
    Friend WithEvents lblPSTInfoExchangeServer As System.Windows.Forms.Label
    Friend WithEvents btnLoadSettingXML As System.Windows.Forms.Button
    Friend WithEvents btnSettingsOK As System.Windows.Forms.Button
    Friend WithEvents tabDB As System.Windows.Forms.TabPage
    Friend WithEvents btnSQLiteDBLocationSelection As System.Windows.Forms.Button
    Friend WithEvents txtSQLiteDBLocation As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnNuixFilesDirSel As System.Windows.Forms.Button
    Friend WithEvents txtNuixFilesDirectory As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents grpDataProcessingSettings As System.Windows.Forms.GroupBox
    Friend WithEvents grpTextIndexingSettings As System.Windows.Forms.GroupBox
    Friend WithEvents chkEnableExactQueries As System.Windows.Forms.CheckBox
    Friend WithEvents chkUsestemming As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseStopWords As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboAnalysisLang As System.Windows.Forms.ComboBox
    Friend WithEvents grpFamilyTextSettings As System.Windows.Forms.GroupBox
    Friend WithEvents chkHideimmaterialItems As System.Windows.Forms.CheckBox
    Friend WithEvents chkCreatefamilySearch As System.Windows.Forms.CheckBox
    Friend WithEvents grpDeletefileSettings As System.Windows.Forms.GroupBox
    Friend WithEvents chkCarveFSunallocated As System.Windows.Forms.CheckBox
    Friend WithEvents chkExtractmailboxslackspace As System.Windows.Forms.CheckBox
    Friend WithEvents chkSmartprocessMSRegistry As System.Windows.Forms.CheckBox
    Friend WithEvents chkExtractEndOfFileSpace As System.Windows.Forms.CheckBox
    Friend WithEvents chkRecoverDeleteFiles As System.Windows.Forms.CheckBox
    Friend WithEvents grpEvidenceSettings As System.Windows.Forms.GroupBox
    Friend WithEvents Maxbinarysize As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblMaxbinarysize As System.Windows.Forms.Label
    Friend WithEvents chkStorebinary As System.Windows.Forms.CheckBox
    Friend WithEvents chkCalculateAuditedSize As System.Windows.Forms.CheckBox
    Friend WithEvents chkReuseEvidenceStores As System.Windows.Forms.CheckBox
    Friend WithEvents lblTraversal As System.Windows.Forms.Label
    Friend WithEvents cboTraversal As System.Windows.Forms.ComboBox
    Friend WithEvents chkCalculateProcessingSize As System.Windows.Forms.CheckBox
    Friend WithEvents chkPerformItemIdentification As System.Windows.Forms.CheckBox
    Friend WithEvents grpDigestSettings As System.Windows.Forms.GroupBox
    Friend WithEvents grpEmailDigestSettings As System.Windows.Forms.GroupBox
    Friend WithEvents chkIncludeItemDate As System.Windows.Forms.CheckBox
    Friend WithEvents chkIncludeBCC As System.Windows.Forms.CheckBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents MaxDigestSize As System.Windows.Forms.NumericUpDown
    Friend WithEvents chkSSDeep As System.Windows.Forms.CheckBox
    Friend WithEvents chkSHA256 As System.Windows.Forms.CheckBox
    Friend WithEvents chkSHA1 As System.Windows.Forms.CheckBox
    Friend WithEvents chkMD5Digest As System.Windows.Forms.CheckBox
    Friend WithEvents grpImageSettings As System.Windows.Forms.GroupBox
    Friend WithEvents chkPerformimagecolouranalysis As System.Windows.Forms.CheckBox
    Friend WithEvents chkGeneratethumbnailsforimagedata As System.Windows.Forms.CheckBox
    Friend WithEvents grpItemContentsSettings As System.Windows.Forms.GroupBox
    Friend WithEvents grpNamedEntitySettings As System.Windows.Forms.GroupBox
    Friend WithEvents chkExtractNamedEntitiesfromProperties As System.Windows.Forms.CheckBox
    Friend WithEvents chkIncludeTextStrippedItems As System.Windows.Forms.CheckBox
    Friend WithEvents chkExtractNamedEntities As System.Windows.Forms.CheckBox
    Friend WithEvents chkEnableTextSummarisation As System.Windows.Forms.CheckBox
    Friend WithEvents chkEnableNearDuplicates As System.Windows.Forms.CheckBox
    Friend WithEvents chkProcesstext As System.Windows.Forms.CheckBox
    Friend WithEvents lblDigestToCompute As System.Windows.Forms.Label
    Friend WithEvents tabArchiveExtractionSettings As System.Windows.Forms.TabPage
    Friend WithEvents lblSourceType As System.Windows.Forms.Label
    Friend WithEvents cboSourceType As System.Windows.Forms.ComboBox
    Friend WithEvents grpArchiveExtractionSettings As System.Windows.Forms.GroupBox
    Friend WithEvents txtMemoryPerNuixInstance As System.Windows.Forms.TextBox
    Friend WithEvents txtMemoryPerWorker As System.Windows.Forms.TextBox
    Friend WithEvents lbNuixAppMemory As System.Windows.Forms.Label
    Friend WithEvents lblAvailableRAMAfterNuix As System.Windows.Forms.Label
    Friend WithEvents lblAvailableAfterNuix As System.Windows.Forms.Label
    Friend WithEvents lblNuixInstances As System.Windows.Forms.Label
    Friend WithEvents cboNumberOfNuixInstances As System.Windows.Forms.ComboBox
    Friend WithEvents lblMemoryPerWorker As System.Windows.Forms.Label
    Friend WithEvents cboNuixWorkers As System.Windows.Forms.ComboBox
    Friend WithEvents lblNumberOfNuixWorkers As System.Windows.Forms.Label
    Friend WithEvents lblAvailableRAM As System.Windows.Forms.Label
    Friend WithEvents lblRAM As System.Windows.Forms.Label
    Friend WithEvents lblAvailableSystemMemory As System.Windows.Forms.Label
    Friend WithEvents lblTotalSystemMemory As System.Windows.Forms.Label
    Friend WithEvents grpEWSThrottling As System.Windows.Forms.GroupBox
    Friend WithEvents lblRetryIncrement As System.Windows.Forms.Label
    Friend WithEvents lblRetryDelay As System.Windows.Forms.Label
    Friend WithEvents lblRetryCount As System.Windows.Forms.Label
    Friend WithEvents grpNuixInstanceSettings As System.Windows.Forms.GroupBox
    Friend WithEvents txtO365NuixAppMemory As System.Windows.Forms.TextBox
    Friend WithEvents txtO365MemoryPerWorker As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblO365AvailableMemoryAfterNuix As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cboO365NumberOfNuixInstances As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cboO365NumberOfNuixWorkers As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lblO365AvailMemory As System.Windows.Forms.Label
    Friend WithEvents lblO365SystemMemory As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents grpRedisConfig As System.Windows.Forms.GroupBox
    Friend WithEvents txtRedisAuth As System.Windows.Forms.TextBox
    Friend WithEvents lblRedis As System.Windows.Forms.Label
    Friend WithEvents txtRedisPort As System.Windows.Forms.TextBox
    Friend WithEvents lblRedisPort As System.Windows.Forms.Label
    Friend WithEvents txtRedisHostName As System.Windows.Forms.TextBox
    Friend WithEvents lblRedisHostName As System.Windows.Forms.Label
    Friend WithEvents lblNMSPort As System.Windows.Forms.Label
    Friend WithEvents txtNMSPort As System.Windows.Forms.TextBox
    Friend WithEvents grpEWSDownloadControl As System.Windows.Forms.GroupBox
    Friend WithEvents grpEWSUploadControl As System.Windows.Forms.GroupBox
    Friend WithEvents lblBulkUploadsze As System.Windows.Forms.Label
    Friend WithEvents numBulkUploadSize As System.Windows.Forms.NumericUpDown
    Friend WithEvents chkEnableBulkUpload As System.Windows.Forms.CheckBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtO365RemovePathPrefix As System.Windows.Forms.TextBox
    Friend WithEvents lblRemovePathPrefix As System.Windows.Forms.Label
    Friend WithEvents lblMaxItemUploadSize As System.Windows.Forms.Label
    Friend WithEvents numEWSMaxUploadSize As System.Windows.Forms.NumericUpDown
    Friend WithEvents chkCollaborativePrefetching As System.Windows.Forms.CheckBox
    Friend WithEvents chkEnablePrefetch As System.Windows.Forms.CheckBox
    Friend WithEvents numDownloadCount As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblMaxDownloadCount As System.Windows.Forms.Label
    Friend WithEvents numDownloadSize As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblMaxMessageDownloadSize As System.Windows.Forms.Label
    Friend WithEvents grpPSTExportOptions As System.Windows.Forms.GroupBox
    Friend WithEvents chkPSTAddDistributionListMetadata As System.Windows.Forms.CheckBox
    Friend WithEvents lblPSTExportSize As System.Windows.Forms.Label
    Friend WithEvents numPSTExportSize As System.Windows.Forms.NumericUpDown
    Friend WithEvents grpEMLExportOptions As System.Windows.Forms.GroupBox
    Friend WithEvents chkEMLAddDistributionListMetadata As System.Windows.Forms.CheckBox
    Friend WithEvents GlobalSettingsToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents numWorkerTimeout As System.Windows.Forms.NumericUpDown
    Friend WithEvents numEWSRetryIncrement As System.Windows.Forms.NumericUpDown
    Friend WithEvents numEWSRetryDelay As System.Windows.Forms.NumericUpDown
    Friend WithEvents numEWSRetryCount As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblExtractWorkerTimeout As System.Windows.Forms.Label
    Friend WithEvents numExtractWorkerTimeout As System.Windows.Forms.NumericUpDown
    Friend WithEvents chkEnableMailboxSlackspace As System.Windows.Forms.CheckBox
    Friend WithEvents lblPSTConsolidation As System.Windows.Forms.Label
    Friend WithEvents cboPSTConsolidation As System.Windows.Forms.ComboBox
End Class
