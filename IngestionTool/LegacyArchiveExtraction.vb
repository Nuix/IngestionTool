
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Xml.Linq
Imports System
Imports System.IO.StreamReader
Imports System.Collections
Imports System.Threading
Imports System.Data.SQLite
Imports System.Data.SqlClient



Public Class LegacyArchiveExtraction
    Public pbNuixSettingsLoaded As Boolean
    Public psSettingsFile As String

    Public pbSettingsLoaded As Boolean

    Public Property ImageList As ImageList
    Public plstAXSOneSIS As List(Of String)
    Public plstSourceFolders As List(Of String)
    Public plstSourceFiles As List(Of String)
    Public plstCenteraClipLists As List(Of String)

    Public psNuixProcessingFilesDir As String
    Public psNuixAppLocation As String
    Public psNMSSourceType As String
    Public psNMSAddress As String
    Public psNMSPort As String
    Public psNMSUserName As String
    Public psNMSAdminInfo As String
    Public psLicenseType As String
    Public psNumberOfWorkers As String
    Public psMemoryPerWorker As String
    Public psNuixLogDir As String
    Public psJavaTempDir As String
    Public piWorkerTimeout As Integer
    Public psNuixFilesDir As String
    Public psWorkerTempDir As String
    Public psNuixAppMemory As String
    Public psSQLiteLocation As String
    Public psExportDir As String
    Public psCNDExportDir As String
    Public psNuixCaseDir As String
    Public psPSTExportSize As String
    Public pbPSTAddDistributionListMetadata As Boolean
    Public pbEMLAddDistributionListMetadata As Boolean
    Public pbExtractFromSlackSpace As Boolean
    Private RunExtractions As System.Threading.Thread
    Private SQLiteDBReadThread As System.Threading.Thread

    Private Sub cboArchiveType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboArchiveType.SelectedIndexChanged
        Dim drives As System.Collections.ObjectModel.ReadOnlyCollection(Of IO.DriveInfo)
        Dim Tnode As TreeNode
        Dim rootDir As String

        grpWSSControl.Hide()
        radMailboxArchive.Enabled = True
        radJournalArchive.Enabled = True

        grpArchiveType.Enabled = True
        grpOutputFormat.Enabled = False
        grpExtractionType.Enabled = False
        radMailboxArchive.Checked = False
        radJournalArchive.Checked = False
        radCentera.Checked = False
        lblTotalBatchSize.Text = "0"

        treeViewFolders.Nodes.Clear()
        drives = My.Computer.FileSystem.Drives
        For i As Integer = 0 To drives.Count - 1
            If Not drives(i).IsReady Then
                Continue For
            End If
            Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
            rootDir = drives(i).Name
            AddAllFolders(Tnode, rootDir)
        Next

        plstSourceFolders.Clear()
        plstSourceFiles.Clear()
        txtSearchTermCSV.Text = vbNullString
        txtCustodianListCSV.Text = vbNullString
        txtSQLDBName.Text = vbNullString
        txtSQLHostName.Text = vbNullString
        txtSQLInfo.Text = vbNullString
        txtSQLPortNumber.Text = vbNullString
        txtSQLUserName.Text = vbNullString
        txtSearchTermCSV.Text = vbNullString
        txtMappingCSV.Text = vbNullString
        radExcludeItemsTrue.Checked = True
        radExcludeItemsFalse.Checked = False
        radVerboseTrue.Checked = False
        radVerboseFalse.Checked = True

        setDateTimePickerBlank(dateFromDate)
        setDateTimePickerBlank(dateToDate)

        btnExpandCollapse.Text = "-"
        grdArchiveExtractionBatch.Top = 430

        lblSQLPWCheck.Text = "X"
        lblSQLPWCheck.ForeColor = Color.Red
        lblHostCheck.Text = "X"
        lblHostCheck.ForeColor = Color.Red
        lblSQLUserNameCheck.Text = "X"
        lblSQLUserNameCheck.ForeColor = Color.Red
        lblDBNameCheck.Text = "X"
        lblDBNameCheck.ForeColor = Color.Red

        Select Case cboArchiveType.Text
            Case "Daegis AXS-One"
                radCentera.Checked = False
                radCentera.Enabled = False
                grpAXSOneSISSettings.Show()
                grpAXSOneSettings.Enabled = True
                grpAXSOneSettings.Show()
                grpAXSOneSISSettings.Enabled = False
                grpAXSOneSettings.Enabled = True
                grpAXSOneSISSettings.Enabled = True
                grpEVSettings.Enabled = False
                grpEMCSettings.Enabled = False
                grpEASSettings.Enabled = False
                grpSQLArchiveSettings.Hide()
                grpEVSettings.Hide()
                grpEMCSettings.Hide()
                grpEASSettings.Hide()
                btnExpandCollapse.Show()
                radJournalArchive.Checked = False
                radMailboxArchive.Checked = True

                treeAXSOneSis.Nodes.Clear()
                Dim Sisnode As TreeNode
                drives = My.Computer.FileSystem.Drives
                rootDir = String.Empty
                For i As Integer = 0 To drives.Count - 1
                    If Not drives(i).IsReady Then
                        Continue For
                    End If
                    Sisnode = treeAXSOneSis.Nodes.Add(drives(i).ToString)
                    rootDir = drives(i).Name
                    AddAllFoldersAndFiles(Sisnode, rootDir)
                Next

                txtSQLDBName.Text = ""
                btnExpandCollapse.Text = "-"
                grpSQLArchiveSettings.Height = 330
                grpSourceInformation.Height = 330
                grpAXSOneSISSettings.Height = 330
                btnAddBatchToGrid.Show()
                grpSourceLocation.Enabled = True
                grpAXSOneSettings.Show()
                grpAXSOneSettings.BringToFront()
                grpAXSOneSettings.Height = 330
                grpAXSOneSettings.Top = 90
                lblFromDate.Show()
                lblToDate.Show()
                dateFromDate.Show()
                dateToDate.Show()
                grpWSSControl.Hide()

            Case "HP/Autonomy EAS"
                radCentera.Enabled = True

                grpAXSOneSISSettings.Hide()
                grpAXSOneSettings.Enabled = False
                grpAXSOneSISSettings.Enabled = False
                grpEVSettings.Enabled = False
                grpEMCSettings.Enabled = False
                grpEASSettings.Enabled = True

                txtSQLDBName.Text = ""
                grpSQLArchiveSettings.Show()
                cboSecurityType.Text = "SQLServer Authentication"
                cboSecurityType.Enabled = True
                grpEMCSettings.Hide()
                grpEVSettings.Hide()
                grpEASSettings.Show()
                grpAXSOneSettings.Hide()
                grpAXSOneSettings.Enabled = False
                grpAXSOneSISSettings.Enabled = False
                btnExpandCollapse.Show()
                radJournalArchive.Checked = True
                radMailboxArchive.Checked = False
                btnExpandCollapse.Text = "-"
                grpSQLArchiveSettings.Height = 330
                grpSourceInformation.Height = 330
                grpEASSettings.Height = 330
                btnAddBatchToGrid.Show()
                lblFromDate.Show()
                lblToDate.Show()
                dateFromDate.Show()
                dateToDate.Show()
            Case "EMC EmailXtender"
                radCentera.Enabled = True

                grpAXSOneSISSettings.Hide()
                grpAXSOneSettings.Enabled = False
                grpAXSOneSISSettings.Enabled = False
                grpEVSettings.Enabled = False
                grpEMCSettings.Enabled = True
                grpEASSettings.Enabled = False

                txtSQLDBName.Text = ""
                grpSQLArchiveSettings.Show()
                cboSecurityType.Text = "SQLServer Authentication"
                cboSecurityType.Enabled = True

                grpEMCSettings.Show()
                grpAXSOneSISSettings.Hide()
                grpEVSettings.Hide()
                grpAXSOneSettings.Hide()
                grpEASSettings.Hide()
                btnExpandCollapse.Show()
                radJournalArchive.Checked = True
                radMailboxArchive.Checked = False
                btnExpandCollapse.Text = "-"
                grpSQLArchiveSettings.Height = 330
                grpSourceInformation.Height = 330
                grpEMCSettings.Height = 330
                btnAddBatchToGrid.Show()
                lblFromDate.Show()
                lblToDate.Show()
                dateFromDate.Show()
                dateToDate.Show()

            Case "EMC SourceOne"
                radCentera.Enabled = True

                grpAXSOneSISSettings.Hide()
                grpAXSOneSettings.Enabled = False
                grpAXSOneSISSettings.Enabled = False
                grpEVSettings.Enabled = False
                grpEMCSettings.Enabled = True
                grpEASSettings.Enabled = False
                txtSQLDBName.Text = ""
                grpSQLArchiveSettings.Show()
                cboSecurityType.Text = "SQLServer Authentication"
                cboSecurityType.Enabled = True
                grpEMCSettings.Show()
                grpAXSOneSISSettings.Hide()
                grpEVSettings.Hide()
                grpAXSOneSettings.Hide()
                grpEASSettings.Hide()
                btnExpandCollapse.Show()
                btnExpandCollapse.Text = "-"
                grpSQLArchiveSettings.Height = 330
                grpSourceInformation.Height = 330
                grpEMCSettings.Height = 330
                btnAddBatchToGrid.Show()
                radJournalArchive.Checked = True
                radMailboxArchive.Checked = False
            Case "Veritas Enterprise Vault"
                radCentera.Enabled = True

                grpAXSOneSISSettings.Hide()
                grpAXSOneSettings.Enabled = False
                grpAXSOneSISSettings.Enabled = False
                grpEVSettings.Enabled = True
                grpEMCSettings.Enabled = False
                grpEASSettings.Enabled = False

                grpSQLArchiveSettings.Show()
                cboSecurityType.Text = "SQLServer Authentication"
                cboSecurityType.Enabled = True

                grpEMCSettings.Hide()
                grpAXSOneSISSettings.Hide()
                grpEVSettings.Show()
                grpEASSettings.Visible = False
                grpAXSOneSettings.Hide()
                grpEASSettings.Hide()
                grpEVSettings.BringToFront()
                grpEVSettings.Height = 330
                grpEVSettings.Top = 90
                btnExpandCollapse.Show()
                btnExpandCollapse.Text = "-"
                grpSQLArchiveSettings.Height = 330
                grpSourceInformation.Height = 330
                grpEVSettings.Height = 330

                btnAddBatchToGrid.Show()
                radJournalArchive.Checked = False
                radMailboxArchive.Checked = False
                grpEVSettings.Height = 330
                grpSQLArchiveSettings.Height = 330
                grpSourceInformation.Height = 330
                btnExpandCollapse.Text = "-"
                btnAddBatchToGrid.Show()
                chkSkipAdditionalSQLLookup.Checked = True
                chkSkipVaultStorePartitionErrors.Checked = False

                txtSQLDBName.Text = "EnterpriseVaultDirectory"
                grpWSSControl.Hide()
            Case Else
                radCentera.Enabled = True

                grpAXSOneSettings.Enabled = False
                grpAXSOneSISSettings.Enabled = False
                grpEVSettings.Enabled = False
                grpEMCSettings.Enabled = False
                grpEASSettings.Enabled = False
                grpAXSOneSISSettings.Hide()
                grpAXSOneSettings.Hide()
                btnExpandCollapse.Hide()
        End Select

    End Sub

    Private Sub setDateTimePickerBlank(ByVal dateTimePicker As DateTimePicker)
        dateTimePicker.Visible = True
        dateTimePicker.Format = DateTimePickerFormat.Custom
        dateTimePicker.CustomFormat = "    "
    End Sub

    Private Sub CheckChildNodes(ByVal iNode As TreeNode)
        Try
            UnCheckParentNodes(iNode)
            For Each sNode As TreeNode In iNode.Nodes
                sNode.Checked = iNode.Checked
                CheckChildNodes(sNode)
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub UnCheckChildNodes(ByVal iNode As TreeNode)
        Try

            For Each sNode As TreeNode In iNode.Nodes
                sNode.Checked = iNode.Checked = False
                UnCheckChildNodes(sNode)
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub LegacyArchiveExtraction_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim msgboxReturn As DialogResult
        Dim sMachineName As String

        Dim common As New Common

        If eMailArchiveMigrationManager.NuixSettingsFile = vbNullString Then
            msgboxReturn = MessageBox.Show("Nuix Email Archive Migration Manager has detected that the Global Settings have not been loaded.  Please load the settings now. ", "Global Settings Not Available", MessageBoxButtons.YesNo)
            If (msgboxReturn = vbYes) Then
                Dim NuixSettings As O365ExtractionSettings
                NuixSettings = New O365ExtractionSettings
                NuixSettings.ShowDialog()
            Else
                Me.Close()

                Exit Sub
            End If
        Else
            psSettingsFile = eMailArchiveMigrationManager.NuixSettingsFile
        End If

        If psSettingsFile = vbNullString Then
            Me.Close()
            Exit Sub
        Else
            grpAXSOneSISSettings.Hide()
            grpSQLArchiveSettings.Hide()
            grpEMCSettings.Hide()
            grpEVSettings.Hide()
            grpEASSettings.Hide()
            grpAXSOneSettings.Hide()
            'txtDomain.Hide()
            'lblDomain.Hide()
            'lblDomainCheck.Hide()

            grpArchiveType.Enabled = False
            grpExtractionType.Enabled = False
            grpOutputFormat.Enabled = False

            lblCustodianListCSV.Hide()
            txtCustodianListCSV.Hide()
            btnCustodianListCSVChooser.Hide()

            psNuixProcessingFilesDir = eMailArchiveMigrationManager.NuixFilesDir & "\Archive Extraction\"
            System.IO.Directory.CreateDirectory(psNuixProcessingFilesDir)

            psNuixAppLocation = eMailArchiveMigrationManager.NuixAppLocation
            psNMSSourceType = eMailArchiveMigrationManager.NMSSourceType

            psNMSAddress = eMailArchiveMigrationManager.NuixNMSAddress
            psNMSPort = eMailArchiveMigrationManager.NuixNMSPort
            psNMSUserName = eMailArchiveMigrationManager.NuixNMSUserName
            psNMSAdminInfo = eMailArchiveMigrationManager.NuixNMSAdminInfo
            psLicenseType = eMailArchiveMigrationManager.NMSSourceType
            psNumberOfWorkers = eMailArchiveMigrationManager.NuixWorkers
            psMemoryPerWorker = eMailArchiveMigrationManager.NuixWorkerMemory
            psNuixLogDir = eMailArchiveMigrationManager.NuixLogDir
            psJavaTempDir = eMailArchiveMigrationManager.NuixJavaTempDir
            psNuixFilesDir = eMailArchiveMigrationManager.NuixFilesDir
            psWorkerTempDir = eMailArchiveMigrationManager.NuixWorkerTempDir
            psNuixAppMemory = eMailArchiveMigrationManager.NuixAppMemory
            psSQLiteLocation = eMailArchiveMigrationManager.SQLiteDBLocation
            psExportDir = eMailArchiveMigrationManager.NuixExportDir
            psNuixAppLocation = eMailArchiveMigrationManager.NuixAppLocation
            psNMSSourceType = eMailArchiveMigrationManager.NMSSourceType
            psNuixCaseDir = eMailArchiveMigrationManager.NuixCaseDir
            piWorkerTimeout = eMailArchiveMigrationManager.WorkerTimeout
            psPSTExportSize = eMailArchiveMigrationManager.PSTExportSize
            pbEMLAddDistributionListMetadata = eMailArchiveMigrationManager.EMLAddDistributionListMetadata
            pbPSTAddDistributionListMetadata = eMailArchiveMigrationManager.PSTAddDistributionListMetadata
            pbExtractFromSlackSpace = eMailArchiveMigrationManager.extractFromSlackSpace
            sMachineName = System.Net.Dns.GetHostName

            psIngestionLogFile = psNuixLogDir & "\" & "Archive Extraction Log - " & sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".log"

            common.Logger(psIngestionLogFile, "Nuix Email Archive Migration - Office 365 Ingestion Log")
            common.Logger(psIngestionLogFile, "Office 365 Exchange Server - " & eMailArchiveMigrationManager.O365ExchangeServer)
            common.Logger(psIngestionLogFile, "Office 365 Domain - " & eMailArchiveMigrationManager.O365Domain)
            common.Logger(psIngestionLogFile, "Office 365 Admin Username - " & eMailArchiveMigrationManager.O365AdminUserName)
            common.Logger(psIngestionLogFile, "Office 365 Admin Info - " & eMailArchiveMigrationManager.O365AdminInfo)
            common.Logger(psIngestionLogFile, "Office 365 Application Impersonation - " & eMailArchiveMigrationManager.O365ApplicationImpersonation.ToString.ToLower)
            common.Logger(psIngestionLogFile, "Office 365 Retry Count - " & eMailArchiveMigrationManager.O365RetryCount)
            common.Logger(psIngestionLogFile, "Office 365 Retry Delay - " & eMailArchiveMigrationManager.O365RetryDelay)
            common.Logger(psIngestionLogFile, "Office 365 Retry Increment - " & eMailArchiveMigrationManager.O365RetryIncrement)
            common.Logger(psIngestionLogFile, "Office 365 File Path Trimming - " & eMailArchiveMigrationManager.O365FilePathTrimming)
            common.Logger(psIngestionLogFile, "Office 365 Nuix App Memory - " & eMailArchiveMigrationManager.O365NuixAppMemory)
            common.Logger(psIngestionLogFile, "Office 365 Nuix Instances - " & eMailArchiveMigrationManager.O365NuixInstances)
            common.Logger(psIngestionLogFile, "Office 365 Nuix Memory Perworker - " & eMailArchiveMigrationManager.O365MemoryPerWorker)
            common.Logger(psIngestionLogFile, "Office 365 Nuix Workers - " & eMailArchiveMigrationManager.O365NumberOfNuixWorkers)
            common.Logger(psIngestionLogFile, "Office 365 File Path Trimming - " & eMailArchiveMigrationManager.NMSSourceType)
            common.Logger(psIngestionLogFile, "Nuix NMS Location - " & eMailArchiveMigrationManager.NuixNMSUserName)
            common.Logger(psIngestionLogFile, "Nuix NMS Username - " & eMailArchiveMigrationManager.NuixNMSUserName)
            common.Logger(psIngestionLogFile, "Nuix NMS Admininfo - " & eMailArchiveMigrationManager.NuixNMSAdminInfo)
            common.Logger(psIngestionLogFile, "Nuix Extraction Instances - " & eMailArchiveMigrationManager.NuixInstances)
            common.Logger(psIngestionLogFile, "Nuix Extraction App Memory- " & eMailArchiveMigrationManager.NuixAppMemory)
            common.Logger(psIngestionLogFile, "Nuix Extraction Workers - " & eMailArchiveMigrationManager.NuixWorkers)
            common.Logger(psIngestionLogFile, "Nuix Extraction Worker Memory - " & eMailArchiveMigrationManager.piNuixMemoryPerWorker)
            common.Logger(psIngestionLogFile, "Settings File location - " & eMailArchiveMigrationManager.NuixSettingsFile)
            common.Logger(psIngestionLogFile, "SQLite DB Location - " & eMailArchiveMigrationManager.SQLiteDBLocation)

            cboExpandDLLocation.Text = """" & "Expanded-DL" & """"
            chkSkipAdditionalSQLLookup.Checked = False
            chkSkipVaultStorePartitionErrors.Checked = False
            Me.Height = 800
            grpEVSettings.Height = 20
            grpSQLArchiveSettings.Height = 20
            grpSourceInformation.Height = 330
            grpAXSOneSISSettings.Height = 20
            btnExpandCollapse.Text = "-"
            grpWSSControl.Hide()
            btnAddBatchToGrid.Hide()

            radExcludeItemsTrue.Checked = True
            radVerboseFalse.Checked = True
            chkEmail.Checked = True
            chkRSSFeed.Checked = True
            chkCalendar.Checked = True
            chkContact.Checked = True

            radAddFolder.Checked = True

            plstAXSOneSIS = New List(Of String)
            plstSourceFiles = New List(Of String)
            plstSourceFolders = New List(Of String)
            plstCenteraClipLists = New List(Of String)
            setDateTimePickerBlank(dateFromDate)
            setDateTimePickerBlank(dateToDate)
            chkAFEX.Checked = True
            chkAFSYS.Checked = True
            chkAFPST.Checked = True
        End If
    End Sub

    Private Sub PopulateTreeView(ByVal dir As String, ByVal parentNode As TreeNode)
        Dim folder As String = String.Empty
        Try
            Dim folders() As String = IO.Directory.GetDirectories(dir)
            If folders.Length <> 0 Then
                Dim childNode As TreeNode = Nothing
                For Each folder In folders
                    childNode = New TreeNode(folder)
                    parentNode.Nodes.Add(childNode)
                Next
            End If
        Catch ex As UnauthorizedAccessException
            parentNode.Nodes.Add(folder & ": Access Denied")
        End Try
    End Sub

    Private Sub radUser_CheckedChanged(sender As Object, e As EventArgs)

        If radUser.Checked = True Then
            txtCustodianListCSV.Enabled = True
            txtCustodianListCSV.Text = ""
            btnCustodianListCSVChooser.Enabled = False
        End If

    End Sub

    Private Sub radLightspeed_CheckedChanged(sender As Object, e As EventArgs)

        If radUser.Checked = True Then
            txtCustodianListCSV.Enabled = True
            txtCustodianListCSV.Text = ""
            btnCustodianListCSVChooser.Enabled = False
        End If

        If ((radMailboxArchive.Checked = True) And (radUser.Checked = True)) Then
            grpSourceInformation.Enabled = False
        Else
            grpSourceInformation.Enabled = True

        End If

    End Sub

    Private Sub btnShowSettings_Click(sender As Object, e As EventArgs) Handles btnShowSettings.Click
        Dim SettingsForm As O365ExtractionSettings
        psSettingsFile = O365ExtractionSettings.NuixSettingsFile
        If psSettingsFile = vbNullString Then

            SettingsForm = New O365ExtractionSettings

            SettingsForm.ShowDialog()
            SettingsForm.NuixSettingsFile = psSettingsFile
        Else
            SettingsForm = New O365ExtractionSettings
            SettingsForm.NuixSettingsFile = psSettingsFile

            SettingsForm.ShowDialog()
            eMailArchiveMigrationManager.NuixSettingsFile = psSettingsFile
        End If
        psSettingsFile = O365ExtractionSettings.NuixSettingsFile

    End Sub

    Private Sub btnBuildFolder_Click(sender As Object, e As EventArgs)
        Dim FolderBrowserDialog As System.Windows.Forms.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog
        FolderBrowserDialog.Description = "Select a root directory for this sample to load the treeview."
        If FolderBrowserDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            TreeViewFolderBuilder.Initialize(treeViewFolders, FolderBrowserDialog.SelectedPath, "Folder_Closed", "Folder_Open", ContextMenuStrip1)
        End If
        FolderBrowserDialog.Dispose()
    End Sub

    Private Sub btnExpandCollapse_Click(sender As Object, e As EventArgs) Handles btnExpandCollapse.Click
        If btnExpandCollapse.Text = "-" Then
            grpEVSettings.Height = 20
            grpSQLArchiveSettings.Height = 20
            grpSourceInformation.Height = 20
            grpAXSOneSISSettings.Height = 20
            grpAXSOneSettings.Height = 20
            grpEASSettings.Height = 20
            grpEMCSettings.Height = 20
            grpWSSControl.Height = 20
            btnExpandCollapse.Text = "+"
            grdArchiveExtractionBatch.Top = 130
            btnAddBatchToGrid.Hide()
        ElseIf btnExpandCollapse.Text = "+" Then
            grpEVSettings.Height = 330
            grpEASSettings.Height = 330
            grpEMCSettings.Height = 330
            grpSQLArchiveSettings.Height = 330
            grpSourceInformation.Height = 330
            grpAXSOneSISSettings.Height = 330
            grpWSSControl.Height = 330
            btnExpandCollapse.Text = "-"
            grdArchiveExtractionBatch.Top = 430
            btnAddBatchToGrid.Show()
        End If

    End Sub

    Private Sub treeViewFolders_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles treeViewFolders.AfterCheck
        Dim dblTotalFolderSize As Double
        Dim bStatus As Boolean
        Dim dblCurrentBatchSize As Double
        Dim dblTotalBatchSize As Double

        If radAddFile.Checked = True Then
            If e.Node.Checked = True Then
                If e.Node.Nodes.Count = 1 AndAlso e.Node.Nodes(0).Text = "{child}" Then
                    e.Node.Nodes.Clear()
                    AddAllFoldersAndFiles(e.Node, CStr(e.Node.Tag))
                    e.Node.SelectedImageIndex = 1
                End If
                dblCurrentBatchSize = CDbl(lblTotalBatchSize.Text)
                CheckChildFileNodes(e.Node, e.Node.Checked, dblCurrentBatchSize, dblTotalBatchSize)
                'attribute = File.GetAttributes(e.Node.Tag)
                'If attribute.HasFlag(FileAttributes.Directory) Then

                '             MessageBox.Show("You can only select files to add as source when Files is selected as the source type above.")
                '              e.Node.Checked = False
                '               Exit Sub
                '            End If
                'fInfo = New FileInfo(e.Node.Tag)
                'dblFileSize = fInfo.Length

                'dblTotalBatchSize = dblCurrentBatchSize + dblFileSize
                Me.lblTotalBatchSize.Text = dblTotalBatchSize.ToString("N0")
                'plstSourceFiles.Add(e.Node.Tag)
            Else
                If e.Node.Nodes.Count = 1 AndAlso e.Node.Nodes(0).Text = "{child}" Then
                    e.Node.Nodes.Clear()
                    AddAllFoldersAndFiles(e.Node, CStr(e.Node.Tag))
                    e.Node.SelectedImageIndex = 1
                End If
                dblCurrentBatchSize = CDbl(lblTotalBatchSize.Text)
                CheckChildFileNodes(e.Node, e.Node.Checked, dblCurrentBatchSize, dblTotalBatchSize)
                If chkComputeBatchSize.Checked = True Then
                    Me.lblTotalBatchSize.Text = dblTotalBatchSize.ToString("N0")
                End If

                '                fInfo = New FileInfo(e.Node.Tag)
                '                If File.Exists(fInfo.FullName) Then
                ' dblFileSize = fInfo.Length
                ' dblTotalBatchSize = dblCurrentBatchSize - dblFileSize
                'End If

                plstSourceFiles.Remove(e.Node.Tag)
            End If
        ElseIf radAddFolder.Checked = True Then
            If e.Node.Nodes.Count = 1 AndAlso e.Node.Nodes(0).Text = "{child}" Then
                e.Node.Nodes.Clear()
                AddAllFoldersAndFiles(e.Node, CStr(e.Node.Tag))
                e.Node.SelectedImageIndex = 1
            End If
            dblCurrentBatchSize = CDbl(lblTotalBatchSize.Text)

            If chkComputeBatchSize.Checked = True Then
                If e.Node.Checked = True Then
                    bStatus = blnComputeFolderSize(e.Node.Tag, dblTotalFolderSize)
                    dblTotalBatchSize = dblCurrentBatchSize + dblTotalFolderSize

                    Me.lblTotalBatchSize.Text = dblTotalBatchSize.ToString("N0")
                    plstSourceFolders.Add(e.Node.Tag)

                Else
                    bStatus = blnComputeFolderSize(e.Node.Tag, dblTotalFolderSize)
                    dblTotalBatchSize = dblCurrentBatchSize - dblTotalFolderSize
                    Me.lblTotalBatchSize.Text = dblTotalBatchSize.ToString("N0")
                    plstSourceFolders.Remove(e.Node.Tag)
                End If
            Else
                If e.Node.Nodes.Count = 1 AndAlso e.Node.Nodes(0).Text = "{child}" Then
                    e.Node.Nodes.Clear()
                    AddAllFoldersAndFiles(e.Node, CStr(e.Node.Tag))
                    e.Node.SelectedImageIndex = 1
                End If
                If e.Node.Checked = True Then
                    plstSourceFolders.Add(e.Node.Tag)

                Else
                    plstSourceFolders.Remove(e.Node.Tag)
                End If
            End If
        ElseIf radCentera.Checked = True Then
            If e.Node.Checked = True Then
                If File.Exists(e.Node.Tag) Then
                    If plstCenteraClipLists.Count >= 1 Then
                        Dim sCurrentlySelected As String
                        Dim CurrentTreeNode() As TreeNode
                        sCurrentlySelected = plstCenteraClipLists(0)
                        CurrentTreeNode = treeViewFolders.Nodes.Find(sCurrentlySelected, True)
                        For Each TreeNode In CurrentTreeNode
                            TreeNode.Checked = False
                        Next


                        plstCenteraClipLists.Remove(sCurrentlySelected)
                        plstCenteraClipLists.Add(e.Node.Tag)

                    Else
                        plstCenteraClipLists.Add(e.Node.Tag)
                    End If

                End If

            Else
                plstCenteraClipLists.Remove(e.Node.Tag)
            End If
        End If
    End Sub

    Private Sub treeViewFolders_AfterExpand(sender As Object, e As TreeViewEventArgs) Handles treeViewFolders.AfterExpand
        If e.Node.Nodes.Count = 1 AndAlso e.Node.Nodes(0).Text = "{child}" Then
            e.Node.Nodes.Clear()
            AddAllFoldersAndFiles(e.Node, CStr(e.Node.Tag))
            e.Node.SelectedImageIndex = 1
        End If
    End Sub
    Sub CheckChildFileNodes(ByVal parent As TreeNode, checked As Boolean, ByVal dblCurrentBatchSize As Double, ByRef dblTotalBatchSize As Double)


        Dim dblFileSize As Double
        Dim attribute As FileAttributes
        Dim fInfo As FileInfo

        Dim bFoundFile As Boolean

        bFoundFile = False

        For Each child As TreeNode In parent.Nodes

            Try
                dblCurrentBatchSize = CDbl(Me.lblTotalBatchSize.Text)
                attribute = File.GetAttributes(child.Tag)
                If attribute.HasFlag(FileAttributes.Directory) Then
                    child.Checked = False
                    dblTotalBatchSize = dblCurrentBatchSize
                    Exit Sub
                End If
                fInfo = New FileInfo(child.Tag)
                dblFileSize = fInfo.Length
                bFoundFile = True
                If chkComputeBatchSize.Checked = True Then
                    If checked = True Then
                        dblTotalBatchSize = dblCurrentBatchSize + dblFileSize
                    Else
                        dblTotalBatchSize = dblCurrentBatchSize - dblFileSize
                    End If
                    Me.lblTotalBatchSize.Text = dblTotalBatchSize.ToString("N0")
                End If
                plstSourceFiles.Add(child.Tag)
                child.Checked = checked
                'If child.Nodes.Count > 0 Then CheckChildFileNodes(child, checked, dblTotalBatchSize)

            Catch ex As Exception

            End Try
        Next

        If bFoundFile = False Then
            dblTotalBatchSize = dblCurrentBatchSize
        End If
    End Sub
    Public Sub UnCheckParentNodes(ByVal iNode As TreeNode)
        Try
            If iNode.Checked = False AndAlso iNode.Parent IsNot Nothing Then

                iNode.Checked = False

                UnCheckParentNodes(iNode)

            End If

        Catch ex As Exception

        End Try

    End Sub
    Private Sub AddAllFoldersAndFiles(ByVal TNode As TreeNode, ByVal FolderPath As String)
        Dim SubFolderNode As TreeNode
        Dim di As IO.DirectoryInfo
        Dim sFileName As String

        '  Create a new ImageList
        Dim MyImages As New ImageList()

        'Add the files to treeview
        di = New DirectoryInfo(FolderPath)
        If di.Attributes <> FileAttributes.ReadOnly Then
            Try
                For Each FolderNode As String In IO.Directory.GetDirectories(FolderPath)
                    If FolderNode <> vbNullString Then '
                        SubFolderNode = TNode.Nodes.Add(FolderNode.Substring(FolderNode.LastIndexOf("\"c) + 1))
                        With SubFolderNode
                            .Tag = FolderNode
                            .Nodes.Add("{child}")
                            treeViewFolders.ImageList = ImgIcons
                            .ImageIndex = 0
                            .SelectedImageIndex = 1
                        End With

                    End If
                Next
            Catch e As IOException
                'Do nothing, since you will get an exception when adding an empty CD/DVD-drive...
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            Try
                Dim files() As String = IO.Directory.GetFiles(FolderPath)
                If files.Length <> 0 Then
                    Dim fileNode As TreeNode = Nothing
                    For Each file In files
                        Try
                            sFileName = IO.Path.GetFileName(file)
                            fileNode = TNode.Nodes.Add(file.ToString, IO.Path.GetFileName(file), 2, 2)
                            fileNode.SelectedImageIndex = 1
                            fileNode.Tag = file.ToString

                        Catch ex As Exception

                        End Try
                    Next
                End If

            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub AddAllFoldersAndCLPFiles(ByVal TNode As TreeNode, ByVal FolderPath As String)
        Dim SubFolderNode As TreeNode
        Dim di As IO.DirectoryInfo
        Dim sFileName As String

        '  Create a new ImageList
        Dim MyImages As New ImageList()

        'Add the files to treeview
        di = New DirectoryInfo(FolderPath)
        If di.Attributes <> FileAttributes.ReadOnly Then
            Try
                For Each FolderNode As String In IO.Directory.GetDirectories(FolderPath)
                    If FolderNode <> vbNullString Then '
                        SubFolderNode = TNode.Nodes.Add(FolderNode.Substring(FolderNode.LastIndexOf("\"c) + 1))
                        With SubFolderNode
                            .Tag = FolderNode
                            .Nodes.Add("{child}")
                            treeViewFolders.ImageList = ImgIcons
                            .ImageIndex = 0
                            .SelectedImageIndex = 1
                        End With

                    End If
                Next
            Catch e As IOException
                'Do nothing, since you will get an exception when adding an empty CD/DVD-drive...
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            Try
                Dim files() As String = IO.Directory.GetFiles(FolderPath)
                If files.Length <> 0 Then
                    Dim fileNode As TreeNode = Nothing
                    For Each file In files
                        Try
                            sFileName = IO.Path.GetFileName(file)
                            If sFileName.Contains(".clp") Then
                                fileNode = TNode.Nodes.Add(file.ToString, IO.Path.GetFileName(file), 2, 2)
                                fileNode.SelectedImageIndex = 1
                                fileNode.Tag = file.ToString

                            End If
                        Catch ex As Exception

                        End Try
                    Next
                End If

            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub AddAllFolders(ByVal TNode As TreeNode, ByVal FolderPath As String)
        Dim SubFolderNode As TreeNode
        Dim di As IO.DirectoryInfo

        '  Create a new ImageList
        Dim MyImages As New ImageList()

        'Add the files to treeview
        di = New DirectoryInfo(FolderPath)
        If di.Attributes <> FileAttributes.ReadOnly Then
            Try
                For Each FolderNode As String In IO.Directory.GetDirectories(FolderPath)
                    If FolderNode <> vbNullString Then '
                        If IO.File.Exists(FolderNode.ToString) Then

                        Else
                            SubFolderNode = TNode.Nodes.Add(FolderNode.Substring(FolderNode.LastIndexOf("\"c) + 1))

                            With SubFolderNode
                                .Tag = FolderNode
                                .Nodes.Add("{child}")
                                treeViewFolders.ImageList = ImgIcons
                                .ImageIndex = 0
                                .SelectedImageIndex = 1
                            End With
                        End If
                    End If
                Next
            Catch e As IOException
                'Do nothing, since you will get an exception when adding an empty CD/DVD-drive...
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End If

    End Sub

    Private Sub treeViewFolders_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles treeViewFolders.BeforeExpand
        If radAddFolder.Checked = True Then
            If e.Node.Nodes.Count = 1 AndAlso e.Node.Nodes(0).Text = "{child}" Then
                e.Node.Nodes.Clear()
                AddAllFolders(e.Node, CStr(e.Node.Tag))
            End If

        ElseIf radAddFile.Checked = True Then
            If e.Node.Nodes.Count = 1 AndAlso e.Node.Nodes(0).Text = "{child}" Then
                e.Node.Nodes.Clear()
                AddAllFoldersAndFiles(e.Node, CStr(e.Node.Tag))
            End If
        ElseIf radCentera.Checked = True Then
            If e.Node.Nodes.Count = 1 AndAlso e.Node.Nodes(0).Text = "{child}" Then
                e.Node.Nodes.Clear()
                AddAllFoldersAndCLPFiles(e.Node, CStr(e.Node.Tag))
            End If
        End If
    End Sub

    Private Sub treeAXSOneSis_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles treeAXSOneSis.AfterCheck
        If e.Node.Checked = True Then

            plstAXSOneSIS.Add(e.Node.Tag)
        Else
            If plstAXSOneSIS.Contains(e.Node.Tag) Then
                plstAXSOneSIS.Remove(e.Node.Tag)
            End If
        End If

    End Sub

    Private Sub treeAXSOneSis_AfterExpand(sender As Object, e As TreeViewEventArgs) Handles treeAXSOneSis.AfterExpand
        If e.Node.Nodes.Count = 1 AndAlso e.Node.Nodes(0).Text = "{child}" Then
            e.Node.Nodes.Clear()
            AddAllFoldersAndFiles(e.Node, CStr(e.Node.Tag))
            e.Node.SelectedImageIndex = 1
        End If
    End Sub

    Private Sub treeAXSOneSis_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles treeAXSOneSis.BeforeExpand
        If e.Node.Nodes.Count = 1 AndAlso e.Node.Nodes(0).Text = "{child}" Then
            e.Node.Nodes.Clear()
            AddAllFoldersAndFiles(e.Node, CStr(e.Node.Tag))
        End If
    End Sub

    Private Sub chkComputeBatchSize_CheckedChanged(sender As Object, e As EventArgs) Handles chkComputeBatchSize.CheckedChanged
        Dim dlgReturn As DialogResult
        Dim bStatus As Boolean
        Dim dblTotalFolderSize As Double
        Dim dblTotalBatchSize As Double
        Dim dblCurrentBatchSize As Double

        If chkComputeBatchSize.Checked = True Then
            dlgReturn = MessageBox.Show("Computing Batch Sizes can dramatically effect performance.  Are you sure you want to compute the Batch Size?", "Compute Batch Size", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            If dlgReturn = Windows.Forms.DialogResult.No Then
                chkComputeBatchSize.Checked = False
            Else
                If plstSourceFolders.Count > 0 Then
                    For Each SourceFolder In plstSourceFolders
                        bStatus = blnComputeFolderSize(SourceFolder.ToString, dblTotalFolderSize)
                        dblTotalBatchSize = dblCurrentBatchSize + dblTotalFolderSize
                        Me.lblTotalBatchSize.Text = dblTotalBatchSize.ToString("N0")
                    Next
                End If
            End If
        Else
            dblTotalBatchSize = 0.0
            dblCurrentBatchSize = 0.0
            dblTotalFolderSize = 0.0

            Me.lblTotalBatchSize.Text = "0"

        End If
    End Sub

    Private Sub radAddFolder_CheckedChanged(sender As Object, e As EventArgs) Handles radAddFolder.CheckedChanged
        Dim Tnode As TreeNode
        Dim drives As System.Collections.ObjectModel.ReadOnlyCollection(Of IO.DriveInfo)
        Dim rootDir As String

        If radAddFolder.Checked = True Then
            chkComputeBatchSize.Enabled = True

            treeViewFolders.Nodes.Clear()
            treeViewFolders.Width = 266
            treeViewFolders.Height = 230

            lblPEAFile.Hide()
            lblIPFile.Hide()
            txtPEAFile.Hide()
            txtIPFile.Hide()
            btnPEAFile.Hide()
            btnIPFile.Hide()
            drives = My.Computer.FileSystem.Drives
            For i As Integer = 0 To drives.Count - 1
                If Not drives(i).IsReady Then
                    Continue For
                End If
                Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
                rootDir = drives(i).Name
                AddAllFolders(Tnode, rootDir)
            Next

        ElseIf radAddFile.Checked = True Then
            chkComputeBatchSize.Enabled = True
            lblPEAFile.Hide()
            lblIPFile.Hide()
            txtPEAFile.Hide()
            txtIPFile.Hide()
            btnPEAFile.Hide()
            btnIPFile.Hide()
            treeViewFolders.Nodes.Clear()

            treeViewFolders.Width = 266
            treeViewFolders.Height = 230

            drives = My.Computer.FileSystem.Drives
            For i As Integer = 0 To drives.Count - 1
                If Not drives(i).IsReady Then
                    Continue For
                End If
                Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
                rootDir = drives(i).Name
                AddAllFoldersAndFiles(Tnode, rootDir)
            Next

        ElseIf radCentera.Checked = True Then

            If cboExtractionOutputType.Text = "pst" And radUser.Checked = True And radMailboxArchive.Checked = True Then
                treeViewFolders.Enabled = False
                treeViewFolders.Height = 170
            Else
                treeViewFolders.Enabled = True
                chkComputeBatchSize.Checked = False

                chkComputeBatchSize.Enabled = False
                treeViewFolders.Nodes.Clear()
                treeViewFolders.Width = 266
                treeViewFolders.Height = 170
                drives = My.Computer.FileSystem.Drives
                For i As Integer = 0 To drives.Count - 1
                    If Not drives(i).IsReady Then
                        Continue For
                    End If
                    Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
                    rootDir = drives(i).Name
                    AddAllFoldersAndCLPFiles(Tnode, rootDir)
                Next

            End If

            lblPEAFile.Show()
            lblIPFile.Show()
            txtPEAFile.Show()
            txtIPFile.Show()
            btnPEAFile.Show()
            btnIPFile.Show()
        End If
    End Sub

    Private Sub radAddFile_CheckedChanged(sender As Object, e As EventArgs) Handles radAddFile.CheckedChanged
        Dim Tnode As TreeNode
        Dim drives As System.Collections.ObjectModel.ReadOnlyCollection(Of IO.DriveInfo)
        Dim rootDir As String

        If radAddFolder.Checked = True Then
            chkComputeBatchSize.Enabled = True

            treeViewFolders.Nodes.Clear()
            treeViewFolders.Width = 266
            treeViewFolders.Height = 230

            lblPEAFile.Hide()
            lblIPFile.Hide()
            txtPEAFile.Hide()
            txtIPFile.Hide()
            btnPEAFile.Hide()
            btnIPFile.Hide()
            drives = My.Computer.FileSystem.Drives
            For i As Integer = 0 To drives.Count - 1
                If Not drives(i).IsReady Then
                    Continue For
                End If
                Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
                rootDir = drives(i).Name
                AddAllFolders(Tnode, rootDir)
            Next

        ElseIf radAddFile.Checked = True Then
            chkComputeBatchSize.Enabled = True
            lblPEAFile.Hide()
            lblIPFile.Hide()
            txtPEAFile.Hide()
            txtIPFile.Hide()
            btnPEAFile.Hide()
            btnIPFile.Hide()
            treeViewFolders.Nodes.Clear()

            treeViewFolders.Width = 266
            treeViewFolders.Height = 230

            drives = My.Computer.FileSystem.Drives
            For i As Integer = 0 To drives.Count - 1
                If Not drives(i).IsReady Then
                    Continue For
                End If
                Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
                rootDir = drives(i).Name
                AddAllFoldersAndFiles(Tnode, rootDir)
            Next

        ElseIf radCentera.Checked = True Then

            If cboExtractionOutputType.Text = "pst" And radUser.Checked = True And radMailboxArchive.Checked = True Then
                treeViewFolders.Enabled = False
                treeViewFolders.Height = 170
            Else
                treeViewFolders.Enabled = True
                chkComputeBatchSize.Checked = False

                chkComputeBatchSize.Enabled = False
                treeViewFolders.Nodes.Clear()
                treeViewFolders.Width = 266
                treeViewFolders.Height = 170
                drives = My.Computer.FileSystem.Drives
                For i As Integer = 0 To drives.Count - 1
                    If Not drives(i).IsReady Then
                        Continue For
                    End If
                    Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
                    rootDir = drives(i).Name
                    AddAllFoldersAndCLPFiles(Tnode, rootDir)
                Next

            End If

            lblPEAFile.Show()
            lblIPFile.Show()
            txtPEAFile.Show()
            txtIPFile.Show()
            btnPEAFile.Show()
            btnIPFile.Show()
        End If
    End Sub

    Private Sub btnAddBatchToGrid_Click(sender As Object, e As EventArgs) Handles btnAddBatchToGrid.Click
        Dim sArchiveName As String
        Dim sSource As String
        Dim sOutputFormat As String
        Dim sSQLSettings As String
        Dim sEVSettings As String
        Dim sEMCSettings As String
        Dim iRowIndex As Integer
        Dim ExtractionRow As DataGridViewRow
        Dim sArchiveType As String
        Dim sOutputType As String
        Dim sCaseTimeStamp As String
        Dim bStatus As Boolean
        Dim sWSSSettings As String
        Dim sFileFiltering As String
        Dim sExcludeItems As String
        Dim sVerbose As String
        Dim sAddressFiltering As String
        Dim sExpandDLLocation As String
        Dim sExportDir As String
        Dim sNuixCaseDir As String
        Dim sAXSOneSisFolders As String
        Dim sExtractFolder As String
        Dim sBatchType As String
        Dim sProcessingFilesDir As String
        Dim sRunDate As String
        Dim sFromDate As String
        Dim sToDate As String
        Dim dFromDate As DateTime
        Dim dToDate As DateTime
        Dim sArchiveSettings As String
        Dim sContentFiltering As String
        Dim sEmailContentFiltering As String
        Dim sRSSFeedContentFiltering As String
        Dim sCalendarContentFiltering As String
        Dim sContactContentFiltering As String
        Dim sHostName As String
        Dim sPort As String
        Dim sDBName As String
        Dim sUserName As String
        Dim sUserInfo As String
        Dim sIPFFile As String
        Dim sLogDirectory As String
        Dim sPEAFile As String
        Dim sPEAVariable As String
        Dim dbService As New DatabaseService
        Dim common As New Common

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"

        sCaseTimeStamp = DateTime.Now.ToString("ddMMyyyyhhmmss")
        sRunDate = DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss")
        If grpSQLArchiveSettings.Visible = True Then
            sHostName = txtSQLHostName.Text
            sPort = txtSQLPortNumber.Text
            sDBName = txtSQLDBName.Text
            sUserName = txtSQLUserName.Text
            sUserInfo = txtSQLInfo.Text


            If cboSecurityType.Text = "Windows Authentication" Then
                If sHostName = vbNullString Then
                    MessageBox.Show("You must enter a SQL Host Name", "Enter SQL Host Name", MessageBoxButtons.OK)
                    txtSQLHostName.Focus()
                    Exit Sub

                End If

                If sPort = vbNullString Then
                    MessageBox.Show("You must enter a Port Number", "Enter Port Number", MessageBoxButtons.OK)
                    txtSQLPortNumber.Focus()
                    Exit Sub

                End If

                If sDBName = vbNullString Then
                    MessageBox.Show("You must enter a Database Name", "Enter DB Name", MessageBoxButtons.OK)
                    txtSQLDBName.Focus()
                    Exit Sub

                End If

            ElseIf cboSecurityType.Text = "SQLServer Authentication" Then
                If sHostName = vbNullString Then
                    MessageBox.Show("You must enter a SQL Host Name", "Enter SQL Host Name", MessageBoxButtons.OK)
                    txtSQLHostName.Focus()
                    Exit Sub

                End If

                If sPort = vbNullString Then
                    MessageBox.Show("You must enter a Port Number", "Enter Port Number", MessageBoxButtons.OK)
                    txtSQLPortNumber.Focus()
                    Exit Sub

                End If

                If sDBName = vbNullString Then
                    MessageBox.Show("You must enter a Database Name", "Enter DB Name", MessageBoxButtons.OK)
                    txtSQLDBName.Focus()
                    Exit Sub

                End If
                If sUserName = vbNullString Then
                    MessageBox.Show("You must enter a SQL UserName", "Enter SQL UserName", MessageBoxButtons.OK)
                    txtSQLUserName.Focus()
                    Exit Sub

                End If

                If sUserInfo = vbNullString Then
                    MessageBox.Show("You must enter a SQL DB Password", "Enter SQL DB Password", MessageBoxButtons.OK)
                    txtSQLInfo.Focus()
                    Exit Sub
                End If
            Else

            End If
        End If

        If Trim(dateFromDate.Text) <> vbNullString Then
            dFromDate = dateFromDate.Text
            dToDate = dateToDate.Text
            sFromDate = dFromDate.ToString("yyyyMMdd") & "T" & dFromDate.ToString("HHmmss") & "Z"
            sToDate = dToDate.ToString("yyyyMMdd") & "T" & dToDate.ToString("HHmmss") & "Z"
        Else
            sFromDate = "NA"
            sToDate = "NA"
        End If


        If cboArchiveType.Text = vbNullString Then
            MessageBox.Show("You must select a Legacy Archive to migrate data from.", "Select Legacy Archive", MessageBoxButtons.OK)
            Exit Sub
        Else
            sArchiveName = cboArchiveType.Text
        End If
        If (radMailboxArchive.Checked = False And radJournalArchive.Checked = False) Then
            MessageBox.Show("You must select an Archive Type to migrate data from.", "Select Archive Type", MessageBoxButtons.OK)
            Exit Sub
        End If
        If grpSourceLocation.Enabled = True Then
            If radAddFile.Checked = True Then
                If plstSourceFiles.Count = 0 Then
                    MessageBox.Show("You must select a source for the archive data", "Select Archive Source Location", MessageBoxButtons.OK)
                    Exit Sub
                End If
            ElseIf radAddFolder.Checked = True Then
                If plstSourceFolders.Count = 0 Then
                    MessageBox.Show("You must select a source for the archive data", "Select Archive Source Location", MessageBoxButtons.OK)
                    Exit Sub
                End If
            End If
        End If
        If cboExtractionOutputType.Text = vbNullString Then
            MessageBox.Show("You must select an output type", "Select output type", MessageBoxButtons.OK)
            Exit Sub
        Else
            sOutputFormat = cboExtractionOutputType.Text
        End If

        If grpAXSOneSISSettings.Enabled = True Then
            If plstAXSOneSIS.Count = 0 Then
                MessageBox.Show("You must select an AXS-One SIS location", "Select AXS-One SIS Location", MessageBoxButtons.OK)
                Exit Sub

            End If
        End If
        Select Case cboArchiveType.Text
            Case "Veritas Enterprise Vault"
                If txtEVUserList.Enabled = True Then
                    If txtEVUserList.Text = vbNullString Then
                        MessageBox.Show("You must select a User List CSV for EV User Mailboxes")
                        Exit Sub
                    End If
                End If
                sBatchType = "EV"
                sExtractFolder = "\Archive Extract\EV\EV_" & sCaseTimeStamp
                sProcessingFilesDir = psNuixProcessingFilesDir & "EV\" & sRunDate
                System.IO.Directory.CreateDirectory(sProcessingFilesDir)
            Case "EMC EmailXtender"
                sBatchType = "EMC"
                sExtractFolder = "\Archive Extract\EMC\EMC_" & sCaseTimeStamp
                sProcessingFilesDir = psNuixProcessingFilesDir & "EMC\" & sRunDate
                System.IO.Directory.CreateDirectory(sProcessingFilesDir)
            Case "EMC SourceOne"
                sBatchType = "EMC"
                sExtractFolder = "\Archive Extract\EMC\EMC_" & sCaseTimeStamp
                sProcessingFilesDir = psNuixProcessingFilesDir & "EMC\" & sRunDate
                System.IO.Directory.CreateDirectory(sProcessingFilesDir)
            Case "HP/Autonomy EAS"
                sBatchType = "EAS"
                sExtractFolder = "\Archive Extract\EAS\EAS_" & sCaseTimeStamp
                sProcessingFilesDir = psNuixProcessingFilesDir & "EAS\" & sRunDate
                System.IO.Directory.CreateDirectory(sProcessingFilesDir)
            Case "Daegis AXS-One"
                sBatchType = "AXS1"
                sExtractFolder = "\Archive Extract\AXS1\AXS1_" & sCaseTimeStamp
                sProcessingFilesDir = psNuixProcessingFilesDir & "AXS1\" & sRunDate
                System.IO.Directory.CreateDirectory(sProcessingFilesDir)
        End Select
        sNuixCaseDir = psNuixCaseDir & sExtractFolder
        System.IO.Directory.CreateDirectory(sNuixCaseDir)
        sExportDir = psExportDir & sExtractFolder
        System.IO.Directory.CreateDirectory(sExportDir)
        sLogDirectory = psNuixLogDir & sExtractFolder
        If sLogDirectory = vbNullString Then
            common.Logger(psIngestionLogFile, "A log Directory must be specified - in btnAddBatchToGrid")
            MessageBox.Show("A Log Directory must be specified - in btnAddBatchToGrid", "Log Directory not specified", MessageBoxButtons.OK)
            Exit Sub
        Else
            System.IO.Directory.CreateDirectory(sLogDirectory)
        End If

        If radMailboxArchive.Checked = True Then
            If radUser.Checked = True Then
                sArchiveType = "Mailbox:User"
            ElseIf radFlat.Checked = True Then
                sArchiveType = "Mailbox:Flat"
            End If
        ElseIf radJournalArchive.Checked = True Then
            If radUser.Checked = True Then
                sArchiveType = "Journal:User"
            ElseIf radFlat.Checked = True Then
                sArchiveType = "Journal:Flat"
            End If
        End If

        If grpWSSControl.Visible = True Then
            sFileFiltering = "false"

            If txtSearchTermCSV.Text = vbNullString Then
                MessageBox.Show("You must choose a Search Term CSV File", "Choose Search Term CSV File", MessageBoxButtons.OK)
                txtSearchTermCSV.Focus()
                Exit Sub
            End If
            If txtMappingCSV.Text = vbNullString Then
                MessageBox.Show("You must choose a Mapping CSV File", "Choose Mapping CSV File", MessageBoxButtons.OK)
                txtMappingCSV.Focus()
                Exit Sub
            End If

            If radExcludeItemsTrue.Checked = True Then
                sExcludeItems = "true"
            Else
                sExcludeItems = "false"
            End If
            If radVerboseTrue.Checked = True Then
                sVerbose = "true"
            Else
                sVerbose = "false"
            End If

            If chkEmail.Checked = True Then
                sEmailContentFiltering = "Email:true"
            Else
                sEmailContentFiltering = "Email:false"
            End If
            If chkRSSFeed.Checked = True Then
                sRSSFeedContentFiltering = "RSSFeed:true"
            Else
                sRSSFeedContentFiltering = "RSSFeed:false"
            End If
            If chkCalendar.Checked = True Then
                sCalendarContentFiltering = "Calendar:true"
            Else
                sCalendarContentFiltering = "Calendar:false"
            End If
            If chkContact.Checked = True Then
                sContactContentFiltering = "Contact:true"
            Else
                sContactContentFiltering = "Contact:false"
            End If

            If sEmailContentFiltering = "Email:false" And sRSSFeedContentFiltering = "RSSFeed:false" And sCalendarContentFiltering = "Calendar:false" And sContactContentFiltering = "Contact:false" Then
                MessageBox.Show("You must select at least one content type for filtering.", "Select File Filtering", MessageBoxButtons.OK)
                Exit Sub
            End If

            sContentFiltering = sEmailContentFiltering & "|" & sRSSFeedContentFiltering & "|" & sCalendarContentFiltering & "|" & sContactContentFiltering

            sWSSSettings = sContentFiltering & "|excludeItems:" & sExcludeItems & "|verbose:" & sVerbose & "|searchTermCSV:" & txtSearchTermCSV.Text & "|mappingCSV:" & txtMappingCSV.Text

        End If

        If grpEMCSettings.Visible = True Then

            If chkAFSYS.Checked = True Then
                If sAddressFiltering = vbNullString Then
                    sAddressFiltering = "SYS;"
                Else
                    sAddressFiltering = ""
                End If
            End If

            If chkAFEX.Checked = True Then
                sAddressFiltering = sAddressFiltering & "EX;"
            End If

            If chkAFPST.Checked = True Then
                sAddressFiltering = sAddressFiltering & "PST"
            End If

            If sAddressFiltering <> vbNullString Then
                If sAddressFiltering.Last() = ";" Then
                    sAddressFiltering = sAddressFiltering.Substring(0, sAddressFiltering.Length - 1)

                End If
            End If

            sExpandDLLocation = cboExpandDLLocation.Text
        End If

        sOutputFormat = cboExtractionOutputType.Text
        If radZipped.Enabled = True Then
            If radZipped.Checked = True Then
                sOutputFormat = sOutputFormat & "|zipped"
            ElseIf radLoose.Checked = True Then
                sOutputFormat = sOutputFormat & "|loose"
            End If
        End If

        If radUser.Checked = True Then
            sOutputType = "User"
        ElseIf radFlat.Checked = True Then
            sOutputType = "Flat"
        End If

        If radAddFolder.Checked = True Then
            If plstSourceFolders.Count > 0 Then
                For Each SourceFolder In plstSourceFolders
                    If sSource = vbNullString Then
                        sSource = SourceFolder.ToString
                    Else
                        sSource = sSource & "|" & SourceFolder.ToString
                    End If
                Next
            End If
        ElseIf radAddFile.Checked = True Then

            If plstSourceFiles.Count > 0 Then
                For Each SourceFile In plstSourceFiles
                    If sSource = vbNullString Then
                        sSource = SourceFile.ToString
                    Else
                        sSource = sSource & "|" & SourceFile.ToString
                    End If
                Next
            End If
        ElseIf radCentera.Checked = True Then
            If txtPEAFile.Text <> vbNullString Then

                sPEAFile = txtPEAFile.Text
                sPEAVariable = Environment.GetEnvironmentVariable("CENTERA_PEA_LOCATION")
            End If
            If plstCenteraClipLists.Count < 1 Then
                MessageBox.Show("You must select at least one Centera Clip List", "Centera Clip List Selection", MessageBoxButtons.OK)
                Exit Sub
            Else
                For Each sourceFolder In plstCenteraClipLists
                    sSource = sourceFolder.ToString
                Next
            End If
            If txtIPFile.Text = vbNullString Then
                MessageBox.Show("You must select an IP File", "Centera IP File Selection", MessageBoxButtons.OK)
                Exit Sub
            Else
                sIPFFile = txtIPFile.Text
            End If
            sSource = "CenteraClip:" & sSource & "|CenteraIP:" & sIPFFile & "|CenteraPEA:" & sPeaFile
        End If
        If ((radMailboxArchive.Checked = True) And (radUser.Checked = True) And (cboArchiveType.Text = "Veritas Enterprise Vault")) Then
            cboExtractionOutputType.Text = "pst"
            sOutputFormat = cboExtractionOutputType.Text
            If radZipped.Enabled = True Then
                If radZipped.Checked = True Then
                    sOutputFormat = sOutputFormat & "|zipped"
                ElseIf radLoose.Checked = True Then
                    sOutputFormat = sOutputFormat & "|loose"
                End If
            End If
        End If

        sEVSettings = "SkipAdditionalSQLLookups:" & chkSkipAdditionalSQLLookup.Checked.ToString.ToLower & "|SkipVaultStorePartitionErrors:" & chkSkipVaultStorePartitionErrors.Checked.ToString.ToLower

        sEMCSettings = "AddressFilter:" & sAddressFiltering & "|ExpandedDLLocation:" & sExpandDLLocation
        sSQLSettings = "Authentication:" & cboSecurityType.Text & "|Domain:" & txtDomain.Text & "|HostName:" & txtSQLHostName.Text & "|PortNumber:" & txtSQLPortNumber.Text & "|DBName:" & txtSQLDBName.Text & "|SQLUsername:" & txtSQLUserName.Text & "|SQLINfo:" & txtSQLInfo.Text & "|Domain:" & txtDomain.Text

        iRowIndex = grdArchiveExtractionBatch.Rows.Add()

        ExtractionRow = grdArchiveExtractionBatch.Rows(iRowIndex)
        common.Logger(psIngestionLogFile, "Adding " & "EV_" & sCaseTimeStamp & " to processing grid")

        With ExtractionRow

            ExtractionRow.Cells("StatusImage").Value = My.Resources.waitingtostart1
            ExtractionRow.Cells("SelectBatch").Value = True
            ExtractionRow.Cells("ArchiveName").Value = sArchiveName
            ExtractionRow.Cells("ExtractionStatus").Value = "Not Started"
            If (cboArchiveType.Text = "Veritas Enterprise Vault") Then
                ExtractionRow.Cells("BatchName").Value = "EV_" & sCaseTimeStamp
                sArchiveSettings = sEVSettings & "|FromDate:" & sFromDate & "|ToDate:" & sToDate
                If txtEVUserList.Enabled = True Then
                    sArchiveSettings = sArchiveSettings & "|EVUserList:" & txtEVUserList.Text
                Else
                    sArchiveSettings = sArchiveSettings & "|EVUserList:" & ""
                End If
                ExtractionRow.Cells("ArchiveSettings").Value = sArchiveSettings
            ElseIf (cboArchiveType.Text = "EMC EmailXtender" Or cboArchiveType.Text = "EMC SourceOne") Then
                ExtractionRow.Cells("BatchName").Value = "EMC_" & sCaseTimeStamp
                sArchiveSettings = sEMCSettings & "|FromDate:" & sFromDate & "|ToDate:" & sToDate
                ExtractionRow.Cells("ArchiveSettings").Value = sArchiveSettings
            ElseIf (cboArchiveType.Text = "HP/Autonomy EAS") Then
                ExtractionRow.Cells("BatchName").Value = "EAS_" & sCaseTimeStamp
                sArchiveSettings = "DocStoreID:" & txtDocServerID.Text & "|FromDate:" & sFromDate & "|ToDate:" & sToDate
                ExtractionRow.Cells("ArchiveSettings").Value = sArchiveSettings
            ElseIf (cboArchiveType.Text = "Daegis AXS-One") Then
                ExtractionRow.Cells("BatchName").Value = "AXS1_" & sCaseTimeStamp
                For Each SisFolder In plstAXSOneSIS
                    sAXSOneSisFolders = sAXSOneSisFolders & SisFolder.ToString & ";"
                Next
                ' remove the last | from the string
                sAXSOneSisFolders = sAXSOneSisFolders.Remove(sAXSOneSisFolders.Length - 1)
                sArchiveSettings = "SkipSISlookups:" & chkAXSOneSkipSISLookups.Checked.ToString.ToLower & "|SISFolder:" & sAXSOneSisFolders & "|FromDate:" & sFromDate & "|ToDate:" & sToDate
                ExtractionRow.Cells("ArchiveSettings").Value = sArchiveSettings
            End If
            ExtractionRow.Cells("TotalBytes").Value = lblTotalBatchSize.Text.Replace(",", "")
            ExtractionRow.Cells("ArchiveType").Value = sArchiveType
            ExtractionRow.Cells("SQLConnectionInfo").Value = sSQLSettings
            ExtractionRow.Cells("LightspeedOutputType").Value = sOutputFormat
            ExtractionRow.Cells("SourceInformation").Value = sSource
            ExtractionRow.Cells("WSSSettings").Value = sWSSSettings
            ExtractionRow.Cells("ProcessingFilesDirectory").Value = sProcessingFilesDir
            ExtractionRow.Cells("CaseDirectory").Value = sNuixCaseDir
            ExtractionRow.Cells("OutputDirectory").Value = sExportDir
            ExtractionRow.Cells("LogDirectory").Value = sLogDirectory


        End With
        common.Logger(psIngestionLogFile, "Updating Database for " & sBatchType & "_" & sCaseTimeStamp)
        bStatus = dbService.UpdateArchiveExtractBatchInfo(sSQLiteDatabaseFullName, sArchiveName, sBatchType & "_" & sCaseTimeStamp, "Not Started", sArchiveSettings, sArchiveType, sOutputFormat, sSQLSettings, sWSSSettings, sSource, "", "", 0, 0, 0, 0, sProcessingFilesDir, sNuixCaseDir, sExportDir, sLogDirectory, "", lblTotalBatchSize.Text.Replace(",", ""))

    End Sub

    Private Sub btnLaunchBatches_Click(sender As Object, e As EventArgs) Handles btnLaunchBatches.Click
        btnLoadPreviousBatches.Enabled = False

        RunExtractions = New System.Threading.Thread(AddressOf Me.LaunchExtractionBatches)
        RunExtractions.Start()

    End Sub
    Private Sub LaunchExtractionBatches()
        Dim bStatus As Boolean

        bStatus = blnBuildArchiveExtractionFiles()

    End Sub

    Private Sub SQLiteUpdates()

        Dim lstExtractionSize As List(Of String)
        Dim lstExtractedItems As List(Of String)
        Dim lstSMTPName As List(Of String)
        Dim lstNotStartedSMTPAddress As List(Of String)
        Dim bStatus As Boolean
        Dim bNoMoreJobs As Boolean
        Dim dblItemsProcessed As Double
        Dim dbService As New DatabaseService

        bNoMoreJobs = False
        Thread.Sleep(10000)
        lstExtractionSize = New List(Of String)
        lstExtractedItems = New List(Of String)
        lstSMTPName = New List(Of String)
        lstNotStartedSMTPAddress = New List(Of String)

        Do While bNoMoreJobs = False


            For Each row In grdArchiveExtractionBatch.Rows
                If row.cells("ExtractionStatus").value = "In Progress" Then
                    row.DefaultCellStyle.Forecolor = Color.Black
                    bStatus = dbService.GetUpdatedProcessingDetails(psSQLiteLocation, row.cells("BatchName").value, dblItemsProcessed)
                    row.cells("ExtractedItems").value = FormatNumber(dblItemsProcessed)
                End If
            Next

            'bNoMoreJobs = pbNoMoreJobs

            Thread.Sleep(10000)
        Loop
        'bStatus = blnRefreshGridFromDB(grdCustodainSMTPInfo, lstNotStartedSMTPAddress, lstInProgresSMTPAddress)
        'common.Logger(psExtractionLogFile, "All Currently select Processing Jobs have completed.  If necessary select more custodians for Processing.")
        MessageBox.Show("All Currently select Processing Jobs have completed.  If necessary select more custodians for Processing.", "All Processing Jobs Completed", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification, False)

    End Sub

    Private Sub radMailboxArchive_CheckedChanged(sender As Object, e As EventArgs) Handles radMailboxArchive.CheckedChanged
        Dim drives As System.Collections.ObjectModel.ReadOnlyCollection(Of IO.DriveInfo)
        Dim Tnode As TreeNode
        Dim rootDir As String

        grpOutputFormat.Enabled = True
        grpExtractionType.Enabled = False
        radUser.Checked = False
        radFlat.Checked = False
        cboExtractionOutputType.Text = ""
        radCentera.Checked = False
        treeViewFolders.Enabled = True
        Select Case cboArchiveType.Text
            Case "Veritas Enterprise Vault"
                If radMailboxArchive.Checked = True And radUser.Checked = True Then
                    cboExtractionOutputType.Text = "pst"
                    cboExtractionOutputType.Enabled = False
                    'grpSourceInformation.Enabled = False
                    radAddFile.Enabled = False
                    radAddFolder.Enabled = False

                    grpWSSControl.Hide()
                    treeViewFolders.Nodes.Clear()
                    drives = My.Computer.FileSystem.Drives
                    For i As Integer = 0 To drives.Count - 1
                        If Not drives(i).IsReady Then
                            Continue For
                        End If
                        Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
                        rootDir = drives(i).Name
                        AddAllFolders(Tnode, rootDir)
                    Next
                Else
                    cboExtractionOutputType.Enabled = True
                End If
            Case "EMC EmailXtender"
                radMailboxArchive.Enabled = False
                radMailboxArchive.Checked = False
                radJournalArchive.Checked = True
                radJournalArchive.Enabled = True
            Case "EMC SourceOne"
                radMailboxArchive.Enabled = False
                radMailboxArchive.Checked = False
                radJournalArchive.Checked = True
                radJournalArchive.Enabled = True
            Case "HP/Autonomy EAS"
                radMailboxArchive.Enabled = False
                radMailboxArchive.Checked = False
                radJournalArchive.Checked = True
                radJournalArchive.Enabled = True
            Case "Daegis AXS-One"
                radMailboxArchive.Enabled = True
                radJournalArchive.Enabled = True
                grpSourceInformation.Enabled = True
                radAddFile.Enabled = True
                radAddFolder.Enabled = True
                If radMailboxArchive.Checked = True And radUser.Checked = True Then
                    grpWSSControl.Show()
                    grpWSSControl.Height = 330
                Else
                    grpWSSControl.Hide()
                End If

        End Select

    End Sub

    Private Sub radUser_CheckedChanged_1(sender As Object, e As EventArgs) Handles radUser.CheckedChanged
        Dim drives As System.Collections.ObjectModel.ReadOnlyCollection(Of IO.DriveInfo)
        Dim Tnode As TreeNode
        Dim rootDir As String

        grpExtractionType.Enabled = True
        cboExtractionOutputType.Enabled = True
        cboExtractionOutputType.Text = ""
        grpWSSControl.Hide()
        cboExtractionOutputType.Enabled = True
        chkComputeBatchSize.Enabled = True
        radCentera.Checked = False
        txtEVUserList.Text = ""


        Select Case cboArchiveType.Text
            Case "Veritas Enterprise Vault"
                If radMailboxArchive.Checked = True And radUser.Checked = True Then
                    chkComputeBatchSize.Checked = False
                    cboExtractionOutputType.Text = "pst"
                    cboExtractionOutputType.Enabled = False
                    chkComputeBatchSize.Enabled = False
                    radCentera.Enabled = False
                    'grpSourceInformation.Enabled = False
                    radAddFile.Enabled = False
                    radAddFile.Checked = False
                    radAddFolder.Enabled = False
                    radAddFolder.Checked = False
                    treeViewFolders.Enabled = False
                    txtEVUserList.Enabled = True
                    btnEVUserListFileChooser.Enabled = True
                    grpWSSControl.Hide()
                    treeViewFolders.Nodes.Clear()
                    drives = My.Computer.FileSystem.Drives
                    For i As Integer = 0 To drives.Count - 1
                        If Not drives(i).IsReady Then
                            Continue For
                        End If
                        Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
                        rootDir = drives(i).Name
                        AddAllFolders(Tnode, rootDir)
                    Next
                Else
                    cboExtractionOutputType.Enabled = True
                    radAddFile.Enabled = True
                    radAddFolder.Enabled = True
                    radAddFolder.Checked = True
                    txtEVUserList.Enabled = False
                    btnEVUserListFileChooser.Enabled = False

                    If radJournalArchive.Checked = True And radUser.Checked = True Then
                        grpWSSControl.Show()

                    End If
                End If
            Case "EMC EmailXtender"
                If radJournalArchive.Checked = True And radUser.Checked = True Then
                    grpWSSControl.Show()
                End If
                radMailboxArchive.Enabled = False
                radMailboxArchive.Checked = False
                radJournalArchive.Checked = True
                radJournalArchive.Enabled = True
                radAddFile.Enabled = True
                radAddFolder.Enabled = True
                radAddFolder.Checked = True

            Case "EMC SourceOne"
                If radJournalArchive.Checked = True And radUser.Checked = True Then
                    grpWSSControl.Show()
                End If
                radMailboxArchive.Enabled = False
                radMailboxArchive.Checked = False
                radJournalArchive.Checked = True
                radJournalArchive.Enabled = True
                radAddFile.Enabled = True
                radAddFolder.Enabled = True
                radAddFolder.Checked = True
            Case "HP/Autonomy EAS"
                If radJournalArchive.Checked = True And radUser.Checked = True Then
                    grpWSSControl.Show()
                End If
                radMailboxArchive.Enabled = False
                radMailboxArchive.Checked = False
                radJournalArchive.Checked = True
                radJournalArchive.Enabled = True
                radAddFile.Enabled = True
                radAddFolder.Enabled = True
                radAddFolder.Checked = True
            Case "Daegis AXS-One"
                If radMailboxArchive.Checked = True And radUser.Checked = True Then
                    grpWSSControl.Show()
                End If
                If radJournalArchive.Checked = True And radUser.Checked = True Then
                    grpWSSControl.Show()
                End If
                grpSourceInformation.Enabled = True
                radAddFile.Enabled = True
                radAddFolder.Enabled = True
                radAddFolder.Checked = True
        End Select

    End Sub

    Private Sub btnCustodianListCSVChooser_Click(sender As Object, e As EventArgs) Handles btnCustodianListCSVChooser.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        With OpenFileDialog1
            .Filter = "*.csv|*.csv"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            txtCustodianListCSV.Text = OpenFileDialog1.FileName.ToString
        End If
    End Sub

    Private Sub radExcludeItemsTrue_CheckedChanged(sender As Object, e As EventArgs) Handles radExcludeItemsTrue.CheckedChanged
        If radExcludeItemsTrue.Checked = True Then
            radExcludeItemsFalse.Checked = False
            radExcludeItemsTrue.BackColor = Color.Green
            radExcludeItemsFalse.BackColor = Color.LightGray
        ElseIf radExcludeItemsTrue.Checked = False Then
            radExcludeItemsFalse.Checked = True
            radExcludeItemsTrue.BackColor = Color.LightGray
            radExcludeItemsFalse.BackColor = Color.Green
        End If
    End Sub

    Private Sub radVerboseTrue_CheckedChanged(sender As Object, e As EventArgs) Handles radVerboseTrue.CheckedChanged
        If radVerboseTrue.Checked = True Then
            radVerboseFalse.Checked = False
            radVerboseTrue.BackColor = Color.Green
            radVerboseFalse.BackColor = Color.LightGray
        ElseIf radVerboseTrue.Checked = False Then
            radVerboseFalse.Checked = True
            radVerboseTrue.BackColor = Color.LightGray
            radVerboseFalse.BackColor = Color.Green
        End If
    End Sub

    Private Sub btnCommCSVSelector_Click(sender As Object, e As EventArgs) Handles btnCommCSVSelector.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        With OpenFileDialog1
            .Filter = "*.csv|*.csv"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            txtSearchTermCSV.Text = OpenFileDialog1.FileName.ToString
        End If
    End Sub

    Private Sub radExcludeItemsFalse_CheckedChanged(sender As Object, e As EventArgs) Handles radExcludeItemsFalse.CheckedChanged
        If radExcludeItemsTrue.Checked = True Then
            radExcludeItemsFalse.Checked = False
            radExcludeItemsTrue.BackColor = Color.Green
            radExcludeItemsFalse.BackColor = Color.LightGray
        ElseIf radExcludeItemsTrue.Checked = False Then
            radExcludeItemsFalse.Checked = True
            radExcludeItemsTrue.BackColor = Color.LightGray
            radExcludeItemsFalse.BackColor = Color.Green
        End If
    End Sub

    Private Sub radVerboseFalse_CheckedChanged(sender As Object, e As EventArgs) Handles radVerboseFalse.CheckedChanged
        If radVerboseTrue.Checked = True Then
            radVerboseFalse.Checked = False
            radVerboseTrue.BackColor = Color.Green
            radVerboseFalse.BackColor = Color.LightGray
        ElseIf radVerboseTrue.Checked = False Then
            radVerboseFalse.Checked = True
            radVerboseTrue.BackColor = Color.LightGray
            radVerboseFalse.BackColor = Color.Green
        End If
    End Sub

    Private Sub radJournalArchive_CheckedChanged(sender As Object, e As EventArgs) Handles radJournalArchive.CheckedChanged
        Dim drives As System.Collections.ObjectModel.ReadOnlyCollection(Of IO.DriveInfo)
        Dim Tnode As TreeNode
        Dim rootDir As String

        grpOutputFormat.Enabled = True
        grpExtractionType.Enabled = False
        radUser.Checked = False
        radFlat.Checked = False
        cboExtractionOutputType.Text = ""
        radCentera.Checked = False
        treeViewFolders.Enabled = True

        Select Case cboArchiveType.Text
            Case "Veritas Enterprise Vault"
                If radMailboxArchive.Checked = True Then
                    cboExtractionOutputType.Text = "pst"
                    cboExtractionOutputType.Enabled = False
                    'grpSourceInformation.Enabled = False
                    radAddFile.Enabled = False
                    radAddFolder.Enabled = False

                    grpWSSControl.Hide()
                    cboExtractionOutputType.Text = "pst"
                    treeViewFolders.Nodes.Clear()
                    drives = My.Computer.FileSystem.Drives
                    For i As Integer = 0 To drives.Count - 1
                        If Not drives(i).IsReady Then
                            Continue For
                        End If
                        Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
                        rootDir = drives(i).Name
                        AddAllFolders(Tnode, rootDir)
                    Next
                Else
                    cboExtractionOutputType.Enabled = True
                End If
            Case "EMC EmailXtender"
                radMailboxArchive.Enabled = False
                radMailboxArchive.Checked = False
                radJournalArchive.Checked = True
                radJournalArchive.Enabled = True
            Case "EMC SourceOne"
                radMailboxArchive.Enabled = False
                radMailboxArchive.Checked = False
                radJournalArchive.Checked = True
                radJournalArchive.Enabled = True
            Case "HP/Autonomy EAS"
                radMailboxArchive.Enabled = False
                radMailboxArchive.Checked = False
                radJournalArchive.Checked = True
                radJournalArchive.Enabled = True
            Case "Daegis AXS-One"
                radMailboxArchive.Enabled = True
                radJournalArchive.Enabled = True

                grpSourceInformation.Enabled = True
                radAddFile.Enabled = True
                radAddFolder.Enabled = True

                'grpWSSControl.Show()
                grpWSSControl.Height = 330
        End Select

    End Sub

    Private Sub radFlat_CheckedChanged(sender As Object, e As EventArgs) Handles radFlat.CheckedChanged
        Dim drives As System.Collections.ObjectModel.ReadOnlyCollection(Of IO.DriveInfo)
        Dim Tnode As TreeNode
        Dim rootDir As String

        grpExtractionType.Enabled = True
        cboExtractionOutputType.Enabled = True
        cboExtractionOutputType.Text = ""
        grpWSSControl.Hide()
        cboExtractionOutputType.Enabled = True
        chkComputeBatchSize.Enabled = True
        radCentera.Checked = False
        txtEVUserList.Text = ""

        Select Case cboArchiveType.Text
            Case "Veritas Enterprise Vault"
                If radMailboxArchive.Checked = True And radUser.Checked = True Then
                    chkComputeBatchSize.Checked = False
                    cboExtractionOutputType.Text = "pst"
                    cboExtractionOutputType.Enabled = False
                    chkComputeBatchSize.Enabled = False
                    radCentera.Enabled = False
                    'grpSourceInformation.Enabled = False
                    radAddFile.Enabled = False
                    radAddFile.Checked = False
                    radAddFolder.Enabled = False
                    radAddFolder.Checked = False
                    treeViewFolders.Enabled = False
                    txtEVUserList.Enabled = True
                    btnEVUserListFileChooser.Enabled = True
                    grpWSSControl.Hide()
                    treeViewFolders.Nodes.Clear()
                    drives = My.Computer.FileSystem.Drives
                    For i As Integer = 0 To drives.Count - 1
                        If Not drives(i).IsReady Then
                            Continue For
                        End If
                        Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
                        rootDir = drives(i).Name
                        AddAllFolders(Tnode, rootDir)
                    Next
                Else
                    cboExtractionOutputType.Enabled = True
                    radAddFile.Enabled = True
                    radAddFolder.Enabled = True
                    radAddFolder.Checked = True
                    txtEVUserList.Enabled = False
                    btnEVUserListFileChooser.Enabled = False

                    If radJournalArchive.Checked = True And radUser.Checked = True Then
                        grpWSSControl.Show()

                    End If
                End If
            Case "EMC EmailXtender"
                If radJournalArchive.Checked = True And radUser.Checked = True Then
                    grpWSSControl.Show()
                End If
                radMailboxArchive.Enabled = False
                radMailboxArchive.Checked = False
                radJournalArchive.Checked = True
                radJournalArchive.Enabled = True
                radAddFile.Enabled = True
                radAddFolder.Enabled = True
                radAddFolder.Checked = True
            Case "EMC SourceOne"
                If radJournalArchive.Checked = True And radUser.Checked = True Then
                    grpWSSControl.Show()
                End If
                radMailboxArchive.Enabled = False
                radMailboxArchive.Checked = False
                radJournalArchive.Checked = True
                radJournalArchive.Enabled = True
                radAddFile.Enabled = True
                radAddFolder.Enabled = True
                radAddFolder.Checked = True
            Case "HP/Autonomy EAS"
                If radJournalArchive.Checked = True And radUser.Checked = True Then
                    grpWSSControl.Show()
                End If
                radMailboxArchive.Enabled = False
                radMailboxArchive.Checked = False
                radJournalArchive.Checked = True
                radJournalArchive.Enabled = True
                radAddFile.Enabled = True
                radAddFolder.Enabled = True
                radAddFolder.Checked = True
            Case "Daegis AXS-One"
                If radMailboxArchive.Checked = True And radUser.Checked = True Then
                    grpWSSControl.Show()
                End If
                If radJournalArchive.Checked = True And radUser.Checked = True Then
                    grpWSSControl.Show()
                End If
                grpSourceInformation.Enabled = True
                radAddFile.Enabled = True
                radAddFolder.Enabled = True
                radAddFolder.Checked = True
        End Select
    End Sub

    Private Sub btnMappingCSV_Click(sender As Object, e As EventArgs) Handles btnMappingCSV.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        With OpenFileDialog1
            .Filter = "*.csv|*.csv"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            txtMappingCSV.Text = OpenFileDialog1.FileName.ToString
        End If
    End Sub

    Private Sub btnClearDates_Click(sender As Object, e As EventArgs) Handles btnClearDates.Click

        setDateTimePickerBlank(dateFromDate)
        setDateTimePickerBlank(dateToDate)
    End Sub

    Private Sub dateFromDate_MouseDown(sender As Object, e As MouseEventArgs) Handles dateFromDate.MouseDown
        dateFromDate.Format = DateTimePickerFormat.Long
    End Sub

    Private Sub dateToDate_MouseDown(sender As Object, e As MouseEventArgs) Handles dateToDate.MouseDown
        dateToDate.Format = DateTimePickerFormat.Long
    End Sub

    Private Sub grdArchiveExtractionBatch_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdArchiveExtractionBatch.CellContentDoubleClick

        If ((e.ColumnIndex = 18) And (grdArchiveExtractionBatch("ProcessingFilesDirectory", e.RowIndex).Value <> vbNullString)) Then
            If Directory.Exists(grdArchiveExtractionBatch("ProcessingFilesDirectory", e.RowIndex).Value) Then
                Process.Start("explorer.exe", grdArchiveExtractionBatch("ProcessingFilesDirectory", e.RowIndex).Value)
            End If
        ElseIf ((e.ColumnIndex = 19) And (grdArchiveExtractionBatch("CaseDirectory", e.RowIndex).Value <> vbNullString)) Then
            If Directory.Exists(grdArchiveExtractionBatch("CaseDirectory", e.RowIndex).Value) Then
                Process.Start("explorer.exe", grdArchiveExtractionBatch("CaseDirectory", e.RowIndex).Value)
            End If
        ElseIf ((e.ColumnIndex = 20) And (grdArchiveExtractionBatch("OutputDirectory", e.RowIndex).Value <> vbNullString)) Then
            If Directory.Exists(grdArchiveExtractionBatch("OutputDirectory", e.RowIndex).Value) Then
                Process.Start("explorer.exe", grdArchiveExtractionBatch("OutputDirectory", e.RowIndex).Value)
            End If
        ElseIf ((e.ColumnIndex = 21) And (grdArchiveExtractionBatch("LogDirectory", e.RowIndex).Value <> vbNullString)) Then
            If Directory.Exists(grdArchiveExtractionBatch("LogDirectory", e.RowIndex).Value) Then
                Process.Start("explorer.exe", grdArchiveExtractionBatch("LogDirectory", e.RowIndex).Value)
            End If
        ElseIf ((e.ColumnIndex = 24) And (grdArchiveExtractionBatch("SummaryReportLocation", e.RowIndex).Value <> vbNullString)) Then
            System.Diagnostics.Process.Start(grdArchiveExtractionBatch("SummaryReportLocation", e.RowIndex).Value.ToString)
        End If
    End Sub

    Private Sub btnLoadPreviousBatches_Click(sender As Object, e As EventArgs) Handles btnLoadPreviousBatches.Click
        Dim bStatus As Boolean
        grdArchiveExtractionBatch.Rows.Clear()

        bStatus = blnLoadGridFromDB(grdArchiveExtractionBatch)
    End Sub

    Private Function blnLoadGridFromDB(ByVal grdArchiveBatches As DataGridView) As Boolean
        Dim mSQL As String
        Dim dt As DataTable
        Dim ds As DataSet
        Dim dataReader As SQLiteDataReader
        Dim sqlCommand As SQLiteCommand
        Dim sqlConnection As SQLiteConnection
        Dim BatchRow As DataGridViewRow
        Dim sArchiveName As String
        Dim iArchiveNameCol As Integer
        Dim sBatchName As String
        Dim iBatchNameCol As Integer
        Dim sExtractionStatus As String
        Dim iExtractionStatusCol As Integer
        Dim sArchiveSettings As String
        Dim iArchiveSettingsCol As Integer
        Dim sArchiveType As String
        Dim iArchiveTypeCol As Integer
        Dim sSQLSettings As String
        Dim iSQLSettingsCol As Integer
        Dim sSourceInformation As String
        Dim iSourceInformationCol As Integer
        Dim sProcessStartTime As String
        Dim iProcessStartTimeCol As Integer
        Dim sProcessEndTime As String
        Dim iProcessEndTimeCol As Integer
        Dim dblItemsProcessed As Double
        Dim iItemsProcessedCol As Integer
        Dim dblItemsExported As Double
        Dim iItemsExportedCol As Integer
        Dim dblItemsSkipped As Double
        Dim iItemsSkippedCol As Integer
        Dim dblItemsFailed As Double
        Dim iItemsFailedCol As Integer
        Dim sProcessingFilesDirectory As String
        Dim iProcessingFilesDirectoryCol As Integer
        Dim sCaseDirectory As String
        Dim iCaseDirectoryCol As Integer
        Dim sOutputDirectory As String
        Dim iOutputDirectoryCol As Integer
        Dim sLogDirectory As String
        Dim iLogDirectoryCol As Integer
        Dim sSummaryReportLocation As String
        Dim iSummaryReportLocationCol As Integer
        Dim sOutputFormat As String
        Dim iOutputFormatCol As Integer
        Dim sWSSSettings As String
        Dim iWSSSettingsCol As Integer
        Dim iTotalBytes As Double
        Dim iTotalBytesCol As Integer
        Dim iBytesProcessed As Double
        Dim iBytesProcessedCol As Integer
        Dim iPercentCompleted As Integer
        Dim iPercentCompletedCol As Integer
        Dim iRowIndex As Integer

        Dim common As New Common

        blnLoadGridFromDB = False

        Try
            dt = Nothing
            ds = New DataSet
            sqlConnection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
            mSQL = "select ArchiveName, BatchName, ExtractionStatus, ArchiveSettings, ArchiveType, OutputFormat, TotalBytes, BytesProcessed, PercentCompleted, SQLSettings, SourceInformation, WSSSettings, ProcessStartTime, ProcessEndTime, ItemsProcessed, ItemsExported, ItemsSkipped, ItemsFailed, ProcessingFilesDirectory, CaseDirectory, OutputDirectory, LogDirectory, SummaryReportLocation from archiveExtractionStats"

            sqlCommand = New SQLiteCommand(mSQL, sqlConnection)
            sqlConnection.Open()

            dataReader = sqlCommand.ExecuteReader

            While dataReader.Read
                iArchiveNameCol = dataReader.GetOrdinal("ArchiveName")
                If dataReader.IsDBNull(iArchiveNameCol) Then
                    sArchiveName = vbNullString
                Else
                    sArchiveName = dataReader.GetString(iArchiveNameCol)
                End If

                iBatchNameCol = dataReader.GetOrdinal("BatchName")
                If dataReader.IsDBNull(iBatchNameCol) Then
                    sBatchName = vbNullString
                Else
                    sBatchName = dataReader.GetString(iBatchNameCol)
                End If

                iExtractionStatusCol = dataReader.GetOrdinal("ExtractionStatus")
                If dataReader.IsDBNull(iExtractionStatusCol) Then
                    sExtractionStatus = vbNullString
                Else
                    sExtractionStatus = dataReader.GetString(iExtractionStatusCol)
                End If

                iTotalBytesCol = dataReader.GetOrdinal("TotalBytes")
                If dataReader.IsDBNull(iTotalBytesCol) Then
                    iTotalBytes = 0
                Else
                    iTotalBytes = dataReader.GetInt64(iTotalBytesCol)
                End If

                iBytesProcessedCol = dataReader.GetOrdinal("BytesProcessed")
                If dataReader.IsDBNull(iBytesProcessedCol) Then
                    iBytesProcessed = 0
                Else
                    iBytesProcessed = dataReader.GetInt64(iBytesProcessedCol)
                End If

                iPercentCompletedCol = dataReader.GetOrdinal("PercentCompleted")
                If dataReader.IsDBNull(iPercentCompletedCol) Then
                    iPercentCompleted = 0
                Else
                    iPercentCompleted = dataReader.GetValue(iPercentCompletedCol)
                End If

                iArchiveSettingsCol = dataReader.GetOrdinal("ArchiveSettings")
                If IsDBNull(iArchiveSettingsCol) Then
                    sArchiveSettings = vbNullString
                Else
                    sArchiveSettings = dataReader.GetString(iArchiveSettingsCol)
                End If

                iArchiveTypeCol = dataReader.GetOrdinal("ArchiveType")
                If IsDBNull(iArchiveTypeCol) Then
                    sArchiveType = vbNullString
                Else
                    sArchiveType = dataReader.GetString(iArchiveTypeCol)
                End If

                iOutputFormatCol = dataReader.GetOrdinal("OutputFormat")
                If IsDBNull(iOutputFormatCol) Then
                    sOutputFormat = vbNullString
                Else
                    sOutputFormat = dataReader.GetString(iOutputFormatCol)
                End If

                iSQLSettingsCol = dataReader.GetOrdinal("SQLSettings")
                If dataReader.IsDBNull(iSQLSettingsCol) Then
                    sSQLSettings = vbNullString
                Else
                    sSQLSettings = dataReader.GetString(iSQLSettingsCol)
                End If

                iWSSSettingsCol = dataReader.GetOrdinal("WSSSettings")
                If dataReader.IsDBNull(iWSSSettingsCol) Then
                    sWSSSettings = vbNullString
                Else
                    sWSSSettings = dataReader.GetString(iWSSSettingsCol)
                End If

                iSourceInformationCol = dataReader.GetOrdinal("SourceInformation")
                If dataReader.IsDBNull(iSourceInformationCol) Then
                    sSourceInformation = vbNullString
                Else
                    sSourceInformation = dataReader.GetString(iSourceInformationCol)
                End If

                iProcessStartTimeCol = dataReader.GetOrdinal("ProcessStartTime")
                If dataReader.IsDBNull(iProcessStartTimeCol) Then
                    sProcessStartTime = vbNullString
                Else
                    sProcessStartTime = dataReader.GetString(iProcessStartTimeCol)
                End If

                iProcessEndTimeCol = dataReader.GetOrdinal("ProcessEndTime")
                If dataReader.IsDBNull(iProcessEndTimeCol) Then
                    sProcessEndTime = vbNullString
                Else
                    sProcessEndTime = dataReader.GetString(iProcessEndTimeCol)
                End If

                iItemsProcessedCol = dataReader.GetOrdinal("ItemsProcessed")
                If dataReader.IsDBNull(iItemsProcessedCol) Then
                    dblItemsProcessed = 0.0
                Else
                    dblItemsProcessed = dataReader.GetInt64(iItemsProcessedCol)
                End If

                iItemsExportedCol = dataReader.GetOrdinal("ItemsExported")
                If dataReader.IsDBNull(iItemsExportedCol) Then
                    dblItemsExported = 0.0
                Else
                    dblItemsExported = dataReader.GetInt64(iItemsExportedCol)
                End If

                iItemsSkippedCol = dataReader.GetOrdinal("ItemsSkipped")
                If dataReader.IsDBNull(iItemsSkippedCol) Then
                    dblItemsSkipped = 0.0
                Else
                    dblItemsSkipped = dataReader.GetInt64(iItemsSkippedCol)
                End If

                iItemsFailedCol = dataReader.GetOrdinal("ItemsFailed")
                If dataReader.IsDBNull(iItemsFailedCol) Then
                    dblItemsFailed = 0.0
                Else
                    dblItemsFailed = dataReader.GetInt64(iItemsFailedCol)
                    '                sBytesUploaded = dblBytesUploaded.ToString("N0")
                End If

                iProcessingFilesDirectoryCol = dataReader.GetOrdinal("ProcessingFilesDirectory")
                If dataReader.IsDBNull(iProcessingFilesDirectoryCol) Then
                    sProcessingFilesDirectory = vbNullString
                Else
                    sProcessingFilesDirectory = dataReader.GetString(iProcessingFilesDirectoryCol)
                End If

                iCaseDirectoryCol = dataReader.GetOrdinal("CaseDirectory")
                If dataReader.IsDBNull(iCaseDirectoryCol) Then
                    sCaseDirectory = vbNullString
                Else
                    sCaseDirectory = dataReader.GetString(iCaseDirectoryCol)
                End If

                iOutputDirectoryCol = dataReader.GetOrdinal("OutputDirectory")
                If dataReader.IsDBNull(iOutputDirectoryCol) Then
                    sOutputDirectory = vbNullString
                Else
                    sOutputDirectory = dataReader.GetString(iOutputDirectoryCol)
                End If

                iLogDirectoryCol = dataReader.GetOrdinal("LogDirectory")
                If dataReader.IsDBNull(iLogDirectoryCol) Then
                    sLogDirectory = vbNullString
                Else
                    sLogDirectory = dataReader.GetString(iLogDirectoryCol)
                End If

                iSummaryReportLocationCol = dataReader.GetOrdinal("SummaryReportLocation")
                If dataReader.IsDBNull(iSummaryReportLocationCol) Then
                    sSummaryReportLocation = vbNullString
                Else
                    sSummaryReportLocation = dataReader.GetString(iSummaryReportLocationCol)
                End If

                iRowIndex = grdArchiveBatches.Rows.Add()

                BatchRow = grdArchiveBatches.Rows(iRowIndex)

                With BatchRow
                    BatchRow.Cells("SelectBatch").Value = False
                    BatchRow.Cells("ArchiveName").Value = sArchiveName
                    BatchRow.Cells("ExtractionStatus").Value = sExtractionStatus
                    If (sExtractionStatus = "Completed") Then
                        BatchRow.Cells("StatusImage").Value = My.Resources.Green_check_small
                        BatchRow.DefaultCellStyle.ForeColor = Color.Green
                    ElseIf (sExtractionStatus = "Failed") Then
                        BatchRow.Cells("StatusImage").Value = My.Resources.red_stop_small
                        BatchRow.DefaultCellStyle.ForeColor = Color.Red
                    ElseIf (sExtractionStatus = "Cancelled by User") Then
                        BatchRow.Cells("StatusImage").Value = My.Resources.yellow_info_small
                        BatchRow.DefaultCellStyle.ForeColor = Color.Orange
                    ElseIf (sExtractionStatus = "Not Started") Then
                        BatchRow.Cells("StatusImage").Value = My.Resources.waitingtostart1
                        BatchRow.DefaultCellStyle.ForeColor = Color.Black
                    ElseIf (sExtractionStatus = "SQL Connection Lost") Then
                        BatchRow.Cells("StatusImage").Value = My.Resources.red_stop_small
                        BatchRow.DefaultCellStyle.ForeColor = Color.Red
                    Else
                        BatchRow.Cells("StatusImage").Value = My.Resources.Default_image
                        BatchRow.DefaultCellStyle.ForeColor = Color.Black
                    End If

                    BatchRow.Cells("PercentCompleted").Value = iPercentCompleted
                    BatchRow.Cells("TotalBytes").Value = iTotalBytes
                    BatchRow.Cells("BytesProcessed").Value = iBytesProcessed
                    BatchRow.Cells("BatchName").Value = sBatchName
                    BatchRow.Cells("ArchiveSettings").Value = sArchiveSettings
                    BatchRow.Cells("ArchiveType").Value = sArchiveType
                    BatchRow.Cells("LightspeedOutputType").Value = sOutputFormat
                    BatchRow.Cells("SQLConnectionInfo").Value = sSQLSettings
                    BatchRow.Cells("SourceInformation").Value = sSourceInformation
                    BatchRow.Cells("WSSSettings").Value = sWSSSettings
                    BatchRow.Cells("ProcessingStartDate").Value = sProcessStartTime
                    BatchRow.Cells("ProcessingEndDate").Value = sProcessEndTime
                    BatchRow.Cells("ItemsProcessed").Value = dblItemsProcessed
                    BatchRow.Cells("ItemsExported").Value = dblItemsExported
                    BatchRow.Cells("ItemsFailed").Value = dblItemsFailed
                    BatchRow.Cells("ItemsSkipped").Value = dblItemsSkipped
                    BatchRow.Cells("ProcessingFilesDirectory").Value = sProcessingFilesDirectory
                    BatchRow.Cells("CaseDirectory").Value = sCaseDirectory
                    BatchRow.Cells("OutputDirectory").Value = sOutputDirectory
                    BatchRow.Cells("LogDirectory").Value = sLogDirectory
                    BatchRow.Cells("SummaryReportLocation").Value = sSummaryReportLocation
                End With
            End While

            sqlConnection.Close()
            grdArchiveBatches.ScrollBars = ScrollBars.Both

        Catch ex As Exception
            common.Logger(psIngestionLogFile, ex.ToString)
        End Try
        blnLoadGridFromDB = True

    End Function

    Private Sub cboExtractionOutputType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboExtractionOutputType.SelectedIndexChanged
        grpOutputFormat.Enabled = True

        If radUser.Checked = True Then
            If ((cboExtractionOutputType.Text = "eml") Or (cboExtractionOutputType.Text = "msg")) Then
                radZipped.Enabled = True
                radLoose.Enabled = True
                radLoose.Checked = True
            ElseIf (cboExtractionOutputType.Text = "pst") Then
                radZipped.Enabled = False
                radLoose.Enabled = False
                radZipped.Checked = False
                radLoose.Checked = False
            End If
        Else
            radZipped.Enabled = False
            radLoose.Enabled = False
            radZipped.Checked = False
            radLoose.Checked = False
        End If
    End Sub

    Private Sub chkAFEX_CheckedChanged(sender As Object, e As EventArgs) Handles chkAFEX.CheckedChanged
        If chkAFEX.Checked = True Then
            chkAFEX.BackColor = Color.Green
            chkAFEX.Checked = True
        ElseIf chkAFEX.Checked = False Then
            chkAFEX.BackColor = Color.LightGray
            chkAFEX.Checked = False
        End If
    End Sub

    Private Sub chkAFPST_CheckedChanged(sender As Object, e As EventArgs) Handles chkAFPST.CheckedChanged
        If chkAFPST.Checked = True Then
            chkAFPST.BackColor = Color.Green
            chkAFPST.Checked = True
        ElseIf chkAFPST.Checked = False Then
            chkAFPST.BackColor = Color.LightGray
            chkAFPST.Checked = False
        End If
    End Sub

    Private Sub chkAFSYS_CheckedChanged(sender As Object, e As EventArgs) Handles chkAFSYS.CheckedChanged
        If chkAFSYS.Checked = True Then
            chkAFSYS.BackColor = Color.Green
            chkAFSYS.Checked = True
        ElseIf chkAFSYS.Checked = False Then
            chkAFSYS.BackColor = Color.LightGray
            chkAFSYS.Checked = False
        End If
    End Sub

    Private Sub btnTestSQLConnection_Click(sender As Object, e As EventArgs) Handles btnTestSQLConnection.Click
        'Create ADO.NET objects.
        Dim sHostName As String
        Dim sPort As String
        Dim sDBName As String
        Dim sUserName As String
        Dim sUserInfo As String
        Dim bStatus As Boolean
        Dim sDomain As String

        Dim dbService As New DatabaseService

        sHostName = txtSQLHostName.Text
        sPort = txtSQLPortNumber.Text
        sDBName = txtSQLDBName.Text
        sUserName = txtSQLUserName.Text
        sUserInfo = txtSQLInfo.Text

        If cboSecurityType.Text = "Windows Authentication" Then
            If sHostName = vbNullString Then
                MessageBox.Show("You must enter a SQL Host Name", "Enter SQL Host Name", MessageBoxButtons.OK)
                txtSQLHostName.Focus()
                Exit Sub

            End If

            If sPort = vbNullString Then
                MessageBox.Show("You must enter a Port Number", "Enter Port Number", MessageBoxButtons.OK)
                txtSQLPortNumber.Focus()
                Exit Sub

            End If

            If sDBName = vbNullString Then
                MessageBox.Show("You must enter a Database Name", "Enter DB Name", MessageBoxButtons.OK)
                txtSQLDBName.Focus()
                Exit Sub

            End If

        ElseIf cboSecurityType.Text = "SQLServer Authentication" Then
            If sHostName = vbNullString Then
                MessageBox.Show("You must enter a SQL Host Name", "Enter SQL Host Name", MessageBoxButtons.OK)
                txtSQLHostName.Focus()
                Exit Sub

            End If

            If sPort = vbNullString Then
                MessageBox.Show("You must enter a Port Number", "Enter Port Number", MessageBoxButtons.OK)
                txtSQLPortNumber.Focus()
                Exit Sub

            End If

            If sDBName = vbNullString Then
                MessageBox.Show("You must enter a Database Name", "Enter DB Name", MessageBoxButtons.OK)
                txtSQLDBName.Focus()
                Exit Sub

            End If
            If sUserName = vbNullString Then
                MessageBox.Show("You must enter a SQL UserName", "Enter SQL UserName", MessageBoxButtons.OK)
                txtSQLUserName.Focus()
                Exit Sub

            End If

            If sUserInfo = vbNullString Then
                MessageBox.Show("You must enter a SQL DB Password", "Enter SQL DB Password", MessageBoxButtons.OK)
                txtSQLInfo.Focus()
                Exit Sub
            End If
        Else

        End If

        bStatus = dbService.CheckSQLConnection(cboSecurityType.Text, sHostName, sPort, sDBName, sUserName, sUserInfo, sDomain)

        If bStatus = True Then
            lblSQLPWCheck.Text = "!"
            lblSQLPWCheck.ForeColor = Color.Green
            lblHostCheck.Text = "!"
            lblHostCheck.ForeColor = Color.Green
            lblSQLUserNameCheck.Text = "!"
            lblSQLUserNameCheck.ForeColor = Color.Green
            lblDBNameCheck.Text = "!"
            lblDBNameCheck.ForeColor = Color.Green
            lblDomainCheck.Text = "!"
            lblDomainCheck.ForeColor = Color.Green
        Else
            lblSQLPWCheck.Text = "X"
            lblSQLPWCheck.ForeColor = Color.Red
            lblHostCheck.Text = "X"
            lblHostCheck.ForeColor = Color.Red
            lblSQLUserNameCheck.Text = "X"
            lblSQLUserNameCheck.ForeColor = Color.Red
            lblDBNameCheck.Text = "X"
            lblDBNameCheck.ForeColor = Color.Red
        End If

    End Sub
    Private Sub chkEmail_CheckedChanged(sender As Object, e As EventArgs) Handles chkEmail.CheckedChanged
        If chkEmail.Checked = True Then
            chkEmail.BackColor = Color.Green
        Else
            chkEmail.BackColor = Color.LightGray
        End If
    End Sub

    Private Sub chkRSSFeed_CheckedChanged(sender As Object, e As EventArgs) Handles chkRSSFeed.CheckedChanged
        If chkRSSFeed.Checked = True Then
            chkRSSFeed.BackColor = Color.Green
        Else
            chkRSSFeed.BackColor = Color.LightGray
        End If

    End Sub

    Private Sub chkCalendar_CheckedChanged(sender As Object, e As EventArgs) Handles chkCalendar.CheckedChanged
        If chkCalendar.Checked = True Then
            chkCalendar.BackColor = Color.Green
        Else
            chkCalendar.BackColor = Color.LightGray
        End If

    End Sub

    Private Sub chkContact_CheckedChanged(sender As Object, e As EventArgs) Handles chkContact.CheckedChanged
        If chkContact.Checked = True Then
            chkContact.BackColor = Color.Green
        Else
            chkContact.BackColor = Color.LightGray
        End If

    End Sub

    Private Sub radCentera_CheckedChanged(sender As Object, e As EventArgs) Handles radCentera.CheckedChanged
        Dim Tnode As TreeNode
        Dim drives As System.Collections.ObjectModel.ReadOnlyCollection(Of IO.DriveInfo)
        Dim rootDir As String

        If radAddFolder.Checked = True Then
            chkComputeBatchSize.Enabled = True

            treeViewFolders.Nodes.Clear()
            treeViewFolders.Width = 266
            treeViewFolders.Height = 230

            lblPEAFile.Hide()
            lblIPFile.Hide()
            txtPEAFile.Hide()
            txtIPFile.Hide()
            btnPEAFile.Hide()
            btnIPFile.Hide()
            drives = My.Computer.FileSystem.Drives
            For i As Integer = 0 To drives.Count - 1
                If Not drives(i).IsReady Then
                    Continue For
                End If
                Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
                rootDir = drives(i).Name
                AddAllFolders(Tnode, rootDir)
            Next

        ElseIf radAddFile.Checked = True Then
            chkComputeBatchSize.Enabled = True
            lblPEAFile.Hide()
            lblIPFile.Hide()
            txtPEAFile.Hide()
            txtIPFile.Hide()
            btnPEAFile.Hide()
            btnIPFile.Hide()
            treeViewFolders.Nodes.Clear()

            treeViewFolders.Width = 266
            treeViewFolders.Height = 230

            drives = My.Computer.FileSystem.Drives
            For i As Integer = 0 To drives.Count - 1
                If Not drives(i).IsReady Then
                    Continue For
                End If
                Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
                rootDir = drives(i).Name
                AddAllFoldersAndFiles(Tnode, rootDir)
            Next

        ElseIf radCentera.Checked = True Then

            If cboExtractionOutputType.Text = "pst" And radUser.Checked = True And radMailboxArchive.Checked = True Then
                treeViewFolders.Enabled = False
                treeViewFolders.Height = 170
            Else
                treeViewFolders.Enabled = True
                chkComputeBatchSize.Checked = False

                chkComputeBatchSize.Enabled = False
                treeViewFolders.Nodes.Clear()
                treeViewFolders.Width = 266
                treeViewFolders.Height = 170
                drives = My.Computer.FileSystem.Drives
                For i As Integer = 0 To drives.Count - 1
                    If Not drives(i).IsReady Then
                        Continue For
                    End If
                    Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
                    rootDir = drives(i).Name
                    AddAllFoldersAndCLPFiles(Tnode, rootDir)
                Next

            End If

            lblPEAFile.Show()
            lblIPFile.Show()
            txtPEAFile.Show()
            txtIPFile.Show()
            btnPEAFile.Show()
            btnIPFile.Show()
        End If
    End Sub

    Private Sub btnIPFile_Click(sender As Object, e As EventArgs) Handles btnIPFile.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        With OpenFileDialog1
            .Filter = "*.ipf|*.ipf"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            txtIPFile.Text = OpenFileDialog1.FileName.ToString
        End If
    End Sub

    Private Sub btnPEAFile_Click(sender As Object, e As EventArgs) Handles btnPEAFile.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        With OpenFileDialog1
            .Filter = "*.pea|*.pea"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            txtPEAFile.Text = OpenFileDialog1.FileName.ToString
        End If
    End Sub

    Private Sub btnShowSettings_MouseHover(sender As Object, e As EventArgs) Handles btnShowSettings.MouseHover
        ArchiveExtractToolTip.Show("Show Settings...", btnShowSettings)
    End Sub

    Private Sub btnLoadPreviousBatches_MouseHover(sender As Object, e As EventArgs) Handles btnLoadPreviousBatches.MouseHover
        ArchiveExtractToolTip.Show("Reload Previous Batches", btnLoadPreviousBatches)

    End Sub

    Private Sub btnAddBatchToGrid_MouseHover(sender As Object, e As EventArgs) Handles btnAddBatchToGrid.MouseHover
        ArchiveExtractToolTip.Show("Add Processing Batch to Grid", btnAddBatchToGrid)

    End Sub

    Private Sub btnLaunchBatches_MouseHover(sender As Object, e As EventArgs) Handles btnLaunchBatches.MouseHover
        ArchiveExtractToolTip.Show("Launch Extraction Batches", btnLaunchBatches)

    End Sub

    Private Sub btnEVUserListFileChooser_Click(sender As Object, e As EventArgs) Handles btnEVUserListFileChooser.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        With OpenFileDialog1
            .Filter = "*.csv|*.csv"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            txtEVUserList.Text = OpenFileDialog1.FileName.ToString
        End If
    End Sub

    Private Sub btnExportProcessingDetails_Click(sender As Object, e As EventArgs) Handles btnExportProcessingDetails.Click
        Dim ReportOutputFile As StreamWriter
        Dim sOutputFileName As String
        Dim sMachineName As String

        Dim sBatchName As String
        Dim sPercentCompleted As String
        Dim sBytesProcessed As String
        Dim sTotalBytes As String
        Dim sExtractionStatus As String
        Dim sItemsProcessed As String
        Dim sArchiveName As String
        Dim sArchiveSettings As String
        Dim sArchiveType As String
        Dim sLightspeedOutputType As String
        Dim sSQLConnectionInfo As String
        Dim sWSSSettings As String
        Dim sProcessingStartDate As String
        Dim sProcessingEndDate As String
        Dim sItemsExported As String
        Dim sItemsFailed As String
        Dim sItemsSkipped As String
        Dim sProcessingFilesDirectory As String
        Dim sCaseDirectory As String
        Dim sOutputDirectory As String
        Dim sLogDirectory As String
        Dim sSummaryReportLocation As String

        sMachineName = System.Net.Dns.GetHostName()
        sOutputFileName = "Legacy Archive Extraction Processing Statistic - " & sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv"
        ReportOutputFile = New StreamWriter(eMailArchiveMigrationManager.NuixFilesDir & "\" & sOutputFileName)

        ReportOutputFile.WriteLine("Batch Name, Percent Completed, Bytes Processed, Total Bytes, Extraction Status, Items Processed, Archive Name, Archive Settings, Archive Type, Output Type, SQL Connection Info, WSS Settings, Start Time, End Time, Success, Failed, Skipped, Processing Files Directory, Case Directory, Output Directory, Log Directory, Summary Report Location")

        For Each row In grdArchiveExtractionBatch.Rows
            sBatchName = row.cells("BatchName").value
            If sBatchName <> vbNullString Then

                sPercentCompleted = row.cells("PercentCompleted").value
                sBytesProcessed = row.cells("BytesProcessed").value
                sTotalBytes = row.cells("TotalBytes").value
                sExtractionStatus = row.cells("ExtractionStatus").value
                sItemsProcessed = row.cells("ItemsProcessed").value
                sArchiveName = row.cells("ArchiveName").value
                sArchiveSettings = row.cells("ArchiveSettings").value
                sArchiveType = row.cells("ArchiveType").value
                sLightspeedOutputType = row.cells("LightspeedOutputType").value
                sSQLConnectionInfo = row.cells("SQLConnectionInfo").value
                sWSSSettings = row.cells("SourceInformation").value
                sProcessingStartDate = row.cells("ProcessingStartDate").value
                sProcessingEndDate = row.cells("ProcessingEndDate").value
                sItemsExported = row.cells("ItemsExported").value
                sItemsFailed = row.cells("ItemsFailed").value
                sItemsSkipped = row.cells("ItemsSkipped").value
                sProcessingFilesDirectory = row.cells("ProcessingFilesDirectory").value
                sCaseDirectory = row.cells("CaseDirectory").value
                sOutputDirectory = row.cells("OutputDirectory").value
                sLogDirectory = row.cells("LogDirectory").value
                sSummaryReportLocation = row.cells("SummaryReportLocation").value

                ReportOutputFile.WriteLine("""" & sBatchName & """" & "," & """" & sPercentCompleted & """" & "," & """" & sBytesProcessed & """" & "," & """" & sTotalBytes & """" & "," & """" & sExtractionStatus & """" & "," & """" & sItemsProcessed & """" & "," & """" & sArchiveName & """" & "," & """" & sArchiveSettings & """" & "," & """" & sArchiveType & """" & "," & """" & sLightspeedOutputType & """" & "," & """" & sSQLConnectionInfo & """" & "," & """" & sWSSSettings & """" & "," & """" & sProcessingStartDate & """" & "," & """" & sProcessingEndDate & """" & "," & """" & sItemsExported & """" & "," & """" & sItemsFailed & """" & "," & """" & sItemsSkipped & """" & "," & """" & sProcessingFilesDirectory & """" & "," & """" & sCaseDirectory & """" & "," & """" & sOutputDirectory & """" & "," & """" & sLogDirectory & """" & "," & """" & sSummaryReportLocation & """")
            End If
        Next

        ReportOutputFile.Close()
        MessageBox.Show("Nuix Case Statistics report finished building.  Report located at: " & eMailArchiveMigrationManager.NuixFilesDir & "\" & sOutputFileName)
    End Sub

    Private Sub btnConsolidateExporterFiles_Click(sender As Object, e As EventArgs) Handles btnConsolidateExporterFiles.Click
        ConsolidateExporterFiles.Show(btnConsolidateExporterFiles, 0, btnConsolidateExporterFiles.Height)
    End Sub

    Private Sub ConsolidateExporterErrorsFilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConolidateExporterErrorsFilesToolStripMenuItem.Click
        Dim bStatus As Boolean
        Dim lstExporterErrorsFiles As List(Of String)
        Dim ExporterErrorsFile As StreamReader
        Dim ConsolidateExporterErrorsFile As StreamWriter
        Dim sCurrentRow As String
        Dim bFirstErrorFile As Boolean

        bFirstErrorFile = True
        lstExporterErrorsFiles = New List(Of String)
        bStatus = blnGetAllExporterErrors(eMailArchiveMigrationManager.NuixCaseDir & "\Archive Extract", lstExporterErrorsFiles)

        ConsolidateExporterErrorsFile = New StreamWriter(eMailArchiveMigrationManager.NuixFilesDir & "\Archive Extraction\consolidated-exporter-errors" & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv")

        If lstExporterErrorsFiles.Count > 0 Then
            For Each ErrorFile In lstExporterErrorsFiles
                ExporterErrorsFile = New StreamReader(ErrorFile.ToString)
                While Not ExporterErrorsFile.EndOfStream
                    sCurrentRow = ExporterErrorsFile.ReadLine
                    If bFirstErrorFile = True Then
                        ConsolidateExporterErrorsFile.WriteLine(sCurrentRow)
                        bFirstErrorFile = False
                    Else
                        If ((sCurrentRow <> "MAILBOX NAME,ERROR,COUNT,SOURCE") And (sCurrentRow <> vbNullString)) Then
                            ConsolidateExporterErrorsFile.WriteLine(sCurrentRow)
                        End If
                    End If
                End While
            Next
        End If
        ConsolidateExporterErrorsFile.Close()
        MessageBox.Show("Consolidated exporter-errors files completed.  Report located at - " & eMailArchiveMigrationManager.NuixFilesDir & "\consolidated-exporter-errors" & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv")

    End Sub

    Private Sub ConsolidateExporterMetricsFilesToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ConsolidateExporterMetricsFilesToolStripMenuItem.Click
        Dim bStatus As Boolean
        Dim lstExporterMetricsFiles As List(Of String)
        Dim ExporterMetricFile As StreamReader
        Dim ConsolidateExporterMetricsFile As StreamWriter
        Dim sCurrentRow As String
        Dim bFirstMetricsFile As Boolean

        bFirstMetricsFile = True
        lstExporterMetricsFiles = New List(Of String)
        ConsolidateExporterMetricsFile = New StreamWriter(eMailArchiveMigrationManager.NuixFilesDir & "\Archive Extraction\consolidated-exporter-metrics" & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv")

        bStatus = blnGetAllExporterMetrics(eMailArchiveMigrationManager.NuixCaseDir & "\Archive Extract", lstExporterMetricsFiles)

        If lstExporterMetricsFiles.Count > 0 Then
            For Each MetricsFile In lstExporterMetricsFiles
                ExporterMetricFile = New StreamReader(MetricsFile.ToString)
                While Not ExporterMetricFile.EndOfStream
                    sCurrentRow = ExporterMetricFile.ReadLine
                    If bFirstMetricsFile = True Then
                        ConsolidateExporterMetricsFile.WriteLine(sCurrentRow)
                        bFirstMetricsFile = False
                    Else
                        If ((sCurrentRow <> "MAILBOX NAME,ITEMS_EXPORTED,ITEMS_FAILED,BYTES_UPLOADED,FOLDERS_FOUND,FOLDERS_CREATED,SOURCE") And (sCurrentRow <> vbNullString)) Then
                            ConsolidateExporterMetricsFile.WriteLine(sCurrentRow)
                        End If
                    End If
                End While
            Next
        End If

        ConsolidateExporterMetricsFile.Close()
        MessageBox.Show("Consolidated exporter-metrics files completed.  Report located at - " & eMailArchiveMigrationManager.NuixFilesDir & "\consolidated-exporter-metrics" & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv")

    End Sub

    Private Function blnBuildArchiveExtractionFiles() As Boolean

        Dim dbService As New DatabaseService
        Dim NuixConsoleProcessStartInfo As ProcessStartInfo
        Dim NuixConsoleProcess As Process
        Dim bStatus As Boolean
        Dim sBatchFileName As String
        Dim sRubyFileName As String
        Dim sSQLAuthentication As String
        Dim sSQLDomain As String
        Dim sSQLUserName As String
        Dim sSQLAdminInfo As String
        Dim sSQLHostName As String
        Dim sRunDate As String
        Dim sNuixAppMemory As String
        Dim sSQLConnectionInfo As String
        Dim sSQLDBName As String
        Dim asSQLConnectionInfo() As String
        Dim asArchiveSettings() As String
        Dim sArchiveSettings As String
        Dim sFailOnPartitionErrors As String
        Dim sEnableMetadataQuery As String
        Dim sNuixConsolePath As String
        Dim sCaseName As String
        Dim sSearchTermCSV As String
        Dim sProcessingType As String
        Dim sWSSSettings As String
        Dim asWSSSettings() As String
        Dim sSourceFolders As String
        Dim sOutputFormat As String
        Dim sMigration As String
        Dim sArchiveName As String
        Dim sAddressFiltering As String
        Dim sExpandDLTo As String
        Dim sDocStoreID As String
        Dim sSkipSISIDLookup As String
        Dim sArchiveType As String
        Dim sSISFolders As String
        Dim sWSSJsonFileName As String
        Dim sCaseJSoNFileName As String
        Dim asAXSOneSettings() As String
        Dim sEvidenceKeyStore As String
        Dim sCustodianMappingFile As String
        Dim sReportPath As String
        Dim sProcessingFilesDir As String
        Dim sManifestFile As String
        Dim sPSTMappingFile As String
        Dim sExportDir As String
        Dim sNuixCaseDir As String
        Dim sFromDate As String
        Dim sToDate As String
        Dim sBatchName As String
        Dim sItemsProcessed As String
        Dim sProcessEndTime As String
        Dim sProcessStartTime As String
        Dim sTotalItemsSkipped As String
        Dim sTotalItemsProcessed As String
        Dim sTotalItemsFailed As String
        Dim sTotalItemsExported As String
        Dim sSourceSize As String
        Dim sEVUserList As String
        Dim sPercentCompleted As String
        Dim sBytesProcessed As String
        Dim asOutputTypes() As String
        Dim sFileStore As String
        Dim sLogDirectory As String
        Dim sWorkerTempDir As String
        Dim sCenteraPEAFile As String
        Dim sCenteraClipFile As String
        Dim sCenteraIPFile As String
        Dim asCenteraFiles() As String
        Dim sSQLPort As String
        Dim common As New Common
        Dim iNuixConsoleVersion As Decimal

        Dim bCentera As Boolean

        Try
            sWorkerTempDir = psWorkerTempDir

            sMigration = "lightspeed"
            sEvidenceKeyStore = ""

            Dim lstWSSSettings As List(Of String)
            lstWSSSettings = New List(Of String)

            For Each folder In plstSourceFolders
                sSourceFolders = sSourceFolders & folder.ToString & "|"
            Next
            'remove the last | from the string
            If sSourceFolders <> vbNullString Then
                sSourceFolders = sSourceFolders.Remove(sSourceFolders.Length - 1)
            End If

            For Each row In grdArchiveExtractionBatch.Rows
                If row.cells("SelectBatch").value = True Then
                    sRunDate = DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss").ToString
                    Invoke(Sub()
                               row.cells("ExtractionStatus").value = "In Progress"
                               row.DefaultCellStyle.Forecolor = Color.Black
                           End Sub)
                    Invoke(Sub()
                               row.cells("StatusImage").value = My.Resources.inprogress_medium
                           End Sub)
                    Invoke(Sub()
                               row.cells("SelectBatch").value = False
                           End Sub)
                    Invoke(Sub()
                               row.readonly = True
                           End Sub)
                    Invoke(Sub()
                               sSourceFolders = row.cells("SourceInformation").value
                               If sSourceFolders <> vbNullString Then
                                   If sSourceFolders.Contains("CenteraClip:") Then
                                       bCentera = True
                                   Else
                                       bCentera = False
                                   End If
                               Else
                                   bCentera = False
                               End If
                           End Sub)
                    Invoke(Sub()
                               sSourceSize = row.cells("TotalBytes").value
                           End Sub)
                    Invoke(Sub()
                               sBatchName = row.cells("BatchName").value
                           End Sub)
                    Invoke(Sub()
                               sArchiveName = row.cells("ArchiveName").value
                           End Sub)
                    Invoke(Sub()
                               sArchiveType = row.cells("ArchiveType").value
                           End Sub)
                    Invoke(Sub()
                               sSQLConnectionInfo = row.Cells("SQLConnectionInfo").value
                           End Sub)
                    Invoke(Sub()
                               sOutputFormat = row.Cells("LightspeedOutputType").value
                           End Sub)
                    Invoke(Sub()
                               sArchiveSettings = row.Cells("ArchiveSettings").value
                           End Sub)
                    Invoke(Sub()
                               sCaseName = row.Cells("BatchName").value
                           End Sub)
                    Invoke(Sub()
                               sWSSSettings = row.cells("WSSSettings").value
                           End Sub)
                    Invoke(Sub()
                               sProcessingFilesDir = row.cells("ProcessingFilesDirectory").value
                           End Sub)
                    Invoke(Sub()
                               sNuixCaseDir = row.cells("CaseDirectory").value
                           End Sub)
                    Invoke(Sub()
                               sExportDir = row.cells("OutputDirectory").value
                           End Sub)
                    Invoke(Sub()
                               sLogDirectory = row.Cells("LogDirectory").value
                           End Sub)
                    sProcessingType = "lightspeed"
                    If sWSSSettings <> vbNullString Then
                        asWSSSettings = Split(sWSSSettings, "|")
                        For iCounter = 0 To UBound(asWSSSettings)
                            If asWSSSettings(iCounter).Contains("mappingCSV") Then
                                sCustodianMappingFile = asWSSSettings(iCounter).Replace("mappingCSV:", "")
                            ElseIf asWSSSettings(iCounter).Contains("searchTermCSV") Then
                                sSearchTermCSV = asWSSSettings(iCounter).Replace("searchTermCSV:", "")
                            End If
                            lstWSSSettings.Add(asWSSSettings(iCounter))
                        Next
                    End If

                    If sOutputFormat.Contains("|") Then
                        asOutputTypes = Split(sOutputFormat, "|")
                        sOutputFormat = asOutputTypes(0)
                        sFileStore = asOutputTypes(1)
                    End If

                    sNuixConsolePath = psNuixAppLocation.Replace("nuix_console.exe", "")
                    Dim regex As Regex = New Regex("\d+.\d+")
                    Dim match As Match = regex.Match(sNuixConsolePath)
                    If match.Success Then
                        iNuixConsoleVersion = match.Value
                    End If

                    Invoke(Sub()
                               sRubyFileName = sProcessingFilesDir & "\" & row.cells("BatchName").value & ".rb"
                           End Sub)
                    Invoke(Sub()
                               sBatchFileName = sProcessingFilesDir & "\" & row.Cells("BatchName").value & ".bat"
                           End Sub)

                    If sSQLConnectionInfo <> vbNullString Then
                        asSQLConnectionInfo = Split(sSQLConnectionInfo, "|")
                        sSQLAuthentication = asSQLConnectionInfo(0).Replace("Authentication:", "")
                        sSQLDomain = asSQLConnectionInfo(1).Replace("Domain:", "")
                        sSQLHostName = asSQLConnectionInfo(2).Replace("HostName:", "") & ":" & asSQLConnectionInfo(3).Replace("PortNumber:", "")
                        sSQLDBName = asSQLConnectionInfo(4).Replace("DBName:", "")
                        sSQLUserName = asSQLConnectionInfo(5).Replace("SQLUsername:", "")
                        sSQLAdminInfo = asSQLConnectionInfo(6).Replace("SQLINfo:", "")
                    End If

                    Select Case sArchiveName
                        Case "Veritas Enterprise Vault"
                            If bCentera = True Then
                                asCenteraFiles = Split(sSourceFolders, "|")
                                For iCounter = 0 To UBound(asCenteraFiles)
                                    If asCenteraFiles(iCounter).Contains("CenteraPEA") Then
                                        sCenteraPEAFile = asCenteraFiles(iCounter).Replace("CenteraPEA:", "")
                                        Try
                                            Environment.SetEnvironmentVariable("CENTERA_PEA_LOCATION", sCenteraPEAFile, EnvironmentVariableTarget.Machine)
                                        Catch ex As Exception
                                            Try
                                                Environment.SetEnvironmentVariable("CENTERA_PEA_LOCATION", sCenteraPEAFile, EnvironmentVariableTarget.User)
                                            Catch ex1 As Exception

                                            End Try

                                        End Try
                                    ElseIf asCenteraFiles(iCounter).Contains("CenteraClip") Then
                                        sCenteraClipFile = asCenteraFiles(iCounter).Replace("CenteraClip:", "")
                                    ElseIf asCenteraFiles(iCounter).Contains("CenteraIP") Then
                                        sCenteraIPFile = asCenteraFiles(iCounter).Replace("CenteraIP:", "")
                                    End If
                                Next
                            End If

                            sCaseJSoNFileName = sProcessingFilesDir & "\" & sCaseName & ".json"
                            bStatus = blnBuildArchiveExtractionJSon(sCaseJSoNFileName, sSourceFolders, sEvidenceKeyStore, sCaseName, sMigration, sSourceSize, bCentera)

                            asArchiveSettings = Split(sArchiveSettings, "|")
                            sEnableMetadataQuery = asArchiveSettings(0).Replace("SkipAdditionalSQLLookups:", "")
                            If sEnableMetadataQuery = "true" Then
                                sEnableMetadataQuery = "false"
                            Else
                                sEnableMetadataQuery = "true"
                            End If

                            sFailOnPartitionErrors = asArchiveSettings(1).Replace("SkipVaultStorePartitionErrors:", "")
                            If (sFailOnPartitionErrors = "true") Then
                                sFailOnPartitionErrors = "true"
                            Else
                                sFailOnPartitionErrors = "false"
                            End If
                            sFromDate = asArchiveSettings(2).Replace("FromDate:", "")
                            sToDate = asArchiveSettings(3).Replace("ToDate:", "")
                            sEVUserList = asArchiveSettings(4).Replace("EVUserList:", "")

                            If (sArchiveType = "Mailbox:User") Then
                                sPSTMappingFile = sProcessingFilesDir & "\PST_Mapping.csv"
                                sManifestFile = sProcessingFilesDir & "\manifestfile.csv"
                                sReportPath = sProcessingFilesDir & "\EV_SQLReports"
                                System.IO.Directory.CreateDirectory(sReportPath)
                                bStatus = blnBuildArchiveEVUserExtractionRuby(sNuixCaseDir, sCaseName, sProcessingFilesDir, sRubyFileName, psNumberOfWorkers, psMemoryPerWorker, psSQLiteLocation, sNuixConsolePath, sCaseJSoNFileName, CDbl(sSourceSize), sWorkerTempDir, pbExtractFromSlackSpace, sSQLAuthentication)
                            ElseIf (sArchiveType = "Mailbox:Flat") Then
                                bStatus = blnBuildArchiveFlatExtractionRubyScript(sArchiveName, sProcessingFilesDir, sSourceFolders, sCaseJSoNFileName, sNuixCaseDir, sRubyFileName, psSQLiteLocation, psNumberOfWorkers, psMemoryPerWorker, CDbl(sSourceSize), sWorkerTempDir, pbExtractFromSlackSpace, sSQLAuthentication)
                            ElseIf (sArchiveType = "Journal:User") Then
                                bStatus = blnBuildArchiveUserExtractionRubyScript(sArchiveName, sProcessingFilesDir, sSourceFolders, sRubyFileName, sCaseJSoNFileName, sNuixCaseDir, psSQLiteLocation, psNumberOfWorkers, psMemoryPerWorker, CDbl(sSourceSize), sWorkerTempDir, pbExtractFromSlackSpace, sSQLAuthentication)
                                sWSSJsonFileName = sProcessingFilesDir & "\wss.json"
                                '                            bStatus = blnBuildArchiveExtractionWSSJSon(sWSSJsonFileName, lstWSSSettings, sProcessingType, sNuixCaseDir)
                                bStatus = blnBuildFinalArchiveExtractionWSSJSon(sWSSJsonFileName, sArchiveName, lstWSSSettings, sProcessingType, sNuixCaseDir)
                                'bStatus = blnBuildWSSRubyScript(sProcessingFilesDir & "\wss.rb", psNuixFilesDir)
                                bStatus = blnBuildFinalWSSRubyScript(sProcessingFilesDir & "\wss.rb", psNuixFilesDir)
                            ElseIf (sArchiveType = "Journal:Flat") Then
                                bStatus = blnBuildArchiveFlatExtractionRubyScript(sArchiveName, sProcessingFilesDir, sSourceFolders, sCaseJSoNFileName, sNuixCaseDir, sRubyFileName, psSQLiteLocation, psNumberOfWorkers, psMemoryPerWorker, CDbl(sSourceSize), sWorkerTempDir, pbExtractFromSlackSpace, sSQLAuthentication)
                            End If

                        Case "EMC EmailXtender"
                            If bCentera = True Then
                                asCenteraFiles = Split(sSourceFolders, "|")
                                For iCounter = 0 To UBound(asCenteraFiles)
                                    If asCenteraFiles(iCounter).Contains("CenteraPEA") Then
                                        sCenteraPEAFile = asCenteraFiles(iCounter).Replace("CenteraPEA:", "")
                                        Try
                                            Environment.SetEnvironmentVariable("CENTERA_PEA_LOCATION", sCenteraPEAFile, EnvironmentVariableTarget.Machine)
                                        Catch ex As Exception
                                            Try
                                                Environment.SetEnvironmentVariable("CENTERA_PEA_LOCATION", sCenteraPEAFile, EnvironmentVariableTarget.User)
                                            Catch ex1 As Exception

                                            End Try

                                        End Try
                                    ElseIf asCenteraFiles(iCounter).Contains("CenteraClip") Then
                                        sCenteraClipFile = asCenteraFiles(iCounter).Replace("CenteraClip:", "")
                                    ElseIf asCenteraFiles(iCounter).Contains("CenteraIP") Then
                                        sCenteraIPFile = asCenteraFiles(iCounter).Replace("CenteraIP:", "")
                                    End If
                                Next
                            End If
                            sCaseJSoNFileName = sProcessingFilesDir & "\" & sCaseName & ".json"
                            bStatus = blnBuildArchiveExtractionJSon(sCaseJSoNFileName, sSourceFolders, sEvidenceKeyStore, sCaseName, sMigration, sSourceSize, bCentera)

                            asArchiveSettings = Split(sArchiveSettings, "|")
                            sAddressFiltering = asArchiveSettings(0).Replace("AddressFilter:", "")
                            sExpandDLTo = asArchiveSettings(1).Replace("ExpandedDLLocation:", "").Trim
                            sFromDate = asArchiveSettings(2).Replace("FromDate:", "")
                            sToDate = asArchiveSettings(3).Replace("ToDate:", "")

                            If (sArchiveType = "Journal:User") Then
                                bStatus = blnBuildArchiveUserExtractionRubyScript(sArchiveName, sProcessingFilesDir, sSourceFolders, sRubyFileName, sCaseJSoNFileName, sNuixCaseDir, psSQLiteLocation, psNumberOfWorkers, psMemoryPerWorker, CDbl(sSourceSize), sWorkerTempDir, pbExtractFromSlackSpace, sSQLAuthentication)
                                sWSSJsonFileName = sProcessingFilesDir & "\wss.json"
                                'bStatus = blnBuildArchiveExtractionWSSJSon(sWSSJsonFileName, lstWSSSettings, sProcessingType, sNuixCaseDir)
                                bStatus = blnBuildFinalArchiveExtractionWSSJSon(sWSSJsonFileName, sArchiveName, lstWSSSettings, sProcessingType, sNuixCaseDir)

                                'bStatus = blnBuildWSSRubyScript(sProcessingFilesDir & "\wss.rb", psNuixFilesDir)
                                bStatus = blnBuildFinalWSSRubyScript(sProcessingFilesDir & "\wss.rb", psNuixFilesDir)
                            ElseIf (sArchiveType = "Journal:Flat") Then
                                bStatus = blnBuildArchiveFlatExtractionRubyScript(sArchiveName, sProcessingFilesDir, sSourceFolders, sCaseJSoNFileName, sNuixCaseDir, sRubyFileName, psSQLiteLocation, psNumberOfWorkers, psMemoryPerWorker, CDbl(sSourceSize), sWorkerTempDir, pbExtractFromSlackSpace, sSQLAuthentication)
                            End If
                        Case "EMC SourceOne"
                            If bCentera = True Then
                                asCenteraFiles = Split(sSourceFolders, "|")
                                For iCounter = 0 To UBound(asCenteraFiles)
                                    If asCenteraFiles(iCounter).Contains("CenteraPEA") Then
                                        sCenteraPEAFile = asCenteraFiles(iCounter).Replace("CenteraPEA:", "")
                                        Try
                                            Environment.SetEnvironmentVariable("CENTERA_PEA_LOCATION", sCenteraPEAFile, EnvironmentVariableTarget.Machine)
                                        Catch ex As Exception
                                            Try
                                                Environment.SetEnvironmentVariable("CENTERA_PEA_LOCATION", sCenteraPEAFile, EnvironmentVariableTarget.User)
                                            Catch ex1 As Exception

                                            End Try

                                        End Try
                                    ElseIf asCenteraFiles(iCounter).Contains("CenteraClip") Then
                                        sCenteraClipFile = asCenteraFiles(iCounter).Replace("CenteraClip:", "")
                                    ElseIf asCenteraFiles(iCounter).Contains("CenteraIP") Then
                                        sCenteraIPFile = asCenteraFiles(iCounter).Replace("CenteraIP:", "")
                                    End If
                                Next
                            End If
                            sCaseJSoNFileName = sProcessingFilesDir & "\" & sCaseName & ".json"
                            bStatus = blnBuildArchiveExtractionJSon(sCaseJSoNFileName, sSourceFolders, sEvidenceKeyStore, sCaseName, sMigration, sSourceSize, bCentera)

                            asArchiveSettings = Split(sArchiveSettings, "|")
                            sAddressFiltering = asArchiveSettings(0).Replace("AddressFilter:", "")
                            sExpandDLTo = asArchiveSettings(1).Replace("ExpandedDLLocation:", "").Trim
                            sFromDate = asArchiveSettings(2).Replace("FromDate:", "")
                            sToDate = asArchiveSettings(3).Replace("ToDate:", "")
                            If (sArchiveType = "Journal:User") Then
                                bStatus = blnBuildArchiveUserExtractionRubyScript(sArchiveName, sProcessingFilesDir, sSourceFolders, sRubyFileName, sCaseJSoNFileName, sNuixCaseDir, psSQLiteLocation, psNumberOfWorkers, psMemoryPerWorker, CDbl(sSourceSize), sWorkerTempDir, pbExtractFromSlackSpace, sSQLAuthentication)
                                sWSSJsonFileName = sProcessingFilesDir & "\wss.json"
                                '                            bStatus = blnBuildArchiveExtractionWSSJSon(sWSSJsonFileName, lstWSSSettings, sProcessingType, sNuixCaseDir)
                                bStatus = blnBuildFinalArchiveExtractionWSSJSon(sWSSJsonFileName, sArchiveName, lstWSSSettings, sProcessingType, sNuixCaseDir)

                                'bStatus = blnBuildWSSRubyScript(sProcessingFilesDir & "\wss.rb", psNuixFilesDir)
                                bStatus = blnBuildFinalWSSRubyScript(sProcessingFilesDir & "\wss.rb", psNuixFilesDir)
                            ElseIf (sArchiveType = "Journal:Flat") Then
                                bStatus = blnBuildArchiveFlatExtractionRubyScript(sArchiveName, sProcessingFilesDir, sSourceFolders, sCaseJSoNFileName, sNuixCaseDir, sRubyFileName, psSQLiteLocation, psNumberOfWorkers, psMemoryPerWorker, CDbl(sSourceSize), sWorkerTempDir, pbExtractFromSlackSpace, sSQLAuthentication)
                            End If
                        Case "HP/Autonomy EAS"
                            If bCentera = True Then
                                asCenteraFiles = Split(sSourceFolders, "|")
                                For iCounter = 0 To UBound(asCenteraFiles)
                                    If asCenteraFiles(iCounter).Contains("CenteraPEA") Then
                                        sCenteraPEAFile = asCenteraFiles(iCounter).Replace("CenteraPEA:", "")
                                        Try
                                            Environment.SetEnvironmentVariable("CENTERA_PEA_LOCATION", sCenteraPEAFile, EnvironmentVariableTarget.Machine)
                                        Catch ex As Exception
                                            Try
                                                Environment.SetEnvironmentVariable("CENTERA_PEA_LOCATION", sCenteraPEAFile, EnvironmentVariableTarget.User)
                                            Catch ex1 As Exception

                                            End Try

                                        End Try
                                    ElseIf asCenteraFiles(iCounter).Contains("CenteraClip") Then
                                        sCenteraClipFile = asCenteraFiles(iCounter).Replace("CenteraClip:", "")
                                    ElseIf asCenteraFiles(iCounter).Contains("CenteraIP") Then
                                        sCenteraIPFile = asCenteraFiles(iCounter).Replace("CenteraIP:", "")
                                    End If
                                Next
                            End If
                            sCaseJSoNFileName = sProcessingFilesDir & "\" & sCaseName & ".json"
                            bStatus = blnBuildArchiveExtractionJSon(sCaseJSoNFileName, sSourceFolders, sEvidenceKeyStore, sCaseName, sMigration, sSourceSize, bCentera)
                            asArchiveSettings = Split(sArchiveSettings, "|")

                            sDocStoreID = asArchiveSettings(0).Replace("DocStoreID:", "")
                            sFromDate = asArchiveSettings(1).Replace("FromDate:", "")
                            sToDate = asArchiveSettings(2).Replace("ToDate:", "")

                            If (sArchiveType = "Journal:User") Then
                                bStatus = blnBuildArchiveUserExtractionRubyScript(sArchiveName, sProcessingFilesDir, sSourceFolders, sRubyFileName, sCaseJSoNFileName, sCaseName, psSQLiteLocation, psNumberOfWorkers, psMemoryPerWorker, CDbl(sSourceSize), sWorkerTempDir, pbExtractFromSlackSpace, sSQLAuthentication)
                                sWSSJsonFileName = sProcessingFilesDir & "\wss.json"
                                'bStatus = blnBuildArchiveExtractionWSSJSon(sWSSJsonFileName, lstWSSSettings, sProcessingType, sNuixCaseDir)
                                bStatus = blnBuildFinalArchiveExtractionWSSJSon(sWSSJsonFileName, sArchiveName, lstWSSSettings, sProcessingType, sNuixCaseDir)

                                'bStatus = blnBuildWSSRubyScript(sProcessingFilesDir & "\wss.rb", psNuixFilesDir)
                                bStatus = blnBuildFinalWSSRubyScript(sProcessingFilesDir & "\wss.rb", psNuixFilesDir)
                            ElseIf (sArchiveType = "Journal:Flat") Then
                                bStatus = blnBuildArchiveFlatExtractionRubyScript(sArchiveName, sProcessingFilesDir, sSourceFolders, sCaseJSoNFileName, sNuixCaseDir, sRubyFileName, psSQLiteLocation, psNumberOfWorkers, psMemoryPerWorker, CDbl(sSourceSize), sWorkerTempDir, pbExtractFromSlackSpace, sSQLAuthentication)
                            End If
                        Case "Daegis AXS-One"
                            If bCentera = True Then
                                asCenteraFiles = Split(sSourceFolders, "|")
                                For iCounter = 0 To UBound(asCenteraFiles)
                                    If asCenteraFiles(iCounter).Contains("CenteraPEA") Then
                                        sCenteraPEAFile = asCenteraFiles(iCounter).Replace("CenteraPEA:", "")
                                        Try
                                            Environment.SetEnvironmentVariable("CENTERA_PEA_LOCATION", sCenteraPEAFile, EnvironmentVariableTarget.Machine)
                                        Catch ex As Exception
                                            Try
                                                Environment.SetEnvironmentVariable("CENTERA_PEA_LOCATION", sCenteraPEAFile, EnvironmentVariableTarget.User)
                                            Catch ex1 As Exception

                                            End Try

                                        End Try
                                    ElseIf asCenteraFiles(iCounter).Contains("CenteraClip") Then
                                        sCenteraClipFile = asCenteraFiles(iCounter).Replace("CenteraClip:", "")
                                    ElseIf asCenteraFiles(iCounter).Contains("CenteraIP") Then
                                        sCenteraIPFile = asCenteraFiles(iCounter).Replace("CenteraIP:", "")
                                    End If
                                Next
                            End If
                            sCaseJSoNFileName = sProcessingFilesDir & "\" & sCaseName & ".json"
                            bStatus = blnBuildArchiveExtractionJSon(sCaseJSoNFileName, sSourceFolders, sEvidenceKeyStore, sCaseName, sMigration, sSourceSize, bCentera)
                            asAXSOneSettings = Split(sArchiveSettings, "|")
                            sSkipSISIDLookup = asAXSOneSettings(0).Replace("SkipSISlookups:", "")
                            sSISFolders = asAXSOneSettings(1).Replace("SISFolder:", "")
                            sFromDate = asAXSOneSettings(2).Replace("FromDate:", "")
                            sToDate = asAXSOneSettings(3).Replace("ToDate:", "")
                            'If (sArchiveType = "Mailbox:User") Then
                            If (sArchiveType.Contains("User")) Then
                                bStatus = blnBuildArchiveUserExtractionRubyScript(sArchiveName, sProcessingFilesDir, sSourceFolders, sRubyFileName, sCaseJSoNFileName, sNuixCaseDir, psSQLiteLocation, psNumberOfWorkers, psMemoryPerWorker, CDbl(sSourceSize), sWorkerTempDir, pbExtractFromSlackSpace, sSQLAuthentication)
                                sWSSJsonFileName = sProcessingFilesDir & "\wss.json"
                                'bStatus = blnBuildArchiveExtractionWSSJSon(sWSSJsonFileName, lstWSSSettings, sProcessingType, sNuixCaseDir)
                                bStatus = blnBuildFinalArchiveExtractionWSSJSon(sWSSJsonFileName, sArchiveName, lstWSSSettings, sProcessingType, sNuixCaseDir)

                                'bStatus = blnBuildWSSRubyScript(sProcessingFilesDir & "\wss.rb", psNuixFilesDir)
                                bStatus = blnBuildFinalWSSRubyScript(sProcessingFilesDir & "\wss.rb", psNuixFilesDir)
                                'ElseIf (sArchiveType = "Mailbox:Flat") Then
                            ElseIf (sArchiveType.Contains("Flat")) Then
                                bStatus = blnBuildArchiveFlatExtractionRubyScript(sArchiveName, sProcessingFilesDir, sSourceFolders, sCaseJSoNFileName, sNuixCaseDir, sRubyFileName, psSQLiteLocation, psNumberOfWorkers, psMemoryPerWorker, CDbl(sSourceSize), sWorkerTempDir, pbExtractFromSlackSpace, sSQLAuthentication)
                            End If
                    End Select

                    sNuixAppMemory = Math.Round(CInt(psNuixAppMemory) / 1000)
                    sNuixAppMemory = "-Xmx" & sNuixAppMemory & "g"
                    blnBuildArchiveExtractionFiles = False

                    bStatus = common.blnBuildArchiveExtractionBatchFile(psNuixAppLocation, sArchiveName, sCaseName, sBatchFileName, psNMSSourceType, psNMSAddress, psNMSPort, psNMSUserName, psNMSAdminInfo, psLicenseType, psNumberOfWorkers, sNuixAppMemory, sLogDirectory, psWorkerTempDir, sRubyFileName, psJavaTempDir, sExportDir, sExportDir, sSQLUserName, sSQLAdminInfo, sSQLHostName, sPSTMappingFile, sManifestFile, sReportPath, sSQLDBName, sFailOnPartitionErrors, sEnableMetadataQuery, sArchiveType, sDocStoreID, sSourceFolders, sSISFolders, sAddressFiltering, sFromDate, sToDate, sOutputFormat, sSkipSISIDLookup, sWSSJsonFileName, sExpandDLTo, sCustodianMappingFile, piWorkerTimeout, sEVUserList, psPSTExportSize, pbPSTAddDistributionListMetadata, pbEMLAddDistributionListMetadata, sFileStore, sSQLAuthentication, sSQLDomain)

                    sProcessStartTime = DateTime.Now.ToString
                    row.cells("ProcessingStartDate").value = sProcessStartTime
                    bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ProcessStartTime", sProcessStartTime)

                    If sSQLDBName <> vbNullString Then
                        sSQLHostName = asSQLConnectionInfo(2).Replace("HostName:", "")
                        sSQLPort = asSQLConnectionInfo(3).Replace("PortNumber:", "")
                        sSQLDomain = asSQLConnectionInfo(1).Replace("Domain:", "")
                        bStatus = dbService.CheckSQLConnection(cboSecurityType.Text, sSQLHostName, sSQLPort, sSQLDBName, sSQLUserName, sSQLAdminInfo, sSQLDomain)
                        If bStatus = True Then
                            common.Logger(psIngestionLogFile, "Starting " & sBatchFileName & " for Nuix Processing")
                            NuixConsoleProcessStartInfo = New ProcessStartInfo(sBatchFileName)
                            NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            'NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Minimized
                            NuixConsoleProcessStartInfo.CreateNoWindow = True
                            NuixConsoleProcess = System.Diagnostics.Process.Start(NuixConsoleProcessStartInfo)
                            Dim bProcessRunning As Boolean
                            bProcessRunning = True
                            Do While bProcessRunning = True
                                bProcessRunning = blnCheckIfProcessIsRunning(NuixConsoleProcess.Id)
                                bStatus = dbService.GetUpdatedArchiveExtractInfo(psSQLiteLocation, sBatchName, "ItemsProcessed", sItemsProcessed, "INT")
                                If sItemsProcessed = vbNullString Then
                                    row.cells("ItemsProcessed").value = 0
                                Else
                                    row.cells("ItemsProcessed").value = sItemsProcessed
                                End If
                                bStatus = dbService.GetUpdatedArchiveExtractInfo(psSQLiteLocation, sBatchName, "PercentCompleted", sPercentCompleted, "INT")
                                If sPercentCompleted = vbNullString Then
                                    row.cells("PercentCompleted").value = 0
                                Else
                                    row.cells("PercentCompleted").value = sPercentCompleted
                                End If
                                bStatus = dbService.GetUpdatedArchiveExtractInfo(psSQLiteLocation, sBatchName, "BytesProcessed", sBytesProcessed, "INT")
                                If sBytesProcessed = vbNullString Then
                                    row.cells("BytesProcessed").value = 0
                                Else
                                    row.cells("BytesProcessed").value = sBytesProcessed
                                End If
                                Thread.Sleep(5000)
                            Loop
                            common.Logger(psIngestionLogFile, "Checking " & sLogDirectory & " nuix.log for errors.")
                            bStatus = blnCheckNuixLogForErrors(sLogDirectory, sTotalItemsProcessed, sTotalItemsExported, sTotalItemsSkipped, sTotalItemsFailed)
                            common.Logger(psIngestionLogFile, "Return from the Check Nuix Log For errors was " & bStatus)
                            common.Logger(psIngestionLogFile, "Return from the Check Nuix Log For errors Total Items Processed was " & sTotalItemsProcessed)
                            common.Logger(psIngestionLogFile, "Return from the Check Nuix Log For errors Total Items Exported was " & sTotalItemsExported)
                            common.Logger(psIngestionLogFile, "Return from the Check Nuix Log For errors Total Items Skipped was " & sTotalItemsSkipped)
                            common.Logger(psIngestionLogFile, "Return from the Check Nuix Log For errors Total Items Skipped was " & sTotalItemsFailed)
                            If bStatus = False Then
                                If sBatchName <> vbNullString Then
                                    bStatus = blnProcessSummaryReportFiles(sNuixCaseDir & "\summary-report.txt", sTotalItemsProcessed, sTotalItemsExported, sTotalItemsSkipped, sTotalItemsFailed)
                                    sItemsProcessed = vbNullString
                                    bStatus = dbService.GetUpdatedArchiveExtractInfo(psSQLiteLocation, sBatchName, "ItemsProcessed", sItemsProcessed, "INT")
                                    bStatus = dbService.GetUpdatedArchiveExtractInfo(psSQLiteLocation, sBatchName, "ProcessEndTime", sProcessEndTime, "TEXT")
                                    If CDbl(sItemsProcessed) > 0 Then
                                        row.cells("ItemsProcessed").value = sItemsProcessed
                                        sProcessEndTime = DateTime.Now.ToString
                                        row.cells("ProcessingEndDate").value = sProcessEndTime
                                        row.cells("ExtractionStatus").value = "Completed"
                                        row.cells("StatusImage").value = My.Resources.Green_check_small
                                        row.cells("ItemsExported").value = sTotalItemsExported
                                        row.cells("ItemsFailed").value = sTotalItemsFailed
                                        row.cells("ItemsSkipped").value = sTotalItemsSkipped
                                        Invoke(Sub()
                                                   row.cells("SummaryReportLocation").value = sNuixCaseDir & "\summary-report.txt"
                                               End Sub
                                            )
                                        Invoke(Sub()
                                                   row.DefaultCellStyle.Forecolor = Color.Green
                                               End Sub
                                            )
                                        bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ItemsProcessed", sItemsProcessed)
                                        bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ItemsExported", sTotalItemsExported)
                                        bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ItemsSkipped", sTotalItemsSkipped)
                                        bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ItemsFailed", sTotalItemsFailed)
                                        bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ExtractionStatus", "Completed")
                                        bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "SummaryReportLocation", sNuixCaseDir & "\summary-report.txt")
                                        bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ProcessEndTime", sProcessEndTime)
                                        Invoke(Sub()
                                                   row.readonly = True
                                               End Sub)
                                    Else
                                        sProcessEndTime = DateTime.Now.ToString
                                        row.cells("ProcessingEndDate").value = sProcessEndTime
                                        row.cells("ExtractionStatus").value = "Cancelled by User"
                                        row.cells("StatusImage").value = My.Resources.yellow_info_small
                                        row.DefaultCellStyle.Forecolor = Color.Orange
                                        Invoke(Sub()
                                                   row.readonly = False
                                               End Sub)
                                        bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ItemsProcessed", sItemsProcessed)
                                        bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ExtractionStatus", "Cancelled by User")
                                        bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "SummaryReportLocation", "")
                                        bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ProcessEndTime", sProcessEndTime)
                                    End If

                                End If
                            Else
                                sProcessEndTime = DateTime.Now.ToString
                                row.cells("ProcessingEndDate").value = sProcessEndTime
                                row.cells("ExtractionStatus").value = "Failed"
                                row.DefaultCellStyle.Forecolor = Color.Red
                                row.cells("StatusImage").value = My.Resources.red_stop_small
                                Invoke(Sub()
                                           row.readonly = False
                                       End Sub)
                                bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ItemsProcessed", sItemsProcessed)
                                bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ExtractionStatus", "Failed")
                                bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "SummaryReportLocation", "")
                                bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ProcessEndTime", sProcessEndTime)

                            End If
                        Else
                            sProcessEndTime = DateTime.Now.ToString
                            row.cells("ProcessingEndDate").value = sProcessEndTime
                            row.cells("ExtractionStatus").value = "SQL Connection Lost"
                            row.cells("StatusImage").value = My.Resources.red_stop_small
                            row.DefaultCellStyle.Forecolor = Color.Red
                            Invoke(Sub()
                                       row.readonly = False
                                   End Sub)
                            bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ItemsProcessed", "0")
                            bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ExtractionStatus", "SQL Connection Lost")
                            bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "SummaryReportLocation", "")
                            bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ProcessEndTime", sProcessEndTime)
                        End If
                    Else
                        common.Logger(psIngestionLogFile, "Starting " & sBatchFileName & " for Nuix Processing")

                        NuixConsoleProcessStartInfo = New ProcessStartInfo(sBatchFileName)
                        NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        'NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Minimized
                        NuixConsoleProcessStartInfo.CreateNoWindow = True
                        NuixConsoleProcess = System.Diagnostics.Process.Start(NuixConsoleProcessStartInfo)
                        Dim bProcessRunning As Boolean
                        bProcessRunning = True
                        Do While bProcessRunning = True
                            bProcessRunning = blnCheckIfProcessIsRunning(NuixConsoleProcess.Id)
                            bStatus = dbService.GetUpdatedArchiveExtractInfo(psSQLiteLocation, sBatchName, "ItemsProcessed", sItemsProcessed, "INT")
                            If sItemsProcessed = vbNullString Then
                                row.cells("ItemsProcessed").value = 0

                            Else
                                row.cells("ItemsProcessed").value = sItemsProcessed

                            End If
                            bStatus = dbService.GetUpdatedArchiveExtractInfo(psSQLiteLocation, sBatchName, "PercentCompleted", sPercentCompleted, "INT")
                            If sPercentCompleted = vbNullString Then
                                row.cells("PercentCompleted").value = 0

                            Else
                                row.cells("PercentCompleted").value = sPercentCompleted
                            End If
                            bStatus = dbService.GetUpdatedArchiveExtractInfo(psSQLiteLocation, sBatchName, "BytesProcessed", sBytesProcessed, "INT")
                            If sBytesProcessed = vbNullString Then
                                row.cells("BytesProcessed").value = 0

                            Else
                                row.cells("BytesProcessed").value = sBytesProcessed
                            End If
                            Thread.Sleep(5000)
                        Loop

                        common.Logger(psIngestionLogFile, "Checking " & sLogDirectory & " nuix.log for errors.")
                        bStatus = blnCheckNuixLogForErrors(sLogDirectory, sTotalItemsProcessed, sTotalItemsExported, sTotalItemsSkipped, sTotalItemsFailed)
                        common.Logger(psIngestionLogFile, "Return from the Check Nuix Log For errors was " & bStatus)
                        common.Logger(psIngestionLogFile, "Return from the Check Nuix Log For errors Total Items Processed was " & sTotalItemsProcessed)
                        common.Logger(psIngestionLogFile, "Return from the Check Nuix Log For errors Total Items Exported was " & sTotalItemsExported)
                        common.Logger(psIngestionLogFile, "Return from the Check Nuix Log For errors Total Items Skipped was " & sTotalItemsSkipped)
                        common.Logger(psIngestionLogFile, "Return from the Check Nuix Log For errors Total Items Skipped was " & sTotalItemsFailed)

                        'bStatus = blnCheckNuixLogForErrors(row.cells("LogDirectory").value, sTotalItemsProcessed, sTotalItemsExported, sTotalItemsSkipped, sTotalItemsFailed)
                        If bStatus = False Then
                            If sBatchName <> vbNullString Then
                                bStatus = blnProcessSummaryReportFiles(sNuixCaseDir & "\summary-report.txt", sTotalItemsProcessed, sTotalItemsExported, sTotalItemsSkipped, sTotalItemsFailed)
                                sItemsProcessed = vbNullString
                                bStatus = dbService.GetUpdatedArchiveExtractInfo(psSQLiteLocation, sBatchName, "ItemsProcessed", sItemsProcessed, "INT")
                                bStatus = dbService.GetUpdatedArchiveExtractInfo(psSQLiteLocation, sBatchName, "ProcessEndTime", sProcessEndTime, "TEXT")
                                If CDbl(sItemsProcessed) > 0 Then
                                    row.cells("ItemsProcessed").value = sItemsProcessed
                                    sProcessEndTime = DateTime.Now.ToString
                                    row.cells("ProcessingEndDate").value = sProcessEndTime
                                    row.cells("ExtractionStatus").value = "Completed"
                                    row.cells("StatusImage").value = My.Resources.Green_check_small
                                    row.cells("ItemsExported").value = sTotalItemsExported
                                    row.cells("ItemsFailed").value = sTotalItemsFailed
                                    row.cells("ItemsSkipped").value = sTotalItemsSkipped
                                    Invoke(Sub()
                                               row.cells("SummaryReportLocation").value = sNuixCaseDir & "\summary-report.txt"
                                           End Sub
                                        )
                                    Invoke(Sub()
                                               row.DefaultCellStyle.Forecolor = Color.Green
                                           End Sub
                                        )
                                    bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ItemsProcessed", sItemsProcessed)
                                    bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ItemsExported", sTotalItemsExported)
                                    bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ItemsSkipped", sTotalItemsSkipped)
                                    bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ItemsFailed", sTotalItemsFailed)
                                    bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ExtractionStatus", "Completed")
                                    bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "SummaryReportLocation", sNuixCaseDir & "\summary-report.txt")
                                    bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ProcessEndTime", sProcessEndTime)
                                    Invoke(Sub()
                                               row.readonly = True
                                           End Sub)
                                Else
                                    sProcessEndTime = DateTime.Now.ToString
                                    row.cells("ProcessingEndDate").value = sProcessEndTime
                                    row.cells("ExtractionStatus").value = "Cancelled by User"
                                    row.cells("StatusImage").value = My.Resources.yellow_info_small
                                    row.DefaultCellStyle.Forecolor = Color.Orange
                                    Invoke(Sub()
                                               row.readonly = False
                                           End Sub)
                                    bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ItemsProcessed", sItemsProcessed)
                                    bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ExtractionStatus", "Cancelled by User")
                                    bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "SummaryReportLocation", "")
                                    bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ProcessEndTime", sProcessEndTime)

                                End If

                            End If
                        Else
                            sProcessEndTime = DateTime.Now.ToString
                            row.cells("ProcessingEndDate").value = sProcessEndTime
                            row.cells("ExtractionStatus").value = "Failed"
                            row.DefaultCellStyle.Forecolor = Color.Red
                            row.cells("StatusImage").value = My.Resources.red_stop_small
                            Invoke(Sub()
                                       row.readonly = False
                                   End Sub)
                            bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ItemsProcessed", sItemsProcessed)
                            bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ExtractionStatus", "Failed")
                            bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "SummaryReportLocation", "")
                            bStatus = dbService.UpdateArchiveExtractCustodianInfo(psSQLiteLocation, sBatchName, "ProcessEndTime", sProcessEndTime)

                        End If

                    End If
                Else
                    sBatchName = vbNullString
                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Exception in Build Archive Extraction Files", MessageBoxButtons.OK)
            common.Logger(psIngestionLogFile, "Error in blnBuildArchiveExtractionFile Function " & ex.ToString)
        End Try

        MessageBox.Show("Finished Processing Selected Archive Migration Batches")
        common.Logger(psIngestionLogFile, "Finished Processing Selected Archive Migration Batches")

        btnLoadPreviousBatches.Enabled = True

        blnBuildArchiveExtractionFiles = True
    End Function

    Private Sub NodesSetChecked(treeNode As TreeNode, checkedState As Boolean)

        For Each childNode As TreeNode In treeNode.Nodes
            childNode.Checked = checkedState
            NodesSetChecked(childNode, checkedState)
        Next

    End Sub

    Private Sub cboSecurityType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSecurityType.SelectedIndexChanged
        If cboSecurityType.Text = "Windows Authentication" Then
            txtDomain.Text = Environment.UserDomainName
            txtSQLUserName.Text = Environment.UserName
            txtDomain.Enabled = False
            txtSQLUserName.Enabled = False
            txtSQLInfo.Enabled = False
            txtSQLDBName.Enabled = True
            txtSQLHostName.Enabled = True
            txtSQLPortNumber.Enabled = True
            txtSQLInfo.Text = ""


        ElseIf cboSecurityType.Text = "SQLServer Authentication" Then
            txtDomain.Text = vbNullString
            txtDomain.Enabled = False
            txtSQLDBName.Enabled = True
            txtSQLHostName.Enabled = True
            txtSQLPortNumber.Enabled = True
            txtSQLUserName.Text = ""
            txtSQLUserName.Enabled = True
            txtSQLInfo.Enabled = True
            txtSQLInfo.Text = ""

        Else
            txtDomain.Text = vbNullString
            txtSQLUserName.Text = ""
        End If
    End Sub
End Class