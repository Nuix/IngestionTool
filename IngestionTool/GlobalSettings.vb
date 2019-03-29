Option Explicit On

Imports System
Imports System.IO
Imports System.IO.StreamReader
Imports System.Collections
Imports System.Threading
Imports System.Data.SQLite

Public Class O365ExtractionSettings
    Public psO365NuixInstances As String
    Public psO365NuixAppMemory As String
    Public psNuixInstances As String
    Public psNuixAppMemory As String
    Public psMemoryPerWorker As String
    Public psNumberOfNuixWorkers As String
    Public psNuixWorkers As String
    Public psNuixMemoryPerWorker As String
    Public psNuixCaseDir As String
    Public psNuixFilesDir As String
    Public psNuixLogDir As String
    Public psJavaTempDirectory As String
    Public psWorkerTempDirectory As String
    Public psNuixExportDir As String
    Public psNuixAppDirectory As String
    Public psO365ExchangeServer As String
    Public psO365Domain As String
    Public psO365AdminUsername As String
    Public psO365AdminInfo As String
    Public pbO365ApplicationImpersonation As Boolean
    Public psO365RetryCount As String
    Public psO365RetryDelay As String
    Public psO365RetryIncrement As String
    Public psO365FilePathTrimming As String
    Public piWorkerTimeout As Integer
    Public piMaxMessageSize As Integer
    Public pbO365EnableBulkUpload As Boolean
    Public piO365BulkUploadSize As Integer
    Public piO365MaxDownloadSize As Integer
    Public piO365MaxDownloadCount As Integer
    Public piPSTExportSize As Integer
    Public pbPSTAddDistributionListMetadata As Boolean
    Public pbEMLAddDistributionListMetadata As Boolean
    Public piExtractWorkerTimeout As Integer
    Public pbEnablePrefetch As Boolean
    Public pbCollaborativeFetch As Boolean
    Public psNMSAddress As String
    Public psNMSPort As String
    Public psNMSUserName As String
    Public psNMSAdminInfo As String
    Public psSettingsFile As String
    Public psSQLiteLocation As String
    Public psRedisHost As String
    Public psRedisPort As String
    Public psRedisAuth As String
    Public psO365MemoryPerWorker As String
    Public psO365NumberOfNuixWorkers As String
    Public psNMSSourceType As String
    Public pbEnableMailboxSlackSpace As Boolean
    Public pbSettingsLoaded As Boolean

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        pbSettingsLoaded = False
        Me.Close()
    End Sub

    Private Sub btnSaveSettings_Click(sender As Object, e As EventArgs) Handles btnSaveSettings.Click

        Dim bStatus As Boolean
        Dim sSettingsLocation As String
        Dim sMachineName As String
        Dim sOutputFileName As String
        Dim sSettingsFile As String
        Dim dbService As New DatabaseService
        Dim common As New Common

        Dim sSqlDataBaseLocation As String


        sSettingsLocation = txtSaveSettingLocation.Text
        If (sSettingsLocation = vbNullString) Then
            MessageBox.Show("You must enter a location to save the setting.")
            Exit Sub
        End If
        bStatus = blnValidateInput()
        If bStatus = True Then
            sMachineName = System.Net.Dns.GetHostName()
            sOutputFileName = "SettingConfiguration-" & sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".xml"
            sSettingsFile = sSettingsLocation & "\" & sOutputFileName
            bStatus = blnBuildNuixSettingsXML(sSettingsFile)

            bStatus = common.BuildSQLiteRubyScript(Me.SQLiteDBLocation)
            bStatus = common.BuildSQLiteDatabaseScript(Me.SQLiteDBLocation)

            eMailArchiveMigrationManager.performItemIdentification = chkPerformItemIdentification.Checked.ToString.ToLower
            eMailArchiveMigrationManager.addBccToEmailDigests = chkIncludeBCC.Checked.ToString.ToLower
            eMailArchiveMigrationManager.addCommunicationDateToEmailDigests = chkIncludeItemDate.Checked.ToString.ToLower
            eMailArchiveMigrationManager.analysisLanguage = cboAnalysisLang.Text
            eMailArchiveMigrationManager.calculateAuditedSize = chkCalculateAuditedSize.Checked.ToString.ToLower
            eMailArchiveMigrationManager.calculateSSDeepFuzzyHash = chkSSDeep.Checked.ToString.ToLower
            eMailArchiveMigrationManager.carveFileSystemUnallocatedSpace = chkCarveFSunallocated.Checked.ToString.ToLower
            eMailArchiveMigrationManager.createThumbnails = chkGeneratethumbnailsforimagedata.Checked.ToString.ToLower
            eMailArchiveMigrationManager.detectFaces = chkPerformimagecolouranalysis.Checked.ToString.ToLower
            If chkMD5Digest.Checked = True Then
                eMailArchiveMigrationManager.digests = "MD5"
            ElseIf chkSHA1.Checked = True Then
                eMailArchiveMigrationManager.digests = "SHA1"

            ElseIf chkSHA256.Checked = True Then
                eMailArchiveMigrationManager.digests = "SHA256"

            ElseIf chkSSDeep.Checked = True Then
                eMailArchiveMigrationManager.digests = "SSDeep"
            End If
            eMailArchiveMigrationManager.enableExactQueries = chkEnableExactQueries.Checked.ToString.ToLower
            eMailArchiveMigrationManager.extractEndOfFileSlackSpace = chkExtractEndOfFileSpace.Checked.ToString.ToLower
            eMailArchiveMigrationManager.extractFromSlackSpace = chkExtractmailboxslackspace.Checked.ToString.ToLower
            eMailArchiveMigrationManager.extractNamedEntitiesFromProperties = chkExtractNamedEntitiesfromProperties.Checked.ToString.ToLower
            eMailArchiveMigrationManager.extractNamedEntitiesFromText = chkExtractNamedEntities.Checked.ToString.ToLower
            eMailArchiveMigrationManager.extractNamedEntitiesFromTextStripped = chkExtractNamedEntities.Checked.ToString.ToLower
            eMailArchiveMigrationManager.extractShingles = "false"
            eMailArchiveMigrationManager.hideEmbeddedImmaterialData = chkHideimmaterialItems.Checked.ToString.ToLower
            eMailArchiveMigrationManager.identifyPhysicalFiles = chkPerformItemIdentification.Checked.ToString.ToLower
            eMailArchiveMigrationManager.maxDigestSize = MaxDigestSize.Value
            eMailArchiveMigrationManager.maxStoredBinarySize = Maxbinarysize.Value
            eMailArchiveMigrationManager.processFamilyFields = chkCreatefamilySearch.Checked.ToString.ToLower
            eMailArchiveMigrationManager.processText = chkProcesstext.Checked.ToString.ToLower
            eMailArchiveMigrationManager.processTextSummaries = chkEnableTextSummarisation.Checked.ToString.ToLower
            eMailArchiveMigrationManager.recoverDeletedFiles = chkRecoverDeleteFiles.Checked.ToString.ToLower
            eMailArchiveMigrationManager.reportProcessingStatus = "false"
            eMailArchiveMigrationManager.reuseEvidenceStores = chkReuseEvidenceStores.Checked.ToString.ToLower
            eMailArchiveMigrationManager.skinToneAnalysis = chkPerformimagecolouranalysis.Checked.ToString.ToLower
            eMailArchiveMigrationManager.smartProcessRegistry = chkSmartprocessMSRegistry.Checked.ToString.ToLower
            eMailArchiveMigrationManager.stemming = chkUsestemming.Checked.ToString.ToLower
            eMailArchiveMigrationManager.stopWords = chkUseStopWords.Checked.ToString.ToLower
            eMailArchiveMigrationManager.storeBinary = chkStorebinary.Checked.ToString.ToLower
            eMailArchiveMigrationManager.traversalScope = cboTraversal.Text
            eMailArchiveMigrationManager.O365NuixInstances = cboO365NumberOfNuixInstances.Text
            eMailArchiveMigrationManager.O365NuixAppMemory = txtO365NuixAppMemory.Text
            eMailArchiveMigrationManager.O365NumberOfNuixWorkers = cboO365NumberOfNuixWorkers.Text
            eMailArchiveMigrationManager.WorkerTimeout = numWorkerTimeout.Value
            eMailArchiveMigrationManager.O365MemoryPerWorker = txtO365MemoryPerWorker.Text
            eMailArchiveMigrationManager.O365ExchangeServer = txtO365ExchangeServer.Text
            eMailArchiveMigrationManager.O365Domain = txtO365Domain.Text
            eMailArchiveMigrationManager.O365AdminUserName = txtO365AdminUserName.Text
            eMailArchiveMigrationManager.O365AdminInfo = txtO365AdminInfo.Text
            eMailArchiveMigrationManager.O365ApplicationImpersonation = chkO365ApplicationImpersonation.Checked
            eMailArchiveMigrationManager.O365RetryCount = numEWSRetryCount.Value
            eMailArchiveMigrationManager.O365RetryDelay = numEWSRetryDelay.Value
            eMailArchiveMigrationManager.O365RetryIncrement = numEWSRetryIncrement.Value
            eMailArchiveMigrationManager.O365FilePathTrimming = txtO365RemovePathPrefix.Text
            eMailArchiveMigrationManager.O365MaxUploadSize = numBulkUploadSize.Value
            eMailArchiveMigrationManager.o365MaxMessageSize = numEWSMaxUploadSize.Value
            eMailArchiveMigrationManager.O365EnableBulkUpload = chkEnableBulkUpload.Checked
            eMailArchiveMigrationManager.O365BulkUploadSize = numBulkUploadSize.Value
            eMailArchiveMigrationManager.O365MaxDownloadSize = numDownloadSize.Value
            eMailArchiveMigrationManager.O365MaxDownloadCount = numDownloadCount.Value
            eMailArchiveMigrationManager.O365EnablePrefetch = chkEnablePrefetch.Checked
            eMailArchiveMigrationManager.O365EnableCollaborativePrefetch = chkCollaborativePrefetching.Checked
            eMailArchiveMigrationManager.EnableMailboxSlackSpace = chkEnableMailboxSlackspace.Checked

            eMailArchiveMigrationManager.NMSSourceType = cboSourceType.Text
            eMailArchiveMigrationManager.NuixNMSAddress = txtNMSAddress.Text
            eMailArchiveMigrationManager.NuixNMSPort = txtNMSPort.Text
            eMailArchiveMigrationManager.NuixNMSUserName = txtNMSUsername.Text
            eMailArchiveMigrationManager.NuixNMSAdminInfo = txtNMSPassword.Text
            eMailArchiveMigrationManager.NuixInstances = cboNumberOfNuixInstances.Text
            eMailArchiveMigrationManager.NuixAppMemory = txtMemoryPerNuixInstance.Text
            eMailArchiveMigrationManager.NuixWorkers = cboNuixWorkers.Text
            eMailArchiveMigrationManager.NuixWorkerMemory = txtMemoryPerWorker.Text
            eMailArchiveMigrationManager.NuixCaseDir = txtNuixCaseDir.Text
            eMailArchiveMigrationManager.NuixFilesDir = txtNuixFilesDirectory.Text
            eMailArchiveMigrationManager.NuixLogDir = txtLogDirectory.Text
            eMailArchiveMigrationManager.NuixJavaTempDir = txtJavaTempDir.Text
            eMailArchiveMigrationManager.NuixWorkerTempDir = txtWorkerTempDir.Text
            eMailArchiveMigrationManager.NuixExportDir = txtExportDir.Text
            eMailArchiveMigrationManager.PSTExportSize = numPSTExportSize.Value
            eMailArchiveMigrationManager.PSTAddDistributionListMetadata = chkPSTAddDistributionListMetadata.Checked
            eMailArchiveMigrationManager.EMLAddDistributionListMetadata = chkEMLAddDistributionListMetadata.Checked
            eMailArchiveMigrationManager.NuixAppLocation = txtNuixAppLocation.Text
            eMailArchiveMigrationManager.SQLiteDBLocation = txtSQLiteDBLocation.Text
            eMailArchiveMigrationManager.RedisHost = txtRedisHostName.Text
            eMailArchiveMigrationManager.RedisPort = txtRedisPort.Text
            eMailArchiveMigrationManager.RedisAuth = txtRedisAuth.Text
            eMailArchiveMigrationManager.PSTConsolidation = cboPSTConsolidation.Text

            sSqlDataBaseLocation = eMailArchiveMigrationManager.SQLiteDBLocation
            Dim sSqlDataBaseName As String = "NuixEmailArchiveMigrationManager.db3"
            Dim sSqlDataBaseFullName As String = System.IO.Path.Combine(sSqlDataBaseLocation, sSqlDataBaseName)

            dbService.CreateDataBase(sSqlDataBaseFullName, sSqlDataBaseLocation)

            Directory.CreateDirectory(eMailArchiveMigrationManager.NuixCaseDir)
            Directory.CreateDirectory(eMailArchiveMigrationManager.NuixFilesDir)
            Directory.CreateDirectory(eMailArchiveMigrationManager.NuixLogDir)
            Directory.CreateDirectory(eMailArchiveMigrationManager.NuixJavaTempDir)
            Directory.CreateDirectory(eMailArchiveMigrationManager.NuixExportDir)
            Directory.CreateDirectory(eMailArchiveMigrationManager.NuixWorkerTempDir)

            pbSettingsLoaded = True

            MessageBox.Show("Successfully saved settings file: " & sSettingsFile)
        End If

    End Sub

    Private Function blnValidateInput() As Boolean
        blnValidateInput = False

        If cboNumberOfNuixInstances.Text = vbNullString Then
            MessageBox.Show("You must select the number of Nuix instances that you will be utilizing.")
            tabNuixDataSettings.BringToFront()
            cboNumberOfNuixInstances.Focus()
            blnValidateInput = False
            Exit Function
        Else

        End If

        If txtMemoryPerNuixInstance.Text = vbNullString Then
            MessageBox.Show("You must select the amount of memory per instance of Nuix that you will be utilizing.")
            tabNuixDataSettings.BringToFront()
            txtMemoryPerNuixInstance.Focus()
            blnValidateInput = False
            Exit Function
        End If

        If cboNuixWorkers.Text = vbNullString Then
            MessageBox.Show("You must select the number of Nuix WORKERS that you will be utilizing.")
            tabNuixDataSettings.BringToFront()
            cboNuixWorkers.Focus()
            blnValidateInput = False
            Exit Function
        End If

        If txtMemoryPerWorker.Text = vbNullString Then
            MessageBox.Show("You must select the amount of memory per Nuix WORKER that you will be utilizing.")
            tabNuixDataSettings.BringToFront()
            cboNuixWorkers.Focus()

            blnValidateInput = False
            Exit Function
        End If
        If txtNuixCaseDir.Text = vbNullString Then
            MessageBox.Show("You must select a Nuix Case directory that you will be utilizing.")
            tabNuixDataSettings.BringToFront()
            txtNuixCaseDir.Focus()

            blnValidateInput = False
            Exit Function
        End If

        If txtWorkerTempDir.Text = vbNullString Then
            MessageBox.Show("You must select a Worker Temp directory that you will be utilizing.")
            tabNuixDataSettings.BringToFront()
            txtWorkerTempDir.Focus()

            blnValidateInput = False
            Exit Function
        End If
        If txtLogDirectory.Text = vbNullString Then
            MessageBox.Show("You must select a Nuix LOG directory that you will be utilizing.")
            tabNuixDataSettings.BringToFront()
            txtLogDirectory.Focus()

            blnValidateInput = False
            Exit Function
        End If

        If txtJavaTempDir.Text = vbNullString Then
            MessageBox.Show("You must select a Nuix Case directory that you will be utilizing.")
            tabNuixDataSettings.BringToFront()
            txtJavaTempDir.Focus()

            blnValidateInput = False
            Exit Function
        End If

        If txtNuixAppLocation.Text = vbNullString Then
            MessageBox.Show("You must select the Location of the Nuix Application.")
            tabNuixDataSettings.BringToFront()
            txtNuixAppLocation.Focus()

            blnValidateInput = False
            Exit Function
        End If

        If txtO365ExchangeServer.Text = vbNullString Then
            MessageBox.Show("You must enter the Exchange Server Information.")
            tabEWSSettings.BringToFront()
            txtO365ExchangeServer.Focus()

            blnValidateInput = False
            Exit Function
        End If

        If txtO365AdminUserName.Text = vbNullString Then
            MessageBox.Show("You must enter the Admin Username.")
            tabEWSSettings.BringToFront()
            txtO365AdminUserName.Focus()

            blnValidateInput = False
            Exit Function
        End If

        If txtO365AdminInfo.Text = vbNullString Then
            MessageBox.Show("You must enter the Admin Password.")
            tabEWSSettings.BringToFront()
            txtO365AdminInfo.Focus()

            blnValidateInput = False
            Exit Function
        End If

        If cboSourceType.Text = "Server" Then
            If txtNMSAddress.Text = vbNullString Then
                MessageBox.Show("You must enter the NMS Location.")
                tabLicensingSettings.BringToFront()
                txtNMSAddress.Focus()

                blnValidateInput = False
                Exit Function
            End If
            If txtNMSUsername.Text = vbNullString Then
                MessageBox.Show("You must enter an NMS UserName.")
                tabLicensingSettings.BringToFront()
                txtNMSUsername.Focus()

                blnValidateInput = False
                Exit Function
            End If

            If txtNMSPassword.Text = vbNullString Then
                MessageBox.Show("You must enter an NMS UserName.")
                tabLicensingSettings.BringToFront()
                txtNMSPassword.Focus()

                blnValidateInput = False
                Exit Function
            End If

        End If

        If txtSQLiteDBLocation.Text = vbNullString Then
            MessageBox.Show("You must enter the location for the SQLite DB to be created.")
            tabDB.BringToFront()
            txtSQLiteDBLocation.Focus()
            blnValidateInput = False
            Exit Function
        End If

        blnValidateInput = True

    End Function
    Private Function blnBuildNuixSettingsXML(ByVal sSettingsFile As String) As Boolean
        Dim SettingsXML As Xml.XmlDocument
        Dim SettingsRoot As Xml.XmlElement

        Dim NuixParallelSettings As Xml.XmlElement
        Dim O365Settings As Xml.XmlElement
        Dim NuixDataProcessinSettings As Xml.XmlElement
        Dim DPPerformItemIdentification As Xml.XmlElement
        Dim DPCalculateProcessingSizeUpFront As Xml.XmlElement
        Dim DPTraversal As Xml.XmlElement
        Dim DPReuseEvidenceStore As Xml.XmlElement
        Dim DPStoreBinaryOfDataItems As Xml.XmlElement
        Dim DPMaximumBinarySize As Xml.XmlElement
        Dim DPRecoverDeletedFilesFromDisk As Xml.XmlElement
        Dim DPExtractEndOfFileSlackSpace As Xml.XmlElement
        Dim DPSmartProcessMSRegistry As Xml.XmlElement
        Dim DPExtractFromMailboxSlackSpace As Xml.XmlElement
        Dim DPCarveFileSystemUnallocatedSpace As Xml.XmlElement
        Dim DPCreateFamilySearchFields As Xml.XmlElement
        Dim DPHideImmaterialItmes As Xml.XmlElement
        Dim DBAnalysisLanguage As Xml.XmlElement
        Dim DPUseStopWords As Xml.XmlElement
        Dim DPUseStemming As Xml.XmlElement
        Dim DPEnableExactQueries As Xml.XmlElement
        Dim DPProcessText As Xml.XmlElement
        Dim DPEnableNearDuplicates As Xml.XmlElement
        Dim DPEnableTextSummarisation As Xml.XmlElement
        Dim DPExtractNamedEntitiesFromText As Xml.XmlElement
        Dim DPIncludeTextStrippedItems As Xml.XmlElement
        Dim DPExractNamedEntitiesFromProperties As Xml.XmlElement
        Dim DPGenerateThumbnails As Xml.XmlElement
        Dim DPPerformImageColourAnalysis As Xml.XmlElement
        Dim DPMD5 As Xml.XmlElement
        Dim DPSHA1 As Xml.XmlElement
        Dim DPSHA256 As Xml.XmlElement
        Dim DPSSDeep As Xml.XmlElement
        Dim DPMaxDigestSize As Xml.XmlElement
        Dim DPIncludeBCC As Xml.XmlElement
        Dim DPIncludeItemDate As Xml.XmlElement

        Dim NuixInstances As Xml.XmlElement
        Dim NuixAppMemory As Xml.XmlElement
        Dim NuixWorkers As Xml.XmlElement
        Dim NuixWorkerMemory As Xml.XmlElement
        Dim NuixCaseDirectory As Xml.XmlElement
        Dim NuixFilesDirectory As Xml.XmlElement
        Dim NuixLogDirectory As Xml.XmlElement
        Dim NuixJavaTempDirectory As Xml.XmlElement
        Dim NuixWorkerTempDirectory As Xml.XmlElement
        Dim NuixExportDirectory As Xml.XmlElement
        Dim NuixAppLocation As Xml.XmlElement
        Dim PSTExportSize As Xml.XmlElement
        Dim PSTAddDistMetaData As Xml.XmlElement
        Dim EMLAddDistMetaData As Xml.XmlElement
        Dim ExtractWorkerTimeout As Xml.XmlElement
        Dim EnableSlackSpace As Xml.XmlElement

        Dim O365ExchangeServer As Xml.XmlElement
        Dim O365Domain As Xml.XmlElement
        Dim O365AdminUserName As Xml.XmlElement
        Dim O365AdminInfo As Xml.XmlElement
        Dim O365ApplicationImpersonation As Xml.XmlElement
        Dim O365RetryCount As Xml.XmlElement
        Dim O365RetryDelay As Xml.XmlElement
        Dim O365RetryIncrement As Xml.XmlElement
        Dim WorkerTimeoutElement As Xml.XmlElement
        Dim O365FilePathTrimming As Xml.XmlElement
        Dim O365NuixInstances As Xml.XmlElement
        Dim O365NuixAppMemory As Xml.XmlElement
        Dim O365NuixWorkers As Xml.XmlElement
        Dim O365NuixMemoryPerWorker As Xml.XmlElement
        Dim O365MaxSize As Xml.XmlElement
        Dim O365BulkUploadEnabled As Xml.XmlElement
        Dim O365BulkSize As Xml.XmlElement
        Dim O365MaxDownSize As Xml.XmlElement
        Dim O365MaxDownCount As Xml.XmlElement
        Dim O365Prefetch As Xml.XmlElement
        Dim O365CollaborativePrefetch As Xml.XmlElement

        Dim NuixNMSSourceType As Xml.XmlElement
        Dim NuixNMSAddress As Xml.XmlElement
        Dim NuixNMSPort As Xml.XmlElement
        Dim NuixNMSUserName As Xml.XmlElement
        Dim NuixNMSInfo As Xml.XmlElement

        Dim DBSettings As Xml.XmlElement
        Dim SQLiteDBLocation As Xml.XmlElement
        Dim RedisHost As Xml.XmlElement
        Dim RedisPort As Xml.XmlElement
        Dim RedisAuth As Xml.XmlElement

        Dim PSTConsolication As Xml.XmlElement

        Dim NuixNMSSettings As Xml.XmlElement
        Dim xmlDeclaration As Xml.XmlDeclaration

        blnBuildNuixSettingsXML = False

        SettingsXML = New Xml.XmlDocument
        xmlDeclaration = SettingsXML.CreateXmlDeclaration("1.0", "UTF-8", "yes")
        SettingsRoot = SettingsXML.CreateElement("AllNuixSettings")
        SettingsXML.AppendChild(SettingsRoot)

        SettingsXML.InsertBefore(xmlDeclaration, SettingsRoot)

        NuixParallelSettings = SettingsXML.CreateElement("NuixParallelSettings")

        SettingsRoot.AppendChild(NuixParallelSettings)
        NuixInstances = SettingsXML.CreateElement("NuixInstances")
        NuixParallelSettings.AppendChild(NuixInstances)
        NuixInstances.InnerText = cboNumberOfNuixInstances.Text

        NuixAppMemory = SettingsXML.CreateElement("NuixAppMemory")
        NuixParallelSettings.AppendChild(NuixAppMemory)
        NuixAppMemory.InnerText = txtMemoryPerNuixInstance.Text

        NuixWorkers = SettingsXML.CreateElement("NuixWorkers")
        NuixParallelSettings.AppendChild(NuixWorkers)
        NuixWorkers.InnerText = cboNuixWorkers.Text

        NuixWorkerMemory = SettingsXML.CreateElement("WorkerMemory")
        NuixParallelSettings.AppendChild(NuixWorkerMemory)
        NuixWorkerMemory.InnerText = txtMemoryPerWorker.Text

        PSTExportSize = SettingsXML.CreateElement("PSTExportSize")
        NuixParallelSettings.AppendChild(PSTExportSize)
        PSTExportSize.InnerText = numPSTExportSize.Value

        PSTAddDistMetaData = SettingsXML.CreateElement("PSTAddDistributionListMetadata")
        NuixParallelSettings.AppendChild(PSTAddDistMetaData)
        PSTAddDistMetaData.InnerText = chkPSTAddDistributionListMetadata.Checked

        EMLAddDistMetaData = SettingsXML.CreateElement("EMLAddDistributionListMetadata")
        NuixParallelSettings.AppendChild(EMLAddDistMetaData)
        EMLAddDistMetaData.InnerText = chkEMLAddDistributionListMetadata.Checked

        ExtractWorkerTimeout = SettingsXML.CreateElement("ExtractWorkerTimeout")
        NuixParallelSettings.AppendChild(ExtractWorkerTimeout)
        ExtractWorkerTimeout.InnerText = numExtractWorkerTimeout.Value

        NuixCaseDirectory = SettingsXML.CreateElement("CaseDirectory")
        NuixParallelSettings.AppendChild(NuixCaseDirectory)
        NuixCaseDirectory.InnerText = txtNuixCaseDir.Text

        NuixFilesDirectory = SettingsXML.CreateElement("NuixProcessingFilesDirectory")
        NuixParallelSettings.AppendChild(NuixFilesDirectory)
        NuixFilesDirectory.InnerText = txtNuixFilesDirectory.Text

        NuixLogDirectory = SettingsXML.CreateElement("LogDirectory")
        NuixParallelSettings.AppendChild(NuixLogDirectory)
        NuixLogDirectory.InnerText = txtLogDirectory.Text

        NuixJavaTempDirectory = SettingsXML.CreateElement("JavaTempDirectory")
        NuixParallelSettings.AppendChild(NuixJavaTempDirectory)
        NuixJavaTempDirectory.InnerText = txtJavaTempDir.Text

        NuixExportDirectory = SettingsXML.CreateElement("ExportDirectory")
        NuixParallelSettings.AppendChild(NuixExportDirectory)
        NuixExportDirectory.InnerText = txtExportDir.Text

        NuixWorkerTempDirectory = SettingsXML.CreateElement("WorkerTempDirectory")
        NuixParallelSettings.AppendChild(NuixWorkerTempDirectory)
        NuixWorkerTempDirectory.InnerText = txtWorkerTempDir.Text

        NuixAppLocation = SettingsXML.CreateElement("NuixAppLocation")
        NuixParallelSettings.AppendChild(NuixAppLocation)
        NuixAppLocation.InnerText = txtNuixAppLocation.Text

        O365Settings = SettingsXML.CreateElement("Office365Settings")

        O365NuixInstances = SettingsXML.CreateElement("O365NuixInstances")
        O365Settings.AppendChild(O365NuixInstances)
        O365NuixInstances.InnerText = cboO365NumberOfNuixInstances.Text

        O365NuixAppMemory = SettingsXML.CreateElement("O365NuixAppMemory")
        O365Settings.AppendChild(O365NuixAppMemory)
        O365NuixAppMemory.InnerText = txtO365NuixAppMemory.Text

        O365NuixWorkers = SettingsXML.CreateElement("O365NuixWorkers")
        O365Settings.AppendChild(O365NuixWorkers)
        O365NuixWorkers.InnerText = cboO365NumberOfNuixWorkers.Text

        WorkerTimeoutElement = SettingsXML.CreateElement("WorkerTimeout")
        O365Settings.AppendChild(WorkerTimeoutElement)
        WorkerTimeoutElement.InnerText = numWorkerTimeout.Value

        O365NuixMemoryPerWorker = SettingsXML.CreateElement("O365MemoryPerWorker")
        O365Settings.AppendChild(O365NuixMemoryPerWorker)
        O365NuixMemoryPerWorker.InnerText = txtO365MemoryPerWorker.Text

        O365ExchangeServer = SettingsXML.CreateElement("ExchangeServer")
        O365Settings.AppendChild(O365ExchangeServer)
        O365ExchangeServer.InnerText = txtO365ExchangeServer.Text

        O365Domain = SettingsXML.CreateElement("Domain")
        O365Settings.AppendChild(O365Domain)
        O365Domain.InnerText = txtO365Domain.Text

        O365AdminUserName = SettingsXML.CreateElement("AdminUserName")
        O365Settings.AppendChild(O365AdminUserName)
        O365AdminUserName.InnerText = txtO365AdminUserName.Text

        O365AdminInfo = SettingsXML.CreateElement("AdminInfo")
        O365Settings.AppendChild(O365AdminInfo)
        O365AdminInfo.InnerText = txtO365AdminInfo.Text

        O365ApplicationImpersonation = SettingsXML.CreateElement("ApplicationImpersonation")
        O365Settings.AppendChild(O365ApplicationImpersonation)
        O365ApplicationImpersonation.InnerText = chkO365ApplicationImpersonation.Checked

        PSTConsolication = SettingsXML.CreateElement("PSTConsolidation")
        O365Settings.AppendChild(PSTConsolication)
        PSTConsolication.InnerText = cboPSTConsolidation.Text

        O365RetryCount = SettingsXML.CreateElement("RetryCount")
        O365Settings.AppendChild(O365RetryCount)
        O365RetryCount.InnerText = numEWSRetryCount.Value

        O365RetryDelay = SettingsXML.CreateElement("RetryDelay")
        O365Settings.AppendChild(O365RetryDelay)
        O365RetryDelay.InnerText = numEWSRetryDelay.Value

        O365RetryIncrement = SettingsXML.CreateElement("RetryIncrement")
        O365Settings.AppendChild(O365RetryIncrement)
        O365RetryIncrement.InnerText = numEWSRetryIncrement.Value

        O365FilePathTrimming = SettingsXML.CreateElement("FilePathTrimming")
        O365Settings.AppendChild(O365FilePathTrimming)
        O365FilePathTrimming.InnerText = txtO365RemovePathPrefix.Text

        O365MaxSize = SettingsXML.CreateElement("EWSMaxMessageSize")
        O365Settings.AppendChild(O365MaxSize)
        O365MaxSize.InnerText = numEWSMaxUploadSize.Value

        O365BulkUploadEnabled = SettingsXML.CreateElement("EnableBulkUpload")
        O365Settings.AppendChild(O365BulkUploadEnabled)
        O365BulkUploadEnabled.InnerText = chkEnableBulkUpload.Checked.ToString

        O365BulkSize = SettingsXML.CreateElement("BulkUploadSize")
        O365Settings.AppendChild(O365BulkSize)
        O365BulkSize.InnerText = numBulkUploadSize.Value

        O365MaxDownSize = SettingsXML.CreateElement("MaxDownloadSize")
        O365Settings.AppendChild(O365MaxDownSize)
        O365MaxDownSize.InnerText = numDownloadSize.Value

        O365MaxDownCount = SettingsXML.CreateElement("MaxDownloadCount")
        O365Settings.AppendChild(O365MaxDownCount)
        O365MaxDownCount.InnerText = numDownloadCount.Value

        O365Prefetch = SettingsXML.CreateElement("Prefetch")
        O365Settings.AppendChild(O365Prefetch)
        O365Prefetch.InnerText = chkEnablePrefetch.Checked

        O365CollaborativePrefetch = SettingsXML.CreateElement("CollaborativePrefetch")
        O365Settings.AppendChild(O365CollaborativePrefetch)
        O365CollaborativePrefetch.InnerText = chkCollaborativePrefetching.Checked

        EnableSlackSpace = SettingsXML.CreateElement("EnableMailboxSlackSapce")
        O365Settings.AppendChild(EnableSlackSpace)
        EnableSlackSpace.InnerText = chkEnableMailboxSlackspace.Checked

        SettingsRoot.AppendChild(O365Settings)

        NuixDataProcessinSettings = SettingsXML.CreateElement("NuixDataProcessinSettings")
        DPPerformItemIdentification = SettingsXML.CreateElement("PerformItemIdentification")
        NuixDataProcessinSettings.AppendChild(DPPerformItemIdentification)
        DPPerformItemIdentification.InnerText = chkPerformItemIdentification.Checked.ToString.ToLower

        DPCalculateProcessingSizeUpFront = SettingsXML.CreateElement("calculateAuditedSize")
        NuixDataProcessinSettings.AppendChild(DPCalculateProcessingSizeUpFront)
        DPCalculateProcessingSizeUpFront.InnerText = chkPerformItemIdentification.Checked.ToString.ToLower

        DPTraversal = SettingsXML.CreateElement("traversalScope")
        NuixDataProcessinSettings.AppendChild(DPTraversal)
        DPTraversal.InnerText = cboTraversal.Text

        DPReuseEvidenceStore = SettingsXML.CreateElement("reuseEvidenceStores")
        NuixDataProcessinSettings.AppendChild(DPReuseEvidenceStore)
        DPReuseEvidenceStore.InnerText = chkReuseEvidenceStores.Checked.ToString.ToLower

        DPStoreBinaryOfDataItems = SettingsXML.CreateElement("storeBinary")
        NuixDataProcessinSettings.AppendChild(DPStoreBinaryOfDataItems)
        DPStoreBinaryOfDataItems.InnerText = chkStorebinary.Checked.ToString.ToLower

        DPMaximumBinarySize = SettingsXML.CreateElement("maxStoredBinarySize")
        NuixDataProcessinSettings.AppendChild(DPMaximumBinarySize)
        DPMaximumBinarySize.InnerText = Maxbinarysize.Value

        DPRecoverDeletedFilesFromDisk = SettingsXML.CreateElement("recoverDeletedFiles")
        NuixDataProcessinSettings.AppendChild(DPRecoverDeletedFilesFromDisk)
        DPRecoverDeletedFilesFromDisk.InnerText = chkRecoverDeleteFiles.Checked.ToString.ToLower

        DPExtractEndOfFileSlackSpace = SettingsXML.CreateElement("extractEndOfFileSlackSpace")
        NuixDataProcessinSettings.AppendChild(DPExtractEndOfFileSlackSpace)
        DPExtractEndOfFileSlackSpace.InnerText = chkExtractEndOfFileSpace.Checked.ToString.ToLower

        DPSmartProcessMSRegistry = SettingsXML.CreateElement("smartProcessRegistry")
        NuixDataProcessinSettings.AppendChild(DPSmartProcessMSRegistry)
        DPSmartProcessMSRegistry.InnerText = chkSmartprocessMSRegistry.Checked.ToString.ToLower

        DPExtractFromMailboxSlackSpace = SettingsXML.CreateElement("extractFromSlackSpace")
        NuixDataProcessinSettings.AppendChild(DPExtractFromMailboxSlackSpace)
        DPExtractFromMailboxSlackSpace.InnerText = chkExtractmailboxslackspace.Checked.ToString.ToLower

        DPCarveFileSystemUnallocatedSpace = SettingsXML.CreateElement("carveFileSystemUnallocatedSpace")
        NuixDataProcessinSettings.AppendChild(DPCarveFileSystemUnallocatedSpace)
        DPCarveFileSystemUnallocatedSpace.InnerText = chkCarveFSunallocated.Checked.ToString.ToLower

        DPCreateFamilySearchFields = SettingsXML.CreateElement("processFamilyFields")
        NuixDataProcessinSettings.AppendChild(DPCreateFamilySearchFields)
        DPCreateFamilySearchFields.InnerText = chkCreatefamilySearch.Checked.ToString.ToLower

        DPHideImmaterialItmes = SettingsXML.CreateElement("hideEmbeddedImmaterialData")
        NuixDataProcessinSettings.AppendChild(DPHideImmaterialItmes)
        DPHideImmaterialItmes.InnerText = chkHideimmaterialItems.Checked.ToString.ToLower

        DBAnalysisLanguage = SettingsXML.CreateElement("analysisLanguage")
        NuixDataProcessinSettings.AppendChild(DBAnalysisLanguage)
        DBAnalysisLanguage.InnerText = cboAnalysisLang.Text

        DPUseStopWords = SettingsXML.CreateElement("stopWords")
        NuixDataProcessinSettings.AppendChild(DPUseStopWords)
        DPUseStopWords.InnerText = chkUseStopWords.Checked.ToString.ToLower

        DPUseStemming = SettingsXML.CreateElement("stemming")
        NuixDataProcessinSettings.AppendChild(DPUseStemming)
        DPUseStemming.InnerText = chkUsestemming.Checked.ToString.ToLower

        DPEnableExactQueries = SettingsXML.CreateElement("enableExactQueries")
        NuixDataProcessinSettings.AppendChild(DPEnableExactQueries)
        DPEnableExactQueries.InnerText = chkEnableExactQueries.Checked.ToString.ToLower

        DPProcessText = SettingsXML.CreateElement("processText")
        NuixDataProcessinSettings.AppendChild(DPProcessText)
        DPProcessText.InnerText = chkProcesstext.Checked.ToString.ToLower

        DPEnableNearDuplicates = SettingsXML.CreateElement("processFamilyFields")
        NuixDataProcessinSettings.AppendChild(DPEnableNearDuplicates)
        DPEnableNearDuplicates.InnerText = chkEnableNearDuplicates.Checked.ToString.ToLower

        DPEnableTextSummarisation = SettingsXML.CreateElement("processTextSummaries")
        NuixDataProcessinSettings.AppendChild(DPEnableTextSummarisation)
        DPEnableTextSummarisation.InnerText = chkEnableTextSummarisation.Checked.ToString.ToLower

        DPExtractNamedEntitiesFromText = SettingsXML.CreateElement("extractNamedEntitiesFromText")
        NuixDataProcessinSettings.AppendChild(DPExtractNamedEntitiesFromText)
        DPExtractNamedEntitiesFromText.InnerText = chkExtractNamedEntities.Checked.ToString.ToLower

        DPIncludeTextStrippedItems = SettingsXML.CreateElement("extractNamedEntitiesFromTextStripped")
        NuixDataProcessinSettings.AppendChild(DPIncludeTextStrippedItems)
        DPIncludeTextStrippedItems.InnerText = chkIncludeTextStrippedItems.Checked.ToString.ToLower

        DPExractNamedEntitiesFromProperties = SettingsXML.CreateElement("extractNamedEntitiesFromProperties")
        NuixDataProcessinSettings.AppendChild(DPExractNamedEntitiesFromProperties)
        DPExractNamedEntitiesFromProperties.InnerText = chkExtractNamedEntitiesfromProperties.Checked.ToString.ToLower

        DPGenerateThumbnails = SettingsXML.CreateElement("createThumbnails")
        NuixDataProcessinSettings.AppendChild(DPGenerateThumbnails)
        DPGenerateThumbnails.InnerText = chkGeneratethumbnailsforimagedata.Checked.ToString.ToLower

        DPPerformImageColourAnalysis = SettingsXML.CreateElement("skinToneAnalysis")
        NuixDataProcessinSettings.AppendChild(DPPerformImageColourAnalysis)
        DPPerformImageColourAnalysis.InnerText = chkPerformimagecolouranalysis.Checked.ToString.ToLower

        DPMD5 = SettingsXML.CreateElement("digestsMD5")
        NuixDataProcessinSettings.AppendChild(DPMD5)
        DPMD5.InnerText = chkMD5Digest.Checked.ToString.ToLower

        DPSHA1 = SettingsXML.CreateElement("digestsSHA1")
        NuixDataProcessinSettings.AppendChild(DPSHA1)
        DPSHA1.InnerText = chkSHA1.Checked.ToString.ToLower

        DPSHA256 = SettingsXML.CreateElement("digestsSHA256")
        NuixDataProcessinSettings.AppendChild(DPSHA256)
        DPSHA256.InnerText = chkSHA256.Checked.ToString.ToLower

        DPSSDeep = SettingsXML.CreateElement("calculateSSDeepFuzzyHash")
        NuixDataProcessinSettings.AppendChild(DPSSDeep)
        DPSSDeep.InnerText = chkSSDeep.Checked.ToString.ToLower

        DPMaxDigestSize = SettingsXML.CreateElement("maxDigestSize")
        NuixDataProcessinSettings.AppendChild(DPMaxDigestSize)
        DPMaxDigestSize.InnerText = DPMaxDigestSize.Value

        DPIncludeBCC = SettingsXML.CreateElement("addBccToEmailDigests")
        NuixDataProcessinSettings.AppendChild(DPIncludeBCC)
        DPIncludeBCC.InnerText = chkIncludeBCC.Checked.ToString.ToLower

        DPIncludeItemDate = SettingsXML.CreateElement("addCommunicationDateToEmailDigests")
        NuixDataProcessinSettings.AppendChild(DPIncludeItemDate)
        DPIncludeItemDate.InnerText = chkIncludeItemDate.Checked.ToString.ToLower

        SettingsRoot.AppendChild(NuixDataProcessinSettings)

        NuixNMSSettings = SettingsXML.CreateElement("NMSSetting")

        NuixNMSSourceType = SettingsXML.CreateElement("NMSSourceType")
        NuixNMSSettings.AppendChild(NuixNMSSourceType)
        NuixNMSSourceType.InnerText = cboSourceType.Text
        NuixNMSAddress = SettingsXML.CreateElement("NMSAddress")
        NuixNMSSettings.AppendChild(NuixNMSAddress)
        NuixNMSAddress.InnerText = txtNMSAddress.Text
        NuixNMSPort = SettingsXML.CreateElement("NMSPort")
        NuixNMSSettings.AppendChild(NuixNMSPort)
        NuixNMSPort.InnerText = txtNMSPort.Text

        NuixNMSUserName = SettingsXML.CreateElement("NMSUserName")
        NuixNMSSettings.AppendChild(NuixNMSUserName)
        NuixNMSUserName.InnerText = txtNMSUsername.Text
        NuixNMSInfo = SettingsXML.CreateElement("NMSPassword")
        NuixNMSSettings.AppendChild(NuixNMSInfo)
        NuixNMSInfo.InnerText = txtNMSPassword.Text

        SettingsRoot.AppendChild(NuixNMSSettings)
        DBSettings = SettingsXML.CreateElement("DBSettings")
        SettingsRoot.AppendChild(DBSettings)

        SQLiteDBLocation = SettingsXML.CreateElement("SQLiteDBLocation")
        DBSettings.AppendChild(SQLiteDBLocation)
        SQLiteDBLocation.InnerText = txtSQLiteDBLocation.Text

        RedisHost = SettingsXML.CreateElement("RedisHost")
        DBSettings.AppendChild(RedisHost)
        RedisHost.InnerText = txtRedisHostName.Text

        RedisPort = SettingsXML.CreateElement("RedisPort")
        DBSettings.AppendChild(RedisPort)
        RedisPort.InnerText = txtRedisPort.Text

        RedisAuth = SettingsXML.CreateElement("RedisAuth")
        DBSettings.AppendChild(RedisAuth)
        RedisAuth.InnerText = txtRedisAuth.Text

        SettingsXML.Save(sSettingsFile)

        blnBuildNuixSettingsXML = True

    End Function
    Public Property NuixSettingsLoaded As Boolean
        Get
            Return pbSettingsLoaded
        End Get
        Set(value As Boolean)
            pbSettingsLoaded = True
        End Set
    End Property

    Public Property O365ExchangeServer As String
        Get
            Return txtO365ExchangeServer.Text
        End Get
        Set(value As String)
            psO365ExchangeServer = value

        End Set
    End Property

    Public Property O365Domain As String
        Get
            Return txtO365Domain.Text
        End Get
        Set(value As String)
            psO365Domain = value
        End Set
    End Property

    Public Property O365AdminUserName As String
        Get
            Return txtO365AdminUserName.Text
        End Get
        Set(value As String)
            psO365AdminUserName = value
        End Set
    End Property

    Public Property O365AdminInfo As String
        Get
            Return txtO365AdminInfo.Text

        End Get
        Set(value As String)
            psO365AdminInfo = value
        End Set
    End Property

    Public Property O365ApplicationImpersonation As Boolean
        Get
            Return chkO365ApplicationImpersonation.Checked

        End Get
        Set(value As Boolean)
            pbO365ApplicationImpersonation = value
        End Set
    End Property

    Public Property O365RetryCount As String
        Get
            Return numEWSRetryCount.Value
        End Get
        Set(value As String)
            psO365RetryCount = value
        End Set
    End Property

    Public Property O365MaxMessageSize As Integer
        Get
            Return numEWSMaxUploadSize.Value
        End Get
        Set(value As Integer)
            piMaxMessageSize = value
        End Set
    End Property
    Public Property O365EnableBulkUpload As Boolean
        Get
            Return chkEnableBulkUpload.Checked
        End Get
        Set(value As Boolean)
            pbO365EnableBulkUpload = value
        End Set
    End Property

    Public Property O365BulkUploadSize As Integer
        Get
            Return numBulkUploadSize.Value
        End Get
        Set(value As Integer)
            piO365BulkUploadSize = value
        End Set
    End Property

    Public Property O365MaxDownloadSize As Integer
        Get
            Return numDownloadSize.Value
        End Get
        Set(value As Integer)
            piO365MaxDownloadSize = value
        End Set
    End Property

    Public Property O365MaxDownloadCount As Integer
        Get
            Return numDownloadCount.Value
        End Get
        Set(value As Integer)
            piO365MaxDownloadCount = value
        End Set
    End Property

    Public Property WorkerTimeout As Integer
        Get
            Return numWorkerTimeout.Value
        End Get
        Set(value As Integer)
            piWorkerTimeout = value
        End Set
    End Property

    Public Property O365Prefetch As Boolean
        Get
            Return chkEnablePrefetch.Checked
        End Get
        Set(value As Boolean)
            pbEnablePrefetch = value
        End Set
    End Property

    Public Property O365CollaborativePrefetch As Boolean
        Get
            Return chkCollaborativePrefetching.Checked
        End Get
        Set(value As Boolean)
            pbCollaborativeFetch = value
        End Set
    End Property

    Public Property NMSSourceType As String
        Get
            Return cboSourceType.Text
        End Get
        Set(value As String)
            psNMSSourceType = value
        End Set
    End Property

    Public Property O365RetryDelay As String
        Get
            Return numEWSRetryDelay.Text
        End Get
        Set(value As String)
            psO365RetryDelay = value
        End Set
    End Property
    Public Property O365RetryIncrement As String
        Get
            Return numEWSRetryIncrement.Text
        End Get
        Set(value As String)
            psO365RetryIncrement = value
        End Set
    End Property
    Public Property O365FilePathTrimming As String
        Get
            Return txtO365RemovePathPrefix.Text
        End Get
        Set(value As String)
            psO365FilePathTrimming = value
        End Set
    End Property
    Public Property NuixNMSAddress As String
        Get
            Return txtNMSAddress.Text
        End Get
        Set(value As String)
            psNMSAddress = value
        End Set
    End Property
    Public Property NuixNMSPort As String
        Get
            Return txtNMSPort.Text
        End Get
        Set(value As String)
            psNMSPort = value
        End Set
    End Property
    Public Property NuixNMSUserName As String
        Get
            Return txtNMSUsername.Text

        End Get
        Set(value As String)
            psNMSUserName = value
        End Set
    End Property
    Public Property NuixNMSAdminInfo As String
        Get
            Return txtNMSPassword.Text
        End Get
        Set(value As String)
            psNMSAdminInfo = value
        End Set
    End Property

    Public Property NuixInstances As String
        Get
            Return cboNumberOfNuixInstances.Text
        End Get
        Set(ByVal value As String)
            psNuixInstances = value
        End Set
    End Property

    Public Property NuixAppMemory As String
        Get
            Return txtMemoryPerNuixInstance.Text
        End Get

        Set(ByVal value As String)
            psNuixAppMemory = value
        End Set
    End Property

    Public Property NuixWorkers() As String
        Get
            Return txtMemoryPerWorker.Text

        End Get
        Set(value As String)
            psNuixWorkers = value
        End Set
    End Property
    Public Property NuixWorkerMemory() As String
        Get
            Return txtMemoryPerWorker.Text

        End Get
        Set(value As String)
            psNuixMemoryPerWorker = value
        End Set
    End Property

    Public Property NuixCaseDir As String
        Get
            Return txtNuixCaseDir.Text

        End Get
        Set(value As String)
            psNuixCaseDir = value
        End Set
    End Property
    Public Property NuixFilesDir As String
        Get
            Return txtNuixFilesDirectory.Text

        End Get
        Set(value As String)
            psNuixFilesDir = value
        End Set
    End Property

    Public Property NuixLogDir As String
        Get
            Return txtLogDirectory.Text

        End Get
        Set(value As String)
            psNuixLogDir = value
        End Set
    End Property

    Public Property NuixJavaTempDir() As String
        Get
            Return txtJavaTempDir.Text

        End Get
        Set(value As String)
            psJavaTempDirectory = value
        End Set
    End Property

    Public Property NuixWorkerTempDir() As String
        Get
            Return txtWorkerTempDir.Text

        End Get
        Set(value As String)
            psWorkerTempDirectory = value
        End Set
    End Property

    Public Property NuixExportDir As String
        Get
            Return txtExportDir.Text

        End Get
        Set(value As String)
            psNuixExportDir = value
        End Set
    End Property

    Public Property NuixAppLocation As String
        Get
            Return txtNuixAppLocation.Text

        End Get
        Set(value As String)
            psNuixAppDirectory = value
        End Set
    End Property

    Public Property NuixSettingsFile() As String
        Get
            Return psSettingsFile
        End Get
        Set(value As String)
            psSettingsFile = value
        End Set
    End Property

    Public Property SQLiteDBLocation As String
        Get
            Return txtSQLiteDBLocation.Text
        End Get
        Set(value As String)
            psSQLiteLocation = value
        End Set
    End Property

    Public Property RedisHost As String
        Get
            Return txtRedisHostName.Text
        End Get
        Set(value As String)
            psRedisHost = value
        End Set
    End Property

    Public Property RedisPort As String
        Get
            Return txtRedisPort.Text
        End Get
        Set(value As String)
            psRedisPort = value
        End Set
    End Property

    Public Property RedisAuth As String
        Get
            Return txtRedisAuth.Text
        End Get
        Set(value As String)
            psRedisAuth = value
        End Set
    End Property
    Public Property O365NuixInstances As String
        Get
            Return cboO365NumberOfNuixInstances.Text
        End Get
        Set(value As String)
            psO365NuixInstances = value
        End Set
    End Property
    Public Property O365NuixAppMemory As String
        Get
            Return txtO365NuixAppMemory.Text
        End Get
        Set(value As String)
            psO365NuixAppMemory = value
        End Set
    End Property
    Public Property O365NumberOfNuixWorkers As String
        Get
            Return cboO365NumberOfNuixWorkers.Text
        End Get
        Set(value As String)
            psO365NumberOfNuixWorkers = value
        End Set
    End Property
    Public Property O365MemoryPerWorker As String
        Get
            Return txtO365MemoryPerWorker.Text
        End Get
        Set(value As String)
            psO365MemoryPerWorker = value
        End Set
    End Property

    Public Property PSTExportSize As Integer
        Get
            Return piPSTExportSize
        End Get
        Set(value As Integer)
            piPSTExportSize = value
        End Set
    End Property

    Public Property ExtractWorkerTimeout As Integer
        Get
            Return piExtractWorkerTimeout
        End Get
        Set(value As Integer)
            piExtractWorkerTimeout = value
        End Set
    End Property

    Public Property PSTAddDistributionListMetadata As Boolean
        Get
            Return pbPSTAddDistributionListMetadata
        End Get
        Set(value As Boolean)
            pbPSTAddDistributionListMetadata = value
        End Set
    End Property

    Public Property EMLAddDistributionListMetadata As Boolean
        Get
            Return pbEMLAddDistributionListMetadata
        End Get
        Set(value As Boolean)
            pbEMLAddDistributionListMetadata = value
        End Set
    End Property

    Public Property EnableMailboxSlackSpace As Boolean
        Get
            Return pbEnableMailboxSlackSpace
        End Get
        Set(value As Boolean)
            pbEnableMailboxSlackSpace = value
        End Set
    End Property

    Private Sub btnSettingsLocation_Click(sender As Object, e As EventArgs) Handles btnSettingsLocation.Click
        Dim fldrDialog As New FolderBrowserDialog
        Dim sSelectedPath As String

        fldrDialog.ShowNewFolderButton = True

        If (fldrDialog.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            sSelectedPath = fldrDialog.SelectedPath
            txtSaveSettingLocation.Text = sSelectedPath
        End If
    End Sub


    Private Sub O365ExtractionSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim dblTotalMemory As Double
        Dim dblAvailableMemory As Double
        Dim NuixDataSettingsTab As TabPage
        Dim common As New Common


        Try
            cboSourceType.Items.Clear()
            cboSourceType.Items.Add("Desktop")
            cboSourceType.Items.Add("Server")
            'cboSourceType.Items.Add("Cloud Server")

            NuixDataSettingsTab = Me.tabParallelProcessingSettings.TabPages(2)
            tabParallelProcessingSettings.Controls.Remove(NuixDataSettingsTab)
            txtO365RemovePathPrefix.Text = 0
            chkEnablePrefetch.Checked = True
            chkPSTAddDistributionListMetadata.Checked = True
            chkEMLAddDistributionListMetadata.Checked = False

            chkEnableBulkUpload.Checked = False
            numBulkUploadSize.Enabled = False

            dblTotalMemory = System.Math.Round(My.Computer.Info.TotalPhysicalMemory / (1024 * 1024), 2).ToString
            dblAvailableMemory = System.Math.Round(My.Computer.Info.AvailablePhysicalMemory / (1024 * 1024), 2).ToString
            lblRAM.Text = dblTotalMemory
            lblAvailableRAM.Text = dblAvailableMemory
            lblO365SystemMemory.Text = dblTotalMemory
            lblO365AvailMemory.Text = dblAvailableMemory

            chkPerformItemIdentification.Checked = CBool(eMailArchiveMigrationManager.performItemIdentification)
            chkIncludeBCC.Checked = CBool(eMailArchiveMigrationManager.addBccToEmailDigests)
            chkIncludeItemDate.Checked = CBool(eMailArchiveMigrationManager.addCommunicationDateToEmailDigests)
            cboAnalysisLang.Text = eMailArchiveMigrationManager.analysisLanguage
            chkCalculateAuditedSize.Checked = CBool(eMailArchiveMigrationManager.calculateAuditedSize)
            chkSSDeep.Checked = CBool(eMailArchiveMigrationManager.calculateSSDeepFuzzyHash)
            chkCarveFSunallocated.Checked = CBool(eMailArchiveMigrationManager.carveFileSystemUnallocatedSpace)
            chkGeneratethumbnailsforimagedata.Checked = CBool(eMailArchiveMigrationManager.createThumbnails)
            chkPerformimagecolouranalysis.Checked = CBool(eMailArchiveMigrationManager.detectFaces)
            If eMailArchiveMigrationManager.digests = "MD5" Then
                chkMD5Digest.Checked = True
            ElseIf eMailArchiveMigrationManager.digests = "SHA1" = True Then
                chkSHA1.Checked = True
            ElseIf eMailArchiveMigrationManager.digests = "SHA256" Then
                chkSHA256.Checked = True
            ElseIf eMailArchiveMigrationManager.digests = "SSDeep" Then
                chkSSDeep.Checked = True
            End If
            chkEnableExactQueries.Checked = CBool(eMailArchiveMigrationManager.enableExactQueries)
            chkExtractEndOfFileSpace.Checked = CBool(eMailArchiveMigrationManager.extractEndOfFileSlackSpace)
            chkExtractmailboxslackspace.Checked = CBool(eMailArchiveMigrationManager.extractFromSlackSpace)
            chkExtractNamedEntitiesfromProperties.Checked = CBool(eMailArchiveMigrationManager.extractNamedEntitiesFromProperties)
            chkExtractNamedEntities.Checked = CBool(eMailArchiveMigrationManager.extractNamedEntitiesFromText)
            chkExtractNamedEntities.Checked = CBool(eMailArchiveMigrationManager.extractNamedEntitiesFromTextStripped)
            'eMailArchiveMigrationManager.extractShingles = "false"
            chkHideimmaterialItems.Checked = CBool(eMailArchiveMigrationManager.hideEmbeddedImmaterialData)
            chkPerformItemIdentification.Checked = CBool(eMailArchiveMigrationManager.identifyPhysicalFiles)
            MaxDigestSize.Value = eMailArchiveMigrationManager.maxDigestSize
            Maxbinarysize.Value = eMailArchiveMigrationManager.maxStoredBinarySize
            chkCreatefamilySearch.Checked = CBool(eMailArchiveMigrationManager.processFamilyFields)
            chkProcesstext.Checked = CBool(eMailArchiveMigrationManager.processText)
            chkEnableTextSummarisation.Checked = CBool(eMailArchiveMigrationManager.processTextSummaries)
            chkRecoverDeleteFiles.Checked = CBool(eMailArchiveMigrationManager.recoverDeletedFiles)
            'eMailArchiveMigrationManager.reportProcessingStatus = "false"
            chkReuseEvidenceStores.Checked = CBool(eMailArchiveMigrationManager.reuseEvidenceStores)
            chkPerformimagecolouranalysis.Checked = CBool(eMailArchiveMigrationManager.skinToneAnalysis)
            chkSmartprocessMSRegistry.Checked = CBool(eMailArchiveMigrationManager.smartProcessRegistry)
            chkUsestemming.Checked = CBool(eMailArchiveMigrationManager.stemming)
            chkUseStopWords.Checked = CBool(eMailArchiveMigrationManager.stopWords)
            chkStorebinary.Checked = CBool(eMailArchiveMigrationManager.storeBinary)
            cboTraversal.Text = eMailArchiveMigrationManager.traversalScope
            cboO365NumberOfNuixInstances.Text = eMailArchiveMigrationManager.O365NuixInstances
            txtO365NuixAppMemory.Text = eMailArchiveMigrationManager.O365NuixAppMemory
            cboO365NumberOfNuixWorkers.Text = eMailArchiveMigrationManager.O365NumberOfNuixWorkers
            txtO365MemoryPerWorker.Text = eMailArchiveMigrationManager.O365MemoryPerWorker
            txtO365ExchangeServer.Text = eMailArchiveMigrationManager.O365ExchangeServer
            txtO365Domain.Text = eMailArchiveMigrationManager.O365Domain
            txtO365AdminUserName.Text = eMailArchiveMigrationManager.O365AdminUserName
            txtO365AdminInfo.Text = eMailArchiveMigrationManager.O365AdminInfo
            chkO365ApplicationImpersonation.Checked = CBool(eMailArchiveMigrationManager.O365ApplicationImpersonation)
            numEWSRetryCount.Value = eMailArchiveMigrationManager.O365RetryCount
            numWorkerTimeout.Value = eMailArchiveMigrationManager.WorkerTimeout
            cboSourceType.Text = eMailArchiveMigrationManager.NMSSourceType
            numEWSRetryDelay.Value = eMailArchiveMigrationManager.O365RetryDelay
            numEWSRetryIncrement.Text = eMailArchiveMigrationManager.O365RetryIncrement
            txtO365RemovePathPrefix.Text = eMailArchiveMigrationManager.O365FilePathTrimming
            chkEnableBulkUpload.Checked = CBool(eMailArchiveMigrationManager.O365EnableBulkUpload)
            numBulkUploadSize.Value = eMailArchiveMigrationManager.O365BulkUploadSize
            numDownloadSize.Value = eMailArchiveMigrationManager.O365MaxDownloadSize
            numDownloadCount.Value = eMailArchiveMigrationManager.O365MaxDownloadCount
            chkEnablePrefetch.Checked = CBool(eMailArchiveMigrationManager.O365EnablePrefetch)
            chkCollaborativePrefetching.Checked = CBool(eMailArchiveMigrationManager.O365EnableCollaborativePrefetch)
            numPSTExportSize.Value = eMailArchiveMigrationManager.PSTExportSize
            chkPSTAddDistributionListMetadata.Checked = CBool(eMailArchiveMigrationManager.PSTAddDistributionListMetadata)
            chkEMLAddDistributionListMetadata.Checked = CBool(eMailArchiveMigrationManager.EMLAddDistributionListMetadata)
            numExtractWorkerTimeout.Value = eMailArchiveMigrationManager.ExtractWorkerTimeout
            txtNMSAddress.Text = eMailArchiveMigrationManager.NuixNMSAddress
            txtNMSPort.Text = eMailArchiveMigrationManager.NuixNMSPort
            txtNMSUsername.Text = eMailArchiveMigrationManager.NuixNMSUserName
            txtNMSPassword.Text = eMailArchiveMigrationManager.NuixNMSAdminInfo
            cboNumberOfNuixInstances.Text = eMailArchiveMigrationManager.NuixInstances
            txtMemoryPerNuixInstance.Text = eMailArchiveMigrationManager.NuixAppMemory
            cboNuixWorkers.Text = eMailArchiveMigrationManager.NuixWorkers
            txtMemoryPerWorker.Text = eMailArchiveMigrationManager.NuixWorkerMemory
            txtNuixCaseDir.Text = eMailArchiveMigrationManager.NuixCaseDir
            txtNuixFilesDirectory.Text = eMailArchiveMigrationManager.NuixFilesDir
            txtLogDirectory.Text = eMailArchiveMigrationManager.NuixLogDir
            txtJavaTempDir.Text = eMailArchiveMigrationManager.NuixJavaTempDir
            txtWorkerTempDir.Text = eMailArchiveMigrationManager.NuixWorkerTempDir
            txtExportDir.Text = eMailArchiveMigrationManager.NuixExportDir
            txtNuixAppLocation.Text = eMailArchiveMigrationManager.NuixAppLocation
            txtSQLiteDBLocation.Text = eMailArchiveMigrationManager.SQLiteDBLocation
            txtRedisAuth.Text = eMailArchiveMigrationManager.RedisAuth
            txtRedisHostName.Text = eMailArchiveMigrationManager.RedisHost
            txtRedisPort.Text = eMailArchiveMigrationManager.RedisPort
            chkEnableMailboxSlackspace.Checked = eMailArchiveMigrationManager.EnableMailboxSlackSpace
            cboPSTConsolidation.Text = eMailArchiveMigrationManager.PSTConsolidation

        Catch ex As Exception
            common.Logger(psIngestionLogFile, ex.ToString)
        End Try

        Me.Height = 575

    End Sub

    Private Sub btnLoadSettingXML_Click(sender As Object, e As EventArgs) Handles btnLoadSettingXML.Click
        Dim OpenFileDialog1 As New OpenFileDialog
        Dim bStatus As Boolean

        With OpenFileDialog1
            .Filter = "xml|*.xml"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            psSettingsFile = OpenFileDialog1.FileName.ToString
            Me.NuixSettingsFile = psSettingsFile
            eMailArchiveMigrationManager.psSettingsFile = psSettingsFile
        Else
            Exit Sub
        End If

        bStatus = blnLoadNuixXMLSettings(psSettingsFile)
        pbSettingsLoaded = True
    End Sub

    Private Function blnLoadNuixXMLSettings(ByVal sSettingConfigurationFile As String) As Boolean
        Dim NuixXMLSettings As Xml.XmlDocument
        Dim oNuixParallelProcessingSettingsNode As Xml.XmlNode
        Dim oNuixDataProcessinSettingsNode As Xml.XmlNode
        Dim oO365SettingsNode As Xml.XmlNode
        Dim oNMSSettingsNode As Xml.XmlNode
        Dim oDBNode As Xml.XmlNode

        blnLoadNuixXMLSettings = False
        NuixXMLSettings = New Xml.XmlDocument

        NuixXMLSettings.Load(sSettingConfigurationFile)

        Me.NuixSettingsLoaded = True

        Me.NuixSettingsFile = sSettingConfigurationFile

        oNuixParallelProcessingSettingsNode = NuixXMLSettings.SelectSingleNode("AllNuixSettings/NuixParallelSettings")
        If oNuixParallelProcessingSettingsNode.HasChildNodes Then
            For Each Child In oNuixParallelProcessingSettingsNode.ChildNodes
                If Child.name = "NuixInstances" Then
                    cboNumberOfNuixInstances.Text = Child.innertext
                    Me.NuixInstances = Child.innertext
                ElseIf Child.name = "NuixAppMemory" Then
                    txtMemoryPerNuixInstance.Text = Child.innertext
                    Me.NuixAppMemory = Child.innertext
                ElseIf Child.name = "NuixWorkers" Then
                    cboNuixWorkers.Text = Child.innertext
                    psNuixWorkers = Child.innertext
                    Me.NuixWorkers = Child.innertext
                ElseIf Child.name = "WorkerMemory" Then
                    txtMemoryPerWorker.Text = Child.innertext
                    Me.NuixWorkerMemory = Child.innertext
                ElseIf Child.name = "CaseDirectory" Then
                    txtNuixCaseDir.Text = Child.innertext
                    Me.NuixCaseDir = Child.innertext
                ElseIf Child.name = "NuixProcessingFilesDirectory" Then
                    txtNuixFilesDirectory.Text = Child.innertext
                    Me.NuixFilesDir = Child.innertext
                ElseIf Child.name = "LogDirectory" Then
                    txtLogDirectory.Text = Child.innertext
                    Me.NuixLogDir = Child.innertext
                ElseIf Child.name = "JavaTempDirectory" Then
                    txtJavaTempDir.Text = Child.innertext
                    Me.NuixJavaTempDir = Child.innertext
                ElseIf Child.name = "ExportDirectory" Then
                    txtExportDir.Text = Child.innertext
                    Me.NuixExportDir = Child.innertext
                ElseIf Child.name = "WorkerTempDirectory" Then
                    txtWorkerTempDir.Text = Child.innertext
                    Me.NuixWorkerTempDir = Child.innertext
                ElseIf Child.name = "NuixAppLocation" Then
                    txtNuixAppLocation.Text = Child.innertext
                    Me.NuixAppLocation = Child.innertext
                ElseIf Child.name = "PSTExportSize" Then
                    numPSTExportSize.Value = Child.innertext
                    Me.PSTExportSize = Child.innertext
                ElseIf Child.name = "PSTAddDistributionListMetadata" Then
                    chkPSTAddDistributionListMetadata.Checked = CBool(Child.innertext)
                    Me.PSTAddDistributionListMetadata = Child.innertext
                ElseIf Child.name = "EMLAddDistributionListMetadata" Then
                    chkEMLAddDistributionListMetadata.Checked = CBool(Child.innertext)
                    Me.EMLAddDistributionListMetadata = Child.innertext
                ElseIf Child.name = "ExtractWorkerTimeout" Then
                    numExtractWorkerTimeout.Value = Child.innertext
                    Me.ExtractWorkerTimeout = Child.innertext
                End If
            Next
        End If

        oO365SettingsNode = NuixXMLSettings.SelectSingleNode("AllNuixSettings/Office365Settings")
        If oO365SettingsNode.HasChildNodes Then
            For Each Child In oO365SettingsNode.ChildNodes
                If Child.name = "O365NuixInstances" Then
                    cboO365NumberOfNuixInstances.Text = Child.Innertext
                    Me.O365NuixInstances = Child.Innertext
                ElseIf Child.name = "O365NuixAppMemory" Then
                    txtO365NuixAppMemory.Text = Child.Innertext
                    Me.O365NuixAppMemory = Child.Innertext
                ElseIf Child.name = "O365NuixWorkers" Then
                    cboO365NumberOfNuixWorkers.Text = Child.Innertext
                    Me.O365NumberOfNuixWorkers = Child.Innertext
                ElseIf Child.name = "O365MemoryPerWorker" Then
                    txtO365MemoryPerWorker.Text = Child.Innertext
                    Me.O365MemoryPerWorker = Child.Innertext
                ElseIf Child.name = "WorkerTimeout" Then
                    numWorkerTimeout.Value = Child.innertext
                    Me.workertimeout = Child.innertext
                ElseIf Child.name = "ExchangeServer" Then
                    txtO365ExchangeServer.Text = Child.Innertext
                    Me.O365ExchangeServer = Child.Innertext
                ElseIf Child.name = "Domain" Then
                    txtO365Domain.Text = Child.Innertext
                    Me.O365Domain = Child.Innertext
                ElseIf Child.name = "AdminUserName" Then
                    txtO365AdminUserName.Text = Child.Innertext
                    Me.O365AdminUserName = Child.Innertext
                ElseIf Child.name = "AdminInfo" Then
                    txtO365AdminInfo.Text = Child.Innertext
                    Me.O365AdminInfo = Child.Innertext
                ElseIf Child.name = "ApplicationImpersonation" Then
                    chkO365ApplicationImpersonation.Checked = Child.Innertext
                    Me.O365ApplicationImpersonation = Child.Innertext
                ElseIf Child.name = "RetryCount" Then
                    numEWSRetryCount.Value = Child.Innertext
                    Me.O365RetryCount = Child.Innertext
                ElseIf Child.name = "RetryDelay" Then
                    numEWSRetryDelay.Value = Child.Innertext
                    Me.O365RetryDelay = Child.Innertext
                ElseIf Child.name = "RetryIncrement" Then
                    numEWSRetryIncrement.Value = Child.Innertext
                    Me.O365RetryIncrement = Child.innertext
                ElseIf Child.name = "FilePathTrimming" Then
                    txtO365RemovePathPrefix.Text = Child.Innertext
                    Me.O365FilePathTrimming = Child.Innertext
                ElseIf Child.name = "EWSMaxMessageSize" Then
                    numEWSMaxUploadSize.Value = Child.innertext
                    Me.O365MaxMessageSize = Child.innertext
                ElseIf Child.name = "EnableBulkUpload" Then
                    chkEnableBulkUpload.Checked = CBool(Child.innertext)
                    Me.O365EnableBulkUpload = CBool(Child.innertext)
                    If CBool(Child.innertext) = True Then
                        chkEnableBulkUpload.Checked = True
                        numBulkUploadSize.Enabled = True
                    Else
                        chkEnableBulkUpload.Checked = False
                        numBulkUploadSize.Enabled = False
                    End If
                ElseIf Child.name = "BulkUploadSize" Then
                    numBulkUploadSize.Value = Child.innertext
                    Me.O365BulkUploadSize = Child.innertext
                ElseIf Child.name = "MaxDownloadSize" Then
                    numDownloadSize.Value = Child.innertext
                    Me.O365MaxDownloadSize = Child.innertext
                ElseIf Child.name = "MaxDownloadCount" Then
                    numDownloadCount.Value = Child.innertext
                    Me.O365MaxDownloadCount = Child.innertext
                ElseIf Child.name = "Prefetch" Then
                    chkEnablePrefetch.Checked = CBool(Child.innertext)
                    Me.O365Prefetch = CBool(Child.innertext)
                ElseIf Child.name = "CollaborativePrefetch" Then
                    chkCollaborativePrefetching.Checked = CBool(Child.innertext)
                    Me.O365CollaborativePrefetch = CBool(Child.innertext)
                ElseIf Child.name = "EnableMailboxSlackSapce" Then
                    chkEnableMailboxSlackspace.Checked = CBool(Child.innertext)
                    Me.EnableMailboxSlackSpace = CBool(Child.innertext)
                ElseIf Child.name = "PSTConsolidation" Then
                    cboPSTConsolidation.Text = Child.innertext
                End If

            Next
        End If
        oNMSSettingsNode = NuixXMLSettings.SelectSingleNode("AllNuixSettings/NMSSetting")
        If (oNMSSettingsNode.HasChildNodes) Then
            For Each child In oNMSSettingsNode.ChildNodes
                If child.name = "NMSAddress" Then
                    txtNMSAddress.Text = child.innertext
                    Me.NuixNMSAddress = child.innertext
                ElseIf child.name = "NMSPort" Then
                    txtNMSPort.Text = child.innertext
                    Me.NuixNMSPort = child.innertext
                ElseIf child.name = "NMSUserName" Then
                    txtNMSUsername.Text = child.innertext
                    Me.NuixNMSUserName = child.innertext
                ElseIf child.name = "NMSPassword" Then
                    txtNMSPassword.Text = child.innertext
                    Me.NuixNMSAdminInfo = child.innertext
                ElseIf child.name = "NMSSourceType" Then
                    cboSourceType.Text = child.innertext
                    Me.NMSSourceType = child.innertext
                End If
            Next
        End If

        oNuixDataProcessinSettingsNode = NuixXMLSettings.SelectSingleNode("AllNuixSettings/NuixDataProcessinSettings")
        If (oNuixDataProcessinSettingsNode.HasChildNodes) Then
            For Each child In oNuixDataProcessinSettingsNode.ChildNodes
                If child.name = "PerformItemIdentification" Then
                    chkPerformItemIdentification.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.PerformItemIdentification = child.innertext
                ElseIf child.name = "recoverDeletedFiles" Then
                    chkRecoverDeleteFiles.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.recoverDeletedFiles = child.innertext
                ElseIf child.name = "calculateAuditedSizeUpFront" Then

                ElseIf child.name = "reportProcessingStatus" Then
                    eMailArchiveMigrationManager.reportProcessingStatus = child.innertext

                ElseIf child.name = "traversalScope" Then
                    cboTraversal.Text = child.innertext
                    eMailArchiveMigrationManager.traversalScope = child.innertext
                ElseIf child.name = "reuseEvidenceStores" Then
                    chkReuseEvidenceStores.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.reuseEvidenceStores = child.innertext
                ElseIf child.name = "storeBinary" Then
                    chkStorebinary.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.storeBinary = chkStorebinary.Checked.ToString.ToLower
                ElseIf child.name = "maxStoredBinarySize" Then
                    Maxbinarysize.Value = child.innertext
                    eMailArchiveMigrationManager.maxStoredBinarySize = Maxbinarysize.Value
                ElseIf child.name = "recoverDeletedFiles" Then
                ElseIf child.name = "extractEndOfFileSlackSpace" Then
                    chkExtractEndOfFileSpace.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.extractEndOfFileSlackSpace = child.innertext
                ElseIf child.name = "smartProcessRegistry" Then
                    chkSmartprocessMSRegistry.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.smartProcessRegistry = child.innertext
                ElseIf child.name = "extractFromSlackSpace" Then
                    eMailArchiveMigrationManager.extractFromSlackSpace = child.innertext
                    chkExtractmailboxslackspace.Checked = CBool(child.innertext)
                ElseIf child.name = "carveFileSystemUnallocatedSpace" Then
                ElseIf child.name = "processFamilyFields" Then
                    chkCreatefamilySearch.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.processFamilyFields = child.innertext
                ElseIf child.name = "hideEmbeddedImmaterialData" Then
                    chkHideimmaterialItems.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.hideEmbeddedImmaterialData = child.innertext
                ElseIf child.name = "analysisLanguage" Then
                ElseIf child.name = "stopWords" Then
                    chkUseStopWords.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.stopWords = child.innertext
                ElseIf child.name = "stemming" Then
                    chkUsestemming.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.stemming = child.innertext
                ElseIf child.name = "enableExactQueries" Then
                    chkEnableExactQueries.Checked = child.innertext
                    eMailArchiveMigrationManager.enableExactQueries = CBool(child.innertext)
                ElseIf child.name = "processText" Then
                    chkProcesstext.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.processText = child.innertext
                ElseIf child.name = "processFamilyFields" Then

                ElseIf child.name = "processTextSummaries" Then
                    chkEnableTextSummarisation.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.processTextSummaries = child.innertext
                ElseIf child.name = "extractNamedEntitiesFromText" Then
                    chkExtractNamedEntities.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.extractNamedEntitiesFromText = child.innertext
                ElseIf child.name = "extractNamedEntitiesFromTextStripped" Then
                    chkIncludeTextStrippedItems.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.extractNamedEntitiesFromTextStripped = child.innertext
                ElseIf child.name = "extractNamedEntitiesFromProperties" Then
                    chkExtractNamedEntitiesfromProperties.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.extractNamedEntitiesFromProperties = child.innertext
                ElseIf child.name = "createThumbnails" Then
                    chkGeneratethumbnailsforimagedata.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.createThumbnails = child.innertext
                ElseIf child.name = "skinToneAnalysis" Then
                    chkPerformimagecolouranalysis.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.skinToneAnalysis = child.innertext
                ElseIf child.name = "digestsMD5" Then
                    chkMD5Digest.Checked = True
                    eMailArchiveMigrationManager.digests = "MD5"
                ElseIf child.name = "digestsSHA1" Then
                    chkSHA1.Checked = True
                    eMailArchiveMigrationManager.digests = "SHA1"
                ElseIf child.name = "digestsSHA256" Then
                    chkSHA256.Checked = True
                    eMailArchiveMigrationManager.digests = "SHA256"
                ElseIf child.name = "calculateSSDeepFuzzyHash" Then
                    chkSSDeep.Checked = True
                    eMailArchiveMigrationManager.digests = "SSDeep"
                ElseIf child.name = "maxDigestSize" Then
                    If child.innertext <> vbNullString Then
                        MaxDigestSize.Value = CInt(child.innertext)
                        eMailArchiveMigrationManager.maxDigestSize = child.innertext
                    Else
                        eMailArchiveMigrationManager.maxDigestSize = 0
                    End If
                ElseIf child.name = "addBccToEmailDigests" Then
                    chkIncludeBCC.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.addBccToEmailDigests = child.innertext
                ElseIf child.name = "addCommunicationDateToEmailDigests" Then
                    chkIncludeItemDate.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.addCommunicationDateToEmailDigests = child.innertext
                ElseIf child.name = "calculateAuditedSize" Then
                    chkCalculateAuditedSize.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.calculateAuditedSize = CBool(child.innertext)
                ElseIf child.name = "extractShingles" Then
                    eMailArchiveMigrationManager.extractShingles = CBool(child.innertext)
                ElseIf child.name = "identifyPhysicalFiles" Then
                    chkPerformItemIdentification.Checked = CBool(child.innertext)
                    eMailArchiveMigrationManager.identifyPhysicalFiles = child.innertext
                End If
            Next
        End If

        oDBNode = NuixXMLSettings.SelectSingleNode("AllNuixSettings/DBSettings")
        If Not IsNothing(oDBNode) Then
            If (oDBNode.HasChildNodes) Then
                For Each child In oDBNode.ChildNodes
                    If child.name = "SQLiteDBLocation" Then
                        txtSQLiteDBLocation.Text = child.innertext
                        Me.SQLiteDBLocation = child.innertext
                    ElseIf child.name = "RedisHost" Then
                        txtRedisHostName.Text = child.innertext
                        Me.RedisHost = child.innertext
                    ElseIf child.name = "RedisPort" Then
                        txtRedisPort.Text = child.innertext
                        Me.RedisPort = child.innertext
                    ElseIf child.name = "RedisAuth" Then
                        txtRedisAuth.Text = child.innertext
                        Me.RedisAuth = child.innertext
                    End If
                Next
            End If
        End If

        Me.NuixSettingsLoaded = True

        blnLoadNuixXMLSettings = True
    End Function

    Private Sub cboNumberOfNuixInstances_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim dblAvailableAfterNuixInstances As Double
        Dim dblMemoryPerNuixInstance As Double
        Dim dblMemoryNecessary As Double


        Try
            If txtMemoryPerNuixInstance.Text <> vbNullString Then
                dblMemoryPerNuixInstance = CInt(txtMemoryPerNuixInstance.Text)

                dblMemoryNecessary = (CInt(cboNumberOfNuixInstances.Text) * dblMemoryPerNuixInstance)

                dblAvailableAfterNuixInstances = CInt(lblAvailableRAM.Text) - dblMemoryNecessary
                If (dblAvailableAfterNuixInstances < 0) Then
                    MessageBox.Show("You do not have enough memory on this machine to run " & cboNumberOfNuixInstances.Text & " instances of Nuix with " & txtMemoryPerNuixInstance.Text & " memory per instance.  Please choose fewer instances of Nuix or less memory per instance.")
                    cboNumberOfNuixInstances.Focus()
                    Exit Sub

                End If

                lblAvailableRAMAfterNuix.Text = dblAvailableAfterNuixInstances

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub cboNuixWorkers_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            txtMemoryPerWorker.Text = Math.Round(CInt(lblAvailableRAMAfterNuix.Text) / cboNuixWorkers.Text)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtMemoryPerNuixInstance_TextChanged(sender As Object, e As EventArgs)
        Dim dblMemoryPerNuixInstance As Double
        Dim dblMemoryNecessary As Double
        Dim dblAvailableAfterNuixInstances As Double


        Try
            If txtMemoryPerNuixInstance.Text <> vbNullString Then
                If cboNumberOfNuixInstances.Text <> vbNullString Then
                    dblMemoryPerNuixInstance = CInt(txtMemoryPerNuixInstance.Text)

                    dblMemoryNecessary = (CInt(cboNumberOfNuixInstances.Text) * dblMemoryPerNuixInstance)

                    dblAvailableAfterNuixInstances = CInt(lblAvailableRAM.Text) - dblMemoryNecessary
                    If (dblAvailableAfterNuixInstances < 0) Then
                        MessageBox.Show("You do not have enough memory on this machine to run " & cboNumberOfNuixInstances.Text & " instances of Nuix with " & txtMemoryPerNuixInstance.Text & " memory per instance.  Please choose fewer instances of Nuix or less memory per instance.")
                        cboNumberOfNuixInstances.Focus()
                        Exit Sub
                    End If
                    lblAvailableRAMAfterNuix.Text = dblAvailableAfterNuixInstances
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnSettingsOK_Click(sender As Object, e As EventArgs) Handles btnSettingsOK.Click
        Dim bStatus As Boolean
        Dim common As New Common

        If eMailArchiveMigrationManager.psSettingsFile = vbNullString Then
            MessageBox.Show("No settings file has been loaded - please load appropriate settings file.", "No Settings file loaded.")
            Exit Sub
        End If
        If txtSQLiteDBLocation.Text = vbNullString Then
            MessageBox.Show("You must enter a location for the SQLite Database Locations", "SQLite Database Location not specified", MessageBoxButtons.OK)
            txtSQLiteDBLocation.Focus()
            Exit Sub
        End If

        pbSettingsLoaded = True
        eMailArchiveMigrationManager.performItemIdentification = chkPerformItemIdentification.Checked.ToString.ToLower
        eMailArchiveMigrationManager.addBccToEmailDigests = chkIncludeBCC.Checked.ToString.ToLower
        eMailArchiveMigrationManager.addCommunicationDateToEmailDigests = chkIncludeItemDate.Checked.ToString.ToLower
        eMailArchiveMigrationManager.analysisLanguage = cboAnalysisLang.Text
        eMailArchiveMigrationManager.calculateAuditedSize = chkCalculateAuditedSize.Checked.ToString.ToLower
        eMailArchiveMigrationManager.calculateSSDeepFuzzyHash = chkSSDeep.Checked.ToString.ToLower
        eMailArchiveMigrationManager.carveFileSystemUnallocatedSpace = chkCarveFSunallocated.Checked.ToString.ToLower
        eMailArchiveMigrationManager.createThumbnails = chkGeneratethumbnailsforimagedata.Checked.ToString.ToLower
        eMailArchiveMigrationManager.detectFaces = chkPerformimagecolouranalysis.Checked.ToString.ToLower
        If chkMD5Digest.Checked = True Then
            eMailArchiveMigrationManager.digests = "MD5"
        ElseIf chkSHA1.Checked = True Then
            eMailArchiveMigrationManager.digests = "SHA1"

        ElseIf chkSHA256.Checked = True Then
            eMailArchiveMigrationManager.digests = "SHA256"

        ElseIf chkSSDeep.Checked = True Then
            eMailArchiveMigrationManager.digests = "SSDeep"
        End If
        eMailArchiveMigrationManager.enableExactQueries = chkEnableExactQueries.Checked.ToString.ToLower
        eMailArchiveMigrationManager.extractEndOfFileSlackSpace = chkExtractEndOfFileSpace.Checked.ToString.ToLower
        eMailArchiveMigrationManager.extractFromSlackSpace = chkExtractmailboxslackspace.Checked.ToString.ToLower
        eMailArchiveMigrationManager.extractNamedEntitiesFromProperties = chkExtractNamedEntitiesfromProperties.Checked.ToString.ToLower
        eMailArchiveMigrationManager.extractNamedEntitiesFromText = chkExtractNamedEntities.Checked.ToString.ToLower
        eMailArchiveMigrationManager.extractNamedEntitiesFromTextStripped = chkExtractNamedEntities.Checked.ToString.ToLower
        eMailArchiveMigrationManager.extractShingles = "false"
        eMailArchiveMigrationManager.hideEmbeddedImmaterialData = chkHideimmaterialItems.Checked.ToString.ToLower
        eMailArchiveMigrationManager.identifyPhysicalFiles = chkPerformItemIdentification.Checked.ToString.ToLower
        eMailArchiveMigrationManager.maxDigestSize = MaxDigestSize.Value
        eMailArchiveMigrationManager.maxStoredBinarySize = Maxbinarysize.Value
        eMailArchiveMigrationManager.processFamilyFields = chkCreatefamilySearch.Checked.ToString.ToLower
        eMailArchiveMigrationManager.processText = chkProcesstext.Checked.ToString.ToLower
        eMailArchiveMigrationManager.processTextSummaries = chkEnableTextSummarisation.Checked.ToString.ToLower
        eMailArchiveMigrationManager.recoverDeletedFiles = chkRecoverDeleteFiles.Checked.ToString.ToLower
        eMailArchiveMigrationManager.reportProcessingStatus = "false"
        eMailArchiveMigrationManager.reuseEvidenceStores = chkReuseEvidenceStores.Checked.ToString.ToLower
        eMailArchiveMigrationManager.skinToneAnalysis = chkPerformimagecolouranalysis.Checked.ToString.ToLower
        eMailArchiveMigrationManager.smartProcessRegistry = chkSmartprocessMSRegistry.Checked.ToString.ToLower
        eMailArchiveMigrationManager.stemming = chkUsestemming.Checked.ToString.ToLower
        eMailArchiveMigrationManager.stopWords = chkUseStopWords.Checked.ToString.ToLower
        eMailArchiveMigrationManager.storeBinary = chkStorebinary.Checked.ToString.ToLower
        eMailArchiveMigrationManager.traversalScope = cboTraversal.Text
        eMailArchiveMigrationManager.O365NuixInstances = cboO365NumberOfNuixInstances.Text
        eMailArchiveMigrationManager.O365NuixAppMemory = txtO365NuixAppMemory.Text
        eMailArchiveMigrationManager.O365NumberOfNuixWorkers = cboO365NumberOfNuixWorkers.Text
        eMailArchiveMigrationManager.O365MemoryPerWorker = txtO365MemoryPerWorker.Text
        eMailArchiveMigrationManager.O365ExchangeServer = txtO365ExchangeServer.Text
        eMailArchiveMigrationManager.O365Domain = txtO365Domain.Text
        eMailArchiveMigrationManager.O365AdminUserName = txtO365AdminUserName.Text
        eMailArchiveMigrationManager.O365AdminInfo = txtO365AdminInfo.Text
        eMailArchiveMigrationManager.O365ApplicationImpersonation = chkO365ApplicationImpersonation.Checked
        eMailArchiveMigrationManager.O365RetryCount = numEWSRetryCount.Value
        eMailArchiveMigrationManager.NMSSourceType = cboSourceType.Text
        eMailArchiveMigrationManager.O365RetryDelay = numEWSRetryDelay.Value
        eMailArchiveMigrationManager.O365RetryIncrement = numEWSRetryIncrement.Value
        eMailArchiveMigrationManager.O365FilePathTrimming = txtO365RemovePathPrefix.Text
        eMailArchiveMigrationManager.O365MaxDownloadSize = numDownloadSize.Value
        eMailArchiveMigrationManager.O365MaxDownloadCount = numDownloadCount.Value
        eMailArchiveMigrationManager.o365MaxMessageSize = numEWSMaxUploadSize.Value
        eMailArchiveMigrationManager.O365EnablePrefetch = chkEnablePrefetch.Checked
        eMailArchiveMigrationManager.O365EnableCollaborativePrefetch = chkCollaborativePrefetching.Checked
        eMailArchiveMigrationManager.O365BulkUploadSize = numBulkUploadSize.Value
        eMailArchiveMigrationManager.O365EnableBulkUpload = chkEnableBulkUpload.Checked
        eMailArchiveMigrationManager.NuixNMSAddress = txtNMSAddress.Text
        eMailArchiveMigrationManager.NuixNMSPort = txtNMSPort.Text
        eMailArchiveMigrationManager.NuixNMSUserName = txtNMSUsername.Text
        eMailArchiveMigrationManager.NuixNMSAdminInfo = txtNMSPassword.Text
        eMailArchiveMigrationManager.NuixInstances = cboNumberOfNuixInstances.Text
        eMailArchiveMigrationManager.NuixAppMemory = txtMemoryPerNuixInstance.Text
        eMailArchiveMigrationManager.NuixWorkers = cboNuixWorkers.Text
        eMailArchiveMigrationManager.NuixWorkerMemory = txtMemoryPerWorker.Text
        eMailArchiveMigrationManager.NuixCaseDir = txtNuixCaseDir.Text
        eMailArchiveMigrationManager.NuixFilesDir = txtNuixFilesDirectory.Text
        eMailArchiveMigrationManager.NuixLogDir = txtLogDirectory.Text
        eMailArchiveMigrationManager.NuixJavaTempDir = txtJavaTempDir.Text
        eMailArchiveMigrationManager.NuixWorkerTempDir = txtWorkerTempDir.Text
        eMailArchiveMigrationManager.NuixExportDir = txtExportDir.Text
        eMailArchiveMigrationManager.NuixAppLocation = txtNuixAppLocation.Text
        eMailArchiveMigrationManager.PSTExportSize = numPSTExportSize.Value
        eMailArchiveMigrationManager.PSTAddDistributionListMetadata = chkPSTAddDistributionListMetadata.Checked
        eMailArchiveMigrationManager.EMLAddDistributionListMetadata = chkEMLAddDistributionListMetadata.Checked
        eMailArchiveMigrationManager.WorkerTimeout = numWorkerTimeout.Value
        eMailArchiveMigrationManager.ExtractWorkerTimeout = numExtractWorkerTimeout.Value
        eMailArchiveMigrationManager.SQLiteDBLocation = txtSQLiteDBLocation.Text
        eMailArchiveMigrationManager.RedisAuth = txtRedisAuth.Text
        eMailArchiveMigrationManager.RedisHost = txtRedisHostName.Text
        eMailArchiveMigrationManager.RedisPort = txtRedisPort.Text
        eMailArchiveMigrationManager.EnableMailboxSlackSpace = chkEnableMailboxSlackspace.Checked
        eMailArchiveMigrationManager.PSTConsolidation = cboPSTConsolidation.Text

        Directory.CreateDirectory(eMailArchiveMigrationManager.NuixCaseDir)
        Directory.CreateDirectory(eMailArchiveMigrationManager.NuixFilesDir)
        Directory.CreateDirectory(eMailArchiveMigrationManager.NuixLogDir)
        Directory.CreateDirectory(eMailArchiveMigrationManager.NuixJavaTempDir)
        Directory.CreateDirectory(eMailArchiveMigrationManager.NuixExportDir)
        Directory.CreateDirectory(eMailArchiveMigrationManager.NuixWorkerTempDir)

        Dim dbService As New DatabaseService
        Dim sSqlDataBaseLocation As String = eMailArchiveMigrationManager.SQLiteDBLocation
        Dim sSqlDataBaseName As String = "NuixEmailArchiveMigrationManager.db3"
        Dim sSqlDataBaseFullName As String = System.IO.Path.Combine(sSqlDataBaseLocation, sSqlDataBaseName)

        bStatus = common.BuildSQLiteRubyScript(txtSQLiteDBLocation.Text)
        bStatus = common.BuildSQLiteDatabaseScript(txtSQLiteDBLocation.Text)
        dbService.CreateDataBase(sSqlDataBaseFullName, sSqlDataBaseLocation)

        Me.Close()

    End Sub

    Private Sub btnNuixAppLocation_Click(sender As Object, e As EventArgs) Handles btnNuixAppLocation.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        With OpenFileDialog1
            .Filter = "nuix_console.exe|nuix_console.exe"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            txtNuixAppLocation.Text = OpenFileDialog1.FileName.ToString
        End If
    End Sub

    Private Sub btnCaseDirSel_Click(sender As Object, e As EventArgs) Handles btnCaseDirSel.Click
        Dim fldrDialog As New FolderBrowserDialog
        Dim sSelectedPath As String

        fldrDialog.ShowNewFolderButton = True

        If (fldrDialog.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            sSelectedPath = fldrDialog.SelectedPath
            txtNuixCaseDir.Text = sSelectedPath
        End If
    End Sub

    Private Sub btnLogDir_Click(sender As Object, e As EventArgs) Handles btnLogDir.Click
        Dim fldrDialog As New FolderBrowserDialog
        Dim sSelectedPath As String

        fldrDialog.ShowNewFolderButton = True

        If (fldrDialog.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            sSelectedPath = fldrDialog.SelectedPath
            txtLogDirectory.Text = sSelectedPath
        End If
    End Sub

    Private Sub btnJavaTempDir_Click(sender As Object, e As EventArgs) Handles btnJavaTempDir.Click
        Dim fldrDialog As New FolderBrowserDialog
        Dim sSelectedPath As String

        fldrDialog.ShowNewFolderButton = True

        If (fldrDialog.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            sSelectedPath = fldrDialog.SelectedPath
            txtJavaTempDir.Text = sSelectedPath
        End If

    End Sub

    Private Sub btnWorkerTempDir_Click(sender As Object, e As EventArgs) Handles btnWorkerTempDir.Click
        Dim fldrDialog As New FolderBrowserDialog
        Dim sSelectedPath As String

        fldrDialog.ShowNewFolderButton = True

        If (fldrDialog.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            sSelectedPath = fldrDialog.SelectedPath
            txtWorkerTempDir.Text = sSelectedPath
        End If

    End Sub

    Private Sub btnExportDir_Click(sender As Object, e As EventArgs) Handles btnExportDir.Click
        Dim fldrDialog As New FolderBrowserDialog
        Dim sSelectedPath As String

        fldrDialog.ShowNewFolderButton = True

        If (fldrDialog.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            sSelectedPath = fldrDialog.SelectedPath
            txtExportDir.Text = sSelectedPath
        End If

    End Sub

    Private Sub btnSQLiteDBLocationSelection_Click(sender As Object, e As EventArgs) Handles btnSQLiteDBLocationSelection.Click
        Dim fldrDialog As New FolderBrowserDialog
        Dim sSelectedPath As String

        fldrDialog.ShowNewFolderButton = True

        If (fldrDialog.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            sSelectedPath = fldrDialog.SelectedPath
            txtSQLiteDBLocation.Text = sSelectedPath
        End If
    End Sub

    Private Sub cboSourceType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSourceType.SelectedIndexChanged
        If cboSourceType.Text = "Desktop" Then
            txtNMSAddress.Enabled = False
            txtNMSPort.Enabled = False
            txtNMSPassword.Enabled = False
            txtNMSUsername.Enabled = False
            txtNMSAddress.Text = ""
            txtNMSPort.Text = ""
            txtNMSPassword.Text = ""
            txtNMSUsername.Text = ""
        ElseIf cboSourceType.Text = "Server" Then
            txtNMSAddress.Text = ""
            txtNMSPort.Text = ""
            txtNMSPassword.Text = ""
            txtNMSUsername.Text = ""
            txtNMSAddress.Enabled = True
            txtNMSPort.Enabled = True
            txtNMSPassword.Enabled = True
            txtNMSUsername.Enabled = True
        ElseIf cboSourceType.Text = "Cloud Server" Then
            txtNMSAddress.Text = ""
            txtNMSPort.Text = ""
            txtNMSPassword.Text = ""
            txtNMSUsername.Text = ""
            txtNMSAddress.Enabled = False
            txtNMSPort.Enabled = False
            txtNMSPassword.Enabled = True
            txtNMSUsername.Enabled = True
        End If
    End Sub

    Private Sub tabArchiveExtractionSettings_Click(sender As Object, e As EventArgs) Handles tabArchiveExtractionSettings.Click
        Dim dblTotalMemory As Double
        Dim dblAvailableMemory As Double

        dblTotalMemory = System.Math.Round(My.Computer.Info.TotalPhysicalMemory / (1024 * 1024), 2).ToString
        dblAvailableMemory = System.Math.Round(My.Computer.Info.AvailablePhysicalMemory / (1024 * 1024), 2).ToString

        lblAvailableRAM.Text = dblTotalMemory
        lblAvailableSystemMemory.Text = dblAvailableMemory
        If cboNumberOfNuixInstances.Text <> vbNullString Then
            If (CInt(cboNumberOfNuixInstances.Text) * CInt(txtMemoryPerNuixInstance.Text) < 0) Then
                MessageBox.Show("You do not have enough memory to run " & cboNumberOfNuixInstances.Text & " with " & txtMemoryPerNuixInstance.Text & " instances of Nuix.  Please adjust instances or memory.", "System Memory Issue.")
                cboNumberOfNuixInstances.Focus()
                lblAvailableRAMAfterNuix.Text = CInt(cboNumberOfNuixInstances.Text * txtMemoryPerNuixInstance.Text)
                Exit Sub
            End If

            lblAvailableRAMAfterNuix.Text = CInt(cboNumberOfNuixInstances.Text * txtMemoryPerNuixInstance.Text)
        End If
        If cboNuixWorkers.Text <> vbNullString Then
            txtMemoryPerWorker.Text = CInt(lblAvailableRAMAfterNuix.Text) / CInt(cboNuixWorkers.Text)
        End If
    End Sub

    Private Sub tabO365Settings_Click(sender As Object, e As EventArgs) Handles tabEWSSettings.Click
        Dim dblTotalMemory As Double
        Dim dblAvailableMemory As Double

        dblTotalMemory = System.Math.Round(My.Computer.Info.TotalPhysicalMemory / (1024 * 1024), 2).ToString
        dblAvailableMemory = System.Math.Round(My.Computer.Info.AvailablePhysicalMemory / (1024 * 1024), 2).ToString

        lblAvailableRAM.Text = dblTotalMemory
        lblAvailableSystemMemory.Text = dblAvailableMemory
        If cboNumberOfNuixInstances.Text <> vbNullString Then
            If (CInt(cboO365NumberOfNuixInstances.Text) * CInt(txtO365NuixAppMemory.Text) < 0) Then
                MessageBox.Show("You do not have enough memory to run " & cboO365NumberOfNuixInstances.Text & " with " & txtO365NuixAppMemory.Text & " instances of Nuix.  Please adjust instances or memory.", "System Memory Issue.")
                cboO365NumberOfNuixInstances.Focus()
                lblAvailableRAMAfterNuix.Text = CInt(cboO365NumberOfNuixInstances.Text * txtO365NuixAppMemory.Text)
                Exit Sub
            End If

            lblAvailableRAMAfterNuix.Text = CInt(cboO365NumberOfNuixInstances.Text * txtO365NuixAppMemory.Text)
        End If
        If cboNuixWorkers.Text <> vbNullString Then
            txtO365MemoryPerWorker.Text = CInt(lblO365AvailableMemoryAfterNuix.Text) / CInt(cboO365NumberOfNuixWorkers.Text)
        End If
    End Sub

    Private Sub cboO365NumberOfNuixInstances_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboO365NumberOfNuixInstances.SelectedIndexChanged
        Dim dblNuixMemory As Double
        Dim dblAvailableMemory As Double
        dblAvailableMemory = CDbl(lblO365AvailMemory.Text)

        If (cboO365NumberOfNuixWorkers.Text <> vbNullString) And (txtO365NuixAppMemory.Text <> vbNullString) Then
            dblNuixMemory = CInt(cboO365NumberOfNuixWorkers.Text) * CInt(txtO365NuixAppMemory.Text)
            If dblAvailableMemory - dblNuixMemory < 0 Then
                MessageBox.Show("You do not have enough memory available to run " & cboO365NumberOfNuixInstances.Text & " with " & txtO365NuixAppMemory.Text, "Memory Available Issue.")
                lblAvailableRAMAfterNuix.Text = dblAvailableMemory - dblNuixMemory
            End If
        End If
    End Sub

    Private Sub txtO365NuixAppMemory_TextChanged(sender As Object, e As EventArgs) Handles txtO365NuixAppMemory.TextChanged
        Dim dblNuixMemory As Double
        Dim dblAvailableMemory As Double
        dblAvailableMemory = CDbl(lblO365AvailMemory.Text)

        If ((cboO365NumberOfNuixWorkers.Text <> vbNullString) And (txtO365NuixAppMemory.Text)) Then

            dblNuixMemory = CInt(cboNumberOfNuixInstances.Text) * CInt(txtO365NuixAppMemory.Text)
            If dblAvailableMemory - dblNuixMemory < 0 Then
                MessageBox.Show("You do not have enough memory available to run " & cboO365NumberOfNuixInstances.Text & " with " & txtO365NuixAppMemory.Text, "Memory Available Issue.")
                lblAvailableRAMAfterNuix.Text = dblAvailableMemory - dblNuixMemory
            End If
        End If

    End Sub

    Private Sub chkEnableBulkUpload_CheckedChanged(sender As Object, e As EventArgs) Handles chkEnableBulkUpload.CheckedChanged
        If chkEnableBulkUpload.Checked = True Then
            numBulkUploadSize.Enabled = True
            numBulkUploadSize.Value = 10
            numEWSMaxUploadSize.Enabled = False
        Else
            numBulkUploadSize.Enabled = False
            numEWSMaxUploadSize.Enabled = True
        End If
    End Sub

    Private Sub btnLoadSettingXML_MouseHover(sender As Object, e As EventArgs) Handles btnLoadSettingXML.MouseHover
        GlobalSettingsToolTip.Show("Load Previous Settings Configuration...", btnLoadSettingXML)
    End Sub

    Private Sub btnSaveSettings_MouseHover(sender As Object, e As EventArgs) Handles btnSaveSettings.MouseHover
        GlobalSettingsToolTip.Show("Save Current Settings Configuration...", btnSaveSettings)

    End Sub

    Private Sub btnSettingsOK_MouseHover(sender As Object, e As EventArgs) Handles btnSettingsOK.MouseHover
        GlobalSettingsToolTip.Show("Load Settings Configuration into memory", btnSettingsOK)

    End Sub

    Private Sub btnCancel_MouseHover(sender As Object, e As EventArgs) Handles btnCancel.MouseHover
        GlobalSettingsToolTip.Show("Cancel", btnCancel)

    End Sub

    Private Sub chkEnablePrefetch_CheckedChanged(sender As Object, e As EventArgs) Handles chkEnablePrefetch.CheckedChanged
        If chkEnablePrefetch.Checked = True Then
            numDownloadSize.Enabled = True
            numDownloadCount.Enabled = True
        Else
            numDownloadSize.Enabled = False
            numDownloadCount.Enabled = False

        End If
    End Sub


End Class