Imports System.IO

Public Class eMailArchiveMigrationManager

    Public piO365NuixInstances As Integer
    Public piO365NuixAppMemory As Integer
    Public piO365MemoryPerWorker As Integer
    Public piO365NumberOfNuixWorkers As Integer
    Public piO365RetryCount As Integer
    Public piO365RetryDelay As Integer
    Public piO365RetryIncrement As Integer
    Public pbO365ApplicationImpersonation As Boolean
    Public psO365FilePathTrimming As String
    Public psO365ExchangeServer As String
    Public psO365Domain As String
    Public psO365AdminUsername As String
    Public psO365AdminInfo As String
    Public piO365MaxUploadSize As Integer
    Public piMaxMessageSize As Integer
    Public pbEnableBulkUpload As Boolean
    Public piMaxDownloadSize As Integer
    Public piMaxDownloadCount As Integer
    Public pbEnablePrefetch As Boolean
    Public pbEnableCollaborativeFetch As Boolean
    Public piBulkUploadSize As Integer
    Public piNuixInstances As Integer
    Public piNuixAppMemory As Integer
    Public psMemoryPerWorker As String
    Public piNumberOfNuixWorkers As Integer
    Public piNuixWorkers As Integer
    Public piNuixMemoryPerWorker As Integer
    Public piPSTExportSize As Integer
    Public pbPSTAddDistributionListMetadata As Boolean
    Public pbEMLAddDistributionListMetadata As Boolean
    Public psNuixCaseDir As String
    Public psNuixFilesDir As String
    Public psNuixLogDir As String
    Public psJavaTempDirectory As String
    Public psWorkerTempDirectory As String
    Public psNuixExportDir As String
    Public psNuixAppDirectory As String
    Public piWorkerTimeout As Integer
    Public psNMSAddress As String
    Public psNMSPort As String
    Public psNMSUserName As String
    Public psNMSAdminInfo As String
    Public psSettingsFile As String
    Public psSQLiteLocation As String
    Public psNMSSourceType As String
    Public pbSettingsLoaded As Boolean
    Public piExtractWorkerTimeout As Integer
    Public psaddBccToEmailDigests As String
    Public psperformItemIdentification As String
    Public psaddCommunicationDateToEmailDigests As String
    Public psanalysisLanguage As String
    Public pscalculateAuditedSize As String
    Public pscalculateSSDeepFuzzyHash As String
    Public pscarveFileSystemUnallocatedSpace As String
    Public pscreateThumbnails As String
    Public psdetectFaces As String
    Public psdigests As String
    Public psenableExactQueries As String
    Public psextractEndOfFileSlackSpace As String
    Public psextractFromSlackSpace As String
    Public psextractNamedEntitiesFromProperties As String
    Public psextractNamedEntitiesFromText As String
    Public psextractNamedEntitiesFromTextStripped As String
    Public psextractShingles As String
    Public pshideEmbeddedImmaterialData As String
    Public psidentifyPhysicalFiles As String
    Public psmaxDigestSize As String
    Public psmaxStoredBinarySize As String
    Public psprocessFamilyFields As String
    Public psprocessText As String
    Public psprocessTextSummaries As String
    Public psrecoverDeletedFiles As String
    Public psreportProcessingStatus As String
    Public psreuseEvidenceStores As String
    Public psskinToneAnalysis As String
    Public pssmartProcessRegistry As String
    Public psstemming As String
    Public psstopWords As String
    Public psstoreBinary As String
    Public pstraversalScope As String
    Public psRedisHost As String
    Public psRedisPort As String
    Public psRedisAuth As String
    Public pbEnableMailboxSlackSpace As Boolean
    Public psPSTConsolidation As String

    Private SummaryReportThread As System.Threading.Thread
    Private SQLiteDBReadThread As System.Threading.Thread


    Private Sub btnUploadtoO365_Click(sender As Object, e As EventArgs) Handles btnUploadtoO365.Click

        Dim NuixO365Ingestion As frmNuixIngestion
        NuixO365Ingestion = New frmNuixIngestion
        NuixO365Ingestion.ShowDialog()
    End Sub

    Private Sub btnExtractFromO365_Click(sender As Object, e As EventArgs) Handles btnExtractFromO365.Click

        Dim NuixO365Extraction As O365Puller
        NuixO365Extraction = New O365Puller
        NuixO365Extraction.ShowDialog()

    End Sub

    Private Sub btnNuixSettings_Click(sender As Object, e As EventArgs) Handles btnNuixSettings.Click

        Dim SettingsForm As O365ExtractionSettings
        psSettingsFile = Me.NuixSettingsFile
        If psSettingsFile = vbNullString Then

            SettingsForm = New O365ExtractionSettings

            SettingsForm.ShowDialog()
            SettingsForm.NuixSettingsFile = psSettingsFile
        Else
            SettingsForm = New O365ExtractionSettings
            SettingsForm.NuixSettingsFile = psSettingsFile

            SettingsForm.ShowDialog()
            Me.NuixSettingsFile = psSettingsFile
        End If
        psSettingsFile = Me.NuixSettingsFile

    End Sub

    Public Property NuixSettingsLoaded As Boolean
        Get
            Return pbSettingsLoaded
        End Get
        Set(value As Boolean)
            pbSettingsLoaded = True
        End Set
    End Property

    Public Property performItemIdentification As String

        Get
            Return psperformItemIdentification
        End Get
        Set(value As String)
            psperformItemIdentification = value
        End Set
    End Property

    Public Property addBccToEmailDigests As String

        Get
            Return psaddBccToEmailDigests
        End Get
        Set(value As String)
            psaddBccToEmailDigests = value
        End Set
    End Property

    Public Property addCommunicationDateToEmailDigests As String

        Get
            Return psaddCommunicationDateToEmailDigests
        End Get
        Set(value As String)
            psaddCommunicationDateToEmailDigests = value
        End Set
    End Property

    Public Property analysisLanguage As String
        Get
            Return psanalysisLanguage
        End Get
        Set(value As String)
            psanalysisLanguage = value
        End Set
    End Property

    Public Property calculateAuditedSize As String

        Get
            Return pscalculateAuditedSize
        End Get
        Set(value As String)
            pscalculateAuditedSize = value
        End Set
    End Property

    Public Property calculateSSDeepFuzzyHash As String

        Get
            Return pscalculateSSDeepFuzzyHash
        End Get
        Set(value As String)
            pscalculateSSDeepFuzzyHash = value
        End Set
    End Property

    Public Property carveFileSystemUnallocatedSpace As String

        Get
            Return pscarveFileSystemUnallocatedSpace
        End Get
        Set(value As String)
            pscarveFileSystemUnallocatedSpace = value
        End Set
    End Property

    Public Property createThumbnails As String

        Get
            Return pscreateThumbnails
        End Get
        Set(value As String)
            pscreateThumbnails = value
        End Set
    End Property

    Public Property detectFaces As String

        Get
            Return psdetectFaces
        End Get
        Set(value As String)
            psdetectFaces = value
        End Set
    End Property

    Public Property digests As String

        Get
            Return psdigests
        End Get
        Set(value As String)
            psdigests = value
        End Set
    End Property

    Public Property enableExactQueries As String

        Get
            Return psenableExactQueries
        End Get
        Set(value As String)
            psenableExactQueries = value
        End Set
    End Property

    Public Property extractEndOfFileSlackSpace As String

        Get
            Return psextractEndOfFileSlackSpace
        End Get
        Set(value As String)
            psextractEndOfFileSlackSpace = value
        End Set
    End Property

    Public Property extractFromSlackSpace As String

        Get
            Return psextractFromSlackSpace
        End Get
        Set(value As String)
            psextractFromSlackSpace = value
        End Set
    End Property

    Public Property extractNamedEntitiesFromProperties As String

        Get
            Return psextractNamedEntitiesFromProperties
        End Get
        Set(value As String)
            psextractNamedEntitiesFromProperties = value
        End Set
    End Property

    Public Property extractNamedEntitiesFromText As String

        Get
            Return psextractNamedEntitiesFromText
        End Get
        Set(value As String)
            psextractNamedEntitiesFromText = value
        End Set
    End Property

    Public Property extractNamedEntitiesFromTextStripped As String

        Get
            Return psextractNamedEntitiesFromTextStripped
        End Get
        Set(value As String)
            psextractNamedEntitiesFromTextStripped = value
        End Set
    End Property

    Public Property extractShingles As String

        Get
            Return psextractShingles
        End Get
        Set(value As String)
            psextractShingles = value
        End Set
    End Property

    Public Property hideEmbeddedImmaterialData As String

        Get
            Return pshideEmbeddedImmaterialData
        End Get
        Set(value As String)
            pshideEmbeddedImmaterialData = value
        End Set
    End Property

    Public Property identifyPhysicalFiles As String

        Get
            Return psidentifyPhysicalFiles
        End Get
        Set(value As String)
            psidentifyPhysicalFiles = value
        End Set
    End Property

    Public Property maxDigestSize As String

        Get
            Return psmaxDigestSize
        End Get
        Set(value As String)
            psmaxDigestSize = value
        End Set
    End Property

    Public Property maxStoredBinarySize As String

        Get
            Return psmaxStoredBinarySize
        End Get
        Set(value As String)
            psmaxStoredBinarySize = value
        End Set
    End Property

    Public Property processFamilyFields As String

        Get
            Return psprocessFamilyFields
        End Get
        Set(value As String)
            psprocessFamilyFields = value
        End Set
    End Property

    Public Property processText As String

        Get
            Return psprocessText
        End Get
        Set(value As String)
            psprocessText = value
        End Set
    End Property

    Public Property processTextSummaries As String

        Get
            Return psprocessTextSummaries
        End Get
        Set(value As String)
            psprocessTextSummaries = value
        End Set
    End Property

    Public Property recoverDeletedFiles As String

        Get
            Return psrecoverDeletedFiles
        End Get
        Set(value As String)
            psrecoverDeletedFiles = value
        End Set
    End Property

    Public Property reportProcessingStatus As String

        Get
            Return psreportProcessingStatus
        End Get
        Set(value As String)
            psreportProcessingStatus = value
        End Set
    End Property

    Public Property reuseEvidenceStores As String

        Get
            Return psreuseEvidenceStores
        End Get
        Set(value As String)
            psreuseEvidenceStores = value
        End Set
    End Property

    Public Property skinToneAnalysis As String

        Get
            Return psskinToneAnalysis
        End Get
        Set(value As String)
            psskinToneAnalysis = value
        End Set
    End Property

    Public Property smartProcessRegistry As String

        Get
            Return pssmartProcessRegistry
        End Get
        Set(value As String)
            pssmartProcessRegistry = value
        End Set
    End Property

    Public Property stemming As String

        Get
            Return psstemming
        End Get
        Set(value As String)
            psstemming = value
        End Set
    End Property

    Public Property stopWords As String

        Get
            Return psstopWords
        End Get
        Set(value As String)
            psstopWords = value
        End Set
    End Property

    Public Property storeBinary As String

        Get
            Return psstoreBinary
        End Get
        Set(value As String)
            psstoreBinary = value
        End Set
    End Property

    Public Property traversalScope As String

        Get
            Return pstraversalScope
        End Get
        Set(value As String)
            pstraversalScope = value
        End Set
    End Property

    Public Property O365ExchangeServer As String
        Get
            Return psO365ExchangeServer
        End Get
        Set(value As String)
            psO365ExchangeServer = value

        End Set
    End Property

    Public Property O365Domain As String
        Get
            Return psO365Domain
        End Get
        Set(value As String)
            psO365Domain = value
        End Set
    End Property

    Public Property O365AdminUserName As String
        Get
            Return psO365AdminUsername
        End Get
        Set(value As String)
            psO365AdminUserName = value
        End Set
    End Property

    Public Property O365AdminInfo As String
        Get
            Return psO365AdminInfo

        End Get
        Set(value As String)
            psO365AdminInfo = value
        End Set
    End Property

    Public Property O365ApplicationImpersonation As Boolean
        Get
            Return pbO365ApplicationImpersonation

        End Get
        Set(value As Boolean)
            pbO365ApplicationImpersonation = value
        End Set
    End Property

    Public Property O365RetryCount As String
        Get
            Return piO365RetryCount
        End Get
        Set(value As String)
            piO365RetryCount = value
        End Set
    End Property

    Public Property NMSSourceType As String
        Get
            Return psNMSSourceType
        End Get
        Set(value As String)
            psNMSSourceType = value
        End Set
    End Property

    Public Property O365RetryDelay As Integer
        Get
            Return piO365RetryDelay
        End Get
        Set(value As Integer)
            piO365RetryDelay = value
        End Set
    End Property

    Public Property O365RetryIncrement As Integer
        Get
            Return piO365RetryIncrement
        End Get
        Set(value As Integer)
            piO365RetryIncrement = value
        End Set
    End Property

    Public Property WorkerTimeout As Integer
        Get
            Return piWorkerTimeout
        End Get
        Set(value As Integer)
            piWorkerTimeout = value
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


    Public Property O365FilePathTrimming As String
        Get
            Return psO365FilePathTrimming
        End Get
        Set(value As String)
            psO365FilePathTrimming = value
        End Set
    End Property

    Public Property O365MaxUploadSize As Integer
        Get
            Return piO365MaxUploadSize
        End Get
        Set(value As Integer)
            piO365MaxUploadSize = value
        End Set
    End Property

    Public Property o365MaxMessageSize As Integer
        Get
            Return piMaxMessageSize
        End Get
        Set(value As Integer)
            piMaxMessageSize = value
        End Set
    End Property

    Public Property O365EnableBulkUpload As Boolean
        Get
            Return pbEnableBulkUpload
        End Get
        Set(value As Boolean)
            pbEnableBulkUpload = value
        End Set
    End Property

    Public Property O365BulkUploadSize As Integer
        Get
            Return piBulkUploadSize
        End Get
        Set(value As Integer)
            piBulkUploadSize = value
        End Set
    End Property

    Public Property O365MaxDownloadSize As Integer
        Get
            Return piMaxDownloadSize
        End Get
        Set(value As Integer)
            piMaxDownloadSize = value
        End Set
    End Property

    Public Property O365MaxDownloadCount As Integer
        Get
            Return piMaxDownloadCount
        End Get
        Set(value As Integer)
            piMaxDownloadCount = value
        End Set
    End Property

    Public Property O365EnablePrefetch As Boolean
        Get
            Return pbEnablePrefetch
        End Get
        Set(value As Boolean)
            pbEnablePrefetch = value
        End Set
    End Property

    Public Property O365EnableCollaborativePrefetch As Boolean
        Get
            Return pbEnableCollaborativeFetch
        End Get
        Set(value As Boolean)
            pbEnableCollaborativeFetch = value
        End Set
    End Property

    Public Property NuixNMSAddress As String
        Get
            Return psNMSAddress
        End Get
        Set(value As String)
            psNMSAddress = value
        End Set
    End Property

    Public Property NuixNMSPort As String
        Get
            Return psNMSPort
        End Get
        Set(value As String)
            psNMSPort = value
        End Set
    End Property

    Public Property NuixNMSUserName As String
        Get
            Return psNMSUserName

        End Get
        Set(value As String)
            psNMSUserName = value
        End Set
    End Property

    Public Property NuixNMSAdminInfo As String
        Get
            Return psNMSAdminInfo
        End Get
        Set(value As String)
            psNMSAdminInfo = value
        End Set
    End Property

    Public Property NuixInstances As Integer
        Get
            Return piNuixInstances
        End Get
        Set(ByVal value As Integer)
            piNuixInstances = value
        End Set
    End Property

    Public Property NuixAppMemory As Integer
        Get
            Return piNuixAppMemory
        End Get

        Set(ByVal value As Integer)
            piNuixAppMemory = value
        End Set
    End Property

    Public Property NuixWorkers As Integer
        Get
            Return piNuixWorkers

        End Get
        Set(value As Integer)
            piNuixWorkers = value
        End Set
    End Property

    Public Property NuixWorkerMemory As Integer
        Get
            Return piNuixMemoryPerWorker
        End Get
        Set(value As Integer)
            piNuixMemoryPerWorker = value
        End Set
    End Property

    Public Property NuixCaseDir As String
        Get
            Return psNuixCaseDir

        End Get
        Set(value As String)
            psNuixCaseDir = value
        End Set
    End Property

    Public Property NuixFilesDir As String
        Get
            Return psNuixFilesDir

        End Get
        Set(value As String)
            psNuixFilesDir = value
        End Set
    End Property

    Public Property NuixLogDir As String
        Get
            Return psNuixLogDir

        End Get
        Set(value As String)
            psNuixLogDir = value
        End Set
    End Property

    Public Property NuixJavaTempDir As String
        Get
            Return psJavaTempDirectory

        End Get
        Set(value As String)
            psJavaTempDirectory = value
        End Set
    End Property

    Public Property NuixWorkerTempDir As String
        Get
            Return psWorkerTempDirectory

        End Get
        Set(value As String)
            psWorkerTempDirectory = value
        End Set
    End Property

    Public Property NuixExportDir As String
        Get
            Return psNuixExportDir

        End Get
        Set(value As String)
            psNuixExportDir = value
        End Set
    End Property

    Public Property NuixAppLocation As String
        Get
            Return psNuixAppDirectory

        End Get
        Set(value As String)
            psNuixAppDirectory = value
        End Set
    End Property

    Public Property SQLiteDBLocation As String
        Get
            Return psSQLiteLocation
        End Get
        Set(value As String)
            psSQLiteLocation = value
        End Set
    End Property

    Public Property O365NuixInstances As Integer
        Get
            Return piO365NuixInstances
        End Get
        Set(value As Integer)
            piO365NuixInstances = value
        End Set
    End Property

    Public Property O365NuixAppMemory As Integer
        Get
            Return piO365NuixAppMemory
        End Get
        Set(value As Integer)
            piO365NuixAppMemory = value
        End Set
    End Property

    Public Property O365NumberOfNuixWorkers As Integer
        Get
            Return piO365NumberOfNuixWorkers
        End Get
        Set(value As Integer)
            piO365NumberOfNuixWorkers = value
        End Set
    End Property

    Public Property O365MemoryPerWorker As Integer
        Get
            Return piO365MemoryPerWorker
        End Get
        Set(value As Integer)
            piO365MemoryPerWorker = value
        End Set
    End Property

    Public Property RedisHost As String
        Get
            Return psRedisHost
        End Get
        Set(value As String)
            psRedisHost = value
        End Set
    End Property

    Public Property RedisPort As String
        Get
            Return psRedisPort
        End Get
        Set(value As String)
            psRedisPort = value
        End Set
    End Property

    Public Property RedisAuth As String
        Get
            Return psRedisAuth
        End Get
        Set(value As String)
            psRedisAuth = value
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
    Public Property NuixSettingsFile As String
        Get
            Return psSettingsFile
        End Get
        Set(value As String)
            psSettingsFile = value
        End Set
    End Property

    Public Property PSTConsolidation As String
        Get
            Return psPSTConsolidation
        End Get
        Set(value As String)
            psPSTConsolidation = value
        End Set
    End Property


    Private Sub btnExtractFromLegacy_Click(sender As Object, e As EventArgs) Handles btnExtractFromLegacy.Click
        Dim LegacyArchive As LegacyArchiveExtraction
        LegacyArchive = New LegacyArchiveExtraction
        LegacyArchive.ShowDialog()

    End Sub

    Private Sub btnCustodianListGenerator_Click(sender As Object, e As EventArgs) Handles btnCustodianListGenerator.Click
        Dim CustomerListGenerator As NuixCustodianListGenerator
        CustomerListGenerator = New NuixCustodianListGenerator
        CustomerListGenerator.ShowDialog()

    End Sub

    Private Sub btnSourceDataConversion_Click(sender As Object, e As EventArgs) Handles btnSourceDataConversion.Click
        Dim EmailConversion As NSFConversion
        EmailConversion = New NSFConversion
        NSFConversion.ShowDialog()

    End Sub

    Private Sub eMailArchiveMigrationManager_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        btnCustodianListGenerator.Hide()
    End Sub

    Private Sub btnNuixSettings_MouseHover(sender As Object, e As EventArgs) Handles btnNuixSettings.MouseHover
        SettingsToolTip.Show("Global Settings", btnNuixSettings)
    End Sub

    Private Sub btnSourceDataConversion_MouseHover(sender As Object, e As EventArgs) Handles btnSourceDataConversion.MouseHover
        SettingsToolTip.Show("Start Legacy Email Conversion", btnSourceDataConversion)
    End Sub

    Private Sub btnUploadtoO365_MouseHover(sender As Object, e As EventArgs) Handles btnUploadtoO365.MouseHover
        SettingsToolTip.Show("Start EWS Ingestion", btnUploadtoO365)
    End Sub

    Private Sub btnExtractFromLegacy_MouseHover(sender As Object, e As EventArgs) Handles btnExtractFromLegacy.MouseHover
        SettingsToolTip.Show("Start Legacy Email Archive Extraction", btnExtractFromLegacy)
    End Sub

    Private Sub btnExtractFromO365_MouseHover(sender As Object, e As EventArgs) Handles btnExtractFromO365.MouseHover
        SettingsToolTip.Show("Start EWS Extraction", btnExtractFromO365)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()

    End Sub

    Private Sub StartEWSExtraction_Click(sender As Object, e As EventArgs) Handles StartEWSExtraction.Click
        Dim NuixO365Extraction As O365Puller
        NuixO365Extraction = New O365Puller
        NuixO365Extraction.ShowDialog()
    End Sub

    Private Sub GlobalSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GlobalSettingsToolStripMenuItem.Click

        Dim SettingsForm As O365ExtractionSettings
        psSettingsFile = Me.NuixSettingsFile
        If psSettingsFile = vbNullString Then

            SettingsForm = New O365ExtractionSettings

            SettingsForm.ShowDialog()
            SettingsForm.NuixSettingsFile = psSettingsFile
        Else
            SettingsForm = New O365ExtractionSettings
            SettingsForm.NuixSettingsFile = psSettingsFile

            SettingsForm.ShowDialog()
            Me.NuixSettingsFile = psSettingsFile
        End If
        psSettingsFile = Me.NuixSettingsFile
    End Sub

    Private Sub AboutNEAMMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutNEAMMToolStripMenuItem.Click
        Dim value As System.Version = My.Application.Info.Version

        MessageBox.Show("Nuix Email Archive Migration Manager - v" & value.ToString, "NEAMM Version", MessageBoxButtons.OK)

    End Sub

    Private Sub OpenLogDirectoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenLogDirectoryToolStripMenuItem.Click
        If Me.NuixFilesDir = vbNullString Then
            MessageBox.Show("The settings have not been loaded yet.  No log directory specified", "No Log Directory specified", MessageBoxButtons.OK)
        Else
            Process.Start("explorer.exe", Me.NuixLogDir)
        End If
    End Sub

    Private Sub StartArchiveExtraction_Click(sender As Object, e As EventArgs) Handles StartArchiveExtraction.Click
        Dim ArchiveExtraction As LegacyArchiveExtraction
        ArchiveExtraction = New LegacyArchiveExtraction
        ArchiveExtraction.ShowDialog()
    End Sub

    Private Sub StartEWSIngestion_Click_1(sender As Object, e As EventArgs) Handles StartEWSIngestion.Click
        Dim EWSIngestion As frmNuixIngestion
        EWSIngestion = New frmNuixIngestion
        EWSIngestion.ShowDialog()

    End Sub

    Private Sub StartEmailConversion_Click_1(sender As Object, e As EventArgs) Handles StartEmailConversion.Click
        Dim EmailConversion As NSFConversion
        EmailConversion = New NSFConversion
        EmailConversion.ShowDialog()

    End Sub
End Class