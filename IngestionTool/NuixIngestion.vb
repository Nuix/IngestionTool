Imports System.IO
Imports System.Diagnostics
Imports System.Net
Imports System.Threading
Imports System.Xml
Imports System.Text
Imports System.Data.SQLite


Public Class frmNuixIngestion

    Public lstSelectedCustodiansForProcessing As List(Of String)
    Public psSettingsFile As String
    Public pbNoMoreJobs As Boolean
    Public psSummaryReportFile As String
    Public psIngestionLogFile As String
    Public psSQLiteLocation As String
    Public psNuixCaseDir As String
    Public plstNuixConsoleProcesses As List(Of String)

    Private SummaryReportThread As System.Threading.Thread
    Private SQLiteDBReadThread As System.Threading.Thread
    Private CaseLockThread As System.Threading.Thread

    Private Sub btnNuixSystemSettings_Click(sender As Object, e As EventArgs) Handles btnNuixSystemSettings.Click
        Dim SettingsForm As O365ExtractionSettings
        psSettingsFile = eMailArchiveMigrationManager.NuixSettingsFile

        If psSettingsFile = vbNullString Then

            SettingsForm = New O365ExtractionSettings

            SettingsForm.ShowDialog()

            '            psSettingsFile = SettingsForm.NuixSettingsFile
            SettingsForm.NuixSettingsFile = psSettingsFile
        Else
            SettingsForm = New O365ExtractionSettings
            SettingsForm.NuixSettingsFile = psSettingsFile

            SettingsForm.ShowDialog()
            'eMailArchiveMigrationManager.NuixSettingsFile = psSettingsFile
        End If
        psSettingsFile = eMailArchiveMigrationManager.NuixSettingsFile

    End Sub

    Private Function blnRedirectionUrlValidationCallback(ByVal redirectionUrl As String) As Boolean
        blnRedirectionUrlValidationCallback = False
        Dim redirectionUri As Uri

        redirectionUri = New Uri(redirectionUrl)

        If (redirectionUri.Scheme = "https") Then
            blnRedirectionUrlValidationCallback = True
        Else
            blnRedirectionUrlValidationCallback = False
        End If

    End Function


    Private Sub btnFileSystemChooser_Click(sender As Object, e As EventArgs) Handles btnFileSystemChooser.Click
        Dim fldrDialog As New FolderBrowserDialog
        Dim sSelectedPath As String

        fldrDialog.ShowNewFolderButton = True

        If (fldrDialog.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            sSelectedPath = fldrDialog.SelectedPath
            txtPSTLocation.Text = sSelectedPath
            btnLoadPSTInfo.Enabled = True
        Else
            btnLoadPSTInfo.Enabled = False
        End If
    End Sub

 
    Public Property NuixSettingsFile As String
        Get
            Return psSettingsFile
        End Get
        Set(value As String)
            psSettingsFile = value
        End Set
    End Property

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim msgboxReturn As DialogResult
        Dim sMachineName As String
        Dim common As New Common

        btnRefreshGridData.Hide()
        btnProcessingDetails.Enabled = False
        psSettingsFile = eMailArchiveMigrationManager.NuixSettingsFile
        If psSettingsFile = vbNullString Then
            msgboxReturn = MessageBox.Show("Nuix Email Archive Migration Manager has detected that the Global Settings have not been loaded.  Please load the settings now. ", "Global Settings Not Available", MessageBoxButtons.YesNo)
            If (msgboxReturn = vbYes) Then
                Dim NuixSettings As O365ExtractionSettings
                NuixSettings = New O365ExtractionSettings
                NuixSettings.ShowDialog()
                eMailArchiveMigrationManager.NuixSettingsFile = psSettingsFile
            Else
                Me.Close()
                Exit Sub

            End If
        Else
            psSettingsFile = eMailArchiveMigrationManager.psSettingsFile
        End If
        psSQLiteLocation = eMailArchiveMigrationManager.SQLiteDBLocation
        psNuixCaseDir = eMailArchiveMigrationManager.NuixCaseDir
        sMachineName = System.Net.Dns.GetHostName

        If eMailArchiveMigrationManager.NuixLogDir = vbNullString Then
            MessageBox.Show("There was no log directory specified.  Please update settings to include log directory.", "No Log Directory Specified")
            Exit Sub
        End If

        psIngestionLogFile = "EWS Ingestion Log - " & sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".log"

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

        ExportItemsNotUploadedToPST.Enabled = False
        ExportFTSDataTooLargeToolStripMenuItem.Enabled = False
        ExportOtherNotUploadedToolStripMenuItem.Enabled = False
        'SelectAllItemsNotProcessed.Enabled = False

    End Sub

    Private Sub btnLoadPSTInfo_Click(sender As Object, e As EventArgs) Handles btnLoadPSTInfo.Click
        Dim sPSTFolderName As String
        Dim lstNotConsolidatedPSTs As List(Of String)
        Dim sSourceFolder As String
        Dim sDestinationFolder As String
        Dim sReportCSVFile As String
        Dim iReturn As DialogResult
        Dim common As New Common

        Dim msgboxReturn As DialogResult

        Dim bStatus As Boolean

        lstNotConsolidatedPSTs = New List(Of String)
        If grdPSTInfo.Rows.Count > 1 Then
            iReturn = MessageBox.Show("Are you sure you want to clear the grid and reload the custodian PST files", "Clear Custodian Grid", MessageBoxButtons.YesNo)
            If iReturn = Windows.Forms.DialogResult.Yes Then
                grdPSTInfo.Rows.Clear()
            Else
                Exit Sub
            End If
        End If
        sPSTFolderName = txtPSTLocation.Text
        If sPSTFolderName = vbNullString Then
            MessageBox.Show("You must select a valid folder to search for the PST Files.", "Select valid PST Folder location")
            txtPSTLocation.Focus()
            Exit Sub
        Else
            common.Logger(psIngestionLogFile, "Loading PST information from - " & sPSTFolderName)
            bStatus = blnLoadPSTInfoFromDirectory(sPSTFolderName, grdPSTInfo, lstNotConsolidatedPSTs)
        End If

        If lstNotConsolidatedPSTs.Count > 0 Then
            'MessageBox.Show("It appears that the PSTs on the file system have not been consolidated.  Please consolidate prior to migrating data.  Removing unconsolidated cusotdian information.")
            msgboxReturn = MessageBox.Show("It appears that the PSTs on the file system have not been consolidated.  Please consolidate prior to migrating data.  Were the PSTs generated via a Nuix process?", "Consolidate Custodian PSTs", MessageBoxButtons.YesNoCancel)
            If msgboxReturn = vbCancel Then
                For Each CustodianNotConsolidated In lstNotConsolidatedPSTs
                    For Each row In grdPSTInfo.Rows
                        If row.cells("CustodianName").value = CustodianNotConsolidated.ToString Then
                            grdPSTInfo.Rows.Remove(row)
                        End If
                    Next
                Next
            ElseIf msgboxReturn = vbNo Then
                grdPSTInfo.Rows.Clear()
                If (txtPSTLocation.Text.Contains(".csv")) Then
                    sSourceFolder = txtPSTLocation.Text.Substring(0, txtPSTLocation.Text.LastIndexOf("\"))
                Else
                    sSourceFolder = txtPSTLocation.Text
                End If

                For Each Custodian In lstNotConsolidatedPSTs
                    sDestinationFolder = sSourceFolder & "\" & Custodian.ToString
                    common.Logger(psIngestionLogFile, "Consolidating Custodian PST Information in - " & sSourceFolder)

                    bStatus = blnConsolidateCustodianPSTs(sSourceFolder, Custodian.ToString, sDestinationFolder)
                Next
                bStatus = blnLoadPSTInfoFromDirectory(sPSTFolderName, grdPSTInfo, lstNotConsolidatedPSTs)
                MessageBox.Show("PST Consolidation complete.", "PST Consolidation")
            ElseIf msgboxReturn = vbYes Then
                bStatus = blnConsolidateNuixPSTs(txtPSTLocation.Text, txtPSTLocation.Text, sReportCSVFile)
                grdPSTInfo.Rows.Clear()
                bStatus = blnLoadPSTInfoFromCSV(sReportCSVFile, grdPSTInfo, True, lstNotConsolidatedPSTs)
                MessageBox.Show("PST Consolidation complete.", "PST Consolidation")

            End If

        End If

        If grdPSTInfo.RowCount > 1 Then
            btnProcessingDetails.Enabled = True
        Else
            btnProcessingDetails.Enabled = False
        End If
    End Sub

    Private Function blnConsolidateNuixPSTs(ByVal sSourceFolder As String, ByVal sDestinationFolder As String, ByRef sReportCSVFile As String) As Boolean
        Dim PSTName As New List(Of String)
        Dim PSTPath As New List(Of String)
        Dim PSTSize As New List(Of String)
        Dim lstCustodian As New List(Of String)
        Dim PSTServerNumber As New List(Of String)
        Dim UniqueName As New List(Of String)
        Dim sFolderName As String
        Dim sFileName As String
        Dim sFilePath As String
        Dim iCounter As Integer
        Dim sGUID As String
        Dim StartDate As DateTime = DateTime.Now
        Dim sNewFileName As String
        Dim sNewGUID As String
        Dim sCustodianName As String
        Dim sMachineName As String
        Dim ReportOutputFile As StreamWriter
        Dim iEmailCounter As Integer

        blnConsolidateNuixPSTs = False
        StartDate = Now

        sMachineName = System.Net.Dns.GetHostName()
        sReportCSVFile = sSourceFolder & "\" & sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv"

        ReportOutputFile = New StreamWriter(sReportCSVFile)
        subPSTDirSearch(sSourceFolder, PSTName, PSTPath, PSTSize)

        For Each item In PSTName
            sFolderName = item.ToString
            sFolderName = Path.GetFileNameWithoutExtension(sFolderName) 'sFolderName.Replace(".pst", "")

            If sFolderName.IndexOf("_") = -1 Then
            Else
                sFolderName = sFolderName.Substring(sFolderName.IndexOf("_"))
            End If

            UniqueName.Add(sFolderName)
            UniqueName = UniqueName.Distinct.ToList
        Next

        For Each foldername In UniqueName
            If Not System.IO.Directory.Exists(sDestinationFolder & "\" & foldername.ToString) Then
                System.IO.Directory.CreateDirectory(sDestinationFolder & "\" & foldername.ToString)
            End If
        Next

        For Each pst In PSTName
            iCounter = iCounter + 1
            sFileName = pst.ToString
            sFileName = Path.GetFileNameWithoutExtension(sFileName)
            sFilePath = PSTPath(iCounter - 1)
            sGUID = sFilePath.Substring(sFilePath.LastIndexOf("\") + 1)
            sNewGUID = sGUID.Substring(0, 11)

            sNewFileName = sFileName & "_" & iEmailCounter & ".pst"

            If sFileName.IndexOf("_") = -1 Then
                sCustodianName = sFileName
            Else

                If sFileName Like "*_#" Then
                    sCustodianName = Microsoft.VisualBasic.Left(sFileName, Len(sFileName) - 2)

                ElseIf sFileName Like "*_##" Then
                    sCustodianName = Microsoft.VisualBasic.Left(sFileName, Len(sFileName) - 3)
                Else
                    sCustodianName = sFileName
                End If
                lstCustodian.Add(sCustodianName)
            End If

            iEmailCounter = iEmailCounter + 1

            My.Computer.FileSystem.MoveFile(sFilePath & "\" & sFileName & ".pst", sDestinationFolder & "\" & sCustodianName & "\" & sNewFileName)
            '            OutputStream.WriteLine(sCustodianName & "," & sFileName & "," & PSTSize(iCounter - 1) & "," & sFilePath & "\" & sFileName & ".pst" & "," & sLegalHold)
            ReportOutputFile.WriteLine(sCustodianName & "," & sFileName & "," & PSTSize(iCounter - 1) & "," & sDestinationFolder & "\" & sCustodianName & "\")

        Next
        ReportOutputFile.Close()

    End Function

    Sub subNuixPSTDirSearch(ByVal sDir As String, ByRef PSTName As List(Of String), ByRef PSTPath As List(Of String), ByRef PSTSize As List(Of String))
        Dim d As String
        Dim Length As String
        Dim currentdirectory As DirectoryInfo
        Dim extension As String

        Try
            currentdirectory = New DirectoryInfo(sDir)

            For Each File In currentdirectory.GetFiles
                extension = Path.GetExtension(File.ToString)
                If extension = ".pst" Then 'If InStr(File.Name, ".pst") > 0 Then
                    PSTName.Add(File.Name)
                    PSTPath.Add(File.DirectoryName)
                    PSTSize.Add(File.Length)
                End If
            Next

            Length = Directory.GetDirectories(sDir).Length
            If Length > 0 Then
                For Each d In Directory.GetDirectories(sDir)
                    subNuixPSTDirSearch(d, PSTName, PSTPath, PSTSize)
                Next
            End If
        Catch ex As Exception
            MsgBox("Directory Search" & ex.ToString)
        End Try

    End Sub

    Public Function blnLoadPSTInfoFromCSV(ByVal sPSTCSVName As String, ByVal grdPSTInfo As DataGridView, ByVal bUpdateGrid As Boolean, ByRef lstNotConsolidatedCustodian As List(Of String)) As Boolean
        Dim CSVMappingFile As Microsoft.VisualBasic.FileIO.TextFieldParser
        Dim bStatus As Boolean
        Dim asCurrentRow() As String
        Dim sCustodian As String
        Dim sCustodianPSTs As String
        Dim asCustodianPSTs() As String
        Dim sCurrentPSTNames As String
        Dim dblCurrentPSTSize As Double
        Dim iCustodianLocation As Integer
        Dim iCounter As Integer
        Dim iNumberOfCustodianPSTs As Integer
        Dim sProcessingFilesDir As String
        Dim sCaseDirectory As String
        Dim sReprocessingDirectory As String
        Dim sOutputDirectory As String
        Dim sLogDirectory As String
        Dim iRowIndex As Integer
        Dim CustodianRow As DataGridViewRow

        Dim lstCustodianName As List(Of String)
        Dim lstCustodianPSTs As List(Of String)
        Dim lstPSTSize As List(Of Double)
        Dim lstPSTPath As List(Of String)
        Dim common As New Common
        Dim dbService As New DatabaseService

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"



        lstCustodianName = New List(Of String)
        lstCustodianPSTs = New List(Of String)
        lstPSTSize = New List(Of Double)
        lstPSTPath = New List(Of String)

        blnLoadPSTInfoFromCSV = False
        CSVMappingFile = New FileIO.TextFieldParser(sPSTCSVName)
        CSVMappingFile.TextFieldType = FileIO.FieldType.Delimited
        CSVMappingFile.SetDelimiters(",")

        Do While Not CSVMappingFile.EndOfData
            asCurrentRow = CSVMappingFile.ReadFields
            If (asCurrentRow(0) <> "Custodian Name") Then
                sCustodian = asCurrentRow(0)
                If lstCustodianName.Contains(sCustodian) Then
                    iCustodianLocation = lstCustodianName.IndexOf(sCustodian)
                    sCurrentPSTNames = lstCustodianPSTs(iCustodianLocation)
                    lstCustodianPSTs(iCustodianLocation) = sCurrentPSTNames & "," & asCurrentRow(1)
                    dblCurrentPSTSize = lstPSTSize(iCustodianLocation)
                    lstPSTSize(iCustodianLocation) = dblCurrentPSTSize + CInt(asCurrentRow(2))
                    If Not LCase(lstPSTPath(iCustodianLocation)).ToString.Contains(LCase(lstCustodianName(iCustodianLocation))) Then
                        lstNotConsolidatedCustodian.Add(asCurrentRow(0))
                    End If
                    If Not lstPSTPath.Contains(asCurrentRow(3)) Then
                        lstNotConsolidatedCustodian.Add(asCurrentRow(0))
                    End If
                Else
                    lstCustodianName.Add(asCurrentRow(0).ToString.ToLower)
                    lstCustodianPSTs.Add(asCurrentRow(1).ToString.ToLower)
                    lstPSTSize.Add(CInt(asCurrentRow(2)))
                    lstPSTPath.Add(asCurrentRow(3).ToString.ToLower)
                End If

            End If
        Loop

        iCounter = 0
        For Each Custodian In lstCustodianName
            sCustodianPSTs = lstCustodianPSTs(iCounter)
            asCustodianPSTs = Split(sCustodianPSTs, ",")
            iNumberOfCustodianPSTs = UBound(asCustodianPSTs)
            iRowIndex = grdPSTInfo.Rows.Add()
            sProcessingFilesDir = eMailArchiveMigrationManager.NuixFilesDir & "\EWS Ingestion\Nuix Files\" & Custodian.ToString
            sCaseDirectory = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Ingestion\" & Custodian.ToString
            sLogDirectory = eMailArchiveMigrationManager.NuixLogDir & "\EWS Ingestion\" & Custodian.ToString
            sOutputDirectory = eMailArchiveMigrationManager.NuixExportDir & "\EWS Ingestion\" & Custodian.ToString
            sReprocessingDirectory = eMailArchiveMigrationManager.NuixExportDir & "\EWS Ingestion\" & Custodian.ToString & "\Reprocessing"

            CustodianRow = grdPSTInfo.Rows(iRowIndex)

            With CustodianRow
                CustodianRow.Cells("StatusImage").Value = My.Resources.not_selected_small
                CustodianRow.Cells("SelectCustodian").Value = False
                CustodianRow.Cells("CustodianName").Value = Custodian.ToString
                CustodianRow.Cells("PSTPath").Value = lstPSTPath(iCounter)
                CustodianRow.Cells("NumberOfPSTs").Value = iNumberOfCustodianPSTs + 1
                CustodianRow.Cells("SizeOfPSTs").Value = lstPSTSize(iCounter).ToString("N0").Replace(".00", "")
                CustodianRow.Cells("ProcessingFilesDirectory").Value = sProcessingFilesDir
                CustodianRow.Cells("CaseDirectory").Value = sCaseDirectory
                CustodianRow.Cells("OutputDirectory").Value = sOutputDirectory
                CustodianRow.Cells("ReprocessingDirectory").Value = sReprocessingDirectory
                CustodianRow.Cells("LogDirectory").Value = sLogDirectory
                CustodianRow.Cells("ProgressStatus").Value = "Staged"
            End With
            common.Logger(psIngestionLogFile, "Updating SQLiteDB with - " & Custodian.ToString & "," & lstPSTPath(iCounter) & "," & iNumberOfCustodianPSTs + 1 & "," & lstPSTSize(iCounter) & ",,,,False,False,Nothing,Nothing,0.0,0,0.0,0.0," & sProcessingFilesDir & "," & sCaseDirectory & "," & sOutputDirectory & "," & sLogDirectory & ",")
            'bStatus = common.blnUpdateSQLiteAllCustodiansInfo(Custodian.ToString, "Staged", lstPSTPath(iCounter), iNumberOfCustodianPSTs + 1, lstPSTSize(iCounter), "", "", "", "", False, False, False, "", Nothing, Nothing, 0.0, 0, 0.0, 0.0, sProcessingFilesDir, sCaseDirectory, sOutputDirectory, sLogDirectory, "")
            bStatus = dbService.UpdateSQLiteEWSDB(sSQLiteDatabaseFullName, Custodian.ToString, "Staged", lstPSTPath(iCounter), iNumberOfCustodianPSTs + 1, lstPSTSize(iCounter), "", "", "", "", "", Nothing, Nothing, 0.0, 0.0, 0.0, sProcessingFilesDir, sCaseDirectory, sOutputDirectory, sLogDirectory, "")
            iCounter = iCounter + 1
        Next
        lstNotConsolidatedCustodian = lstNotConsolidatedCustodian.Distinct.ToList
        blnLoadPSTInfoFromCSV = True
    End Function
    Public Function blnLoadPSTInfoFromDirectory(ByVal sPSTFolderName As String, ByVal grdPSTInfo As DataGridView, ByRef lstNotConsolidatedPST As List(Of String)) As Boolean
        Dim bStatus As Boolean
        Dim ReportOutputFile As StreamWriter
        Dim sMachineName As String
        Dim sOutputFileName As String

        sMachineName = System.Net.Dns.GetHostName()
        sOutputFileName = sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv"

        ReportOutputFile = New StreamWriter(sPSTFolderName & "\" & sOutputFileName)
        ReportOutputFile.WriteLine("Custodian Name" & "," & "PST Name" & "," & "PST Size" & "," & "Original Folder")

        blnLoadPSTInfoFromDirectory = False
        bStatus = blnBuildPSTReports(sPSTFolderName, ReportOutputFile)

        bStatus = blnLoadPSTInfoFromCSV(sPSTFolderName & "\" & sOutputFileName, grdPSTInfo, True, lstNotConsolidatedPST)

        blnLoadPSTInfoFromDirectory = True

    End Function

    Private Sub OnCreated(sender As Object, e As IO.FileSystemEventArgs)
        Dim sCustodianName As String
        Dim lstCustodianName As List(Of String)
        Dim lstCustodianItemsExported As List(Of String)
        Dim lstCustodianItemsFailed As List(Of String)
        Dim lstCustodianBytesUploaded As List(Of String)
        Dim sInvalidFolder As String
        Dim NuixConsoleProcess As Process
        Dim NuixConsoleProcessStartInfo As ProcessStartInfo
        Dim sReprocessingFilesDirectory As String

        Dim sCompletedFolder As String
        Dim sCurrentFolder As String
        Dim sInProgressFolder As String

        Dim sItemsExported As String
        Dim sItemsFailed As String
        Dim sBytesUploaded As String

        Dim iNumberItemsFailed As Double
        Dim iNumberItemsExported As Double
        Dim dPercentFailed As Decimal
        Dim sProcessID As String
        Dim bStatus As Boolean
        Dim sAllFailuresLog As String
        Dim sOtherFailuresLog As String
        Dim sFTSFailuresLog As String

        Dim iNumberOfNuixInstancesRunning As Integer
        Dim iNuixInstancesRequested As Integer

        Dim lstProcessesNotRunning As List(Of String)

        Dim dEndTime As Date
        Dim dEndDate As Date
        Dim dStartDate As DateTime = DateTime.Now
        Dim dbService As New DatabaseService
        Dim common As New Common

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"


        lstCustodianName = New List(Of String)
        lstCustodianItemsExported = New List(Of String)
        lstCustodianItemsFailed = New List(Of String)
        lstCustodianBytesUploaded = New List(Of String)
        lstProcessesNotRunning = New List(Of String)

        Try
            Threading.Thread.Sleep(2000)
            common.Logger(psIngestionLogFile, "Loading Data from - " & e.FullPath)
            bStatus = blnGetSummaryReportInfo(e.FullPath, sCustodianName, sItemsExported, sItemsFailed, sBytesUploaded, dStartDate, dEndDate)
            iNuixInstancesRequested = Convert.ToInt16(eMailArchiveMigrationManager.O365NuixInstances)

            For Each Row In grdPSTInfo.Rows
                If Row.Cells("CustodianName").value = sCustodianName Then
                    sProcessID = Row.Cells("ProcessID").value
                    dEndTime = Now
                    iNumberItemsFailed = CInt(sItemsFailed)
                    iNumberItemsExported = CInt(sItemsExported)
                    If (iNumberItemsExported > 0) Then
                        dPercentFailed = (iNumberItemsFailed / iNumberItemsExported)
                    Else
                        dPercentFailed = 1.0
                    End If
                    bStatus = dbService.GetCustodianPSTLocation(sSQLiteDatabaseFullName, Row.Cells("CustodianName").value, sCurrentFolder)
                    sCompletedFolder = sCurrentFolder.Replace("In Progress", "Completed")
                    sCompletedFolder = sCurrentFolder.Replace("in progress", "completed")
                    common.Logger(psIngestionLogFile, "Moving - " & sCurrentFolder & " to " & sCompletedFolder)
                    bStatus = blnMoveCustodianFolder(sCustodianName, sCurrentFolder, sCompletedFolder, sInvalidFolder)
                    Row.cells("PSTPath").value = sCompletedFolder
                    Row.cells("ProgressStatus").value = "Completed"
                    Row.cells("SelectCustodian").readonly = True
                    Row.cells("ProgressStatus").readonly = True
                    Row.cells("ProcessID").value = vbNullString
                    If (dPercentFailed > 0.25) Then
                        Row.DefaultCellStyle.Forecolor = Color.Red
                        Row.cells("StatusImage").value = My.Resources.red_stop_small
                    ElseIf (dPercentFailed > 0.1) Then
                        Row.DefaultCellStyle.Forecolor = Color.Orange
                        Row.cells("StatusImage").value = My.Resources.yellow_info_small
                    Else
                        Row.defaultcellstyle.Forecolor = Color.Green
                        Row.cells("StatusImage").value = My.Resources.Green_check_small
                    End If
                    Row.Cells("IngestionEndTime").value = dEndTime
                    dbService.UpdateCustodianIngestionValues(psSQLiteLocation, sCustodianName, "IngestionEndTime", dEndTime.ToString)
                    Row.cells("BytesUploaded").value = FormatNumber(Convert.ToDouble(sBytesUploaded), , , , TriState.True)
                    Row.Cells("NumberSuccess").value = sItemsExported
                    Row.Cells("NumberFailed").value = sItemsFailed
                    dbService.UpdateCustodianIngestionValues(psSQLiteLocation, sCustodianName, "Success", sItemsExported)
                    dbService.UpdateCustodianIngestionValues(psSQLiteLocation, sCustodianName, "Failed", sItemsFailed)
                    sReprocessingFilesDirectory = Row.cells("ProcessingFilesDirectory").value & "\Reprocessing\"
                    Directory.CreateDirectory(sReprocessingFilesDirectory)
                    Row.cells("SummaryReportLocation").value = e.FullPath
                    If Row.cells("NumberFailed").value > 0 Then

                        sAllFailuresLog = Row.Cells("ProcessingFilesDirectory").value & "\Reprocessing\AllFailures.log"
                        sFTSFailuresLog = Row.Cells("ProcessingFilesDirectory").value & "\Reprocessing\FTSFailures.log"
                        sOtherFailuresLog = Row.Cells("ProcessingFilesDirectory").value & "\Reprocessing\OtherFailures.log"

                        bStatus = blnConsolidateFailureLogsAndFiles(Row.cells("OutputDirectory").value, Row.cells("ReprocessingDirectory").value, Row.Cells("ProcessingFilesDirectory").value & "Reprocessing", sAllFailuresLog, sFTSFailuresLog, sOtherFailuresLog)
                    End If

                    bStatus = blnUpdateSQLiteDBCustodianInfo(sCustodianName, Row.cells("PSTPath").value, Row.cells("ProcessID").value, Row.cells("IngestionEndTime").value, Row.cells("NumberSuccess").value.ToString.Replace(",", ""), Row.cells("NumberFailed").value.ToString.Replace(",", ""), Row.cells("SummaryReportLocation").value)
                End If
            Next

            plstNuixConsoleProcesses.RemoveAt(plstNuixConsoleProcesses.IndexOf(sProcessID))
            common.Logger(psIngestionLogFile, "Checking Number of Processes Running ")

            For Each ConsoleProcess In plstNuixConsoleProcesses
                bStatus = blnCheckIfProcessIsRunning(ConsoleProcess.ToString)
                If bStatus = False Then
                    lstProcessesNotRunning.Add(ConsoleProcess.ToString)
                    For Each custodianrow In grdPSTInfo.Rows
                        If custodianrow.cells("ProcessID").value = ConsoleProcess.ToString Then
                            If (File.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\" & custodianrow.cells("CustodianName").value & "\summary-report.txt")) Then
                                bStatus = blnGetSummaryReportInfo(eMailArchiveMigrationManager.NuixCaseDir & "\" & custodianrow.cells("CustodianName").value & "\summary-report.txt", sCustodianName, sItemsExported, sItemsFailed, sBytesUploaded, dStartDate, dEndDate)
                                dEndTime = Now
                                iNumberItemsFailed = CInt(sItemsFailed)
                                iNumberItemsExported = CInt(sItemsExported)
                                If (iNumberItemsExported > 0) Then
                                    dPercentFailed = (iNumberItemsFailed / iNumberItemsExported)
                                Else
                                    dPercentFailed = 1.0
                                End If
                                custodianrow.cells("ProcessID").value = vbNullString
                                custodianrow.cells("StatusImage").value = My.Resources.yellow_info_small
                                custodianrow.cells("ProgressStatus").value = "Completed"
                                bStatus = dbService.UpdateCustodianIngestionValues(eMailArchiveMigrationManager.SQLiteDBLocation, custodianrow.Cells("CustodianName").Value, "ProgressStatus", custodianrow.Cells("ProgressStatus").Value)
                            Else
                                custodianrow.cells("ProcessID").value = vbNullString
                                custodianrow.cells("StatusImage").value = My.Resources.yellow_info_small
                                custodianrow.cells("ProgressStatus").value = "Process Terminated by User"
                                bStatus = dbService.UpdateCustodianIngestionValues(eMailArchiveMigrationManager.SQLiteDBLocation, custodianrow.Cells("CustodianName").Value, "ProgressStatus", custodianrow.Cells("ProgressStatus").Value)
                            End If

                        End If
                    Next
                End If
            Next

            For Each NotRunningProcess In lstProcessesNotRunning
                plstNuixConsoleProcesses.RemoveAt(plstNuixConsoleProcesses.IndexOf(NotRunningProcess.ToString))
            Next

            iNumberOfNuixInstancesRunning = plstNuixConsoleProcesses.Count

            common.Logger(psIngestionLogFile, "Summary Report Creation Instances Running = " & plstNuixConsoleProcesses.Count)
            For Each row In grdPSTInfo.Rows
                If iNumberOfNuixInstancesRunning < iNuixInstancesRequested Then
                    If row.Cells("ProgressStatus").value = "Not Started" Then
                        dStartDate = Now
                        sCustodianName = row.Cells("CustodianName").Value
                        bStatus = dbService.GetCustodianPSTLocation(sSQLiteDatabaseFullName, sCustodianName, sCurrentFolder)
                        sInProgressFolder = sCurrentFolder.Replace("Not Started", "In Progress")
                        sInProgressFolder = sCurrentFolder.Replace("not started", "in progress")
                        common.Logger(psIngestionLogFile, "Moving - " & sCustodianName & " from " & sCurrentFolder & " to " & sInProgressFolder)
                        bStatus = blnMoveCustodianFolder(sCustodianName, sCurrentFolder, sInProgressFolder, sInvalidFolder)
                        common.Logger(psIngestionLogFile, "Launching Nuix...")
                        NuixConsoleProcessStartInfo = New ProcessStartInfo(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Ingestion\Nuix Files\" & sCustodianName & "\" & sCustodianName & ".bat")
                        NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        'NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Minimized

                        NuixConsoleProcess = System.Diagnostics.Process.Start(NuixConsoleProcessStartInfo)
                        row.Cells("StatusImage").Value = My.Resources.inprogress_medium
                        row.Cells("ProgressStatus").Value = "In Progress"
                        row.Cells("ProcessID").Value = NuixConsoleProcess.Id
                        row.Cells("IngestionStartTime").Value = dStartDate
                        row.Cells("PSTPath").Value = sInProgressFolder
                        common.Logger(psIngestionLogFile, "Updating Status for " & row.Cells("CustodianName").Value)
                        bStatus = dbService.UpdateCustodianIngestionValues(eMailArchiveMigrationManager.SQLiteDBLocation, row.Cells("CustodianName").Value, "ProgressStatus", row.Cells("ProgressStatus").Value)
                        common.Logger(psIngestionLogFile, "Updating Start Time for " & row.Cells("CustodianName").Value & " to " & row.Cells("IngestionStartTime").Value)
                        bStatus = dbService.UpdateCustodianDBStartTime(sSQLiteDatabaseFullName, row.Cells("CustodianName").Value, row.Cells("IngestionStartTime").Value)
                        plstNuixConsoleProcesses.Add(NuixConsoleProcess.Id.ToString)
                        iNumberOfNuixInstancesRunning = plstNuixConsoleProcesses.Count
                    End If
                End If
            Next

        Catch ex As Exception
            Dim Result As String
            Dim hr As Integer = Runtime.InteropServices.Marshal.GetHRForException(ex)
            Result = ex.GetType.ToString & "(0x" & hr.ToString("X8") & "): " & ex.Message & Environment.NewLine & ex.StackTrace & Environment.NewLine
            Dim st As StackTrace = New StackTrace(ex, True)
            For Each sf As StackFrame In st.GetFrames
                If sf.GetFileLineNumber() > 0 Then
                    Result &= "Line:" & sf.GetFileLineNumber() & " Filename: " & IO.Path.GetFileName(sf.GetFileName) & Environment.NewLine
                End If
            Next
            common.Logger(psIngestionLogFile, "On SummaryReport Create: " & Result.ToString)
        End Try

    End Sub

    Private Function blnConsolidateFailureLogsAndFiles(ByVal sExportDirectory As String, ByVal sReprocessingFilesDirectory As String, ByVal sReprocessingProcessingFilesDirectory As String, ByRef sAllFailuresLog As String, ByRef sFTSFailuresLog As String, ByRef sOtherFailuresLog As String) As Boolean

        Dim bStatus As Boolean
        Dim lstFailureLogs As List(Of String)
        lstFailureLogs = New List(Of String)
        Dim AllFailuresLog As StreamWriter
        Dim FTSFailuresLog As StreamWriter
        Dim OtherFailuresLog As StreamWriter
        Dim FailuresFiles As StreamReader
        Dim sCurrentRow As String
        Dim asFailuresInfo() As String
        Dim sMessageName As String
        Dim sFTSDirectory As String
        Dim sOtherDirectory As String

        AllFailuresLog = New StreamWriter(sAllFailuresLog)
        FTSFailuresLog = New StreamWriter(sFTSFailuresLog)
        OtherFailuresLog = New StreamWriter(sOtherFailuresLog)

        sFTSDirectory = sReprocessingFilesDirectory & "\FTS"
        sOtherDirectory = sReprocessingFilesDirectory & "\Other"
        Directory.CreateDirectory(sReprocessingFilesDirectory & "\FTS")
        Directory.CreateDirectory(sReprocessingFilesDirectory & "\Other")

        bStatus = blnGetAllFailureLogs(sExportDirectory, lstFailureLogs)
        For Each FailureLog In lstFailureLogs
            FailuresFiles = New StreamReader(FailureLog.ToString)
            While Not FailuresFiles.EndOfStream

                sCurrentRow = FailuresFiles.ReadLine
                If sCurrentRow <> vbNullString Then
                    asFailuresInfo = Split(sCurrentRow, ",")
                    sMessageName = asFailuresInfo(0)
                    'subFindMSGFile(sExportDirectory, sMessageName, sMessageFullName)
                    AllFailuresLog.WriteLine(sCurrentRow)
                    If (sCurrentRow.Contains("FTS data is too large")) Then
                        FTSFailuresLog.WriteLine(sCurrentRow)
                        'System.IO.File.Move(sMessageFullName, sFTSDirectory & "\" & sMessageName)
                    Else
                        OtherFailuresLog.WriteLine(sCurrentRow)
                        'System.IO.File.Move(sMessageFullName.ToString, sOtherDirectory & "\" & sMessageName)
                    End If
                End If
            End While
        Next
        AllFailuresLog.Close()
        FTSFailuresLog.Close()
        OtherFailuresLog.Close()
        blnConsolidateFailureLogsAndFiles = True
    End Function

    Sub subFindMSGFile(ByVal sDir As String, ByRef sMSGFileName As String, ByRef sMSGFullFileName As String)
        Dim d As String
        Dim Length As String
        Dim currentdirectory As DirectoryInfo
        Dim common As New Common

        Try
            currentdirectory = New DirectoryInfo(sDir)

            For Each File In currentdirectory.GetFiles
                If File.Name = sMSGFileName Then
                    sMSGFullFileName = File.FullName

                End If
            Next

            Length = Directory.GetDirectories(sDir).Length
            If Length > 0 Then
                For Each d In Directory.GetDirectories(sDir)
                    subFindMSGFile(d, sMSGFileName, sMSGFullFileName)
                Next
            End If
        Catch ex As Exception
            common.Logger(psIngestionLogFile, "Error 2 - Directory Search" & ex.ToString)
        End Try

    End Sub

    Private Function blnCheckIfProcessIsRunning(ByVal sProcessID As String) As Boolean

        Dim NuixProcess As System.Diagnostics.Process

        blnCheckIfProcessIsRunning = False

        Try
            NuixProcess = Process.GetProcessById(CInt(sProcessID))
            blnCheckIfProcessIsRunning = True
        Catch ex As Exception
            blnCheckIfProcessIsRunning = False
        End Try

    End Function

    Private Sub CaseLock_Deleted(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs)
        Dim iNumberOfNuixInstancesRunning As Integer
        Dim bStatus As Boolean

        Dim common As New Common

        bStatus = blnCheckIfNuixIsRunning("nuix_console", iNumberOfNuixInstancesRunning)
        bStatus = blnCheckIfNuixIsRunning("nuix_app", iNumberOfNuixInstancesRunning)
        common.Logger(psIngestionLogFile, "CaseLock deleted Instances Running = " & iNumberOfNuixInstancesRunning)

    End Sub
    Public Sub WatchCaseLockFile()

        CaseLockWatcher.Path = psNuixCaseDir

        AddHandler CaseLockWatcher.Deleted, AddressOf CaseLock_Deleted

    End Sub

    Private Sub btnProcessSelectedUsers_Click(sender As Object, e As EventArgs) Handles btnProcessandRunSelectedUsers.Click
        Dim bStatus As Boolean
        Dim lstProcessingCustodians As List(Of String)
        Dim lstRestartProcessingCustodian As List(Of String)
        Dim lstRestartSelectedCustodianInfo As List(Of String)
        Dim lstInvalidCustodian As List(Of String)
        Dim sPSTDirectory As String
        Dim sNuixAppMemory As String
        Dim iNumberofNuixInstances As Integer
        Dim lstCurrentCustodianFolder As List(Of String)
        Dim bFoundSelectedCustodian As Boolean
        Dim msgboxReturn As DialogResult

        Dim common As New Common

        btnLoadPreviousConfig.Enabled = False
        btnProcessandRunSelectedUsers.Enabled = False
        btnExportNotUploaded.Enabled = False

        Try
            If psSettingsFile = vbNullString Then
                msgboxReturn = MessageBox.Show("The settings have not been loaded. Would you like to load the setting dialog.", "Settings not loaded!", MessageBoxButtons.YesNo)
                If msgboxReturn = vbYes Then
                    Dim SettingsForm As O365ExtractionSettings
                    SettingsForm = New O365ExtractionSettings
                    SettingsForm.ShowDialog()
                    Cursor.Current = Cursors.Default
                    Exit Sub
                Else
                    Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            End If

            plstNuixConsoleProcesses = New List(Of String)
            lstProcessingCustodians = New List(Of String)
            lstRestartSelectedCustodianInfo = New List(Of String)
            lstRestartProcessingCustodian = New List(Of String)
            lstInvalidCustodian = New List(Of String)
            lstCurrentCustodianFolder = New List(Of String)

            sPSTDirectory = txtPSTLocation.Text

            bFoundSelectedCustodian = False

            For Each Row In grdPSTInfo.Rows
                If Row.cells("SelectCustodian").value = True Then
                    If (Row.cells("ProgressStatus").value = "Staged") Then
                        If ((Row.cells("DestinationFolder").value = vbNullString) Or (Row.cells("DestinationRoot").value = vbNullString) Or (Row.cells("DestinationSMTP").value = vbNullString)) Then
                            MessageBox.Show("You have not Added the appropriate Office 365 ingestion details.  Please update custodian mapping file and reload.")
                            btnProcessandRunSelectedUsers.Enabled = True
                            Exit Sub
                        Else
                            lstProcessingCustodians.Add(Row.cells("CustodianName").value)
                            Row.cells("ProgressStatus").value = "Not Started"
                        End If
                    End If
                End If
            Next

            If lstProcessingCustodians.Count <= 0 Then
                MessageBox.Show("You must select at least one custodian for EWS Ingestion.")
                btnProcessandRunSelectedUsers.Enabled = True
                Exit Sub
            End If

            sNuixAppMemory = CInt(eMailArchiveMigrationManager.O365NuixAppMemory) / 1000
            sNuixAppMemory = "-Xmx" & sNuixAppMemory & "g"

            grdPSTInfo.Sort(grdPSTInfo.Columns("SelectCustodian"), System.ComponentModel.ListSortDirection.Descending)

            bStatus = blnProcessSelectedCustodians(lstProcessingCustodians, grdPSTInfo, iNumberofNuixInstances, sPSTDirectory & "\ews ingestion\not started", sNuixAppMemory)

            SummaryReportThread = New System.Threading.Thread(AddressOf WatchCaseDirectory)
            SummaryReportThread.Start()

            SQLiteDBReadThread = New System.Threading.Thread(AddressOf SQLiteUpdates)
            SQLiteDBReadThread.Start()

            Cursor.Current = Cursors.Default

        Catch ex As Exception
            Dim Result As String
            Dim hr As Integer = Runtime.InteropServices.Marshal.GetHRForException(ex)
            Result = ex.GetType.ToString & "(0x" & hr.ToString("X8") & "): " & ex.Message & Environment.NewLine & ex.StackTrace & Environment.NewLine
            Dim st As StackTrace = New StackTrace(ex, True)
            For Each sf As StackFrame In st.GetFrames
                If sf.GetFileLineNumber() > 0 Then
                    Result &= "Line:" & sf.GetFileLineNumber() & " Filename: " & IO.Path.GetFileName(sf.GetFileName) & Environment.NewLine
                End If
            Next
            common.Logger(psIngestionLogFile, Result.ToString)
        End Try

        Cursor.Current = Cursors.WaitCursor
    End Sub
    Public Sub WatchCaseDirectory()
        ' set the path to be watched
        FileSystemWatcher1.Path = psNuixCaseDir

        ' add event handlers
        AddHandler FileSystemWatcher1.Created, AddressOf OnCreated
    End Sub

    Public Sub SQLiteUpdates()

        Dim lstBytesUploaded As List(Of String)
        Dim lstPercentCompleted As List(Of String)
        Dim lstProcessingCustodians As List(Of String)
        Dim lstNotStartedCustodians As List(Of String)
        Dim lstInProgressCustodians As List(Of String)
        Dim lstActiveCustodians As List(Of String)
        Dim lstCompletedCustodians As List(Of String)
        Dim lstCompletedBytesUploaded As List(Of String)
        Dim lstCompletedPercentCompleted As List(Of String)

        Dim bStatus As Boolean
        Dim bNoMoreJobs As Boolean
        Dim dPercentCompleted As Decimal
        Dim sCaseDirectory As String
        Dim bStillJobsRunning As Boolean
        Dim bUpdateReturn As Boolean
        Dim dbService As New DatabaseService

        Dim common As New Common

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"


        bNoMoreJobs = False
        Thread.Sleep(10000)
        lstBytesUploaded = New List(Of String)
        lstPercentCompleted = New List(Of String)
        lstProcessingCustodians = New List(Of String)
        lstCompletedPercentCompleted = New List(Of String)
        lstCompletedBytesUploaded = New List(Of String)
        lstActiveCustodians = New List(Of String)
        lstInProgressCustodians = New List(Of String)
        lstNotStartedCustodians = New List(Of String)
        lstCompletedCustodians = New List(Of String)
        bStillJobsRunning = True
        Try
            Do While bStillJobsRunning = True
                lstBytesUploaded.Clear()
                lstPercentCompleted.Clear()
                lstProcessingCustodians.Clear()

                bStatus = dbService.GetUpdatedEWSIngestionCustodianInfo(sSQLiteDatabaseFullName, lstProcessingCustodians, lstBytesUploaded, lstPercentCompleted, bUpdateReturn)
                If bUpdateReturn = False Then
                    bStillJobsRunning = False
                Else
                    bStillJobsRunning = True
                End If
                If lstProcessingCustodians.Count > 0 Then
                    For Each Custodian In lstProcessingCustodians
                        For Each row In grdPSTInfo.Rows
                            If row.cells("CustodianName").value = Custodian.ToString Then
                                row.cells("BytesUploaded").value = FormatNumber(lstBytesUploaded(lstProcessingCustodians.IndexOf(Custodian.ToString)), 0, , , TriState.True)
                                dPercentCompleted = CInt(lstPercentCompleted(lstProcessingCustodians.IndexOf(Custodian.ToString)))
                                dPercentCompleted = dPercentCompleted
                                row.cells("PercentageCompleted").value = dPercentCompleted.ToString.Replace(".00", "")
                            End If
                        Next
                    Next
                Else
                    'MessageBox.Show(lstProcessingCustodians.Count & "-" & bStillJobsRunning.ToString, "Processing Custodians - Jobs Still Running")
                End If


                bStatus = common.GetCompletedCustodianInfo(lstCompletedCustodians, lstCompletedBytesUploaded, lstCompletedPercentCompleted)
                If lstCompletedCustodians.Count > 0 Then
                    For Each custodian In lstCompletedCustodians
                        For Each row In grdPSTInfo.Rows
                            If row.cells("CustodianName").value = custodian.ToString Then
                                row.cells("BytesUploaded").value = FormatNumber(lstCompletedBytesUploaded(lstCompletedCustodians.IndexOf(custodian.ToString)), 0, , , TriState.True)
                                dPercentCompleted = CInt(lstCompletedPercentCompleted(lstCompletedCustodians.IndexOf(custodian.ToString)))
                                dPercentCompleted = dPercentCompleted / 100
                                row.cells("PercentageCompleted").value = "100"
                                bStatus = dbService.GetUpdatedIngestionDBInfo(psSQLiteLocation, custodian.ToString, "CaseDirectory", sCaseDirectory, "TEXT")
                                bStatus = blnCheckIfProcessIsRunning(row.cells("ProcessID").value)
                                If bStatus = False Then
                                    bStatus = dbService.UpdateCustodianIngestionValues(psSQLiteLocation, custodian.ToString, "ProcessID", "")
                                    row.cells("ProcessID").value = ""
                                End If
                                If File.Exists(sCaseDirectory & "\summary-report.txt") Then
                                    row.cells("SummaryReportLocation").value = sCaseDirectory & "\summary-report.txt"
                                    row.cells("ProcessID").value = ""
                                    bStatus = dbService.UpdateCustodianIngestionValues(psSQLiteLocation, custodian.ToString, "SummaryReportLocation", sCaseDirectory & "\summary-report.txt")
                                End If
                            End If
                        Next
                    Next
                End If

                Thread.Sleep(10000)
            Loop

        Catch ex As Exception
            Dim Result As String
            Dim hr As Integer = Runtime.InteropServices.Marshal.GetHRForException(ex)
            Result = ex.GetType.ToString & "(0x" & hr.ToString("X8") & "): " & ex.Message & Environment.NewLine & ex.StackTrace & Environment.NewLine
            Dim st As StackTrace = New StackTrace(ex, True)
            For Each sf As StackFrame In st.GetFrames
                If sf.GetFileLineNumber() > 0 Then
                    Result &= "Line:" & sf.GetFileLineNumber() & " Filename: " & IO.Path.GetFileName(sf.GetFileName) & Environment.NewLine
                End If
            Next
            common.Logger(psIngestionLogFile, "SQLiteUpdates: " & Result.ToString)
        End Try
        '        MessageBox.Show(lstProcessingCustodians.Count & "-" & bStillJobsRunning.ToString, "Processing Custodians - Jobs Still Running")

        bStatus = blnRefreshGridFromDB(grdPSTInfo, lstNotStartedCustodians, lstInProgressCustodians)
        common.Logger(psIngestionLogFile, "All Currently select Processing Jobs have completed.  If necessary select more custodians for Processing.")
        MessageBox.Show("All Currently select Processing Jobs have completed.  If necessary select more custodians for Processing.", "All Processing Jobs Completed", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification, False)
        btnLoadPreviousConfig.Enabled = True
        btnProcessandRunSelectedUsers.Enabled = True
        btnExportNotUploaded.Enabled = True

    End Sub

    Private Function blnBuildNuixFilesForCustodians(ByVal sCustodianName As String, ByVal sNuixAppMemory As String, ByVal sWorkerMemory As String, ByVal sWorkerCount As String, ByRef sPSTDirectory As String, ByVal sCaseDirectory As String, ByVal sLogDir As String, ByVal sJavaTempDir As String, ByVal sTimeout As String, ByVal sRetries As String, ByVal sFailureIncrement As String, ByVal sRetryDelay As String, ByVal sRemovePathPrefix As String, bApplicationImpersonation As Boolean, ByVal sNuixAppLocation As String, ByVal sExchangeServer As String, ByVal sDomain As String, ByVal sAdminUserName As String, ByVal sAdminInfo As String, ByVal sNMSLocation As String, ByVal sWorkerTempDir As String, ByVal sExportDir As String, ByRef lstCurrentCustodianFolder As List(Of String), ByVal sNMSUsername As String, ByVal sNMSPassword As String) As Boolean
        Dim bStatus As Boolean
        Dim sRubyFileName As String
        Dim sMappingFileName As String
        Dim sRubyFilePSTLocation As String
        Dim sNewPSTLocation As String
        Dim sJSonFileName As String
        Dim sNuixCase As String
        Dim sCustodianPSTPath As String
        Dim dblPSTSourceSize As Double
        Dim sDestinationFolder As String
        Dim sDestinationRoot As String
        Dim sDestinationSMTPAddress As String
        Dim sLogDirectory As String
        Dim sProcessingFilesDirectory As String
        Dim sOutputDirectory As String
        Dim sCustodianPSTDirectory As String
        Dim common As New Common
        Dim dbService As New DatabaseService

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"


        blnBuildNuixFilesForCustodians = False

        My.Computer.FileSystem.CreateDirectory(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Ingestion\Nuix Files")

        bStatus = dbService.GetCustodianPSTLocation(sSQLiteDatabaseFullName, sCustodianName, sCustodianPSTDirectory)
        sPSTDirectory = sPSTDirectory.ToLower
        sNuixCase = eMailArchiveMigrationManager.NuixCaseDir & "\ews ingestion\" & sCustodianName
        'If Not (sPSTDirectory.Contains("ews ingestion")) Then
        'sPSTDirectory = sPSTDirectory.Replace(sCustodianName.ToString.ToLower, "")
        'sPSTDirectory = sPSTDirectory.Replace("\\", "\")
        'My.Computer.FileSystem.CreateDirectory(sPSTDirectory & "\ews ingestion\not started\" & sCustodianName)
        'End If
        sPSTDirectory = sPSTDirectory & "\" & sCustodianName
        sPSTDirectory = sPSTDirectory.Replace("\\", "\")

        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "ProcessingFilesDirectory", sProcessingFilesDirectory, "TEXT")
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "OutputDirectory", sOutputDirectory, "TEXT")
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "LogDirectory", sLogDirectory, "TEXT")

        My.Computer.FileSystem.CreateDirectory(sProcessingFilesDirectory)
        sMappingFileName = sProcessingFilesDirectory & "\" & sCustodianName & "_mapping.csv"
        sRubyFileName = sProcessingFilesDirectory & "\" & sCustodianName & ".rb"
        sJSonFileName = sProcessingFilesDirectory & "\" & sCustodianName & ".json"

        bStatus = BuildIngestionBatchFiles(sCustodianName, sProcessingFilesDirectory & "\", sRubyFileName, sMappingFileName, sNuixAppMemory, sLogDirectory, sOutputDirectory)
        sNewPSTLocation = sPSTDirectory.Substring(0, sPSTDirectory.LastIndexOf("\"))
        sRubyFilePSTLocation = sPSTDirectory.Replace("not started", "in progress")

        If sRubyFilePSTLocation.Contains("\\") Then
            sRubyFilePSTLocation.Replace("\\", "\")
        End If

        bStatus = blnBuildUpdatedRubyScript(sNuixCase, sCustodianName, sRubyFileName, sRubyFilePSTLocation, sJSonFileName)
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "DestinationFolder", sDestinationFolder, "TEXT")
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "DestinationRoot", sDestinationRoot, "TEXT")
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "DestinationSMTP", sDestinationSMTPAddress, "TEXT")
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "TotalSizeOfPSTs", dblPSTSourceSize, "INT")
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "PSTPath", sCustodianPSTPath, "TEXT")

        bStatus = blnBuildJSonFile(sCustodianName, sJSonFileName, sRubyFilePSTLocation, eMailArchiveMigrationManager.O365MemoryPerWorker, eMailArchiveMigrationManager.O365NuixInstances, eMailArchiveMigrationManager.NuixCaseDir, sRubyFilePSTLocation.Replace("\", "\\"), sDestinationFolder, sDestinationRoot, sDestinationSMTPAddress, dblPSTSourceSize, sCustodianPSTPath.Replace("\", "\\"))
        bStatus = blnBuildIngestionCustodianMappingFile(sCustodianName, sMappingFileName, sDestinationFolder, sDestinationRoot, sDestinationSMTPAddress, dblPSTSourceSize, sCustodianPSTPath)

        lstCurrentCustodianFolder.Add(sPSTDirectory)

        blnBuildNuixFilesForCustodians = True
    End Function


    Public Function BuildIngestionBatchFiles(ByVal sCustodianName As String, ByVal sDirectoryName As String, ByVal sRubyFileName As String, ByVal sMappingFileName As String, ByVal sNuixAppMemory As String, ByVal sLogDirectory As String, ByVal sOutputDirectory As String) As Boolean

        BuildIngestionBatchFiles = False

        Dim CustodianBatchFile As StreamWriter

        Dim sStartUpBatchFileName As String
        Dim sLicenceSourceType As String

        sStartUpBatchFileName = sDirectoryName & "\" & sCustodianName & ".bat"
        sStartUpBatchFileName = sStartUpBatchFileName.Replace("\\", "\")

        'My.Computer.FileSystem.CreateDirectory(eMailArchiveMigrationManager.NuixLogDir & "\" & sCustodianName)

        If eMailArchiveMigrationManager.NMSSourceType = "Desktop" Then
            sLicenceSourceType = " -licencesourcetype dongle"
        ElseIf eMailArchiveMigrationManager.NMSSourceType = "Server" Then
            sLicenceSourceType = " -licencesourcetype server -licencesourcelocation " & eMailArchiveMigrationManager.NuixNMSAddress & ":" & eMailArchiveMigrationManager.NuixNMSPort & " -Dnuix.registry.servers=" & eMailArchiveMigrationManager.NuixNMSAddress
        ElseIf eMailArchiveMigrationManager.NMSSourceType = "Cloud Server" Then
            sLicenceSourceType = " -Dnuix.product.server.addClsProdEndpoint=true -licencesourcetype cloud-server"
        End If
        CustodianBatchFile = New StreamWriter(sStartUpBatchFileName)
        CustodianBatchFile.WriteLine("::TITLE is the destination SMTP Address")
        CustodianBatchFile.WriteLine("@TITLE " & sCustodianName)
        CustodianBatchFile.WriteLine("::Enter NMS Username on Line 4")
        CustodianBatchFile.WriteLine("@SET NUIX_USERNAME=" & eMailArchiveMigrationManager.NuixNMSUserName)
        CustodianBatchFile.WriteLine("::Enter NMS Username on Line 6")
        CustodianBatchFile.WriteLine("@SET NUIX_PASSWORD=" & eMailArchiveMigrationManager.NuixNMSAdminInfo)

        CustodianBatchFile.Write("""" & eMailArchiveMigrationManager.NuixAppLocation & """" & sLicenceSourceType)
        CustodianBatchFile.Write(" -licencetype email-archive-examiner -licenceworkers " & eMailArchiveMigrationManager.O365NumberOfNuixWorkers & " " & sNuixAppMemory & " -Dnuix.logdir=" & """" & sLogDirectory & """" & " -Djava.io.tmpdir=" & """" & eMailArchiveMigrationManager.NuixJavaTempDir & """" & " -Dnuix.worker.tmpdir=" & """" & eMailArchiveMigrationManager.NuixWorkerTempDir & """" & " -Dnuix.crackAndDump.exportDir=" & """" & sOutputDirectory & """")
        CustodianBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & "ews:configFile=>" & sMappingFileName & """")
        CustodianBatchFile.Write(" -Dnuix.processing.crackAndDump.useRelativePath=true -Dnuix.crackAndDump.fallbackExporterType=MSG -Dnuix.crackAndDump.compressFailures=false")
        CustodianBatchFile.Write(" -Dnuix.processing.worker.timeout=" & eMailArchiveMigrationManager.WorkerTimeout & " -Dnuix.processing.crackAndDump.prependEvidenceName=true -Dnuix.processing.crackAndDump.removePathPrefix=" & eMailArchiveMigrationManager.O365FilePathTrimming)
        CustodianBatchFile.Write(" -Dnuix.data.exchangews.retryCount=" & eMailArchiveMigrationManager.O365RetryCount & " -Dnuix.data.exchangews.retryDelay=" & eMailArchiveMigrationManager.O365RetryDelay)
        CustodianBatchFile.Write(" -Dnuix.data.exchangews.retryDelayIncrement=" & eMailArchiveMigrationManager.O365RetryIncrement & " -Dnuix.data.exchangews.impersonate=" & eMailArchiveMigrationManager.O365ApplicationImpersonation.ToString.ToLower)
        If (CBool(eMailArchiveMigrationManager.O365EnableBulkUpload) = True) Then
            CustodianBatchFile.Write(" -Dnuix.export.exchangews.bulkUpload=" & eMailArchiveMigrationManager.O365EnableBulkUpload)
            CustodianBatchFile.Write(" -Dnuix.export.exchangews.maxUploadSize=" & eMailArchiveMigrationManager.O365BulkUploadSize * 1000 & " ")
        Else
            CustodianBatchFile.Write(" -Dnuix.export.exchangews.messageSizeLimit=" & eMailArchiveMigrationManager.o365MaxMessageSize * 1000 & " ")
        End If
        CustodianBatchFile.WriteLine("""" & sRubyFileName & """")

        'CustodianBatchFile.WriteLine("@pause")
        CustodianBatchFile.WriteLine("@exit")

        CustodianBatchFile.Close()

        BuildIngestionBatchFiles = True
    End Function

    Private Function blnBuildReprocessingFilesForCustodians(ByVal sCustodianName As String, ByVal sExportDirectory As String, ByVal sTotalExportSize As String, ByVal sNuixAppMemory As String, ByVal sWorkerMemory As String, ByVal sWorkerCount As String, ByRef sPSTDirectory As String, ByVal sCaseDirectory As String, ByVal sLogDir As String, ByVal sJavaTempDir As String, ByVal sTimeout As String, ByVal sRetries As String, ByVal sFailureIncrement As String, ByVal sRetryDelay As String, ByVal sRemovePathPrefix As String, bApplicationImpersonation As Boolean, ByVal sNuixAppLocation As String, ByVal sExchangeServer As String, ByVal sDomain As String, ByVal sAdminUserName As String, ByVal sAdminInfo As String, ByVal sNMSLocation As String, ByVal sWorkerTempDir As String, ByVal sExportDir As String, ByRef lstCurrentCustodianFolder As List(Of String), ByVal sNMSUsername As String, ByVal sNMSPassword As String, ByVal sFailureLog As String) As Boolean
        Dim bStatus As Boolean
        Dim sRubyFileName As String
        Dim sMappingFileName As String
        Dim sNewPSTLocation As String
        Dim sJSonFileName As String
        Dim sNuixCase As String
        Dim sCustodianPSTPath As String
        Dim dblPSTSourceSize As Double
        Dim sDestinationFolder As String
        Dim sDestinationRoot As String
        Dim sDestinationSMTPAddress As String
        Dim sReprocessingLogDirectory As String
        Dim sReprocessingFilesDirectory As String
        Dim sReprocessingOutputDirectory As String
        Dim sReprocessingCaseDirectory As String
        Dim CustodianMappingStream As StreamWriter

        Dim dbService As New DatabaseService

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"

        blnBuildReprocessingFilesForCustodians = False

        My.Computer.FileSystem.CreateDirectory(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Ingestion\Nuix Files")

        bStatus = dbService.GetCustodianPSTLocation(sSQLiteDatabaseFullName, sCustodianName, sPSTDirectory)
        sPSTDirectory = sPSTDirectory.ToLower
        sNuixCase = eMailArchiveMigrationManager.NuixCaseDir & "\ews ingestion\" & sCustodianName
        If Not (sPSTDirectory.Contains("ews ingestion")) Then
            sPSTDirectory = sPSTDirectory.Replace(sCustodianName.ToString.ToLower, "")
            sPSTDirectory = sPSTDirectory.Replace("\\", "\")
            My.Computer.FileSystem.CreateDirectory(sPSTDirectory & "\ews ingestion\not started\" & sCustodianName)
        End If
        sPSTDirectory = sPSTDirectory & "\ews ingestion\not started\" & sCustodianName
        sPSTDirectory = sPSTDirectory.Replace("\\", "\")

        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "ProcessingFilesDirectory", sReprocessingFilesDirectory, "TEXT")
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "OutputDirectory", sReprocessingOutputDirectory, "TEXT")
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "LogDirectory", sReprocessingLogDirectory, "TEXT")
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "CaseDirectory", sReprocessingCaseDirectory, "TEXT")

        sReprocessingFilesDirectory = sReprocessingFilesDirectory & "\Reprocessing\"
        sReprocessingLogDirectory = sReprocessingLogDirectory & "\Reprocessing"
        sReprocessingCaseDirectory = sReprocessingCaseDirectory & "\Reprocessing"
        sReprocessingOutputDirectory = sReprocessingOutputDirectory & "\Reprocessing"


        My.Computer.FileSystem.CreateDirectory(sReprocessingFilesDirectory)
        My.Computer.FileSystem.CreateDirectory(sReprocessingLogDirectory)
        My.Computer.FileSystem.CreateDirectory(sReprocessingCaseDirectory)
        My.Computer.FileSystem.CreateDirectory(sReprocessingOutputDirectory)


        sMappingFileName = sReprocessingFilesDirectory & sCustodianName & "_mapping.csv"
        sRubyFileName = sReprocessingFilesDirectory & sCustodianName & ".rb"
        sJSonFileName = sReprocessingFilesDirectory & sCustodianName & ".json"
        CustodianMappingStream = New StreamWriter(sMappingFileName)
        CustodianMappingStream.WriteLine(sCustodianName & "," & sCustodianName & ".pst")
        CustodianMappingStream.Close()

        bStatus = blnBuildReprocessingBatchFiles(sCustodianName, sReprocessingLogDirectory, sReprocessingOutputDirectory, sReprocessingFilesDirectory, sRubyFileName, sMappingFileName, eMailArchiveMigrationManager.NuixAppMemory, sFailureLog)
        sNewPSTLocation = sPSTDirectory.Substring(0, sPSTDirectory.LastIndexOf("\"))

        bStatus = blnBuildEWSReprocessingRubyScript(sRubyFileName, psSQLiteLocation, sJSonFileName, sReprocessingCaseDirectory, sWorkerCount, sWorkerMemory, sExportDirectory, sWorkerTempDir, eMailArchiveMigrationManager.extractFromSlackSpace)
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "DestinationFolder", sDestinationFolder, "TEXT")
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "DestinationRoot", sDestinationRoot, "TEXT")
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "DestinationSMTP", sDestinationSMTPAddress, "TEXT")
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "TotalSizeOfPSTs", dblPSTSourceSize, "INT")
        bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "PSTPath", sCustodianPSTPath, "TEXT")

        'bStatus = blnBuildJSonFile(sCustodianName, sJSonFileName, "", eMailArchiveMigrationManager.O365MemoryPerWorker, 1, sReprocessingCaseDirectory, "", sDestinationFolder, sDestinationRoot, sDestinationSMTPAddress, sTotalExportSize, sExportDirectory)
        bStatus = blnBuildReprocessingJSonFile(sJSonFileName, sExportDirectory, "", "", sCustodianName, sTotalExportSize, "lightspeed")
        lstCurrentCustodianFolder.Add(sPSTDirectory & "\EWS Ingestion\Not Started\" & sCustodianName)

        blnBuildReprocessingFilesForCustodians = True
    End Function
    Private Function blnBuildNuixFilesForSelectedCustodians(ByVal lstSelectedCustodiansInfo As List(Of String), ByVal sNuixAppMemory As String, ByVal sWorkerMemory As String, ByVal sWorkerCount As String, ByRef sPSTDirectory As String, ByVal sCaseDirectory As String, ByVal sLogDir As String, ByVal sJavaTempDir As String, ByVal sTimeout As String, ByVal sRetries As String, ByVal sFailureIncrement As String, ByVal sRetryDelay As String, ByVal sRemovePathPrefix As String, bApplicationImpersonation As Boolean, ByVal sNuixAppLocation As String, ByVal sExchangeServer As String, ByVal sDomain As String, ByVal sAdminUserName As String, ByVal sAdminInfo As String, ByVal sNMSLocation As String, ByVal sWorkerTempDir As String, ByVal sExportDir As String, ByRef lstCurrentCustodianFolder As List(Of String), ByVal sNMSUsername As String, ByVal sNMSPassword As String) As Boolean
        Dim bStatus As Boolean
        Dim asCustodianInfo() As String
        Dim sRubyFileName As String
        Dim sMappingFileName As String
        Dim sRubyFilePSTLocation As String
        Dim sNewPSTLocation As String
        Dim sJSonFileName As String
        Dim sNuixCase As String
        Dim sCustodianPSTPath As String
        Dim dblPSTSourceSize As Double
        Dim sDestinationFolder As String
        Dim sDestinationRoot As String
        Dim sDestinationSMTPAddress As String
        Dim sLogDirectory As String
        Dim sProcessingFilesDirectory As String
        Dim sOutputDirectory As String

        Dim common As New Common
        Dim dbService As New DatabaseService

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"

        blnBuildNuixFilesForSelectedCustodians = False

        My.Computer.FileSystem.CreateDirectory(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Ingestion\Nuix Files")

        For Each Custodian In lstSelectedCustodiansInfo
            asCustodianInfo = Split(Custodian.ToString, ",")
            bStatus = dbService.GetCustodianPSTLocation(sSQLiteDatabaseFullName, asCustodianInfo(0), sPSTDirectory)
            sPSTDirectory = sPSTDirectory.ToLower
            sNuixCase = eMailArchiveMigrationManager.NuixCaseDir & "\ews ingestion\" & asCustodianInfo(0)
            If Not (sPSTDirectory.Contains("ews ingestion")) Then
                sPSTDirectory = sPSTDirectory.Replace(asCustodianInfo(0).ToString.ToLower, "")
                sPSTDirectory = sPSTDirectory.Replace("\\", "\")
                My.Computer.FileSystem.CreateDirectory(sPSTDirectory & "\ews ingestion\not started\" & asCustodianInfo(0))
            End If

            bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, asCustodianInfo(0), "ProcessingFilesDirectory", sProcessingFilesDirectory, "TEXT")
            bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, asCustodianInfo(0), "OutputDirectory", sOutputDirectory, "TEXT")
            bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, asCustodianInfo(0), "LogDirectory", sLogDirectory, "TEXT")

            My.Computer.FileSystem.CreateDirectory(sProcessingFilesDirectory)
            sMappingFileName = sProcessingFilesDirectory & "\" & asCustodianInfo(0) & "_mapping.csv"
            sRubyFileName = sProcessingFilesDirectory & "\" & asCustodianInfo(0) & ".rb"
            sJSonFileName = sProcessingFilesDirectory & "\" & asCustodianInfo(0) & ".json"

            bStatus = common.BuildIngestionBatchFiles(asCustodianInfo(0), sProcessingFilesDirectory & "\", sRubyFileName, sMappingFileName, sNuixAppMemory, sLogDirectory, sOutputDirectory)
            sNewPSTLocation = sPSTDirectory.Substring(0, sPSTDirectory.LastIndexOf("\"))
            sRubyFilePSTLocation = sPSTDirectory.Replace("not started", "in progress")

            If sRubyFilePSTLocation.Contains("\\") Then
                sRubyFilePSTLocation.Replace("\\", "\")
            End If

            bStatus = blnBuildUpdatedRubyScript(sNuixCase, asCustodianInfo(0), sRubyFileName, sRubyFilePSTLocation, sJSonFileName)
            bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, asCustodianInfo(0), "DestinationFolder", sDestinationFolder, "TEXT")
            bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, asCustodianInfo(0), "DestinationRoot", sDestinationRoot, "TEXT")
            bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, asCustodianInfo(0), "DestinationSMTP", sDestinationSMTPAddress, "TEXT")
            bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, asCustodianInfo(0), "TotalSizeOfPSTs", dblPSTSourceSize, "INT")
            bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, asCustodianInfo(0), "PSTPath", sCustodianPSTPath, "TEXT")

            bStatus = blnBuildJSonFile(asCustodianInfo(0), sJSonFileName, sRubyFilePSTLocation, eMailArchiveMigrationManager.O365MemoryPerWorker, eMailArchiveMigrationManager.O365NuixInstances, eMailArchiveMigrationManager.NuixCaseDir, sRubyFilePSTLocation.Replace("\", "\\"), sDestinationFolder, sDestinationRoot, sDestinationSMTPAddress, dblPSTSourceSize, sCustodianPSTPath.Replace("\", "\\"))
            bStatus = blnBuildIngestionCustodianMappingFile(asCustodianInfo(0), sMappingFileName, sDestinationFolder, sDestinationRoot, sDestinationSMTPAddress, dblPSTSourceSize, sCustodianPSTPath)

            lstCurrentCustodianFolder.Add(sPSTDirectory & "\EWS Ingestion\Not Started\" & asCustodianInfo(0))
            asCustodianInfo = Nothing
        Next

        blnBuildNuixFilesForSelectedCustodians = True
    End Function


    Public Function blnBuildReprocessingBatchFiles(ByVal sCustodianName As String, ByVal sLogDirectory As String, ByVal sReprocessingExportDirectory As String, ByVal sDirectoryName As String, ByVal sRubyFileName As String, ByVal sMappingFileName As String, ByVal sNuixAppMemory As String, ByVal sFailureLog As String) As Boolean

        blnBuildReprocessingBatchFiles = False

        Dim CustodianBatchFile As StreamWriter

        Dim sStartUpBatchFileName As String
        Dim sLicenceSourceType As String

        sStartUpBatchFileName = sDirectoryName & "\" & sCustodianName & ".bat"
        sStartUpBatchFileName = sStartUpBatchFileName.Replace("\\", "\")

        My.Computer.FileSystem.CreateDirectory(eMailArchiveMigrationManager.NuixLogDir & "\" & sCustodianName)

        If eMailArchiveMigrationManager.NMSSourceType = "Desktop" Then
            sLicenceSourceType = " -licencesourcetype dongle"
        ElseIf eMailArchiveMigrationManager.NMSSourceType = "Server" Then
            sLicenceSourceType = " -licencesourcetype server -licencesourcelocation " & eMailArchiveMigrationManager.NuixNMSAddress & ":" & eMailArchiveMigrationManager.NuixNMSPort & " -Dnuix.registry.servers=" & eMailArchiveMigrationManager.NuixNMSAddress
        ElseIf eMailArchiveMigrationManager.NMSSourceType = "Cloud Server" Then
            sLicenceSourceType = " -Dnuix.product.server.addClsProdEndpoint=true -licencesourcetype cloud-server"
        End If
        CustodianBatchFile = New StreamWriter(sStartUpBatchFileName)
        CustodianBatchFile.WriteLine("::Title will be the NSF Custodian")
        CustodianBatchFile.WriteLine("@TITLE " & sCustodianName)
        CustodianBatchFile.WriteLine("::Enter NMS Username on Line 4")
        CustodianBatchFile.WriteLine("@SET NUIX_USERNAME=" & eMailArchiveMigrationManager.NuixNMSUserName)
        CustodianBatchFile.WriteLine("::Enter NMS Username on Line 6")
        CustodianBatchFile.WriteLine("@SET NUIX_PASSWORD=" & eMailArchiveMigrationManager.NuixNMSAdminInfo)

        CustodianBatchFile.Write("""" & eMailArchiveMigrationManager.NuixAppLocation & """" & sLicenceSourceType)
        CustodianBatchFile.Write(" -licencetype email-archive-examiner -licenceworkers 1 -Xmx8g -Dnuix.logdir=" & """" & sLogDirectory & """" & " -Djava.io.tmpdir=" & """" & eMailArchiveMigrationManager.NuixJavaTempDir & """" & " -Dnuix.worker.tmpdir=" & """" & eMailArchiveMigrationManager.NuixWorkerTempDir & """" & " -Dnuix.crackAndDump.exportDir=" & """" & sReprocessingExportDirectory & """")
        CustodianBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & "pst:configFile=>" & sMappingFileName & """")
        CustodianBatchFile.Write(" -Dnuix.processing.crackAndDump.useRelativePath=true -Dnuix.processing.worker.timeout=" & eMailArchiveMigrationManager.WorkerTimeout & " -Dnuix.processing.crackAndDump.prependEvidenceName=true ")
        CustodianBatchFile.Write(" -Dnuix.export.mailbox.maximumFileSizePerMailbox=" & eMailArchiveMigrationManager.PSTExportSize & "GB -Dnuix.crackAndDump.sourceDataMapping=" & """" & sFailureLog & """" & " ")
        CustodianBatchFile.WriteLine("""" & sRubyFileName & """")

        'CustodianBatchFile.WriteLine("@pause")
        CustodianBatchFile.WriteLine("@exit")

        CustodianBatchFile.Close()

        blnBuildReprocessingBatchFiles = True
    End Function

    Public Function blnBuildEWSReprocessingRubyScript(ByVal sRubyScriptFileName As String, ByVal sSQLiteDBLocation As String, ByVal sEvidenceJSon As String, ByVal sCaseDir As String, ByVal sNumberOfWorkers As String, ByVal sMemoryPerWorker As String, ByVal sExportDirectory As String, ByVal sWorkerTempDir As String, ByVal bExtractFromSlackSpace As Boolean) As Boolean
        blnBuildEWSReprocessingRubyScript = False

        Dim ReprocessingRuby As StreamWriter

        ReprocessingRuby = New StreamWriter(sRubyScriptFileName)

        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("require 'thread'")
        ReprocessingRuby.WriteLine("require 'json'")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("CALLBACK_FREQUENCY = 100")
        ReprocessingRuby.WriteLine("callback_count = 0")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("load """ & sSQLiteDBLocation.Replace("\", "\\") & "\\Database.rb_""")
        ReprocessingRuby.WriteLine("load """ & sSQLiteDBLocation.Replace("\", "\\") & "\\SQLite.rb_""")
        ReprocessingRuby.WriteLine("db = SQLite.new(""" & sSQLiteDBLocation.Replace("\", "\\") & "\\NuixEmailArchiveMigrationManager.db3" & """" & ")")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("#######################################")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("file = File.read('" & sEvidenceJSon.Replace("\", "\\") & "')")
        ReprocessingRuby.WriteLine("parsed = JSON.parse(file)")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("    archive_file = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_file" & """" & "]")
        ReprocessingRuby.WriteLine("    archive_keystore = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_keystore" & """" & "]")
        ReprocessingRuby.WriteLine("    archive_password = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_password" & """" & "]")
        ReprocessingRuby.WriteLine("    archive_name = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_name" & """" & "]")
        ReprocessingRuby.WriteLine("    archive_centera = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera" & """" & "]")
        ReprocessingRuby.WriteLine("    archive_centera_ip = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera_ip" & """" & "]")
        ReprocessingRuby.WriteLine("    archive_centera_clip = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera_clip" & """" & "]")
        ReprocessingRuby.WriteLine("    archive_source_size = parsed[" & """" & "email_archive" & """" & "][" & """" & "source_size" & """" & "]")
        ReprocessingRuby.WriteLine("    archive_migration = parsed[" & """" & "email_archive" & """" & "][" & """" & "migration" & """" & "]")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("caseFactory = $utilities.getCaseFactory()")
        ReprocessingRuby.WriteLine("case_settings = {")
        ReprocessingRuby.WriteLine("    :compound => false,")
        ReprocessingRuby.WriteLine("    :name => " & """" & "#{archive_name}" & """" & ",")
        ReprocessingRuby.WriteLine("    :description => " & """" & "Created using Nuix Email Archive Migration Manager" & """" & ",")
        ReprocessingRuby.WriteLine("    :investigator => " & """" & "NEAMM EWS Reprocess" & """")
        ReprocessingRuby.WriteLine("}")
        ReprocessingRuby.WriteLine("$current_case = caseFactory.create(" & """" & sCaseDir.Replace("\", "\\") & """" & ", case_settings)")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("processor = $current_case.createProcessor")
        ReprocessingRuby.WriteLine("if archive_migration == " & """" & "lightspeed" & """")
        ReprocessingRuby.WriteLine("	processing_settings = {")
        ReprocessingRuby.WriteLine("		:traversalScope => " & """" & "full_traversal" & """" & ",")
        ReprocessingRuby.WriteLine("		:analysisLanguage => " & """" & "en" & """" & ",")
        ReprocessingRuby.WriteLine("		:identifyPhysicalFiles => true,")
        ReprocessingRuby.WriteLine("		:reuseEvidenceStores => true,")
        ReprocessingRuby.WriteLine("		:reportProcessingStatus => " & """" & "physical_files" & """")
        ReprocessingRuby.WriteLine("	}")
        ReprocessingRuby.WriteLine("end")
        ReprocessingRuby.WriteLine("processor.setProcessingSettings(processing_settings)")
        ReprocessingRuby.WriteLine("parallel_processing_settings = {")
        ReprocessingRuby.WriteLine("	:workerCount => " & sNumberOfWorkers & ",")
        ReprocessingRuby.WriteLine("	:workerMemory => " & sMemoryPerWorker & ",")
        ReprocessingRuby.WriteLine("	:embedBroker => true,")
        ReprocessingRuby.WriteLine("	:brokerMemory => " & sMemoryPerWorker & ",")
        ReprocessingRuby.WriteLine("   :workerTemp => " & """" & sWorkerTempDir.Replace("\", "\\") & """")

        ReprocessingRuby.WriteLine("}")
        ReprocessingRuby.WriteLine("processor.setParallelProcessingSettings(parallel_processing_settings)")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("cust_name = archive_name")
        ReprocessingRuby.WriteLine("evidence_container = processor.newEvidenceContainer(cust_name)")
        ReprocessingRuby.WriteLine("evidence_container.addFile(" & """" & sExportDirectory.Replace("\", "\\") & """" & ")")
        ReprocessingRuby.WriteLine("	evidence_container.setEncoding(" & """" & "utf-8" & """" & ")")
        ReprocessingRuby.WriteLine("	evidence_container.save")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("start_time = Time.now")
        ReprocessingRuby.WriteLine("last_progress = Time.now")
        ReprocessingRuby.WriteLine("semaphore = Mutex.new")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("puts")
        ReprocessingRuby.WriteLine("puts " & """" & "EWS Reprocessing for #{archive_name} has started..." & """")
        ReprocessingRuby.WriteLine("puts")
        ReprocessingRuby.WriteLine("printf " & """" & "%-40s %-25s %-25s %-25s" & """" & "," & """" & "Timestamp" & """" & ", " & """" & "Bytes Converted" & """" & ", " & """" & "Total Bytes" & """" & ", " & """" & "Percent (%) Completed" & """")
        ReprocessingRuby.WriteLine("puts")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("processor.when_progress_updated do |progress|")
        ReprocessingRuby.WriteLine("   semaphore.synchronize {")
        ReprocessingRuby.WriteLine("	class Numeric")
        ReprocessingRuby.WriteLine("	  def percent_of(n)")
        ReprocessingRuby.WriteLine("        self.to_f / n.to_f * 100")
        ReprocessingRuby.WriteLine("      end")
        ReprocessingRuby.WriteLine("    end")
        ReprocessingRuby.WriteLine("    # Progress message every 15 seconds")
        ReprocessingRuby.WriteLine("	current_size = progress.get_current_size")
        ReprocessingRuby.WriteLine("	total_size = progress.get_total_size")
        ReprocessingRuby.WriteLine("	percent = current_size.percent_of(total_size).round(1)")
        ReprocessingRuby.WriteLine("    if (Time.now - last_progress) > 15")
        ReprocessingRuby.WriteLine("         last_progress = Time.now")
        ReprocessingRuby.WriteLine("			printf " & """" & "\r%-40s %-25s %-25s %-25s" & """" & ", Time.now, current_size, total_size, percent")
        ReprocessingRuby.WriteLine("			updated_callback = [current_size,percent]")
        ReprocessingRuby.WriteLine("			db.update(" & """" & "UPDATE EWSReprocessingStats SET BytesProcessed = ?, PercentCompleted = ? WHERE CustodianName = '#{archive_name}'" & """" & ",updated_callback)")
        ReprocessingRuby.WriteLine("	end")
        ReprocessingRuby.WriteLine("    }")
        ReprocessingRuby.WriteLine("end")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("processor.process")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("puts")
        ReprocessingRuby.WriteLine("end_time = Time.now")
        ReprocessingRuby.WriteLine("puts")
        ReprocessingRuby.WriteLine("puts " & """" & "EWS Reprocessing for #{archive_name} has finished!" & """")
        ReprocessingRuby.WriteLine("puts")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("updated_callback = [" & """" & "Reprocessing Completed" & """" & ", 100, end_time, archive_source_size]")
        ReprocessingRuby.WriteLine("db.update(" & """" & "UPDATE EWSReprocessingStats SET ReprocessingStatus = ?, PercentCompleted = ?, ReprocessingEndTime = ?, BytesProcessed = ? WHERE CustodianName = '#{archive_name}'" & """" & ",updated_callback)")
        ReprocessingRuby.WriteLine("")
        ReprocessingRuby.WriteLine("$current_case.close")
        ReprocessingRuby.WriteLine("return")
        ReprocessingRuby.Close()

        blnBuildEWSReprocessingRubyScript = True

    End Function


    Private Function blnBuildUpdatedRubyScript(ByVal sNuixCase As String, ByVal sCustodianName As String, ByVal sRubyScriptFileName As String, sPSTLocation As String, ByVal sCustodianJSonFile As String) As Boolean
        blnBuildUpdatedRubyScript = False

        Dim CustodianRuby As StreamWriter
        Dim sCaseDir As String

        If eMailArchiveMigrationManager.NuixCaseDir.Contains("\") Then
            sCaseDir = eMailArchiveMigrationManager.NuixCaseDir.Replace("\", "\\")
        ElseIf eMailArchiveMigrationManager.NuixCaseDir.Contains("/") Then
            sCaseDir = eMailArchiveMigrationManager.NuixCaseDir.Replace("/", "\\")
        End If

        If sPSTLocation.Contains("\") Then
            sPSTLocation = sPSTLocation.Replace("\", "\\")
        ElseIf sPSTLocation.Contains("/") Then
            sPSTLocation = sPSTLocation.Replace("/", "\\")
        End If
        CustodianRuby = New StreamWriter(sRubyScriptFileName)
        CustodianRuby.WriteLine("# Menu Title: EWS Ingestion")
        CustodianRuby.WriteLine("# Needs Selected Items: false")
        CustodianRuby.WriteLine("# ")
        CustodianRuby.WriteLine("# This script expects a JSON configured with O365 parameters completed in order")
        CustodianRuby.WriteLine("# to automatically ingest data from PSTs to an O365 mailbox, archive or purges.")
        CustodianRuby.WriteLine("# ")
        CustodianRuby.WriteLine("# Version 1.6")
        CustodianRuby.WriteLine("# Jan 3 2017 - Alex Chatzistamatis, Nuix")
        CustodianRuby.WriteLine("# ")
        CustodianRuby.WriteLine("#######################################")

        CustodianRuby.WriteLine("require 'thread'")
        CustodianRuby.WriteLine("require 'json'")

        CustodianRuby.WriteLine("load """ & eMailArchiveMigrationManager.SQLiteDBLocation.Replace("\", "\\") & "\\Database.rb_""")
        CustodianRuby.WriteLine("load """ & eMailArchiveMigrationManager.SQLiteDBLocation.Replace("\", "\\") & "\\SQLite.rb_""")
        CustodianRuby.WriteLine("db = SQLite.new(""" & eMailArchiveMigrationManager.SQLiteDBLocation.Replace("\", "\\") & "\\NuixEmailArchiveMigrationManager.db3" & """" & ")")

        CustodianRuby.WriteLine("#######################################")

        If sCustodianJSonFile.Contains("Not Started") Then
            sCustodianJSonFile = sCustodianJSonFile.Replace("Not Started", "In Progress")
        End If

        CustodianRuby.WriteLine("file = File.read('" & sCustodianJSonFile.Replace("\", "\\") & "')")
        CustodianRuby.WriteLine("parsed = JSON.parse(file)")

        CustodianRuby.WriteLine("  o365_mailbox = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "mailbox" & """" & "]")
        CustodianRuby.WriteLine("  o365_custodian_name = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "custodian_name" & """" & "]")
        CustodianRuby.WriteLine("  o365_source_data = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "source_data" & """]")
        CustodianRuby.WriteLine("  o365_source_size = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "source_size" & """]")
        CustodianRuby.WriteLine("  o365_migration = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "migration" & """]")

        CustodianRuby.WriteLine("caseFactory = $utilities.getCaseFactory()")
        CustodianRuby.WriteLine("case_settings = {")
        CustodianRuby.WriteLine("    :compound => false,")
        CustodianRuby.WriteLine("    :name => " & """#{o365_custodian_name}" & """,")
        CustodianRuby.WriteLine("    :description => " & """" & """,")
        CustodianRuby.WriteLine("    :investigator => " & """EWS Pusher" & """")
        CustodianRuby.WriteLine("}")
        CustodianRuby.WriteLine("$current_case = caseFactory.create(""" & sNuixCase.Replace("\", "\\") & """" & ", case_settings)")

        If (eMailArchiveMigrationManager.NuixAppLocation.Contains("Nuix 6")) Then
            CustodianRuby.WriteLine("processor = $current_case.getProcessor")
        Else
            CustodianRuby.WriteLine("processor = $current_case.createProcessor")
        End If
        CustodianRuby.WriteLine("if o365_migration == """ & "lightspeed" & """")
        CustodianRuby.WriteLine("	processing_settings = {")
        CustodianRuby.WriteLine("		:traversalScope => """ & "full_traversal""" & ",")
        CustodianRuby.WriteLine("		:analysisLanguage => """ & "en" & """" & ",")
        CustodianRuby.WriteLine("		:identifyPhysicalFiles => true,")
        CustodianRuby.WriteLine("		:reuseEvidenceStores => true,")
        CustodianRuby.WriteLine("		:reportProcessingStatus => " & """" & "physical_files" & """")
        CustodianRuby.WriteLine("	}")
        CustodianRuby.WriteLine("end")
        CustodianRuby.WriteLine("processor.setProcessingSettings(processing_settings)")
        CustodianRuby.WriteLine("parallel_processing_settings = {")
        CustodianRuby.WriteLine("	:workerCount => " & eMailArchiveMigrationManager.O365NumberOfNuixWorkers & ",")
        CustodianRuby.WriteLine("	:workerMemory => " & eMailArchiveMigrationManager.O365MemoryPerWorker & ",")
        CustodianRuby.WriteLine("	:embedBroker => true,")
        CustodianRuby.WriteLine("	:brokerMemory => " & eMailArchiveMigrationManager.O365MemoryPerWorker & ",")
        CustodianRuby.WriteLine("   :workerTemp => " & """" & eMailArchiveMigrationManager.NuixWorkerTempDir.Replace("\", "\\") & """")
        CustodianRuby.WriteLine("}")
        CustodianRuby.WriteLine("processor.setParallelProcessingSettings(parallel_processing_settings)")
        CustodianRuby.WriteLine("# MIME Type Fiter for row-based items")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/csv" & """" & ", { process_embedded: false })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/tab-separated-values" & """" & ", { process_embedded: false })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.sqlite-database" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-registry" & """" & ", { text_strip: true})")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/x-plist" & """" & ", { process_embedded: false })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-logfile" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-mft" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-usnjrnl" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/exe" & """" & ", { process_text: false })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.tcpdump.pcap" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/x-common-log" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-iis-log" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-windows-event-log-record" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-windows-event-logx" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/x-pcapng" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.symantec-vault-stream-data" & """" & ", { enabled: false })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-cab-compressed" & """" & ", { process_embedded: true })")
        CustodianRuby.WriteLine("")
        CustodianRuby.WriteLine("cust_name = o365_custodian_name")
        CustodianRuby.WriteLine("evidence_container = processor.newEvidenceContainer(cust_name)")
        CustodianRuby.WriteLine("evidence_container.addFile(""" & "#{o365_source_data}" & """)")
        CustodianRuby.WriteLine("evidence_container.setEncoding(""" & "utf-8" & """)")
        CustodianRuby.WriteLine("evidence_container.save")
        CustodianRuby.WriteLine("")
        CustodianRuby.WriteLine("start_time = Time.now")
        CustodianRuby.WriteLine("last_progress = Time.now")
        CustodianRuby.WriteLine("semaphore = Mutex.new")
        CustodianRuby.WriteLine("")
        CustodianRuby.WriteLine("puts")
        CustodianRuby.WriteLine("puts """ & "Office 365 Ingestion for #{o365_mailbox} started at #{start_time}..." & """")
        CustodianRuby.WriteLine("puts")
        CustodianRuby.WriteLine("printf """ & "%-40s %-25s %-25s %-25s""" & "," & """" & "Timestamp" & """" & ", " & """" & "Bytes Uploaded" & """" & "," & """" & "Total Bytes" & """" & "," & """" & "Percent (%) Completed" & """")
        CustodianRuby.WriteLine("puts")

        CustodianRuby.WriteLine("processor.when_progress_updated do |progress|")
        CustodianRuby.WriteLine("    semaphore.synchronize {")
        CustodianRuby.WriteLine("	class Numeric")
        CustodianRuby.WriteLine("	  def percent_of(n)")
        CustodianRuby.WriteLine("        self.to_f / n.to_f * 100")
        CustodianRuby.WriteLine("      end")
        CustodianRuby.WriteLine("    end")
        CustodianRuby.WriteLine("    # Progress message every 15 seconds")
        CustodianRuby.WriteLine("	current_size = progress.get_current_size")
        CustodianRuby.WriteLine("	total_size = progress.get_total_size")
        CustodianRuby.WriteLine("	percent = current_size.percent_of(total_size).round(1)")
        CustodianRuby.WriteLine("    if (Time.now - last_progress) > 15")
        CustodianRuby.WriteLine("         last_progress = Time.now")
        CustodianRuby.WriteLine("			printf """ & "\r%-40s %-25s %-25s %-25s""" & ", Time.now, current_size, total_size, percent")
        CustodianRuby.WriteLine("			updated_callback = [current_size,percent]")
        CustodianRuby.WriteLine("			begin")
        CustodianRuby.WriteLine("			    db.update(""" & "UPDATE ewsIngestionStats SET BytesUploaded = ?, PercentCompleted = ? WHERE CustodianName = '#{o365_custodian_name}'" & """" & ",updated_callback)")
        CustodianRuby.WriteLine("			rescue")
        CustodianRuby.WriteLine("				processor.when_progress_updated")
        CustodianRuby.WriteLine("			end")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("    }")
        CustodianRuby.WriteLine("end")

        CustodianRuby.WriteLine("processor.process")
        CustodianRuby.WriteLine("puts")
        CustodianRuby.WriteLine("end_time = Time.now")
        CustodianRuby.WriteLine("updated_callback = [""" & "Completed" & """" & "," & """" & "100" & """" & ",o365_source_size" & "]")
        CustodianRuby.WriteLine("db.update(""" & "UPDATE ewsIngestionStats SET ProgressStatus = ?, PercentCompleted = ?, BytesUploaded = ? WHERE CustodianName = '#{o365_custodian_name}'" & """" & ",updated_callback)")

        CustodianRuby.WriteLine("$current_case.close")
        CustodianRuby.WriteLine("return")
        CustodianRuby.Close()

        blnBuildUpdatedRubyScript = True

    End Function
    Private Function blnBuildJSonFile(ByVal sCustodianName As String, ByVal sJSonFileName As String, sPSTLocation As String, ByVal sWorkerMemory As String, ByVal sWorkerCount As String, ByVal sCaseDir As String, ByVal sRubyFilePSTLocation As String, ByVal sDestinationFolder As String, ByVal sDestinationRoot As String, ByVal sDestinationSMTPAddress As String, ByVal dblPSTSourceSize As Double, ByVal sCustodianPSTPath As String) As Boolean
        blnBuildJSonFile = False

        Dim CustodianJSon As StreamWriter

        If sCaseDir.Contains("\") Then
            sCaseDir = sCaseDir.Replace("\", "\\")
        ElseIf sCaseDir.Contains("/") Then
            sCaseDir = sCaseDir.Replace("/", "\\")
        End If

        If sPSTLocation.Contains("\") Then
            sPSTLocation = sPSTLocation.Replace("\", "\\")
        ElseIf sPSTLocation.Contains("/") Then
            sPSTLocation = sPSTLocation.Replace("/", "\\")
        End If

        If sPSTLocation.Contains("\\\\") Then
            sPSTLocation = sPSTLocation.Replace("\\\\", "\\")
        End If

        CustodianJSon = New StreamWriter(sJSonFileName)

        CustodianJSon.WriteLine("{")
        CustodianJSon.WriteLine("  " & """" & "o365_mailbox" & """: {")
        CustodianJSon.WriteLine("    " & """server""" & ": " & """" & eMailArchiveMigrationManager.O365ExchangeServer & """" & ",")
        CustodianJSon.WriteLine("    " & """user""" & ": " & """" & eMailArchiveMigrationManager.O365AdminUserName & """" & ",")
        CustodianJSon.WriteLine("    " & """password""" & ": " & """" & eMailArchiveMigrationManager.O365AdminInfo & """" & ",")
        CustodianJSon.WriteLine("    " & """mailbox""" & ": " & """" & sDestinationSMTPAddress & """" & ",")
        CustodianJSon.WriteLine("    " & """folder""" & ": " & """" & sDestinationFolder & """" & ",")
        If sDestinationRoot.Contains("root_mailbox") Then
            CustodianJSon.WriteLine("    " & """root_mailbox""" & ": " & """mailbox""" & ",")
        Else
            CustodianJSon.WriteLine("    " & """root_mailbox""" & ": " & """" & vbNullString & """" & ",")
        End If

        If sDestinationRoot.Contains("root_archive") Then
            CustodianJSon.WriteLine("    " & """root_archive""" & ": " & """archive""" & ",")
        Else
            CustodianJSon.WriteLine("    " & """root_archive""" & ": " & """" & vbNullString & """" & ",")
        End If

        If sDestinationRoot.Contains("mailbox_purges") Then
            CustodianJSon.WriteLine("    " & """" & "mailbox_purges""" & ": " & """purges""" & ",")
        Else
            CustodianJSon.WriteLine("    " & """" & "mailbox_purges""" & ": " & """" & vbNullString & """" & ",")
        End If
        If sDestinationRoot.Contains("archive_purges") Then
            CustodianJSon.WriteLine("    " & """" & "archive_purges""" & ": " & """archive_purges""" & ",")
        Else
            CustodianJSon.WriteLine("    " & """" & "archive_purges""" & ": " & """" & vbNullString & """" & ",")
        End If
        If sDestinationRoot.Contains("mailbox_recoverable") Then
            CustodianJSon.WriteLine("    " & """" & "mailbox_recoverable""" & ": " & """recoverable_items""" & ",")
        Else
            CustodianJSon.WriteLine("    " & """" & "mailbox_recoverable""" & ": " & """" & vbNullString & """" & ",")
        End If
        If sDestinationRoot.Contains("archive_recoverable") Then
            CustodianJSon.WriteLine("    " & """" & "archive_recoverable""" & ": " & """archive_recoverable_items""" & ",")
        Else
            CustodianJSon.WriteLine("    " & """" & "archive_recoverable""" & ": " & """" & vbNullString & """" & ",")
        End If
        If sDestinationRoot.Contains("public_folders") Then
            CustodianJSon.WriteLine("    " & """" & "root_public""" & ": " & """public_folders""" & ",")
        Else
            CustodianJSon.WriteLine("    " & """" & "root_public""" & ": " & """" & vbNullString & """" & ",")
        End If

        CustodianJSon.WriteLine("    " & """from_date""" & ": " & """" & """" & ",")
        CustodianJSon.WriteLine("    " & """to_date""" & ": " & """" & """" & ",")
        CustodianJSon.WriteLine("	" & """custodian_name""" & " :" & """" & sCustodianName & """" & ",")
        CustodianJSon.WriteLine("	" & """source_data""" & ": " & """" & sRubyFilePSTLocation & """,")
        CustodianJSon.WriteLine("	" & """source_size""" & ": " & """" & dblPSTSourceSize & """" & ",")
        CustodianJSon.WriteLine("	" & """migration""" & ": " & """lightspeed""")
        CustodianJSon.WriteLine("  }")
        CustodianJSon.WriteLine("}")

        CustodianJSon.Close()

        blnBuildJSonFile = True

    End Function

    Public Function blnBuildReprocessingJSonFile(ByVal sJSonFileName As String, sSourceFolders As String, ByVal sEvidenceKeyStore As String, ByVal sEvidencePassword As String, ByVal sCaseName As String, ByVal sSourceSize As String, ByVal sMigrationType As String) As Boolean
        blnBuildReprocessingJSonFile = False

        Dim CustodianJSon As StreamWriter


        If sEvidenceKeyStore.Contains("\") Then
            sEvidenceKeyStore = sEvidenceKeyStore.Replace("\", "\\")
        ElseIf sEvidenceKeyStore.Contains("/") Then
            sEvidenceKeyStore = sEvidenceKeyStore.Replace("/", "\\")
        End If

        If sSourceFolders.Contains("\") Then
            sSourceFolders = sSourceFolders.Replace("\", "\\")
        ElseIf sSourceFolders.Contains("/") Then
            sSourceFolders = sSourceFolders.Replace("/", "\\")
        End If

        If sSourceFolders.Contains("\\\\") Then
            sSourceFolders = sSourceFolders.Replace("\\\\", "\\")
        End If


        CustodianJSon = New StreamWriter(sJSonFileName)

        CustodianJSon.WriteLine("{")
        CustodianJSon.WriteLine("  " & """" & "email_archive" & """" & ": {")

        CustodianJSon.WriteLine("     " & """" & "evidence_file" & """" & ": " & """" & sSourceFolders & """" & ",")
        CustodianJSon.WriteLine("     " & """" & "evidence_keystore" & """" & ": " & """" & sEvidenceKeyStore & """" & ",")
        CustodianJSon.WriteLine("     " & """" & "evidence_password" & """" & ": " & """" & sEvidencePassword & """" & ",")
        CustodianJSon.WriteLine("     " & """" & "evidence_name" & """" & ": " & """" & sCaseName.Replace("\", "\\") & """" & ",")
        CustodianJSon.WriteLine("     " & """" & "centera" & """" & ": " & """" & "no" & """" & ",")
        CustodianJSon.WriteLine("     " & """" & "centera_ip" & """" & ": " & """" & """" & ",")
        CustodianJSon.WriteLine("     " & """" & "centera_clip" & """" & ": " & """" & """" & ",")
        CustodianJSon.WriteLine("     " & """" & "source_size" & """" & ": " & """" & sSourceSize & """" & ",")
        CustodianJSon.WriteLine("     " & """" & "migration" & """" & ": " & """" & sMigrationType & """")
        CustodianJSon.WriteLine(" }")
        CustodianJSon.WriteLine("}")

        CustodianJSon.Close()

        blnBuildReprocessingJSonFile = True

    End Function
    Private Function blnBuildRubyScript(ByVal sCustodianName As String, ByVal sRubyScriptFileName As String, sPSTLocation As String, ByVal sWorkerMemory As String, ByVal sWorkerCount As String, ByVal sCaseDir As String, ByVal sNuixAppLocation As String, ByVal sWorkerTempDir As String) As Boolean
        blnBuildRubyScript = False

        Dim CustodianRuby As StreamWriter


        If sCaseDir.Contains("\") Then
            sCaseDir = sCaseDir.Replace("\", "\\")
        ElseIf sCaseDir.Contains("/") Then
            sCaseDir = sCaseDir.Replace("/", "\\")
        End If

        If sPSTLocation.Contains("\") Then
            sPSTLocation = sPSTLocation.Replace("\", "\\")
        ElseIf sPSTLocation.Contains("/") Then
            sPSTLocation = sPSTLocation.Replace("/", "\\")
        End If
        CustodianRuby = New StreamWriter(sRubyScriptFileName)

        CustodianRuby.WriteLine("caseFactory = $utilities.getCaseFactory()")
        CustodianRuby.WriteLine("case_settings = {")
        CustodianRuby.WriteLine("    :compound => false,")
        CustodianRuby.WriteLine("    :name => """ & sCustodianName & """" & ",")
        CustodianRuby.WriteLine("    :description => """"" & ",")
        CustodianRuby.WriteLine("    :investigator => """ & "PST Pusher" & """")
        CustodianRuby.WriteLine("}")
        CustodianRuby.WriteLine("$current_case = caseFactory.create(" & """" & sCaseDir & "\\" & sCustodianName & """" & ", case_settings)")
        CustodianRuby.WriteLine("")
        If (sNuixAppLocation.Contains("Nuix 6")) Then
            CustodianRuby.WriteLine("processor = $current_case.getProcessor")
        Else
            CustodianRuby.WriteLine("processor = $current_case.createProcessor")
        End If

        CustodianRuby.WriteLine("processing_settings = {")
        CustodianRuby.WriteLine("	:processText => false,")
        CustodianRuby.WriteLine("	:processLooseFileContents => false,")
        CustodianRuby.WriteLine("	:processForensicImages => false,")
        CustodianRuby.WriteLine("	:analysisLanguage => " & """" & "en" & """" & ",")
        CustodianRuby.WriteLine("	:stopWords => false,")
        CustodianRuby.WriteLine("	:stemming => false,")
        CustodianRuby.WriteLine("	:enableExactQueries => false,")
        CustodianRuby.WriteLine("	:extractNamedEntities => false,")
        CustodianRuby.WriteLine("	:extractShingles => false,")
        CustodianRuby.WriteLine("	:processTextSummaries => false,")
        CustodianRuby.WriteLine("	:extractFromSlackSpace => false,")
        CustodianRuby.WriteLine("	:carveFileSystemUnallocatedSpace => false,")
        CustodianRuby.WriteLine("	:carveUnidentifiedData => false,")
        CustodianRuby.WriteLine("	:recoverDeletedFiles => false,")
        CustodianRuby.WriteLine("	:identifyPhysicalFiles => true,")
        CustodianRuby.WriteLine("	:createThumbnails => false,")
        CustodianRuby.WriteLine("	:skinToneAnalysis => false,")
        CustodianRuby.WriteLine("	:calculateAuditedSize => false,")
        CustodianRuby.WriteLine("	:storeBinary => false,")
        CustodianRuby.WriteLine("	:reuseEvidenceStores => true,")
        CustodianRuby.WriteLine("	:processFamilyFields => false,")
        CustodianRuby.WriteLine("	:hideEmbeddedImmaterialData => false,")
        CustodianRuby.WriteLine("	:reportProcessingStatus => " & """" & "physical_files" & """")
        CustodianRuby.WriteLine("}")
        CustodianRuby.WriteLine("processor.setProcessingSettings(processing_settings)")
        CustodianRuby.WriteLine("parallel_processing_settings = {")
        CustodianRuby.WriteLine("	:workerCount => " & sWorkerCount & ",")
        CustodianRuby.WriteLine("	:workerMemory => " & sWorkerMemory & ",")
        CustodianRuby.WriteLine("	:embedBroker => true,")
        CustodianRuby.WriteLine("	:brokerMemory => " & sWorkerMemory & ",")
        CustodianRuby.WriteLine("   :workerTemp => " & """" & sWorkerTempDir.Replace("\", "\\") & """")
        CustodianRuby.WriteLine("}")
        CustodianRuby.WriteLine("processor.setParallelProcessingSettings(parallel_processing_settings)")
        CustodianRuby.WriteLine("# MIME Type Fiter for row-based items")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/csv" & """" & ", { process_embedded: false })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/tab-separated-values" & """" & ", { process_embedded: false })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.sqlite-database" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-registry" & """" & ", { text_strip: true})")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/x-plist" & """" & ", { process_embedded: false })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-logfile" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-mft" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-usnjrnl" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/exe" & """" & ", { process_text: false })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.tcpdump.pcap" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/x-common-log" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-iis-log" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-windows-event-log-record" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-windows-event-logx" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/x-pcapng" & """" & ", { process_embedded: false, text_strip: true })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.symantec-vault-stream-data" & """" & ", { enabled: false })")
        CustodianRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-cab-compressed" & """" & ", { process_embedded: true })")
        CustodianRuby.WriteLine(" ")
        CustodianRuby.WriteLine("evidence_name = " & """" & sCustodianName & """")
        CustodianRuby.WriteLine("evidence_container = processor.newEvidenceContainer(evidence_name)")
        CustodianRuby.WriteLine("evidence_container.addFile(" & """" & sPSTLocation & """" & ")")
        CustodianRuby.WriteLine("evidence_container.setEncoding(""" & "utf-8" & """" & ")")
        CustodianRuby.WriteLine("evidence_container.save")
        CustodianRuby.WriteLine("last_progress = Time.now")
        CustodianRuby.WriteLine("semaphore = Mutex.new")
        CustodianRuby.WriteLine("puts")
        CustodianRuby.WriteLine("puts " & """" & "Office 365 Upload for #{evidence_name} has begun..." & """")
        CustodianRuby.WriteLine("processor.when_progress_updated do |progress|")
        CustodianRuby.WriteLine("	semaphore.synchronize {")
        CustodianRuby.WriteLine("	# Progress message every 5 seconds")
        CustodianRuby.WriteLine("	if (Time.now - last_progress) > 5")
        CustodianRuby.WriteLine("		last_progress = Time.now")
        CustodianRuby.WriteLine("			puts")
        CustodianRuby.WriteLine("			puts " & """" & "#{Time.now}: Processed #{progress.get_current_size} bytes of #{progress.get_total_size} bytes for #{evidence_name}" & """")
        'CustodianRuby.WriteLine("			puts " & """" & "#{Time.now}: Processed #{progress.get_current_size} bytes of #{progress.get_total_size} bytes for " & """" & sCustodianName & """")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	}")
        CustodianRuby.WriteLine("end")
        CustodianRuby.WriteLine("processor.process")
        CustodianRuby.WriteLine("puts")
        CustodianRuby.WriteLine("puts " & """" & "Office 365 Upload for #{evidence_name} has completed!" & """")
        CustodianRuby.WriteLine("puts")
        CustodianRuby.WriteLine("$current_case.close")
        CustodianRuby.WriteLine("return")

        CustodianRuby.Close()

        blnBuildRubyScript = True

    End Function

    Private Function blnBuildIngestionCustodianMappingFile(ByVal sCustodianName As String, ByVal sMappingFileName As String, ByVal sDestinationFolder As String, ByVal sDestinationRoot As String, ByVal sDestinationSMTPAddress As String, ByVal dblPSTSourceSize As Double, ByVal sCustodianPSTPath As String) As Boolean
        blnBuildIngestionCustodianMappingFile = False

        Dim CustodianMappingFile As StreamWriter

        CustodianMappingFile = New StreamWriter(sMappingFileName)

        CustodianMappingFile.WriteLine(sCustodianName & "," & eMailArchiveMigrationManager.O365ExchangeServer & "," & eMailArchiveMigrationManager.O365Domain & "," & eMailArchiveMigrationManager.O365AdminUserName & "," & eMailArchiveMigrationManager.O365AdminInfo & "," & sDestinationFolder & "," & sDestinationRoot & "," & sDestinationSMTPAddress)

        CustodianMappingFile.Close()

        blnBuildIngestionCustodianMappingFile = True
    End Function
    Private Sub btnCustomerMappingFile_Click(sender As Object, e As EventArgs)
        Dim OpenFileDialog1 As New OpenFileDialog

        With OpenFileDialog1
            .Filter = "CSV|*.csv"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            txtCustodianMappingFile.Text = OpenFileDialog1.FileName.ToString
        End If
    End Sub
    Private Sub txtMemoryPerNuixInstance_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtMemoryPerWorker_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtTimeout_KeyPress(sender As Object, e As KeyPressEventArgs)

        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtRetries_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtFailureIncrement_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
    Private Function blnMoveConsolidatedPSTs(ByRef lstSourceFolders As List(Of String), ByRef lstSourceServers As List(Of String), ByVal sReportDirectory As String) As Boolean
        Dim currentdirectory As DirectoryInfo
        Dim PSTName As New List(Of String)
        Dim PSTPath As New List(Of String)
        Dim PSTSize As New List(Of String)
        Dim PSTServerNumber As New List(Of String)
        Dim UniqueName As New List(Of String)

        Dim sMachineName As String
        Dim sOutputFileName As String
        Dim iIndexOfSourceFolder As Integer

        Dim sServerNumber As String

        Dim OutputStream As StreamWriter

        sMachineName = System.Net.Dns.GetHostName()
        sOutputFileName = sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv"

        PSTName = New List(Of String)
        PSTPath = New List(Of String)
        PSTSize = New List(Of String)
        PSTServerNumber = New List(Of String)
        UniqueName = New List(Of String)

        Dim StartDate As DateTime = DateTime.Now
        StartDate = Now

        OutputStream = New StreamWriter(sReportDirectory & "\" & sOutputFileName)
        OutputStream.WriteLine("Custodian Name" & "," & "Original Folder" & "," & "Destination Folder")

        For Each sSourceFolder In lstSourceFolders
            iIndexOfSourceFolder = lstSourceFolders.IndexOf(sSourceFolder)
            sServerNumber = lstSourceServers(iIndexOfSourceFolder)
            currentdirectory = New DirectoryInfo(sSourceFolder)
            For Each CustodianDir In currentdirectory.GetDirectories
                My.Computer.FileSystem.MoveDirectory(CustodianDir.FullName, sReportDirectory & "\" & CustodianDir.Name)
                OutputStream.WriteLine(CustodianDir.Name & "," & CustodianDir.FullName & "," & sReportDirectory & "\" & CustodianDir.Name)
            Next
        Next

        OutputStream.Close()
        blnMoveConsolidatedPSTs = False

    End Function

    Private Function blnBuildPSTReports(ByVal sPSTSourceFolder As String, ByVal ReportOutputFile As StreamWriter) As Boolean

        Dim PSTName As New List(Of String)
        Dim PSTPath As New List(Of String)
        Dim PSTSize As New List(Of String)
        Dim PSTServerNumber As New List(Of String)
        Dim UniqueName As New List(Of String)
        Dim sRemainingFileName As String
        Dim sRemainingNumber As String

        Dim sFileName As String
        Dim sFilePath As String
        Dim sCustodianName As String

        Dim iCounter As Integer

        PSTName = New List(Of String)
        PSTPath = New List(Of String)
        PSTSize = New List(Of String)
        UniqueName = New List(Of String)

        Dim StartDate As DateTime = DateTime.Now
        StartDate = Now

        subPSTDirSearch(sPSTSourceFolder.ToString, PSTName, PSTPath, PSTSize)

        For Each pst In PSTName
            iCounter = iCounter + 1
            sFileName = pst.ToString
            sFileName = Path.GetFileNameWithoutExtension(sFileName)

            sFilePath = PSTPath(iCounter - 1)

            If sFileName.IndexOf("_") = -1 Then
                sCustodianName = sFileName

            Else
                If sFileName.LastIndexOf("_") > 0 Then
                    sRemainingFileName = sFileName.Substring(sFileName.LastIndexOf("_"))
                    sRemainingNumber = sRemainingFileName.Replace("_", "")
                    Try
                        If IsNumeric(CInt(sRemainingNumber)) Then
                            sCustodianName = sFileName.Substring(0, sFileName.Length - sRemainingFileName.Length)
                        Else
                            If sRemainingFileName.Length = 12 Then
                                sCustodianName = sFileName.Replace(sRemainingFileName, "")
                            End If
                        End If
                    Catch ex As Exception
                        sCustodianName = sCustodianName
                    End Try
                End If
            End If

            '            OutputStream.WriteLine(sCustodianName & "," & sFileName & "," & PSTSize(iCounter - 1) & "," & sFilePath & "\" & sFileName & ".pst" & "," & sLegalHold)
            ReportOutputFile.WriteLine(sCustodianName & "," & sFileName & "," & PSTSize(iCounter - 1) & "," & sFilePath & "\")

        Next
        'OutputStream.WriteLine("Start Time = " & StartDate & ", End Time = " & EndDate & " , Total Time (In Seconds) = " & Duration.Seconds)

        ReportOutputFile.Close()
        blnBuildPSTReports = False

    End Function

    Private Function blnConsolidateCustodianPSTs(ByRef sSourceFolder As String, sCustodianName As String, ByVal sDestinationFolder As String) As Boolean
        Dim lstPSTName As New List(Of String)
        Dim lstPSTPath As New List(Of String)
        Dim lstPSTSize As New List(Of Double)
        Dim UniqueName As New List(Of String)
        Dim StartDate As DateTime = DateTime.Now
        Dim sNewFileName As String
        Dim iEmailCounter As Integer
        Dim sPSTName As String

        blnConsolidateCustodianPSTs = True
        StartDate = Now

        iEmailCounter = 0

        subCustodianPSTSearch(sSourceFolder, sCustodianName, lstPSTName, lstPSTPath, lstPSTSize)

        For Each item In lstPSTName
            If item.ToString.LastIndexOf("_") > 0 Then
                sCustodianName = item.ToString.Substring(0, item.ToString.LastIndexOf("_"))
                sPSTName = Path.GetFileNameWithoutExtension(item.ToString)

                sNewFileName = sPSTName.Substring(0, sPSTName.LastIndexOf("_")) & "_" & iEmailCounter & ".pst"
            Else
                sPSTName = Path.GetFileNameWithoutExtension(item.ToString)
                sNewFileName = sPSTName & "_" & iEmailCounter & ".pst"
            End If

            If Not System.IO.Directory.Exists(sDestinationFolder) Then
                System.IO.Directory.CreateDirectory(sDestinationFolder)
            End If

            If eMailArchiveMigrationManager.PSTConsolidation = "Move" Then
                My.Computer.FileSystem.MoveFile(lstPSTPath(iEmailCounter) & "\" & sPSTName & ".pst", sDestinationFolder & "\" & sNewFileName)
            ElseIf eMailArchiveMigrationManager.PSTConsolidation = "Copy" Then
                My.Computer.FileSystem.CopyFile(lstPSTPath(iEmailCounter) & "\" & sPSTName & ".pst", sDestinationFolder & "\" & sNewFileName)

            End If
            iEmailCounter = iEmailCounter + 1
        Next
        blnConsolidateCustodianPSTs = True
    End Function
    Sub subPSTDirSearch(ByVal sDir As String, ByRef PSTName As List(Of String), ByRef PSTPath As List(Of String), ByRef PSTSize As List(Of Double))
        Dim d As String
        Dim Length As String
        Dim currentdirectory As DirectoryInfo
        Dim extension As String

        Dim common As New Common

        Try
            currentdirectory = New DirectoryInfo(sDir)

            For Each File In currentdirectory.GetFiles
                extension = Path.GetExtension(File.ToString)
                If extension = ".pst" Then 'If InStr(File.Name, ".pst") > 0 Then
                    PSTName.Add(File.Name)
                    PSTPath.Add(File.DirectoryName)
                    PSTSize.Add(File.Length)
                End If
            Next

            Length = Directory.GetDirectories(sDir).Length
            If Length > 0 Then
                For Each d In Directory.GetDirectories(sDir)
                    subPSTDirSearch(d, PSTName, PSTPath, PSTSize)
                Next
            End If
        Catch ex As Exception
            common.Logger(psIngestionLogFile, "Error 1 - Directory Search" & ex.ToString)
        End Try

    End Sub

    Sub subCustodianPSTSearch(ByVal sDir As String, ByRef sCustodianName As String, ByRef PSTName As List(Of String), ByRef PSTPath As List(Of String), ByRef PSTSize As List(Of Double))
        Dim d As String
        Dim Length As String
        Dim currentdirectory As DirectoryInfo
        Dim extension As String

        Dim common As New Common

        Try
            currentdirectory = New DirectoryInfo(sDir)

            For Each File In currentdirectory.GetFiles
                If File.Name.ToLower.Contains(sCustodianName.ToLower) Then
                    extension = Path.GetExtension(File.ToString)
                    If extension = ".pst" Then
                        PSTName.Add(File.Name)
                        PSTPath.Add(File.DirectoryName)
                        PSTSize.Add(File.Length)
                    End If
                End If
            Next

            Length = Directory.GetDirectories(sDir).Length
            If Length > 0 Then
                For Each d In Directory.GetDirectories(sDir)
                    subCustodianPSTSearch(d, sCustodianName, PSTName, PSTPath, PSTSize)
                Next
            End If
        Catch ex As Exception
            common.Logger(psIngestionLogFile, "Error 2 - Directory Search" & ex.ToString)
        End Try

    End Sub

    Private Sub btnBuildNuixFiles_Click(sender As Object, e As EventArgs)
        Dim bStatus As Boolean
        Dim lstSelectedCustodiansInfo As List(Of String)
        Dim sTimeout As String
        Dim msgboxReturn As DialogResult

        Dim sNuixAppMemory As String
        Dim lstCurrentCustodianFolder As List(Of String)
        Dim sInvalidFolder As String
        Dim lstAllCustodiansInfo As List(Of String)

        Dim bSelectedCustodian As Boolean
        Dim sCustodianName As String
        Dim sPSTPath As String
        Dim iTotalNumPSTs As Integer
        Dim dTotalPSTSize As Double
        Dim sGroupID As String
        Dim sDestinationFolder As String
        Dim sDestinationRoot As String
        Dim sDestinationSMTP As String
        Dim sProcessID As String
        Dim sIngestionStartTime As String
        Dim sIngestionEndTime As String
        Dim dBytesUploaded As Double
        Dim sProgressStatus As String
        Dim iPercentCompleted As Integer
        Dim dSuccess As Double
        Dim dFailed As Double
        Dim sSummaryReportLocation As String

        Dim sCurrentPSTLocation As String
        Dim sNewFolder As String
        Dim sNewCustodianFolder As String
        Dim sCustodianPSTPath As String
        Dim sProcessingFilesDirectory As String
        Dim sCaseDirectory As String
        Dim sOutputDirectory As String
        Dim sLogDirectory As String

        Dim asSelectedCustodian() As String
        Dim asCustodianInfo() As String

        Dim sPSTDirectory As String

        Dim common As New Common
        Dim dbService As DatabaseService

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"

        If psSettingsFile = vbNullString Then
            msgboxReturn = MessageBox.Show("The settings have not been loaded. Would you like to load the setting dialog.", "Settings not loaded!", MessageBoxButtons.YesNo)
            If msgboxReturn = vbYes Then
                Dim SettingsForm As O365ExtractionSettings
                SettingsForm = New O365ExtractionSettings
                SettingsForm.ShowDialog()
                Exit Sub
            Else
                Exit Sub
            End If
        End If

        lstAllCustodiansInfo = New List(Of String)
        lstSelectedCustodiansInfo = New List(Of String)
        lstCurrentCustodianFolder = New List(Of String)

        For Each Row In grdPSTInfo.Rows
            If Row.cells("SelectCustodian").value = True Then
                If ((Row.cells("DestinationFolder").value = vbNullString) Or (Row.cells("DestinationRoot").value = vbNullString) Or (Row.cells("DestinationSMTP").value = vbNullString)) Then
                    MessageBox.Show("You have not Added the appropriate Office 365 ingestion details.  Please update custodian mapping file and reload.")
                ElseIf (Row.cells("ProgressStatus").value = "In Progress") Then
                    Row.cells("SelectCustodian").value = False
                    Row.cells("SelectCustodian").readonly = True
                ElseIf (Row.cells("ProgressStatus").value = "Not Started") Then
                    Row.cells("SelectCustodian").value = False
                    Row.cells("SelectCustodian").readonly = True
                Else
                    lstSelectedCustodiansInfo.Add(Row.cells("CustodianName").value & "," & Row.cells("PSTPath").value & "," & Row.cells("SizeOfPSTs").value.ToString.Replace(",", "") & "," & Row.cells("DestinationFolder").value & "," & Row.cells("DestinationRoot").value & "," & Row.cells("DestinationSMTP").value)
                End If
            Else
                If Row.cells("CustodianName").value <> vbNullString Then
                    lstAllCustodiansInfo.Add(Row.cells("SelectCustodian").value & "," & Row.cells("CustodianName").value & "," & Row.cells("PSTPath").value & "," & Row.cells("NumberOfPSTs").value & "," & Row.cells("SizeOfPSTs").value.ToString.Replace(",", "") & "," & Row.cells("GroupID").value & "," & Row.cells("DestinationFolder").value & "," & Row.cells("DestinationRoot").value & "," & Row.cells("DestinationSMTP").value & "," & "" & "," & Row.cells("ProcessID").value & "," & Row.cells("IngestionStartTime").value & "," & Row.cells("IngestionEndTime").value & "," & Row.cells("BytesUploaded").value & "," & Row.cells("PercentageCompleted").value & "," & Row.cells("NumberSuccess").value & "," & Row.cells("NumberFailed").value & "," & Row.cells("SummaryReportLocation").value)
                End If
            End If
        Next

        sNuixAppMemory = CInt(eMailArchiveMigrationManager.O365NuixAppMemory) / 1000
        sNuixAppMemory = Math.Round(CInt(sNuixAppMemory))
        sNuixAppMemory = "-Xmx" & sNuixAppMemory & "g"
        sTimeout = eMailArchiveMigrationManager.WorkerTimeout

        My.Computer.FileSystem.CreateDirectory(eMailArchiveMigrationManager.NuixCaseDir)

        For Each Custodian In lstSelectedCustodiansInfo
            asCustodianInfo = Split(Custodian.ToString, ",")
            bStatus = common.BuildIngestionCustodianNuixFiles(asCustodianInfo, sPSTDirectory, sTimeout, sNuixAppMemory, lstCurrentCustodianFolder, asCustodianInfo(2), asCustodianInfo(3), asCustodianInfo(4), asCustodianInfo(5))
            bStatus = dbService.GetCustodianPSTLocation(sSQLiteDatabaseFullName, asCustodianInfo(0), sCurrentPSTLocation)
            sNewFolder = sPSTDirectory & "EWS Ingestion\Not Started\"
            If sNewFolder.Contains("\\") Then
                sNewFolder = sNewFolder.Replace("\\", "\")
            End If
            If sNewFolder.Substring(sNewFolder.Length - 1) = "\" Then
                sNewCustodianFolder = sNewFolder & asCustodianInfo(0) & "\"
            Else
                sNewCustodianFolder = sNewFolder & "\" & asCustodianInfo(0) & "\"
            End If
            common.Logger(psIngestionLogFile, "Moving " & sPSTDirectory & asCustodianInfo(0) & " to " & sNewCustodianFolder)
            bStatus = blnMoveCustodianFolder(asCustodianInfo(0), sPSTDirectory & asCustodianInfo(0), sNewCustodianFolder, sInvalidFolder)
        Next

        For Each selectedCustodian In lstSelectedCustodiansInfo
            asSelectedCustodian = Split(selectedCustodian.ToString, ",")

            For Each row In grdPSTInfo.Rows
                If row.cells("CustodianName").value = asSelectedCustodian(0) Then
                    row.cells("StatusImage").value = My.Resources.waitingtostart1
                    row.cells("SelectCustodian").Value = False
                    bStatus = dbService.GetCustodianPSTLocation(sSQLiteDatabaseFullName, asSelectedCustodian(0), sCustodianPSTPath)
                    row.cells("PSTPath").value = sCustodianPSTPath
                    row.cells("ProgressStatus").value = "Not Started"
                    row.cells("SelectCustodian").readonly = True

                    bSelectedCustodian = row.cells("SelectCustodian").value
                    sCustodianName = row.cells("CustodianName").value
                    sPSTPath = row.cells("PSTPath").value
                    iTotalNumPSTs = row.cells("NumberOfPSTs").value
                    dTotalPSTSize = row.cells("SizeOfPSTs").value
                    sGroupID = row.cells("GroupID").value
                    sDestinationFolder = row.cells("DestinationFolder").value
                    sDestinationRoot = row.cells("DestinationRoot").value
                    sDestinationSMTP = row.cells("DestinationSMTP").value
                    sProgressStatus = row.cells("ProgressStatus").value
                    sProcessingFilesDirectory = row.cells("ProcessingFilesDirectory").value
                    sCaseDirectory = row.cells("CaseDirectory").value
                    sOutputDirectory = row.cells("OutputDirectory").value
                    sLogDirectory = row.cells("LogDirectory").value

                    sProcessID = row.cells("ProcessID").value
                    sIngestionStartTime = row.cells("IngestionStartTime").value
                    sIngestionEndTime = row.cells("IngestionEndTime").value
                    If row.cells("BytesUploaded").value = vbNullString Then
                        dBytesUploaded = 0.0
                    Else
                        dBytesUploaded = row.cells("BytesUploaded").value
                    End If
                    If row.cells("PercentageCompleted").value = vbNullString Then
                        iPercentCompleted = 0
                    Else
                        iPercentCompleted = CInt(row.cells("PercentageCompleted").value.ToString.Replace("%", ""))
                        iPercentCompleted = iPercentCompleted.ToString.Replace(".00", "")
                    End If
                    If row.cells("NumberSuccess").value = vbNullString Then
                        dSuccess = 0.0
                    Else
                        dSuccess = row.cells("NumberSuccess").value
                    End If
                    If row.cells("NumberFailed").value = vbNullString Then
                        dFailed = 0.0
                    Else
                        dFailed = row.cells("NumberFailed").value
                    End If
                    sSummaryReportLocation = row.cells("SummaryReportLocation").value
                    common.Logger(psIngestionLogFile, "Updating database for " & sCustodianName)
                    bStatus = blnUpdateSQLiteDB(sCustodianName, sProgressStatus, sPSTPath, iTotalNumPSTs, dTotalPSTSize, sGroupID, sDestinationFolder, sDestinationRoot, sDestinationSMTP, sProcessID, sIngestionStartTime, sIngestionEndTime, dBytesUploaded, dSuccess, dFailed, sProcessingFilesDirectory, sCaseDirectory, sOutputDirectory, sLogDirectory, sSummaryReportLocation)
                End If
            Next
        Next

        MessageBox.Show("Finished building Nuix files for selected custodians")
    End Sub

    Private Function blnUpdateSQLiteDB(ByVal sCustodianName As String, ByVal sProgressStatus As String, ByVal sPSTPath As String, ByVal iNumberOfPSTs As Integer, ByVal dblTotalPSTSize As Double, ByVal sGroupID As String, ByVal sDestinationFolder As String, ByVal sDestinationRoot As String, ByVal sDestinationSMTP As String, ByVal sProcessID As String, ByVal sIngestionStartTime As String, ByVal sIngestionEndTime As String, ByVal dblBytesUploaded As Double, ByVal dblSuccess As Double, ByVal dblFailed As Double, ByVal sProcessingFilesDirectory As String, ByVal sCaseDirectory As String, ByVal sOutputDirectory As String, ByVal sLogDirectory As String, ByVal sSummaryReportLocation As String) As Boolean
        blnUpdateSQLiteDB = False
        Dim SQLConnection As SQLiteConnection
        Dim sInsertEWSIngestionData As String
        Dim sUpdateEWSIngestionData As String
        Dim iPercentCompleted As Integer

        Dim common As New Common

        If Not sCustodianName = vbNullString Then
            SQLConnection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
            SQLConnection.Open()

            Using SQLSelectCommand As New SQLiteCommand("SELECT CustodianName FROM ewsIngestionStats WHERE CustodianName='" & sCustodianName & "'")
                With SQLSelectCommand
                    .Connection = SQLConnection
                End With
                Using readerObject As SQLiteDataReader = SQLSelectCommand.ExecuteReader
                    While readerObject.Read
                        sCustodianName = readerObject("CustodianName").ToString
                    End While
                End Using
            End Using

            SQLConnection.Close()

            If sCustodianName = vbNullString Then
                Using Connection As New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;New=False;Compress=True;")
                    sInsertEWSIngestionData = "Insert into ewsIngestionStats (CustodianName, PSTPath, NumberOfPSTs, TotalSizeOfPSTs, GroupID, DestinationFolder, DestinationRoot, DestinationSMTP ProcessID, IngestionStartTime, IngestionEndTime, BytesUploaded, Success, Failed, ProcessingFilesDirectory, CaseDirectory, OutputDirectory, LogDirectory, SummaryReportLocation) Values "
                    sInsertEWSIngestionData = sInsertEWSIngestionData & "(@CustodianName, @PSTPath, @NumberOfPSTs, @TotalSizeOfPSTs, @GroupID, @DestinationFolder, @DestinationRoot, @DestinationSMTP, @ProcessID, @IngestionStartTime, @IngestionEndTime, @BytesUploaded, @Success, @Failed, @ProcessingFilesDirectory, @CaseDirectory, @OutputDirectory, @LogDirectory, @SummaryReportLocation)"
                    Using oInsertEWSIngestionDataCommand As New SQLiteCommand()
                        With oInsertEWSIngestionDataCommand
                            .Connection = Connection
                            .CommandText = sInsertEWSIngestionData
                            .Parameters.AddWithValue("@CustodianName", sCustodianName)
                            .Parameters.AddWithValue("@PSTPath", sPSTPath)
                            .Parameters.AddWithValue("@NumberOfPSTs", iNumberOfPSTs)
                            .Parameters.AddWithValue("@TotalSizeOfPSTs", dblTotalPSTSize)
                            .Parameters.AddWithValue("@GroupID", sGroupID)
                            .Parameters.AddWithValue("@DestinationFolder", sDestinationFolder)
                            .Parameters.AddWithValue("@DestinationRoot", sDestinationRoot)
                            .Parameters.AddWithValue("@DestinationSMTP", sDestinationSMTP)
                            .Parameters.AddWithValue("@ProcessID", sProcessID)
                            .Parameters.AddWithValue("@IngestionStartTime", sIngestionStartTime)
                            .Parameters.AddWithValue("@IngestionEndTime", sIngestionEndTime)
                            .Parameters.AddWithValue("@BytesUploaded", dblBytesUploaded)
                            .Parameters.AddWithValue("@PercentCompleted", iPercentCompleted)
                            .Parameters.AddWithValue("@Success", dblSuccess)
                            .Parameters.AddWithValue("@Failed", dblFailed)
                            .Parameters.AddWithValue("@ProcessingFilesDirectory", sProcessingFilesDirectory)
                            .Parameters.AddWithValue("@CaseDirectory", sCaseDirectory)
                            .Parameters.AddWithValue("@OutputDirectory", sOutputDirectory)
                            .Parameters.AddWithValue("@LogDirectory", sLogDirectory)
                            .Parameters.AddWithValue("@SummaryReportLocation", sSummaryReportLocation)
                        End With
                        Try
                            Connection.Open()
                            oInsertEWSIngestionDataCommand.ExecuteNonQuery()
                            Connection.Close()
                        Catch ex As Exception
                            common.Logger(psIngestionLogFile, "Error 3 - " & ex.Message.ToString())
                        End Try
                    End Using
                End Using

            Else
                Using Connection As New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;New=False;Compress=True;")
                    sUpdateEWSIngestionData = "Update ewsIngestionStats set PSTPath = @PSTPath, ProcessID = @ProcessID, IngestionStartTime = @IngestionStartTime, IngestionEndTime = @IngestionEndTime, "
                    sUpdateEWSIngestionData = sUpdateEWSIngestionData & "BytesUploaded = @BytesUploaded, PercentCompleted = @PercentCompleted, Success = @Success, Failed = @Failed, ProcessingFilesDirectory = @ProcessingFilesDirectory, CaseDirectory = @CaseDirectory, OutputDirectory = @OutputDirectory, LogDirectory = @LogDirectory, SummaryReportLocation = @SummaryReportLocation WHERE CustodianName = @CustodianName"

                    Using oUpdateEWSIngestionDataCommand As New SQLiteCommand()
                        With oUpdateEWSIngestionDataCommand
                            .Connection = Connection
                            .CommandText = sUpdateEWSIngestionData
                            .Parameters.AddWithValue("@CustodianName", sCustodianName)
                            .Parameters.AddWithValue("@ProgressStatus", sProgressStatus)
                            .Parameters.AddWithValue("@PSTPath", sPSTPath)
                            .Parameters.AddWithValue("@ProcessID", sProcessID)
                            .Parameters.AddWithValue("@IngestionStartTime", sIngestionStartTime)
                            .Parameters.AddWithValue("@IngestionEndTime", sIngestionEndTime)
                            .Parameters.AddWithValue("@BytesUploaded", dblBytesUploaded)
                            .Parameters.AddWithValue("@PercentCompleted", iPercentCompleted)
                            .Parameters.AddWithValue("@Success", dblSuccess)
                            .Parameters.AddWithValue("@Failed", dblFailed)
                            .Parameters.AddWithValue("@ProcessingFilesDirectory", sProcessingFilesDirectory)
                            .Parameters.AddWithValue("@CaseDirectory", sCaseDirectory)
                            .Parameters.AddWithValue("@OutputDirectory", sOutputDirectory)
                            .Parameters.AddWithValue("@LogDirectory", sLogDirectory)
                            .Parameters.AddWithValue("@SummaryReportLocation", sSummaryReportLocation)
                        End With
                        Try
                            Connection.Open()
                            oUpdateEWSIngestionDataCommand.ExecuteNonQuery()
                            Connection.Close()
                        Catch ex As Exception
                            common.Logger(psIngestionLogFile, "Error 4 - " & ex.Message.ToString())
                        End Try
                    End Using
                End Using

            End If
        End If

        blnUpdateSQLiteDB = True
    End Function


    Private Function blnUpdateSQLiteDBCustodianInfo(ByVal sCustodianName As String, ByVal sPSTPath As String, ByVal sProcessID As String, ByVal sIngestionEndTime As String, ByVal dSuccessItems As Double, ByVal dFailedItems As Double, ByVal sSummaryReportLocation As String) As Boolean
        blnUpdateSQLiteDBCustodianInfo = False
        Dim sUpdateEWSIngestionData As String

        Dim common As New Common

        Using Connection As New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;New=False;Compress=True;")
            sUpdateEWSIngestionData = "Update ewsIngestionStats set PSTPath = @PSTPath, ProcessID = @ProcessID, IngestionEndTime = @IngestionEndTime, "
            sUpdateEWSIngestionData = sUpdateEWSIngestionData & "PercentCompleted = @PercentCompleted, Success = @Success, Failed = @Failed, SummaryReportLocation = @SummaryReportLocation WHERE CustodianName = @CustodianName"

            Using oUpdateEWSIngestionDataCommand As New SQLiteCommand()
                With oUpdateEWSIngestionDataCommand
                    .Connection = Connection
                    .CommandText = sUpdateEWSIngestionData
                    .Parameters.AddWithValue("@CustodianName", sCustodianName)
                    .Parameters.AddWithValue("@PSTPath", sPSTPath)
                    .Parameters.AddWithValue("@ProcessID", "")
                    .Parameters.AddWithValue("@IngestionEndTime", sIngestionEndTime)
                    .Parameters.AddWithValue("@PercentCompleted", 100)
                    .Parameters.AddWithValue("@Success", dSuccessItems)
                    .Parameters.AddWithValue("@Failed", dFailedItems)
                    .Parameters.AddWithValue("@SummaryReportLocation", sSummaryReportLocation)
                    Try
                        Connection.Open()
                        oUpdateEWSIngestionDataCommand.ExecuteNonQuery()
                        Connection.Close()
                    Catch ex As Exception
                        common.Logger(psIngestionLogFile, "Error updating custodian PST Location in the database.")
                        common.Logger(psIngestionLogFile, ex.ToString)
                    End Try
                End With
            End Using
        End Using

        blnUpdateSQLiteDBCustodianInfo = True
    End Function

    Private Sub btnSelectAll_Click(sender As Object, e As EventArgs) Handles btnSelectAll.Click
        If grdPSTInfo.RowCount > 0 Then
            If btnSelectAll.Text = "Select All" Then

                For Each row In grdPSTInfo.Rows
                    If row.cells("CustodianName").value <> vbNullString Then
                        row.cells("SelectCustodian").value = True
                    End If
                Next
                btnSelectAll.Text = "Deselect All"
            Else

                For Each row In grdPSTInfo.Rows
                    row.cells("SelectCustodian").value = False
                Next
                btnSelectAll.Text = "Select All"

            End If
        End If

    End Sub

    Private Sub btnLoadCustodianMappingData_Click(sender As Object, e As EventArgs)
        Dim lstMappingCustodianName As List(Of String)
        Dim lstMappingDestinationRoot As List(Of String)
        Dim lstMappingDestinationFolder As List(Of String)
        Dim lstMappingDestinationSMTP As List(Of String)
        Dim lstMappingCustodianGroupID As List(Of String)
        Dim lstDuplicateMappingCustodianName As List(Of String)
        Dim lstDuplicateMappingDestinationRoot As List(Of String)
        Dim lstDuplicateMappingDestinationFolder As List(Of String)
        Dim lstDuplicateMappingDestinationSMTP As List(Of String)
        Dim lstDuplicateMappingCustodianGroupID As List(Of String)
        Dim lstUniqueGroupID As List(Of String)

        Dim iCustodianIndex As Integer

        Dim sCurrentRow() As String

        lstMappingCustodianName = New List(Of String)
        lstMappingDestinationRoot = New List(Of String)
        lstMappingDestinationFolder = New List(Of String)
        lstMappingDestinationSMTP = New List(Of String)
        lstMappingCustodianGroupID = New List(Of String)
        lstDuplicateMappingCustodianName = New List(Of String)
        lstDuplicateMappingDestinationRoot = New List(Of String)
        lstDuplicateMappingDestinationFolder = New List(Of String)
        lstDuplicateMappingDestinationSMTP = New List(Of String)
        lstDuplicateMappingCustodianGroupID = New List(Of String)
        lstUniqueGroupID = New List(Of String)

        lstDuplicateMappingCustodianName = New List(Of String)


        If txtCustodianMappingFile.Text = vbNullString Then
            MessageBox.Show("You must select a Custodian Mapping File.")
        Else

            Dim fileCSVFile As New Microsoft.VisualBasic.FileIO.TextFieldParser(txtCustodianMappingFile.Text)
            fileCSVFile.TextFieldType = FileIO.FieldType.Delimited
            fileCSVFile.SetDelimiters(",")

            'FileReader = New StreamReader(Files.FullName)
            Do While Not fileCSVFile.EndOfData

                sCurrentRow = fileCSVFile.ReadFields

                If lstMappingCustodianName.Contains(sCurrentRow(0)) Then
                    lstDuplicateMappingCustodianName.Add(sCurrentRow(0))
                    lstDuplicateMappingDestinationFolder.Add(sCurrentRow(1))
                    lstDuplicateMappingDestinationRoot.Add(sCurrentRow(2))
                    lstDuplicateMappingDestinationSMTP.Add(sCurrentRow(3))
                    If sCurrentRow.Count > 4 Then
                        lstDuplicateMappingCustodianGroupID.Add(sCurrentRow(4))
                    End If
                Else
                    lstMappingCustodianName.Add(sCurrentRow(0))
                    lstMappingDestinationFolder.Add(sCurrentRow(1))
                    lstMappingDestinationRoot.Add(sCurrentRow(2))
                    lstMappingDestinationSMTP.Add(sCurrentRow(3))
                    If sCurrentRow.Count > 4 Then
                        lstMappingCustodianGroupID.Add(sCurrentRow(4))
                    End If
                End If
            Loop
            cboGroupIDs.Items.Add("")
            lstUniqueGroupID = lstMappingCustodianGroupID.Distinct.ToList
            For Each UniqueGroup In lstUniqueGroupID
                cboGroupIDs.Items.Add(UniqueGroup.ToString)

            Next

            For Each row In grdPSTInfo.Rows
                If lstMappingCustodianName.Contains(row.cells("CustodianName").Value) Then
                    iCustodianIndex = lstMappingCustodianName.IndexOf(row.cells("CustodianName").Value)
                    row.cells("GroupID").value = lstMappingCustodianGroupID(iCustodianIndex)
                    row.cells("DestinationFolder").value = lstMappingDestinationFolder(iCustodianIndex)
                    row.cells("DestinationRoot").value = lstMappingDestinationRoot(iCustodianIndex)
                    row.cells("DestinationSMTP").value = lstMappingDestinationSMTP(iCustodianIndex)
                End If

            Next

        End If
    End Sub

    Private Sub cboGroupIDs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboGroupIDs.SelectedIndexChanged
        Dim sGroupID As String

        sGroupID = cboGroupIDs.Text
        For Each row In grdPSTInfo.Rows
            If row.cells("GroupID").value = sGroupID Then
                row.cells("SelectCustodian").value = True
            Else
                row.cells("SelectCustodian").value = False
            End If
        Next

    End Sub

    Sub subPSTDirSearch(ByVal sDir As String, ByRef PSTName As List(Of String), ByRef PSTPath As List(Of String), ByRef PSTSize As List(Of String))
        Dim d As String
        Dim Length As String
        Dim currentdirectory As DirectoryInfo
        Dim extension As String

        Dim common As New Common

        Try
            currentdirectory = New DirectoryInfo(sDir)

            For Each File In currentdirectory.GetFiles
                extension = Path.GetExtension(File.ToString)
                If extension = ".pst" Then 'If InStr(File.Name, ".pst") > 0 Then
                    PSTName.Add(File.Name.ToString.ToLower)
                    PSTPath.Add(File.DirectoryName.ToString.ToLower)
                    PSTSize.Add(File.Length)
                End If
            Next

            Length = Directory.GetDirectories(sDir).Length
            If Length > 0 Then
                For Each d In Directory.GetDirectories(sDir)
                    subPSTDirSearch(d, PSTName, PSTPath, PSTSize)
                Next
            End If
        Catch ex As Exception
            common.Logger(psIngestionLogFile, "Error 7 - Directory Search" & ex.ToString)
        End Try

    End Sub

    Private Function blnGetCustodianNameFromFolderContents(ByVal sPSTFolderName As String, ByVal lstPSTName As List(Of String), ByVal lstCustodianNames As List(Of String))
        blnGetCustodianNameFromFolderContents = False

        Dim lstPSTPath As New List(Of String)
        Dim lstPSTSize As New List(Of String)
        Dim lstPSTServerName As New List(Of String)
        Dim sCustodianName As String


        lstPSTPath = New List(Of String)
        lstPSTSize = New List(Of String)
        lstPSTServerName = New List(Of String)

        subPSTDirSearch(sPSTFolderName, lstPSTName, lstPSTPath, lstPSTSize)

        For Each PSTName In lstPSTName
            sCustodianName = PSTName.ToString
            sCustodianName = sCustodianName.Replace(".pst", "")

            If Not (lstCustodianNames.Contains(sCustodianName)) Then
                lstCustodianNames.Add(sCustodianName)
            End If
        Next
        blnGetCustodianNameFromFolderContents = True
    End Function

    Public Function blnMoveSelectedCustodianPST(ByVal lstSelectedCustodianInfo As List(Of String), ByVal sNewFolder As String, ByRef lstInvalidFolder As List(Of String)) As Boolean
        blnMoveSelectedCustodianPST = False
        Dim asSelectedCustodianInfo() As String
        Dim bStatus As Boolean
        Dim sCurrentPSTLocation As String
        Dim sNewCustodianFolder As String
        Dim dbService As New DatabaseService
        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"


        For Each SelectedCustodian In lstSelectedCustodianInfo
            asSelectedCustodianInfo = Split(SelectedCustodian.ToString, ",")
            bStatus = dbService.GetCustodianPSTLocation(sSQLiteDatabaseFullName, asSelectedCustodianInfo(0), sCurrentPSTLocation)

            If sNewFolder.Contains("\\") Then
                sNewFolder = sNewFolder.Replace("\\", "\")
            End If
            If sNewFolder.Substring(sNewFolder.Length - 1) = "\" Then
                sNewCustodianFolder = sNewFolder & asSelectedCustodianInfo(0) & "\"
            Else
                sNewCustodianFolder = sNewFolder & "\" & asSelectedCustodianInfo(0) & "\"
            End If

            If My.Computer.FileSystem.DirectoryExists(sCurrentPSTLocation) Then
                My.Computer.FileSystem.MoveDirectory(sCurrentPSTLocation, sNewCustodianFolder)
                bStatus = blnUpdateCustodianDBPSTLocation(asSelectedCustodianInfo(0), sNewCustodianFolder)
            Else
                lstInvalidFolder.Add(SelectedCustodian)
            End If
        Next

        blnMoveSelectedCustodianPST = True
    End Function
    Public Function blnUpdateCustodianDBPSTLocation(ByVal sCustodianName As String, ByVal sNewFolder As String)

        Dim sUpdateEWSIngestionData As String

        Dim common As Common

        blnUpdateCustodianDBPSTLocation = True

        Using Connection As New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;New=False;Compress=True;")
            sUpdateEWSIngestionData = "Update ewsIngestionStats set PSTPath = @PSTPath"
            sUpdateEWSIngestionData = sUpdateEWSIngestionData & " WHERE CustodianName = @CustodianName"

            Using oUpdateEWSIngestionDataCommand As New SQLiteCommand()
                With oUpdateEWSIngestionDataCommand
                    .Connection = Connection
                    .CommandText = sUpdateEWSIngestionData
                    .Parameters.AddWithValue("@CustodianName", sCustodianName)
                    .Parameters.AddWithValue("@PSTPath", sNewFolder)
                End With
                Try
                    Connection.Open()
                    oUpdateEWSIngestionDataCommand.ExecuteNonQuery()
                    Connection.Close()
                Catch ex As Exception
                    common.Logger(psIngestionLogFile, "Error updating custodian PST Location in the database.")
                    common.Logger(psIngestionLogFile, ex.ToString)
                    blnUpdateCustodianDBPSTLocation = False
                End Try
            End Using
        End Using
    End Function

    Public Function blnMoveCustodianFolder(ByVal sCustodianName As String, ByVal sCustodianFolder As String, ByVal sNewFolder As String, ByRef sInvalidFolder As String) As Boolean
        blnMoveCustodianFolder = False
        Dim bStatus As Boolean

        Dim common As New Common

        Try
            My.Computer.FileSystem.CreateDirectory(sNewFolder)
            'now move the rest of the files in the troot folder
            Dim filesInRoot = From file In Directory.EnumerateFiles(sCustodianFolder, "*.pst")
            For Each filename As String In filesInRoot
                Try
                    Dim file As FileInfo
                    file = New System.IO.FileInfo(filename)
                    IO.File.Move(file.FullName, Path.Combine(sNewFolder, file.Name))
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)

                    ' log or whatever, here just ignore and continue ...
                End Try
            Next
            '            If My.Computer.FileSystem.DirectoryExists(sCustodianFolder) Then
            'My.Computer.FileSystem.MoveDirectory(sCustodianFolder, sNewFolder)
            'bStatus = blnUpdateCustodianDBPSTLocation(sCustodianName, sNewFolder)
            'Else
            'sInvalidFolder = sCustodianFolder
            'End If

        Catch ex As Exception
            common.Logger(psIngestionLogFile, "Error Moving - " & sCustodianFolder & " - to - " & sNewFolder)
            common.Logger(psIngestionLogFile, ex.Message)
        End Try

        blnMoveCustodianFolder = True
    End Function

    '    Public Function blnProcessSelectedCustodians(ByVal sPSTLocation As String, ByVal lstProcessingCustodians As List(Of String), ByVal grdPSTInfo As DataGridView, ByVal iNuixInstances As Integer, ByVal lstCurrentCustodianFolder As List(Of String)) As Boolean
    Public Function blnProcessSelectedCustodians(ByVal lstProcessingCustodians As List(Of String), ByVal grdPSTInfo As DataGridView, ByVal iNuixInstances As Integer, ByVal sNewPSTDirectory As String, ByVal sNuixAppMemory As String) As Boolean
        Dim bStatus As Boolean
        Dim sCustodianName As String
        Dim sCurrentCustodianFolderLocation As String
        Dim sNewCustodianFolderLocation As String
        Dim bFoundProcessedCustodian As Boolean
        Dim NuixConsoleProcess As Process
        Dim NuixConsoleProcessStartInfo As ProcessStartInfo
        Dim iNumberOfNuixInstancesLaunched As Integer
        Dim iNuixInstancesRequested As Integer
        Dim iNumberOfNuixProcessesRunning As Integer
        Dim sCurrentPSTLocation As String
        Dim msgboxReturn As DialogResult
        Dim sInvalidFolder As String
        Dim sPSTDirectory As String
        Dim lstCurrerntCustodianFolder As List(Of String)
        Dim dbService As New DatabaseService

        Dim common As New Common

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"


        Dim dStartDate As Date
        common.Logger(psIngestionLogFile, "Processing Selected Custodians...")

        blnProcessSelectedCustodians = False

        iNuixInstancesRequested = eMailArchiveMigrationManager.O365NuixInstances
        lstCurrerntCustodianFolder = New List(Of String)

        bStatus = blnCheckIfNuixIsRunning("nuix_app", iNumberOfNuixProcessesRunning)
        bStatus = blnCheckIfNuixIsRunning("nuix_console", iNumberOfNuixProcessesRunning)

        If (iNumberOfNuixProcessesRunning >= CInt(eMailArchiveMigrationManager.O365NuixInstances)) Then
            MessageBox.Show("There are already " & iNumberOfNuixProcessesRunning & ". Please wait until processes have finished or manually end processes")
            Exit Function
        ElseIf (iNumberOfNuixProcessesRunning > 0) Then
            msgboxReturn = MessageBox.Show("There are already " & iNumberOfNuixProcessesRunning & " and you have requested to launch " & eMailArchiveMigrationManager.O365NuixInstances & " instances of Nuix. Would you like to launch " & CInt(eMailArchiveMigrationManager.O365NuixInstances) - iNumberOfNuixProcessesRunning & "?", "Nuix Instances Running", MessageBoxButtons.YesNo)
            If msgboxReturn = vbYes Then
                iNuixInstances = CInt(eMailArchiveMigrationManager.O365NuixInstances) - iNumberOfNuixProcessesRunning
            Else
                Exit Function
            End If
        Else
            iNuixInstances = eMailArchiveMigrationManager.O365NuixInstances
        End If

        If grdPSTInfo.RowCount <= iNuixInstances Then
            iNuixInstances = grdPSTInfo.RowCount
        End If
        iNumberOfNuixInstancesLaunched = 0

        For Each custodian In lstProcessingCustodians
            For Each row In grdPSTInfo.Rows
                If row.Cells("CustodianName").value = custodian.ToString Then
                    If File.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\EWS Ingestion\" & custodian.ToString & "\case.fbi2") Then
                        row.Cells("ProgressStatus").value = "Failed (Case Already Exists)"
                        row.cells("StatusImage").value = My.Resources.red_stop_small
                        row.cells("ProcessID").value = ""
                    ElseIf (Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\" & custodian.ToString)) Then
                        row.cells("ProgressStatus").value = "Failed (Export Already Exists)"
                        row.cells("StatusImage").value = My.Resources.red_stop_small
                    ElseIf (File.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\" & custodian.ToString & "\case.fbi2") And (Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\" & custodian.ToString))) Then
                        row.cells("ProgressStatus").value = "Failed (Case and Export Already Exist)"
                        row.cells("StatusImage").value = My.Resources.red_stop_small
                    Else
                        bStatus = blnBuildNuixFilesForCustodians(custodian.ToString, sNuixAppMemory, eMailArchiveMigrationManager.NuixWorkerMemory, eMailArchiveMigrationManager.NuixWorkers, sNewPSTDirectory, row.cells("CaseDirectory").value, row.cells("LogDirectory").value, eMailArchiveMigrationManager.NuixJavaTempDir, eMailArchiveMigrationManager.WorkerTimeout, eMailArchiveMigrationManager.O365RetryCount, eMailArchiveMigrationManager.O365RetryIncrement, eMailArchiveMigrationManager.O365RetryDelay, eMailArchiveMigrationManager.O365FilePathTrimming, eMailArchiveMigrationManager.O365ApplicationImpersonation, eMailArchiveMigrationManager.NuixAppLocation, eMailArchiveMigrationManager.O365ExchangeServer, eMailArchiveMigrationManager.O365Domain, eMailArchiveMigrationManager.O365AdminUserName, eMailArchiveMigrationManager.O365AdminInfo, eMailArchiveMigrationManager.NuixNMSAddress & ":" & eMailArchiveMigrationManager.NuixNMSPort, eMailArchiveMigrationManager.NuixWorkerTempDir, row.cells("OutputDirectory").value, lstCurrerntCustodianFolder, eMailArchiveMigrationManager.NuixNMSUserName, eMailArchiveMigrationManager.NuixNMSAdminInfo)
                        Directory.CreateDirectory(eMailArchiveMigrationManager.NuixExportDir & "\EWS Ingestion\" & custodian.ToString & "\Reprocessing")

                        If iNumberOfNuixInstancesLaunched < iNuixInstances Then
                            bStatus = dbService.GetCustodianPSTLocation(sSQLiteDatabaseFullName, custodian.ToString, sCurrentPSTLocation)
                            bStatus = blnMoveCustodianFolder(custodian.ToString, sCurrentPSTLocation, sNewPSTDirectory, sInvalidFolder)
                            bStatus = blnUpdateCustodianDBPSTLocation(custodian.ToString, sNewPSTDirectory)
                            If row.cells("ProgressStatus").value = "Not Started" Then
                                dStartDate = Now
                                sCustodianName = row.Cells("CustodianName").Value
                                common.Logger(psIngestionLogFile, "Processing Custodian - " & sCustodianName & "...")

                                bStatus = dbService.GetCustodianPSTLocation(sSQLiteDatabaseFullName, sCustodianName, sCurrentCustodianFolderLocation)
                                sNewCustodianFolderLocation = sCurrentCustodianFolderLocation.Replace("Not Started", "In Progress")
                                sNewCustodianFolderLocation = sCurrentCustodianFolderLocation.Replace("not started", "in progress")
                                common.Logger(psIngestionLogFile, "Moving " & sCurrentCustodianFolderLocation & " to " & sNewCustodianFolderLocation & "...")
                                bStatus = blnMoveCustodianFolder(sCustodianName, sCurrentCustodianFolderLocation, sNewCustodianFolderLocation, sInvalidFolder)
                                common.Logger(psIngestionLogFile, "Launching Nuix...")
                                NuixConsoleProcessStartInfo = New ProcessStartInfo(row.cells("ProcessingFilesDirectory").value & "\" & sCustodianName & ".bat")
                                NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                'NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Minimized

                                NuixConsoleProcess = System.Diagnostics.Process.Start(NuixConsoleProcessStartInfo)
                                plstNuixConsoleProcesses.Add(NuixConsoleProcess.Id.ToString)
                                row.cells("StatusImage").value = My.Resources.inprogress_medium
                                row.Cells("SelectCustodian").Value = False
                                row.Cells("ProgressStatus").Value = "In Progress"
                                row.cells("SelectCustodian").readonly = True
                                row.cells("ProgressStatus").readonly = True
                                row.Cells("ProcessID").Value = NuixConsoleProcess.Id
                                row.cells("PSTPath").value = sNewCustodianFolderLocation
                                row.Cells("IngestionStartTime").Value = dStartDate
                                bStatus = dbService.UpdateCustodianIngestionValues(eMailArchiveMigrationManager.SQLiteDBLocation, row.Cells("CustodianName").Value, "ProgressStatus", row.Cells("ProgressStatus").Value)
                                bStatus = dbService.UpdateCustodianDBStartTime(sSQLiteDatabaseFullName, row.cells("CustodianName").value, row.cells("IngestionStartTime").value)
                                bFoundProcessedCustodian = True
                                iNumberOfNuixInstancesLaunched = iNumberOfNuixInstancesLaunched + 1
                                Thread.Sleep(8000)
                            End If
                        Else
                            If row.cells("CustodianName").value <> vbNullString Then
                                If row.cells("SelectCustodian").value = True Then
                                    row.cells("StatusImage").value = My.Resources.waitingtostart1
                                    row.cells("SelectCustodian").value = False
                                    row.cells("ProgressStatus").value = "Not Started"
                                    bStatus = dbService.GetCustodianPSTLocation(sSQLiteDatabaseFullName, row.cells("CustodianName").value, sCurrentPSTLocation)
                                    bStatus = blnMoveCustodianFolder(custodian.ToString, sCurrentPSTLocation, sNewPSTDirectory & "\" & custodian.ToString, sInvalidFolder)
                                    bStatus = dbService.UpdateCustodianIngestionValues(eMailArchiveMigrationManager.SQLiteDBLocation, row.Cells("CustodianName").Value, "ProgressStatus", row.Cells("ProgressStatus").Value)
                                    row.cells("SelectCustodian").readonly = True
                                Else
                                    row.cells("StatusImage").value = My.Resources.not_selected_small
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        Next

        blnProcessSelectedCustodians = True
    End Function

    Public Function blnReProcessSelectedCustodians(ByVal lstReProcessingCustodians As List(Of String), ByVal grdPSTInfo As DataGridView, ByVal iNuixInstances As Integer, ByVal sNewPSTDirectory As String, ByVal sNuixAppMemory As String) As Boolean
        Dim bStatus As Boolean
        Dim sCustodianName As String

        Dim NuixConsoleProcess As Process
        Dim NuixConsoleProcessStartInfo As ProcessStartInfo

        Dim iNuixInstancesRequested As Integer
        Dim iNumberOfNuixProcessesRunning As Integer
        Dim sPSTDirectory As String
        Dim lstCurrerntCustodianFolder As List(Of String)
        Dim sExportDirectory As String
        Dim bProcessIsRunning As Boolean
        Dim sReprocessingFilesDirectory As String
        Dim sReprocessingOutputDirectory As String
        Dim sReprocessingLogDirectory As String
        Dim sReprocessingCaseDirectory As String
        Dim dblTotalExportSize As Double
        Dim common As New Common
        Dim dbService As New DatabaseService

        Dim dStartDate As Date
        Dim dEndDate As Date
        common.Logger(psIngestionLogFile, "ReProcessing Selected Custodians...")

        blnReProcessSelectedCustodians = False

        iNuixInstancesRequested = eMailArchiveMigrationManager.O365NuixInstances
        lstCurrerntCustodianFolder = New List(Of String)

        bStatus = blnCheckIfNuixIsRunning("nuix_app", iNumberOfNuixProcessesRunning)
        bStatus = blnCheckIfNuixIsRunning("nuix_console", iNumberOfNuixProcessesRunning)

        Try
            For Each custodian In lstReProcessingCustodians
                For Each row In grdPSTInfo.Rows
                    If row.Cells("CustodianName").value = custodian.ToString Then
                        sExportDirectory = row.Cells("OutputDirectory").value
                        bStatus = blnComputeFolderSize(sExportDirectory, dblTotalExportSize)
                        bStatus = blnBuildReprocessingFilesForCustodians(custodian.ToString, sExportDirectory, dblTotalExportSize.ToString, sNuixAppMemory, eMailArchiveMigrationManager.NuixWorkerMemory, eMailArchiveMigrationManager.NuixWorkers, sPSTDirectory, row.cells("CaseDirectory").value, row.cells("LogDirectory").value, eMailArchiveMigrationManager.NuixJavaTempDir, eMailArchiveMigrationManager.WorkerTimeout, eMailArchiveMigrationManager.O365RetryCount, eMailArchiveMigrationManager.O365RetryIncrement, eMailArchiveMigrationManager.O365RetryDelay, eMailArchiveMigrationManager.O365FilePathTrimming, eMailArchiveMigrationManager.O365ApplicationImpersonation, eMailArchiveMigrationManager.NuixAppLocation, eMailArchiveMigrationManager.O365ExchangeServer, eMailArchiveMigrationManager.O365Domain, eMailArchiveMigrationManager.O365AdminUserName, eMailArchiveMigrationManager.O365AdminInfo, eMailArchiveMigrationManager.NuixNMSAddress & ":" & eMailArchiveMigrationManager.NuixNMSPort, eMailArchiveMigrationManager.NuixWorkerTempDir, row.cells("OutputDirectory").value, lstCurrerntCustodianFolder, eMailArchiveMigrationManager.NuixNMSUserName, eMailArchiveMigrationManager.NuixNMSAdminInfo, row.cells("ReprocessingDirectory").value & "\AllFailures.log")
                        If row.cells("ProgressStatus").value = "Not Started Items Not Uploaded" Then
                            dStartDate = Now
                            sCustodianName = row.Cells("CustodianName").Value
                            row.cells("StatusImage").value = My.Resources.inprogress_medium
                            row.Cells("SelectCustodian").Value = False
                            row.cells("ProgressStatus").value = "In Progress Items Not Uploaded"
                            row.cells("SelectCustodian").readonly = True
                            row.cells("ProgressStatus").readonly = True
                            row.Cells("IngestionStartTime").Value = dStartDate
                            common.Logger(psIngestionLogFile, "ReProcessing Custodian - " & sCustodianName & "...")
                            bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "ProcessingFilesDirectory", sReprocessingFilesDirectory, "TEXT")
                            bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "OutputDirectory", sReprocessingOutputDirectory, "TEXT")
                            bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "LogDirectory", sReprocessingLogDirectory, "TEXT")
                            bStatus = dbService.GetUpdatedIngestionDBInfo(eMailArchiveMigrationManager.SQLiteDBLocation, sCustodianName, "CaseDirectory", sReprocessingCaseDirectory, "TEXT")

                            sReprocessingFilesDirectory = sReprocessingFilesDirectory & "\Reprocessing\"
                            sReprocessingLogDirectory = sReprocessingLogDirectory & "\Reprocessing\"
                            sReprocessingCaseDirectory = sReprocessingCaseDirectory & "\Reprocessing\"
                            sReprocessingOutputDirectory = sReprocessingOutputDirectory & "\Reprocessing\"


                            common.Logger(psIngestionLogFile, "Launching Nuix...")
                            NuixConsoleProcessStartInfo = New ProcessStartInfo(sReprocessingFilesDirectory & "\" & sCustodianName & ".bat")
                            NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            'NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Minimized

                            NuixConsoleProcess = System.Diagnostics.Process.Start(NuixConsoleProcessStartInfo)
                            row.Cells("ProcessID").Value = NuixConsoleProcess.Id
                            plstNuixConsoleProcesses.Add(NuixConsoleProcess.Id.ToString)
                            Thread.Sleep(8000)
                            bProcessIsRunning = True
                            While bProcessIsRunning = True
                                bProcessIsRunning = blnCheckIfProcessIsRunning(NuixConsoleProcess.Id.ToString)
                                Thread.Sleep(10000)
                            End While
                            row.cells("ProgressStatus").value = "Completed Processing Items Not Uploaded"
                            row.cells("SelectCustodian").readonly = False
                            row.cells("ProgressStatus").readonly = False
                            row.Cells("ProcessID").Value = ""
                            dEndDate = Now
                            row.Cells("IngestionEndTime").Value = dEndDate
                        End If
                    End If
                Next
            Next

        Catch ex As Exception
            common.Logger(psIngestionLogFile, ex.ToString)
        End Try

        blnReProcessSelectedCustodians = True

        MessageBox.Show("All Reprocessing jobs have completed", "Reprocessing Jobs Completed", MessageBoxButtons.OK)
        btnProcessandRunSelectedUsers.Enabled = True
        btnLoadPreviousConfig.Enabled = True
        btnProcessandRunSelectedUsers.Enabled = True


    End Function
    Public Function blnProcessNotStartedCustodians(ByVal grdPSTInfo As DataGridView, ByVal iNuixInstances As Integer, sCaseDir As String) As Boolean
        Dim iCounter As Integer
        Dim bStatus As Boolean
        Dim sCustodianName As String
        Dim sNewCustodianFolderLocation As String
        Dim bFoundProcessedCustodian As Boolean
        Dim lstSelectedCustodians As List(Of String)
        Dim sInvalidFolder As String
        Dim NuixConsoleProcess As Process
        Dim NuixConsoleProcessStartInfo As ProcessStartInfo
        Dim sCurrentCustodianFolderLocation As String
        Dim dbService As New DatabaseService
        Dim common As New Common

        lstSelectedCustodians = New List(Of String)

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"


        Dim dStartDate As Date

        blnProcessNotStartedCustodians = False
        common.Logger(psIngestionLogFile, "Processing Not Started Custodians...")

        If grdPSTInfo.RowCount - 1 <= iNuixInstances Then
            iNuixInstances = grdPSTInfo.RowCount - 1
        End If

        For iCounter = 0 To iNuixInstances
            bFoundProcessedCustodian = False

            For Each row In grdPSTInfo.Rows
                If row.cells("ProgressStatus").value = "Not Started" Then
                    If lstSelectedCustodians.Count < iNuixInstances Then
                        lstSelectedCustodians.Add(row.cells("CustodianName").value)
                        dStartDate = Now
                        sCustodianName = row.Cells("CustodianName").Value
                        If Directory.Exists(sCaseDir & "\" & sCustodianName) Then
                            row.Cells("ProcessID").value = "Case Already Exists"
                        ElseIf (Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\" & sCustodianName)) Then
                            row.cells("ProcessID").value = "Export Directory Exists"
                        ElseIf (Directory.Exists(sCaseDir & "\" & sCustodianName) And (Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\" & sCustodianName))) Then
                            row.cells("ProcessID").value = "Case Directory and Export Directory Exists"
                        Else
                            common.Logger(psIngestionLogFile, "Processing " & sCustodianName & "...")
                            bStatus = dbService.GetCustodianPSTLocation(sSQLiteDatabaseFullName, row.cells("CustodianName").value, sCurrentCustodianFolderLocation)
                            sNewCustodianFolderLocation = sCurrentCustodianFolderLocation.Replace("Not Started", "In Progress")
                            sNewCustodianFolderLocation = sCurrentCustodianFolderLocation.Replace("not started", "in progress")
                            common.Logger(psIngestionLogFile, "Moving " & sCurrentCustodianFolderLocation & " to " & sNewCustodianFolderLocation & "...")
                            bStatus = blnMoveCustodianFolder(sCustodianName, sCurrentCustodianFolderLocation, sNewCustodianFolderLocation, sInvalidFolder)
                            common.Logger(psIngestionLogFile, "Launching Nuix...")
                            NuixConsoleProcessStartInfo = New ProcessStartInfo(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Ingestion\Nuix Files\" & sCustodianName & "\" & sCustodianName & ".bat")
                            NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            'NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Minimized

                            NuixConsoleProcess = System.Diagnostics.Process.Start(NuixConsoleProcessStartInfo)
                            plstNuixConsoleProcesses.Add(NuixConsoleProcess.Id.ToString)
                            row.cells("StatusImage").value = My.Resources.inprogress_medium
                            row.Cells("SelectCustodian").Value = False
                            row.cells("PSTPath").value = sNewCustodianFolderLocation
                            row.cells("SelectCustodian").readonly = True
                            row.Cells("ProgressStatus").Value = "In Progress"
                            row.Cells("ProcessID").Value = NuixConsoleProcess.Id
                            row.Cells("IngestionStartTime").Value = dStartDate
                            bStatus = dbService.UpdateCustodianIngestionValues(eMailArchiveMigrationManager.SQLiteDBLocation, row.Cells("CustodianName").Value, "ProgressStatus", row.Cells("ProgressStatus").Value)
                            bStatus = dbService.UpdateCustodianDBStartTime(sSQLiteDatabaseFullName, row.cells("CustodianName").value, row.cells("IngestionStartTime").value)
                            bFoundProcessedCustodian = True
                        End If
                    Else
                        row.cells("SelectCustodian").value = False
                        row.cells("ProgressStatus").value = "Not Started"
                        row.cells("SelectCustodian").readonly = True
                    End If
                End If
            Next
        Next

        blnProcessNotStartedCustodians = True
    End Function

    Public Function blnProcessNextJob(ByVal sCustodianName As String, ByVal sCurrentCustodianFolderLocation As String, ByVal sNewCustodianFolderLocation As String, ByVal grdPSTInfo As DataGridView) As Boolean
        blnProcessNextJob = False
        Dim bStatus As Boolean
        Dim dStartDate As Date
        Dim sInvalidFolder As String
        Dim NuixConsoleProcess As Process
        Dim NuixConsoleProcessStartInfo As ProcessStartInfo
        Dim dbService As New DatabaseService
        Dim common As New Common

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"


        common.Logger(psIngestionLogFile, "Processing " & sCustodianName & "...")
        common.Logger(psIngestionLogFile, "Moving " & sCurrentCustodianFolderLocation & " to " & sNewCustodianFolderLocation & "...")
        bStatus = blnMoveCustodianFolder(sCustodianName, sCurrentCustodianFolderLocation, sNewCustodianFolderLocation, sInvalidFolder)
        common.Logger(psIngestionLogFile, "Launching Nuix...")
        NuixConsoleProcessStartInfo = New ProcessStartInfo(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Ingestion\Nuix Files\" & sCustodianName & "\" & sCustodianName & ".bat")
        NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
        'NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Minimized

        NuixConsoleProcess = System.Diagnostics.Process.Start(NuixConsoleProcessStartInfo)
        plstNuixConsoleProcesses.Add(NuixConsoleProcess.Id.ToString)
        For Each row In grdPSTInfo.Rows
            If row("CustodianName").Value = sCustodianName Then
                row.cells("StatusImage").value = My.Resources.inprogress_medium
                row.cells("SelectCustodian").value = False
                row.cells("ProgressStatus").value = "In Progress"
                row.cells("SelectCustodian").readonly = True
                row.cells("ProcessID").value = NuixConsoleProcess.Id
                dStartDate = Now
                row.cells("IngestionStartTime").value = dStartDate
                bStatus = dbService.UpdateCustodianIngestionValues(eMailArchiveMigrationManager.SQLiteDBLocation, row.Cells("CustodianName").Value, "ProgressStatus", row.Cells("ProgressStatus").Value)
                bStatus = dbService.UpdateCustodianDBStartTime(sSQLiteDatabaseFullName, row.Cells("CustodianName").value, row.cells("IngestionStartTime").value)
            End If
        Next

        blnProcessNextJob = True
    End Function

    Private Sub btnLoadPreviousConfig_Click(sender As Object, e As EventArgs) Handles btnLoadPreviousConfig.Click
        Dim OpenFileDialog1 As New OpenFileDialog
        Dim iNumberOfNuixInstancesRunning As Integer
        Dim dlgReturn As DialogResult

        Dim lstActiveProcesses As List(Of String)
        Dim lstInProgressCustodians As List(Of String)
        Dim lstNotStartedCustodians As List(Of String)
        Dim bStatus As Boolean
        Dim common As New Common

        lstActiveProcesses = New List(Of String)
        lstInProgressCustodians = New List(Of String)
        lstNotStartedCustodians = New List(Of String)
        grdPSTInfo.Rows.Clear()

        common.Logger(psIngestionLogFile, "Loading previous configuration from database")
        bStatus = blnLoadGridFromDB(grdPSTInfo, lstNotStartedCustodians, lstInProgressCustodians)
        If grdPSTInfo.RowCount > 1 Then
            btnProcessingDetails.Enabled = True
        Else
            btnProcessingDetails.Enabled = False
        End If
        If lstNotStartedCustodians.Count > 0 Then
            bStatus = blnCheckIfNuixIsRunning("nuix_app", iNumberOfNuixInstancesRunning)
            bStatus = blnCheckIfNuixIsRunning("nuix_console", iNumberOfNuixInstancesRunning)
            If (iNumberOfNuixInstancesRunning < eMailArchiveMigrationManager.O365NuixInstances) Then
                dlgReturn = MessageBox.Show("You have requested - " & eMailArchiveMigrationManager.O365NuixInstances & " - instances of Nuix.  There are - " & iNumberOfNuixInstancesRunning & " instances running - would you like to launch - " & eMailArchiveMigrationManager.O365NuixInstances - iNumberOfNuixInstancesRunning & "?", "Launch Nuix Instances", MessageBoxButtons.YesNo)
                If (dlgReturn = vbYes) Then
                    bStatus = blnCheckIfNuixIsRunning("nuix_app", iNumberOfNuixInstancesRunning)
                    bStatus = blnCheckIfNuixIsRunning("nuix_console", iNumberOfNuixInstancesRunning)
                    If (iNumberOfNuixInstancesRunning < eMailArchiveMigrationManager.O365NuixInstances) Then
                        common.Logger(psIngestionLogFile, "Starting " & eMailArchiveMigrationManager.O365NuixInstances - iNumberOfNuixInstancesRunning & " instances of Nuix...")
                        bStatus = blnProcessNotStartedCustodians(grdPSTInfo, eMailArchiveMigrationManager.O365NuixInstances - iNumberOfNuixInstancesRunning, eMailArchiveMigrationManager.NuixCaseDir)

                        common.Logger(psIngestionLogFile, "Watching directory " & FileSystemWatcher1.Path)
                        SummaryReportThread = New System.Threading.Thread(AddressOf WatchCaseDirectory)
                        SummaryReportThread.Start()

                        common.Logger(psIngestionLogFile, "Polling SQLite database...")
                        SQLiteDBReadThread = New System.Threading.Thread(AddressOf SQLiteUpdates)
                        SQLiteDBReadThread.Start()
                    End If
                End If
            End If
        End If

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub grdPSTInfo_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdPSTInfo.CellContentDoubleClick

        Try
            If ((e.ColumnIndex = 16) And (grdPSTInfo("ProcessingFilesDirectory", e.RowIndex).Value <> vbNullString)) Then
                If Directory.Exists(grdPSTInfo("ProcessingFilesDirectory", e.RowIndex).Value) Then
                    Process.Start("explorer.exe", grdPSTInfo("ProcessingFilesDirectory", e.RowIndex).Value)
                End If
            ElseIf ((e.ColumnIndex = 17) And (grdPSTInfo("CaseDirectory", e.RowIndex).Value <> vbNullString)) Then
                If Directory.Exists(grdPSTInfo("CaseDirectory", e.RowIndex).Value) Then
                    Process.Start("explorer.exe", grdPSTInfo("CaseDirectory", e.RowIndex).Value)
                End If
            ElseIf ((e.ColumnIndex = 18) And (grdPSTInfo("OutputDirectory", e.RowIndex).Value <> vbNullString)) Then
                If Directory.Exists(grdPSTInfo("OutputDirectory", e.RowIndex).Value) Then
                    Process.Start("explorer.exe", grdPSTInfo("OutputDirectory", e.RowIndex).Value)
                End If

            ElseIf ((e.ColumnIndex = 19) And (grdPSTInfo("ReprocessingDirectory", e.RowIndex).Value <> vbNullString)) Then
                If Directory.Exists(grdPSTInfo("OutputDirectory", e.RowIndex).Value) Then
                    Process.Start("explorer.exe", grdPSTInfo("OutputDirectory", e.RowIndex).Value)
                End If
            ElseIf ((e.ColumnIndex = 20) And (grdPSTInfo("LogDirectory", e.RowIndex).Value <> vbNullString)) Then
                If Directory.Exists(grdPSTInfo("LogDirectory", e.RowIndex).Value) Then
                    Process.Start("explorer.exe", grdPSTInfo("LogDirectory", e.RowIndex).Value)
                End If
            ElseIf ((e.ColumnIndex = 23) And (grdPSTInfo("SummaryReportLocation", e.RowIndex).Value <> vbNullString)) Then
                System.Diagnostics.Process.Start(grdPSTInfo("SummaryReportLocation", e.RowIndex).Value.ToString)
            ElseIf ((e.ColumnIndex = 15) And (grdPSTInfo("IngestionEndTime", e.RowIndex).Value <> vbNullString)) Then
                If grdPSTInfo("NumberFailed", e.RowIndex).Value <> vbNullString Then
                    If (grdPSTInfo("NumberFailed", e.RowIndex).Value > 0) Then
                        psSummaryReportFile = grdPSTInfo(18, e.RowIndex).Value
                        Dim FailuresForm As FailuresForm
                        FailuresForm = New FailuresForm
                        FailuresForm.CustodianSummaryReportFile = psSummaryReportFile
                        FailuresForm.Text = grdPSTInfo("CustodianName", e.RowIndex).Value
                        FailuresForm.ShowDialog()
                    End If

                End If
            End If

        Catch ex As Exception

        End Try
    End Sub


    Private Sub btnClearGrid_Click(sender As Object, e As EventArgs) Handles btnClearGrid.Click
        grdPSTInfo.Rows.Clear()
    End Sub

    Private Function blnCheckIfNuixIsRunning(ByVal sProcessName As String, ByRef iNumberOfNuixInstancesRunning As Integer) As Boolean

        Dim NuixProcess() As System.Diagnostics.Process

        blnCheckIfNuixIsRunning = False

        NuixProcess = Process.GetProcessesByName(sProcessName)

        If NuixProcess.Length > 0 Then
            iNumberOfNuixInstancesRunning = iNumberOfNuixInstancesRunning + NuixProcess.Length
        End If

        blnCheckIfNuixIsRunning = True

    End Function

    Public Function blnProcessSummaryReportFiles(ByVal sSummaryReportFile As String, ByRef lstUniqueSummaryErrors As List(Of String), ByRef lstUniqueSummaryErrorsCount As List(Of Double)) As Boolean
        Dim sCurrentRow As String
        Dim asErrorArray() As String
        Dim fileSummaryFile As StreamReader
        Dim iCounter As Integer
        Dim bFailedItemStart As Boolean

        Dim iErrorIndex As Integer
        Dim iArrayCount As Integer
        Dim dblCurrentCount As Integer
        Dim sCurrentErrorDetail As String
        Dim common As New Common

        blnProcessSummaryReportFiles = True

        Try
            iCounter = iCounter + 1
            bFailedItemStart = False
            fileSummaryFile = New StreamReader(sSummaryReportFile)
            common.Logger(psIngestionLogFile, "Reading " & sSummaryReportFile)
            While Not fileSummaryFile.EndOfStream

                sCurrentRow = fileSummaryFile.ReadLine
                If (sCurrentRow.Contains("Failed Item Statistics")) Then
                    bFailedItemStart = True
                ElseIf (sCurrentRow.Contains("File Type Statistics")) Then
                    bFailedItemStart = False
                End If

                If (bFailedItemStart = True) Then

                    If sCurrentRow.Contains("GUID:") Then
                        asErrorArray = Split(sCurrentRow, ":")
                        iArrayCount = asErrorArray.Count
                        sCurrentErrorDetail = asErrorArray(iArrayCount - 1)
                        If (sCurrentErrorDetail.Length <= 2) Then
                            sCurrentErrorDetail = asErrorArray(iArrayCount - 2)
                        End If
                        If lstUniqueSummaryErrors.Contains(sCurrentErrorDetail) Then
                            iErrorIndex = lstUniqueSummaryErrors.IndexOf(sCurrentErrorDetail)
                            dblCurrentCount = lstUniqueSummaryErrorsCount(iErrorIndex)
                            lstUniqueSummaryErrorsCount(iErrorIndex) = dblCurrentCount + 1
                        Else
                            lstUniqueSummaryErrors.Add(sCurrentErrorDetail)
                            lstUniqueSummaryErrorsCount.Add(1)
                        End If
                    End If
                End If

            End While

            blnProcessSummaryReportFiles = True
        Catch ex As Exception
            common.Logger(psIngestionLogFile, ex.Message)
            blnProcessSummaryReportFiles = False
        End Try

    End Function

    Public Function blnConsolidateSummaryErrors(ByRef lstUniqueSummaryErrors As List(Of String), ByRef lstUniqueSummaryErrorsCount As List(Of Double), ByRef lstConsolidatedSummaryErrors As List(Of String), ByRef lstConsolidatedSummaryErrorCount As List(Of Integer), ByRef lstConsolidatedOtherErrors As List(Of String), ByRef lstConsolidatedOtherErrorCount As List(Of Integer))

        Dim sCurrentError As String

        Dim iErrorIndex As Integer

        For Each sError In lstUniqueSummaryErrors

            If sError.Contains("Skipping item for") Or sError.Contains("No mapping was found") Then
                If lstConsolidatedSummaryErrors.Contains(sError) Then

                    iErrorIndex = lstConsolidatedSummaryErrors.IndexOf(sError)
                    lstConsolidatedSummaryErrorCount(iErrorIndex) = lstUniqueSummaryErrorsCount(iErrorIndex)
                Else

                    lstConsolidatedSummaryErrors.Add(sError)
                    lstConsolidatedSummaryErrorCount.Add(lstUniqueSummaryErrorsCount(lstUniqueSummaryErrors.IndexOf(sError)))
                End If
            Else
                If lstConsolidatedOtherErrors.Contains(sError) Then

                    iErrorIndex = lstConsolidatedOtherErrors.IndexOf(sError)
                    lstConsolidatedOtherErrorCount(iErrorIndex) = lstUniqueSummaryErrorsCount(iErrorIndex)
                Else

                    lstConsolidatedOtherErrors.Add(sCurrentError)
                    lstConsolidatedOtherErrorCount.Add(lstUniqueSummaryErrorsCount(lstUniqueSummaryErrors.IndexOf(sError)))
                End If
            End If
        Next

        blnConsolidateSummaryErrors = True

    End Function

    Private Sub btnRefreshGridData_Click(sender As Object, e As EventArgs) Handles btnRefreshGridData.Click
        Dim bStatus As Boolean

        Dim dlgReturn As DialogResult
        Dim iNumberOfNuixInstancesRunning As Integer
        Dim lstSummaryReportFiles As List(Of String)
        Dim sCSVFileLocation As String
        Dim sPSTFolderName As String
        Dim msgboxReturn As DialogResult
        Dim lstSummaryCustodians As List(Of String)
        Dim lstSummaryItemsExported As List(Of String)
        Dim lstSummaryItemsFailed As List(Of String)
        Dim lstSummaryBytesUploaded As List(Of String)
        Dim lstSummaryStartDate As List(Of String)
        Dim lstSummaryEndDate As List(Of String)
        Dim lstSummaryReports As List(Of String)
        Dim lstInProgressCustodians As List(Of String)
        Dim lstNotStartedCustodians As List(Of String)
        Dim lstSelectedCustodiansInfo As List(Of String)
        Dim lstProcessingCustodians As List(Of String)
        Dim lstRestartSelectedCustodianInfo As List(Of String)
        Dim lstRestartProcessingCustodian As List(Of String)
        Dim lstInvalidFolder As List(Of String)

        lstSummaryReportFiles = New List(Of String)
        lstSummaryCustodians = New List(Of String)
        lstSummaryItemsExported = New List(Of String)
        lstSummaryItemsFailed = New List(Of String)
        lstSummaryBytesUploaded = New List(Of String)
        lstSummaryStartDate = New List(Of String)
        lstSummaryEndDate = New List(Of String)
        lstSummaryReports = New List(Of String)
        lstInProgressCustodians = New List(Of String)
        lstNotStartedCustodians = New List(Of String)
        lstSelectedCustodiansInfo = New List(Of String)
        lstProcessingCustodians = New List(Of String)
        lstRestartProcessingCustodian = New List(Of String)
        lstRestartSelectedCustodianInfo = New List(Of String)
        lstInvalidFolder = New List(Of String)

        If eMailArchiveMigrationManager.NuixCaseDir = vbNullString Then
            msgboxReturn = MessageBox.Show("You must choose a Case Directory to refresh the data from. Would you like to show the settings dialog?", "No Case Directory Selected", MessageBoxButtons.YesNo)
            If msgboxReturn = vbYes Then
                Dim SettingsForm As O365ExtractionSettings
                SettingsForm = New O365ExtractionSettings
                SettingsForm.ShowDialog()
                Exit Sub
            Else
                Exit Sub
            End If
            Exit Sub
        End If
        If grdPSTInfo.RowCount <= 0 Then
            MessageBox.Show("You must load the appropriate Custodian Data into the grid.")
            Exit Sub
        End If

        msgboxReturn = MessageBox.Show("Would you like to refresh the information from the Nuix eMail Archive Manager DB? (No will refresh from the file system)?", "Refresh Custodian Information", MessageBoxButtons.YesNo)
        If (msgboxReturn = vbYes) Then
            bStatus = blnRefreshGridFromDB(grdPSTInfo, lstNotStartedCustodians, lstInProgressCustodians)
            bStatus = blnCheckIfNuixIsRunning("nuix_app", iNumberOfNuixInstancesRunning)
            bStatus = blnCheckIfNuixIsRunning("nuix_console", iNumberOfNuixInstancesRunning)
            If (iNumberOfNuixInstancesRunning < eMailArchiveMigrationManager.O365NuixInstances) Then
                dlgReturn = MessageBox.Show("You have requested - " & eMailArchiveMigrationManager.O365NuixInstances & " - instances of Nuix.  There are - " & iNumberOfNuixInstancesRunning & " instances running - would you like to launch - " & eMailArchiveMigrationManager.O365NuixInstances - iNumberOfNuixInstancesRunning & "?", "Launch Nuix Instances", MessageBoxButtons.YesNo)
                If (dlgReturn = vbYes) Then
                    bStatus = blnCheckIfNuixIsRunning("nuix_app", iNumberOfNuixInstancesRunning)
                    bStatus = blnCheckIfNuixIsRunning("nuix_console", iNumberOfNuixInstancesRunning)
                    If (iNumberOfNuixInstancesRunning < eMailArchiveMigrationManager.O365NuixInstances) Then
                        bStatus = blnProcessNotStartedCustodians(grdPSTInfo, eMailArchiveMigrationManager.O365NuixInstances - iNumberOfNuixInstancesRunning, eMailArchiveMigrationManager.NuixCaseDir)
                        SummaryReportThread = New System.Threading.Thread(AddressOf WatchCaseDirectory)
                        SummaryReportThread.Start()

                        SQLiteDBReadThread = New System.Threading.Thread(AddressOf SQLiteUpdates)
                        SQLiteDBReadThread.Start()
                    End If
                End If
            End If

        Else
            sPSTFolderName = txtPSTLocation.Text

            If (sCSVFileLocation = vbNullString And sPSTFolderName = vbNullString) Then
                MessageBox.Show("You must select the location of the PST Files to refresh the grid from the file system")
                txtPSTLocation.Focus()
                Exit Sub
            End If
        End If

        ' set the path to be watched
        FileSystemWatcher1.Path = psNuixCaseDir

        ' add event handlers
        AddHandler FileSystemWatcher1.Created, AddressOf OnCreated

    End Sub
    Private Function blnRefreshGridFromDB(ByVal grdPSTInfo As DataGridView, ByRef lstNotStartedCustodians As List(Of String), ByRef lstInProgressCustodians As List(Of String)) As Boolean
        Dim mSQL As String
        Dim dt As DataTable
        Dim ds As DataSet
        Dim dataReader As SQLiteDataReader
        Dim sqlCommand As SQLiteCommand
        Dim sqlConnection As SQLiteConnection
        Dim sCustodianName As String
        Dim sProgressStatus As String
        Dim dPercentFailed As Decimal
        Dim bStatus As Boolean
        Dim sSummaryReportLocation As String
        Dim sSummaryReportCustodian As String
        Dim sSummaryReportItemsExported As String
        Dim sSummaryReportItemsFailed As String
        Dim sSummaryReportBytesUploaded As String
        Dim dSummaryReportStartDate As Date
        Dim dSummaryReportEndDate As Date
        Dim common As New Common

        blnRefreshGridFromDB = False

        dt = Nothing
        ds = New DataSet
        common.Logger(psIngestionLogFile, "Refreshing Data grid from database")
        sqlConnection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")

        For Each row In grdPSTInfo.Rows
            sCustodianName = row.cells("CustodianName").value
            mSQL = "select CustodianName, ProgressStatus, PSTPath, NumberOfPSTs, TotalSizeOfPSTs, GroupID, DestinationFolder, DestinationRoot, DestinationSMTP, ProcessID, IngestionStartTime, IngestionEndTime, BytesUploaded, PercentCompleted, Success, Failed, SummaryReportLocation from ewsIngestionStats where CustodianName = '" & sCustodianName & "'"
            sqlCommand = New SQLiteCommand(mSQL, sqlConnection)
            sqlConnection.Open()

            dataReader = sqlCommand.ExecuteReader

            While dataReader.Read

                sProgressStatus = vbNullString

                row.cells("CustodianName").value = dataReader.GetValue(0)
                row.cells("ProgressStatus").value = dataReader.GetValue(1)
                row.cells("PSTPath").value = dataReader.GetValue(2)
                row.cells("NumberOfPSTs").value = dataReader.GetValue(3)
                row.cells("SizeOfPSTs").value = FormatNumber(dataReader.GetValue(4), , , , TriState.True).ToString.Replace(".00", "")
                row.cells("GroupID").value = dataReader.GetValue(5)
                row.cells("DestinationFolder").value = dataReader.GetValue(6)
                row.cells("DestinationRoot").value = dataReader.GetValue(7)
                row.cells("DestinationSMTP").value = dataReader.GetValue(8)
                row.cells("ProcessID").value = dataReader.GetValue(9)
                row.cells("IngestionStartTime").value = dataReader.GetValue(10)
                row.cells("IngestionEndTime").value = dataReader.GetValue(11)
                row.cells("BytesUploaded").value = FormatNumber(dataReader.GetValue(12), , , , TriState.True).ToString.Replace(".00", "")
                row.cells("PercentageCompleted").value = dataReader.GetValue(13).ToString.Replace(".00", "")
                row.cells("NumberSuccess").value = dataReader.GetValue(14)
                row.cells("NumberFailed").value = dataReader.GetValue(15)
                row.cells("SummaryReportLocation").value = dataReader.GetValue(16)

                If (dataReader.IsDBNull(16)) Then
                    If File.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\" & sCustodianName & "\case.lock") Then
                        bStatus = blnCheckIfProcessIsRunning(dataReader.GetValue(11))
                        If bStatus = False Then
                            sProgressStatus = "Process Terminated"
                            bStatus = blnUpdateSQLiteDBCustodianInfo(sCustodianName, dataReader.GetValue(1), "", "", 0, 0, "")
                        Else
                            sProgressStatus = "In Progress"
                            bStatus = blnUpdateSQLiteDBCustodianInfo(sCustodianName, row.cells("PSTPath").value, "", "", 0, 0, "")
                        End If
                    End If
                End If
                If File.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\" & sCustodianName & "\summary-report.txt") Then
                    sProgressStatus = "Completed"
                    bStatus = blnGetSummaryReportInfo(sSummaryReportLocation, sSummaryReportCustodian, sSummaryReportItemsExported, sSummaryReportItemsFailed, sSummaryReportBytesUploaded, dSummaryReportStartDate, dSummaryReportEndDate)
                    bStatus = blnUpdateSQLiteDBCustodianInfo(sCustodianName, row.cells("PSTPath").value, "", dSummaryReportEndDate, sSummaryReportItemsExported.Replace(",", ""), sSummaryReportItemsFailed.Replace(",", ""), sSummaryReportLocation)
                End If

                Select Case sProgressStatus
                    Case "Not Started"
                        row.cells("StatusImage").value = My.Resources.waitingtostart1
                        row.cells("SelectCustodian").readonly = True
                    Case "In Progress"
                        row.cells("StatusImage").value = My.Resources.inprogress_medium
                        row.cells("SelectCustodian").readonly = True
                    Case "Completed"
                        If (CDbl(dataReader.GetValue(16)) > 0) Then
                            dPercentFailed = (CDbl(dataReader.GetValue(17)) / CDbl(dataReader.GetValue(16)))
                        Else
                            dPercentFailed = 1.0
                        End If

                        If (dPercentFailed > 0.25) Then
                            row.DefaultCellStyle.Forecolor = Color.Red
                            row.cells("StatusImage").value = My.Resources.red_stop_small
                            row.cells("SelectCustodian").readonly = True
                        ElseIf (dPercentFailed > 0.1) Then
                            row.DefaultCellStyle.Forecolor = Color.Orange
                            row.cells("StatusImage").value = My.Resources.yellow_info_small
                            row.cells("SelectCustodian").readonly = True
                        Else
                            row.defaultcellstyle.Forecolor = Color.Green
                            row.cells("StatusImage").value = My.Resources.Green_check_small
                            row.cells("SelectCustodian").readonly = True
                        End If
                    Case Else
                        row.cells("StatusImage").value = My.Resources.not_selected_small
                End Select

            End While
            sqlConnection.Close()

        Next
        blnRefreshGridFromDB = True
    End Function

    Private Function blnLoadGridFromDB(ByVal grdPSTInfo As DataGridView, ByRef lstNotStartedCustodians As List(Of String), ByRef lstInProgressCustodians As List(Of String)) As Boolean
        Dim mSQL As String
        Dim dt As DataTable
        Dim ds As DataSet
        Dim dataReader As SQLiteDataReader
        Dim sqlCommand As SQLiteCommand
        Dim sqlConnection As SQLiteConnection
        Dim sProgressStatus As String
        Dim iCustodianNameCol As Integer
        Dim iPSTPathCol As Integer
        Dim iNumberOfPSTsCol As Integer
        Dim iTotalSizeOfPSTsCol As Integer
        Dim iGroupIDCol As Integer
        Dim iDestinationFolderCol As Integer
        Dim iDestinationRootCol As Integer
        Dim iDestinationSMTPCol As Integer
        Dim iProcessIDCol As Integer
        Dim iIngestionStartTimeCol As Integer
        Dim iIngestionEndTimeCol As Integer
        Dim iBytesUploadedCol As Integer
        Dim iPercentCompletedCol As Integer
        Dim iSuccessCol As Integer
        Dim iFailedCol As Integer
        Dim iCaseDirectoryCol As Integer
        Dim iOutputDirectoryCol As Integer
        Dim iLogDirectoryCol As Integer
        Dim iSummaryReportLocationCol As Integer
        Dim iProgressStatusCol As Integer
        Dim CustodianRow As DataGridViewRow

        Dim sCustodianName As String
        Dim sPSTPath As String
        Dim iNumberOfPSTs As Integer
        Dim dblTotalSizeOfPSTs As Double
        Dim sTotalSizeOfPSTs As String
        Dim sGroupID As String
        Dim sDestinationFolder As String
        Dim sDestinationRoot As String
        Dim sDestinationSMTP As String
        Dim sProcessID As String
        Dim sIngestionStartTime As String
        Dim sIngestionEndTime As String
        Dim dblBytesUploaded As Double
        Dim sBytesUploaded As String
        Dim iPercentCompleted As Integer
        Dim sPercentCompleted As String
        Dim dblSuccess As Integer
        Dim sSuccess As String
        Dim sCaseDirectory As String
        Dim sOutputDirectory As String
        Dim sLogDirectory As String
        Dim dblFailed As Integer
        Dim dSuccessPercent As Decimal
        Dim sFailed As String
        Dim sSummaryReportLocation As String

        Dim iRowIndex As Integer
        Dim sProcessingFilesDirectory As String
        Dim iProcessingFilesDirectoryCol As Integer

        Dim lstDBGroups As List(Of String)
        Dim lstUniqueGroups As List(Of String)

        Dim iCounter As Integer

        blnLoadGridFromDB = False

        lstDBGroups = New List(Of String)
        lstUniqueGroups = New List(Of String)

        cboGroupIDs.Items.Clear()
        dt = Nothing
        ds = New DataSet
        sqlConnection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")

        mSQL = "select CustodianName, ProgressStatus, PSTPath, NumberOfPSTs, TotalSizeOfPSTs, GroupID, DestinationFolder, DestinationRoot, DestinationSMTP, ProcessID, IngestionStartTime, IngestionEndTime, BytesUploaded, PercentCompleted, Success, Failed, CaseDirectory, OutputDirectory, LogDirectory, ProcessingFilesDirectory, SummaryReportLocation from ewsIngestionStats "
        sqlCommand = New SQLiteCommand(mSQL, sqlConnection)
        sqlConnection.Open()

        dataReader = sqlCommand.ExecuteReader
        grdPSTInfo.ScrollBars = ScrollBars.None
        iCounter = 0

        While dataReader.Read
            sProgressStatus = ""
            iCustodianNameCol = dataReader.GetOrdinal("CustodianName")
            If dataReader.IsDBNull(iCustodianNameCol) Then
                sCustodianName = vbNullString
            Else
                sCustodianName = dataReader.GetString(iCustodianNameCol)
            End If

            iProgressStatusCol = dataReader.GetOrdinal("ProgressStatus")
            If dataReader.IsDBNull(iProgressStatusCol) Then
                sProgressStatus = vbNullString
            Else
                sProgressStatus = dataReader.GetString(iProgressStatusCol)
            End If

            iPSTPathCol = dataReader.GetOrdinal("PSTPath")
            If dataReader.IsDBNull(iPSTPathCol) Then
                sPSTPath = vbNullString
            Else
                sPSTPath = dataReader.GetString(iPSTPathCol)
            End If

            iNumberOfPSTsCol = dataReader.GetOrdinal("NumberOfPSTs")
            If dataReader.IsDBNull(iNumberOfPSTsCol) Then
                iNumberOfPSTs = 0
            Else
                iNumberOfPSTs = dataReader.GetInt16(iNumberOfPSTsCol)
            End If

            iTotalSizeOfPSTsCol = dataReader.GetOrdinal("TotalSizeOfPSTs")
            If IsDBNull(iTotalSizeOfPSTsCol) Then
                dblTotalSizeOfPSTs = 0.0
            Else
                dblTotalSizeOfPSTs = dataReader.GetInt64(iTotalSizeOfPSTsCol)
                sTotalSizeOfPSTs = dblTotalSizeOfPSTs.ToString("N0")
            End If

            iGroupIDCol = dataReader.GetOrdinal("GroupID")
            If dataReader.IsDBNull(iGroupIDCol) Then
                sGroupID = vbNullString
            Else
                sGroupID = dataReader.GetString(iGroupIDCol)
                lstDBGroups.Add(sGroupID)
            End If

            iDestinationFolderCol = dataReader.GetOrdinal("DestinationFolder")
            If dataReader.IsDBNull(iDestinationFolderCol) Then
                sDestinationFolder = vbNullString
            Else
                sDestinationFolder = dataReader.GetString(iDestinationFolderCol)
            End If

            iDestinationRootCol = dataReader.GetOrdinal("DestinationRoot")
            If dataReader.IsDBNull(iDestinationRootCol) Then
                sDestinationRoot = vbNullString
            Else
                sDestinationRoot = dataReader.GetString(iDestinationRootCol)
            End If

            iDestinationSMTPCol = dataReader.GetOrdinal("DestinationSMTP")
            If dataReader.IsDBNull(iDestinationSMTPCol) Then
                sDestinationSMTP = vbNullString
            Else
                sDestinationSMTP = dataReader.GetString(iDestinationSMTPCol)
            End If

            iProcessIDCol = dataReader.GetOrdinal("ProcessID")
            If dataReader.IsDBNull(iProcessIDCol) Then
                sProcessID = vbNullString
            Else
                sProcessID = dataReader.GetString(iProcessIDCol)
            End If

            iIngestionStartTimeCol = dataReader.GetOrdinal("IngestionStartTime")
            If dataReader.IsDBNull(iIngestionStartTimeCol) Then
                sIngestionStartTime = vbNullString
            Else
                sIngestionStartTime = dataReader.GetString(iIngestionStartTimeCol)
            End If

            iIngestionEndTimeCol = dataReader.GetOrdinal("IngestionEndTime")
            If dataReader.IsDBNull(iIngestionEndTimeCol) Then
                sIngestionEndTime = vbNullString
            Else
                sIngestionEndTime = dataReader.GetString(iIngestionEndTimeCol)
            End If

            iBytesUploadedCol = dataReader.GetOrdinal("BytesUploaded")
            If dataReader.IsDBNull(iBytesUploadedCol) Then
                dblBytesUploaded = 0.0
            Else
                dblBytesUploaded = dataReader.GetInt64(iBytesUploadedCol).ToString("N0")
                sBytesUploaded = dblBytesUploaded.ToString("N0")
            End If

            iPercentCompletedCol = dataReader.GetOrdinal("PercentCompleted")
            If dataReader.IsDBNull(iPercentCompletedCol) Then
                iPercentCompleted = 0
            Else
                iPercentCompleted = dataReader.GetValue(iPercentCompletedCol)
                'sPercentCompleted = dataReader.GetString(iPercentCompletedCol)
                iPercentCompleted = iPercentCompleted / 100
                sPercentCompleted = FormatPercent(iPercentCompleted)
            End If

            iSuccessCol = dataReader.GetOrdinal("Success")
            If dataReader.IsDBNull(iSuccessCol) Then
                dblSuccess = 0.0
            Else
                dblSuccess = dataReader.GetInt64(iSuccessCol)
                sSuccess = dblSuccess.ToString("N0")
            End If

            iFailedCol = dataReader.GetOrdinal("Failed")
            If dataReader.IsDBNull(iFailedCol) Then
                dblFailed = 0.0
            Else
                dblFailed = dataReader.GetInt64(iFailedCol).ToString("N0")
                sFailed = dblFailed.ToString("N0")
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

            iProcessingFilesDirectoryCol = dataReader.GetOrdinal("ProcessingFilesDirectory")
            If dataReader.IsDBNull(iProcessingFilesDirectoryCol) Then
                sProcessingFilesDirectory = vbNullString
            Else
                sProcessingFilesDirectory = dataReader.GetString(iProcessingFilesDirectoryCol)
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

            Select Case sProgressStatus
                Case "Not Started"
                    iRowIndex = grdPSTInfo.Rows.Add()

                    CustodianRow = grdPSTInfo.Rows(iRowIndex)

                    With CustodianRow
                        CustodianRow.Cells("StatusImage").Value = My.Resources.waitingtostart1
                        CustodianRow.Cells("SelectCustodian").Value = False
                        CustodianRow.Cells("ProgressStatus").Value = sProgressStatus
                        CustodianRow.Cells("CustodianName").Value = sCustodianName
                        CustodianRow.Cells("PSTPath").Value = sPSTPath
                        CustodianRow.Cells("PercentageCompleted").Value = sPercentCompleted.Replace(".00", "")
                        CustodianRow.Cells("BytesUploaded").Value = sBytesUploaded
                        CustodianRow.Cells("NumberOfPSTs").Value = iNumberOfPSTs
                        CustodianRow.Cells("SizeOfPSTs").Value = sTotalSizeOfPSTs
                        CustodianRow.Cells("DestinationFolder").Value = sDestinationFolder
                        CustodianRow.Cells("DestinationRoot").Value = sDestinationRoot
                        CustodianRow.Cells("DestinationSMTP").Value = sDestinationSMTP
                        CustodianRow.Cells("GroupID").Value = sGroupID
                        CustodianRow.Cells("ProcessID").Value = sProcessID
                        CustodianRow.Cells("IngestionStartTime").Value = sIngestionStartTime
                        CustodianRow.Cells("IngestionEndTime").Value = sIngestionEndTime
                        CustodianRow.Cells("NumberSuccess").Value = sSuccess
                        CustodianRow.Cells("NumberFailed").Value = sFailed
                        CustodianRow.Cells("CaseDirectory").Value = sCaseDirectory
                        CustodianRow.Cells("OutputDirectory").Value = sOutputDirectory
                        CustodianRow.Cells("LogDirectory").Value = sLogDirectory
                        CustodianRow.Cells("ProcessingFilesDirectory").Value = sProcessingFilesDirectory
                        CustodianRow.Cells("ReprocessingDirectory").Value = sOutputDirectory & "\Reprocessing\"
                        CustodianRow.Cells("SummaryReportLocation").Value = sSummaryReportLocation
                    End With
                    'grdPSTInfo.Rows.Add(CustodianRow)
                Case "In Progress"
                    iRowIndex = grdPSTInfo.Rows.Add()

                    CustodianRow = grdPSTInfo.Rows(iRowIndex)

                    With CustodianRow
                        CustodianRow.Cells("StatusImage").Value = My.Resources.inprogress_medium
                        CustodianRow.Cells("SelectCustodian").Value = False
                        CustodianRow.Cells("ProgressStatus").Value = sProgressStatus
                        CustodianRow.Cells("CustodianName").Value = sCustodianName
                        CustodianRow.Cells("PSTPath").Value = sPSTPath
                        CustodianRow.Cells("PercentageCompleted").Value = sPercentCompleted.Replace(".00", "")
                        CustodianRow.Cells("BytesUploaded").Value = sBytesUploaded
                        CustodianRow.Cells("NumberOfPSTs").Value = iNumberOfPSTs
                        CustodianRow.Cells("SizeOfPSTs").Value = sTotalSizeOfPSTs
                        CustodianRow.Cells("DestinationFolder").Value = sDestinationFolder
                        CustodianRow.Cells("DestinationRoot").Value = sDestinationRoot
                        CustodianRow.Cells("DestinationSMTP").Value = sDestinationSMTP
                        CustodianRow.Cells("GroupID").Value = sGroupID
                        CustodianRow.Cells("ProcessID").Value = sProcessID
                        CustodianRow.Cells("IngestionStartTime").Value = sIngestionStartTime
                        CustodianRow.Cells("IngestionEndTime").Value = sIngestionEndTime
                        CustodianRow.Cells("NumberSuccess").Value = sSuccess
                        CustodianRow.Cells("NumberFailed").Value = sFailed
                        CustodianRow.Cells("CaseDirectory").Value = sCaseDirectory
                        CustodianRow.Cells("OutputDirectory").Value = sOutputDirectory
                        CustodianRow.Cells("LogDirectory").Value = sLogDirectory
                        CustodianRow.Cells("ProcessingFilesDirectory").Value = sProcessingFilesDirectory
                        CustodianRow.Cells("ReprocessingDirectory").Value = sOutputDirectory & "\Reprocessing\"
                        CustodianRow.Cells("SummaryReportLocation").Value = sSummaryReportLocation
                    End With

                Case "Completed"
                    iRowIndex = grdPSTInfo.Rows.Add()

                    CustodianRow = grdPSTInfo.Rows(iRowIndex)

                    With CustodianRow
                        CustodianRow.Cells("SelectCustodian").Value = False
                        CustodianRow.Cells("CustodianName").Value = sCustodianName
                        CustodianRow.Cells("ProgressStatus").Value = sProgressStatus
                        CustodianRow.Cells("PSTPath").Value = sPSTPath
                        CustodianRow.Cells("PercentageCompleted").Value = sPercentCompleted.Replace(".00", "")
                        CustodianRow.Cells("BytesUploaded").Value = sBytesUploaded
                        CustodianRow.Cells("NumberOfPSTs").Value = iNumberOfPSTs
                        CustodianRow.Cells("SizeOfPSTs").Value = sTotalSizeOfPSTs
                        CustodianRow.Cells("DestinationFolder").Value = sDestinationFolder
                        CustodianRow.Cells("DestinationRoot").Value = sDestinationRoot
                        CustodianRow.Cells("DestinationSMTP").Value = sDestinationSMTP
                        CustodianRow.Cells("GroupID").Value = sGroupID
                        CustodianRow.Cells("ProcessID").Value = sProcessID
                        CustodianRow.Cells("IngestionStartTime").Value = sIngestionStartTime
                        CustodianRow.Cells("IngestionEndTime").Value = sIngestionEndTime
                        CustodianRow.Cells("NumberSuccess").Value = sSuccess
                        CustodianRow.Cells("NumberFailed").Value = sFailed
                        CustodianRow.Cells("CaseDirectory").Value = sCaseDirectory
                        CustodianRow.Cells("OutputDirectory").Value = sOutputDirectory
                        CustodianRow.Cells("LogDirectory").Value = sLogDirectory
                        CustodianRow.Cells("ProcessingFilesDirectory").Value = sProcessingFilesDirectory
                        CustodianRow.Cells("ReprocessingDirectory").Value = sOutputDirectory & "\Reprocessing\"
                        CustodianRow.Cells("SummaryReportLocation").Value = sSummaryReportLocation
                        If dblFailed > 0 Then
                            If dblSuccess > 0 Then
                                dSuccessPercent = 1 - (CInt(dblFailed) / CInt(dblSuccess))
                            Else
                                dSuccessPercent = 1 - (CInt(dblFailed) / 1)

                            End If
                            If (dSuccessPercent > 0.9) Then
                                CustodianRow.Cells("StatusImage").Value = My.Resources.Green_check_small
                                grdPSTInfo.Rows(iCounter).DefaultCellStyle.ForeColor = Color.Green
                            ElseIf (dSuccessPercent > 0.75) Then
                                CustodianRow.Cells("StatusImage").Value = My.Resources.yellow_info_small
                                grdPSTInfo.Rows(iCounter).DefaultCellStyle.ForeColor = Color.Orange
                            Else
                                CustodianRow.Cells("StatusImage").Value = My.Resources.red_stop_small
                                grdPSTInfo.Rows(iCounter).DefaultCellStyle.ForeColor = Color.Red
                            End If
                        Else
                            CustodianRow.Cells("StatusImage").Value = My.Resources.Green_check_small
                            grdPSTInfo.Rows(iCounter).DefaultCellStyle.ForeColor = Color.Green
                        End If
                    End With
                Case "Process Termniated"
                    iRowIndex = grdPSTInfo.Rows.Add()

                    CustodianRow = grdPSTInfo.Rows(iRowIndex)

                    With CustodianRow
                        CustodianRow.Cells("StatusImage").Value = My.Resources.yellow_info_small
                        CustodianRow.Cells("SelectCustodian").Value = False
                        CustodianRow.Cells("ProgressStatus").Value = sProgressStatus
                        CustodianRow.Cells("CustodianName").Value = sCustodianName
                        CustodianRow.Cells("PSTPath").Value = sPSTPath
                        CustodianRow.Cells("PercentageCompleted").Value = sPercentCompleted.Replace(".00", "")
                        CustodianRow.Cells("BytesUploaded").Value = sBytesUploaded
                        CustodianRow.Cells("NumberOfPSTs").Value = iNumberOfPSTs
                        CustodianRow.Cells("SizeOfPSTs").Value = sTotalSizeOfPSTs
                        CustodianRow.Cells("DestinationFolder").Value = sDestinationFolder
                        CustodianRow.Cells("DestinationRoot").Value = sDestinationRoot
                        CustodianRow.Cells("DestinationSMTP").Value = sDestinationSMTP
                        CustodianRow.Cells("GroupID").Value = sGroupID
                        CustodianRow.Cells("ProcessID").Value = sProcessID
                        CustodianRow.Cells("IngestionStartTime").Value = sIngestionStartTime
                        CustodianRow.Cells("IngestionEndTime").Value = sIngestionEndTime
                        CustodianRow.Cells("NumberSuccess").Value = sSuccess
                        CustodianRow.Cells("NumberFailed").Value = sFailed
                        CustodianRow.Cells("CaseDirectory").Value = sCaseDirectory
                        CustodianRow.Cells("OutputDirectory").Value = sOutputDirectory
                        CustodianRow.Cells("LogDirectory").Value = sLogDirectory
                        CustodianRow.Cells("ProcessingFilesDirectory").Value = sProcessingFilesDirectory
                        CustodianRow.Cells("ReprocessingDirectory").Value = sOutputDirectory & "\Reprocessing\"
                        CustodianRow.Cells("SummaryReportLocation").Value = sSummaryReportLocation
                        If dblFailed > 0 Then
                            If dblSuccess > 0 Then
                                dSuccessPercent = 1 - (CInt(dblFailed) / CInt(dblSuccess))
                            Else
                                dSuccessPercent = 1 - (CInt(dblFailed) / 1)

                            End If
                            If (dSuccessPercent > 0.9) Then
                                CustodianRow.Cells("StatusImage").Value = My.Resources.Green_check_small
                                grdPSTInfo.Rows(iCounter).DefaultCellStyle.ForeColor = Color.Green
                            ElseIf (dSuccessPercent > 0.75) Then
                                CustodianRow.Cells("StatusImage").Value = My.Resources.yellow_info_small
                                grdPSTInfo.Rows(iCounter).DefaultCellStyle.ForeColor = Color.Orange
                            Else
                                CustodianRow.Cells("StatusImage").Value = My.Resources.red_stop_small
                                grdPSTInfo.Rows(iCounter).DefaultCellStyle.ForeColor = Color.Red
                            End If
                        Else
                            CustodianRow.Cells("StatusImage").Value = My.Resources.Green_check_small
                            grdPSTInfo.Rows(iCounter).DefaultCellStyle.ForeColor = Color.Green
                        End If
                    End With

                Case Else
                    iRowIndex = grdPSTInfo.Rows.Add()

                    CustodianRow = grdPSTInfo.Rows(iRowIndex)

                    With CustodianRow
                        CustodianRow.Cells("StatusImage").Value = My.Resources.not_selected_small
                        CustodianRow.Cells("SelectCustodian").Value = False
                        CustodianRow.Cells("CustodianName").Value = sCustodianName
                        CustodianRow.Cells("ProgressStatus").Value = sProgressStatus
                        CustodianRow.Cells("PSTPath").Value = sPSTPath
                        CustodianRow.Cells("PercentageCompleted").Value = sPercentCompleted.Replace(".00", "")
                        CustodianRow.Cells("BytesUploaded").Value = sBytesUploaded
                        CustodianRow.Cells("NumberOfPSTs").Value = iNumberOfPSTs
                        CustodianRow.Cells("SizeOfPSTs").Value = sTotalSizeOfPSTs
                        CustodianRow.Cells("DestinationFolder").Value = sDestinationFolder
                        CustodianRow.Cells("DestinationRoot").Value = sDestinationRoot
                        CustodianRow.Cells("DestinationSMTP").Value = sDestinationSMTP
                        CustodianRow.Cells("GroupID").Value = sGroupID
                        CustodianRow.Cells("ProcessID").Value = sProcessID
                        CustodianRow.Cells("IngestionStartTime").Value = sIngestionStartTime
                        CustodianRow.Cells("IngestionEndTime").Value = sIngestionEndTime
                        CustodianRow.Cells("NumberSuccess").Value = sSuccess
                        CustodianRow.Cells("NumberFailed").Value = sFailed
                        CustodianRow.Cells("CaseDirectory").Value = sCaseDirectory
                        CustodianRow.Cells("OutputDirectory").Value = sOutputDirectory
                        CustodianRow.Cells("LogDirectory").Value = sLogDirectory
                        CustodianRow.Cells("ProcessingFilesDirectory").Value = sProcessingFilesDirectory
                        CustodianRow.Cells("ReprocessingDirectory").Value = sOutputDirectory & "\Reprocessing\"
                        CustodianRow.Cells("SummaryReportLocation").Value = sSummaryReportLocation

                    End With
            End Select
            iCounter = iCounter + 1
        End While

        cboGroupIDs.Items.Add("")
        lstUniqueGroups = lstDBGroups.Distinct.ToList
        For Each UniqueGroup In lstUniqueGroups
            cboGroupIDs.Items.Add(UniqueGroup.ToString)
        Next

        sqlConnection.Close()
        grdPSTInfo.ScrollBars = ScrollBars.Both
        blnLoadGridFromDB = True

    End Function

    Private Function blnGetFolderCustodians(ByVal sFolderName As String, ByRef lstFolderContents As List(Of String)) As Boolean
        blnGetFolderCustodians = False
        Dim sFullPath As String
        Dim sCustodian As String

        If My.Computer.FileSystem.DirectoryExists(sFolderName) Then
            For Each d In Directory.GetDirectories(sFolderName)
                sFullPath = Path.GetFullPath(d)
                sCustodian = d.Substring(d.LastIndexOf("\") + 1)

                lstFolderContents.Add(sCustodian)
            Next
        End If

        blnGetFolderCustodians = True
    End Function
    Private Function blnGetAllSummaryReports(ByVal sCaseDirectory As String, ByRef lstSummaryReportFiles As List(Of String)) As Boolean
        blnGetAllSummaryReports = False
        Dim CurrentDirectory As DirectoryInfo
        Dim Length As Integer

        CurrentDirectory = New DirectoryInfo(sCaseDirectory)

        For Each File In CurrentDirectory.GetFiles
            If File.Name = "summary-report.txt" Then
                lstSummaryReportFiles.Add(File.FullName)
            End If
        Next

        Length = Directory.GetDirectories(sCaseDirectory).Length
        If Length > 0 Then
            For Each d In Directory.GetDirectories(sCaseDirectory)
                blnGetAllSummaryReports(d, lstSummaryReportFiles)
            Next
        End If

        blnGetAllSummaryReports = True
    End Function

    Private Function blnRefreshSummaryData(ByVal sSummaryReportFile As String, ByRef lstSummaryReportInfo As List(Of String)) As Boolean
        Dim lstCustodianName As List(Of String)
        Dim lstCustodianItemsExported As List(Of String)
        Dim lstCustodianItemsFailed As List(Of String)
        Dim lstCustodianBytesUploaded As List(Of String)
        Dim lstInvalidCustodian As List(Of String)
        Dim FileStream As StreamReader

        Dim sCurrentRow As String
        Dim sMailboxName As String
        Dim sCustodianName As String
        Dim asCustodianValues() As String
        Dim asItemsExported() As String
        Dim sItemsExported As String
        Dim asItemsFailed() As String
        Dim sItemsFailed As String
        Dim asBytesUploaded() As String
        Dim sBytesUploaded As String

        blnRefreshSummaryData = False

        lstCustodianName = New List(Of String)
        lstCustodianItemsExported = New List(Of String)
        lstCustodianItemsFailed = New List(Of String)
        lstCustodianBytesUploaded = New List(Of String)
        lstInvalidCustodian = New List(Of String)

        Threading.Thread.Sleep(2000)
        FileStream = New StreamReader(sSummaryReportFile)
        While Not FileStream.EndOfStream
            sCurrentRow = FileStream.ReadLine
            If sCurrentRow.Contains("Mailbox name: ") Then
                sMailboxName = sCurrentRow.Substring(sCurrentRow.IndexOf(": ") + 2)
                sMailboxName = sMailboxName.Substring(0, sMailboxName.IndexOf("\"))
            ElseIf sCurrentRow.Contains("SOURCE:") Then
                asCustodianValues = Split(sCurrentRow, ":")
                sCustodianName = Trim(asCustodianValues(1))
            ElseIf (sCurrentRow.Contains("ITEMS_EXPORTED: ")) Then
                If sCurrentRow.Contains("count:") Then
                    asItemsExported = Split(sCurrentRow, ":")
                    sItemsExported = asItemsExported(UBound(asItemsExported))
                Else
                    asItemsExported = Split(sCurrentRow, ":")
                    sItemsExported = asItemsExported(1)
                    sItemsExported = sItemsExported.Substring(0, sItemsExported.IndexOf(" in"))
                End If
            ElseIf (sCurrentRow.Contains("ITEMS_FAILED: ")) Then
                If sCurrentRow.Contains("count:") Then
                    asItemsFailed = Split(sCurrentRow, ":")
                    sItemsFailed = asItemsFailed(UBound(asItemsFailed))
                Else
                    asItemsFailed = Split(sCurrentRow, ":")
                    sItemsFailed = asItemsFailed(1).Substring(0, asItemsFailed(1).IndexOf(" in"))
                End If
            ElseIf (sCurrentRow.Contains("BYTES_UPLOADED: ")) Then
                If sCurrentRow.Contains("count:") Then
                    asBytesUploaded = Split(sCurrentRow, ":")
                    sBytesUploaded = asBytesUploaded(UBound(asBytesUploaded))
                Else
                    asBytesUploaded = Split(sCurrentRow, ":")
                    sBytesUploaded = asBytesUploaded(1).Substring(0, asBytesUploaded(1).IndexOf(" in"))

                End If
            End If

        End While

        If (sItemsExported = vbNullString) Then
            sItemsExported = 0
        End If

        If (sItemsFailed = vbNullString) Then
            sItemsFailed = 0
        End If

        If (sBytesUploaded = vbNullString) Then
            sBytesUploaded = 0
        End If

        lstSummaryReportInfo.Add(sCustodianName & "," & sMailboxName & "," & sItemsExported & "," & sItemsExported & "," & FormatNumber(sBytesUploaded, , , , TriState.True))

        blnRefreshSummaryData = True

    End Function

    Private Sub btnCustomerMappingFile_Click_1(sender As Object, e As EventArgs) Handles btnCustomerMappingFile.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        With OpenFileDialog1
            .Filter = "CSV|*.csv"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            txtCustodianMappingFile.Text = OpenFileDialog1.FileName.ToString
        End If
    End Sub

    Private Sub btnLoadCustodianMappingData_Click_1(sender As Object, e As EventArgs) Handles btnLoadCustodianMappingData.Click
        Dim lstMappingCustodianName As List(Of String)
        Dim lstMappingDestinationRoot As List(Of String)
        Dim lstMappingDestinationFolder As List(Of String)
        Dim lstMappingDestinationSMTP As List(Of String)
        Dim lstMappingCustodianGroupID As List(Of String)
        Dim lstDuplicateMappingCustodianName As List(Of String)
        Dim lstDuplicateMappingDestinationRoot As List(Of String)
        Dim lstDuplicateMappingDestinationFolder As List(Of String)
        Dim lstDuplicateMappingDestinationSMTP As List(Of String)
        Dim lstDuplicateMappingCustodianGroupID As List(Of String)
        Dim lstUniqueGroupID As List(Of String)
        Dim bStatus As Boolean
        Dim iPercentComplete As Integer
        Dim sProcessingFilesDir As String
        Dim sCaseDirectory As String
        Dim sOutputDirectory As String
        Dim sLogDirectory As String

        Dim iCustodianIndex As Integer

        Dim sCurrentRow() As String
        Dim common As New Common
        Dim dbService As New DatabaseService

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"

        lstMappingCustodianName = New List(Of String)
        lstMappingDestinationRoot = New List(Of String)
        lstMappingDestinationFolder = New List(Of String)
        lstMappingDestinationSMTP = New List(Of String)
        lstMappingCustodianGroupID = New List(Of String)
        lstDuplicateMappingCustodianName = New List(Of String)
        lstDuplicateMappingDestinationRoot = New List(Of String)
        lstDuplicateMappingDestinationFolder = New List(Of String)
        lstDuplicateMappingDestinationSMTP = New List(Of String)
        lstDuplicateMappingCustodianGroupID = New List(Of String)
        lstUniqueGroupID = New List(Of String)

        lstDuplicateMappingCustodianName = New List(Of String)

        common.Logger(psIngestionLogFile, "Loading Custodian Mapping File from " & txtCustodianMappingFile.Text & "...")
        If txtCustodianMappingFile.Text = vbNullString Then
            MessageBox.Show("You must select a Custodian Mapping File.")
        Else

            Dim fileCSVFile As New Microsoft.VisualBasic.FileIO.TextFieldParser(txtCustodianMappingFile.Text)
            fileCSVFile.TextFieldType = FileIO.FieldType.Delimited
            fileCSVFile.SetDelimiters(",")

            Do While Not fileCSVFile.EndOfData
                sCurrentRow = fileCSVFile.ReadFields

                If lstMappingCustodianName.Contains(sCurrentRow(0)) Then
                    lstDuplicateMappingCustodianName.Add(sCurrentRow(0))
                    lstDuplicateMappingDestinationFolder.Add(sCurrentRow(1))
                    lstDuplicateMappingDestinationRoot.Add(sCurrentRow(2))
                    lstDuplicateMappingDestinationSMTP.Add(sCurrentRow(3))
                    If sCurrentRow.Count > 4 Then
                        lstDuplicateMappingCustodianGroupID.Add(sCurrentRow(4))
                    End If
                Else
                    lstMappingCustodianName.Add(LCase(sCurrentRow(0)))
                    lstMappingDestinationFolder.Add(sCurrentRow(1))
                    lstMappingDestinationRoot.Add(sCurrentRow(2))
                    lstMappingDestinationSMTP.Add(sCurrentRow(3))
                    If sCurrentRow.Count > 4 Then
                        lstMappingCustodianGroupID.Add(sCurrentRow(4))
                    End If
                End If
            Loop
            cboGroupIDs.Items.Add("")
            lstUniqueGroupID = lstMappingCustodianGroupID.Distinct.ToList
            For Each UniqueGroup In lstUniqueGroupID
                cboGroupIDs.Items.Add(UniqueGroup.ToString)
            Next


            For Each row In grdPSTInfo.Rows
                Try
                    sProcessingFilesDir = row.cells("ProcessingFilesDirectory").value
                    sCaseDirectory = row.cells("CaseDirectory").value
                    sOutputDirectory = row.cells("OutputDirectory").value
                    sLogDirectory = row.cells("LogDirectory").value
                    If lstMappingCustodianName.Contains(LCase(row.cells("CustodianName").Value)) Then
                        iCustodianIndex = lstMappingCustodianName.IndexOf(LCase(row.cells("CustodianName").Value))
                        row.cells("GroupID").value = lstMappingCustodianGroupID(iCustodianIndex)
                        row.cells("DestinationFolder").value = lstMappingDestinationFolder(iCustodianIndex)
                        row.cells("DestinationRoot").value = lstMappingDestinationRoot(iCustodianIndex)
                        row.cells("DestinationSMTP").value = lstMappingDestinationSMTP(iCustodianIndex)
                        If row.cells("PercentageCompleted").value <> vbNullString Then
                            iPercentComplete = CInt(row.cells("PercentageCompleted").value.ToString.Replace("%", ""))
                            iPercentComplete = CInt(iPercentComplete.ToString.Replace(".00", ""))
                        Else
                            iPercentComplete = 0
                        End If
                        'bStatus = blnUpdateSQLiteAllCustodiansInfo(row.cells("CustodianName").value, "Staged", row.cells("PSTPath").value, row.cells("NumberOfPSTs").value, row.cells("SizeOfPSTs").value, row.cells("GroupID").value, row.cells("DestinationFolder").value, row.cells("DestinationRoot").value, row.cells("DestinationSMTP").value, False, False, False, row.cells("ProcessID").value, row.cells("IngestionStartTime").value, row.cells("IngestionEndTime").value, row.cells("BytesUploaded").value, iPercentComplete, row.cells("NumberSuccess").value, row.cells("NumberFailed").value, sProcessingFilesDir, sCaseDirectory, sOutputDirectory, sLogDirectory, row.cells("SummaryReportLocation").value)
                        'bStatus = dbService.UpdateSQLiteEWSDB(row.cells("CustodianName").value, "Staged", row.cells("PSTPath").value, row.cells("NumberOfPSTs").value, row.cells("SizeOfPSTs").value, row.cells("GroupID").value, row.cells("DestinationFolder").value, row.cells("DestinationRoot").value, row.cells("DestinationSMTP").value, False, False, False, row.cells("ProcessID").value, row.cells("IngestionStartTime").value, row.cells("IngestionEndTime").value, row.cells("BytesUploaded").value, iPercentComplete, row.cells("NumberSuccess").value, row.cells("NumberFailed").value, sProcessingFilesDir, sCaseDirectory, sOutputDirectory, sLogDirectory, row.cells("SummaryReportLocation").value)
                        bStatus = dbService.UpdateSQLiteEWSDB(sSQLiteDatabaseFullName, row.cells("CustodianName").value, "Staged", row.cells("PSTPath").value, row.cells("NumberOfPSTs").value, row.cells("SizeOfPSTs").value, row.cells("GroupID").value, row.cells("DestinationFolder").value, row.cells("DestinationRoot").value, row.cells("DestinationSMTP").value, row.cells("ProcessID").value, row.cells("IngestionStartTime").value, row.cells("IngestionEndTime").value, row.cells("BytesUploaded").value, row.cells("NumberSuccess").value, row.cells("NumberFailed").value, sProcessingFilesDir, sCaseDirectory, sOutputDirectory, sLogDirectory, row.cells("SummaryReportLocation").value)
                    End If

                Catch ex As Exception

                End Try
            Next
        End If
    End Sub

    Public Function blnGetSummaryReportInfo(ByVal sSummaryReport As String, ByRef sCustodianName As String, ByRef sItemsExported As String, ByRef sItemsFailed As String, ByRef sBytesUploaded As String, ByRef dStartDate As Date, ByRef dEndDate As Date) As Boolean
        blnGetSummaryReportInfo = False

        Dim FileStream As StreamReader
        Dim sCurrentRow As String
        Dim sMailboxName As String
        Dim asStartDateValues() As String
        Dim asEndDateValues() As String
        Dim asCustodianValues() As String
        Dim asItemsExported() As String
        Dim asItemsFailed() As String
        Dim asBytesUploaded() As String
        Dim sStartDate As String
        Dim sEndDate As String
        Dim common As New Common

        Try

            FileStream = New StreamReader(sSummaryReport)
            While Not FileStream.EndOfStream
                sCurrentRow = FileStream.ReadLine
                If sCurrentRow.Contains("Started at:") Then
                    sStartDate = sCurrentRow.Substring(sCurrentRow.IndexOf(":"))
                    sStartDate = sStartDate.Substring(2, sStartDate.IndexOf("(") - 3)
                    asStartDateValues = Split(sStartDate, " ")
                    sStartDate = asStartDateValues(1) & " " & asStartDateValues(2) & "," & asStartDateValues(4) & " " & asStartDateValues(3)
                    dStartDate = FormatDateTime(sStartDate, DateFormat.GeneralDate)

                ElseIf sCurrentRow.Contains("Completed at:") Then
                    sEndDate = sCurrentRow.Substring(sCurrentRow.IndexOf(":"))
                    sEndDate = sEndDate.Substring(2, sEndDate.IndexOf("(") - 3)
                    asEndDateValues = Split(sEndDate, " ")
                    sEndDate = asEndDateValues(1) & " " & asEndDateValues(2) & "," & asEndDateValues(4) & " " & asEndDateValues(3)
                    dEndDate = FormatDateTime(sEndDate, DateFormat.GeneralDate)
                ElseIf sCurrentRow.Contains("Mailbox name: ") Then
                    If sCurrentRow.Contains("\") Then
                        sMailboxName = sCurrentRow.Substring(sCurrentRow.IndexOf(": ") + 2)
                        sMailboxName = sMailboxName.Substring(0, sMailboxName.IndexOf("\"))
                    Else
                        asCustodianValues = Split(sCurrentRow, ":")
                        sCustodianName = Trim(asCustodianValues(1))
                    End If
                ElseIf sCurrentRow.Contains("SOURCE:") Then
                    asCustodianValues = Split(sCurrentRow, ":")
                    sCustodianName = Trim(asCustodianValues(1))
                ElseIf sCurrentRow.Contains("Case: ") Then
                    asCustodianValues = Split(sCurrentRow, ":")
                    sCustodianName = Trim(asCustodianValues(1))
                ElseIf sCurrentRow.Contains("TARGET:") Then
                    If sCurrentRow.Contains("\") Then
                        sMailboxName = sCurrentRow.Substring(sCurrentRow.IndexOf(": ") + 2)
                        sMailboxName = sMailboxName.Substring(0, sMailboxName.IndexOf("\"))
                    Else
                        asCustodianValues = Split(sCurrentRow, ":")
                        sCustodianName = Trim(asCustodianValues(1))
                    End If
                ElseIf (sCurrentRow.Contains("ITEMS_EXPORTED: ")) Then
                    If sCurrentRow.Contains("count:") Then
                        asItemsExported = Split(sCurrentRow, ":")
                        sItemsExported = asItemsExported(UBound(asItemsExported))
                    Else
                        asItemsExported = Split(sCurrentRow, ":")
                        sItemsExported = asItemsExported(1)
                        sItemsExported = sItemsExported.Substring(0, sItemsExported.IndexOf(" in"))
                    End If
                ElseIf (sCurrentRow.Contains("ITEMS_FAILED: ")) Then
                    If sCurrentRow.Contains("count:") Then
                        asItemsFailed = Split(sCurrentRow, ":")
                        sItemsFailed = asItemsFailed(UBound(asItemsFailed))
                    Else
                        asItemsFailed = Split(sCurrentRow, ":")
                        sItemsFailed = asItemsFailed(1).Substring(0, asItemsFailed(1).IndexOf(" in"))
                    End If

                ElseIf (sCurrentRow.Contains("BYTES_UPLOADED: ")) Then
                    If sCurrentRow.Contains("count:") Then
                        asBytesUploaded = Split(sCurrentRow, ":")
                        sBytesUploaded = asBytesUploaded(UBound(asBytesUploaded))
                    Else
                        asBytesUploaded = Split(sCurrentRow, ":")
                        sBytesUploaded = asBytesUploaded(1).Substring(0, asBytesUploaded(1).IndexOf(" in"))
                    End If
                ElseIf (sCurrentRow.Contains("Total items exported")) Then
                    asItemsExported = Split(sCurrentRow, ":")
                    sItemsExported = asItemsExported(UBound(asItemsExported))
                ElseIf (sCurrentRow.Contains("Total number of failed items")) Then
                    asItemsFailed = Split(sCurrentRow, ":")
                    sItemsFailed = asItemsFailed(UBound(asItemsFailed))
                End If
            End While

        Catch ex As Exception
            common.Logger(psIngestionLogFile, "Error In blnGetSummaryReportInfo")
            common.Logger(psIngestionLogFile, ex.Message)
            sCustodianName = vbNullString
        End Try


        If (sItemsExported = vbNullString) Then
            sItemsExported = 0
        End If

        If (sItemsFailed = vbNullString) Then
            sItemsFailed = 0
        End If

        If (sBytesUploaded = vbNullString) Then
            sBytesUploaded = 0
        End If

        blnGetSummaryReportInfo = True

    End Function


    Private Sub btnProcessingDetails_Click(sender As Object, e As EventArgs) Handles btnProcessingDetails.Click
        Dim bStatus As Boolean
        Dim sSummaryReportFiles As String
        Dim common As New Common

        If grdPSTInfo.RowCount > 1 Then
            bStatus = common.GetAllSummaryReportFiles(sSummaryReportFiles)

            If sSummaryReportFiles <> vbNullString Then
                Dim FailuresForm As FailuresForm
                FailuresForm = New FailuresForm
                FailuresForm.CustodianSummaryReportFile = sSummaryReportFiles
                FailuresForm.Text = "All Custodians Completed Processing"

                FailuresForm.ShowDialog()

            Else
                MessageBox.Show("There are no summary report files to process.", "No Summary Report Files", MessageBoxButtons.OK)
            End If
        End If

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles txtFilter.KeyDown
        Dim lstRowIndex As List(Of Integer)
        Dim iColumns As Integer

        lstRowIndex = New List(Of Integer)

        If e.KeyCode = Keys.Enter Then
            If Me.txtFilter.Text = vbNullString Then
                For Each row In grdPSTInfo.Rows
                    grdPSTInfo.Rows(row.cells(0).rowindex()).Visible = True

                Next
            Else
                For Each row In grdPSTInfo.Rows
                    iColumns = grdPSTInfo.ColumnCount
                    For iCounter = 0 To iColumns
                        Try
                            If (row.cells(iCounter).value.ToString.ToLower.Contains(Me.txtFilter.Text.ToString.ToLower)) Then
                                lstRowIndex.Add(row.cells(0).rowindex())
                            End If
                        Catch ex As Exception

                        End Try
                    Next
                Next

                lstRowIndex.Distinct()
                For Each row In grdPSTInfo.Rows
                    If lstRowIndex.Contains(row.cells(0).rowindex()) Then
                        grdPSTInfo.Rows(row.cells(0).rowindex()).Visible = True
                    Else
                        If row.cells("CustodianName").value <> vbNullString Then
                            grdPSTInfo.Rows(row.cells(0).rowindex()).Visible = False
                        End If


                    End If
                Next
            End If
        End If
    End Sub

    Private Sub grdPSTInfo_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdPSTInfo.CellContentClick

    End Sub

    Private Sub btnNuixSystemSettings_MouseHover(sender As Object, e As EventArgs) Handles btnNuixSystemSettings.MouseHover
        EWSIngestToolTip.Show("Show Settings...", btnNuixSystemSettings)
    End Sub

    Private Sub btnLoadPreviousConfig_MouseHover(sender As Object, e As EventArgs) Handles btnLoadPreviousConfig.MouseHover
        EWSIngestToolTip.Show("Load Previous Configuration", btnLoadPreviousConfig)

    End Sub

    Private Sub btnProcessingDetails_MouseHover(sender As Object, e As EventArgs) Handles btnProcessingDetails.MouseHover
        EWSIngestToolTip.Show("Show Processing Details...", btnProcessingDetails)

    End Sub

    Private Sub btnRefreshGridData_MouseHover(sender As Object, e As EventArgs) Handles btnRefreshGridData.MouseHover
        EWSIngestToolTip.Show("Refresh Data from Database", btnRefreshGridData)

    End Sub

    Private Sub btnProcessandRunSelectedUsers_MouseHover(sender As Object, e As EventArgs) Handles btnProcessandRunSelectedUsers.MouseHover
        EWSIngestToolTip.Show("Process Selected Users", btnProcessandRunSelectedUsers)

    End Sub

    Private Sub btnExportProcessingDetails_Click(sender As Object, e As EventArgs) Handles btnExportProcessingDetails.Click
        Dim ReportOutputFile As StreamWriter
        Dim sOutputFileName As String
        Dim sMachineName As String

        Dim sCustodianName As String
        Dim sPercentCompleted As String
        Dim sBytesUploaded As String
        Dim sProgressStatus As String
        Dim sNumberOfPSTs As String
        Dim sSizeOfPSTs As String
        Dim sPSTPath As String
        Dim sDestinationFolder As String
        Dim sDestinationRoot As String
        Dim sDestinationSMTP As String
        Dim sGroupID As String
        Dim sProcessID As String
        Dim sIngestionStartTime As String
        Dim sIngestionEndTime As String
        Dim sNumberSuccess As String
        Dim sNumberFailed As String
        Dim sCaseDirectory As String
        Dim sOutputDirectory As String
        Dim sLogDirectory As String
        Dim sSummaryReportLocation As String


        sMachineName = System.Net.Dns.GetHostName()
        sOutputFileName = "EWS Processing Statistic - " & sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv"
        ReportOutputFile = New StreamWriter(eMailArchiveMigrationManager.NuixFilesDir & "\" & sOutputFileName)

        ReportOutputFile.WriteLine("Custodian Name, Percent Completed, Bytes Uploaded, Progress Status, Number of PSTs, Size of PSTs, PST Path, Destination Folder, Destination Root, Destination SMTP, Group ID, Process ID, Ingestion Start Time, Success, Failed, Summary Report Location")


        For Each row In grdPSTInfo.Rows
            sCustodianName = row.cells("CustodianName").value
            If sCustodianName <> vbNullString Then

                sPercentCompleted = row.cells("PercentageCompleted").value
                sBytesUploaded = row.cells("BytesUploaded").value
                sProgressStatus = row.cells("ProgressStatus").value
                sNumberOfPSTs = row.cells("NumberOfPSTs").value
                sSizeOfPSTs = row.cells("SizeOfPSTs").value
                sPSTPath = row.cells("PSTPath").value
                sDestinationFolder = row.cells("DestinationFolder").value
                sDestinationRoot = row.cells("DestinationRoot").value
                sDestinationSMTP = row.cells("DestinationSMTP").value
                sGroupID = row.cells("GroupID").value
                sProcessID = row.cells("ProcessID").value
                sIngestionStartTime = row.cells("IngestionStartTime").value
                sIngestionEndTime = row.cells("IngestionEndTime").value
                sNumberSuccess = row.cells("NumberSuccess").value
                sNumberFailed = row.cells("NumberFailed").value
                sCaseDirectory = row.cells("CaseDirectory").value
                sOutputDirectory = row.cells("OutputDirectory").value
                sLogDirectory = row.cells("LogDirectory").value
                sSummaryReportLocation = row.cells("SummaryReportLocation").value

                ReportOutputFile.WriteLine("""" & sCustodianName & """" & "," & """" & sPercentCompleted & """" & "," & """" & sBytesUploaded & """" & "," & """" & sProgressStatus & """" & "," & """" & sNumberOfPSTs & """" & "," & """" & sSizeOfPSTs & """" & "," & """" & sPSTPath & """" & "," & """" & sDestinationFolder & """" & "," & """" & sDestinationRoot & """" & "," & """" & sDestinationSMTP & """" & "," & """" & sGroupID & """" & "," & """" & sProcessID & """" & "," & """" & sProcessID & """" & "," & """" & sIngestionStartTime & """" & "," & """" & sIngestionEndTime & """" & "," & """" & sNumberSuccess & """" & "," & """" & sNumberFailed & """" & "," & """" & sCaseDirectory & """" & """" & "," & """" & sLogDirectory & """" & """" & "," & """" & sOutputDirectory & """" & """" & "," & """" & sSummaryReportLocation & """")
            End If
        Next

        ReportOutputFile.Close()
        MessageBox.Show("Nuix Case Statistics report finished building.  Report located at: " & eMailArchiveMigrationManager.NuixFilesDir & "\" & sOutputFileName)
    End Sub

    Private Sub btnExportProcessingDetails_MouseHover(sender As Object, e As EventArgs) Handles btnExportProcessingDetails.MouseHover
        EWSIngestToolTip.Show("Export Processing Details to CSV", btnExportProcessingDetails)
    End Sub

    Private Sub btnExportNotUploaded_Click(sender As Object, e As EventArgs) Handles btnExportNotUploaded.Click
        ItemsNotUploadedMenuStrip.Show(btnExportNotUploaded, 0, btnExportNotUploaded.Height)
    End Sub

    Private Sub btnExportNotUploaded_MouseHover(sender As Object, e As EventArgs) Handles btnExportNotUploaded.MouseHover
        EWSIngestToolTip.Show("Export Not Uploaded Email To PST", btnExportNotUploaded)
    End Sub

    Private Sub SelectAllItemsNotProcessed_Click(sender As Object, e As EventArgs) Handles SelectAllItemsNotProcessed.Click
        Dim lstCustodianNotUploaded As List(Of String)

        lstCustodianNotUploaded = New List(Of String)

        For Each row In grdPSTInfo.Rows
            If row.cells("NumberFailed").value > 0 Then
                row.cells("SelectCustodian").value = True
                lstCustodianNotUploaded.Add(row.cells("CustodianName").value)
                row.cells("ProgressStatus").value = "Not Started Items Not Uploaded"
            Else
                row.cells("SelectCustodian").value = False
            End If
        Next
        If lstCustodianNotUploaded.Count = 0 Then
            MessageBox.Show("There are no custodians with items not uploaded", "No custodians selected", MessageBoxButtons.OK)
            ExportItemsNotUploadedToPST.Enabled = False
            ExportFTSDataTooLargeToolStripMenuItem.Enabled = False
            ExportOtherNotUploadedToolStripMenuItem.Enabled = False
        Else
            ExportItemsNotUploadedToPST.Enabled = True
            ExportFTSDataTooLargeToolStripMenuItem.Enabled = False
            ExportOtherNotUploadedToolStripMenuItem.Enabled = False
        End If
    End Sub

    Private Sub btnConsolidateExporterFiles_Click(sender As Object, e As EventArgs) Handles btnConsolidateExporterFiles.Click
        ConsolidateExporterFiles.Show(btnConsolidateExporterFiles, 0, btnConsolidateExporterFiles.Height)

    End Sub

    Private Sub ConsolidateExporterMetricsFilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsolidateExporterMetricsFilesToolStripMenuItem.Click
        Dim bStatus As Boolean
        Dim lstExporterMetricsFiles As List(Of String)
        Dim ExporterMetricFile As StreamReader
        Dim ConsolidateExporterMetricsFile As StreamWriter
        Dim sCurrentRow As String
        Dim bFirstMetricsFile As Boolean

        bFirstMetricsFile = True
        lstExporterMetricsFiles = New List(Of String)
        ConsolidateExporterMetricsFile = New StreamWriter(eMailArchiveMigrationManager.NuixFilesDir & "\ews ingestion\consolidated-exporter-metrics" & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv")

        bStatus = blnGetAllExporterMetrics(eMailArchiveMigrationManager.NuixCaseDir & "\ews ingestion", lstExporterMetricsFiles)

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

    Private Sub ConsolidateExporterErrorsFilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsolidateExporterErrorsFilesToolStripMenuItem.Click
        Dim bStatus As Boolean
        Dim lstExporterErrorsFiles As List(Of String)
        Dim ExporterErrorsFile As StreamReader
        Dim ConsolidateExporterErrorsFile As StreamWriter
        Dim sCurrentRow As String
        Dim bFirstErrorFile As Boolean

        bFirstErrorFile = True
        lstExporterErrorsFiles = New List(Of String)
        bStatus = blnGetAllExporterErrors(eMailArchiveMigrationManager.NuixCaseDir & "\ews ingestion", lstExporterErrorsFiles)

        ConsolidateExporterErrorsFile = New StreamWriter(eMailArchiveMigrationManager.NuixFilesDir & "\ews ingestion\consolidated-exporter-errors" & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv")

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

    Private Sub ExportItemsNotUploadedToPST_Click(sender As Object, e As EventArgs) Handles ExportItemsNotUploadedToPST.Click

        Dim bStatus As Boolean
        Dim lstReProcessingCustodians As List(Of String)
        Dim lstRestartProcessingCustodian As List(Of String)
        Dim lstRestartSelectedCustodianInfo As List(Of String)
        Dim lstInvalidCustodian As List(Of String)
        Dim sPSTDirectory As String
        Dim sNuixAppMemory As String
        Dim iNumberofNuixInstances As Integer
        Dim lstCurrentCustodianFolder As List(Of String)
        Dim bFoundSelectedCustodian As Boolean
        Dim msgboxReturn As DialogResult
        Dim sReprocessingDirectory As String
        Dim common As New Common

        btnLoadPreviousConfig.Enabled = False
        btnProcessandRunSelectedUsers.Enabled = False
        btnProcessandRunSelectedUsers.Enabled = False
        btnProcessandRunSelectedUsers.Enabled = False

        Try
            If psSettingsFile = vbNullString Then
                msgboxReturn = MessageBox.Show("The settings have not been loaded. Would you like to load the setting dialog.", "Settings not loaded!", MessageBoxButtons.YesNo)
                If msgboxReturn = vbYes Then
                    Dim SettingsForm As O365ExtractionSettings
                    SettingsForm = New O365ExtractionSettings
                    SettingsForm.ShowDialog()
                    Cursor.Current = Cursors.Default
                    Exit Sub
                Else
                    Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            End If

            plstNuixConsoleProcesses = New List(Of String)
            lstReProcessingCustodians = New List(Of String)
            lstRestartSelectedCustodianInfo = New List(Of String)
            lstRestartProcessingCustodian = New List(Of String)
            lstInvalidCustodian = New List(Of String)
            lstCurrentCustodianFolder = New List(Of String)

            sPSTDirectory = txtPSTLocation.Text

            bFoundSelectedCustodian = False

            For Each Row In grdPSTInfo.Rows
                If Row.cells("SelectCustodian").value = True Then
                    If (Row.cells("ProgressStatus").value = "Not Started Items Not Uploaded") Then
                        lstReProcessingCustodians.Add(Row.cells("CustodianName").value)
                        sReprocessingDirectory = Row.Cells("ReprocessingDirectory").value
                    End If
                End If
            Next

            If lstReProcessingCustodians.Count <= 0 Then
                MessageBox.Show("You must select at least one custodian for EWS Ingestion.")
                Exit Sub
            End If

            sNuixAppMemory = CInt(eMailArchiveMigrationManager.O365NuixAppMemory) / 1000
            sNuixAppMemory = "-Xmx" & sNuixAppMemory & "g"

            grdPSTInfo.Sort(grdPSTInfo.Columns("SelectCustodian"), System.ComponentModel.ListSortDirection.Descending)

            bStatus = blnReProcessSelectedCustodians(lstReProcessingCustodians, grdPSTInfo, iNumberofNuixInstances, sReprocessingDirectory, sNuixAppMemory)

            Cursor.Current = Cursors.Default

        Catch ex As Exception
            Dim Result As String
            Dim hr As Integer = Runtime.InteropServices.Marshal.GetHRForException(ex)
            Result = ex.GetType.ToString & "(0x" & hr.ToString("X8") & "): " & ex.Message & Environment.NewLine & ex.StackTrace & Environment.NewLine
            Dim st As StackTrace = New StackTrace(ex, True)
            For Each sf As StackFrame In st.GetFrames
                If sf.GetFileLineNumber() > 0 Then
                    Result &= "Line:" & sf.GetFileLineNumber() & " Filename: " & IO.Path.GetFileName(sf.GetFileName) & Environment.NewLine
                End If
            Next
            common.Logger(psIngestionLogFile, Result.ToString)
        End Try
        btnProcessandRunSelectedUsers.Enabled = True
    End Sub
End Class
