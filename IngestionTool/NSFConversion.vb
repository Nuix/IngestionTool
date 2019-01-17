Imports System.IO
Imports System.Diagnostics
Imports System.Net
Imports System.Threading
Imports System.Xml
Imports System.Text
Imports System.Data.SQLite
Imports CommonFunctions


Public Class NSFConversion
    Public psSettingsFile As String
    Public psIngestionLogFile As String
    Public psSQLiteLocation As String
    Public pbNoMoreJobs As Boolean
    Public psSummaryReportFile As String
    Public psNuixCaseDir As String
    Public plstNuixConsoleProcesses As List(Of String)

    Private SummaryReportThread As System.Threading.Thread
    Private SQLiteDBReadThread As System.Threading.Thread
    Private CaseLockThread As System.Threading.Thread

    Private Sub btnFileSystemChooser_Click(sender As Object, e As EventArgs) Handles btnFileSystemChooser.Click
        Dim fldrDialog As New FolderBrowserDialog
        Dim sSelectedPath As String

        fldrDialog.ShowNewFolderButton = True

        If (fldrDialog.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            sSelectedPath = fldrDialog.SelectedPath
            txtNSFLocation.Text = sSelectedPath
            btnLoadPSTInfo.Enabled = True
        Else
            btnLoadPSTInfo.Enabled = False
        End If
    End Sub

    Private Sub btnShowSettings_Click(sender As Object, e As EventArgs) Handles btnShowSettings.Click
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

    Private Sub btnLoadPSTInfo_Click(sender As Object, e As EventArgs) Handles btnLoadPSTInfo.Click
        Dim bStatus As Boolean

        If txtNSFLocation.Text = vbNullString Then
            MessageBox.Show("You must select a Source Data Location.", "Select Source Data Location", MessageBoxButtons.OK)
            txtNSFLocation.Focus()
            Exit Sub
        Else
            bStatus = blnLoadCustodianSourceInfo(txtNSFLocation.Text, grdConversionDataInfo, cboSourceDataType.Text)

        End If

    End Sub

    Private Function blnLoadCustodianSourceInfo(ByVal sSourceLocation As String, ByVal grdConversionDataInfo As DataGridView, sSourceDataType As String) As Boolean
        blnLoadCustodianSourceInfo = False

        Dim bStatus As Boolean
        Dim sCustodianName As String
        Dim iCustodianID As Integer
        Dim iRowIndex As Integer
        Dim ConversionRow As DataGridViewRow
        Dim sSourceFormat As String
        Dim sOutputFormat As String
        Dim CustodianDir As DirectoryInfo
        Dim iNumberofSourceDataFiles As Integer
        Dim dblTotalSizeofSourceData As Double
        Dim dbService As New DatabaseService

        If radSourceUser.Checked = True Then
            sSourceFormat = "User"
        Else
            sSourceFormat = "Flat"
        End If

        If radOutputUser.Checked = True Then
            sOutputFormat = "User"
        Else
            sOutputFormat = "Flat"
        End If

        'diDirectories = New DirectoryInfo(sNSFLocation)
        For Each custdir In Directory.GetDirectories(sSourceLocation)
            CustodianDir = New DirectoryInfo(custdir)
            sCustodianName = CustodianDir.Name.ToString

            subSourceDirSearch(custdir, iNumberofSourceDataFiles, dblTotalSizeofSourceData, sSourceDataType)

            If iNumberofSourceDataFiles > 0 Then
                iRowIndex = grdConversionDataInfo.Rows.Add()
                ConversionRow = grdConversionDataInfo.Rows(iRowIndex)
                bStatus = dbService.UpdateDataConversionStatBatchInfo(psSQLiteLocation, sCustodianName, iCustodianID, "", "Conversion Not Started", "", "Host:" & eMailArchiveMigrationManager.RedisHost & "|Port:" & eMailArchiveMigrationManager.RedisPort & "|Auth:" & eMailArchiveMigrationManager.RedisAuth & "|key_to_use:" & sCustodianName, cboSourceDataType.Text, sSourceFormat, cboOutputDataType.Text, sOutputFormat, "", "", 0, 0, 0, 0, "")

                iNumberofSourceDataFiles = 0
                dblTotalSizeofSourceData = 0

                With ConversionRow
                    ConversionRow.Cells("SelectCustodian").Value = False
                    ConversionRow.Cells("CustodianName").Value = sCustodianName
                    ConversionRow.Cells("ConversionStatus").Value = "Conversion Not Started"

                    ConversionRow.Cells("GroupID").Value = ""
                    ConversionRow.Cells("PercentComplete").Value = 0
                    ConversionRow.Cells("BytesProcessed").Value = 0
                    ConversionRow.Cells("NumberOfSourceFiles").Value = iNumberofSourceDataFiles
                    ConversionRow.Cells("SizeOfSourceFiles").Value = dblTotalSizeofSourceData
                    ConversionRow.Cells("SourcePath").Value = CustodianDir.FullName.ToString
                    ConversionRow.Cells("SourceDataType").Value = cboSourceDataType.Text
                    ConversionRow.Cells("SourceFormat").Value = sSourceFormat
                    ConversionRow.Cells("OutputDataType").Value = cboOutputDataType.Text
                    ConversionRow.Cells("OutputFormat").Value = sOutputFormat
                    ConversionRow.Cells("RedisSettings").Value = eMailArchiveMigrationManager.RedisHost & "|Port:" & eMailArchiveMigrationManager.RedisPort & "|Auth:" & eMailArchiveMigrationManager.RedisAuth & "|key_to_use:" & sCustodianName
                    ConversionRow.Cells("ProcessID").Value = ""
                    ConversionRow.Cells("ConversionStartTime").Value = ""
                    ConversionRow.Cells("ConversionEndTime").Value = ""
                    ConversionRow.Cells("NumberSuccess").Value = 0
                    ConversionRow.Cells("NumberFailed").Value = 0
                    ConversionRow.Cells("SummaryReportLocation").Value = ""
                End With
            End If
        Next

        blnLoadCustodianSourceInfo = True
    End Function


    Private Sub btnCustomerMappingFile_Click(sender As Object, e As EventArgs) Handles btnCustomerMappingFile.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        With OpenFileDialog1
            .Filter = "CSV|*.csv"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            txtCustodianMappingFile.Text = OpenFileDialog1.FileName.ToString
        End If
    End Sub

    Private Sub btnLoadCustodianMappingData_Click(sender As Object, e As EventArgs) Handles btnLoadCustodianMappingData.Click
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

        Dim iCustodianIndex As Integer

        Dim sCurrentRow() As String
        Dim common As New Common

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
                    If sCurrentRow.Count > 4 Then
                        lstDuplicateMappingCustodianGroupID.Add(sCurrentRow(4))
                    End If
                Else
                    lstMappingCustodianName.Add(LCase(sCurrentRow(0)))
                    lstMappingCustodianGroupID.Add(sCurrentRow(1))
                End If
            Loop
            cboGroupIDs.Items.Add("")
            lstUniqueGroupID = lstMappingCustodianGroupID.Distinct.ToList
            For Each UniqueGroup In lstUniqueGroupID
                cboGroupIDs.Items.Add(UniqueGroup.ToString)
            Next

            For Each row In grdConversionDataInfo.Rows
                Try
                    If lstMappingCustodianName.Contains(LCase(row.cells("CustodianName").Value)) Then
                        iCustodianIndex = lstMappingCustodianName.IndexOf(LCase(row.cells("CustodianName").Value))
                        row.cells("GroupID").value = lstMappingCustodianGroupID(iCustodianIndex)
                        If row.cells("PercentageCompleted").value <> vbNullString Then
                            iPercentComplete = CInt(row.cells("PercentageCompleted").value.ToString.Replace("%", ""))
                        Else
                            iPercentComplete = 0
                        End If
                        'bStatus = blnUpdateSQLiteAllCustodiansInfo(row.cells("CustodianName").value, row.cells("PSTPath").value, row.cells("NumberOfPSTs").value, row.cells("SizeOfPSTs").value, row.cells("GroupID").value, row.cells("DestinationFolder").value, row.cells("DestinationRoot").value, row.cells("DestinationSMTP").value, False, False, False, row.cells("ProcessID").value, row.cells("IngestionStartTime").value, row.cells("IngestionEndTime").value, row.cells("BytesUploaded").value, iPercentComplete, row.cells("NumberSuccess").value, row.cells("NumberFailed").value, row.cells("SummaryReportLocation").value, row.cells("ProcessingCompleted").value, row.cells("ReprocessingInProgress").value, row.cells("ReprocessingCompleted").value)
                    End If

                Catch ex As Exception

                End Try
            Next
        End If
    End Sub

    Private Sub btnClearDates_Click(sender As Object, e As EventArgs) Handles btnClearDates.Click
        setDateTimePickerBlank(dateFromDate)
        setDateTimePickerBlank(dateToDate)
        lstDateRangeFilters.Items.Clear()
    End Sub
    Private Sub setDateTimePickerBlank(ByVal dateTimePicker As DateTimePicker)
        dateTimePicker.Visible = True
        dateTimePicker.Format = DateTimePickerFormat.Custom
        dateTimePicker.CustomFormat = "    "
    End Sub

    Private Sub dateFromDate_MouseDown(sender As Object, e As MouseEventArgs) Handles dateFromDate.MouseDown
        dateFromDate.Format = DateTimePickerFormat.Long
    End Sub

    Private Sub dateToDate_MouseDown(sender As Object, e As MouseEventArgs) Handles dateToDate.MouseDown
        dateToDate.Format = DateTimePickerFormat.Long
    End Sub

    Private Sub NSFConversion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim sMachineName As String
        Dim msgboxReturn As DialogResult
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

        sMachineName = System.Net.Dns.GetHostName
        psIngestionLogFile = eMailArchiveMigrationManager.NuixLogDir & "\" & "Email Conversion - " & sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".log"
        psSQLiteLocation = eMailArchiveMigrationManager.SQLiteDBLocation

        grpOutputFormat.Hide()
        grpSourceFormat.Hide()

    End Sub

    Private Sub btnLaunchSourceConversionBatches_Click(sender As Object, e As EventArgs) Handles btnLaunchSourceConversionBatches.Click
        Dim lstSelectedCustodiansInfo As List(Of String)
        Dim lstProcessingCustodians As List(Of String)
        Dim asCustodianInfo() As String
        Dim sSourceDirectory As String

        Dim sNuixAppMemory As String

        Dim bStatus As Boolean

        btnReloadEmailConversionData.Enabled = False

        plstNuixConsoleProcesses = New List(Of String)
        lstSelectedCustodiansInfo = New List(Of String)
        lstProcessingCustodians = New List(Of String)


        For Each Row In grdConversionDataInfo.Rows
            If Row.cells("SelectCustodian").value = True Then
                If (Row.cells("ConversionStatus").value <> "Conversion Completed" And Row.cells("ConversionStatus").value <> "Conversion In Progress") Then
                    lstSelectedCustodiansInfo.Add(Row.cells("CustodianName").value & "," & Row.cells("SourcePath").value & "," & Row.cells("SizeOfSourceFiles").value.ToString.Replace(",", ""))
                    lstProcessingCustodians.Add(Row.cells("CustodianName").value)
                    Row.cells("ConversionStatus").value = "Not Started"
                End If
            End If
        Next


        If lstSelectedCustodiansInfo.Count <= 0 Then
            MessageBox.Show("You must select at least one custodian to ingest data into Office 365 for.")
            Exit Sub
        End If

        sNuixAppMemory = CInt(eMailArchiveMigrationManager.O365NuixAppMemory) / 1000
        sNuixAppMemory = "-Xmx" & sNuixAppMemory & "g"

        My.Computer.FileSystem.CreateDirectory(eMailArchiveMigrationManager.NuixCaseDir)

        My.Computer.FileSystem.CreateDirectory(eMailArchiveMigrationManager.NuixLogDir)
        My.Computer.FileSystem.CreateDirectory(eMailArchiveMigrationManager.NuixJavaTempDir)
        My.Computer.FileSystem.CreateDirectory(eMailArchiveMigrationManager.NuixWorkerTempDir)
        My.Computer.FileSystem.CreateDirectory(eMailArchiveMigrationManager.NuixWorkerTempDir)

        grdConversionDataInfo.Sort(grdConversionDataInfo.Columns("SelectCustodian"), System.ComponentModel.ListSortDirection.Descending)

        For Each Custodian In lstSelectedCustodiansInfo
            asCustodianInfo = Split(Custodian.ToString, ",")
            bStatus = blnBuildCustodianNuixFiles(asCustodianInfo, sNuixAppMemory, cboOutputDataType.Text)

        Next

        SummaryReportThread = New System.Threading.Thread(AddressOf WatchCaseDirectory)
        SummaryReportThread.Start()

        SQLiteDBReadThread = New System.Threading.Thread(AddressOf SQLiteUpdates)
        SQLiteDBReadThread.Start()

    End Sub

    Public Sub WatchCaseDirectory()
        ' set the path to be watched
        FileSystemWatcher1.Path = psNuixCaseDir

        ' add event handlers
        AddHandler FileSystemWatcher1.Created, AddressOf OnCreated
    End Sub

    Private Sub OnCreated(sender As Object, e As IO.FileSystemEventArgs)
        Dim sCustodianName As String
        Dim lstCustodianName As List(Of String)
        Dim lstCustodianItemsExported As List(Of String)
        Dim lstCustodianItemsFailed As List(Of String)
        Dim lstCustodianBytesUploaded As List(Of String)
        Dim lstInvalidCustodian As List(Of String)
        Dim NuixConsoleProcess As Process
        Dim NuixConsoleProcessStartInfo As ProcessStartInfo

        Dim sCompletedFolder As String
        Dim sCurrentFolder As String
        Dim sInProgressFolder As String

        Dim sItemsExported As String
        Dim sItemsFailed As String
        Dim sBytesUploaded As String
        Dim sItemsProcessed As String
        Dim sItemsSkipped As String

        Dim iNumberItemsFailed As Double
        Dim iNumberItemsExported As Double
        Dim dPercentFailed As Decimal
        Dim sProcessID As String
        Dim aProcess As System.Diagnostics.Process
        Dim bStatus As Boolean

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
        lstInvalidCustodian = New List(Of String)
        lstProcessesNotRunning = New List(Of String)

        Try
            'Threading.Thread.Sleep(2000)
            common.Logger(psIngestionLogFile, "Loading Data from - " & e.FullPath)
            bStatus = blnGetSummaryReportInfo(e.FullPath, sCustodianName, sItemsProcessed, sItemsExported, sItemsFailed, sItemsSkipped, sBytesUploaded, dStartDate, dEndDate)

            For Each Row In grdConversionDataInfo.Rows
                If Row.Cells("CustodianName").value = sCustodianName Then
                    sProcessID = Row.Cells("ProcessID").value
                    dEndTime = Now
                    Row.cells("ConversionStatus").readonly = True
                    Row.cells("ConversionStatus").value = "Conversion Completed"
                    Row.cells("ProcessID").value = vbNullString
                    Row.defaultcellstyle.Forecolor = Color.Green
                    Row.cells("StatusImage").value = My.Resources.Green_check_small
                    Row.Cells("ConversionEndTime").value = dEndTime
                    Row.cells("BytesProcessed").value = FormatNumber(Convert.ToDouble(sBytesUploaded), , , , TriState.True)
                    Row.Cells("NumberSuccess").value = FormatNumber(Convert.ToInt32(sItemsExported), , , , TriState.True)
                    Row.Cells("NumberFailed").value = FormatNumber(Convert.ToInt32(sItemsFailed), , , , TriState.True)
                    Row.cells("SummaryReportLocation").value = e.FullPath
                    Row.cells("ConversionCompleted").value = True
                    'bStatus = blnUpdateSQLiteDBCustodianInfo(sCustodianName, Row.cells("PSTPath").value, False, False, True, Row.cells("ProcessID").value, Row.cells("IngestionEndTime").value, Row.cells("NumberSuccess").value.ToString.Replace(",", ""), Row.cells("NumberFailed").value.ToString.Replace(",", ""), Row.cells("SummaryReportLocation").value)
                End If
            Next

            iNuixInstancesRequested = Convert.ToInt16(eMailArchiveMigrationManager.O365NuixInstances)
            plstNuixConsoleProcesses.RemoveAt(plstNuixConsoleProcesses.IndexOf(sProcessID))
            common.Logger(psIngestionLogFile, "Checking Number of Processes Running ")

            For Each ConsoleProcess In plstNuixConsoleProcesses
                bStatus = blnCheckIfProcessIsRunning(ConsoleProcess.ToString)
                If bStatus = False Then
                    lstProcessesNotRunning.Add(ConsoleProcess.ToString)
                    For Each custodianrow In grdConversionDataInfo.Rows
                        If custodianrow.cells("ProcessID").value = ConsoleProcess.ToString Then
                            If (File.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\" & custodianrow.cells("CustodianName").value & "\summary-report.txt")) Then
                                bStatus = blnGetSummaryReportInfo(eMailArchiveMigrationManager.NuixCaseDir & "\" & custodianrow.cells("CustodianName").value & "\summary-report.txt", sCustodianName, sItemsProcessed, sItemsExported, sItemsFailed, sItemsSkipped, sBytesUploaded, dStartDate, dEndDate)
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
                            Else
                                custodianrow.cells("ProcessID").value = vbNullString
                                custodianrow.cells("StatusImage").value = My.Resources.yellow_info_small
                                custodianrow.cells("ProgressStatus").value = "Process Terminated"
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
            For Each row In grdConversionDataInfo.Rows
                If iNumberOfNuixInstancesRunning < iNuixInstancesRequested Then
                    If row.Cells("ProgressStatus").value = "Not Started" Then
                        dStartDate = Now
                        sCustodianName = row.Cells("CustodianName").Value
                        'bStatus = blnGetCustodianPSTLocation(sCustodianName, sCurrentFolder)
                        sInProgressFolder = sCurrentFolder.Replace("Not Started", "In Progress")
                        sInProgressFolder = sCurrentFolder.Replace("not started", "in progress")
                        common.Logger(psIngestionLogFile, "Moving - " & sCustodianName & " from " & sCurrentFolder & " to " & sInProgressFolder)
                        'bStatus = blnMoveCustodianFolder(sCustodianName, sCurrentFolder, sInProgressFolder, lstInvalidCustodian)
                        common.Logger(psIngestionLogFile, "Launching Nuix...")
                        NuixConsoleProcessStartInfo = New ProcessStartInfo(eMailArchiveMigrationManager.NuixFilesDir & "\email conversion\Nuix Files\" & sCustodianName & "\" & sCustodianName & ".bat")
                        NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        'NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Minimized
                        NuixConsoleProcess = System.Diagnostics.Process.Start(NuixConsoleProcessStartInfo)
                        row.Cells("StatusImage").Value = My.Resources.inprogress_medium
                        row.Cells("ProgressStatus").Value = "In Progress"
                        row.Cells("ProcessID").Value = NuixConsoleProcess.Id
                        row.Cells("IngestionStartTime").Value = dStartDate
                        row.Cells("PSTPath").Value = sInProgressFolder
                        common.Logger(psIngestionLogFile, "Updating Status for " & row.Cells("CustodianName").Value)
                        bStatus = dbService.UpdateCustodianDBInfo(psSQLiteLocation, row.Cells("CustodianName").Value, "ConversionStatus", row.Cells("ConversionStatus").Value)
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

    Public Sub SQLiteUpdates()

        Dim lstBytesUploaded As List(Of String)
        Dim lstPercentCompleted As List(Of String)
        Dim lstProcessingCustodians As List(Of String)
        Dim lstNotStartedCustodians As List(Of String)
        Dim lstInProgressCustodians As List(Of String)
        Dim lstActiveCustodians As List(Of String)
        Dim bStatus As Boolean
        Dim bNoMoreJobs As Boolean
        Dim dPercentCompleted As Decimal
        Dim dbService As New DatabaseService
        Dim common As New Common

        bNoMoreJobs = False
        Thread.Sleep(10000)
        lstBytesUploaded = New List(Of String)
        lstPercentCompleted = New List(Of String)
        lstProcessingCustodians = New List(Of String)
        lstActiveCustodians = New List(Of String)
        lstInProgressCustodians = New List(Of String)
        lstNotStartedCustodians = New List(Of String)

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"


        Try
            Do While bNoMoreJobs = False

                bStatus = dbService.GetUpdatedDataConversionCustodianInfo(sSQLiteDatabaseFullName, lstProcessingCustodians, lstBytesUploaded, lstPercentCompleted)
                If lstProcessingCustodians.Count > 0 Then
                    For Each Custodian In lstProcessingCustodians
                        For Each row In grdConversionDataInfo.Rows
                            If row.cells("CustodianName").value = Custodian.ToString Then

                                row.cells("BytesProcessed").value = FormatNumber(lstBytesUploaded(lstProcessingCustodians.IndexOf(Custodian.ToString)), 0, , , TriState.True)
                                dPercentCompleted = CInt(lstPercentCompleted(lstProcessingCustodians.IndexOf(Custodian.ToString)))
                                dPercentCompleted = dPercentCompleted / 100
                                row.cells("PercentageCompleted").value = FormatPercent(dPercentCompleted)
                            End If
                        Next
                    Next
                    pbNoMoreJobs = False
                Else
                    pbNoMoreJobs = True
                End If
                lstBytesUploaded.Clear()
                lstPercentCompleted.Clear()
                lstProcessingCustodians.Clear()

                bNoMoreJobs = pbNoMoreJobs

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

        bStatus = blnRefreshGridFromDB(psSQLiteLocation, grdConversionDataInfo, lstNotStartedCustodians, lstInProgressCustodians)
        common.Logger(psIngestionLogFile, "All Currently select Processing Jobs have completed.  If necessary select more custodians for Processing.")
        MessageBox.Show("All Currently select Processing Jobs have completed.  If necessary select more custodians for Processing.", "All Processing Jobs Completed", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification, False)
        btnReloadEmailConversionData.Enabled = True


    End Sub



    Public Function blnInsertDataConversionStats(ByVal sSQLLiteLocation As String, ByVal sCustodianName As String, ByRef iCustodianID As Integer, ByVal sGroupID As String, ByVal sStatus As String, ByVal sProcessID As String, ByVal sRedisSettings As String, ByVal sSourceType As String, ByVal sSourceFormat As String, ByVal sOutputType As String, ByVal sOutputFormat As String, ByVal sConversionStartTime As String, ByVal sConversionEndTime As String, ByVal iBytesProcessed As Integer, ByVal iPercentComplete As Integer, ByVal iSuccess As Integer, ByVal iFailed As Integer, ByVal sSummaryReportLocation As String) As Boolean
        blnInsertDataConversionStats = False
        Dim sInsertArchiveExtractionData As String
        Dim sUpdateEArchiveExtractionData As String
        Dim sQueryReturnedCustodianID As String
        Dim SQLiteConnection As SQLiteConnection

        SQLiteConnection = New SQLiteConnection("Data Source=" & sSQLLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
        SQLiteConnection.Open()

        Using Connection As New SQLiteConnection("Data Source=" & sSQLLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;")
            Using SQLSelectCommand As New SQLiteCommand("SELECT RowID FROM DataConversionStats WHERE CutodianName='" & sCustodianName & "'")
                With SQLSelectCommand
                    .Connection = SQLiteConnection
                    Using readerObject As SQLiteDataReader = SQLSelectCommand.ExecuteReader
                        While readerObject.Read
                            sQueryReturnedCustodianID = readerObject("RowID").ToString
                        End While
                    End Using
                End With
            End Using

            If sQueryReturnedCustodianID = vbNullString Then
                sInsertArchiveExtractionData = "Insert into DataConversionStats (CustodianName,ConversionStatus,GroupID,PercentCompleted,ProcessID,RedisSettings,SourceType,SourceFormat,OutputType,OutputFormat,ConversionStartTime,ConversionEndTime,BytesProcessed,PercentCompleted,Success,Failed,SummaryReportLocation) Values "
                sInsertArchiveExtractionData = sInsertArchiveExtractionData & "(@CustodianName,@ConversionStatus,@GroupID,@PercentCompleted,@ProcessID,@RedisSettings,@SourceType,@SourceFormat,@OutputType,@OutputFormat,@ConversionStartTime,@ConversionEndTime,@BytesProcessed,@PercentCompleted,@Success,@Failed,@SummaryReportLocation)"
                Using oInsertArchiveExtractionDataCommand As New SQLiteCommand()
                    With oInsertArchiveExtractionDataCommand
                        .Connection = Connection
                        .CommandText = sInsertArchiveExtractionData
                        .Parameters.AddWithValue("@CustodianName", sCustodianName)
                        .Parameters.AddWithValue("@ConversionStatus", sStatus)
                        .Parameters.AddWithValue("@GroupID", sGroupID)
                        .Parameters.AddWithValue("@PercentCompleted", iPercentComplete)
                        .Parameters.AddWithValue("@ProcessID", sProcessID)
                        .Parameters.AddWithValue("@RedisSettings", sRedisSettings)
                        .Parameters.AddWithValue("@SourceType", sSourceType)
                        .Parameters.AddWithValue("@SourceFormat", sSourceFormat)
                        .Parameters.AddWithValue("@OutputType", sOutputType)
                        .Parameters.AddWithValue("@OutputFormat", sOutputFormat)
                        .Parameters.AddWithValue("@ConversionStartTime", sConversionStartTime)
                        .Parameters.AddWithValue("@ConversionEndTime", sConversionEndTime)
                        .Parameters.AddWithValue("@BytesProcessed", iBytesProcessed)
                        .Parameters.AddWithValue("@PercentCompleted", iPercentComplete)
                        .Parameters.AddWithValue("@Success", iSuccess)
                        .Parameters.AddWithValue("@Failed", iFailed)
                        .Parameters.AddWithValue("@SummaryReportLocation", sSummaryReportLocation)
                    End With
                    Try
                        Connection.Open()
                        oInsertArchiveExtractionDataCommand.ExecuteNonQuery()
                        Connection.Close()
                    Catch ex As Exception
                        'common.Logger(psIngestionLogFile, "Error 5 - " & ex.Message.ToString())
                    End Try
                End Using
            Else
                iCustodianID = CInt(sQueryReturnedCustodianID)
            End If
            SQLiteConnection.Close()
        End Using
        blnInsertDataConversionStats = True
    End Function

    Public Function blnInsertSourceDataStats(ByVal sSQLLiteLocation As String, ByVal sSourceFileName As String, ByVal sSourceFilePath As String, ByVal sSourceFileCreateDate As String, ByVal sSourceFileModifiedDate As String, ByVal dblSourceFileSize As Double, ByVal iCustodianID As Integer) As Boolean
        blnInsertSourceDataStats = False
        Dim sInsertSourceDataStatsQuery As String
        Dim sQueryReturnedCustodianID As String
        Dim SQLiteConnection As SQLiteConnection

        SQLiteConnection = New SQLiteConnection("Data Source=" & sSQLLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
        SQLiteConnection.Open()

        Using Connection As New SQLiteConnection("Data Source=" & sSQLLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;")
            Using SQLSelectCommand As New SQLiteCommand("SELECT CustodianID, SourceFileName, SourceFilePath, SourceFileCreateDate, SourceFileSize FROM SourceDataInfoStats WHERE CustodianID='" & iCustodianID & "' and SourceFileName='" & sSourceFileName & "' and SourceFilePath='" & sSourceFilePath & "' and SourceFileCreateDate='" & sSourceFileCreateDate & "' and SourceFileSize=" & dblSourceFileSize)
                With SQLSelectCommand
                    .Connection = SQLiteConnection
                    Using readerObject As SQLiteDataReader = SQLSelectCommand.ExecuteReader
                        While readerObject.Read
                            sQueryReturnedCustodianID = readerObject("CustodianID").ToString
                        End While
                    End Using
                End With

            End Using

            If sQueryReturnedCustodianID = vbNullString Then
                sInsertSourceDataStatsQuery = "Insert into SourceDataInfoStats (SourceFileName,SourceFilePath,SourceFileCreateDate,SourceFileModifiedDate,SourceFileSize,CustodianID) Values "
                sInsertSourceDataStatsQuery = sInsertSourceDataStatsQuery & "(@SourceFileName,@SourceFilePath,@SourceFileCreateDate,@SourceFileModifiedDate,@SourceFileSize,@CustodianID)"
                Using oInsertSourceDataStats As New SQLiteCommand()
                    With oInsertSourceDataStats
                        .Connection = Connection
                        .CommandText = sInsertSourceDataStatsQuery
                        .Parameters.AddWithValue("@SourceFileName", sSourceFileName)
                        .Parameters.AddWithValue("@SourceFilePath", sSourceFilePath)
                        .Parameters.AddWithValue("@SourceFileCreateDate", sSourceFileCreateDate)
                        .Parameters.AddWithValue("@SourceFileModifiedDate", sSourceFileModifiedDate)
                        .Parameters.AddWithValue("@SourceFileSize", dblSourceFileSize)
                        .Parameters.AddWithValue("@CustodianID", iCustodianID)
                    End With
                    Try
                        Connection.Open()
                        oInsertSourceDataStats.ExecuteNonQuery()
                        Connection.Close()
                    Catch ex As Exception
                        'common.Logger(psIngestionLogFile, "Error 5 - " & ex.Message.ToString())
                    End Try
                End Using
            Else
                iCustodianID = CInt(sQueryReturnedCustodianID)
            End If
            SQLiteConnection.Close()
        End Using
        blnInsertSourceDataStats = True
    End Function


    Private Function blnRefreshGridFromDB(ByVal sSQLLiteLocation As String, ByVal grdConversionDataInfo As DataGridView, ByRef lstNotStartedCustodians As List(Of String), ByRef lstInProgressCustodians As List(Of String)) As Boolean
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
        sqlConnection = New SQLiteConnection("Data Source=" & sSQLLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")

        For Each row In grdConversionDataInfo.Rows
            sCustodianName = row.cells("CustodianName").value
            mSQL = "select CustodianName, GroupID, ConversionStatus, ProcessID, RedisSettings, SourceType, SourceFormat, OutputType, OutputFormat, ConversionStartTime, ConversionEndTime, BytesProcessed, PercentCompleted, Success, Failed, SummaryReportLocation from DataConversionStats where CustodianName = '" & sCustodianName & "'"
            sqlCommand = New SQLiteCommand(mSQL, sqlConnection)
            sqlConnection.Open()

            dataReader = sqlCommand.ExecuteReader

            While dataReader.Read

                row.cells("CustodianName").value = dataReader.GetValue(0)
                row.cells("GroupID").value = dataReader.GetValue(1)
                row.cells("ConversionStatus").value = dataReader.GetValue(2)
                row.cells("ProcessID").value = dataReader.GetValue(3)
                row.cells("RedisSettings").value = dataReader.GetValue(4)
                row.cells("SourceType").value = dataReader.GetValue(5)
                row.cells("SourceFormat").value = dataReader.GetValue(6)
                row.cells("OutputType").value = dataReader.GetValue(7)
                row.cells("OutputFormat").value = sProgressStatus
                row.cells("ConversionStartTime").value = dataReader.GetValue(11)
                row.cells("ConversionEndTime").value = dataReader.GetValue(12)
                row.cells("BytesProcessed").value = dataReader.GetValue(13)
                row.cells("PercentCompleted").value = FormatPercent(dataReader.GetValue(14))
                row.cells("NumberSuccess").value = FormatNumber(dataReader.GetValue(15), 0, , , TriState.True)
                row.cells("NumberFailed").value = FormatNumber(dataReader.GetValue(16), 0, , , TriState.True)
                row.cells("SummaryReportLocation").value = dataReader.GetValue(18)

                If (CDbl(dataReader.GetValue(15)) > 0) Then
                    dPercentFailed = (CDbl(dataReader.GetValue(16)) / CDbl(dataReader.GetValue(15)))
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
                row.cells("StatusImage").value = My.Resources.not_selected_small

            End While
            sqlConnection.Close()

        Next
        blnRefreshGridFromDB = True
    End Function

    Private Sub btnShowSettings_MouseHover(sender As Object, e As EventArgs) Handles btnShowSettings.MouseHover
        EmailConversionToolTip.Show("Show Settings...", btnShowSettings)

    End Sub

    Private Sub btnLaunchSourceConversionBatches_MouseHover(sender As Object, e As EventArgs) Handles btnLaunchSourceConversionBatches.MouseHover
        EmailConversionToolTip.Show("Launch Conversion Batches", btnLaunchSourceConversionBatches)

    End Sub

    Private Sub btnReloadEmailConversionData_Click(sender As Object, e As EventArgs) Handles btnReloadEmailConversionData.Click
        Dim bStatus As Boolean
        grdConversionDataInfo.Rows.Clear()

        bStatus = blnLoadGridFromDB(grdConversionDataInfo)

    End Sub

    Private Function blnLoadGridFromDB(ByVal grdArchiveBatches As DataGridView) As Boolean
        Dim mSQL As String
        Dim bStatus As Boolean
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
        Dim iRowIndex As Integer
        Dim sCustodianName As String
        Dim sProgressStatus As String
        Dim dPercentFailed As Decimal

        Dim sSummaryReportCustodian As String
        Dim sSummaryReportItemsExported As String
        Dim sSummaryReportItemsFailed As String
        Dim sSummaryReportBytesUploaded As String
        Dim dSummaryReportStartDate As Date
        Dim dSummaryReportEndDate As Date
        blnLoadGridFromDB = False

        dt = Nothing
        ds = New DataSet
        sqlConnection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
        mSQL = "select CustodianName, GroupID, ConversionStatus, ProcessID, RedisSettings, SourceType, SourceFormat, OutputType, OutputFormat, ConversionStartTime, ConversionEndTime, BytesProcessed, PercentCompleted, Success, Failed, CaseDirectory, OutputDirectory, LogDirectory, SummaryReportLocation from DataConversionStats"
        sqlCommand = New SQLiteCommand(mSQL, sqlConnection)
        sqlConnection.Open()

        dataReader = sqlCommand.ExecuteReader

        While dataReader.Read
            iRowIndex = grdArchiveBatches.Rows.Add()

            BatchRow = grdArchiveBatches.Rows(iRowIndex)

            BatchRow.Cells("CustodianName").Value = dataReader.GetValue(0)
            BatchRow.Cells("GroupID").Value = dataReader.GetValue(1)
            BatchRow.Cells("ConversionStatus").Value = dataReader.GetValue(2)
            BatchRow.Cells("ProcessID").Value = dataReader.GetValue(3)
            BatchRow.Cells("RedisSettings").Value = dataReader.GetValue(4)
            BatchRow.Cells("SourceType").Value = dataReader.GetValue(5)
            BatchRow.Cells("SourceFormat").Value = dataReader.GetValue(6)
            BatchRow.Cells("OutputType").Value = dataReader.GetValue(7)
            BatchRow.Cells("OutputFormat").Value = sProgressStatus
            BatchRow.Cells("ConversionStartTime").Value = dataReader.GetValue(11)
            BatchRow.Cells("ConversionEndTime").Value = dataReader.GetValue(12)
            BatchRow.Cells("BytesProcessed").Value = dataReader.GetValue(13)
            BatchRow.Cells("PercentCompleted").Value = FormatPercent(dataReader.GetValue(14))
            BatchRow.Cells("NumberSuccess").Value = FormatNumber(dataReader.GetValue(15), 0, , , TriState.True)
            BatchRow.Cells("NumberFailed").Value = FormatNumber(dataReader.GetValue(16), 0, , , TriState.True)
            BatchRow.Cells("CaseDirectory").Value = dataReader.GetValue(18)
            BatchRow.Cells("OutputDirectory").Value = dataReader.GetValue(19)
            BatchRow.Cells("LogDirectory").Value = dataReader.GetValue(20)
            BatchRow.Cells("SummaryReportLocation").Value = dataReader.GetValue(21)

            If (CDbl(dataReader.GetValue(15)) > 0) Then
                dPercentFailed = (CDbl(dataReader.GetValue(16)) / CDbl(dataReader.GetValue(15)))
            Else
                dPercentFailed = 1.0
            End If

            If (dPercentFailed > 0.25) Then
                BatchRow.DefaultCellStyle.ForeColor = Color.Red
                BatchRow.Cells("StatusImage").Value = My.Resources.red_stop_small
                BatchRow.Cells("SelectCustodian").ReadOnly = True
            ElseIf (dPercentFailed > 0.1) Then
                BatchRow.DefaultCellStyle.ForeColor = Color.Orange
                BatchRow.Cells("StatusImage").Value = My.Resources.yellow_info_small
                BatchRow.Cells("SelectCustodian").ReadOnly = True
            Else
                BatchRow.DefaultCellStyle.ForeColor = Color.Green
                BatchRow.Cells("StatusImage").Value = My.Resources.Green_check_small
                BatchRow.Cells("SelectCustodian").ReadOnly = True
            End If


        End While
        sqlConnection.Close()

        sqlConnection.Close()
        grdArchiveBatches.ScrollBars = ScrollBars.Both
        blnLoadGridFromDB = True

    End Function

    Private Sub btnExportProcessingDetails_Click(sender As Object, e As EventArgs) Handles btnExportProcessingDetails.Click
        Dim sReportFilePath As String
        Dim ReportOutputFile As StreamWriter
        Dim sReportType As String
        Dim sOutputFileName As String
        Dim sMachineName As String
        Dim sCustodianName As String

        Dim sPercentCompleted As String
        Dim sBytesProcessed As String
        Dim sConversionStatus As String
        Dim sGroupID As String
        Dim sNumberOfSourceFiles As String
        Dim sSourceFileSize As String
        Dim sSourcePath As String
        Dim sSourceDataType As String
        Dim sSourceFormat As String
        Dim sOutputDataType As String
        Dim sOuptutFormat As String
        Dim sRedisSettings As String
        Dim sProcessID As String
        Dim sConversionStartTime As String
        Dim sConversionEndTime As String
        Dim sSuccess As String
        Dim sFailed As String
        Dim sSummaryReportLocation As String
        Dim sConversionCompleted As String

        sMachineName = System.Net.Dns.GetHostName()
        sOutputFileName = "Legacy Email Conversion Processing Statistic - " & sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv"
        ReportOutputFile = New StreamWriter(eMailArchiveMigrationManager.NuixFilesDir & "\" & sOutputFileName)

        ReportOutputFile.WriteLine("Custodian Name, Percent Completed, Bytes Uploaded, Progress Status, Number of PSTs, Size of PSTs, PST Path, Destination Folder, Destination Root, Destination SMTP, Group ID, Process ID, Ingestion Start Time, Success, Failed, Summary Report Location")


        For Each row In grdConversionDataInfo.Rows
            sCustodianName = row.cells("BatchName").value
            If sCustodianName <> vbNullString Then

                sPercentCompleted = row.cells("PercentComplete").value
                sBytesProcessed = row.cells("BytesProcessed").value
                sConversionStatus = row.cells("ConversionStatus").value
                sGroupID = row.cells("GroupID").value
                sNumberOfSourceFiles = row.cells("NumberOfSourceFiles").value
                sSourceFileSize = row.cells("SizeOfSourceFiles").value
                sSourcePath = row.cells("SourcePath").value
                sSourceDataType = row.cells("SourceDataType").value
                sSourceFormat = row.cells("SourceFormat").value
                sOutputDataType = row.cells("OutputDataType").value
                sOuptutFormat = row.cells("OutputFormat").value
                sRedisSettings = row.cells("RedisSettings").value
                sProcessID = row.cells("ProcessID").value
                sConversionStartTime = row.cells("ConversionStartTime").value
                sConversionEndTime = row.cells("ConversionEndTime").value
                sSuccess = row.cells("NumberSuccess").value
                sFailed = row.cells("NumberFailed").value
                sSummaryReportLocation = row.cells("SummaryReportLocation").value
                sConversionCompleted = row.cells("ConversionCompleted").value

                ReportOutputFile.WriteLine("""" & sCustodianName & """" & "," & """" & sPercentCompleted & """" & "," & """" & sBytesProcessed & """" & "," & """" & sConversionStatus & """" & "," & """" & sGroupID & """" & "," & """" & sNumberOfSourceFiles & """" & "," & """" & sSourceFileSize & """" & "," & """" & sSourcePath & """" & "," & """" & sSourceDataType & """" & "," & """" & sSourceFormat & """" & "," & """" & sOutputDataType & """" & "," & """" & sOuptutFormat & """" & "," & """" & sRedisSettings & """" & "," & """" & sProcessID & """" & "," & """" & sConversionStartTime & """" & "," & """" & sConversionEndTime & """" & "," & """" & sSuccess & """" & "," & """" & sFailed & """" & "," & """" & sSummaryReportLocation & """" & "," & """" & sConversionCompleted & """")
            End If
        Next

        ReportOutputFile.Close()
        MessageBox.Show("Nuix Case Statistics report finished building.  Report located at: " & eMailArchiveMigrationManager.NuixFilesDir & "\" & sOutputFileName)
    End Sub



    Private Sub grdConversionDataInfo_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdConversionDataInfo.CellContentDoubleClick
        If ((e.ColumnIndex = 18) And (grdConversionDataInfo("ProcessingFilesDirectory", e.RowIndex).Value <> vbNullString)) Then
            If Directory.Exists(grdConversionDataInfo("ProcessingFilesDirectory", e.RowIndex).Value) Then
                Process.Start("explorer.exe", grdConversionDataInfo("ProcessingFilesDirectory", e.RowIndex).Value)
            End If

        ElseIf ((e.ColumnIndex = 19) And (grdConversionDataInfo("CaseDirectory", e.RowIndex).Value <> vbNullString)) Then
            If Directory.Exists(grdConversionDataInfo("CaseDirectory", e.RowIndex).Value) Then
                Process.Start("explorer.exe", grdConversionDataInfo("CaseDirectory", e.RowIndex).Value)
            End If
        ElseIf ((e.ColumnIndex = 20) And (grdConversionDataInfo("OutputDirectory", e.RowIndex).Value <> vbNullString)) Then
            If Directory.Exists(grdConversionDataInfo("OutputDirectory", e.RowIndex).Value) Then
                Process.Start("explorer.exe", grdConversionDataInfo("OutputDirectory", e.RowIndex).Value)
            End If
        ElseIf ((e.ColumnIndex = 21) And (grdConversionDataInfo("LogDirectory", e.RowIndex).Value <> vbNullString)) Then
            If Directory.Exists(grdConversionDataInfo("LogDirectory", e.RowIndex).Value) Then
                Process.Start("explorer.exe", grdConversionDataInfo("LogDirectory", e.RowIndex).Value)
            End If
        ElseIf ((e.ColumnIndex = 24) And (grdConversionDataInfo("SummaryReportLocation", e.RowIndex).Value <> vbNullString)) Then
            System.Diagnostics.Process.Start(grdConversionDataInfo("SummaryReportLocation", e.RowIndex).Value.ToString)
        End If
    End Sub

    Private Sub btnAddDateRangeFilter_Click(sender As Object, e As EventArgs) Handles btnAddDateRangeFilter.Click
        If ((dateFromDate.Text <> vbNullString) And (dateToDate.Text <> vbNullString)) Then
            lstDateRangeFilters.Items.Add(FormatDateTime(dateFromDate.Text, DateFormat.ShortDate) & "-" & FormatDateTime(dateToDate.Text, DateFormat.ShortDate))
        End If
    End Sub

    Private Sub btnAddToSelected_Click(sender As Object, e As EventArgs) Handles btnAddToSelected.Click
        Dim sDateRangeFilters As String

        For i = 0 To lstDateRangeFilters.Items.Count - 1
            If sDateRangeFilters = vbNullString Then
                sDateRangeFilters = lstDateRangeFilters.Items.Item(i)
            Else
                sDateRangeFilters = sDateRangeFilters & "|" & lstDateRangeFilters.Items.Item(i)
            End If
        Next i
        For Each row In grdConversionDataInfo.Rows
            If row.cells("SelectCustodian").value = True Then
                row.cells("DateRangeFilter").value = sDateRangeFilters

            End If
        Next
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
        ConsolidateExporterMetricsFile = New StreamWriter(eMailArchiveMigrationManager.NuixFilesDir & "\email conversion\consolidated-exporter-metrics" & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv")

        bStatus = blnGetAllExporterMetrics(eMailArchiveMigrationManager.NuixCaseDir & "\email conversion", lstExporterMetricsFiles)

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
        bStatus = blnGetAllExporterErrors(eMailArchiveMigrationManager.NuixCaseDir & "\email conversion", lstExporterErrorsFiles)

        ConsolidateExporterErrorsFile = New StreamWriter(eMailArchiveMigrationManager.NuixFilesDir & "\email conversion\consolidated-exporter-errors" & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv")

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
        MessageBox.Show("Consolidated exporter-error files completed.  Report located at - " & eMailArchiveMigrationManager.NuixFilesDir & "\consolidated-exporter-error" & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv")

    End Sub
End Class