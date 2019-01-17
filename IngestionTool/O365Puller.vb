Imports System.IO
Imports System.Diagnostics
Imports System.Net
Imports System.Threading
Imports System.Xml
Imports System.Text
Imports System.Data.SQLite

Public Class O365Puller
    Public psSettingsFile As String
    Public pbSettingsLoaded As Boolean
    Public psNuixCaseDir As String
    Public plstCustodianGridInfo As List(Of String)
    Public plstCustodianGridIndex As List(Of Integer)
    Public psExtractionLogFile As String
    Public piNumberOfNuixInstancesRequested As Integer
    Public psSQLiteLocation As String
    Public pbNoMoreJobs As Boolean

    Private SummaryReportThread As System.Threading.Thread
    Private SQLiteDBReadThread As System.Threading.Thread
    Private CaseLockThread As System.Threading.Thread



    Private Sub O365Puller_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim msgboxReturn As DialogResult
        Dim sMachineName As String
        Dim common As New Common

        psSettingsFile = eMailArchiveMigrationManager.NuixSettingsFile
        If psSettingsFile = vbNullString Then
            msgboxReturn = MessageBox.Show("Nuix Email Archive Migration Manager has detected that the Global Settings have not been loaded.  Please load the settings now.", "Global Settings Not Available", MessageBoxButtons.YesNo)
            If (msgboxReturn = vbYes) Then
                Dim NuixSettings As O365ExtractionSettings
                NuixSettings = New O365ExtractionSettings
                NuixSettings.ShowDialog()
            Else
                Me.Close()
                Exit Sub

            End If
        End If
        sMachineName = System.Net.Dns.GetHostName

        FromDateTimePicker.Value = New Date(1990, 1, 1)
        ToDateTimePicker.Value = New Date(Today.Year, Today.Month, Today.Day)

        psExtractionLogFile = "EWS Extraction Log - " & sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".log"

        common.Logger(psExtractionLogFile, "Nuix Case Directory - " & eMailArchiveMigrationManager.NuixCaseDir)
        common.Logger(psExtractionLogFile, "Nuix Processing Files Directory - " & eMailArchiveMigrationManager.NuixFilesDir)
        common.Logger(psExtractionLogFile, "Nuix Log Directory - " & eMailArchiveMigrationManager.NuixLogDir)
        common.Logger(psExtractionLogFile, "Nuix Java Temp Directory - " & eMailArchiveMigrationManager.NuixJavaTempDir)
        common.Logger(psExtractionLogFile, "Nuix Worker Temp Directory - " & eMailArchiveMigrationManager.NuixWorkerTempDir)
        common.Logger(psExtractionLogFile, "Nuix Export Directory - " & eMailArchiveMigrationManager.NuixExportDir)
        common.Logger(psExtractionLogFile, "Nuix App Location - " & eMailArchiveMigrationManager.NuixAppLocation)
        common.Logger(psExtractionLogFile, "O365 Nuix Instances - " & eMailArchiveMigrationManager.O365NuixInstances)
        common.Logger(psExtractionLogFile, "O365 Nuix App Memory - " & eMailArchiveMigrationManager.O365NuixAppMemory)
        common.Logger(psExtractionLogFile, "O365 Number of Nuix Workers - " & eMailArchiveMigrationManager.O365NumberOfNuixWorkers)
        common.Logger(psExtractionLogFile, "O365 Memory per worker - " & eMailArchiveMigrationManager.O365MemoryPerWorker)
        common.Logger(psExtractionLogFile, "O365 Exchange Server - " & eMailArchiveMigrationManager.O365ExchangeServer)
        common.Logger(psExtractionLogFile, "O365 Domain - " & eMailArchiveMigrationManager.O365Domain)
        common.Logger(psExtractionLogFile, "O365 Admin Username - " & eMailArchiveMigrationManager.O365AdminUserName)
        common.Logger(psExtractionLogFile, "O365 Admin info - " & eMailArchiveMigrationManager.O365AdminInfo)
        common.Logger(psExtractionLogFile, "O365 Application Impersonation - " & eMailArchiveMigrationManager.O365ApplicationImpersonation)
        common.Logger(psExtractionLogFile, "O365 Retry Count - " & eMailArchiveMigrationManager.O365RetryCount)
        common.Logger(psExtractionLogFile, "O365 Retry Delay - " & eMailArchiveMigrationManager.O365RetryDelay)
        common.Logger(psExtractionLogFile, "O365 Retry Incrment - " & eMailArchiveMigrationManager.O365RetryIncrement)
        common.Logger(psExtractionLogFile, "O365 File Path Trimming - " & eMailArchiveMigrationManager.O365FilePathTrimming)
        common.Logger(psExtractionLogFile, "Archive Extration Nuix Workers - " & eMailArchiveMigrationManager.NuixWorkers)
        common.Logger(psExtractionLogFile, "Archive Extraction Nuix Instances - " & eMailArchiveMigrationManager.NuixInstances)
        common.Logger(psExtractionLogFile, "Archive Extraction Nuix App Memory - " & eMailArchiveMigrationManager.NuixAppMemory)
        common.Logger(psExtractionLogFile, "Archive Extraction Nuix Workers - " & eMailArchiveMigrationManager.NuixWorkerMemory)
        common.Logger(psExtractionLogFile, "Nuix NMS Source Type - " & eMailArchiveMigrationManager.NMSSourceType)
        common.Logger(psExtractionLogFile, "Nuix NMS Address - " & eMailArchiveMigrationManager.NuixNMSAddress)
        common.Logger(psExtractionLogFile, "Nuix NMS Port - " & eMailArchiveMigrationManager.NuixNMSPort)
        common.Logger(psExtractionLogFile, "Nuix NMS Username - " & eMailArchiveMigrationManager.NuixNMSUserName)
        common.Logger(psExtractionLogFile, "Nuix NMS info - " & eMailArchiveMigrationManager.NuixNMSAdminInfo)
        common.Logger(psExtractionLogFile, "SQLite DB Location - " & eMailArchiveMigrationManager.SQLiteDBLocation)

        psNuixCaseDir = eMailArchiveMigrationManager.NuixCaseDir
        psSQLiteLocation = eMailArchiveMigrationManager.SQLiteDBLocation
        piNumberOfNuixInstancesRequested = eMailArchiveMigrationManager.O365NuixInstances
        Me.StartPosition = FormStartPosition.CenterScreen
        'Me.Width = 1200
        'Me.Height = 600
        'grdCustodainSMTPInfo.Width = 1100
        'grdCustodainSMTPInfo.Height = 450

    End Sub

    Private Sub btnPullerCSVChooser_Click(sender As Object, e As EventArgs) Handles btnPullerCSVChooser.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        With OpenFileDialog1
            .Filter = "CSV|*.csv"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            txtPullerSMTPCSVFile.Text = OpenFileDialog1.FileName.ToString
        End If
    End Sub

    Private Sub btnLoadCustodianData_Click(sender As Object, e As EventArgs) Handles btnLoadCustodianData.Click
        Dim bStatus As Boolean
        If cboOutputType.Text = vbNullString Then
            MessageBox.Show("You must select a Lightspeed output format prior to loading the Custodian SMTP Information", "Select output format", MessageBoxButtons.OK)
            cboOutputType.Focus()
            Exit Sub

        End If
        cboPullerGroupIDs.Items.Add("")
        If txtPullerSMTPCSVFile.Text = vbNullString Then
            MessageBox.Show("You must select a custodian CSV file to load.", "Load Custodian CSV File")
            txtPullerSMTPCSVFile.Focus()
            Exit Sub
        End If
        bStatus = blnLoadSMTPInfoFromCSV(txtPullerSMTPCSVFile.Text, grdCustodainSMTPInfo, True, cboOutputType.Text)
    End Sub

    Public Function blnLoadSMTPInfoFromCSV(ByVal sSMTPCSVName As String, ByVal grdCustodainSMTPInfo As DataGridView, ByVal bUpdateGrid As Boolean, ByVal sOutputFormat As String) As Boolean
        Dim CSVMappingFile As Microsoft.VisualBasic.FileIO.TextFieldParser
        Dim asCurrentRow() As String
        Dim sCustodianPSTs As String
        Dim asCustodianPSTs() As String
        Dim iCounter As Integer
        Dim iNumberOfCustodianPSTs As Integer
        Dim bStatus As Boolean

        Dim lstCustodianName As List(Of String)
        Dim lstCustodianPSTs As List(Of String)
        Dim lstPSTSize As List(Of Double)
        Dim lstPSTPath As List(Of String)
        Dim lstGroupNumber As List(Of String)
        Dim CustodianRow As DataGridViewRow
        Dim sFromDate As String
        Dim sToDate As String
        Dim iRowIndex As Integer

        lstCustodianName = New List(Of String)
        lstCustodianPSTs = New List(Of String)
        lstPSTSize = New List(Of Double)
        lstPSTPath = New List(Of String)
        lstGroupNumber = New List(Of String)

        sFromDate = FromDateTimePicker.Value.ToString("yyyy-MM-dd")
        sToDate = ToDateTimePicker.Value.ToString("yyyy-MM-dd")


        blnLoadSMTPInfoFromCSV = False
        CSVMappingFile = New FileIO.TextFieldParser(sSMTPCSVName)
        CSVMappingFile.TextFieldType = FileIO.FieldType.Delimited
        CSVMappingFile.SetDelimiters(",")

        Try
            Do While Not CSVMappingFile.EndOfData
                asCurrentRow = CSVMappingFile.ReadFields
                If (asCurrentRow(0) <> "Custodian Name") Then

                    iRowIndex = grdCustodainSMTPInfo.Rows.Add()
                    CustodianRow = grdCustodainSMTPInfo.Rows(iRowIndex)

                    If asCurrentRow.Length >= 2 Then
                        CustodianRow.Cells("StatusImage").Value = My.Resources.not_selected_small
                        CustodianRow.Cells("SelectCustodian").Value = False
                        CustodianRow.Cells("SMTPAddress").Value = asCurrentRow(0)
                        CustodianRow.Cells("ExtractionRoot").Value = asCurrentRow(1)
                        CustodianRow.Cells("GroupID").Value = asCurrentRow(2)
                        CustodianRow.Cells("OutputFormat").Value = sOutputFormat

                        If Not (lstGroupNumber.Contains(asCurrentRow(2))) Then
                            cboPullerGroupIDs.Items.Add(asCurrentRow(2))
                            lstGroupNumber.Add(asCurrentRow(2))
                        End If
                    Else
                        CustodianRow.Cells("SelectCustodian").Value = False
                        CustodianRow.Cells("SMTPAddress").Value = asCurrentRow(0)
                        CustodianRow.Cells("ExtractionRoot").Value = asCurrentRow(1)
                        CustodianRow.Cells("OutputFormat").Value = sOutputFormat
                    End If
                    bStatus = blnUpdateSQLiteAllCustodiansInfo(CustodianRow.Cells("SMTPAddress").Value, CustodianRow.Cells("ExtractionRoot").Value, CustodianRow.Cells("GroupID").Value, CustodianRow.Cells("CaseDirectory").Value, CustodianRow.Cells("OutputDirectory").Value, CustodianRow.Cells("ExtractionStatus").Value, CustodianRow.Cells("ProcessID").Value, CustodianRow.Cells("ExtractionStartTime").Value, CustodianRow.Cells("ExtractionEndTime").Value, CustodianRow.Cells("ExtractedItems").Value, CustodianRow.Cells("ExtractedSize").Value, CustodianRow.Cells("SummaryReportLocation").Value, sFromDate, sToDate, sOutputFormat)
                End If
            Loop

            iCounter = 0
            For Each Custodian In lstCustodianName
                sCustodianPSTs = lstCustodianPSTs(iCounter)
                asCustodianPSTs = Split(sCustodianPSTs, ",")
                iNumberOfCustodianPSTs = UBound(asCustodianPSTs)
                iCounter = iCounter + 1
            Next

            blnLoadSMTPInfoFromCSV = True

        Catch ex As Exception
            MessageBox.Show("An invalid CSV file has been loaded.  Please check the csv file for the appropriate format.", "Invalid CSV File", MessageBoxButtons.OK)
        End Try
    End Function

    Private Sub btnSelectAll_Click(sender As Object, e As EventArgs) Handles btnSelectAll.Click
        If grdCustodainSMTPInfo.RowCount > 0 Then
            If btnSelectAll.Text = "Select All" Then

                For Each row In grdCustodainSMTPInfo.Rows
                    If row.cells("SMTPAddress").value <> vbNullString Then
                        row.cells("SelectCustodian").value = True
                    End If
                Next
                btnSelectAll.Text = "Deselect All"
            Else

                For Each row In grdCustodainSMTPInfo.Rows
                    row.cells("SelectCustodian").value = False
                Next
                btnSelectAll.Text = "Select All"

            End If
        End If
    End Sub

    Private Sub cboPullerGroupIDs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPullerGroupIDs.SelectedIndexChanged
        Dim sGroupID As String

        sGroupID = cboPullerGroupIDs.Text
        For Each row In grdCustodainSMTPInfo.Rows
            If sGroupID = vbNullString Then
                row.cells("SelectCustodian").value = False
            Else
                If row.cells("GroupID").value = sGroupID Then
                    row.cells("SelectCustodian").value = True
                Else
                    row.cells("SelectCustodian").value = False
                End If
            End If

        Next
    End Sub

    Private Sub btnExtractShowSettings_Click(sender As Object, e As EventArgs) Handles btnExtractShowSettings.Click
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
            eMailArchiveMigrationManager.NuixSettingsFile = psSettingsFile
        End If
        psSettingsFile = O365ExtractionSettings.NuixSettingsFile

    End Sub


    Private Sub btnLaunchO365Extractions_Click(sender As Object, e As EventArgs) Handles btnLaunchO365Extractions.Click
        Dim bStatus As Boolean
        Dim sDirectoryName As String
        Dim sRubyFileName As String
        Dim sJSonFileName As String
        Dim lstSelectedCustodiansInfo As List(Of String)
        Dim sFromDate As String
        Dim sToDate As String
        Dim sMigrationType As String
        Dim sOutputType As String
        Dim iNumberOfNuixProcessesRunning As Integer
        Dim iNumberofNuixInstances As Integer
        Dim msgboxReturn As DialogResult
        Dim sNuixAppMemory As String
        Dim sTimeout As String
        Dim iRunningInstances As Integer
        Dim dStartDate As DateTime
        Dim NuixConsoleProcess As Process
        Dim NuixConsoleProcessStartInfo As ProcessStartInfo
        Dim dFromDate As Date
        Dim dToDate As Date
        Dim common As New Common

        sMigrationType = "lightspeed"
        sOutputType = cboOutputType.Text

        btnLoadPreviousBatches.Enabled = True

        sNuixAppMemory = CInt(eMailArchiveMigrationManager.O365NuixAppMemory) / 1000
        sNuixAppMemory = Math.Round(CInt(sNuixAppMemory))
        sNuixAppMemory = "-Xmx" & sNuixAppMemory & "g"
        sTimeout = "100000"


        If Trim(FromDateTimePicker.Text) <> vbNullString Then
            dFromDate = FromDateTimePicker.Text
            dToDate = ToDateTimePicker.Text
            sFromDate = dFromDate.ToString("yyyy-MM-dd")
            sToDate = dToDate.ToString("yyyy-MM-dd")
        Else
            sFromDate = "NA"
            sToDate = "NA"
        End If

        lstSelectedCustodiansInfo = New List(Of String)

        For Each row In grdCustodainSMTPInfo.Rows
            If row.cells("SelectCustodian").value = True Then
                lstSelectedCustodiansInfo.Add(row.cells("SMTPAddress").value & "," & row.cells("ExtractionRoot").value)
            End If
        Next

        bStatus = blnCheckIfNuixIsRunning("nuix_app", iNumberOfNuixProcessesRunning)
        bStatus = blnCheckIfNuixIsRunning("nuix_console", iNumberOfNuixProcessesRunning)

        If (iNumberOfNuixProcessesRunning >= CInt(eMailArchiveMigrationManager.O365NuixInstances)) Then
            MessageBox.Show("There are already " & iNumberOfNuixProcessesRunning & ". Please wait until processes have finished or manually end processes")
            Exit Sub
        ElseIf (iNumberOfNuixProcessesRunning > 0) Then
            msgboxReturn = MessageBox.Show("There are already " & iNumberOfNuixProcessesRunning & " Nuix instances running and you have requested to launch " & eMailArchiveMigrationManager.O365NuixInstances & " instances of Nuix. Would you like to launch " & CInt(eMailArchiveMigrationManager.O365NuixInstances) - iNumberOfNuixProcessesRunning & " instance of Nuix?", "Nuix Instances Running", MessageBoxButtons.YesNo)
            If msgboxReturn = vbYes Then
                iNumberofNuixInstances = CInt(eMailArchiveMigrationManager.O365NuixInstances) - iNumberOfNuixProcessesRunning
            Else
                Exit Sub
            End If
        Else
            iNumberofNuixInstances = eMailArchiveMigrationManager.O365NuixInstances
        End If

        If lstSelectedCustodiansInfo.Count < iNumberofNuixInstances Then
            iNumberofNuixInstances = lstSelectedCustodiansInfo.Count
        End If

        grdCustodainSMTPInfo.Sort(grdCustodainSMTPInfo.Columns("SelectCustodian"), System.ComponentModel.ListSortDirection.Descending)
        iRunningInstances = 0

        For Each CustodianInfo In lstSelectedCustodiansInfo
            For Each row In grdCustodainSMTPInfo.Rows
                If row.cells("SelectCustodian").value = True Then
                    If iRunningInstances < iNumberofNuixInstances Then
                        dStartDate = Now
                        If File.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\" & row.cells("SMTPAddress").value & "\case.fbi2") Then
                            row.Cells("ProcessID").value = ""
                            row.cells("ExtractionStatus").value = "Failed (Case Exists)"
                            row.cells("StatusImage").value = My.Resources.red_stop_small
                            If Directory.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value) Then
                                row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value
                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value) Then
                                row.cells("OutputDirectory").value = eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value
                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value) Then
                                row.cells("LogDirectory").value = eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value

                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\") Then
                                row.cells("ProcessingFilesDirectory").value = eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                            End If
                        ElseIf (Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.cells("SMTPAddress").value)) Then
                            row.cells("ProcessID").value = ""
                            row.cells("StatusImage").value = My.Resources.red_stop_small
                            row.cells("ExtractionStatus").value = "Failed (Export Exists)"
                            If Directory.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value) Then
                                row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value
                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value) Then
                                row.cells("OutputDirectory").value = eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value
                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value) Then
                                row.cells("LogDirectory").value = eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value

                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\") Then
                                row.cells("ProcessingFilesDirectory").value = eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ProcessingFilesDirectory", eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                            End If
                        ElseIf (File.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\" & row.cells("SMTPAddress").value & "\case.fbi2") And (Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\" & row.cells("SMTPAddress").value))) Then
                            row.cells("ProcessID").value = ""
                            row.cells("StatusImage").value = My.Resources.red_stop_small
                            row.cells("ExtractionStatus").value = "Failed (Case and Export Exists)"
                            If Directory.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value) Then
                                row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value
                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value) Then
                                row.cells("OutputDirectory").value = eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value
                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value) Then
                                row.cells("LogDirectory").value = eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value

                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\") Then
                                row.cells("ProcessingFilesDirectory").value = eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ProcessingFilesDirectory", eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                            End If
                        Else
                            If Len(eMailArchiveMigrationManager.NuixCaseDir & "\" & row.cells("SMTPAddress").value) > 65 Then
                                row.cells("ExtractionStatus").value = "Failed (Case Path Too Long)"
                                row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                row.cells("OutputDirectory").value = eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                row.cells("ProcessingFilesDirectory").value = eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                row.cells("LogDirectory").value = eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                row.cells("StatusImage").value = My.Resources.red_stop_small
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ProcessingFilesDirectory", eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "CaseDirectory", eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "OutputDirectory", eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "LogDirectory", eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")

                            Else
                                common.Logger(psExtractionLogFile, "Processing Custodian - " & row.cells("SMTPAddress").value & "...")
                                common.Logger(psExtractionLogFile, "Launching Nuix...")

                                Directory.CreateDirectory(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                Directory.CreateDirectory(eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                Directory.CreateDirectory(eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                Directory.CreateDirectory(eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                sDirectoryName = eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                row.cells("OutputDirectory").value = eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                row.cells("ProcessingFilesDirectory").value = eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                row.cells("LogDirectory").value = eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ProcessingFilesDirectory", eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "CaseDirectory", eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "OutputDirectory", eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "LogDirectory", eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                sRubyFileName = sDirectoryName & row.cells("SMTPAddress").value & ".rb"
                                sJSonFileName = sDirectoryName & row.cells("SMTPAddress").value & ".json"
                                bStatus = blnBuildO365ExtractionBatchFiles(row.cells("SMTPAddress").value, sDirectoryName, sRubyFileName, sMigrationType, sOutputType, sNuixAppMemory)
                                bStatus = blnBuildEWSExtractJSonFile(row.cells("SMTPAddress").value, "", 0, "", row.cells("ExtractionRoot").value, "", sJSonFileName, "", 0, eMailArchiveMigrationManager.O365NumberOfNuixWorkers, row.cells("CaseDirectory").value, sFromDate, sToDate)
                                bStatus = blnBuildO365ExtractionRubyFiles(row.cells("SMTPAddress").value, sRubyFileName, sJSonFileName, row.cells("CaseDirectory").value)
                                NuixConsoleProcessStartInfo = New ProcessStartInfo(row.cells("ProcessingFilesDirectory").value & "\" & row.cells("SMTPAddress").value & ".bat")
                                NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                'NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Minimized

                                NuixConsoleProcess = System.Diagnostics.Process.Start(NuixConsoleProcessStartInfo)
                                iRunningInstances = iRunningInstances + 1
                                row.cells("StatusImage").value = My.Resources.inprogress_medium
                                row.cells("ProcessID").value = NuixConsoleProcess.Id
                                row.cells("ExtractionStatus").value = "In Progress"
                                row.DefaultCellStyle.Forecolor = Color.Black
                                row.cells("ExtractionStartTime").value = dStartDate
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ProcessID", NuixConsoleProcess.Id)
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ExtractionStatus", "In Progress")
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ExtractionStartTime", dStartDate)
                                row.cells("SelectCustodian").value = False
                                row.cells("SelectCustodian").readonly = True
                                Threading.Thread.Sleep(5000)
                            End If
                        End If
                    Else
                        If File.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\" & row.cells("SMTPAddress").value & "\case.fbi2") Then
                            row.Cells("ProcessID").value = ""
                            row.cells("ExtractionStatus").value = "Failed (Case Exists)"
                            row.cells("StatusImage").value = My.Resources.red_stop_small
                            If Directory.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value) Then
                                row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value
                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value) Then
                                row.cells("OutputDirectory").value = eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value
                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value) Then
                                row.cells("LogDirectory").value = eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value

                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\") Then
                                row.cells("ProcessingFilesDirectory").value = eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ProcessingFilesDirectory", eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                            End If
                        ElseIf (Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.cells("SMTPAddress").value)) Then
                            row.cells("ProcessID").value = ""
                            row.cells("StatusImage").value = My.Resources.red_stop_small
                            row.cells("ExtractionStatus").value = "Failed (Export Exists)"
                            If Directory.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value) Then
                                row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value
                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value) Then
                                row.cells("OutputDirectory").value = eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value
                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value) Then
                                row.cells("LogDirectory").value = eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value

                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\") Then
                                row.cells("ProcessingFilesDirectory").value = eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ProcessingFilesDirectory", eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                            End If
                        ElseIf (File.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\" & row.cells("SMTPAddress").value & "\case.fbi2") And (Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\" & row.cells("SMTPAddress").value))) Then
                            row.cells("ProcessID").value = ""
                            row.cells("StatusImage").value = My.Resources.red_stop_small
                            row.cells("ExtractionStatus").value = "Failed (Case and Export Exists)"
                            If Directory.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value) Then
                                row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value
                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value) Then
                                row.cells("OutputDirectory").value = eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value
                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value) Then
                                row.cells("LogDirectory").value = eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value

                            End If
                            If Directory.Exists(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\") Then
                                row.cells("ProcessingFilesDirectory").value = eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ProcessingFilesDirectory", eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                            End If

                        Else
                            If Len(eMailArchiveMigrationManager.NuixCaseDir & "\" & row.cells("SMTPAddress").value) > 65 Then
                                row.cells("ExtractionStatus").value = "Failed (Case Path Too Long)"
                                row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                row.cells("OutputDirectory").value = eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                row.cells("ProcessingFilesDirectory").value = eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                row.cells("LogDirectory").value = eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                row.cells("StatusImage").value = My.Resources.red_stop_small
                            Else
                                row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value
                                row.cells("OutputDirectory").value = eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value
                                row.cells("LogDirectory").value = eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.Cells("SMTPAddress").Value
                                row.cells("ProcessingFilesDirectory").value = eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ProcessingFilesDirectory", eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                row.cells("StatusImage").value = My.Resources.waitingtostart1
                                row.cells("ExtractionStatus").value = "Not Started"
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ExtractionStatus", "Not Started")
                                row.cells("SelectCustodian").value = False
                                row.cells("SelectCustodian").readonly = True
                                common.Logger(psExtractionLogFile, "Building Extraction Files for " & row.cells("SMTPAddress").value & "...")

                                Directory.CreateDirectory(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                sDirectoryName = eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\"
                                sRubyFileName = sDirectoryName & row.cells("SMTPAddress").value & ".rb"
                                sJSonFileName = sDirectoryName & row.cells("SMTPAddress").value & ".json"
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ProcessingFilesDirectory", eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "CaseDirectory", eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "OutputDirectory", eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")
                                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "LogDirectory", eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & row.cells("SMTPAddress").value & "\")

                                bStatus = blnBuildO365ExtractionBatchFiles(row.cells("SMTPAddress").value, sDirectoryName, sRubyFileName, sMigrationType, sOutputType, sNuixAppMemory)
                                bStatus = blnBuildEWSExtractJSonFile(row.cells("SMTPAddress").value, "", 0, "", row.cells("ExtractionRoot").value, "", sJSonFileName, "", 0, eMailArchiveMigrationManager.O365NumberOfNuixWorkers, eMailArchiveMigrationManager.NuixCaseDir, sFromDate, sToDate)
                                bStatus = blnBuildO365ExtractionRubyFiles(row.cells("SMTPAddress").value, sRubyFileName, sJSonFileName, row.cells("CaseDirectory").value)
                            End If
                        End If
                    End If
                End If
            Next
        Next

        grdCustodainSMTPInfo.Sort(grdCustodainSMTPInfo.Columns("ExtractionStatus"), System.ComponentModel.ListSortDirection.Descending)

        SQLiteDBReadThread = New System.Threading.Thread(AddressOf SQLiteUpdates)
        SQLiteDBReadThread.Start()

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

    Public Function blnCheckNuixLogForErrors(ByVal sNuixLogFileDir As String, ByRef sNuixLogFile As String, ByRef sErrorType As String) As Boolean
        Dim bStatus As Boolean
        Dim NuixLogStreamReader As StreamReader
        Dim sCurrentRow As String

        blnCheckNuixLogForErrors = False

        bStatus = blnGetNuixLogFiles(sNuixLogFileDir, sNuixLogFile)
        Try
            NuixLogStreamReader = New StreamReader(sNuixLogFile)
            While Not NuixLogStreamReader.EndOfStream
                sCurrentRow = NuixLogStreamReader.ReadLine
                If sCurrentRow.Contains("Error running script:") Then
                    blnCheckNuixLogForErrors = True
                    sErrorType = "Error Running Script"
                    Exit Function
                ElseIf sCurrentRow.Contains("FATAL com.nuix.investigator.main.b - Couldn't acquire a licence") Then
                    sErrorType = "No License"
                    blnCheckNuixLogForErrors = True
                    Exit Function
                ElseIf sCurrentRow.Contains("No licences were found.") Then
                    sErrorType = "No License"
                    blnCheckNuixLogForErrors = True
                    Exit Function
                ElseIf sCurrentRow.Contains("Error opening case") Then
                    blnCheckNuixLogForErrors = True
                    sErrorType = "Could not open Case"
                    Exit Function
                ElseIf sCurrentRow.Contains("Items have an invalid product class") Then
                    blnCheckNuixLogForErrors = True
                    sErrorType = "Invalid Product Class"
                    Exit Function
                ElseIf sCurrentRow.Contains("Couldn't acquire a licence") Then
                    sErrorType = "No License"
                    blnCheckNuixLogForErrors = True
                    Exit Function
                End If

            End While

        Catch ex As Exception
            'common.Logger(psUCRTLogFile, "Error in blnCheckNuixLogForErrors - " & ex.ToString)
        End Try
    End Function

    Private Function blnUpdateSQLiteAllCustodiansInfo(ByVal sCustodianName As String, ByVal sExtractionRoot As String, ByVal sGroupID As String, ByVal sCaseDirectory As String, ByVal sOutputDirectory As String, ByVal sExtractionStatus As String, ByVal sProcessID As String, ByVal dStartDate As Date, ByVal dEndDate As Date, ByVal dblExtractedItemCount As Double, ByVal dblExtractionSize As Double, ByVal sSummaryReportLocation As String, ByVal sFromDate As String, ByVal sToDate As String, ByVal sOutputFormat As String) As Boolean
        blnUpdateSQLiteAllCustodiansInfo = False
        Dim sInsertEWSExtractionDate As String
        Dim sUpdateEWSExtractionDate As String
        Dim sQueryReturnedCustodian As String
        Dim SQLiteConnection As SQLiteConnection
        Dim common As New Common

        SQLiteConnection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;New=False;Compress=True;")
        SQLiteConnection.Open()

        Using Connection As New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3")
            Using SQLSelectCommand As New SQLiteCommand("SELECT CustodianName FROM ewsExtractionStats WHERE CustodianName='" & sCustodianName & "'")
                With SQLSelectCommand
                    .Connection = SQLiteConnection
                    Using readerObject As SQLiteDataReader = SQLSelectCommand.ExecuteReader
                        While readerObject.Read
                            sQueryReturnedCustodian = readerObject("CustodianName").ToString
                        End While
                    End Using
                End With
            End Using

            SQLiteConnection.Close()

            If sQueryReturnedCustodian = vbNullString Then
                sInsertEWSExtractionDate = "Insert into ewsExtractionStats (CustodianName, ExtractionRoot, GroupID, CaseDirectory, OutputDirectory, ExtractionStatus, ProcessID, ExtractionStartTime, ExtractionEndTime, ProcessedItems, ExtractionSize, SummaryReportLocation, FromDate, ToDate, OutputFormat) Values "
                sInsertEWSExtractionDate = sInsertEWSExtractionDate & "(@CustodianName, @ExtractionRoot, @GroupID, @CaseDirectory, @OutputDirectory, @ExtractionStatus, @ProcessID, @ExtractionStartTime, @ExtractionEndTime, @ProcessedItems, @ExtractionSize, @SummaryReportLocation, @FromDate, @ToDate, @OutputFormat)"
                Using oInsertEWSExtractionDataCommand As New SQLiteCommand()
                    With oInsertEWSExtractionDataCommand
                        .Connection = Connection
                        .CommandText = sInsertEWSExtractionDate
                        .Parameters.AddWithValue("@CustodianName", sCustodianName)
                        .Parameters.AddWithValue("@ExtractionRoot", sExtractionRoot)
                        .Parameters.AddWithValue("@GroupID", sGroupID)
                        .Parameters.AddWithValue("@CaseDirectory", sCaseDirectory)
                        .Parameters.AddWithValue("@OutputDirectory", sOutputDirectory)
                        .Parameters.AddWithValue("@ExtractionStatus", sExtractionStatus)
                        .Parameters.AddWithValue("@ProcessID", sProcessID)
                        .Parameters.AddWithValue("@ExtractionStartTime", dStartDate.ToString)
                        .Parameters.AddWithValue("@ExtractionEndTime", dStartDate.ToString)
                        .Parameters.AddWithValue("@ProcessedItems", dblExtractedItemCount)
                        .Parameters.AddWithValue("@ExtractionSize", dblExtractionSize)
                        .Parameters.AddWithValue("@SummaryReportLocation", sSummaryReportLocation)
                        .Parameters.AddWithValue("@FromDate", sFromDate)
                        .Parameters.AddWithValue("@ToDate", sToDate)
                        .Parameters.AddWithValue("@OutputFormat", sOutputFormat)
                    End With
                    Try
                        Connection.Open()
                        oInsertEWSExtractionDataCommand.ExecuteNonQuery()
                        Connection.Close()
                    Catch ex As Exception
                        common.Logger(psExtractionLogFile, "Error 5 - " & ex.Message.ToString())
                    End Try
                End Using
            Else
                sUpdateEWSExtractionDate = "Update ewsExtractionStats set ExtractionRoot = @ExtractionRoot, GroupID = @GroupID, CaseDirectory = @CaseDirectory, OutputDirectory = @OutputDirectory, "
                sUpdateEWSExtractionDate = sUpdateEWSExtractionDate & "ExtractionStatus = @ExtractionStatus, ProcessID = @ProcessID, ExtractionStartTime = @ExtractionStartTime, ExtractionEndTime = @ExtractionEndTime, ProcessedItems = @ProcessedItems, ExtractionSize = @ExtractionSize, SummaryReportLocation = @SummaryReportLocation, FromDate = @FromDate, ToDate = @ToDate, OutputFormat = @OutputFormat "
                sUpdateEWSExtractionDate = sUpdateEWSExtractionDate & "WHERE CustodianName = @CustodianName"

                Using oUpdateEWSExtractionDataCommand As New SQLiteCommand()
                    With oUpdateEWSExtractionDataCommand
                        .Connection = Connection
                        .CommandText = sUpdateEWSExtractionDate
                        .Parameters.AddWithValue("@CustodianName", sCustodianName)
                        .Parameters.AddWithValue("@ExtractionRoot", sExtractionRoot)
                        .Parameters.AddWithValue("@GroupID", sGroupID)
                        .Parameters.AddWithValue("@CaseDirectory", sCaseDirectory)
                        .Parameters.AddWithValue("@GroupID", sGroupID)
                        .Parameters.AddWithValue("@OutputDirectory", sOutputDirectory)
                        .Parameters.AddWithValue("@ExtractionStatus", sExtractionStatus)
                        .Parameters.AddWithValue("@ProcessID", sProcessID)
                        .Parameters.AddWithValue("@ExtractionStartTime", dStartDate.ToString)
                        .Parameters.AddWithValue("@ExtractionEndTime", dStartDate.ToString)
                        .Parameters.AddWithValue("@ProcessedItems", dblExtractedItemCount)
                        .Parameters.AddWithValue("@ExtractionSize", dblExtractionSize)
                        .Parameters.AddWithValue("@SummaryReportLocation", sSummaryReportLocation)
                        .Parameters.AddWithValue("@FromDate", sFromDate)
                        .Parameters.AddWithValue("@ToDate", sToDate)
                        .Parameters.AddWithValue("@OutputFormat", sOutputFormat)
                    End With
                    Try
                        Connection.Open()
                        oUpdateEWSExtractionDataCommand.ExecuteNonQuery()
                        Connection.Close()
                    Catch ex As Exception
                        common.Logger(psExtractionLogFile, "Error 6 - " & ex.Message.ToString())
                    End Try
                End Using
            End If
            Connection.Close()
        End Using
        blnUpdateSQLiteAllCustodiansInfo = True
    End Function

    Private Sub CaseLock_Deleted(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs)
        Dim asFolderPath() As String
        Dim sSMTPAddress As String
        Dim dEndDate As DateTime
        Dim iNumberOfNuixInstancesRunning As Integer
        Dim bNextJobFound As Boolean
        Dim dStartDate As DateTime
        Dim iCounter As Integer
        Dim iTotalRows As Integer
        Dim iNuixInstancesRequested As Integer
        Dim bStatus As Boolean
        Dim NuixConsoleProcessStartInfo As ProcessStartInfo
        Dim NuixConsoleProcess As Process
        Dim sProcessedItems As String
        Dim dblTotalFolderSize As Double
        Dim sExportFolder As String
        Dim dbService As New DatabaseService

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"


        asFolderPath = Split(e.FullPath.ToString, "\")
        sSMTPAddress = asFolderPath(UBound(asFolderPath) - 1)
        For Each row In grdCustodainSMTPInfo.Rows
            If row.cells("SMTPAddress").value = sSMTPAddress Then
                dEndDate = Now
                row.cells("StatusImage").value = My.Resources.Green_check_small
                row.cells("ExtractionEndTime").value = dEndDate
                row.cells("ProcessID").value = vbNullString
                row.cells("ExtractionStatus").value = "Completed"
                bStatus = blnUpdateSQLiteExtrationDB(sSMTPAddress, "ExtractionStatus", "Completed")
                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ExtractionEndTime", dEndDate)
                bStatus = dbService.GetUpdatedEWSExtractInfo(sSQLiteDatabaseFullName, sSMTPAddress, "ProcessedItems", sProcessedItems)
                row.cells("ExtractedItems").value = FormatNumber(Convert.ToDouble(sProcessedItems), -1, , , TriState.True)
                sExportFolder = row.cells("OutputDirectory").value
                bStatus = blnComputeFolderSize(sExportFolder, dblTotalFolderSize)
                bStatus = blnUpdateSQLiteExtrationDB(row.cells("SMTPAddress").value, "ExtractionSize", dblTotalFolderSize)

                row.cells("ExtractedSize").value = dblTotalFolderSize.ToString("N0")
            End If
        Next

        bNextJobFound = False
        iCounter = 0
        iTotalRows = grdCustodainSMTPInfo.Rows.Count

        iNuixInstancesRequested = Convert.ToInt16(eMailArchiveMigrationManager.O365NuixInstances)

        If iNumberOfNuixInstancesRunning < iNuixInstancesRequested Then
            Do While bNextJobFound = False
                If iCounter < iTotalRows - 1 Then
                    If grdCustodainSMTPInfo.Rows(iCounter).Cells("ExtractionStatus").Value = "Not Started" Then
                        dStartDate = Now
                        sSMTPAddress = grdCustodainSMTPInfo.Rows(iCounter).Cells("SMTPAddress").Value
                        NuixConsoleProcessStartInfo = New ProcessStartInfo(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & sSMTPAddress & "\" & sSMTPAddress & ".bat")
                        NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        'NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Minimized

                        NuixConsoleProcess = System.Diagnostics.Process.Start(NuixConsoleProcessStartInfo)
                        grdCustodainSMTPInfo.Rows(iCounter).Cells("StatusImage").Value = My.Resources.inprogress_medium
                        grdCustodainSMTPInfo.Rows(iCounter).Cells("ExtractionStatus").Value = "In Progress"
                        grdCustodainSMTPInfo.Rows(iCounter).DefaultCellStyle.ForeColor = Color.Black
                        grdCustodainSMTPInfo.Rows(iCounter).Cells("ProcessID").Value = NuixConsoleProcess.Id
                        grdCustodainSMTPInfo.Rows(iCounter).Cells("ExtractionStartTime").Value = dStartDate
                        grdCustodainSMTPInfo.Rows(iCounter).Cells("CaseDirectory").Value = eMailArchiveMigrationManager.NuixCaseDir & "\" & grdCustodainSMTPInfo.Rows(iCounter).Cells("SMTPAddress").Value
                        grdCustodainSMTPInfo.Rows(iCounter).Cells("OutputDirectory").Value = eMailArchiveMigrationManager.NuixExportDir & "\" & grdCustodainSMTPInfo.Rows(iCounter).Cells("SMTPAddress").Value
                        bStatus = blnUpdateSQLiteExtrationDB(grdCustodainSMTPInfo.Rows(iCounter).Cells("SMTPAddress").Value, "ExtractionStatus", "In Progress")
                        bStatus = blnUpdateSQLiteExtrationDB(grdCustodainSMTPInfo.Rows(iCounter).Cells("SMTPAddress").Value, "ProcessID", NuixConsoleProcess.Id)
                        bStatus = blnUpdateSQLiteExtrationDB(grdCustodainSMTPInfo.Rows(iCounter).Cells("SMTPAddress").Value, "ExtractionStartTime", dStartDate)
                        bStatus = blnUpdateSQLiteExtrationDB(grdCustodainSMTPInfo.Rows(iCounter).Cells("SMTPAddress").Value, "CaseDirectory", eMailArchiveMigrationManager.NuixCaseDir & "\" & grdCustodainSMTPInfo.Rows(iCounter).Cells("SMTPAddress").Value)
                        bStatus = blnUpdateSQLiteExtrationDB(grdCustodainSMTPInfo.Rows(iCounter).Cells("SMTPAddress").Value, "OutputDirectory", eMailArchiveMigrationManager.NuixExportDir & "\" & grdCustodainSMTPInfo.Rows(iCounter).Cells("SMTPAddress").Value)
                        bNextJobFound = True
                    End If
                Else
                    bNextJobFound = True
                    Exit Sub
                End If
                iCounter = iCounter + 1
            Loop
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
            common.Logger(psExtractionLogFile, "Error In blnGetSummaryReportInfo")
            common.Logger(psExtractionLogFile, ex.Message)
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

    Private Function blnUpdateSQLiteExtrationDB(ByVal sCustodianName As String, ByVal sColumnName As String, ByVal sColumnValue As String) As Boolean
        blnUpdateSQLiteExtrationDB = False

        Dim sUpdateEWSExtractionDate As String
        Dim sParameterValue As String

        Dim SQLiteConnection As SQLiteConnection
        Dim common As New Common


        sParameterValue = "" & "@" & sColumnName & ""
        SQLiteConnection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;New=False;Compress=True;")
        SQLiteConnection.Open()
        Using Connection As New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3")


            sUpdateEWSExtractionDate = "Update ewsExtractionStats set " & sColumnName & " = " & sParameterValue
            sUpdateEWSExtractionDate = sUpdateEWSExtractionDate & " WHERE CustodianName = @CustodianName"

            Using oUpdateEWSExtractionDataCommand As New SQLiteCommand()
                With oUpdateEWSExtractionDataCommand
                    .Connection = Connection
                    .CommandText = sUpdateEWSExtractionDate
                    .Parameters.AddWithValue("@CustodianName", sCustodianName)
                    .Parameters.AddWithValue(sParameterValue, sColumnValue)
                End With
                Try
                    Connection.Open()
                    oUpdateEWSExtractionDataCommand.ExecuteNonQuery()
                    Connection.Close()
                Catch ex As Exception
                    common.Logger(psExtractionLogFile, "Error 6 - " & ex.Message.ToString())
                End Try
            End Using
            Connection.Close()
        End Using

        blnUpdateSQLiteExtrationDB = True
    End Function

    Private Function blnCheckIfNuixIsRunning(ByVal sProcessName As String, ByRef iNumberOfNuixInstancesRunning As Integer) As Boolean

        Dim NuixProcess() As System.Diagnostics.Process

        blnCheckIfNuixIsRunning = False

        NuixProcess = Process.GetProcessesByName(sProcessName)

        If NuixProcess.Length > 0 Then
            iNumberOfNuixInstancesRunning = iNumberOfNuixInstancesRunning + NuixProcess.Length
        End If

        blnCheckIfNuixIsRunning = True

    End Function
    Public Function blnProcessCustodian(ByVal sSMTPAddress As String, ByVal grdCustodianSMTPInfo As DataGridView) As Boolean
        Dim bFoundProcessedCustodian As Boolean
        Dim lstSelectedCustodians As List(Of String)
        Dim lstInvalidCustodian As List(Of String)
        Dim NuixConsoleProcess As Process
        Dim common As New Common

        lstSelectedCustodians = New List(Of String)
        lstInvalidCustodian = New List(Of String)

        Dim dStartDate As Date
        common.Logger(psExtractionLogFile, "Processing Selected Custodians...")

        blnProcessCustodian = False
        For Each row In grdCustodainSMTPInfo.Rows
            If row.cells("SMTPAddress").value = sSMTPAddress Then
                dStartDate = Now
                If Directory.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\" & sSMTPAddress) Then
                    row.Cells("ProcessID").value = "Case Already Exists"
                    row.cells("StatusImage").value = My.Resources.red_stop_small
                    row.cells("CaseDirectory").value = row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\" & sSMTPAddress
                ElseIf (Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\" & sSMTPAddress)) Then
                    row.cells("ProcessID").value = "Export Directory Exists"
                    row.cells("StatusImage").value = My.Resources.red_stop_small
                    row.cells("CaseDirectory").value = row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\" & sSMTPAddress
                ElseIf (Directory.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\" & sSMTPAddress) And (Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\" & sSMTPAddress))) Then
                    row.cells("ProcessID").value = "Case Directory and Export Directory Exists"
                    row.cells("StatusImage").value = My.Resources.red_stop_small
                    row.cells("CaseDirectory").value = row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\" & sSMTPAddress
                Else
                    common.Logger(psExtractionLogFile, "Processing Custodian - " & sSMTPAddress & "...")
                    common.Logger(psExtractionLogFile, "Launching Nuix...")
                    NuixConsoleProcess = System.Diagnostics.Process.Start(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & sSMTPAddress & "\" & sSMTPAddress & ".bat")
                    row.cells("StatusImage").value = My.Resources.inprogress_medium
                    row.cells("ProcessID").value = NuixConsoleProcess.Id
                    row.cells("ExtractionStatus").value = "In Progress"

                    row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\" & sSMTPAddress
                    row.cells("ExtractionStartTime").value = dStartDate
                    row.cells("SelectCustodian").readonly = True
                    bFoundProcessedCustodian = True
                End If
            End If
        Next

        blnProcessCustodian = True
    End Function

    Public Function blnProcessSelectedCustodians(ByVal lstProcessingCustodians As List(Of String), ByVal grdCustodianSMTPInfo As DataGridView, ByVal iNuixInstances As Integer) As Boolean
        Dim iCounter As Integer
        Dim sCustodianName As String
        Dim bFoundProcessedCustodian As Boolean
        Dim lstSelectedCustodians As List(Of String)
        Dim lstInvalidCustodian As List(Of String)
        Dim NuixConsoleProcess As Process
        Dim common As New Common

        lstSelectedCustodians = New List(Of String)
        lstInvalidCustodian = New List(Of String)

        Dim dStartDate As Date
        common.Logger(psExtractionLogFile, "Processing Selected Custodians...")

        blnProcessSelectedCustodians = False

        If grdCustodianSMTPInfo.RowCount - 1 <= iNuixInstances Then
            iNuixInstances = grdCustodianSMTPInfo.RowCount - 1
        End If

        For iCounter = 0 To iNuixInstances - 1
            bFoundProcessedCustodian = False

            For Each row In grdCustodainSMTPInfo.Rows
                If row.cells("SelectCustodian").value = True Then
                    If lstSelectedCustodians.Count < iNuixInstances Then
                        lstSelectedCustodians.Add(row.cells("SMTPAddress").value)
                        dStartDate = Now
                        sCustodianName = row.Cells("SMTPAddress").Value
                        If Directory.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & sCustodianName) Then
                            row.Cells("ProcessID").value = "Case Already Exists"
                            row.cells("StatusImage").value = My.Resources.red_stop_small
                            row.cells("CaseDirectory").value = row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & sCustodianName
                        ElseIf (Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & sCustodianName)) Then
                            row.cells("ProcessID").value = "Export Directory Exists"
                            row.cells("StatusImage").value = My.Resources.red_stop_small
                            row.cells("CaseDirectory").value = row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & sCustodianName
                        ElseIf (Directory.Exists(eMailArchiveMigrationManager.NuixCaseDir & "\" & sCustodianName) And (Directory.Exists(eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & sCustodianName))) Then
                            row.cells("ProcessID").value = "Case Directory and Export Directory Exists"
                            row.cells("StatusImage").value = My.Resources.red_stop_small
                            row.cells("CaseDirectory").value = row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & sCustodianName
                        Else
                            common.Logger(psExtractionLogFile, "Processing Custodian - " & sCustodianName & "...")
                            common.Logger(psExtractionLogFile, "Launching Nuix...")
                            NuixConsoleProcess = System.Diagnostics.Process.Start(eMailArchiveMigrationManager.NuixFilesDir & "\EWS Extraction\" & sCustodianName & "\" & sCustodianName & ".bat")
                            row.cells("StatusImage").value = My.Resources.inprogress_medium
                            row.cells("ProcessID").value = NuixConsoleProcess.Id
                            row.cells("ExtractionStatus").value = "In Progress"
                            row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\" & sCustodianName
                            row.cells("ExtractionStartTime").value = dStartDate
                            row.cells("SelectCustodian").readonly = True
                            bFoundProcessedCustodian = True
                        End If
                    Else
                        row.cells("CaseDirectory").value = eMailArchiveMigrationManager.NuixCaseDir & "\EWS Extraction\" & sCustodianName
                        row.cells("SelectCustodian").value = False
                        row.cells("ExtractionStatus").value = "Not Started"

                        row.cells("SelectCustodian").readonly = True
                    End If
                End If
            Next
        Next

        blnProcessSelectedCustodians = True
    End Function


    Private Function blnBuildO365ExtractionBatchFiles(ByVal sCustodianName As String, ByVal sDirectoryName As String, ByVal sRubyFileName As String, ByVal sMigrationType As String, ByVal sOutputType As String, ByVal sNuixAppMemory As String) As Boolean

        blnBuildO365ExtractionBatchFiles = False

        Dim CustodianBatchFile As StreamWriter

        Dim sStartUpBatchFileName As String
        Dim sLicenceSourceType As String

        sStartUpBatchFileName = sDirectoryName & "\" & sCustodianName & ".bat"
        sStartUpBatchFileName = sStartUpBatchFileName.Replace("\\", "\")

        If eMailArchiveMigrationManager.NMSSourceType = "Desktop" Then
            sLicenceSourceType = " -licencesourcetype dongle"
        ElseIf eMailArchiveMigrationManager.NMSSourceType = "Server" Then
            sLicenceSourceType = " -licencesourcetype server -licencesourcelocation " & eMailArchiveMigrationManager.NuixNMSAddress & ":" & eMailArchiveMigrationManager.NuixNMSPort & " -Dnuix.registry.servers=" & eMailArchiveMigrationManager.NuixNMSAddress
        ElseIf eMailArchiveMigrationManager.NMSSourceType = "Cloud Server" Then
            sLicenceSourceType = " -Dnuix.product.server.addClsProdEndpoint=true -licencesourcetype cloud-server"
        End If
        CustodianBatchFile = New StreamWriter(sStartUpBatchFileName)
        CustodianBatchFile.WriteLine("::Title will be the source mailbox in Office 365")
        CustodianBatchFile.WriteLine("@TITLE " & sCustodianName)
        CustodianBatchFile.WriteLine("::Enter NMS Username on Line 4")
        CustodianBatchFile.WriteLine("@SET NUIX_USERNAME=" & eMailArchiveMigrationManager.NuixNMSUserName)
        CustodianBatchFile.WriteLine("::Enter NMS Username on Line 6")
        CustodianBatchFile.WriteLine("@SET NUIX_PASSWORD=" & eMailArchiveMigrationManager.NuixNMSAdminInfo)

        CustodianBatchFile.Write("""" & eMailArchiveMigrationManager.NuixAppLocation & """" & sLicenceSourceType)
        CustodianBatchFile.Write(" -licencetype email-archive-examiner -licenceworkers " & eMailArchiveMigrationManager.O365NumberOfNuixWorkers & " " & sNuixAppMemory & " -Dnuix.logdir=" & """" & eMailArchiveMigrationManager.NuixLogDir & "\EWS Extraction\" & sCustodianName & """" & " -Djava.io.tmpdir=" & """" & eMailArchiveMigrationManager.NuixJavaTempDir & """" & " -Dnuix.worker.tmpdir=" & """" & eMailArchiveMigrationManager.NuixWorkerTempDir & """" & " -Dnuix.crackAndDump.exportDir=" & """" & eMailArchiveMigrationManager.NuixExportDir & "\EWS Extraction\" & sCustodianName & """")
        CustodianBatchFile.Write(" -Dnuix.processing.crackAndDump.useRelativePath=true -Dnuix.processing.worker.timeout=" & eMailArchiveMigrationManager.WorkerTimeout)
        If eMailArchiveMigrationManager.O365EnablePrefetch = True Then
            CustodianBatchFile.Write(" -Dnuix.data.exchangews.prefetchDisabled=false -Dnuix.data.exchangews.ewsMaxDownloadCount=" & eMailArchiveMigrationManager.O365MaxDownloadCount & " -Dnuix.data.exchangews.ewsMaxDownloadSize=" & eMailArchiveMigrationManager.O365MaxDownloadSize)
        Else
            CustodianBatchFile.Write(" -Dnuix.data.exchangews.prefetchDisabled=true")
        End If
        If eMailArchiveMigrationManager.O365EnableCollaborativePrefetch = True Then
            CustodianBatchFile.Write(" -Dnuix.data.exchangews.collaborativeFetching=true")
        Else
            CustodianBatchFile.Write(" -Dnuix.data.exchangews.collaborativeFetching=false")
        End If
        CustodianBatchFile.Write(" -Dnuix.processing.crackAndDump.prependEvidenceName=true")
        If sMigrationType = "lightspeed" Then
            CustodianBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & sOutputType)
        End If
        CustodianBatchFile.Write(" -Dnuix.data.exchangews.retryCount=" & eMailArchiveMigrationManager.O365RetryCount & " -Dnuix.data.exchangews.retryDelay=" & eMailArchiveMigrationManager.O365RetryDelay & " -Dnuix.data.exchangews.retryDelayIncrement=" & eMailArchiveMigrationManager.O365RetryIncrement & " -Dnuix.data.exchangews.impersonate=" & eMailArchiveMigrationManager.O365ApplicationImpersonation.ToString.ToLower & " ")
        CustodianBatchFile.WriteLine("""" & sRubyFileName & """")

        CustodianBatchFile.WriteLine("@exit")
        'CustodianBatchFile.WriteLine("@pause")

        CustodianBatchFile.Close()

        blnBuildO365ExtractionBatchFiles = True
    End Function

    Private Function blnBuildO365ExtractionRubyFiles(ByVal sCustodianName As String, ByVal sRubyScriptFileName As String, ByVal sCustodianJSonFile As String, ByVal sNuixCase As String) As Boolean
        blnBuildO365ExtractionRubyFiles = False

        Dim CustodianRuby As StreamWriter
        Dim sCaseDir As String

        If O365ExtractionSettings.NuixCaseDir.Contains("\") Then
            sCaseDir = eMailArchiveMigrationManager.NuixCaseDir.Replace("\", "\\")
        ElseIf O365ExtractionSettings.NuixCaseDir.Contains("/") Then
            sCaseDir = eMailArchiveMigrationManager.NuixCaseDir.Replace("/", "\\")
        End If

        CustodianRuby = New StreamWriter(sRubyScriptFileName)

        CustodianRuby.WriteLine("# Menu Title: EWS Puller")
        CustodianRuby.WriteLine("# Needs Selected Items: false")
        CustodianRuby.WriteLine("# ")
        CustodianRuby.WriteLine("# This script expects a JSON configured with O365 parameters completed in order")
        CustodianRuby.WriteLine("# to automatically extract data inside of an O365 mailbox, archive or purges.")
        CustodianRuby.WriteLine("# ")
        CustodianRuby.WriteLine("# Version 1.5")
        CustodianRuby.WriteLine("# Dec 16 2016 - Alex Chatzistamatis, Nuix")
        CustodianRuby.WriteLine("# ")
        CustodianRuby.WriteLine("#######################################")

        CustodianRuby.WriteLine("require 'thread'")
        CustodianRuby.WriteLine("require 'json'")

        CustodianRuby.WriteLine("CALLBACK_FREQUENCY = 100")
        CustodianRuby.WriteLine("callback_count = 0")

        CustodianRuby.WriteLine("#######################################")


        CustodianRuby.WriteLine("load """ & psSQLiteLocation.Replace("\", "\\") & "\\Database.rb_""")
        CustodianRuby.WriteLine("load """ & psSQLiteLocation.Replace("\", "\\") & "\\SQLite.rb_""")
        CustodianRuby.WriteLine("db = SQLite.new(""" & psSQLiteLocation.Replace("\", "\\") & "\\NuixEmailArchiveMigrationManager.db3" & """" & ")")

        CustodianRuby.WriteLine("#######################################")

        If sCustodianJSonFile.Contains("Not Started") Then
            sCustodianJSonFile = sCustodianJSonFile.Replace("Not Started", "In Progress")
        End If

        CustodianRuby.WriteLine("file = File.read('" & sCustodianJSonFile.Replace("\", "\\") & "')")
        CustodianRuby.WriteLine("parsed = JSON.parse(file)")

        CustodianRuby.WriteLine("  o365_server = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "server" & """" & "]")
        CustodianRuby.WriteLine("  o365_username = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "user" & """" & "]")
        CustodianRuby.WriteLine("  o365_password = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "password" & """" & "]")
        CustodianRuby.WriteLine("  o365_mailbox = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "mailbox" & """" & "]")
        CustodianRuby.WriteLine("  o365_domain = o365_mailbox.split('@').last")
        CustodianRuby.WriteLine("  o365_folder = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "folder" & """" & "]")
        CustodianRuby.WriteLine("  o365_root_mailbox = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "root_mailbox" & """" & "]")
        CustodianRuby.WriteLine("  o365_root_archive = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "root_archive" & """" & "]")
        CustodianRuby.WriteLine("  o365_mailbox_purges = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "mailbox_purges" & """" & "]")
        CustodianRuby.WriteLine("  o365_archive_purges = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "archive_purges" & """" & "]")
        CustodianRuby.WriteLine("  o365_mailbox_recoverable = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "mailbox_recoverable" & """" & "]")
        CustodianRuby.WriteLine("  o365_archive_recoverable = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "archive_recoverable" & """" & "]")
        CustodianRuby.WriteLine("  o365_root_public = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "root_public" & """" & "]")
        CustodianRuby.WriteLine("  o365_from_date = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "from_date" & """]")
        CustodianRuby.WriteLine("  o365_to_date = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "to_date" & """]")
        CustodianRuby.WriteLine("  o365_custodian_name = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "custodian_name" & """" & "]")
        CustodianRuby.WriteLine("  o365_migration = parsed[" & """" & "o365_mailbox" & """" & "][" & """" & "migration" & """]")
        CustodianRuby.WriteLine("caseFactory = $utilities.getCaseFactory()")
        CustodianRuby.WriteLine("case_settings = {")
        CustodianRuby.WriteLine("    :compound => false,")
        CustodianRuby.WriteLine("    :name => " & """#{o365_mailbox}" & """,")
        CustodianRuby.WriteLine("    :description => " & """" & """,")
        CustodianRuby.WriteLine("    :investigator => " & """EWS Puller" & """")
        CustodianRuby.WriteLine("}")
        CustodianRuby.WriteLine("$current_case = caseFactory.create(""" & sNuixCase.Replace("\", "\\") & """" & ", case_settings)")
        'CustodianRuby.WriteLine("        kinds_to_keep = {")
        'CustodianRuby.WriteLine("    " & """" & "email" & """" & "=> true,")
        'CustodianRuby.WriteLine("    " & """" & "calendar" & """" & "=> true,")
        'CustodianRuby.WriteLine("    " & """" & "contact" & """" & "=> true,")
        'CustodianRuby.WriteLine("    " & """" & "container" & """" & "=> true,")
        'CustodianRuby.WriteLine("    " & """" & "no-data" & """" & "=> true,")
        'CustodianRuby.WriteLine("    " & """" & "system" & """" & "=> true,")
        'CustodianRuby.WriteLine("    " & """" & "unrecognised" & """" & "=> true")
        'CustodianRuby.WriteLine("}")
        If (O365ExtractionSettings.NuixAppLocation.Contains("Nuix 6")) Then
            CustodianRuby.WriteLine("processor = $current_case.getProcessor")
        Else
            CustodianRuby.WriteLine("processor = $current_case.createProcessor")
        End If
        CustodianRuby.WriteLine("if o365_migration == """ & "lightspeed" & """")
        CustodianRuby.WriteLine("	processing_settings = {")
        CustodianRuby.WriteLine("		:traversalScope => " & """" & "full_traversal" & """" & ",")
        CustodianRuby.WriteLine("		:analysisLanguage => """ & "en" & """" & ",")
        CustodianRuby.WriteLine("		:identifyPhysicalFiles => " & "true" & ",")
        CustodianRuby.WriteLine("		:reuseEvidenceStores => " & "true" & ",")
        CustodianRuby.WriteLine("		:extractFromSlackSpace=> " & eMailArchiveMigrationManager.EnableMailboxSlackSpace.ToString.ToLower & ",")
        CustodianRuby.WriteLine("		:reportProcessingStatus => " & """" & "none" & """")
        CustodianRuby.WriteLine("	}")
        CustodianRuby.WriteLine("	processor.setProcessingSettings(processing_settings)")
        CustodianRuby.WriteLine("end")
        CustodianRuby.WriteLine("parallel_processing_settings = {")
        CustodianRuby.WriteLine("	:workerCount => " & eMailArchiveMigrationManager.O365NumberOfNuixWorkers & ",")
        CustodianRuby.WriteLine("	:workerMemory => " & eMailArchiveMigrationManager.O365MemoryPerWorker & ",")
        CustodianRuby.WriteLine("	:embedBroker => true,")
        CustodianRuby.WriteLine("	:brokerMemory => " & eMailArchiveMigrationManager.O365MemoryPerWorker & ",")
        CustodianRuby.WriteLine("	:workerTemp => " & """" & eMailArchiveMigrationManager.NuixWorkerTempDir.Replace("\", "\\") & """")
        CustodianRuby.WriteLine("}")
        CustodianRuby.WriteLine("processor.setParallelProcessingSettings(parallel_processing_settings)")
        CustodianRuby.WriteLine("")
        CustodianRuby.WriteLine("if o365_root_mailbox == " & """" & "mailbox" & """")
        CustodianRuby.WriteLine("	evidence_root_mailbox = " & """" & "#{o365_root_mailbox}" & """")
        CustodianRuby.WriteLine("	evidence_container = processor.newEvidenceContainer(evidence_root_mailbox)")
        CustodianRuby.WriteLine("	network_location_settings = {")
        CustodianRuby.WriteLine("		:type => " & """" & "exchange" & """" & ",")
        CustodianRuby.WriteLine("		:uri => " & """" & "#{o365_server}" & """" & ",")
        CustodianRuby.WriteLine("		:domain => " & """" & "#{o365_domain}" & """" & ",")
        CustodianRuby.WriteLine("		:username => " & """" & "#{o365_username}" & """" & ",")
        CustodianRuby.WriteLine("		:password => " & """" & "#{o365_password}" & """" & ",")
        CustodianRuby.WriteLine("		:mailboxRetrieval => " & """" & "#{o365_root_mailbox}" & """" & ",")
        CustodianRuby.WriteLine("		:mailbox => " & """" & "#{o365_mailbox}" & """" & ",")
        CustodianRuby.WriteLine("	}")
        CustodianRuby.WriteLine("	if !o365_from_date.empty?")
        CustodianRuby.WriteLine("		network_location_settings[:from] = o365_from_date")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	if !o365_to_date.empty?")
        CustodianRuby.WriteLine("		network_location_settings[:to] = o365_to_date")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	evidence_container.addNetworkLocation(network_location_settings)")
        CustodianRuby.WriteLine("	evidence_container.setEncoding(""" & "utf-8" & """)")
        CustodianRuby.WriteLine("	evidence_container.save")
        CustodianRuby.WriteLine("end")
        CustodianRuby.WriteLine("if o365_root_archive  == " & """" & "archive" & """")
        CustodianRuby.WriteLine("	evidence_root_archive = " & """" & "#{o365_root_archive}" & """")
        CustodianRuby.WriteLine("	evidence_container = processor.newEvidenceContainer(evidence_root_archive)")
        CustodianRuby.WriteLine("	network_location_settings = {")
        CustodianRuby.WriteLine("		:type => " & """" & "exchange" & """" & ",")
        CustodianRuby.WriteLine("		:uri => " & """" & "#{o365_server}" & """" & ",")
        CustodianRuby.WriteLine("		:domain => " & """" & "#{o365_domain}" & """" & ",")
        CustodianRuby.WriteLine("		:username => " & """" & "#{o365_username}" & """" & ",")
        CustodianRuby.WriteLine("		:password => " & """" & "#{o365_password}" & """" & ",")
        CustodianRuby.WriteLine("		:mailboxRetrieval => " & """" & "#{o365_root_archive}" & """" & ",")
        CustodianRuby.WriteLine("		:mailbox => " & """" & "#{o365_mailbox}" & """" & ",")
        CustodianRuby.WriteLine("	}")
        CustodianRuby.WriteLine("	if !o365_from_date.empty?")
        CustodianRuby.WriteLine("		network_location_settings[:from] = o365_from_date")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	if !o365_to_date.empty?")
        CustodianRuby.WriteLine("		network_location_settings[:to] = o365_to_date")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	evidence_container.addNetworkLocation(network_location_settings)")
        CustodianRuby.WriteLine("	evidence_container.setEncoding(""" & "utf-8" & """)")
        CustodianRuby.WriteLine("	evidence_container.save")
        CustodianRuby.WriteLine("end")
        CustodianRuby.WriteLine("if o365_mailbox_purges == " & """" & "purges" & """")
        CustodianRuby.WriteLine("	evidence_mailbox_purges = " & """" & "#{o365_mailbox_purges}" & """")
        CustodianRuby.WriteLine("	evidence_container = processor.newEvidenceContainer(evidence_mailbox_purges)")
        CustodianRuby.WriteLine("	network_location_settings = {")
        CustodianRuby.WriteLine("		:type => " & """" & "exchange" & """" & ",")
        CustodianRuby.WriteLine("		:uri => " & """" & "#{o365_server}" & """" & ",")
        CustodianRuby.WriteLine("		:domain => " & """" & "#{o365_domain}" & """" & ",")
        CustodianRuby.WriteLine("		:username => " & """" & "#{o365_username}" & """" & ",")
        CustodianRuby.WriteLine("		:password => " & """" & "#{o365_password}" & """" & ",")
        CustodianRuby.WriteLine("		:mailboxRetrieval => " & """" & "#{o365_mailbox_purges}" & """" & ",")
        CustodianRuby.WriteLine("		:mailbox => " & """" & "#{o365_mailbox}" & """" & ",")
        CustodianRuby.WriteLine("	}")
        CustodianRuby.WriteLine("	if !o365_from_date.empty?")
        CustodianRuby.WriteLine("		network_location_settings[:from] = o365_from_date")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	if !o365_to_date.empty?")
        CustodianRuby.WriteLine("		network_location_settings[:to] = o365_to_date")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	evidence_container.addNetworkLocation(network_location_settings)")
        CustodianRuby.WriteLine("	evidence_container.setEncoding(""" & "utf-8" & """)")
        CustodianRuby.WriteLine("	evidence_container.save")
        CustodianRuby.WriteLine("end")
        CustodianRuby.WriteLine("if o365_archive_purges == " & """" & "archive_purges" & """")
        CustodianRuby.WriteLine("	evidence_archive_purges = " & """" & "#{o365_archive_purges}" & """")
        CustodianRuby.WriteLine("	evidence_container = processor.newEvidenceContainer(evidence_archive_purges)")
        CustodianRuby.WriteLine("	network_location_settings = {")
        CustodianRuby.WriteLine("		:type => " & """" & "exchange" & """" & ",")
        CustodianRuby.WriteLine("		:uri => " & """" & "#{o365_server}" & """" & ",")
        CustodianRuby.WriteLine("		:domain => " & """" & "#{o365_domain}" & """" & ",")
        CustodianRuby.WriteLine("		:username => " & """" & "#{o365_username}" & """" & ",")
        CustodianRuby.WriteLine("		:password => " & """" & "#{o365_password}" & """" & ",")
        CustodianRuby.WriteLine("		:mailboxRetrieval => " & """" & "#{o365_archive_purges}" & """" & ",")
        CustodianRuby.WriteLine("		:mailbox => " & """" & "#{o365_mailbox}" & """" & ",")
        CustodianRuby.WriteLine("	}")
        CustodianRuby.WriteLine("	if !o365_from_date.empty?")
        CustodianRuby.WriteLine("		network_location_settings[:from] = o365_from_date")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	if !o365_to_date.empty?")
        CustodianRuby.WriteLine("		network_location_settings[:to] = o365_to_date")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	evidence_container.addNetworkLocation(network_location_settings)")
        CustodianRuby.WriteLine("	evidence_container.setEncoding(""" & "utf-8" & """)")
        CustodianRuby.WriteLine("	evidence_container.save")
        CustodianRuby.WriteLine("end")
        CustodianRuby.WriteLine("if o365_mailbox_recoverable == " & """" & "recoverable_items" & """")
        CustodianRuby.WriteLine("	evidence_mailbox_recoverable = " & """" & "#{o365_mailbox_recoverable}" & """")
        CustodianRuby.WriteLine("	evidence_container = processor.newEvidenceContainer(evidence_mailbox_recoverable)")
        CustodianRuby.WriteLine("	network_location_settings = {")
        CustodianRuby.WriteLine("		:type => " & """" & "exchange" & """" & ",")
        CustodianRuby.WriteLine("		:uri => " & """" & "#{o365_server}" & """" & ",")
        CustodianRuby.WriteLine("		:domain => " & """" & "#{o365_domain}" & """" & ",")
        CustodianRuby.WriteLine("		:username => " & """" & "#{o365_username}" & """" & ",")
        CustodianRuby.WriteLine("		:password => " & """" & "#{o365_password}" & """" & ",")
        CustodianRuby.WriteLine("		:mailboxRetrieval => " & """" & "#{o365_mailbox_recoverable}" & """" & ",")
        CustodianRuby.WriteLine("		:mailbox => " & """" & "#{o365_mailbox}" & """" & ",")
        CustodianRuby.WriteLine("	}")
        CustodianRuby.WriteLine("	if !o365_from_date.empty?")
        CustodianRuby.WriteLine("		network_location_settings[:from] = o365_from_date")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	if !o365_to_date.empty?")
        CustodianRuby.WriteLine("		network_location_settings[:to] = o365_to_date")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	evidence_container.addNetworkLocation(network_location_settings)")
        CustodianRuby.WriteLine("	evidence_container.setEncoding(""" & "utf-8" & """)")
        CustodianRuby.WriteLine("	evidence_container.save")
        CustodianRuby.WriteLine("end")
        CustodianRuby.WriteLine("if o365_archive_recoverable == " & """" & "archive_recoverable_items" & """")
        CustodianRuby.WriteLine("	evidence_archive_recoverable = " & """" & "#{o365_archive_recoverable}" & """")
        CustodianRuby.WriteLine("	evidence_container = processor.newEvidenceContainer(evidence_archive_recoverable)")
        CustodianRuby.WriteLine("	network_location_settings = {")
        CustodianRuby.WriteLine("		:type => " & """" & "exchange" & """" & ",")
        CustodianRuby.WriteLine("		:uri => " & """" & "#{o365_server}" & """" & ",")
        CustodianRuby.WriteLine("		:domain => " & """" & "#{o365_domain}" & """" & ",")
        CustodianRuby.WriteLine("		:username => " & """" & "#{o365_username}" & """" & ",")
        CustodianRuby.WriteLine("		:password => " & """" & "#{o365_password}" & """" & ",")
        CustodianRuby.WriteLine("		:mailboxRetrieval => " & """" & "#{o365_archive_recoverable}" & """" & ",")
        CustodianRuby.WriteLine("		:mailbox => " & """" & "#{o365_mailbox}" & """" & ",")
        CustodianRuby.WriteLine("	}")
        CustodianRuby.WriteLine("	if !o365_from_date.empty?")
        CustodianRuby.WriteLine("		network_location_settings[:from] = o365_from_date")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	if !o365_to_date.empty?")
        CustodianRuby.WriteLine("		network_location_settings[:to] = o365_to_date")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	evidence_container.addNetworkLocation(network_location_settings)")
        CustodianRuby.WriteLine("	evidence_container.setEncoding(""" & "utf-8" & """)")
        CustodianRuby.WriteLine("	evidence_container.save")
        CustodianRuby.WriteLine("end")
        CustodianRuby.WriteLine("if o365_root_public == " & """" & "public_folders" & """")
        CustodianRuby.WriteLine("	evidence_root_public = " & """" & "#{o365_root_public}" & """")
        CustodianRuby.WriteLine("	evidence_container = processor.newEvidenceContainer(evidence_root_public)")
        CustodianRuby.WriteLine("	network_location_settings = {")
        CustodianRuby.WriteLine("		:type => " & """" & "exchange" & """" & ",")
        CustodianRuby.WriteLine("		:uri => " & """" & "#{o365_server}" & """" & ",")
        CustodianRuby.WriteLine("		:domain => " & """" & "#{o365_domain}" & """" & ",")
        CustodianRuby.WriteLine("		:username => " & """" & "#{o365_username}" & """" & ",")
        CustodianRuby.WriteLine("		:password => " & """" & "#{o365_password}" & """" & ",")
        CustodianRuby.WriteLine("		:mailboxRetrieval => " & """" & "#{o365_root_public}" & """" & ",")
        CustodianRuby.WriteLine("		:mailbox => " & """" & "#{o365_mailbox}" & """" & ",")
        CustodianRuby.WriteLine("	}")
        CustodianRuby.WriteLine("	if !o365_from_date.empty?")
        CustodianRuby.WriteLine("		network_location_settings[:from] = o365_from_date")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	if !o365_to_date.empty?")
        CustodianRuby.WriteLine("		network_location_settings[:to] = o365_to_date")
        CustodianRuby.WriteLine("	end")
        CustodianRuby.WriteLine("	evidence_container.addNetworkLocation(network_location_settings)")
        CustodianRuby.WriteLine("	evidence_container.setEncoding(""" & "utf-8" & """)")
        CustodianRuby.WriteLine("	evidence_container.save")
        CustodianRuby.WriteLine("end")
        CustodianRuby.WriteLine("start_time = Time.now")
        CustodianRuby.WriteLine("")
        CustodianRuby.WriteLine("puts")
        CustodianRuby.WriteLine("puts " & """" & "Office 365 Extraction for #{o365_mailbox} has started..." & """")
        CustodianRuby.WriteLine("puts")
        CustodianRuby.WriteLine("printf """ & "%-40s %-25s """ & "," & """" & "Timestamp" & """" & ", " & """" & "Processed Items" & """")
        CustodianRuby.WriteLine("puts")
        CustodianRuby.WriteLine("processor.when_item_processed do |event|")
        CustodianRuby.WriteLine("  if callback_count % CALLBACK_FREQUENCY == 0")
        CustodianRuby.WriteLine("			updated_callback = [callback_count]")
        CustodianRuby.WriteLine("			begin")
        CustodianRuby.WriteLine("			    db.update(""" & "UPDATE ewsExtractionStats SET ProcessedItems = ? WHERE CustodianName = '#{o365_custodian_name}'" & """" & ",updated_callback)")
        CustodianRuby.WriteLine("			rescue")
        CustodianRuby.WriteLine("				when_item_processed")
        CustodianRuby.WriteLine("			end")
        CustodianRuby.WriteLine("	printf " & """" & "\r%-40s %-25s" & """" & ", Time.now, callback_count")
        CustodianRuby.WriteLine("  end")
        CustodianRuby.WriteLine("  callback_count += 1")
        CustodianRuby.WriteLine("end")
        CustodianRuby.WriteLine("")
        CustodianRuby.WriteLine("processor.process")
        'CustodianRuby.WriteLine("updated_callback = [callback_count]")
        'CustodianRuby.WriteLine("db.update(""" & "UPDATE ewsExtractionStats SET ProcessedItems = ? WHERE CustodianName = '#{o365_custodian_name}'" & """" & ",updated_callback)")
        'CustodianRuby.WriteLine("")
        CustodianRuby.WriteLine("puts")
        'CustodianRuby.WriteLine("items = $current_case.search('flag:top_level AND kind:(email OR calendar OR contact)');")

        CustodianRuby.WriteLine("#search_file_size = $current_case.getStatistics.getAuditSize('flag:top_level AND kind:(email OR calendar OR contact)')")
        CustodianRuby.WriteLine("#puts " & """" & "Callback Count #{callback_count}" & """")
        CustodianRuby.WriteLine("#puts " & """" & "Search File size #{search_file_size}" & """")
        CustodianRuby.WriteLine("#updated_callback = [callback_count, search_file_size]")
        CustodianRuby.WriteLine("#   begin")
        CustodianRuby.WriteLine("#       db.update(" & """" & "UPDATE ewsExtractionStats SET ProcessedItems = ?, ExtractionSize = ? WHERE CustodianName = '#{o365_custodian_name}'" & """" & ", updated_callback)")
        CustodianRuby.WriteLine("#   rescue")
        CustodianRuby.WriteLine("#       sleep 2")
        CustodianRuby.WriteLine("#       db.update(" & """" & "UPDATE ewsExtractionStats SET ProcessedItems = ?, ExtractionSize = ? WHERE CustodianName = '#{o365_custodian_name}'" & """" & ", updated_callback)")
        CustodianRuby.WriteLine("#   end")
        'CustodianRuby.WriteLine("DBFinalUpdate(updated_callback)")
        '        CustodianRuby.WriteLine("digestSize = 0")
        '        CustodianRuby.WriteLine("items.each do |i|")
        '        CustodianRuby.WriteLine("     digestSize += i.getDigests().getInputSize();")
        '        CustodianRuby.WriteLine("end")
        '        CustodianRuby.WriteLine("updated_callback = [digestSize]")
        CustodianRuby.WriteLine("end_time = Time.now")
        CustodianRuby.WriteLine("updated_callback = [" & """" & "Completed" & """" & ",end_time]")
        CustodianRuby.WriteLine("db.update(" & """" & "UPDATE ewsExtractionStats SET ExtractionStatus = ?, ExtractionEndTime = ? WHERE CustodianName = '#{o365_custodian_name}'" & """" & ", updated_callback)")
        CustodianRuby.WriteLine("$current_case.close")
        CustodianRuby.WriteLine("return")
        CustodianRuby.Close()

        blnBuildO365ExtractionRubyFiles = True
    End Function


    Private Function blnBuildEWSExtractJSonFile(ByVal sCustodianName As String, ByVal sCustodianPath As String, dblPSTSourceSize As Double, ByVal sDestinationFolder As String, ByVal sDestinationRoot As String, ByVal sDestinationSMTPAddress As String, ByVal sJSonFileName As String, sPSTLocation As String, ByVal sWorkerMemory As String, ByVal sWorkerCount As String, ByVal sCaseDir As String, ByVal sFromDate As String, ByVal sToDate As String) As Boolean
        blnBuildEWSExtractJSonFile = False

        Dim ArchiveExtractCustodianJSon As StreamWriter


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

        ArchiveExtractCustodianJSon = New StreamWriter(sJSonFileName)

        ArchiveExtractCustodianJSon.WriteLine("{")
        ArchiveExtractCustodianJSon.WriteLine("  " & """" & "o365_mailbox" & """: {")
        ArchiveExtractCustodianJSon.WriteLine("    " & """server""" & ": " & """" & eMailArchiveMigrationManager.O365ExchangeServer & """" & ",")
        ArchiveExtractCustodianJSon.WriteLine("    " & """user""" & ": " & """" & eMailArchiveMigrationManager.O365AdminUserName & """" & ",")
        ArchiveExtractCustodianJSon.WriteLine("    " & """password""" & ": " & """" & eMailArchiveMigrationManager.O365AdminInfo & """" & ",")
        ArchiveExtractCustodianJSon.WriteLine("    " & """mailbox""" & ": " & """" & sCustodianName & """" & ",")
        ArchiveExtractCustodianJSon.WriteLine("    " & """folder""" & ": " & """" & sDestinationFolder & """" & ",")

        Select Case sDestinationRoot
            Case "root_mailbox"
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_mailbox""" & ": " & """mailbox""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_archive""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_public""" & ": " & """" & vbNullString & """" & ",")

            Case "root_archive"
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_mailbox""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_archive""" & ": " & """archive""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_public""" & ": " & """" & vbNullString & """" & ",")

            Case "mailbox_purges"
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_mailbox""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_archive""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_purges""" & ": " & """purges""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_public""" & ": " & """" & vbNullString & """" & ",")
            Case "archive_purges"
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_mailbox""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_archive""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_purges""" & ": " & """archive_purges""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_public""" & ": " & """" & vbNullString & """" & ",")

            Case "mailbox_recoverable"
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_mailbox""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_archive""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_recoverable""" & ": " & """recoverable_items""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_public""" & ": " & """" & vbNullString & """" & ",")

            Case "archive_recoverable"
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_mailbox""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_archive""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_recoverable""" & ": " & """archive_recoverable_items""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_public""" & ": " & """" & vbNullString & """" & ",")
            Case "public folders"
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_mailbox""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_archive""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_public""" & ": " & """public_folders""" & ",")
            Case "mailbox/archive"
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_mailbox""" & ": " & """mailbox""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_archive""" & ": " & """archive""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_public""" & ": " & """" & vbNullString & """" & ",")
            Case "mailbox/mailbox_recoverable"
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_mailbox""" & ": " & """mailbox""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_archive""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_recoverable""" & ": " & """recoverable_items""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_public""" & ": " & """" & vbNullString & """" & ",")
            Case "archive/archive_recoverable"
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_mailbox""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_archive""" & ": " & """archive""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_recoverable""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_recoverable""" & ": " & """archive_recoverable_items""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_public""" & ": " & """" & vbNullString & """" & ",")
            Case "mailbox/mailbox_recoverable/archive/archive_recoverable"
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_mailbox""" & ": " & """mailbox""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_archive""" & ": " & """archive""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_purges""" & ": " & """" & vbNullString & """" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "mailbox_recoverable""" & ": " & """recoverable_items""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "archive_recoverable""" & ": " & """archive_recoverable_items""" & ",")
                ArchiveExtractCustodianJSon.WriteLine("    " & """" & "root_public""" & ": " & """" & vbNullString & """" & ",")
        End Select
        If sFromDate <> "NA" Then
            ArchiveExtractCustodianJSon.WriteLine("    " & """from_date""" & ": " & """" & sFromDate & """" & ",")
            ArchiveExtractCustodianJSon.WriteLine("    " & """to_date""" & ": " & """" & sToDate & """" & ",")
        End If
        ArchiveExtractCustodianJSon.WriteLine("	" & """custodian_name""" & ": " & """" & sCustodianName & """" & ",")

        ArchiveExtractCustodianJSon.WriteLine("	" & """source_data""" & ": """ & sPSTLocation & """,")
        ArchiveExtractCustodianJSon.WriteLine("	" & """source_size""" & ": " & """" & dblPSTSourceSize & """" & ",")
        ArchiveExtractCustodianJSon.WriteLine("	" & """migration""" & ": " & """lightspeed""")
        ArchiveExtractCustodianJSon.WriteLine("  }")
        ArchiveExtractCustodianJSon.WriteLine("}")

        ArchiveExtractCustodianJSon.Close()

        blnBuildEWSExtractJSonFile = True

    End Function


    Public Sub SQLiteUpdates()

        Dim lstExtractionSize As List(Of String)
        Dim lstExtractedItems As List(Of String)
        Dim lstSMTPName As List(Of String)
        Dim lstNotStartedSMTPAddress As List(Of String)
        Dim lstEndTime As List(Of String)
        Dim bStatus As Boolean
        Dim bNoMoreJobs As Boolean
        Dim sProcessID As String
        Dim sErrorType As String
        Dim sNuixLogFileDir As String
        Dim sNuixLogFile As String
        Dim lstCustodianName As List(Of String)
        Dim lstItemsProcessed As List(Of String)
        Dim sCaseDirectory As String
        Dim dblTotalFolderSize As Double
        Dim iNumberOfNuixInstancesRunning As Integer
        Dim dStartDate As Date
        Dim sSMTPAddress As String
        Dim NuixConsoleProcess As Process
        Dim NuixConsoleProcessStartInfo As ProcessStartInfo
        Dim sCustodianName As String
        Dim sExtractionSize As String
        Dim dbService As New DatabaseService
        Dim common As New Common

        bNoMoreJobs = False
        Thread.Sleep(10000)
        lstExtractionSize = New List(Of String)
        lstExtractedItems = New List(Of String)
        lstSMTPName = New List(Of String)
        lstNotStartedSMTPAddress = New List(Of String)
        lstCustodianName = New List(Of String)
        lstItemsProcessed = New List(Of String)
        lstEndTime = New List(Of String)

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"


        Do While bNoMoreJobs = False

            bStatus = blnGetUpdatedSMTPInfo(lstSMTPName, lstExtractionSize, lstExtractedItems)
            If lstSMTPName.Count > 0 Then
                For Each SMPTAddress In lstSMTPName
                    For Each row In grdCustodainSMTPInfo.Rows

                        If row.cells("SMTPAddress").value = SMPTAddress.ToString Then
                            Try
                                row.cells("ExtractedItems").value = lstExtractedItems(lstSMTPName.IndexOf(SMPTAddress.ToString)).Replace(".00", "")
                                row.cells("ExtractedSize").value = lstExtractionSize(lstSMTPName.IndexOf(SMPTAddress.ToString)).Replace(".00", "")
                                sProcessID = row.cells("ProcessID").value
                                bStatus = blnCheckIfProcessIsRunning(sProcessID)
                                If bStatus = False Then
                                    row.cells("ProcessID").value = ""
                                    sNuixLogFileDir = row.cells("LogDirectory").value
                                    bStatus = blnCheckNuixLogForErrors(sNuixLogFileDir, sNuixLogFile, sErrorType)
                                    If bStatus = True Then
                                        If sErrorType <> vbNullString Then
                                            Select Case sErrorType
                                                Case "Error Running Script"
                                                    row.cells("ExtractionStatus").value = "Failed (review logs)"
                                                    bStatus = dbService.UpdateExtractionDBInfo(psSQLiteLocation, row.cells("SMTPAddress").value, "ExtractionStatus", "Failed (review logs)")
                                                    row.cells("StatusImage").value = My.Resources.red_stop_small
                                                    row.DefaultCellStyle.Forecolor = Color.Red
                                                Case "No License"
                                                    row.cells("ExtractionStatus").value = "Failed (No License Available)"
                                                    row.cells("StatusImage").value = My.Resources.red_stop_small
                                                    bStatus = dbService.UpdateExtractionDBInfo(psSQLiteLocation, row.cells("SMTPAddress").value, "ExtractionStatus", "Failed (No License Available)")
                                                    row.DefaultCellStyle.Forecolor = Color.Red
                                                Case "Cound not open Case"
                                                    row.cells("ExtractionStatus").value = "Failed (review logs)"
                                                    bStatus = dbService.UpdateExtractionDBInfo(psSQLiteLocation, row.cells("SMTPAddress").value, "ExtractionStatus", "Failed (review logs)")
                                                    row.cells("StatusImage").value = My.Resources.red_stop_small
                                                    row.DefaultCellStyle.Forecolor = Color.Red
                                                Case "Invalid Product Class"
                                                    row.cells("ExtractionStatus").value = "Failed (review logs)"
                                                    bStatus = dbService.UpdateExtractionDBInfo(psSQLiteLocation, row.cells("SMTPAddress").value, "ExtractionStatus", "Failed (review logs)")
                                                    row.cells("StatusImage").value = My.Resources.red_stop_small
                                                    row.DefaultCellStyle.Forecolor = Color.Red
                                            End Select
                                        End If
                                    Else
                                        row.cells("StatusImage").value = My.Resources.Green_check_small
                                        row.DefaultCellStyle.Forecolor = Color.Green
                                        row.cells("ExtractionStatus").value = "Completed"
                                        bStatus = dbService.UpdateExtractionDBInfo(psSQLiteLocation, row.cells("SMTPAddress").value, "ExtractionStatus", "Completed")
                                    End If
                                End If

                            Catch ex As Exception
                                MessageBox.Show("SQLiteUpdates - IN Progress - " & ex.ToString, "SQLiteUpdates - In Progress")
                                common.Logger(psIngestionLogFile, ex.ToString)

                            End Try
                        End If
                    Next
                Next
            Else
                pbNoMoreJobs = True
            End If
            lstExtractionSize.Clear()
            lstExtractedItems.Clear()
            lstSMTPName.Clear()
            lstCustodianName.Clear()
            lstItemsProcessed.Clear()
            lstExtractionSize.Clear()
            lstEndTime.Clear()

            bStatus = blnGetCompletedExtractionCustodianInfo(lstCustodianName, lstItemsProcessed, lstEndTime, lstExtractionSize)
            If lstCustodianName.Count > 0 Then
                For Each custodian In lstCustodianName
                    For Each row In grdCustodainSMTPInfo.Rows
                        If row.cells("SMTPAddress").value = custodian.ToString Then
                            Try
                                row.cells("ExtractionStatus").value = "Completed"
                                row.cells("ExtractedItems").value = lstItemsProcessed(lstCustodianName.IndexOf(custodian.ToString))
                                row.cells("ExtractionEndTime").value = lstEndTime(lstCustodianName.IndexOf(custodian.ToString))
                                sCaseDirectory = row.Cells("CaseDirectory").value
                                bStatus = blnCheckIfProcessIsRunning(row.cells("ProcessID").value)
                                If bStatus = False Then
                                    bStatus = dbService.UpdateExtractionDBInfo(psSQLiteLocation, custodian.ToString, "ProcessID", "")
                                    row.cells("ProcessID").value = ""
                                End If
                                If File.Exists(sCaseDirectory & "\summary-report.txt") Then
                                    row.cells("SummaryReportLocation").value = sCaseDirectory & "\summary-report.txt"
                                    row.cells("ProcessID").value = ""
                                    bStatus = dbService.UpdateExtractionDBInfo(psSQLiteLocation, custodian.ToString, "SummaryReportLocation", sCaseDirectory & "\summary-report.txt")
                                End If
                                bStatus = dbService.GetUpdatedEWSExtractInfo(sSQLiteDatabaseFullName, custodian.ToString, "ExtractionSize", sExtractionSize)
                                If ((sExtractionSize = vbNullString) Or (sExtractionSize = "0")) Then
                                    bStatus = blnComputeFolderSize(row.Cells("OutputDirectory").value, dblTotalFolderSize)
                                    row.cells("ExtractedSize").value = dblTotalFolderSize.ToString("N0")
                                    bStatus = dbService.UpdateExtractionDBInfo(psSQLiteLocation, custodian.ToString, "ExtractionSize", dblTotalFolderSize)
                                Else
                                    row.cells("ExtractedSize").value = CDbl(sExtractionSize).ToString("N0")
                                End If
                                row.defaultcellstyle.Forecolor = Color.Green
                                row.cells("StatusImage").value = My.Resources.Green_check_small

                            Catch ex As Exception
                                MessageBox.Show("SQLiteUpdates - Completed - " & ex.ToString, "SQLiteUpdates - Completed")
                                common.Logger(psIngestionLogFile, ex.ToString)
                            End Try
                        End If
                    Next
                Next
            End If
            iNumberOfNuixInstancesRunning = 0

            bStatus = blnCheckIfNuixIsRunning("nuix_console", iNumberOfNuixInstancesRunning)
            bStatus = blnCheckIfNuixIsRunning("nuix_app", iNumberOfNuixInstancesRunning)
            If iNumberOfNuixInstancesRunning < piNumberOfNuixInstancesRequested Then
                sCustodianName = vbNullString
                bStatus = blnGetNextExtractionNotStartedCustodianInfo(sCustodianName)
                If sCustodianName <> vbNullString Then
                    For Each row In grdCustodainSMTPInfo.Rows
                        If row.cells("SMTPAddress").value = sCustodianName Then
                            Try
                                dStartDate = Now
                                sSMTPAddress = row.Cells("SMTPAddress").Value

                                NuixConsoleProcessStartInfo = New ProcessStartInfo(row.Cells("ProcessingFilesDirectory").Value & "\" & sSMTPAddress & ".bat")
                                NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                'NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Minimized

                                NuixConsoleProcess = System.Diagnostics.Process.Start(NuixConsoleProcessStartInfo)

                                row.Cells("StatusImage").Value = My.Resources.inprogress_medium
                                row.Cells("ExtractionStatus").Value = "In Progress"
                                row.Cells("ProcessID").Value = NuixConsoleProcess.Id
                                row.Cells("ExtractionStartTime").Value = dStartDate
                                bStatus = dbService.UpdateExtractionDBInfo(psSQLiteLocation, sSMTPAddress, "ExtractionStatus", "In Progress")
                                bStatus = dbService.UpdateExtractionDBInfo(psSQLiteLocation, sSMTPAddress, "ExtractionStartTime", dStartDate)

                            Catch ex As Exception
                                MessageBox.Show("SQLiteUpdates - Next Instance - " & ex.ToString, "SQLiteUpdates - Next Instance")
                                common.Logger(psIngestionLogFile, ex.ToString)

                            End Try
                        End If
                    Next
                End If
            End If

            Thread.Sleep(10000)
            bNoMoreJobs = pbNoMoreJobs
        Loop

        'bStatus = blnRefreshGridFromDB(grdCustodainSMTPInfo, lstNotStartedSMTPAddress, lstInProgresSMTPAddress)
        common.Logger(psExtractionLogFile, "All Currently select Processing Jobs have completed.  If necessary select more custodians for Processing.")
        MessageBox.Show("All Currently select Processing Jobs have completed.  If necessary select more custodians for Processing.", "All Processing Jobs Completed", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification, False)
        btnLoadPreviousBatches.Enabled = True
    End Sub


    Public Function blnGetCompletedExtractionCustodianInfo(ByRef lstCustodianName As List(Of String), ByRef lstItemsProcessed As List(Of String), ByRef lstEndTime As List(Of String), ByRef lstExtractionSize As List(Of String)) As Boolean
        blnGetCompletedExtractionCustodianInfo = False

        Dim Connection As SQLiteConnection
        Dim SQLCommand As SQLiteCommand
        Dim sCustodianQuery As String
        Dim dataReader As SQLiteDataReader

        Connection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
        Connection.Open()

        sCustodianQuery = "SELECT CustodianName, ProcessedItems, ExtractionEndTime FROM ewsExtractionStats WHERE ExtractionStatus = 'Completed'"

        SQLCommand = New SQLiteCommand(sCustodianQuery, Connection)
        dataReader = SQLCommand.ExecuteReader
        If dataReader.HasRows Then
            While dataReader.Read
                lstCustodianName.Add(dataReader.GetValue(0))
                lstItemsProcessed.Add(dataReader.GetInt64(1).ToString)
                lstEndTime.Add(dataReader.GetValue(2))
                If dataReader.IsDBNull(3) Then
                    lstExtractionSize.Add("0")
                Else
                    lstExtractionSize.Add(dataReader.GetInt64(3).ToString)
                End If
            End While
        End If
        Connection.Close()
        blnGetCompletedExtractionCustodianInfo = True
    End Function

    Public Function blnGetNextExtractionNotStartedCustodianInfo(ByRef sCustodianName As String) As Boolean
        blnGetNextExtractionNotStartedCustodianInfo = False

        Dim Connection As SQLiteConnection
        Dim SQLCommand As SQLiteCommand
        Dim sCustodianQuery As String
        Dim dataReader As SQLiteDataReader

        Connection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
        Connection.Open()

        sCustodianQuery = "SELECT CustodianName from ewsExtractionStats WHERE ExtractionStatus = 'Not Started'"

        SQLCommand = New SQLiteCommand(sCustodianQuery, Connection)
        dataReader = SQLCommand.ExecuteReader
        If dataReader.HasRows Then
            While dataReader.Read
                sCustodianName = dataReader.GetValue(0)
            End While
        End If
        Connection.Close()
        blnGetNextExtractionNotStartedCustodianInfo = True
    End Function
    Private Function blnRefreshGridFromDB(ByVal grdCustodainSMTPInfo As DataGridView, ByRef lstNotStartedCustodians As List(Of String), ByRef lstInProgressCustodians As List(Of String)) As Boolean
        Dim mSQL As String
        Dim dt As DataTable
        Dim ds As DataSet
        Dim dataReader As SQLiteDataReader
        Dim sqlCommand As SQLiteCommand
        Dim sqlConnection As SQLiteConnection
        Dim sSMTPAddress As String
        Dim sProgressStatus As String
        Dim common As New Common

        blnRefreshGridFromDB = False

        dt = Nothing
        ds = New DataSet
        common.Logger(psExtractionLogFile, "Refreshing Data grid from database")
        sqlConnection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")

        For Each row In grdCustodainSMTPInfo.Rows
            sSMTPAddress = row.cells("SMTPAddress").value
            mSQL = "select CustodianName, ExtractionRoot, GroupID, CaseDirectory, OutputDirectory, ExtractionStatus, ProcessID, ExtractionStartTime, ExtractionEndTime, ProcessedItems, ExtractionSize, SummaryReportLocation from ewsExtractionStats where CustodianName = '" & sSMTPAddress & "'"
            sqlCommand = New SQLiteCommand(mSQL, sqlConnection)
            sqlConnection.Open()

            dataReader = sqlCommand.ExecuteReader

            While dataReader.Read

                sProgressStatus = vbNullString

                row.cells("SMTPAddress").value = dataReader.GetValue(0)
                row.cells("ExtractionRoot").value = dataReader.GetValue(1)
                row.cells("GroupID").value = dataReader.GetValue(2)
                row.cells("CaseDirectory").value = dataReader.GetValue(3)
                row.cells("OutputDirectory").value = dataReader.GetValue(4)
                row.cells("ExtractionStatus").value = dataReader.GetValue(5)
                If dataReader.GetValue(5) = "In Progress" Then
                    row.cells("StatusImage").value = My.Resources.inprogress_medium
                ElseIf (dataReader.GetValue(5).value = "Complete") Then
                    row.cells("StatusImage").value = My.Resources.Green_check_small
                ElseIf (dataReader.GetValue(5).value = "Not Started") Then
                    row.cells("StatusImage").value = My.Resources.waitingtostart1
                Else
                    row.cells("StatusImage").value = My.Resources.not_selected_small
                End If
                row.cells("ProcessID").value = dataReader.GetValue(6)
                row.cells("ExtractionStartTime").value = dataReader.GetValue(7)
                row.cells("ExtractionEndTime").value = sProgressStatus
                row.cells("ExtractedItems").value = FormatNumber(dataReader.GetValue(11), -1, , , TriState.True)
                row.cells("ExtractionSize").value = FormatNumber(dataReader.GetValue(12), -1, , , TriState.True)
                row.cells("SummaryReportLocation").value = dataReader.GetValue(13)

            End While
            sqlConnection.Close()

        Next
        blnRefreshGridFromDB = True
    End Function

    Public Function blnGetUpdatedSMTPInfo(ByRef lstSMTPName As List(Of String), ByRef lstExtractionSize As List(Of String), ByRef lstExtractedItems As List(Of String)) As Boolean
        blnGetUpdatedSMTPInfo = False

        Dim Connection As SQLiteConnection
        Dim SQLCommand As SQLiteCommand
        Dim sCustodianQuery As String
        Dim dataReader As SQLiteDataReader
        Try

            Connection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
            Connection.Open()

            sCustodianQuery = "SELECT CustodianName, ExtractionSize, ProcessedItems FROM ewsExtractionStats WHERE ExtractionStatus in ('In Progress', 'Not Started') "

            SQLCommand = New SQLiteCommand(sCustodianQuery, Connection)
            dataReader = SQLCommand.ExecuteReader
            If dataReader.HasRows Then
                While dataReader.Read
                    lstSMTPName.Add(dataReader.GetValue(0))
                    lstExtractionSize.Add(dataReader.GetInt64(1).ToString)
                    lstExtractedItems.Add(dataReader.GetValue(2))
                End While
            End If
            Connection.Close()

        Catch ex As Exception
            MessageBox.Show("Error in: blnGetUpdatedSMTPInfo - " & ex.ToString)

        End Try
        blnGetUpdatedSMTPInfo = True
    End Function



    Private Function blnLoadGridFromDB(ByVal grdCustodainSMTPInfo As DataGridView) As Boolean
        Dim mSQL As String
        Dim dt As DataTable
        Dim ds As DataSet
        Dim dataReader As SQLiteDataReader
        Dim sqlCommand As SQLiteCommand
        Dim sqlConnection As SQLiteConnection
        Dim BatchRow As DataGridViewRow
        Dim sCustodianName As String
        Dim iCustodianNameCol As Integer
        Dim sExtractionRoot As String
        Dim iExtractionRootCol As Integer
        Dim sExtractionStatus As String
        Dim iExtractionStatusCol As Integer
        Dim sGroupID As String
        Dim iGroupIDCol As Integer
        Dim dblProcessedItems As Double
        Dim iProcessedItemsCol As Integer
        Dim dblExtractionSize As Double
        Dim iExtractionSizeCol As Integer
        Dim sCaseDirectory As String
        Dim iCaseDirectoryCol As Integer
        Dim iLogDirectoryCol As Integer
        Dim sLogDirectory As String
        Dim sOutputDirectory As String
        Dim iOutputDirectoryCol As Integer
        Dim sSummaryReportLocation As String
        Dim iSummaryReportLocationCol As Integer
        Dim sExtractionStartTime As String
        Dim iExtractionStartTimeCol As Integer
        Dim sExtractionEndTime As String
        Dim iExtractionEndTimeCol As Integer
        Dim iProcessingFilesDirCol As Integer
        Dim sProcessingFilesDirectory As String
        Dim iOutputFormatCol As Integer
        Dim sOutputFormat As String
        Dim iRowIndex As Integer

        blnLoadGridFromDB = False

        Try

            dt = Nothing
            ds = New DataSet
            sqlConnection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
            mSQL = "select CustodianName, ExtractionStatus, ExtractionRoot, GroupID, ProcessedItems, ExtractionSize, OutputFormat, CaseDirectory, OutputDirectory, LogDirectory, ProcessingFilesDirectory, SummaryReportLocation, ExtractionStartTime, ExtractionEndTime from ewsExtractionStats"

            sqlCommand = New SQLiteCommand(mSQL, sqlConnection)
            sqlConnection.Open()

            dataReader = sqlCommand.ExecuteReader

            While dataReader.Read
                iCustodianNameCol = dataReader.GetOrdinal("CustodianName")
                If dataReader.IsDBNull(iCustodianNameCol) Then
                    sCustodianName = vbNullString
                Else
                    Try
                        sCustodianName = dataReader.GetString(iCustodianNameCol)
                    Catch ex As Exception
                        sCustodianName = vbNullString
                    End Try
                End If

                iExtractionStatusCol = dataReader.GetOrdinal("ExtractionStatus")
                If dataReader.IsDBNull(iExtractionStatusCol) Then
                    sExtractionStatus = vbNullString
                Else
                    Try
                        sExtractionStatus = dataReader.GetString(iExtractionStatusCol)

                    Catch ex As Exception
                        sExtractionStatus = vbNullString
                    End Try
                End If

                iExtractionRootCol = dataReader.GetOrdinal("ExtractionRoot")
                If dataReader.IsDBNull(iExtractionRootCol) Then
                    sExtractionRoot = vbNullString
                Else
                    Try
                        sExtractionRoot = dataReader.GetString(iExtractionRootCol)

                    Catch ex As Exception
                        sExtractionRoot = vbNullString
                    End Try
                End If

                iGroupIDCol = dataReader.GetOrdinal("GroupID")
                If dataReader.IsDBNull(iGroupIDCol) Then
                    sGroupID = vbNullString
                Else
                    Try
                        sGroupID = dataReader.GetString(iGroupIDCol)

                    Catch ex As Exception
                        sGroupID = vbNullString
                    End Try
                End If

                iOutputFormatCol = dataReader.GetOrdinal("OutputFormat")
                If dataReader.IsDBNull(iOutputFormatCol) Then
                    sOutputFormat = vbNullString
                Else
                    Try
                        sOutputFormat = dataReader.GetString(iOutputFormatCol)

                    Catch ex As Exception
                        sOutputFormat = vbNullString
                    End Try
                End If


                iProcessedItemsCol = dataReader.GetOrdinal("ProcessedItems")
                If IsDBNull(iProcessedItemsCol) Then
                    dblProcessedItems = 0
                Else
                    Try
                        dblProcessedItems = dataReader.GetInt64(iProcessedItemsCol)

                    Catch ex As Exception
                        dblProcessedItems = 0
                    End Try
                End If

                iExtractionSizeCol = dataReader.GetOrdinal("ExtractionSize")
                If IsDBNull(iExtractionSizeCol) Then
                    dblExtractionSize = 0
                Else
                    Try
                        dblExtractionSize = dataReader.GetInt64(iExtractionSizeCol)

                    Catch ex As Exception
                        dblExtractionSize = 0

                    End Try
                End If

                iCaseDirectoryCol = dataReader.GetOrdinal("CaseDirectory")
                If dataReader.IsDBNull(iCaseDirectoryCol) Then
                    sCaseDirectory = vbNullString
                Else
                    Try
                        sCaseDirectory = dataReader.GetString(iCaseDirectoryCol)

                    Catch ex As Exception
                        sCaseDirectory = vbNullString
                    End Try
                End If

                iOutputDirectoryCol = dataReader.GetOrdinal("OutputDirectory")
                If dataReader.IsDBNull(iOutputDirectoryCol) Then
                    sOutputDirectory = vbNullString
                Else
                    Try
                        sOutputDirectory = dataReader.GetString(iOutputDirectoryCol)

                    Catch ex As Exception
                        sOutputDirectory = vbNullString
                    End Try
                End If

                iLogDirectoryCol = dataReader.GetOrdinal("LogDirectory")
                If dataReader.IsDBNull(iLogDirectoryCol) Then
                    sLogDirectory = vbNullString
                Else
                    Try
                        sLogDirectory = dataReader.GetString(iLogDirectoryCol)

                    Catch ex As Exception
                        sLogDirectory = vbNullString
                    End Try
                End If

                iProcessingFilesDirCol = dataReader.GetOrdinal("ProcessingFilesDirectory")
                If dataReader.IsDBNull(iProcessingFilesDirCol) Then
                    sProcessingFilesDirectory = vbNullString
                Else
                    Try
                        sProcessingFilesDirectory = dataReader.GetString(iProcessingFilesDirCol)

                    Catch ex As Exception
                        sProcessingFilesDirectory = vbNullString
                    End Try
                End If

                iSummaryReportLocationCol = dataReader.GetOrdinal("SummaryReportLocation")
                If dataReader.IsDBNull(iSummaryReportLocationCol) Then
                    sSummaryReportLocation = vbNullString
                Else
                    Try
                        sSummaryReportLocation = dataReader.GetString(iSummaryReportLocationCol)

                    Catch ex As Exception
                        sSummaryReportLocation = vbNullString

                    End Try
                End If

                iExtractionStartTimeCol = dataReader.GetOrdinal("ExtractionStartTime")
                If dataReader.IsDBNull(iExtractionStartTimeCol) Then
                    sExtractionStartTime = vbNullString
                Else
                    Try
                        sExtractionStartTime = dataReader.GetString(iExtractionStartTimeCol)

                    Catch ex As Exception
                        sExtractionStartTime = vbNullString

                    End Try
                End If

                iExtractionEndTimeCol = dataReader.GetOrdinal("ExtractionEndTime")
                If dataReader.IsDBNull(iExtractionEndTimeCol) Then
                    sExtractionEndTime = vbNullString
                Else
                    Try
                        sExtractionEndTime = dataReader.GetString(iExtractionEndTimeCol)

                    Catch ex As Exception
                        sExtractionEndTime = vbNullString

                    End Try
                End If

                iRowIndex = grdCustodainSMTPInfo.Rows.Add()

                BatchRow = grdCustodainSMTPInfo.Rows(iRowIndex)

                With BatchRow
                    BatchRow.Cells("SelectCustodian").Value = False
                    BatchRow.Cells("SMTPAddress").Value = sCustodianName
                    BatchRow.Cells("ExtractionRoot").Value = sExtractionRoot
                    BatchRow.Cells("OutputFormat").Value = sOutputFormat
                    BatchRow.Cells("GroupID").Value = sGroupID
                    BatchRow.Cells("ExtractedItems").Value = dblProcessedItems.ToString("N0")
                    BatchRow.Cells("ExtractedSize").Value = dblExtractionSize.ToString("N0")
                    BatchRow.Cells("CaseDirectory").Value = sCaseDirectory
                    BatchRow.Cells("OutputDirectory").Value = sOutputDirectory
                    BatchRow.Cells("LogDirectory").Value = sLogDirectory
                    BatchRow.Cells("ExtractionStatus").Value = sExtractionStatus
                    BatchRow.Cells("ProcessID").Value = ""
                    BatchRow.Cells("ExtractionStartTime").Value = sExtractionEndTime
                    BatchRow.Cells("ExtractionEndTime").Value = sExtractionEndTime
                    BatchRow.Cells("SummaryReportLocation").Value = sSummaryReportLocation
                    BatchRow.Cells("ExtractionStatus").Value = sExtractionStatus
                    If sExtractionStatus = vbNullString Then
                        BatchRow.Cells("StatusImage").Value = My.Resources.not_selected_small
                        BatchRow.DefaultCellStyle.ForeColor = Color.Black
                    ElseIf sExtractionStatus = "Completed" Then
                        BatchRow.Cells("StatusImage").Value = My.Resources.Green_check_small
                        BatchRow.DefaultCellStyle.ForeColor = Color.Green
                    ElseIf sExtractionStatus.Contains("Failed") Then
                        BatchRow.Cells("StatusImage").Value = My.Resources.red_stop_small
                        BatchRow.DefaultCellStyle.ForeColor = Color.Red
                    End If
                    BatchRow.Cells("ProcessingFilesDirectory").Value = sProcessingFilesDirectory
                End With
            End While

            sqlConnection.Close()
            grdCustodainSMTPInfo.ScrollBars = ScrollBars.Both
            blnLoadGridFromDB = True

        Catch ex As Exception
            MessageBox.Show("Load Data from Grid Error - " & ex.ToString)
        End Try

    End Function

    Private Sub btnLoadPreviousBatches_Click(sender As Object, e As EventArgs) Handles btnLoadPreviousBatches.Click
        Dim bstatus As Boolean
        grdCustodainSMTPInfo.Rows.Clear()
        bstatus = blnLoadGridFromDB(grdCustodainSMTPInfo)
    End Sub

    Private Sub btnExtractShowSettings_MouseHover(sender As Object, e As EventArgs) Handles btnExtractShowSettings.MouseHover
        EWSExtractToolTip.Show("Show Settings...", btnExtractShowSettings)
    End Sub

    Private Sub btnLoadPreviousBatches_MouseHover(sender As Object, e As EventArgs) Handles btnLoadPreviousBatches.MouseHover
        EWSExtractToolTip.Show("Load Previous Batches", btnLoadPreviousBatches)

    End Sub

    Private Sub btnLaunchO365Extractions_MouseHover(sender As Object, e As EventArgs) Handles btnLaunchO365Extractions.MouseHover
        EWSExtractToolTip.Show("Launch EWS Extraction", btnLaunchO365Extractions)
    End Sub

    Private Sub grdCustodainSMTPInfo_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdCustodainSMTPInfo.CellContentDoubleClick

        If ((e.ColumnIndex = 10) And (grdCustodainSMTPInfo("ProcessingFilesDirectory", e.RowIndex).Value <> vbNullString)) Then
            If Directory.Exists(grdCustodainSMTPInfo("ProcessingFilesDirectory", e.RowIndex).Value) Then
                Process.Start("explorer.exe", grdCustodainSMTPInfo("ProcessingFilesDirectory", e.RowIndex).Value)
            End If
        ElseIf ((e.ColumnIndex = 11) And (grdCustodainSMTPInfo("CaseDirectory", e.RowIndex).Value <> vbNullString)) Then
            If Directory.Exists(grdCustodainSMTPInfo("CaseDirectory", e.RowIndex).Value) Then
                Process.Start("explorer.exe", grdCustodainSMTPInfo("CaseDirectory", e.RowIndex).Value)
            End If
        ElseIf ((e.ColumnIndex = 12) And (grdCustodainSMTPInfo("OutputDirectory", e.RowIndex).Value <> vbNullString)) Then
            If Directory.Exists(grdCustodainSMTPInfo("OutputDirectory", e.RowIndex).Value) Then
                Process.Start("explorer.exe", grdCustodainSMTPInfo("OutputDirectory", e.RowIndex).Value)
            End If
        ElseIf ((e.ColumnIndex = 13) And (grdCustodainSMTPInfo("LogDirectory", e.RowIndex).Value <> vbNullString)) Then
            If Directory.Exists(grdCustodainSMTPInfo("LogDirectory", e.RowIndex).Value) Then
                Process.Start("explorer.exe", grdCustodainSMTPInfo("LogDirectory", e.RowIndex).Value)
            End If
        ElseIf ((e.ColumnIndex = 16) And (grdCustodainSMTPInfo("SummaryReportLocation", e.RowIndex).Value <> vbNullString)) Then
            System.Diagnostics.Process.Start(grdCustodainSMTPInfo("SummaryReportLocation", e.RowIndex).Value.ToString)
        End If

    End Sub

    Private Sub btnClearDates_Click(sender As Object, e As EventArgs) Handles btnClearDates.Click
        setDateTimePickerBlank(FromDateTimePicker)
        setDateTimePickerBlank(ToDateTimePicker)
    End Sub

    Private Sub setDateTimePickerBlank(ByVal dateTimePicker As DateTimePicker)
        dateTimePicker.Visible = True
        dateTimePicker.Format = DateTimePickerFormat.Custom
        dateTimePicker.CustomFormat = "    "
    End Sub

    Private Sub btnExportProcessingDetails_Click(sender As Object, e As EventArgs) Handles btnExportProcessingDetails.Click
        Dim ReportOutputFile As StreamWriter
        Dim sOutputFileName As String
        Dim sMachineName As String


        Dim sSMTPAddress As String
        Dim sExtractionRoot As String
        Dim sExtractionStatus As String
        Dim sGroupID As String
        Dim sExtractedItems As String
        Dim sExtractedSize As String
        Dim sProcessID As String
        Dim sCaseLocation As String
        Dim sOutputDirectory As String
        Dim sLogDirectory As String
        Dim sExtractionStartTime As String
        Dim sExtractionEndTime As String
        Dim sSummaryReportLocation As String

        sMachineName = System.Net.Dns.GetHostName()
        sOutputFileName = "EWS Processing Statistic - " & sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv"
        ReportOutputFile = New StreamWriter(eMailArchiveMigrationManager.NuixFilesDir & "\" & sOutputFileName)

        ReportOutputFile.WriteLine("Custodian Name, Percent Completed, Bytes Uploaded, Progress Status, Number of PSTs, Size of PSTs, PST Path, Destination Folder, Destination Root, Destination SMTP, Group ID, Process ID, Ingestion Start Time, Success, Failed, Summary Report Location")

        For Each row In grdCustodainSMTPInfo.Rows
            sSMTPAddress = row.cells("SMTPAddress").value
            If sSMTPAddress <> vbNullString Then

                sExtractionRoot = row.cells("ExtractionRoot").value
                sExtractionStatus = row.cells("ExtractionStatus").value
                sGroupID = row.cells("GroupID").value
                sExtractedItems = row.cells("ExtractedItems").value
                sExtractedSize = row.cells("ExtractedSize").value
                sProcessID = row.cells("ProcessID").value
                sCaseLocation = row.cells("CaseDirectory").value
                sOutputDirectory = row.cells("OutputDirectory").value
                sLogDirectory = row.cells("LogDirectory").value
                sExtractionStartTime = row.cells("ExtractionStartTime").value
                sExtractionEndTime = row.cells("ExtractionEndTime").value
                sSummaryReportLocation = row.cells("SummaryReportLocation").value

                ReportOutputFile.WriteLine("""" & sSMTPAddress & """" & "," & """" & sExtractionRoot & """" & "," & """" & sExtractionStatus & """" & "," & """" & sGroupID & """" & "," & """" & sExtractedItems & """" & "," & """" & sExtractedSize & """" & "," & """" & sProcessID & """" & "," & """" & sCaseLocation & """" & "," & """" & sOutputDirectory & """" & "," & """" & sLogDirectory & """" & "," & """" & sExtractionStartTime & """" & "," & """" & sExtractionEndTime & """" & "," & """" & sSummaryReportLocation & """")
            End If
        Next

        ReportOutputFile.Close()
        MessageBox.Show("Nuix Case Statistics report finished building.  Report located at: " & eMailArchiveMigrationManager.NuixFilesDir & "\" & sOutputFileName)
    End Sub


    Private Sub grdCustodainSMTPInfo_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdCustodainSMTPInfo.CellContentClick

    End Sub
End Class