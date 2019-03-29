Imports System.IO
Imports System.Diagnostics
Imports System.Net
Imports System.Threading
Imports System.Xml
Imports System.Text
Imports System.Data.SQLite
Imports System.Data.SqlClient

Module CommonFunctions
    Public psSettingsFile As String
    Public psIngestionLogFile As String
    Public psSQLiteLocation As String
    Public pbNoMoreJobs As Boolean
    Public psSummaryReportFile As String
    Public psNuixCaseDir As String
    Public plstNuixConsoleProcesses As List(Of String)


    Public Function blnUpdateCustodianDBInfo(ByVal sSQLiteLocation As String, ByVal sCustodianName As String, ByVal sFieldName As String, sFieldValue As String) As Boolean
        blnUpdateCustodianDBInfo = False
        Dim sUpdateArchiveExtractionData As String

        Using Connection As New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;New=False;Compress=True;")
            sUpdateArchiveExtractionData = "Update DataConversionStats set " & sFieldName & " = " & "@" & sFieldName
            sUpdateArchiveExtractionData = sUpdateArchiveExtractionData & " WHERE CustodianName = @CustodianName"

            Using oUpdateArchiveExtractionDataCommand As New SQLiteCommand()
                With oUpdateArchiveExtractionDataCommand
                    .Parameters.AddWithValue("@CustodianName", sCustodianName)
                    .Parameters.AddWithValue("@" & sFieldName, sFieldValue)
                    .CommandText = sUpdateArchiveExtractionData
                    .Connection = Connection
                End With
                Try
                    Connection.Open()
                    oUpdateArchiveExtractionDataCommand.ExecuteNonQuery()
                    Connection.Close()
                Catch ex As Exception

                End Try
            End Using
        End Using

        blnUpdateCustodianDBInfo = True
    End Function


    Public Function blnUpdateExtractionDBInfo(ByVal sSQLiteLocation As String, ByVal sCustodianName As String, ByVal sFieldName As String, sFieldValue As String) As Boolean
        blnUpdateExtractionDBInfo = False
        Dim sUpdateArchiveExtractionData As String

        Using Connection As New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;New=False;Compress=True;")
            sUpdateArchiveExtractionData = "Update ewsExtractionStats set " & sFieldName & " = " & "@" & sFieldName
            sUpdateArchiveExtractionData = sUpdateArchiveExtractionData & " WHERE CustodianName = @CustodianName"

            Using oUpdateArchiveExtractionDataCommand As New SQLiteCommand()
                With oUpdateArchiveExtractionDataCommand
                    .Parameters.AddWithValue("@CustodianName", sCustodianName)
                    .Parameters.AddWithValue("@" & sFieldName, sFieldValue)
                    .CommandText = sUpdateArchiveExtractionData
                    .Connection = Connection
                End With
                Try
                    Connection.Open()
                    oUpdateArchiveExtractionDataCommand.ExecuteNonQuery()
                    Connection.Close()
                Catch ex As Exception

                End Try
            End Using
        End Using

        blnUpdateExtractionDBInfo = True
    End Function


    Public Function blnInsertDataConversionStats(ByVal sCustodianName As String, ByRef iCustodianID As Integer, ByVal sGroupID As String, ByVal sStatus As String, ByVal sProcessID As String, ByVal sRedisSettings As String, ByVal sSourceType As String, ByVal sSourceFormat As String, ByVal sOutputType As String, ByVal sOutputFormat As String, ByVal sConversionStartTime As String, ByVal sConversionEndTime As String, ByVal iBytesProcessed As Integer, ByVal iPercentComplete As Integer, ByVal iSuccess As Integer, ByVal iFailed As Integer, ByVal sSummaryReportLocation As String) As Boolean
        blnInsertDataConversionStats = False
        Dim sInsertArchiveExtractionData As String
        Dim sQueryReturnedCustodianID As String
        Dim SQLiteConnection As SQLiteConnection

        SQLiteConnection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
        SQLiteConnection.Open()

        Using Connection As New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;")
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

    Public Function blnInsertSourceDataStats(ByVal sSourceFileName As String, ByVal sSourceFilePath As String, ByVal sSourceFileCreateDate As String, ByVal sSourceFileModifiedDate As String, ByVal dblSourceFileSize As Double, ByVal iCustodianID As Integer) As Boolean
        blnInsertSourceDataStats = False
        Dim sInsertSourceDataStatsQuery As String
        Dim sQueryReturnedCustodianID As String
        Dim SQLiteConnection As SQLiteConnection

        SQLiteConnection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
        SQLiteConnection.Open()

        Using Connection As New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;")
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

    Public Function blnLoadPNSFInfoFromDirectory(ByVal sNSFFolderName As String, ByVal grdNSFInfo As DataGridView) As Boolean
        Dim ReportOutputFile As StreamWriter
        Dim sMachineName As String
        Dim sOutputFileName As String

        sMachineName = System.Net.Dns.GetHostName()
        sOutputFileName = sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv"

        ReportOutputFile = New StreamWriter(sNSFFolderName & "\" & sOutputFileName)
        ReportOutputFile.WriteLine("Custodian Name" & "," & "NSF Name" & "," & "NSF Size" & "," & "Folder")

        blnLoadPNSFInfoFromDirectory = False
        'bStatus = blnBuildNSFReports(sNSFFolderName, ReportOutputFile)

        blnLoadPNSFInfoFromDirectory = True

    End Function

    Sub subSourceDirSearch(ByVal sDir As String, ByRef iTotalNumberOfSourceDataFiles As Integer, ByRef dblTotalSizeOfSourceData As Double, ByVal sSourceDataType As String)
        Dim d As String
        Dim Length As String
        Dim currentdirectory As DirectoryInfo
        Dim extension As String
        Dim sFileName As String
        Dim sFilePath As String
        Dim sFileCreateDate As String
        Dim sFileModifiedDate As String
        Dim dblFileSize As Double
        Dim iCustodianRowID As Integer
        Dim sCustodianName As String
        Dim bStatus As Boolean
        Dim dbService As New DatabaseService
        Dim common As New Common

        Dim sSQLiteDatabaseFullName As String

        sSQLiteDatabaseFullName = eMailArchiveMigrationManager.SQLiteDBLocation & "\NuixEmailArchiveMigrationManager.db3"

        Try
            currentdirectory = New DirectoryInfo(sDir)

            For Each File In currentdirectory.GetFiles
                extension = Path.GetExtension(File.ToString)
                sCustodianName = Path.GetFileNameWithoutExtension(File.ToString)
                If extension = "." & sSourceDataType Then
                    sFileName = File.Name.ToString
                    sFilePath = File.DirectoryName.ToString
                    dblFileSize = File.Length
                    sFileCreateDate = File.CreationTime
                    sFileModifiedDate = File.LastWriteTime

                    bStatus = dbService.GetCustodianRowID(sSQLiteDatabaseFullName, sCustodianName, iCustodianRowID)

                    bStatus = dbService.InsertSourceDataStats(sSQLiteDatabaseFullName, sFileName, sFilePath, sFileCreateDate, sFileModifiedDate, dblFileSize, iCustodianRowID)
                    iTotalNumberOfSourceDataFiles = iTotalNumberOfSourceDataFiles + 1
                    dblTotalSizeOfSourceData = dblTotalSizeOfSourceData + dblFileSize
                End If
            Next

            Length = Directory.GetDirectories(sDir).Length
            If Length > 0 Then
                For Each d In Directory.GetDirectories(sDir)
                    subSourceDirSearch(d, iTotalNumberOfSourceDataFiles, dblTotalSizeOfSourceData, sSourceDataType)
                Next
            End If
        Catch ex As Exception
            common.Logger(psIngestionLogFile, "Error 7 - Directory Search" & ex.ToString)
        End Try

    End Sub

    Public Function blnGetCustodianRowID(ByVal sCustodianName As String, ByRef iCustodianRowID As Integer) As Boolean

        blnGetCustodianRowID = False

        Dim sQueryReturnedCustodianID As String
        Dim SQLiteConnection As SQLiteConnection

        SQLiteConnection = New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
        SQLiteConnection.Open()

        Using Connection As New SQLiteConnection("Data Source=" & psSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;")
            Using SQLSelectCommand As New SQLiteCommand("SELECT ROWID FROM DataConversionStats WHERE CustodianName='" & sCustodianName & "'")
                With SQLSelectCommand
                    .Connection = SQLiteConnection
                    Using readerObject As SQLiteDataReader = SQLSelectCommand.ExecuteReader
                        While readerObject.Read
                            sQueryReturnedCustodianID = readerObject("ROWID").ToString
                        End While
                    End Using
                End With

            End Using

            If sQueryReturnedCustodianID = vbNullString Then
                iCustodianRowID = vbNull
            Else
                iCustodianRowID = CInt(sQueryReturnedCustodianID)
            End If
            SQLiteConnection.Close()
        End Using

        blnGetCustodianRowID = True
    End Function

    Public Function blnBuildCustodianNuixFiles(ByVal asCustodianInfo() As String, ByVal sNuixAppMemory As String, sDataOutputType As String) As Boolean
        Dim bStatus As Boolean
        Dim sRubyFileName As String
        Dim sMappingFileName As String
        Dim sJSonFileName As String
        Dim sEvidenceKeyStore As String
        Dim sEvidencePassword As String
        Dim sCaseName As String
        Dim sEvidenceLocation As String
        Dim sWorkerTempDir As String
        Dim bExtractFromSlackSpace As Boolean

        Dim common As New Common

        blnBuildCustodianNuixFiles = False

        My.Computer.FileSystem.CreateDirectory(eMailArchiveMigrationManager.NuixFilesDir & "\Email Conversion\Nuix Files")
        sMappingFileName = eMailArchiveMigrationManager.NuixFilesDir & "\Email Conversion\Nuix Files\" & asCustodianInfo(0) & "\" & asCustodianInfo(0) & "_mapping.csv"
        sRubyFileName = eMailArchiveMigrationManager.NuixFilesDir & "\Email Conversion\Nuix Files\" & asCustodianInfo(0) & "\" & asCustodianInfo(0) & ".rb"
        sJSonFileName = eMailArchiveMigrationManager.NuixFilesDir & "\Email Conversion\Nuix Files\" & asCustodianInfo(0) & "\" & asCustodianInfo(0) & ".json"
        sWorkerTempDir = eMailArchiveMigrationManager.NuixWorkerTempDir
        bExtractFromSlackSpace = eMailArchiveMigrationManager.extractFromSlackSpace

        sCaseName = eMailArchiveMigrationManager.NuixCaseDir & "\Email Conversion\" & asCustodianInfo(0)
        sEvidenceLocation = asCustodianInfo(1)
        My.Computer.FileSystem.CreateDirectory(eMailArchiveMigrationManager.NuixFilesDir & "\Email Conversion\Nuix Files\" & asCustodianInfo(0))

        bStatus = blnBuildBatchFiles(asCustodianInfo(0), eMailArchiveMigrationManager.NuixFilesDir & "\Email Conversion\Nuix Files\" & asCustodianInfo(0) & "\", sRubyFileName, sMappingFileName, sNuixAppMemory)

        bStatus = blnBuildNSFConversionRubyScript(sRubyFileName, psSQLiteLocation, sJSonFileName, sCaseName, eMailArchiveMigrationManager.O365NumberOfNuixWorkers, eMailArchiveMigrationManager.O365MemoryPerWorker, sEvidenceLocation, sWorkerTempDir, bExtractFromSlackSpace)

        sEvidenceKeyStore = ""
        sEvidencePassword = ""
        bStatus = blnBuildJSonFile(sJSonFileName, asCustodianInfo(1), sEvidenceKeyStore, sEvidencePassword, asCustodianInfo(0), asCustodianInfo(2), "lightspeed")
        bStatus = common.BuildExtractionCustodianMappingFile(asCustodianInfo, sMappingFileName, sDataOutputType)

        asCustodianInfo = Nothing

        blnBuildCustodianNuixFiles = True
    End Function

    Public Function blnBuildBatchFiles(ByVal sCustodianName As String, ByVal sDirectoryName As String, ByVal sRubyFileName As String, ByVal sMappingFileName As String, ByVal sNuixAppMemory As String) As Boolean

        blnBuildBatchFiles = False

        Dim CustodianBatchFile As StreamWriter

        Dim sStartUpBatchFileName As String
        Dim sLicenceSourceType As String

        sStartUpBatchFileName = sDirectoryName & "\" & sCustodianName & ".bat"
        sStartUpBatchFileName = sStartUpBatchFileName.Replace("\\", "\")

        My.Computer.FileSystem.CreateDirectory(eMailArchiveMigrationManager.NuixLogDir & "\" & sCustodianName)

        If eMailArchiveMigrationManager.NMSSourceType = "Desktop" Then
            sLicenceSourceType = " -licencesourcetype dongle"
        Else
            sLicenceSourceType = " -licencesourcetype server -licencesourcelocation " & eMailArchiveMigrationManager.NuixNMSAddress & ":" & eMailArchiveMigrationManager.NuixNMSPort & " -Dnuix.registry.servers=" & eMailArchiveMigrationManager.NuixNMSAddress

        End If
        CustodianBatchFile = New StreamWriter(sStartUpBatchFileName)
        CustodianBatchFile.WriteLine("::Title will be the NSF Custodian")
        CustodianBatchFile.WriteLine("@TITLE " & sCustodianName)
        CustodianBatchFile.WriteLine("::Enter NMS Username on Line 4")
        CustodianBatchFile.WriteLine("@SET NUIX_USERNAME=" & eMailArchiveMigrationManager.NuixNMSUserName)
        CustodianBatchFile.WriteLine("::Enter NMS Username on Line 6")
        CustodianBatchFile.WriteLine("@SET NUIX_PASSWORD=" & eMailArchiveMigrationManager.NuixNMSAdminInfo)

        CustodianBatchFile.Write("""" & eMailArchiveMigrationManager.NuixAppLocation & """" & sLicenceSourceType)
        CustodianBatchFile.Write(" -licencetype email-archive-examiner -licenceworkers " & eMailArchiveMigrationManager.O365NumberOfNuixWorkers & " " & sNuixAppMemory & " -Dnuix.logdir=" & """" & eMailArchiveMigrationManager.NuixLogDir & "\" & sCustodianName & """" & " -Djava.io.tmpdir=" & """" & eMailArchiveMigrationManager.NuixJavaTempDir & """" & " -Dnuix.worker.tmpdir=" & """" & eMailArchiveMigrationManager.NuixWorkerTempDir & """" & " -Dnuix.crackAndDump.exportDir=" & """" & eMailArchiveMigrationManager.NuixExportDir & "\" & sCustodianName & """")
        CustodianBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & "pst:configFile=>" & sMappingFileName & """")
        CustodianBatchFile.Write(" -Dnuix.processing.crackAndDump.useRelativePath=true -Dnuix.processing.worker.timeout=" & eMailArchiveMigrationManager.WorkerTimeout & " -Dnuix.processing.crackAndDump.prependEvidenceName=true ")
        CustodianBatchFile.Write(" -Dnuix.export.mailbox.maximumFileSizePerMailbox=" & eMailArchiveMigrationManager.PSTExportSize & "GB")
        CustodianBatchFile.WriteLine("""" & sRubyFileName & """")

        'CustodianBatchFile.WriteLine("@pause")
        CustodianBatchFile.WriteLine("@exit")

        CustodianBatchFile.Close()

        blnBuildBatchFiles = True
    End Function

    Public Function blnBuildUpdatedRubyScript(ByVal sCustodianName As String, ByVal sRubyScriptFileName As String, sPSTLocation As String, ByVal sCustodianJSonFile As String, ByVal sWSSFileName As String) As Boolean
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

        CustodianRuby.WriteLine("load """ & eMailArchiveMigrationManager.SQLiteDBLocation.Replace("\", "\\") & "\\SQLite.rb_""")
        CustodianRuby.WriteLine("load """ & eMailArchiveMigrationManager.SQLiteDBLocation.Replace("\", "\\") & "\\Database.rb_""")
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
        CustodianRuby.WriteLine("$current_case = caseFactory.create(""" & eMailArchiveMigrationManager.NuixCaseDir.Replace("\", "\\") & "\\#{o365_custodian_name}" & """" & ", case_settings)")

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
        CustodianRuby.WriteLine("		:reportProcessingStatus => " & """" & "physical_files," & """")
        CustodianRuby.WriteLine("       :wss =>" & """" & sWSSFileName & """")
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
        CustodianRuby.WriteLine("total_time = '%.2f' % ((end_time-start_time)/60)")
        CustodianRuby.WriteLine("display_time =  '%.0f' % (total_time)")
        CustodianRuby.WriteLine("length = 60 / display_time.to_i")
        CustodianRuby.WriteLine("size = length.to_i * o365_source_size.to_i")
        CustodianRuby.WriteLine("speed = size.to_i / 1024 / 1024 / 1024.round(2)")
        CustodianRuby.WriteLine("final = speed.round(2)")
        CustodianRuby.WriteLine("puts")
        CustodianRuby.WriteLine("puts """ & "Office 365 Ingestion for #{o365_mailbox} finished at #{end_time}""")
        CustodianRuby.WriteLine("puts")
        CustodianRuby.WriteLine("puts """ & "Completed in #{display_time} minutes and averaged #{final} GB/hr.""")
        CustodianRuby.WriteLine("puts")

        CustodianRuby.WriteLine("updated_callback = [""" & "100" & """]")
        CustodianRuby.WriteLine("db.update(""" & "UPDATE ewsIngestionStats SET PercentCompleted = ? WHERE CustodianName = '#{o365_custodian_name}'" & """" & ",updated_callback)")

        CustodianRuby.WriteLine("$current_case.close")
        CustodianRuby.WriteLine("return")
        CustodianRuby.Close()

        blnBuildUpdatedRubyScript = True

    End Function

    Public Function blnBuildJSonFile(ByVal sJSonFileName As String, sSourceFolders As String, ByVal sEvidenceKeyStore As String, ByVal sEvidencePassword As String, ByVal sCaseName As String, ByVal sSourceSize As String, ByVal sMigrationType As String) As Boolean
        blnBuildJSonFile = False


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
        CustodianJSon.WriteLine("     " & """" & "evidence_name" & """" & ": " & """" & sCaseName & """" & ",")
        CustodianJSon.WriteLine("     " & """" & "centera" & """" & ": " & """" & "no" & """" & ",")
        CustodianJSon.WriteLine("     " & """" & "centera_ip" & """" & ": " & """" & """" & ",")
        CustodianJSon.WriteLine("     " & """" & "centera_clip" & """" & ": " & """" & """" & ",")
        CustodianJSon.WriteLine("     " & """" & "source_size" & """" & ": " & """" & sSourceSize & """" & ",")
        CustodianJSon.WriteLine("     " & """" & "migration" & """" & ": " & """" & sMigrationType & """")
        CustodianJSon.WriteLine(" }")
        CustodianJSon.WriteLine("}")

        CustodianJSon.Close()

        blnBuildJSonFile = True

    End Function


    Public Function blnBuildNSFConversionRubyScript(ByVal sRubyScriptFileName As String, ByVal sSQLiteDBLocation As String, ByVal sEvidenceJSon As String, ByVal sCaseDir As String, ByVal sNumberOfWorkers As String, ByVal sMemoryPerWorker As String, ByVal sCustodianEvidenceLocation As String, ByVal sWorkerTempDir As String, ByVal bExtractFromSlackSpace As Boolean) As Boolean
        blnBuildNSFConversionRubyScript = False

        Dim NSFConversionRuby As StreamWriter

        NSFConversionRuby = New StreamWriter(sRubyScriptFileName)

        NSFConversionRuby.WriteLine("# Menu Title: NEAMM NSF Conversion")
        NSFConversionRuby.WriteLine("# Needs Selected Items: false")
        NSFConversionRuby.WriteLine("# ")
        NSFConversionRuby.WriteLine("# This script expects a JSON configured with NSF parameters completed in order")
        NSFConversionRuby.WriteLine("# to automatically convert data from NSF to EML or PST.   ")
        NSFConversionRuby.WriteLine("# ")
        NSFConversionRuby.WriteLine("# Version 2.0")
        NSFConversionRuby.WriteLine("# March 22 2017 - Alex Chatzistamatis, Nuix")
        NSFConversionRuby.WriteLine("# ")
        NSFConversionRuby.WriteLine("#######################################")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("require 'thread'")
        NSFConversionRuby.WriteLine("require 'json'")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("CALLBACK_FREQUENCY = 100")
        NSFConversionRuby.WriteLine("callback_count = 0")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("load """ & sSQLiteDBLocation.Replace("\", "\\") & "\\Database.rb_""")
        NSFConversionRuby.WriteLine("load """ & sSQLiteDBLocation.Replace("\", "\\") & "\\SQLite.rb_""")
        NSFConversionRuby.WriteLine("db = SQLite.new(""" & sSQLiteDBLocation.Replace("\", "\\") & "\\NuixEmailArchiveMigrationManager.db3" & """" & ")")
        NSFConversionRuby.WriteLine("worker_side_script = " & """" & "I:\\NEAMM\\NuixFiles\\Archive Extraction\\AXS1\\2017-07-21 03-42-31\\wss.rb" & """")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("#######################################")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("file = File.read('" & sEvidenceJSon.Replace("\", "\\") & "')")
        NSFConversionRuby.WriteLine("parsed = JSON.parse(file)")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("    archive_file = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_file" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_keystore = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_keystore" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_password = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_password" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_name = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_name" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_centera = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_centera_ip = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera_ip" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_centera_clip = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera_clip" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_source_size = parsed[" & """" & "email_archive" & """" & "][" & """" & "source_size" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_migration = parsed[" & """" & "email_archive" & """" & "][" & """" & "migration" & """" & "]")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("caseFactory = $utilities.getCaseFactory()")
        NSFConversionRuby.WriteLine("case_settings = {")
        NSFConversionRuby.WriteLine("    :compound => false,")
        NSFConversionRuby.WriteLine("    :name => " & """" & "#{archive_name}" & """" & ",")
        NSFConversionRuby.WriteLine("    :description => " & """" & "Created using Nuix Email Archive Migration Manager" & """" & ",")
        NSFConversionRuby.WriteLine("    :investigator => " & """" & "NEAMM NSF Conversion" & """")
        NSFConversionRuby.WriteLine("}")
        NSFConversionRuby.WriteLine("$current_case = caseFactory.create(" & """" & sCaseDir.Replace("\", "\\") & """" & ", case_settings)")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("processor = $current_case.createProcessor")
        NSFConversionRuby.WriteLine("if archive_migration == " & """" & "lightspeed" & """")
        NSFConversionRuby.WriteLine("	processing_settings = {")
        NSFConversionRuby.WriteLine("		:traversalScope => " & """" & "full_traversal" & """" & ",")
        NSFConversionRuby.WriteLine("		:analysisLanguage => " & """" & "en" & """" & ",")
        NSFConversionRuby.WriteLine("		:identifyPhysicalFiles => true,")
        NSFConversionRuby.WriteLine("		:reuseEvidenceStores => true,")
        NSFConversionRuby.WriteLine("		:reportProcessingStatus => " & """" & "physical_files," & """")
        NSFConversionRuby.WriteLine("       :workerItemCallback => " & """" & "ruby:#{IO.read(worker_side_script)}" & """")
        NSFConversionRuby.WriteLine("	}")
        NSFConversionRuby.WriteLine("end")
        NSFConversionRuby.WriteLine("processor.setProcessingSettings(processing_settings)")
        NSFConversionRuby.WriteLine("parallel_processing_settings = {")
        NSFConversionRuby.WriteLine("	:workerCount => " & sNumberOfWorkers & ",")
        NSFConversionRuby.WriteLine("	:workerMemory => " & sMemoryPerWorker & ",")
        NSFConversionRuby.WriteLine("	:embedBroker => true,")
        NSFConversionRuby.WriteLine("	:brokerMemory => " & sMemoryPerWorker & ",")
        NSFConversionRuby.WriteLine("   :workerTemp => " & """" & sWorkerTempDir.Replace("\", "\\") & """")

        NSFConversionRuby.WriteLine("}")
        NSFConversionRuby.WriteLine("processor.setParallelProcessingSettings(parallel_processing_settings)")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("# MIME Type Fiter for row-based items")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/csv" & """" & ", { process_embedded: false })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/tab-separated-values" & """" & ", { process_embedded: false })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.sqlite-database" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-registry" & """" & ", { text_strip: true})")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/x-plist" & """" & ", { process_embedded: false })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-logfile" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-mft" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-usnjrnl" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/exe" & """" & ", { process_text: false })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.tcpdump.pcap" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/x-common-log" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-iis-log" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-windows-event-log-record" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-windows-event-logx" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/x-pcapng" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.symantec-vault-stream-data" & """" & ", { enabled: false })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-cab-compressed" & """" & ", { process_embedded: true })")
        NSFConversionRuby.WriteLine(" ")
        NSFConversionRuby.WriteLine("cust_name = archive_name")
        NSFConversionRuby.WriteLine("evidence_container = processor.newEvidenceContainer(cust_name)")
        NSFConversionRuby.WriteLine("evidence_container.addFile(" & """" & sCustodianEvidenceLocation & """" & ")")
        NSFConversionRuby.WriteLine("	evidence_container.setEncoding(" & """" & "utf-8" & """" & ")")
        NSFConversionRuby.WriteLine("	evidence_container.save")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("	id_file = archive_keystore")
        NSFConversionRuby.WriteLine("   id_password = archive_password")
        NSFConversionRuby.WriteLine("   processor.addKeyStore(id_file,{")
        NSFConversionRuby.WriteLine("        :filePassword => id_password,")
        NSFConversionRuby.WriteLine("        :target => " & """" & "" & """")
        NSFConversionRuby.WriteLine("        })")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("start_time = Time.now")
        NSFConversionRuby.WriteLine("last_progress = Time.now")
        NSFConversionRuby.WriteLine("semaphore = Mutex.new")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("puts")
        NSFConversionRuby.WriteLine("puts " & """" & "NSF Conversion for #{archive_name} has started..." & """")
        NSFConversionRuby.WriteLine("puts")
        NSFConversionRuby.WriteLine("printf " & """" & "%-40s %-25s %-25s %-25s" & """" & "," & """" & "Timestamp" & """" & ", " & """" & "Bytes Converted" & """" & ", " & """" & "Total Bytes" & """" & ", " & """" & "Percent (%) Completed" & """")
        NSFConversionRuby.WriteLine("puts")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("processor.when_item_processed do |event|")
        NSFConversionRuby.WriteLine("   semaphore.synchronize {")
        NSFConversionRuby.WriteLine("	class Numeric")
        NSFConversionRuby.WriteLine("	  def percent_of(n)")
        NSFConversionRuby.WriteLine("        self.to_f / n.to_f * 100")
        NSFConversionRuby.WriteLine("      end")
        NSFConversionRuby.WriteLine("    end")
        NSFConversionRuby.WriteLine("    # Progress message every 15 seconds")
        NSFConversionRuby.WriteLine("	current_size = progress.get_current_size")
        NSFConversionRuby.WriteLine("	total_size = progress.get_total_size")
        NSFConversionRuby.WriteLine("	percent = current_size.percent_of(total_size).round(1)")
        NSFConversionRuby.WriteLine("    if (Time.now - last_progress) > 15")
        NSFConversionRuby.WriteLine("         last_progress = Time.now")
        NSFConversionRuby.WriteLine("			printf " & """" & "\r%-40s %-25s %-25s %-25s" & """" & ", Time.now, current_size, total_size, percent")
        NSFConversionRuby.WriteLine("			updated_callback = [current_size,percent]")
        NSFConversionRuby.WriteLine("			db.update(" & """" & "UPDATE DataConversionStats BytesProcessed = ?, PercentCompleted = ? WHERE CustodianName = '#{archive_name}'" & """" & ",updated_callback)")
        NSFConversionRuby.WriteLine("	end")
        NSFConversionRuby.WriteLine("    }")
        NSFConversionRuby.WriteLine("end")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("processor.process")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("puts")
        NSFConversionRuby.WriteLine("end_time = Time.now")
        NSFConversionRuby.WriteLine("total_time = '%.2f' % ((end_time-start_time)/60)")
        NSFConversionRuby.WriteLine("display_time =  '%.0f' % (total_time)")
        NSFConversionRuby.WriteLine("length = 60 / display_time.to_i")
        NSFConversionRuby.WriteLine("size = length.to_i * archive_source_size.to_i")
        NSFConversionRuby.WriteLine("speed = size.to_i / 1024 / 1024 / 1024.round(2)")
        NSFConversionRuby.WriteLine("final = speed.round(2)")
        NSFConversionRuby.WriteLine("puts")
        NSFConversionRuby.WriteLine("puts " & """" & "NSF Conversion for #{archive_name} has finished!" & """")
        NSFConversionRuby.WriteLine("puts")
        NSFConversionRuby.WriteLine("puts " & """" & "Completed in #{display_time} minutes and averaged #{final} GB/hr." & """")
        NSFConversionRuby.WriteLine("puts")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("updated_callback = [" & """" & "Conversion Completed" & """" & "]")
        NSFConversionRuby.WriteLine("db.update(" & """" & "UPDATE DataConversionStats SET ConversionStatus = ? WHERE CustodianName = '#{archive_name}'" & """" & ",updated_callback)")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("$current_case.close")
        NSFConversionRuby.WriteLine("return")
        NSFConversionRuby.Close()

        blnBuildNSFConversionRubyScript = True

    End Function

    Public Function blnGetSummaryReportInfo(ByVal sSummaryReport As String, ByRef sCustodianName As String, ByRef sItemsProcessed As String, ByRef sItemsExported As String, ByRef sItemsFailed As String, ByRef sItemsSkipped As String, ByRef sBytesUploaded As String, ByRef dStartDate As Date, ByRef dEndDate As Date) As Boolean
        blnGetSummaryReportInfo = False

        Dim FileStream As StreamReader
        Dim sCurrentRow As String
        Dim sMailboxName As String
        Dim asStartDateValues() As String
        Dim asEndDateValues() As String
        Dim asCustodianValues() As String
        Dim asItemsExported() As String
        Dim asItemsFailed() As String
        Dim asItemsProcessed() As String
        Dim asItemsSkipped() As String
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
                ElseIf (sCurrentRow.Contains("Items processed for export:")) Then
                    asItemsProcessed = Split(sCurrentRow, ":")
                    sItemsProcessed = asItemsProcessed(UBound(asItemsProcessed))

                ElseIf sCurrentRow.Contains("Total items skipped:") Then
                    asItemsSkipped = Split(sCurrentRow, ":")
                    sItemsSkipped = asItemsSkipped(UBound(asItemsSkipped))
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


    Public Function blnComputeFolderSize(ByVal sFolderName As String, ByRef dblTotalFolderSize As Double) As Boolean
        blnComputeFolderSize = False
        Dim FolderInfo = New IO.DirectoryInfo(sFolderName)
        'Add the files to treeview
        If FolderInfo.Attributes <> FileAttributes.ReadOnly Then
            For Each File In FolderInfo.GetFiles : dblTotalFolderSize += File.Length
            Next
            For Each SubFolderInfo In FolderInfo.GetDirectories : blnComputeFolderSize(SubFolderInfo.FullName, dblTotalFolderSize)
            Next
        End If
        blnComputeFolderSize = True
    End Function

    Public Function blnProcessSummaryReportFiles(ByVal sSummaryReportFile As String, ByRef sTotalItemsProcessed As String, ByRef sTotalItemsExported As String, ByRef sTotalItemsSkipped As String, ByRef sTotalItemsFailed As String) As Boolean
        Dim sCurrentRow As String
        Dim fileSummaryFile As StreamReader
        Dim iCounter As Integer
        Dim bFailedItemStart As Boolean

        Dim common As New Common

        blnProcessSummaryReportFiles = True

        Try
            iCounter = iCounter + 1
            bFailedItemStart = False
            common.Logger(psIngestionLogFile, "Checking " & sSummaryReportFile & " for items processed.")
            If File.Exists(sSummaryReportFile) Then
                common.Logger(psIngestionLogFile, "Opening " & sSummaryReportFile & " for processing")
                fileSummaryFile = New StreamReader(sSummaryReportFile)
                While Not fileSummaryFile.EndOfStream
                    sCurrentRow = fileSummaryFile.ReadLine
                    If sCurrentRow.Contains("Total items exported:") Then
                        sTotalItemsExported = sCurrentRow.Replace("Total items exported:", "")
                        blnProcessSummaryReportFiles = False
                    ElseIf sCurrentRow.Contains("Total items skipped:") Then
                        sTotalItemsSkipped = sCurrentRow.Replace("Total items skipped:", "")
                        blnProcessSummaryReportFiles = False
                    ElseIf sCurrentRow.Contains("Total number of failed items:") Then
                        sTotalItemsFailed = sCurrentRow.Replace("Total number of failed items:", "")
                        blnProcessSummaryReportFiles = False
                    ElseIf sCurrentRow.Contains("Items processed for export:") Then
                        sTotalItemsProcessed = sCurrentRow.Replace("Items processed for export:", "")
                        blnProcessSummaryReportFiles = False
                    End If
                End While
            Else
                common.Logger(psIngestionLogFile, "The summary report file specified at " & sSummaryReportFile & " does not exist")
                blnProcessSummaryReportFiles = False
            End If
            blnProcessSummaryReportFiles = True
        Catch ex As Exception
            MessageBox.Show("Counter = " & iCounter & ": Error in Process Summary Report Files " & ex.ToString)
            common.Logger(psIngestionLogFile, "Error Checking " & sSummaryReportFile & ex.ToString)
            blnProcessSummaryReportFiles = False
        End Try

    End Function

    Public Function blnCheckNuixLogForErrors(ByVal sNuixLogFileDir As String, ByRef sTotalItemsProcessed As String, ByRef sTotalItemsExported As String, ByRef sTotalItemsSkipped As String, ByRef sTotalItemsFailed As String) As Boolean
        Dim bStatus As Boolean
        Dim sNuixLogFile As String
        Dim NuixLogStreamReader As StreamReader
        Dim sCurrentRow As String
        Dim common As New Common

        blnCheckNuixLogForErrors = False

        bStatus = blnGetNuixLogFiles(sNuixLogFileDir, sNuixLogFile)
        If sNuixLogFile <> vbNullString Then
            Try
                NuixLogStreamReader = New StreamReader(sNuixLogFile)
                While Not NuixLogStreamReader.EndOfStream
                    sCurrentRow = NuixLogStreamReader.ReadLine
                    If sCurrentRow.Contains("Error running script:") Then
                        common.Logger(psIngestionLogFile, sNuixLogFile & " - Error " & sCurrentRow)
                        blnCheckNuixLogForErrors = True
                        Exit Function
                    ElseIf sCurrentRow.Contains("FATAL com.nuix.investigator.main.b - Couldn't acquire a licence") Then
                        blnCheckNuixLogForErrors = True
                        common.Logger(psIngestionLogFile, sNuixLogFile & " - Error " & sCurrentRow)
                        Exit Function
                    ElseIf sCurrentRow.Contains("No licences were found.") Then
                        blnCheckNuixLogForErrors = True
                        common.Logger(psIngestionLogFile, sNuixLogFile & " - Error " & sCurrentRow)
                        Exit Function
                    ElseIf sCurrentRow.Contains("Error opening case") Then
                        blnCheckNuixLogForErrors = True
                        common.Logger(psIngestionLogFile, sNuixLogFile & " - Error " & sCurrentRow)
                        Exit Function
                    ElseIf sCurrentRow.Contains("Items have an invalid product class") Then
                        blnCheckNuixLogForErrors = True
                        common.Logger(psIngestionLogFile, sNuixLogFile & " - Error " & sCurrentRow)
                        Exit Function
                    ElseIf sCurrentRow.Contains("Couldn't acquire a licence") Then
                        blnCheckNuixLogForErrors = True
                        common.Logger(psIngestionLogFile, sNuixLogFile & " - Error " & sCurrentRow)
                        Exit Function
                    ElseIf sCurrentRow.Contains("Total items exported:") Then
                        sTotalItemsExported = sCurrentRow.Replace("Total items exported:", "")
                        common.Logger(psIngestionLogFile, sNuixLogFile & " - Total Items Exported " & sCurrentRow)
                        blnCheckNuixLogForErrors = False
                    ElseIf sCurrentRow.Contains("Total items skipped:") Then
                        sTotalItemsSkipped = sCurrentRow.Replace("Total items skipped:", "")
                        common.Logger(psIngestionLogFile, sNuixLogFile & " - Total Items Skipped " & sCurrentRow)
                        blnCheckNuixLogForErrors = False
                    ElseIf sCurrentRow.Contains("Total number of failed items:") Then
                        sTotalItemsFailed = sCurrentRow.Replace("Total number of failed items:", "")
                        common.Logger(psIngestionLogFile, sNuixLogFile & " - Total Items Failed " & sCurrentRow)
                        blnCheckNuixLogForErrors = False
                    End If
                End While

            Catch ex As Exception
                MessageBox.Show("CheckNuixLog - " & ex.ToString, "blnCheckNuixLogForErrors", MessageBoxButtons.OK)
                common.Logger(psIngestionLogFile, "blnCheckNuixLogForErrors - " & ex.ToString)
                blnCheckNuixLogForErrors = False
            End Try
        End If

        'blnCheckNuixLogForErrors = True
    End Function

    Public Function blnGetNuixLogFiles(ByVal sNuixLogFileDir As String, ByRef sNuixLogFile As String) As Boolean
        Dim CurrentDirectory As DirectoryInfo

        CurrentDirectory = New DirectoryInfo(sNuixLogFileDir)
        If Not CurrentDirectory.Attributes.HasFlag(FileAttributes.ReadOnly) Then
            Try
                For Each Directory In CurrentDirectory.GetDirectories
                    blnGetNuixLogFiles(Directory.FullName, sNuixLogFile)
                Next

                For Each Files In CurrentDirectory.GetFiles
                    If Files.Name = "nuix.log" Then
                        sNuixLogFile = Files.FullName
                    End If
                Next
            Catch ex As Exception
                MessageBox.Show("GetNuixLogFiles -" & ex.Message)
            End Try
        End If

        blnGetNuixLogFiles = True

    End Function

    Public Function blnBuildWSSRubyScript(ByVal sRubyScriptFileName As String, ByVal sNuixFilesDir As String) As Boolean
        blnBuildWSSRubyScript = False

        Dim WSSRuby As StreamWriter

        WSSRuby = New StreamWriter(sRubyScriptFileName)
        WSSRuby.WriteLine("#ENV_JAVA[" & """" & "nuix.wss.config" & """" & "] = " & """" & sNuixFilesDir.Replace("\", "\\") & "\\wss.json" & """")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("require 'csv'")
        WSSRuby.WriteLine("require 'fileutils'")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("#Class for report files")
        WSSRuby.WriteLine("class WSS_Report")
        WSSRuby.WriteLine("  #makes directory if it does not exist.")
        WSSRuby.WriteLine("  def initialize(location)")
        WSSRuby.WriteLine("    @file_location = location")
        WSSRuby.WriteLine("    FileUtils.mkdir_p File.dirname(location)")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("#Class for per-item report")
        WSSRuby.WriteLine("# Reports: GUID, PATH, FLAGS")
        WSSRuby.WriteLine("class Item_Report < WSS_Report")
        WSSRuby.WriteLine("  def initialize(location)")
        WSSRuby.WriteLine("    super(location)")
        WSSRuby.WriteLine("    #Create CSV and write header")
        WSSRuby.WriteLine("    WSS.log(" & """" & "Creating item report at #{@file_location}" & """" & ", true)")
        WSSRuby.WriteLine("    CSV.open(@file_location, " & """" & "w" & """" & ") { |csv| csv << [" & """" & "GUID" & """" & ", " & """" & "Path" & """" & ", " & """" & "Flags" & """" & "] }")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Appends item to report.")
        WSSRuby.WriteLine("  def add(wss_item)")
        WSSRuby.WriteLine("    guid = wss_item.worker_item.get_guid_path.last")
        WSSRuby.WriteLine("    path = File.join(wss_item.source_item.get_path_names.to_a)")
        WSSRuby.WriteLine("    flags = wss_item.flags.to_a.sort.join(" & """" & ";" & """" & ")")
        WSSRuby.WriteLine("    row = [guid, path, flags]")
        WSSRuby.WriteLine("    CSV.open(@file_location, " & """" & "a" & """" & ") { |csv| csv << row }")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("#Class for per-query report")
        WSSRuby.WriteLine("# Reports: CATEGORY, NUMBER OF ITEMS")
        WSSRuby.WriteLine("#includes TOTAL to tally items that hit on any category.")
        WSSRuby.WriteLine("class Query_Report < WSS_Report")
        WSSRuby.WriteLine("  def initialize(location)")
        WSSRuby.WriteLine("    super(location)")
        WSSRuby.WriteLine("    @queries = Hash.new(0)")
        WSSRuby.WriteLine("    @total = 0")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Adds item to report.")
        WSSRuby.WriteLine("  def add(wss_item)")
        WSSRuby.WriteLine("    @total += 1")
        WSSRuby.WriteLine("    wss_item.flags.each do |f|")
        WSSRuby.WriteLine("      @queries[f] += 1")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Writes report from @queries and @total.")
        WSSRuby.WriteLine("  def write()")
        WSSRuby.WriteLine("    WSS.log(" & """" & "Creating query report at #{@file_location}" & """" & ", true)")
        WSSRuby.WriteLine("    CSV.open(@file_location, " & """" & "w" & """" & ") do |csv|")
        WSSRuby.WriteLine("      csv << [" & """" & "Category" & """" & ", " & """" & "Count" & """" & "]")
        WSSRuby.WriteLine("      csv << [" & """" & "TOTAL RESPONSIVE ITEMS" & """" & ", @total]")
        WSSRuby.WriteLine("      @queries.each do |k, v|")
        WSSRuby.WriteLine("        csv << [k.to_s, v]")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("#Module for normalizing metadata property object to string.")
        WSSRuby.WriteLine("module Metadata")
        WSSRuby.WriteLine("  #Returns property value as normalized string.")
        WSSRuby.WriteLine("  # param:  obj - Object from metadata properties Hash.")
        WSSRuby.WriteLine("  def self.stringify(obj)")
        WSSRuby.WriteLine("    case obj")
        WSSRuby.WriteLine("    when String")
        WSSRuby.WriteLine("      return obj")
        WSSRuby.WriteLine("    when TrueClass, FalseClass")
        WSSRuby.WriteLine("      return obj.to_s")
        WSSRuby.WriteLine("    when Fixnum, Float")
        WSSRuby.WriteLine("      return obj.to_s")
        WSSRuby.WriteLine("    when Java::OrgJodaTime::Duration")
        WSSRuby.WriteLine("      return obj.to_string")
        WSSRuby.WriteLine("    when org.joda.time.DateTime")
        WSSRuby.WriteLine("      return obj.to_string(" & """" & "YmdHMS" & """" & ")")
        WSSRuby.WriteLine("    when Java::byte[]")
        WSSRuby.WriteLine("      return obj.to_s.unpack(" & """" & "H*" & """" & ")[0]")
        WSSRuby.WriteLine("    when Java::ComNuixUtilExpression::b")
        WSSRuby.WriteLine("      return obj.to_s")
        WSSRuby.WriteLine("    else")
        WSSRuby.WriteLine("      return obj.to_a.map{ |e| stringify(e) }.join(" & """" & ";" & """" & ") if obj.respond_to?(:each)")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("    return obj.to_s")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("#Module for the Worker-Side Script, contains the configuration and search parameters, and methods for comparing against the parameters.")
        WSSRuby.WriteLine("module WSS")
        WSSRuby.WriteLine("  require 'json'")
        WSSRuby.WriteLine("  require 'csv'")
        WSSRuby.WriteLine("  require 'set'")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  @@config = nil")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  @@file_filters = Hash.new() # mime/kind/ext => Array of options")
        WSSRuby.WriteLine("  @@addresses = Hash.new{|h,k| h[k]=[]}    # tag => Array of { from/to/cc/bcc/recipient/all => Array of terms }")
        WSSRuby.WriteLine("  @@text_queries = Hash.new{|h,k| h[k]=[]}  # tag => Array of patterns")
        WSSRuby.WriteLine("  @@properties = Hash.new{|h,k| h[k]=[]}  # tag => Array of patterns (pattern could be " & """" & "field:value" & """" & ")")
        WSSRuby.WriteLine("  @@report_directory = nil")
        WSSRuby.WriteLine("  @@reports_hash = Hash.new() # type => WSS_Report")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  class << self")
        WSSRuby.WriteLine("    #Returns addresses Hash.")
        WSSRuby.WriteLine("    def addresses()")
        WSSRuby.WriteLine("      return @@addresses")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns if non-hit items should be case sensitive")
        WSSRuby.WriteLine("    def case_sensitive()")
        WSSRuby.WriteLine("      return @@config[:caseSensitive]")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns set of tags from @@addresses that matched comm.")
        WSSRuby.WriteLine("    # param:  comm -  an item's communication")
        WSSRuby.WriteLine("    def check_comm(comm)")
        WSSRuby.WriteLine("      tags = Set.new()")
        WSSRuby.WriteLine("      @@addresses.each do |tag, query_array|")
        WSSRuby.WriteLine("        query_array.each do |query_hash|")
        WSSRuby.WriteLine("          query_hash.each do |field, term|")
        WSSRuby.WriteLine("            next if tags.include?(tag)  #item has already hit on one of the terms")
        WSSRuby.WriteLine("            #Get Addresses")
        WSSRuby.WriteLine("            addys = []")
        WSSRuby.WriteLine("            case field")
        WSSRuby.WriteLine("            when :from")
        WSSRuby.WriteLine("              addys = comm.get_from().to_a")
        WSSRuby.WriteLine("            when :to")
        WSSRuby.WriteLine("              addys = comm.get_to().to_a")
        WSSRuby.WriteLine("            when :cc")
        WSSRuby.WriteLine("              addys = comm.get_cc().to_a")
        WSSRuby.WriteLine("            when :bcc")
        WSSRuby.WriteLine("              addys = comm.get_bcc().to_a")
        WSSRuby.WriteLine("            when :recipient")
        WSSRuby.WriteLine("              addys = [comm.get_to(), comm.get_cc(), comm.get_bcc()].flatten")
        WSSRuby.WriteLine("            else  #:address")
        WSSRuby.WriteLine("              addys = [comm.get_from(), comm.get_to(), comm.get_cc(), comm.get_bcc()].flatten")
        WSSRuby.WriteLine("            end")
        WSSRuby.WriteLine("            #Format Address")
        WSSRuby.WriteLine("            search_text = addys.map {|a| a.to_rfc822_string}")
        WSSRuby.WriteLine("            search_text = search_text.map {|e| e.downcase} if !case_sensitive()")
        WSSRuby.WriteLine("            #Add matches")
        WSSRuby.WriteLine("            tags.add(tag) if contains_term(search_text, term)")
        WSSRuby.WriteLine("          end")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      return tags")
        WSSRuby.WriteLine("     end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns T/F if @@file_filters contains one of the parameters.")
        WSSRuby.WriteLine("    # params: mime  - MIME-type string")
        WSSRuby.WriteLine("    #         kind  - kind string")
        WSSRuby.WriteLine("    #         ext   - file extension String")
        WSSRuby.WriteLine("    def check_file(mime, kind, ext)")
        WSSRuby.WriteLine("      mime = mime.downcase if !case_sensitive()")
        WSSRuby.WriteLine("      kind = kind.downcase if !case_sensitive()")
        WSSRuby.WriteLine("      ext = ext.downcase if !case_sensitive()")
        WSSRuby.WriteLine("      return ( @@file_filters[:mime].include?(mime) || @@file_filters[:kind].include?(kind) || @@file_filters[:ext].include?(ext) )")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns set of tags from @@properties that matched props parameter.")
        WSSRuby.WriteLine("    # param:  props - Hash of metadata properties")
        WSSRuby.WriteLine("    def check_props(props)")
        WSSRuby.WriteLine("      tags = Set.new()")
        WSSRuby.WriteLine("      @@properties.each do |tag, query_array|")
        WSSRuby.WriteLine("        query_array.each do |query_string|")
        WSSRuby.WriteLine("          next if tags.include?(tag)  #item has already hit on one of the terms")
        WSSRuby.WriteLine("          #Determine search_text and term to check.")
        WSSRuby.WriteLine("          query = query_string.split(" & """" & ":" & """" & ")")
        WSSRuby.WriteLine("          term = nil")
        WSSRuby.WriteLine("          search_text = Set.new()")
        WSSRuby.WriteLine("          if query.size == 2")
        WSSRuby.WriteLine("            name = query[0]")
        WSSRuby.WriteLine("            term = query[1]")
        WSSRuby.WriteLine("            #find relevant values based on query[0] to search with query[1]")
        WSSRuby.WriteLine("            fields = props.keys()")
        WSSRuby.WriteLine("            fields.each do |f|")
        WSSRuby.WriteLine("              f_name = f")
        WSSRuby.WriteLine("              f_name = f.downcase if !case_sensitive()")
        WSSRuby.WriteLine("              if !f_name.match(name).nil?")
        WSSRuby.WriteLine("                property = props[f]")
        WSSRuby.WriteLine("                property = property.to_a.join(" & """" & ";" & """" & ") if !property.is_a?(String)")
        WSSRuby.WriteLine("                search_text.add(Metadata.stringify(property))")
        WSSRuby.WriteLine("              end")
        WSSRuby.WriteLine("            end")
        WSSRuby.WriteLine("          else")
        WSSRuby.WriteLine("            term = query_string")
        WSSRuby.WriteLine("            search_text = props.map{ |k, v| " & """" & "#{k}:#{Metadata.stringify(v)}" & """" & " }")
        WSSRuby.WriteLine("          end")
        WSSRuby.WriteLine("          #Check for term within search_text")
        WSSRuby.WriteLine("          search_text = search_text.map {|e| e.downcase} if !case_sensitive()")
        WSSRuby.WriteLine("          tags.add(tag) if contains_term(search_text, term)")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      return tags")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns set of tags from @@text_queries that matched text.")
        WSSRuby.WriteLine("    # param:  text -  a String")
        WSSRuby.WriteLine("    def check_text(text)")
        WSSRuby.WriteLine("      tags = Set.new()")
        WSSRuby.WriteLine("      text = text.downcase if !case_sensitive()")
        WSSRuby.WriteLine("      @@text_queries.each do |tag, query_hash|")
        WSSRuby.WriteLine("        query_hash.each do |query|")
        WSSRuby.WriteLine("          next if tags.include?(tag)  #item has already hit on one of the terms")
        WSSRuby.WriteLine("          log(" & """" & "Searching text for: #{query}" & """" & ", nil)")
        WSSRuby.WriteLine("          tags.add(tag) if !text.match(query).nil?")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      return tags")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns T/F if the term matches any of the text the list.")
        WSSRuby.WriteLine("    # params: list  - Enumerable of items that can be to_string'ed")
        WSSRuby.WriteLine("    #         term  - The term to look for")
        WSSRuby.WriteLine("    def contains_term(list, term)")
        WSSRuby.WriteLine("      log(" & """" & "Looking for: #{term}" & """" & ", nil)")
        WSSRuby.WriteLine("      list.each do |text|")
        WSSRuby.WriteLine("        return true if !text.to_s.match(term).nil?")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      return false  #if there were no matches")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns name of custom metadata field to generate.")
        WSSRuby.WriteLine("    # nil or "" if no custom metadata should be created.")
        WSSRuby.WriteLine("    def custom_metadata()")
        WSSRuby.WriteLine("      return @@custom_metadata")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns if non-hit items should be excluded.")
        WSSRuby.WriteLine("    def exclude_items()")
        WSSRuby.WriteLine("      return @@config[:excludeItems]")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns flagUnresponsive bool.")
        WSSRuby.WriteLine("    def flag_unresponsive()")
        WSSRuby.WriteLine("      return @@config[:flagUnresponsive]")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns fileFiltering bool.")
        WSSRuby.WriteLine("    def file_filter()")
        WSSRuby.WriteLine("      return @@config[:fileFiltering]")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns Hash of " & """" & "communication field" & """" & " => pattern")
        WSSRuby.WriteLine("    #takes query like " & """" & "from:foo@bar.com" & """" & "or " & """" & "from-mail-domain:bar.com" & """" & " and converts to field=>pattern.")
        WSSRuby.WriteLine("    # param:  query_string  - string in Nuix Communication Fields search syntax")
        WSSRuby.WriteLine("    def get_field_query(query_string)")
        WSSRuby.WriteLine("      field = nil")
        WSSRuby.WriteLine("      pattern = nil")
        WSSRuby.WriteLine("      query = query_string.to_s.split(" & """" & ":" & """" & ")")
        WSSRuby.WriteLine("     if query.size == 2")
        WSSRuby.WriteLine("        if query[0].start_with?(" & """" & "from" & """" & ")")
        WSSRuby.WriteLine("          field = :from")
        WSSRuby.WriteLine("        elsif query[0].start_with?(" & """" & "to" & """" & ")")
        WSSRuby.WriteLine("          field = :to")
        WSSRuby.WriteLine("        elsif query[0].start_with?(" & """" & "cc" & """" & ")")
        WSSRuby.WriteLine("          field = :cc")
        WSSRuby.WriteLine("        elsif query[0].start_with?(" & """" & "bcc" & """" & ")")
        WSSRuby.WriteLine("          field = :bcc")
        WSSRuby.WriteLine("        elsif query[0].start_with?(" & """" & "recipient" & """" & ")")
        WSSRuby.WriteLine("          field = :recipient")
        WSSRuby.WriteLine("        elsif query[0].start_with?(" & """" & "address" & """" & ")")
        WSSRuby.WriteLine("          field = :all")
        WSSRuby.WriteLine("        else")
        WSSRuby.WriteLine("          #Trying to search unknown field")
        WSSRuby.WriteLine("          field = :all")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        #trim quotation marks from query if required")
        WSSRuby.WriteLine("        query[1] = query[1].slice(1..-2) if query[1].start_with?(" & """" & "\" & """" & """" & ") && query[1].end_with?(" & """" & "\" & """" & """" & ")")
        WSSRuby.WriteLine("        #handle special fields")
        WSSRuby.WriteLine("        if query[0].end_with?(" & """" & "mail-address" & """" & ")")
        WSSRuby.WriteLine("          pattern = " & """" & "/^#{query[1]}/" & """" & " #convert to Regex to capture bar@foo.com but not foobar@foo.com")
        WSSRuby.WriteLine("        elsif query[0].end_with?(" & """" & "mail-domain" & """" & ")")
        WSSRuby.WriteLine("          pattern = " & """" & "@#{query[1]}" & """")
        WSSRuby.WriteLine("        else")
        WSSRuby.WriteLine("          pattern = query[1]")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      else")
        WSSRuby.WriteLine("        #No field specified.")
        WSSRuby.WriteLine("        field = :all")
        WSSRuby.WriteLine("        pattern = query_string")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      pattern = Regexp.new( pattern.slice(1..-2) ) if pattern.start_with?(" & """" & "/" & """" & ") && pattern.end_with?(" & """" & "/" & """" & ")")
        WSSRuby.WriteLine("      return {field => pattern}")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Sets @@settings to a Hash from JSON.")
        WSSRuby.WriteLine("    # param:  json  - location of JSON file")
        WSSRuby.WriteLine("    def load(json)")
        WSSRuby.WriteLine("      log(" & """" & "Loading settings from #{json}" & """" & ", true)")
        WSSRuby.WriteLine("      json_file = File.read(json)")
        WSSRuby.WriteLine("      @@config = JSON.parse(json_file, :symbolize_names => true)")
        WSSRuby.WriteLine("      log(" & """" & "Configuration is: #{@@config}" & """" & ", nil)")
        WSSRuby.WriteLine("      @@file_filters = load_filter(@@config[:filterMimeTypes], @@config[:filterKinds], @@config[:filterFileExt]) if file_filter()")
        WSSRuby.WriteLine("      @@addresses = load_addresses(@@config[:communicationCSV]) if File.exist?(@@config[:communicationCSV])")
        WSSRuby.WriteLine("      @@properties = load_queries(@@config[:propertiesCSV]) if File.exist?(@@config[:propertiesCSV])")
        WSSRuby.WriteLine("      @@text_queries = load_queries(@@config[:textQueriesCSV]) if File.exist?(@@config[:textQueriesCSV])")
        WSSRuby.WriteLine("      if (report_items() || report_queries())")
        WSSRuby.WriteLine("        @@config[:reportDirectory] = ENV_JAVA[" & """" & "nuix.wss.reports" & """" & "] if !ENV_JAVA[" & """" & "nuix.wss.reports" & """" & "].nil?")
        WSSRuby.WriteLine("        @@report_directory = @@config[:reportDirectory]")
        WSSRuby.WriteLine("        log(" & """" & "Reporting to: #{@@report_directory}" & """" & ", true)")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      @@custom_metadata = load_custom_metadata(@@config[:customMetadata])")
        WSSRuby.WriteLine("      log(" & """" & "RegEx Enabled." & """" & ", nil) if use_regex()")
        WSSRuby.WriteLine("      log(" & """" & "Case Sensitive." & """" & ", nil) if case_sensitive()")
        WSSRuby.WriteLine("      log(" & """" & "Items will be excluded." & """" & ", nil) if exclude_items()")
        WSSRuby.WriteLine("      log(" & """" & "Items will be tagged." & """" & ", nil) if tag_items()")
        WSSRuby.WriteLine("      log(" & """" & "Custom Metadata field #{@@custom_metadata} will be created." & """" & ", nil) if !@@custom_metadata.nil?")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns Hash of tag=>{communication query}.")
        WSSRuby.WriteLine("    # param:  file_name - CSV file to load")
        WSSRuby.WriteLine("    def load_addresses(file_name)")
        WSSRuby.WriteLine("      queries = Hash.new{|h,k| h[k]=[]}")
        WSSRuby.WriteLine("      csv_file = load_queries(file_name)")
        WSSRuby.WriteLine("      csv_file.each do |k, a|")
        WSSRuby.WriteLine("        a.each do |v|")
        WSSRuby.WriteLine("          queries[k] << get_field_query(v)")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      return queries")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      #Returns nil or name of custom metadata field to generate.")
        WSSRuby.WriteLine("      # param:  value - the :customMetadata value from @@config  ")
        WSSRuby.WriteLine("      def load_custom_metadata(value)")
        WSSRuby.WriteLine("         value.strip! if !value.nil?")
        WSSRuby.WriteLine("         value = nil if value.empty?")
        WSSRuby.WriteLine("         return value")
        WSSRuby.WriteLine("     end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns Hash for file filtering")
        WSSRuby.WriteLine("    # params: mime  - MIME-types")
        WSSRuby.WriteLine("    #         kind  - kinds")
        WSSRuby.WriteLine("    #         ext   - file extensions")
        WSSRuby.WriteLine("    def load_filter(mime, kind, ext)")
        WSSRuby.WriteLine("      WSS.log(" & """" & "Loading file filter." & """" & ", nil)")
        WSSRuby.WriteLine("      [mime, kind, ext].each{ |f| f.each { |e| e.downcase! } } if !case_sensitive()")
        WSSRuby.WriteLine("      filter = Hash.new")
        WSSRuby.WriteLine("      filter[:mime] = mime")
        WSSRuby.WriteLine("      filter[:kind] = kind")
        WSSRuby.WriteLine("      filter[:ext] = ext")
        WSSRuby.WriteLine("      return filter")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns Hash of tag=>query from CSV.")
        WSSRuby.WriteLine("    # param:  file_name - CSV file to load")
        WSSRuby.WriteLine("    def load_queries(file_name)")
        WSSRuby.WriteLine("      log(" & """" & "Loading queries from: #{file_name}" & """" & ", nil)")
        WSSRuby.WriteLine("      queries = Hash.new{|h,k| h[k]=[]}")
        WSSRuby.WriteLine("      CSV.foreach(file_name) do |row|")
        WSSRuby.WriteLine("        tag = row[0]")
        WSSRuby.WriteLine("        query = row[1]")
        WSSRuby.WriteLine("        query = query.downcase if !case_sensitive()")
        WSSRuby.WriteLine("        query = Regexp.new( query.slice(1..-2) ) if use_regex() && query.start_with?(" & """" & "/" & """" & ") && query.end_with?(" & """" & "/" & """" & ")")
        WSSRuby.WriteLine("        queries[tag.to_sym] << query")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      return queries")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Puts log of datetime and string")
        WSSRuby.WriteLine("    # params: str - String to log")
        WSSRuby.WriteLine("    #         log - if verbose() should be overridden")
        WSSRuby.WriteLine("    def log(str, log)")
        WSSRuby.WriteLine("      puts " & """" & "WSS #{Time.now}: #{str}" & """" & " if ( log || verbose() )")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns report directory")
        WSSRuby.WriteLine("    def report_directory()")
        WSSRuby.WriteLine("      return @@report_directory")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns if per-item reports should be generated.")
        WSSRuby.WriteLine("    def report_items()")
        WSSRuby.WriteLine("      return @@config[:reportItems]")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns if per-query reports should be generated.")
        WSSRuby.WriteLine("    def report_queries()")
        WSSRuby.WriteLine("      return @@config[:reportQueries]")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns reports Hash")
        WSSRuby.WriteLine("    def reports()")
        WSSRuby.WriteLine("      return @@reports_hash")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Initializes reports")
        WSSRuby.WriteLine("    def reports_init(worker_guid)")
        WSSRuby.WriteLine("      report_dir = File.join(@@report_directory, worker_guid)")
        WSSRuby.WriteLine("      log(" & """" & "Report folder created at #{report_dir}" & """" & ", nil)")
        WSSRuby.WriteLine("      @@reports_hash[:item] = Item_Report.new(File.join(report_dir, " & """" & "items.csv" & """" & ")) if report_items()")
        WSSRuby.WriteLine("      @@reports_hash[:query] = Query_Report.new(File.join(report_dir, " & """" & "queries.csv" & """" & ")) if report_queries()")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns tagItems bool.")
        WSSRuby.WriteLine("    def tag_items()")
        WSSRuby.WriteLine("      return @@config[:tagItems]")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns tagUnique bool.")
        WSSRuby.WriteLine("    def tag_unique()")
        WSSRuby.WriteLine("      return @@config[:tagUnique]")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns useRegEx bool.")
        WSSRuby.WriteLine("    def use_regex()")
        WSSRuby.WriteLine("      return @@config[:useRegEx]")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Returns fileFiltering bool.")
        WSSRuby.WriteLine("    def verbose()")
        WSSRuby.WriteLine("      return @@config[:verbose]")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("#Class for each WorkerItem, containing its details and methods for comparing itself against the WSS it is a part of.")
        WSSRuby.WriteLine("class WSSitem")
        WSSRuby.WriteLine("  extend WSS")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  attr_accessor :container, :flags, :source_item, :worker_item")
        WSSRuby.WriteLine("  def initialize(worker_item)")
        WSSRuby.WriteLine("    @worker_item = worker_item")
        WSSRuby.WriteLine("    @source_item = worker_item.get_source_item()")
        WSSRuby.WriteLine("    WSS.log(" & """" & "Processing Item: #{@source_item.get_name}" & """" & ", true)")
        WSSRuby.WriteLine("    @container = @source_item.is_kind(" & """" & "container" & """" & ")")
        WSSRuby.WriteLine("    if !@container")
        WSSRuby.WriteLine("      @communication = @source_item.get_communication()")
        WSSRuby.WriteLine("      @text = @source_item.get_text()")
        WSSRuby.WriteLine("      @props = @source_item.get_properties()")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("    @flags = Set.new() #tags (as Symbol) to apply")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Analyzes items")
        WSSRuby.WriteLine("  def analyze()")
        WSSRuby.WriteLine("    return nil if @container")
        WSSRuby.WriteLine("    if WSS.file_filter() && !analyze_file()")
        WSSRuby.WriteLine("      WSS.log(" & """" & "Filtered out #{@source_item.get_name}" & """" & ", nil)")
        WSSRuby.WriteLine("      @worker_item.set_process_item(false) if WSS.exclude_items()")
        WSSRuby.WriteLine("      return false")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("    analyze_comm()")
        WSSRuby.WriteLine("    analyze_props()")
        WSSRuby.WriteLine("    analyze_text()")
        WSSRuby.WriteLine("    if @flags.empty?")
        WSSRuby.WriteLine("      WSS.log(" & """" & "#{@source_item.get_name} did not hit on any terms." & """" & ", nil)")
        WSSRuby.WriteLine("      if WSS.flag_unresponsive()")
        WSSRuby.WriteLine("        @flags.add(:unresponsive)")
        WSSRuby.WriteLine("      else")
        WSSRuby.WriteLine("        @worker_item.set_process_item(false) if WSS.exclude_items()")
        WSSRuby.WriteLine("        return false")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    else")
        WSSRuby.WriteLine("      WSS.log(" & """" & "#{@source_item.get_name} matched the following queries: #{@flags.to_a.join(" & """" & ";" & """" & ")}" & """" & ", nil)")
        WSSRuby.WriteLine("      @worker_item.add_tag(" & """" & "unique|#{@flags.first.to_s}" & """" & ") if ( WSS.tag_unique && @flags.size == 1 )")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("    @flags.each { |tag| @worker_item.add_tag(tag.to_s) } if WSS.tag_items()")
        WSSRuby.WriteLine("    @worker_item.add_custom_metadata(WSS.custom_metadata(), @flags.to_a.join(" & """" & "|" & """" & "), " & """" & "text" & """" & ", " & """" & "user" & """" & ") if !WSS.custom_metadata().nil?")
        WSSRuby.WriteLine("    WSS.reports[:item].add(self) if WSS.report_items")
        WSSRuby.WriteLine("    WSS.reports[:query].add(self) if WSS.report_queries")
        WSSRuby.WriteLine("    return true")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Analyze communication fields")
        WSSRuby.WriteLine("  #Returns nil if @source_item is not a communication, T/F if addresses matched.")
        WSSRuby.WriteLine("  # @flags is updated for matching items")
        WSSRuby.WriteLine("  def analyze_comm()")
        WSSRuby.WriteLine("    return nil if @communication.nil?")
        WSSRuby.WriteLine("    tags = WSS.check_comm(@communication)")
        WSSRuby.WriteLine("    return false if tags.empty?")
        WSSRuby.WriteLine("    @flags = @flags.merge(tags)")
        WSSRuby.WriteLine("    return true")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Analyze MIME-type, kind, and extension.")
        WSSRuby.WriteLine("  #Returns nil if filter is disabled, T/F if @source_item matches.")
        WSSRuby.WriteLine("  def analyze_file()")
        WSSRuby.WriteLine("    return nil if !WSS.file_filter()")
        WSSRuby.WriteLine("    mime = @source_item.get_type().get_name()")
        WSSRuby.WriteLine("    kind = @source_item.get_kind().get_name()")
        WSSRuby.WriteLine("    ext = File.extname(@source_item.get_name)")
        WSSRuby.WriteLine("    return WSS.check_file(mime, kind, ext)")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Analyze properties")
        WSSRuby.WriteLine("  #Returns nil if @properties aren't set, T/F if @properties matched.")
        WSSRuby.WriteLine("  # @flags is updated for matching items")
        WSSRuby.WriteLine("  def analyze_props()")
        WSSRuby.WriteLine("    return nil if @props.nil?")
        WSSRuby.WriteLine("    tags = WSS.check_props(@props)")
        WSSRuby.WriteLine("    return false if tags.empty?")
        WSSRuby.WriteLine("    @flags = @flags.merge(tags)")
        WSSRuby.WriteLine("    return true")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Analyze text content")
        WSSRuby.WriteLine("  #Returns nil if @source_item had empty or no text. T/F if it matched.")
        WSSRuby.WriteLine("  # @flags is updated for matching items")
        WSSRuby.WriteLine("  def analyze_text()")
        WSSRuby.WriteLine("    if @text.is_available()")
        WSSRuby.WriteLine("      begin #to handle OutOfMemoryError when text is too big.")
        WSSRuby.WriteLine("        text = @text.to_string.strip")
        WSSRuby.WriteLine("        if text.empty?")
        WSSRuby.WriteLine("          #Text was empty/OCR candidate.")
        WSSRuby.WriteLine("          WSS.log(" & """" & "Empty text found for #{@source_item.get_name}" & """" & ", nil)")
        WSSRuby.WriteLine("          return nil")
        WSSRuby.WriteLine("        else")
        WSSRuby.WriteLine("          tags = WSS.check_text(text)")
        WSSRuby.WriteLine("          return false if tags.empty?")
        WSSRuby.WriteLine("          @flags = @flags.merge(tags)")
        WSSRuby.WriteLine("          return true")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      rescue java.lang.OutOfMemoryError")
        WSSRuby.WriteLine("        WSS.log(" & """" & "There was not enough memory to read text of #{@source_item.get_name}" & """" & ", true)")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    else")
        WSSRuby.WriteLine("      #the item does not have text")
        WSSRuby.WriteLine("      WSS.log(" & """" & "No text found for #{@source_item.get_name}" & """" & ", false)")
        WSSRuby.WriteLine("      return nil")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("#Loads settings from JSON.")
        WSSRuby.WriteLine("def nuixWorkerItemCallbackInit()")
        WSSRuby.WriteLine("  config_json = ENV_JAVA[" & """" & "nuix.wss.config" & """" & "]")
        WSSRuby.WriteLine("  if File.exist?(config_json)")
        WSSRuby.WriteLine("    WSS.load(config_json)")
        WSSRuby.WriteLine("  else")
        WSSRuby.WriteLine("    puts " & """" & "No configuration file selected." & """")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("#Creates per-worker reports if required.")
        WSSRuby.WriteLine("#Creates new WSSitem from item")
        WSSRuby.WriteLine("#Analyzes items")
        WSSRuby.WriteLine("def nuixWorkerItemCallback(worker_item)")
        WSSRuby.WriteLine("  WSS.reports_init(worker_item.get_worker_guid()) if ( WSS.reports.empty? && !WSS.report_directory().nil? )")
        WSSRuby.WriteLine("  wss_item = WSSitem.new(worker_item)")
        WSSRuby.WriteLine("  wss_item.analyze()")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("#Generates per-query report.")
        WSSRuby.WriteLine("def nuixWorkerItemCallbackClose")
        WSSRuby.WriteLine("  WSS.reports[:query].write() if WSS.report_queries")
        WSSRuby.WriteLine("end")

        WSSRuby.Close()

        blnBuildWSSRubyScript = True

    End Function

    Public Function blnBuildProcessingRubyScript(ByVal sProcessingRuby As String) As Boolean
        blnBuildProcessingRubyScript = False

        Dim ProcessingRuby As StreamWriter

        ProcessingRuby = New StreamWriter(sProcessingRuby)

        ProcessingRuby.WriteLine("require 'java'")
        ProcessingRuby.WriteLine("require 'thread'")
        ProcessingRuby.WriteLine("require 'json'")
        ProcessingRuby.WriteLine("require 'csv'")
        ProcessingRuby.WriteLine("")
        ProcessingRuby.WriteLine("module Chooser")
        ProcessingRuby.WriteLine("  java_import javax.swing.JFileChooser")
        ProcessingRuby.WriteLine("  java_import javax.swing.filechooser.FileNameExtensionFilter")
        ProcessingRuby.WriteLine("  class << self")
        ProcessingRuby.WriteLine("    def dialog(initial_directory=nil, title=nil, filter=nil)")
        ProcessingRuby.WriteLine("      chooser = JFileChooser.new")
        ProcessingRuby.WriteLine("      if filter == JFileChooser::DIRECTORIES_ONLY")
        ProcessingRuby.WriteLine("        chooser.set_file_selection_mode(JFileChooser::DIRECTORIES_ONLY)")
        ProcessingRuby.WriteLine("      elsif !filter.nil?")
        ProcessingRuby.WriteLine("        chooser.set_file_filter(filter)")
        ProcessingRuby.WriteLine("      end")
        ProcessingRuby.WriteLine("      chooser.set_current_directory(java.io.File.new(initial_directory)) if !initial_directory.nil?")
        ProcessingRuby.WriteLine("      chooser.set_dialog_title(title) if !title.nil?")
        ProcessingRuby.WriteLine("      return chooser.get_selected_file.get_absolute_path if (chooser.show_open_dialog(nil) == JFileChooser::APPROVE_OPTION)")
        ProcessingRuby.WriteLine("      return nil")
        ProcessingRuby.WriteLine("    end")
        ProcessingRuby.WriteLine("")
        ProcessingRuby.WriteLine("    def dir(initial_directory=nil, title=nil)")
        ProcessingRuby.WriteLine("      return dialog(initial_directory, title, JFileChooser::DIRECTORIES_ONLY)")
        ProcessingRuby.WriteLine("    end")
        ProcessingRuby.WriteLine("")
        ProcessingRuby.WriteLine("    def ext(extension, initial_directory=nil, title=nil)")
        ProcessingRuby.WriteLine("      return dialog(initial_directory, title, FileNameExtensionFilter.new(extension, extension))")
        ProcessingRuby.WriteLine("    end")
        ProcessingRuby.WriteLine("  end")
        ProcessingRuby.WriteLine("end")
        ProcessingRuby.WriteLine("")
        ProcessingRuby.WriteLine("case_dir = ENV_JAVA[" & """" & "case" & """" & "] || Chooser.dir(nil, " & """" & "Select Case Directory" & """" & ")")
        ProcessingRuby.WriteLine("case_name = ENV_JAVA[" & """" & "case.name" & """" & "] || Time.now.strftime(" & """" & "%Y%m%d%H%M%S" & """" & ")")
        ProcessingRuby.WriteLine("case_description = ENV_JAVA[" & """" & "case.description" & """" & "] || " & """" & "custom" & """")
        ProcessingRuby.WriteLine("case_investigator = ENV_JAVA[" & """" & "case.investigator" & """" & "] || " & """" & "script" & """")
        ProcessingRuby.WriteLine("source_dir = ENV_JAVA[" & """" & """" & "source" & """" & "] || Chooser.dialog(nil," & """" & "Select Source Evidence" & """" & ", nil)")
        ProcessingRuby.WriteLine("parallel_json = ENV_JAVA[" & """" & "parallel" & """" & "] || Chooser.ext(" & """" & "json" & """" & ", Dir.pwd, " & """" & "Select Parallel Processing Settings" & """" & ")")
        ProcessingRuby.WriteLine("processor_json = ENV_JAVA[" & """" & "processing" & """" & "] || Chooser.ext(" & """" & "json" & """" & ", Dir.pwd, " & """" & "Select Data Processing Settings" & """" & ")")
        ProcessingRuby.WriteLine("worker_side_script = ENV_JAVA[" & """" & "wss" & """" & "] ||  Chooser.ext(" & """" & "rb" & """" & ", Dir.pwd, " & """" & "Select Worker Side Script" & """" & ")")
        ProcessingRuby.WriteLine("ENV_JAVA[" & """" & "nuix.wss.config" & """" & "] = Chooser.ext(" & """" & "json" & """" & ", Dir.pwd, " & """" & "Select WSS Settings" & """" & ") if File.exist?(worker_side_script) && !File.exist?(ENV_JAVA[" & """" & "nuix.wss.config" & """" & "])")
        ProcessingRuby.WriteLine("reporting = false")
        ProcessingRuby.WriteLine("if !worker_side_script.nil? && !ENV_JAVA[" & """" & "nuix.wss.config" & """" & "].nil? && File.exist?(ENV_JAVA[" & """" & "nuix.wss.config" & """" & "])")
        ProcessingRuby.WriteLine("  wss_settings = JSON.parse(IO.read(ENV_JAVA[" & """" & "nuix.wss.config" & """" & "]))")
        ProcessingRuby.WriteLine("  if (wss_settings[" & """" & "reportItems" & """" & "] == true) || (wss_settings[" & """" & "reportQueries" & """" & "] == true)")
        ProcessingRuby.WriteLine("    ENV_JAVA[" & """" & "nuix.wss.reports" & """" & "] = Chooser.dir(wss_settings[" & """" & "reportDirectory" & """" & "], " & """" & "Select Reporting Directory" & """" & ") if ENV_JAVA[" & """" & "nuix.wss.reports" & """" & "].nil?")
        ProcessingRuby.WriteLine("    if !ENV_JAVA[" & """" & "nuix.wss.reports" & """" & "].nil?")
        ProcessingRuby.WriteLine("      reporting = true")
        ProcessingRuby.WriteLine("    end")
        ProcessingRuby.WriteLine("  end")
        ProcessingRuby.WriteLine("end")
        ProcessingRuby.WriteLine("")
        ProcessingRuby.WriteLine("begin")
        ProcessingRuby.WriteLine("  if !case_dir.nil?")
        ProcessingRuby.WriteLine("    if !source_dir.nil?")
        ProcessingRuby.WriteLine("      if File.exist?(processor_json)")
        ProcessingRuby.WriteLine("        #Create case")
        ProcessingRuby.WriteLine("        case_settings = {" & """" & "compound" & """" & "=> false, " & """" & "name" & """" & " => case_name, " & """" & "description" & """" & " => case_description, " & """" & "investigator" & """" & " => case_investigator}")
        ProcessingRuby.WriteLine("        case_location = File.join(case_dir, case_name)")
        ProcessingRuby.WriteLine("        case1 = utilities.get_case_factory().create(case_location, case_settings)")
        ProcessingRuby.WriteLine("        #Set processor")
        ProcessingRuby.WriteLine("        processor = case1.get_processor()")
        ProcessingRuby.WriteLine("        processing_settings = JSON.parse(IO.read(processor_json))")
        ProcessingRuby.WriteLine("        #Add Worker-Side Script")
        ProcessingRuby.WriteLine("        processing_settings[" & """" & "workerItemCallback" & """" & "] = " & """" & "ruby:#{IO.read(worker_side_script)}" & """" & " if File.exist?(worker_side_script)")
        ProcessingRuby.WriteLine("        processor.set_processing_settings(processing_settings)")
        ProcessingRuby.WriteLine("        #Load parallel settings")
        ProcessingRuby.WriteLine("        parallel_settings = {}  #runs with default parallel settings if none provided")
        ProcessingRuby.WriteLine("        parallel_settings = JSON.parse(IO.read(parallel_json)) if !parallel_json.nil? && File.exist?(parallel_json)")
        ProcessingRuby.WriteLine("        processor.set_parallel_processing_settings(parallel_settings)")
        ProcessingRuby.WriteLine("        #Add Evidence Container")
        ProcessingRuby.WriteLine("        evidence_container = processor.new_evidence_container(" & """" & "container" & """" & ")")
        ProcessingRuby.WriteLine("        evidence_container.add_file(source_dir)")
        ProcessingRuby.WriteLine("        evidence_container.save()")
        ProcessingRuby.WriteLine("        #Set progress callback")
        ProcessingRuby.WriteLine("        last_progress = Time.now")
        ProcessingRuby.WriteLine("        semaphore = Mutex.new")
        ProcessingRuby.WriteLine("        processor.when_progress_updated do |progress|")
        ProcessingRuby.WriteLine("          semaphore.synchronize {")
        ProcessingRuby.WriteLine("            # Progress message every 60 seconds")
        ProcessingRuby.WriteLine("            if (Time.now - last_progress) > 60")
        ProcessingRuby.WriteLine("              last_progress = Time.now")
        ProcessingRuby.WriteLine("              puts " & """" & "#{Time.now}: Processed #{progress.get_current_size} of #{progress.get_total_size}" & """")
        ProcessingRuby.WriteLine("            end")
        ProcessingRuby.WriteLine("          }")
        ProcessingRuby.WriteLine("        end")
        ProcessingRuby.WriteLine("")
        ProcessingRuby.WriteLine("        #Set cleanup")
        ProcessingRuby.WriteLine("        processor.when_cleaning_up do |cleaning|")
        ProcessingRuby.WriteLine("          puts " & """" & "#{Time.now}: Cleaning up." & """")
        ProcessingRuby.WriteLine("          Outputter.close_files()")
        ProcessingRuby.WriteLine("        end")
        ProcessingRuby.WriteLine("")
        ProcessingRuby.WriteLine("        #Start processing")
        ProcessingRuby.WriteLine("        puts " & """" & "#{Time.now}: Starting processing..." & """")
        ProcessingRuby.WriteLine("        processor.process()")
        ProcessingRuby.WriteLine("")
        ProcessingRuby.WriteLine("        case1.close")
        ProcessingRuby.WriteLine("        puts " & """" & "#{Time.now}: Processing complete." & """")
        ProcessingRuby.WriteLine("")
        ProcessingRuby.WriteLine("        if reporting")
        ProcessingRuby.WriteLine("          Dir.glob(" & """" & "#{ENV_JAVA[" & """" & "nuix.wss.reports" & """" & "]}#{File::SEPARATOR}*#{File::SEPARATOR}" & """" & ").each do |folder|")
        ProcessingRuby.WriteLine("            #Get each report")
        ProcessingRuby.WriteLine("            Dir.glob(File.join(folder, " & """" & "*.csv" & """" & ")).each do |csv_file|")
        ProcessingRuby.WriteLine("              csv = CSV.read(csv_file, :headers => true, :header_converters => :symbol)")
        ProcessingRuby.WriteLine("              type = File.basename(csv_file, " & """" & ".csv" & """" & ")")
        ProcessingRuby.WriteLine("              case type")
        ProcessingRuby.WriteLine("              when " & """" & "items" & """")
        ProcessingRuby.WriteLine("                #Append to per-item report")
        ProcessingRuby.WriteLine("                item_report = File.join(ENV_JAVA[" & """" & "nuix.wss.reports" & """" & "], " & """" & "items-summary.csv" & """" & ")")
        ProcessingRuby.WriteLine("                #Create the report if it doesn't already exist")
        ProcessingRuby.WriteLine("                CSV.open(item_report, " & """" & "w" & """" & ") { |c| c << csv.headers } if !File.exist?(item_report)")
        ProcessingRuby.WriteLine("                CSV.open(item_report, " & """" & "a" & """" & ") { |c| csv.each{ |r| c << r } }")
        ProcessingRuby.WriteLine("              when " & """" & "queries" & """")
        ProcessingRuby.WriteLine("                #Merge with summary @query_report Hash")
        ProcessingRuby.WriteLine("                #Create the hash if it doesn't already exist")
        ProcessingRuby.WriteLine("                @query_report = Hash.new(0) if @query_report.nil?")
        ProcessingRuby.WriteLine("                csv.each { |row| @query_report[row[0]] += row[1].to_i }")
        ProcessingRuby.WriteLine("              else")
        ProcessingRuby.WriteLine("                puts " & """" & "Unknown Report Type: #{csv_file}" & """")
        ProcessingRuby.WriteLine("              end")
        ProcessingRuby.WriteLine("            end")
        ProcessingRuby.WriteLine("          end")
        ProcessingRuby.WriteLine("          if !@query_report.nil?")
        ProcessingRuby.WriteLine("            #If there is a query_report, write it")
        ProcessingRuby.WriteLine("            CSV.open(File.join(ENV_JAVA[" & """" & "nuix.wss.reports" & """" & "], " & """" & "queries-summary.csv" & """" & "), " & """" & "w" & """" & ") do |c|")
        ProcessingRuby.WriteLine("              c << [" & """" & "Category" & """" & ", " & """" & "Count" & """" & "]")
        ProcessingRuby.WriteLine("              @query_report.each { |row| c << row }")
        ProcessingRuby.WriteLine("            end")
        ProcessingRuby.WriteLine("          end")
        ProcessingRuby.WriteLine("        end")
        ProcessingRuby.WriteLine("      else")
        ProcessingRuby.WriteLine("        puts " & """" & "No processing settings selected." & """")
        ProcessingRuby.WriteLine("      end")
        ProcessingRuby.WriteLine("    else")
        ProcessingRuby.WriteLine("      puts " & """" & "No source data found." & """")
        ProcessingRuby.WriteLine("    end")
        ProcessingRuby.WriteLine("  else")
        ProcessingRuby.WriteLine("    puts " & """" & "No case directory selected." & """")
        ProcessingRuby.WriteLine("  end")
        ProcessingRuby.WriteLine("end")
        blnBuildProcessingRubyScript = False
    End Function

    Public Function blnBuildReportingRubyScript(ByVal sReportingRubyScriptFileName As String) As Boolean
        blnBuildReportingRubyScript = False

        Dim ReportingRuby As StreamWriter
        ReportingRuby = New StreamWriter(sReportingRubyScriptFileName)

        ReportingRuby.WriteLine("require 'java'")
        ReportingRuby.WriteLine("require 'csv'")
        ReportingRuby.WriteLine("")
        ReportingRuby.WriteLine("module Chooser")
        ReportingRuby.WriteLine("  java_import javax.swing.JFileChooser")
        ReportingRuby.WriteLine("  def Chooser::dir(initial_directory=nil, title=nil)")
        ReportingRuby.WriteLine("    dir_chooser = JFileChooser.new")
        ReportingRuby.WriteLine("    dir_chooser.setFileSelectionMode(JFileChooser::DIRECTORIES_ONLY)")
        ReportingRuby.WriteLine("    dir_chooser.setCurrentDirectory(java.io.File.new(initial_directory)) if !initial_directory.nil?")
        ReportingRuby.WriteLine("    dir_chooser.setDialogTitle(title) if !title.nil?")
        ReportingRuby.WriteLine("    return dir_chooser.getSelectedFile.getAbsolutePath if (dir_chooser.showOpenDialog(nil) == JFileChooser::APPROVE_OPTION)")
        ReportingRuby.WriteLine("    return nil")
        ReportingRuby.WriteLine("  end")
        ReportingRuby.WriteLine("end")
        ReportingRuby.WriteLine("")
        ReportingRuby.WriteLine("begin")
        ReportingRuby.WriteLine("  #Select Reporting Directory")
        ReportingRuby.WriteLine("  dir = Chooser.dir(Dir.pwd, " & """" & "Select Report Directory" & """" & ")")
        ReportingRuby.WriteLine("  if !dir.nil?")
        ReportingRuby.WriteLine("    #Get each per-worker folder")
        ReportingRuby.WriteLine("    Dir.glob(" & """" & "#{dir}#{File::SEPARATOR}*#{File::SEPARATOR}" & """" & ").each do |folder|")
        ReportingRuby.WriteLine("      #Get each report")
        ReportingRuby.WriteLine("      Dir.glob(File.join(folder, " & """" & "*.csv" & """" & ")).each do |csv_file|")
        ReportingRuby.WriteLine("        csv = CSV.read(csv_file, :headers => true, :header_converters => :symbol)")
        ReportingRuby.WriteLine("        type = File.basename(csv_file, " & """" & ".csv" & """" & ")")
        ReportingRuby.WriteLine("        case type")
        ReportingRuby.WriteLine("        when " & """" & "items" & """")
        ReportingRuby.WriteLine("          #Append to per-item report")
        ReportingRuby.WriteLine("          item_report = File.join(dir, " & """" & "items-summary.csv" & """" & ")")
        ReportingRuby.WriteLine("          #Create the report if it doesn't already exist")
        ReportingRuby.WriteLine("          CSV.open(item_report, " & """" & "w" & """" & ") { |c| c << csv.headers } if !File.exist?(item_report)")
        ReportingRuby.WriteLine("          CSV.open(item_report, " & """" & "a" & """" & ") { |c| csv.each{ |r| c << r } }")
        ReportingRuby.WriteLine("        when " & """" & "queries" & """")
        ReportingRuby.WriteLine("          #Merge with summary @query_report Hash")
        ReportingRuby.WriteLine("          #Create the hash if it doesn't already exist")
        ReportingRuby.WriteLine("          @query_report = Hash.new(0) if @query_report.nil?")
        ReportingRuby.WriteLine("          csv.each { |row| @query_report[row[0]] += row[1].to_i }")
        ReportingRuby.WriteLine("        else")
        ReportingRuby.WriteLine("          puts " & """" & "Unknown Report Type: #{csv_file}" & """")
        ReportingRuby.WriteLine("        end")
        ReportingRuby.WriteLine("      end")
        ReportingRuby.WriteLine("    end")
        ReportingRuby.WriteLine("    if !@query_report.nil?")
        ReportingRuby.WriteLine("      #If there is a query_report, write it")
        ReportingRuby.WriteLine("      CSV.open(File.join(dir, " & """" & "queries-summary.csv" & """" & "), " & """" & "w" & """" & ") do |c|")
        ReportingRuby.WriteLine("        c << [" & """" & "Category" & """" & ", " & """" & "Count" & """" & "]")
        ReportingRuby.WriteLine("        @query_report.each { |row| c << row }")
        ReportingRuby.WriteLine("      end")
        ReportingRuby.WriteLine("    end")
        ReportingRuby.WriteLine("  end")
        ReportingRuby.WriteLine("end")

        blnBuildReportingRubyScript = True
    End Function

    Public Function blnBuildUpdatedWSSRubyScript(ByVal sRubyScriptFileName As String, ByVal sNuixFilesDir As String) As Boolean
        blnBuildUpdatedWSSRubyScript = False

        Dim WSSRuby As StreamWriter

        WSSRuby = New StreamWriter(sRubyScriptFileName)

        WSSRuby.WriteLine("ENV_JAVA[" & """" & "nuix.wss.config" & """" & "] = " & """" & "E:/wss/wss.json" & """")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("#Load settings from JSON for WSS.")
        WSSRuby.WriteLine("def nuixWorkerItemCallbackInit()")
        WSSRuby.WriteLine("  WSS.load( ENV_JAVA[" & """" & "nuix.wss.config" & """" & "] )")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("#Process item.")
        WSSRuby.WriteLine("def nuixWorkerItemCallback(worker_item)")
        WSSRuby.WriteLine("  WSSitem.new(worker_item)")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("#Create per-query report.")
        WSSRuby.WriteLine("def nuixWorkerItemCallbackClose")
        WSSRuby.WriteLine("  WSS.close()")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("#Module for the Worker-Side Script")
        WSSRuby.WriteLine("#contains:  @@settings  - the settings")
        WSSRuby.WriteLine("#           @@queries   - WSSqueries object; Compare items against queries using .compare(item)")
        WSSRuby.WriteLine("#           @@reports   - WSSreports object; Add items to reports using .add(item)")
        WSSRuby.WriteLine("module WSS")
        WSSRuby.WriteLine("  require 'json'")
        WSSRuby.WriteLine("  require 'csv'")
        WSSRuby.WriteLine("  require 'set'")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  @@settings = {}")
        WSSRuby.WriteLine("  #settings[:verbose] = Bool")
        WSSRuby.WriteLine("  #settings[:case_sensitive] = Bool")
        WSSRuby.WriteLine("  #settings[:use_regex] = Bool")
        WSSRuby.WriteLine("  #settings[:exclude_items] = Bool")
        WSSRuby.WriteLine("  #settings[:flag_unresponsive] = Bool")
        WSSRuby.WriteLine("  #settings[:custom_metadata] = String")
        WSSRuby.WriteLine("  #settings[:category] = String")
        WSSRuby.WriteLine("  #settings[:tag_items] = Bool")
        WSSRuby.WriteLine("  #settings[:tag_unique] = Bool")
        WSSRuby.WriteLine("  #settings[:file_filter] = Hash of :mime/:kind/:ext => Array of options")
        WSSRuby.WriteLine("  #settings[:address_queries] = CSV")
        WSSRuby.WriteLine("  #settings[:properties_queries] = CSV")
        WSSRuby.WriteLine("  #settings[:text_queries] = CSV")
        WSSRuby.WriteLine("  #settings[:report_items] = Bool")
        WSSRuby.WriteLine("  #settings[:report_queries] = Bool")
        WSSRuby.WriteLine("  #settings[:report_directory] = String")
        WSSRuby.WriteLine("  @@queries = nil #WSSqueries object")
        WSSRuby.WriteLine("  @@reports = nil #WSSreports object")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  class << self")
        WSSRuby.WriteLine("    #Initializes WSS")
        WSSRuby.WriteLine("    # param:  json  - location of JSON file")
        WSSRuby.WriteLine("    def load(json)")
        WSSRuby.WriteLine("      if json.nil? || !File.exist?(json)")
        WSSRuby.WriteLine("        Log.out(" & """" & "No Configuration File!" & """" & ", true)")
        WSSRuby.WriteLine("      else")
        WSSRuby.WriteLine("        Log.out(" & """" & "Loading settings from #{json}" & """" & ", true)")
        WSSRuby.WriteLine("        config = JSON.parse(File.read(json), :symbolize_names => true)")
        WSSRuby.WriteLine("        @@settings[:verbose] = config[:verbose]")
        WSSRuby.WriteLine("        Log.set(@@settings[:verbose])")
        WSSRuby.WriteLine("        Log.out(" & """" & "Config: #{JSON.pretty_generate(config)}" & """" & ", nil)")
        WSSRuby.WriteLine("        @@settings[:case_sensitive] = config[:caseSensitive]")
        WSSRuby.WriteLine("        @@settings[:use_regex] = config[:useRegEx]")
        WSSRuby.WriteLine("        @@settings[:exclude_items] = config[:excludeItems]")
        WSSRuby.WriteLine("        @@settings[:flag_unresponsive] = config[:flagUnresponsive]")
        WSSRuby.WriteLine("        if !config[:customMetadata].nil? && !config[:customMetadata].empty?")
        WSSRuby.WriteLine("          @@settings[:custom_metadata] = config[:customMetadata]")
        WSSRuby.WriteLine("          @@settings[:category] = config[:categoryProperty] if !config[:categoryProperty].nil? && !config[:categoryProperty].empty?")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        @@settings[:tag_items] = config[:tagItems]")
        WSSRuby.WriteLine("        @@settings[:tag_unique] = config[:tagUnique]")
        WSSRuby.WriteLine("        @@settings[:file_filter] = {)")
        WSSRuby.WriteLine("          :mime => config[:filterMimeTypes],")
        WSSRuby.WriteLine("          :kind => config[:filterKinds],")
        WSSRuby.WriteLine("          :ext => config[:filterFileExt]")
        WSSRuby.WriteLine("        }.reject{|k, v| v.nil? || v.empty?} if config[:fileFiltering]")
        WSSRuby.WriteLine("        if !@@settings[:case_sensitive]")
        WSSRuby.WriteLine("          @@settings[:category].downcase! if !@@settings[:category].nil?")
        WSSRuby.WriteLine("          @@settings[:file_filter].each_value{|a| a.map!{|v| v.downcase}} if !@@settings[:file_filter].nil?")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        @@settings[:address_queries] = config[:communicationCSV]")
        WSSRuby.WriteLine("        @@settings[:properties_queries] = config[:propertiesCSV]")
        WSSRuby.WriteLine("        @@settings[:text_queries] = config[:textQueriesCSV]")
        WSSRuby.WriteLine("        @@settings[:report_items] = config[:reportItems]")
        WSSRuby.WriteLine("        @@settings[:report_queries] = config[:reportQueries]")
        WSSRuby.WriteLine("        if @@settings[:report_items] || @@settings[:report_queries]")
        WSSRuby.WriteLine("          @@settings[:report_directory] = config[:reportDirectory]")
        WSSRuby.WriteLine("          @@reports = WSSreports.new( @@settings )")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        Log.out(" & """" & "Settings: #{JSON.pretty_generate(@@settings)}" & """" & ", true)")
        WSSRuby.WriteLine("        Log.out(" & """" & "Loading CSV queries" & """" & ", true)")
        WSSRuby.WriteLine("        @@queries = WSSqueries.new(@@settings)")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Write Query Report")
        WSSRuby.WriteLine("    def close()")
        WSSRuby.WriteLine("      Log.out(" & """" & "Closing." & """" & ", nil)")
        WSSRuby.WriteLine("      @@reports.close() if !@@reports.nil?")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Class for handling the queries and comparing items against them.")
        WSSRuby.WriteLine("  # contains: @run - Hash containing Queries.")
        WSSRuby.WriteLine("  # WSSqueries.compare(item) does Queries.compare(item) for each in @run.")
        WSSRuby.WriteLine("  class WSSqueries")
        WSSRuby.WriteLine("    #@run contains non-empty Queries.")
        WSSRuby.WriteLine("    # params: settings  - settings Hash from WSS")
        WSSRuby.WriteLine("    def initialize(settings)")
        WSSRuby.WriteLine("      @run = {)")
        WSSRuby.WriteLine("        :address => Address_Queries.new(settings, :address_queries),")
        WSSRuby.WriteLine("        :property => Property_Queries.new(settings, :properties_queries),")
        WSSRuby.WriteLine("        :text => Queries.new(settings, :text_queries)")
        WSSRuby.WriteLine("      }.reject{|k, q| q.queries.empty?} #reject empty queries")
        WSSRuby.WriteLine("      Log.out(" & """" & "Loaded queries for: #{@run.keys.join(" & """" & ", " & """" & ")}" & """" & ", true)")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Compares the item against the queries.")
        WSSRuby.WriteLine("    #Returns number of flags the item has matched.")
        WSSRuby.WriteLine("    # param:  item  - a WSSitem")
        WSSRuby.WriteLine("    def compare(item)")
        WSSRuby.WriteLine("      @run.each_value { |q| q.compare(item) }")
        WSSRuby.WriteLine("      return item.flags.size")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Class for Queries. Queries.compare(item) compares the item against the queries.")
        WSSRuby.WriteLine("    # contains: @settings - settings Hash from WSS")
        WSSRuby.WriteLine("    #           @queries  - Hash of queries to run.")
        WSSRuby.WriteLine("    class Queries")
        WSSRuby.WriteLine("      attr_reader :queries")
        WSSRuby.WriteLine("      #@queries is a Hash of :tag => Array of patterns")
        WSSRuby.WriteLine("      # params: settings  - settings Hash from WSS")
        WSSRuby.WriteLine("      #         type      - key for file to load")
        WSSRuby.WriteLine("      def initialize(settings, type)")
        WSSRuby.WriteLine("        @settings = settings")
        WSSRuby.WriteLine("        @queries = load_csv(settings[type])")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      #Returns Hash of tag => Array of query from CSV.")
        WSSRuby.WriteLine("      # param:  file_name - CSV file to load, formatted as: tag, query")
        WSSRuby.WriteLine("      def load_csv(file_name)")
        WSSRuby.WriteLine("        h = Hash.new{|h,k| h[k]=[]}")
        WSSRuby.WriteLine("        if !file_name.nil? && File.exist?(file_name)")
        WSSRuby.WriteLine("          CSV.foreach(file_name) do |row|")
        WSSRuby.WriteLine("            tag = row[0]")
        WSSRuby.WriteLine("            query = row[1]")
        WSSRuby.WriteLine("            query = query.downcase if !@settings[:case_sensitive]")
        WSSRuby.WriteLine("            query = Regexp.new( query.slice(1..-2) ) if @settings[:use_regex] && query.start_with?(" & """" & "/" & """" & ") && query.end_with?(" & """" & "/" & """" & ")")
        WSSRuby.WriteLine("            h[tag.to_sym] << query")
        WSSRuby.WriteLine("          end")
        WSSRuby.WriteLine("          Log.out(" & """" & "Loaded #{file_name}" & """" & ", true)")
        WSSRuby.WriteLine("        else")
        WSSRuby.WriteLine("          Log.out(" & """" & "No file found at: #{file_name}" & """" & ", true)")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        return h")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      #Returns String of item's text to be queried.")
        WSSRuby.WriteLine("      # param:  item  - a WSSitem")
        WSSRuby.WriteLine("      def cache(item)")
        WSSRuby.WriteLine("        text = item.source_item.get_text().to_string")
        WSSRuby.WriteLine("        if text.nil?")
        WSSRuby.WriteLine("          #the item does not have text")
        WSSRuby.WriteLine("          Log.out(" & """" & "No text found for #{item.source_item.get_name}" & """" & ", nil)")
        WSSRuby.WriteLine("        else")
        WSSRuby.WriteLine("          text.strip!")
        WSSRuby.WriteLine("          if text.empty?")
        WSSRuby.WriteLine("            #Text was empty/OCR candidate.")
        WSSRuby.WriteLine("            Log.out(" & """" & "Empty text found for #{item.source_item.get_name}" & """" & ", nil)")
        WSSRuby.WriteLine("            text = nil")
        WSSRuby.WriteLine("          else")
        WSSRuby.WriteLine("            text.downcase! if !@settings[:case_sensitive]")
        WSSRuby.WriteLine("          end")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        return text")
        WSSRuby.WriteLine("      rescue java.lang.OutOfMemoryError")
        WSSRuby.WriteLine("        Log.out(" & """" & "ERROR - There was not enough memory to read text of #{item.source_item.get_name}" & """" & ", true)")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      #Compares queries against item's text and updates item's flags.")
        WSSRuby.WriteLine("      # param:  item  - a WSSitem")
        WSSRuby.WriteLine("      def compare(item)")
        WSSRuby.WriteLine("        search_data = cache(item)")
        WSSRuby.WriteLine("        if !search_data.nil?")
        WSSRuby.WriteLine("          @queries.each do |tag, query_array|")
        WSSRuby.WriteLine("            next if item.flags.include?(tag)  #item has already hit for the tag")
        WSSRuby.WriteLine("            #compare text against queries from array until there is a match")
        WSSRuby.WriteLine("            item.flags.add(tag) if query_array.any?{|query| !search_data.match(query).nil?}")
        WSSRuby.WriteLine("          end")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Queries Class for queries against Address/Communication fields.")
        WSSRuby.WriteLine("    #@queries is a Hash of :tag => Hash of :field => Array of patterns.")
        WSSRuby.WriteLine("    # query strings are converted to :field => [pattern] from " & """" & "from:foo@bar.com" & """" & " or " & """" & "from-mail-domain:bar.com.")
        WSSRuby.WriteLine("    class Address_Queries < Queries")
        WSSRuby.WriteLine("      # params: settings  - settings Hash from WSS")
        WSSRuby.WriteLine("      #         type      - key for file to load")
        WSSRuby.WriteLine("      def initialize(settings, type)")
        WSSRuby.WriteLine("        super(settings, type)")
        WSSRuby.WriteLine("        address_queries = Hash.new{|h,k| h[k] = Hash.new{|h2,k2| h2[k2] = []}}")
        WSSRuby.WriteLine("        @queries.each do |tag, terms|")
        WSSRuby.WriteLine("          #convert query_string to :field => Array of patterns")
        WSSRuby.WriteLine("          terms.each do |query_string|")
        WSSRuby.WriteLine("            field = nil")
        WSSRuby.WriteLine("            pattern = nil")
        WSSRuby.WriteLine("            query = query_string.to_s.split(" & """" & ":" & """" & ")")
        WSSRuby.WriteLine("            if query.size == 2")
        WSSRuby.WriteLine("              #contained " & """" & "field:pattern" & """" & ")")
        WSSRuby.WriteLine("              if query[0].start_with?(" & """" & "from" & """" & ")")
        WSSRuby.WriteLine("                field = :from")
        WSSRuby.WriteLine("              elsif query[0].start_with?(" & """" & "to" & """" & ")")
        WSSRuby.WriteLine("                field = :to")
        WSSRuby.WriteLine("              elsif query[0].start_with?(" & """" & "cc" & """" & ")")
        WSSRuby.WriteLine("                field = :cc")
        WSSRuby.WriteLine("              elsif query[0].start_with?(" & """" & "bcc" & """" & ")")
        WSSRuby.WriteLine("                field = :bcc")
        WSSRuby.WriteLine("              elsif query[0].start_with?(" & """" & "recipient" & """" & ")")
        WSSRuby.WriteLine("                field = :recipient")
        WSSRuby.WriteLine("              elsif query[0].start_with?(" & """" & "address" & """" & ")")
        WSSRuby.WriteLine("                field = :all")
        WSSRuby.WriteLine("              else")
        WSSRuby.WriteLine("                #Trying to search unknown field")
        WSSRuby.WriteLine("                Log.out(" & """" & "Unknown communication field: #{query[0]}" & """" & ", true)")
        WSSRuby.WriteLine("                field = :all")
        WSSRuby.WriteLine("              end")
        WSSRuby.WriteLine("              #trim quotation marks from query if required")
        WSSRuby.WriteLine("              query[1] = query[1].slice(1..-2) if query[1].start_with?(" & """" & "\" & """" & "" & """" & ") && query[1].end_with?(" & """" & "\" & """" & "" & """" & ")")
        WSSRuby.WriteLine("              #handle special fields")
        WSSRuby.WriteLine("              if query[0].end_with?(" & """" & "mail-address" & """" & ")")
        WSSRuby.WriteLine("                pattern = " & """" & "/^#{query[1]}/" & """" & "  #convert to Regex to capture bar@foo.com but not foobar@foo.com")
        WSSRuby.WriteLine("              elsif query[0].end_with?(" & """" & "mail-domain" & """" & ")")
        WSSRuby.WriteLine("                pattern = " & """" & "@#{query[1]}" & """" & "")
        WSSRuby.WriteLine("              else")
        WSSRuby.WriteLine("                pattern = query[1]")
        WSSRuby.WriteLine("              end")
        WSSRuby.WriteLine("            else")
        WSSRuby.WriteLine("              #No field specified.")
        WSSRuby.WriteLine("              field = :all")
        WSSRuby.WriteLine("              pattern = query_string")
        WSSRuby.WriteLine("            end")
        WSSRuby.WriteLine("            #convert to Regular Expression if " & """" & "/.../")
        WSSRuby.WriteLine("            pattern = Regexp.new( pattern.slice(1..-2) ) if pattern.start_with?(" & """" & "/" & """" & ") && pattern.end_with?(" & """" & "/" & """" & ")")
        WSSRuby.WriteLine("            address_queries[tag][field] << pattern")
        WSSRuby.WriteLine("          end")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        @queries = address_queries")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      #Returns Hash of cached values to be queried.")
        WSSRuby.WriteLine("      # param:  comm  - a Communication")
        WSSRuby.WriteLine("      def cache(comm)")
        WSSRuby.WriteLine("        address_list = Hash.new{|h,k| h[k] = []}")
        WSSRuby.WriteLine("        #convert to rfc822")
        WSSRuby.WriteLine("        { :from => comm.get_from(),")
        WSSRuby.WriteLine("          :to => comm.get_to(),")
        WSSRuby.WriteLine("          :cc => comm.get_cc(),")
        WSSRuby.WriteLine("          :bcc => comm.get_bcc()")
        WSSRuby.WriteLine("        }.each do |k, l|")
        WSSRuby.WriteLine("          l.each do |e|")
        WSSRuby.WriteLine("            v = e.to_rfc822_string")
        WSSRuby.WriteLine("            v.downcase! if !@settings[:case_sensitive]")
        WSSRuby.WriteLine("            address_list[k] << v")
        WSSRuby.WriteLine("          end")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        #add :recipient (to/cc/bcc) and :all (from/to/cc/bcc)")
        WSSRuby.WriteLine("        address_list[:recipient] = address_list[:to] + address_list[:cc] + address_list[:bcc]")
        WSSRuby.WriteLine("        address_list[:all] = address_list[:from] + address_list[:recipient]")
        WSSRuby.WriteLine("        return address_list")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      #Compares queries against item's communication fields and update's item's flags.")
        WSSRuby.WriteLine("      # param:  item  - a WSSitem")
        WSSRuby.WriteLine("      def compare(item)")
        WSSRuby.WriteLine("        comm = item.source_item.get_communication()")
        WSSRuby.WriteLine("        if !comm.nil?")
        WSSRuby.WriteLine("          search_data = cache(comm)")
        WSSRuby.WriteLine("          @queries.each do |tag, query_hash|")
        WSSRuby.WriteLine("            next if item.flags.include?(tag)  #item has already hit for the tag")
        WSSRuby.WriteLine("            #compare cached addresses by field with each query from array, until there is a match")
        WSSRuby.WriteLine("            item.flags.add(tag) if query_hash.any?{|field, query_array| query_array.any?{ |query| search_data[field].any?{|a| !a.to_s.match(query).nil?} }}")
        WSSRuby.WriteLine("          end")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Queries Class for queries against metadata properties.")
        WSSRuby.WriteLine("    #@queries is a Hash of property_name => Hash of tag => Array of value_query (or True if just looking for the property name}")
        WSSRuby.WriteLine("    # nil property_name key contains queries against all values.")
        WSSRuby.WriteLine("    class Property_Queries < Queries")
        WSSRuby.WriteLine("      # params: settings  - settings Hash from WSS")
        WSSRuby.WriteLine("      #         type      - key for file to load")
        WSSRuby.WriteLine("      def initialize(settings, type)")
        WSSRuby.WriteLine("        super(settings, type)")
        WSSRuby.WriteLine("        property_queries = Hash.new{|h,k| h[k] = Hash.new{|h2,k2| h2[k2] = []}}")
        WSSRuby.WriteLine("        #load CSV  and generate property queries")
        WSSRuby.WriteLine("        @queries.each do |tag, query_array|")
        WSSRuby.WriteLine("          #convert to :property_name => :tag => Array of value_query (or True)")
        WSSRuby.WriteLine("          query_array.each do |query_string|")
        WSSRuby.WriteLine("            p_name = nil")
        WSSRuby.WriteLine("            value_query = query_string")
        WSSRuby.WriteLine("            #Determine if there's a specific property name.")
        WSSRuby.WriteLine("            query = query_string.split(" & """" & ":" & """" & ")")
        WSSRuby.WriteLine("            if query.size != 1")
        WSSRuby.WriteLine("              p_name = query[0]")
        WSSRuby.WriteLine("              value_query = query[1..-1].join(" & """" & ":" & """" & ") #recombine property values with colons.")
        WSSRuby.WriteLine("            elsif query_string.end_with?(" & """" & ":" & """" & ")")
        WSSRuby.WriteLine("              #looking for the presence of this property")
        WSSRuby.WriteLine("              property_queries[ query[0] ][tag] = true")
        WSSRuby.WriteLine("              next")
        WSSRuby.WriteLine("            end")
        WSSRuby.WriteLine("            property_queries[p_name][tag] << value_query")
        WSSRuby.WriteLine("          end")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        @queries = property_queries")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      #Returns Hash of queries relevant to the property.")
        WSSRuby.WriteLine("      # param:  item  - a WSSitem")
        WSSRuby.WriteLine("      #         prop_name - Name of a metadata property")
        WSSRuby.WriteLine("      def cache(item, prop_name)")
        WSSRuby.WriteLine("        queries = @queries[nil] || Hash.new{|h,k| h[k]=[]}")
        WSSRuby.WriteLine("        @queries.select{ |k, v| !k.nil? && !k.match(prop_name).nil? }.each_value do |tag_hash|")
        WSSRuby.WriteLine("          tag_hash.each do |t, a|")
        WSSRuby.WriteLine("            next if item.flags.include?(t)  #item has already hit for the tag")
        WSSRuby.WriteLine("            if a == true #just looking for the presence of the property")
        WSSRuby.WriteLine("              item.flags.add(t)")
        WSSRuby.WriteLine("            else")
        WSSRuby.WriteLine("              queries[t] += a")
        WSSRuby.WriteLine("            end")
        WSSRuby.WriteLine("          end")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        return queries")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      #Compares queries against item's properties and update's item's flags.")
        WSSRuby.WriteLine("      # param:  item  - a WSSitem")
        WSSRuby.WriteLine("      def compare(item)")
        WSSRuby.WriteLine("        item.props.each do |prop_name, prop_value|")
        WSSRuby.WriteLine("          #determine queries to run")
        WSSRuby.WriteLine("          queries = cache(item, prop_name)")
        WSSRuby.WriteLine("          #run queries")
        WSSRuby.WriteLine("          if !queries.empty?  #there are queries for this property")
        WSSRuby.WriteLine("            search_data = stringify(prop_value)")
        WSSRuby.WriteLine("            #downcase value if it is needed")
        WSSRuby.WriteLine("            search_data.downcase! if !@settings[:case_sensitive]")
        WSSRuby.WriteLine("            queries.each do |tag, query_array|")
        WSSRuby.WriteLine("              next if item.flags.include?(tag)  #item has already hit for the tag")
        WSSRuby.WriteLine("              #compare property value with each query from array, until there is a match")
        WSSRuby.WriteLine("              item.flags.add(tag) if query_array.any?{|query| !search_data.match(query).nil?}")
        WSSRuby.WriteLine("            end")
        WSSRuby.WriteLine("          end")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      #Returns metadata property value as normalized string.")
        WSSRuby.WriteLine("      # param:  obj - Object from metadata properties Hash.")
        WSSRuby.WriteLine("      def stringify(obj)")
        WSSRuby.WriteLine("        case obj")
        WSSRuby.WriteLine("        when String")
        WSSRuby.WriteLine("          return obj")
        WSSRuby.WriteLine("        when TrueClass, FalseClass")
        WSSRuby.WriteLine("          return obj.to_s")
        WSSRuby.WriteLine("        when Fixnum, Float")
        WSSRuby.WriteLine("          return obj.to_s")
        WSSRuby.WriteLine("        when Java::OrgJodaTime::Duration")
        WSSRuby.WriteLine("          return obj.to_string")
        WSSRuby.WriteLine("        when org.joda.time.DateTime")
        WSSRuby.WriteLine("          return obj.to_string(" & """" & "YmdHMS" & """" & ")")
        WSSRuby.WriteLine("        when Java::byte[]")
        WSSRuby.WriteLine("          return obj.to_s.unpack(" & """" & "H*" & """" & ")[0]")
        WSSRuby.WriteLine("        when Java::ComNuixUtilExpression::b")
        WSSRuby.WriteLine("          return obj.to_s")
        WSSRuby.WriteLine("        else")
        WSSRuby.WriteLine("          return obj.to_a.map{ |e| stringify(e) }.join(" & """" & ";" & """" & ") if obj.respond_to?(:each)")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        return obj.to_s")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Class for handling the reports")
        WSSRuby.WriteLine("  # contains: @reports  - Hash containg the Reports.")
        WSSRuby.WriteLine("  # WSSreports.add(item) does Report.add(item) for each in @reports.")
        WSSRuby.WriteLine("  class WSSreports")
        WSSRuby.WriteLine("    require 'fileutils'")
        WSSRuby.WriteLine("    require 'csv'")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Creates the folder if required.")
        WSSRuby.WriteLine("    # params: settings  - settings Hash from WSS.")
        WSSRuby.WriteLine("    def initialize(settings)")
        WSSRuby.WriteLine("      items = settings[:report_items]")
        WSSRuby.WriteLine("      queries = settings[:report_queries]")
        WSSRuby.WriteLine("      if items || settings")
        WSSRuby.WriteLine("        guid = ENV_JAVA[" & """" & "nuix.simple.worker.guid" & """" & "]")
        WSSRuby.WriteLine("        directory = File.join( settings[:report_directory] , guid)")
        WSSRuby.WriteLine("        Log.out(" & """" & "Reporting to: #{directory}" & """" & ", true)")
        WSSRuby.WriteLine("        FileUtils.mkdir_p directory")
        WSSRuby.WriteLine("        @reports = {}")
        WSSRuby.WriteLine("        @reports[:items] = Item_Report.new(File.join(directory, " & """" & "items.csv" & """" & ")) if items")
        WSSRuby.WriteLine("        @reports[:query] = Query_Report.new(File.join(directory, " & """" & "queries.csv" & """" & ")) if queries")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Adds item to each in @reports using Report.add(item)")
        WSSRuby.WriteLine("    # param:  wss_item  - a WWSitem")
        WSSRuby.WriteLine("    def add(wss_item)")
        WSSRuby.WriteLine("      @reports.each_value { |v| v.add(wss_item) }")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Write Query Report")
        WSSRuby.WriteLine("    def close()")
        WSSRuby.WriteLine("      @reports[:query].write() if @reports.has_key?(:query)")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Class for report files")
        WSSRuby.WriteLine("    # reports should define add(item) to report an item.")
        WSSRuby.WriteLine("    class Report")
        WSSRuby.WriteLine("      # param:  location - location of report")
        WSSRuby.WriteLine("      def initialize(location)")
        WSSRuby.WriteLine("        @file_location = location")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      #Appends row to report.")
        WSSRuby.WriteLine("      # param:  row - row for CSV.")
        WSSRuby.WriteLine("      def add(row)")
        WSSRuby.WriteLine("        CSV.open(@file_location, " & """" & "a" & """" & ") { |csv| csv << row }")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Report Class for per-item report")
        WSSRuby.WriteLine("    # Reports: GUID, PATH, FLAGS")
        WSSRuby.WriteLine("    class Item_Report < Report")
        WSSRuby.WriteLine("      def initialize(location)")
        WSSRuby.WriteLine("        super(location)")
        WSSRuby.WriteLine("        #Create CSV and write header")
        WSSRuby.WriteLine("        @header = [" & """" & "GUID" & """" & ", " & """" & "MD5" & """" & ", " & """" & "Path" & """" & ", " & """" & "Flags" & """" & "]")
        WSSRuby.WriteLine("        Log.out(" & """" & "Creating item report at #{@file_location}" & """" & ", true)")
        WSSRuby.WriteLine("        CSV.open(@file_location, " & """" & "w" & """" & ") { |csv| csv << @header }")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      #Appends item to report.")
        WSSRuby.WriteLine("      # param:  wss_item  - a WWSitem")
        WSSRuby.WriteLine("      def add(wss_item)")
        WSSRuby.WriteLine("        #get values")
        WSSRuby.WriteLine("        guid = wss_item.worker_item.get_guid_path.last")
        WSSRuby.WriteLine("        md5 = wss_item.worker_item.get_digests.get_md5")
        WSSRuby.WriteLine("        path = File.join(wss_item.source_item.get_path_names.to_a)")
        WSSRuby.WriteLine("        flags = wss_item.flags.to_a.sort.join(" & """" & ";" & """" & ")")
        WSSRuby.WriteLine("        #add row")
        WSSRuby.WriteLine("        super( [guid, md5, path, flags] )")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Report Class for per-query report")
        WSSRuby.WriteLine("    # Reports: CATEGORY, NUMBER OF ITEMS")
        WSSRuby.WriteLine("    #includes TOTAL to tally items that hit on any category.")
        WSSRuby.WriteLine("    class Query_Report < Report")
        WSSRuby.WriteLine("      def initialize(location)")
        WSSRuby.WriteLine("        super(location)")
        WSSRuby.WriteLine("        @queries = Hash.new(0)")
        WSSRuby.WriteLine("        @total = 0")
        WSSRuby.WriteLine("        @header = [" & """" & "Category" & """" & ", " & """" & "Count" & """" & "]")
        WSSRuby.WriteLine("        Log.out(" & """" & "Query report will be created at: #{@file_location}" & """" & ", true)")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      #Adds item to report.")
        WSSRuby.WriteLine("      # param:  wss_item  - a WWSitem")
        WSSRuby.WriteLine("      def add(wss_item)")
        WSSRuby.WriteLine("        @total += 1")
        WSSRuby.WriteLine("        wss_item.flags.each { |f| @queries[f] += 1 }")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      #Writes report from @queries and @total.")
        WSSRuby.WriteLine("      def write()")
        WSSRuby.WriteLine("        Log.out(" & """" & "Creating query report at #{@file_location}" & """" & ", true)")
        WSSRuby.WriteLine("        CSV.open(@file_location, " & """" & "w" & """" & ") do |csv|")
        WSSRuby.WriteLine("          csv << @header")
        WSSRuby.WriteLine("          csv << [" & """" & "TOTAL RESPONSIVE ITEMS" & """" & ", @total]")
        WSSRuby.WriteLine("          @queries.each { |k, v| csv << [k.to_s, v] }")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("#Module for outputting log and error messages.")
        WSSRuby.WriteLine("module Log")
        WSSRuby.WriteLine("  @log = true")
        WSSRuby.WriteLine("  class << self")
        WSSRuby.WriteLine("    #Sets the log level.")
        WSSRuby.WriteLine("    # param:  b - True if verbose logging.")
        WSSRuby.WriteLine("    def set(b)")
        WSSRuby.WriteLine("      @log = b")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    #Puts log of string according to log level.")
        WSSRuby.WriteLine("    # params: str - String to log")
        WSSRuby.WriteLine("    #         level - if @@log should be overridden")
        WSSRuby.WriteLine("    def out(str, level)")
        WSSRuby.WriteLine("      puts " & """" & "WSS: #{str}" & """" & " if ( level || @log )")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("#Class for each WorkerItem, containing its details and methods for comparing itself against the WSS it is a part of.")
        WSSRuby.WriteLine("class WSSitem")
        WSSRuby.WriteLine("  include WSS")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  attr_accessor :flags, :source_item, :worker_item")
        WSSRuby.WriteLine("  def initialize(worker_item)")
        WSSRuby.WriteLine("    @flags = Set.new()")
        WSSRuby.WriteLine("    @worker_item = worker_item")
        WSSRuby.WriteLine("    @source_item = worker_item.get_source_item()")
        WSSRuby.WriteLine("    @properties = nil #loaded by props() when needed")
        WSSRuby.WriteLine("    process()")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Returns True if the item should be filtered out or is a container...")
        WSSRuby.WriteLine("  #Returns False if the item matched the whitelist.")
        WSSRuby.WriteLine("  #Returns nil if there is no filter.")
        WSSRuby.WriteLine("  def filter()")
        WSSRuby.WriteLine("    return nil if @@settings[:file_filter].nil?")
        WSSRuby.WriteLine("    if !@source_item.is_kind(" & """" & "container" & """" & ")")
        WSSRuby.WriteLine("      #check if item matches filter criteria")
        WSSRuby.WriteLine("      @@settings[:file_filter].each do |field, values|")
        WSSRuby.WriteLine("        case field")
        WSSRuby.WriteLine("        when :mime")
        WSSRuby.WriteLine("          v = @source_item.get_type().get_name()")
        WSSRuby.WriteLine("        when :kind")
        WSSRuby.WriteLine("          v = @source_item.get_kind().get_name()")
        WSSRuby.WriteLine("        when :ext")
        WSSRuby.WriteLine("          v = File.extname(@source_item.get_name)")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        v.downcase! if !@@settings[:case_sensitive]")
        WSSRuby.WriteLine("        return false if values.include?(v)  #item matched whitelist")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("    return true #item should be filtered out")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Returns String for custom metadata value.")
        WSSRuby.WriteLine("  #joins using " & """" & "|" & """" & " to work with Lightspeed")
        WSSRuby.WriteLine("  # appends property value from @@settings[:category] if required")
        WSSRuby.WriteLine("  def custom_metadata_value()")
        WSSRuby.WriteLine("    values = @flags.to_a")
        WSSRuby.WriteLine("    if !@@settings[:category].nil?")
        WSSRuby.WriteLine("      category_value = props()[@@settings[:category]]")
        WSSRuby.WriteLine("      values.map!{|n| " & """" & "#{n}_{category_value}" & """" & " } if !category_value.nil?")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("    return values.join(" & """" & "|" & """" & ")")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Checks the item against the filter, runs queries, and applies results.")
        WSSRuby.WriteLine("  #Returns True if the item was processed.")
        WSSRuby.WriteLine("  #Returns False if the item did not match.")
        WSSRuby.WriteLine("  #Returns nil if the item did not pass the filter.")
        WSSRuby.WriteLine("  def process()")
        WSSRuby.WriteLine("    #Check filter")
        WSSRuby.WriteLine("    if filter()")
        WSSRuby.WriteLine("      Log.out(" & """" & "Filtered out #{@source_item.get_name}" & """" & ", nil)")
        WSSRuby.WriteLine("      @worker_item.set_process_item(false) if @@settings[:exclude_items]")
        WSSRuby.WriteLine("      return nil")
        WSSRuby.WriteLine("    end:")
        WSSRuby.WriteLine("    #Run queries")
        WSSRuby.WriteLine("    Log.out(" & """" & "Querying #{@source_item.get_name}" & """" & ", nil)")
        WSSRuby.WriteLine("    if !matches().nil?")
        WSSRuby.WriteLine("      #Add tags")
        WSSRuby.WriteLine("      @flags.each { |tag| @worker_item.add_tag(tag.to_s) } if @@settings[:tag_items]")
        WSSRuby.WriteLine("      #Add custom metadata")
        WSSRuby.WriteLine("      @worker_item.add_custom_metadata(@@settings[:custom_metadata], custom_metadata_value(), " & """" & "text" & """" & ", " & """" & "user" & """" & ") if !@@settings[:custom_metadata].nil?")
        WSSRuby.WriteLine("      #Add to reports")
        WSSRuby.WriteLine("      @@reports.add(self) if !@@reports.nil?")
        WSSRuby.WriteLine("      return true")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("    return false #item did not match any criteria")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Returns the item's properties, keys are down-cased if necessary")
        WSSRuby.WriteLine("  def props()")
        WSSRuby.WriteLine("    if @properties.nil?")
        WSSRuby.WriteLine("      @properties = @source_item.get_properties()")
        WSSRuby.WriteLine("      @properties = @properties.map{ |k, v| [k.downcase, v] }.to_h if !@@settings[:case_sensitive]")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("    return @properties")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  #Compares item against the queries present in WSS. Marks unresponsive if required.")
        WSSRuby.WriteLine("  #Returns number of matches, or nil if it did not match any criteria. (i.e. unresponsive not marked)")
        WSSRuby.WriteLine("  def matches()")
        WSSRuby.WriteLine("    match_num = @@queries.compare(self)")
        WSSRuby.WriteLine("    if match_num == 0")
        WSSRuby.WriteLine("      Log.out(" & """" & "#{@source_item.get_name} did not hit on any terms." & """" & ", nil)")
        WSSRuby.WriteLine("      if !@@settings[:flag_unresponsive]")
        WSSRuby.WriteLine("        @worker_item.set_process_item(false) if @@settings[:exclude_items]")
        WSSRuby.WriteLine("        return nil  #nothing flagged")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      @flags.add(:unresponsive)")
        WSSRuby.WriteLine("    else")
        WSSRuby.WriteLine("      Log.out(" & """" & "#{@source_item.get_name} has #{match_num} matches: #{@flags.to_a.join(" & """" & ", " & """" & ")}" & """" & ", nil)")
        WSSRuby.WriteLine("      #Check for uniques")
        WSSRuby.WriteLine("      @worker_item.add_tag(" & """" & "unique|#{@flags.first.to_s}" & """" & ") if @@settings[:tag_unique] && match_num == 1")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("    return match_num")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("end")
        WSSRuby.Close()


        blnBuildUpdatedWSSRubyScript = True
    End Function

    Public Function blnBuildAddressQueriesExpandedRuby(ByVal sRubyScriptFileName As String, ByVal sNuixFilesDir As String) As Boolean
        blnBuildAddressQueriesExpandedRuby = False

        Dim AddressQueriesExpanded As StreamWriter

        AddressQueriesExpanded = New StreamWriter(sRubyScriptFileName)
        AddressQueriesExpanded.WriteLine("require File.join(ENV_JAVA[" & """" & "nuix.wss.scripts" & """" & "]," & """" & "WSSqueries.rb" & """" & ")")
        AddressQueriesExpanded.WriteLine("#Queries against addresses and expanded properties")
        AddressQueriesExpanded.WriteLine("class Address_Queries_Expanded < WSSqueries")
        AddressQueriesExpanded.WriteLine("	#Initialize and load queries from CSV.")
        AddressQueriesExpanded.WriteLine("	# param:  settings  - Hash of parameters")
        AddressQueriesExpanded.WriteLine("	#         :csv    - CSV input file")
        AddressQueriesExpanded.WriteLine("  #         :expanded_properties  - Array of strings of metadata property names")
        AddressQueriesExpanded.WriteLine("  #         :both        - Bool;")
        AddressQueriesExpanded.WriteLine("  #                         T  if search both props and comm.")
        AddressQueriesExpanded.WriteLine("  #                         F  if search only props (if expanded_properties contains queries).")
        AddressQueriesExpanded.WriteLine("  #                        nil if only search comm if props came back empty.")
        AddressQueriesExpanded.WriteLine("  def initialize(settings)")
        AddressQueriesExpanded.WriteLine("    super(settings[:csv])")
        AddressQueriesExpanded.WriteLine("    @expanded_properties = settings[:expanded_properties]")
        AddressQueriesExpanded.WriteLine("    @both = settings[:both]")
        AddressQueriesExpanded.WriteLine("  end")
        AddressQueriesExpanded.WriteLine("")
        AddressQueriesExpanded.WriteLine("  #Compares queries against expanded addresses and update's item's flags.")
        AddressQueriesExpanded.WriteLine("  # param:  item  - a WSSitem")
        AddressQueriesExpanded.WriteLine("  def compare(item)")
        AddressQueriesExpanded.WriteLine("    super(item, cache(item))")
        AddressQueriesExpanded.WriteLine("  end")
        AddressQueriesExpanded.WriteLine("")
        AddressQueriesExpanded.WriteLine("  #Returns array of data to be queried.")
        AddressQueriesExpanded.WriteLine("  # param:  item  - a WSSitem")
        AddressQueriesExpanded.WriteLine("  def cache(item)")
        AddressQueriesExpanded.WriteLine("    search_data = []")
        AddressQueriesExpanded.WriteLine("    #Get expanded property values, if they exist")
        AddressQueriesExpanded.WriteLine("    if !@expanded_properties.empty?")
        AddressQueriesExpanded.WriteLine("      search_data = cache_props(item.props)")
        AddressQueriesExpanded.WriteLine("      #Exit if only looking at properties")
        AddressQueriesExpanded.WriteLine("      if !@both #if (false or nil)")
        AddressQueriesExpanded.WriteLine("        if (@both.nil? && search_data.empty?)")
        AddressQueriesExpanded.WriteLine("          Log.out(" & """" & "No properties matches. Searching Communication." & """" & ", nil)")
        AddressQueriesExpanded.WriteLine("        else")
        AddressQueriesExpanded.WriteLine("          return search_data")
        AddressQueriesExpanded.WriteLine("        end")
        AddressQueriesExpanded.WriteLine("      end")
        AddressQueriesExpanded.WriteLine("    end")
        AddressQueriesExpanded.WriteLine("    #Add communication values")
        AddressQueriesExpanded.WriteLine("    return search_data + cache_comm( item.source_item.get_communication() )")
        AddressQueriesExpanded.WriteLine("  end")
        AddressQueriesExpanded.WriteLine("")
        AddressQueriesExpanded.WriteLine("		#Returns Array of cached rfc822 address values to be queried.")
        AddressQueriesExpanded.WriteLine("  # param:  comm  - a Communication:)")
        AddressQueriesExpanded.WriteLine("  def cache_comm(comm)")
        AddressQueriesExpanded.WriteLine("    #convert to rfc822")
        AddressQueriesExpanded.WriteLine("    return [] if comm.nil?")
        AddressQueriesExpanded.WriteLine("    return [:get_from, :get_to, :get_cc, :get_bcc].map{|c| comm.send(c).to_a.map(&:to_rfc822_string)}.flatten")
        AddressQueriesExpanded.WriteLine("  end")
        AddressQueriesExpanded.WriteLine("")
        AddressQueriesExpanded.WriteLine("  #Returns Array of cached property values to be queried.")
        AddressQueriesExpanded.WriteLine("  # param:  props - Map from item.props")
        AddressQueriesExpanded.WriteLine("  def cache_props(props)")
        AddressQueriesExpanded.WriteLine("    values = []")
        AddressQueriesExpanded.WriteLine("    @expanded_properties.each do |p|")
        AddressQueriesExpanded.WriteLine("      if props.has_key?(p)")
        AddressQueriesExpanded.WriteLine("        Log.out(" & """" & "Found metadata #{p}" & """" & ", nil)")
        AddressQueriesExpanded.WriteLine("        v = expanded_prop(props[p])")
        AddressQueriesExpanded.WriteLine("        Log.out(" & """" & "Adding: #{v}" & """" & ", nil)")
        AddressQueriesExpanded.WriteLine("        values << v")
        AddressQueriesExpanded.WriteLine("      end")
        AddressQueriesExpanded.WriteLine("    end")
        AddressQueriesExpanded.WriteLine("    return values")
        AddressQueriesExpanded.WriteLine("  end")

        AddressQueriesExpanded.WriteLine("  #Converts object from metadata properties map to searchable string.")
        AddressQueriesExpanded.WriteLine("  # param:  obj - an object")
        AddressQueriesExpanded.WriteLine("  def expanded_prop(obj)")
        AddressQueriesExpanded.WriteLine("    return obj.to_a.map{ |e| expanded_prop(e) }.join(" & """" & "; " & """" & ") if obj.respond_to?(:each)")
        AddressQueriesExpanded.WriteLine("    return obj.to_s")
        AddressQueriesExpanded.WriteLine("  end")
        AddressQueriesExpanded.WriteLine("end")
        blnBuildAddressQueriesExpandedRuby = True
    End Function

    Public Function blnBuildLogRuby(ByVal sRubyScriptFileName As String, ByVal sNuixFilesDir As String) As Boolean
        blnBuildLogRuby = False

        Dim LogRuby As StreamWriter

        LogRuby = New StreamWriter(sRubyScriptFileName)

        LogRuby.WriteLine("#Module for outputting log and error messages.")
        LogRuby.WriteLine("#contains:  @log  - the log level (T = verbose logging).")
        LogRuby.WriteLine("#")
        LogRuby.WriteLine("#methods:   .set(bool)        - sets the log level.")
        LogRuby.WriteLine("#           .out(str, level)  - outputs str, if @log or level is T.")
        LogRuby.WriteLine("module Log")
        LogRuby.WriteLine("  @log = true")
        LogRuby.WriteLine("  class << self")
        LogRuby.WriteLine("    #Sets the log level.")
        LogRuby.WriteLine("    # param:  b - True if verbose logging.")
        LogRuby.WriteLine("    def set(b)")
        LogRuby.WriteLine("      @log = b")
        LogRuby.WriteLine("    end")
        LogRuby.WriteLine("")
        LogRuby.WriteLine("    #Puts log of string according to log level.")
        LogRuby.WriteLine("    # params: str - String to log")
        LogRuby.WriteLine("    #         level - if @log should be overridden")
        LogRuby.WriteLine("    def out(str, level)")
        LogRuby.WriteLine("      puts " & """" & "WSS: #{str}" & """" & "if ( level || @log )")
        LogRuby.WriteLine("    end")
        LogRuby.WriteLine("  end")
        LogRuby.WriteLine("end")

        blnBuildLogRuby = True
    End Function

    Public Function blnBuildWSSItemRuby(ByVal sRubyScriptFileName As String, ByVal sNuixFilesDir As String) As Boolean
        blnBuildWSSItemRuby = False

        Dim WSSItemRuby As StreamWriter

        WSSItemRuby = New StreamWriter(sRubyScriptFileName)
        WSSItemRuby.WriteLine("#Class for each WorkerItem, containing its details and methods for comparing itself against the WSS it is a part of.")
        WSSItemRuby.WriteLine("class WSSitem")
        WSSItemRuby.WriteLine("  include WSS")
        WSSItemRuby.WriteLine("  attr_accessor :flags, :source_item, :worker_item")
        WSSItemRuby.WriteLine("  #@flags is a Set of string that will be applied as tags/custom metadata")
        WSSItemRuby.WriteLine("  #Initializes and processes the item.")
        WSSItemRuby.WriteLine("  # param: worker_item  - WorkerItem from Nuix")
        WSSItemRuby.WriteLine("  def initialize(worker_item)")
        WSSItemRuby.WriteLine("    @worker_item = worker_item")
        WSSItemRuby.WriteLine("    if @@ready")
        WSSItemRuby.WriteLine("      @flags = Set.new()")
        WSSItemRuby.WriteLine("      @source_item = worker_item.get_source_item()")
        WSSItemRuby.WriteLine("      @properties = nil #retrieved by props() when needed")
        WSSItemRuby.WriteLine("      process()")
        WSSItemRuby.WriteLine("    else  #fail case")
        WSSItemRuby.WriteLine("     Log.out(" & """" & "Invalid Script Configuration - skipping item" & """" & ", true)")
        WSSItemRuby.WriteLine("      @worker_item.set_process_item(false)")
        WSSItemRuby.WriteLine("    end")
        WSSItemRuby.WriteLine("  end")
        WSSItemRuby.WriteLine("")
        WSSItemRuby.WriteLine("  #Returns category property value.")
        WSSItemRuby.WriteLine("  #Returns nil if no category or no property.")
        WSSItemRuby.WriteLine("  # uses the first value identified")
        WSSItemRuby.WriteLine("  def category_value()")
        WSSItemRuby.WriteLine("    if !@@settings[:category_property].nil?")
        WSSItemRuby.WriteLine("      @@settings[:category_property].each do |category|")
        WSSItemRuby.WriteLine("        category_value = props()[category]")
        WSSItemRuby.WriteLine("        if !category_value.nil?")
        WSSItemRuby.WriteLine("          Log.out(" & """" & "Identified #{category}: #{category_value}" & """" & ", nil)")
        WSSItemRuby.WriteLine("          return category_value")
        WSSItemRuby.WriteLine("        end")
        WSSItemRuby.WriteLine("      end")
        WSSItemRuby.WriteLine("    end")
        WSSItemRuby.WriteLine("    return nil")
        WSSItemRuby.WriteLine("  end")

        WSSItemRuby.WriteLine("  #Returns String for custom metadata value.")
        WSSItemRuby.WriteLine("  #joins using " & """" & "|" & """" & "to work with Lightspeed")
        WSSItemRuby.WriteLine("  # appends property value from @@settings[:category_property] if required")
        WSSItemRuby.WriteLine("  def custom_metadata_value()")
        WSSItemRuby.WriteLine("    values = @flags.to_a")
        WSSItemRuby.WriteLine("    category = category_value()")
        WSSItemRuby.WriteLine("    values.map!{|n| " & """" & "#{n}_#{category}" & """" & " } if !category.nil?")
        WSSItemRuby.WriteLine("    return values.join(" & """" & "|" & """" & ")")
        WSSItemRuby.WriteLine("  end")

        WSSItemRuby.WriteLine("  #Returns False if the item matched the whitelist.")
        WSSItemRuby.WriteLine("  #Returns nil if there is no filter.")
        WSSItemRuby.WriteLine("  #Returns True if the item should be filtered out.")
        WSSItemRuby.WriteLine("  def filter()")
        WSSItemRuby.WriteLine("    #check if item matches filter criteria")
        WSSItemRuby.WriteLine("    return nil if @@settings[:file_filter].empty?")
        WSSItemRuby.WriteLine("    @@settings[:file_filter].each do |field, values|")
        WSSItemRuby.WriteLine("      case field")
        WSSItemRuby.WriteLine("      when :mime")
        WSSItemRuby.WriteLine("        v = @source_item.get_type().get_name()")
        WSSItemRuby.WriteLine("      when :kind")
        WSSItemRuby.WriteLine("        v = @source_item.get_kind().get_name()")
        WSSItemRuby.WriteLine("      when :ext")
        WSSItemRuby.WriteLine("        #downcase and ignore leading '.'")
        WSSItemRuby.WriteLine("        v = File.extname(@source_item.get_name).downcase[1..-1]")
        WSSItemRuby.WriteLine("      end")
        WSSItemRuby.WriteLine("      return false if values.include?(v)  #item matched whitelist")
        WSSItemRuby.WriteLine("    end")
        WSSItemRuby.WriteLine("    return true #item should be filtered out")
        WSSItemRuby.WriteLine("  end")
        WSSItemRuby.WriteLine("")
        WSSItemRuby.WriteLine("  #Checks the item against the filter, runs queries, and applies results.")
        WSSItemRuby.WriteLine("  #Compares item against the queries present in WSS. Marks unresponsive if required.")
        WSSItemRuby.WriteLine("  #Returns True if the item was processed")
        WSSItemRuby.WriteLine("  #Returns False if the item did not match")
        WSSItemRuby.WriteLine("  #Returns nil if the item did not get queried")
        WSSItemRuby.WriteLine("  def process()")
        WSSItemRuby.WriteLine("    #Check filter")
        WSSItemRuby.WriteLine("    if filter()")
        WSSItemRuby.WriteLine("      #Skip item that should be filtered out")
        WSSItemRuby.WriteLine("      Log.out(" & """" & "Filtered out #{@source_item.get_name} (#{@source_item.get_type().get_name()})" & """" & ", nil)")
        WSSItemRuby.WriteLine("      @worker_item.set_process_item(false) if @@settings[:exclude_items]")
        WSSItemRuby.WriteLine("    elsif @source_item.is_kind(" & """" & "container" & """" & ")")
        WSSItemRuby.WriteLine("      #Skip containers")
        WSSItemRuby.WriteLine("      Log.out(" & """" & "Skipping container: #{@source_item.get_name}" & """" & ", nil)")
        WSSItemRuby.WriteLine("    else")
        WSSItemRuby.WriteLine("      #Compare against the queries")
        WSSItemRuby.WriteLine("      Log.out(" & """" & "Querying #{@source_item.get_name}" & """" & ", nil)")
        WSSItemRuby.WriteLine("      @@queries.compare(self)")
        WSSItemRuby.WriteLine("      #Check results")
        WSSItemRuby.WriteLine("      if !@flags.empty?")
        WSSItemRuby.WriteLine("        Log.out(" & """" & "Found #{@flags.size} matches for #{@source_item.get_name}:\n#{@flags.to_a.join(" & """" & ", " & """" & ")}" & """" & ", nil)")
        WSSItemRuby.WriteLine("        #Check for uniques")
        WSSItemRuby.WriteLine("        @worker_item.add_tag(" & """" & "unique|#{@flags.first.to_s}" & """" & ") if @@settings[:tag_unique] && match_num == 1")
        WSSItemRuby.WriteLine("      else")
        WSSItemRuby.WriteLine("#Handle unresponsive")
        WSSItemRuby.WriteLine("        Log.out(" & """" & "No matches for #{@source_item.get_name}" & """" & ", nil)")
        WSSItemRuby.WriteLine("        if @@settings[:flag_unresponsive]")
        WSSItemRuby.WriteLine("          @flags.add(:unresponsive)")
        WSSItemRuby.WriteLine("        else")
        WSSItemRuby.WriteLine("          @worker_item.set_process_item(false) if @@settings[:exclude_items]")
        WSSItemRuby.WriteLine("          return false")
        WSSItemRuby.WriteLine("        end")
        WSSItemRuby.WriteLine("      end")
        WSSItemRuby.WriteLine("      #Sort item according to flags")
        WSSItemRuby.WriteLine("      # should only get here if item has flags (unresponsive was added or returned false above)")
        WSSItemRuby.WriteLine("      #Add tags")
        WSSItemRuby.WriteLine("      @flags.each { |tag| @worker_item.add_tag(tag.to_s) } if @@settings[:tag_items]")
        WSSItemRuby.WriteLine("      #Add custom metadata")
        WSSItemRuby.WriteLine("      @worker_item.add_custom_metadata(@@settings[:custom_metadata], custom_metadata_value(), " & """" & "text" & """" & ", " & """" & "user" & """" & ") if !@@settings[:custom_metadata].nil?")
        WSSItemRuby.WriteLine("      #Add to reports")
        WSSItemRuby.WriteLine("      @@reports.add(self) if !@@reports.nil?")
        WSSItemRuby.WriteLine("      return true")
        WSSItemRuby.WriteLine("    end")
        WSSItemRuby.WriteLine("    return nil  #item filtered or is container (i.e. was not queried)")
        WSSItemRuby.WriteLine("  end")
        WSSItemRuby.WriteLine("")
        WSSItemRuby.WriteLine("  #Returns the item's properties")
        WSSItemRuby.WriteLine("  def props()")
        WSSItemRuby.WriteLine("    @properties = @source_item.get_properties() if @properties.nil?")
        WSSItemRuby.WriteLine("    return @properties")
        WSSItemRuby.WriteLine("  end")
        WSSItemRuby.WriteLine("end")

        blnBuildWSSItemRuby = True
    End Function

    Public Function blnBuildWSSQueriesRuby(ByVal sRubyScriptFileName As String, ByVal sNuixFilesDir As String) As Boolean
        blnBuildWSSQueriesRuby = False

        Dim WSSQueriesRuby As StreamWriter

        WSSQueriesRuby = New StreamWriter(sRubyScriptFileName)
        WSSQueriesRuby.WriteLine("")
        WSSQueriesRuby.WriteLine("#Class for querying items.")
        WSSQueriesRuby.WriteLine("class WSSqueries")
        WSSQueriesRuby.WriteLine("  attr_reader :queries")
        WSSQueriesRuby.WriteLine("  #@queries is a Hash of :tag => pattern. The pattern may be the union of multiple queries.")
        WSSQueriesRuby.WriteLine("  #Initialize and loads @queries from CSV file.")
        WSSQueriesRuby.WriteLine("  #Multiple search terms are converted into a single pattern up-front.")
        WSSQueriesRuby.WriteLine("  # param: location  - CSV input file")
        WSSQueriesRuby.WriteLine("  def initialize(location)")
        WSSQueriesRuby.WriteLine("    @queries = {}")
        WSSQueriesRuby.WriteLine("    load_csv(location).each do |tag, terms|")
        WSSQueriesRuby.WriteLine("      @queries[tag] = Regexp.union( terms.map!{|s| to_pattern(s)} )")
        WSSQueriesRuby.WriteLine("    end")
        WSSQueriesRuby.WriteLine("    Log.out(" & """" & "Queries: #{JSON.pretty_generate(@queries)}" & """" & ", nil)")
        WSSQueriesRuby.WriteLine("  end")

        WSSQueriesRuby.WriteLine("  #Compares list of search_data against queries using String.match()")
        WSSQueriesRuby.WriteLine("  # search_data is converted to single string using .join(" & """" & "; " & """" & ")")
        WSSQueriesRuby.WriteLine("  #Updates item's flags on matches")
        WSSQueriesRuby.WriteLine("  # params: item        - the WSSitem")
        WSSQueriesRuby.WriteLine("  #         search_data - array of strings to be searched")
        WSSQueriesRuby.WriteLine("  def compare(item, search_data)")
        WSSQueriesRuby.WriteLine("    if !search_data.empty?")
        WSSQueriesRuby.WriteLine("      search_string = search_data.join(" & """" & "; " & """" & ")")
        WSSQueriesRuby.WriteLine("      Log.out(" & """" & "Searching: #{search_string}" & """" & ", nil)")
        WSSQueriesRuby.WriteLine("      @queries.each do |tag, query|")
        WSSQueriesRuby.WriteLine("        next if item.flags.include?(tag)")
        WSSQueriesRuby.WriteLine("        item.flags.add(tag) if !search_string.match(query).nil?")
        WSSQueriesRuby.WriteLine("      end")
        WSSQueriesRuby.WriteLine("    end")
        WSSQueriesRuby.WriteLine("  end")

        WSSQueriesRuby.WriteLine("  #Returns Hash of tag => Array of query Strings from CSV.")
        WSSQueriesRuby.WriteLine("  # param:  file_name - CSV file to load, formatted as: tag, query")
        WSSQueriesRuby.WriteLine("  def load_csv(file_name)")
        WSSQueriesRuby.WriteLine("    h = Hash.new{|h,k| h[k]=[]}")
        WSSQueriesRuby.WriteLine("    if !file_name.nil? && File.exist?(file_name)")
        WSSQueriesRuby.WriteLine("      CSV.foreach(file_name) do |row|")
        WSSQueriesRuby.WriteLine("        tag = row[0]")
        WSSQueriesRuby.WriteLine("        query = row[1]")
        WSSQueriesRuby.WriteLine("        h[tag.to_sym] << query")
        WSSQueriesRuby.WriteLine("      end")
        WSSQueriesRuby.WriteLine("      Log.out(" & """" & "Loaded #{file_name}" & """" & ", true)")
        WSSQueriesRuby.WriteLine("    else")
        WSSQueriesRuby.WriteLine("      Log.out(" & """" & "No file found at: #{file_name}" & """" & ", true)")
        WSSQueriesRuby.WriteLine("    end")
        WSSQueriesRuby.WriteLine("    return h")
        WSSQueriesRuby.WriteLine("  end")
        WSSQueriesRuby.WriteLine("")
        WSSQueriesRuby.WriteLine("  #Returns Regexp pattern from string.")
        WSSQueriesRuby.WriteLine("  # param:  string  - a string, stars and ends with " & """" & "/" & """" & "if it's a regexp")
        WSSQueriesRuby.WriteLine("  def to_pattern(string)")
        WSSQueriesRuby.WriteLine("    if string.start_with?(" & """" & "/" & """" & ") && string.end_with?(" & """" & "/" & """" & ")")
        WSSQueriesRuby.WriteLine("      string = string.slice(1..-2)")
        WSSQueriesRuby.WriteLine("    else")
        WSSQueriesRuby.WriteLine("      string = Regexp.escape(string)")
        WSSQueriesRuby.WriteLine("    end")
        WSSQueriesRuby.WriteLine("    return Regexp.new( string, true )")
        WSSQueriesRuby.WriteLine("  end")
        WSSQueriesRuby.WriteLine("end")

        blnBuildWSSQueriesRuby = True
    End Function

    Public Function blnBuildWSSFullRuby(ByVal sRubyScriptFileName As String, ByVal sNuixFilesDir As String) As Boolean
        blnBuildWSSFullRuby = False


        '      WSSFullRuby = New StreamWriter(sRubyScriptFileName)
        '      WSSFullRuby.WriteLine("")

        '      WSSFullRuby.WriteLine("#Module for the Worker-Side Script")
        '      WSSFullRuby.WriteLine("module WSS")
        '      WSSFullRuby.WriteLine("require 'json'")
        '      WSSFullRuby.WriteLine("require 'csv'")
        '      WSSFullRuby.WriteLine("require 'set'")
        '      WSSFullRuby.WriteLine("require File.join(ENV_JAVA[" & """" & "nuix.wss.scripts" & """" & "], " & """" & "Log.rb" & """" & ")")
        '      WSSFullRuby.WriteLine("require File.join(ENV_JAVA[" & """" & "nuix.wss.scripts" & """" & "], " & """" & "WSSitem.rb" & """" & ")")
        '      WSSFullRuby.WriteLine("require File.join(ENV_JAVA[" & """" & "nuix.wss.scripts" & """" & "], " & """" & "WSSreports.rb" & """" & ")")
        '      WSSFullRuby.WriteLine("")
        '      WSSFullRuby.WriteLine("@@settings = {} #the actions to perform after querying")
        '      WSSFullRuby.WriteLine("@@queries = nil #a subclass of WSSqueries, containing the queries and a compare(item) method to match the items against the queries.")
        '      WSSFullRuby.WriteLine("@@reports = nil #WSSreports object; Add items to reports using @@reports.add(item)")
        '      WSSFullRuby.WriteLine("#@@ready - True if in a valid run state, otherwise the items will be skipped")
        '      WSSFullRuby.WriteLine("")
        '      WSSFullRuby.WriteLine("class << self")
        '      WSSFullRuby.WriteLine("    #Initializes WSS")
        '      WSSFullRuby.WriteLine("    # param:  json    - location of JSON file")
        '      WSSFullRuby.WriteLine("      #settings_json[:verbose]                        = Bool")
        '      WSSFullRuby.WriteLine("      #settings_json[:settings][:exclude_items]       = Bool")
        '      WSSFullRuby.WriteLine("      #settings_json[:settings][:flag_unresponsive]   = Bool")
        '      WSSFullRuby.WriteLine("      #settings_json[:settings][:tag_items]           = Bool")
        '      WSSFullRuby.WriteLine("      #settings_json[:settings][:tag_unique]          = Bool")
        '      WSSFullRuby.WriteLine("      #settings_json[:settings][:custom_metadata]     = String")
        '      WSSFullRuby.WriteLine("      #settings_json[:settings][:category_property]   = Array of Strings")
        '      WSSFullRuby.WriteLine("      #settings_json[:settings][:file_filter]         = Hash")
        '      WSSFullRuby.WriteLine("      #settings_json[:settings][:file_filter][:mime]  == Array of Strings")
        '      WSSFullRuby.WriteLine("      #settings_json[:settings][:file_filter][:kind]  == Array of Strings")
        '      WSSFullRuby.WriteLine("      #settings_json[:settings][:file_filter][:ext]   == Array of Strings")
        '      WSSFullRuby.WriteLine("      #settings_json[:queries][:type]                 = 'Address_Queries_Expanded'")
        '      WSSFullRuby.WriteLine("      #settings_json[:queries][:csv]                  = path to CSV")
        '      WSSFullRuby.WriteLine("      #settings_json[:queries][:expanded_properties]  = Array of Strings")
        '      WSSFullRuby.WriteLine("      #settings_json[:queries][:both]                 = Bool (T to search both comm and props. False to only search props. Nil to only search comm if props comes back empty)")
        '      WSSFullRuby.WriteLine("      #settings_json[:reporting][:report_items]       = Bool")
        '      WSSFullRuby.WriteLine("      #settings_json[:reporting][:report_queries]     = Bool")
        '      WSSFullRuby.WriteLine("      #settings_json[:reporting][:report_directory]   = String")
        '      WSSFullRuby.WriteLine("    def load(json)")
        '      WSSFullRuby.WriteLine("      if json.nil? || !File.exist?(json)")
        '      WSSFullRuby.WriteLine("        Log.out(" & """" & "No Configuration File" & """" & ", true)")
        '      WSSFullRuby.WriteLine("      else")
        '      WSSFullRuby.WriteLine("        Log.out(" & """" & "Loading configuration from #{json}" & """" & ", true)")
        '      WSSFullRuby.WriteLine("        settings_json = JSON.parse(File.read(json), :symbolize_names => true)")
        '      WSSFullRuby.WriteLine("        Log.out(" & """" & "Configuration: #{JSON.pretty_generate(settings_json)}" & """" & ", true)")
        '      WSSFullRuby.WriteLine("        Log.set(settings_json[:verbose])")
        '      WSSFullRuby.WriteLine("        @@settings = settings_json[:settings]")
        '      WSSFullRuby.WriteLine("        #Sanitize settings input (i.e. replace empty with nil)")
        '      WSSFullRuby.WriteLine("        @@settings[:custom_metadata] = nil if empty_nil(@@settings[:custom_metadata])")
        '      WSSFullRuby.WriteLine("        @@settings[:category_property] = nil if @@settings[:custom_metadata].nil? || @@settings[:category_property].empty?")
        '      WSSFullRuby.WriteLine("        if !@@settings[:file_filter].nil?")
        '      WSSFullRuby.WriteLine("          @@settings[:file_filter].delete_if{|k,v| v.empty?}")
        '      WSSFullRuby.WriteLine("          @@settings[:file_filter].each_value{|v| v.each(&:downcase!)}")
        '      WSSFullRuby.WriteLine("        end")
        '      WSSFullRuby.WriteLine("        #Initialize Queries")
        '      WSSFullRuby.WriteLine("        query_settings = settings_json[:queries]")
        '      WSSFullRuby.WriteLine("        type = query_settings.delete(:type)")
        '      WSSFullRuby.WriteLine("        queries_file = File.join(ENV_JAVA[" & """" & "nuix.wss.scripts" & """" & "], type + " & """" & ".rb" & """" & ")")
        '      WSSFullRuby.WriteLine("        Log.out(" & """" & "Queries Script: #{queries_file}" & """" & ", true)")
        '      WSSFullRuby.WriteLine("        require queries_file")
        '      WSSFullRuby.WriteLine("        @@queries = Object.const_get(type).new(query_settings)")
        '      WSSFullRuby.WriteLine("        #Initialize Reports")
        '      WSSFullRuby.WriteLine("        @@reports = WSSreports.new(settings_json[:reporting])")
        '      WSSFullRuby.WriteLine("        #Validate")
        '      WSSFullRuby.WriteLine("        if !(@@settings[:tag_unique] || @@settings[:tag_items] || @@settings[:exclude_items] || @@settings[:flag_unresponsive] || !@@settings[:custom_metadata].nil)")
        '      WSSFullRuby.WriteLine("          Log.out(" & """" & "ERROR - Script isn't configured to do anything." & """" & ", true)")
        '      WSSFullRuby.WriteLine("        else")
        '      WSSFullRuby.WriteLine("          if @@queries.queries.empty?")
        '      WSSFullRuby.WriteLine("            Log.out(" & """" & "ERROR - There are no queries." & """" & ", true)")
        '      WSSFullRuby.WriteLine("          else")
        '      WSSFullRuby.WriteLine("            Log.out(" & """" & "Ready to go…" & """" & ", true)")
        '      WSSFullRuby.WriteLine("            @@ready = true")
        '      WSSFullRuby.WriteLine("          end")
        '      WSSFullRuby.WriteLine("        end")
        '      WSSFullRuby.WriteLine("      end")
        '      WSSFullRuby.WriteLine("      Log.out(" & """" & "ERROR - Invalid Script Configuration!" & """" & ", true) if !@@ready")
        '      WSSFullRuby.WriteLine("end")
        '      WSSFullRuby.WriteLine("")
        '      WSSFullRuby.WriteLine("    #Processes the worker_item by initializing a WSSitem.")
        '      WSSFullRuby.WriteLine("    # param:  worker_item - Nuix WorkerItem")
        '      WSSFullRuby.WriteLine("    def process(worker_item)")
        '      WSSFullRuby.WriteLine("      WSSitem.new(worker_item)")
        '      WSSFullRuby.WriteLine("    end")
        '      WSSFullRuby.WriteLine("")
        '      WSSFullRuby.WriteLine("    #Close Reports")
        '      WSSFullRuby.WriteLine("    def close()")
        '      WSSFullRuby.WriteLine("      Log.out(" & """" & "Closing…" & """" & ", nil)")
        '      WSSFullRuby.WriteLine("      @@reports.close() if !@@reports.nil?")
        '      WSSFullRuby.WriteLine("    end")

        '      WSSFullRuby.WriteLine("    #Returns T/F if the string is empty after strip.")
        '      WSSFullRuby.WriteLine("    # param:  string  - a String.")
        '      WSSFullRuby.WriteLine("    def empty_nil(string)")
        '      WSSFullRuby.WriteLine("      return !string.nil? && string.strip.empty?")
        '      WSSFullRuby.WriteLine("    end")
        '      WSSFullRuby.WriteLine("  end")
        '      WSSFullRuby.WriteLine("end")
        '      WSSFullRuby.WriteLine("		WSSReportsRuby.WriteLine(" & """" & "#Class for handling the reports" & """" & ")")
        '      WSSFullRuby.WriteLine("		WSSReportsRuby.WriteLine(" & """" & "#contains: @reports  - Array containing the Reports." & """" & ")")
        '      WSSFullRuby.WriteLine("		WSSReportsRuby.WriteLine(" & """" & "# WSSreports.add(item) does Report.add(item) for each in @reports." & """" & ")")
        'WSSFullRuby.WriteLine("		WSSReportsRuby.WriteLine(" & """" & "class WSSreports" & """" & ")
        blnBuildWSSFullRuby = True
    End Function
    Public Function blnBuildFinalWSSRubyScript(ByVal sRubyScriptFileName As String, ByVal sString As String) As FunctionType
        blnBuildFinalWSSRubyScript = False

        Dim WSSRuby As StreamWriter

        WSSRuby = New StreamWriter(sRubyScriptFileName)

        WSSRuby.WriteLine("#Worker-Side Script for NEAMM")
        WSSRuby.WriteLine("# v1.0")
        WSSRuby.WriteLine("#requires: ENV_JAVA[" & """" & "nuix.wss.config" & """" & "]")
        WSSRuby.WriteLine("def nuixWorkerItemCallbackInit()")
        WSSRuby.WriteLine("  WSS.load( ENV_JAVA[" & """" & "nuix.wss.config" & """" & "] )")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("def nuixWorkerItemCallback(worker_item)")
        WSSRuby.WriteLine("  WSSitem.new(worker_item)")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("def nuixWorkerItemCallbackClose")
        WSSRuby.WriteLine("  WSS.close()")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("module WSS")
        WSSRuby.WriteLine("  require 'json'")
        WSSRuby.WriteLine("  require 'csv'")
        WSSRuby.WriteLine("  require 'set'")
        WSSRuby.WriteLine("  @@settings = {}")
        WSSRuby.WriteLine("  @@queries = nil")
        WSSRuby.WriteLine("  @@reports = nil")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  class << self")
        WSSRuby.WriteLine("    def load(json)")
        WSSRuby.WriteLine("      if json.nil? || !File.exist?(json)")
        WSSRuby.WriteLine("        Log.out(" & """" & "No Configuration File!" & """" & ", true)")
        WSSRuby.WriteLine("      else")
        WSSRuby.WriteLine("        Log.out(" & """" & "Loading configuration from #{json}" & """" & ", true)")
        WSSRuby.WriteLine("        @@settings = JSON.parse(File.read(json), :symbolize_names => true)")
        WSSRuby.WriteLine("        Log.out(" & """" & "Configuration: #{JSON.pretty_generate(@@settings)}" & """" & ", true)")
        WSSRuby.WriteLine("        Log.set(@@settings[:verbose])")
        WSSRuby.WriteLine("        @@settings[:custom_metadata] = nil if empty_nil(@@settings[:custom_metadata])")
        WSSRuby.WriteLine("        @@settings[:category_property] = nil if @@settings[:custom_metadata].nil? || @@settings[:category_property].empty?")
        WSSRuby.WriteLine("        if !@@settings[:file_filter].nil?")
        WSSRuby.WriteLine("          @@settings[:file_filter].delete_if{|k,v| v.empty?}")
        WSSRuby.WriteLine("          @@settings[:file_filter].each_value{|v| v.each(&:downcase!)}")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        query_settings = [:csv_queries, :expanded_properties].map{|f| @@settings.delete(f)}")
        WSSRuby.WriteLine("        @@queries = Address_Queries_Expanded.new(*query_settings)")
        WSSRuby.WriteLine("        report_settings = [:report_directory, :report_items, :report_queries].map{|f| @@settings.delete(f)}")
        WSSRuby.WriteLine("        @@reports = WSSreports.new(*report_settings)")
        WSSRuby.WriteLine("        @@ready = valid()")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    def valid()")
        WSSRuby.WriteLine("      if !(@@settings[:tag_unique] || @@settings[:tag_items] || @@settings[:exclude_items] || @@settings[:flag_unresponsive] || !@@settings[:custom_metadata].nil)")
        WSSRuby.WriteLine("        Log.out(" & """" & "ERROR - Script isn't configured to do anything." & """" & ", true)")
        WSSRuby.WriteLine("      else")
        WSSRuby.WriteLine("        if @@queries.queries.empty?")
        WSSRuby.WriteLine("          Log.out(" & """" & "ERROR - There are no queries." & """" & ", true)")
        WSSRuby.WriteLine("        else")
        WSSRuby.WriteLine("          Log.out(" & """" & "Ready to go…" & """" & ", true)")
        WSSRuby.WriteLine("          return true")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      return false")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    def empty_nil(string)")
        WSSRuby.WriteLine("      return !string.nil? && string.strip.empty?")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    def close()")
        WSSRuby.WriteLine("      Log.out(" & """" & "Closing…" & """" & ", nil) & """" & ")
        WSSRuby.WriteLine("      @@reports.close() if !@@reports.nil?")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  class Address_Queries_Expanded")
        WSSRuby.WriteLine("    attr_reader :queries")
        WSSRuby.WriteLine("    def initialize(location, properties)")
        WSSRuby.WriteLine("      @expanded_properties = properties")
        WSSRuby.WriteLine("      @queries = {}")
        WSSRuby.WriteLine("      load_csv(location).each do |tag, terms|")
        WSSRuby.WriteLine("        @queries[tag] = Regexp.union( terms.map!{|s| to_pattern(s)} )")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      Log.out(" & """" & "Queries: #{JSON.pretty_generate(@queries)}" & """" & ", nil)")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    def cache(comm)")
        WSSRuby.WriteLine("      return [] if comm.nil?")
        WSSRuby.WriteLine("      return [:get_from, :get_to, :get_cc, :get_bcc].map{|c| comm.send(c).to_a.map(&:to_rfc822_string)}.flatten")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    def compare(item)")
        WSSRuby.WriteLine("      search_data = cache( item.source_item.get_communication() )")
        WSSRuby.WriteLine("      @expanded_properties.each do |p|")
        WSSRuby.WriteLine("        if item.props.has_key?(p)")
        WSSRuby.WriteLine("          Log.out(" & """" & "#{item.source_item.get_name} has #{p}" & """" & ", nil) & """" & ")
        WSSRuby.WriteLine("          v = expanded_prop(item.props[p])")
        WSSRuby.WriteLine("          Log.out(" & """" & "Adding: #{v}" & """" & ", nil)")
        WSSRuby.WriteLine("          search_data << v")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      compare_data(item, search_data)")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    def compare_data(item, search_data)")
        WSSRuby.WriteLine("      if !search_data.empty?")
        WSSRuby.WriteLine("        search_string = search_data.join(" & """" & "; " & """" & ")")
        WSSRuby.WriteLine("        Log.out(" & """" & "Searching: #{search_string}" & """" & ", nil)")
        WSSRuby.WriteLine("        @queries.each do |tag, query|")
        WSSRuby.WriteLine("          next if item.flags.include?(tag)")
        WSSRuby.WriteLine("          item.flags.add(tag) if !search_string.match(query).nil?")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    def expanded_prop(obj)")
        WSSRuby.WriteLine("      return obj.to_a.map{ |e| expanded_prop(e) }.join(" & """" & "; " & """" & ") if obj.respond_to?(:each)")
        WSSRuby.WriteLine("      return obj.to_s")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    def load_csv(file_name)")
        WSSRuby.WriteLine("      h = Hash.new{|h,k| h[k]=[]}")
        WSSRuby.WriteLine("      if !file_name.nil? && File.exist?(file_name)")
        WSSRuby.WriteLine("        CSV.foreach(file_name) do |row|")
        WSSRuby.WriteLine("          tag = row[0]")
        WSSRuby.WriteLine("          query = row[1]")
        WSSRuby.WriteLine("          h[tag.to_sym] << query")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        Log.out(" & """" & "Loaded #{file_name}" & """" & ", true)")
        WSSRuby.WriteLine("      else")
        WSSRuby.WriteLine("        Log.out(" & """" & "No file found at: #{file_name}" & """" & ", true)")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      return h")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    def to_pattern(string)")
        WSSRuby.WriteLine("      if string.start_with?(" & """" & "/" & """" & ") && string.end_with?(" & """" & "/" & """" & ")")
        WSSRuby.WriteLine("        string = string.slice(1..-2)")
        WSSRuby.WriteLine("      else")
        WSSRuby.WriteLine("        string = Regexp.escape(string)")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      return Regexp.new( string, true )")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  class WSSreports")
        WSSRuby.WriteLine("    require 'fileutils'")
        WSSRuby.WriteLine("    attr_reader :reports")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    def initialize(report_directory, items, queries)")
        WSSRuby.WriteLine("      @reports = []")
        WSSRuby.WriteLine("      if items || queries")
        WSSRuby.WriteLine("        guid = ENV_JAVA[" & """" & "nuix.simple.worker.guid" & """" & "]")
        WSSRuby.WriteLine("        directory = File.join( report_directory , guid)")
        WSSRuby.WriteLine("        Log.out(" & """" & "Reporting to: #{directory}" & """" & ", true)")
        WSSRuby.WriteLine("        begin")
        WSSRuby.WriteLine("          FileUtils.mkdir_p directory")
        WSSRuby.WriteLine("          @reports << Item_Report.new(directory) if items")
        WSSRuby.WriteLine("          @reports << Query_Report.new(directory) if queries")
        WSSRuby.WriteLine("        rescue Errno::ENOENT => e")
        WSSRuby.WriteLine("          Log.out(" & """" & "ERROR - Failed to initialize reports: #{e}" & """" & ", true)")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      else")
        WSSRuby.WriteLine("        Log.out(" & """" & "No reports selected." & """" & ", true)")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    def add(wss_item)")
        WSSRuby.WriteLine("      @reports.each { |v| v.add(wss_item) }")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    def close()")
        WSSRuby.WriteLine("      @reports.each { |v| v.close }")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    class Report")
        WSSRuby.WriteLine("      def initialize(directory, name)")
        WSSRuby.WriteLine("        @name = name")
        WSSRuby.WriteLine("        file_location = File.join(directory, " & """" & "#{@name}.csv" & """" & ")")
        WSSRuby.WriteLine("        Log.out(" & """" & "Creating report #{file_location}" & """" & ", true)")
        WSSRuby.WriteLine("        @csv = CSV.open(file_location, " & """" & "w" & """" & ")")
        WSSRuby.WriteLine("        @csv << @header if !@header.nil?")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      def add(row)")
        WSSRuby.WriteLine("        @csv << row")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      def close")
        WSSRuby.WriteLine("        @csv.close")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    class Item_Report < Report")
        WSSRuby.WriteLine("      def initialize(directory)")
        WSSRuby.WriteLine("        @header = [" & """" & "GUID" & """" & ", " & """" & "MD5" & """" & ", " & """" & "Path" & """" & ", " & """" & "Flags" & """" & "]")
        WSSRuby.WriteLine("        super(directory, " & """" & "items" & """" & ")")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      def add(wss_item)")
        WSSRuby.WriteLine("        guid = wss_item.worker_item.get_guid_path.last")
        WSSRuby.WriteLine("        md5 = wss_item.worker_item.get_digests.get_md5")
        WSSRuby.WriteLine("        path = File.join(wss_item.source_item.get_path_names.to_a)")
        WSSRuby.WriteLine("        flags = wss_item.flags.to_a.sort.join(" & """" & ";" & """" & ")")
        WSSRuby.WriteLine("        super( [guid, md5, path, flags] )")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    class Query_Report < Report")
        WSSRuby.WriteLine("      def initialize(directory)")
        WSSRuby.WriteLine("        @header = [" & """" & "Flag" & """" & ", " & """" & "Count" & """" & "]")
        WSSRuby.WriteLine("        @queries = Hash.new(0)")
        WSSRuby.WriteLine("        @total = 0")
        WSSRuby.WriteLine("        super(directory, " & """" & "queries" & """" & ")")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      def add(wss_item)")
        WSSRuby.WriteLine("        @total += 1")
        WSSRuby.WriteLine("        wss_item.flags.each { |f| @queries[f] += 1 }")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("      def close()")
        WSSRuby.WriteLine("        Log.out(" & """" & "Reporting query results from #{@total} items" & """" & ", nil)")
        WSSRuby.WriteLine("        @csv << [" & """" & "TOTAL ITEMS" & """" & ", @total]")
        WSSRuby.WriteLine("        @queries.each { |k, v| @csv << [k.to_s, v] }")
        WSSRuby.WriteLine("        super()")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("module Log")
        WSSRuby.WriteLine("  @log = true")
        WSSRuby.WriteLine("  class << self")
        WSSRuby.WriteLine("    def set(b)")
        WSSRuby.WriteLine("      @log = b")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("    def out(str, level)")
        WSSRuby.WriteLine("      puts " & """" & "WSS: #{str}" & """" & " if ( level || @log )")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("class WSSitem")
        WSSRuby.WriteLine("  include WSS")
        WSSRuby.WriteLine("  attr_accessor :flags, :source_item, :worker_item")
        WSSRuby.WriteLine("  def initialize(worker_item)")
        WSSRuby.WriteLine("    @worker_item = worker_item")
        WSSRuby.WriteLine("    if @@ready")
        WSSRuby.WriteLine("      @flags = Set.new()")
        WSSRuby.WriteLine("      @source_item = worker_item.get_source_item()")
        WSSRuby.WriteLine("      @properties = nil")
        WSSRuby.WriteLine("      process()")
        WSSRuby.WriteLine("    else")
        WSSRuby.WriteLine("      Log.out(" & """" & "ERROR - Invalid Script Configuration" & """" & ", true)")
        WSSRuby.WriteLine("      @worker_item.set_process_item(false)")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  def category_value()")
        WSSRuby.WriteLine("    if !@@settings[:category_property].nil?")
        WSSRuby.WriteLine("      @@settings[:category_property].each do |category|")
        WSSRuby.WriteLine("        category_value = props()[category]")
        WSSRuby.WriteLine("        if !category_value.nil?")
        WSSRuby.WriteLine("          Log.out(" & """" & "Identified #{category}: #{category_value}" & """" & ", nil)")
        WSSRuby.WriteLine("          return category_value")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("    return nil")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  def custom_metadata_value()")
        WSSRuby.WriteLine("    values = @flags.to_a")
        WSSRuby.WriteLine("    category = category_value()")
        WSSRuby.WriteLine("    values.map!{|n| " & """" & "#{n}_#{category}" & """" & " } if !category.nil?")
        WSSRuby.WriteLine("    return values.join(" & """" & "|" & """" & ")")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  def filter()")
        WSSRuby.WriteLine("    return nil if @@settings[:file_filter].empty?")
        WSSRuby.WriteLine("    @@settings[:file_filter].each do |field, values|")
        WSSRuby.WriteLine("      case field")
        WSSRuby.WriteLine("      when :mime")
        WSSRuby.WriteLine("        v = @source_item.get_type().get_name()")
        WSSRuby.WriteLine("      when :kind")
        WSSRuby.WriteLine("        v = @source_item.get_kind().get_name()")
        WSSRuby.WriteLine("      when :ext")
        WSSRuby.WriteLine("        v = File.extname(@source_item.get_name).downcase[1..-1]")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("      return false if values.include?(v)")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("    return true")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  def process()")
        WSSRuby.WriteLine("    if filter()")
        WSSRuby.WriteLine("      Log.out(" & """" & "Filtered out #{@source_item.get_name} (#{@source_item.get_type().get_name()})" & """" & ", nil)")
        WSSRuby.WriteLine("      @worker_item.set_process_item(false) if @@settings[:exclude_items]")
        WSSRuby.WriteLine("    else")
        WSSRuby.WriteLine("      if @source_item.is_kind(" & """" & "container" & """" & ")")
        WSSRuby.WriteLine("        Log.out(" & """" & "Skipping container: #{@source_item.get_name}" & """" & ", nil)")
        WSSRuby.WriteLine("      else")
        WSSRuby.WriteLine("        Log.out(" & """" & "Querying #{@source_item.get_name}" & """" & ", nil)")
        WSSRuby.WriteLine("        @@queries.compare(self)")
        WSSRuby.WriteLine("        if !@flags.empty?")
        WSSRuby.WriteLine("          Log.out(" & """" & "Found #{@flags.size} matches for #{@source_item.get_name}:\n#{@flags.to_a.join(" & """" & ", " & """" & ")}" & """" & ", nil)")
        WSSRuby.WriteLine("          @worker_item.add_tag(" & """" & "unique|#{@flags.first.to_s}" & """" & ") if @@settings[:tag_unique] && match_num == 1")
        WSSRuby.WriteLine("        else")
        WSSRuby.WriteLine("          Log.out(" & """" & "No matches for #{@source_item.get_name}" & """" & ", nil)")
        WSSRuby.WriteLine("          if @@settings[:flag_unresponsive]")
        WSSRuby.WriteLine("            @flags.add(:unresponsive)")
        WSSRuby.WriteLine("          else")
        WSSRuby.WriteLine("            @worker_item.set_process_item(false) if @@settings[:exclude_items]")
        WSSRuby.WriteLine("            return false")
        WSSRuby.WriteLine("          end")
        WSSRuby.WriteLine("        end")
        WSSRuby.WriteLine("        @flags.each { |tag| @worker_item.add_tag(tag.to_s) } if @@settings[:tag_items]")
        WSSRuby.WriteLine("        @worker_item.add_custom_metadata(@@settings[:custom_metadata], custom_metadata_value(), " & """" & "text" & """" & ", " & """" & "user" & """" & ") if !@@settings[:custom_metadata].nil?")
        WSSRuby.WriteLine("        @@reports.add(self) if !@@reports.nil?")
        WSSRuby.WriteLine("        return true")
        WSSRuby.WriteLine("      end")
        WSSRuby.WriteLine("    end")
        WSSRuby.WriteLine("    return nil")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("")
        WSSRuby.WriteLine("  def props()")
        WSSRuby.WriteLine("    @properties = @source_item.get_properties() if @properties.nil?")
        WSSRuby.WriteLine("    return @properties")
        WSSRuby.WriteLine("  end")
        WSSRuby.WriteLine("end")

        WSSRuby.Close()
        blnBuildFinalWSSRubyScript = True
    End Function

    Public Function blnBuildFinalArchiveExtractionWSSJSon(ByVal sWSSJsonFileName As String, ByVal sArchiveType As String, ByVal lstWSSSettings As List(Of String), ByVal sProcessingType As String, ByVal sCaseDirectory As String) As Boolean
        Dim sMachineName As String

        Dim sFileFiltering As String
        Dim sFilterFileExt As String
        Dim sMappingCSV As String
        Dim sVerbose As String
        Dim sUnresponsive As String
        Dim sTagItems As String
        Dim sTagUnique As String
        Dim sExcludeItem As String
        Dim sFilterKinds As String
        Dim asEmailContent() As String
        Dim sEmailContent As String
        Dim asRSSFeedContent() As String
        Dim sRSSFeedContent As String
        Dim asCalendarContent() As String
        Dim sCalendarContent As String
        Dim asContactContent() As String
        Dim sContactContent As String
        Dim sMimeType As String
        Dim sKind As String
        Dim sExt As String
        Dim sSearchTermCSV As String
        Dim WSSJSonFile As StreamWriter

        blnBuildFinalArchiveExtractionWSSJSon = False

        For Each setting In lstWSSSettings
            If setting.ToString.Contains("Email:") Then
                asEmailContent = Split(setting.ToString, ":")
                sEmailContent = asEmailContent(1)
            ElseIf setting.ToString.Contains("RSSFeed") Then
                asRSSFeedContent = Split(setting.ToString, ":")
                sRSSFeedContent = asRSSFeedContent(1)
            ElseIf setting.ToString.Contains("Contact") Then
                asContactContent = Split(setting.ToString, ":")
                sContactContent = asContactContent(1)
            ElseIf setting.ToString.Contains("Calendar") Then
                asCalendarContent = Split(setting.ToString, ":")
                sCalendarCOntent = asCalendarContent(1)
            ElseIf setting.ToString.Contains("filterKinds:") Then
                sFilterKinds = setting.ToString.Replace("filterKinds:", "")
            ElseIf setting.ToString.Contains("filterFileExt:") Then
                sFilterFileExt = setting.ToString.Replace("filterFileExt:", "")
            ElseIf setting.ToString.Contains("searchTermCSV:") Then
                sSearchTermCSV = setting.ToString.Replace("searchTermCSV:", "")
            ElseIf setting.ToString.Contains("mappingCSV:") Then
                sMappingCSV = setting.ToString.Replace("mappingCSV:", "")
            ElseIf setting.ToString.Contains("excludeItems:") Then
                sExcludeItem = setting.ToString.Replace("excludeItems:", "")
            ElseIf setting.ToString.Contains("verbose:") Then
                sVerbose = setting.ToString.Replace("verbose:", "")
            ElseIf setting.ToString.Contains("flagUnresponsive:") Then
                sUnresponsive = setting.ToString.Replace("flagUnresponsive:", "")
            End If
        Next

        If (sEmailContent = "true") And (sRSSFeedContent = "true") And (sCalendarCOntent = "true") And (sContactContent = "true") Then
            sMimeType = vbNullString
            sKind = """" & "email" & """" & ", " & """" & "calendar" & """" & ", " & """" & "contact" & """" & ", " & """" & "container" & """"
            sExt = vbNullString
        ElseIf (sEmailContent = "false") And (sRSSFeedContent = "true") And (sCalendarCOntent = "true") And (sContactContent = "true") Then
            sMimeType = """" & "application/vnd.ms-outlook-item" & """"
            sKind = """" & "calendar" & """" & ", " & """" & "contact" & """" & ", " & """" & "container" & """"
            sExt = vbNullString
        ElseIf (sEmailContent = "true") And (sRSSFeedContent = "false") And (sCalendarCOntent = "true") And (sContactContent = "true") Then
            sMimeType = """" & "application/pdf-mail" & """" & ", " & """" & "application/vnd.aftermail-email" & """" & ", " & """" & "application/vnd.hp-trim-email" & """" & ", " & """" & "application/vnd.lotus-domino-xml-mail-document" & """" & ", " & """" & "application/vnd.lotus-notes-document" & """" & ", " & """" & "application/vnd.ms-outlook-mac-email" & """" & ", " & """" & "application/vnd.ms-outlook-note" & """" & ", " & """" & "application/vnd.ms-entourage-message" & """" & ", " & """" & "application/x-microsoft-restricted-permission-message" & """" & ", " & """" & "application/pcm-email" & """" & ", " & """" & "application/vnd.rim-blackberry-email" & """" & ", " & """" & "application/vnd.rim-blackberry-sms" & """" & ", " & """" & "application/vnd.rimarts-becky-email" & """" & ", " & """" & "message/news" & """" & ", " & """" & "message/rfc822" & """" & ", " & """" & "message/rfc822-headers" & """" & ", " & """" & "message/x-scraped" & """" & ", " & """" & "application/vnd.lotus-domino-xml-other-document" & """" & ", " & """" & "application/vnd.lotus-domino-xml-task-document" & """" & ", " & """" & "application/vnd.ms-outlook-activity" & """" & ", " & """" & "application/vnd.ms-outlook-journal" & """" & ", " & """" & "application/vnd.ms-outlook-mac-note" & """" & ", " & """" & "application/vnd.ms-outlook-mac-task" & """" & ", " & """" & "application/vnd.ms-outlook-task" & """" & ", " & """" & "application/vnd.ms-outlook-stickynote" & """" & ", " & """" & "application/vnd.ms-outlook-property-block" & """"
            sKind = """" & "calendar" & """" & ", " & """" & "contact" & """" & ", " & """" & "container" & """"
            sExt = vbNullString
        ElseIf (sEmailContent = "true") And (sRSSFeedContent = "true") And (sCalendarContent = "false") And (sContactContent = "true") Then
            sMimeType = ""
            sKind = """" & "email" & """" & ", " & """" & "contact" & """" & ", " & """" & "container" & """"
            sExt = ""
        ElseIf (sEmailContent = "true") And (sRSSFeedContent = "true") And (sCalendarContent = "true") And (sContactContent = "false") Then
            sMimeType = ""
            sKind = """" & "email" & """" & ", " & """" & "calendar" & """" & ", " & """" & "container" & """"
            sExt = ""
        ElseIf (sEmailContent = "true") And (sRSSFeedContent = "false") And (sCalendarContent = "false") And (sContactContent = "false") Then
            sMimeType = """" & "application/pdf-mail" & """" & ", " & """" & "application/vnd.aftermail-email" & """" & ", " & """" & "application/vnd.hp-trim-email" & """" & ", " & """" & "application/vnd.lotus-domino-xml-mail-document" & """" & ", " & """" & "application/vnd.lotus-notes-document" & """" & ", " & """" & "application/vnd.ms-outlook-mac-email" & """" & ", " & """" & "application/vnd.ms-outlook-note" & """" & ", " & """" & "application/vnd.ms-entourage-message" & """" & ", " & """" & "application/x-microsoft-restricted-permission-message" & """" & ", " & """" & "application/pcm-email" & """" & ", " & """" & "application/vnd.rim-blackberry-email" & """" & ", " & """" & "application/vnd.rim-blackberry-sms" & """" & ", " & """" & "application/vnd.rimarts-becky-email" & """" & ", " & """" & "message/news" & """" & ", " & """" & "message/rfc822" & """" & ", " & """" & "message/rfc822-headers" & """" & ", " & """" & "message/x-scraped" & """" & ", " & """" & "application/vnd.lotus-domino-xml-other-document" & """" & ", " & """" & "application/vnd.lotus-domino-xml-task-document" & """" & ", " & """" & "application/vnd.ms-outlook-activity" & """" & ", " & """" & "application/vnd.ms-outlook-journal" & """" & ", " & """" & "application/vnd.ms-outlook-mac-note" & """" & ", " & """" & "application/vnd.ms-outlook-mac-task" & """" & ", " & """" & "application/vnd.ms-outlook-task" & """" & ", " & """" & "application/vnd.ms-outlook-stickynote" & """" & ", " & """" & "application/vnd.ms-outlook-property-block" & """"
            sKind = """" & "container" & """"
            sExt = ""
        ElseIf (sEmailContent = "true") And (sRSSFeedContent = "true") And (sCalendarContent = "false") And (sContactContent = "false") Then
            sMimeType = ""
            sKind = """" & "email" & """" & ", " & """" & "container" & """"
            sExt = ""
        ElseIf (sEmailContent = "false") And (sRSSFeedContent = "true") And (sCalendarContent = "false") And (sContactContent = "false") Then
            sMimeType = """" & "application/vnd.ms-outlook-item" & """"
            sKind = """" & "container" & """"
            sExt = ""
        ElseIf (sEmailContent = "false") And (sRSSFeedContent = "false") And (sCalendarContent = "true") And (sContactContent = "true") Then
            sMimeType = ""
            sKind = """" & "calendar" & """" & ", " & """" & "contact" & """" & ", " & """" & "container" & """"
            sExt = ""
        ElseIf (sEmailContent = "false") And (sRSSFeedContent = "false") And (sCalendarContent = "false") And (sContactContent = "true") Then
            sMimeType = ""
            sKind = """" & "contact" & """" & ", " & """" & "container"
            sExt = ""
        ElseIf (sEmailContent = "false") And (sRSSFeedContent = "true") And (sCalendarContent = "false") And (sContactContent = "true") Then
            sMimeType = """" & "application/vnd.ms-outlook-item" & """"
            sKind = """" & "contact" & """" & ", " & """" & "container"
            sExt = ""
        ElseIf (sEmailContent = "false") And (sRSSFeedContent = "true") And (sCalendarContent = "true") And (sContactContent = "false") Then
            sMimeType = """" & "application/vnd.ms-outlook-item" & """"
            sKind = """" & "calendar" & """" & ", " & """" & "container" & """"
            sExt = ""
        ElseIf (sEmailContent = "true") And (sRSSFeedContent = "false") And (sCalendarContent = "false") And (sContactContent = "true") Then
            sMimeType = """" & "application/pdf-mail" & """" & ", " & """" & "application/vnd.aftermail-email" & """" & ", " & """" & "application/vnd.hp-trim-email" & """" & ", " & """" & "application/vnd.lotus-domino-xml-mail-document" & """" & ", " & """" & "application/vnd.lotus-notes-document" & """" & ", " & """" & "application/vnd.ms-outlook-mac-email" & """" & ", " & """" & "application/vnd.ms-outlook-note" & """" & ", " & """" & "application/vnd.ms-entourage-message" & """" & ", " & """" & "application/x-microsoft-restricted-permission-message" & """" & ", " & """" & "application/pcm-email" & """" & ", " & """" & "application/vnd.rim-blackberry-email" & """" & ", " & """" & "application/vnd.rim-blackberry-sms" & """" & ", " & """" & "application/vnd.rimarts-becky-email" & """" & ", " & """" & "message/news" & """" & ", " & """" & "message/rfc822" & """" & ", " & """" & "message/rfc822-headers" & """" & ", " & """" & "message/x-scraped" & """" & ", " & """" & "application/vnd.lotus-domino-xml-other-document" & """" & ", " & """" & "application/vnd.lotus-domino-xml-task-document" & """" & ", " & """" & "application/vnd.ms-outlook-activity" & """" & ", " & """" & "application/vnd.ms-outlook-journal" & """" & ", " & """" & "application/vnd.ms-outlook-mac-note" & """" & ", " & """" & "application/vnd.ms-outlook-mac-task" & """" & ", " & """" & "application/vnd.ms-outlook-task" & """" & ", " & """" & "application/vnd.ms-outlook-stickynote" & """" & ", " & """" & "application/vnd.ms-outlook-property-block" & """"
            sKind = """" & "contact" & """" & ", " & """" & "container" & """"
            sExt = ""
        ElseIf (sEmailContent = "true") And (sRSSFeedContent = "false") And (sCalendarContent = "true") And (sContactContent = "false") Then
            sMimeType = """" & "application/pdf-mail" & """" & ", " & """" & "application/vnd.aftermail-email" & """" & ", " & """" & "application/vnd.hp-trim-email" & """" & ", " & """" & "application/vnd.lotus-domino-xml-mail-document" & """" & ", " & """" & "application/vnd.lotus-notes-document" & """" & ", " & """" & "application/vnd.ms-outlook-mac-email" & """" & ", " & """" & "application/vnd.ms-outlook-note" & """" & ", " & """" & "application/vnd.ms-entourage-message" & """" & ", " & """" & "application/x-microsoft-restricted-permission-message" & """" & ", " & """" & "application/pcm-email" & """" & ", " & """" & "application/vnd.rim-blackberry-email" & """" & ", " & """" & "application/vnd.rim-blackberry-sms" & """" & ", " & """" & "application/vnd.rimarts-becky-email" & """" & ", " & """" & "message/news" & """" & ", " & """" & "message/rfc822" & """" & ", " & """" & "message/rfc822-headers" & """" & ", " & """" & "message/x-scraped" & """" & ", " & """" & "application/vnd.lotus-domino-xml-other-document" & """" & ", " & """" & "application/vnd.lotus-domino-xml-task-document" & """" & ", " & """" & "application/vnd.ms-outlook-activity" & """" & ", " & """" & "application/vnd.ms-outlook-journal" & """" & ", " & """" & "application/vnd.ms-outlook-mac-note" & """" & ", " & """" & "application/vnd.ms-outlook-mac-task" & """" & ", " & """" & "application/vnd.ms-outlook-task" & """" & ", " & """" & "application/vnd.ms-outlook-stickynote" & """" & ", " & """" & "application/vnd.ms-outlook-property-block" & """"
            sKind = """" & "calendar" & """" & ", " & """" & "container" & """"
            sExt = ""
        ElseIf (sEmailContent = "false") And (sRSSFeedContent = "false") And (sCalendarContent = "true") And (sContactContent = "false") Then
            sMimeType = ""
            sKind = """" & "calendar" & """" & ", " & """" & "container" & """"
            sExt = ""
        ElseIf (sEmailContent = "false") And (sRSSFeedContent = "false") And (sCalendarContent = "false") And (sContactContent = "false") Then
            sMimeType = ""
            sKind = ""
            sExt = ""
        End If

        If sFileFiltering = vbNullString Then
            sFileFiltering = "false"
        End If

        If sTagItems = vbNullString Then
            sTagItems = "false"
        End If

        If sTagUnique = vbNullString Then
            sTagUnique = "false"
        End If

        sMachineName = System.Net.Dns.GetHostName()
        WSSJSonFile = New StreamWriter(sWSSJsonFileName)

        Select Case sArchiveType

            Case "Veritas Enterprise Vault"
                WSSJSonFile.WriteLine("{")
                WSSJSonFile.WriteLine(" " & """" & "verbose" & """" & ": " & sVerbose & ",")
                If sExcludeItem = "true" Then
                    WSSJSonFile.WriteLine(" " & """" & "exclude_items" & """" & ": true,")
                    WSSJSonFile.WriteLine(" " & """" & "flag_unresponsive" & """" & ": false,")
                ElseIf (sExcludeItem = "false") Then
                    WSSJSonFile.WriteLine(" " & """" & "exclude_items" & """" & ": false,")
                    WSSJSonFile.WriteLine(" " & """" & "flag_unresponsive" & """" & ": true,")

                End If
                WSSJSonFile.WriteLine(" " & """" & "tag_items" & """" & ": false,")
                WSSJSonFile.WriteLine(" " & """" & "tag_unique" & """" & ": false,")
                WSSJSonFile.WriteLine(" " & """" & "custom_metadata" & """" & ": " & """" & "Mailboxes" & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "category_property" & """" & ": [],")
                WSSJSonFile.WriteLine(" " & """" & "file_filter" & """" & ": {")
                WSSJSonFile.WriteLine("     " & """" & "mime" & """" & ": [" & sMimeType & "],")
                WSSJSonFile.WriteLine("     " & """" & "kind" & """" & ": [" & sKind & "],")
                WSSJSonFile.WriteLine("     " & """" & "ext" & """" & ": [" & sExt & "]")
                WSSJSonFile.WriteLine("  },")
                WSSJSonFile.WriteLine(" " & """" & "csv_queries" & """" & ": " & """" & sSearchTermCSV.Replace("\", "\\") & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "expanded_properties" & """" & ": [" & """" & "Expanded DL" & """" & "],")
                WSSJSonFile.WriteLine(" " & """" & "report_directory" & """" & ": " & """" & Trim(sCaseDirectory.Replace("\", "\\")) & "\\WSSReports" & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "report_items" & """" & ": true,")
                WSSJSonFile.WriteLine(" " & """" & "report_queries" & """" & ": true")
                WSSJSonFile.WriteLine("}")
            Case "EMC EmailXtender"
                WSSJSonFile.WriteLine("{")
                WSSJSonFile.WriteLine(" " & """" & "verbose" & """" & ": " & sVerbose & ",")
                WSSJSonFile.WriteLine(" " & """" & "exclude_items" & """" & ": true,")
                If sExcludeItem = "true" Then
                    WSSJSonFile.WriteLine(" " & """" & "flag_unresponsive" & """" & ": false,")
                ElseIf (sExcludeItem = "false") Then
                    WSSJSonFile.WriteLine(" " & """" & "flag_unresponsive" & """" & ": true,")

                End If
                WSSJSonFile.WriteLine(" " & """" & "tag_items" & """" & ": false,")
                WSSJSonFile.WriteLine(" " & """" & "tag_unique" & """" & ": false,")
                WSSJSonFile.WriteLine(" " & """" & "custom_metadata" & """" & ": " & """" & "Mailboxes" & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "category_property" & """" & ": [],")
                WSSJSonFile.WriteLine(" " & """" & "file_filter" & """" & ": {")
                WSSJSonFile.WriteLine("     " & """" & "mime" & """" & ": [" & sMimeType & "],")
                WSSJSonFile.WriteLine("     " & """" & "kind" & """" & ": [" & sKind & "],")
                WSSJSonFile.WriteLine("     " & """" & "ext" & """" & ": [" & sExt & "]")
                WSSJSonFile.WriteLine("  },")
                WSSJSonFile.WriteLine(" " & """" & "csv_queries" & """" & ": " & """" & sSearchTermCSV.Replace("\", "\\") & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "expanded_properties" & """" & ": [" & """" & "Expanded DL" & """" & "],")
                WSSJSonFile.WriteLine(" " & """" & "report_directory" & """" & ": " & """" & Trim(sCaseDirectory.Replace("\", "\\")) & "\\WSSReports" & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "report_items" & """" & ": true,")
                WSSJSonFile.WriteLine(" " & """" & "report_queries" & """" & ": true")
                WSSJSonFile.WriteLine("}")
            Case "EMC SourceOne"
                WSSJSonFile.WriteLine("{")
                WSSJSonFile.WriteLine(" " & """" & "verbose" & """" & ": " & sVerbose & ",")
                If sExcludeItem = "true" Then
                    WSSJSonFile.WriteLine(" " & """" & "exclude_items" & """" & ": true,")
                    WSSJSonFile.WriteLine(" " & """" & "flag_unresponsive" & """" & ": false,")
                ElseIf (sExcludeItem = "false") Then
                    WSSJSonFile.WriteLine(" " & """" & "exclude_items" & """" & ": false,")
                    WSSJSonFile.WriteLine(" " & """" & "flag_unresponsive" & """" & ": true,")

                End If
                WSSJSonFile.WriteLine(" " & """" & "tag_items" & """" & ": false,")
                WSSJSonFile.WriteLine(" " & """" & "tag_unique" & """" & ": false,")
                WSSJSonFile.WriteLine(" " & """" & "custom_metadata" & """" & ": " & """" & "Mailboxes" & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "category_property" & """" & ": [],")
                WSSJSonFile.WriteLine(" " & """" & "file_filter" & """" & ": {")
                WSSJSonFile.WriteLine("     " & """" & "mime" & """" & ": [" & sMimeType & "],")
                WSSJSonFile.WriteLine("     " & """" & "kind" & """" & ": [" & sKind & "],")
                WSSJSonFile.WriteLine("     " & """" & "ext" & """" & ": [" & sExt & "]")
                WSSJSonFile.WriteLine("  },")
                WSSJSonFile.WriteLine(" " & """" & "csv_queries" & """" & ": " & """" & sSearchTermCSV.Replace("\", "\\") & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "expanded_properties" & """" & ": [" & """" & "Expanded DL" & """" & "],")
                WSSJSonFile.WriteLine(" " & """" & "report_directory" & """" & ": " & """" & Trim(sCaseDirectory.Replace("\", "\\")) & "\\WSSReports" & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "report_items" & """" & ": true,")
                WSSJSonFile.WriteLine(" " & """" & "report_queries" & """" & ": true")
                WSSJSonFile.WriteLine("}")
            Case "HP/Autonomy EAS"
                WSSJSonFile.WriteLine("{")
                WSSJSonFile.WriteLine(" " & """" & "verbose" & """" & ": " & sVerbose & ",")
                If sExcludeItem = "true" Then
                    WSSJSonFile.WriteLine(" " & """" & "exclude_items" & """" & ": true,")
                    WSSJSonFile.WriteLine(" " & """" & "flag_unresponsive" & """" & ": false,")
                ElseIf (sExcludeItem = "false") Then
                    WSSJSonFile.WriteLine(" " & """" & "exclude_items" & """" & ": false,")
                    WSSJSonFile.WriteLine(" " & """" & "flag_unresponsive" & """" & ": true,")

                End If
                WSSJSonFile.WriteLine(" " & """" & "tag_items" & """" & ": false,")
                WSSJSonFile.WriteLine(" " & """" & "tag_unique" & """" & ": false,")
                WSSJSonFile.WriteLine(" " & """" & "custom_metadata" & """" & ": " & """" & "Mailboxes" & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "category_property" & """" & ": [],")
                WSSJSonFile.WriteLine(" " & """" & "file_filter" & """" & ": {")
                WSSJSonFile.WriteLine("     " & """" & "mime" & """" & ": [" & """" & sMimeType & """" & "],")
                WSSJSonFile.WriteLine("     " & """" & "kind" & """" & ": [" & """" & "email" & """" & ", " & """" & "calendar" & """" & ", " & """" & "contact" & """" & ", " & """" & "container" & """" & "],")
                WSSJSonFile.WriteLine("     " & """" & "ext" & """" & ": []")
                WSSJSonFile.WriteLine("  },")
                WSSJSonFile.WriteLine(" " & """" & "csv_queries" & """" & ": " & """" & sSearchTermCSV.Replace("\", "\\") & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "expanded_properties" & """" & ": [" & """" & "Expanded DL" & """" & "],")
                WSSJSonFile.WriteLine(" " & """" & "report_directory" & """" & ": " & """" & Trim(sCaseDirectory.Replace("\", "\\")) & "\\WSSReports" & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "report_items" & """" & ": true,")
                WSSJSonFile.WriteLine(" " & """" & "report_queries" & """" & ": true")
                WSSJSonFile.WriteLine("}")
            Case "Daegis AXS-One"
                WSSJSonFile.WriteLine("{")
                WSSJSonFile.WriteLine(" " & """" & "verbose" & """" & ": " & sVerbose & ",")
                If sExcludeItem = "true" Then
                    WSSJSonFile.WriteLine(" " & """" & "exclude_items" & """" & ": true,")
                    WSSJSonFile.WriteLine(" " & """" & "flag_unresponsive" & """" & ": false,")
                ElseIf (sExcludeItem = "false") Then
                    WSSJSonFile.WriteLine(" " & """" & "exclude_items" & """" & ": false,")
                    WSSJSonFile.WriteLine(" " & """" & "flag_unresponsive" & """" & ": true,")

                End If
                WSSJSonFile.WriteLine(" " & """" & "tag_items" & """" & ": false,")
                WSSJSonFile.WriteLine(" " & """" & "tag_unique" & """" & ": false,")
                WSSJSonFile.WriteLine(" " & """" & "custom_metadata" & """" & ": " & """" & "Mailboxes" & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "category_property" & """" & ": [],")
                WSSJSonFile.WriteLine(" " & """" & "file_filter" & """" & ": {")
                If sMimeType = vbNullString Then
                    sMimeType = """" & "application/vnd.axs-one-pgi-email-archive" & """"
                Else
                    sMimeType = """" & "application/vnd.axs-one-pgi-email-archive" & """" & ", " & sMimeType
                End If
                WSSJSonFile.WriteLine("     " & """" & "mime" & """" & ": [" & sMimeType & "],")
                WSSJSonFile.WriteLine("     " & """" & "kind" & """" & ": [" & sKind & "],")
                WSSJSonFile.WriteLine("     " & """" & "ext" & """" & ": [" & """" & ".pgi" & """" & "]")
                WSSJSonFile.WriteLine("  },")
                WSSJSonFile.WriteLine(" " & """" & "csv_queries" & """" & ": " & """" & sSearchTermCSV.Replace("\", "\\") & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "expanded_properties" & """" & ": [" & """" & "AXS-One-Mailbox" & """" & ", " & """" & "AXS-One-mailbox" & """" & "],")
                WSSJSonFile.WriteLine(" " & """" & "report_directory" & """" & ": " & """" & Trim(sCaseDirectory.Replace("\", "\\")) & "\\WSSReports" & """" & ",")
                WSSJSonFile.WriteLine(" " & """" & "report_items" & """" & ": true,")
                WSSJSonFile.WriteLine(" " & """" & "report_queries" & """" & ": true")
                WSSJSonFile.WriteLine("}")
        End Select

        WSSJSonFile.Close()

        blnBuildFinalArchiveExtractionWSSJSon = True
    End Function


    Public Function blnBuildArchiveUserExtractionRubyScript(ByVal sArchiveName As String, ByVal sProcessingFileDir As String, ByVal sSourceLocations As String, sRubyScriptFileName As String, ByVal sEvidenceJSon As String, ByVal sCaseDir As String, ByVal sSQLiteDBLocation As String, ByVal sNumberOfWorkers As String, ByVal sMemoryPerWorker As String, ByVal iEvidenceSize As Double, ByVal sWorkerTempDir As String, ByVal bExtractFromSlackSpace As Boolean, ByVal sSQLAuthentication As String) As Boolean
        blnBuildArchiveUserExtractionRubyScript = False

        Dim asSourceLocations() As String
        Dim ArchiveUserExtractionRuby As StreamWriter

        ArchiveUserExtractionRuby = New StreamWriter(sRubyScriptFileName)
        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("# Menu Title: NEAMM WSS Extraction")
        ArchiveUserExtractionRuby.WriteLine("# Needs Selected Items: false")
        ArchiveUserExtractionRuby.WriteLine("# ")
        ArchiveUserExtractionRuby.WriteLine("# This script expects a JSON configured with archive parameters completed in order to automatically process data inside from a legacy email archive. ")
        ArchiveUserExtractionRuby.WriteLine("# ")
        ArchiveUserExtractionRuby.WriteLine("# Version 2.0")
        ArchiveUserExtractionRuby.WriteLine("# March 22 2017 - Alex Chatzistamatis, Nuix")
        ArchiveUserExtractionRuby.WriteLine("# ")
        ArchiveUserExtractionRuby.WriteLine("#######################################")
        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("require 'thread'")
        ArchiveUserExtractionRuby.WriteLine("require 'json'")
        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("CALLBACK_FREQUENCY = 100")
        ArchiveUserExtractionRuby.WriteLine("callback_count = 0")
        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("load """ & sSQLiteDBLocation.Replace("\", "\\") & "\\Database.rb_""")
        ArchiveUserExtractionRuby.WriteLine("load """ & sSQLiteDBLocation.Replace("\", "\\") & "\\SQLite.rb_""")
        ArchiveUserExtractionRuby.WriteLine("db = SQLite.new(""" & sSQLiteDBLocation.Replace("\", "\\") & "\\NuixEmailArchiveMigrationManager.db3" & """" & ")")
        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("worker_side_script = " & """" & sProcessingFileDir.Replace("\", "\\") & "\\wss.rb" & """")
        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("#######################################")
        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("file = File.read('" & sEvidenceJSon.Replace("\", "\\") & "')")
        ArchiveUserExtractionRuby.WriteLine("parsed = JSON.parse(file)")
        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("    archive_file = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_file" & """" & "]")
        ArchiveUserExtractionRuby.WriteLine("    archive_keystore = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_keystore" & """" & "]")
        ArchiveUserExtractionRuby.WriteLine("    archive_password = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_password" & """" & "]")
        ArchiveUserExtractionRuby.WriteLine("    archive_name = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_name" & """" & "]")
        ArchiveUserExtractionRuby.WriteLine("    archive_centera = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera" & """" & "]")
        ArchiveUserExtractionRuby.WriteLine("    archive_centera_ip = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera_ip" & """" & "]")
        ArchiveUserExtractionRuby.WriteLine("    archive_centera_clip = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera_clip" & """" & "]")
        ArchiveUserExtractionRuby.WriteLine("    archive_source_size = parsed[" & """" & "email_archive" & """" & "][" & """" & "source_size" & """" & "]")
        ArchiveUserExtractionRuby.WriteLine("    archive_migration = parsed[" & """" & "email_archive" & """" & "][" & """" & "migration" & """" & "]")
        ArchiveUserExtractionRuby.WriteLine("caseFactory = $utilities.getCaseFactory()")
        ArchiveUserExtractionRuby.WriteLine("case_settings = {")
        ArchiveUserExtractionRuby.WriteLine("    :compound => false,")
        ArchiveUserExtractionRuby.WriteLine("    :name => " & """" & "#{archive_name}" & """" & ",")
        ArchiveUserExtractionRuby.WriteLine("    :description => " & """" & "Created using Nuix Email Archive Migration Manager" & """" & ",")
        ArchiveUserExtractionRuby.WriteLine("    :investigator => " & """" & "NEAMM WSS Extraction" & """")
        ArchiveUserExtractionRuby.WriteLine("}")
        ArchiveUserExtractionRuby.WriteLine("$current_case = caseFactory.create(" & """" & sCaseDir.Replace("\", "\\") & """" & ", case_settings)")
        ArchiveUserExtractionRuby.WriteLine("processor = $current_case.createProcessor")
        ArchiveUserExtractionRuby.WriteLine("if archive_migration == " & """" & "lightspeed" & """")
        ArchiveUserExtractionRuby.WriteLine("	processing_settings = {")
        ArchiveUserExtractionRuby.WriteLine("		:traversalScope => " & """" & "full_traversal" & """" & ",")
        ArchiveUserExtractionRuby.WriteLine("		:analysisLanguage => " & """" & "en" & """" & ",")
        ArchiveUserExtractionRuby.WriteLine("		:identifyPhysicalFiles => true,")
        ArchiveUserExtractionRuby.WriteLine("		:reuseEvidenceStores => true,")
        If iEvidenceSize = 0 Then
            ArchiveUserExtractionRuby.WriteLine("			:reportProcessingStatus => " & """" & "none" & """" & ",")
        Else
            ArchiveUserExtractionRuby.WriteLine("			:reportProcessingStatus => " & """" & "physical_files" & """" & ",")
        End If

        ArchiveUserExtractionRuby.WriteLine("		:workerItemCallback => " & """" & "ruby:#{IO.read(worker_side_script)}" & """")
        ArchiveUserExtractionRuby.WriteLine("	}")
        ArchiveUserExtractionRuby.WriteLine("	processor.setProcessingSettings(processing_settings)")
        ArchiveUserExtractionRuby.WriteLine("end")
        ArchiveUserExtractionRuby.WriteLine("parallel_processing_settings = {")
        ArchiveUserExtractionRuby.WriteLine("	:workerCount => " & sNumberOfWorkers & ",")
        ArchiveUserExtractionRuby.WriteLine("	:workerMemory => " & sMemoryPerWorker & ",")
        ArchiveUserExtractionRuby.WriteLine("	:embedBroker => true,")
        ArchiveUserExtractionRuby.WriteLine("	:brokerMemory => " & sMemoryPerWorker & ",")
        ArchiveUserExtractionRuby.WriteLine("   :workerTemp => " & """" & sWorkerTempDir.Replace("\", "\\") & """")
        ArchiveUserExtractionRuby.WriteLine("}")
        ArchiveUserExtractionRuby.WriteLine("processor.setParallelProcessingSettings(parallel_processing_settings)")
        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("if archive_centera == " & """" & "no" & """")
        ArchiveUserExtractionRuby.WriteLine("	evidence_archive = archive_name")
        ArchiveUserExtractionRuby.WriteLine("	evidence_container = processor.newEvidenceContainer(evidence_archive)")
        If sSourceLocations.Contains("|") Then
            asSourceLocations = Split(sSourceLocations, "|")
            For iCounter = 0 To UBound(asSourceLocations)
                If asSourceLocations(iCounter) <> vbNullString Then
                    ArchiveUserExtractionRuby.WriteLine("	evidence_container.addFile(" & """" & asSourceLocations(iCounter).Replace("\", "\\") & """" & ")")
                End If
            Next
        Else
            ArchiveUserExtractionRuby.WriteLine("	evidence_container.addFile(" & """" & sSourceLocations.Replace("\", "\\") & """" & ")")
        End If
        ArchiveUserExtractionRuby.WriteLine("	evidence_container.setEncoding(" & """" & "utf-8" & """" & ")")
        ArchiveUserExtractionRuby.WriteLine("	evidence_container.save")
        ArchiveUserExtractionRuby.WriteLine("end")

        ArchiveUserExtractionRuby.WriteLine("if archive_centera == " & """" & "yes" & """")
        ArchiveUserExtractionRuby.WriteLine("	evidence_archive = archive_name")
        ArchiveUserExtractionRuby.WriteLine("	evidence_container = processor.newEvidenceContainer(evidence_archive)")
        ArchiveUserExtractionRuby.WriteLine("	centera_cluster_settings = {")
        ArchiveUserExtractionRuby.WriteLine("		:ipsFile => " & """" & "#{archive_centera_ip}" & """" & ",")
        ArchiveUserExtractionRuby.WriteLine("		:clipsFile => " & """" & "#{archive_centera_clip}" & """" & ",")
        ArchiveUserExtractionRuby.WriteLine("	}")
        ArchiveUserExtractionRuby.WriteLine("	evidence_container.addCenteraCluster(centera_cluster_settings)")
        ArchiveUserExtractionRuby.WriteLine("	evidence_container.setEncoding(" & """" & "utf-8" & """" & ")")
        ArchiveUserExtractionRuby.WriteLine("	evidence_container.save")
        ArchiveUserExtractionRuby.WriteLine("end")

        ArchiveUserExtractionRuby.WriteLine("   id_password = archive_password")
        ArchiveUserExtractionRuby.WriteLine("	id_file = archive_keystore")
        If sArchiveName = "NSF" Then
            ArchiveUserExtractionRuby.WriteLine("   processor.addKeyStore(id_file,{")
            ArchiveUserExtractionRuby.WriteLine("        :filePassword =>id_password,")
            ArchiveUserExtractionRuby.WriteLine("        :target => " & """" & "" & """")
            ArchiveUserExtractionRuby.WriteLine("        })")
        End If

        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("start_time = Time.now")
        ArchiveUserExtractionRuby.WriteLine("last_progress = Time.now")
        ArchiveUserExtractionRuby.WriteLine("semaphore = Mutex.new")
        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("puts")
        ArchiveUserExtractionRuby.WriteLine("puts " & """" & "Email Archive Extraction for #{archive_name} started at #{start_time}..." & """")
        ArchiveUserExtractionRuby.WriteLine("puts")
        ArchiveUserExtractionRuby.WriteLine("printf " & """" & "\r%-40s %-25s %-25s %-25s %-25s" & """" & ", " & """" & "Timestamp" & """" & ", " & """" & "Processed Items" & """" & ", " & """" & "Current Bytes" & """" & ", " & """" & "Total Bytes" & """" & ", " & """" & "Percent (%) Completed" & """")
        ArchiveUserExtractionRuby.WriteLine("puts")
        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("if archive_source_size == " & """" & "0" & """")
        ArchiveUserExtractionRuby.WriteLine("processor.when_progress_updated do |progress|")
        ArchiveUserExtractionRuby.WriteLine("	semaphore.synchronize {")
        ArchiveUserExtractionRuby.WriteLine("		class Numeric")
        ArchiveUserExtractionRuby.WriteLine("    def percent_of(n)")
        ArchiveUserExtractionRuby.WriteLine("        self.to_f / n.to_f * 100")
        ArchiveUserExtractionRuby.WriteLine("      end")
        ArchiveUserExtractionRuby.WriteLine("    end")
        ArchiveUserExtractionRuby.WriteLine("    # Progress message every 15 seconds")
        ArchiveUserExtractionRuby.WriteLine("current_size = progress.get_current_size")
        ArchiveUserExtractionRuby.WriteLine("total_size = progress.get_total_size")
        ArchiveUserExtractionRuby.WriteLine("percent_completed = current_size.percent_of(total_size).round(1)")
        ArchiveUserExtractionRuby.WriteLine("  if callback_count % CALLBACK_FREQUENCY == 0")
        ArchiveUserExtractionRuby.WriteLine("       last_progress = Time.now")
        ArchiveUserExtractionRuby.WriteLine("  		printf " & """" & "\r%-40s %-25s %-25s %-25s %-25s" & """" & ", Time.now, callback_count, current_size, total_size,  0")
        ArchiveUserExtractionRuby.WriteLine("       updated_callback = [callback_count, current_size, " & """" & "0" & """" & "]")
        ArchiveUserExtractionRuby.WriteLine("       db.update(" & """" & "UPDATE archiveExtractionStats SET ItemsProcessed = ?, BytesProcessed = ?, PercentCompleted = ? WHERE BatchName = '#{archive_name}'" & """" & ",updated_callback)")
        ArchiveUserExtractionRuby.WriteLine("		end")
        ArchiveUserExtractionRuby.WriteLine("  }")
        ArchiveUserExtractionRuby.WriteLine("  callback_count += 1")
        ArchiveUserExtractionRuby.WriteLine("  end")
        ArchiveUserExtractionRuby.WriteLine("end")
        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("if archive_source_size != " & """" & "0" & """")
        ArchiveUserExtractionRuby.WriteLine("processor.when_progress_updated do |progress|")
        ArchiveUserExtractionRuby.WriteLine("    semaphore.synchronize {")
        ArchiveUserExtractionRuby.WriteLine("    class Numeric")
        ArchiveUserExtractionRuby.WriteLine("    def percent_of(n)")
        ArchiveUserExtractionRuby.WriteLine("        self.to_f / n.to_f * 100")
        ArchiveUserExtractionRuby.WriteLine("      end")
        ArchiveUserExtractionRuby.WriteLine("    end")
        ArchiveUserExtractionRuby.WriteLine("    # Progress message every 15 seconds")
        ArchiveUserExtractionRuby.WriteLine("    current_size = progress.get_current_size")
        ArchiveUserExtractionRuby.WriteLine("    total_size = progress.get_total_size")
        ArchiveUserExtractionRuby.WriteLine("    percent_completed = current_size.percent_of(total_size).round(1)")
        ArchiveUserExtractionRuby.WriteLine("  if callback_count % CALLBACK_FREQUENCY == 0")
        ArchiveUserExtractionRuby.WriteLine("       last_progress = Time.now")
        ArchiveUserExtractionRuby.WriteLine("  		printf " & """" & "\r%-40s %-25s %-25s %-25s %-25s" & """" & ", Time.now, callback_count, current_size, total_size, percent_completed")
        ArchiveUserExtractionRuby.WriteLine("		updated_callback = [callback_count, current_size, percent_completed]")
        ArchiveUserExtractionRuby.WriteLine("       db.update(" & """" & "UPDATE archiveExtractionStats SET ItemsProcessed = ?, BytesProcessed = ?, PercentCompleted = ? WHERE BatchName = '#{archive_name}'" & """" & ",updated_callback)")
        ArchiveUserExtractionRuby.WriteLine("	end")
        ArchiveUserExtractionRuby.WriteLine("	}")
        ArchiveUserExtractionRuby.WriteLine("   callback_count += 1")
        ArchiveUserExtractionRuby.WriteLine("  end")
        ArchiveUserExtractionRuby.WriteLine("end")
        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("processor.process")
        ArchiveUserExtractionRuby.WriteLine("puts")
        ArchiveUserExtractionRuby.WriteLine("end_time = Time.now")
        ArchiveUserExtractionRuby.WriteLine("total_time = '%.2f' % ((end_time-start_time)/60)")
        ArchiveUserExtractionRuby.WriteLine("display_time =  '%.0f' % (total_time)")
        ArchiveUserExtractionRuby.WriteLine("length = 60 / display_time.to_i")
        ArchiveUserExtractionRuby.WriteLine("size = length.to_i * archive_source_size.to_i")
        ArchiveUserExtractionRuby.WriteLine("speed = size.to_i / 1024 / 1024 / 1024.round(2)")
        ArchiveUserExtractionRuby.WriteLine("final = speed.round(2)")
        ArchiveUserExtractionRuby.WriteLine("puts")
        ArchiveUserExtractionRuby.WriteLine("puts " & """" & "Email Archive Extraction for #{archive_name} finished at #{end_time}..." & """")
        ArchiveUserExtractionRuby.WriteLine("puts")
        ArchiveUserExtractionRuby.WriteLine("puts " & """" & "Completed in #{display_time} minutes and averaged #{final} GB/hr." & """")
        ArchiveUserExtractionRuby.WriteLine("puts")
        ArchiveUserExtractionRuby.WriteLine("updated_callback = [" & """" & "Completed" & """" & ", callback_count, archive_source_size, 100, end_time]")
        ArchiveUserExtractionRuby.WriteLine("db.update(" & """" & "UPDATE archiveExtractionStats SET ExtractionStatus = ?, ItemsProcessed = ?, BytesProcessed = ?, PercentCompleted = ?,ProcessEndTIme = ? WHERE BatchName = '#{archive_name}'" & """" & ",updated_callback)")
        ArchiveUserExtractionRuby.WriteLine("")
        ArchiveUserExtractionRuby.WriteLine("$current_case.close")
        ArchiveUserExtractionRuby.WriteLine("return")
        ArchiveUserExtractionRuby.Close()

        blnBuildArchiveUserExtractionRubyScript = True

    End Function

    Public Function blnBuildArchiveFlatExtractionRubyScript(ByVal sArchiveName As String, ByVal sProcessingFileDir As String, ByVal sSoureLocations As String, ByVal sEvidenceJSoN As String, ByVal sCaseDir As String, ByVal sRubyScriptFileName As String, ByVal sSQLiteDBLocation As String, ByVal sNumberOfWorkers As String, ByVal sMemoryPerWorker As String, ByVal iEvidenceSize As Double, ByVal sWorkerTempDir As String, ByVal bExtractFromSlackSpace As Boolean, ByVal sSQLAuthentication As String) As Boolean
        blnBuildArchiveFlatExtractionRubyScript = False

        Dim ArchiveFlatExtractionRuby As StreamWriter
        Dim asSourceLocations() As String

        ArchiveFlatExtractionRuby = New StreamWriter(sRubyScriptFileName)

        ArchiveFlatExtractionRuby.WriteLine("# Menu Title: NEAMM Flat Archive Extraction")
        ArchiveFlatExtractionRuby.WriteLine("# Needs Selected Items: false")
        ArchiveFlatExtractionRuby.WriteLine("# ")
        ArchiveFlatExtractionRuby.WriteLine("# This script expects a JSON configured with archive parameters completed in order to automatically process data inside from a legacy email archive. ")
        ArchiveFlatExtractionRuby.WriteLine("# ")
        ArchiveFlatExtractionRuby.WriteLine("# Version 2.0")
        ArchiveFlatExtractionRuby.WriteLine("# March 22 2017 - Alex Chatzistamatis, Nuix")
        ArchiveFlatExtractionRuby.WriteLine("# ")
        ArchiveFlatExtractionRuby.WriteLine("#######################################")
        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("require 'thread'")
        ArchiveFlatExtractionRuby.WriteLine("require 'json'")
        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("CALLBACK_FREQUENCY = 100")
        ArchiveFlatExtractionRuby.WriteLine("callback_count = 0")
        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("")

        ArchiveFlatExtractionRuby.WriteLine("load """ & sSQLiteDBLocation.Replace("\", "\\") & "\\Database.rb_""")
        ArchiveFlatExtractionRuby.WriteLine("load """ & sSQLiteDBLocation.Replace("\", "\\") & "\\SQLite.rb_""")
        ArchiveFlatExtractionRuby.WriteLine("db = SQLite.new(""" & sSQLiteDBLocation.Replace("\", "\\") & "\\NuixEmailArchiveMigrationManager.db3" & """" & ")")
        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("#######################################")
        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("file = File.read('" & sEvidenceJSoN.Replace("\", "\\") & "')")
        ArchiveFlatExtractionRuby.WriteLine("parsed = JSON.parse(file)")
        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("    archive_file = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_file" & """" & "]")
        ArchiveFlatExtractionRuby.WriteLine("    archive_keystore = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_keystore" & """" & "]")
        ArchiveFlatExtractionRuby.WriteLine("    archive_password = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_password" & """" & "]")
        ArchiveFlatExtractionRuby.WriteLine("    archive_name = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_name" & """" & "]")
        ArchiveFlatExtractionRuby.WriteLine("    archive_centera = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera" & """" & "]")
        ArchiveFlatExtractionRuby.WriteLine("    archive_centera_ip = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera_ip" & """" & "]")
        ArchiveFlatExtractionRuby.WriteLine("    archive_centera_clip = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera_clip" & """" & "]")
        ArchiveFlatExtractionRuby.WriteLine("    archive_source_size = parsed[" & """" & "email_archive" & """" & "][" & """" & "source_size" & """" & "]")
        ArchiveFlatExtractionRuby.WriteLine("    archive_migration = parsed[" & """" & "email_archive" & """" & "][" & """" & "migration" & """" & "]")
        ArchiveFlatExtractionRuby.WriteLine("caseFactory = $utilities.getCaseFactory()")
        ArchiveFlatExtractionRuby.WriteLine("case_settings = {")
        ArchiveFlatExtractionRuby.WriteLine("    :compound => false,")
        ArchiveFlatExtractionRuby.WriteLine("    :name => " & """" & "#{archive_name}" & """" & ",")
        ArchiveFlatExtractionRuby.WriteLine("    :description => " & """" & "Created using Nuix Email Archive Migration Manager" & """" & ",")
        ArchiveFlatExtractionRuby.WriteLine("    :investigator => " & """" & "NEAMM Flat Archive Extraction" & """")
        ArchiveFlatExtractionRuby.WriteLine("}")
        ArchiveFlatExtractionRuby.WriteLine("$current_case = caseFactory.create(" & """" & sCaseDir.Replace("\", "\\") & """" & ", case_settings)")
        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("processor = $current_case.createProcessor")
        ArchiveFlatExtractionRuby.WriteLine("if archive_migration == " & """" & "lightspeed" & """")
        ArchiveFlatExtractionRuby.WriteLine("	processing_settings = {")
        ArchiveFlatExtractionRuby.WriteLine("		:traversalScope => " & """" & "full_traversal" & """" & ",")
        ArchiveFlatExtractionRuby.WriteLine("		:analysisLanguage => " & """" & "en" & """" & ",")
        ArchiveFlatExtractionRuby.WriteLine("		:identifyPhysicalFiles => true,")
        ArchiveFlatExtractionRuby.WriteLine("		:reuseEvidenceStores => true,")
        If iEvidenceSize = 0 Then
            ArchiveFlatExtractionRuby.WriteLine("			:reportProcessingStatus => " & """" & "none" & """" & ",")
        Else
            ArchiveFlatExtractionRuby.WriteLine("			:reportProcessingStatus => " & """" & "physical_files" & """" & ",")
        End If
        ArchiveFlatExtractionRuby.WriteLine("	}")
        ArchiveFlatExtractionRuby.WriteLine("	processor.setProcessingSettings(processing_settings)")
        ArchiveFlatExtractionRuby.WriteLine("end")
        ArchiveFlatExtractionRuby.WriteLine("parallel_processing_settings = {")
        ArchiveFlatExtractionRuby.WriteLine("	:workerCount => " & sNumberOfWorkers & ",")
        ArchiveFlatExtractionRuby.WriteLine("	:workerMemory => " & sMemoryPerWorker & ",")
        ArchiveFlatExtractionRuby.WriteLine("	:embedBroker => true,")
        ArchiveFlatExtractionRuby.WriteLine("	:brokerMemory => " & sMemoryPerWorker & ",")
        ArchiveFlatExtractionRuby.WriteLine("   :workerTemp => " & """" & sWorkerTempDir.Replace("\", "\\") & """")
        ArchiveFlatExtractionRuby.WriteLine("}")
        ArchiveFlatExtractionRuby.WriteLine("processor.setParallelProcessingSettings(parallel_processing_settings)")
        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("if archive_centera == " & """" & "no" & """")
        ArchiveFlatExtractionRuby.WriteLine("	evidence_archive = archive_name")
        ArchiveFlatExtractionRuby.WriteLine("	evidence_container = processor.newEvidenceContainer(evidence_archive)")
        If sSoureLocations.Contains("|") Then
            asSourceLocations = Split(sSoureLocations, "|")
            For iCounter = 0 To UBound(asSourceLocations)
                If asSourceLocations(iCounter) <> vbNullString Then
                    ArchiveFlatExtractionRuby.WriteLine("	evidence_container.addFile(" & """" & asSourceLocations(iCounter).Replace("\", "\\") & """" & ")")
                End If
            Next
        Else
            ArchiveFlatExtractionRuby.WriteLine("	evidence_container.addFile(" & """" & sSoureLocations.Replace("\", "\\") & """" & ")")
        End If
        ArchiveFlatExtractionRuby.WriteLine("	evidence_container.setEncoding(" & """" & "utf-8" & """" & ")")
        ArchiveFlatExtractionRuby.WriteLine("	evidence_container.save")
        ArchiveFlatExtractionRuby.WriteLine("end")
        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("if archive_centera == " & """" & "yes" & """")
        ArchiveFlatExtractionRuby.WriteLine("	evidence_archive = archive_name")
        ArchiveFlatExtractionRuby.WriteLine("	evidence_container = processor.newEvidenceContainer(evidence_archive)")
        ArchiveFlatExtractionRuby.WriteLine("	centera_cluster_settings = {")
        ArchiveFlatExtractionRuby.WriteLine("		:ipsFile => " & """" & "#{archive_centera_ip}" & """" & ",")
        ArchiveFlatExtractionRuby.WriteLine("		:clipsFile => " & """" & "#{archive_centera_clip}" & """" & ", ")
        ArchiveFlatExtractionRuby.WriteLine("	}")
        ArchiveFlatExtractionRuby.WriteLine("	evidence_container.addCenteraCluster(centera_cluster_settings)	")
        ArchiveFlatExtractionRuby.WriteLine("	evidence_container.setEncoding(" & """" & "utf-8" & """" & ")")
        ArchiveFlatExtractionRuby.WriteLine("	evidence_container.save")
        ArchiveFlatExtractionRuby.WriteLine("end")
        ArchiveFlatExtractionRuby.WriteLine("   id_password = archive_password")
        ArchiveFlatExtractionRuby.WriteLine("	id_file = archive_keystore")
        If sArchiveName = "NSF" Then
            ArchiveFlatExtractionRuby.WriteLine("   processor.addKeyStore(id_file,{")
            ArchiveFlatExtractionRuby.WriteLine("        :filePassword => id_password,")
            ArchiveFlatExtractionRuby.WriteLine("        :target => " & """" & "" & """")
            ArchiveFlatExtractionRuby.WriteLine("        })")
        End If

        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("start_time = Time.now")
        ArchiveFlatExtractionRuby.WriteLine("last_progress = Time.now")
        ArchiveFlatExtractionRuby.WriteLine("semaphore = Mutex.new")
        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("puts")
        ArchiveFlatExtractionRuby.WriteLine("puts")
        ArchiveFlatExtractionRuby.WriteLine("puts " & """" & "Email Archive Extraction for #{archive_name} started at #{start_time}..." & """")
        ArchiveFlatExtractionRuby.WriteLine("puts")
        ArchiveFlatExtractionRuby.WriteLine("printf " & """" & "\r%-40s %-25s %-25s %-25s %-25s" & """" & ", " & """" & "Timestamp" & """" & ", " & """" & "Processed Items" & """" & ", " & """" & "Current Bytes" & """" & ", " & """" & "Total Bytes" & """" & ", " & """" & "Percent (%) Completed" & """")
        ArchiveFlatExtractionRuby.WriteLine("puts")
        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("if archive_source_size == " & """" & "0" & """")
        ArchiveFlatExtractionRuby.WriteLine("processor.when_progress_updated do |progress|")
        ArchiveFlatExtractionRuby.WriteLine("	semaphore.synchronize {")
        ArchiveFlatExtractionRuby.WriteLine("		class Numeric")
        ArchiveFlatExtractionRuby.WriteLine("    def percent_of(n)")
        ArchiveFlatExtractionRuby.WriteLine("        self.to_f / n.to_f * 100")
        ArchiveFlatExtractionRuby.WriteLine("      end")
        ArchiveFlatExtractionRuby.WriteLine("    end")
        ArchiveFlatExtractionRuby.WriteLine("    # Progress message every 15 seconds")
        ArchiveFlatExtractionRuby.WriteLine("current_size = progress.get_current_size")
        ArchiveFlatExtractionRuby.WriteLine("total_size = progress.get_total_size")
        ArchiveFlatExtractionRuby.WriteLine("percent_completed = current_size.percent_of(total_size).round(1)")
        ArchiveFlatExtractionRuby.WriteLine("  if callback_count % CALLBACK_FREQUENCY == 0")
        ArchiveFlatExtractionRuby.WriteLine("       last_progress = Time.now")
        ArchiveFlatExtractionRuby.WriteLine("  		printf " & """" & "\r%-40s %-25s %-25s %-25s %-25s" & """" & ", Time.now, callback_count, current_size, total_size, 0")
        ArchiveFlatExtractionRuby.WriteLine("       updated_callback = [callback_count, current_size, " & """" & "0" & """" & "]")
        ArchiveFlatExtractionRuby.WriteLine("       db.update(" & """" & "UPDATE archiveExtractionStats SET ItemsProcessed = ?, BytesProcessed = ?, PercentCompleted = ? WHERE BatchName = '#{archive_name}'" & """" & ",updated_callback)")
        ArchiveFlatExtractionRuby.WriteLine("		end")
        ArchiveFlatExtractionRuby.WriteLine("  }")
        ArchiveFlatExtractionRuby.WriteLine("  callback_count += 1")
        ArchiveFlatExtractionRuby.WriteLine("  end")
        ArchiveFlatExtractionRuby.WriteLine("end")
        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("if archive_source_size != " & """" & "0" & """")
        ArchiveFlatExtractionRuby.WriteLine("processor.when_progress_updated do |progress|")
        ArchiveFlatExtractionRuby.WriteLine("    semaphore.synchronize {")
        ArchiveFlatExtractionRuby.WriteLine("    class Numeric")
        ArchiveFlatExtractionRuby.WriteLine("    def percent_of(n)")
        ArchiveFlatExtractionRuby.WriteLine("        self.to_f / n.to_f * 100")
        ArchiveFlatExtractionRuby.WriteLine("      end")
        ArchiveFlatExtractionRuby.WriteLine("    end")
        ArchiveFlatExtractionRuby.WriteLine("    # Progress message every 15 seconds")
        ArchiveFlatExtractionRuby.WriteLine("    current_size = progress.get_current_size")
        ArchiveFlatExtractionRuby.WriteLine("    total_size = progress.get_total_size")
        ArchiveFlatExtractionRuby.WriteLine("    percent_completed = current_size.percent_of(total_size).round(1)")
        ArchiveFlatExtractionRuby.WriteLine("  if callback_count % CALLBACK_FREQUENCY == 0")
        ArchiveFlatExtractionRuby.WriteLine("       last_progress = Time.now")
        ArchiveFlatExtractionRuby.WriteLine("  		printf " & """" & "\r%-40s %-25s %-25s %-25s %-25s" & """" & ", Time.now, callback_count, current_size, total_size, percent_completed")
        ArchiveFlatExtractionRuby.WriteLine("		updated_callback = [callback_count, current_size, percent_completed]")
        ArchiveFlatExtractionRuby.WriteLine("       db.update(" & """" & "UPDATE archiveExtractionStats SET ItemsProcessed = ?, BytesProcessed = ?, PercentCompleted = ? WHERE BatchName = '#{archive_name}'" & """" & ",updated_callback)")
        ArchiveFlatExtractionRuby.WriteLine("	end")
        ArchiveFlatExtractionRuby.WriteLine("	}")
        ArchiveFlatExtractionRuby.WriteLine("   callback_count += 1")
        ArchiveFlatExtractionRuby.WriteLine("  end")
        ArchiveFlatExtractionRuby.WriteLine("end")
        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("processor.process")
        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("puts")
        ArchiveFlatExtractionRuby.WriteLine("end_time = Time.now")
        ArchiveFlatExtractionRuby.WriteLine("puts " & """" & "Email Archive Extraction for #{archive_name} finished at #{end_time}..." & """")
        ArchiveFlatExtractionRuby.WriteLine("puts")
        ArchiveFlatExtractionRuby.WriteLine("updated_callback = [" & """" & "Completed" & """" & ", callback_count, archive_source_size, 100, end_time]")
        ArchiveFlatExtractionRuby.WriteLine("db.update(" & """" & "UPDATE archiveExtractionStats SET ExtractionStatus = ?, ItemsProcessed = ?, BytesProcessed = ?, PercentCompleted = ?,ProcessEndTIme = ? WHERE BatchName = '#{archive_name}'" & """" & ",updated_callback)")

        ArchiveFlatExtractionRuby.WriteLine("")
        ArchiveFlatExtractionRuby.WriteLine("$current_case.close")
        ArchiveFlatExtractionRuby.WriteLine("return")
        ArchiveFlatExtractionRuby.Close()

        blnBuildArchiveFlatExtractionRubyScript = True

    End Function

    Public Function blnBuildNSFConversionRubyScript(ByVal sArchiveName As String, ByVal sProcessingFileDir As String, ByVal sCustodianEvidenceLocation As String, ByVal sEvidenceJSoN As String, ByVal sCaseDir As String, ByVal sRubyScriptFileName As String, ByVal sSQLiteDBLocation As String, ByVal sNumberOfWorkers As String, ByVal sMemoryPerWorker As String, ByVal sWorkerTempDir As String, ByVal bExtractFromSlackSpace As Boolean) As Boolean
        blnBuildNSFConversionRubyScript = False

        Dim NSFConversionRuby As StreamWriter

        NSFConversionRuby = New StreamWriter(sRubyScriptFileName)

        NSFConversionRuby.WriteLine("# Menu Title: NEAMM NSF Conversion")
        NSFConversionRuby.WriteLine("# Needs Selected Items: false")
        NSFConversionRuby.WriteLine("# ")
        NSFConversionRuby.WriteLine("# This script expects a JSON configured with NSF parameters completed in order to automatically convert data from NSF to EML or PST.   ")
        NSFConversionRuby.WriteLine("# ")
        NSFConversionRuby.WriteLine("# Version 2.0")
        NSFConversionRuby.WriteLine("# March 22 2017 - Alex Chatzistamatis, Nuix")
        NSFConversionRuby.WriteLine("# ")
        NSFConversionRuby.WriteLine("#######################################")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("require 'thread'")
        NSFConversionRuby.WriteLine("require 'json'")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("CALLBACK_FREQUENCY = 100")
        NSFConversionRuby.WriteLine("callback_count = 0")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("load """ & sSQLiteDBLocation.Replace("\", "\\") & "\\Database.rb_""")
        NSFConversionRuby.WriteLine("load """ & sSQLiteDBLocation.Replace("\", "\\") & "\\SQLite.rb_""")
        NSFConversionRuby.WriteLine("db = SQLite.new(""" & sSQLiteDBLocation.Replace("\", "\\") & "\\NuixEmailArchiveMigrationManager.db3" & """" & ")")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("#######################################")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("file = File.read('" & sEvidenceJSoN.Replace("\", "\\") & "')")
        NSFConversionRuby.WriteLine("parsed = JSON.parse(file)")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("    archive_file = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_file" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_keystore = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_keystore" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_password = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_password" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_name = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_name" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_centera = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_centera_ip = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera_ip" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_centera_clip = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera_clip" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_source_size = parsed[" & """" & "email_archive" & """" & "][" & """" & "source_size" & """" & "]")
        NSFConversionRuby.WriteLine("    archive_migration = parsed[" & """" & "email_archive" & """" & "][" & """" & "migration" & """" & "]")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("caseFactory = $utilities.getCaseFactory()")
        NSFConversionRuby.WriteLine("case_settings = {")
        NSFConversionRuby.WriteLine("    :compound => false,")
        NSFConversionRuby.WriteLine("    :name => " & """" & "#{archive_name}" & """" & ",")
        NSFConversionRuby.WriteLine("    :description => " & """" & "Created using Nuix Email Archive Migration Manager" & """" & ",")
        NSFConversionRuby.WriteLine("    :investigator => " & """" & "NEAMM NSF Conversion" & """")
        NSFConversionRuby.WriteLine("}")
        NSFConversionRuby.WriteLine("$current_case = caseFactory.create(" & """" & sCaseDir.Replace("\", "\\") & """" & ", case_settings)")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("processor = $current_case.createProcessor")
        NSFConversionRuby.WriteLine("if archive_migration == " & """" & "lightspeed" & """")
        NSFConversionRuby.WriteLine("	processing_settings = {")
        NSFConversionRuby.WriteLine("		:traversalScope => " & """" & "full_traversal" & """" & ",")
        NSFConversionRuby.WriteLine("		:analysisLanguage => " & """" & "en" & """" & ",")
        NSFConversionRuby.WriteLine("		:identifyPhysicalFiles => true,")
        NSFConversionRuby.WriteLine("		:reuseEvidenceStores => true,")

        NSFConversionRuby.WriteLine("		:reportProcessingStatus => " & """" & "physical_files" & """")
        NSFConversionRuby.WriteLine("	}")
        NSFConversionRuby.WriteLine("end")
        NSFConversionRuby.WriteLine("processor.setProcessingSettings(processing_settings)")
        NSFConversionRuby.WriteLine("parallel_processing_settings = {")
        NSFConversionRuby.WriteLine("	:workerCount => " & sNumberOfWorkers & ",")
        NSFConversionRuby.WriteLine("	:workerMemory => " & sMemoryPerWorker & ",")
        NSFConversionRuby.WriteLine("	:embedBroker => true,")
        NSFConversionRuby.WriteLine("	:brokerMemory => " & sMemoryPerWorker & ",")
        NSFConversionRuby.WriteLine("   :workerTemp => " & """" & sWorkerTempDir.Replace("\", "\\") & """")

        NSFConversionRuby.WriteLine("}")
        NSFConversionRuby.WriteLine("processor.setParallelProcessingSettings(parallel_processing_settings)")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("# MIME Type Fiter for row-based items")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/csv" & """" & ", { process_embedded: false })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/tab-separated-values" & """" & ", { process_embedded: false })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.sqlite-database" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-registry" & """" & ", { text_strip: true})")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/x-plist" & """" & ", { process_embedded: false })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-logfile" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-mft" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-usnjrnl" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/exe" & """" & ", { process_text: false })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.tcpdump.pcap" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/x-common-log" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-iis-log" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-windows-event-log-record" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-windows-event-logx" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/x-pcapng" & """" & ", { process_embedded: false, text_strip: true })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.symantec-vault-stream-data" & """" & ", { enabled: false })")
        NSFConversionRuby.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-cab-compressed" & """" & ", { process_embedded: true })")
        NSFConversionRuby.WriteLine(" ")
        NSFConversionRuby.WriteLine(" ")
        NSFConversionRuby.WriteLine("cust_name = archive_name")
        NSFConversionRuby.WriteLine("evidence_container = processor.newEvidenceContainer(cust_name)")
        NSFConversionRuby.WriteLine("evidence_container.addFile(" & """" & sCustodianEvidenceLocation & """" & ")")
        NSFConversionRuby.WriteLine("	evidence_container.setEncoding(" & """" & "utf-8" & """" & ")")
        NSFConversionRuby.WriteLine("	evidence_container.save")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("	id_file = archive_keystore")
        NSFConversionRuby.WriteLine("   id_password = archive_password")
        NSFConversionRuby.WriteLine("   processor.addKeyStore(id_file,{")
        NSFConversionRuby.WriteLine("        :filePassword => id_password,")
        NSFConversionRuby.WriteLine("        :target => " & """" & "" & """")
        NSFConversionRuby.WriteLine("        })")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("start_time = Time.now")
        NSFConversionRuby.WriteLine("last_progress = Time.now")
        NSFConversionRuby.WriteLine("semaphore = Mutex.new")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("puts")
        NSFConversionRuby.WriteLine("puts " & """" & "NSF Conversion for #{archive_name} has started..." & """")
        NSFConversionRuby.WriteLine("puts")
        NSFConversionRuby.WriteLine("printf " & """" & "%-40s %-25s %-25s %-25s" & """" & "," & """" & "Timestamp" & """" & ", " & """" & "Bytes Converted" & """" & ", " & """" & "Total Bytes" & """" & ", " & """" & "Percent (%) Completed" & """")
        NSFConversionRuby.WriteLine("puts")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("processor.when_item_processed do |event|")
        NSFConversionRuby.WriteLine("   semaphore.synchronize {")
        NSFConversionRuby.WriteLine("	class Numeric")
        NSFConversionRuby.WriteLine("	  def percent_of(n)")
        NSFConversionRuby.WriteLine("        self.to_f / n.to_f * 100")
        NSFConversionRuby.WriteLine("      end")
        NSFConversionRuby.WriteLine("    end")
        NSFConversionRuby.WriteLine("    # Progress message every 15 seconds")
        NSFConversionRuby.WriteLine("	current_size = progress.get_current_size")
        NSFConversionRuby.WriteLine("	total_size = progress.get_total_size")
        NSFConversionRuby.WriteLine("	percent = current_size.percent_of(total_size).round(1)")
        NSFConversionRuby.WriteLine("    if (Time.now - last_progress) > 15")
        NSFConversionRuby.WriteLine("         last_progress = Time.now")
        NSFConversionRuby.WriteLine("			printf " & """" & "\r%-40s %-25s %-25s %-25s" & """" & ", Time.now, current_size, total_size, percent")
        NSFConversionRuby.WriteLine("			updated_callback = [Time.now,current_size,percent]")
        NSFConversionRuby.WriteLine("			db.update(" & """" & "UPDATE ewsIngestionStats SET TimeStamp = ?, BytesUploaded = ?, PercentCompleted = ? WHERE CustodianName = '#{archive_name}'" & """" & ",updated_callback)")
        NSFConversionRuby.WriteLine("	end")
        NSFConversionRuby.WriteLine("    }")
        NSFConversionRuby.WriteLine("end")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("processor.process")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("puts")
        NSFConversionRuby.WriteLine("end_time = Time.now")
        NSFConversionRuby.WriteLine("total_time = '%.2f' % ((end_time-start_time)/60)")
        NSFConversionRuby.WriteLine("display_time =  '%.0f' % (total_time)")
        NSFConversionRuby.WriteLine("length = 60 / display_time.to_i")
        NSFConversionRuby.WriteLine("size = length.to_i * archive_source_size.to_i")
        NSFConversionRuby.WriteLine("speed = size.to_i / 1024 / 1024 / 1024.round(2)")
        NSFConversionRuby.WriteLine("final = speed.round(2)")
        NSFConversionRuby.WriteLine("puts")
        NSFConversionRuby.WriteLine("puts " & """" & "NSF Conversion for #{archive_name} has finished!" & """")
        NSFConversionRuby.WriteLine("puts")
        NSFConversionRuby.WriteLine("puts " & """" & "Completed in #{display_time} minutes and averaged #{final} GB/hr." & """")
        NSFConversionRuby.WriteLine("puts")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("updated_callback = [" & """" & "True" & """" & "]")
        NSFConversionRuby.WriteLine("db.update(" & """" & "UPDATE ewsIngestionStats SET = ? WHERE CustodianName = '#{archive_name}'" & """" & ",updated_callback)")
        NSFConversionRuby.WriteLine("")
        NSFConversionRuby.WriteLine("$current_case.close")
        NSFConversionRuby.WriteLine("return")
        NSFConversionRuby.Close()

        blnBuildNSFConversionRubyScript = True

    End Function

    Public Function blnBuildArchiveExtractionJSon(ByVal sJSonFileName As String, ByVal sSourceFolders As String, ByVal sEvidenceKeyStore As String, ByVal sCaseName As String, ByVal sMigrationType As String, ByVal sSourceSize As String, ByVal bCentera As Boolean) As Boolean
        Dim sMachineName As String

        Dim asCenteraInfo() As String
        Dim sClipFile As String
        Dim sIPFile As String
        Dim FlatExtractJSonFile As StreamWriter

        blnBuildArchiveExtractionJSon = False

        sMachineName = System.Net.Dns.GetHostName()
        FlatExtractJSonFile = New StreamWriter(sJSonFileName)

        If sSourceFolders <> vbNullString Then
            If sSourceFolders.Contains("Centera") Then
                asCenteraInfo = Split(sSourceFolders, "|")
                sClipFile = asCenteraInfo(0)
                sIPFile = asCenteraInfo(1)
                sClipFile = sClipFile.Replace("CenteraClip:", "")
                sIPFile = sIPFile.Replace("CenteraIP:", "")
            End If
        End If

        FlatExtractJSonFile.WriteLine("{")
        FlatExtractJSonFile.WriteLine("  " & """" & "email_archive" & """" & ": {")

        If sSourceFolders <> vbNullString Then
            FlatExtractJSonFile.WriteLine("     " & """" & "evidence_file" & """" & ": " & """" & sSourceFolders.Replace("\", "\\") & """" & ",")
        Else
            FlatExtractJSonFile.WriteLine("     " & """" & "evidence_file" & """" & ": " & """" & """" & ",")
        End If
        FlatExtractJSonFile.WriteLine("     " & """" & "evidence_keystore" & """" & ": " & """" & sEvidenceKeyStore & """" & ",")
        FlatExtractJSonFile.WriteLine("     " & """" & "evidence_name" & """" & ": " & """" & sCaseName & """" & ",")
        If bCentera = True Then
            FlatExtractJSonFile.WriteLine("     " & """" & "centera" & """" & ": " & """" & "yes" & """" & ",")
            FlatExtractJSonFile.WriteLine("     " & """" & "centera_ip" & """" & ": " & """" & sIPFile.Replace("\", "\\") & """" & ",")
            FlatExtractJSonFile.WriteLine("     " & """" & "centera_clip" & """" & ": " & """" & sClipFile.Replace("\", "\\") & """" & ",")
        Else
            FlatExtractJSonFile.WriteLine("     " & """" & "centera" & """" & ": " & """" & "no" & """" & ",")
            FlatExtractJSonFile.WriteLine("     " & """" & "centera_ip" & """" & ": " & """" & """" & ",")
            FlatExtractJSonFile.WriteLine("     " & """" & "centera_clip" & """" & ": " & """" & """" & ",")
        End If

        FlatExtractJSonFile.WriteLine("     " & """" & "source_size" & """" & ": " & """" & sSourceSize & """" & ",")
        FlatExtractJSonFile.WriteLine("     " & """" & "migration" & """" & ": " & """" & sMigrationType & """")
        FlatExtractJSonFile.WriteLine(" }")
        FlatExtractJSonFile.WriteLine("}")

        FlatExtractJSonFile.Close()

        blnBuildArchiveExtractionJSon = True
    End Function

    Public Function blnBuildArchiveExtractionBatchFile(ByVal sNuixAppLocation As String, ByVal sArchiveName As String, ByVal sCaseName As String, ByVal sBatchFileName As String, ByVal sNMSSourceType As String, ByVal sNMSAddress As String, ByVal sNMSPort As String, ByVal sNMSUserName As String, ByVal sNMSAdminInfo As String, ByVal sLicenseType As String, ByVal sNumberOfWorkers As String, ByVal sNuixAppMemory As String, ByVal sNuixLogDir As String, ByVal sWorkerTempDir As String, ByVal sRubyFileName As String, ByVal sJavaTempDir As String, ByVal sExportDir As String, ByVal sCNDExportDir As String, ByVal sSQLUserName As String, ByVal sSQLAdminInfo As String, ByVal sServerNameandPort As String, ByVal sCSVMappingFile As String, ByVal sManifestFile As String, ByVal sReportPath As String, ByVal sDBName As String, ByVal sFailOnPartionErrors As String, ByVal sEnableMetadataQuery As String, ByVal sArchiveType As String, ByVal sDocServerID As String, ByVal sSourceFolder As String, ByVal sSisFolder As String, ByVal sFilteredAddressTypes As String, ByVal sFromDate As String, ByVal sToDate As String, ByVal sOutputFormat As String, ByVal sSkipSISIDLookup As String, ByVal sJSonFileName As String, ByVal sExpandDLTo As String, ByVal sCustodianMappingFile As String, ByVal iWorkerTimeout As Integer, ByVal sEVUserList As String, ByVal sPSTExportSize As String, ByVal bPSTAddDistributionListMetadata As Boolean, ByVal bEMLAddDistributionListMetadata As Boolean, ByVal sFileStore As String, ByVal sSQLAuthentication As String, ByVal sSQLDomain As String) As Boolean
        blnBuildArchiveExtractionBatchFile = False

        Dim ExtractionBatchFile As StreamWriter

        Dim sLicenceSourceType As String

        If sNMSSourceType = "Desktop" Then
            sLicenceSourceType = " -licencesourcetype dongle"
        Else
            sLicenceSourceType = " -licencesourcetype server -licencesourcelocation " & sNMSAddress & ":" & sNMSPort & " -Dnuix.registry.servers=" & sNMSAddress & " -licencetype email-archive-examiner"

        End If
        ExtractionBatchFile = New StreamWriter(sBatchFileName)
        ExtractionBatchFile.WriteLine("::TITLE is the destination SMTP Address")
        ExtractionBatchFile.WriteLine("@SET NUIX_USERNAME=" & sNMSUserName)
        ExtractionBatchFile.WriteLine("@SET NUIX_PASSWORD=" & sNMSAdminInfo)

        ' generic nuix startup info
        ExtractionBatchFile.Write("""" & sNuixAppLocation & """")
        ExtractionBatchFile.Write(sLicenceSourceType)
        ExtractionBatchFile.Write(" -licenceworkers " & sNumberOfWorkers)
        ExtractionBatchFile.Write(" " & sNuixAppMemory)
        ExtractionBatchFile.Write(" -Dnuix.logdir=" & """" & sNuixLogDir & """")
        ExtractionBatchFile.Write(" -Dnuix.worker.tmpdir=" & """" & sWorkerTempDir & """")
        ExtractionBatchFile.Write(" -Djava.io.tmpdir=" & """" & sJavaTempDir & """")
        ExtractionBatchFile.Write(" -Dnuix.processing.worker.timeout=" & iWorkerTimeout)

        If sOutputFormat = "pst" Then
            ExtractionBatchFile.Write(" -Dnuix.export.mailbox.maximumFileSizePerMailbox=" & sPSTExportSize & "GB")
        End If

        If ((sOutputFormat = "pst" Or sOutputFormat = "msg") And sArchiveType.Contains("Journal")) Then
            ExtractionBatchFile.Write(" -Dnuix.data.mapi.addDistributionListMetadata=" & bPSTAddDistributionListMetadata.ToString.ToLower)
        End If

        If (sOutputFormat = "eml") And (sArchiveName <> "Daegis AXS-One") Then
            ExtractionBatchFile.Write(" -Dnuix.rfc822.addDeliveredToHeaders=" & bEMLAddDistributionListMetadata.ToString.ToLower)
        End If

        ' lightspeed specific startup switches
        ExtractionBatchFile.Write(" -Dnuix.crackAndDump.exportDir=" & """" & sExportDir & """")
        If (sFromDate <> "NA") Then
            ExtractionBatchFile.Write(" -Dnuix.crackAndDump.filterFromDate=" & sFromDate)
            ExtractionBatchFile.Write(" -Dnuix.crackAndDump.filterToDate=" & sToDate)

        End If

        ' archive specific startup switches
        Select Case sArchiveName
            Case "Daegis AXS-One"
                ExtractionBatchFile.Write(" -Dnuix.data.axsone.sisFolders=" & """" & sSisFolder.Replace(";", "|") & """")
                ExtractionBatchFile.Write(" -Dnuix.data.axsOneProcessing=true")
                ExtractionBatchFile.Write(" -Dnuix.data.axsone.allowNoSis=" & sSkipSISIDLookup)
                If sArchiveType.Contains("Flat") Then
                    ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & """")
                ElseIf sArchiveType.Contains("User") Then
                    If sFileStore = vbNullString Then
                        ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & ":configFile=>" & sCustodianMappingFile & """")
                    Else
                        ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & ":configFile=>" & sCustodianMappingFile & ";manifest=>" & sFileStore & """")
                    End If
                    ExtractionBatchFile.Write(" -Dnuix.wss.config=" & """" & sJSonFileName & """")
                End If

            Case "HP/Autonomy EAS"
                ExtractionBatchFile.Write(" -Dnuix.zantaz.serverAndPort=" & sServerNameandPort)
                ExtractionBatchFile.Write(" -Dnuix.zantaz.databaseName=" & sDBName)
                ExtractionBatchFile.Write(" -Dnuix.zantaz.userName=" & sSQLUserName)
                ExtractionBatchFile.Write(" -Dnuix.zantaz.password=" & sSQLAdminInfo)
                ExtractionBatchFile.Write(" -Dnuix.zantaz.docserverID=" & sDocServerID)
                If sArchiveType = "Journal:Flat" Then
                    ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & """")
                ElseIf sArchiveType = "Journal:User" Then
                    If sFileStore = vbNullString Then
                        ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & ":configFile=>" & sCustodianMappingFile & """")
                    Else
                        ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & ":configFile=>" & sCustodianMappingFile & ";manifest=>" & sFileStore & """")
                    End If
                    ExtractionBatchFile.Write(" -Dnuix.wss.config=" & """" & sJSonFileName & """")
                End If
            Case "EMC EmailXtender"
                ExtractionBatchFile.Write(" -Dnuix.xtender.serverAndPort=" & sServerNameandPort)
                ExtractionBatchFile.Write(" -Dnuix.xtender.databaseName=" & sDBName)
                ExtractionBatchFile.Write(" -Dnuix.xtender.userName=" & sSQLUserName)
                ExtractionBatchFile.Write(" -Dnuix.xtender.password=" & sSQLAdminInfo)
                ExtractionBatchFile.Write(" -Dnuix.data.xtender.filteredAddressTypes=" & sFilteredAddressTypes)
                If sExpandDLTo = "Bcc: + " & """" & "Expanded-DL" & """" Then
                    ExtractionBatchFile.Write(" -Dnuix.data.xtender.excludeRoutableAddressesForToField=false")
                    ExtractionBatchFile.Write(" -Dnuix.data.xtender.routableRecipientTypeHidden=true")
                ElseIf sExpandDLTo = """" & "Expanded-DL" & """" Then
                    ExtractionBatchFile.Write(" -Dnuix.data.xtender.excludeRoutableAddressesForToField=true")
                    ExtractionBatchFile.Write(" -Dnuix.data.xtender.routableRecipientTypeHidden=false")
                End If
                If sArchiveType = "Journal:Flat" Then
                    ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & """")
                ElseIf sArchiveType = "Journal:User" Then
                    If sFileStore = vbNullString Then
                        ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & ":configFile=>" & sCustodianMappingFile & """")
                    Else
                        ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & ":configFile=>" & sCustodianMappingFile & ";manifest=>" & sFileStore & """")
                    End If
                    ExtractionBatchFile.Write(" -Dnuix.wss.config=" & """" & sJSonFileName & """")
                End If
            Case "EMC SourceOne"
                ExtractionBatchFile.Write(" -Dnuix.xtender.serverAndPort=" & sServerNameandPort)
                ExtractionBatchFile.Write(" -Dnuix.xtender.databaseName=" & sDBName)
                ExtractionBatchFile.Write(" -Dnuix.xtender.userName=" & sSQLUserName)
                ExtractionBatchFile.Write(" -Dnuix.xtender.password=" & sSQLAdminInfo)
                ExtractionBatchFile.Write(" -Dnuix.data.xtender.addressDbSchema=sourceOne")
                ExtractionBatchFile.Write(" -Dnuix.data.xtender.filteredAddressTypes=" & sFilteredAddressTypes)
                If sExpandDLTo = "Bcc: + " & """" & "Expanded-DL" & """" Then
                    ExtractionBatchFile.Write(" -Dnuix.data.xtender.excludeRoutableAddressesForToField=false")
                    ExtractionBatchFile.Write(" -Dnuix.data.xtender.routableRecipientTypeHidden=true")
                ElseIf sExpandDLTo = """" & "Expanded-DL" & """" Then
                    ExtractionBatchFile.Write(" -Dnuix.data.xtender.excludeRoutableAddressesForToField=true")
                    ExtractionBatchFile.Write(" -Dnuix.data.xtender.routableRecipientTypeHidden=false")
                End If
                If sArchiveType = "Journal:Flat" Then
                    ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & """")
                ElseIf sArchiveType = "Journal:User" Then
                    If sFileStore = vbNullString Then
                        ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & ":configFile=>" & sCustodianMappingFile & """")
                    Else
                        ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & ":configFile=>" & sCustodianMappingFile & ";manifest=>" & sFileStore & """")
                    End If
                    ExtractionBatchFile.Write(" -Dnuix.wss.config=" & """" & sJSonFileName & """")
                End If
            Case "Veritas Enterprise Vault"
                If sSQLAuthentication = "Windows Authentication" Then
                    ExtractionBatchFile.Write(" -Dnuix.symantecVault.serverAndPort=" & sServerNameandPort)
                    ExtractionBatchFile.Write(" -Dnuix.symantecVault.databaseName=" & sDBName)
                    ExtractionBatchFile.Write(" -Dnuix.symantecVault.domain=" & sSQLDomain)
                    ExtractionBatchFile.Write(" -Dnuix.symantecVault.userName=" & sSQLUserName)
                    ExtractionBatchFile.Write(" -Dnuix.symantecVault.password=" & sSQLAdminInfo)
                ElseIf sSQLAuthentication = "SQLServer Authentication" Then
                    ExtractionBatchFile.Write(" -Dnuix.symantecVault.serverAndPort=" & sServerNameandPort)
                    ExtractionBatchFile.Write(" -Dnuix.symantecVault.databaseName=" & sDBName)
                    ExtractionBatchFile.Write(" -Dnuix.symantecVault.userName=" & sSQLUserName)
                    ExtractionBatchFile.Write(" -Dnuix.symantecVault.password=" & sSQLAdminInfo)
                Else
                    ExtractionBatchFile.Write(" -Dnuix.symantecVault.serverAndPort=" & sServerNameandPort)
                    ExtractionBatchFile.Write(" -Dnuix.symantecVault.databaseName=" & sDBName)
                    ExtractionBatchFile.Write(" -Dnuix.symantecVault.userName=" & sSQLUserName)
                    ExtractionBatchFile.Write(" -Dnuix.symantecVault.password=" & sSQLAdminInfo)
                End If
                ExtractionBatchFile.Write(" -Dnuix.symantecVault.useFileTransactionId=" & sFailOnPartionErrors)
                'ExtractionBatchFile.Write(" -Dnuix.symantecVault.failOnPartitionErrors=" & sFailOnPartionErrors)
                ExtractionBatchFile.Write(" -Dnuix.symantecVault.enableMetadataQueries=" & sEnableMetadataQuery)
                If sArchiveType = "Mailbox:User" Then
                    ExtractionBatchFile.Write(" -Dnuix.crackAndDump.sourceDataMapping=" & """" & sManifestFile & """")
                    ExtractionBatchFile.Write(" -Dnuix.processing.crackAndDump.useRelativePath=true")

                    ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & ":configFile=>" & sCSVMappingFile & """")

                    ExtractionBatchFile.Write(" -Dcsv_path=" & """" & sEVUserList & """")
                    ExtractionBatchFile.Write(" -Dcreate_config=" & sOutputFormat)
                    ExtractionBatchFile.Write(" -Dnuix.manifest.evidence_size=10000 ")
                    ExtractionBatchFile.Write(" -Dnuix.export.mail.mailbox.coerceLooseEmailToMailbox=true")
                    ExtractionBatchFile.Write(" -Dreport_path=" & """" & sReportPath & """")
                ElseIf sArchiveType = "Journal:Flat" Then
                    ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & """")
                ElseIf sArchiveType = "Mailbox:Flat" Then
                    ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & """")
                ElseIf sArchiveType = "Journal:User" Then
                    If sFileStore = vbNullString Then
                        ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & ":configFile=>" & sCustodianMappingFile & """")
                    Else
                        ExtractionBatchFile.Write(" -Dnuix.processing.enableCrackAndDump=" & """" & sOutputFormat & ":configFile=>" & sCustodianMappingFile & ";manifest=>" & sFileStore & """")
                    End If
                    ExtractionBatchFile.Write(" -Dnuix.wss.config=" & """" & sJSonFileName & """")
                End If
        End Select
        ExtractionBatchFile.WriteLine("-Dnuix.symantecVault.options=useNTLMv2=true -Dnuix.xtender.options=useNTLMv2=true -Dnuix.zantaz.options=useNTLMv2=true")

        ExtractionBatchFile.WriteLine(" " & """" & sRubyFileName & """")
        ExtractionBatchFile.WriteLine("@exit")

        ExtractionBatchFile.Close()

        blnBuildArchiveExtractionBatchFile = True
    End Function

    Public Function blnBuildArchiveExtractionWSSJSon(ByVal sWSSJsonFileName As String, ByVal lstWSSSettings As List(Of String), ByVal sProcessingType As String, ByVal sCaseDirectory As String) As Boolean
        Dim sMachineName As String
        Dim sFileFiltering As String
        Dim sFilterFileExt As String
        Dim sCommunicationCSV As String
        Dim sPropertiesCSV As String
        Dim sMappingCSV As String
        Dim sCaseSensitive As String
        Dim sVerbose As String
        Dim sUnresponsive As String
        Dim sTagItems As String
        Dim sTagUnique As String
        Dim sReportItems As String
        Dim sReportQueries As String
        Dim sExcludeItem As String
        Dim sFilterMimeTypes As String
        Dim sFilterKinds As String
        Dim sTextQueriesCSV As String
        Dim sUseRegEx As String
        Dim sReportDirectory As String
        Dim WSSJSonFile As StreamWriter

        blnBuildArchiveExtractionWSSJSon = False

        For Each setting In lstWSSSettings
            If setting.ToString.Contains("fileFiltering:") Then
                sFileFiltering = setting.ToString.Replace("fileFiltering:", "")
            ElseIf setting.ToString.Contains("filterMimeTypes:") Then
                sFilterMimeTypes = setting.ToString.Replace("filterMimeTypes:", "")
            ElseIf setting.ToString.Contains("filterKinds:") Then
                sFilterKinds = setting.ToString.Replace("filterKinds:", "")
            ElseIf setting.ToString.Contains("filterFileExt:") Then
                sFilterFileExt = setting.ToString.Replace("filterFileExt:", "")
            ElseIf setting.ToString.Contains("communicationsCSV:") Then
                sCommunicationCSV = setting.ToString.Replace("communicationsCSV:", "")
            ElseIf setting.ToString.Contains("textQueriesCSV:") Then
                sTextQueriesCSV = setting.ToString.Replace("textQueriesCSV:", "")
            ElseIf setting.ToString.Contains("mappingCSV:") Then
                sMappingCSV = setting.ToString.Replace("mappingCSV:", "")
            ElseIf setting.ToString.Contains("useRegEx:") Then
                sUseRegEx = setting.ToString.Replace("useRegEx:", "")
            ElseIf setting.ToString.Contains("propertiesCSV:") Then
                sPropertiesCSV = setting.ToString.Replace("propertiesCSV:", "")
            ElseIf setting.ToString.Contains("excludeItems:") Then
                sExcludeItem = setting.ToString.Replace("excludeItems:", "")
            ElseIf setting.ToString.Contains("caseSensitive:") Then
                sCaseSensitive = setting.ToString.Replace("caseSensitive:", "")
            ElseIf setting.ToString.Contains("verbose:") Then
                sVerbose = setting.ToString.Replace("verbose:", "")
            ElseIf setting.ToString.Contains("flagUnresponsive:") Then
                sUnresponsive = setting.ToString.Replace("flagUnresponsive:", "")
            ElseIf setting.ToString.Contains("tagUnique:") Then
                sTagUnique = setting.ToString.Replace("tagUnique:", "")
            ElseIf setting.ToString.Contains("tagItems:") Then
                sTagItems = setting.ToString.Replace("tagItems:", "")
            ElseIf setting.ToString.Contains("reportDirectory:") Then
                sReportDirectory = setting.ToString.Replace("reportDirectory:", "")
            ElseIf setting.ToString.Contains("reportItems:") Then
                sReportItems = setting.ToString.Replace("reportItems:", "")
            ElseIf setting.ToString.Contains("reportQueries:") Then
                sReportQueries = setting.ToString.Replace("reportQueries:", "")
            End If
        Next

        If sFileFiltering = vbNullString Then
            sFileFiltering = "false"
        End If

        If sTagItems = vbNullString Then
            sTagItems = "false"
        End If

        If sTagUnique = vbNullString Then
            sTagUnique = "false"
        End If

        sMachineName = System.Net.Dns.GetHostName()
        WSSJSonFile = New StreamWriter(sWSSJsonFileName)

        WSSJSonFile.WriteLine("{")
        WSSJSonFile.WriteLine("  " & """" & "fileFiltering" & """" & ":" & Trim(sFileFiltering) & ",")
        WSSJSonFile.WriteLine("  " & """" & "filterMimeTypes" & """" & ": [" & sFilterMimeTypes & "],")
        WSSJSonFile.WriteLine("  " & """" & "filterKinds" & """" & ": [" & sFilterKinds & "],")
        If sFilterFileExt = """" & """" Then
            WSSJSonFile.WriteLine("  " & """" & "filterFileExt" & """" & ": [],")
        Else
            WSSJSonFile.WriteLine("  " & """" & "filterFileExt" & """" & ": [" & """" & Trim(sFilterFileExt) & """" & "],")
        End If

        WSSJSonFile.WriteLine("  " & """" & "communicationCSV" & """" & ": " & """" & Trim(sCommunicationCSV.Replace("\", "\\")) & """" & ",")
        WSSJSonFile.WriteLine("  " & """" & "textQueriesCSV" & """" & ": " & """" & """" & ",")
        WSSJSonFile.WriteLine("  " & """" & "mappingCSV" & """" & ": " & """" & Trim(sMappingCSV.Replace("\", "\\")) & """" & ",")
        WSSJSonFile.WriteLine("  " & """" & "useRegEx" & """" & ": false,")
        WSSJSonFile.WriteLine("  " & """" & "propertiesCSV" & """" & ": " & """" & Trim(sPropertiesCSV.Replace("\", "\\")) & """" & ",")
        WSSJSonFile.WriteLine("  " & """" & "excludeItems" & """" & ": " & Trim(sExcludeItem & ","))
        WSSJSonFile.WriteLine("  " & """" & "caseSensitive" & """" & ": false,")
        WSSJSonFile.WriteLine("  " & """" & "verbose" & """" & ": " & Trim(sVerbose & ","))
        WSSJSonFile.WriteLine("  " & """" & "flagUnresponsive" & """" & ": " & Trim(sUnresponsive) & ",")
        WSSJSonFile.WriteLine("  " & """" & "tagUnique" & """" & ": " & Trim(sTagUnique) & ",")
        WSSJSonFile.WriteLine("  " & """" & "tagItems" & """" & ": " & Trim(sTagItems) & ",")
        If sProcessingType = "lightspeed" Then
            WSSJSonFile.WriteLine("  " & """" & "customMetadata" & """" & ": " & """" & "Mailboxes" & """" & ",")
        Else
            WSSJSonFile.WriteLine("  " & """" & "customMetadata" & """" & ": " & """" & "" & """" & ",")

        End If
        WSSJSonFile.WriteLine("  " & """" & "report_directory" & """" & ": " & """" & Trim(sCaseDirectory.Replace("\", "\\")) & "\\WSSReports" & """" & ",")
        WSSJSonFile.WriteLine("  " & """" & "report_items" & """" & ": " & Trim(sReportItems) & ",")
        WSSJSonFile.WriteLine("  " & """" & "report_queries" & """" & ": " & Trim(sReportQueries))
        WSSJSonFile.WriteLine("}")

        WSSJSonFile.Close()

        blnBuildArchiveExtractionWSSJSon = True
    End Function

    Public Function blnBuildUpdatedArchiveExtractionWSSJSon(ByVal sWSSJsonFileName As String, ByVal lstWSSSettings As List(Of String), ByVal sProcessingType As String, ByVal sCaseDirectory As String) As Boolean
        Dim sMachineName As String

        Dim sFileFiltering As String
        Dim sFilterFileExt As String
        Dim sCommunicationCSV As String
        Dim sPropertiesCSV As String
        Dim sMappingCSV As String
        Dim sCaseSensitive As String
        Dim sVerbose As String
        Dim sUnresponsive As String
        Dim sTagItems As String
        Dim sTagUnique As String
        Dim sReportItems As String
        Dim sReportQueries As String
        Dim sExcludeItem As String
        Dim sFilterMimeTypes As String
        Dim sFilterKinds As String
        Dim sTextQueriesCSV As String
        Dim sUseRegEx As String
        Dim sReportDirectory As String
        Dim bVerbose As Boolean
        Dim bExcludeItems As Boolean
        Dim bFlagUnresponsive As Boolean
        Dim bTagItems As Boolean
        Dim bTagUnique As Boolean
        Dim sCustomMetadata As String
        Dim sCategoryProperty As String
        Dim sMimeType As String
        Dim sExt As String
        Dim bCaseInsensitive As Boolean
        Dim bUseRegExp As Boolean
        Dim sSearch As String
        Dim sAddress As String
        Dim sAddressField As String
        Dim sExpanded As String
        Dim sText As String
        Dim sProperties As String
        Dim bReportItems As Boolean
        Dim bReportQueries As Boolean
        Dim WSSJSonFile As StreamWriter

        blnBuildUpdatedArchiveExtractionWSSJSon = False

        bVerbose = False
        bExcludeItems = False
        bFlagUnresponsive = True
        bTagItems = True
        bTagUnique = False
        sCustomMetadata = ""
        sCategoryProperty = ""
        sMimeType = ""
        sExt = ""
        bCaseInsensitive = True
        bUseRegExp = True
        sSearch = ""
        sAddress = ""
        sAddressField = ""
        sExpanded = ""
        sText = ""
        sProperties = ""
        bReportItems = False
        bReportQueries = False

        For Each setting In lstWSSSettings
            If setting.ToString.Contains("fileFiltering:") Then
                sFileFiltering = setting.ToString.Replace("fileFiltering:", "")
            ElseIf setting.ToString.Contains("filterMimeTypes:") Then
                sFilterMimeTypes = setting.ToString.Replace("filterMimeTypes:", "")
            ElseIf setting.ToString.Contains("filterKinds:") Then
                sFilterKinds = setting.ToString.Replace("filterKinds:", "")
            ElseIf setting.ToString.Contains("filterFileExt:") Then
                sFilterFileExt = setting.ToString.Replace("filterFileExt:", "")
            ElseIf setting.ToString.Contains("communicationsCSV:") Then
                sCommunicationCSV = setting.ToString.Replace("communicationsCSV:", "")
            ElseIf setting.ToString.Contains("textQueriesCSV:") Then
                sTextQueriesCSV = setting.ToString.Replace("textQueriesCSV:", "")
            ElseIf setting.ToString.Contains("mappingCSV:") Then
                sMappingCSV = setting.ToString.Replace("mappingCSV:", "")
            ElseIf setting.ToString.Contains("useRegEx:") Then
                sUseRegEx = setting.ToString.Replace("useRegEx:", "")
            ElseIf setting.ToString.Contains("propertiesCSV:") Then
                sPropertiesCSV = setting.ToString.Replace("propertiesCSV:", "")
            ElseIf setting.ToString.Contains("excludeItems:") Then
                sExcludeItem = setting.ToString.Replace("excludeItems:", "")
            ElseIf setting.ToString.Contains("caseSensitive:") Then
                sCaseSensitive = setting.ToString.Replace("caseSensitive:", "")
            ElseIf setting.ToString.Contains("verbose:") Then
                sVerbose = setting.ToString.Replace("verbose:", "")
            ElseIf setting.ToString.Contains("flagUnresponsive:") Then
                sUnresponsive = setting.ToString.Replace("flagUnresponsive:", "")
            ElseIf setting.ToString.Contains("tagUnique:") Then
                sTagUnique = setting.ToString.Replace("tagUnique:", "")
            ElseIf setting.ToString.Contains("tagItems:") Then
                sTagItems = setting.ToString.Replace("tagItems:", "")
            ElseIf setting.ToString.Contains("reportDirectory:") Then
                sReportDirectory = setting.ToString.Replace("reportDirectory:", "")
            ElseIf setting.ToString.Contains("reportItems:") Then
                sReportItems = setting.ToString.Replace("reportItems:", "")
            ElseIf setting.ToString.Contains("reportQueries:") Then
                sReportQueries = setting.ToString.Replace("reportQueries:", "")
            End If
        Next

        If sFileFiltering = vbNullString Then
            sFileFiltering = "false"
        End If

        If sTagItems = vbNullString Then
            sTagItems = "false"
        End If

        If sTagUnique = vbNullString Then
            sTagUnique = "false"
        End If

        sMachineName = System.Net.Dns.GetHostName()
        WSSJSonFile = New StreamWriter(sWSSJsonFileName)

        WSSJSonFile.WriteLine("{")
        WSSJSonFile.WriteLine("  " & """" & "verbose" & """" & ": " & bVerbose.toString.tolower & ",")
        WSSJSonFile.WriteLine("  " & """" & "exclude_items" & """" & ": " & bExcludeItems.toString.tolower & ",")
        WSSJSonFile.WriteLine("  " & """" & "flag_unresponsive" & """" & ": " & bFlagUnresponsive.toString.tolower & ",")
        WSSJSonFile.WriteLine("  " & """" & "tag_items" & """" & ": " & bTagItems.toString.tolower & ",")
        WSSJSonFile.WriteLine("  " & """" & "tag_unique" & """" & ": " & bTagUnique.toString.tolower & ",")
        WSSJSonFile.WriteLine("  " & """" & "custom_metadata" & """" & ": " & """" & sCustomMetadata & """" & ",")
        WSSJSonFile.WriteLine("  " & """" & "category_property" & """" & ": " & """" & sCategoryProperty & """" & ",")
        WSSJSonFile.WriteLine("  " & """" & "file_filter" & """" & ": {")
        WSSJSonFile.WriteLine("      " & """" & "mime" & """" & ": [" & sMimeType & "],")
        WSSJSonFile.WriteLine("      " & """" & "kind" & """" & ": [" & """" & "email" & """" & "," & """" & "calendar" & """" & "," & """" & "contact" & """" & "],")
        WSSJSonFile.WriteLine("      " & """" & "ext" & """" & ": " & sExt & "]")
        WSSJSonFile.WriteLine("  },")
        WSSJSonFile.WriteLine("  " & """" & "case_insensitive" & """" & ": " & bCaseInsensitive.tostring.tolower & ",")
        WSSJSonFile.WriteLine("  " & """" & "use_regexp" & """" & ": " & bUseRegExp.toString.ToLower & ",")
        WSSJSonFile.WriteLine("  " & """" & "csv_queries" & """" & ": {")
        WSSJSonFile.WriteLine("      " & """" & "search" & """" & ": " & """" & sSearch & """" & ",")
        WSSJSonFile.WriteLine("      " & """" & "address" & """" & ": " & """" & sAddress & """" & ",")
        WSSJSonFile.WriteLine("      " & """" & "address_field" & """" & ": " & sAddressField & """" & ",")
        WSSJSonFile.WriteLine("      " & """" & "expanded" & """" & ": " & """" & sExpanded & """" & ",")
        WSSJSonFile.WriteLine("      " & """" & "text" & """" & ": " & sText & """" & ",")
        WSSJSonFile.WriteLine("      " & """" & "properties" & """" & ": " & sProperties & """")
        WSSJSonFile.WriteLine("  },")
        WSSJSonFile.WriteLine("  " & """" & "report_directory" & """" & ": " & sReportDirectory & ",")
        WSSJSonFile.WriteLine("  " & """" & "report_items" & """" & ": " & bReportItems.toString.tolower & ",")
        WSSJSonFile.WriteLine("  " & """" & "report_queries" & """" & ": " & bReportQueries.toString.tolower & ",")
        WSSJSonFile.WriteLine("}")

        WSSJSonFile.Close()

        blnBuildUpdatedArchiveExtractionWSSJSon = True
    End Function

    Public Function blnCheckIfProcessIsRunning(ByVal sProcessID As String) As Boolean

        Dim NuixProcess As System.Diagnostics.Process

        blnCheckIfProcessIsRunning = False

        Try
            NuixProcess = Process.GetProcessById(CInt(sProcessID))
            blnCheckIfProcessIsRunning = True
        Catch ex As Exception
            blnCheckIfProcessIsRunning = False
        End Try
    End Function

    Public Function blnBuildArchiveCommunicationCSVFile() As Boolean
        blnBuildArchiveCommunicationCSVFile = False


        blnBuildArchiveCommunicationCSVFile = True
    End Function

    Public Function blnBuildArchiveTextCSVFile() As Boolean
        blnBuildArchiveTextCSVFile = False


        blnBuildArchiveTextCSVFile = True
    End Function

    Public Function blnBuildArchivePropertiesCSVFile() As Boolean
        blnBuildArchivePropertiesCSVFile = False


        blnBuildArchivePropertiesCSVFile = True
    End Function

    Public Function blnBuildArchiveEVUserExtractionRuby(ByVal sNuixCaseDir As String, ByVal sCaseName As String, ByVal sNuixFilesDirectory As String, ByVal sRubyFileName As String, ByVal sNumberOfWorkers As String, ByVal sMemoryPerWorker As String, ByVal sSQLiteDBLocation As String, ByVal sNuixConsolePath As String, ByVal sEvidenceJSoN As String, ByVal iEvidenceSize As Double, ByVal sWorkerTempDir As String, ByVal bExtractFromSlackSpace As Boolean, ByVal sSQLAuthentication As String) As Boolean
        Dim ArchiveEVUserExtraction As StreamWriter

        blnBuildArchiveEVUserExtractionRuby = False


        ArchiveEVUserExtraction = New StreamWriter(sRubyFileName)

        ArchiveEVUserExtraction.WriteLine("# encoding: UTF-8")
        ArchiveEVUserExtraction.WriteLine("#libraries")
        ArchiveEVUserExtraction.WriteLine("# ")
        ArchiveEVUserExtraction.WriteLine("# Menu Title: Archive Processor")
        ArchiveEVUserExtraction.WriteLine("# Needs Selected Items: false")
        ArchiveEVUserExtraction.WriteLine("# ")
        ArchiveEVUserExtraction.WriteLine("# This script expects a JSON configured with archive parameters completed in order to automatically process data inside from a legacy email archive. ")
        ArchiveEVUserExtraction.WriteLine("# ")
        ArchiveEVUserExtraction.WriteLine("# Feb 23 2017 - Alex Chatzistamatis, Nuix")
        ArchiveEVUserExtraction.WriteLine("# ")
        ArchiveEVUserExtraction.WriteLine("#######################################")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("require 'fileutils'")
        ArchiveEVUserExtraction.WriteLine("require 'thread'")
        ArchiveEVUserExtraction.WriteLine("require 'set'")
        ArchiveEVUserExtraction.WriteLine("#require " & """" & sNuixConsolePath.Replace("\", "\\") & "OptionSelectorDialog.jar" & """")
        ArchiveEVUserExtraction.WriteLine("require " & """" & "rubygems" & """")
        ArchiveEVUserExtraction.WriteLine("require 'java'")

        If (sNuixConsolePath.Contains("Nuix 7")) Then
            '            ArchiveEVUserExtraction.WriteLine("require " & """" & sNuixConsolePath.Replace("\", "\\") & "lib\\sqljdbc4-N1.0.jar" & """")
            ArchiveEVUserExtraction.WriteLine("require " & """" & sNuixConsolePath.Replace("\", "\\") & "lib\\sqljdbc41.jar" & """")
        Else
            ArchiveEVUserExtraction.WriteLine("require " & """" & sNuixConsolePath.Replace("\", "\\") & "lib\\sqljdbc4.jar" & """")
        End If

        ArchiveEVUserExtraction.WriteLine("# encoding: UTF-8")
        ArchiveEVUserExtraction.WriteLine("#libraries")
        ArchiveEVUserExtraction.WriteLine("require 'json'")
        ArchiveEVUserExtraction.WriteLine("require 'csv'")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("load """ & sSQLiteDBLocation.Replace("\", "\\") & "\\Database.rb_""")
        ArchiveEVUserExtraction.WriteLine("load """ & sSQLiteDBLocation.Replace("\", "\\") & "\\SQLite.rb_""")
        ArchiveEVUserExtraction.WriteLine("db = SQLite.new(""" & sSQLiteDBLocation.Replace("\", "\\") & "\\NuixEmailArchiveMigrationManager.db3" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("Java::com.microsoft.sqlserver.jdbc.SQLServerDriver")
        ArchiveEVUserExtraction.WriteLine("java_import javax.swing.SwingUtilities")
        ArchiveEVUserExtraction.WriteLine("java_import javax.swing.JFileChooser")
        ArchiveEVUserExtraction.WriteLine("java_import javax.swing.filechooser.FileNameExtensionFilter")
        ArchiveEVUserExtraction.WriteLine("#java_import 'nuix.plugins.OptionSelectorDialog'")
        ArchiveEVUserExtraction.WriteLine("java_import java.util.ArrayList")
        ArchiveEVUserExtraction.WriteLine("java_import javax.swing.UIManager")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("CALLBACK_FREQUENCY = 100")
        ArchiveEVUserExtraction.WriteLine("callback_count = 0")
        ArchiveEVUserExtraction.WriteLine("evidence_size = 1000")
        ArchiveEVUserExtraction.WriteLine("evidence_size = ENV_JAVA['nuix.manifest.evidence_size'] if !(ENV_JAVA['nuix.manifest.evidence_size'].nil?)")
        ArchiveEVUserExtraction.WriteLine("file = File.read('" & sEvidenceJSoN.Replace("\", "\\") & "')")
        ArchiveEVUserExtraction.WriteLine("parsed = JSON.parse(file)")

        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("#Environmental Variables from batch file")
        ArchiveEVUserExtraction.WriteLine("manifest_dir = ENV_JAVA['nuix.crackAndDump.sourceDataMapping']")
        ArchiveEVUserExtraction.WriteLine("CND_config = ENV_JAVA['nuix.processing.enableCrackAndDump']")
        ArchiveEVUserExtraction.WriteLine("export_path = ENV_JAVA['nuix.crackAndDump.exportDir']")
        ArchiveEVUserExtraction.WriteLine("report_path = ENV_JAVA['report_path']")
        ArchiveEVUserExtraction.WriteLine("working_path = ENV_JAVA['working_path']")
        ArchiveEVUserExtraction.WriteLine("csv_path = ENV_JAVA['csv_path']")
        ArchiveEVUserExtraction.WriteLine("create_config = ENV_JAVA['create_config']")
        ArchiveEVUserExtraction.WriteLine("noCnD = ENV_JAVA['nuix.manifest.noCnD']")
        ArchiveEVUserExtraction.WriteLine("$Centera = false")
        ArchiveEVUserExtraction.WriteLine("$Centera = ENV_JAVA['nuix.manifest.Centera'] if !(ENV_JAVA['nuix.manifest.Centera'].nil?)")
        ArchiveEVUserExtraction.WriteLine("$Centera_ipFile = ENV_JAVA['nuix.manifest.Centera_IPFile']")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("#SQL variables")
        ArchiveEVUserExtraction.WriteLine("vaulthost = ENV_JAVA['nuix.symantecVault.serverAndPort']")
        ArchiveEVUserExtraction.WriteLine("#If SQL db not specified, default to EnterpriseVaultDirectory as consistent with Nuix App")
        ArchiveEVUserExtraction.WriteLine("if (ENV_JAVA['nuix.symantecVault.databaseName'].nil?)")
        ArchiveEVUserExtraction.WriteLine("	$vaultdb1 = " & """" & "EnterpriseVaultDirectory" & """")
        ArchiveEVUserExtraction.WriteLine("	else")
        ArchiveEVUserExtraction.WriteLine("	$vaultdb1 = ENV_JAVA['nuix.symantecVault.databaseName']")
        ArchiveEVUserExtraction.WriteLine("	end")
        ArchiveEVUserExtraction.WriteLine("sa = ENV_JAVA['nuix.symantecVault.userName']")
        ArchiveEVUserExtraction.WriteLine("sa_password = ENV_JAVA['nuix.symantecVault.password']")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("#connection")
        If sSQLAuthentication = "Windows Authentication" Then
            ArchiveEVUserExtraction.WriteLine("url = 'jdbc:sqlserver://'+" & """" & "#{vaulthost}" & """" & "+';databaseName='+" & """" & "#{$vaultdb1}" & """" & "+';integratedSecurity=true'")
        ElseIf sSQLAuthentication = "SQLServer Authentication" Then
            ArchiveEVUserExtraction.WriteLine("url = 'jdbc:sqlserver://'+" & """" & "#{vaulthost}" & """" & "+';databaseName='+" & """" & "#{$vaultdb1}" & """" & "+';integratedSecurity=false'")
        End If
        ArchiveEVUserExtraction.WriteLine("conn = java.sql.DriverManager.get_connection(url, " & """" & "#{sa}" & """" & ", " & """" & "#{sa_password}" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("$statement1 = conn.create_statement")
        ArchiveEVUserExtraction.WriteLine("puts " & """" & "Connected to main EV DB..." & """")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("#######################################")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def getFromCSV(csv_path)")
        ArchiveEVUserExtraction.WriteLine("	archives = Array.new")
        ArchiveEVUserExtraction.WriteLine("	if csv_path.downcase == " & """" & "user" & """")
        ArchiveEVUserExtraction.WriteLine("	chooser = javax.swing.JFileChooser.new")
        ArchiveEVUserExtraction.WriteLine("	chooser.dialog_title = 'Choose a CSV of target archive names'")
        ArchiveEVUserExtraction.WriteLine("		if chooser.show_open_dialog(nil) == javax.swing.JFileChooser::APPROVE_OPTION")
        ArchiveEVUserExtraction.WriteLine("		  csv_file = chooser.selected_file.path")
        ArchiveEVUserExtraction.WriteLine("		  puts " & """" & "Using archive file: #{csv_file}" & """")
        ArchiveEVUserExtraction.WriteLine("		else")
        ArchiveEVUserExtraction.WriteLine("		  return")
        ArchiveEVUserExtraction.WriteLine("		end")
        ArchiveEVUserExtraction.WriteLine("	else")
        ArchiveEVUserExtraction.WriteLine("		csv_file = csv_path")
        ArchiveEVUserExtraction.WriteLine("	end")
        ArchiveEVUserExtraction.WriteLine("	CSV.foreach(" & """" & "#{csv_file}" & """" & ") do |line|")
        ArchiveEVUserExtraction.WriteLine("		archives << line[0]")
        ArchiveEVUserExtraction.WriteLine("	end")
        ArchiveEVUserExtraction.WriteLine("	return archives")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def getArchiveVaultIDs(archives)")
        ArchiveEVUserExtraction.WriteLine("#Lookup archive vault IDs")
        ArchiveEVUserExtraction.WriteLine("	archive_list = Array.new")
        ArchiveEVUserExtraction.WriteLine("	archives.each do |query|")
        ArchiveEVUserExtraction.WriteLine("		rs = $statement1.execute_query(" & """" & "SELECT ArchiveName, VaultStoreEntryID FROM dbo.Archive where ArchiveName = '#{query}'" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("		while (rs.next) do")
        ArchiveEVUserExtraction.WriteLine("			archive = Array.new")
        ArchiveEVUserExtraction.WriteLine("			archive << " & """" & "#{rs.getString('ArchiveName')}" & """")
        ArchiveEVUserExtraction.WriteLine("			archive << " & """" & "#{rs.getString('VaultStoreEntryID')}" & """")
        ArchiveEVUserExtraction.WriteLine("			archive_list << archive")
        ArchiveEVUserExtraction.WriteLine("		end")
        ArchiveEVUserExtraction.WriteLine("	end")
        ArchiveEVUserExtraction.WriteLine("	return archive_list, archives")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def getStores()")
        ArchiveEVUserExtraction.WriteLine("#This looks up the available Vault Stores from EV to be used for user selection")
        ArchiveEVUserExtraction.WriteLine("stores_list = Array.new")
        ArchiveEVUserExtraction.WriteLine("  rs = $statement1.execute_query(" & """" & "SELECT VaultStoreDescription, VaultStoreEntryId FROM dbo.VaultStoreEntry" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("  while (rs.next) do")
        ArchiveEVUserExtraction.WriteLine("    store = Array.new")
        ArchiveEVUserExtraction.WriteLine("    store << " & """" & "#{rs.getString('VaultStoreDescription')}" & """")
        ArchiveEVUserExtraction.WriteLine("	store << " & """" & "#{rs.getString('VaultStoreEntryId')}" & """")
        ArchiveEVUserExtraction.WriteLine("	stores_list << store")
        ArchiveEVUserExtraction.WriteLine("  end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("  return stores_list")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def getArchives()")
        ArchiveEVUserExtraction.WriteLine("#Depreciated, this was for when the user did not specify the Vault Store and just went from a dump of archive names")
        ArchiveEVUserExtraction.WriteLine("archive_list = Array.new")
        ArchiveEVUserExtraction.WriteLine("  rs = $statement1.execute_query(" & """" & "SELECT ArchiveName, RootIdentity FROM dbo.Archive" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("  while (rs.next) do")
        ArchiveEVUserExtraction.WriteLine("    archive = " & """" & "#{rs.getString('ArchiveName')}" & """")
        ArchiveEVUserExtraction.WriteLine("	archive_list << archive")
        ArchiveEVUserExtraction.WriteLine("  end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("  return archive_list")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def getStoreArchives(stores)")
        ArchiveEVUserExtraction.WriteLine("#Once we have selected the vault stores, we look up the archives and pass back the names and vault store IDs so we can later look them up in their specific dbs")
        ArchiveEVUserExtraction.WriteLine("archive_list = Array.new")
        ArchiveEVUserExtraction.WriteLine("query = Array.new")
        ArchiveEVUserExtraction.WriteLine("stores.each do |store|")
        ArchiveEVUserExtraction.WriteLine("	query << store[1]")
        ArchiveEVUserExtraction.WriteLine("	end")
        ArchiveEVUserExtraction.WriteLine("query = " & """" & "'#{query.join(" & """" & "','" & """" & ")}'" & """")
        ArchiveEVUserExtraction.WriteLine("  rs = $statement1.execute_query(" & """" & "SELECT ArchiveName, VaultStoreEntryID FROM dbo.Archive where VaultStoreEntryID in (#{query}) order by ArchiveName" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("  while (rs.next) do")
        ArchiveEVUserExtraction.WriteLine("    archive = Array.new")
        ArchiveEVUserExtraction.WriteLine("    archive << " & """" & "#{rs.getString('ArchiveName')}" & """")
        ArchiveEVUserExtraction.WriteLine("	archive << " & """" & "#{rs.getString('VaultStoreEntryID')}" & """")
        ArchiveEVUserExtraction.WriteLine("	archive_list << archive")
        ArchiveEVUserExtraction.WriteLine("  end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("  return archive_list")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def getUserSelectedArchives()")
        ArchiveEVUserExtraction.WriteLine("#Simple user select from a list of archives")
        ArchiveEVUserExtraction.WriteLine("	archives = getArchives()")
        ArchiveEVUserExtraction.WriteLine("	dialog_input = ArrayList.new")
        ArchiveEVUserExtraction.WriteLine("	dialog_input.addAll(archives)")
        ArchiveEVUserExtraction.WriteLine("#	dialog = OptionSelectorDialog.new(" & """" & "select archives to export" & """" & ", dialog_input, false)")
        ArchiveEVUserExtraction.WriteLine("	return dialog.getUserSelectedItems()")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def getUserSelectedStores()")
        ArchiveEVUserExtraction.WriteLine("#Select from a list of Vault stores")
        ArchiveEVUserExtraction.WriteLine("	stores = getStores()")
        ArchiveEVUserExtraction.WriteLine("	dialog_input = ArrayList.new")
        ArchiveEVUserExtraction.WriteLine("	stores.each do |item|")
        ArchiveEVUserExtraction.WriteLine("	dialog_input << item[0]")
        ArchiveEVUserExtraction.WriteLine("	end")
        ArchiveEVUserExtraction.WriteLine("#	dialog = OptionSelectorDialog.new(" & """" & "select source vault stores:" & """" & ", dialog_input, false)")
        ArchiveEVUserExtraction.WriteLine("	dialog = dialog.getUserSelectedItems()")
        ArchiveEVUserExtraction.WriteLine("	stores.delete_if do |item|")
        ArchiveEVUserExtraction.WriteLine("		 if !dialog.include?(item[0])")
        ArchiveEVUserExtraction.WriteLine("			true ")
        ArchiveEVUserExtraction.WriteLine("		  end")
        ArchiveEVUserExtraction.WriteLine("		end")
        ArchiveEVUserExtraction.WriteLine("	return stores")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def getStoreSelectedArchives(stores)")
        ArchiveEVUserExtraction.WriteLine("#Select archives filtered by Vault Store")
        ArchiveEVUserExtraction.WriteLine("	archives = getStoreArchives(stores)")
        ArchiveEVUserExtraction.WriteLine("	dialog_input = ArrayList.new")
        ArchiveEVUserExtraction.WriteLine("	archives.each do |item|")
        ArchiveEVUserExtraction.WriteLine("	dialog_input << item[0]")
        ArchiveEVUserExtraction.WriteLine("	end")
        ArchiveEVUserExtraction.WriteLine("# dialog = OptionSelectorDialog.new(" & """" & "select archives to export" & """" & ", dialog_input, false)")
        ArchiveEVUserExtraction.WriteLine("	dialog = dialog.getUserSelectedItems()")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("		archives.delete_if do |item|")
        ArchiveEVUserExtraction.WriteLine("		 if !dialog.include?(item[0])")
        ArchiveEVUserExtraction.WriteLine("			true ")
        ArchiveEVUserExtraction.WriteLine("		  end")
        ArchiveEVUserExtraction.WriteLine("		end")
        ArchiveEVUserExtraction.WriteLine("	return archives, dialog")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def getVaults(archive_list)")
        ArchiveEVUserExtraction.WriteLine("#This maps the Archives to each Vault Store DB as needed so the script can query the appropriate SQL dbs for target files and manifest info")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("		archive_list.sort_by! {|a| a[1]}")
        ArchiveEVUserExtraction.WriteLine("		vaultmap = Hash.new {|h,k| h[k] = []}")
        ArchiveEVUserExtraction.WriteLine("		dbmap = Hash.new {|h,k| h[k] = " & """" & """" & "}")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("		archive_list.each do |item|")
        ArchiveEVUserExtraction.WriteLine("			vaultmap[" & """" & "#{item[1]}" & """" & "] << item[0]")
        ArchiveEVUserExtraction.WriteLine("		end")
        ArchiveEVUserExtraction.WriteLine("		vaultmap.each {|vault, names|")
        ArchiveEVUserExtraction.WriteLine("			rs = $statement1.execute_query(" & """" & "SELECT DatabaseName FROM dbo.VaultStoreEntry where VaultStoreEntryID = '#{vault}'" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("			while (rs.next) do")
        ArchiveEVUserExtraction.WriteLine("				db = " & """" & "#{rs.getString('DatabaseName')}" & """")
        ArchiveEVUserExtraction.WriteLine("			end")
        ArchiveEVUserExtraction.WriteLine("			dbmap[" & """" & "#{db}" & """" & "] << " & """" & "'#{names.join(" & """" & "','" & """" & ")}'" & """")
        ArchiveEVUserExtraction.WriteLine("		}")
        ArchiveEVUserExtraction.WriteLine("	return dbmap")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def getTargetFiles(dbmap)")
        ArchiveEVUserExtraction.WriteLine("	targets = Array.new")
        ArchiveEVUserExtraction.WriteLine("	#We need to run the query once for each vault store db")
        ArchiveEVUserExtraction.WriteLine("	dbmap.each {|db, names|")
        ArchiveEVUserExtraction.WriteLine("	if $Centera == false")
        ArchiveEVUserExtraction.WriteLine("	#Centera uses a different query to pull target clips")
        ArchiveEVUserExtraction.WriteLine("	#This heinous query essentially transforms the various pieces into EV7 or EV8 target file names for CAB files when they are in a collection or loose DVS files when not")
        ArchiveEVUserExtraction.WriteLine("		rs = $statement1.execute_query(" & """" & "SELECT  g.PartitionRootPath + '\\,'+isnull(f.RelativeFileName, ")
        ArchiveEVUserExtraction.WriteLine("		(CASE ")
        ArchiveEVUserExtraction.WriteLine("			WHEN d.IdUniqueNo = -1 THEN ")
        ArchiveEVUserExtraction.WriteLine("				cast(year(d.archiveddate) as nvarchar(4))+'\\'+ ")
        ArchiveEVUserExtraction.WriteLine("				right('0'+cast(month(d.archiveddate) as nvarchar(2)),2)+'-'+ ")
        ArchiveEVUserExtraction.WriteLine("				right('0'+cast(day(d.archiveddate) as nvarchar(2)),2)+'\\'+ ")
        ArchiveEVUserExtraction.WriteLine("				substring(cast(d.IdTransaction as nvarchar(99)),1,1)+'\\'+ ")
        ArchiveEVUserExtraction.WriteLine("				substring(cast(d.IdTransaction as nvarchar(99)),2,3)+'\\'+ ")
        ArchiveEVUserExtraction.WriteLine("				replace(cast(d.IdTransaction as nvarchar(99))+'.DVS','-','')")
        ArchiveEVUserExtraction.WriteLine("			ELSE ")
        ArchiveEVUserExtraction.WriteLine("				cast(year(d.IdDateTime) as nvarchar(4))+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				right('0'+cast(month(d.IdDateTime) as nvarchar(2)),2)+'-'+")
        ArchiveEVUserExtraction.WriteLine("				right('0'+cast(day(d.IdDateTime) as nvarchar(2)),2)+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				cast(datepart(hour,d.IdDateTime) as nvarchar(2))+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				(CAST (d.IdChecksumHigh AS VARCHAR(7)) + ")
        ArchiveEVUserExtraction.WriteLine("				(CASE WHEN Len(d.IdChecksumHigh) = 1 THEN '000000' ")
        ArchiveEVUserExtraction.WriteLine("					  ELSE '' END)) + ")
        ArchiveEVUserExtraction.WriteLine("				(CASE WHEN Len(d.IdChecksumLow) = 1 THEN CAST(d.IdChecksumLow as nvarchar(1))+'0000000'")
        ArchiveEVUserExtraction.WriteLine("					  WHEN Len(d.IdChecksumLow) = 6 THEN '00'+CAST(d.IdChecksumLow as nvarchar(6))")
        ArchiveEVUserExtraction.WriteLine("					  WHEN Len(d.IdChecksumLow) = 7 THEN '0'+CAST(d.IdChecksumLow as nvarchar(7))")
        ArchiveEVUserExtraction.WriteLine("					  ELSE CAST(IdChecksumLow as nvarchar(8)) END)+'~'+")
        ArchiveEVUserExtraction.WriteLine("				REPLACE(REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(23),d.IdDateTime, 121),'-',''),' ',''),'.',''),':','')+'~'+")
        ArchiveEVUserExtraction.WriteLine("				(CAST (d.IdUniqueNo AS VARCHAR(36)) )+'.DVS'")
        ArchiveEVUserExtraction.WriteLine("			END)) as target")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("		  FROM #{$vaultdb1}.[dbo].[Archive] a")
        ArchiveEVUserExtraction.WriteLine("		  left join #{$vaultdb1}.dbo.Root b")
        ArchiveEVUserExtraction.WriteLine("		  on a.RootIdentity = ISNULL(b.containerrootidentity,b.rootidentity)")
        ArchiveEVUserExtraction.WriteLine("		  inner join #{db}.dbo.Vault c")
        ArchiveEVUserExtraction.WriteLine("		  on b.VaultEntryId = c.VaultId")
        ArchiveEVUserExtraction.WriteLine("		  inner join #{db}.dbo.Saveset d")
        ArchiveEVUserExtraction.WriteLine("		  on c.VaultIdentity = d.VaultIdentity")
        ArchiveEVUserExtraction.WriteLine("		  left join #{$vaultdb1}.dbo.ArchiveFolder e")
        ArchiveEVUserExtraction.WriteLine("		  on b.RootIdentity = e.RootIdentity")
        ArchiveEVUserExtraction.WriteLine("		  left join #{db}.dbo.Collection f")
        ArchiveEVUserExtraction.WriteLine("		  on d.CollectionIdentity = f.CollectionIdentity")
        ArchiveEVUserExtraction.WriteLine("		  left join #{$vaultdb1}.dbo.PartitionEntry g")
        ArchiveEVUserExtraction.WriteLine("		  on d.IdPartition = g.IdPartition and a.VaultStoreEntryId = g.VaultStoreEntryId")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("		 where a.ArchiveName in (#{names})")
        ArchiveEVUserExtraction.WriteLine("		 group by g.PartitionRootPath + '\\,'+isnull(f.RelativeFileName,")
        ArchiveEVUserExtraction.WriteLine("		(CASE ")
        ArchiveEVUserExtraction.WriteLine("			WHEN d.IdUniqueNo = -1 THEN")
        ArchiveEVUserExtraction.WriteLine("				cast(year(d.archiveddate) as nvarchar(4))+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				right('0'+cast(month(d.archiveddate) as nvarchar(2)),2)+'-'+")
        ArchiveEVUserExtraction.WriteLine("				right('0'+cast(day(d.archiveddate) as nvarchar(2)),2)+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				substring(cast(d.IdTransaction as nvarchar(99)),1,1)+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				substring(cast(d.IdTransaction as nvarchar(99)),2,3)+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				replace(cast(d.IdTransaction as nvarchar(99))+'.DVS','-','')")
        ArchiveEVUserExtraction.WriteLine("			ELSE ")
        ArchiveEVUserExtraction.WriteLine("				cast(year(d.IdDateTime) as nvarchar(4))+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				right('0'+cast(month(d.IdDateTime) as nvarchar(2)),2)+'-'+")
        ArchiveEVUserExtraction.WriteLine("				right('0'+cast(day(d.IdDateTime) as nvarchar(2)),2)+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				cast(datepart(hour,d.IdDateTime) as nvarchar(2))+'\\'+")
        ArchiveEVUserExtraction.WriteLine("             (CAST (d.IdChecksumHigh AS VARCHAR(7)) + ")
        ArchiveEVUserExtraction.WriteLine("				(CASE WHEN Len(d.IdChecksumHigh) = 1 THEN '000000' ")
        ArchiveEVUserExtraction.WriteLine("					  ELSE '' END)) +")
        ArchiveEVUserExtraction.WriteLine("				(CASE WHEN Len(d.IdChecksumLow) = 1 THEN CAST(d.IdChecksumLow as nvarchar(1))+'0000000'")
        ArchiveEVUserExtraction.WriteLine("					  WHEN Len(d.IdChecksumLow) = 6 THEN '00'+CAST(d.IdChecksumLow as nvarchar(6))")
        ArchiveEVUserExtraction.WriteLine("					  WHEN Len(d.IdChecksumLow) = 7 THEN '0'+CAST(d.IdChecksumLow as nvarchar(7))")
        ArchiveEVUserExtraction.WriteLine("					  ELSE CAST(IdChecksumLow as nvarchar(8)) END)+'~'+")
        ArchiveEVUserExtraction.WriteLine("				REPLACE(REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(23),d.IdDateTime, 121),'-',''),' ',''),'.',''),':','')+'~'+")
        ArchiveEVUserExtraction.WriteLine("				(CAST (d.IdUniqueNo AS VARCHAR(36)) )+'.DVS'")
        ArchiveEVUserExtraction.WriteLine("			END))")
        ArchiveEVUserExtraction.WriteLine("		 order by g.PartitionRootPath + '\\,'+isnull(f.RelativeFileName,")
        ArchiveEVUserExtraction.WriteLine("		(CASE ")
        ArchiveEVUserExtraction.WriteLine("			WHEN d.IdUniqueNo = -1 THEN")
        ArchiveEVUserExtraction.WriteLine("				cast(year(d.archiveddate) as nvarchar(4))+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				right('0'+cast(month(d.archiveddate) as nvarchar(2)),2)+'-'+")
        ArchiveEVUserExtraction.WriteLine("				right('0'+cast(day(d.archiveddate) as nvarchar(2)),2)+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				substring(cast(d.IdTransaction as nvarchar(99)),1,1)+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				substring(cast(d.IdTransaction as nvarchar(99)),2,3)+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				replace(cast(d.IdTransaction as nvarchar(99))+'.DVS','-','')")
        ArchiveEVUserExtraction.WriteLine("			ELSE ")
        ArchiveEVUserExtraction.WriteLine("				cast(year(d.IdDateTime) as nvarchar(4))+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				right('0'+cast(month(d.IdDateTime) as nvarchar(2)),2)+'-'+")
        ArchiveEVUserExtraction.WriteLine("				right('0'+cast(day(d.IdDateTime) as nvarchar(2)),2)+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				cast(datepart(hour,d.IdDateTime) as nvarchar(2))+'\\'+")
        ArchiveEVUserExtraction.WriteLine("				(CAST (d.IdChecksumHigh AS VARCHAR(7)) + ")
        ArchiveEVUserExtraction.WriteLine("				(CASE WHEN Len(d.IdChecksumHigh) = 1 THEN '000000'")
        ArchiveEVUserExtraction.WriteLine("					  ELSE '' END)) +")
        ArchiveEVUserExtraction.WriteLine("				(CASE WHEN Len(d.IdChecksumLow) = 1 THEN CAST(d.IdChecksumLow as nvarchar(1))+'0000000'")
        ArchiveEVUserExtraction.WriteLine("					  WHEN Len(d.IdChecksumLow) = 6 THEN '00'+CAST(d.IdChecksumLow as nvarchar(6))")
        ArchiveEVUserExtraction.WriteLine("					  WHEN Len(d.IdChecksumLow) = 7 THEN '0'+CAST(d.IdChecksumLow as nvarchar(7))")
        ArchiveEVUserExtraction.WriteLine("					  ELSE CAST(IdChecksumLow as nvarchar(8)) END)+'~'+")
        ArchiveEVUserExtraction.WriteLine("				REPLACE(REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(23),d.IdDateTime, 121),'-',''),' ',''),'.',''),':','')+'~'+")
        ArchiveEVUserExtraction.WriteLine("				(CAST (d.IdUniqueNo AS VARCHAR(36)) )+'.DVS'")
        ArchiveEVUserExtraction.WriteLine("			END))" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("	else")
        ArchiveEVUserExtraction.WriteLine("		rs = $statement1.execute_query(" & """" & "SELECT  d.SavesetIdentity, ISNULL(f.RelativeFileName,g.StoreIdentifier) as target")
        ArchiveEVUserExtraction.WriteLine("			 ")
        ArchiveEVUserExtraction.WriteLine("		  FROM #{$vaultdb1}.[dbo].[Archive] a")
        ArchiveEVUserExtraction.WriteLine("		  left join #{$vaultdb1}.dbo.Root b")
        ArchiveEVUserExtraction.WriteLine("		  on a.RootIdentity = ISNULL(b.containerrootidentity,b.rootidentity)")
        ArchiveEVUserExtraction.WriteLine("		  inner join #{db}.dbo.Vault c")
        ArchiveEVUserExtraction.WriteLine("		  on b.VaultEntryId = c.VaultId")
        ArchiveEVUserExtraction.WriteLine("		  inner join #{db}.dbo.Saveset d")
        ArchiveEVUserExtraction.WriteLine("		  on c.VaultIdentity = d.VaultIdentity")
        ArchiveEVUserExtraction.WriteLine("		  left join #{$vaultdb1}.dbo.ArchiveFolder e")
        ArchiveEVUserExtraction.WriteLine("		  on b.RootIdentity = e.RootIdentity")
        ArchiveEVUserExtraction.WriteLine("		  left join #{db}.dbo.Collection f")
        ArchiveEVUserExtraction.WriteLine("		  on d.CollectionIdentity = f.CollectionIdentity")
        ArchiveEVUserExtraction.WriteLine("		  left join #{db}.dbo.SavesetStore g")
        ArchiveEVUserExtraction.WriteLine("		  on d.SavesetIdentity = g.SavesetIdentity")
        ArchiveEVUserExtraction.WriteLine("		  ")
        ArchiveEVUserExtraction.WriteLine("		  where a.ArchiveName in (#{names})" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("	end")
        ArchiveEVUserExtraction.WriteLine("	while (rs.next) do")
        ArchiveEVUserExtraction.WriteLine("		targets << " & """" & "#{rs.getString('target')}" & """")
        ArchiveEVUserExtraction.WriteLine("		end")
        ArchiveEVUserExtraction.WriteLine("	}")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("	return targets.uniq")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def filesExist(targets)")
        ArchiveEVUserExtraction.WriteLine("#Nuix does not like you to pass it target files which don't exist which is an issue for EV so we break out missing files into a report")
        ArchiveEVUserExtraction.WriteLine("if $Centera == false")
        ArchiveEVUserExtraction.WriteLine("	missing = Array.new")
        ArchiveEVUserExtraction.WriteLine("	targets.delete_if do |target|")
        ArchiveEVUserExtraction.WriteLine("		file = target.delete(',')")
        ArchiveEVUserExtraction.WriteLine("		  if !File.file?(file)")
        ArchiveEVUserExtraction.WriteLine("			missing << file")
        ArchiveEVUserExtraction.WriteLine("			true ")
        ArchiveEVUserExtraction.WriteLine("		  end")
        ArchiveEVUserExtraction.WriteLine("		end")
        ArchiveEVUserExtraction.WriteLine("	return targets,missing")
        ArchiveEVUserExtraction.WriteLine("else")
        ArchiveEVUserExtraction.WriteLine("	missing = Array.new")
        ArchiveEVUserExtraction.WriteLine("	missing << " & """" & "Centera cluster targeted, missing file check skipped" & """")
        ArchiveEVUserExtraction.WriteLine("	missing << " & """" & "Centera IP file not specified/missing" & """" & "if $Centera_IPFile.nil? || !File.file?($Centera_ipFile)")
        ArchiveEVUserExtraction.WriteLine("	return targets, missing")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def mailboxFolders(dbmap)")
        ArchiveEVUserExtraction.WriteLine("#A report of user mailbox message counts by folder.  Eventually this could be expanded to include file sizes and other data from EV SQL")
        ArchiveEVUserExtraction.WriteLine("	mailboxes = Array.new")
        ArchiveEVUserExtraction.WriteLine("		dbmap.each {|db, names|")
        ArchiveEVUserExtraction.WriteLine("			rs = $statement1.execute_query(" & """" & "SELECT  ")
        ArchiveEVUserExtraction.WriteLine("				a.ArchiveName as Archive,")
        ArchiveEVUserExtraction.WriteLine("				ISNULL(e.RootIdentity, a.RootIdentity) as ID,")
        ArchiveEVUserExtraction.WriteLine("				ISNULL(cast(e.FolderPath as nvarchar(max)),'Archive '+a.ArchiveDescription) as Path,")
        ArchiveEVUserExtraction.WriteLine("				e.FolderName as FolderName,")
        ArchiveEVUserExtraction.WriteLine("				COUNT(d.idtransaction) as Count")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("			  FROM [#{$vaultdb1}].[dbo].[Archive] a")
        ArchiveEVUserExtraction.WriteLine("			  left join #{$vaultdb1}.dbo.Root b")
        ArchiveEVUserExtraction.WriteLine("			  on a.RootIdentity = ISNULL(b.containerrootidentity,b.rootidentity)")
        ArchiveEVUserExtraction.WriteLine("			  LEFT join #{db}.dbo.Vault c")
        ArchiveEVUserExtraction.WriteLine("			  on b.VaultEntryId = c.VaultId")
        ArchiveEVUserExtraction.WriteLine("			  left join #{db}.dbo.Saveset d")
        ArchiveEVUserExtraction.WriteLine("			  on c.VaultIdentity = d.VaultIdentity")
        ArchiveEVUserExtraction.WriteLine("			  left join #{$vaultdb1}.dbo.ArchiveFolder e")
        ArchiveEVUserExtraction.WriteLine("			  on b.RootIdentity = e.RootIdentity")
        ArchiveEVUserExtraction.WriteLine("			  left join #{db}.dbo.Collection f")
        ArchiveEVUserExtraction.WriteLine("			  on d.CollectionIdentity = f.CollectionIdentity")
        ArchiveEVUserExtraction.WriteLine("			  left join #{$vaultdb1}.dbo.PartitionEntry g")
        ArchiveEVUserExtraction.WriteLine("			  on d.IdPartition = g.IdPartition and a.VaultStoreEntryId = g.VaultStoreEntryId")
        ArchiveEVUserExtraction.WriteLine("			 where a.ArchiveName in (#{names})")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("			 group by a.ArchiveName, ISNULL(e.RootIdentity, a.RootIdentity),ISNULL(cast(e.FolderPath as nvarchar(max)),'Archive '+a.ArchiveDescription), e.FolderName")
        ArchiveEVUserExtraction.WriteLine("			 order by a.ArchiveName, ISNULL(e.RootIdentity, a.RootIdentity),ISNULL(cast(e.FolderPath as nvarchar(max)),'Archive '+a.ArchiveDescription), e.FolderName" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("				while (rs.next) do")
        ArchiveEVUserExtraction.WriteLine("					box = Array.new")
        ArchiveEVUserExtraction.WriteLine("					box << " & """" & "#{rs.getString('Archive')}" & """")
        ArchiveEVUserExtraction.WriteLine("					box << " & """" & "#{rs.getString('ID')}" & """")
        ArchiveEVUserExtraction.WriteLine("					box << " & """" & "#{rs.getString('Path')}" & """")
        ArchiveEVUserExtraction.WriteLine("					box << " & """" & "#{rs.getString('FolderName')}" & """")
        ArchiveEVUserExtraction.WriteLine("					box << " & """" & "#{rs.getString('Count')}" & """")
        ArchiveEVUserExtraction.WriteLine("					mailboxes << box.to_csv")
        ArchiveEVUserExtraction.WriteLine("				end")
        ArchiveEVUserExtraction.WriteLine("			}")
        ArchiveEVUserExtraction.WriteLine("	return mailboxes")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def mailboxVolumes(dbmap)")
        ArchiveEVUserExtraction.WriteLine("#A quick report of user mailbox message totals.  Eventually this could be expanded to include file sizes and other data")
        ArchiveEVUserExtraction.WriteLine("	mailboxes = Array.new")
        ArchiveEVUserExtraction.WriteLine("		dbmap.each {|db, names|")
        ArchiveEVUserExtraction.WriteLine("		rs = $statement1.execute_query(" & """" & "SELECT  ")
        ArchiveEVUserExtraction.WriteLine("			a.ArchiveName as Archive,")
        ArchiveEVUserExtraction.WriteLine("			COUNT(d.idtransaction) as Count, SUM(d.ItemSize)/1024 AS ArchivedItemSizeMB")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("		  FROM [#{$vaultdb1}].[dbo].[Archive] a")
        ArchiveEVUserExtraction.WriteLine("		  left join #{$vaultdb1}.dbo.Root b")
        ArchiveEVUserExtraction.WriteLine("		  on a.RootIdentity = ISNULL(b.containerrootidentity,b.rootidentity)")
        ArchiveEVUserExtraction.WriteLine("		  LEFT join #{db}.dbo.Vault c")
        ArchiveEVUserExtraction.WriteLine("			on b.VaultEntryId = c.VaultId")
        ArchiveEVUserExtraction.WriteLine("		  left join #{db}.dbo.Saveset d")
        ArchiveEVUserExtraction.WriteLine("		  on c.VaultIdentity = d.VaultIdentity")
        ArchiveEVUserExtraction.WriteLine("		  left join #{$vaultdb1}.dbo.ArchiveFolder e")
        ArchiveEVUserExtraction.WriteLine("		  on b.RootIdentity = e.RootIdentity")
        ArchiveEVUserExtraction.WriteLine("		  left join #{db}.dbo.Collection f")
        ArchiveEVUserExtraction.WriteLine("		  on d.CollectionIdentity = f.CollectionIdentity")
        ArchiveEVUserExtraction.WriteLine("		  left join #{$vaultdb1}.dbo.PartitionEntry g")
        ArchiveEVUserExtraction.WriteLine("		  on d.IdPartition = g.IdPartition and a.VaultStoreEntryId = g.VaultStoreEntryId")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("		where a.ArchiveName in (#{names})")
        ArchiveEVUserExtraction.WriteLine("		 group by a.ArchiveName")
        ArchiveEVUserExtraction.WriteLine("		 order by a.ArchiveName" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("			while (rs.next) do")
        ArchiveEVUserExtraction.WriteLine("				box = Array.new")
        ArchiveEVUserExtraction.WriteLine("				box << " & """" & "#{rs.getString('Archive')}" & """")
        ArchiveEVUserExtraction.WriteLine("				box << " & """" & "#{rs.getString('Count')}" & """")
        ArchiveEVUserExtraction.WriteLine("				box << " & """" & "#{rs.getString('ArchivedItemSizeMB')}" & """")
        ArchiveEVUserExtraction.WriteLine("				mailboxes << box.to_csv")
        ArchiveEVUserExtraction.WriteLine("			end")
        ArchiveEVUserExtraction.WriteLine("			}")
        ArchiveEVUserExtraction.WriteLine("	return mailboxes")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def makeManifest(dbmap)")
        ArchiveEVUserExtraction.WriteLine("#This queries EV for the Transaction IDs used to pass mailbox and folder information and filter content from CAB containers so only targeted messages are exported")
        ArchiveEVUserExtraction.WriteLine("	targets = Array.new")
        ArchiveEVUserExtraction.WriteLine("		dbmap.each {|db, names|")
        ArchiveEVUserExtraction.WriteLine("		rs = $statement1.execute_query(" & """" & "SELECT  d.IdTransaction, a.ArchiveName, e.FolderPath")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("		  FROM [#{$vaultdb1}].[dbo].[Archive] a")
        ArchiveEVUserExtraction.WriteLine("		  left join #{$vaultdb1}.dbo.Root b")
        ArchiveEVUserExtraction.WriteLine("		  on a.RootIdentity = ISNULL(b.containerrootidentity,b.rootidentity)")
        ArchiveEVUserExtraction.WriteLine("		  inner join #{db}.dbo.Vault c")
        ArchiveEVUserExtraction.WriteLine("		  on b.VaultEntryId = c.VaultId")
        ArchiveEVUserExtraction.WriteLine("		  inner join #{db}.dbo.Saveset d")
        ArchiveEVUserExtraction.WriteLine("		  on c.VaultIdentity = d.VaultIdentity")
        ArchiveEVUserExtraction.WriteLine("		  left join #{$vaultdb1}.dbo.ArchiveFolder e")
        ArchiveEVUserExtraction.WriteLine("		  on b.RootIdentity = e.RootIdentity")
        ArchiveEVUserExtraction.WriteLine("		  left join #{db}.dbo.Collection f")
        ArchiveEVUserExtraction.WriteLine("		  on d.CollectionIdentity = f.CollectionIdentity")
        ArchiveEVUserExtraction.WriteLine("		  left join #{$vaultdb1}.dbo.PartitionEntry g")
        ArchiveEVUserExtraction.WriteLine("		  on d.IdPartition = g.IdPartition and a.VaultStoreEntryId = g.VaultStoreEntryId")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("		  where a.ArchiveName in (#{names})" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("			while (rs.next) do")
        ArchiveEVUserExtraction.WriteLine("				target = Array.new")
        ArchiveEVUserExtraction.WriteLine("				target << " & """" & "#{rs.getString('IdTransaction').delete(" & """" & "-" & """" & ")}" & """")
        ArchiveEVUserExtraction.WriteLine("				target << " & """" & "#{rs.getString('ArchiveName')}" & """")
        ArchiveEVUserExtraction.WriteLine("				target << " & """" & "#{rs.getString('FolderPath')}" & """")
        ArchiveEVUserExtraction.WriteLine("				targets << target.to_csv")
        ArchiveEVUserExtraction.WriteLine("			end")
        ArchiveEVUserExtraction.WriteLine("		}")
        ArchiveEVUserExtraction.WriteLine("	return targets")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def makeConfig(targets,create_config)")
        ArchiveEVUserExtraction.WriteLine("#This creates the PST mapping file automatically, normailizing commas and spaces.  The PST mapping could be made manually, this was just to see if the script itself could create all of the necessary config info from SQL and user input")
        ArchiveEVUserExtraction.WriteLine("psts = Array.new")
        ArchiveEVUserExtraction.WriteLine("file_extension = " & """" & """")
        ArchiveEVUserExtraction.WriteLine("file_extension = " & """" & ".pst" & """" & " if create_config.downcase == " & """" & "pst" & """")
        ArchiveEVUserExtraction.WriteLine("targets.each do |target|")
        ArchiveEVUserExtraction.WriteLine("	pst = Array.new")
        ArchiveEVUserExtraction.WriteLine("	pst << CSV.parse_line(target)[1]")
        ArchiveEVUserExtraction.WriteLine("	pst << CSV.parse_line(target)[1].delete(" & """" & "," & """" & ").gsub(/[\x00\/\\:\*\?\" & """" & "<>\|]/, '-')")
        ArchiveEVUserExtraction.WriteLine("	if !psts.include? (pst.to_csv)")
        ArchiveEVUserExtraction.WriteLine("		psts << pst.to_csv")
        ArchiveEVUserExtraction.WriteLine("	end")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("return psts")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def makeJSON(targets)")
        ArchiveEVUserExtraction.WriteLine("	hash = Hash.new {|h,k| h[k] = []}")
        ArchiveEVUserExtraction.WriteLine("	json = Hash.new")
        ArchiveEVUserExtraction.WriteLine("	targets.each do |target|")
        ArchiveEVUserExtraction.WriteLine("		root = target.split(',')[0]")
        ArchiveEVUserExtraction.WriteLine("		path =  target.split(',')[1]")
        ArchiveEVUserExtraction.WriteLine("		hash[" & """" & "#{root}" & """" & "] << path")
        ArchiveEVUserExtraction.WriteLine("	end")
        ArchiveEVUserExtraction.WriteLine(" hash.each {|key, value| hash[ " & """" & "#{key}" & """" & "] = Hash[(1...hash[ " & """" & "#{key}" & """" & "].size+1).zip hash[ " & """" & "#{key}" & """" & "]] }")
        ArchiveEVUserExtraction.WriteLine("	json[" & """" & "dvs_list" & """" & "] = hash")
        ArchiveEVUserExtraction.WriteLine("	return json")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("def writeReport(name, filename, content)")
        ArchiveEVUserExtraction.WriteLine("#Just takes an array of data (formatted to CSV already in my functions) and dumps it to a file on disk, used to spit out the various CSV reports and config files")
        ArchiveEVUserExtraction.WriteLine("	report_file = File.new(" & """" & "#{filename}" & """" & ", " & """" & "w+" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("	report_file.puts content")
        ArchiveEVUserExtraction.WriteLine("	report_file.close")
        ArchiveEVUserExtraction.WriteLine("	puts " & """" & "#{name}: #{filename}" & """")
        ArchiveEVUserExtraction.WriteLine("end ")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("#script variables")
        ArchiveEVUserExtraction.WriteLine("names = ''")
        ArchiveEVUserExtraction.WriteLine("stores = Array.new")
        ArchiveEVUserExtraction.WriteLine("selected = Array.new")
        ArchiveEVUserExtraction.WriteLine("targets = Array.new")
        ArchiveEVUserExtraction.WriteLine("list = Array.new")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("#User Selected Archives")
        ArchiveEVUserExtraction.WriteLine("if !(csv_path.nil?)")
        ArchiveEVUserExtraction.WriteLine("	selected,list = getArchiveVaultIDs(getFromCSV(csv_path))")
        ArchiveEVUserExtraction.WriteLine("else")
        ArchiveEVUserExtraction.WriteLine("stores = getUserSelectedStores()")
        ArchiveEVUserExtraction.WriteLine("selected,list = getStoreSelectedArchives(stores)")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("names = " & """" & "'#{list.join(" & """" & "','" & """" & ")}'" & """")
        ArchiveEVUserExtraction.WriteLine("puts " & """" & "Searching for target files for archives: #{names}" & """")
        ArchiveEVUserExtraction.WriteLine("dbmap = getVaults(selected)")
        ArchiveEVUserExtraction.WriteLine("targets = getTargetFiles(dbmap)")
        ArchiveEVUserExtraction.WriteLine("targets,missing = filesExist(targets)")
        ArchiveEVUserExtraction.WriteLine("puts " & """" & "---" & """")
        ArchiveEVUserExtraction.WriteLine("puts " & """" & "Targeting Centera clips" & """" & "if $Centera")
        ArchiveEVUserExtraction.WriteLine("#Write reports")
        ArchiveEVUserExtraction.WriteLine("puts " & """" & "Writing Reports" & """")
        ArchiveEVUserExtraction.WriteLine("writeReport(" & """" & "Target files" & """" & ", " & """" & "#{report_path}" & """" & " + " & """" & "\\target_files.csv" & """" & ", targets)")
        ArchiveEVUserExtraction.WriteLine("writeReport(" & """" & "Missing files" & """" & ", " & """" & "#{report_path}" & """" & " + " & """" & "\\missing_files.csv" & """" & ", missing)")
        ArchiveEVUserExtraction.WriteLine("writeReport(" & """" & "Mailbox Folders" & """" & ", " & """" & "#{report_path}" & """" & " + " & """" & "\\mailbox_report.csv" & """" & ", mailboxFolders(dbmap))")
        ArchiveEVUserExtraction.WriteLine("writeReport(" & """" & "Mailbox Totals" & """" & ", " & """" & "#{report_path}" & """" & " + " & """" & "\\mailbox_totals.csv" & """" & ", mailboxVolumes(dbmap))")
        ArchiveEVUserExtraction.WriteLine("puts " & """" & "---" & """")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("manifest = makeManifest(dbmap)")
        ArchiveEVUserExtraction.WriteLine("if (noCnD.nil?)")
        ArchiveEVUserExtraction.WriteLine("	puts(" & """" & "Export to #{CND_config}" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("	puts " & """" & "---" & """")
        ArchiveEVUserExtraction.WriteLine("	#Prep CrackAndDump config files")
        ArchiveEVUserExtraction.WriteLine("	puts " & """" & "Prepping config files" & """")
        ArchiveEVUserExtraction.WriteLine("	writeReport(" & """" & "Manifest files" & """" & "," & """" & "#{manifest_dir}" & """" & ",manifest)")
        ArchiveEVUserExtraction.WriteLine("	if CND_config.include?(" & """" & "configFile=>" & """" & ") & !(create_config.nil?)")
        ArchiveEVUserExtraction.WriteLine("		config_name = CND_config.split(" & """" & "configFile=>" & """" & ")[1].split(" & """" & ";" & """" & ")[0]")
        ArchiveEVUserExtraction.WriteLine("		writeReport(" & """" & "Auto-generating config file from target archive names to" & """" & "," & """" & "#{config_name}" & """" & ", makeConfig(manifest,create_config))")
        ArchiveEVUserExtraction.WriteLine("	end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("	puts " & """" & "Target export folder: #{export_path}" & """")
        ArchiveEVUserExtraction.WriteLine("	puts " & """" & "---" & """")
        ArchiveEVUserExtraction.WriteLine("else")
        ArchiveEVUserExtraction.WriteLine("	writeReport(" & """" & "Manifest files" & """" & "," & """" & "#{report_path}" & """" & "+" & """" & "\\manifest.csv" & """" & ",manifest)")
        ArchiveEVUserExtraction.WriteLine("	writeReport(" & """" & "Auto-generating config file from target archive names to" & """" & "," & """" & "#{report_path}" & """" & "+" & """" & "\\auto_config.csv" & """" & ",makeConfig(manifest,create_config)) if !(create_config.nil?)")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("# Export Folders")
        ArchiveEVUserExtraction.WriteLine("callback_count = 0")
        ArchiveEVUserExtraction.WriteLine("d = DateTime.now")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("   archive_file = parsed[" & """" & "email_archive" & """" & "][" & """" & "folder" & """" & "]")
        ArchiveEVUserExtraction.WriteLine("   archive_name = parsed[" & """" & "email_archive" & """" & "][" & """" & "evidence_name" & """" & "]")
        ArchiveEVUserExtraction.WriteLine("   archive_centera = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera" & """" & "]")
        ArchiveEVUserExtraction.WriteLine("   archive_centera_ip = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera_ip" & """" & "]")
        ArchiveEVUserExtraction.WriteLine("   archive_centera_clip = parsed[" & """" & "email_archive" & """" & "][" & """" & "centera_clip" & """" & "]")
        ArchiveEVUserExtraction.WriteLine("   archive_migration = parsed[" & """" & "email_archive" & """" & "][" & """" & "migration" & """" & "]")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("caseFactory = $utilities.getCaseFactory()")
        ArchiveEVUserExtraction.WriteLine("case_settings = {")
        ArchiveEVUserExtraction.WriteLine("    :compound => false,")
        ArchiveEVUserExtraction.WriteLine("    :name => " & """" & sCaseName & """" & ",")
        ArchiveEVUserExtraction.WriteLine("    :description => " & """" & """" & ",")
        ArchiveEVUserExtraction.WriteLine("    :investigator => " & """" & "Email Archive Migration Manager" & """")
        ArchiveEVUserExtraction.WriteLine(" }")
        ArchiveEVUserExtraction.WriteLine("$current_case = caseFactory.create(" & """" & sNuixCaseDir.Replace("\", "\\") & """" & ", case_settings)")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("kinds_to_keep = {")
        ArchiveEVUserExtraction.WriteLine("	" & """" & "email" & """" & " => true,")
        ArchiveEVUserExtraction.WriteLine("	" & """" & "calendar" & """" & " => true,")
        ArchiveEVUserExtraction.WriteLine("	" & """" & "contact" & """" & "=> true,")
        ArchiveEVUserExtraction.WriteLine("	" & """" & "container" & """" & " => true,")
        ArchiveEVUserExtraction.WriteLine("	" & """" & "no-data" & """" & "=> true,")
        ArchiveEVUserExtraction.WriteLine("	" & """" & "system" & """" & "=> true,")
        ArchiveEVUserExtraction.WriteLine("	" & """" & "unrecognised" & """" & "=> true")
        ArchiveEVUserExtraction.WriteLine(" }")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("processor = $current_case.createProcessor")
        ArchiveEVUserExtraction.WriteLine("#if archive_migration == " & """" & "lightspeed" & """")
        ArchiveEVUserExtraction.WriteLine("	processing_settings = {")
        ArchiveEVUserExtraction.WriteLine("		:traversalScope => " & """" & "full_traversal" & """" & ",")
        ArchiveEVUserExtraction.WriteLine("		:analysisLanguage => " & """" & "en" & """" & ",")
        ArchiveEVUserExtraction.WriteLine("		:identifyPhysicalFiles => true,")
        ArchiveEVUserExtraction.WriteLine("		:reuseEvidenceStores => true,")
        ArchiveEVUserExtraction.WriteLine("		:reportProcessingStatus => " & """" & "none" & """")
        ArchiveEVUserExtraction.WriteLine("	}")
        ArchiveEVUserExtraction.WriteLine("	processor.setProcessingSettings(processing_settings)")
        ArchiveEVUserExtraction.WriteLine("#	else")
        ArchiveEVUserExtraction.WriteLine("		# processing_settings = {")
        ArchiveEVUserExtraction.WriteLine("			# :processText => true,")
        ArchiveEVUserExtraction.WriteLine("			# :traversalScope => " & """" & "full_traversal" & """" & ",")
        ArchiveEVUserExtraction.WriteLine("			# :analysisLanguage => " & """" & "en" & """" & ",")
        ArchiveEVUserExtraction.WriteLine("			# :stopWords => false,")
        ArchiveEVUserExtraction.WriteLine("			# :stemming => false,")
        ArchiveEVUserExtraction.WriteLine("			# :enableExactQueries => false,")
        ArchiveEVUserExtraction.WriteLine("			# :extractNamedEntitiesFromText => false,")
        ArchiveEVUserExtraction.WriteLine("			# :extractNamedEntitiesFromProperties => false,")
        ArchiveEVUserExtraction.WriteLine("			# :extractNamedEntitiesFromTextStripped => false,")
        ArchiveEVUserExtraction.WriteLine("			# :extractShingles => false,")
        ArchiveEVUserExtraction.WriteLine("			# :processTextSummaries => false,")
        ArchiveEVUserExtraction.WriteLine("			# :extractFromSlackSpace => false,")
        ArchiveEVUserExtraction.WriteLine("			# :carveUnidentifiedData => false,")
        ArchiveEVUserExtraction.WriteLine("			# :recoverDeletedFiles => false,")
        ArchiveEVUserExtraction.WriteLine("			# :extractEndOfFileSlackSpace => false,")
        ArchiveEVUserExtraction.WriteLine("			# :identifyPhysicalFiles => true,")
        ArchiveEVUserExtraction.WriteLine("			# :createThumbnails => false,")
        ArchiveEVUserExtraction.WriteLine("			# :skinToneAnalysis => false,")
        ArchiveEVUserExtraction.WriteLine("			# :calculateAuditedSize => true,")
        ArchiveEVUserExtraction.WriteLine("			# :storeBinary => false,")
        ArchiveEVUserExtraction.WriteLine("			# :maxStoredBinarySize => 250000000,")
        ArchiveEVUserExtraction.WriteLine("			# :maxDigestSize => 250000000,")
        ArchiveEVUserExtraction.WriteLine("			# :digests => " & """" & "MD5" & """" & ",")
        ArchiveEVUserExtraction.WriteLine("			# :addBccToEmailDigests => false,")
        ArchiveEVUserExtraction.WriteLine("			# :addCommunicationDateToEmailDigests => false,")
        ArchiveEVUserExtraction.WriteLine("			# :reuseEvidenceStores => true,")
        ArchiveEVUserExtraction.WriteLine("			# :processFamilyFields => false,")
        ArchiveEVUserExtraction.WriteLine("			# :hideEmbeddedImmaterialData => false,")
        If iEvidenceSize = 0 Then
            ArchiveEVUserExtraction.WriteLine("			# :reportProcessingStatus => " & """" & "none," & """")
        Else
            ArchiveEVUserExtraction.WriteLine("			# :reportProcessingStatus => " & """" & "physical_files," & """")
        End If
        ArchiveEVUserExtraction.WriteLine("		# }")
        ArchiveEVUserExtraction.WriteLine("		# processor.setProcessingSettings(processing_settings)")
        ArchiveEVUserExtraction.WriteLine("# end")
        ArchiveEVUserExtraction.WriteLine("parallel_processing_settings = {")
        ArchiveEVUserExtraction.WriteLine("	:workerCount => " & sNumberOfWorkers & ",")
        ArchiveEVUserExtraction.WriteLine("	:workerMemory => " & sMemoryPerWorker & ",")
        ArchiveEVUserExtraction.WriteLine("	:embedBroker => true,")
        ArchiveEVUserExtraction.WriteLine("	:brokerMemory => " & sMemoryPerWorker & ",")
        ArchiveEVUserExtraction.WriteLine(" :workerTemp => " & """" & sWorkerTempDir.Replace("\", "\\") & """")

        ArchiveEVUserExtraction.WriteLine("}")
        ArchiveEVUserExtraction.WriteLine("processor.setParallelProcessingSettings(parallel_processing_settings)")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("# MIME Type Fiter for row-based items")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/csv" & """" & ", { process_embedded: false })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/tab-separated-values" & """" & ", { process_embedded: false })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.sqlite-database" & """" & ", { process_embedded: false, text_strip: true })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-registry" & """" & ", { text_strip: true})")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/x-plist" & """" & ", { process_embedded: false })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-logfile" & """" & ", { process_embedded: false, text_strip: true })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-mft" & """" & ", { process_embedded: false, text_strip: true })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "filesystem/x-ntfs-usnjrnl" & """" & ", { process_embedded: false, text_strip: true })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/exe" & """" & ", { process_text: false })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.tcpdump.pcap" & """" & ", { process_embedded: false, text_strip: true })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "text/x-common-log" & """" & ", { process_embedded: false, text_strip: true })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-iis-log" & """" & ", { process_embedded: false, text_strip: true })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-windows-event-log-record" & """" & ", { process_embedded: false, text_strip: true })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-windows-event-logx" & """" & ", { process_embedded: false, text_strip: true })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/x-pcapng" & """" & ", { process_embedded: false, text_strip: true })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.symantec-vault-stream-data" & """" & ", { enabled: false })")
        ArchiveEVUserExtraction.WriteLine("processor.set_mime_type_processing_settings(" & """" & "application/vnd.ms-cab-compressed" & """" & ", { process_embedded: true })")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("if archive_centera == " & """" & "no" & """")
        ArchiveEVUserExtraction.WriteLine("	file_count = 0")
        ArchiveEVUserExtraction.WriteLine("	evidence_num = 1")
        ArchiveEVUserExtraction.WriteLine("	folder = processor.newEvidenceContainer(" & """" & "Evidence" & """" & "+evidence_num.to_s)")
        ArchiveEVUserExtraction.WriteLine("	targets.each do |target|")
        ArchiveEVUserExtraction.WriteLine("		file_count = file_count + 1")
        ArchiveEVUserExtraction.WriteLine("		if file_count % evidence_size.to_i == 0")
        ArchiveEVUserExtraction.WriteLine("			folder.save")
        ArchiveEVUserExtraction.WriteLine("			evidence_num = evidence_num + 1")
        ArchiveEVUserExtraction.WriteLine("			folder = processor.newEvidenceContainer(" & """" & "Evidence" & """" & "+evidence_num.to_s)")
        ArchiveEVUserExtraction.WriteLine("		end")
        ArchiveEVUserExtraction.WriteLine("		file = target.delete(',')")
        ArchiveEVUserExtraction.WriteLine("		folder.setEncoding(" & """" & "utf-8" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("		folder.add_file(file)")
        ArchiveEVUserExtraction.WriteLine("	end")
        ArchiveEVUserExtraction.WriteLine("	folder.save")
        ArchiveEVUserExtraction.WriteLine("	end")

        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("if archive_centera == " & """" & "yes" & """")
        ArchiveEVUserExtraction.WriteLine("	evidence_num = 1")
        ArchiveEVUserExtraction.WriteLine("	folder = processor.newEvidenceContainer(" & """" & "Evidence" & """" & "+evidence_num.to_s)")
        ArchiveEVUserExtraction.WriteLine("	folder.setEncoding(" & """" & "utf-8" & """" & ")")
        ArchiveEVUserExtraction.WriteLine("	folder.addCenteraCluster({")
        ArchiveEVUserExtraction.WriteLine("		:ipsFile => " & """" & "#{archive_centera_ip}" & """" & ",")
        ArchiveEVUserExtraction.WriteLine("		:clipsFile => report_path+" & """" & "\\target_files.csv" & """")
        ArchiveEVUserExtraction.WriteLine("	 })")
        ArchiveEVUserExtraction.WriteLine("	 folder.save")
        ArchiveEVUserExtraction.WriteLine("	 end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("start_time = Time.now")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("puts")
        ArchiveEVUserExtraction.WriteLine("puts " & """" & "Email Archive Extraction has started..." & """")
        ArchiveEVUserExtraction.WriteLine("puts")
        ArchiveEVUserExtraction.WriteLine("printf " & """" & "%-40s %-25s" & """" & ", " & """" & "Timestamp" & """" & ", " & """" & "Processed Items" & """")
        ArchiveEVUserExtraction.WriteLine("puts")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("processor.when_item_processed do |event|")
        ArchiveEVUserExtraction.WriteLine("  if callback_count % CALLBACK_FREQUENCY == 0")
        ArchiveEVUserExtraction.WriteLine("	printf " & """" & "\r%-40s %-25s" & """" & ", Time.now, callback_count")
        ArchiveEVUserExtraction.WriteLine("	updated_callback = [callback_count]")
        ArchiveEVUserExtraction.WriteLine("    db.update(" & """" & "UPDATE archiveExtractionStats SET ItemsProcessed = ? WHERE BatchName = '" & sCaseName & "'" & """" & ",updated_callback)")
        ArchiveEVUserExtraction.WriteLine("  end")
        ArchiveEVUserExtraction.WriteLine("  callback_count += 1")
        ArchiveEVUserExtraction.WriteLine("end")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("processor.process")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("puts")
        ArchiveEVUserExtraction.WriteLine("end_time = Time.now")
        ArchiveEVUserExtraction.WriteLine("total_time = '%.2f' % ((end_time-start_time)/60)")
        ArchiveEVUserExtraction.WriteLine("display_time =  '%.0f' % (total_time)")
        ArchiveEVUserExtraction.WriteLine("puts")
        ArchiveEVUserExtraction.WriteLine("puts " & """" & "Email Archive Extraction has finished!" & """")
        ArchiveEVUserExtraction.WriteLine("puts")
        ArchiveEVUserExtraction.WriteLine("puts " & """" & "Completed in #{display_time} minutes." & """")
        ArchiveEVUserExtraction.WriteLine("puts")
        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("updated_callback = [" & """" & "Completed" & """" & ", callback_count, 0, 100, end_time]")
        ArchiveEVUserExtraction.WriteLine("db.update(" & """" & "UPDATE archiveExtractionStats SET ExtractionStatus = ?, ItemsProcessed = ?, BytesProcessed = ?, PercentCompleted = ?,ProcessEndTIme = ? WHERE BatchName = '#{archive_name}'" & """" & ",updated_callback)")

        ArchiveEVUserExtraction.WriteLine("")
        ArchiveEVUserExtraction.WriteLine("$current_case.close")
        ArchiveEVUserExtraction.WriteLine("return")

        ArchiveEVUserExtraction.Close()


        blnBuildArchiveEVUserExtractionRuby = True
    End Function

    Public Function blnGetAllExporterMetrics(ByVal sCaseDirectory As String, ByRef lstExporterMetrics As List(Of String)) As Boolean
        Dim CurrentDirectory As DirectoryInfo
        Dim asSubDirectories(0) As String

        Array.Clear(asSubDirectories, 0, asSubDirectories.Length)
        ReDim asSubDirectories(0)

        Try


            CurrentDirectory = New DirectoryInfo(sCaseDirectory)
            If Not CurrentDirectory.Attributes.HasFlag(FileAttributes.ReadOnly) Then
                Try
                    For Each Directory In CurrentDirectory.GetDirectories
                        blnGetAllExporterMetrics(Directory.FullName, lstExporterMetrics)
                    Next

                    For Each Files In CurrentDirectory.GetFiles
                        If Files.Name = "exporter-metrics.csv" Then
                            lstExporterMetrics.Add(Files.FullName)
                        End If
                    Next
                Catch ex As Exception
                    MessageBox.Show("blnGetExporterMetrics  -" & ex.Message)
                End Try
            End If

        Catch ex As Exception
            MessageBox.Show("Error in blnGetExporterMetrics - " & ex.ToString, "blnGetExporterMetrics Error")
        End Try

        blnGetAllExporterMetrics = True

    End Function

    Public Function blnGetAllExporterErrors(ByVal sCaseDirectory As String, ByRef lstExportErrors As List(Of String)) As Boolean
        Dim CurrentDirectory As DirectoryInfo
        Dim asSubDirectories(0) As String

        Array.Clear(asSubDirectories, 0, asSubDirectories.Length)
        ReDim asSubDirectories(0)

        Try


            CurrentDirectory = New DirectoryInfo(sCaseDirectory)
            If Not CurrentDirectory.Attributes.HasFlag(FileAttributes.ReadOnly) Then
                Try
                    For Each Directory In CurrentDirectory.GetDirectories
                        blnGetAllExporterErrors(Directory.FullName, lstExportErrors)
                    Next

                    For Each Files In CurrentDirectory.GetFiles
                        If Files.Name = "exporter-error.csv" Then
                            lstExportErrors.Add(Files.FullName)
                        End If
                    Next
                Catch ex As Exception
                    MessageBox.Show("blnGetAllExporterErrors  -" & ex.Message)
                End Try
            End If

        Catch ex As Exception
            MessageBox.Show("Error in blnGetAllExporterErrors - " & ex.ToString, "blnGetAllExporter Error")
        End Try

        blnGetAllExporterErrors = True

    End Function


    Public Function blnGetAllFailureLogs(ByVal sExportDirectory As String, ByRef lstFailureLogs As List(Of String)) As Boolean
        Dim CurrentDirectory As DirectoryInfo
        Dim asSubDirectories(0) As String

        Array.Clear(asSubDirectories, 0, asSubDirectories.Length)
        ReDim asSubDirectories(0)

        Try


            CurrentDirectory = New DirectoryInfo(sExportDirectory)
            If Not CurrentDirectory.Attributes.HasFlag(FileAttributes.ReadOnly) Then
                Try
                    For Each Directory In CurrentDirectory.GetDirectories
                        blnGetAllFailureLogs(Directory.FullName, lstFailureLogs)
                    Next

                    For Each Files In CurrentDirectory.GetFiles
                        If Files.Name = "Mailbox.log" Then
                            lstFailureLogs.Add(Files.FullName)
                        End If
                    Next
                Catch ex As Exception
                    MessageBox.Show("GetNuixLogFiles -" & ex.Message)
                End Try
            End If

        Catch ex As Exception
            MessageBox.Show("Error in blnGetAllFailureLogs - " & ex.ToString, "blnGetAllFailuresLogs Error")
        End Try

        blnGetAllFailureLogs = True

    End Function

End Module
