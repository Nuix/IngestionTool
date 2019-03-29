Imports System.Data.SQLite
Imports System.Data.SqlClient


Public Class DatabaseService
    'Dim sSqlDataBaseLocation As String = eMailArchiveMigrationManager.SQLiteDBLocation


    Public Function duplicateDatabase(ByVal sDatabaseLocation As String) As Boolean
        Return System.IO.File.Exists(sDatabaseLocation)
    End Function

    Public Sub CreateDataBase(ByVal sSqlDataBaseFullName As String, ByVal sSqlDataBaseLocation As String)

        Dim connectionString As String = String.Format("Data Source = {0}", sSqlDataBaseFullName)

        Dim sArchiveTableQuery As String
        Dim sEWSExtractionTableQuery As String
        Dim sEWSIngestionTableQuery As String
        Dim sDataConversionTableQuery As String
        Dim sSourceDataInformationTableQuery As String
        Dim sEWSReprocessingTableQuery As String
        Dim Connection As SQLiteConnection


        sArchiveTableQuery = "Create Table archiveExtractionStats(ArchiveName TEXT, BatchName TEXT, ExtractionStatus TEXT, PercentCompleted Int, TotalBytes Int, BytesProcessed Int, ArchiveSettings TEXT, ArchiveType TEXT, OutputFormat TEXT, SQLSettings TEXT, SourceInformation TEXT, WSSSettings TEXT, ProcessStartTime TEXT, ProcessEndTime TEXT, ItemsProcessed INT, ItemsExported INT, ItemsSkipped INT, ItemsFailed INT, ProcessingFilesDirectory TEXT, CaseDirectory TEXT, OutputDirectory TEXT, LogDirectory TEXT, SummaryReportLocation TEXT)"
        sEWSExtractionTableQuery = "Create Table ewsExtractionStats(CustodianName TEXT, ExtractionRoot VARCHAR(25), GroupID VARCHAR (50), ExtractionStatus VARCHAR(25), FromDate VARCHAR(25), ToDate VARCHAR(25), ProcessID VARCHAR(10), ExtractionStartTime VARCHAR(50), ExtractionEndTime VARCHAR(50), ProcessedItems INT, ExtractionSize INT, ProcessingFilesDirectory TEXT, CaseDirectory TEXT, OutputDirectory TEXT, LogDirectory TEXT, SummaryReportLocation TEXT, OutputFormat TEXT)"
        sEWSIngestionTableQuery = "Create table ewsIngestionStats (CustodianName Text, ProgressStatus Text, PSTPath Text, NumberOfPSTs Int, TotalSizeOfPSTs Int, GroupID Text, DestinationFolder Text, DestinationRoot Text, DestinationSMTP Text, ProcessID Text, IngestionStartTime VARCHAR(50), IngestionEndTime VARCHAR(50), BytesUploaded Int, PercentCompleted Int, Success Int, Failed Int, ProcessingFilesDirectory TEXT, CaseDirectory TEXT, OutputDirectory TEXT, LogDirectory TEXT, SummaryReportLocation TEXT)"
        sDataConversionTableQuery = "Create table DataConversionStats (CustodianName Text, GroupID Text, ConversionStatus Text, ProcessID Text, RedisSettings Text, SourceType TEXT, SourceFormat TExT, OutputType Text, OutputFormat Text, ItemsProcessed INT, TotalSizeOfSource INT, BytesProcessed Int, PercentCompleted Int, Success Int, Failed Int, ConversionStartTime VARCHAR(50), ConversionEndTime VARCHAR(50), ProcessingFilesDirectory TEXT, CaseDirectory TEXT, OutputDirectory TEXT, LogDirectory TEXT, SummaryReportLocation TEXT)"
        sEWSReprocessingTableQuery = "Create table EWSReprocessingStats (CustodianName Text, GroupID Text, ReprocessingStatus Text, ProcessID Text, ItemsProcessed INT, TotalSizeOfFailures INT, SourceFormat TEXT, OutputFormat Text, ReprocessingStartTime VARCHAR(50), ReprocessingEndTime VARCHAR(50), BytesProcessed Int, PercentCompleted Int, Success Int, Failed Int, ProcessingFilesDirectory TEXT, CaseDirectory TEXT, OutputDirectory TEXT, LogDirectory TEXT, SummaryReportLocation TEXT)"
        sSourceDataInformationTableQuery = "Create table SourceDataInfoStats (SourceFileName Text, SourceFilePath Text, SourceFileCreateDate VARCHAR(50), SourceFileModifiedDate VARCHAR(50), SourceFileSize Int, CustodianID Int)"

        If Not duplicateDatabase(sSqlDataBaseFullName) Then
            SQLiteConnection.CreateFile(sSqlDataBaseFullName)

            Connection = New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;New=False;Compress=True;")
            Connection.Open()

            System.IO.Directory.CreateDirectory(sSqlDataBaseLocation)
            Using SqlConn As New SQLiteConnection(connectionString)
                SqlConn.Open()
                Dim ArchiveCommand As New SQLiteCommand(sArchiveTableQuery, SqlConn)
                ArchiveCommand.ExecuteNonQuery()

                Dim EWSExtractCommand As New SQLiteCommand(sEWSExtractionTableQuery, SqlConn)
                EWSExtractCommand.ExecuteNonQuery()

                Dim EWSIngestCommand As New SQLiteCommand(sEWSIngestionTableQuery, SqlConn)
                EWSIngestCommand.ExecuteNonQuery()

                Dim DataConversionCommand As New SQLiteCommand(sDataConversionTableQuery, SqlConn)
                DataConversionCommand.ExecuteNonQuery()

                Dim EWSReprocessingCommand As New SQLiteCommand(sEWSReprocessingTableQuery, SqlConn)
                EWSReprocessingCommand.ExecuteNonQuery()

                Dim SourceDataInfoCommand As New SQLiteCommand(sSourceDataInformationTableQuery, SqlConn)
                SourceDataInfoCommand.ExecuteNonQuery()

            End Using

        End If

    End Sub

    Public Function UpdateCustodianDBInfo(ByVal sSQLiteLocation As String, ByVal sCustodianName As String, ByVal sFieldName As String, sFieldValue As String) As Boolean
        UpdateCustodianDBInfo = False
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

        UpdateCustodianDBInfo = True
    End Function

    Public Function UpdateExtractionDBInfo(ByVal sSQLiteLocation As String, ByVal sCustodianName As String, ByVal sFieldName As String, sFieldValue As String) As Boolean
        UpdateExtractionDBInfo = False
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

        UpdateExtractionDBInfo = True
    End Function

    Public Function UpdateSQLiteAllBatchInfo(ByVal sSqlDataBaseFullName As String, ByVal sCustodianName As String, ByRef iCustodianID As Integer, ByVal sGroupID As String, ByVal sStatus As String, ByVal sProcessID As String, ByVal sRedisSettings As String, ByVal sSourceType As String, ByVal sSourceFormat As String, ByVal sOutputType As String, ByVal sOutputFormat As String, ByVal sConversionStartTime As String, ByVal sConversionEndTime As String, ByVal iBytesProcessed As Integer, ByVal iPercentComplete As Integer, ByVal iSuccess As Integer, ByVal iFailed As Integer, ByVal sSummaryReportLocation As String) As Boolean
        UpdateSQLiteAllBatchInfo = False
        Dim sInsertArchiveExtractionData As String
        Dim sQueryReturnedCustodianID As String
        Dim SQLiteConnection As SQLiteConnection

        SQLiteConnection = New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;Read Only=True;New=False;Compress=True;")
        SQLiteConnection.Open()

        Using Connection As New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;")
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
        UpdateSQLiteAllBatchInfo = True
    End Function

    Public Function InsertDataConversionStats(ByVal sSqlDataBaseFullName As String, ByVal sCustodianName As String, ByRef iCustodianID As Integer, ByVal sGroupID As String, ByVal sStatus As String, ByVal sProcessID As String, ByVal sRedisSettings As String, ByVal sSourceType As String, ByVal sSourceFormat As String, ByVal sOutputType As String, ByVal sOutputFormat As String, ByVal sConversionStartTime As String, ByVal sConversionEndTime As String, ByVal iBytesProcessed As Integer, ByVal iPercentComplete As Integer, ByVal iSuccess As Integer, ByVal iFailed As Integer, ByVal sSummaryReportLocation As String) As Boolean
        InsertDataConversionStats = False
        Dim sInsertArchiveExtractionData As String
        Dim sQueryReturnedCustodianID As String
        Dim SQLiteConnection As SQLiteConnection

        SQLiteConnection = New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;Read Only=True;New=False;Compress=True;")
        SQLiteConnection.Open()

        Using Connection As New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;")
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
        InsertDataConversionStats = True
    End Function

    Public Function InsertSourceDataStats(ByVal sSqlDataBaseFullName As String, ByVal sSourceFileName As String, ByVal sSourceFilePath As String, ByVal sSourceFileCreateDate As String, ByVal sSourceFileModifiedDate As String, ByVal dblSourceFileSize As Double, ByVal iCustodianID As Integer) As Boolean
        InsertSourceDataStats = False
        Dim sInsertSourceDataStatsQuery As String
        Dim sQueryReturnedCustodianID As String
        Dim SQLiteConnection As SQLiteConnection

        SQLiteConnection = New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;Read Only=True;New=False;Compress=True;")
        SQLiteConnection.Open()

        Using Connection As New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;")
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
        InsertSourceDataStats = True
    End Function

    Public Function GetUpdatedIngestionDBInfo(ByRef sSQLiteLocation As String, ByRef sCustodianName As String, ByRef sFieldName As String, ByRef sFieldValue As String, ByVal sFieldType As String) As Boolean
        GetUpdatedIngestionDBInfo = False

        Dim Connection As SQLiteConnection
        Dim SQLCommand As SQLiteCommand
        Dim sCustodianQuery As String
        Dim dataReader As SQLiteDataReader
        Dim dFieldValue As Double

        Try
            Connection = New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
            Connection.Open()

            sCustodianQuery = "SELECT " & sFieldName & " FROM ewsIngestionStats WHERE CustodianName = '" & sCustodianName & "'"

            SQLCommand = New SQLiteCommand(sCustodianQuery, Connection)
            dataReader = SQLCommand.ExecuteReader
            If dataReader.HasRows Then
                While dataReader.Read
                    If sFieldType = "TEXT" Then
                        If Not IsDBNull(dataReader(sFieldName)) Then
                            sFieldValue = dataReader.GetValue(0)
                        Else
                            sFieldValue = vbNullString
                        End If
                    ElseIf sFieldType = "INT" Then

                        If Not IsDBNull(dataReader(sFieldName)) Then
                            'dFieldValue = dataReader.GetValue(0)
                            Try
                                dFieldValue = dataReader.GetInt64(0)
                                sFieldValue = dFieldValue.ToString
                            Catch ex1 As Exception
                                Try
                                    dFieldValue = dataReader.GetValue(0)
                                    sFieldValue = dFieldValue.ToString
                                Catch ex2 As Exception

                                End Try
                            End Try

                        Else
                            dFieldValue = 0
                            sFieldValue = dFieldValue.ToString
                        End If

                    End If
                End While
            End If
            Connection.Close()
            GetUpdatedIngestionDBInfo = True

        Catch ex As Exception
            MessageBox.Show("Error in DatabaseService.GetUpdatedIngestionDBInfo - DB NAME - " & sSQLiteLocation & " " & ex.ToString & "Field Name = " & sFieldName, "DatabaseService GetUpdatedIngestionDBInfo", MessageBoxButtons.OK)
        End Try
    End Function

    Public Function GetCustodianPSTLocation(ByVal sSqlDataBaseFullName As String, ByVal sCustodianName As String, ByRef sPSTLocation As String) As Boolean
        GetCustodianPSTLocation = False

        Dim Connection As SQLiteConnection
        Dim SQLCommand As SQLiteCommand
        Dim sCustodianQuery As String
        Dim dataReader As SQLiteDataReader

        Connection = New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;Read Only=True;New=False;Compress=True;")
        Connection.Open()

        sCustodianQuery = "SELECT PSTPath FROM ewsIngestionStats WHERE CustodianName = '" & sCustodianName & "'"

        SQLCommand = New SQLiteCommand(sCustodianQuery, Connection)
        dataReader = SQLCommand.ExecuteReader
        If dataReader.HasRows Then
            While dataReader.Read
                sPSTLocation = dataReader.GetValue(0)
            End While

        End If
        Connection.Close()
        GetCustodianPSTLocation = True
    End Function
    Public Function GetCustodianRowID(ByVal sSqlDataBaseFullName As String, ByVal sCustodianName As String, ByRef iCustodianRowID As Integer) As Boolean

        GetCustodianRowID = False

        Dim sQueryReturnedCustodianID As String
        Dim SQLiteConnection As SQLiteConnection

        SQLiteConnection = New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;Read Only=True;New=False;Compress=True;")
        SQLiteConnection.Open()

        Using Connection As New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;")
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

        GetCustodianRowID = True
    End Function

    Public Function UpdateCustodianDBStartTime(ByVal sSqlDataBaseFullName As String, ByVal sCustodianName As String, ByVal sStartTime As String) As Boolean
        UpdateCustodianDBStartTime = False
        Dim iReturnValue As Integer
        Dim Connection As SQLiteConnection
        Dim sUpdateEWSIngestionData As String
        Dim oUpdateEWSIngestionDataCommand As SQLiteCommand

        Connection = New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;New=False;Compress=True;")
        'Connection.ConnectionString = "Data Source=c:\mydatabasefile.db3;Version=3;New=False;Compress=True;"
        Connection.Open()

        sUpdateEWSIngestionData = "Update DataConversionStats set ConversionStartTime = @ConversionStartTime"
        sUpdateEWSIngestionData = sUpdateEWSIngestionData & " WHERE CustodianName = @CustodianName"

        oUpdateEWSIngestionDataCommand = New SQLiteCommand(sUpdateEWSIngestionData)
        oUpdateEWSIngestionDataCommand.Parameters.AddWithValue("@CustodianName", sCustodianName)
        oUpdateEWSIngestionDataCommand.Parameters.AddWithValue("@ConversionStartTime", sStartTime)
        oUpdateEWSIngestionDataCommand.Connection = Connection
        iReturnValue = oUpdateEWSIngestionDataCommand.ExecuteNonQuery()

        If Not IsNothing(Connection) Then
            Connection.Close()
        End If

        UpdateCustodianDBStartTime = True
    End Function


    Public Function GetUpdatedDataConversionCustodianInfo(ByVal sSqlDataBaseFullName As String, ByRef lstCustodianName As List(Of String), ByRef lstBytesUploaded As List(Of String), ByRef lstPercentCompleted As List(Of String)) As Boolean
        GetUpdatedDataConversionCustodianInfo = False

        Dim Connection As SQLiteConnection
        Dim SQLCommand As SQLiteCommand
        Dim sCustodianQuery As String
        Dim dataReader As SQLiteDataReader

        Connection = New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;Read Only=True;New=False;Compress=True;")
        Connection.Open()

        sCustodianQuery = "SELECT CustodianName, BytesProcessed, PercentCompleted FROM DataConversionStats WHERE ConverstionStatus = 'Conversion In Progress'"

        SQLCommand = New SQLiteCommand(sCustodianQuery, Connection)
        dataReader = SQLCommand.ExecuteReader
        If dataReader.HasRows Then
            While dataReader.Read
                lstCustodianName.Add(dataReader.GetValue(0))
                lstBytesUploaded.Add(dataReader.GetInt64(1).ToString)
                If dataReader.IsDBNull(2) Then
                    lstPercentCompleted.Add(0)
                Else
                    lstPercentCompleted.Add(dataReader.GetValue(2))
                End If
            End While
        End If
        Connection.Close()
        GetUpdatedDataConversionCustodianInfo = True
    End Function

    Public Function GetUpdatedEWSIngestionCustodianInfo(ByVal sSqlDataBaseFullName As String, ByRef lstCustodianName As List(Of String), ByRef lstBytesUploaded As List(Of String), ByRef lstPercentCompleted As List(Of String), ByRef bUpdateReturn As Boolean) As Boolean
        GetUpdatedEWSIngestionCustodianInfo = False

        Dim Connection As SQLiteConnection
        Dim SQLCommand As SQLiteCommand
        Dim sCustodianQuery As String
        Dim dataReader As SQLiteDataReader

        Connection = New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;Read Only=True;New=False;")
        Connection.Open()

        sCustodianQuery = "SELECT CustodianName, BytesUploaded, PercentCompleted FROM ewsIngestionStats WHERE ProgressStatus = 'In Progress'"

        SQLCommand = New SQLiteCommand(sCustodianQuery, Connection)
        dataReader = SQLCommand.ExecuteReader

        While dataReader.Read
            lstCustodianName.Add(dataReader.GetValue(0))
            lstBytesUploaded.Add(dataReader.GetInt64(1).ToString)
            If dataReader.IsDBNull(2) Then
                lstPercentCompleted.Add(0)
            Else
                lstPercentCompleted.Add(dataReader.GetValue(2))
            End If
        End While
        If lstCustodianName.Count > 0 Then
            bUpdateReturn = True
        Else
            bUpdateReturn = False
        End If
        Connection.Close()
        GetUpdatedEWSIngestionCustodianInfo = True
    End Function
    Public Function UpdateArchiveExtractCustodianInfo(ByVal sSQLiteLocation As String, ByVal sBatchName As String, ByVal sFieldName As String, sFieldValue As String) As Boolean
        UpdateArchiveExtractCustodianInfo = False
        Dim sUpdateArchiveExtractionData As String

        Using Connection As New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;New=False;Compress=True;")
            sUpdateArchiveExtractionData = "Update archiveExtractionStats set " & sFieldName & " = " & "@" & sFieldName
            sUpdateArchiveExtractionData = sUpdateArchiveExtractionData & " WHERE BatchName = @BatchName"

            Using oUpdateArchiveExtractionDataCommand As New SQLiteCommand()
                With oUpdateArchiveExtractionDataCommand
                    .Parameters.AddWithValue("@BatchName", sBatchName)
                    .Parameters.AddWithValue("@" & sFieldName, sFieldValue)
                    .CommandText = sUpdateArchiveExtractionData
                    .Connection = Connection
                End With
                Try
                    Connection.Open()
                    oUpdateArchiveExtractionDataCommand.ExecuteNonQuery()
                    Connection.Close()
                Catch ex As Exception
                    MessageBox.Show("Error in UpdateCustodianStatus" & ex.ToString, " UpdateArchiveExtractCustodianInfo Error", MessageBoxButtons.OK)
                End Try
            End Using
        End Using

        UpdateArchiveExtractCustodianInfo = True
    End Function

    Public Function UpdateCustodianIngestionValues(ByVal sSQLiteLocation As String, ByVal sCustodianName As String, ByVal sFieldName As String, sFieldValue As String) As Boolean
        UpdateCustodianIngestionValues = False
        Dim sUpdateArchiveExtractionData As String

        Using Connection As New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;New=False;Compress=True;")
            sUpdateArchiveExtractionData = "Update ewsIngestionStats set " & sFieldName & " = " & "@" & sFieldName
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
                    MessageBox.Show("Error in UpdateCustodianStatus" & ex.ToString, " UpdateCustodianIngestionValues Error", MessageBoxButtons.OK)
                End Try
            End Using
        End Using

        UpdateCustodianIngestionValues = True
    End Function

    Public Function GetUpdatedProcessingDetails(ByRef sSQLiteLocation As String, ByRef sBatchName As String, ByRef dblItemsProcessed As Double) As Boolean
        GetUpdatedProcessingDetails = False

        Dim Connection As SQLiteConnection
        Dim SQLCommand As SQLiteCommand
        Dim sCustodianQuery As String
        Dim dataReader As SQLiteDataReader
        Try

            Connection = New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
            Connection.Open()

            sCustodianQuery = "SELECT BatchName, ItemsProcessed FROM archiveExtractionStats WHERE ExtractionStatus = 'Extraction In Progress'"

            SQLCommand = New SQLiteCommand(sCustodianQuery, Connection)
            dataReader = SQLCommand.ExecuteReader
            If dataReader.HasRows Then
                While dataReader.Read
                    dblItemsProcessed = dataReader.GetInt64(1).ToString
                End While
            End If
            Connection.Close()

        Catch ex As Exception
            MessageBox.Show("Error in: GetUpdatedProcessingDetails - " & ex.ToString)

        End Try
        GetUpdatedProcessingDetails = True
    End Function

    Public Function GetUpdatedArchiveExtractInfo(ByRef sSQLiteLocation As String, ByRef sBatchName As String, ByRef sFieldName As String, ByRef sFieldValue As String, ByVal sFieldType As String) As Boolean
        GetUpdatedArchiveExtractInfo = False

        Dim Connection As SQLiteConnection
        Dim SQLCommand As SQLiteCommand
        Dim sCustodianQuery As String
        Dim dataReader As SQLiteDataReader
        Dim dFieldValue As Double

        Try
            Connection = New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
            Connection.Open()

            sCustodianQuery = "SELECT " & sFieldName & " FROM archiveExtractionStats WHERE BatchName = '" & sBatchName & "'"

            SQLCommand = New SQLiteCommand(sCustodianQuery, Connection)
            dataReader = SQLCommand.ExecuteReader
            If dataReader.HasRows Then
                While dataReader.Read
                    If sFieldType = "TEXT" Then
                        If Not IsDBNull(dataReader(sFieldName)) Then
                            sFieldValue = dataReader.GetValue(0)
                        Else
                            sFieldValue = vbNullString
                        End If
                    ElseIf sFieldType = "INT" Then

                        If Not IsDBNull(dataReader(sFieldName)) Then
                            'dFieldValue = dataReader.GetValue(0)
                            Try
                                dFieldValue = dataReader.GetInt64(0)
                                sFieldValue = dFieldValue.ToString
                            Catch ex1 As Exception
                                Try
                                    dFieldValue = dataReader.GetValue(0)
                                    sFieldValue = dFieldValue.ToString
                                Catch ex2 As Exception

                                End Try
                            End Try

                        Else
                            dFieldValue = 0
                            sFieldValue = dFieldValue.ToString
                        End If

                    End If
                End While
            End If
            Connection.Close()
            GetUpdatedArchiveExtractInfo = True

        Catch ex As Exception
            MessageBox.Show("Error in GetUpdatedArchiveExtractInfo - " & ex.ToString & "Field Name = " & sFieldName, " GetUpdatedArchiveExtractInfo", MessageBoxButtons.OK)
        End Try
    End Function

    Public Function GetUpdatedEWSExtractInfo(ByVal sSqlDataBaseFullName As String, ByRef sCustodianName As String, ByRef sFieldName As String, ByRef sFieldValue As String) As Boolean
        GetUpdatedEWSExtractInfo = False

        Dim Connection As SQLiteConnection
        Dim SQLCommand As SQLiteCommand
        Dim sCustodianQuery As String
        Dim dataReader As SQLiteDataReader

        Connection = New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;Read Only=True;New=False;Compress=True;")
        Connection.Open()

        sCustodianQuery = "SELECT " & sFieldName & " FROM ewsExtractionStats WHERE CustodianName = '" & sCustodianName & "'"

        SQLCommand = New SQLiteCommand(sCustodianQuery, Connection)
        dataReader = SQLCommand.ExecuteReader
        If dataReader.HasRows Then
            While dataReader.Read
                sFieldValue = dataReader.GetValue(0)
            End While
        End If
        Connection.Close()
        GetUpdatedEWSExtractInfo = True
    End Function

    Public Function CheckSQLConnection(ByVal sAuthType As String, ByVal sHostName As String, ByVal sPort As String, ByVal sDBName As String, ByVal sSQLUserName As String, ByVal sSQLPW As String, ByVal sDomain As String) As Boolean
        Dim Connection As SqlConnection
        Dim sConnection As String
        Dim common As New Common

        If sAuthType = "Windows Authentication" Then
            sConnection = "Data Source=" & sHostName & "," & sPort & ";Database=" & sDBName & ";Integrated Security=true;"
            Connection = New SqlConnection(sConnection)

        ElseIf sAuthType = "SQLServer Authentication" Then
            sConnection = "Data Source=" & sHostName & "," & sPort & ";Database=" & sDBName & ";User ID=" & sSQLUserName & ";Password=" & sSQLPW & ";"
            Connection = New SqlConnection("Data Source=" & sHostName & "," & sPort & ";Database=" & sDBName & ";User ID=" & sSQLUserName & ";Password=" & sSQLPW & ";")
        End If

        Try
            Connection.Open()

            Connection.Close()
            CheckSQLConnection = True
        Catch ex As Exception
            common.Logger(psIngestionLogFile, ex.ToString)
            CheckSQLConnection = False
        End Try

    End Function

    Public Function UpdateArchiveExtractBatchInfo(ByVal sSQLiteDatabaseFullName As String, ByVal sArchiveName As String, ByVal sBatchName As String, ByVal sExtractionStatus As String, ByVal sArchiveSettings As String, ByVal sArchiveType As String, ByVal sOutputFormat As String, sSQLSettings As String, ByVal sWSSSettings As String, ByVal sSourceInformation As String, ByVal sProcessStartTime As String, ByVal sProcessEndTime As String, ByVal dblItemsProcessed As Double, ByVal dblItemsExported As Double, ByVal dblItemsSkipped As Double, ByVal dblItemsFailed As Double, ByVal sProcessingFilesDirectory As String, ByVal sCaseDirectory As String, ByVal sOutputDirectory As String, sLogDirectory As String, ByVal sSummaryReportLocation As String, ByVal sTotalBytes As String) As Boolean
        UpdateArchiveExtractBatchInfo = False
        Dim sInsertArchiveExtractionData As String
        Dim sQueryReturnedBatches As String
        Dim SQLiteConnection As SQLiteConnection
        Dim common As New Common

        SQLiteConnection = New SQLiteConnection("Data Source=" & sSQLiteDatabaseFullName & ";Version=3;Read Only=True;New=False;Compress=True;")
        SQLiteConnection.Open()

        Using Connection As New SQLiteConnection("Data Source=" & sSQLiteDatabaseFullName & ";Version=3;")
            Using SQLSelectCommand As New SQLiteCommand("SELECT BatchName FROM ArchiveExtractionStats WHERE BatchName='" & sBatchName & "'")
                With SQLSelectCommand
                    .Connection = SQLiteConnection
                    Using readerObject As SQLiteDataReader = SQLSelectCommand.ExecuteReader
                        While readerObject.Read
                            sQueryReturnedBatches = readerObject("BatchName").ToString
                        End While
                    End Using
                End With

            End Using

            If sQueryReturnedBatches = vbNullString Then
                sInsertArchiveExtractionData = "Insert into ArchiveExtractionStats (ArchiveName, BatchName, ExtractionStatus, ArchiveSettings, ArchiveType, OutputFormat, SQLSettings, SourceInformation, WSSSettings, ProcessStartTime, ProcessEndTime, ItemsProcessed, ItemsExported, ItemsSkipped, ItemsFailed, ProcessingFilesDirectory, CaseDirectory, OutputDirectory, LogDirectory, SummaryReportLocation, TotalBytes) Values "
                sInsertArchiveExtractionData = sInsertArchiveExtractionData & "(@ArchiveName, @BatchName, @ExtractionStatus, @ArchiveSettings, @ArchiveType, @OutputFormat, @SQLSettings, @SourceInformation, @WSSSettings, @ProcessStartTime, @ProcessEndTime, @ItemsProcessed, @ItemsExported, @ItemsSkipped, @ItemsFailed, @ProcessingFilesDirectory, @CaseDirectory, @OutputDirectory, @LogDirectory, @SummaryReportLocation, @TotalBytes)"
                Using oInsertArchiveExtractionDataCommand As New SQLiteCommand()
                    With oInsertArchiveExtractionDataCommand
                        .Connection = Connection
                        .CommandText = sInsertArchiveExtractionData
                        .Parameters.AddWithValue("@ArchiveName", sArchiveName)
                        .Parameters.AddWithValue("@BatchName", sBatchName)
                        .Parameters.AddWithValue("@ExtractionStatus", sExtractionStatus)
                        .Parameters.AddWithValue("@ArchiveSettings", sArchiveSettings)
                        .Parameters.AddWithValue("@ArchiveType", sArchiveType)
                        .Parameters.AddWithValue("@OutputFormat", sOutputFormat)
                        .Parameters.AddWithValue("@SQLSettings", sSQLSettings)
                        .Parameters.AddWithValue("@SourceInformation", sSourceInformation)
                        .Parameters.AddWithValue("@WSSSettings", sWSSSettings)
                        .Parameters.AddWithValue("@ProcessStartTime", sProcessStartTime)
                        .Parameters.AddWithValue("@ProcessEndTime", sProcessEndTime)
                        .Parameters.AddWithValue("@ItemsProcessed", dblItemsProcessed)
                        .Parameters.AddWithValue("@ItemsExported", dblItemsExported)
                        .Parameters.AddWithValue("@ItemsSkipped", dblItemsSkipped)
                        .Parameters.AddWithValue("@ItemsFailed", dblItemsFailed)
                        .Parameters.AddWithValue("@ProcessingFilesDirectory", sProcessingFilesDirectory)
                        .Parameters.AddWithValue("@CaseDirectory", sCaseDirectory)
                        .Parameters.AddWithValue("@OutputDirectory", sOutputDirectory)
                        .Parameters.AddWithValue("@LogDirectory", sLogDirectory)
                        .Parameters.AddWithValue("@SummaryReportLocation", sSummaryReportLocation)
                        .Parameters.AddWithValue("@TotalBytes", sTotalBytes)
                    End With
                    Try
                        Connection.Open()
                        oInsertArchiveExtractionDataCommand.ExecuteNonQuery()
                        Connection.Close()
                    Catch ex As Exception
                        common.Logger(psIngestionLogFile, "Error 5 - " & ex.Message.ToString())
                    End Try
                End Using
            Else

            End If
            SQLiteConnection.Close()
        End Using
        UpdateArchiveExtractBatchInfo = True
    End Function

    Public Function UpdateDataConversionStatBatchInfo(ByVal sSQLLiteLocation As String, ByVal sCustodianName As String, ByRef iCustodianID As Integer, ByVal sGroupID As String, ByVal sStatus As String, ByVal sProcessID As String, ByVal sRedisSettings As String, ByVal sSourceType As String, ByVal sSourceFormat As String, ByVal sOutputType As String, ByVal sOutputFormat As String, ByVal sConversionStartTime As String, ByVal sConversionEndTime As String, ByVal iBytesProcessed As Integer, ByVal iPercentComplete As Integer, ByVal iSuccess As Integer, ByVal iFailed As Integer, ByVal sSummaryReportLocation As String) As Boolean
        UpdateDataConversionStatBatchInfo = False
        Dim sInsertArchiveExtractionData As String
        Dim sQueryReturnedCustodianID As String
        Dim SQLiteConnection As SQLiteConnection

        SQLiteConnection = New SQLiteConnection("Data Source=" & sSQLLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;Read Only=True;New=False;Compress=True;")
        SQLiteConnection.Open()

        Using Connection As New SQLiteConnection("Data Source=" & sSQLLiteLocation & "\NuixEmailArchiveMigrationManager.db3;Version=3;")
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
        UpdateDataConversionStatBatchInfo = True
    End Function

    Public Function UpdateEWSIngestionAllCustodiansInfo(ByVal sSqlDataBaseFullName As String, ByVal sCustodianName As String, ByVal sProgressStatus As String, ByVal sPSTPath As String, ByVal iNumberofPST As Integer, ByVal dblPSTSize As Double, ByVal sGroupID As String, ByVal sDestinationFolder As String, ByVal sDestinationRoot As String, ByVal sDestinationSMTP As String, ByVal bUploadNotStart As Boolean, ByVal bUploadInProgress As Boolean, ByVal bUploadCompleted As Boolean, ByVal sProcessID As String, ByVal dIngestionStartTime As DateTime, ByVal dIngestionEndTime As DateTime, ByVal dblBytesUploaded As Double, ByVal iPercentCompleted As Integer, ByVal dblSuccessItems As Double, ByVal dblFailedItems As Double, ByVal sProcessingFilesDir As String, ByVal sCaseDirectory As String, ByVal sOutputDirectory As String, ByVal sLogDirectory As String, ByVal sSummaryReport As String) As Boolean
        UpdateEWSIngestionAllCustodiansInfo = False
        Dim sInsertEWSIngestionData As String
        Dim sUpdateEWSIngestionData As String
        Dim sQueryReturnedCustodian As String
        Dim SQLiteConnection As SQLiteConnection
        Dim common As New Common

        SQLiteConnection = New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;Read Only=False;New=False;Compress=True;")
        SQLiteConnection.Open()


        Using Connection As New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;")
            Using SQLSelectCommand As New SQLiteCommand("SELECT CustodianName FROM ewsIngestionStats WHERE CustodianName='" & sCustodianName & "'")
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
                sInsertEWSIngestionData = "Insert into ewsIngestionStats (CustodianName, ProgressStatus, PSTPath, NumberOfPSTs, TotalSizeOfPSTs, GroupID, DestinationFolder, DestinationRoot, DestinationSMTP, ProcessID, IngestionStartTime, IngestionEndTime, BytesUploaded, PercentCompleted, Success, Failed, ProcessingFilesDirectory, CaseDirectory, OutputDirectory, LogDirectory, SummaryReportLocation) Values "
                sInsertEWSIngestionData = sInsertEWSIngestionData & "(@CustodianName, @ProgressStatus, @PSTPath, @NumberOfPSTs, @TotalSizeOfPSTs, @GroupID, @DestinationFolder, @DestinationRoot, @DestinationSMTP, @ProcessID, @IngestionStartTime, @IngestionEndTime, @BytesUploaded, @PercentCompleted, @Success, @Failed, @ProcessingFilesDirectory, @CaseDirectory, @OutputDirectory, @LogDirectory, @SummaryReportLocation)"
                Using oInsertEWSIngestionDataCommand As New SQLiteCommand()
                    With oInsertEWSIngestionDataCommand
                        .Connection = Connection
                        .CommandText = sInsertEWSIngestionData
                        .Parameters.AddWithValue("@CustodianName", sCustodianName)
                        .Parameters.AddWithValue("@ProgressStatus", sProgressStatus)
                        .Parameters.AddWithValue("@PSTPath", sPSTPath)
                        .Parameters.AddWithValue("@NumberOfPSTs", iNumberofPST)
                        .Parameters.AddWithValue("@TotalSizeOfPSTs", dblPSTSize)
                        .Parameters.AddWithValue("@GroupID", sGroupID)
                        .Parameters.AddWithValue("@DestinationFolder", sDestinationFolder)
                        .Parameters.AddWithValue("@DestinationRoot", sDestinationRoot)
                        .Parameters.AddWithValue("@DestinationSMTP", sDestinationSMTP)
                        .Parameters.AddWithValue("@ProcessID", sProcessID)
                        .Parameters.AddWithValue("@IngestionStartTime", dIngestionStartTime.ToString)
                        .Parameters.AddWithValue("@IngestionEndTime", dIngestionEndTime.ToString)
                        .Parameters.AddWithValue("@BytesUploaded", dblBytesUploaded)
                        .Parameters.AddWithValue("@PercentCompleted", iPercentCompleted)
                        .Parameters.AddWithValue("@Success", dblSuccessItems)
                        .Parameters.AddWithValue("@Failed", dblFailedItems)
                        .Parameters.AddWithValue("@ProcessingFilesDirectory", sProcessingFilesDir)
                        .Parameters.AddWithValue("@CaseDirectory", sCaseDirectory)
                        .Parameters.AddWithValue("@OutputDirectory", sOutputDirectory)
                        .Parameters.AddWithValue("@LogDirectory", sLogDirectory)
                        .Parameters.AddWithValue("@SummaryReportLocation", sSummaryReport)
                    End With
                    Try
                        Connection.Open()
                        oInsertEWSIngestionDataCommand.ExecuteNonQuery()
                        Connection.Close()
                    Catch ex As Exception
                        common.Logger(psIngestionLogFile, "Error 5 - " & ex.Message.ToString())
                    End Try
                End Using
            Else
                sUpdateEWSIngestionData = "Update ewsIngestionStats set ProgressStatus = @ProgressStatus, PSTPath = @PSTPath, ProcessID = @ProcessID, GroupID = @GroupID, DestinationFolder = @DestinationFolder, DestinationRoot = @DestinationRoot, DestinationSMTP = @DestinationSMTP, IngestionStartTime = @IngestionStartTime, IngestionEndTime = @IngestionEndTime, "
                sUpdateEWSIngestionData = sUpdateEWSIngestionData & "BytesUploaded = @BytesUploaded, PercentCompleted = @PercentCompleted, Success = @Success, Failed = @Failed, ProcessingFilesDirectory = @ProcessingFilesDirectory, CaseDirectory = @CaseDirectory, OutputDirectory = @OutputDirectory, LogDirectory = @LogDirectory, SummaryReportLocation = @SummaryReportLocation WHERE CustodianName = @CustodianName"

                Using oUpdateEWSIngestionDataCommand As New SQLiteCommand()
                    With oUpdateEWSIngestionDataCommand
                        .Connection = Connection
                        .CommandText = sUpdateEWSIngestionData
                        .Parameters.AddWithValue("@CustodianName", sCustodianName)
                        .Parameters.AddWithValue("@ProgressStatus", sProgressStatus)
                        .Parameters.AddWithValue("@PSTPath", sPSTPath)
                        .Parameters.AddWithValue("@GroupID", sGroupID)
                        .Parameters.AddWithValue("@DestinationFolder", sDestinationFolder)
                        .Parameters.AddWithValue("@DestinationRoot", sDestinationRoot)
                        .Parameters.AddWithValue("@DestinationSMTP", sDestinationSMTP)
                        .Parameters.AddWithValue("@ProcessID", sProcessID)
                        .Parameters.AddWithValue("@IngestionStartTime", dIngestionStartTime.ToString)
                        .Parameters.AddWithValue("@IngestionEndTime", dIngestionEndTime.ToString)
                        .Parameters.AddWithValue("@BytesUploaded", dblBytesUploaded)
                        .Parameters.AddWithValue("@PercentCompleted", iPercentCompleted)
                        .Parameters.AddWithValue("@Success", dblSuccessItems)
                        .Parameters.AddWithValue("@Failed", dblFailedItems)
                        .Parameters.AddWithValue("@ProcessingFilesDirectory", sProcessingFilesDir)
                        .Parameters.AddWithValue("@CaseDirectory", sCaseDirectory)
                        .Parameters.AddWithValue("@OutputDirectory", sOutputDirectory)
                        .Parameters.AddWithValue("@LogDirectory", sLogDirectory)
                        .Parameters.AddWithValue("@SummaryReportLocation", sSummaryReport)
                    End With
                    Try
                        Connection.Open()
                        oUpdateEWSIngestionDataCommand.ExecuteNonQuery()
                        Connection.Close()
                    Catch ex As Exception
                        common.Logger(psIngestionLogFile, "Error 6 - " & ex.Message.ToString())
                    End Try
                End Using
            End If
            Connection.Close()
        End Using
        UpdateEWSIngestionAllCustodiansInfo = True
    End Function

    Public Function UpdateSQLiteEWSDB(ByVal sSqlDataBaseFullName As String, ByVal sCustodianName As String, ByVal sProgressStatus As String, ByVal sPSTPath As String, ByVal iNumberOfPSTs As Integer, ByVal dblTotalPSTSize As Double, ByVal sGroupID As String, ByVal sDestinationFolder As String, ByVal sDestinationRoot As String, ByVal sDestinationSMTP As String, ByVal sProcessID As String, ByVal sIngestionStartTime As String, ByVal sIngestionEndTime As String, ByVal dblBytesUploaded As Double, ByVal dblSuccess As Double, ByVal dblFailed As Double, ByVal sProcessingFilesDirectory As String, ByVal sCaseDirectory As String, ByVal sOutputDirectory As String, ByVal sLogDirectory As String, ByVal sSummaryReportLocation As String) As Boolean
        UpdateSQLiteEWSDB = False
        Dim SQLConnection As SQLiteConnection
        Dim sInsertEWSIngestionData As String
        Dim sUpdateEWSIngestionData As String
        Dim iPercentCompleted As Integer
        Dim sCustodainDBName As String

        Dim common As New Common

        If Not sCustodianName = vbNullString Then
            SQLConnection = New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;Read Only=True;New=False;Compress=True;")
            SQLConnection.Open()

            Using SQLSelectCommand As New SQLiteCommand("SELECT CustodianName FROM ewsIngestionStats WHERE CustodianName='" & sCustodianName & "'")
                With SQLSelectCommand
                    .Connection = SQLConnection
                End With
                Using readerObject As SQLiteDataReader = SQLSelectCommand.ExecuteReader
                    While readerObject.Read
                        sCustodainDBName = readerObject("CustodianName").ToString
                    End While
                End Using
            End Using

            SQLConnection.Close()

            If sCustodainDBName = vbNullString Then
                Using Connection As New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;New=False;Compress=True;")
                    sInsertEWSIngestionData = "Insert into ewsIngestionStats (CustodianName, ProgressStatus, PSTPath, NumberOfPSTs, TotalSizeOfPSTs, GroupID, DestinationFolder, DestinationRoot, DestinationSMTP, ProcessID, IngestionStartTime, IngestionEndTime, BytesUploaded, Success, Failed, ProcessingFilesDirectory, CaseDirectory, OutputDirectory, LogDirectory, SummaryReportLocation) Values "
                    sInsertEWSIngestionData = sInsertEWSIngestionData & "(@CustodianName, @ProgressStatus, @PSTPath, @NumberOfPSTs, @TotalSizeOfPSTs, @GroupID, @DestinationFolder, @DestinationRoot, @DestinationSMTP, @ProcessID, @IngestionStartTime, @IngestionEndTime, @BytesUploaded, @Success, @Failed, @ProcessingFilesDirectory, @CaseDirectory, @OutputDirectory, @LogDirectory, @SummaryReportLocation)"
                    Using oInsertEWSIngestionDataCommand As New SQLiteCommand()
                        With oInsertEWSIngestionDataCommand
                            .Connection = Connection
                            .CommandText = sInsertEWSIngestionData
                            .Parameters.AddWithValue("@CustodianName", sCustodianName)
                            .Parameters.AddWithValue("@ProgressStatus", sProgressStatus)
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
                Using Connection As New SQLiteConnection("Data Source=" & sSqlDataBaseFullName & ";Version=3;New=False;Compress=True;")
                    sUpdateEWSIngestionData = "Update ewsIngestionStats set PSTPath = @PSTPath,ProgressStatus = @ProgressStatus,  ProcessID = @ProcessID, IngestionStartTime = @IngestionStartTime, IngestionEndTime = @IngestionEndTime, "
                    sUpdateEWSIngestionData = sUpdateEWSIngestionData & "BytesUploaded = @BytesUploaded, DestinationFolder = @DestinationFolder, DestinationRoot = @DestinationRoot, DestinationSMTP = @DestinationSMTP, PercentCompleted = @PercentCompleted, Success = @Success, Failed = @Failed, ProcessingFilesDirectory = @ProcessingFilesDirectory, CaseDirectory = @CaseDirectory, OutputDirectory = @OutputDirectory, LogDirectory = @LogDirectory, SummaryReportLocation = @SummaryReportLocation WHERE CustodianName = @CustodianName"

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
                            .Parameters.AddWithValue("@DestinationFolder", sDestinationFolder)
                            .Parameters.AddWithValue("@DestinationRoot", sDestinationRoot)
                            .Parameters.AddWithValue("@DestinationSMTP", sDestinationSMTP)
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

        UpdateSQLiteEWSDB = True
    End Function
End Class
