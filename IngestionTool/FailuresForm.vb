Imports System.IO
Public Class FailuresForm


    Private psSummaryReportFiles As String

    Public Property CustodianSummaryReportFile As String
        Get
            Return psSummaryReportFiles
        End Get
        Set(value As String)
            psSummaryReportFiles = value
        End Set
    End Property
    Private Sub FailuresForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim sSummaryReportFile As String
        Dim iCounter As Integer
        Dim iErrorCounter As Integer
        Dim dblSuccessfulItems As Double
        Dim asSummaryReportFiles() As String
        Dim sCustodianName As String

        Dim RootNode As TreeNode
        Dim custodianNode As TreeNode

        Dim value As System.Version = My.Application.Info.Version

        Me.Text = "Processing Details - " & value.ToString

        asSummaryReportFiles = Split(CustodianSummaryReportFile, ",")

        Dim lstUniqueSummaryErrors As List(Of String)
        Dim lstUniqueSummaryErrorsCount As List(Of Double)

        Dim bStatus As Boolean

        lstUniqueSummaryErrors = New List(Of String)
        lstUniqueSummaryErrorsCount = New List(Of Double)

        RootNode = New TreeNode
        RootNode.Name = "custodiansNode"
        RootNode.Text = "Custodian Information"
        Me.treeFailuresView.Nodes.Add(RootNode)

        For iCounter = 0 To asSummaryReportFiles.Length - 1
            sSummaryReportFile = asSummaryReportFiles(iCounter)
            bStatus = blnProcessSummaryReportFiles(sSummaryReportFile, sCustodianName, dblSuccessfulItems, lstUniqueSummaryErrors, lstUniqueSummaryErrorsCount)

            custodianNode = New TreeNode
            custodianNode.Name = "custodianNode"
            custodianNode.Text = sCustodianName
            RootNode.Nodes.Add(custodianNode)
            custodianNode.Nodes.Add("Successfuly Exported : " & dblSuccessfulItems)
            iErrorCounter = 0
            For Each SummaryError In lstUniqueSummaryErrors
                custodianNode.Nodes.Add(lstUniqueSummaryErrors(iErrorCounter) & " : " & lstUniqueSummaryErrorsCount(iErrorCounter))
                iErrorCounter = iErrorCounter + 1
            Next
            lstUniqueSummaryErrors.Clear()
            lstUniqueSummaryErrorsCount.Clear()
            sCustodianName = vbNullString

        Next iCounter
        RootNode.Expand()

    End Sub

    Public Function blnProcessSummaryReportFiles(ByVal sSummaryReportFile As String, ByRef sCustodianName As String, ByRef dblSuccessfulItems As Double, ByRef lstUniqueSummaryErrors As List(Of String), ByRef lstUniqueSummaryErrorsCount As List(Of Double)) As Boolean
        Dim sCurrentRow As String
        Dim asErrorArray() As String
        Dim asSuccessItems() As String
        Dim asCustodianInfo() As String
        Dim fileSummaryFile As StreamReader
        Dim iCounter As Integer
        Dim bFailedItemStart As Boolean

        Dim iErrorIndex As Integer
        Dim iArrayCount As Integer
        Dim dblCurrentCount As Integer
        Dim sCurrentErrorDetail As String

        blnProcessSummaryReportFiles = True

        Try
            iCounter = iCounter + 1
            bFailedItemStart = False
            If File.Exists(sSummaryReportFile) Then
                fileSummaryFile = New StreamReader(sSummaryReportFile)
                While Not fileSummaryFile.EndOfStream

                    sCurrentRow = fileSummaryFile.ReadLine
                    If (sCurrentRow.Contains("Failed Item Statistics")) Then
                        bFailedItemStart = True
                    ElseIf (sCurrentRow.Contains("File Type Statistics")) Then
                        bFailedItemStart = False
                    ElseIf (sCurrentRow.Contains("Total items exported:")) Then
                        asSuccessItems = Split(sCurrentRow, ":")
                        dblSuccessfulItems = CDbl(asSuccessItems(1))
                    ElseIf (sCurrentRow.Contains("   SOURCE: ")) Then
                        asCustodianInfo = Split(sCurrentRow, ":")
                        sCustodianName = asCustodianInfo(1)
                    ElseIf (sCurrentRow.Contains("Case: ")) Then
                        asCustodianInfo = Split(sCurrentRow, ":")
                        sCustodianName = asCustodianInfo(1)
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
            End If

            blnProcessSummaryReportFiles = True
        Catch ex As Exception
            MessageBox.Show("Counter = " & iCounter & ": Error in Process Summary Report Files " & ex.ToString)
            blnProcessSummaryReportFiles = False

        End Try

    End Function

    Private Sub treeFailuresView_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles treeFailuresView.AfterCheck

        subCheckAllSubNodes(e.Node, e.Node.Checked)
    End Sub

    Public Sub subCheckAllSubNodes(ByVal custodianNode As TreeNode, ByVal bCheckedMode As Boolean)
        For Each node In custodianNode.Nodes
            node.checked = bCheckedMode
            If node.Nodes.count > 0 Then
                subCheckAllSubNodes(node, bCheckedMode)
            End If

        Next
    End Sub

    Private Sub btnExportExceptionDetails_Click(sender As Object, e As EventArgs) Handles btnExportExceptionDetails.Click
        Dim bStatus As Boolean
        Dim lstCustodianIngestionDetails As List(Of String)

        lstCustodianIngestionDetails = New List(Of String)
        GetAllSelectedTreeNodes(treeFailuresView, lstCustodianIngestionDetails)

        bStatus = blnCustodianIngestionOutput(lstCustodianIngestionDetails)

    End Sub

    Private Sub GetNodeValues(ByVal n As TreeNode, ByRef lstCustodianIngestionDetails As List(Of String))


        Dim aNode As TreeNode
        For Each aNode In n.Nodes
            If aNode.Checked = True Then
                If aNode.Parent.Text <> "Custodian Information" Then
                    lstCustodianIngestionDetails.Add(aNode.Parent.Text & "," & aNode.Text)
                End If
            End If
            GetNodeValues(aNode, lstCustodianIngestionDetails)
        Next
    End Sub

    ' Call the procedure using the top nodes of the treeview.
    Private Sub GetAllSelectedTreeNodes(ByVal failuresTreeView As TreeView, ByRef lstCustodianIngestionDetails As List(Of String))

        For Each n In failuresTreeView.Nodes
            If n.checked = True Then

            End If
            GetNodeValues(n, lstCustodianIngestionDetails)
        Next
    End Sub

    Private Function blnCustodianIngestionOutput(ByVal lstCustodianIngestionDetails As List(Of String)) As Boolean

        Dim ReportOutputFile As StreamWriter
        Dim sMachineName As String
        Dim asIngestionDetails() As String
        Dim asCustodianDetails() As String
        Dim sCustodianName As String
        Dim sType As String
        Dim dblCount As Double
        Dim sOutputFileName As String
        Dim sOutputFileLocation As String


        sMachineName = System.Net.Dns.GetHostName()
        sOutputFileName = sMachineName & "-CustodianIngestion" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv"
        sOutputFileLocation = eMailArchiveMigrationManager.NuixLogDir
        ReportOutputFile = New StreamWriter(sOutputFileLocation & "\" & sOutputFileName)
        ReportOutputFile.WriteLine("Custodian Name, Type, Count")

        For Each IngestionDetail In lstCustodianIngestionDetails
            asCustodianDetails = Split(IngestionDetail.ToString, ",")

            sCustodianName = asCustodianDetails(0)
            asIngestionDetails = Split(asCustodianDetails(1), ":")
            sType = asIngestionDetails(0)
            dblCount = CDbl(asIngestionDetails(1))

            '            OutputStream.WriteLine(sCustodianName & "," & sFileName & "," & PSTSize(iCounter - 1) & "," & sFilePath & "\" & sFileName & ".pst" & "," & sLegalHold)
            ReportOutputFile.WriteLine(sCustodianName & "," & sType & "," & dblCount)

        Next

        ReportOutputFile.Close()
        MessageBox.Show("Custodian Ingestion details report exported - report can be found at " & sOutputFileLocation & "\" & sOutputFileName)
        blnCustodianIngestionOutput = True

    End Function

    Private Sub btnExportExceptionDetails_MouseHover(sender As Object, e As EventArgs) Handles btnExportExceptionDetails.MouseHover
        FailureDetailsToolTip.Show("Export Exception Details to CSV", btnExportExceptionDetails)
    End Sub
End Class