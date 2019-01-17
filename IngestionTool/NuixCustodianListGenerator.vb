Imports System.IO


Public Class NuixCustodianListGenerator
    Public psSettingsFile As String


    Private Sub btnCustodianListSelector_Click(sender As Object, e As EventArgs) Handles btnCustodianListSelector.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        With OpenFileDialog1
            .Filter = "*.csv|*.csv"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            txtCustomerCustodianList.Text = OpenFileDialog1.FileName.ToString
        End If
    End Sub

    Private Sub btnLoadCustomerCustodians_Click(sender As Object, e As EventArgs) Handles btnLoadCustomerCustodians.Click
        Dim sCustodianList As String
        Dim bStatus As Boolean

        sCustodianList = txtCustomerCustodianList.Text
        If (sCustodianList = vbNullString) Then
            MessageBox.Show("You must select a custodian list to import.", "Import Custodian List")
            txtCustomerCustodianList.Focus()
            Exit Sub
        Else
            bStatus = blnLoadCustomerCustodianList(sCustodianList, grdCustomerCustodianList, chkFirstRowHeader.Checked)

        End If
    End Sub

    Private Function blnLoadCustomerCustodianList(ByVal sCustodianList As String, ByVal grdCustodianListGrid As DataGridView, ByVal bFirstRowHeader As Boolean) As Boolean
        blnLoadCustomerCustodianList = False
        Dim sCurrentRow() As String
        Dim bFirstRow As Boolean
        Dim asFields() As String
        Dim bStatus As Boolean
        Dim asUpdatedRow() As String


        Dim fileCSVFile As New Microsoft.VisualBasic.FileIO.TextFieldParser(sCustodianList)
        fileCSVFile.TextFieldType = FileIO.FieldType.Delimited
        fileCSVFile.SetDelimiters(",")

        bFirstRow = True
        Do While Not fileCSVFile.EndOfData
            sCurrentRow = fileCSVFile.ReadFields
            If bFirstRow = True Then
                If bFirstRowHeader = True Then

                    bStatus = blnBuildHeaderRow(grdCustodianListGrid, sCurrentRow)
                    bStatus = blnPopulateFieldMappingGrid(grdCustodianFields, sCurrentRow)

                Else
                    ReDim Preserve asUpdatedRow(0)
                    asUpdatedRow(0) = "False"
                    For iCounter = 0 To UBound(sCurrentRow)
                        ReDim Preserve asFields(iCounter)
                        ReDim Preserve asFields(iCounter + 1)
                        asFields(iCounter) = "Column-" & iCounter
                        asUpdatedRow(iCounter + 1) = sCurrentRow(iCounter)
                    Next
                    bStatus = blnPopulateFieldMappingGrid(grdCustodianFields, asFields)
                    bStatus = blnBuildHeaderRow(grdCustodianListGrid, asFields)
                    grdCustodianListGrid.Rows.Add(asUpdatedRow)

                End If
                bFirstRow = False
            Else
                ReDim asUpdatedRow(0)
                asUpdatedRow(0) = "False"
                For iCounter = 0 To UBound(sCurrentRow)
                    ReDim Preserve asUpdatedRow(iCounter + 1)
                    asUpdatedRow(iCounter + 1) = sCurrentRow(iCounter)
                Next
                grdCustodianListGrid.Rows.Add(asUpdatedRow)
            End If

        Loop


        blnLoadCustomerCustodianList = True
    End Function

    Private Function blnBuildHeaderRow(ByVal grdCustodianListGrid As DataGridView, ByVal sHeaderRow() As String) As Boolean
        blnBuildHeaderRow = False
        Dim sHeaderValue As String
        Dim HeaderTextBox As DataGridViewTextBoxColumn
        Dim HeaderCell As DataGridViewColumnHeaderCell
        Dim HeaderCheckBox As DataGridViewCheckBoxColumn

        grdCustodianListGrid.Columns.Clear()

        HeaderCheckBox = New DataGridViewCheckBoxColumn
        HeaderCell = New DataGridViewColumnHeaderCell
        HeaderCheckBox.HeaderText = "Select Custodian"
        grdCustodianListGrid.Columns.Add(HeaderCheckBox)

        For iCounter = 0 To UBound(sHeaderRow) - 1
            sHeaderValue = sHeaderRow(iCounter)
            HeaderTextBox = New DataGridViewTextBoxColumn
            HeaderCell = New DataGridViewColumnHeaderCell
            HeaderTextBox.HeaderCell = HeaderCell
            HeaderTextBox.HeaderText = sHeaderValue
            grdCustodianListGrid.Columns.Add(HeaderTextBox)

        Next
        blnBuildHeaderRow = True

    End Function

    Private Function blnPopulateFieldMappingGrid(ByVal grdFieldMapping As DataGridView, ByVal asFields() As String) As Boolean
        blnPopulateFieldMappingGrid = False
        grdFieldMapping.Rows.Clear()

        For iCounter = 0 To UBound(asFields) - 1
            If iCounter = 0 Then
                grdFieldMapping.Rows.Add(asFields(iCounter), True)
                grdFieldMapping.Rows(0).ReadOnly = True

            Else
                grdFieldMapping.Rows.Add(asFields(iCounter), False)
            End If

        Next

        blnPopulateFieldMappingGrid = True
    End Function

    Private Sub btnSelectAllFields_Click(sender As Object, e As EventArgs) Handles btnSelectAllFields.Click
        If btnSelectAllFields.Text = "Select All" Then
            For Each row In grdCustodianFields.Rows
                If row.cells(0).value <> vbNullString Then
                    row.cells(1).value = True
                End If
            Next
            btnSelectAllFields.Text = "Deselect All"
        Else
            For Each row In grdCustodianFields.Rows
                If row.cells(0).value <> vbNullString Then
                    row.cells(1).value = False
                End If
            Next
            btnSelectAllFields.Text = "Select All"

        End If
    End Sub

    Private Sub btnSelectAllCustodians_Click(sender As Object, e As EventArgs) Handles btnSelectAllCustodians.Click
        If btnSelectAllCustodians.Text = "Select All" Then
            For Each row In grdCustomerCustodianList.Rows
                If row.cells(1).value <> vbNullString Then
                    row.cells(0).value = True
                End If
            Next
            btnSelectAllCustodians.Text = "Deselect All"
        Else
            For Each row In grdCustomerCustodianList.Rows
                If row.cells(1).value <> vbNullString Then
                    row.cells(0).value = False
                End If
            Next
            btnSelectAllCustodians.Text = "Select All"

        End If
    End Sub

    Private Sub btnExportCustodianInfo_Click(sender As Object, e As EventArgs) Handles btnExportCustodianInfo.Click

        Dim Custodianlist As StreamWriter

        Dim sCustodianListPath As String
        Dim sCustodianName As String
        Dim sValue As String

        sCustodianListPath = txtCustomerCustodianList.Text.Substring(0, txtCustomerCustodianList.Text.LastIndexOf("\"))

        Custodianlist = New StreamWriter(sCustodianListPath & "\NuixCustodianList.csv")

        Dim lstFieldColumns As List(Of Integer)
        lstFieldColumns = New List(Of Integer)

        For Each row In grdCustodianFields.Rows
            If row.cells(1).Value = True Then
                lstFieldColumns.Add(row.cells(0).rowindex + 1)
            End If
        Next

        For Each row In grdCustomerCustodianList.Rows
            If row.cells(0).value = True Then
                sCustodianName = row.cells(1).value
                For Each Column In lstFieldColumns
                    sValue = row.cells(CInt(Column.ToString)).value
                    Custodianlist.WriteLine(sCustodianName & "," & sValue)
                Next
            End If
        Next

        Custodianlist.Close()

        MessageBox.Show("Finished building NuixCustodianList.csv.  File is Located at - " & sCustodianListPath & "\" & "NuixCustodianList.csv", "Nuix Custodian List Generator")
    End Sub

    Private Sub NuixCustodianListGenerator_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class