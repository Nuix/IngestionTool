<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NuixCustodianListGenerator
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(NuixCustodianListGenerator))
        Me.txtCustomerCustodianList = New System.Windows.Forms.TextBox()
        Me.lblActiveDirectoryDumpFile = New System.Windows.Forms.Label()
        Me.btnCustodianListSelector = New System.Windows.Forms.Button()
        Me.grdCustomerCustodianList = New System.Windows.Forms.DataGridView()
        Me.btnLoadCustomerCustodians = New System.Windows.Forms.Button()
        Me.grdCustodianFields = New System.Windows.Forms.DataGridView()
        Me.FieldName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SelectForMapping = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.chkFirstRowHeader = New System.Windows.Forms.CheckBox()
        Me.btnSelectAllFields = New System.Windows.Forms.Button()
        Me.btnSelectAllCustodians = New System.Windows.Forms.Button()
        Me.btnExportCustodianInfo = New System.Windows.Forms.Button()
        CType(Me.grdCustomerCustodianList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdCustodianFields, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtCustomerCustodianList
        '
        Me.txtCustomerCustodianList.Location = New System.Drawing.Point(143, 20)
        Me.txtCustomerCustodianList.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtCustomerCustodianList.Name = "txtCustomerCustodianList"
        Me.txtCustomerCustodianList.Size = New System.Drawing.Size(277, 20)
        Me.txtCustomerCustodianList.TabIndex = 0
        '
        'lblActiveDirectoryDumpFile
        '
        Me.lblActiveDirectoryDumpFile.AutoSize = True
        Me.lblActiveDirectoryDumpFile.Location = New System.Drawing.Point(15, 20)
        Me.lblActiveDirectoryDumpFile.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblActiveDirectoryDumpFile.Name = "lblActiveDirectoryDumpFile"
        Me.lblActiveDirectoryDumpFile.Size = New System.Drawing.Size(123, 13)
        Me.lblActiveDirectoryDumpFile.TabIndex = 1
        Me.lblActiveDirectoryDumpFile.Text = "Customer Custodian List:"
        '
        'btnCustodianListSelector
        '
        Me.btnCustodianListSelector.Location = New System.Drawing.Point(423, 20)
        Me.btnCustodianListSelector.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCustodianListSelector.Name = "btnCustodianListSelector"
        Me.btnCustodianListSelector.Size = New System.Drawing.Size(28, 23)
        Me.btnCustodianListSelector.TabIndex = 2
        Me.btnCustodianListSelector.Text = "..."
        Me.btnCustodianListSelector.UseVisualStyleBackColor = True
        '
        'grdCustomerCustodianList
        '
        Me.grdCustomerCustodianList.BackgroundColor = System.Drawing.SystemColors.Control
        Me.grdCustomerCustodianList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdCustomerCustodianList.Location = New System.Drawing.Point(250, 92)
        Me.grdCustomerCustodianList.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.grdCustomerCustodianList.Name = "grdCustomerCustodianList"
        Me.grdCustomerCustodianList.RowTemplate.Height = 28
        Me.grdCustomerCustodianList.Size = New System.Drawing.Size(652, 347)
        Me.grdCustomerCustodianList.TabIndex = 3
        '
        'btnLoadCustomerCustodians
        '
        Me.btnLoadCustomerCustodians.Location = New System.Drawing.Point(565, 21)
        Me.btnLoadCustomerCustodians.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnLoadCustomerCustodians.Name = "btnLoadCustomerCustodians"
        Me.btnLoadCustomerCustodians.Size = New System.Drawing.Size(97, 23)
        Me.btnLoadCustomerCustodians.TabIndex = 4
        Me.btnLoadCustomerCustodians.Text = "Load Custodians"
        Me.btnLoadCustomerCustodians.UseVisualStyleBackColor = True
        '
        'grdCustodianFields
        '
        Me.grdCustodianFields.BackgroundColor = System.Drawing.SystemColors.Control
        Me.grdCustodianFields.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdCustodianFields.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.FieldName, Me.SelectForMapping})
        Me.grdCustodianFields.Location = New System.Drawing.Point(8, 92)
        Me.grdCustodianFields.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.grdCustodianFields.Name = "grdCustodianFields"
        Me.grdCustodianFields.RowTemplate.Height = 28
        Me.grdCustodianFields.Size = New System.Drawing.Size(238, 347)
        Me.grdCustodianFields.TabIndex = 5
        '
        'FieldName
        '
        Me.FieldName.HeaderText = "Field Name"
        Me.FieldName.Name = "FieldName"
        Me.FieldName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'SelectForMapping
        '
        Me.SelectForMapping.HeaderText = "Select For Mapping"
        Me.SelectForMapping.Name = "SelectForMapping"
        '
        'chkFirstRowHeader
        '
        Me.chkFirstRowHeader.AutoSize = True
        Me.chkFirstRowHeader.Location = New System.Drawing.Point(455, 26)
        Me.chkFirstRowHeader.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkFirstRowHeader.Name = "chkFirstRowHeader"
        Me.chkFirstRowHeader.Size = New System.Drawing.Size(108, 17)
        Me.chkFirstRowHeader.TabIndex = 6
        Me.chkFirstRowHeader.Text = "First Row Header"
        Me.chkFirstRowHeader.UseVisualStyleBackColor = True
        '
        'btnSelectAllFields
        '
        Me.btnSelectAllFields.Location = New System.Drawing.Point(8, 57)
        Me.btnSelectAllFields.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnSelectAllFields.Name = "btnSelectAllFields"
        Me.btnSelectAllFields.Size = New System.Drawing.Size(73, 31)
        Me.btnSelectAllFields.TabIndex = 7
        Me.btnSelectAllFields.Text = "Select All"
        Me.btnSelectAllFields.UseVisualStyleBackColor = True
        '
        'btnSelectAllCustodians
        '
        Me.btnSelectAllCustodians.Location = New System.Drawing.Point(250, 57)
        Me.btnSelectAllCustodians.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnSelectAllCustodians.Name = "btnSelectAllCustodians"
        Me.btnSelectAllCustodians.Size = New System.Drawing.Size(75, 31)
        Me.btnSelectAllCustodians.TabIndex = 8
        Me.btnSelectAllCustodians.Text = "Select All"
        Me.btnSelectAllCustodians.UseVisualStyleBackColor = True
        '
        'btnExportCustodianInfo
        '
        Me.btnExportCustodianInfo.Location = New System.Drawing.Point(803, 443)
        Me.btnExportCustodianInfo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnExportCustodianInfo.Name = "btnExportCustodianInfo"
        Me.btnExportCustodianInfo.Size = New System.Drawing.Size(99, 40)
        Me.btnExportCustodianInfo.TabIndex = 9
        Me.btnExportCustodianInfo.Text = "Export Custodian Data"
        Me.btnExportCustodianInfo.UseVisualStyleBackColor = True
        '
        'NuixCustodianListGenerator
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(919, 525)
        Me.Controls.Add(Me.btnExportCustodianInfo)
        Me.Controls.Add(Me.btnSelectAllCustodians)
        Me.Controls.Add(Me.btnSelectAllFields)
        Me.Controls.Add(Me.chkFirstRowHeader)
        Me.Controls.Add(Me.grdCustodianFields)
        Me.Controls.Add(Me.btnLoadCustomerCustodians)
        Me.Controls.Add(Me.grdCustomerCustodianList)
        Me.Controls.Add(Me.btnCustodianListSelector)
        Me.Controls.Add(Me.lblActiveDirectoryDumpFile)
        Me.Controls.Add(Me.txtCustomerCustodianList)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Name = "NuixCustodianListGenerator"
        Me.Text = "NuixCustodianListGenerator"
        CType(Me.grdCustomerCustodianList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdCustodianFields, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtCustomerCustodianList As System.Windows.Forms.TextBox
    Friend WithEvents lblActiveDirectoryDumpFile As System.Windows.Forms.Label
    Friend WithEvents btnCustodianListSelector As System.Windows.Forms.Button
    Friend WithEvents grdCustomerCustodianList As System.Windows.Forms.DataGridView
    Friend WithEvents btnLoadCustomerCustodians As System.Windows.Forms.Button
    Friend WithEvents grdCustodianFields As System.Windows.Forms.DataGridView
    Friend WithEvents chkFirstRowHeader As System.Windows.Forms.CheckBox
    Friend WithEvents FieldName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SelectForMapping As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents btnSelectAllFields As System.Windows.Forms.Button
    Friend WithEvents btnSelectAllCustodians As System.Windows.Forms.Button
    Friend WithEvents btnExportCustodianInfo As System.Windows.Forms.Button
End Class
