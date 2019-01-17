<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FailuresForm
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FailuresForm))
        Me.treeFailuresView = New System.Windows.Forms.TreeView()
        Me.btnExportExceptionDetails = New System.Windows.Forms.Button()
        Me.NuixLogoPicture = New System.Windows.Forms.PictureBox()
        Me.FailureDetailsToolTip = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.NuixLogoPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'treeFailuresView
        '
        Me.treeFailuresView.CheckBoxes = True
        Me.treeFailuresView.Location = New System.Drawing.Point(8, 8)
        Me.treeFailuresView.Margin = New System.Windows.Forms.Padding(2)
        Me.treeFailuresView.Name = "treeFailuresView"
        Me.treeFailuresView.Size = New System.Drawing.Size(546, 362)
        Me.treeFailuresView.TabIndex = 0
        '
        'btnExportExceptionDetails
        '
        Me.btnExportExceptionDetails.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportExceptionDetails.BackColor = System.Drawing.Color.White
        Me.btnExportExceptionDetails.Image = CType(resources.GetObject("btnExportExceptionDetails.Image"), System.Drawing.Image)
        Me.btnExportExceptionDetails.Location = New System.Drawing.Point(472, 384)
        Me.btnExportExceptionDetails.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExportExceptionDetails.Name = "btnExportExceptionDetails"
        Me.btnExportExceptionDetails.Size = New System.Drawing.Size(80, 60)
        Me.btnExportExceptionDetails.TabIndex = 1
        Me.btnExportExceptionDetails.UseVisualStyleBackColor = False
        '
        'NuixLogoPicture
        '
        Me.NuixLogoPicture.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.NuixLogoPicture.Image = CType(resources.GetObject("NuixLogoPicture.Image"), System.Drawing.Image)
        Me.NuixLogoPicture.Location = New System.Drawing.Point(12, 375)
        Me.NuixLogoPicture.Name = "NuixLogoPicture"
        Me.NuixLogoPicture.Size = New System.Drawing.Size(75, 100)
        Me.NuixLogoPicture.TabIndex = 2
        Me.NuixLogoPicture.TabStop = False
        '
        'FailuresForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(563, 490)
        Me.Controls.Add(Me.NuixLogoPicture)
        Me.Controls.Add(Me.btnExportExceptionDetails)
        Me.Controls.Add(Me.treeFailuresView)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "FailuresForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Processing Details"
        CType(Me.NuixLogoPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents treeFailuresView As System.Windows.Forms.TreeView
    Friend WithEvents btnExportExceptionDetails As System.Windows.Forms.Button
    Friend WithEvents NuixLogoPicture As System.Windows.Forms.PictureBox
    Friend WithEvents FailureDetailsToolTip As System.Windows.Forms.ToolTip
End Class
