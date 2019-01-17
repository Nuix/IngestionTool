<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.btnNuixSystemSettings = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnNuixSystemSettings
        '
        Me.btnNuixSystemSettings.FlatAppearance.BorderSize = 3
        Me.btnNuixSystemSettings.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnNuixSystemSettings.Location = New System.Drawing.Point(17, 402)
        Me.btnNuixSystemSettings.Name = "btnNuixSystemSettings"
        Me.btnNuixSystemSettings.Size = New System.Drawing.Size(130, 51)
        Me.btnNuixSystemSettings.TabIndex = 0
        Me.btnNuixSystemSettings.Text = "Nuix Settings..."
        Me.btnNuixSystemSettings.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1427, 465)
        Me.Controls.Add(Me.btnNuixSystemSettings)
        Me.Name = "Form1"
        Me.Text = "Nuix Office 365 Ingestion Tool"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnNuixSystemSettings As System.Windows.Forms.Button

End Class
