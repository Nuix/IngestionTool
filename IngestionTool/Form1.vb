Public Class Form1

    Private Sub Form1_BindingContextChanged(sender As Object, e As EventArgs) Handles Me.BindingContextChanged
        MsgBox(String.Format("TotalPhysicalMemory: {0} MBytes", System.Math.Round(My.Computer.Info.TotalPhysicalMemory / (1024 * 1024)), 2).ToString)
        MsgBox(String.Format("AvailablePhysicalMemory: {0} MBytes", System.Math.Round(My.Computer.Info.AvailablePhysicalMemory / (1024 * 1024)), 2).ToString)
    End Sub

    Private Sub btnNuixSystemSettings_Click(sender As Object, e As EventArgs) Handles btnNuixSystemSettings.Click
        NuixSettings.Show()
    End Sub
End Class
