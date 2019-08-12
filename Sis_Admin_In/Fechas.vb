Public Class Fechas
    Private Sub btn_aceptar_Click(sender As Object, e As EventArgs) Handles btn_aceptar.Click
        Dim f1, f2 As String
        f1 = dtp_inicio.Value.Day & "/" & dtp_inicio.Value.Month & "/" & dtp_inicio.Value.Year
        f2 = dtp_final.Value.Day & "/" & dtp_final.Value.Month & "/" & dtp_final.Value.Year
        Reportes.lbl_f1.Text = f1
        Reportes.lbl_f2.Text = f2
        Dispose()
    End Sub

    Private Sub Fechas_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class