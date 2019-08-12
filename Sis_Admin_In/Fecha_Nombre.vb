Imports System.Data.SqlClient
Public Class Fecha_Nombre
    Dim conexionsql As New SqlConnection("Data Source ='.'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
    Dim comando As SqlCommand = conexionsql.CreateCommand
    Dim lector As SqlDataReader
    Private Sub Fecha_Nombre_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conexionsql.Open()
        cb_nombre.Items.Clear()
        If Reportes.lbl_f1.Text = "1" Then
            comando.CommandText = "SELECT nombre FROM EMPLEADO"
            lector = comando.ExecuteReader
            While lector.Read
                cb_nombre.Items.Add(lector(0))
            End While
            lector.Close()
        ElseIf Reportes.lbl_f1.Text = "2" Then
            comando.CommandText = "SELECT nombre FROM PROVEEDOR"
            lector = comando.ExecuteReader
            While lector.Read
                cb_nombre.Items.Add(lector(0))
            End While
            lector.Close()
        End If
    End Sub

    Private Sub btn_aceptar_Click(sender As Object, e As EventArgs) Handles btn_aceptar.Click
        Dim f1, f2 As String
        f1 = dtp_inicio.Value.Day & "/" & dtp_inicio.Value.Month & "/" & dtp_inicio.Value.Year
        f2 = dtp_final.Value.Day & "/" & dtp_final.Value.Month & "/" & dtp_final.Value.Year
        Reportes.lbl_f1.Text = f1
        Reportes.lbl_f2.Text = f2
        Reportes.lbl_n1.Text = cb_nombre.Text
        Dispose()
    End Sub
End Class