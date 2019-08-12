Imports System.Data.SqlClient

Public Class Bitacora
    Dim conexionsql As New SqlConnection("Data Source ='.'; Initial Catalog = 'BITACORA_DB'; Integrated security = true")
    Dim comando As SqlCommand = conexionsql.CreateCommand
    Dim lector As SqlDataReader
    Private Sub Bitacora_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dg_datos.Rows.Clear()
        conexionsql.Open()
        comando.CommandText = "SELECT * FROM Bitacora"
        lector = comando.ExecuteReader
        While lector.Read()
            dg_datos.Rows.Add(lector(0), lector(1), lector(2), lector(3), lector(4), lector(5), lector(6), lector(7))
        End While
        lector.Close()
        conexionsql.Close()
    End Sub
End Class