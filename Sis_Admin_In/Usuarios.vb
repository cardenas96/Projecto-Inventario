Imports System.Data.SqlClient

Public Class Usuarios

    Private Sub btn_nuevo_Click(sender As Object, e As EventArgs) Handles btn_nuevo.Click
        Dim n As Integer
        InventariO_DBDataSet1.USUARIOS.Clear()
        SqlDataAdapter1.Fill(InventariO_DBDataSet1.USUARIOS)
        dg_datos.Refresh()
        n = 0
        '
        n = USUARIOSBindingSource.Count + 1
        USUARIOSBindingSource.AddNew()
        txt_id.Text = n
        habilitar()
        btn_nuevo.Enabled = False
        btn_grabar.Enabled = True
    End Sub

    Private Sub Usuarios_Load(sender As Object, e As EventArgs) Handles Me.Load
        'TODO: esta línea de código carga datos en la tabla 'InventariO_DBDataSet1.USUARIOS' Puede moverla o quitarla según sea necesario.
        SqlDataAdapter1.Fill(InventariO_DBDataSet1.USUARIOS)
        deshabilitar()
        dg_datos.Refresh()
    End Sub
    Sub habilitar()
        txt_usuario.Enabled = True
        txt_contraseña.Enabled = True
    End Sub
    Sub deshabilitar()
        txt_usuario.Enabled = False
        txt_contraseña.Enabled = False
    End Sub

    Private Sub btn_grabar_Click(sender As Object, e As EventArgs) Handles btn_grabar.Click
        Try
            USUARIOSBindingSource.EndEdit()
            SqlDataAdapter1.Update(InventariO_DBDataSet1.USUARIOS)
            InventariO_DBDataSet1.USUARIOS.Clear()
            SqlDataAdapter1.Fill(InventariO_DBDataSet1.USUARIOS)
        Catch ex As System.Data.SqlClient.SqlException
            MessageBox.Show("Ya existe un producto con este ID." & USUARIOSBindingSource.Count & " " & vbCrLf &
                            "Se Almacenara en el ID " & USUARIOSBindingSource.Count + 1, "Error al registrar")
            USUARIOSBindingSource.Current(0) += 1
            USUARIOSBindingSource.EndEdit()
            SqlDataAdapter1.Update(InventariO_DBDataSet1.USUARIOS)
            InventariO_DBDataSet1.USUARIOS.Clear()
            SqlDataAdapter1.Fill(InventariO_DBDataSet1.USUARIOS)
        End Try
        btn_nuevo.Enabled = True
        btn_grabar.Enabled = False
        deshabilitar()
        dg_datos.Refresh()
    End Sub

    Private Sub btn_salir_Click(sender As Object, e As EventArgs) Handles btn_salir.Click
        Dispose()
    End Sub

    Private Sub SqlDataAdapter1_RowUpdated(sender As Object, e As SqlRowUpdatedEventArgs) Handles SqlDataAdapter1.RowUpdated
        If e.Status = UpdateStatus.ErrorsOccurred Then
            MessageBox.Show(e.Errors.Message & vbCrLf & vbCrLf &
                            e.Row.Item("NOMBRE", DataRowVersion.Original) & vbCrLf &
                            e.Row.Item("NOMBRE", DataRowVersion.Current))
            e.Status = UpdateStatus.SkipCurrentRow
        End If
        btn_nuevo.Enabled = True
    End Sub
End Class