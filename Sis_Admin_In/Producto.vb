Imports System.Data.SqlClient

Public Class Producto
    Private Sub Producto_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'InventariO_DBDataSet1.PRODUCTO' Puede moverla o quitarla según sea necesario.
        'PRODUCTOTableAdapter.Fill(Me.InventariO_DBDataSet1.PRODUCTO)
        SqlDataAdapter1.Fill(InventariO_DBDataSet1.PRODUCTO)
        dg_datos.Refresh()
        deshabilitar()
    End Sub
    Private Sub cmd_inicio_Click(sender As Object, e As EventArgs) Handles cmd_inicio.Click
        PRODUCTOBindingSource.MoveFirst()
    End Sub

    Private Sub cmd_siguiente_Click(sender As Object, e As EventArgs) Handles cmd_siguiente.Click
        PRODUCTOBindingSource.MoveNext()
    End Sub

    Private Sub cmd_anterior_Click(sender As Object, e As EventArgs) Handles cmd_anterior.Click
        PRODUCTOBindingSource.MovePrevious()
    End Sub

    Private Sub cmd_ultimo_Click(sender As Object, e As EventArgs) Handles cmd_ultimo.Click
        PRODUCTOBindingSource.MoveLast()
    End Sub

    Private Sub btn_nuevo_Click(sender As Object, e As EventArgs) Handles btn_nuevo.Click
        Dim n As Integer
        InventariO_DBDataSet1.PRODUCTO.Clear()
        SqlDataAdapter1.Fill(InventariO_DBDataSet1.PRODUCTO)
        dg_datos.Refresh()
        n = 0
        n = PRODUCTOBindingSource.Count + 1
        PRODUCTOBindingSource.AddNew()
        txt_id.Text = n
        habilitar()
        btn_nuevo.Enabled = False
        btn_grabar.Enabled = True
        btn_modificar.Enabled = False
    End Sub
    Sub deshabilitar()
        txt_nombre.Enabled = False
        txt_marca.Enabled = False
        cb_unidad.Enabled = False
        txt_max.Enabled = False
        txt_min.Enabled = False
        txt_existencias.Enabled = False
        txt_costo.Enabled = False
        btn_aceptar.Enabled = False
        btn_grabar.Enabled = False
    End Sub
    Sub habilitar()
        txt_nombre.Enabled = True
        txt_marca.Enabled = True
        cb_unidad.Enabled = True
        txt_max.Enabled = True
        txt_min.Enabled = True
        txt_existencias.Enabled = True
        txt_costo.Enabled = True
    End Sub
    Sub desactivar()
        Me.cmd_inicio.Enabled = False
        Me.cmd_anterior.Enabled = False
        Me.cmd_siguiente.Enabled = False
        Me.cmd_ultimo.Enabled = False
    End Sub
    Sub activar()
        Me.cmd_inicio.Enabled = True
        Me.cmd_anterior.Enabled = True
        Me.cmd_siguiente.Enabled = True
        Me.cmd_ultimo.Enabled = True
    End Sub

    Private Sub btn_grabar_Click(sender As Object, e As EventArgs) Handles btn_grabar.Click
        deshabilitar()
        activar()
        'EMPLEADOTableAdapter.Update(InventariO_DBDataSet1.EMPLEADO)
        Try
            PRODUCTOBindingSource.EndEdit()
            PRODUCTOTableAdapter.Update(InventariO_DBDataSet1.PRODUCTO)
            SqlDataAdapter1.Update(InventariO_DBDataSet1.PRODUCTO)
            InventariO_DBDataSet1.PRODUCTO.Clear()
            SqlDataAdapter1.Fill(InventariO_DBDataSet1.PRODUCTO)
        Catch ex As System.Data.SqlClient.SqlException
            MessageBox.Show("Ya existe un producto con este ID " & PRODUCTOBindingSource.Count & " " & vbCrLf &
                            "Se Almacenara en el ID " & PRODUCTOBindingSource.Count + 1, "Error al registrar")
            PRODUCTOBindingSource.Current(0) += 1
            PRODUCTOBindingSource.EndEdit()
            PRODUCTOTableAdapter.Update(InventariO_DBDataSet1.PRODUCTO)
            SqlDataAdapter1.Update(InventariO_DBDataSet1.PRODUCTO)
            InventariO_DBDataSet1.PRODUCTO.Clear()
            SqlDataAdapter1.Fill(InventariO_DBDataSet1.PRODUCTO)
        End Try
        btn_grabar.Enabled = False
        btn_nuevo.Enabled = True
        btn_modificar.Enabled = True
        dg_datos.Refresh()
    End Sub

    Private Sub btn_salir_Click(sender As Object, e As EventArgs) Handles btn_salir.Click
        Dispose()
    End Sub

    Private Sub btn_modificar_Click(sender As Object, e As EventArgs) Handles btn_modificar.Click
        habilitar()
        btn_nuevo.Enabled = False
        btn_grabar.Enabled = True
        btn_modificar.Enabled = False
        btn_aceptar.Enabled = True
        btn_aceptar.Visible = True
        desactivar()
    End Sub

    Private Sub btn_cancelar_Click(sender As Object, e As EventArgs)
        PRODUCTOBindingSource.EndEdit()
        deshabilitar()
        activar()
        btn_grabar.Enabled = False
        btn_nuevo.Enabled = True
        btn_modificar.Enabled = True
        dg_datos.Refresh()
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

    Private Sub btn_aceptar_Click(sender As Object, e As EventArgs) Handles btn_aceptar.Click
        Try
            PRODUCTOBindingSource.EndEdit()
            'EMPLEADOTableAdapter.Update(InventariO_DBDataSet1.EMPLEADO)
            SqlDataAdapter1.Update(InventariO_DBDataSet1.PRODUCTO)
            ' EMPLEADOBindingSource.EndEdit()
            InventariO_DBDataSet1.PRODUCTO.Clear()
            SqlDataAdapter1.Fill(InventariO_DBDataSet1.PRODUCTO)
        Catch ex As System.Data.DBConcurrencyException
            MessageBox.Show("El Usuario Ah Sido Modificado Desde Otra Terminal")
            InventariO_DBDataSet1.PRODUCTO.Clear()
            SqlDataAdapter1.Fill(InventariO_DBDataSet1.PRODUCTO)
        End Try
        btn_aceptar.Visible = False
        btn_modificar.Enabled = True
        deshabilitar()
    End Sub

End Class