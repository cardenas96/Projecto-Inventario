Imports System.Data.SqlClient

Public Class Proveedor
    Private Sub Proveedor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'INVENTARIO_DBDataSet.PROVEEDOR' Puede moverla o quitarla según sea necesario.
        'Me.PROVEEDORTableAdapter.Fill(Me.INVENTARIO_DBDataSet.PROVEEDOR)
        SqlDataAdapter1.Fill(INVENTARIO_DBDataSet.PROVEEDOR)
        deshabilitar()
        If PROVEEDORBindingSource.Count = 0 Then
            btn_modificar.Enabled = False
            desactivar()
        End If
        dg_datos.Refresh()
    End Sub
    Sub habilitar()
        txt_nombre.Enabled = True
        txt_ciudad.Enabled = True
        txt_colonia.Enabled = True
        txt_domicilio.Enabled = True
        txt_telefono.Enabled = True
        txt_contacto.Enabled = True
        txt_fax.Enabled = True
        txt_cp.Enabled = True
    End Sub
    Sub deshabilitar()
        txt_nombre.Enabled = False
        txt_ciudad.Enabled = False
        txt_colonia.Enabled = False
        txt_domicilio.Enabled = False
        txt_telefono.Enabled = False
        txt_contacto.Enabled = False
        txt_fax.Enabled = False
        txt_cp.Enabled = False
        btn_aceptar.Enabled = False
        btn_grabar.Enabled = False
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
    Private Sub cmd_inicio_Click(sender As Object, e As EventArgs) Handles cmd_inicio.Click
        PROVEEDORBindingSource.MoveFirst()
    End Sub

    Private Sub cmd_siguiente_Click(sender As Object, e As EventArgs) Handles cmd_siguiente.Click
        PROVEEDORBindingSource.MoveNext()
    End Sub

    Private Sub cmd_anterior_Click(sender As Object, e As EventArgs) Handles cmd_anterior.Click
        PROVEEDORBindingSource.MovePrevious()
    End Sub

    Private Sub cmd_ultimo_Click(sender As Object, e As EventArgs) Handles cmd_ultimo.Click
        PROVEEDORBindingSource.MoveLast()
    End Sub

    Private Sub btn_nuevo_Click(sender As Object, e As EventArgs) Handles btn_nuevo.Click
        Dim n As Integer
        INVENTARIO_DBDataSet.PROVEEDOR.Clear()
        SqlDataAdapter1.Fill(INVENTARIO_DBDataSet.PROVEEDOR)
        dg_datos.Refresh()
        n = 0
        '
        n = PROVEEDORBindingSource.Count + 1
        PROVEEDORBindingSource.AddNew()
        txt_id.Text = n
        habilitar()
        btn_nuevo.Enabled = False
        btn_grabar.Enabled = True
        btn_modificar.Enabled = False
    End Sub

    Private Sub btn_grabar_Click(sender As Object, e As EventArgs) Handles btn_grabar.Click
        deshabilitar()
        activar()
        Try
            PROVEEDORBindingSource.EndEdit()
            PROVEEDORTableAdapter.Update(INVENTARIO_DBDataSet.PROVEEDOR)
            SqlDataAdapter1.Update(INVENTARIO_DBDataSet.PROVEEDOR)
            INVENTARIO_DBDataSet.PROVEEDOR.Clear()
            SqlDataAdapter1.Fill(INVENTARIO_DBDataSet.PROVEEDOR)
        Catch ex As System.Data.SqlClient.SqlException
            MessageBox.Show("Ya existe un producto con este ID." & PROVEEDORBindingSource.Count & " " & vbCrLf &
                            "Se Almacenara en el ID " & PROVEEDORBindingSource.Count + 1, "Error al registrar")
            PROVEEDORBindingSource.Current(0) += 1
            PROVEEDORBindingSource.EndEdit()
            PROVEEDORTableAdapter.Update(INVENTARIO_DBDataSet.PROVEEDOR)
            SqlDataAdapter1.Update(INVENTARIO_DBDataSet.PROVEEDOR)
            INVENTARIO_DBDataSet.PROVEEDOR.Clear()
            SqlDataAdapter1.Fill(INVENTARIO_DBDataSet.PROVEEDOR)
        End Try
        btn_nuevo.Enabled = True
        btn_modificar.Enabled = True
        dg_datos.Refresh()
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


    Private Sub btn_salir_Click(sender As Object, e As EventArgs) Handles btn_salir.Click
        Dispose()
    End Sub

    Private Sub btn_aceptar_Click(sender As Object, e As EventArgs) Handles btn_aceptar.Click
        Try
            PROVEEDORBindingSource.EndEdit()
            'EMPLEADOTableAdapter.Update(InventariO_DBDataSet1.EMPLEADO)
            SqlDataAdapter1.Update(INVENTARIO_DBDataSet.PROVEEDOR)
            ' EMPLEADOBindingSource.EndEdit()
            INVENTARIO_DBDataSet.PROVEEDOR.Clear()
            SqlDataAdapter1.Fill(INVENTARIO_DBDataSet.PROVEEDOR)
        Catch ex As System.Data.DBConcurrencyException
            MessageBox.Show("El Usuario Ah Sido Modificado Desde Otra Terminal")
            INVENTARIO_DBDataSet.PROVEEDOR.Clear()
            SqlDataAdapter1.Fill(INVENTARIO_DBDataSet.PROVEEDOR)
        End Try
        btn_aceptar.Visible = False
        btn_modificar.Enabled = True
        deshabilitar()
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