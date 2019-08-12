Imports System.Data.SqlClient

Public Class frm_Empleados
    Private Sub frm_Empleados_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'InventariO_DBDataSet1.EMPLEADO' Puede moverla o quitarla según sea necesario.
        'Me.EMPLEADOTableAdapter.Fill(Me.InventariO_DBDataSet1.EMPLEADO)
        SqlDataAdapter1.Fill(InventariO_DBDataSet1.EMPLEADO)
        deshabilitar()
    End Sub
    Private Sub cmd_inicio_Click(sender As Object, e As EventArgs) Handles cmd_inicio.Click
        EMPLEADOBindingSource.MoveFirst()
    End Sub

    Private Sub cmd_siguiente_Click(sender As Object, e As EventArgs) Handles cmd_siguiente.Click
        EMPLEADOBindingSource.MoveNext()
    End Sub

    Private Sub cmd_anterior_Click(sender As Object, e As EventArgs) Handles cmd_anterior.Click
        EMPLEADOBindingSource.MovePrevious()
    End Sub

    Private Sub cmd_ultimo_Click(sender As Object, e As EventArgs) Handles cmd_ultimo.Click
        EMPLEADOBindingSource.MoveLast()
    End Sub

    Private Sub btn_nuevo_Click(sender As Object, e As EventArgs) Handles btn_nuevo.Click
        'comando.CommandText = "SELECT COUNT(*) FROM EMPLEADO"
        ' Dim n As Integer = comando.ExecuteScalar() + 1
        Dim n As Integer
        InventariO_DBDataSet1.EMPLEADO.Clear()
        SqlDataAdapter1.Fill(InventariO_DBDataSet1.EMPLEADO)
        dg_datos.Refresh()
        n = 0
        n = EMPLEADOBindingSource.Count + 1
        EMPLEADOBindingSource.AddNew()
        txt_id.Text = n
        habilitar()
        btn_nuevo.Enabled = False
        btn_grabar.Enabled = True
        btn_modificar.Enabled = False
    End Sub
    Sub limpiar()
        txt_id.Text = ""
        txt_nombre.Text = ""
        txt_ciudad.Text = ""
        txt_colonia.Text = ""
        txt_domicilio.Text = ""
        txt_telefono.Text = ""
        cb_sex.Text = ""
        dtp_fecha.Text = Date.Today
    End Sub
    Sub habilitar()
        txt_nombre.Enabled = True
        txt_ciudad.Enabled = True
        txt_colonia.Enabled = True
        txt_domicilio.Enabled = True
        txt_telefono.Enabled = True
        cb_sex.Enabled = True
        dtp_fecha.Enabled = True
    End Sub
    Sub deshabilitar()
        txt_nombre.Enabled = False
        txt_ciudad.Enabled = False
        txt_colonia.Enabled = False
        txt_domicilio.Enabled = False
        txt_telefono.Enabled = False
        cb_sex.Enabled = False
        dtp_fecha.Enabled = False
        btn_Aceptar.Enabled = False
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
    Private Sub btn_grabar_Click(sender As Object, e As EventArgs) Handles btn_grabar.Click
        deshabilitar()
        activar()
        Try
            EMPLEADOBindingSource.EndEdit()
            EMPLEADOTableAdapter.Update(InventariO_DBDataSet1.EMPLEADO)
            SqlDataAdapter1.Update(InventariO_DBDataSet1.EMPLEADO)
            InventariO_DBDataSet1.EMPLEADO.Clear()
            SqlDataAdapter1.Fill(InventariO_DBDataSet1.EMPLEADO)
        Catch ex As System.Data.SqlClient.SqlException
            MessageBox.Show("Ya existe un producto con este ID." & EMPLEADOBindingSource.Count & " " & vbCrLf &
                                "Se Almacenara en el ID " & EMPLEADOBindingSource.Count + 1, "Error al registrar")
            EMPLEADOBindingSource.Current(0) += 1
            EMPLEADOBindingSource.EndEdit()
            EMPLEADOTableAdapter.Update(InventariO_DBDataSet1.EMPLEADO)
            SqlDataAdapter1.Update(InventariO_DBDataSet1.EMPLEADO)
            InventariO_DBDataSet1.EMPLEADO.Clear()
            SqlDataAdapter1.Fill(InventariO_DBDataSet1.EMPLEADO)
        End Try

        btn_grabar.Enabled = False
        btn_nuevo.Enabled = True
        btn_modificar.Enabled = True
        dg_datos.Refresh()

    End Sub

    Private Sub btn_modificar_Click(sender As Object, e As EventArgs) Handles btn_modificar.Click
        habilitar()
        btn_nuevo.Enabled = False
        btn_grabar.Enabled = True
        btn_modificar.Enabled = False
        btn_Aceptar.Enabled = True
        btn_Aceptar.Visible = True
        desactivar()
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

    Private Sub btn_Aceptar_Click(sender As Object, e As EventArgs) Handles btn_Aceptar.Click
        Try
            EMPLEADOBindingSource.EndEdit()
            'EMPLEADOTableAdapter.Update(InventariO_DBDataSet1.EMPLEADO)
            SqlDataAdapter1.Update(InventariO_DBDataSet1.EMPLEADO)
            ' EMPLEADOBindingSource.EndEdit()
            InventariO_DBDataSet1.EMPLEADO.Clear()
            SqlDataAdapter1.Fill(InventariO_DBDataSet1.EMPLEADO)
        Catch ex As System.Data.DBConcurrencyException
            MessageBox.Show("El Usuario Ah Sido Modificado Desde Otra Terminal")
            InventariO_DBDataSet1.EMPLEADO.Clear()
            SqlDataAdapter1.Fill(InventariO_DBDataSet1.EMPLEADO)
        End Try
        btn_Aceptar.Visible = False
        btn_modificar.Enabled = True
        deshabilitar()
    End Sub

    Private Sub frm_Empleados_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 Then
            System.Diagnostics.Process.Start("D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Ayuda.chm")
        End If
    End Sub
End Class