Imports System.Data.SqlClient
Public Class Consulta_Ventas
    Dim conexionsql As New SqlConnection("Data Source ='.'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true")
    Dim comando As SqlCommand = conexionsql.CreateCommand
    Dim lector As SqlDataReader
    Private Sub Consulta_Ventas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conexionsql.Open()
        'llenar combo proveedor
        comando.CommandText = "SELECT nombre FROM EMPLEADO"
        lector = comando.ExecuteReader
        While lector.Read()
            cb_empleado.Items.Add(lector(0))
        End While
        lector.Close()
        'llenar como IDCompra
        comando.CommandText = "SELECT idVenta FROM VENTA"
        lector = comando.ExecuteReader
        While lector.Read()
            cb_idventa.Items.Add(lector(0))
        End While
        lector.Close()
        'Llenar DataGrill
        comando.CommandText = "SELECT v.idVenta,v.fecha,p.nombre,dv.Cantidad,Precio FROM VENTA v 
                                INNER JOIN DETALLEVENTA dv ON v.idVenta = dv.idVenta
                                INNER JOIN PRODUCTO p ON p.idProducto = dv.idProducto
                                INNER JOIN EMPLEADO e ON v.idEmpleado = e.idEmpleado"
        lector = comando.ExecuteReader
        While lector.Read()
            dg_datos.Rows.Add(lector(0), lector(1), lector(2), lector(3), lector(4), (Val(lector(3)) * Val(lector(4))))
        End While
        lector.Close()
    End Sub

    Private Sub rb_empleado_CheckedChanged(sender As Object, e As EventArgs) Handles rb_empleado.CheckedChanged
        cb_empleado.Enabled = True
        dtp_inicio.Enabled = False
        dtp_final.Enabled = False
        cb_idventa.Enabled = False
        cb_empleado.Text = ""
        cb_idventa.Text = ""
    End Sub

    Private Sub rb_fecha_CheckedChanged(sender As Object, e As EventArgs) Handles rb_fecha.CheckedChanged
        cb_empleado.Enabled = False
        dtp_inicio.Enabled = True
        dtp_final.Enabled = True
        cb_idventa.Enabled = False
        cb_empleado.Text = ""
        cb_idventa.Text = ""
    End Sub

    Private Sub rb_todos_CheckedChanged(sender As Object, e As EventArgs) Handles rb_todos.CheckedChanged
        cb_empleado.Enabled = False
        dtp_inicio.Enabled = False
        dtp_final.Enabled = False
        cb_idventa.Enabled = False
        cb_empleado.Text = ""
        cb_idventa.Text = ""
    End Sub

    Private Sub rb_idventa_CheckedChanged(sender As Object, e As EventArgs) Handles rb_idventa.CheckedChanged, rb_idventa.CheckedChanged
        cb_empleado.Enabled = False
        dtp_inicio.Enabled = False
        dtp_final.Enabled = False
        cb_idventa.Enabled = True
        cb_empleado.Text = ""
        cb_idventa.Text = ""
    End Sub
    Sub Consulta(cadena As String)
        dg_datos.Rows.Clear()
        Dim suma As Double = 0.0
        comando.CommandText = "Select v.idVenta,v.fecha,p.nombre,dv.Cantidad,Precio FROM VENTA v 
                                INNER Join DETALLEVENTA dv ON v.idVenta = dv.idVenta
                                INNER Join PRODUCTO p ON p.idProducto = dv.idProducto
                                INNER Join EMPLEADO e ON v.idEmpleado = e.idEmpleado " + cadena
        lector = comando.ExecuteReader
        While lector.Read()
            dg_datos.Rows.Add(lector(0), lector(1), lector(2), lector(3), lector(4), (Val(lector(3)) * Val(lector(4))))
            suma = suma + (Val(lector(3)) * Val(lector(4)))
        End While
        lector.Close()
        lbl_total.Text = suma.ToString
    End Sub

    Private Sub btn_buscar_Click(sender As Object, e As EventArgs) Handles btn_buscar.Click
        Dim cadena As String
        If rb_empleado.Checked = True Then
            If cb_empleado.Text = "" Then
                MessageBox.Show("Tenemos Un Problema", "No Hay Proveedor Seleccionado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                cadena = "WHERE e.nombre = '" & cb_empleado.Text & "' "
                Consulta(cadena)
            End If
        ElseIf rb_fecha.Checked = True Then
            If dtp_inicio.Text = "" Or dtp_final.Text = "" Then
                dtp_inicio.Text = Today
                dtp_final.Text = Today
            ElseIf dtp_inicio.Text = dtp_final.Text Then
                cadena = "WHERE v.fecha = '" & dtp_inicio.Text & "' "
                Consulta(cadena)
            Else
                cadena = "WHERE v.fecha BETWEEN '" & dtp_inicio.Text & "' AND '" & dtp_final.Text & "'"
                Consulta(cadena)
            End If
        ElseIf rb_idventa.Checked = True Then
            If cb_idventa.Text = "" Then
                MessageBox.Show("Tenemos Un Problema", "No Hay Venta Seleccionada", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                cadena = "WHERE v.idVenta = '" & cb_idventa.Text & "' "
                Consulta(cadena)
            End If
        Else
            Consulta(" ")
        End If
    End Sub

    Private Sub Consulta_Ventas_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        conexionsql.Close()
        Dispose()
    End Sub

    Private Sub btn_salir_Click(sender As Object, e As EventArgs) Handles btn_salir.Click
        conexionsql.Close()
        Dispose()
    End Sub

    Private Sub cb_empleado_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cb_empleado.KeyPress
        e.KeyChar = Chr(0)
    End Sub

    Private Sub cb_idventa_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cb_idventa.KeyPress
        e.KeyChar = Chr(0)
    End Sub
End Class