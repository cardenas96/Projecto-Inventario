Imports System.Data.SqlClient
Public Class Consulta_Compra
    Dim conexionsql As New SqlConnection("Data Source ='.'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true")
    Dim comando As SqlCommand = conexionsql.CreateCommand
    Dim lector As SqlDataReader
    Private Sub Consulta_Compra_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conexionsql.Open()
        'llenar combo proveedor
        comando.CommandText = "SELECT nombre FROM PROVEEDOR"
        lector = comando.ExecuteReader
        While lector.Read()
            cb_proveedor.Items.Add(lector(0))
        End While
        lector.Close()
        'llenar como IDCompra
        comando.CommandText = "SELECT idCompra FROM COMPRA"
        lector = comando.ExecuteReader
        While lector.Read()
            cb_idcompra.Items.Add(lector(0))
        End While
        lector.Close()
        'Llenar DataGrill
        comando.CommandText = "SELECT c.idCompra,pr.nombre,c.fecha,p.nombre,dc.cantidad,dc.costo FROM COMPRA c 
                                INNER JOIN DETALLECOMPRA dc ON c.idCompra = dc.idCompra 
                                INNER JOIN PRODUCTO p ON p.idProducto = dc.idProducto 
                                INNER JOIN PROVEEDOR pr on pr.idProveedor	= c.idProveedor"
        lector = comando.ExecuteReader
        While lector.Read()
            dg_datos.Rows.Add(lector(0), lector(1), lector(2), lector(3), lector(4), lector(5))
        End While
        lector.Close()
    End Sub

    Private Sub rb_proveedor_CheckedChanged(sender As Object, e As EventArgs) Handles rb_proveedor.CheckedChanged
        cb_proveedor.Enabled = True
        dtp_inicio.Enabled = False
        dtp_final.Enabled = False
        cb_idcompra.Enabled = False
        cb_proveedor.Text = ""
        cb_idcompra.Text = ""
    End Sub

    Private Sub rb_fecha_CheckedChanged(sender As Object, e As EventArgs) Handles rb_fecha.CheckedChanged
        cb_proveedor.Enabled = False
        dtp_inicio.Enabled = True
        dtp_final.Enabled = True
        cb_idcompra.Enabled = False
        cb_proveedor.Text = ""
        cb_idcompra.Text = ""
    End Sub

    Private Sub rb_idcompra_CheckedChanged(sender As Object, e As EventArgs) Handles rb_idcompra.CheckedChanged
        cb_proveedor.Enabled = False
        dtp_inicio.Enabled = False
        dtp_final.Enabled = False
        cb_idcompra.Enabled = True
        cb_proveedor.Text = ""
        cb_idcompra.Text = ""
    End Sub

    Private Sub rb_todos_CheckedChanged(sender As Object, e As EventArgs) Handles rb_todos.CheckedChanged
        cb_proveedor.Enabled = False
        dtp_inicio.Enabled = False
        dtp_final.Enabled = False
        cb_idcompra.Enabled = False
        cb_proveedor.Text = ""
        cb_idcompra.Text = ""
    End Sub

    Private Sub btn_buscar_Click(sender As Object, e As EventArgs) Handles btn_buscar.Click
        Dim cadena As String
        If rb_proveedor.Checked = True Then
            If cb_proveedor.Text = "" Then
                MessageBox.Show("Tenemos Un Problema", "No Hay Proveedor Seleccionado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                cadena = "WHERE pr.nombre = '" & cb_proveedor.Text & "' "
                consulta(cadena)
            End If
        ElseIf rb_fecha.Checked = True Then
            If dtp_inicio.Text = "" Or dtp_final.Text = "" Then
                dtp_inicio.Text = Today
                dtp_final.Text = Today
            ElseIf dtp_inicio.Text = dtp_final.Text Then
                cadena = "WHERE c.fecha = '" & dtp_inicio.Text & "' "
                consulta(cadena)
            Else
                cadena = "WHERE c.fecha BETWEEN '" & dtp_inicio.Text & "' AND '" & dtp_final.Text & "'"
                consulta(cadena)
            End If
        ElseIf rb_idcompra.Checked = True Then
            If cb_idcompra.Text = "" Then
                MessageBox.Show("Tenemos Un Problema", "No Hay Compra Seleccionada", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                cadena = "WHERE c.idCompra = '" & cb_idcompra.Text & "' "
                consulta(cadena)
            End If
        Else
            consulta(" ")
        End If
    End Sub
    Sub consulta(cadena As String)
        dg_datos.Rows.Clear()
        comando.CommandText = "SELECT c.idCompra,pr.nombre,c.fecha,p.nombre,dc.cantidad,dc.costo FROM COMPRA c 
                                INNER JOIN DETALLECOMPRA dc ON c.idCompra = dc.idCompra 
                                INNER JOIN PRODUCTO p ON p.idProducto = dc.idProducto 
                                INNER JOIN PROVEEDOR pr on pr.idProveedor	= c.idProveedor " + cadena
        lector = comando.ExecuteReader
        While lector.Read()
            dg_datos.Rows.Add(lector(0), lector(1), lector(2), lector(3), lector(4), lector(5))
        End While
        lector.Close()
    End Sub

    Private Sub cb_idcompra_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cb_idcompra.KeyPress
        e.KeyChar = Chr(0)
    End Sub

    Private Sub cb_proveedor_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cb_proveedor.KeyPress
        e.KeyChar = Chr(0)
    End Sub

    Private Sub btn_salir_Click(sender As Object, e As EventArgs) Handles btn_salir.Click
        Dispose()
        conexionsql.Close()
    End Sub

    Private Sub Consulta_Compra_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        conexionsql.Close()
    End Sub

End Class