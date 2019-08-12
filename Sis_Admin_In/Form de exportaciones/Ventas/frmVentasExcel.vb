Imports System.Data.SqlClient
Public Class frmVentasExcel
    'Cadena para obtener los campos seleccionados de la consulta
    Dim contador As Integer

    Private Sub frmVentasExcel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Using conexionRemota As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true")
            Dim comandoRemoto As SqlCommand = conexionRemota.CreateCommand
            Dim lectorRemoto As SqlDataReader

            conexionRemota.Open()
            comandoRemoto.CommandText = "SELECT
		                                    t.name, c.name, ti.name, c.max_length
	                                    FROM
		                                    sys.tables t left join
		                                    sys.all_columns c on (c.object_id = t.object_id) left join
		                                    sys.types ti on (c.system_type_id = ti.system_type_id)
	                                    WHERE 
		                                    NOT t.name ='PROVEEDOR' 
		                                    AND NOT t.name ='COMPRA' 
		                                    AND NOT t.name ='DETALLECOMPRA' 
		                                    AND NOT t.name ='USUARIOS' 
		                                    AND NOT c.name ='idProducto' 
		                                    AND NOT c.name ='idEmpleado'
		                                    AND NOT c.name ='idVenta'
		                                    AND NOT c.name ='costo'
		                                    AND NOT c.name ='ciudad'
		                                    AND NOT c.name ='colonia'
		                                    AND NOT c.name ='telefono'
		                                    AND NOT c.name ='nacimiento'
		                                    AND NOT c.name ='sexo'
		                                    AND NOT c.name ='existencas'
		                                    AND NOT c.name ='minimo'
		                                    AND NOT c.name ='maximmo'
		                                    AND NOT c.name ='direccion'
	                                    ORDER BY t.name DESC;"
            lectorRemoto = comandoRemoto.ExecuteReader
            While lectorRemoto.Read
                dgTablas.Rows.Add(lectorRemoto(0), lectorRemoto(1))
            End While

            lectorRemoto.Close()


            comandoRemoto.CommandText = "SELECT  
	                                        v.idVenta,   e.nombre,
	                                        v.fecha,     p.nombre,
	                                        dv.Cantidad, dv.Precio,
	                                        p.marca,     p.unidad,
	                                        v.total
                                        FROM VENTA v 
	                                        INNER JOIN DETALLEVENTA dv ON v.idVenta = dv.idVenta 
                                            INNER JOIN PRODUCTO p ON p.idProducto = dv.idProducto 
	                                        INNER JOIN EMPLEADO e ON e.idEmpleado = v.idEmpleado 
                                        ORDER BY v.idVenta ASC;"
            lectorRemoto = comandoRemoto.ExecuteReader
            While lectorRemoto.Read
                dgFinal.Rows.Add(lectorRemoto(0), lectorRemoto(1), lectorRemoto(2),
                                 lectorRemoto(3), lectorRemoto(4), lectorRemoto(5),
                                 lectorRemoto(6), lectorRemoto(7), lectorRemoto(8))
            End While

            lectorRemoto.Close()

            Dim t As Integer


            For t = 0 To dgFinal.ColumnCount - 1
                    dgFinal.Columns(t).Visible = False
                Next

            conexionRemota.Close()
        End Using





    End Sub

    Private Sub btn_salir_Click(sender As Object, e As EventArgs) Handles btn_salir.Click
        'Cerrar esta conexion puede llegar a generar errores si es que no se cierran las conexiones en
        'otros lados
        conexionsql.Close()
        Dispose()
    End Sub

    Private Sub btnGenerar_Click(sender As Object, e As EventArgs) Handles btnGenerar.Click

        If contador > 0 Then
            dgFinal.Columns(0).Visible = True
            llenarExcel(dgFinal)

        Else
                MsgBox("NO SE HAN SELECCIONADO CAMPOS PARA LA CONSULTA")
        End If



    End Sub

    Private Sub dgTablas_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgTablas.CellContentClick
        Dim seleccionado As Boolean = dgTablas.Rows(e.RowIndex).Cells(2).Value

        If (seleccionado) Then
            'MsgBox(seleccionado)
            'MessageBox.Show("UnChecked", "Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Dim terminacion As String = dgTablas.Rows(e.RowIndex).Cells(0).Value
            Dim nombreCampo As String = dgTablas.Rows(e.RowIndex).Cells(1).Value + terminacion.Chars(0)

            '    'dgFinal.Columns.Remove(nombreCampo)
            dgFinal.Columns(nombreCampo).Visible = False



        Else
            ' MsgBox(seleccionado)
            'MessageBox.Show("The Value is Checked", "Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Dim col As New DataGridViewTextBoxColumn
            Dim terminacion As String = dgTablas.Rows(e.RowIndex).Cells(0).Value
            Dim nombreCampo As String = dgTablas.Rows(e.RowIndex).Cells(1).Value + terminacion.Chars(0)
            'col.Name = dgTablas.Rows(e.RowIndex).Cells(1).Value
            'dgFinal.Columns.Add(col)
            dgFinal.Columns(nombreCampo).Visible = True
            contador += 1
        End If



    End Sub
End Class