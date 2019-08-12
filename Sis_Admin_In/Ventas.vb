Imports System.Data.SqlClient
Imports System.Configuration
Public Class Ventas
    Public suma As Integer

    Private Sub Ventas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Using Conexion As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true")
            Using Adaptador As New SqlDataAdapter("EMPLE", Conexion)
                Adaptador.SelectCommand.CommandType = CommandType.StoredProcedure
                Dim Data As New DataSet
                Adaptador.Fill(Data, "EMPLEADO")
                cb_empleado.DataSource = Data.Tables("EMPLEADO")
                cb_empleado.ValueMember = "nombre"
            End Using
            Using Adaptador As New SqlDataAdapter("PRODUC", Conexion)
                Adaptador.SelectCommand.CommandType = CommandType.StoredProcedure
                Dim Data As New DataSet
                Adaptador.Fill(Data, "PRODUCTO")
                cb_articulo.DataSource = Data.Tables("PRODUCTO")
                cb_articulo.ValueMember = "nombre"
            End Using
        End Using

    End Sub
    Sub insertarError(linea As Integer, corrida As String, ex As Exception, formulario As String, usuario As String, fecha As String, hora As String)
        Dim con As New SqlConnection("Data Source ='.'; Initial Catalog = 'BITACORA_DB'; Integrated security = true")
        con.Open()
        Dim com As New SqlCommand
        com.Connection = con
        com.CommandText = "SELECT COUNT(*) FROM Bitacora"
        Dim no As Integer = com.ExecuteScalar() + 1

        Try
            com.CommandText = "INSERT INTO Bitacora VALUES(" & no & ",'" & corrida & "','" & ex.Message.Replace("'", " ") & "','" & formulario & "'," & Val(linea) & ",'" & usuario & "','" & fecha & "','" & hora & "') "
            com.ExecuteNonQuery()

        Catch e As Exception
            MsgBox(e.Message, MsgBoxStyle.MsgBoxHelp, "Error De Bitacora")
        End Try
        con.Close()
    End Sub

    Private Sub cb_empleado_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_empleado.SelectedIndexChanged
        Using Conexion As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true")
            Using comando As New SqlCommand
                comando.Connection = Conexion
                Try
                    comando.CommandType = CommandType.StoredProcedure
                    comando.CommandText = "BUSCAEMPLEADO"
                    comando.Parameters.Add("@EMPLEADO", SqlDbType.VarChar).Value = cb_empleado.Text
                Catch ex As Exception
                    Dim linea As Integer, corrida, formulario, usuario, fecha, hora As String
                    linea = 48
                    corrida = "Error En Obtencion De Datos"
                    'descripcion = ex.Message
                    formulario = "Ventas.vb"
                    usuario = Inicio_Sesion.txt_usuario.Text
                    fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
                    hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
                    insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
                End Try
                Using Adaptador As New SqlDataAdapter
                    Adaptador.SelectCommand = comando
                    Dim Data As New DataSet
                    Try
                        Adaptador.Fill(Data, "EMPLEADO")
                        txt_ide.Text = CStr(Data.Tables("EMPLEADO").Rows(0).Item(0))
                        txt_direccion.Text = CStr(Data.Tables("EMPLEADO").Rows(0).Item(2))
                        txt_telefono.Text = CStr(Data.Tables("EMPLEADO").Rows(0).Item(1))
                    Catch ex As Exception
                        'MessageBox.Show("Error: " & ex.Message & " ", "Error")
                    End Try
                End Using
            End Using
        End Using
    End Sub

    Private Sub cb_articulo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_articulo.SelectedIndexChanged
        Using Conexion As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true")
            Using comando As New SqlCommand
                comando.Connection = Conexion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "BUSCAPRODUCTO"
                Try
                    comando.Parameters.Add("@ARTICULO", SqlDbType.VarChar).Value = cb_articulo.Text
                Catch ex As Exception
                    Dim linea As Integer, corrida, formulario, usuario, fecha, hora As String
                    linea = 83
                    corrida = "Error En Obtencion De Datos"
                    'descripcion = ex.Message
                    formulario = "Ventas.vb"
                    usuario = Inicio_Sesion.txt_usuario.Text
                    fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
                    hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
                    insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
                End Try
                Using Adaptador As New SqlDataAdapter
                    Adaptador.SelectCommand = comando
                    Dim Data As New DataSet
                    Try
                        Adaptador.Fill(Data, "PRODUCTO")
                        txt_ida.Text = CStr(Data.Tables("PRODUCTO").Rows(0).Item(0))
                        txt_marca.Text = CStr(Data.Tables("PRODUCTO").Rows(0).Item(2))
                        txt_unidad.Text = CStr(Data.Tables("PRODUCTO").Rows(0).Item(3))
                        txt_existencias.Text = CStr(Data.Tables("PRODUCTO").Rows(0).Item(5))
                        txt_precio.Text = CStr(Data.Tables("PRODUCTO").Rows(0).Item(4) * 1.2)
                    Catch ex As Exception
                        Dim linea As Integer, corrida, formulario, usuario, fecha, hora As String
                        linea = 99
                        corrida = "Error En Obtencion De Datos"
                        'descripcion = ex.Message
                        formulario = "Ventas.vb"
                        usuario = Inicio_Sesion.txt_usuario.Text
                        fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
                        hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
                        insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
                    End Try
                End Using
            End Using
        End Using
    End Sub

    Private Sub btn_agregar_Click(sender As Object, e As EventArgs) Handles btn_agregar.Click
        If txt_cantidad.Text = "" Then
            MessageBox.Show("Introduzca Cantidad A Vender", "¡Error Campo Vacio!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txt_cantidad.Focus()
        ElseIf cb_articulo.Text = "" Then
            MessageBox.Show("Seleccione Un Producto...", "¡Error Campo Vacio!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cb_articulo.Focus()
        Else
            Dim ban As Boolean = False
            For i = 0 To dg_datos.RowCount - 2
                If Val(txt_ida.Text) = Val(dg_datos.Item(0, i).Value) Then
                    ban = True
                End If
            Next
            If ban = False Then
                dg_datos.Rows.Add(txt_ida.Text, cb_articulo.Text, txt_unidad.Text, txt_cantidad.Text, txt_precio.Text, (Val(txt_precio.Text) * Val(txt_cantidad.Text)))
                lblSub.Text = Val(lblSub.Text) + (Val(txt_precio.Text) * Val(txt_cantidad.Text))
                lblIva.Text = CDbl(lblSub.Text) * 0.16
                lbl_total.Text = CDbl(lblSub.Text) + CDbl(lblIva.Text)
                limpiar()
            Else
                For i = 0 To dg_datos.RowCount - 2
                    If Val(txt_ida.Text) = Val(dg_datos.Item(0, i).Value) Then
                        dg_datos.Item(3, i).Value = Val(dg_datos.Item(3, i).Value) + Val(txt_cantidad.Text)
                        lblSub.Text = Val(lblSub.Text) + (Val(txt_cantidad.Text) * Val(dg_datos.Item(4, i).Value))
                        dg_datos.Item(5, i).Value = (Val(dg_datos.Item(3, i).Value) * Val(dg_datos.Item(4, i).Value))
                        lblIva.Text = CDbl(lblSub.Text) * 0.16
                        lbl_total.Text = CDbl(lblSub.Text) + CDbl(lblIva.Text)
                        limpiar()
                    End If
                Next
            End If
        End If

    End Sub
    Private Sub btn_limpiar_Click(sender As Object, e As EventArgs) Handles btn_limpiar.Click
        limpiar()
    End Sub

    Private Sub btn_nuevo_Click(sender As Object, e As EventArgs) Handles btn_nuevo.Click
        suma = 0
        Using Conexion As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true")
            Using comando As New SqlCommand
                comando.Connection = Conexion
                Conexion.Open()
                comando.CommandText = "SELECT IDENT_CURRENT('VENTA')"
                Dim lector As SqlDataReader
                lector = comando.ExecuteReader()
                lector.Read()
                'If Val(txt_id.Text) = lector(0) Then
                txt_id.Text = lector(0) + 1
                'Else
                '    txt_id.Text = lector(0)
                'End If

                lector.Close()
                Conexion.Close()
                'comando.CommandType = CommandType.StoredProcedure
                'comando.CommandText = "VENTAS"
                ' Using Adaptador As New SqlDataAdapter
                'Adaptador.SelectCommand = comando
                'Dim Data As New DataSet
                'Adaptador.Fill(Data, "VENTA")
                'txt_id.Text = Data.Tables("VENTA").Rows.Count + 1
                '   End Using
            End Using
        End Using
        btn_agregar.Enabled = True
        btn_grabar.Enabled = True
        btn_nuevo.Enabled = False
    End Sub
    Sub limpiar()
        txt_ida.Text = ""
        cb_articulo.Text = ""
        txt_marca.Text = ""
        txt_precio.Text = ""
        txt_unidad.Text = ""
        txt_existencias.Text = ""
        txt_cantidad.Text = ""
    End Sub

    Private Sub btn_grabar_Click(sender As Object, e As EventArgs) Handles btn_grabar.Click
        Using Conexion As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true")
            Using comando As New SqlCommand
                Conexion.Open()
                comando.Connection = Conexion
                'comando.CommandType = CommandType.StoredProcedure
                Try
                    comando.CommandText = "INSERT INTO VENTA VALUES(" & Val(txt_ide.Text) & ",'" & dtp_fecha.Text & "'," & Val(lbl_total.Text) & ")"
                    comando.ExecuteNonQuery()
                Catch ex As Exception
                    Dim linea As Integer, corrida, formulario, usuario, fecha, hora As String
                    linea = 158
                    corrida = "Error En Trasaccion"
                    'descripcion = ex.Message
                    formulario = "Venta.vb"
                    usuario = Inicio_Sesion.txt_usuario.Text
                    fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
                    hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
                    insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
                End Try
                comando.CommandText = "SELECT IDENT_CURRENT('VENTA')"
                Dim lector As SqlDataReader
                lector = comando.ExecuteReader()
                lector.Read()
                txt_id.Text = lector(0)
                lector.Close()
                Dim a, b As Integer, c, d As Double

                For i = 0 To dg_datos.RowCount - 2
                    'Obtencion de valores
                    a = Val(txt_id.Text)
                    b = Val(dg_datos.Item(0, i).Value)
                    c = Val(dg_datos.Item(3, i).Value)
                    d = Val(dg_datos.Item(4, i).Value)
                    'Inserta Detalle Compra
                    Try
                        comando.CommandText = "INSERT INTO DETALLEVENTA VALUES (" & a & "," & b & "," & c & "," & d & ")"
                        comando.ExecuteNonQuery()
                    Catch ex As Exception
                        Dim linea As Integer, corrida, formulario, usuario, fecha, hora As String
                        linea = 168
                        corrida = "Error De Insercion de Detalle"
                        'descripcion = ex.Message
                        formulario = "Ventas.vb"
                        usuario = Inicio_Sesion.txt_usuario.Text
                        fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
                        hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
                        insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
                    End Try
                    'Actualiza Articulos
                    Try
                        comando.CommandText = "UPDATE PRODUCTO SET  existencas = existencas - " & c & " WHERE idProducto = " & b & ""
                        comando.ExecuteNonQuery()
                    Catch ex As Exception
                        Dim linea As Integer, corrida, formulario, usuario, fecha, hora As String
                        linea = 171
                        corrida = "Error De Actualizacion"
                        'descripcion = ex.Message
                        formulario = "Venta.vb"
                        usuario = Inicio_Sesion.txt_usuario.Text
                        fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
                        hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
                        insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
                    End Try
                Next
            End Using
            Conexion.Close()
        End Using
        dg_datos.Rows.Clear()
        lbl_total.Text = 0.0
        btn_agregar.Enabled = False
        btn_grabar.Enabled = False
        btn_nuevo.Enabled = True
    End Sub

    Private Sub btn_salir_Click(sender As Object, e As EventArgs) Handles btn_salir.Click
        Dispose()
    End Sub
End Class