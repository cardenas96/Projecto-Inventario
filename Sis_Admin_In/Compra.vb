Imports System.Data.SqlClient
Public Class Compra
    Dim conexionsql As New SqlConnection("Data Source ='.'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
    Dim comando As SqlCommand = conexionsql.CreateCommand
    Dim lector As SqlDataReader
    Dim transaccion As SqlTransaction
    Sub limpiar()
        txt_idcompra.Text = ""
        txt_folio.Text = ""
        txt_subtotal.Text = ""
        txt_iva.Text = ""
        txt_total.Text = ""
        'Campos de Proveedor
        txt_idp.Text = ""
        cb_proveedor.Text = ""
        txt_contacto.Text = ""
        txt_telefono.Text = ""
        txt_fax.Text = ""
        txt_ciudad.Text = ""
        txt_colonia.Text = ""
        txt_direccion.Text = ""

        'Campos de Articulo
        txt_ida.Text = ""
        cb_articulo.Text = ""
        txt_marca.Text = ""
        txt_unidad.Text = ""
        txt_existencias.Text = ""
        txt_costo_a.Text = ""
        txt_min.Text = ""
        txt_max.Text = ""
        txt_costo.Text = ""
        txt_cantidad.Text = ""
    End Sub

    Private Sub Compra_Load(sender As Object, e As EventArgs) Handles Me.Load
        conexionsql.Open()
        'CARGAR COMBO PROVEEDOR
        Try
            comando.CommandText = "SELECT nombre FROM PROVEEDOR"
            lector = comando.ExecuteReader
            While lector.Read
                cb_proveedor.Items.Add(lector(0))
            End While
        Catch ex As Exception
            Dim linea As Integer, corrida, formulario, usuario, fecha, hora As String
            linea = 40
            corrida = "Error Cargar Proveedor"
            'descripcion = ex.Message
            formulario = "Compra.vb"
            usuario = Inicio_Sesion.txt_usuario.Text
            fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
            hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
            insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
            MessageBox.Show(ex.Message, "Tenemos Un Problema", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
        lector.Close()

        'CARGAR COMO ARTICULOS
        Try
            comando.CommandText = "SELECT nombre FROM PRODUCTO"
            lector = comando.ExecuteReader
            While lector.Read
                cb_articulo.Items.Add(lector(0))
            End While
        Catch ex As Exception
            Dim linea As Integer, corrida, formulario, usuario, fecha, hora As String
            linea = 61
            corrida = "Error Cargar Producto"
            'descripcion = ex.Message
            formulario = "Compra.vb"
            usuario = Inicio_Sesion.txt_usuario.Text
            fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
            hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
            insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
            MessageBox.Show(ex.Message, "Tenemos Un Problema", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
        lector.Close()
    End Sub

    Private Sub cb_proveedor_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_proveedor.SelectedIndexChanged
        Try
            comando.CommandText = "SELECT * FROM PROVEEDOR WHERE nombre = '" & cb_proveedor.Text & "'"
            lector = comando.ExecuteReader
        Catch ex As Exception
            Dim linea As Integer, corrida, formulario, usuario, fecha, hora As String
            linea = 84
            corrida = "Error Cargar Combo Proveedor"
            ' descripcion = ex.Message
            formulario = "Compra.vb"
            usuario = Inicio_Sesion.txt_usuario.Text
            fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
            hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
            insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
            MessageBox.Show(ex.Message, "Error Tenemos Un Problema", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
        lector.Read()
        txt_idp.Text = lector(0)
        txt_contacto.Text = lector(2)
        txt_telefono.Text = lector(6)
        txt_fax.Text = lector(7)
        txt_ciudad.Text = lector(4)
        txt_colonia.Text = lector(5)
        txt_direccion.Text = lector(3)
        lector.Close()
    End Sub

    Private Sub cb_articulo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_articulo.SelectedIndexChanged
        Try
            comando.CommandText = "SELECT * FROM PRODUCTO WHERE nombre = '" & cb_articulo.Text & "'"
            lector = comando.ExecuteReader
        Catch ex As Exception
            Dim linea As Integer, corrida, formulario, usuario, fecha, hora As String
            linea = 110
            corrida = "Error Cargar Combo Producto"
            'descripcion = ex.Message
            formulario = "Compra.vb"
            usuario = Inicio_Sesion.txt_usuario.Text
            fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
            hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
            insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
            MessageBox.Show(ex.Message, "Error Tenemos Un Problema", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
        lector.Read()
        txt_ida.Text = lector(0)
        txt_marca.Text = lector(2)
        txt_unidad.Text = lector(3)
        txt_existencias.Text = lector(5)
        txt_costo_a.Text = lector(4)
        txt_min.Text = lector(6)
        txt_max.Text = lector(7)
        lector.Close()

    End Sub

    Private Sub btn_nuevo_Click(sender As Object, e As EventArgs) Handles btn_nuevo.Click
        Dim n As Integer
        txt_folio.Enabled = True
        btn_nuevo.Enabled = False
        btn_grabar.Enabled = True
        txt_costo.Enabled = True
        txt_cantidad.Enabled = True
        Try
            comando.CommandText = "SELECT COUNT(*) FROM COMPRA"
            n = comando.ExecuteScalar() + 1
        Catch ex As Exception
            Dim linea As Integer, corrida, descripcion, formulario, usuario, fecha, hora As String
            linea = 144
            corrida = "Error Contador De Compra"
            descripcion = ex.Message
            formulario = "Compra.vb"
            usuario = Inicio_Sesion.txt_usuario.Text
            fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
            hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
            insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
            MessageBox.Show(descripcion, "Error Tenemos Un Problema", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
        txt_idcompra.Text = n
        txt_subtotal.Text = "0"
        txt_total.Text = "0"
        btn_agregar.Enabled = True
    End Sub

    Private Sub Compra_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        conexionsql.Close()
    End Sub

    Private Sub cb_articulo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cb_articulo.KeyPress
        e.KeyChar = Chr(0)
    End Sub

    Private Sub cb_proveedor_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cb_proveedor.KeyPress
        e.KeyChar = Chr(0)
    End Sub

    Private Sub btn_agregar_Click(sender As Object, e As EventArgs) Handles btn_agregar.Click
        If (txt_cantidad.Text = "" Or txt_costo.Text = "" Or cb_articulo.Text = "") Then
            MessageBox.Show("Error", "Faltan Campos", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            Dim ban As Boolean = False
            For i = 0 To dg_datos.RowCount - 2
                If Val(txt_ida.Text) = Val(dg_datos.Item(6, i).Value) Then
                    ban = True
                End If
            Next
            If ban = False Then
                Dim c1, c5 As Integer, c2 As String, c3, c4 As Double
                Try
                    c1 = Val(txt_idcompra.Text)
                    c2 = cb_articulo.Text
                    c3 = CDbl(txt_cantidad.Text)
                    c4 = CDbl(txt_costo.Text)
                    c5 = Val(txt_ida.Text)
                    'Agrega A Regilla
                    dg_datos.Rows.Add(c1, c2, txt_unidad.Text, c3, c4, CDbl(c3 * c4), c5)
                Catch ex As Exception
                    Dim linea As Integer, corrida, descripcion, formulario, usuario, fecha, hora As String
                    linea = 190
                    corrida = "Error Al En Inserción de Datos"
                    descripcion = ex.Message
                    formulario = "Compra.vb"
                    usuario = Inicio_Sesion.txt_usuario.Text
                    fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
                    hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
                    insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
                    MessageBox.Show(descripcion, "Error De Inserción", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Try

            Else
                For i = 0 To dg_datos.RowCount - 2
                    If Val(txt_ida.Text) = Val(dg_datos.Item(6, i).Value) Then
                        dg_datos.Item(3, i).Value = Val(dg_datos.Item(3, i).Value) + Val(txt_cantidad.Text)
                        'lbl_total.Text = Val(lbl_total.Text) + (Val(txt_cantidad.Text) * Val(dg_datos.Item(4, i).Value))
                        dg_datos.Item(5, i).Value = (Val(dg_datos.Item(3, i).Value) * Val(dg_datos.Item(4, i).Value))
                    End If
                Next
            End If

            'Campos de Articulo
            txt_ida.Text = ""
            cb_articulo.Text = ""
            txt_marca.Text = ""
            txt_unidad.Text = ""
            txt_existencias.Text = ""
            txt_costo_a.Text = ""
            txt_min.Text = ""
            txt_max.Text = ""
            txt_costo.Text = ""
            txt_cantidad.Text = ""
            Calcular()
        End If
    End Sub
    Sub Calcular()
        Dim i As Integer, subto As Double = 0
        For i = 0 To dg_datos.Rows.Count - 2
            subto = subto + Format(CDbl(dg_datos.Item(5, i).Value), "0000.00")
        Next
        txt_subtotal.Text = subto.ToString
        txt_iva.Text = Format(CDec(txt_subtotal.Text) * CDec(0.16), "0.00")
        txt_total.Text = Format(CDbl(txt_subtotal.Text) + CDbl(txt_iva.Text), "0.00")
    End Sub

    Private Sub btn_grabar_Click(sender As Object, e As EventArgs) Handles btn_grabar.Click
        If cb_proveedor.Text = "" Or txt_folio.Text = "" Then
            MessageBox.Show("Los Campos Folio O Proveedor Estan Vacios Revisar", "Faltas Campos Por Llenar", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            transaccion = conexionsql.BeginTransaction("COMPRA")
            comando.Transaction = transaccion
            Try
                comando.CommandText = "INSERT INTO COMPRA VALUES (" & Val(txt_idcompra.Text) & "," & Val(txt_idp.Text) & "," & Val(txt_folio.Text) & ",'" & dtp_fecha.Value.Date & "'," & Val(txt_subtotal.Text) & "," & CDbl(txt_iva.Text) & "," & CDec(txt_total.Text) & ")"
                comando.ExecuteNonQuery()
            Catch ex As Exception
                Dim linea As Integer, corrida, descripcion, formulario, usuario, fecha, hora As String
                linea = 250
                corrida = "Error Al Insertar Compra"
                descripcion = ex.Message
                formulario = "Compra.vb"
                usuario = Inicio_Sesion.txt_usuario.Text
                fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
                hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
                insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
                MessageBox.Show(descripcion, "Error Tenemos Un Problema", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Try
            'transaccion.Commit()
            'transaccion.Dispose()
            Dim i, a, b As Integer, c, d As Double
            For i = 0 To dg_datos.RowCount - 2
                ' transaccion = conexionsql.BeginTransaction("DETALLECOMPRA")
                'comando.Transaction = transaccion
                a = Val(dg_datos.Item(0, i).Value)
                b = Val(dg_datos.Item(6, i).Value)
                c = Val(dg_datos.Item(3, i).Value)
                d = Val(dg_datos.Item(4, i).Value)
                'Inserta Detalle Compra
                Try
                    comando.CommandText = "INSERT INTO DETALLECOMPRA VALUES (" & a & "," & b & "," & c & "," & d & ")"
                    comando.ExecuteNonQuery()
                Catch ex As Exception
                    Dim linea As Integer, corrida, descripcion, formulario, usuario, fecha, hora As String
                    linea = 276
                    corrida = "Error Al Insertar Detalle Compra"
                    descripcion = ex.Message
                    formulario = "Compra.vb"
                    usuario = Inicio_Sesion.txt_usuario.Text
                    fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
                    hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
                    insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
                    MessageBox.Show(descripcion, "Error Tenemos Un Problema", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Try
                'Actualiza Articulos
                Try
                    comando.CommandText = "UPDATE PRODUCTO SET costo = " & d & ", existencas = existencas + " & c & " WHERE idProducto = " & b & ""
                    comando.ExecuteNonQuery()
                Catch ex As Exception
                    Dim linea As Integer, corrida, descripcion, formulario, usuario, fecha, hora As String
                    linea = 291
                    corrida = "Error Al Actualizar Producto"
                    descripcion = ex.Message
                    formulario = "Compra.vb"
                    usuario = Inicio_Sesion.txt_usuario.Text
                    fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
                    hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
                    insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
                    MessageBox.Show(descripcion, "Error Tenemos Un Problema", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Try
            Next
            If MsgBox("Desea Ejecutar Transaccion", MsgBoxStyle.YesNo, "Ejecutar") = MsgBoxResult.Yes Then
                Try
                    transaccion.Commit()
                    transaccion.Dispose()
                Catch ex As Exception
                    Dim linea As Integer, corrida, descripcion, formulario, usuario, fecha, hora As String
                    linea = 308
                    corrida = "Error En Trasaccion"
                    descripcion = ex.Message
                    formulario = "Compra.vb"
                    usuario = Inicio_Sesion.txt_usuario.Text
                    fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
                    hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
                    insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
                    MessageBox.Show(descripcion, "Error Tenemos Un Problema", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Try
                MessageBox.Show("Los Campos Se Han Almacenado Correctamente", "Información Grabada", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                Try
                    transaccion.Rollback()
                Catch ex As Exception
                    Dim linea As Integer, corrida, formulario, usuario, fecha, hora As String
                    linea = 325
                    corrida = "Error En Trasaccion"
                    'descripcion = ex.Message
                    formulario = "Compra.vb"
                    usuario = Inicio_Sesion.txt_usuario.Text
                    fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
                    hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
                    insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
                End Try
                MsgBox("Transaccion Cancelada")
            End If
            limpiar()
            dg_datos.Rows.Clear()
            btn_nuevo.Enabled = True
            btn_grabar.Enabled = False
            btn_agregar.Enabled = False
        End If

    End Sub

    Private Sub Compra_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        conexionsql.Close()
    End Sub

    Private Sub btn_salir_Click(sender As Object, e As EventArgs) Handles btn_salir.Click
        Dispose()
        conexionsql.Close()
    End Sub

    Private Sub txt_costo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_costo.KeyPress

        Dim cadenaValida As String = "1234567890."
        e.KeyChar = UCase(e.KeyChar)
        If e.KeyChar > ChrW(26) Then
            If InStr(cadenaValida, e.KeyChar) = 0 Then
                e.KeyChar = ChrW(0)
            End If
        End If
    End Sub

    Private Sub txt_cantidad_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_cantidad.KeyPress
        Dim cadenaValida As String = "1234567890"
        e.KeyChar = UCase(e.KeyChar)
        If e.KeyChar > ChrW(26) Then
            If InStr(cadenaValida, e.KeyChar) = 0 Then
                e.KeyChar = ChrW(0)
            End If
        End If
    End Sub

    Sub insertarError(linea As Integer, corrida As String, ex As Exception, formulario As String, usuario As String, fecha As String, hora As String)
        Dim con As New SqlConnection("Data Source ='.'; Initial Catalog = 'BITACORA_DB'; Integrated security = true")
        con.Open()
        Dim com As New SqlCommand
        com.Connection = con
                com.CommandText = "SELECT COUNT(*) FROM Bitacora"
                Dim no As Integer = com.ExecuteScalar() + 1
        ' Dim cont As Integer = 0, cadena As String
        ' cadena = ex.Message
        ' While cadena.Contains("'")
        'Dim x As Integer = cadena.IndexOf("'")
        '    cadena = cadena.Remove(x, 1)
        'End While

        Try
            'Dim st As New StackTrace(True)
            'st = New StackTrace(ex, True)
            'Dim frame = st.GetFrame(st.FrameCount - 1)
            com.CommandText = "INSERT INTO Bitacora VALUES(" & no & ",'" & corrida & "','" & ex.Message.Replace("'", " ") & "','" & formulario & "'," & Val(linea) & ",'" & usuario & "','" & fecha & "','" & hora & "') "
            ' com.Parameters.Add("@Descripcion", SqlDbType.VarChar).Value = ex.Message
            com.ExecuteNonQuery()

        Catch e As Exception
            MsgBox(e.Message, MsgBoxStyle.MsgBoxHelp, "Error De Bitacora")
        End Try
        con.Close()
    End Sub

    Private Sub Compra_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 Then
            System.Diagnostics.Process.Start("D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Ayuda.chm")
        End If
    End Sub
End Class