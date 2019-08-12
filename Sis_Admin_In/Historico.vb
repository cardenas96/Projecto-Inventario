Imports System.Data.SqlClient

Public Class Historico
    Dim conexion As New SqlConnection("Data Source ='CARDENAS-PC'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
    Dim segconexion As New SqlConnection("Data Source ='CARDENAS-PC'; Initial Catalog = 'INVENTARIO_HIST'; Integrated security = true ")
    Dim com As SqlCommand = conexion.CreateCommand
    Dim com2 As SqlCommand = segconexion.CreateCommand
    Dim lector As SqlDataReader

    Private Sub btn_aceptar_Click(sender As Object, e As EventArgs) Handles btn_aceptar.Click
        Dim lista As New List(Of Int32)
        conexion.Open()
        segconexion.Open()
        com.Connection = conexion
        com2.Connection = segconexion

        'INSERTAR CATALOGOS'
        'Insertar Producto'
        com.CommandText = "SELECT * FROM PRODUCTO"
        lector = com.ExecuteReader()
        com2.CommandText = "TRUNCATE TABLE PRODUCTO"
        com2.ExecuteNonQuery()
        While lector.Read()
            com2.CommandText = "INSERT INTO PRODUCTO VALUES(" & lector(0) & ",'" & lector(1) & "','" & lector(2) & "','" & lector(3) & "'," & lector(4) & "," & lector(5) & "," & lector(6) & "," & lector(7) & ")"
            com2.ExecuteNonQuery()
        End While
        lector.Close()

        'Insertar Proveedor'
        com.CommandText = "SELECT * FROM PROVEEDOR"
        lector = com.ExecuteReader()
        com2.CommandText = "TRUNCATE TABLE PROVEEDOR"
        com2.ExecuteNonQuery()
        While lector.Read()
            com2.CommandText = "INSERT INTO PROVEEDOR VALUES(" & lector(0) & ",'" & lector(1) & "','" & lector(2) & "','" & lector(3) & "','" & lector(4) & "','" & lector(5) & "'," & lector(6) & "," & lector(7) & "," & lector(8) & ")"
            com2.ExecuteNonQuery()
        End While
        lector.Close()

        'Insertar Empleado'
        com.CommandText = "SELECT * FROM EMPLEADO"
        lector = com.ExecuteReader()
        com2.CommandText = "TRUNCATE TABLE EMPLEADO"
        com2.ExecuteNonQuery()
        While lector.Read()
            com2.CommandText = "INSERT INTO EMPLEADO VALUES(" & lector(0) & ",'" & lector(1) & "','" & lector(2) & "','" & lector(3) & "','" & lector(4) & "'," & lector(5) & ",'" & lector(6) & "','" & lector(7) & "')"
            com2.ExecuteNonQuery()
        End While
        lector.Close()

        'INSERTAR MOVIMIENTOS'
        com.CommandText = "SELECT * FROM COMPRA WHERE YEAR(fecha) = '" & cb_año.Text & "'"
        lector = com.ExecuteReader()
        'com2.CommandText = "DELETE COMPRA"
        'com2.ExecuteNonQuery()
        While lector.Read()
            com2.CommandText = "INSERT INTO COMPRA VALUES(" & lector(0) & "," & lector(1) & "," & lector(2) & ",'" & lector(3) & "'," & lector(4) & "," & lector(5) & "," & lector(6) & ")"
            com2.ExecuteNonQuery()
            lista.Add(lector(0))
        End While
        lector.Close()
        com.CommandText = "SELECT * FROM DETALLECOMPRA"
        lector = com.ExecuteReader()
        'com2.CommandText = "DELETE DETALLECOMPRA"
        'com2.ExecuteNonQuery()
        While lector.Read()
            If lista.Contains(lector(0)) Then
                com2.CommandText = "INSERT INTO DETALLECOMPRA VALUES(" & lector(0) & "," & lector(1) & "," & lector(2) & "," & lector(3) & ")"
                com2.ExecuteNonQuery()
            End If
        End While
        lector.Close()
        For Each n In lista
            com.CommandText = "DELETE FROM DETALLECOMPRA WHERE idCompra = " & n & ""
            com.ExecuteNonQuery()
            com.CommandText = "DELETE FROM COMPRA WHERE idCompra = " & n & ""
            com.ExecuteNonQuery()

        Next

        com.CommandText = "SELECT * FROM VENTA WHERE YEAR(fecha) = '" & cb_año.Text & "'"
        lector = com.ExecuteReader()
        'com2.CommandText = "DELETE VENTA"
        'com2.ExecuteNonQuery()
        While lector.Read()
            com2.CommandText = "INSERT INTO VENTA VALUES(" & lector(0) & "," & lector(1) & ",'" & lector(2) & "'," & lector(3) & ")"
            com2.ExecuteNonQuery()
            lista.Add(lector(0))
        End While
        lector.Close()
        com.CommandText = "SELECT * FROM DETALLEVENTA"
        lector = com.ExecuteReader()
        ' com2.CommandText = "DELETE DETALLEVENTA"
        'com2.ExecuteNonQuery()
        While lector.Read()
            If lista.Contains(lector(0)) Then
                com2.CommandText = "INSERT INTO DETALLEVENTA VALUES(" & lector(0) & "," & lector(1) & "," & lector(2) & "," & lector(3) & ")"
                com2.ExecuteNonQuery()
            End If
        End While
        lector.Close()
        For Each n In lista
            com.CommandText = "DELETE FROM DETALLEVENTA WHERE idVenta = " & n & ""
            com.ExecuteNonQuery()
            com.CommandText = "DELETE FROM VENTA WHERE idVenta = " & n & ""
            com.ExecuteNonQuery()
        Next
        MessageBox.Show("Se Ah trasado a Historicos Los Datos Del Año" & cb_año.Text, "¡Historico Completo!", MessageBoxButtons.OK, MessageBoxIcon.Information)
        conexion.Close()
        segconexion.Close()
    End Sub
End Class