Imports System.ComponentModel
Imports System.Data.SqlClient
Imports Microsoft.Reporting.WinForms
Public Class Reportes
    Private Sub Reportes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ReportViewer1.LocalReport.DataSources.Clear()
        Me.ReportViewer1.RefreshReport()
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
    Private Sub EmpleadosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EmpleadosToolStripMenuItem.Click
        Dim conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
        Dim comando As New SqlCommand("REPORTEMPLEADO", conexionsql)
        comando.CommandType = CommandType.StoredProcedure
        Dim adaptador As New SqlDataAdapter(comando)
        Dim data As New DataSet
        Try
            adaptador.Fill(data)
            data.DataSetName = "DataSet1"
            Dim reportes As New ReportDataSource("DataSet1", data.Tables(0))
            Dim p1 As New ReportParameter("P1", dtp_fecha.Text)
            ReportViewer1.LocalReport.DataSources.Clear()
            ReportViewer1.LocalReport.DataSources.Add(reportes)
            ReportViewer1.LocalReport.ReportPath = "D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Sis_Admin_In\Reporte_Empleados.rdlc"
            ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p1})
        Catch ex As Exception
            Dim linea As Integer, corrida, formulario, usuario, fecha, hora As String
            linea = 32
            corrida = "Error En Reportes"
            'descripcion = ex.Message
            formulario = "Reportes.vb"
            usuario = Inicio_Sesion.txt_usuario.Text
            fecha = Today.Date 'Today.Day & "/" & Today.Month & "/" & Today.Year
            hora = TimeOfDay.Hour & ":" & TimeOfDay.Minute & ":" & TimeOfDay.Second
            insertarError(linea, corrida, ex, formulario, usuario, fecha, hora)
            MessageBox.Show(ex.Message, "Error Tenemos Un Problema", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
        ReportViewer1.RefreshReport()
        conexionsql.Close()
    End Sub

    Private Sub ProveedoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProveedoresToolStripMenuItem.Click
        Dim conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
        Dim comando As New SqlCommand("REPORTEPROVEEDOR", conexionsql)
        comando.CommandType = CommandType.StoredProcedure
        Dim adaptador As New SqlDataAdapter(comando)
        Dim data As New DataSet
        adaptador.Fill(data)
        data.DataSetName = "DataSet1"
        Dim reportes As New ReportDataSource("DataSet1", data.Tables(0))
        Dim p1 As New ReportParameter("P1", dtp_fecha.Text)
        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.DataSources.Add(reportes)
        ReportViewer1.LocalReport.ReportPath = "D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Sis_Admin_In\Reporte_Proveedor.rdlc"
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p1})
        ReportViewer1.RefreshReport()
        conexionsql.Close()
    End Sub

    Private Sub ProductoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductoToolStripMenuItem.Click
        Dim conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
        Dim comando As New SqlCommand("REPORTEPRODUCTO", conexionsql)
        comando.CommandType = CommandType.StoredProcedure
        Dim adaptador As New SqlDataAdapter(comando)
        Dim data As New DataSet
        adaptador.Fill(data)
        data.DataSetName = "DataSet1"
        Dim reportes As New ReportDataSource("DataSet1", data.Tables(0))
        Dim p1 As New ReportParameter("P1", dtp_fecha.Text)
        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.DataSources.Add(reportes)
        ReportViewer1.LocalReport.ReportPath = "D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Sis_Admin_In\Reporte_Producto.rdlc"
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p1})
        ReportViewer1.RefreshReport()
        conexionsql.Close()
    End Sub

    Private Sub VentasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VentasToolStripMenuItem.Click
        Fechas.ShowDialog()
        Dim f1, f2 As String
        Dim conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
        f1 = lbl_f1.Text
        f2 = lbl_f2.Text
        Dim cmd As New SqlCommand("FECHAVENTA", conexionsql)
        cmd.CommandType = CommandType.StoredProcedure
        Dim adaptador As New SqlDataAdapter()
        adaptador.SelectCommand = New SqlCommand
        adaptador.SelectCommand.Connection = conexionsql
        adaptador.SelectCommand.CommandText = "FECHAVENTA"
        adaptador.SelectCommand.CommandType = CommandType.StoredProcedure
        Dim param1 = New SqlParameter("@FECHA1", SqlDbType.Date)
        Dim param2 = New SqlParameter("@FECHA2", SqlDbType.Date)
        param1.Direction = ParameterDirection.Input
        param2.Direction = ParameterDirection.Input
        param1.Value = CDate(f1)
        param2.Value = CDate(f2)
        adaptador.SelectCommand.Parameters.Add(param1)
        adaptador.SelectCommand.Parameters.Add(param2)
        Dim dataset As New DataSet
        adaptador.Fill(dataset)
        dataset.DataSetName = "DataSet1"
        Dim datasource As New ReportDataSource("DataSet1", dataset.Tables(0))
        datasource.Name = "DataSet1"
        datasource.Value = dataset.Tables(0)

        Dim p1 As New ReportParameter("P1", f1)

        Dim p2 As New ReportParameter("P2", f2)

        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.DataSources.Add(datasource)
        ReportViewer1.LocalReport.ReportPath = "D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Sis_Admin_In\Reporte_Fecha_Venta.rdlc"
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p1, p2})

        ReportViewer1.RefreshReport()
        conexionsql.Close()
    End Sub

    Private Sub VentaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VentaToolStripMenuItem.Click
        Dim conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
        Dim comando As New SqlCommand("REPORTEVENTAS", conexionsql)
        comando.CommandType = CommandType.StoredProcedure
        Dim adaptador As New SqlDataAdapter(comando)
        Dim data As New DataSet
        adaptador.Fill(data)
        data.DataSetName = "DataSet1"
        Dim reportes As New ReportDataSource("DataSet1", data.Tables(0))
        Dim p1 As New ReportParameter("P1", dtp_fecha.Text)
        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.DataSources.Add(reportes)
        ReportViewer1.LocalReport.ReportPath = "D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Sis_Admin_In\Reporte_Ventas.rdlc"
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p1})
        ReportViewer1.RefreshReport()
        conexionsql.Close()
    End Sub

    Private Sub ProductosFaltantesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductosFaltantesToolStripMenuItem.Click
        Dim conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
        Dim comando As New SqlCommand("PRODUCTOSFALTANTES", conexionsql)
        comando.CommandType = CommandType.StoredProcedure
        Dim adaptador As New SqlDataAdapter(comando)
        Dim data As New DataSet
        adaptador.Fill(data)
        data.DataSetName = "DataSet1"
        Dim reportes As New ReportDataSource("DataSet1", data.Tables(0))
        Dim p1 As New ReportParameter("P1", dtp_fecha.Text)
        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.DataSources.Add(reportes)
        ReportViewer1.LocalReport.ReportPath = "D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Sis_Admin_In\Reporte_Faltantes.rdlc"
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p1})
        ReportViewer1.RefreshReport()
        conexionsql.Close()
    End Sub

    Private Sub CompraDetalleProductoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CompraDetalleProductoToolStripMenuItem.Click
        Dim conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
        Dim comando As New SqlCommand("REPORTECOMPRAPRODUCTO", conexionsql)
        comando.CommandType = CommandType.StoredProcedure
        Dim adaptador As New SqlDataAdapter(comando)
        Dim data As New DataSet
        adaptador.Fill(data)
        data.DataSetName = "DataSet1"
        Dim reportes As New ReportDataSource("DataSet1", data.Tables(0))
        Dim p1 As New ReportParameter("P1", dtp_fecha.Text)
        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.DataSources.Add(reportes)
        ReportViewer1.LocalReport.ReportPath = "D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Sis_Admin_In\Reporte_Compra_Producto.rdlc"
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p1})
        ReportViewer1.RefreshReport()
        conexionsql.Close()
    End Sub

    Private Sub ComprasRealizadasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ComprasRealizadasToolStripMenuItem.Click
        Dim conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
        Dim comando As New SqlCommand("REPORTECOMPRA", conexionsql)
        comando.CommandType = CommandType.StoredProcedure
        Dim adaptador As New SqlDataAdapter(comando)
        Dim data As New DataSet
        adaptador.Fill(data)
        data.DataSetName = "DataSet1"
        Dim reportes As New ReportDataSource("DataSet1", data.Tables(0))
        Dim p1 As New ReportParameter("P1", dtp_fecha.Text)
        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.DataSources.Add(reportes)
        ReportViewer1.LocalReport.ReportPath = "D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Sis_Admin_In\Reporte_Compra.rdlc"
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p1})
        ReportViewer1.RefreshReport()
        conexionsql.Close()
    End Sub

    Private Sub ProductosCompradosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductosCompradosToolStripMenuItem.Click
        Fechas.ShowDialog()
        Dim f1, f2 As String
        Dim conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
        f1 = lbl_f1.Text
        f2 = lbl_f2.Text
        Dim cmd As New SqlCommand("FECHAREPORTECOMPRAPRODUCTO", conexionsql)
        cmd.CommandType = CommandType.StoredProcedure
        Dim adaptador As New SqlDataAdapter()
        adaptador.SelectCommand = New SqlCommand
        adaptador.SelectCommand.Connection = conexionsql
        adaptador.SelectCommand.CommandText = "FECHAREPORTECOMPRAPRODUCTO"
        adaptador.SelectCommand.CommandType = CommandType.StoredProcedure
        Dim param1 = New SqlParameter("@FECHA1", SqlDbType.Date)
        Dim param2 = New SqlParameter("@FECHA2", SqlDbType.Date)
        param1.Direction = ParameterDirection.Input
        param2.Direction = ParameterDirection.Input
        param1.Value = CDate(f1)
        param2.Value = CDate(f2)
        adaptador.SelectCommand.Parameters.Add(param1)
        adaptador.SelectCommand.Parameters.Add(param2)
        Dim dataset As New DataSet
        adaptador.Fill(dataset)
        dataset.DataSetName = "DataSet1"
        Dim datasource As New ReportDataSource("DataSet1", dataset.Tables(0))
        datasource.Name = "DataSet1"
        datasource.Value = dataset.Tables(0)

        Dim p1 As New ReportParameter("P1", f1)

        Dim p2 As New ReportParameter("P2", f2)

        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.DataSources.Add(datasource)
        ReportViewer1.LocalReport.ReportPath = "D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Sis_Admin_In\Reporte_Com_Prod_Fech.rdlc"
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p1, p2})

        ReportViewer1.RefreshReport()
        conexionsql.Close()
    End Sub

    Private Sub ComprasRealizadasToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ComprasRealizadasToolStripMenuItem1.Click
        Fechas.ShowDialog()
        Dim f1, f2 As String
        Dim conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
        f1 = lbl_f1.Text
        f2 = lbl_f2.Text
        Dim cmd As New SqlCommand("FECHAREPORTECOMPRA", conexionsql)
        cmd.CommandType = CommandType.StoredProcedure
        Dim adaptador As New SqlDataAdapter()
        adaptador.SelectCommand = New SqlCommand
        adaptador.SelectCommand.Connection = conexionsql
        adaptador.SelectCommand.CommandText = "FECHAREPORTECOMPRA"
        adaptador.SelectCommand.CommandType = CommandType.StoredProcedure
        Dim param1 = New SqlParameter("@FECHA1", SqlDbType.Date)
        Dim param2 = New SqlParameter("@FECHA2", SqlDbType.Date)
        param1.Direction = ParameterDirection.Input
        param2.Direction = ParameterDirection.Input
        param1.Value = CDate(f1)
        param2.Value = CDate(f2)
        adaptador.SelectCommand.Parameters.Add(param1)
        adaptador.SelectCommand.Parameters.Add(param2)
        Dim dataset As New DataSet
        adaptador.Fill(dataset)
        dataset.DataSetName = "DataSet1"
        Dim datasource As New ReportDataSource("DataSet1", dataset.Tables(0))
        datasource.Name = "DataSet1"
        datasource.Value = dataset.Tables(0)

        Dim p1 As New ReportParameter("P1", f1)

        Dim p2 As New ReportParameter("P2", f2)

        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.DataSources.Add(datasource)
        ReportViewer1.LocalReport.ReportPath = "D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Sis_Admin_In\Reporte_Fecha_Compra.rdlc"
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p1, p2})

        ReportViewer1.RefreshReport()
        conexionsql.Close()
    End Sub

    Private Sub VentasPorEmpleadoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VentasPorEmpleadoToolStripMenuItem.Click
        lbl_f1.Text = "1"
        Fecha_Nombre.gb_nombre.Text = "Seleccione Empleado"
        Fecha_Nombre.ShowDialog()
        Dim f1, f2, n1 As String
        Dim conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
        f1 = lbl_f1.Text
        f2 = lbl_f2.Text
        n1 = lbl_n1.Text
        Dim cmd As New SqlCommand("VENTASPOREMPLEADOFECHA", conexionsql)
        cmd.CommandType = CommandType.StoredProcedure
        Dim adaptador As New SqlDataAdapter()
        adaptador.SelectCommand = New SqlCommand
        adaptador.SelectCommand.Connection = conexionsql
        adaptador.SelectCommand.CommandText = "VENTASPOREMPLEADOFECHA"
        adaptador.SelectCommand.CommandType = CommandType.StoredProcedure
        Dim param1 = New SqlParameter("@EMPLEADO", SqlDbType.VarChar)
        Dim param2 = New SqlParameter("@FECHA1", SqlDbType.Date)
        Dim param3 = New SqlParameter("@FECHA2", SqlDbType.Date)
        param1.Direction = ParameterDirection.Input
        param2.Direction = ParameterDirection.Input
        param3.Direction = ParameterDirection.Input
        param1.Value = n1
        param2.Value = CDate(f1)
        param3.Value = CDate(f2)
        adaptador.SelectCommand.Parameters.Add(param1)
        adaptador.SelectCommand.Parameters.Add(param2)
        adaptador.SelectCommand.Parameters.Add(param3)
        Dim dataset As New DataSet
        adaptador.Fill(dataset)
        dataset.DataSetName = "DataSet1"
        Dim datasource As New ReportDataSource("DataSet1", dataset.Tables(0))
        datasource.Name = "DataSet1"
        datasource.Value = dataset.Tables(0)

        Dim p1 As New ReportParameter("P1", f1)

        Dim p2 As New ReportParameter("P2", f2)

        Dim p3 As New ReportParameter("P3", n1)

        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.DataSources.Add(datasource)
        ReportViewer1.LocalReport.ReportPath = "D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Sis_Admin_In\Reporte_Venta_Empleado.rdlc"
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p1, p2, p3})

        ReportViewer1.RefreshReport()
        conexionsql.Close()
    End Sub

    Private Sub ComprasPorProveedorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ComprasPorProveedorToolStripMenuItem.Click
        lbl_f1.Text = "2"
        Fecha_Nombre.gb_nombre.Text = "Seleccione Proveedor"
        Fecha_Nombre.ShowDialog()
        Dim f1, f2, n1 As String
        Dim conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
        f1 = lbl_f1.Text
        f2 = lbl_f2.Text
        n1 = lbl_n1.Text
        Dim cmd As New SqlCommand("COMPRASPORPROVEEDORFECHA", conexionsql)
        cmd.CommandType = CommandType.StoredProcedure
        Dim adaptador As New SqlDataAdapter()
        adaptador.SelectCommand = New SqlCommand
        adaptador.SelectCommand.Connection = conexionsql
        adaptador.SelectCommand.CommandText = "COMPRASPORPROVEEDORFECHA"
        adaptador.SelectCommand.CommandType = CommandType.StoredProcedure
        Dim param1 = New SqlParameter("@PROVEEDOR", SqlDbType.VarChar)
        Dim param2 = New SqlParameter("@FECHA1", SqlDbType.Date)
        Dim param3 = New SqlParameter("@FECHA2", SqlDbType.Date)
        param1.Direction = ParameterDirection.Input
        param2.Direction = ParameterDirection.Input
        param3.Direction = ParameterDirection.Input
        param1.Value = n1
        param2.Value = CDate(f1)
        param3.Value = CDate(f2)
        adaptador.SelectCommand.Parameters.Add(param1)
        adaptador.SelectCommand.Parameters.Add(param2)
        adaptador.SelectCommand.Parameters.Add(param3)
        Dim dataset As New DataSet
        adaptador.Fill(dataset)
        dataset.DataSetName = "DataSet1"
        Dim datasource As New ReportDataSource("DataSet1", dataset.Tables(0))
        datasource.Name = "DataSet1"
        datasource.Value = dataset.Tables(0)

        Dim p1 As New ReportParameter("P1", f1)

        Dim p2 As New ReportParameter("P2", f2)

        Dim p3 As New ReportParameter("P3", n1)

        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.DataSources.Add(datasource)
        ReportViewer1.LocalReport.ReportPath = "D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Sis_Admin_In\Reporte_Compra_Proveedor.rdlc"
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p1, p2, p3})

        ReportViewer1.RefreshReport()
        conexionsql.Close()
    End Sub

    Private Sub Reportes_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 Then
            System.Diagnostics.Process.Start("D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Ayuda.chm")
        End If
    End Sub

    Private Sub ProductosMasVendidosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductosMasVendidosToolStripMenuItem.Click
        Dim conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
        Dim comando As New SqlCommand("PRODUCTOMASVENDIDO", conexionsql)
        comando.CommandType = CommandType.StoredProcedure
        Dim adaptador As New SqlDataAdapter(comando)
        Dim data As New DataSet
        adaptador.Fill(data)
        data.DataSetName = "DataSet1"
        Dim reportes As New ReportDataSource("DataSet1", data.Tables(0))
        Dim p1 As New ReportParameter("P1", dtp_fecha.Text)
        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.DataSources.Add(reportes)
        ReportViewer1.LocalReport.ReportPath = "C:\PruebasPH\ProyectoHope\Sis_Admin_In\Sis_Admin_In\Producto_Mas_Vendido.rdlc"
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p1})
        ReportViewer1.RefreshReport()
        conexionsql.Close()
    End Sub

    Private Sub Reportes_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Dispose()
        conexionsql.Close()
    End Sub

    Private Sub ProductoMasVendidosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductoMasVendidosToolStripMenuItem.Click
        Fechas.ShowDialog()
        Dim f1, f2 As String
        Dim conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
        f1 = lbl_f1.Text
        f2 = lbl_f2.Text
        Dim cmd As New SqlCommand("PRODUCTOMASVENDIDOFECHA", conexionsql)
        cmd.CommandType = CommandType.StoredProcedure
        Dim adaptador As New SqlDataAdapter()
        adaptador.SelectCommand = New SqlCommand
        adaptador.SelectCommand.Connection = conexionsql
        adaptador.SelectCommand.CommandText = "PRODUCTOMASVENDIDOFECHA"
        adaptador.SelectCommand.CommandType = CommandType.StoredProcedure
        Dim param1 = New SqlParameter("@FECHA1", SqlDbType.Date)
        Dim param2 = New SqlParameter("@FECHA2", SqlDbType.Date)
        param1.Direction = ParameterDirection.Input
        param2.Direction = ParameterDirection.Input
        param1.Value = CDate(f1)
        param2.Value = CDate(f2)
        adaptador.SelectCommand.Parameters.Add(param1)
        adaptador.SelectCommand.Parameters.Add(param2)
        Dim dataset As New DataSet
        adaptador.Fill(dataset)
        dataset.DataSetName = "DataSet1"
        Dim datasource As New ReportDataSource("DataSet1", dataset.Tables(0))
        datasource.Name = "DataSet1"
        datasource.Value = dataset.Tables(0)

        Dim p1 As New ReportParameter("P1", f1)

        Dim p2 As New ReportParameter("P2", f2)

        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.DataSources.Add(datasource)
        ReportViewer1.LocalReport.ReportPath = "C:\PruebasPH\ProyectoHope\Sis_Admin_In\Sis_Admin_In\Reporte_Producto_Mas_Vendido_Fecha.rdlc"
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p1, p2})

        ReportViewer1.RefreshReport()
        conexionsql.Close()
    End Sub
End Class