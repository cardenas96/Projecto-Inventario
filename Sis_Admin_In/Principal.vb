Imports System.ComponentModel

Public Class Principal
    Private Sub EmpleadosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EmpleadosToolStripMenuItem.Click
        frm_Empleados.ShowDialog()
    End Sub

    Private Sub ProveedoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProveedoresToolStripMenuItem.Click
        Proveedor.ShowDialog()
    End Sub

    Private Sub MaterialToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MaterialToolStripMenuItem.Click
        Producto.ShowDialog()
    End Sub

    Private Sub CompraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CompraToolStripMenuItem.Click
        Compra.ShowDialog()
    End Sub

    Private Sub ConsultaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsultaToolStripMenuItem.Click
        Consulta_Compra.ShowDialog()
    End Sub

    Private Sub VentaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VentaToolStripMenuItem.Click
        Ventas.ShowDialog()
    End Sub

    Private Sub ConsultaToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ConsultaToolStripMenuItem1.Click
        Consulta_Ventas.ShowDialog()
    End Sub

    Private Sub UsuariosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles tsm_usuarios.Click, tsm_usuarios.Click
        Usuarios.ShowDialog()
    End Sub

    Private Sub Principal_Closed(sender As Object, e As EventArgs) Handles Me.Closed
    End Sub

    Private Sub Principal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If tipo.Equals("Administrador") Then
            tsm_usuarios.Visible = True
            UtileriaToolStripMenuItem.Enabled = True
        ElseIf tipo.Equals("Gerente") Then
            UtileriaToolStripMenuItem.Enabled = False
            UtileriaToolStripMenuItem.Enabled = False
            VentaToolStripMenuItem.Enabled = False
            EmpleadosToolStripMenuItem.Enabled = False
        Else
            tsm_usuarios.Visible = False
            UtileriaToolStripMenuItem.Enabled = False
            ExportarAExcelToolStripMenuItem.Enabled = False
            UtileriaToolStripMenuItem.Enabled = False
            ReportesToolStripMenuItem.Enabled = False
            ConsultaToolStripMenuItem1.Enabled = False
            ConsultaToolStripMenuItem.Enabled = False
            CompraToolStripMenuItem.Enabled = False
            MaterialToolStripMenuItem.Enabled = False
            ProveedoresToolStripMenuItem.Enabled = False
            EmpleadosToolStripMenuItem.Enabled = False
        End If

    End Sub

    Private Sub ReportesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportesToolStripMenuItem.Click
        Reportes.ShowDialog()
    End Sub

    Private Sub RespaldoYRestauracionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RespaldoYRestauracionToolStripMenuItem.Click
        Respaldo_Restauracion.ShowDialog()
    End Sub

    Private Sub ErroresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ErroresToolStripMenuItem.Click
        Bitacora.ShowDialog()
    End Sub

    Private Sub Principal_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 Then
            System.Diagnostics.Process.Start("D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Ayuda.chm")
        End If
    End Sub

    Private Sub HistoricoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HistoricoToolStripMenuItem.Click
        Historico.ShowDialog()
    End Sub

    Private Sub VentasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VentasToolStripMenuItem.Click
        frmVentasExcel.ShowDialog()
    End Sub

    Private Sub CerrarSesionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CerrarSesionToolStripMenuItem.Click
        Dispose()
        Inicio_Sesion.Visible = True
        Inicio_Sesion.bandera = True
        Inicio_Sesion.txt_usuario.Text = ""
        Inicio_Sesion.txt_cusuario.Text = ""
    End Sub

    Private Sub ImportarMaterialesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportarMaterialesToolStripMenuItem.Click
        importarTxt()
    End Sub

    Private Sub Principal_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Dispose()
        Inicio_Sesion.Visible = True
        Inicio_Sesion.bandera = True
        Inicio_Sesion.txt_usuario.Text = ""
        Inicio_Sesion.txt_cusuario.Text = ""
    End Sub

End Class

