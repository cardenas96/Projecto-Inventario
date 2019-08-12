<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Principal
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.ms_menu = New System.Windows.Forms.MenuStrip()
        Me.CatalogosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EmpleadosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProveedoresToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MaterialToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportarMaterialesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsm_usuarios = New System.Windows.Forms.ToolStripMenuItem()
        Me.MovimientosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CompraToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ConsultaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VentaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ConsultaToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UtileriaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RespaldoYRestauracionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ErroresToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HistoricoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExportarAExcelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VentasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CerrarSesionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.lbl_portada = New System.Windows.Forms.Label()
        Me.ms_menu.SuspendLayout()
        Me.SuspendLayout()
        '
        'ms_menu
        '
        Me.ms_menu.Font = New System.Drawing.Font("Segoe UI", 20.0!)
        Me.ms_menu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CatalogosToolStripMenuItem, Me.MovimientosToolStripMenuItem, Me.ReportesToolStripMenuItem, Me.UtileriaToolStripMenuItem, Me.ExportarAExcelToolStripMenuItem, Me.CerrarSesionToolStripMenuItem})
        Me.ms_menu.Location = New System.Drawing.Point(0, 0)
        Me.ms_menu.Name = "ms_menu"
        Me.ms_menu.Size = New System.Drawing.Size(1001, 45)
        Me.ms_menu.TabIndex = 0
        Me.ms_menu.Text = "MenuStrip1"
        '
        'CatalogosToolStripMenuItem
        '
        Me.CatalogosToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EmpleadosToolStripMenuItem, Me.ProveedoresToolStripMenuItem, Me.MaterialToolStripMenuItem, Me.tsm_usuarios})
        Me.CatalogosToolStripMenuItem.Name = "CatalogosToolStripMenuItem"
        Me.CatalogosToolStripMenuItem.Size = New System.Drawing.Size(149, 41)
        Me.CatalogosToolStripMenuItem.Text = "Catalogos"
        '
        'EmpleadosToolStripMenuItem
        '
        Me.EmpleadosToolStripMenuItem.Name = "EmpleadosToolStripMenuItem"
        Me.EmpleadosToolStripMenuItem.Size = New System.Drawing.Size(241, 42)
        Me.EmpleadosToolStripMenuItem.Text = "Empleados"
        '
        'ProveedoresToolStripMenuItem
        '
        Me.ProveedoresToolStripMenuItem.Name = "ProveedoresToolStripMenuItem"
        Me.ProveedoresToolStripMenuItem.Size = New System.Drawing.Size(241, 42)
        Me.ProveedoresToolStripMenuItem.Text = "Proveedores"
        '
        'MaterialToolStripMenuItem
        '
        Me.MaterialToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportarMaterialesToolStripMenuItem})
        Me.MaterialToolStripMenuItem.Name = "MaterialToolStripMenuItem"
        Me.MaterialToolStripMenuItem.Size = New System.Drawing.Size(241, 42)
        Me.MaterialToolStripMenuItem.Text = "Material"
        '
        'ImportarMaterialesToolStripMenuItem
        '
        Me.ImportarMaterialesToolStripMenuItem.Name = "ImportarMaterialesToolStripMenuItem"
        Me.ImportarMaterialesToolStripMenuItem.Size = New System.Drawing.Size(327, 42)
        Me.ImportarMaterialesToolStripMenuItem.Text = "Importar Materiales"
        '
        'tsm_usuarios
        '
        Me.tsm_usuarios.Name = "tsm_usuarios"
        Me.tsm_usuarios.Size = New System.Drawing.Size(241, 42)
        Me.tsm_usuarios.Text = "Usuarios"
        Me.tsm_usuarios.Visible = False
        '
        'MovimientosToolStripMenuItem
        '
        Me.MovimientosToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CompraToolStripMenuItem, Me.VentaToolStripMenuItem})
        Me.MovimientosToolStripMenuItem.Name = "MovimientosToolStripMenuItem"
        Me.MovimientosToolStripMenuItem.Size = New System.Drawing.Size(184, 41)
        Me.MovimientosToolStripMenuItem.Text = "Movimientos"
        '
        'CompraToolStripMenuItem
        '
        Me.CompraToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ConsultaToolStripMenuItem})
        Me.CompraToolStripMenuItem.Name = "CompraToolStripMenuItem"
        Me.CompraToolStripMenuItem.Size = New System.Drawing.Size(189, 42)
        Me.CompraToolStripMenuItem.Text = "Compra"
        '
        'ConsultaToolStripMenuItem
        '
        Me.ConsultaToolStripMenuItem.Name = "ConsultaToolStripMenuItem"
        Me.ConsultaToolStripMenuItem.Size = New System.Drawing.Size(198, 42)
        Me.ConsultaToolStripMenuItem.Text = "Consulta"
        '
        'VentaToolStripMenuItem
        '
        Me.VentaToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ConsultaToolStripMenuItem1})
        Me.VentaToolStripMenuItem.Name = "VentaToolStripMenuItem"
        Me.VentaToolStripMenuItem.Size = New System.Drawing.Size(189, 42)
        Me.VentaToolStripMenuItem.Text = "Venta"
        '
        'ConsultaToolStripMenuItem1
        '
        Me.ConsultaToolStripMenuItem1.Name = "ConsultaToolStripMenuItem1"
        Me.ConsultaToolStripMenuItem1.Size = New System.Drawing.Size(198, 42)
        Me.ConsultaToolStripMenuItem1.Text = "Consulta"
        '
        'ReportesToolStripMenuItem
        '
        Me.ReportesToolStripMenuItem.Name = "ReportesToolStripMenuItem"
        Me.ReportesToolStripMenuItem.Size = New System.Drawing.Size(133, 41)
        Me.ReportesToolStripMenuItem.Text = "Reportes"
        '
        'UtileriaToolStripMenuItem
        '
        Me.UtileriaToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RespaldoYRestauracionToolStripMenuItem, Me.ErroresToolStripMenuItem, Me.HistoricoToolStripMenuItem})
        Me.UtileriaToolStripMenuItem.Enabled = False
        Me.UtileriaToolStripMenuItem.Name = "UtileriaToolStripMenuItem"
        Me.UtileriaToolStripMenuItem.Size = New System.Drawing.Size(115, 41)
        Me.UtileriaToolStripMenuItem.Text = "Utileria"
        '
        'RespaldoYRestauracionToolStripMenuItem
        '
        Me.RespaldoYRestauracionToolStripMenuItem.Name = "RespaldoYRestauracionToolStripMenuItem"
        Me.RespaldoYRestauracionToolStripMenuItem.Size = New System.Drawing.Size(383, 42)
        Me.RespaldoYRestauracionToolStripMenuItem.Text = "Respaldo Y Restauracion"
        '
        'ErroresToolStripMenuItem
        '
        Me.ErroresToolStripMenuItem.Name = "ErroresToolStripMenuItem"
        Me.ErroresToolStripMenuItem.Size = New System.Drawing.Size(383, 42)
        Me.ErroresToolStripMenuItem.Text = "Bitacora"
        '
        'HistoricoToolStripMenuItem
        '
        Me.HistoricoToolStripMenuItem.Name = "HistoricoToolStripMenuItem"
        Me.HistoricoToolStripMenuItem.Size = New System.Drawing.Size(383, 42)
        Me.HistoricoToolStripMenuItem.Text = "Historico"
        '
        'ExportarAExcelToolStripMenuItem
        '
        Me.ExportarAExcelToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.VentasToolStripMenuItem})
        Me.ExportarAExcelToolStripMenuItem.Name = "ExportarAExcelToolStripMenuItem"
        Me.ExportarAExcelToolStripMenuItem.Size = New System.Drawing.Size(215, 41)
        Me.ExportarAExcelToolStripMenuItem.Text = "Exportar a Excel"
        '
        'VentasToolStripMenuItem
        '
        Me.VentasToolStripMenuItem.Name = "VentasToolStripMenuItem"
        Me.VentasToolStripMenuItem.Size = New System.Drawing.Size(172, 42)
        Me.VentasToolStripMenuItem.Text = "Ventas"
        '
        'CerrarSesionToolStripMenuItem
        '
        Me.CerrarSesionToolStripMenuItem.Name = "CerrarSesionToolStripMenuItem"
        Me.CerrarSesionToolStripMenuItem.Size = New System.Drawing.Size(182, 41)
        Me.CerrarSesionToolStripMenuItem.Text = "Cerrar sesión"
        '
        'lbl_portada
        '
        Me.lbl_portada.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbl_portada.Font = New System.Drawing.Font("Century Gothic", 27.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_portada.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.lbl_portada.Location = New System.Drawing.Point(0, 45)
        Me.lbl_portada.Name = "lbl_portada"
        Me.lbl_portada.Size = New System.Drawing.Size(1001, 371)
        Me.lbl_portada.TabIndex = 1
        Me.lbl_portada.Text = "Administrador De Inventario De Refaccionaria"
        '
        'Principal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DarkSlateGray
        Me.ClientSize = New System.Drawing.Size(1001, 416)
        Me.Controls.Add(Me.lbl_portada)
        Me.Controls.Add(Me.ms_menu)
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.ms_menu
        Me.Name = "Principal"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sistema Administrativo De Inventario"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ms_menu.ResumeLayout(False)
        Me.ms_menu.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ms_menu As MenuStrip
    Friend WithEvents CatalogosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents EmpleadosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProveedoresToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MaterialToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MovimientosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CompraToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents VentaToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ReportesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ConsultaToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ConsultaToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents tsm_usuarios As ToolStripMenuItem
    Friend WithEvents lbl_portada As Label
    Friend WithEvents UtileriaToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents RespaldoYRestauracionToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ErroresToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents HistoricoToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExportarAExcelToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents VentasToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents CerrarSesionToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportarMaterialesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
End Class
