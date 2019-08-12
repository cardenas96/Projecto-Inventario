<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Reportes
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ReportViewer1 = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ReporteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EmpleadosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProveedoresToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProductoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProductosMasVendidosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CompraToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CompraDetalleProductoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ComprasRealizadasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VentaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProductosFaltantesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PorPeriodoDeTiempoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VentasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProductoMasVendidosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ComprasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProductosCompradosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ComprasRealizadasToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.VentasPorEmpleadoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ComprasPorProveedorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.dtp_fecha = New System.Windows.Forms.DateTimePicker()
        Me.lbl_f1 = New System.Windows.Forms.Label()
        Me.lbl_f2 = New System.Windows.Forms.Label()
        Me.lbl_n1 = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ReportViewer1
        '
        Me.ReportViewer1.AutoSize = True
        Me.ReportViewer1.Location = New System.Drawing.Point(1, 28)
        Me.ReportViewer1.Name = "ReportViewer1"
        Me.ReportViewer1.Size = New System.Drawing.Size(874, 500)
        Me.ReportViewer1.TabIndex = 0
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReporteToolStripMenuItem, Me.PorPeriodoDeTiempoToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(875, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ReporteToolStripMenuItem
        '
        Me.ReporteToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EmpleadosToolStripMenuItem, Me.ProveedoresToolStripMenuItem, Me.ProductoToolStripMenuItem, Me.CompraToolStripMenuItem, Me.VentaToolStripMenuItem, Me.ProductosFaltantesToolStripMenuItem})
        Me.ReporteToolStripMenuItem.Name = "ReporteToolStripMenuItem"
        Me.ReporteToolStripMenuItem.Size = New System.Drawing.Size(60, 20)
        Me.ReporteToolStripMenuItem.Text = "Reporte"
        '
        'EmpleadosToolStripMenuItem
        '
        Me.EmpleadosToolStripMenuItem.Name = "EmpleadosToolStripMenuItem"
        Me.EmpleadosToolStripMenuItem.Size = New System.Drawing.Size(178, 22)
        Me.EmpleadosToolStripMenuItem.Text = "Empleados"
        '
        'ProveedoresToolStripMenuItem
        '
        Me.ProveedoresToolStripMenuItem.Name = "ProveedoresToolStripMenuItem"
        Me.ProveedoresToolStripMenuItem.Size = New System.Drawing.Size(178, 22)
        Me.ProveedoresToolStripMenuItem.Text = "Proveedores"
        '
        'ProductoToolStripMenuItem
        '
        Me.ProductoToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProductosMasVendidosToolStripMenuItem})
        Me.ProductoToolStripMenuItem.Name = "ProductoToolStripMenuItem"
        Me.ProductoToolStripMenuItem.Size = New System.Drawing.Size(178, 22)
        Me.ProductoToolStripMenuItem.Text = "Producto"
        '
        'ProductosMasVendidosToolStripMenuItem
        '
        Me.ProductosMasVendidosToolStripMenuItem.Name = "ProductosMasVendidosToolStripMenuItem"
        Me.ProductosMasVendidosToolStripMenuItem.Size = New System.Drawing.Size(204, 22)
        Me.ProductosMasVendidosToolStripMenuItem.Text = "Productos Mas Vendidos"
        '
        'CompraToolStripMenuItem
        '
        Me.CompraToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CompraDetalleProductoToolStripMenuItem, Me.ComprasRealizadasToolStripMenuItem})
        Me.CompraToolStripMenuItem.Name = "CompraToolStripMenuItem"
        Me.CompraToolStripMenuItem.Size = New System.Drawing.Size(178, 22)
        Me.CompraToolStripMenuItem.Text = "Compra"
        '
        'CompraDetalleProductoToolStripMenuItem
        '
        Me.CompraDetalleProductoToolStripMenuItem.Name = "CompraDetalleProductoToolStripMenuItem"
        Me.CompraDetalleProductoToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.CompraDetalleProductoToolStripMenuItem.Text = "Productos Comprados"
        '
        'ComprasRealizadasToolStripMenuItem
        '
        Me.ComprasRealizadasToolStripMenuItem.Name = "ComprasRealizadasToolStripMenuItem"
        Me.ComprasRealizadasToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.ComprasRealizadasToolStripMenuItem.Text = "Compras Realizadas"
        '
        'VentaToolStripMenuItem
        '
        Me.VentaToolStripMenuItem.Name = "VentaToolStripMenuItem"
        Me.VentaToolStripMenuItem.Size = New System.Drawing.Size(178, 22)
        Me.VentaToolStripMenuItem.Text = "Venta"
        '
        'ProductosFaltantesToolStripMenuItem
        '
        Me.ProductosFaltantesToolStripMenuItem.Name = "ProductosFaltantesToolStripMenuItem"
        Me.ProductosFaltantesToolStripMenuItem.Size = New System.Drawing.Size(178, 22)
        Me.ProductosFaltantesToolStripMenuItem.Text = "Productos Faltantes"
        '
        'PorPeriodoDeTiempoToolStripMenuItem
        '
        Me.PorPeriodoDeTiempoToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.VentasToolStripMenuItem, Me.ComprasToolStripMenuItem, Me.VentasPorEmpleadoToolStripMenuItem, Me.ComprasPorProveedorToolStripMenuItem})
        Me.PorPeriodoDeTiempoToolStripMenuItem.Name = "PorPeriodoDeTiempoToolStripMenuItem"
        Me.PorPeriodoDeTiempoToolStripMenuItem.Size = New System.Drawing.Size(141, 20)
        Me.PorPeriodoDeTiempoToolStripMenuItem.Text = "Por Periodo de Tiempo"
        '
        'VentasToolStripMenuItem
        '
        Me.VentasToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProductoMasVendidosToolStripMenuItem})
        Me.VentasToolStripMenuItem.Name = "VentasToolStripMenuItem"
        Me.VentasToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.VentasToolStripMenuItem.Text = "Ventas"
        '
        'ProductoMasVendidosToolStripMenuItem
        '
        Me.ProductoMasVendidosToolStripMenuItem.Name = "ProductoMasVendidosToolStripMenuItem"
        Me.ProductoMasVendidosToolStripMenuItem.Size = New System.Drawing.Size(199, 22)
        Me.ProductoMasVendidosToolStripMenuItem.Text = "Producto Mas Vendidos"
        '
        'ComprasToolStripMenuItem
        '
        Me.ComprasToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProductosCompradosToolStripMenuItem, Me.ComprasRealizadasToolStripMenuItem1})
        Me.ComprasToolStripMenuItem.Name = "ComprasToolStripMenuItem"
        Me.ComprasToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ComprasToolStripMenuItem.Text = "Compras"
        '
        'ProductosCompradosToolStripMenuItem
        '
        Me.ProductosCompradosToolStripMenuItem.Name = "ProductosCompradosToolStripMenuItem"
        Me.ProductosCompradosToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.ProductosCompradosToolStripMenuItem.Text = "Productos Comprados"
        '
        'ComprasRealizadasToolStripMenuItem1
        '
        Me.ComprasRealizadasToolStripMenuItem1.Name = "ComprasRealizadasToolStripMenuItem1"
        Me.ComprasRealizadasToolStripMenuItem1.Size = New System.Drawing.Size(193, 22)
        Me.ComprasRealizadasToolStripMenuItem1.Text = "Compras Realizadas"
        '
        'VentasPorEmpleadoToolStripMenuItem
        '
        Me.VentasPorEmpleadoToolStripMenuItem.Name = "VentasPorEmpleadoToolStripMenuItem"
        Me.VentasPorEmpleadoToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.VentasPorEmpleadoToolStripMenuItem.Text = "Ventas Por Empleado"
        '
        'ComprasPorProveedorToolStripMenuItem
        '
        Me.ComprasPorProveedorToolStripMenuItem.Name = "ComprasPorProveedorToolStripMenuItem"
        Me.ComprasPorProveedorToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ComprasPorProveedorToolStripMenuItem.Text = "Compras Por Proveedor"
        '
        'dtp_fecha
        '
        Me.dtp_fecha.Enabled = False
        Me.dtp_fecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtp_fecha.Location = New System.Drawing.Point(773, 0)
        Me.dtp_fecha.Name = "dtp_fecha"
        Me.dtp_fecha.Size = New System.Drawing.Size(102, 20)
        Me.dtp_fecha.TabIndex = 2
        '
        'lbl_f1
        '
        Me.lbl_f1.AutoSize = True
        Me.lbl_f1.Location = New System.Drawing.Point(112, 531)
        Me.lbl_f1.Name = "lbl_f1"
        Me.lbl_f1.Size = New System.Drawing.Size(10, 13)
        Me.lbl_f1.TabIndex = 3
        Me.lbl_f1.Text = "-"
        Me.lbl_f1.Visible = False
        '
        'lbl_f2
        '
        Me.lbl_f2.AutoSize = True
        Me.lbl_f2.Location = New System.Drawing.Point(186, 531)
        Me.lbl_f2.Name = "lbl_f2"
        Me.lbl_f2.Size = New System.Drawing.Size(10, 13)
        Me.lbl_f2.TabIndex = 4
        Me.lbl_f2.Text = "-"
        Me.lbl_f2.Visible = False
        '
        'lbl_n1
        '
        Me.lbl_n1.AutoSize = True
        Me.lbl_n1.Location = New System.Drawing.Point(249, 531)
        Me.lbl_n1.Name = "lbl_n1"
        Me.lbl_n1.Size = New System.Drawing.Size(10, 13)
        Me.lbl_n1.TabIndex = 5
        Me.lbl_n1.Text = "-"
        Me.lbl_n1.Visible = False
        '
        'Reportes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(875, 554)
        Me.Controls.Add(Me.lbl_n1)
        Me.Controls.Add(Me.lbl_f2)
        Me.Controls.Add(Me.lbl_f1)
        Me.Controls.Add(Me.dtp_fecha)
        Me.Controls.Add(Me.ReportViewer1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Reportes"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Reportes"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ReportViewer1 As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents ReporteToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents EmpleadosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProveedoresToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProductoToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents dtp_fecha As DateTimePicker
    Friend WithEvents PorPeriodoDeTiempoToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents lbl_f1 As Label
    Friend WithEvents lbl_f2 As Label
    Friend WithEvents VentasToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CompraToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents VentaToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ComprasToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents VentasPorEmpleadoToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ComprasPorProveedorToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProductosFaltantesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CompraDetalleProductoToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ComprasRealizadasToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProductosCompradosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ComprasRealizadasToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents lbl_n1 As Label
    Friend WithEvents ProductosMasVendidosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProductoMasVendidosToolStripMenuItem As ToolStripMenuItem
End Class
