<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Ventas
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ventas))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txt_telefono = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txt_direccion = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cb_empleado = New System.Windows.Forms.ComboBox()
        Me.txt_ide = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.cb_articulo = New System.Windows.Forms.ComboBox()
        Me.btn_limpiar = New System.Windows.Forms.Button()
        Me.btn_agregar = New System.Windows.Forms.Button()
        Me.txt_cantidad = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txt_precio = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txt_existencias = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txt_unidad = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txt_marca = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txt_ida = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.dtp_fecha = New System.Windows.Forms.DateTimePicker()
        Me.txt_id = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.dg_datos = New System.Windows.Forms.DataGridView()
        Me.no_venta = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.articulo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.uni = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cantidad = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.costo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.importe = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btn_salir = New System.Windows.Forms.Button()
        Me.btn_grabar = New System.Windows.Forms.Button()
        Me.btn_nuevo = New System.Windows.Forms.Button()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.lbl_total = New System.Windows.Forms.Label()
        Me.lbl_titulo = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.lblSub = New System.Windows.Forms.Label()
        Me.lblIva = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txt_telefono)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txt_direccion)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.cb_empleado)
        Me.GroupBox1.Controls.Add(Me.txt_ide)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 42)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(503, 104)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Empleado"
        '
        'txt_telefono
        '
        Me.txt_telefono.Enabled = False
        Me.txt_telefono.Location = New System.Drawing.Point(338, 60)
        Me.txt_telefono.Name = "txt_telefono"
        Me.txt_telefono.Size = New System.Drawing.Size(149, 27)
        Me.txt_telefono.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(262, 63)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 20)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Telefono:"
        '
        'txt_direccion
        '
        Me.txt_direccion.Enabled = False
        Me.txt_direccion.Location = New System.Drawing.Point(87, 60)
        Me.txt_direccion.Name = "txt_direccion"
        Me.txt_direccion.Size = New System.Drawing.Size(172, 27)
        Me.txt_direccion.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 63)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 20)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Direccion:"
        '
        'cb_empleado
        '
        Me.cb_empleado.FormattingEnabled = True
        Me.cb_empleado.Location = New System.Drawing.Point(196, 22)
        Me.cb_empleado.Name = "cb_empleado"
        Me.cb_empleado.Size = New System.Drawing.Size(291, 28)
        Me.cb_empleado.TabIndex = 3
        '
        'txt_ide
        '
        Me.txt_ide.Enabled = False
        Me.txt_ide.Location = New System.Drawing.Point(36, 22)
        Me.txt_ide.Name = "txt_ide"
        Me.txt_ide.Size = New System.Drawing.Size(85, 27)
        Me.txt_ide.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(127, 25)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 20)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Nombre:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(27, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ID:"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cb_articulo)
        Me.GroupBox2.Controls.Add(Me.btn_limpiar)
        Me.GroupBox2.Controls.Add(Me.btn_agregar)
        Me.GroupBox2.Controls.Add(Me.txt_cantidad)
        Me.GroupBox2.Controls.Add(Me.Label14)
        Me.GroupBox2.Controls.Add(Me.txt_precio)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.txt_existencias)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.txt_unidad)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.txt_marca)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.txt_ida)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 158)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(725, 143)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Articulos"
        '
        'cb_articulo
        '
        Me.cb_articulo.FormattingEnabled = True
        Me.cb_articulo.Location = New System.Drawing.Point(196, 23)
        Me.cb_articulo.Name = "cb_articulo"
        Me.cb_articulo.Size = New System.Drawing.Size(291, 28)
        Me.cb_articulo.Sorted = True
        Me.cb_articulo.TabIndex = 23
        '
        'btn_limpiar
        '
        Me.btn_limpiar.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_limpiar.Location = New System.Drawing.Point(560, 19)
        Me.btn_limpiar.Name = "btn_limpiar"
        Me.btn_limpiar.Size = New System.Drawing.Size(110, 34)
        Me.btn_limpiar.TabIndex = 22
        Me.btn_limpiar.Text = "Limpiar"
        Me.btn_limpiar.UseVisualStyleBackColor = True
        '
        'btn_agregar
        '
        Me.btn_agregar.Enabled = False
        Me.btn_agregar.Image = CType(resources.GetObject("btn_agregar.Image"), System.Drawing.Image)
        Me.btn_agregar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_agregar.Location = New System.Drawing.Point(252, 94)
        Me.btn_agregar.Name = "btn_agregar"
        Me.btn_agregar.Size = New System.Drawing.Size(169, 43)
        Me.btn_agregar.TabIndex = 21
        Me.btn_agregar.Text = "Agregar"
        Me.btn_agregar.UseVisualStyleBackColor = True
        '
        'txt_cantidad
        '
        Me.txt_cantidad.Location = New System.Drawing.Point(79, 99)
        Me.txt_cantidad.Name = "txt_cantidad"
        Me.txt_cantidad.Size = New System.Drawing.Size(88, 27)
        Me.txt_cantidad.TabIndex = 20
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(5, 102)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(72, 20)
        Me.Label14.TabIndex = 19
        Me.Label14.Text = "Cantidad:"
        '
        'txt_precio
        '
        Me.txt_precio.Enabled = False
        Me.txt_precio.Location = New System.Drawing.Point(626, 59)
        Me.txt_precio.Name = "txt_precio"
        Me.txt_precio.Size = New System.Drawing.Size(93, 27)
        Me.txt_precio.TabIndex = 17
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(567, 62)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(53, 20)
        Me.Label12.TabIndex = 16
        Me.Label12.Text = "Precio:"
        '
        'txt_existencias
        '
        Me.txt_existencias.Enabled = False
        Me.txt_existencias.Location = New System.Drawing.Point(458, 59)
        Me.txt_existencias.Name = "txt_existencias"
        Me.txt_existencias.Size = New System.Drawing.Size(88, 27)
        Me.txt_existencias.TabIndex = 15
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(378, 62)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(83, 20)
        Me.Label11.TabIndex = 14
        Me.Label11.Text = "Existencias:"
        '
        'txt_unidad
        '
        Me.txt_unidad.Enabled = False
        Me.txt_unidad.Location = New System.Drawing.Point(249, 59)
        Me.txt_unidad.Name = "txt_unidad"
        Me.txt_unidad.Size = New System.Drawing.Size(120, 27)
        Me.txt_unidad.TabIndex = 13
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(190, 62)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(60, 20)
        Me.Label10.TabIndex = 12
        Me.Label10.Text = "Unidad:"
        '
        'txt_marca
        '
        Me.txt_marca.Enabled = False
        Me.txt_marca.Location = New System.Drawing.Point(65, 59)
        Me.txt_marca.Name = "txt_marca"
        Me.txt_marca.Size = New System.Drawing.Size(120, 27)
        Me.txt_marca.TabIndex = 11
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 62)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(53, 20)
        Me.Label9.TabIndex = 10
        Me.Label9.Text = "Marca:"
        '
        'txt_ida
        '
        Me.txt_ida.Enabled = False
        Me.txt_ida.Location = New System.Drawing.Point(36, 26)
        Me.txt_ida.Name = "txt_ida"
        Me.txt_ida.Size = New System.Drawing.Size(85, 27)
        Me.txt_ida.TabIndex = 8
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(127, 29)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(67, 20)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "Nombre:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 29)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(27, 20)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = "ID:"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.dtp_fecha)
        Me.GroupBox3.Controls.Add(Me.txt_id)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Location = New System.Drawing.Point(517, 42)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(215, 103)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 67)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(50, 20)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "Fecha:"
        '
        'dtp_fecha
        '
        Me.dtp_fecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtp_fecha.Location = New System.Drawing.Point(78, 63)
        Me.dtp_fecha.Name = "dtp_fecha"
        Me.dtp_fecha.Size = New System.Drawing.Size(126, 27)
        Me.dtp_fecha.TabIndex = 8
        '
        'txt_id
        '
        Me.txt_id.Enabled = False
        Me.txt_id.Location = New System.Drawing.Point(78, 18)
        Me.txt_id.Name = "txt_id"
        Me.txt_id.Size = New System.Drawing.Size(126, 27)
        Me.txt_id.TabIndex = 7
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 20)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "No.Venta:"
        '
        'dg_datos
        '
        Me.dg_datos.BackgroundColor = System.Drawing.SystemColors.ControlDarkDark
        Me.dg_datos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg_datos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.no_venta, Me.articulo, Me.uni, Me.cantidad, Me.costo, Me.importe})
        Me.dg_datos.Location = New System.Drawing.Point(8, 312)
        Me.dg_datos.Name = "dg_datos"
        Me.dg_datos.Size = New System.Drawing.Size(725, 101)
        Me.dg_datos.TabIndex = 3
        '
        'no_venta
        '
        Me.no_venta.HeaderText = "No."
        Me.no_venta.Name = "no_venta"
        Me.no_venta.ReadOnly = True
        Me.no_venta.Width = 75
        '
        'articulo
        '
        Me.articulo.HeaderText = "Articulo"
        Me.articulo.Name = "articulo"
        Me.articulo.ReadOnly = True
        Me.articulo.Width = 170
        '
        'uni
        '
        Me.uni.HeaderText = "Unidad"
        Me.uni.Name = "uni"
        Me.uni.ReadOnly = True
        Me.uni.Width = 90
        '
        'cantidad
        '
        Me.cantidad.HeaderText = "Cantidad"
        Me.cantidad.Name = "cantidad"
        Me.cantidad.ReadOnly = True
        '
        'costo
        '
        Me.costo.HeaderText = "Precio"
        Me.costo.Name = "costo"
        Me.costo.ReadOnly = True
        '
        'importe
        '
        Me.importe.HeaderText = "Importe"
        Me.importe.Name = "importe"
        Me.importe.ReadOnly = True
        Me.importe.Width = 120
        '
        'btn_salir
        '
        Me.btn_salir.Image = CType(resources.GetObject("btn_salir.Image"), System.Drawing.Image)
        Me.btn_salir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_salir.Location = New System.Drawing.Point(550, 525)
        Me.btn_salir.Name = "btn_salir"
        Me.btn_salir.Size = New System.Drawing.Size(177, 48)
        Me.btn_salir.TabIndex = 27
        Me.btn_salir.Text = "Salir"
        Me.btn_salir.UseVisualStyleBackColor = True
        '
        'btn_grabar
        '
        Me.btn_grabar.Enabled = False
        Me.btn_grabar.Image = CType(resources.GetObject("btn_grabar.Image"), System.Drawing.Image)
        Me.btn_grabar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_grabar.Location = New System.Drawing.Point(274, 525)
        Me.btn_grabar.Name = "btn_grabar"
        Me.btn_grabar.Size = New System.Drawing.Size(177, 48)
        Me.btn_grabar.TabIndex = 26
        Me.btn_grabar.Text = "Grabar"
        Me.btn_grabar.UseVisualStyleBackColor = True
        '
        'btn_nuevo
        '
        Me.btn_nuevo.Image = CType(resources.GetObject("btn_nuevo.Image"), System.Drawing.Image)
        Me.btn_nuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_nuevo.Location = New System.Drawing.Point(8, 525)
        Me.btn_nuevo.Name = "btn_nuevo"
        Me.btn_nuevo.Size = New System.Drawing.Size(177, 48)
        Me.btn_nuevo.TabIndex = 25
        Me.btn_nuevo.Text = "Nuevo"
        Me.btn_nuevo.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(562, 493)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(62, 24)
        Me.Label13.TabIndex = 28
        Me.Label13.Text = "Total:"
        '
        'lbl_total
        '
        Me.lbl_total.AutoSize = True
        Me.lbl_total.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_total.Location = New System.Drawing.Point(624, 493)
        Me.lbl_total.Name = "lbl_total"
        Me.lbl_total.Size = New System.Drawing.Size(38, 24)
        Me.lbl_total.TabIndex = 29
        Me.lbl_total.Text = "0.0"
        '
        'lbl_titulo
        '
        Me.lbl_titulo.BackColor = System.Drawing.Color.LightSeaGreen
        Me.lbl_titulo.Dock = System.Windows.Forms.DockStyle.Top
        Me.lbl_titulo.Font = New System.Drawing.Font("Century Gothic", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_titulo.Location = New System.Drawing.Point(0, 0)
        Me.lbl_titulo.Name = "lbl_titulo"
        Me.lbl_titulo.Size = New System.Drawing.Size(739, 43)
        Me.lbl_titulo.TabIndex = 30
        Me.lbl_titulo.Text = "Venta de Articulos"
        Me.lbl_titulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold)
        Me.Label15.Location = New System.Drawing.Point(534, 434)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(91, 24)
        Me.Label15.TabIndex = 31
        Me.Label15.Text = "Subtotal:"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold)
        Me.Label16.Location = New System.Drawing.Point(575, 463)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(49, 24)
        Me.Label16.TabIndex = 32
        Me.Label16.Text = "IVA:"
        '
        'lblSub
        '
        Me.lblSub.AutoSize = True
        Me.lblSub.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold)
        Me.lblSub.Location = New System.Drawing.Point(624, 434)
        Me.lblSub.Name = "lblSub"
        Me.lblSub.Size = New System.Drawing.Size(38, 24)
        Me.lblSub.TabIndex = 33
        Me.lblSub.Text = "0.0"
        '
        'lblIva
        '
        Me.lblIva.AutoSize = True
        Me.lblIva.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold)
        Me.lblIva.Location = New System.Drawing.Point(624, 463)
        Me.lblIva.Name = "lblIva"
        Me.lblIva.Size = New System.Drawing.Size(38, 24)
        Me.lblIva.TabIndex = 34
        Me.lblIva.Text = "0.0"
        '
        'Ventas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Highlight
        Me.ClientSize = New System.Drawing.Size(739, 579)
        Me.Controls.Add(Me.lblIva)
        Me.Controls.Add(Me.lblSub)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.lbl_titulo)
        Me.Controls.Add(Me.lbl_total)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.btn_salir)
        Me.Controls.Add(Me.btn_grabar)
        Me.Controls.Add(Me.btn_nuevo)
        Me.Controls.Add(Me.dg_datos)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Ventas"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Ventas"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents txt_telefono As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents txt_direccion As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents cb_empleado As ComboBox
    Friend WithEvents txt_ide As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents dtp_fecha As DateTimePicker
    Friend WithEvents txt_id As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txt_ida As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents btn_limpiar As Button
    Friend WithEvents btn_agregar As Button
    Friend WithEvents txt_cantidad As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents txt_precio As TextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents txt_existencias As TextBox
    Friend WithEvents Label11 As Label
    Friend WithEvents txt_unidad As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents txt_marca As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents dg_datos As DataGridView
    Friend WithEvents btn_salir As Button
    Friend WithEvents btn_grabar As Button
    Friend WithEvents btn_nuevo As Button
    Friend WithEvents Label13 As Label
    Friend WithEvents lbl_total As Label
    Friend WithEvents no_venta As DataGridViewTextBoxColumn
    Friend WithEvents articulo As DataGridViewTextBoxColumn
    Friend WithEvents uni As DataGridViewTextBoxColumn
    Friend WithEvents cantidad As DataGridViewTextBoxColumn
    Friend WithEvents costo As DataGridViewTextBoxColumn
    Friend WithEvents importe As DataGridViewTextBoxColumn
    Friend WithEvents cb_articulo As ComboBox
    Friend WithEvents lbl_titulo As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label16 As Label
    Friend WithEvents lblSub As Label
    Friend WithEvents lblIva As Label
End Class
