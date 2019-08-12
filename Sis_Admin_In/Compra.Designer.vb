<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Compra
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Compra))
        Me.lbl_id = New System.Windows.Forms.Label()
        Me.txt_idcompra = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txt_folio = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtp_fecha = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lbl_proveedor = New System.Windows.Forms.Label()
        Me.cb_proveedor = New System.Windows.Forms.ComboBox()
        Me.txt_direccion = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txt_fax = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txt_telefono = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txt_colonia = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txt_contacto = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txt_ciudad = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txt_idp = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txt_cantidad = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.txt_costo = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.btn_agregar = New System.Windows.Forms.Button()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cb_articulo = New System.Windows.Forms.ComboBox()
        Me.txt_costo_a = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txt_max = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txt_min = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txt_existencias = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txt_unidad = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txt_marca = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txt_ida = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.dg_datos = New System.Windows.Forms.DataGridView()
        Me.id = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.nombre = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.unidad = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Costo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cantidad = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.subtotal = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ida = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txt_total = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.txt_iva = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.txt_subtotal = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.btn_nuevo = New System.Windows.Forms.Button()
        Me.btn_grabar = New System.Windows.Forms.Button()
        Me.btn_salir = New System.Windows.Forms.Button()
        Me.lbl_titulo = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lbl_id
        '
        Me.lbl_id.AutoSize = True
        Me.lbl_id.Location = New System.Drawing.Point(642, 50)
        Me.lbl_id.Name = "lbl_id"
        Me.lbl_id.Size = New System.Drawing.Size(23, 17)
        Me.lbl_id.TabIndex = 0
        Me.lbl_id.Text = "ID:"
        '
        'txt_idcompra
        '
        Me.txt_idcompra.Enabled = False
        Me.txt_idcompra.Location = New System.Drawing.Point(678, 47)
        Me.txt_idcompra.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_idcompra.Name = "txt_idcompra"
        Me.txt_idcompra.ReadOnly = True
        Me.txt_idcompra.Size = New System.Drawing.Size(115, 25)
        Me.txt_idcompra.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(626, 83)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 17)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Folio:"
        '
        'txt_folio
        '
        Me.txt_folio.Enabled = False
        Me.txt_folio.Location = New System.Drawing.Point(678, 80)
        Me.txt_folio.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_folio.Name = "txt_folio"
        Me.txt_folio.Size = New System.Drawing.Size(116, 25)
        Me.txt_folio.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(506, 115)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(44, 17)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Fecha:"
        '
        'dtp_fecha
        '
        Me.dtp_fecha.Location = New System.Drawing.Point(558, 112)
        Me.dtp_fecha.Name = "dtp_fecha"
        Me.dtp_fecha.Size = New System.Drawing.Size(236, 25)
        Me.dtp_fecha.TabIndex = 3
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lbl_proveedor)
        Me.GroupBox1.Controls.Add(Me.cb_proveedor)
        Me.GroupBox1.Controls.Add(Me.txt_direccion)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.txt_fax)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.txt_telefono)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.txt_colonia)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txt_contacto)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.txt_ciudad)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txt_idp)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Cursor = System.Windows.Forms.Cursors.Default
        Me.GroupBox1.Location = New System.Drawing.Point(12, 46)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox1.Size = New System.Drawing.Size(490, 190)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Proveedor"
        '
        'lbl_proveedor
        '
        Me.lbl_proveedor.AutoSize = True
        Me.lbl_proveedor.Location = New System.Drawing.Point(128, 28)
        Me.lbl_proveedor.Name = "lbl_proveedor"
        Me.lbl_proveedor.Size = New System.Drawing.Size(72, 17)
        Me.lbl_proveedor.TabIndex = 17
        Me.lbl_proveedor.Text = "Proveedor:"
        '
        'cb_proveedor
        '
        Me.cb_proveedor.FormattingEnabled = True
        Me.cb_proveedor.Location = New System.Drawing.Point(203, 25)
        Me.cb_proveedor.Name = "cb_proveedor"
        Me.cb_proveedor.Size = New System.Drawing.Size(253, 25)
        Me.cb_proveedor.TabIndex = 16
        '
        'txt_direccion
        '
        Me.txt_direccion.Location = New System.Drawing.Point(78, 157)
        Me.txt_direccion.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_direccion.Name = "txt_direccion"
        Me.txt_direccion.ReadOnly = True
        Me.txt_direccion.Size = New System.Drawing.Size(378, 25)
        Me.txt_direccion.TabIndex = 15
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(13, 160)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(65, 17)
        Me.Label10.TabIndex = 14
        Me.Label10.Text = "Direccion:"
        '
        'txt_fax
        '
        Me.txt_fax.Location = New System.Drawing.Point(293, 91)
        Me.txt_fax.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_fax.Name = "txt_fax"
        Me.txt_fax.ReadOnly = True
        Me.txt_fax.Size = New System.Drawing.Size(163, 25)
        Me.txt_fax.TabIndex = 13
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(257, 94)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(30, 17)
        Me.Label9.TabIndex = 12
        Me.Label9.Text = "Fax:"
        '
        'txt_telefono
        '
        Me.txt_telefono.Location = New System.Drawing.Point(78, 91)
        Me.txt_telefono.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_telefono.Name = "txt_telefono"
        Me.txt_telefono.ReadOnly = True
        Me.txt_telefono.Size = New System.Drawing.Size(151, 25)
        Me.txt_telefono.TabIndex = 11
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(16, 94)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(61, 17)
        Me.Label8.TabIndex = 10
        Me.Label8.Text = "Telefono:"
        '
        'txt_colonia
        '
        Me.txt_colonia.Location = New System.Drawing.Point(293, 124)
        Me.txt_colonia.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_colonia.Name = "txt_colonia"
        Me.txt_colonia.ReadOnly = True
        Me.txt_colonia.Size = New System.Drawing.Size(163, 25)
        Me.txt_colonia.TabIndex = 9
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(241, 127)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(55, 17)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "Colonia:"
        '
        'txt_contacto
        '
        Me.txt_contacto.Location = New System.Drawing.Point(82, 58)
        Me.txt_contacto.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_contacto.Name = "txt_contacto"
        Me.txt_contacto.ReadOnly = True
        Me.txt_contacto.Size = New System.Drawing.Size(374, 25)
        Me.txt_contacto.TabIndex = 7
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(16, 61)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(63, 17)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Contacto:"
        '
        'txt_ciudad
        '
        Me.txt_ciudad.Location = New System.Drawing.Point(78, 124)
        Me.txt_ciudad.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_ciudad.Name = "txt_ciudad"
        Me.txt_ciudad.ReadOnly = True
        Me.txt_ciudad.Size = New System.Drawing.Size(151, 25)
        Me.txt_ciudad.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(26, 127)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(52, 17)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Ciudad:"
        '
        'txt_idp
        '
        Me.txt_idp.Location = New System.Drawing.Point(52, 25)
        Me.txt_idp.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_idp.Name = "txt_idp"
        Me.txt_idp.ReadOnly = True
        Me.txt_idp.Size = New System.Drawing.Size(71, 25)
        Me.txt_idp.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 28)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 17)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "IDP:"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txt_cantidad)
        Me.GroupBox2.Controls.Add(Me.Label20)
        Me.GroupBox2.Controls.Add(Me.txt_costo)
        Me.GroupBox2.Controls.Add(Me.Label21)
        Me.GroupBox2.Controls.Add(Me.btn_agregar)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.cb_articulo)
        Me.GroupBox2.Controls.Add(Me.txt_costo_a)
        Me.GroupBox2.Controls.Add(Me.Label13)
        Me.GroupBox2.Controls.Add(Me.txt_max)
        Me.GroupBox2.Controls.Add(Me.Label14)
        Me.GroupBox2.Controls.Add(Me.txt_min)
        Me.GroupBox2.Controls.Add(Me.Label15)
        Me.GroupBox2.Controls.Add(Me.txt_existencias)
        Me.GroupBox2.Controls.Add(Me.Label16)
        Me.GroupBox2.Controls.Add(Me.txt_unidad)
        Me.GroupBox2.Controls.Add(Me.Label17)
        Me.GroupBox2.Controls.Add(Me.txt_marca)
        Me.GroupBox2.Controls.Add(Me.Label18)
        Me.GroupBox2.Controls.Add(Me.txt_ida)
        Me.GroupBox2.Controls.Add(Me.Label19)
        Me.GroupBox2.Cursor = System.Windows.Forms.Cursors.Default
        Me.GroupBox2.Location = New System.Drawing.Point(12, 242)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox2.Size = New System.Drawing.Size(781, 148)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Articulo"
        '
        'txt_cantidad
        '
        Me.txt_cantidad.Enabled = False
        Me.txt_cantidad.Location = New System.Drawing.Point(643, 60)
        Me.txt_cantidad.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_cantidad.Name = "txt_cantidad"
        Me.txt_cantidad.Size = New System.Drawing.Size(132, 25)
        Me.txt_cantidad.TabIndex = 22
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(583, 63)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(63, 17)
        Me.Label20.TabIndex = 21
        Me.Label20.Text = "Cantidad:"
        '
        'txt_costo
        '
        Me.txt_costo.Enabled = False
        Me.txt_costo.Location = New System.Drawing.Point(643, 24)
        Me.txt_costo.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_costo.Name = "txt_costo"
        Me.txt_costo.Size = New System.Drawing.Size(132, 25)
        Me.txt_costo.TabIndex = 20
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(585, 29)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(45, 17)
        Me.Label21.TabIndex = 19
        Me.Label21.Text = "Costo:"
        '
        'btn_agregar
        '
        Me.btn_agregar.Enabled = False
        Me.btn_agregar.Image = CType(resources.GetObject("btn_agregar.Image"), System.Drawing.Image)
        Me.btn_agregar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_agregar.Location = New System.Drawing.Point(623, 93)
        Me.btn_agregar.Name = "btn_agregar"
        Me.btn_agregar.Size = New System.Drawing.Size(151, 48)
        Me.btn_agregar.TabIndex = 18
        Me.btn_agregar.Text = "Agregar"
        Me.btn_agregar.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(144, 32)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(60, 17)
        Me.Label12.TabIndex = 17
        Me.Label12.Text = "Nombre:"
        '
        'cb_articulo
        '
        Me.cb_articulo.FormattingEnabled = True
        Me.cb_articulo.Location = New System.Drawing.Point(210, 28)
        Me.cb_articulo.Name = "cb_articulo"
        Me.cb_articulo.Size = New System.Drawing.Size(194, 25)
        Me.cb_articulo.TabIndex = 16
        '
        'txt_costo_a
        '
        Me.txt_costo_a.Location = New System.Drawing.Point(460, 62)
        Me.txt_costo_a.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_costo_a.Name = "txt_costo_a"
        Me.txt_costo_a.ReadOnly = True
        Me.txt_costo_a.Size = New System.Drawing.Size(116, 25)
        Me.txt_costo_a.TabIndex = 15
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(408, 65)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(45, 17)
        Me.Label13.TabIndex = 14
        Me.Label13.Text = "Costo:"
        '
        'txt_max
        '
        Me.txt_max.Location = New System.Drawing.Point(276, 95)
        Me.txt_max.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_max.Name = "txt_max"
        Me.txt_max.ReadOnly = True
        Me.txt_max.Size = New System.Drawing.Size(116, 25)
        Me.txt_max.TabIndex = 13
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(241, 98)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(36, 17)
        Me.Label14.TabIndex = 12
        Me.Label14.Text = "Max:"
        '
        'txt_min
        '
        Me.txt_min.Location = New System.Drawing.Point(68, 95)
        Me.txt_min.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_min.Name = "txt_min"
        Me.txt_min.ReadOnly = True
        Me.txt_min.Size = New System.Drawing.Size(116, 25)
        Me.txt_min.TabIndex = 11
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(36, 98)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(33, 17)
        Me.Label15.TabIndex = 10
        Me.Label15.Text = "Min:"
        '
        'txt_existencias
        '
        Me.txt_existencias.Location = New System.Drawing.Point(288, 62)
        Me.txt_existencias.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_existencias.Name = "txt_existencias"
        Me.txt_existencias.ReadOnly = True
        Me.txt_existencias.Size = New System.Drawing.Size(116, 25)
        Me.txt_existencias.TabIndex = 9
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(214, 67)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(73, 17)
        Me.Label16.TabIndex = 8
        Me.Label16.Text = "Existencias:"
        '
        'txt_unidad
        '
        Me.txt_unidad.Location = New System.Drawing.Point(68, 62)
        Me.txt_unidad.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_unidad.Name = "txt_unidad"
        Me.txt_unidad.ReadOnly = True
        Me.txt_unidad.Size = New System.Drawing.Size(140, 25)
        Me.txt_unidad.TabIndex = 7
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(16, 65)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(53, 17)
        Me.Label17.TabIndex = 6
        Me.Label17.Text = "Unidad:"
        '
        'txt_marca
        '
        Me.txt_marca.Location = New System.Drawing.Point(460, 29)
        Me.txt_marca.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_marca.Name = "txt_marca"
        Me.txt_marca.ReadOnly = True
        Me.txt_marca.Size = New System.Drawing.Size(116, 25)
        Me.txt_marca.TabIndex = 5
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(408, 32)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(48, 17)
        Me.Label18.TabIndex = 4
        Me.Label18.Text = "Marca:"
        '
        'txt_ida
        '
        Me.txt_ida.Location = New System.Drawing.Point(45, 29)
        Me.txt_ida.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_ida.Name = "txt_ida"
        Me.txt_ida.ReadOnly = True
        Me.txt_ida.Size = New System.Drawing.Size(86, 25)
        Me.txt_ida.TabIndex = 3
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(16, 32)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(23, 17)
        Me.Label19.TabIndex = 2
        Me.Label19.Text = "ID:"
        '
        'dg_datos
        '
        Me.dg_datos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg_datos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.id, Me.nombre, Me.unidad, Me.Costo, Me.cantidad, Me.subtotal, Me.ida})
        Me.dg_datos.Location = New System.Drawing.Point(15, 396)
        Me.dg_datos.Name = "dg_datos"
        Me.dg_datos.ReadOnly = True
        Me.dg_datos.Size = New System.Drawing.Size(595, 136)
        Me.dg_datos.TabIndex = 6
        '
        'id
        '
        Me.id.HeaderText = "ID"
        Me.id.Name = "id"
        Me.id.ReadOnly = True
        Me.id.Width = 40
        '
        'nombre
        '
        Me.nombre.HeaderText = "Articulo"
        Me.nombre.Name = "nombre"
        Me.nombre.ReadOnly = True
        Me.nombre.Width = 165
        '
        'unidad
        '
        Me.unidad.HeaderText = "Unidad"
        Me.unidad.Name = "unidad"
        Me.unidad.ReadOnly = True
        '
        'Costo
        '
        Me.Costo.HeaderText = "Cantidad"
        Me.Costo.Name = "Costo"
        Me.Costo.ReadOnly = True
        Me.Costo.Width = 70
        '
        'cantidad
        '
        Me.cantidad.HeaderText = "Costo"
        Me.cantidad.Name = "cantidad"
        Me.cantidad.ReadOnly = True
        Me.cantidad.Width = 75
        '
        'subtotal
        '
        Me.subtotal.HeaderText = "Subtotal"
        Me.subtotal.Name = "subtotal"
        Me.subtotal.ReadOnly = True
        '
        'ida
        '
        Me.ida.HeaderText = "ida"
        Me.ida.Name = "ida"
        Me.ida.ReadOnly = True
        Me.ida.Visible = False
        '
        'txt_total
        '
        Me.txt_total.Location = New System.Drawing.Point(677, 482)
        Me.txt_total.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_total.Name = "txt_total"
        Me.txt_total.ReadOnly = True
        Me.txt_total.Size = New System.Drawing.Size(116, 25)
        Me.txt_total.TabIndex = 21
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(625, 485)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(39, 17)
        Me.Label22.TabIndex = 20
        Me.Label22.Text = "Total:"
        '
        'txt_iva
        '
        Me.txt_iva.Location = New System.Drawing.Point(677, 447)
        Me.txt_iva.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_iva.Name = "txt_iva"
        Me.txt_iva.ReadOnly = True
        Me.txt_iva.Size = New System.Drawing.Size(116, 25)
        Me.txt_iva.TabIndex = 19
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(625, 452)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(29, 17)
        Me.Label23.TabIndex = 18
        Me.Label23.Text = "IVA:"
        '
        'txt_subtotal
        '
        Me.txt_subtotal.Location = New System.Drawing.Point(677, 413)
        Me.txt_subtotal.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_subtotal.Name = "txt_subtotal"
        Me.txt_subtotal.ReadOnly = True
        Me.txt_subtotal.Size = New System.Drawing.Size(116, 25)
        Me.txt_subtotal.TabIndex = 17
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(616, 421)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(59, 17)
        Me.Label24.TabIndex = 16
        Me.Label24.Text = "Subtotal:"
        '
        'btn_nuevo
        '
        Me.btn_nuevo.Font = New System.Drawing.Font("Segoe UI", 11.0!)
        Me.btn_nuevo.Image = CType(resources.GetObject("btn_nuevo.Image"), System.Drawing.Image)
        Me.btn_nuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_nuevo.Location = New System.Drawing.Point(15, 546)
        Me.btn_nuevo.Name = "btn_nuevo"
        Me.btn_nuevo.Size = New System.Drawing.Size(177, 48)
        Me.btn_nuevo.TabIndex = 22
        Me.btn_nuevo.Text = "Nuevo"
        Me.btn_nuevo.UseVisualStyleBackColor = True
        '
        'btn_grabar
        '
        Me.btn_grabar.Enabled = False
        Me.btn_grabar.Font = New System.Drawing.Font("Segoe UI", 11.0!)
        Me.btn_grabar.Image = CType(resources.GetObject("btn_grabar.Image"), System.Drawing.Image)
        Me.btn_grabar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_grabar.Location = New System.Drawing.Point(325, 546)
        Me.btn_grabar.Name = "btn_grabar"
        Me.btn_grabar.Size = New System.Drawing.Size(177, 48)
        Me.btn_grabar.TabIndex = 23
        Me.btn_grabar.Text = "Grabar"
        Me.btn_grabar.UseVisualStyleBackColor = True
        '
        'btn_salir
        '
        Me.btn_salir.Font = New System.Drawing.Font("Segoe UI", 11.0!)
        Me.btn_salir.Image = CType(resources.GetObject("btn_salir.Image"), System.Drawing.Image)
        Me.btn_salir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_salir.Location = New System.Drawing.Point(618, 546)
        Me.btn_salir.Name = "btn_salir"
        Me.btn_salir.Size = New System.Drawing.Size(177, 48)
        Me.btn_salir.TabIndex = 24
        Me.btn_salir.Text = "Salir"
        Me.btn_salir.UseVisualStyleBackColor = True
        '
        'lbl_titulo
        '
        Me.lbl_titulo.BackColor = System.Drawing.Color.LightSeaGreen
        Me.lbl_titulo.Dock = System.Windows.Forms.DockStyle.Top
        Me.lbl_titulo.Font = New System.Drawing.Font("Century Gothic", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_titulo.Location = New System.Drawing.Point(0, 0)
        Me.lbl_titulo.Name = "lbl_titulo"
        Me.lbl_titulo.Size = New System.Drawing.Size(807, 43)
        Me.lbl_titulo.TabIndex = 31
        Me.lbl_titulo.Text = "Compra de Articulos"
        Me.lbl_titulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Compra
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Highlight
        Me.ClientSize = New System.Drawing.Size(807, 609)
        Me.Controls.Add(Me.lbl_titulo)
        Me.Controls.Add(Me.btn_salir)
        Me.Controls.Add(Me.btn_grabar)
        Me.Controls.Add(Me.btn_nuevo)
        Me.Controls.Add(Me.txt_total)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.txt_iva)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.txt_subtotal)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.dg_datos)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.dtp_fecha)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txt_folio)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txt_idcompra)
        Me.Controls.Add(Me.lbl_id)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "Compra"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Compra"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub



    Friend WithEvents lbl_id As Label
    Friend WithEvents txt_idcompra As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txt_folio As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents dtp_fecha As DateTimePicker
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents lbl_proveedor As Label
    Friend WithEvents cb_proveedor As ComboBox
    Friend WithEvents txt_direccion As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents txt_fax As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents txt_telefono As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents txt_colonia As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents txt_contacto As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents txt_ciudad As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txt_idp As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents txt_cantidad As TextBox
    Friend WithEvents Label20 As Label
    Friend WithEvents txt_costo As TextBox
    Friend WithEvents Label21 As Label
    Friend WithEvents btn_agregar As Button
    Friend WithEvents Label12 As Label
    Friend WithEvents cb_articulo As ComboBox
    Friend WithEvents txt_costo_a As TextBox
    Friend WithEvents Label13 As Label
    Friend WithEvents txt_max As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents txt_min As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents txt_existencias As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents txt_unidad As TextBox
    Friend WithEvents Label17 As Label
    Friend WithEvents txt_marca As TextBox
    Friend WithEvents Label18 As Label
    Friend WithEvents txt_ida As TextBox
    Friend WithEvents Label19 As Label
    Friend WithEvents dg_datos As DataGridView
    Friend WithEvents txt_total As TextBox
    Friend WithEvents Label22 As Label
    Friend WithEvents txt_iva As TextBox
    Friend WithEvents Label23 As Label
    Friend WithEvents txt_subtotal As TextBox
    Friend WithEvents Label24 As Label
    Friend WithEvents btn_nuevo As Button
    Friend WithEvents btn_grabar As Button
    Friend WithEvents btn_salir As Button
    Friend WithEvents id As DataGridViewTextBoxColumn
    Friend WithEvents nombre As DataGridViewTextBoxColumn
    Friend WithEvents unidad As DataGridViewTextBoxColumn
    Friend WithEvents Costo As DataGridViewTextBoxColumn
    Friend WithEvents cantidad As DataGridViewTextBoxColumn
    Friend WithEvents subtotal As DataGridViewTextBoxColumn
    Friend WithEvents ida As DataGridViewTextBoxColumn
    Friend WithEvents lbl_titulo As Label
End Class
