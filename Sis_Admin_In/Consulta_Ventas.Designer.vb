<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Consulta_Ventas
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Consulta_Ventas))
        Me.btn_salir = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cb_idventa = New System.Windows.Forms.ComboBox()
        Me.rb_idventa = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cb_empleado = New System.Windows.Forms.ComboBox()
        Me.dtp_final = New System.Windows.Forms.DateTimePicker()
        Me.dtp_inicio = New System.Windows.Forms.DateTimePicker()
        Me.rb_todos = New System.Windows.Forms.RadioButton()
        Me.rb_fecha = New System.Windows.Forms.RadioButton()
        Me.rb_empleado = New System.Windows.Forms.RadioButton()
        Me.btn_buscar = New System.Windows.Forms.Button()
        Me.dg_datos = New System.Windows.Forms.DataGridView()
        Me.idcompra = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.fecha = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.articulo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cantidad = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.costo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.empleado = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lbl_total = New System.Windows.Forms.Label()
        Me.lbl_titulo = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btn_salir
        '
        Me.btn_salir.BackColor = System.Drawing.SystemColors.Control
        Me.btn_salir.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btn_salir.FlatAppearance.BorderSize = 2
        Me.btn_salir.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btn_salir.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.btn_salir.Image = CType(resources.GetObject("btn_salir.Image"), System.Drawing.Image)
        Me.btn_salir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_salir.Location = New System.Drawing.Point(242, 408)
        Me.btn_salir.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btn_salir.Name = "btn_salir"
        Me.btn_salir.Size = New System.Drawing.Size(222, 49)
        Me.btn_salir.TabIndex = 5
        Me.btn_salir.Text = "Salir"
        Me.btn_salir.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cb_idventa)
        Me.GroupBox1.Controls.Add(Me.rb_idventa)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.cb_empleado)
        Me.GroupBox1.Controls.Add(Me.dtp_final)
        Me.GroupBox1.Controls.Add(Me.dtp_inicio)
        Me.GroupBox1.Controls.Add(Me.rb_todos)
        Me.GroupBox1.Controls.Add(Me.rb_fecha)
        Me.GroupBox1.Controls.Add(Me.rb_empleado)
        Me.GroupBox1.Controls.Add(Me.btn_buscar)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.GroupBox1.Location = New System.Drawing.Point(16, 46)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox1.Size = New System.Drawing.Size(665, 179)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Consulta De Ventas"
        '
        'cb_idventa
        '
        Me.cb_idventa.Enabled = False
        Me.cb_idventa.FormattingEnabled = True
        Me.cb_idventa.Location = New System.Drawing.Point(122, 105)
        Me.cb_idventa.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cb_idventa.Name = "cb_idventa"
        Me.cb_idventa.Size = New System.Drawing.Size(221, 28)
        Me.cb_idventa.TabIndex = 10
        '
        'rb_idventa
        '
        Me.rb_idventa.AutoSize = True
        Me.rb_idventa.Location = New System.Drawing.Point(8, 105)
        Me.rb_idventa.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.rb_idventa.Name = "rb_idventa"
        Me.rb_idventa.Size = New System.Drawing.Size(94, 24)
        Me.rb_idventa.TabIndex = 9
        Me.rb_idventa.TabStop = True
        Me.rb_idventa.Text = "No. Venta:"
        Me.rb_idventa.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(255, 73)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 20)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Hasta:"
        '
        'cb_empleado
        '
        Me.cb_empleado.Enabled = False
        Me.cb_empleado.FormattingEnabled = True
        Me.cb_empleado.Location = New System.Drawing.Point(122, 30)
        Me.cb_empleado.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cb_empleado.Name = "cb_empleado"
        Me.cb_empleado.Size = New System.Drawing.Size(352, 28)
        Me.cb_empleado.TabIndex = 7
        '
        'dtp_final
        '
        Me.dtp_final.Enabled = False
        Me.dtp_final.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtp_final.Location = New System.Drawing.Point(332, 68)
        Me.dtp_final.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dtp_final.Name = "dtp_final"
        Me.dtp_final.Size = New System.Drawing.Size(116, 27)
        Me.dtp_final.TabIndex = 6
        '
        'dtp_inicio
        '
        Me.dtp_inicio.Enabled = False
        Me.dtp_inicio.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtp_inicio.Location = New System.Drawing.Point(122, 68)
        Me.dtp_inicio.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dtp_inicio.Name = "dtp_inicio"
        Me.dtp_inicio.Size = New System.Drawing.Size(113, 27)
        Me.dtp_inicio.TabIndex = 5
        '
        'rb_todos
        '
        Me.rb_todos.AutoSize = True
        Me.rb_todos.Checked = True
        Me.rb_todos.Location = New System.Drawing.Point(8, 139)
        Me.rb_todos.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.rb_todos.Name = "rb_todos"
        Me.rb_todos.Size = New System.Drawing.Size(67, 24)
        Me.rb_todos.TabIndex = 4
        Me.rb_todos.TabStop = True
        Me.rb_todos.Text = "Todos"
        Me.rb_todos.UseVisualStyleBackColor = True
        '
        'rb_fecha
        '
        Me.rb_fecha.AutoSize = True
        Me.rb_fecha.Location = New System.Drawing.Point(8, 68)
        Me.rb_fecha.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.rb_fecha.Name = "rb_fecha"
        Me.rb_fecha.Size = New System.Drawing.Size(97, 24)
        Me.rb_fecha.TabIndex = 3
        Me.rb_fecha.TabStop = True
        Me.rb_fecha.Text = "Fechas De:"
        Me.rb_fecha.UseVisualStyleBackColor = True
        '
        'rb_empleado
        '
        Me.rb_empleado.AutoSize = True
        Me.rb_empleado.Location = New System.Drawing.Point(8, 30)
        Me.rb_empleado.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.rb_empleado.Name = "rb_empleado"
        Me.rb_empleado.Size = New System.Drawing.Size(95, 24)
        Me.rb_empleado.TabIndex = 2
        Me.rb_empleado.TabStop = True
        Me.rb_empleado.Text = "Empleado"
        Me.rb_empleado.UseVisualStyleBackColor = True
        '
        'btn_buscar
        '
        Me.btn_buscar.Image = CType(resources.GetObject("btn_buscar.Image"), System.Drawing.Image)
        Me.btn_buscar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_buscar.Location = New System.Drawing.Point(469, 68)
        Me.btn_buscar.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btn_buscar.Name = "btn_buscar"
        Me.btn_buscar.Size = New System.Drawing.Size(173, 41)
        Me.btn_buscar.TabIndex = 0
        Me.btn_buscar.Text = "Buscar"
        Me.btn_buscar.UseVisualStyleBackColor = True
        '
        'dg_datos
        '
        Me.dg_datos.BackgroundColor = System.Drawing.SystemColors.ControlDarkDark
        Me.dg_datos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg_datos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.idcompra, Me.fecha, Me.articulo, Me.cantidad, Me.costo, Me.empleado})
        Me.dg_datos.GridColor = System.Drawing.SystemColors.ControlDarkDark
        Me.dg_datos.Location = New System.Drawing.Point(16, 264)
        Me.dg_datos.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dg_datos.Name = "dg_datos"
        Me.dg_datos.Size = New System.Drawing.Size(665, 119)
        Me.dg_datos.TabIndex = 3
        '
        'idcompra
        '
        Me.idcompra.HeaderText = "ID"
        Me.idcompra.Name = "idcompra"
        Me.idcompra.ReadOnly = True
        Me.idcompra.Width = 50
        '
        'fecha
        '
        Me.fecha.HeaderText = "Fecha"
        Me.fecha.Name = "fecha"
        Me.fecha.ReadOnly = True
        '
        'articulo
        '
        Me.articulo.HeaderText = "Articulo"
        Me.articulo.Name = "articulo"
        Me.articulo.ReadOnly = True
        Me.articulo.Width = 150
        '
        'cantidad
        '
        Me.cantidad.HeaderText = "Cantidad"
        Me.cantidad.Name = "cantidad"
        Me.cantidad.ReadOnly = True
        '
        'costo
        '
        Me.costo.HeaderText = "Costo"
        Me.costo.Name = "costo"
        Me.costo.ReadOnly = True
        '
        'empleado
        '
        Me.empleado.HeaderText = "Importe"
        Me.empleado.Name = "empleado"
        Me.empleado.ReadOnly = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(513, 402)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 20)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Total:"
        '
        'lbl_total
        '
        Me.lbl_total.AutoSize = True
        Me.lbl_total.Font = New System.Drawing.Font("SimSun-ExtB", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_total.Location = New System.Drawing.Point(590, 402)
        Me.lbl_total.Name = "lbl_total"
        Me.lbl_total.Size = New System.Drawing.Size(49, 20)
        Me.lbl_total.TabIndex = 7
        Me.lbl_total.Text = "0.00"
        '
        'lbl_titulo
        '
        Me.lbl_titulo.BackColor = System.Drawing.Color.LightSeaGreen
        Me.lbl_titulo.Dock = System.Windows.Forms.DockStyle.Top
        Me.lbl_titulo.Font = New System.Drawing.Font("Century Gothic", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_titulo.Location = New System.Drawing.Point(0, 0)
        Me.lbl_titulo.Name = "lbl_titulo"
        Me.lbl_titulo.Size = New System.Drawing.Size(700, 43)
        Me.lbl_titulo.TabIndex = 31
        Me.lbl_titulo.Text = "Consulta De Ventas"
        Me.lbl_titulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Consulta_Ventas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Highlight
        Me.ClientSize = New System.Drawing.Size(700, 467)
        Me.Controls.Add(Me.lbl_titulo)
        Me.Controls.Add(Me.lbl_total)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btn_salir)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.dg_datos)
        Me.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Consulta_Ventas"
        Me.Text = "Consulta_Ventas"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btn_salir As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents cb_idventa As ComboBox
    Friend WithEvents rb_idventa As RadioButton
    Friend WithEvents Label1 As Label
    Friend WithEvents cb_empleado As ComboBox
    Friend WithEvents dtp_final As DateTimePicker
    Friend WithEvents dtp_inicio As DateTimePicker
    Friend WithEvents rb_todos As RadioButton
    Friend WithEvents rb_fecha As RadioButton
    Friend WithEvents rb_empleado As RadioButton
    Friend WithEvents btn_buscar As Button
    Friend WithEvents dg_datos As DataGridView
    Friend WithEvents idcompra As DataGridViewTextBoxColumn
    Friend WithEvents fecha As DataGridViewTextBoxColumn
    Friend WithEvents articulo As DataGridViewTextBoxColumn
    Friend WithEvents cantidad As DataGridViewTextBoxColumn
    Friend WithEvents costo As DataGridViewTextBoxColumn
    Friend WithEvents empleado As DataGridViewTextBoxColumn
    Friend WithEvents Label2 As Label
    Friend WithEvents lbl_total As Label
    Friend WithEvents lbl_titulo As Label
End Class
