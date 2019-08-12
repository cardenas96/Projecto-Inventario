<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Consulta_Compra
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Consulta_Compra))
        Me.dg_datos = New System.Windows.Forms.DataGridView()
        Me.idcompra = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.proveedor = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.fecha = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.articulo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cantidad = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.costo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cb_idcompra = New System.Windows.Forms.ComboBox()
        Me.rb_idcompra = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cb_proveedor = New System.Windows.Forms.ComboBox()
        Me.dtp_final = New System.Windows.Forms.DateTimePicker()
        Me.dtp_inicio = New System.Windows.Forms.DateTimePicker()
        Me.rb_todos = New System.Windows.Forms.RadioButton()
        Me.rb_fecha = New System.Windows.Forms.RadioButton()
        Me.rb_proveedor = New System.Windows.Forms.RadioButton()
        Me.btn_buscar = New System.Windows.Forms.Button()
        Me.btn_salir = New System.Windows.Forms.Button()
        Me.lbl_titulo = New System.Windows.Forms.Label()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dg_datos
        '
        Me.dg_datos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg_datos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.idcompra, Me.proveedor, Me.fecha, Me.articulo, Me.cantidad, Me.costo})
        Me.dg_datos.Location = New System.Drawing.Point(18, 251)
        Me.dg_datos.Name = "dg_datos"
        Me.dg_datos.Size = New System.Drawing.Size(682, 135)
        Me.dg_datos.TabIndex = 0
        '
        'idcompra
        '
        Me.idcompra.HeaderText = "ID"
        Me.idcompra.Name = "idcompra"
        Me.idcompra.ReadOnly = True
        Me.idcompra.Width = 50
        '
        'proveedor
        '
        Me.proveedor.HeaderText = "Proveedor"
        Me.proveedor.Name = "proveedor"
        Me.proveedor.ReadOnly = True
        Me.proveedor.Width = 120
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
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cb_idcompra)
        Me.GroupBox1.Controls.Add(Me.rb_idcompra)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.cb_proveedor)
        Me.GroupBox1.Controls.Add(Me.dtp_final)
        Me.GroupBox1.Controls.Add(Me.dtp_inicio)
        Me.GroupBox1.Controls.Add(Me.rb_todos)
        Me.GroupBox1.Controls.Add(Me.rb_fecha)
        Me.GroupBox1.Controls.Add(Me.rb_proveedor)
        Me.GroupBox1.Controls.Add(Me.btn_buscar)
        Me.GroupBox1.Location = New System.Drawing.Point(18, 58)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(681, 174)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Consulta De Compras"
        '
        'cb_idcompra
        '
        Me.cb_idcompra.Enabled = False
        Me.cb_idcompra.FormattingEnabled = True
        Me.cb_idcompra.Location = New System.Drawing.Point(162, 107)
        Me.cb_idcompra.Name = "cb_idcompra"
        Me.cb_idcompra.Size = New System.Drawing.Size(167, 28)
        Me.cb_idcompra.TabIndex = 10
        '
        'rb_idcompra
        '
        Me.rb_idcompra.AutoSize = True
        Me.rb_idcompra.Location = New System.Drawing.Point(27, 108)
        Me.rb_idcompra.Name = "rb_idcompra"
        Me.rb_idcompra.Size = New System.Drawing.Size(110, 24)
        Me.rb_idcompra.TabIndex = 9
        Me.rb_idcompra.TabStop = True
        Me.rb_idcompra.Text = "No. Compra:"
        Me.rb_idcompra.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(279, 71)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 20)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Hasta:"
        '
        'cb_proveedor
        '
        Me.cb_proveedor.Enabled = False
        Me.cb_proveedor.FormattingEnabled = True
        Me.cb_proveedor.Location = New System.Drawing.Point(162, 25)
        Me.cb_proveedor.Name = "cb_proveedor"
        Me.cb_proveedor.Size = New System.Drawing.Size(285, 28)
        Me.cb_proveedor.TabIndex = 7
        '
        'dtp_final
        '
        Me.dtp_final.Enabled = False
        Me.dtp_final.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtp_final.Location = New System.Drawing.Point(331, 66)
        Me.dtp_final.Name = "dtp_final"
        Me.dtp_final.Size = New System.Drawing.Size(116, 27)
        Me.dtp_final.TabIndex = 6
        '
        'dtp_inicio
        '
        Me.dtp_inicio.Enabled = False
        Me.dtp_inicio.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtp_inicio.Location = New System.Drawing.Point(162, 66)
        Me.dtp_inicio.Name = "dtp_inicio"
        Me.dtp_inicio.Size = New System.Drawing.Size(116, 27)
        Me.dtp_inicio.TabIndex = 5
        '
        'rb_todos
        '
        Me.rb_todos.AutoSize = True
        Me.rb_todos.Checked = True
        Me.rb_todos.Location = New System.Drawing.Point(27, 144)
        Me.rb_todos.Name = "rb_todos"
        Me.rb_todos.Size = New System.Drawing.Size(68, 24)
        Me.rb_todos.TabIndex = 4
        Me.rb_todos.TabStop = True
        Me.rb_todos.Text = "Todos"
        Me.rb_todos.UseVisualStyleBackColor = True
        '
        'rb_fecha
        '
        Me.rb_fecha.AutoSize = True
        Me.rb_fecha.Location = New System.Drawing.Point(27, 69)
        Me.rb_fecha.Name = "rb_fecha"
        Me.rb_fecha.Size = New System.Drawing.Size(97, 24)
        Me.rb_fecha.TabIndex = 3
        Me.rb_fecha.TabStop = True
        Me.rb_fecha.Text = "Fechas De:"
        Me.rb_fecha.UseVisualStyleBackColor = True
        '
        'rb_proveedor
        '
        Me.rb_proveedor.AutoSize = True
        Me.rb_proveedor.Location = New System.Drawing.Point(27, 29)
        Me.rb_proveedor.Name = "rb_proveedor"
        Me.rb_proveedor.Size = New System.Drawing.Size(95, 24)
        Me.rb_proveedor.TabIndex = 2
        Me.rb_proveedor.TabStop = True
        Me.rb_proveedor.Text = "Proveedor"
        Me.rb_proveedor.UseVisualStyleBackColor = True
        '
        'btn_buscar
        '
        Me.btn_buscar.Image = CType(resources.GetObject("btn_buscar.Image"), System.Drawing.Image)
        Me.btn_buscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btn_buscar.Location = New System.Drawing.Point(482, 59)
        Me.btn_buscar.Name = "btn_buscar"
        Me.btn_buscar.Size = New System.Drawing.Size(153, 47)
        Me.btn_buscar.TabIndex = 0
        Me.btn_buscar.Text = "Buscar"
        Me.btn_buscar.UseVisualStyleBackColor = True
        '
        'btn_salir
        '
        Me.btn_salir.Image = CType(resources.GetObject("btn_salir.Image"), System.Drawing.Image)
        Me.btn_salir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_salir.Location = New System.Drawing.Point(279, 392)
        Me.btn_salir.Name = "btn_salir"
        Me.btn_salir.Size = New System.Drawing.Size(186, 50)
        Me.btn_salir.TabIndex = 2
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
        Me.lbl_titulo.Size = New System.Drawing.Size(726, 43)
        Me.lbl_titulo.TabIndex = 31
        Me.lbl_titulo.Text = "Consulta de Compras"
        Me.lbl_titulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Consulta_Compra
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Highlight
        Me.ClientSize = New System.Drawing.Size(726, 454)
        Me.Controls.Add(Me.lbl_titulo)
        Me.Controls.Add(Me.btn_salir)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.dg_datos)
        Me.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Consulta_Compra"
        Me.Text = "Consulta_Compra"
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents dg_datos As DataGridView
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents rb_fecha As RadioButton
    Friend WithEvents rb_proveedor As RadioButton
    Friend WithEvents btn_buscar As Button
    Friend WithEvents idcompra As DataGridViewTextBoxColumn
    Friend WithEvents proveedor As DataGridViewTextBoxColumn
    Friend WithEvents fecha As DataGridViewTextBoxColumn
    Friend WithEvents articulo As DataGridViewTextBoxColumn
    Friend WithEvents cantidad As DataGridViewTextBoxColumn
    Friend WithEvents costo As DataGridViewTextBoxColumn
    Friend WithEvents cb_idcompra As ComboBox
    Friend WithEvents rb_idcompra As RadioButton
    Friend WithEvents Label1 As Label
    Friend WithEvents cb_proveedor As ComboBox
    Friend WithEvents dtp_final As DateTimePicker
    Friend WithEvents dtp_inicio As DateTimePicker
    Friend WithEvents rb_todos As RadioButton
    Friend WithEvents btn_salir As Button
    Friend WithEvents lbl_titulo As Label
End Class
