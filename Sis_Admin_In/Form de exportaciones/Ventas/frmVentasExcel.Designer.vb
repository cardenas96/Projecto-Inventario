<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmVentasExcel
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVentasExcel))
        Me.dgTablas = New System.Windows.Forms.DataGridView()
        Me.btn_salir = New System.Windows.Forms.Button()
        Me.btnGenerar = New System.Windows.Forms.Button()
        Me.dgFinal = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.seleccionar = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.idVentaV = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.nombreE = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.fechaV = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.nombreP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CantidadD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PrecioD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.marcaP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.unidadP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.totalV = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.dgTablas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgFinal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgTablas
        '
        Me.dgTablas.AllowUserToAddRows = False
        Me.dgTablas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgTablas.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column2, Me.Column3, Me.seleccionar})
        Me.dgTablas.Location = New System.Drawing.Point(35, 59)
        Me.dgTablas.Name = "dgTablas"
        Me.dgTablas.Size = New System.Drawing.Size(354, 266)
        Me.dgTablas.TabIndex = 0
        '
        'btn_salir
        '
        Me.btn_salir.Image = CType(resources.GetObject("btn_salir.Image"), System.Drawing.Image)
        Me.btn_salir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_salir.Location = New System.Drawing.Point(275, 353)
        Me.btn_salir.Name = "btn_salir"
        Me.btn_salir.Size = New System.Drawing.Size(177, 48)
        Me.btn_salir.TabIndex = 30
        Me.btn_salir.Text = "Salir"
        Me.btn_salir.UseVisualStyleBackColor = True
        '
        'btnGenerar
        '
        Me.btnGenerar.Image = CType(resources.GetObject("btnGenerar.Image"), System.Drawing.Image)
        Me.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnGenerar.Location = New System.Drawing.Point(35, 353)
        Me.btnGenerar.Name = "btnGenerar"
        Me.btnGenerar.Size = New System.Drawing.Size(177, 48)
        Me.btnGenerar.TabIndex = 29
        Me.btnGenerar.Text = "Exportar a excel"
        Me.btnGenerar.UseVisualStyleBackColor = True
        '
        'dgFinal
        '
        Me.dgFinal.AllowUserToAddRows = False
        Me.dgFinal.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgFinal.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.idVentaV, Me.nombreE, Me.fechaV, Me.nombreP, Me.CantidadD, Me.PrecioD, Me.marcaP, Me.unidadP, Me.totalV})
        Me.dgFinal.Location = New System.Drawing.Point(422, 59)
        Me.dgFinal.Name = "dgFinal"
        Me.dgFinal.Size = New System.Drawing.Size(474, 266)
        Me.dgFinal.TabIndex = 31
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(416, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(489, 31)
        Me.Label1.TabIndex = 32
        Me.Label1.Text = "Estructura previa del archivo EXCEL"
        '
        'Column2
        '
        Me.Column2.HeaderText = "Tabla"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        '
        'Column3
        '
        Me.Column3.HeaderText = "Campo"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        '
        'seleccionar
        '
        Me.seleccionar.HeaderText = "seleccionar"
        Me.seleccionar.Name = "seleccionar"
        '
        'idVentaV
        '
        Me.idVentaV.HeaderText = "idVenta"
        Me.idVentaV.Name = "idVentaV"
        '
        'nombreE
        '
        Me.nombreE.HeaderText = "Empleado"
        Me.nombreE.Name = "nombreE"
        Me.nombreE.ReadOnly = True
        '
        'fechaV
        '
        Me.fechaV.HeaderText = "fecha"
        Me.fechaV.Name = "fechaV"
        '
        'nombreP
        '
        Me.nombreP.HeaderText = "Producto"
        Me.nombreP.Name = "nombreP"
        Me.nombreP.ReadOnly = True
        '
        'CantidadD
        '
        Me.CantidadD.HeaderText = "Cantidad"
        Me.CantidadD.Name = "CantidadD"
        '
        'PrecioD
        '
        Me.PrecioD.HeaderText = "Precio"
        Me.PrecioD.Name = "PrecioD"
        Me.PrecioD.ReadOnly = True
        '
        'marcaP
        '
        Me.marcaP.HeaderText = "Marca"
        Me.marcaP.Name = "marcaP"
        Me.marcaP.ReadOnly = True
        '
        'unidadP
        '
        Me.unidadP.HeaderText = "unidad"
        Me.unidadP.Name = "unidadP"
        Me.unidadP.ReadOnly = True
        '
        'totalV
        '
        Me.totalV.HeaderText = "Total"
        Me.totalV.Name = "totalV"
        '
        'frmVentasExcel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(924, 404)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgFinal)
        Me.Controls.Add(Me.btn_salir)
        Me.Controls.Add(Me.btnGenerar)
        Me.Controls.Add(Me.dgTablas)
        Me.Name = "frmVentasExcel"
        Me.Text = "frmVentasExcel"
        CType(Me.dgTablas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgFinal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dgTablas As DataGridView
    Friend WithEvents btn_salir As Button
    Friend WithEvents btnGenerar As Button
    Friend WithEvents dgFinal As DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents Column2 As DataGridViewTextBoxColumn
    Friend WithEvents Column3 As DataGridViewTextBoxColumn
    Friend WithEvents seleccionar As DataGridViewCheckBoxColumn
    Friend WithEvents idVentaV As DataGridViewTextBoxColumn
    Friend WithEvents nombreE As DataGridViewTextBoxColumn
    Friend WithEvents fechaV As DataGridViewTextBoxColumn
    Friend WithEvents nombreP As DataGridViewTextBoxColumn
    Friend WithEvents CantidadD As DataGridViewTextBoxColumn
    Friend WithEvents PrecioD As DataGridViewTextBoxColumn
    Friend WithEvents marcaP As DataGridViewTextBoxColumn
    Friend WithEvents unidadP As DataGridViewTextBoxColumn
    Friend WithEvents totalV As DataGridViewTextBoxColumn
End Class
