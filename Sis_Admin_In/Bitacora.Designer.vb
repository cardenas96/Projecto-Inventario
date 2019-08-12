<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Bitacora
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
        Me.dg_datos = New System.Windows.Forms.DataGridView()
        Me.No = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Erro = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Descripcion = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Formula = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Linea = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Usuario = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Fecha = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Hora = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.lbl_titulo = New System.Windows.Forms.Label()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dg_datos
        '
        Me.dg_datos.BackgroundColor = System.Drawing.SystemColors.ControlDarkDark
        Me.dg_datos.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.Sunken
        Me.dg_datos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg_datos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.No, Me.Erro, Me.Descripcion, Me.Formula, Me.Linea, Me.Usuario, Me.Fecha, Me.Hora})
        Me.dg_datos.Location = New System.Drawing.Point(5, 59)
        Me.dg_datos.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dg_datos.Name = "dg_datos"
        Me.dg_datos.Size = New System.Drawing.Size(830, 252)
        Me.dg_datos.TabIndex = 0
        '
        'No
        '
        Me.No.HeaderText = "No Falla"
        Me.No.Name = "No"
        Me.No.ReadOnly = True
        Me.No.Width = 80
        '
        'Erro
        '
        Me.Erro.HeaderText = "Error de Corrida"
        Me.Erro.Name = "Erro"
        Me.Erro.ReadOnly = True
        Me.Erro.Width = 130
        '
        'Descripcion
        '
        Me.Descripcion.HeaderText = "Descripcion"
        Me.Descripcion.Name = "Descripcion"
        Me.Descripcion.ReadOnly = True
        Me.Descripcion.Width = 120
        '
        'Formula
        '
        Me.Formula.HeaderText = "Formulario"
        Me.Formula.Name = "Formula"
        Me.Formula.ReadOnly = True
        '
        'Linea
        '
        Me.Linea.HeaderText = "Linea"
        Me.Linea.Name = "Linea"
        Me.Linea.ReadOnly = True
        Me.Linea.Width = 60
        '
        'Usuario
        '
        Me.Usuario.HeaderText = "Usuario"
        Me.Usuario.Name = "Usuario"
        '
        'Fecha
        '
        Me.Fecha.HeaderText = "Fecha"
        Me.Fecha.Name = "Fecha"
        Me.Fecha.ReadOnly = True
        '
        'Hora
        '
        Me.Hora.HeaderText = "Hora"
        Me.Hora.Name = "Hora"
        Me.Hora.ReadOnly = True
        Me.Hora.Width = 80
        '
        'lbl_titulo
        '
        Me.lbl_titulo.BackColor = System.Drawing.Color.LightSeaGreen
        Me.lbl_titulo.Dock = System.Windows.Forms.DockStyle.Top
        Me.lbl_titulo.Font = New System.Drawing.Font("Century Gothic", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_titulo.Location = New System.Drawing.Point(0, 0)
        Me.lbl_titulo.Name = "lbl_titulo"
        Me.lbl_titulo.Size = New System.Drawing.Size(847, 43)
        Me.lbl_titulo.TabIndex = 32
        Me.lbl_titulo.Text = "Consulta de Bitacora"
        Me.lbl_titulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Bitacora
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Highlight
        Me.ClientSize = New System.Drawing.Size(847, 327)
        Me.Controls.Add(Me.lbl_titulo)
        Me.Controls.Add(Me.dg_datos)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "Bitacora"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Bitacora"
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents dg_datos As DataGridView
    Friend WithEvents No As DataGridViewTextBoxColumn
    Friend WithEvents Erro As DataGridViewTextBoxColumn
    Friend WithEvents Descripcion As DataGridViewTextBoxColumn
    Friend WithEvents Formula As DataGridViewTextBoxColumn
    Friend WithEvents Linea As DataGridViewTextBoxColumn
    Friend WithEvents Usuario As DataGridViewTextBoxColumn
    Friend WithEvents Fecha As DataGridViewTextBoxColumn
    Friend WithEvents Hora As DataGridViewTextBoxColumn
    Friend WithEvents lbl_titulo As Label
End Class
