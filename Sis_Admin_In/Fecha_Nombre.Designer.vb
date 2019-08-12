<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Fecha_Nombre
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Fecha_Nombre))
        Me.dtp_final = New System.Windows.Forms.DateTimePicker()
        Me.dtp_inicio = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.btn_aceptar = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.gb_nombre = New System.Windows.Forms.GroupBox()
        Me.lbl_nombre = New System.Windows.Forms.Label()
        Me.cb_nombre = New System.Windows.Forms.ComboBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.gb_nombre.SuspendLayout()
        Me.SuspendLayout()
        '
        'dtp_final
        '
        Me.dtp_final.Location = New System.Drawing.Point(161, 87)
        Me.dtp_final.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtp_final.Name = "dtp_final"
        Me.dtp_final.Size = New System.Drawing.Size(315, 26)
        Me.dtp_final.TabIndex = 10
        '
        'dtp_inicio
        '
        Me.dtp_inicio.Location = New System.Drawing.Point(161, 30)
        Me.dtp_inicio.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtp_inicio.Name = "dtp_inicio"
        Me.dtp_inicio.Size = New System.Drawing.Size(311, 26)
        Me.dtp_inicio.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 91)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(145, 19)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Seleccione Fecha Final:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 34)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(149, 19)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Seleccione Fecha Inicio:"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(547, 18)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(129, 142)
        Me.PictureBox1.TabIndex = 11
        Me.PictureBox1.TabStop = False
        '
        'btn_aceptar
        '
        Me.btn_aceptar.Image = CType(resources.GetObject("btn_aceptar.Image"), System.Drawing.Image)
        Me.btn_aceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_aceptar.Location = New System.Drawing.Point(249, 278)
        Me.btn_aceptar.Margin = New System.Windows.Forms.Padding(4, 6, 4, 6)
        Me.btn_aceptar.Name = "btn_aceptar"
        Me.btn_aceptar.Size = New System.Drawing.Size(255, 45)
        Me.btn_aceptar.TabIndex = 6
        Me.btn_aceptar.Text = "Aceptar"
        Me.btn_aceptar.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dtp_final)
        Me.GroupBox1.Controls.Add(Me.dtp_inicio)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(18, 18)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(500, 142)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Seleccion De Periodo"
        '
        'gb_nombre
        '
        Me.gb_nombre.BackColor = System.Drawing.SystemColors.Highlight
        Me.gb_nombre.Controls.Add(Me.cb_nombre)
        Me.gb_nombre.Controls.Add(Me.lbl_nombre)
        Me.gb_nombre.Location = New System.Drawing.Point(18, 183)
        Me.gb_nombre.Name = "gb_nombre"
        Me.gb_nombre.Size = New System.Drawing.Size(499, 77)
        Me.gb_nombre.TabIndex = 13
        Me.gb_nombre.TabStop = False
        Me.gb_nombre.Text = "GroupBox2"
        '
        'lbl_nombre
        '
        Me.lbl_nombre.AutoSize = True
        Me.lbl_nombre.Location = New System.Drawing.Point(22, 33)
        Me.lbl_nombre.Name = "lbl_nombre"
        Me.lbl_nombre.Size = New System.Drawing.Size(62, 19)
        Me.lbl_nombre.TabIndex = 0
        Me.lbl_nombre.Text = "Nombre:"
        '
        'cb_nombre
        '
        Me.cb_nombre.FormattingEnabled = True
        Me.cb_nombre.Location = New System.Drawing.Point(90, 30)
        Me.cb_nombre.Name = "cb_nombre"
        Me.cb_nombre.Size = New System.Drawing.Size(359, 27)
        Me.cb_nombre.TabIndex = 1
        '
        'Fecha_Nombre
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.MenuHighlight
        Me.ClientSize = New System.Drawing.Size(743, 339)
        Me.Controls.Add(Me.gb_nombre)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.btn_aceptar)
        Me.Font = New System.Drawing.Font("Segoe UI", 10.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "Fecha_Nombre"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Fecha_Nombre"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.gb_nombre.ResumeLayout(False)
        Me.gb_nombre.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents dtp_final As DateTimePicker
    Friend WithEvents dtp_inicio As DateTimePicker
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents btn_aceptar As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents gb_nombre As GroupBox
    Friend WithEvents cb_nombre As ComboBox
    Friend WithEvents lbl_nombre As Label
End Class
