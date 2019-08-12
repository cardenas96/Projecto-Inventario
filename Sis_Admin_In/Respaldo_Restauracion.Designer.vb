<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Respaldo_Restauracion
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Respaldo_Restauracion))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btn_respaldo = New System.Windows.Forms.Button()
        Me.lbl_ruta = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btn_seleccion = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btn_restaura = New System.Windows.Forms.Button()
        Me.lbl_ruta_a = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btn_seleccionar = New System.Windows.Forms.Button()
        Me.sfd_grabar = New System.Windows.Forms.SaveFileDialog()
        Me.ofd_abrir = New System.Windows.Forms.OpenFileDialog()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btn_respaldo)
        Me.GroupBox1.Controls.Add(Me.lbl_ruta)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.btn_seleccion)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(586, 122)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Crear Respaldo De Base De Datos"
        '
        'btn_respaldo
        '
        Me.btn_respaldo.Enabled = False
        Me.btn_respaldo.Image = CType(resources.GetObject("btn_respaldo.Image"), System.Drawing.Image)
        Me.btn_respaldo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_respaldo.Location = New System.Drawing.Point(316, 60)
        Me.btn_respaldo.Name = "btn_respaldo"
        Me.btn_respaldo.Size = New System.Drawing.Size(234, 46)
        Me.btn_respaldo.TabIndex = 3
        Me.btn_respaldo.Text = "Crear Respaldo"
        Me.btn_respaldo.UseVisualStyleBackColor = True
        '
        'lbl_ruta
        '
        Me.lbl_ruta.AutoSize = True
        Me.lbl_ruta.Location = New System.Drawing.Point(135, 29)
        Me.lbl_ruta.Name = "lbl_ruta"
        Me.lbl_ruta.Size = New System.Drawing.Size(26, 19)
        Me.lbl_ruta.TabIndex = 2
        Me.lbl_ruta.Text = "C:\"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(116, 19)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Ruta de Creación:"
        '
        'btn_seleccion
        '
        Me.btn_seleccion.Image = CType(resources.GetObject("btn_seleccion.Image"), System.Drawing.Image)
        Me.btn_seleccion.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_seleccion.Location = New System.Drawing.Point(17, 58)
        Me.btn_seleccion.Name = "btn_seleccion"
        Me.btn_seleccion.Size = New System.Drawing.Size(246, 49)
        Me.btn_seleccion.TabIndex = 0
        Me.btn_seleccion.Text = "Seleccionar Ruta"
        Me.btn_seleccion.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btn_restaura)
        Me.GroupBox2.Controls.Add(Me.lbl_ruta_a)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.btn_seleccionar)
        Me.GroupBox2.Location = New System.Drawing.Point(11, 149)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(586, 137)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Restaurar Base De Datos Anterior"
        '
        'btn_restaura
        '
        Me.btn_restaura.Enabled = False
        Me.btn_restaura.Image = CType(resources.GetObject("btn_restaura.Image"), System.Drawing.Image)
        Me.btn_restaura.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_restaura.Location = New System.Drawing.Point(317, 68)
        Me.btn_restaura.Name = "btn_restaura"
        Me.btn_restaura.Size = New System.Drawing.Size(234, 46)
        Me.btn_restaura.TabIndex = 7
        Me.btn_restaura.Text = "Restaurar Base De Datos"
        Me.btn_restaura.UseVisualStyleBackColor = True
        '
        'lbl_ruta_a
        '
        Me.lbl_ruta_a.AutoSize = True
        Me.lbl_ruta_a.Location = New System.Drawing.Point(136, 37)
        Me.lbl_ruta_a.Name = "lbl_ruta_a"
        Me.lbl_ruta_a.Size = New System.Drawing.Size(26, 19)
        Me.lbl_ruta_a.TabIndex = 6
        Me.lbl_ruta_a.Text = "C:\"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(14, 37)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(118, 19)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Ruta de Respaldo:"
        '
        'btn_seleccionar
        '
        Me.btn_seleccionar.Image = CType(resources.GetObject("btn_seleccionar.Image"), System.Drawing.Image)
        Me.btn_seleccionar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_seleccionar.Location = New System.Drawing.Point(18, 66)
        Me.btn_seleccionar.Name = "btn_seleccionar"
        Me.btn_seleccionar.Size = New System.Drawing.Size(246, 49)
        Me.btn_seleccionar.TabIndex = 4
        Me.btn_seleccionar.Text = "Seleccionar Ruta"
        Me.btn_seleccionar.UseVisualStyleBackColor = True
        '
        'ofd_abrir
        '
        Me.ofd_abrir.Title = "Seleccione Respaldo"
        '
        'Respaldo_Restauracion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Highlight
        Me.ClientSize = New System.Drawing.Size(612, 301)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Respaldo_Restauracion"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Respaldo & Restauracion"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents lbl_ruta As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents btn_seleccion As Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents btn_respaldo As Button
    Friend WithEvents btn_restaura As Button
    Friend WithEvents lbl_ruta_a As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents btn_seleccionar As Button
    Friend WithEvents sfd_grabar As SaveFileDialog
    Friend WithEvents ofd_abrir As OpenFileDialog
End Class
