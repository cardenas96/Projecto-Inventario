<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Usuarios
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Usuarios))
        Me.lbl_titulo = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txt_id = New System.Windows.Forms.TextBox()
        Me.USUARIOSBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.InventariO_DBDataSet1 = New Sis_Admin_In.INVENTARIO_DBDataSet()
        Me.txt_usuario = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txt_contraseña = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btn_salir = New System.Windows.Forms.Button()
        Me.btn_grabar = New System.Windows.Forms.Button()
        Me.btn_nuevo = New System.Windows.Forms.Button()
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlConnection1 = New System.Data.SqlClient.SqlConnection()
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlDeleteCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlDataAdapter1 = New System.Data.SqlClient.SqlDataAdapter()
        Me.dg_datos = New System.Windows.Forms.DataGridView()
        Me.IdUsuarioDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NUsuarioDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CUsuarioDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cbTipo = New System.Windows.Forms.ComboBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.USUARIOSBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.InventariO_DBDataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lbl_titulo
        '
        Me.lbl_titulo.BackColor = System.Drawing.Color.LightSeaGreen
        Me.lbl_titulo.Dock = System.Windows.Forms.DockStyle.Top
        Me.lbl_titulo.Font = New System.Drawing.Font("Century Gothic", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_titulo.Location = New System.Drawing.Point(0, 0)
        Me.lbl_titulo.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbl_titulo.Name = "lbl_titulo"
        Me.lbl_titulo.Size = New System.Drawing.Size(743, 66)
        Me.lbl_titulo.TabIndex = 21
        Me.lbl_titulo.Text = "Agregar Usuarios"
        Me.lbl_titulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.Sis_Admin_In.My.Resources.Resources.usuario
        Me.PictureBox1.Location = New System.Drawing.Point(16, 71)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(159, 160)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 22
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(218, 90)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(27, 20)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "ID:"
        '
        'txt_id
        '
        Me.txt_id.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.USUARIOSBindingSource, "idUsuario", True))
        Me.txt_id.Enabled = False
        Me.txt_id.Location = New System.Drawing.Point(305, 87)
        Me.txt_id.Name = "txt_id"
        Me.txt_id.Size = New System.Drawing.Size(100, 27)
        Me.txt_id.TabIndex = 24
        '
        'USUARIOSBindingSource
        '
        Me.USUARIOSBindingSource.DataMember = "USUARIOS"
        Me.USUARIOSBindingSource.DataSource = Me.InventariO_DBDataSet1
        '
        'InventariO_DBDataSet1
        '
        Me.InventariO_DBDataSet1.DataSetName = "INVENTARIO_DBDataSet"
        Me.InventariO_DBDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'txt_usuario
        '
        Me.txt_usuario.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.USUARIOSBindingSource, "nUsuario", True))
        Me.txt_usuario.Enabled = False
        Me.txt_usuario.Location = New System.Drawing.Point(305, 129)
        Me.txt_usuario.Name = "txt_usuario"
        Me.txt_usuario.Size = New System.Drawing.Size(202, 27)
        Me.txt_usuario.TabIndex = 26
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(218, 132)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 20)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "Usuario:"
        '
        'txt_contraseña
        '
        Me.txt_contraseña.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.USUARIOSBindingSource, "cUsuario", True))
        Me.txt_contraseña.Enabled = False
        Me.txt_contraseña.Location = New System.Drawing.Point(305, 176)
        Me.txt_contraseña.Name = "txt_contraseña"
        Me.txt_contraseña.Size = New System.Drawing.Size(202, 27)
        Me.txt_contraseña.TabIndex = 28
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(218, 179)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 20)
        Me.Label3.TabIndex = 27
        Me.Label3.Text = "Contraseña:"
        '
        'btn_salir
        '
        Me.btn_salir.Image = CType(resources.GetObject("btn_salir.Image"), System.Drawing.Image)
        Me.btn_salir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_salir.Location = New System.Drawing.Point(552, 336)
        Me.btn_salir.Name = "btn_salir"
        Me.btn_salir.Size = New System.Drawing.Size(177, 48)
        Me.btn_salir.TabIndex = 31
        Me.btn_salir.Text = "Salir"
        Me.btn_salir.UseVisualStyleBackColor = True
        '
        'btn_grabar
        '
        Me.btn_grabar.Enabled = False
        Me.btn_grabar.Image = CType(resources.GetObject("btn_grabar.Image"), System.Drawing.Image)
        Me.btn_grabar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_grabar.Location = New System.Drawing.Point(276, 336)
        Me.btn_grabar.Name = "btn_grabar"
        Me.btn_grabar.Size = New System.Drawing.Size(177, 48)
        Me.btn_grabar.TabIndex = 30
        Me.btn_grabar.Text = "Grabar"
        Me.btn_grabar.UseVisualStyleBackColor = True
        '
        'btn_nuevo
        '
        Me.btn_nuevo.Image = CType(resources.GetObject("btn_nuevo.Image"), System.Drawing.Image)
        Me.btn_nuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_nuevo.Location = New System.Drawing.Point(10, 336)
        Me.btn_nuevo.Name = "btn_nuevo"
        Me.btn_nuevo.Size = New System.Drawing.Size(177, 48)
        Me.btn_nuevo.TabIndex = 29
        Me.btn_nuevo.Text = "Nuevo"
        Me.btn_nuevo.UseVisualStyleBackColor = True
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT        USUARIOS.*" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "FROM            USUARIOS"
        Me.SqlSelectCommand1.Connection = Me.SqlConnection1
        '
        'SqlConnection1
        '
        Me.SqlConnection1.ConnectionString = "Data Source=CARDENAS-PC;Initial Catalog=INVENTARIO_DB;Integrated Security=True"
        Me.SqlConnection1.FireInfoMessageEventOnUserErrors = False
        '
        'SqlInsertCommand1
        '
        Me.SqlInsertCommand1.CommandText = "INSERT INTO [USUARIOS] ([idUsuario], [nUsuario], [cUsuario]) VALUES (@idUsuario, " &
    "@nUsuario, @cUsuario);" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "SELECT idUsuario, nUsuario, cUsuario FROM USUARIOS WHERE" &
    " (idUsuario = @idUsuario)"
        Me.SqlInsertCommand1.Connection = Me.SqlConnection1
        Me.SqlInsertCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@idUsuario", System.Data.SqlDbType.Int, 0, "idUsuario"), New System.Data.SqlClient.SqlParameter("@nUsuario", System.Data.SqlDbType.VarChar, 0, "nUsuario"), New System.Data.SqlClient.SqlParameter("@cUsuario", System.Data.SqlDbType.VarChar, 0, "cUsuario")})
        '
        'SqlUpdateCommand1
        '
        Me.SqlUpdateCommand1.CommandText = resources.GetString("SqlUpdateCommand1.CommandText")
        Me.SqlUpdateCommand1.Connection = Me.SqlConnection1
        Me.SqlUpdateCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@idUsuario", System.Data.SqlDbType.Int, 0, "idUsuario"), New System.Data.SqlClient.SqlParameter("@nUsuario", System.Data.SqlDbType.VarChar, 0, "nUsuario"), New System.Data.SqlClient.SqlParameter("@cUsuario", System.Data.SqlDbType.VarChar, 0, "cUsuario"), New System.Data.SqlClient.SqlParameter("@Original_idUsuario", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "idUsuario", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_nUsuario", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "nUsuario", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_cUsuario", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "cUsuario", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlDeleteCommand1
        '
        Me.SqlDeleteCommand1.CommandText = "DELETE FROM [USUARIOS] WHERE (([idUsuario] = @Original_idUsuario) AND ([nUsuario]" &
    " = @Original_nUsuario) AND ([cUsuario] = @Original_cUsuario))"
        Me.SqlDeleteCommand1.Connection = Me.SqlConnection1
        Me.SqlDeleteCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_idUsuario", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "idUsuario", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_nUsuario", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "nUsuario", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_cUsuario", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "cUsuario", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlDataAdapter1
        '
        Me.SqlDataAdapter1.DeleteCommand = Me.SqlDeleteCommand1
        Me.SqlDataAdapter1.InsertCommand = Me.SqlInsertCommand1
        Me.SqlDataAdapter1.SelectCommand = Me.SqlSelectCommand1
        Me.SqlDataAdapter1.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "USUARIOS", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("idUsuario", "idUsuario"), New System.Data.Common.DataColumnMapping("nUsuario", "nUsuario"), New System.Data.Common.DataColumnMapping("cUsuario", "cUsuario")})})
        Me.SqlDataAdapter1.UpdateCommand = Me.SqlUpdateCommand1
        '
        'dg_datos
        '
        Me.dg_datos.AutoGenerateColumns = False
        Me.dg_datos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg_datos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IdUsuarioDataGridViewTextBoxColumn, Me.NUsuarioDataGridViewTextBoxColumn, Me.CUsuarioDataGridViewTextBoxColumn})
        Me.dg_datos.DataSource = Me.USUARIOSBindingSource
        Me.dg_datos.Location = New System.Drawing.Point(213, 227)
        Me.dg_datos.Name = "dg_datos"
        Me.dg_datos.Size = New System.Drawing.Size(506, 95)
        Me.dg_datos.TabIndex = 32
        '
        'IdUsuarioDataGridViewTextBoxColumn
        '
        Me.IdUsuarioDataGridViewTextBoxColumn.DataPropertyName = "idUsuario"
        Me.IdUsuarioDataGridViewTextBoxColumn.HeaderText = "ID"
        Me.IdUsuarioDataGridViewTextBoxColumn.Name = "IdUsuarioDataGridViewTextBoxColumn"
        Me.IdUsuarioDataGridViewTextBoxColumn.Width = 80
        '
        'NUsuarioDataGridViewTextBoxColumn
        '
        Me.NUsuarioDataGridViewTextBoxColumn.DataPropertyName = "nUsuario"
        Me.NUsuarioDataGridViewTextBoxColumn.HeaderText = "Nombre"
        Me.NUsuarioDataGridViewTextBoxColumn.Name = "NUsuarioDataGridViewTextBoxColumn"
        Me.NUsuarioDataGridViewTextBoxColumn.Width = 175
        '
        'CUsuarioDataGridViewTextBoxColumn
        '
        Me.CUsuarioDataGridViewTextBoxColumn.DataPropertyName = "cUsuario"
        Me.CUsuarioDataGridViewTextBoxColumn.HeaderText = "Contraseña"
        Me.CUsuarioDataGridViewTextBoxColumn.Name = "CUsuarioDataGridViewTextBoxColumn"
        Me.CUsuarioDataGridViewTextBoxColumn.Width = 175
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(548, 132)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(115, 20)
        Me.Label4.TabIndex = 33
        Me.Label4.Text = "Tipo de usuario:"
        '
        'cbTipo
        '
        Me.cbTipo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbTipo.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cbTipo.FormattingEnabled = True
        Me.cbTipo.Items.AddRange(New Object() {"Administrador", "Gerente", "Empleado"})
        Me.cbTipo.Location = New System.Drawing.Point(552, 171)
        Me.cbTipo.Name = "cbTipo"
        Me.cbTipo.Size = New System.Drawing.Size(167, 28)
        Me.cbTipo.TabIndex = 34
        '
        'Usuarios
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(743, 396)
        Me.Controls.Add(Me.cbTipo)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.dg_datos)
        Me.Controls.Add(Me.btn_salir)
        Me.Controls.Add(Me.btn_grabar)
        Me.Controls.Add(Me.btn_nuevo)
        Me.Controls.Add(Me.txt_contraseña)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txt_usuario)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txt_id)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.lbl_titulo)
        Me.Font = New System.Drawing.Font("Segoe UI", 11.0!)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Usuarios"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Usuarios"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.USUARIOSBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.InventariO_DBDataSet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lbl_titulo As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txt_id As TextBox
    Friend WithEvents txt_usuario As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txt_contraseña As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents btn_salir As Button
    Friend WithEvents btn_grabar As Button
    Friend WithEvents btn_nuevo As Button
    Friend WithEvents SqlSelectCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlConnection1 As SqlClient.SqlConnection
    Friend WithEvents SqlInsertCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlUpdateCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlDeleteCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlDataAdapter1 As SqlClient.SqlDataAdapter
    Friend WithEvents dg_datos As DataGridView
    Friend WithEvents InventariO_DBDataSet1 As INVENTARIO_DBDataSet
    Friend WithEvents USUARIOSBindingSource As BindingSource
    Friend WithEvents IdUsuarioDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents NUsuarioDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents CUsuarioDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Label4 As Label
    Friend WithEvents cbTipo As ComboBox
End Class
