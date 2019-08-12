<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Proveedor
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Proveedor))
        Me.txt_domicilio = New System.Windows.Forms.TextBox()
        Me.PROVEEDORBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.INVENTARIO_DBDataSet = New Sis_Admin_In.INVENTARIO_DBDataSet()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txt_telefono = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txt_colonia = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txt_contacto = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txt_id = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txt_ciudad = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txt_fax = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txt_cp = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cmd_siguiente = New System.Windows.Forms.Button()
        Me.cmd_ultimo = New System.Windows.Forms.Button()
        Me.cmd_anterior = New System.Windows.Forms.Button()
        Me.cmd_inicio = New System.Windows.Forms.Button()
        Me.btn_salir = New System.Windows.Forms.Button()
        Me.btn_grabar = New System.Windows.Forms.Button()
        Me.btn_nuevo = New System.Windows.Forms.Button()
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlConnection1 = New System.Data.SqlClient.SqlConnection()
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlDeleteCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlDataAdapter1 = New System.Data.SqlClient.SqlDataAdapter()
        Me.PROVEEDORTableAdapter = New Sis_Admin_In.INVENTARIO_DBDataSetTableAdapters.PROVEEDORTableAdapter()
        Me.txt_nombre = New System.Windows.Forms.TextBox()
        Me.dg_datos = New System.Windows.Forms.DataGridView()
        Me.btn_modificar = New System.Windows.Forms.Button()
        Me.btn_aceptar = New System.Windows.Forms.Button()
        Me.IdProveedorDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NombreDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ContactoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DireccionDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CiudadDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColoniaDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TelefonoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FaxDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CpDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.PROVEEDORBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.INVENTARIO_DBDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txt_domicilio
        '
        Me.txt_domicilio.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PROVEEDORBindingSource, "direccion", True))
        Me.txt_domicilio.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_domicilio.Location = New System.Drawing.Point(93, 203)
        Me.txt_domicilio.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_domicilio.Name = "txt_domicilio"
        Me.txt_domicilio.Size = New System.Drawing.Size(245, 25)
        Me.txt_domicilio.TabIndex = 10
        '
        'PROVEEDORBindingSource
        '
        Me.PROVEEDORBindingSource.DataMember = "PROVEEDOR"
        Me.PROVEEDORBindingSource.DataSource = Me.INVENTARIO_DBDataSet
        '
        'INVENTARIO_DBDataSet
        '
        Me.INVENTARIO_DBDataSet.DataSetName = "INVENTARIO_DBDataSet"
        Me.INVENTARIO_DBDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(15, 204)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(75, 20)
        Me.Label8.TabIndex = 34
        Me.Label8.Text = "Dirección:"
        '
        'txt_telefono
        '
        Me.txt_telefono.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PROVEEDORBindingSource, "telefono", True))
        Me.txt_telefono.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_telefono.Location = New System.Drawing.Point(467, 89)
        Me.txt_telefono.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_telefono.Name = "txt_telefono"
        Me.txt_telefono.Size = New System.Drawing.Size(160, 25)
        Me.txt_telefono.TabIndex = 11
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(380, 92)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(70, 20)
        Me.Label6.TabIndex = 30
        Me.Label6.Text = "Telefono:"
        '
        'txt_colonia
        '
        Me.txt_colonia.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PROVEEDORBindingSource, "colonia", True))
        Me.txt_colonia.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_colonia.Location = New System.Drawing.Point(93, 171)
        Me.txt_colonia.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_colonia.Name = "txt_colonia"
        Me.txt_colonia.Size = New System.Drawing.Size(245, 25)
        Me.txt_colonia.TabIndex = 9
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(17, 173)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(63, 20)
        Me.Label5.TabIndex = 28
        Me.Label5.Text = "Colonia:"
        '
        'txt_contacto
        '
        Me.txt_contacto.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PROVEEDORBindingSource, "contacto", True))
        Me.txt_contacto.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_contacto.Location = New System.Drawing.Point(93, 107)
        Me.txt_contacto.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_contacto.Name = "txt_contacto"
        Me.txt_contacto.Size = New System.Drawing.Size(245, 25)
        Me.txt_contacto.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(17, 109)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 20)
        Me.Label4.TabIndex = 26
        Me.Label4.Text = "Contacto:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(17, 78)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 20)
        Me.Label2.TabIndex = 24
        Me.Label2.Text = "Nombre:"
        '
        'txt_id
        '
        Me.txt_id.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PROVEEDORBindingSource, "idProveedor", True))
        Me.txt_id.Enabled = False
        Me.txt_id.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_id.Location = New System.Drawing.Point(93, 45)
        Me.txt_id.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_id.Name = "txt_id"
        Me.txt_id.ReadOnly = True
        Me.txt_id.Size = New System.Drawing.Size(78, 25)
        Me.txt_id.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 48)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(27, 20)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "ID:"
        '
        'txt_ciudad
        '
        Me.txt_ciudad.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PROVEEDORBindingSource, "ciudad", True))
        Me.txt_ciudad.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_ciudad.Location = New System.Drawing.Point(93, 138)
        Me.txt_ciudad.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_ciudad.Name = "txt_ciudad"
        Me.txt_ciudad.Size = New System.Drawing.Size(245, 25)
        Me.txt_ciudad.TabIndex = 8
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(17, 140)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(59, 20)
        Me.Label9.TabIndex = 42
        Me.Label9.Text = "Ciudad:"
        '
        'txt_fax
        '
        Me.txt_fax.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PROVEEDORBindingSource, "fax", True))
        Me.txt_fax.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_fax.Location = New System.Drawing.Point(467, 123)
        Me.txt_fax.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_fax.Name = "txt_fax"
        Me.txt_fax.Size = New System.Drawing.Size(160, 25)
        Me.txt_fax.TabIndex = 12
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(380, 126)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(33, 20)
        Me.Label3.TabIndex = 44
        Me.Label3.Text = "Fax:"
        '
        'txt_cp
        '
        Me.txt_cp.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PROVEEDORBindingSource, "cp", True))
        Me.txt_cp.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_cp.Location = New System.Drawing.Point(467, 158)
        Me.txt_cp.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_cp.MaxLength = 5
        Me.txt_cp.Name = "txt_cp"
        Me.txt_cp.Size = New System.Drawing.Size(160, 25)
        Me.txt_cp.TabIndex = 13
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(380, 161)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(76, 20)
        Me.Label7.TabIndex = 46
        Me.Label7.Text = "Cod. Post.:"
        '
        'cmd_siguiente
        '
        Me.cmd_siguiente.Location = New System.Drawing.Point(365, 8)
        Me.cmd_siguiente.Name = "cmd_siguiente"
        Me.cmd_siguiente.Size = New System.Drawing.Size(153, 30)
        Me.cmd_siguiente.TabIndex = 3
        Me.cmd_siguiente.Text = "Siguiente"
        Me.cmd_siguiente.UseVisualStyleBackColor = True
        '
        'cmd_ultimo
        '
        Me.cmd_ultimo.Location = New System.Drawing.Point(541, 7)
        Me.cmd_ultimo.Name = "cmd_ultimo"
        Me.cmd_ultimo.Size = New System.Drawing.Size(153, 31)
        Me.cmd_ultimo.TabIndex = 4
        Me.cmd_ultimo.Text = "Final"
        Me.cmd_ultimo.UseVisualStyleBackColor = True
        '
        'cmd_anterior
        '
        Me.cmd_anterior.Location = New System.Drawing.Point(186, 8)
        Me.cmd_anterior.Name = "cmd_anterior"
        Me.cmd_anterior.Size = New System.Drawing.Size(153, 30)
        Me.cmd_anterior.TabIndex = 2
        Me.cmd_anterior.Text = "Anterior"
        Me.cmd_anterior.UseVisualStyleBackColor = True
        '
        'cmd_inicio
        '
        Me.cmd_inicio.Location = New System.Drawing.Point(12, 7)
        Me.cmd_inicio.Name = "cmd_inicio"
        Me.cmd_inicio.Size = New System.Drawing.Size(154, 31)
        Me.cmd_inicio.TabIndex = 1
        Me.cmd_inicio.Text = "Inicio"
        Me.cmd_inicio.UseVisualStyleBackColor = True
        '
        'btn_salir
        '
        Me.btn_salir.Location = New System.Drawing.Point(550, 251)
        Me.btn_salir.Name = "btn_salir"
        Me.btn_salir.Size = New System.Drawing.Size(153, 36)
        Me.btn_salir.TabIndex = 18
        Me.btn_salir.Text = "Salir"
        Me.btn_salir.UseVisualStyleBackColor = True
        '
        'btn_grabar
        '
        Me.btn_grabar.Location = New System.Drawing.Point(195, 252)
        Me.btn_grabar.Name = "btn_grabar"
        Me.btn_grabar.Size = New System.Drawing.Size(153, 35)
        Me.btn_grabar.TabIndex = 15
        Me.btn_grabar.Text = "Grabar"
        Me.btn_grabar.UseVisualStyleBackColor = True
        '
        'btn_nuevo
        '
        Me.btn_nuevo.Location = New System.Drawing.Point(21, 251)
        Me.btn_nuevo.Name = "btn_nuevo"
        Me.btn_nuevo.Size = New System.Drawing.Size(154, 36)
        Me.btn_nuevo.TabIndex = 14
        Me.btn_nuevo.Text = "Nuevo"
        Me.btn_nuevo.UseVisualStyleBackColor = True
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT        PROVEEDOR.*" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "FROM            PROVEEDOR"
        Me.SqlSelectCommand1.Connection = Me.SqlConnection1
        '
        'SqlConnection1
        '
        Me.SqlConnection1.ConnectionString = "Data Source=CARDENAS-PC;Initial Catalog=INVENTARIO_DB;Integrated Security=True"
        Me.SqlConnection1.FireInfoMessageEventOnUserErrors = False
        '
        'SqlInsertCommand1
        '
        Me.SqlInsertCommand1.CommandText = resources.GetString("SqlInsertCommand1.CommandText")
        Me.SqlInsertCommand1.Connection = Me.SqlConnection1
        Me.SqlInsertCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@idProveedor", System.Data.SqlDbType.Int, 0, "idProveedor"), New System.Data.SqlClient.SqlParameter("@nombre", System.Data.SqlDbType.VarChar, 0, "nombre"), New System.Data.SqlClient.SqlParameter("@contacto", System.Data.SqlDbType.VarChar, 0, "contacto"), New System.Data.SqlClient.SqlParameter("@direccion", System.Data.SqlDbType.VarChar, 0, "direccion"), New System.Data.SqlClient.SqlParameter("@ciudad", System.Data.SqlDbType.VarChar, 0, "ciudad"), New System.Data.SqlClient.SqlParameter("@colonia", System.Data.SqlDbType.VarChar, 0, "colonia"), New System.Data.SqlClient.SqlParameter("@telefono", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), "telefono", System.Data.DataRowVersion.Current, Nothing), New System.Data.SqlClient.SqlParameter("@fax", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), "fax", System.Data.DataRowVersion.Current, Nothing), New System.Data.SqlClient.SqlParameter("@cp", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(5, Byte), CType(0, Byte), "cp", System.Data.DataRowVersion.Current, Nothing)})
        '
        'SqlUpdateCommand1
        '
        Me.SqlUpdateCommand1.CommandText = resources.GetString("SqlUpdateCommand1.CommandText")
        Me.SqlUpdateCommand1.Connection = Me.SqlConnection1
        Me.SqlUpdateCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@idProveedor", System.Data.SqlDbType.Int, 0, "idProveedor"), New System.Data.SqlClient.SqlParameter("@nombre", System.Data.SqlDbType.VarChar, 0, "nombre"), New System.Data.SqlClient.SqlParameter("@contacto", System.Data.SqlDbType.VarChar, 0, "contacto"), New System.Data.SqlClient.SqlParameter("@direccion", System.Data.SqlDbType.VarChar, 0, "direccion"), New System.Data.SqlClient.SqlParameter("@ciudad", System.Data.SqlDbType.VarChar, 0, "ciudad"), New System.Data.SqlClient.SqlParameter("@colonia", System.Data.SqlDbType.VarChar, 0, "colonia"), New System.Data.SqlClient.SqlParameter("@telefono", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), "telefono", System.Data.DataRowVersion.Current, Nothing), New System.Data.SqlClient.SqlParameter("@fax", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), "fax", System.Data.DataRowVersion.Current, Nothing), New System.Data.SqlClient.SqlParameter("@cp", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(5, Byte), CType(0, Byte), "cp", System.Data.DataRowVersion.Current, Nothing), New System.Data.SqlClient.SqlParameter("@Original_idProveedor", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "idProveedor", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_nombre", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "nombre", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_nombre", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "nombre", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_contacto", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "contacto", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_contacto", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "contacto", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_direccion", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "direccion", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_direccion", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "direccion", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_ciudad", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "ciudad", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_ciudad", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ciudad", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_colonia", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "colonia", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_colonia", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "colonia", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_telefono", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "telefono", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_telefono", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), "telefono", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_fax", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "fax", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_fax", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), "fax", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_cp", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "cp", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_cp", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(5, Byte), CType(0, Byte), "cp", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlDeleteCommand1
        '
        Me.SqlDeleteCommand1.CommandText = resources.GetString("SqlDeleteCommand1.CommandText")
        Me.SqlDeleteCommand1.Connection = Me.SqlConnection1
        Me.SqlDeleteCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_idProveedor", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "idProveedor", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_nombre", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "nombre", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_nombre", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "nombre", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_contacto", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "contacto", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_contacto", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "contacto", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_direccion", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "direccion", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_direccion", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "direccion", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_ciudad", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "ciudad", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_ciudad", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ciudad", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_colonia", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "colonia", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_colonia", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "colonia", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_telefono", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "telefono", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_telefono", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), "telefono", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_fax", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "fax", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_fax", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), "fax", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_cp", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "cp", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_cp", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(5, Byte), CType(0, Byte), "cp", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlDataAdapter1
        '
        Me.SqlDataAdapter1.DeleteCommand = Me.SqlDeleteCommand1
        Me.SqlDataAdapter1.InsertCommand = Me.SqlInsertCommand1
        Me.SqlDataAdapter1.SelectCommand = Me.SqlSelectCommand1
        Me.SqlDataAdapter1.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "PROVEEDOR", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("idProveedor", "idProveedor"), New System.Data.Common.DataColumnMapping("nombre", "nombre"), New System.Data.Common.DataColumnMapping("contacto", "contacto"), New System.Data.Common.DataColumnMapping("direccion", "direccion"), New System.Data.Common.DataColumnMapping("ciudad", "ciudad"), New System.Data.Common.DataColumnMapping("colonia", "colonia"), New System.Data.Common.DataColumnMapping("telefono", "telefono"), New System.Data.Common.DataColumnMapping("fax", "fax"), New System.Data.Common.DataColumnMapping("cp", "cp")})})
        Me.SqlDataAdapter1.UpdateCommand = Me.SqlUpdateCommand1
        '
        'PROVEEDORTableAdapter
        '
        Me.PROVEEDORTableAdapter.ClearBeforeFill = True
        '
        'txt_nombre
        '
        Me.txt_nombre.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PROVEEDORBindingSource, "nombre", True))
        Me.txt_nombre.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_nombre.Location = New System.Drawing.Point(92, 77)
        Me.txt_nombre.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_nombre.Name = "txt_nombre"
        Me.txt_nombre.Size = New System.Drawing.Size(245, 25)
        Me.txt_nombre.TabIndex = 6
        '
        'dg_datos
        '
        Me.dg_datos.AutoGenerateColumns = False
        Me.dg_datos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg_datos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IdProveedorDataGridViewTextBoxColumn, Me.NombreDataGridViewTextBoxColumn, Me.ContactoDataGridViewTextBoxColumn, Me.DireccionDataGridViewTextBoxColumn, Me.CiudadDataGridViewTextBoxColumn, Me.ColoniaDataGridViewTextBoxColumn, Me.TelefonoDataGridViewTextBoxColumn, Me.FaxDataGridViewTextBoxColumn, Me.CpDataGridViewTextBoxColumn})
        Me.dg_datos.DataSource = Me.PROVEEDORBindingSource
        Me.dg_datos.Location = New System.Drawing.Point(23, 297)
        Me.dg_datos.Name = "dg_datos"
        Me.dg_datos.ReadOnly = True
        Me.dg_datos.Size = New System.Drawing.Size(679, 109)
        Me.dg_datos.TabIndex = 57
        '
        'btn_modificar
        '
        Me.btn_modificar.Location = New System.Drawing.Point(374, 251)
        Me.btn_modificar.Name = "btn_modificar"
        Me.btn_modificar.Size = New System.Drawing.Size(153, 35)
        Me.btn_modificar.TabIndex = 16
        Me.btn_modificar.Text = "Modificar"
        Me.btn_modificar.UseVisualStyleBackColor = True
        '
        'btn_aceptar
        '
        Me.btn_aceptar.Location = New System.Drawing.Point(195, 251)
        Me.btn_aceptar.Name = "btn_aceptar"
        Me.btn_aceptar.Size = New System.Drawing.Size(153, 35)
        Me.btn_aceptar.TabIndex = 58
        Me.btn_aceptar.Text = "Aceptar"
        Me.btn_aceptar.UseVisualStyleBackColor = True
        Me.btn_aceptar.Visible = False
        '
        'IdProveedorDataGridViewTextBoxColumn
        '
        Me.IdProveedorDataGridViewTextBoxColumn.DataPropertyName = "idProveedor"
        Me.IdProveedorDataGridViewTextBoxColumn.HeaderText = "ID"
        Me.IdProveedorDataGridViewTextBoxColumn.Name = "IdProveedorDataGridViewTextBoxColumn"
        Me.IdProveedorDataGridViewTextBoxColumn.ReadOnly = True
        Me.IdProveedorDataGridViewTextBoxColumn.Width = 50
        '
        'NombreDataGridViewTextBoxColumn
        '
        Me.NombreDataGridViewTextBoxColumn.DataPropertyName = "nombre"
        Me.NombreDataGridViewTextBoxColumn.HeaderText = "Nombre"
        Me.NombreDataGridViewTextBoxColumn.Name = "NombreDataGridViewTextBoxColumn"
        Me.NombreDataGridViewTextBoxColumn.ReadOnly = True
        Me.NombreDataGridViewTextBoxColumn.Width = 140
        '
        'ContactoDataGridViewTextBoxColumn
        '
        Me.ContactoDataGridViewTextBoxColumn.DataPropertyName = "contacto"
        Me.ContactoDataGridViewTextBoxColumn.HeaderText = "Contacto"
        Me.ContactoDataGridViewTextBoxColumn.Name = "ContactoDataGridViewTextBoxColumn"
        Me.ContactoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DireccionDataGridViewTextBoxColumn
        '
        Me.DireccionDataGridViewTextBoxColumn.DataPropertyName = "direccion"
        Me.DireccionDataGridViewTextBoxColumn.HeaderText = "Direccion"
        Me.DireccionDataGridViewTextBoxColumn.Name = "DireccionDataGridViewTextBoxColumn"
        Me.DireccionDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CiudadDataGridViewTextBoxColumn
        '
        Me.CiudadDataGridViewTextBoxColumn.DataPropertyName = "ciudad"
        Me.CiudadDataGridViewTextBoxColumn.HeaderText = "Ciudad"
        Me.CiudadDataGridViewTextBoxColumn.Name = "CiudadDataGridViewTextBoxColumn"
        Me.CiudadDataGridViewTextBoxColumn.ReadOnly = True
        '
        'ColoniaDataGridViewTextBoxColumn
        '
        Me.ColoniaDataGridViewTextBoxColumn.DataPropertyName = "colonia"
        Me.ColoniaDataGridViewTextBoxColumn.HeaderText = "Colonia"
        Me.ColoniaDataGridViewTextBoxColumn.Name = "ColoniaDataGridViewTextBoxColumn"
        Me.ColoniaDataGridViewTextBoxColumn.ReadOnly = True
        Me.ColoniaDataGridViewTextBoxColumn.Visible = False
        '
        'TelefonoDataGridViewTextBoxColumn
        '
        Me.TelefonoDataGridViewTextBoxColumn.DataPropertyName = "telefono"
        Me.TelefonoDataGridViewTextBoxColumn.HeaderText = "Telefono"
        Me.TelefonoDataGridViewTextBoxColumn.Name = "TelefonoDataGridViewTextBoxColumn"
        Me.TelefonoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'FaxDataGridViewTextBoxColumn
        '
        Me.FaxDataGridViewTextBoxColumn.DataPropertyName = "fax"
        Me.FaxDataGridViewTextBoxColumn.HeaderText = "Fax"
        Me.FaxDataGridViewTextBoxColumn.Name = "FaxDataGridViewTextBoxColumn"
        Me.FaxDataGridViewTextBoxColumn.ReadOnly = True
        Me.FaxDataGridViewTextBoxColumn.Visible = False
        '
        'CpDataGridViewTextBoxColumn
        '
        Me.CpDataGridViewTextBoxColumn.DataPropertyName = "cp"
        Me.CpDataGridViewTextBoxColumn.HeaderText = "Cod Post"
        Me.CpDataGridViewTextBoxColumn.Name = "CpDataGridViewTextBoxColumn"
        Me.CpDataGridViewTextBoxColumn.ReadOnly = True
        Me.CpDataGridViewTextBoxColumn.Visible = False
        '
        'Proveedor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(722, 423)
        Me.Controls.Add(Me.btn_aceptar)
        Me.Controls.Add(Me.btn_modificar)
        Me.Controls.Add(Me.dg_datos)
        Me.Controls.Add(Me.txt_nombre)
        Me.Controls.Add(Me.cmd_siguiente)
        Me.Controls.Add(Me.cmd_ultimo)
        Me.Controls.Add(Me.cmd_anterior)
        Me.Controls.Add(Me.cmd_inicio)
        Me.Controls.Add(Me.btn_salir)
        Me.Controls.Add(Me.btn_grabar)
        Me.Controls.Add(Me.btn_nuevo)
        Me.Controls.Add(Me.txt_cp)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txt_fax)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txt_ciudad)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.txt_domicilio)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txt_telefono)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txt_colonia)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txt_contacto)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txt_id)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Proveedor"
        Me.Text = "Proveedor"
        CType(Me.PROVEEDORBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.INVENTARIO_DBDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txt_domicilio As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents txt_telefono As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents txt_colonia As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txt_contacto As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents txt_id As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txt_ciudad As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents txt_fax As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txt_cp As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents cmd_siguiente As Button
    Friend WithEvents cmd_ultimo As Button
    Friend WithEvents cmd_anterior As Button
    Friend WithEvents cmd_inicio As Button
    Friend WithEvents btn_salir As Button
    Friend WithEvents btn_grabar As Button
    Friend WithEvents btn_nuevo As Button
    Friend WithEvents SqlSelectCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlConnection1 As SqlClient.SqlConnection
    Friend WithEvents SqlInsertCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlUpdateCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlDeleteCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlDataAdapter1 As SqlClient.SqlDataAdapter
    Friend WithEvents INVENTARIO_DBDataSet As INVENTARIO_DBDataSet
    Friend WithEvents PROVEEDORBindingSource As BindingSource
    Friend WithEvents PROVEEDORTableAdapter As INVENTARIO_DBDataSetTableAdapters.PROVEEDORTableAdapter
    Friend WithEvents txt_nombre As TextBox
    Friend WithEvents dg_datos As DataGridView
    Friend WithEvents btn_modificar As Button
    Friend WithEvents btn_aceptar As Button
    Friend WithEvents IdProveedorDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents NombreDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents ContactoDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents DireccionDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents CiudadDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents ColoniaDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents TelefonoDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents FaxDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents CpDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
End Class
