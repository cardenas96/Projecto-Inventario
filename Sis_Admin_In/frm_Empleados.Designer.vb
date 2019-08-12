<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_Empleados
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Empleados))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txt_id = New System.Windows.Forms.TextBox()
        Me.EMPLEADOBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.InventariO_DBDataSet1 = New Sis_Admin_In.INVENTARIO_DBDataSet()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txt_ciudad = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txt_colonia = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txt_telefono = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cb_sex = New System.Windows.Forms.ComboBox()
        Me.txt_domicilio = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.dtp_fecha = New System.Windows.Forms.DateTimePicker()
        Me.btn_nuevo = New System.Windows.Forms.Button()
        Me.btn_grabar = New System.Windows.Forms.Button()
        Me.btn_salir = New System.Windows.Forms.Button()
        Me.btn_modificar = New System.Windows.Forms.Button()
        Me.txt_nombre = New System.Windows.Forms.TextBox()
        Me.cmd_siguiente = New System.Windows.Forms.Button()
        Me.cmd_ultimo = New System.Windows.Forms.Button()
        Me.cmd_anterior = New System.Windows.Forms.Button()
        Me.cmd_inicio = New System.Windows.Forms.Button()
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlConnection1 = New System.Data.SqlClient.SqlConnection()
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlDeleteCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlDataAdapter1 = New System.Data.SqlClient.SqlDataAdapter()
        Me.dg_datos = New System.Windows.Forms.DataGridView()
        Me.IdEmpleadoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NombreDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CiudadDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColoniaDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DireccionDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TelefonoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NacimientoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SexoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EMPLEADOTableAdapter = New Sis_Admin_In.INVENTARIO_DBDataSetTableAdapters.EMPLEADOTableAdapter()
        Me.btn_Aceptar = New System.Windows.Forms.Button()
        CType(Me.EMPLEADOBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.InventariO_DBDataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(31, 67)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(27, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ID:"
        '
        'txt_id
        '
        Me.txt_id.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.EMPLEADOBindingSource, "idEmpleado", True))
        Me.txt_id.Enabled = False
        Me.txt_id.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_id.Location = New System.Drawing.Point(116, 62)
        Me.txt_id.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_id.Name = "txt_id"
        Me.txt_id.Size = New System.Drawing.Size(78, 25)
        Me.txt_id.TabIndex = 5
        '
        'EMPLEADOBindingSource
        '
        Me.EMPLEADOBindingSource.DataMember = "EMPLEADO"
        Me.EMPLEADOBindingSource.DataSource = Me.InventariO_DBDataSet1
        '
        'InventariO_DBDataSet1
        '
        Me.InventariO_DBDataSet1.DataSetName = "INVENTARIO_DBDataSet"
        Me.InventariO_DBDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(31, 97)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 20)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Nombre:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(403, 77)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(89, 20)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Nacimiento:"
        '
        'txt_ciudad
        '
        Me.txt_ciudad.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.EMPLEADOBindingSource, "ciudad", True))
        Me.txt_ciudad.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_ciudad.Location = New System.Drawing.Point(116, 130)
        Me.txt_ciudad.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_ciudad.Name = "txt_ciudad"
        Me.txt_ciudad.Size = New System.Drawing.Size(245, 25)
        Me.txt_ciudad.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(31, 130)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 20)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Ciudad:"
        '
        'txt_colonia
        '
        Me.txt_colonia.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.EMPLEADOBindingSource, "colonia", True))
        Me.txt_colonia.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_colonia.Location = New System.Drawing.Point(116, 162)
        Me.txt_colonia.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_colonia.Name = "txt_colonia"
        Me.txt_colonia.Size = New System.Drawing.Size(245, 25)
        Me.txt_colonia.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(31, 165)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(63, 20)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Colonia:"
        '
        'txt_telefono
        '
        Me.txt_telefono.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.EMPLEADOBindingSource, "telefono", True))
        Me.txt_telefono.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_telefono.Location = New System.Drawing.Point(490, 106)
        Me.txt_telefono.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_telefono.Name = "txt_telefono"
        Me.txt_telefono.Size = New System.Drawing.Size(160, 25)
        Me.txt_telefono.TabIndex = 11
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(403, 111)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(70, 20)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Telefono:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(403, 141)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(44, 20)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "Sexo:"
        '
        'cb_sex
        '
        Me.cb_sex.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.EMPLEADOBindingSource, "sexo", True))
        Me.cb_sex.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cb_sex.FormattingEnabled = True
        Me.cb_sex.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cb_sex.Items.AddRange(New Object() {"Femenino", "Masculino"})
        Me.cb_sex.Location = New System.Drawing.Point(490, 140)
        Me.cb_sex.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cb_sex.Name = "cb_sex"
        Me.cb_sex.Size = New System.Drawing.Size(160, 25)
        Me.cb_sex.TabIndex = 12
        '
        'txt_domicilio
        '
        Me.txt_domicilio.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.EMPLEADOBindingSource, "direccion", True))
        Me.txt_domicilio.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_domicilio.Location = New System.Drawing.Point(116, 196)
        Me.txt_domicilio.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_domicilio.Name = "txt_domicilio"
        Me.txt_domicilio.Size = New System.Drawing.Size(245, 25)
        Me.txt_domicilio.TabIndex = 9
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(31, 199)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(77, 20)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Domicilio:"
        '
        'dtp_fecha
        '
        Me.dtp_fecha.DataBindings.Add(New System.Windows.Forms.Binding("Value", Me.EMPLEADOBindingSource, "nacimiento", True))
        Me.dtp_fecha.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtp_fecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtp_fecha.Location = New System.Drawing.Point(499, 72)
        Me.dtp_fecha.Name = "dtp_fecha"
        Me.dtp_fecha.Size = New System.Drawing.Size(141, 25)
        Me.dtp_fecha.TabIndex = 10
        '
        'btn_nuevo
        '
        Me.btn_nuevo.Location = New System.Drawing.Point(21, 241)
        Me.btn_nuevo.Name = "btn_nuevo"
        Me.btn_nuevo.Size = New System.Drawing.Size(154, 36)
        Me.btn_nuevo.TabIndex = 13
        Me.btn_nuevo.Text = "Nuevo"
        Me.btn_nuevo.UseVisualStyleBackColor = True
        '
        'btn_grabar
        '
        Me.btn_grabar.Location = New System.Drawing.Point(195, 242)
        Me.btn_grabar.Name = "btn_grabar"
        Me.btn_grabar.Size = New System.Drawing.Size(153, 35)
        Me.btn_grabar.TabIndex = 14
        Me.btn_grabar.Text = "Grabar"
        Me.btn_grabar.UseVisualStyleBackColor = True
        '
        'btn_salir
        '
        Me.btn_salir.Location = New System.Drawing.Point(550, 241)
        Me.btn_salir.Name = "btn_salir"
        Me.btn_salir.Size = New System.Drawing.Size(153, 36)
        Me.btn_salir.TabIndex = 16
        Me.btn_salir.Text = "Salir"
        Me.btn_salir.UseVisualStyleBackColor = True
        '
        'btn_modificar
        '
        Me.btn_modificar.Location = New System.Drawing.Point(374, 242)
        Me.btn_modificar.Name = "btn_modificar"
        Me.btn_modificar.Size = New System.Drawing.Size(153, 35)
        Me.btn_modificar.TabIndex = 21
        Me.btn_modificar.Text = "Modificar"
        Me.btn_modificar.UseVisualStyleBackColor = True
        '
        'txt_nombre
        '
        Me.txt_nombre.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.EMPLEADOBindingSource, "nombre", True))
        Me.txt_nombre.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_nombre.Location = New System.Drawing.Point(115, 96)
        Me.txt_nombre.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_nombre.Name = "txt_nombre"
        Me.txt_nombre.Size = New System.Drawing.Size(245, 25)
        Me.txt_nombre.TabIndex = 6
        '
        'cmd_siguiente
        '
        Me.cmd_siguiente.Location = New System.Drawing.Point(374, 12)
        Me.cmd_siguiente.Name = "cmd_siguiente"
        Me.cmd_siguiente.Size = New System.Drawing.Size(153, 30)
        Me.cmd_siguiente.TabIndex = 3
        Me.cmd_siguiente.Text = "Siguiente"
        Me.cmd_siguiente.UseVisualStyleBackColor = True
        '
        'cmd_ultimo
        '
        Me.cmd_ultimo.Location = New System.Drawing.Point(550, 11)
        Me.cmd_ultimo.Name = "cmd_ultimo"
        Me.cmd_ultimo.Size = New System.Drawing.Size(153, 31)
        Me.cmd_ultimo.TabIndex = 4
        Me.cmd_ultimo.Text = "Final"
        Me.cmd_ultimo.UseVisualStyleBackColor = True
        '
        'cmd_anterior
        '
        Me.cmd_anterior.Location = New System.Drawing.Point(195, 12)
        Me.cmd_anterior.Name = "cmd_anterior"
        Me.cmd_anterior.Size = New System.Drawing.Size(153, 30)
        Me.cmd_anterior.TabIndex = 2
        Me.cmd_anterior.Text = "Anterior"
        Me.cmd_anterior.UseVisualStyleBackColor = True
        '
        'cmd_inicio
        '
        Me.cmd_inicio.Location = New System.Drawing.Point(21, 11)
        Me.cmd_inicio.Name = "cmd_inicio"
        Me.cmd_inicio.Size = New System.Drawing.Size(154, 31)
        Me.cmd_inicio.TabIndex = 1
        Me.cmd_inicio.Text = "Inicio"
        Me.cmd_inicio.UseVisualStyleBackColor = True
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT        EMPLEADO.*" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "FROM            EMPLEADO"
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
        Me.SqlInsertCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@idEmpleado", System.Data.SqlDbType.Int, 0, "idEmpleado"), New System.Data.SqlClient.SqlParameter("@nombre", System.Data.SqlDbType.VarChar, 0, "nombre"), New System.Data.SqlClient.SqlParameter("@ciudad", System.Data.SqlDbType.VarChar, 0, "ciudad"), New System.Data.SqlClient.SqlParameter("@colonia", System.Data.SqlDbType.VarChar, 0, "colonia"), New System.Data.SqlClient.SqlParameter("@direccion", System.Data.SqlDbType.VarChar, 0, "direccion"), New System.Data.SqlClient.SqlParameter("@telefono", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), "telefono", System.Data.DataRowVersion.Current, Nothing), New System.Data.SqlClient.SqlParameter("@nacimiento", System.Data.SqlDbType.[Date], 0, "nacimiento"), New System.Data.SqlClient.SqlParameter("@sexo", System.Data.SqlDbType.VarChar, 0, "sexo")})
        '
        'SqlUpdateCommand1
        '
        Me.SqlUpdateCommand1.CommandText = resources.GetString("SqlUpdateCommand1.CommandText")
        Me.SqlUpdateCommand1.Connection = Me.SqlConnection1
        Me.SqlUpdateCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@idEmpleado", System.Data.SqlDbType.Int, 0, "idEmpleado"), New System.Data.SqlClient.SqlParameter("@nombre", System.Data.SqlDbType.VarChar, 0, "nombre"), New System.Data.SqlClient.SqlParameter("@ciudad", System.Data.SqlDbType.VarChar, 0, "ciudad"), New System.Data.SqlClient.SqlParameter("@colonia", System.Data.SqlDbType.VarChar, 0, "colonia"), New System.Data.SqlClient.SqlParameter("@direccion", System.Data.SqlDbType.VarChar, 0, "direccion"), New System.Data.SqlClient.SqlParameter("@telefono", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), "telefono", System.Data.DataRowVersion.Current, Nothing), New System.Data.SqlClient.SqlParameter("@nacimiento", System.Data.SqlDbType.[Date], 0, "nacimiento"), New System.Data.SqlClient.SqlParameter("@sexo", System.Data.SqlDbType.VarChar, 0, "sexo"), New System.Data.SqlClient.SqlParameter("@Original_idEmpleado", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "idEmpleado", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_nombre", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "nombre", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_nombre", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "nombre", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_ciudad", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "ciudad", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_ciudad", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ciudad", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_colonia", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "colonia", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_colonia", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "colonia", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_direccion", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "direccion", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_direccion", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "direccion", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_telefono", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "telefono", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_telefono", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), "telefono", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_nacimiento", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "nacimiento", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_nacimiento", System.Data.SqlDbType.[Date], 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "nacimiento", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_sexo", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "sexo", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_sexo", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "sexo", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlDeleteCommand1
        '
        Me.SqlDeleteCommand1.CommandText = resources.GetString("SqlDeleteCommand1.CommandText")
        Me.SqlDeleteCommand1.Connection = Me.SqlConnection1
        Me.SqlDeleteCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_idEmpleado", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "idEmpleado", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_nombre", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "nombre", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_nombre", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "nombre", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_ciudad", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "ciudad", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_ciudad", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ciudad", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_colonia", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "colonia", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_colonia", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "colonia", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_direccion", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "direccion", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_direccion", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "direccion", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_telefono", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "telefono", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_telefono", System.Data.SqlDbType.[Decimal], 0, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), "telefono", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_nacimiento", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "nacimiento", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_nacimiento", System.Data.SqlDbType.[Date], 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "nacimiento", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_sexo", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "sexo", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_sexo", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "sexo", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlDataAdapter1
        '
        Me.SqlDataAdapter1.DeleteCommand = Me.SqlDeleteCommand1
        Me.SqlDataAdapter1.InsertCommand = Me.SqlInsertCommand1
        Me.SqlDataAdapter1.SelectCommand = Me.SqlSelectCommand1
        Me.SqlDataAdapter1.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "EMPLEADO", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("idEmpleado", "idEmpleado"), New System.Data.Common.DataColumnMapping("nombre", "nombre"), New System.Data.Common.DataColumnMapping("ciudad", "ciudad"), New System.Data.Common.DataColumnMapping("colonia", "colonia"), New System.Data.Common.DataColumnMapping("direccion", "direccion"), New System.Data.Common.DataColumnMapping("telefono", "telefono"), New System.Data.Common.DataColumnMapping("nacimiento", "nacimiento"), New System.Data.Common.DataColumnMapping("sexo", "sexo")})})
        Me.SqlDataAdapter1.UpdateCommand = Me.SqlUpdateCommand1
        '
        'dg_datos
        '
        Me.dg_datos.AutoGenerateColumns = False
        Me.dg_datos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg_datos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IdEmpleadoDataGridViewTextBoxColumn, Me.NombreDataGridViewTextBoxColumn, Me.CiudadDataGridViewTextBoxColumn, Me.ColoniaDataGridViewTextBoxColumn, Me.DireccionDataGridViewTextBoxColumn, Me.TelefonoDataGridViewTextBoxColumn, Me.NacimientoDataGridViewTextBoxColumn, Me.SexoDataGridViewTextBoxColumn})
        Me.dg_datos.DataSource = Me.EMPLEADOBindingSource
        Me.dg_datos.Location = New System.Drawing.Point(13, 289)
        Me.dg_datos.Name = "dg_datos"
        Me.dg_datos.ReadOnly = True
        Me.dg_datos.Size = New System.Drawing.Size(705, 135)
        Me.dg_datos.TabIndex = 22
        '
        'IdEmpleadoDataGridViewTextBoxColumn
        '
        Me.IdEmpleadoDataGridViewTextBoxColumn.DataPropertyName = "idEmpleado"
        Me.IdEmpleadoDataGridViewTextBoxColumn.HeaderText = "ID"
        Me.IdEmpleadoDataGridViewTextBoxColumn.Name = "IdEmpleadoDataGridViewTextBoxColumn"
        Me.IdEmpleadoDataGridViewTextBoxColumn.ReadOnly = True
        Me.IdEmpleadoDataGridViewTextBoxColumn.Width = 30
        '
        'NombreDataGridViewTextBoxColumn
        '
        Me.NombreDataGridViewTextBoxColumn.DataPropertyName = "nombre"
        Me.NombreDataGridViewTextBoxColumn.HeaderText = "Nombre"
        Me.NombreDataGridViewTextBoxColumn.Name = "NombreDataGridViewTextBoxColumn"
        Me.NombreDataGridViewTextBoxColumn.ReadOnly = True
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
        '
        'DireccionDataGridViewTextBoxColumn
        '
        Me.DireccionDataGridViewTextBoxColumn.DataPropertyName = "direccion"
        Me.DireccionDataGridViewTextBoxColumn.HeaderText = "Direccion"
        Me.DireccionDataGridViewTextBoxColumn.Name = "DireccionDataGridViewTextBoxColumn"
        Me.DireccionDataGridViewTextBoxColumn.ReadOnly = True
        '
        'TelefonoDataGridViewTextBoxColumn
        '
        Me.TelefonoDataGridViewTextBoxColumn.DataPropertyName = "telefono"
        Me.TelefonoDataGridViewTextBoxColumn.HeaderText = "Telefono"
        Me.TelefonoDataGridViewTextBoxColumn.Name = "TelefonoDataGridViewTextBoxColumn"
        Me.TelefonoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'NacimientoDataGridViewTextBoxColumn
        '
        Me.NacimientoDataGridViewTextBoxColumn.DataPropertyName = "nacimiento"
        Me.NacimientoDataGridViewTextBoxColumn.HeaderText = "Nacimiento"
        Me.NacimientoDataGridViewTextBoxColumn.Name = "NacimientoDataGridViewTextBoxColumn"
        Me.NacimientoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'SexoDataGridViewTextBoxColumn
        '
        Me.SexoDataGridViewTextBoxColumn.DataPropertyName = "sexo"
        Me.SexoDataGridViewTextBoxColumn.HeaderText = "Sexo"
        Me.SexoDataGridViewTextBoxColumn.Name = "SexoDataGridViewTextBoxColumn"
        Me.SexoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'EMPLEADOTableAdapter
        '
        Me.EMPLEADOTableAdapter.ClearBeforeFill = True
        '
        'btn_Aceptar
        '
        Me.btn_Aceptar.Location = New System.Drawing.Point(195, 241)
        Me.btn_Aceptar.Name = "btn_Aceptar"
        Me.btn_Aceptar.Size = New System.Drawing.Size(153, 35)
        Me.btn_Aceptar.TabIndex = 23
        Me.btn_Aceptar.Text = "Aceptar"
        Me.btn_Aceptar.UseVisualStyleBackColor = True
        Me.btn_Aceptar.Visible = False
        '
        'frm_Empleados
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Highlight
        Me.ClientSize = New System.Drawing.Size(736, 436)
        Me.Controls.Add(Me.btn_Aceptar)
        Me.Controls.Add(Me.dg_datos)
        Me.Controls.Add(Me.cmd_siguiente)
        Me.Controls.Add(Me.cmd_ultimo)
        Me.Controls.Add(Me.cmd_anterior)
        Me.Controls.Add(Me.cmd_inicio)
        Me.Controls.Add(Me.txt_nombre)
        Me.Controls.Add(Me.btn_modificar)
        Me.Controls.Add(Me.btn_salir)
        Me.Controls.Add(Me.btn_grabar)
        Me.Controls.Add(Me.btn_nuevo)
        Me.Controls.Add(Me.dtp_fecha)
        Me.Controls.Add(Me.txt_domicilio)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.cb_sex)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txt_telefono)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txt_colonia)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txt_ciudad)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txt_id)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "frm_Empleados"
        Me.Text = "Empleados"
        CType(Me.EMPLEADOBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.InventariO_DBDataSet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents txt_id As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents txt_ciudad As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents txt_colonia As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txt_telefono As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents cb_sex As ComboBox
    Friend WithEvents txt_domicilio As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents dtp_fecha As DateTimePicker
    Friend WithEvents btn_nuevo As Button
    Friend WithEvents btn_grabar As Button
    Friend WithEvents btn_salir As Button
    Friend WithEvents btn_modificar As Button
    Friend WithEvents txt_nombre As TextBox
    Friend WithEvents cmd_siguiente As Button
    Friend WithEvents cmd_ultimo As Button
    Friend WithEvents cmd_anterior As Button
    Friend WithEvents cmd_inicio As Button
    Friend WithEvents SqlSelectCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlConnection1 As SqlClient.SqlConnection
    Friend WithEvents SqlInsertCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlUpdateCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlDeleteCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlDataAdapter1 As SqlClient.SqlDataAdapter
    Friend WithEvents InventariO_DBDataSet1 As INVENTARIO_DBDataSet
    Friend WithEvents EMPLEADOBindingSource As BindingSource
    Friend WithEvents EMPLEADOTableAdapter As INVENTARIO_DBDataSetTableAdapters.EMPLEADOTableAdapter
    Friend WithEvents dg_datos As DataGridView
    Friend WithEvents IdEmpleadoDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents NombreDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents CiudadDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents ColoniaDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents DireccionDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents TelefonoDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents NacimientoDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents SexoDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents btn_Aceptar As Button
End Class
