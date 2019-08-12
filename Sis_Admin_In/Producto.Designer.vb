<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Producto
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Producto))
        Me.txt_costo = New System.Windows.Forms.TextBox()
        Me.PRODUCTOBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.InventariO_DBDataSet1 = New Sis_Admin_In.INVENTARIO_DBDataSet()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txt_min = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txt_existencias = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txt_max = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txt_marca = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txt_id = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cb_unidad = New System.Windows.Forms.ComboBox()
        Me.cmd_siguiente = New System.Windows.Forms.Button()
        Me.cmd_ultimo = New System.Windows.Forms.Button()
        Me.cmd_anterior = New System.Windows.Forms.Button()
        Me.cmd_inicio = New System.Windows.Forms.Button()
        Me.txt_nombre = New System.Windows.Forms.TextBox()
        Me.btn_modificar = New System.Windows.Forms.Button()
        Me.btn_salir = New System.Windows.Forms.Button()
        Me.btn_grabar = New System.Windows.Forms.Button()
        Me.btn_nuevo = New System.Windows.Forms.Button()
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlConnection1 = New System.Data.SqlClient.SqlConnection()
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlDeleteCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlDataAdapter1 = New System.Data.SqlClient.SqlDataAdapter()
        Me.PRODUCTOTableAdapter = New Sis_Admin_In.INVENTARIO_DBDataSetTableAdapters.PRODUCTOTableAdapter()
        Me.dg_datos = New System.Windows.Forms.DataGridView()
        Me.IdProductoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NombreDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MarcaDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.UnidadDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CostoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ExistencasDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MinimoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MaximmoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btn_aceptar = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        CType(Me.PRODUCTOBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.InventariO_DBDataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txt_costo
        '
        Me.txt_costo.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PRODUCTOBindingSource, "costo", True))
        Me.txt_costo.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_costo.Location = New System.Drawing.Point(476, 154)
        Me.txt_costo.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_costo.Name = "txt_costo"
        Me.txt_costo.Size = New System.Drawing.Size(160, 25)
        Me.txt_costo.TabIndex = 11
        '
        'PRODUCTOBindingSource
        '
        Me.PRODUCTOBindingSource.DataMember = "PRODUCTO"
        Me.PRODUCTOBindingSource.DataSource = Me.InventariO_DBDataSet1
        '
        'InventariO_DBDataSet1
        '
        Me.InventariO_DBDataSet1.DataSetName = "INVENTARIO_DBDataSet"
        Me.InventariO_DBDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(389, 157)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(50, 20)
        Me.Label3.TabIndex = 65
        Me.Label3.Text = "Costo:"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(20, 194)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(60, 20)
        Me.Label9.TabIndex = 63
        Me.Label9.Text = "Unidad:"
        '
        'txt_min
        '
        Me.txt_min.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PRODUCTOBindingSource, "minimo", True))
        Me.txt_min.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_min.Location = New System.Drawing.Point(105, 260)
        Me.txt_min.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_min.Name = "txt_min"
        Me.txt_min.Size = New System.Drawing.Size(245, 25)
        Me.txt_min.TabIndex = 9
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(20, 263)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(63, 20)
        Me.Label8.TabIndex = 57
        Me.Label8.Text = "Minimo:"
        '
        'txt_existencias
        '
        Me.txt_existencias.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PRODUCTOBindingSource, "existencas", True))
        Me.txt_existencias.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_existencias.Location = New System.Drawing.Point(476, 120)
        Me.txt_existencias.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_existencias.Name = "txt_existencias"
        Me.txt_existencias.Size = New System.Drawing.Size(160, 25)
        Me.txt_existencias.TabIndex = 10
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(389, 123)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(83, 20)
        Me.Label6.TabIndex = 55
        Me.Label6.Text = "Existencias:"
        '
        'txt_max
        '
        Me.txt_max.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PRODUCTOBindingSource, "maximmo", True))
        Me.txt_max.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_max.Location = New System.Drawing.Point(105, 228)
        Me.txt_max.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_max.Name = "txt_max"
        Me.txt_max.Size = New System.Drawing.Size(245, 25)
        Me.txt_max.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(20, 228)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(66, 20)
        Me.Label5.TabIndex = 53
        Me.Label5.Text = "Maximo:"
        '
        'txt_marca
        '
        Me.txt_marca.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PRODUCTOBindingSource, "marca", True))
        Me.txt_marca.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_marca.Location = New System.Drawing.Point(105, 163)
        Me.txt_marca.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_marca.Name = "txt_marca"
        Me.txt_marca.Size = New System.Drawing.Size(245, 25)
        Me.txt_marca.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(20, 163)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 20)
        Me.Label4.TabIndex = 51
        Me.Label4.Text = "Marca:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(20, 137)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 20)
        Me.Label2.TabIndex = 50
        Me.Label2.Text = "Nombre:"
        '
        'txt_id
        '
        Me.txt_id.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PRODUCTOBindingSource, "idProducto", True))
        Me.txt_id.Enabled = False
        Me.txt_id.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_id.Location = New System.Drawing.Point(105, 100)
        Me.txt_id.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_id.Name = "txt_id"
        Me.txt_id.ReadOnly = True
        Me.txt_id.Size = New System.Drawing.Size(78, 25)
        Me.txt_id.TabIndex = 49
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 105)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(27, 20)
        Me.Label1.TabIndex = 48
        Me.Label1.Text = "ID:"
        '
        'cb_unidad
        '
        Me.cb_unidad.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PRODUCTOBindingSource, "unidad", True))
        Me.cb_unidad.FormattingEnabled = True
        Me.cb_unidad.Items.AddRange(New Object() {"Pieza", "Litro", "Paquete", "Metro", "Kilo"})
        Me.cb_unidad.Location = New System.Drawing.Point(105, 194)
        Me.cb_unidad.Name = "cb_unidad"
        Me.cb_unidad.Size = New System.Drawing.Size(245, 28)
        Me.cb_unidad.TabIndex = 7
        '
        'cmd_siguiente
        '
        Me.cmd_siguiente.Location = New System.Drawing.Point(376, 64)
        Me.cmd_siguiente.Name = "cmd_siguiente"
        Me.cmd_siguiente.Size = New System.Drawing.Size(153, 30)
        Me.cmd_siguiente.TabIndex = 3
        Me.cmd_siguiente.Text = "Siguiente"
        Me.cmd_siguiente.UseVisualStyleBackColor = True
        '
        'cmd_ultimo
        '
        Me.cmd_ultimo.Location = New System.Drawing.Point(552, 63)
        Me.cmd_ultimo.Name = "cmd_ultimo"
        Me.cmd_ultimo.Size = New System.Drawing.Size(153, 31)
        Me.cmd_ultimo.TabIndex = 4
        Me.cmd_ultimo.Text = "Final"
        Me.cmd_ultimo.UseVisualStyleBackColor = True
        '
        'cmd_anterior
        '
        Me.cmd_anterior.Location = New System.Drawing.Point(197, 64)
        Me.cmd_anterior.Name = "cmd_anterior"
        Me.cmd_anterior.Size = New System.Drawing.Size(153, 30)
        Me.cmd_anterior.TabIndex = 2
        Me.cmd_anterior.Text = "Anterior"
        Me.cmd_anterior.UseVisualStyleBackColor = True
        '
        'cmd_inicio
        '
        Me.cmd_inicio.Location = New System.Drawing.Point(23, 63)
        Me.cmd_inicio.Name = "cmd_inicio"
        Me.cmd_inicio.Size = New System.Drawing.Size(154, 31)
        Me.cmd_inicio.TabIndex = 1
        Me.cmd_inicio.Text = "Inicio"
        Me.cmd_inicio.UseVisualStyleBackColor = True
        '
        'txt_nombre
        '
        Me.txt_nombre.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.PRODUCTOBindingSource, "nombre", True))
        Me.txt_nombre.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_nombre.Location = New System.Drawing.Point(105, 132)
        Me.txt_nombre.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txt_nombre.Name = "txt_nombre"
        Me.txt_nombre.Size = New System.Drawing.Size(245, 25)
        Me.txt_nombre.TabIndex = 5
        '
        'btn_modificar
        '
        Me.btn_modificar.Location = New System.Drawing.Point(375, 297)
        Me.btn_modificar.Name = "btn_modificar"
        Me.btn_modificar.Size = New System.Drawing.Size(153, 35)
        Me.btn_modificar.TabIndex = 14
        Me.btn_modificar.Text = "Modificar"
        Me.btn_modificar.UseVisualStyleBackColor = True
        '
        'btn_salir
        '
        Me.btn_salir.Location = New System.Drawing.Point(548, 296)
        Me.btn_salir.Name = "btn_salir"
        Me.btn_salir.Size = New System.Drawing.Size(153, 36)
        Me.btn_salir.TabIndex = 16
        Me.btn_salir.Text = "Salir"
        Me.btn_salir.UseVisualStyleBackColor = True
        '
        'btn_grabar
        '
        Me.btn_grabar.Location = New System.Drawing.Point(193, 297)
        Me.btn_grabar.Name = "btn_grabar"
        Me.btn_grabar.Size = New System.Drawing.Size(153, 35)
        Me.btn_grabar.TabIndex = 13
        Me.btn_grabar.Text = "Grabar"
        Me.btn_grabar.UseVisualStyleBackColor = True
        '
        'btn_nuevo
        '
        Me.btn_nuevo.Location = New System.Drawing.Point(19, 296)
        Me.btn_nuevo.Name = "btn_nuevo"
        Me.btn_nuevo.Size = New System.Drawing.Size(154, 36)
        Me.btn_nuevo.TabIndex = 12
        Me.btn_nuevo.Text = "Nuevo"
        Me.btn_nuevo.UseVisualStyleBackColor = True
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT        PRODUCTO.*" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "FROM            PRODUCTO"
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
        Me.SqlInsertCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@idProducto", System.Data.SqlDbType.Int, 0, "idProducto"), New System.Data.SqlClient.SqlParameter("@nombre", System.Data.SqlDbType.VarChar, 0, "nombre"), New System.Data.SqlClient.SqlParameter("@marca", System.Data.SqlDbType.VarChar, 0, "marca"), New System.Data.SqlClient.SqlParameter("@unidad", System.Data.SqlDbType.VarChar, 0, "unidad"), New System.Data.SqlClient.SqlParameter("@costo", System.Data.SqlDbType.Float, 0, "costo"), New System.Data.SqlClient.SqlParameter("@existencas", System.Data.SqlDbType.Int, 0, "existencas"), New System.Data.SqlClient.SqlParameter("@minimo", System.Data.SqlDbType.Int, 0, "minimo"), New System.Data.SqlClient.SqlParameter("@maximmo", System.Data.SqlDbType.Int, 0, "maximmo")})
        '
        'SqlUpdateCommand1
        '
        Me.SqlUpdateCommand1.CommandText = resources.GetString("SqlUpdateCommand1.CommandText")
        Me.SqlUpdateCommand1.Connection = Me.SqlConnection1
        Me.SqlUpdateCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@idProducto", System.Data.SqlDbType.Int, 0, "idProducto"), New System.Data.SqlClient.SqlParameter("@nombre", System.Data.SqlDbType.VarChar, 0, "nombre"), New System.Data.SqlClient.SqlParameter("@marca", System.Data.SqlDbType.VarChar, 0, "marca"), New System.Data.SqlClient.SqlParameter("@unidad", System.Data.SqlDbType.VarChar, 0, "unidad"), New System.Data.SqlClient.SqlParameter("@costo", System.Data.SqlDbType.Float, 0, "costo"), New System.Data.SqlClient.SqlParameter("@existencas", System.Data.SqlDbType.Int, 0, "existencas"), New System.Data.SqlClient.SqlParameter("@minimo", System.Data.SqlDbType.Int, 0, "minimo"), New System.Data.SqlClient.SqlParameter("@maximmo", System.Data.SqlDbType.Int, 0, "maximmo"), New System.Data.SqlClient.SqlParameter("@Original_idProducto", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "idProducto", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_nombre", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "nombre", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_nombre", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "nombre", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_marca", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "marca", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_marca", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "marca", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_unidad", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "unidad", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_unidad", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "unidad", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_costo", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "costo", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_costo", System.Data.SqlDbType.Float, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "costo", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_existencas", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "existencas", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_existencas", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "existencas", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_minimo", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "minimo", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_minimo", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "minimo", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_maximmo", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "maximmo", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_maximmo", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "maximmo", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlDeleteCommand1
        '
        Me.SqlDeleteCommand1.CommandText = resources.GetString("SqlDeleteCommand1.CommandText")
        Me.SqlDeleteCommand1.Connection = Me.SqlConnection1
        Me.SqlDeleteCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_idProducto", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "idProducto", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_nombre", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "nombre", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_nombre", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "nombre", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_marca", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "marca", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_marca", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "marca", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_unidad", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "unidad", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_unidad", System.Data.SqlDbType.VarChar, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "unidad", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_costo", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "costo", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_costo", System.Data.SqlDbType.Float, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "costo", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_existencas", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "existencas", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_existencas", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "existencas", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_minimo", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "minimo", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_minimo", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "minimo", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@IsNull_maximmo", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, CType(0, Byte), CType(0, Byte), "maximmo", System.Data.DataRowVersion.Original, True, Nothing, "", "", ""), New System.Data.SqlClient.SqlParameter("@Original_maximmo", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "maximmo", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlDataAdapter1
        '
        Me.SqlDataAdapter1.DeleteCommand = Me.SqlDeleteCommand1
        Me.SqlDataAdapter1.InsertCommand = Me.SqlInsertCommand1
        Me.SqlDataAdapter1.SelectCommand = Me.SqlSelectCommand1
        Me.SqlDataAdapter1.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "PRODUCTO", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("idProducto", "idProducto"), New System.Data.Common.DataColumnMapping("nombre", "nombre"), New System.Data.Common.DataColumnMapping("marca", "marca"), New System.Data.Common.DataColumnMapping("unidad", "unidad"), New System.Data.Common.DataColumnMapping("costo", "costo"), New System.Data.Common.DataColumnMapping("existencas", "existencas"), New System.Data.Common.DataColumnMapping("minimo", "minimo"), New System.Data.Common.DataColumnMapping("maximmo", "maximmo")})})
        Me.SqlDataAdapter1.UpdateCommand = Me.SqlUpdateCommand1
        '
        'PRODUCTOTableAdapter
        '
        Me.PRODUCTOTableAdapter.ClearBeforeFill = True
        '
        'dg_datos
        '
        Me.dg_datos.AutoGenerateColumns = False
        Me.dg_datos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg_datos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IdProductoDataGridViewTextBoxColumn, Me.NombreDataGridViewTextBoxColumn, Me.MarcaDataGridViewTextBoxColumn, Me.UnidadDataGridViewTextBoxColumn, Me.CostoDataGridViewTextBoxColumn, Me.ExistencasDataGridViewTextBoxColumn, Me.MinimoDataGridViewTextBoxColumn, Me.MaximmoDataGridViewTextBoxColumn})
        Me.dg_datos.DataSource = Me.PRODUCTOBindingSource
        Me.dg_datos.Location = New System.Drawing.Point(16, 345)
        Me.dg_datos.Name = "dg_datos"
        Me.dg_datos.ReadOnly = True
        Me.dg_datos.Size = New System.Drawing.Size(688, 111)
        Me.dg_datos.TabIndex = 77
        '
        'IdProductoDataGridViewTextBoxColumn
        '
        Me.IdProductoDataGridViewTextBoxColumn.DataPropertyName = "idProducto"
        Me.IdProductoDataGridViewTextBoxColumn.HeaderText = "ID"
        Me.IdProductoDataGridViewTextBoxColumn.Name = "IdProductoDataGridViewTextBoxColumn"
        Me.IdProductoDataGridViewTextBoxColumn.ReadOnly = True
        Me.IdProductoDataGridViewTextBoxColumn.Width = 50
        '
        'NombreDataGridViewTextBoxColumn
        '
        Me.NombreDataGridViewTextBoxColumn.DataPropertyName = "nombre"
        Me.NombreDataGridViewTextBoxColumn.HeaderText = "Nombre"
        Me.NombreDataGridViewTextBoxColumn.Name = "NombreDataGridViewTextBoxColumn"
        Me.NombreDataGridViewTextBoxColumn.ReadOnly = True
        '
        'MarcaDataGridViewTextBoxColumn
        '
        Me.MarcaDataGridViewTextBoxColumn.DataPropertyName = "marca"
        Me.MarcaDataGridViewTextBoxColumn.HeaderText = "Marca"
        Me.MarcaDataGridViewTextBoxColumn.Name = "MarcaDataGridViewTextBoxColumn"
        Me.MarcaDataGridViewTextBoxColumn.ReadOnly = True
        '
        'UnidadDataGridViewTextBoxColumn
        '
        Me.UnidadDataGridViewTextBoxColumn.DataPropertyName = "unidad"
        Me.UnidadDataGridViewTextBoxColumn.HeaderText = "Unidad"
        Me.UnidadDataGridViewTextBoxColumn.Name = "UnidadDataGridViewTextBoxColumn"
        Me.UnidadDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CostoDataGridViewTextBoxColumn
        '
        Me.CostoDataGridViewTextBoxColumn.DataPropertyName = "costo"
        Me.CostoDataGridViewTextBoxColumn.HeaderText = "Costo"
        Me.CostoDataGridViewTextBoxColumn.Name = "CostoDataGridViewTextBoxColumn"
        Me.CostoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'ExistencasDataGridViewTextBoxColumn
        '
        Me.ExistencasDataGridViewTextBoxColumn.DataPropertyName = "existencas"
        Me.ExistencasDataGridViewTextBoxColumn.HeaderText = "Existencas"
        Me.ExistencasDataGridViewTextBoxColumn.Name = "ExistencasDataGridViewTextBoxColumn"
        Me.ExistencasDataGridViewTextBoxColumn.ReadOnly = True
        '
        'MinimoDataGridViewTextBoxColumn
        '
        Me.MinimoDataGridViewTextBoxColumn.DataPropertyName = "minimo"
        Me.MinimoDataGridViewTextBoxColumn.HeaderText = "Minimo"
        Me.MinimoDataGridViewTextBoxColumn.Name = "MinimoDataGridViewTextBoxColumn"
        Me.MinimoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'MaximmoDataGridViewTextBoxColumn
        '
        Me.MaximmoDataGridViewTextBoxColumn.DataPropertyName = "maximmo"
        Me.MaximmoDataGridViewTextBoxColumn.HeaderText = "Maximo"
        Me.MaximmoDataGridViewTextBoxColumn.Name = "MaximmoDataGridViewTextBoxColumn"
        Me.MaximmoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'btn_aceptar
        '
        Me.btn_aceptar.Location = New System.Drawing.Point(193, 297)
        Me.btn_aceptar.Name = "btn_aceptar"
        Me.btn_aceptar.Size = New System.Drawing.Size(153, 35)
        Me.btn_aceptar.TabIndex = 78
        Me.btn_aceptar.Text = "Aceptar"
        Me.btn_aceptar.UseVisualStyleBackColor = True
        Me.btn_aceptar.Visible = False
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label7.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label7.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(0, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(721, 38)
        Me.Label7.TabIndex = 79
        Me.Label7.Text = "Registro de Productos"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Producto
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DarkCyan
        Me.ClientSize = New System.Drawing.Size(721, 469)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btn_aceptar)
        Me.Controls.Add(Me.dg_datos)
        Me.Controls.Add(Me.btn_modificar)
        Me.Controls.Add(Me.btn_salir)
        Me.Controls.Add(Me.btn_grabar)
        Me.Controls.Add(Me.btn_nuevo)
        Me.Controls.Add(Me.txt_nombre)
        Me.Controls.Add(Me.cmd_siguiente)
        Me.Controls.Add(Me.cmd_ultimo)
        Me.Controls.Add(Me.cmd_anterior)
        Me.Controls.Add(Me.cmd_inicio)
        Me.Controls.Add(Me.cb_unidad)
        Me.Controls.Add(Me.txt_costo)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.txt_min)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txt_existencias)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txt_max)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txt_marca)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txt_id)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Producto"
        Me.Text = "Producto"
        CType(Me.PRODUCTOBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.InventariO_DBDataSet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dg_datos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txt_costo As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents txt_min As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents txt_existencias As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents txt_max As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txt_marca As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents txt_id As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents cb_unidad As ComboBox
    Friend WithEvents cmd_siguiente As Button
    Friend WithEvents cmd_ultimo As Button
    Friend WithEvents cmd_anterior As Button
    Friend WithEvents cmd_inicio As Button
    Friend WithEvents txt_nombre As TextBox
    Friend WithEvents btn_modificar As Button
    Friend WithEvents btn_salir As Button
    Friend WithEvents btn_grabar As Button
    Friend WithEvents btn_nuevo As Button
    Friend WithEvents SqlSelectCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlConnection1 As SqlClient.SqlConnection
    Friend WithEvents SqlInsertCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlUpdateCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlDeleteCommand1 As SqlClient.SqlCommand
    Friend WithEvents SqlDataAdapter1 As SqlClient.SqlDataAdapter
    Friend WithEvents InventariO_DBDataSet1 As INVENTARIO_DBDataSet
    Friend WithEvents PRODUCTOBindingSource As BindingSource
    Friend WithEvents PRODUCTOTableAdapter As INVENTARIO_DBDataSetTableAdapters.PRODUCTOTableAdapter
    Friend WithEvents dg_datos As DataGridView
    Friend WithEvents IdProductoDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents NombreDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents MarcaDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents UnidadDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents CostoDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents ExistencasDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents MinimoDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents MaximmoDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents btn_aceptar As Button
    Friend WithEvents Label7 As Label
End Class
