Imports System.IO 'esta libreria nos va a servir para poder activar el commandialog
Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Module conexionGlobal
    Public tipo As String
    Public servidor As String = "."
    Public conexionsql As New SqlConnection("Data Source ='" & servidor & "'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true")
    'Dim conexionsql As New SqlConnection("Data Source ='.'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true ")
    Dim comando As SqlCommand = conexionsql.CreateCommand

    Function llenarExcel(ByVal ElGrid As DataGridView) As Boolean
        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        Try
            'Añadimos el Libro al programa, y la hoja al libro
            exLibro = exApp.Workbooks.Add
            exHoja = exLibro.Worksheets.Add()

            ' ¿Cuantas columnas y cuantas filas?
            Dim NCol As Integer = ElGrid.ColumnCount
            Dim NRow As Integer = ElGrid.RowCount

            'Aqui recorremos todas las filas, y por cada fila todas las columnas
            'y vamos escribiendo.
            Dim cont As Integer = 1
            For i As Integer = 0 To NCol - 1
                If (ElGrid.Columns(i).Visible = True) Then
                    exHoja.Cells.Item(1, cont) = ElGrid.Columns(i).HeaderText

                    cont += 1
                End If
            Next


            For Fila As Integer = 0 To NRow - 1
                Dim c As Integer = 0
                For Col As Integer = 0 To NCol - 1
                    If (ElGrid.Rows(0).Cells(Col).Visible) Then
                        exHoja.Cells.Item(Fila + 2, c + 1) = ElGrid.Item(Col, Fila).Value
                        c += 1
                    End If
                Next
            Next
            'Titulo en negrita, Alineado al centro y que el tamaño de la columna
            'se ajuste al texto
            exHoja.Rows.Item(1).Font.Bold = 1
            exHoja.Rows.Item(1).HorizontalAlignment = 3
            exHoja.Columns.AutoFit()
            'Aplicación visible
            exApp.Application.Visible = True
            exHoja = Nothing
            exLibro = Nothing
            exApp = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")
            Return False
        End Try
        Return True
    End Function

    Sub importarTxt()
        If conexionsql.State <> ConnectionState.Open Then
            conexionsql.Open()
        End If
        Dim cad As String = Nothing
        Dim DocumentoDialog As New OpenFileDialog()

        With DocumentoDialog
            .InitialDirectory = "D:\ProyectoHope"
            .Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
            .FilterIndex = 2
            .RestoreDirectory = True
            .Title = "Abrir Archivo"
            .ShowDialog()
        End With
        If DocumentoDialog.FileName.ToString <> "" Then
            Try
                'conexionsql.ConnectionString = "Data Source ='.'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true "
                'conexionsql.Open()
                cad = DocumentoDialog.FileName
                If (cad IsNot Nothing) Then
                    comando.CommandText = "BULK INSERT PRODUCTO FROM '" & cad & "' WITH ( FIELDTERMINATOR= ',', ROWTERMINATOR = '\n' );"
                    comando.ExecuteNonQuery()
                    MessageBox.Show("SE INSERTÓ EL REGISTRO DE FORMA CORRECTA", "ÉXITO")
                    Producto.ShowDialog()

                Else
                    MessageBox.Show("NO SE SELECCIONO ALGUN ARCHIVO", "ATENCIÓN")
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Information, "Informacion")
            Finally
                conexionsql.Close()
            End Try
        End If
        ' MsgBox("Se ha cargado la importacion correctamente", MsgBoxStyle.Information, "Importado con exito")
    End Sub
End Module
