
Imports System.ComponentModel
Imports System.Data.SqlClient

Public Class Respaldo_Restauracion
    'Dim conexionsql As New SqlConnection("Data Source ='CARDENAS-PC'; Initial Catalog = 'master'; Integrated security = true")
    ' Dim comando As SqlCommand = conexionsql.CreateCommand
    Private Sub btn_seleccion_Click(sender As Object, e As EventArgs) Handles btn_seleccion.Click
        sfd_grabar.Filter = "Archivos Respaldo (*.bak)|*.bak| Todos (*.*)|*.* "
        sfd_grabar.FilterIndex = 1
        sfd_grabar.FileName = "Respado_" & Today.Date.Day & "_" & Today.Date.Month & "_" & Today.Date.Year
        If (sfd_grabar.ShowDialog = DialogResult.OK) Then
            lbl_ruta.Text = sfd_grabar.FileName
            If lbl_ruta.Text.Length > 3 Then
                btn_respaldo.Enabled = True
            Else
                btn_respaldo.Enabled = False
            End If
        End If
    End Sub

    Private Sub btn_respaldo_Click(sender As Object, e As EventArgs) Handles btn_respaldo.Click
        Dim resultado As DialogResult
        Try
            resultado = MessageBox.Show("¿Estas Seguro De Hacer El Respaldo?", "Respaldo", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (resultado = DialogResult.Yes) Then
                Cursor.Current = Cursors.WaitCursor
                Using conexion As New SqlConnection("Data Source ='.'; Initial Catalog = 'master'; Integrated security = true")
                    conexion.Open()
                    Using comando As New SqlCommand
                        comando.Connection = conexion
                        'comando.CommandText = "use master"
                        'comando.ExecuteNonQuery()
                        comando.CommandText = "BACKUP DATABASE [INVENTARIO_DB] TO DISK = '" & lbl_ruta.Text & "' WITH STATS = 5"
                        comando.ExecuteNonQuery()
                        MessageBox.Show("El Respaldo de la base de datos fue realizado satisfactoriamente", "Respaldo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Cursor.Current = Cursors.Default
                    End Using
                    conexion.Close()
                End Using
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            '   Finally
            ' conexionsql.Close()
            ' conexionsql.Dispose()
        End Try
    End Sub

    Private Sub btn_seleccionar_Click(sender As Object, e As EventArgs) Handles btn_seleccionar.Click
        ofd_abrir.Filter = "Archivos Respaldo (*.bak)|*.bak| Todos (*.*)|*.* "
        ofd_abrir.FilterIndex = 1
        If (ofd_abrir.ShowDialog = DialogResult.OK) Then
            lbl_ruta_a.Text = ofd_abrir.FileName
            If lbl_ruta_a.Text.Length > 3 Then
                btn_restaura.Enabled = True
            Else
                btn_restaura.Enabled = False
            End If
        End If
    End Sub

    Private Sub Respaldo_Restauracion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lbl_ruta.Text = "C:\"
        lbl_ruta_a.Text = "C:\"
        btn_respaldo.Enabled = False
        btn_restaura.Enabled = False
    End Sub

    Private Sub btn_restaura_Click(sender As Object, e As EventArgs) Handles btn_restaura.Click
        SqlConnection.ClearAllPools()
        Dim resultado As DialogResult
        Try
            resultado = MessageBox.Show("¿Estas De Restaurar La Base De Datos?", "Restauración", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (resultado = DialogResult.Yes) Then
                Cursor.Current = Cursors.WaitCursor
                Using conexion As New SqlConnection("Data Source ='.'; Initial Catalog = 'master'; Integrated security = true")
                    conexion.Open()
                    Using comando As New SqlCommand
                        comando.Connection = conexion
                        comando.CommandText = "use master"
                        comando.ExecuteNonQuery()
                        comando.CommandText = "RESTORE DATABASE INVENTARIO_DB FROM DISK = '" & lbl_ruta_a.Text & "' WITH  STATS = 5, REPLACE"
                        comando.ExecuteNonQuery()
                        Cursor.Current = Cursors.Default
                        MessageBox.Show("✌ Se Restauro la base de datos satisfactoriamente 😉", "Restauración", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End Using
                    conexion.Close()
                End Using
            End If
        Catch ex As Exception
            MsgBox(ex.Message & "😒😢")
            '       Finally
            '          conexionsql.Close()
            '          conexionsql.Dispose()
        End Try
    End Sub
    Private Sub Respaldo_Restauracion_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Dispose()
    End Sub

    Private Sub Respaldo_Restauracion_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 Then
            System.Diagnostics.Process.Start("D:\Desa e Imple de Sis\Proyecto\Sis_Admin_In\Ayuda.chm")
        End If
    End Sub
End Class