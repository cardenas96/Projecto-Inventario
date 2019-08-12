Imports System.Data.SqlClient
Public Class Inicio_Sesion
    'Dim conexionsqls2 As New SqlConnection("Data Source ='.'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true")
    Dim comando As SqlCommand = conexionsql.CreateCommand
    Dim lector As SqlDataReader
    Public bandera As Boolean = False
    Private Sub Inicio_Sesion_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Sub btn_aceptar_Click(sender As Object, e As EventArgs) Handles btn_aceptar.Click
        If txt_usuario.Text = "" Or txt_cusuario.Text = "" Then
            MessageBox.Show("Llene el campo vacio", "¡Error Campo Vacio!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            If txt_usuario.Text = "" Then
                txt_usuario.Focus()
            Else
                txt_cusuario.Focus()
            End If
        Else
            If bandera = False Then
                conexionsql.Open()
            Else
                conexionsql.ConnectionString = "Data Source ='CARDENAS-PC'; Initial Catalog = 'INVENTARIO_DB'; Integrated security = true"
                conexionsql.Open()
                bandera = False
            End If

            Dim usuario, pass, u, c, t As String
            usuario = txt_usuario.Text
            pass = txt_cusuario.Text
            Try
                comando.CommandText = "SELECT nUsuario,cUsuario, tipo FROM USUARIOS WHERE nusuario = '" & usuario & "'"
                lector = comando.ExecuteReader
                lector.Read()
                u = lector(0).ToString
                c = lector(1).ToString
                t = lector(2).ToString

                lector.Close()
                conexionsql.Close()
                If usuario = u And pass = c Then
                    Visible = False
                    'SqlConnection.ClearAllPools()
                    tipo = t
                    Principal.ShowDialog()
                    MsgBox("Termine")
                    conexionsql.Dispose()
                    conexionsql.Close()
                Else
                    MessageBox.Show("Usuario o Contraseña Incorrecto", "¡Error De Informacion!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    conexionsql.Close()
                End If
            Catch ex As Exception
            End Try

        End If
    End Sub

    Private Sub btn_salir_Click(sender As Object, e As EventArgs) Handles btn_salir.Click
        Dispose()
    End Sub

    Private Sub txt_cusuario_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_cusuario.KeyPress
        If Asc(e.KeyChar) = 13 Then
            btn_aceptar.PerformClick()
        End If
    End Sub

    Private Sub txt_usuario_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_usuario.KeyPress
        If Asc(e.KeyChar) = 13 Then
            txt_cusuario.Focus()
        End If
    End Sub
End Class