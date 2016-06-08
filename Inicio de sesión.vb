Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop
Public Class IniciarSesion

    Private Sub IniciarSesion_Enter(sender As Object, e As EventArgs) Handles Me.Enter
        Button1_Click(sender, e)
    End Sub

    Private Sub IniciarSesion_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        SalirDelSistema()
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Conectar(TextBox1.Text.ToString, TextBox2.Text.ToString) = 0 Then

            Dim dr As SqlDataReader
            Try
                cmd = New SqlCommand("VerRoles", conn)
                cmd.CommandType = CommandType.StoredProcedure


                dr = cmd.ExecuteReader()

                While (dr.Read)
                    If (dr(1) = "8") OrElse (dr(1) = "16384") Then
                        dr.Close()
                        Interfaz_Supervisor.Show()
                        Me.Hide()
                        Exit While
                    Else
                        If (dr(0).ToString = "MédicosGrupo") Then

                            dr.Close()
                            Médico_1.Show()
                            ActivarMED()
                            Me.Hide()
                            Exit While
                        Else
                            If (dr(0).ToString = "SecretariaGrupo") Then
                                dr.Close()
                                Interfaz_Secretaria.Show()
                                ActivarSEC()
                                Me.Hide()
                                Exit While
                            End If
                        End If
                    End If
                End While



            Catch ex As Exception
                MsgBox(ex.ToString)
                Desconectar()
            End Try
        Else
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        SalirDelSistema()
    End Sub

    Private Sub SALIRDELSISTEMAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SALIRDELSISTEMAToolStripMenuItem.Click
        SalirDelSistema()
    End Sub

    Private Sub IniciarSesion_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub AyudaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AyudaToolStripMenuItem.Click
        Try
            Dim path As String = PathAyudas.FullName + "\Login.docx"
            Process.Start(path)
        Catch ex As Exception

        End Try
    End Sub
End Class
