Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports System.IO

Module Conexiones
    Public conn As New SqlClient.SqlConnection
    Public cmd As New SqlClient.SqlCommand
    Public MSword As New Word.Application

    Public PathSalidas As New DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "\Documentos de salida")
    Public PathAyudas As New DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "\Ayudas")
    Public Function Conectar(Usuario As String, Contraseña As String)
        Try
            Dim sqlStr As String = "User ID=" & Usuario & "; Password=" & Contraseña & "; Initial Catalog=GestionClinica; Data Source=.\SQLEXPRESS"
            conn.ConnectionString = sqlStr
            conn.Open()

            Return 0
        Catch ex As Exception
            If (ex.ToString.Contains("0x80131904")) Then
                MsgBox("Usuario o Contraseña incorrecto, intete de nuevo ", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Inicio de sesión fallido.")
            Else
                MsgBox(ex.Message + vbCrLf + "Contacte con el departamento de informatica")
            End If

            Return 1
        End Try
    End Function

    Public Sub Desconectar()
        Try
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Public Sub SalirDelSistema()
        Try
            Dim Resultado As MsgBoxResult

            Resultado = MsgBox("¿Seguro que quiere salir del sistema?", MsgBoxStyle.YesNo, "SALIR DEL SISTEMA")

            If (Resultado = MsgBoxResult.Yes) Then
                End
            End If
        Catch ex As AggregateException

        End Try
      

    End Sub

    Public Function RegistrarMed(ByVal Nick As String, ByVal Contraseña As String, ByVal Cédula As String, ByVal Nombre As String, ByVal Apellido As String, ByVal Teléfono As String, ByVal Celular As String, ByVal Correo As String, ByVal Dirección As String)

        Dim Transac As SqlTransaction = conn.BeginTransaction()

        Try
            cmd = New SqlCommand("RegistrarMED", conn, Transac)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@Nick", Nick)
            cmd.Parameters.AddWithValue("@Contraseña", Contraseña)
            cmd.Parameters.AddWithValue("@Cédula", Cédula)
            cmd.Parameters.AddWithValue("@Nombre", Nombre)
            cmd.Parameters.AddWithValue("@Apellido", Apellido)
            cmd.Parameters.AddWithValue("@Teléfono", Teléfono)
            cmd.Parameters.AddWithValue("@Celular", Celular)
            cmd.Parameters.AddWithValue("@Correo", Correo)
            cmd.Parameters.AddWithValue("@Dirección", Dirección)


            cmd.ExecuteNonQuery()

            Transac.Commit()

            Return 0

        Catch ex As Exception

            MsgBox(ex.Message)
            Transac.Rollback()

            Return 1
        End Try

    End Function
    Public Function EditarSEC(ByVal Cédula As String, ByVal Nombre As String, ByVal Apellido As String, ByVal Teléfono As String, ByVal Celular As String, ByVal Correo As String, ByVal Dirección As String)
        Dim Transac As SqlTransaction

        Transac = conn.BeginTransaction()

        Try
            cmd = New SqlCommand("EditarSEC", conn, Transac)
            cmd.CommandType = CommandType.StoredProcedure

            cmd.Parameters.AddWithValue("@Cédula", Cédula)
            cmd.Parameters.AddWithValue("@Nombre", Nombre)
            cmd.Parameters.AddWithValue("@Apellido", Apellido)
            cmd.Parameters.AddWithValue("@Teléfono", Teléfono)
            cmd.Parameters.AddWithValue("@Celular", Celular)
            cmd.Parameters.AddWithValue("@Correo", Correo)
            cmd.Parameters.AddWithValue("@Dirección", Dirección)

            cmd.ExecuteNonQuery()
            Transac.Commit()

            Return 0
        Catch ex As Exception

            MsgBox(ex.Message)
            Transac.Rollback()
            Return 1
        End Try

        Transac.Dispose()
    End Function
    Public Function RegistrarSec(ByVal Nick As String, ByVal Contraseña As String, ByVal Cédula As String, ByVal Nombre As String, ByVal Apellido As String, ByVal Teléfono As String, ByVal Celular As String, ByVal Correo As String, ByVal Dirección As String)
        Dim Transac As SqlTransaction

        Transac = conn.BeginTransaction()

        Try



            cmd = New SqlCommand("RegistrarSEC", conn, Transac)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@Nick", Nick)
            cmd.Parameters.AddWithValue("@Contraseña", Contraseña)
            cmd.Parameters.AddWithValue("@Cédula", Cédula)
            cmd.Parameters.AddWithValue("@Nombre", Nombre)
            cmd.Parameters.AddWithValue("@Apellido", Apellido)
            cmd.Parameters.AddWithValue("@Teléfono", Teléfono)
            cmd.Parameters.AddWithValue("@Celular", Celular)
            cmd.Parameters.AddWithValue("@Correo", Correo)
            cmd.Parameters.AddWithValue("@Dirección", Dirección)

            cmd.ExecuteNonQuery()

            Transac.Commit()

            Return 0
        Catch ex As Exception
            MsgBox(ex.Message)
            Transac.Rollback()
            Return 1
        End Try
        Transac.Dispose()
    End Function

    Public Function VerificarTextbox(ByVal Form As TabPage, ByVal GroupBox As GroupBox) As Integer
        Dim control As Integer = 0
        For Each c As Control In Form.Controls
            If TypeOf c Is TextBox OrElse TypeOf c Is MaskedTextBox Then
                If c.Text = "" Then
                    control = control + 1

                    Exit For
                End If
            End If
        Next

        For Each c As Control In GroupBox.Controls
            If TypeOf c Is TextBox OrElse TypeOf c Is MaskedTextBox Then
                If c.Text = "" Then
                    control = control + 1
                End If
            End If
        Next

        If control <> 0 Then
            MsgBox("No pueden haber campos sin llenar!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ATENCIÓN!")
            Return 1
        Else
            Return 0
        End If

    End Function

    Public Sub LimpiarTextbox(ByVal TabPage, ByVal Group)
        Try
            For Each c As Control In TabPage.Controls
                If TypeOf c Is TextBox OrElse TypeOf c Is MaskedTextBox Then
                    c.Text = ""
                End If
            Next

            For Each c As Control In Group.Controls
                If TypeOf c Is TextBox OrElse TypeOf c Is MaskedTextBox Then
                    c.Text = ""
                End If
            Next


        Catch ex As Exception

        End Try


    End Sub

    Public Sub ActivarMED()
        Try
            cmd = New SqlCommand("ActivarMED", conn)
            cmd.CommandType = CommandType.StoredProcedure

            cmd.ExecuteNonQuery()
            Interfaz_Supervisor.CargarSupervisor()
            Interfaz_Secretaria.MedicosActivos()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub ActivarSEC()
        Try
            cmd = New SqlCommand("ActivarSEC", conn)
            cmd.CommandType = CommandType.StoredProcedure

            cmd.ExecuteNonQuery()

            Interfaz_Supervisor.CargarSupervisor()


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub DesactivarMED()
        Try
            cmd = New SqlCommand("DesactivarMED", conn)
            cmd.CommandType = CommandType.StoredProcedure

            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub DesactivarSEC()
        Try
            cmd = New SqlCommand("DesactivarSEC", conn)
            cmd.CommandType = CommandType.StoredProcedure

            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub ConsultarEspecialidades(ByVal CB)
        If (TypeOf CB Is ComboBox) OrElse (TypeOf CB Is ListBox) Then

            Dim dr As SqlDataReader

            Try
                cmd = New SqlCommand("ConsultarEspecialidades", conn)
                dr = cmd.ExecuteReader
                While dr.Read
                    CB.Items.Add(dr(0).ToString + " " + dr(1).ToString)
                End While

                dr.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

    End Sub

    Public Sub InsertarEspecialidad(ByVal Espe As String, ByVal Cédula As String)
        Try
            cmd.CommandText = "INSERT INTO dbo.Médicos_Especialidades_MSTR(Cédula_MED, Código_ESP) VALUES ('" + Cédula + "', " + Espe + ")"

            cmd.CommandType = CommandType.Text

            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub BorrarEspecialidad(ByVal Cédula As String)
        Try
            cmd.CommandText = " DELETE dbo.Médicos_Especialidades_MSTR WHERE Cédula_MED = '" + Cédula + "'"

            cmd.CommandType = CommandType.Text

            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Public Sub Medicos_Especialidades(ByVal Cédula As String, ByVal LB As ListBox, ByVal CB As ComboBox)
        Dim dr As SqlDataReader
        Dim count As Integer = 0

        Try
            cmd = New SqlCommand("Especialidades_Medicos", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@Cédula", Cédula)

            dr = cmd.ExecuteReader

            While (dr.Read)
                LB.Items.Add(dr(0).ToString + " " + dr(1).ToString)
            End While

            dr.Close()

            cmd.CommandText = "SELECT * FROM dbo.Especialidades_MSTR"
            cmd.CommandType = CommandType.Text

            dr = cmd.ExecuteReader

            While (dr.Read)
                count = 0
                For i = 0 To LB.Items.Count - 1

                    If (dr(0).ToString + " " + dr(1).ToString) = (LB.Items.Item(i)) Then
                        count = count + 1
                    End If
                Next
                If (count = 0) Then
                    CB.Items.Add(dr(0).ToString + " " + dr(1).ToString)
                End If
            End While

            dr.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Public Function EditarMedico(ByVal Cédula As String, ByVal Nombre As String, ByVal Apellido As String, ByVal Teléfono As String, ByVal Celular As String, ByVal Correo As String, ByVal Dirección As String) As Integer
        Dim Transac As SqlTransaction

        Transac = conn.BeginTransaction()

        Try

            cmd = New SqlCommand("EdiMED", conn, Transac)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@Cédula", Cédula)
            cmd.Parameters.AddWithValue("@Nombre", Nombre)
            cmd.Parameters.AddWithValue("@Apellido", Apellido)
            cmd.Parameters.AddWithValue("@Teléfono", Teléfono)
            cmd.Parameters.AddWithValue("@Celular", Celular)
            cmd.Parameters.AddWithValue("@Correo", Correo)
            cmd.Parameters.AddWithValue("@Dirección", Dirección)

            cmd.ExecuteNonQuery()

            Transac.Commit()

            Return 0
        Catch ex As Exception
            MsgBox(ex.Message + vbCrLf + "Error al Editar el médico!", MsgBoxStyle.Critical + MsgBoxStyle.RetryCancel, "ERROR!")
            Transac.Rollback()
            Return 1
        End Try
    End Function
    

    Public Function EliminarSEC(ByVal Nick As String, ByVal Cédula As String)
        Dim Transac As SqlTransaction

        Transac = conn.BeginTransaction

        Try


            cmd = New SqlCommand("EliminarSEC", conn, Transac)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@Nick", Nick)
            cmd.Parameters.AddWithValue("@Cédula", Cédula)

            cmd.ExecuteNonQuery()

            Transac.Commit()
            Return 0
        Catch ex As Exception
            Transac.Rollback()
            Return 1
        End Try
    End Function

    Public Sub DBUSER_DEL(ByVal Cédula As String)
        Try
            cmd.CommandText = "DROP USER [" + Cédula + "]"
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()

        Catch ex As Exception
        End Try
    End Sub

    Public Sub CambiarContrase(Nick As String, Contraseña As String)
        Dim Transac As SqlTransaction = conn.BeginTransaction

        Try
            cmd = New SqlCommand("ContraEdi", conn, Transac)
            cmd.CommandType = CommandType.StoredProcedure
            With cmd.Parameters
                .AddWithValue("@Nick", Nick)
                .AddWithValue("@Contraseña", Contraseña)
            End With

            cmd.ExecuteNonQuery()

            Transac.Commit()
        Catch ex As Exception
            MsgBox(ex.Message)
            Transac.Rollback()
        End Try
    End Sub
End Module
