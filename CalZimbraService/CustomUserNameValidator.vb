
Imports System.Data.SqlClient
Imports System.IdentityModel.Selectors
Imports System.Security.Cryptography
Public Class CustomUserNameValidator
    Inherits UserNamePasswordValidator

    Dim connString As String = "Data Source=.;server=sqlcls;Initial Catalog=ScaDataDB;User ID=sa;Password=sqladmin"

    Public Overrides Sub Validate(userName As String, password As String)
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка логина и пароля
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim ds As New DataSet()
        Dim MyPassword As String

        If ((userName.Equals("") = False) And (password.Equals("") = False)) Then
            MyPassword = GetMD5Str(password)
            '---проверка логина и пароля----------------------
            Try
                MySQLStr = "dbo.spp_Services_GetLoginPasswords"
                Using MyConn As SqlConnection = New SqlConnection(connString)
                    Try
                        Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                            cmd.CommandType = CommandType.StoredProcedure
                            cmd.CommandTimeout = 1800
                            cmd.Parameters.AddWithValue("@MyLogin", userName)
                            cmd.Parameters.AddWithValue("@MyMD5Password", MyPassword)
                            Using da As New SqlDataAdapter()
                                da.SelectCommand = cmd
                                da.Fill(ds)
                                If ds.Tables(0).Rows.Count <> 0 Then
                                Else
                                    Throw New FaultException("Unknown Username or Incorrect Password")
                                End If
                            End Using
                        End Using
                    Catch ex As Exception
                        EventLog.WriteEntry("CalZimbraService", "CheckLoginPassword --1--> " & ex.Message)
                    Finally
                        MyConn.Close()
                    End Try
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("CalZimbraService", "CheckLoginPassword --2--> " & ex.Message)
            End Try
        Else
            Throw New FaultException("Unknown Username or Incorrect Password")
        End If
    End Sub

    Private Function GetMD5Str(MyStr As String) As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение строки MD5 для строки - параметра
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Using md5Hash As MD5 = MD5.Create()


            Dim data As Byte() = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(MyStr))
            Dim sBuilder As New StringBuilder()
            Dim i As Integer
            For i = 0 To data.Length - 1
                sBuilder.Append(data(i).ToString("x2"))
            Next i
            GetMD5Str = sBuilder.ToString
        End Using
    End Function
End Class
