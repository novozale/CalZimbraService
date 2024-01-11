' NOTE: You can use the "Rename" command on the context menu to change the class name "Service1" in code, svc and config file together.
' NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.vb at the Solution Explorer and start debugging.
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Xml

Public Class Service1
    Implements ICalendarZimbraService
    Dim connString As String = "Data Source=.;server=sqlcls;Initial Catalog=ScaDataDB;User ID=sa;Password=sqladmin"
    Dim AccName As String = "admin"
    Dim AccPass As String = "QpRa2OKlI+@X_"
    Dim requestUriString As String = "https://mail.skandikagroup.ru/service/soap"

    Public Sub New()
    End Sub

    Private Function IsAuthorised(MyLogin As String, MyService As String) As Boolean
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка - есть ли у данного пользователя право на использование сервиса
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim ds As New DataSet()
        Dim MyAuth As Boolean = False

        Try
            MySQLStr = "dbo.spp_Services_GetAuthInfo"
            Using MyConn As SqlConnection = New SqlConnection(connString)
                Try
                    Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandTimeout = 1800
                        cmd.Parameters.AddWithValue("@MyLogin", MyLogin)
                        cmd.Parameters.AddWithValue("@MyService", MyService)
                        Using da As New SqlDataAdapter()
                            da.SelectCommand = cmd
                            da.Fill(ds)
                            If ds.Tables(0).Rows.Count <> 0 Then
                                MyAuth = True
                            End If
                        End Using
                    End Using
                Catch ex As Exception
                    EventLog.WriteEntry("CalZimbraService", "IsAuthorised --1--> " & ex.Message)
                Finally
                    MyConn.Close()
                End Try
            End Using
        Catch ex As Exception
            EventLog.WriteEntry("CalZimbraService", "IsAuthorised --2--> " & ex.Message)
        End Try
        Return MyAuth
    End Function

    Public Function CreateCalendarEvent(MyEvent As CreateCalendarEventType) As String Implements ICalendarZimbraService.CreateCalendarEvent
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание события в календаре
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MyLogin As String
        Dim MyService As String
        Dim MyId As String
        Dim MyToken As String

        MyLogin = MyEvent.Login
        MyService = "CalZimbraService"
        MyId = ""

        If IsAuthorised(MyLogin, MyService) Then
            '------------в случае авторизации - создание события
            Try
                MyToken = ZGetToken()
                If MyToken.Equals("") Then
                    Return MyId
                Else
                    If MyEvent.CalendarEventIDOld.Equals("") Then
                    Else
                        MyId = ZDeleteCalendarEvent(MyEvent.CalendarEventIDOld, MyEvent.Email, MyToken)
                    End If
                    MyId = ""
                    MyId = ZCreateCalendarEvent(MyEvent, MyToken)
                    Return MyId
                End If
            Catch ex As Exception
                EventLog.WriteEntry("CalendarZimbraServices", "CreateCalendarEvent --1--> " & ex.Message)
                Return ""
            End Try
        Else
            EventLog.WriteEntry("CalendarZimbraServices", "CreateCalendarEvent --2--> " & "Not authorized")
            Return ""
        End If
    End Function

    Public Function DeleteCalendarEvent(MyEvent As DeleteCalendarEventType) As String Implements ICalendarZimbraService.DeleteCalendarEvent
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление события в календаре
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MyLogin As String
        Dim MyService As String
        Dim MyId As String
        Dim MyToken As String

        MyLogin = MyEvent.Login
        MyService = "CalZimbraService"
        MyId = ""

        If IsAuthorised(MyLogin, MyService) Then
            '------------в случае авторизации - удаление события
            Try
                MyToken = ZGetToken()
                If MyToken.Equals("") Then
                    Return MyId
                Else
                    MyId = ZDeleteCalendarEvent(MyEvent.CalendarEventIDOld, MyEvent.Email, MyToken)
                    Return MyEvent.CalendarEventIDOld
                End If
            Catch ex As Exception
                EventLog.WriteEntry("CalendarZimbraServices", "DeleteCalendarEvent --1--> " & ex.Message)
                Return ""
            End Try
        Else
            EventLog.WriteEntry("CalendarZimbraServices", "DeleteCalendarEvent --2--> " & "Not authorized")
            Return ""
        End If
    End Function

    Private Function ZCreateCalendarEvent(MyEvent As CreateCalendarEventType, MyToken As String) As String
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Непосредственное создание события в календаре
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MyId As String
        Dim MyRez As String
        Dim str As String

        MyId = ""
        str = "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"">"
        str = str & "<SOAP-ENV:Header>"
        str = str & "<context xmlns=""urn:zimbra"">"
        str = str & "<authToken>" & MyToken & "</authToken>"
        str = str & "</context>"
        str = str & "</SOAP-ENV:Header>"
        str = str & "<SOAP-ENV:Body>"

        str = str & "<CreateAppointmentRequest xmlns=""urn:zimbraMail"" forcesend=""1"">"
        str = str & "<m su=""" & MyEvent.Subject & """>"
        str = str & "<e a=""" & MyEvent.Email & """ t=""t""/>"
        str = str & "<mp ct=""multipart/alternative"">"
        str = str & "<mp content=""" & MyEvent.Body & """ ct=""text/plain""/>"
        str = str & "</mp>"
        str = str & "<inv status=""CONF"" isOrg=""1"" allDay=""0"" draft=""0"" fb=""B"">"
        str = str & "<comp xmlns=""urn:zimbraMail"" loc=""" & "Russia" & """ name=""" & MyEvent.Subject & """ noBlob=""1"">"
        str = str & "<s d=""" & MyEvent.Start.Year.ToString & Right("00" & MyEvent.Start.Month.ToString, 2) & Right("00" & MyEvent.Start.Day.ToString, 2) & "T090000"" tz=""Europe/Moscow""/>"
        str = str & "<e d=""" & MyEvent.Finish.Year.ToString & Right("00" & MyEvent.Finish.Month.ToString, 2) & Right("00" & MyEvent.Finish.Day.ToString, 2) & "T173000"" tz=""Europe/Moscow""/>"
        str = str & "<alarm action=""DISPLAY"">"
        str = str & "<trigger>"
        str = str & "<rel neg=""1"" m=""15"" related=""START""/>"
        str = str & "</trigger>"
        str = str & "</alarm>"
        str = str & "</comp>"
        str = str & "</inv>"
        str = str & "</m>"
        str = str & "</CreateAppointmentRequest>"

        str = str & "</SOAP-ENV:Body>"
        str = str & "</SOAP-ENV:Envelope>"

        Dim uTF8Encoding As New System.Text.UTF8Encoding()
        Dim bytes As Byte() = uTF8Encoding.GetBytes(str)

        ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

        Dim httpWebRequest As Net.HttpWebRequest = CType(WebRequest.Create(requestUriString), HttpWebRequest)
        httpWebRequest.Headers.Add(HttpRequestHeader.AcceptEncoding, "gzip, deflate")
        httpWebRequest.Headers.Add("SOAPAction", requestUriString)
        httpWebRequest.Method = "POST"
        httpWebRequest.ContentType = "Text/ Xml; charset=UTF-8"
        httpWebRequest.ContentLength = CLng(bytes.Length)

        Try
            Dim requestStream As Stream = httpWebRequest.GetRequestStream()
            requestStream.Write(bytes, 0, bytes.Length)
            requestStream.Close()
            Dim httpWebResponse As HttpWebResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)
            Dim streamReader As StreamReader = New StreamReader(httpWebResponse.GetResponseStream(), uTF8Encoding)

            Dim xmlDocument As New System.Xml.XmlDocument()
            xmlDocument.LoadXml(streamReader.ReadToEnd())
            httpWebResponse.Close()

            Dim root As XmlNode = xmlDocument.DocumentElement
            MyId = root.ChildNodes(1).ChildNodes(0).Attributes.ItemOf("invId").Value
            Return MyId
        Catch ex As Exception
            EventLog.WriteEntry("CalendarZimbraServices", "ZCreateCalendarEvent --1--> " & ex.Message)
            Return MyId
        End Try
    End Function

    Private Function ZDeleteCalendarEvent(MyId As String, MyEmail As String, MyToken As String) As String
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Непосредственное удаление события в календаре
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim Str As String
        Dim MyRez As String

        MyRez = ""
        Str = "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"">"
        Str = Str & "<SOAP-ENV:Header>"
        Str = Str & "<context xmlns=""urn:zimbra"">"
        Str = Str & "<authToken>" & MyToken & "</authToken>"
        Str = Str & "</context>"
        Str = Str & "</SOAP-ENV:Header>"
        Str = Str & "<SOAP-ENV:Body>"

        Str = Str & "<CancelAppointmentRequest xmlns=""urn:zimbraMail"" id=""" & MyId & """ comp=""0"">"
        Str = Str & "<m su=""" & "CRM APPOINTMENT: " & "Удаление CRM события из календаря" & """>"
        Str = Str & "<e a=""" & MyEmail & """ t=""t""/>"
        Str = Str & "<mp ct=""multipart/alternative"">"
        Str = Str & "<mp content=""" & "Удаление CRM события из календаря" & """ ct=""text/plain""/>"
        Str = Str & "</mp>"
        Str = Str & "</m>"
        Str = Str & "</CancelAppointmentRequest>"

        Str = Str & "</SOAP-ENV:Body>"
        Str = Str & "</SOAP-ENV:Envelope>"

        Dim uTF8Encoding As New System.Text.UTF8Encoding()
        Dim bytes As Byte() = uTF8Encoding.GetBytes(Str)

        ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

        Dim httpWebRequest As Net.HttpWebRequest = CType(WebRequest.Create(requestUriString), HttpWebRequest)
        httpWebRequest.Headers.Add(HttpRequestHeader.AcceptEncoding, "gzip, deflate")
        httpWebRequest.Headers.Add("SOAPAction", requestUriString)
        httpWebRequest.Method = "POST"
        httpWebRequest.ContentType = "Text/ Xml; charset=UTF-8"
        httpWebRequest.ContentLength = CLng(bytes.Length)

        Try
            Dim requestStream As Stream = httpWebRequest.GetRequestStream()
            requestStream.Write(bytes, 0, bytes.Length)
            requestStream.Close()
            Dim httpWebResponse As HttpWebResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)
            Dim streamReader As StreamReader = New StreamReader(httpWebResponse.GetResponseStream(), uTF8Encoding)

            Dim xmlDocument As New System.Xml.XmlDocument()
            xmlDocument.LoadXml(streamReader.ReadToEnd())
            httpWebResponse.Close()
            Return MyRez
        Catch ex As Exception
            EventLog.WriteEntry("CalendarZimbraServices", "ZDeleteCalendarEvent --1--> " & ex.Message)
            Return ex.Message
        End Try
    End Function

    Private Function ZGetToken() As String
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение токена для работы с Zimbra
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim str As String
        Dim MyToken As String

        MyToken = ""
        str = "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"">"
        str = str & "<SOAP-ENV:Header>"
        str = str & "<context xmlns=""urn:zimbra"">"
        str = str & "</context>"
        str = str & "</SOAP-ENV:Header>"
        str = str & "<SOAP-ENV:Body>"
        str = str & "<AuthRequest xmlns=""urn:zimbraAccount"">"
        str = str & "<account by=""name"">" & AccName & "</account>"
        str = str & "<password>" & AccPass & "</password>"
        str = str & "</AuthRequest>"
        str = str & "</SOAP-ENV:Body>"
        str = str & "</SOAP-ENV:Envelope>"

        Dim uTF8Encoding As New System.Text.UTF8Encoding()
        Dim bytes As Byte() = uTF8Encoding.GetBytes(str)

        ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

        Dim httpWebRequest As Net.HttpWebRequest = CType(WebRequest.Create(requestUriString), HttpWebRequest)
        httpWebRequest.Headers.Add(HttpRequestHeader.AcceptEncoding, "gzip, deflate")
        httpWebRequest.Headers.Add("SOAPAction", requestUriString)
        httpWebRequest.Method = "POST"
        httpWebRequest.ContentType = "Text/ Xml; charset=UTF-8"
        httpWebRequest.ContentLength = CLng(bytes.Length)

        Try
            Dim requestStream As Stream = httpWebRequest.GetRequestStream()
            requestStream.Write(bytes, 0, bytes.Length)
            requestStream.Close()
            Dim httpWebResponse As HttpWebResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)
            Dim streamReader As StreamReader = New StreamReader(httpWebResponse.GetResponseStream(), uTF8Encoding)

            Dim xmlDocument As New System.Xml.XmlDocument()
            xmlDocument.LoadXml(streamReader.ReadToEnd())
            httpWebResponse.Close()

            Dim root As XmlNode = xmlDocument.DocumentElement
            MyToken = root.ChildNodes(1).ChildNodes(0).ChildNodes(0).InnerText
        Catch ex As Exception
            EventLog.WriteEntry("CalendarZimbraServices", "ZGetToken --1--> " & ex.Message)
        End Try
        Return MyToken
    End Function

End Class
