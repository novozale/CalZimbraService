' NOTE: You can use the "Rename" command on the context menu to change the interface name "IService1" in both code and config file together.
<ServiceContract()>
Public Interface ICalendarZimbraService
    '////////////////////////////////////////////////////////////////////////////////////
    '//
    '// Функции
    '//
    '////////////////////////////////////////////////////////////////////////////////////

    <OperationContract()>
    Function CreateCalendarEvent(ByVal MyEvent As CreateCalendarEventType) As String

    <OperationContract()>
    Function DeleteCalendarEvent(ByVal MyEvent As DeleteCalendarEventType) As String

    ' TODO: Add your service operations here

End Interface

'////////////////////////////////////////////////////////////////////////////////////
'//
'// Параметры
'//
'////////////////////////////////////////////////////////////////////////////////////

<DataContract()>
Public Class CreateCalendarEventType
    Private CalendarEventIDOld_value As String                      '--ID существующего события (если нет - пустая строка)
    Private Subject_value As String                                 '--Subject
    Private Body_value As String                                    '--Body
    Private Start_value As DateTime                                 '--начало события
    Private Finish_value As DateTime                                '--конец события
    Private Timezone_value As String                                '--Timezone
    Private Email_value As String                                   '--Email пользователя
    Private Login_value As String                                   '--Login сервисов

    <DataMember()>
    Public Property CalendarEventIDOld() As String                  '--ID существующего события (если нет - пустая строка)
        Get
            Return CalendarEventIDOld_value
        End Get
        Set(value As String)
            CalendarEventIDOld_value = value
        End Set
    End Property

    <DataMember()>
    Public Property Subject() As String                             '--Subject
        Get
            Return Subject_value
        End Get
        Set(value As String)
            Subject_value = value
        End Set
    End Property

    <DataMember()>
    Public Property Body() As String                                '--Body
        Get
            Return Body_value
        End Get
        Set(value As String)
            Body_value = value
        End Set
    End Property

    <DataMember()>
    Public Property Start() As DateTime                              '--Начало события
        Get
            Return Start_value
        End Get
        Set(value As DateTime)
            Start_value = value
        End Set
    End Property

    <DataMember()>
    Public Property Finish() As DateTime                              '--Конец события
        Get
            Return Finish_value
        End Get
        Set(value As DateTime)
            Finish_value = value
        End Set
    End Property

    <DataMember()>
    Public Property Timezone() As String                              '--Timezone
        Get
            Return Timezone_value
        End Get
        Set(value As String)
            Timezone_value = value
        End Set
    End Property

    <DataMember()>
    Public Property Email() As String                                  '--Email
        Get
            Return Email_value
        End Get
        Set(value As String)
            Email_value = value
        End Set
    End Property

    <DataMember()>
    Public Property Login() As String                                  '--Login
        Get
            Return Login_value
        End Get
        Set(value As String)
            Login_value = value
        End Set
    End Property
End Class

<DataContract()>
Public Class DeleteCalendarEventType
    Private CalendarEventIDOld_value As String                      '--ID существующего события (если нет - пустая строка)
    Private Email_value As String                                   '--Email пользователя
    Private Login_value As String                                   '--Login сервисов

    <DataMember()>
    Public Property CalendarEventIDOld() As String                  '--ID существующего события (если нет - пустая строка)
        Get
            Return CalendarEventIDOld_value
        End Get
        Set(value As String)
            CalendarEventIDOld_value = value
        End Set
    End Property

    <DataMember()>
    Public Property Email() As String                                  '--Email
        Get
            Return Email_value
        End Get
        Set(value As String)
            Email_value = value
        End Set
    End Property

    <DataMember()>
    Public Property Login() As String                                  '--Login
        Get
            Return Login_value
        End Get
        Set(value As String)
            Login_value = value
        End Set
    End Property
End Class
