Imports System.Net.Mail
Imports System.Configuration
Imports System.Web.Configuration

Public Class Settings
    Private Shared mconfig As Configuration = Nothing


    Public Shared Property AdminEmail() As MailAddress
        Get
            Dim s = ConfigProperty("AdminEmailAddress")
            If String.IsNullOrWhiteSpace(s) Then
                s = "none@none.com"
                ConfigProperty("AdminEmailAddress") = s
            End If
            Dim n = ConfigProperty("AdminEmailName")
            If String.IsNullOrWhiteSpace(n) Then
                n = "none"
                ConfigProperty("AdminEmailName") = n
            End If
            Return New MailAddress(s, n)
        End Get
        Set(ByVal value As MailAddress)
            ConfigProperty("AdminEmailAddress") = value.Address
            ConfigProperty("AdminEmailName") = value.DisplayName
        End Set
    End Property

    Public Shared ReadOnly Property ConnStringName() As String
        Get
            Return "MainDBConnString"
        End Get
    End Property
    Public Shared Function GetConfiguration() As Configuration
        If mconfig Is Nothing Then
            Try
                mconfig = WebConfigurationManager.OpenWebConfiguration("~")
            Catch ex As Exception
                Try
                    mconfig = ConfigurationManager.OpenExeConfiguration("")
                Catch ex2 As Exception
                    mconfig = WebConfigurationManager.OpenWebConfiguration(Nothing)
                End Try
                'config = WebConfigurationManager.OpenWebConfiguration("c:\")
            End Try
        End If
        Return mconfig
    End Function
    Private Shared Property NamedConnString(ByVal ConnName As String) As String
        Get
            Dim config = GetConfiguration()
            Dim CSS As ConnectionStringSettings = config.ConnectionStrings.ConnectionStrings(ConnName)
            If CSS IsNot Nothing Then
                Return CSS.ConnectionString
            End If
            Return ""
        End Get
        Set(ByVal value As String)
            Dim config = GetConfiguration()
            Dim CSS As ConnectionStringSettings = config.ConnectionStrings.ConnectionStrings(ConnName)
            If CSS Is Nothing Then
                CSS = New ConnectionStringSettings(ConnName, value, "System.Data.SqlClient")
                config.ConnectionStrings.ConnectionStrings.Add(CSS)
            Else
                CSS.ConnectionString = value
            End If
            config.Save()

        End Set
    End Property
    Public Shared Property ConfigProperty(ByVal Name As String) As String
        Get
            Dim config = GetConfiguration()
            Dim kvce As KeyValueConfigurationElement = config.AppSettings.Settings(Name)

            If kvce IsNot Nothing Then
                Return kvce.Value
            End If
            kvce = New KeyValueConfigurationElement(Name, "")
            config.AppSettings.Settings.Add(kvce)
            Try
                config.Save()
            Catch ex As Exception

            End Try
            Return ""
        End Get
        Set(ByVal value As String)
            Dim config = GetConfiguration()
            Dim kvce As KeyValueConfigurationElement = config.AppSettings.Settings(Name)
            If kvce Is Nothing Then
                kvce = New KeyValueConfigurationElement(Name, value)
                config.AppSettings.Settings.Add(kvce)
            Else
                kvce.Value = value
            End If
            Try
                config.Save()
            Catch ex As Exception

            End Try
        End Set
    End Property
    Public Shared Function GetConfigProperty(Name As String) As String
        Return ConfigProperty(Name)
    End Function
    Public Shared Sub SetConfigProperty(Name As String, Value As String)
        ConfigProperty(Name) = Value
    End Sub
    Public Shared Property SiteName() As String
        Get
            Dim s = ConfigProperty("SiteName")
            If String.IsNullOrWhiteSpace(s) Then
                s = "Unknown"
                ConfigProperty("SiteName") = s
            End If

            Return s
        End Get
        Set(ByVal value As String)
            ConfigProperty("SiteName") = value
        End Set
    End Property

    Public Shared ReadOnly Property ConnectionString() As String
        Get
            Return NamedConnString(ConnStringName)
        End Get
    End Property

End Class
