Imports EaTools
Imports EaMySql
Public Class DcSettings

#Region " Notes "

    ' settings file is DataConfig.xml saved in the application folder, not the application version folder
    ' C:\Users\username\AppData\Roaming\Company\Product\

#End Region

#Region " Class Constants, ENums and Variables "

#Region " Constants "

    Private Const AppNameText As String = "AppName"

    Private Const fxDataFile As String = "mdb"
    Private Const fxXml As String = "xml"

    Private Const snDataConnection As String = "DataConnection"
    Private Const snSettings As String = "Settings"

    Private Const vkDatabase As String = "Database"
    Private Const vkEncryptedPassword As String = "EncryptedPassword"
    Private Const vkPersisSecInfo As String = "PersisSecInfo"
    Private Const vkServer As String = "Server"
    Private Const vkSslMode As String = "SslMode"
    Private Const vkUserName As String = "UserName"

    Private Const DefaultConfigName As String = "DataConfig." & fxXml
    Private Const dfPassword As String = "password"
    Private Const dfPerSecInfo As Boolean = True
    Private Const dfServer As String = "localhost"
    Private Const dfUserName As String = "user"

    Public Const NoErrors As Integer = 0
    Public Const ioFileNotFoundErr As Integer = -4031
    Public Const ioFolderNotFoundErr As Integer = -4032
    Public Const OtherExErr As Integer = -5000

    Public Const AppRootFolder As String = "C:\Projects 2017\PcdbToSql\PcdbToSql\bin\Debug"

#End Region

#Region " Enums "

#End Region

#Region " Private Vars "

    Private _encryptedPassword As String = String.Empty

    Private wrapper As Simple3Des

#End Region

#End Region

#Region " New "

    Public Sub New()

        ' vars passed:
        '   CurrentSettings - values for the settings class

        Try
            wrapper = New Simple3Des(vkEncryptedPassword)   ' create encryption wrapper
            SetDefaults()                                   ' set the default values for all 
            ReadUserSettings()                              ' reads the user settings
        Catch ex As Exception
            Return
        End Try
    End Sub

#End Region

#Region " Properties & Methods "

#Region " Properties "

    Shared ReadOnly Property ConfigFileName As String
        Get
            'Return String.Format("{0}\{1}", MiniEsaIo.MyLibrary.CurrentUserAppDataRoot, DefaultConfigName)
            Return String.Format("{0}\{1}", AppRootFolder, DefaultConfigName)
        End Get
    End Property

    Property Database As String = String.Empty

    Property Password As String
        Get
            Return wrapper.DecryptData(_encryptedPassword)
        End Get
        Set(value As String)
            _encryptedPassword = wrapper.EncryptData(value)
        End Set
    End Property

    Property PerSecInfo As Boolean = True
    Property Server As String = dfServer
    Property sslMode As MySql.Data.MySqlClient.MySqlSslMode = MySql.Data.MySqlClient.MySqlSslMode.None
    Property UserName As String = dfUserName

#End Region

#Region " Methods "

    Public Function ReadUserSettings() As Integer

        ' reads the user settings from the PCDB config xml file
        '
        ' returns:
        '   NoErrors - no errors
        '   ioFileNotFoundErr - could not find file
        '   OtherExErr - other error

        Dim reader As Xml.XmlReader = Nothing
        Try
            Dim CfgFileName As String = ConfigFileName                              ' get the file name
            If Not My.Computer.FileSystem.FileExists(CfgFileName) Then              ' if file does not exits
                Return ioFileNotFoundErr                                            ' return error
            End If

            Dim rs As Xml.XmlReaderSettings = New Xml.XmlReaderSettings() With {.IgnoreWhitespace = True} ' ignore whitespace
            reader = Xml.XmlTextReader.Create(CfgFileName, rs)                      ' create xml reader
            Dim ElementName As String = String.Empty                                ' init element name to blank
            Dim ReaderData As String
            While reader.Read()                                                     ' read an item from the xml file
                ReaderData = String.Format("reader.Read(): Name: {0}, Value: {1}", reader.Name, reader.Value)
                Select Case reader.NodeType
                    Case Xml.XmlNodeType.Element                                    ' if an element header
                        ElementName = reader.Name                                   ' get the element name
                    Case Xml.XmlNodeType.Text                                       ' if element value
                        SetUserSettingValueFromXml(ElementName, reader.Value)       ' set the value in user settings
                    Case Xml.XmlNodeType.EndElement                                 ' if an end element
                        ElementName = String.Empty                                  ' clear element name
                End Select
            End While
            Return NoErrors
        Catch ex As Exception
            Return OtherExErr
        Finally
            If reader IsNot Nothing Then                                            ' if got a reader
                reader.Close()                                                      ' always close it
            End If
        End Try
    End Function

    Private Sub SetDefaults()

        ' sets the default settings values

        Try
            Server = dfServer
            UserName = dfUserName
            Password = dfPassword
            PerSecInfo = dfPerSecInfo
            sslMode = MySql.Data.MySqlClient.MySqlSslMode.Preferred
        Catch ex As Exception
            Return
        End Try
    End Sub

    Private Sub SetUserSettingValueFromXml(Name As String,
                                           ValueText As String)

        ' sets the user setting value 
        '
        ' vars passed:
        '   Name - element name read from xml file
        '   ValueText - value read from xml file

        Try
            Select Case Name

                ' Settings

                Case vkDatabase
                    Database = ValueText
                Case vkEncryptedPassword
                    _encryptedPassword = ValueText
                Case vkServer
                    Server = ValueText
                Case vkSslMode
                    sslMode = MySqlTools.SslModeFromString(ValueText)
                Case vkPersisSecInfo
                    PerSecInfo = MySqlTools.PrecisionSecurityInfoFromString(ValueText)
                Case vkUserName
                    UserName = ValueText
            End Select
        Catch ex As Exception
            Return
        End Try
    End Sub

    Public Function WriteUserSettings() As Integer

        ' writes settings to user settings file
        '
        ' returns:
        '   NoErrors - no errors
        '   ioFolderNotFoundErr - could not find folder
        '   OtherExErr - other error

        Dim writer As Xml.XmlTextWriter = Nothing
        Try
            Dim CfgFileName As String = ConfigFileName                                          ' get the file name
            Dim FInfo As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(CfgFileName)   ' get the file info for config file

            If Not My.Computer.FileSystem.DirectoryExists(FInfo.DirectoryName) Then
                Return ioFolderNotFoundErr
            End If

            writer = New Xml.XmlTextWriter(CfgFileName, Nothing)                                ' create the xml writer 
            writer.Formatting = Xml.Formatting.Indented                                         ' use indenting
            writer.WriteStartElement(snDataConnection)

            ' settings
            Try
                writer.WriteStartElement(snSettings)
                writer.WriteElementString(vkServer, Server)
                writer.WriteElementString(vkUserName, UserName)
                writer.WriteElementString(vkEncryptedPassword, _encryptedPassword)
                writer.WriteElementString(vkDatabase, Database)
                writer.WriteElementString(vkPersisSecInfo, PerSecInfo.ToString.ToUpper)
                writer.WriteElementString(vkSslMode, sslMode.ToString)
            Finally
                writer.WriteEndElement()
            End Try

            Return NoErrors
        Catch ex As Exception
            Return OtherExErr
        Finally
            writer.WriteEndElement()
            If writer IsNot Nothing Then
                writer.Close()
            End If
        End Try
    End Function

#End Region

#End Region

End Class

