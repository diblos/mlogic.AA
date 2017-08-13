Imports BlackBerry.Workspaces
Imports BlackBerry.Workspaces.Json
Imports BlackBerry.Workspaces.Resource
Imports BlackBerry.Workspaces.Enums

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Xml
Imports System.Data.OleDb 'For Excel

Imports System.Threading
Imports System.ServiceProcess

Imports MassLogicService2.MassLogicConsole
Imports System.Numerics

Public Class Service1

    Public watchfolder() As FileSystemWatcher
    Public NO_OF_APP As Integer = 1

    Public DEV_MODE As Boolean = False
    Public APP_STATUS() As ApplicationObject

    'Dim Logger As New EventsLogger("Application", ".")
    Dim Logger As New LogWriter("Application Start")

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        Try
            mainLoader() 'Load Embedded Libraries

            InitAPP()
            StartWatching()
        Catch ex As Exception
            Logger.LogWrite("OnStart: " & ex.Message)
        End Try


    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        Try
            StopWatching()
        Catch ex As Exception
            Logger.LogWrite("OnStop: " & ex.Message)
        Finally
            Logger.LogWrite("Applications Stop")
        End Try

    End Sub

#Region "MassLogic Methods"

    Public Enum ModuleResult
        FAIL = 0
        OK = 1
    End Enum

    Private Const DateStr As String = "yyyy-MM-dd HH:mm"
    Private Const driveLetter As String = "C:\\"
    Private Const dirPathXML As String = "C:\Airbus\LPC-NG\report"
    Private Const extToSearch As String = "*.xml"
    Private Const filenameSplitChar As Char = "-"c
    Private Const numberOfSplits As Integer = 4
    Private Const WORKSPACE_ROOM_ID_ONE As Integer = 339569
    Private Const WORKSPACE_ROOM_ID_TWO As Integer = WORKSPACE_ROOM_ID_ONE '340450
    Private Const certFilename As String = "MassLogicCert.pfx"
    Private Const certPassword As String = "masslogicshukor"
    Private Const workspaceServerUrl As String = "shukor.watchdox.com"
    'Private Const userEmail As String = "msahmad82@gmail.com"
    Private Const userEmail As String = "rudi@masslogic.net"
    Private Const serviceAccountIssuerName As String = "com.watchdox.system.0367.3855"
    Private Const tokenExpiresInMinutes As Integer = 5

    Private Const ExcelFilename As String = "text_excel.xlsx"
    Private Const ExcelWorkspace As String = "Sheet1"

    Private Const LOGFOLDER_POSTFIX As String = "log"
    Private Const XMLFOLDER_POSTFIX As String = "xml"

    Private apiSession As ApiSession
    Private VolumeSerialNumber As String
    Private VolumeSerialNumberHex As String
    Private liReportFile As List(Of ReportFile)
    Private liGroups As List(Of String)
    Private liDomains As List(Of String)
    Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias _
    "GetVolumeInformationA" (PathName As String, VolumeNameBuffer As StringBuilder, VolumeNameSize As UInteger, ByRef VolumeSerialNumber As UInteger, ByRef MaximumComponentLength As UInteger, ByRef FileSystemFlags As UInteger, FileSystemNameBuffer As StringBuilder, FileSystemNameSize As UInteger) As Long

    Private Sub MainProcess(ByVal Filename As String)
        VolumeSerialNumber = Nothing
        VolumeSerialNumberHex = Nothing
        liReportFile = New List(Of ReportFile)()
        liGroups = New List(Of String)()
        liDomains = New List(Of String)()

        getVolumeSerialNumber()
        generateReportFiles(dirPathXML, System.IO.Path.GetFileNameWithoutExtension(Filename) & extToSearch.Replace("*", "")) 'FIND SPECIFIC FILE
        parseReportFile()
        Dim text As String = authenticateAndGetToken(apiSession)

        'MAP SERIAL WITH USERNAME START
        MapWithUsername()
        'MAP SERIAL WITH USERNAME END
        generateStringReportFile(True) 'Default without username entry in file content 

        lstMsgs(liReportFile.Count)

        If text IsNot Nothing AndAlso text.Length <> 0 Then
            For Each current As ReportFile In liReportFile
                uploadReportFile(apiSession, current, liGroups, liDomains)
            Next
        End If
    End Sub

    Private Sub getVolumeSerialNumber()
        Dim num As UInteger = 0UI
        Dim num2 As UInteger = 0UI
        Dim stringBuilder As StringBuilder = New StringBuilder(256)
        Dim num3 As UInteger = 0UI
        Dim stringBuilder2 As StringBuilder = New StringBuilder(256)
        If GetVolumeInformation("C:\\", stringBuilder, CUInt(stringBuilder.Capacity), num, num2, num3, stringBuilder2, CUInt(stringBuilder2.Capacity)) <> 0L Then
            VolumeSerialNumber = num.ToString()
            If VolumeSerialNumber IsNot Nothing Then
                Dim value As ULong = 0UL
                ULong.TryParse(VolumeSerialNumber, value)
                Dim bigInteger As BigInteger = New BigInteger(value)
                VolumeSerialNumberHex = bigInteger.ToString("X")
                VolumeSerialNumberHex = VolumeSerialNumberHex.TrimStart(New Char() {"0"c})
                Dim length As Integer = VolumeSerialNumberHex.Length
                If length >= 8 Then
                    Dim arg_D5_0 As String = VolumeSerialNumberHex.Substring(length - 8, 4)
                    Dim str As String = VolumeSerialNumberHex.Substring(length - 4, 4)
                    VolumeSerialNumberHex = arg_D5_0 + "-" + str
                End If
            End If
        End If
    End Sub

    Private Sub generateReportFiles(dirPath As String, extToSearch As String)

        lstMsgs(extToSearch)

        Dim files As String() = Directory.GetFiles(dirPath, extToSearch)
        For i As Integer = 0 To files.Length - 1
            Dim text As String = files(i)
            Dim fileNameWithoutExtension As String = Path.GetFileNameWithoutExtension(text)
            Dim array As String() = fileNameWithoutExtension.Split(New Char() {"-"c})

            If 4 = array.Count() Then
                Dim vSNfromFilename As String = array(0) + array(1)
                liReportFile.Add(New ReportFile(text, fileNameWithoutExtension, vSNfromFilename, array(3)))
            End If
        Next
    End Sub

    Private Sub parseReportFile()
        For Each current As ReportFile In liReportFile
            Using xmlReader As XmlReader = XmlReader.Create(New StringReader(File.ReadAllText(current.absolutePath, Encoding.UTF8)))
                Try
                    If xmlReader.ReadToFollowing("ConfigurationReport") Then
                        current.dateTimeFromContent = xmlReader.GetAttribute("DateTime")
                    End If
                Catch ex_53 As InvalidOperationException
                Catch ex_56 As ArgumentException
                End Try
                Try
                    If xmlReader.ReadToDescendant("Platform") Then
                        current.platformName = xmlReader.GetAttribute("Name")
                        current.platformType = xmlReader.GetAttribute("Type")
                    End If
                Catch ex_8B As InvalidOperationException
                Catch ex_8E As ArgumentException
                End Try
                Try
                    If xmlReader.ReadToFollowing("ConfigurationDescription") Then
                        current.OISVersion = xmlReader.GetAttribute("OISVersion")
                    End If
                Catch ex_B2 As InvalidOperationException
                Catch ex_B5 As ArgumentException
                End Try
            End Using
        Next
    End Sub

    Private Sub MapWithUsername()

        Try
            Try
                File.Delete(Path.Combine(LocalDir, ExcelFilename)) 'delete local copy excel file
            Catch ex As Exception
                Logger.LogWrite("MapWithUsername: " & ex.Message)
            End Try

            If DownloadFileByName(WORKSPACE_ROOM_ID_ONE, "/", ExcelFilename, Path.Combine(LocalDir, ExcelFilename), Now) = ModuleResult.OK Then

                Dim dt As DataTable
                If Path.GetExtension(ExcelFilename) = "xls" Then
                    'ReadExcel(ExcelFilename, ExcelWorkspace)
                    dt = ReadExcelToTable(Path.Combine(LocalDir, ExcelFilename))
                Else
                    dt = EPPlusClass.ReadExcelToTable(Path.Combine(LocalDir, ExcelFilename))
                End If

                If dt.Rows.Count > 0 Then
                    For Each current As ReportFile In liReportFile
                        Dim foundRows As DataRow()
                        foundRows = dt.Select("F1='" & current.platformName & "'")
                        current.Username = foundRows(0).Item("F2")
                    Next
                End If
            End If

        Catch ex As Exception
            Logger.LogWrite("MapWithUsername: " & ex.Message)
        End Try
    End Sub

    Private Function ReadExcelToTable(excelpath As String) As DataTable

        'CONNECTION STRING
        'Dim connstring As String = (Convert.ToString("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=") & excelpath) + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"
        'THE SAME NAME 
        Dim connstring As String = (Convert.ToString("Provider=Microsoft.JET.OLEDB.4.0;Data Source=") & excelpath) + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"

        Using conn As New OleDbConnection(connstring)
            conn.Open()
            'Get All Sheets Name
            Dim sheetsName As DataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "Table"})

            'Get the First Sheet Name
            Dim firstSheetName As String = sheetsName.Rows(0)(2).ToString()

            'Query String 
            Dim sql As String = String.Format("SELECT * FROM [{0}]", firstSheetName)
            Dim ada As New OleDbDataAdapter(sql, connstring)
            Dim [set] As New DataSet()
            ada.Fill([set])
            Return [set].Tables(0)
        End Using
    End Function

    Private Sub generateStringReportFile(Optional ByVal withUsername As Boolean = False)
        For Each current As ReportFile In liReportFile
            Dim stringBuilder As StringBuilder = New StringBuilder()
            stringBuilder.Append("DateTime:")
            stringBuilder.Append(current.dateTimeFromContent)
            stringBuilder.Append(Environment.NewLine)
            stringBuilder.Append("OISVersion:")
            stringBuilder.Append(current.OISVersion)
            stringBuilder.Append(Environment.NewLine)
            stringBuilder.Append("PlatformName:")
            stringBuilder.Append(current.platformName)
            stringBuilder.Append(Environment.NewLine)
            stringBuilder.Append("PlatformType:")
            stringBuilder.Append(current.platformType)
            stringBuilder.Append(Environment.NewLine)
            If withUsername Then
                stringBuilder.Append("Username:")
                stringBuilder.Append(current.Username)
                stringBuilder.Append(Environment.NewLine)
            End If
            current.WatchdoxFileContent = stringBuilder.ToString()
        Next
    End Sub

    Private Function DownloadFileByName(ByVal roomId As Integer, ByVal folderPath As String, ByVal docName As String, ByVal destinationPath As String, ByVal lastUpdateTime As Date) As ModuleResult
        Try
            ' Get an instance of DownloadManager    
            Dim downloadManager As DownloadManager = apiSession.GetDownloadManager()
            ' A call to the DownloadFileByName            
            downloadManager.DownloadFileByName(roomId, folderPath, docName, destinationPath, lastUpdateTime)
            Return ModuleResult.OK
        Catch ex As Exception
            Logger.LogWrite("DownloadFileByName: " & ex.Message)
            Return ModuleResult.FAIL
        End Try
    End Function

    Private Sub uploadReportFile(ByRef apiSession As ApiSession, reportFile As ReportFile, liGroups As List(Of String), liDomains As List(Of String))
        Dim text As String = generateRandomAlphaString(10)

        Using expr_0Fx As StreamWriter = New StreamWriter(text, True)
            expr_0Fx.Write(reportFile.WatchdoxFileContent)
            expr_0Fx.Close()
        End Using


        Dim UFC As UploadFilesClass = New UploadFilesClass(apiSession)

        Dim r As UploadResult = UFC.UploadDocumentToRoom(WORKSPACE_ROOM_ID_TWO, reportFile.getDstFilename, text, reportFile.getDstFolder & "_" & LOGFOLDER_POSTFIX, liGroups, liDomains)
        Logger.LogWrite("Upload log: " & r.Status.ToString)
        Dim s As UploadResult = UFC.UploadDocumentToRoom(WORKSPACE_ROOM_ID_TWO, Path.ChangeExtension(reportFile.getDstFilename, "xml"), reportFile.absolutePath, reportFile.getDstFolder & "_" & XMLFOLDER_POSTFIX, liGroups, liDomains)
        Logger.LogWrite("Upload xml: " & s.Status.ToString)

        Try
            File.Delete(text)
        Catch ex_48 As ArgumentException
        Catch ex_4B As DirectoryNotFoundException
        Catch ex_4E As IOException
        Catch ex_51 As NotSupportedException
        Catch ex_54 As UnauthorizedAccessException
        End Try
    End Sub

    Private Function generateRandomAlphaString(length As Integer) As String
        Dim stringBuilder As StringBuilder = New StringBuilder()
        Dim random As Random = New Random()
        While True
            Dim arg_36_0 As Integer = 0
            Dim expr_31 As Integer = length
            length = expr_31 - 1
            If arg_36_0 >= expr_31 Then
                Exit While
            End If
            stringBuilder.Append("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"(random.[Next]("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ".Length)))
        End While
        Return stringBuilder.ToString()
    End Function

    Private Function authenticateAndGetToken(ByRef apiSession As ApiSession) As String
        Dim cert As X509Certificate2 = Nothing
        Try
            'Logger.LogWrite("authenticateAndGetToken: " & certFilename & "|" & certPassword)
            'cert = New X509Certificate2(certFilename, certPassword, X509KeyStorageFlags.Exportable)
            cert = New X509Certificate2(Path.Combine(LocalDir, certFilename), certPassword, X509KeyStorageFlags.Exportable)
        Catch ex_15 As CryptographicException
            Logger.LogWrite("authenticateAndGetToken: " & ex_15.Message)
        End Try
        apiSession = New ApiSession(workspaceServerUrl, Nothing)
        apiSession.GetWorkspacesResource()
        Dim arg_47_0 As Integer = CInt(apiSession.StartSessionWithServiceAccount(userEmail, serviceAccountIssuerName, tokenExpiresInMinutes, cert))
        Dim arg_45_0 As String = String.Empty
        If arg_47_0 = 1 Then
            Return apiSession.GetToken()
        End If
        Return Nothing
    End Function

#End Region

#Region "Watcher Methods"

    Private Function LocalDir() As String
        Dim P As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)
        P = New Uri(P).LocalPath
        Return P
    End Function

    Public Sub InitAPP()

        ReDim APP_STATUS(NO_OF_APP - 1)

        For x = 0 To NO_OF_APP - 1

            APP_STATUS(x) = New ApplicationObject

            'SET APPLICATION OBJECT VALUES
            '=============================
            APP_STATUS(x).EXEPath = LocalDir()
            APP_STATUS(x).Time = Now
            APP_STATUS(x).WatchFolder = dirPathXML
            APP_STATUS(x).ArchivePath = dirPathXML
            APP_STATUS(x).Status = AppStatus.Active

        Next
    End Sub

    Private Sub StartWatching()
        ReDim watchfolder(NO_OF_APP - 1)

        For x = 0 To UBound(watchfolder)
            watchfolder(x) = New System.IO.FileSystemWatcher()

            'this is the path we want to monitor
            watchfolder(x).Path = APP_STATUS(x).WatchFolder

            'Add a list of Filter we want to specify

            'make sure you use OR for each Filter as we need to

            'all of those
            'watchfolder(x).NotifyFilter = IO.NotifyFilters.DirectoryName
            'watchfolder(x).NotifyFilter = watchfolder(x).NotifyFilter Or _
            '                              IO.NotifyFilters.FileName
            'watchfolder(x).NotifyFilter = watchfolder(x).NotifyFilter Or _
            '                              IO.NotifyFilters.Attributes
            '-------------------------------------------------------------
            'watchfolder(x).NotifyFilter = (NotifyFilters.LastAccess Or _
            '             NotifyFilters.LastWrite Or _
            '             NotifyFilters.FileName Or _
            '             NotifyFilters.DirectoryName)
            watchfolder(x).NotifyFilter = IO.NotifyFilters.DirectoryName
            'watchfolder(x).NotifyFilter = watchfolder(x).NotifyFilter Or _
            '                              IO.NotifyFilters.LastWrite
            watchfolder(x).NotifyFilter = (NotifyFilters.LastAccess Or
             NotifyFilters.LastWrite Or
             NotifyFilters.FileName Or
             NotifyFilters.DirectoryName)
            '-------------------------------------------------------------
            ' add the handler to each event

            'AddHandler watchfolder(x).Changed, AddressOf logchange
            AddHandler watchfolder(x).Created, AddressOf logchange
            'AddHandler watchfolder(x).Deleted, AddressOf logchange

            ' add the rename handler as the signature is different
            'AddHandler watchfolder(x).Renamed, AddressOf logrename

            'Set this property to true to start watching

            watchfolder(x).EnableRaisingEvents = True
        Next

    End Sub

    Private Sub StopWatching()
        Try
            For x = 0 To UBound(watchfolder)
                watchfolder(x).EnableRaisingEvents = False
            Next
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub logchange(ByVal source As Object, ByVal e As _
                    System.IO.FileSystemEventArgs)
        If e.ChangeType = IO.WatcherChangeTypes.Changed Then
            If DEV_MODE = True Then
                lstMsgs("logchange: File " & Path.GetFileName(e.FullPath) &
                        " has been arrived")
            End If
        End If

        If (isExtensionRight(e.FullPath, ".xml")) Then MainProcess(e.FullPath)

    End Sub

    Public Function isExtensionRight(ByVal fPath As String, ByVal ext As String) As Boolean

        Dim extension As String = Path.GetExtension(fPath)
        If extension.ToUpper = ext.Trim().ToUpper Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub lstMsgs(ByVal str As String)
        Try
            Logger.LogWrite(str)
        Catch ex As Exception

        End Try
    End Sub

#End Region

End Class

Public Enum AppStatus
    Active = 1
    InActive = 2
End Enum

Public Class ApplicationObject
    Private _Status As AppStatus
    Private _Time As DateTime
    Private _WatchFolder As String
    Private _Path As String
    Private _ArchivePath As String

    Public Property Status() As AppStatus
        Get
            Return Me._Status
        End Get
        Set(ByVal value As AppStatus)
            Me._Status = value
        End Set
    End Property

    Public Property Time() As DateTime
        Get
            Return Me._Time
        End Get
        Set(ByVal value As DateTime)
            Me._Time = value
        End Set
    End Property

    Public Property WatchFolder() As String
        Get
            Return Me._WatchFolder
        End Get
        Set(ByVal value As String)
            Me._WatchFolder = value
        End Set
    End Property

    Public Property ArchivePath() As String
        Get
            Return Me._ArchivePath
        End Get
        Set(ByVal value As String)
            Me._ArchivePath = value
        End Set
    End Property

    Public Property EXEPath() As String
        Get
            Return Me._Path
        End Get
        Set(ByVal value As String)
            Me._Path = value
        End Set
    End Property

End Class