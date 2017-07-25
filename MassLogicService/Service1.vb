Imports BlackBerry.Workspaces
Imports BlackBerry.Workspaces.Json
Imports BlackBerry.Workspaces.Resource
Imports BlackBerry.Workspaces.Enums

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Numerics
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Xml
Imports System.Data.OleDb 'For Excel

Imports System.Threading
Imports System.ServiceProcess

Imports MassLogicService.MassLogicConsole

Public Class Service1

    Public watchfolder() As FileSystemWatcher
    Public NO_OF_APP As Integer = 1

    Public DEV_MODE As Boolean = False
    Public APP_STATUS() As ApplicationObject

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
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

    Private Const ExcelFilename As String = "text_excel.xls" ' xls is OK, xlsx need to be checked
    Private Const ExcelWorkspace As String = "Sheet1"


    Private apiSession As ApiSession
    Private VolumeSerialNumber As String
    Private VolumeSerialNumberHex As String
    Private liReportFile As List(Of ReportFile)
    Private liGroups As List(Of String)
    Private liDomains As List(Of String)
    Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias _
"GetVolumeInformationA" (PathName As String, VolumeNameBuffer As StringBuilder, VolumeNameSize As UInteger, ByRef VolumeSerialNumber As UInteger, ByRef MaximumComponentLength As UInteger, ByRef FileSystemFlags As UInteger, FileSystemNameBuffer As StringBuilder, FileSystemNameSize As UInteger) As Long

#End Region

#Region "Watcher Methods"

    Public Sub InitAPP()
        Dim tmpStr As String
        ReDim APP_STATUS(NO_OF_APP - 1)

        For x = 0 To NO_OF_APP - 1

            'tmpStr = AppSettings("APP." & x + 1)

            APP_STATUS(x) = New ApplicationObject

            'SET APPLICATION OBJECT VALUES
            '=============================
            'APP_STATUS(x).EXEPath = AppSettings("RESPONSE." & x + 1)
            APP_STATUS(x).Time = Now
            APP_STATUS(x).WatchFolder = tmpStr
            'tmpStr = AppSettings("ARCHIVE." & x + 1)
            APP_STATUS(x).ArchivePath = tmpStr
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
            watchfolder(x).NotifyFilter = (NotifyFilters.LastAccess Or _
             NotifyFilters.LastWrite Or _
             NotifyFilters.FileName Or _
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

    Private Sub logchange(ByVal source As Object, ByVal e As  _
                    System.IO.FileSystemEventArgs)
        If e.ChangeType = IO.WatcherChangeTypes.Changed Then
            If DEV_MODE = True Then
                lstMsgs(Now.ToString("yyyy-MM-dd HH:mm:ss") & " File " & Path.GetFileName(e.FullPath) & _
                        " has been arrived")
            End If
        End If

        If (isExtensionRight(e.FullPath, ".txt")) Then ParseFileName(e.FullPath)

    End Sub

    Public Function isExtensionRight(ByVal fPath As String, ByVal ext As String) As Boolean

        Dim extension As String = Path.GetExtension(fPath)
        If extension.ToUpper = ext.Trim().ToUpper Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub ParseFileName(ByVal FILENAME As String)
        Dim tmpName = System.IO.Path.GetFileNameWithoutExtension(FILENAME)
        Dim ServiceCode As String = tmpName.Split("_")(1)

        Select Case ServiceCode
            Case "RQCA"
                'ReadCAFile(FILENAME)

            Case "K1BTC"
                'ReadK1File(FILENAME)
            Case Else
                lstMsgs("unrecognised file: " & FILENAME)
        End Select

    End Sub

    Private Sub lstMsgs(ByVal str As String)
        Try

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