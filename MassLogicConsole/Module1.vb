Imports BlackBerry.Workspaces
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


Namespace MassLogicConsole
    Module Module1

        Private Const DateStr As String = "yyyy-MM-dd HH:mm"

        Private Const driveLetter As String = "C:\\"

        Private Const dirPath As String = "C:\Airbus\LPC-NG\report"

        Private Const extToSearch As String = "*.xml"

        Private Const filenameSplitChar As Char = "-"c

        Private Const numberOfSplits As Integer = 4

        Private Const workspaceRoomId As Integer = 339569

        Private Const certFilename As String = "MassLogicCert.pfx"

        Private Const certPassword As String = "masslogicshukor"

        Private Const workspaceServerUrl As String = "shukor.watchdox.com"

        Private Const userEmail As String = "msahmad82@gmail.com"

        Private Const serviceAccountIssuerName As String = "com.watchdox.system.0367.3855"

        Private Const tokenExpiresInMinutes As Integer = 5

        'Shared?
        'Private apiSession As ApiSession
        'Private VolumeSerialNumber As String
        'Private VolumeSerialNumberHex As String
        'Private liReportFile As List(Of ReportFile)
        'Private liGroups As List(Of String)
        'Private liDomains As List(Of String)
        'Private Declare Function GetVolumeInformation Lib "kernel32.dll" (PathName As String, VolumeNameBuffer As StringBuilder, VolumeNameSize As UInteger, ByRef VolumeSerialNumber As UInteger, ByRef MaximumComponentLength As UInteger, ByRef FileSystemFlags As UInteger, FileSystemNameBuffer As StringBuilder, FileSystemNameSize As UInteger) As Long
        'Shared?


        Sub Main()
            Console.WriteLine(Now.ToString(DateStr))
            HappyEnd() 'Wait input to end
        End Sub

        Sub HappyEnd()
            Dim r = Console.ReadLine()
            Console.WriteLine(r)
        End Sub

    End Module
End Namespace