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
        Private Const dirPathXML As String = "C:\Airbus\LPC-NG\report"
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

        Private apiSession As ApiSession
        Private VolumeSerialNumber As String
        Private VolumeSerialNumberHex As String
        Private liReportFile As List(Of ReportFile)
        Private liGroups As List(Of String)
        Private liDomains As List(Of String)
        'Private Declare Function GetVolumeInformation Lib "kernel32.dll" (PathName As String, VolumeNameBuffer As StringBuilder, VolumeNameSize As UInteger, ByRef VolumeSerialNumber As UInteger, ByRef MaximumComponentLength As UInteger, ByRef FileSystemFlags As UInteger, FileSystemNameBuffer As StringBuilder, FileSystemNameSize As UInteger) As Long
        Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias _
    "GetVolumeInformationA" (PathName As String, VolumeNameBuffer As StringBuilder, VolumeNameSize As UInteger, ByRef VolumeSerialNumber As UInteger, ByRef MaximumComponentLength As UInteger, ByRef FileSystemFlags As UInteger, FileSystemNameBuffer As StringBuilder, FileSystemNameSize As UInteger) As Long
        'http://www.jasinskionline.com/windowsapi/ref/g/getvolumeinformation.html

        Sub Main()
            Console.WriteLine(Now.ToString(DateStr))
            theMain(Nothing)
            HappyEnd() 'Wait input to end
        End Sub

        Private Sub theMain(args As String())
            VolumeSerialNumber = Nothing
            VolumeSerialNumberHex = Nothing
            liReportFile = New List(Of ReportFile)()
            liGroups = New List(Of String)()
            liDomains = New List(Of String)()

            getVolumeSerialNumber()
            generateReportFiles(dirPathXML, extToSearch)
            parseReportFile()
            generateStringReportFile()

            Dim text As String = authenticateAndGetToken(apiSession)
            If text IsNot Nothing AndAlso text.Length <> 0 Then
                For Each current As ReportFile In liReportFile
                    uploadReportFile(apiSession, current, liGroups, liDomains)
                Next
            End If
        End Sub

        Sub HappyEnd()
            Dim r = Console.ReadLine()
            Console.WriteLine(r)
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
                Using xmlReader As XmlReader = xmlReader.Create(New StringReader(File.ReadAllText(current.absolutePath, Encoding.UTF8)))
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

        Private Sub generateStringReportFile()
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
                current.WatchdoxFileContent = stringBuilder.ToString()
            Next
        End Sub

        Private Function authenticateAndGetToken(ByRef apiSession As ApiSession) As String
            Dim cert As X509Certificate2 = Nothing
            Try
                cert = New X509Certificate2(certFilename, certPassword, X509KeyStorageFlags.Exportable)
            Catch ex_15 As CryptographicException
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

        Private Sub uploadReportFile(ByRef apiSession As ApiSession, reportFile As ReportFile, liGroups As List(Of String), liDomains As List(Of String))
            Dim text As String = generateRandomAlphaString(10)
            Dim expr_0F As StreamWriter = New StreamWriter(text, True)
            expr_0F.Write(reportFile.WatchdoxFileContent)
            expr_0F.Close()
            Dim UFC As UploadFilesClass = New UploadFilesClass(apiSession)
            UFC.UploadDocumentToRoom(workspaceRoomId, reportFile.getDstFilename(), text, VolumeSerialNumberHex, liGroups, liDomains)

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

    End Module
End Namespace