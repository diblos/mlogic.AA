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

Namespace MassLogicConsole
    Module Module1

        Private Const DateStr As String = "yyyy-MM-dd HH:mm"
        Private Const driveLetter As String = "C:\\"
        Private Const dirPathXML As String = "C:\Airbus\LPC-NG\report"
        Private Const extToSearch As String = "*.xml"
        Private Const filenameSplitChar As Char = "-"c
        Private Const numberOfSplits As Integer = 4
        Private Const WORKSPACE_ROOM_ID_ONE As Integer = 339569
        Private Const WORKSPACE_ROOM_ID_TWO As Integer = 340450
        Private Const certFilename As String = "MassLogicCert.pfx"
        Private Const certPassword As String = "masslogicshukor"
        Private Const workspaceServerUrl As String = "shukor.watchdox.com"
        'Private Const userEmail As String = "msahmad82@gmail.com"
        Private Const userEmail As String = "rudi@masslogic.net"
        Private Const serviceAccountIssuerName As String = "com.watchdox.system.0367.3855"
        Private Const tokenExpiresInMinutes As Integer = 5

        Private Const ExcelFilename As String = "text_excel.xlsx"
        Private Const ExcelWorkspace As String = "Sheet1"


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

            'ReadExcel(ExcelFilename, ExcelWorkspace)
            'Dim dt As DataTable = ReadExcelToTable(ExcelFilename)

            Dim text As String = authenticateAndGetToken(apiSession)

            'GetWorkspace()
            'GetFolder()
            'GetFile(workspaceRoomId)

            'DownloadFileByName(WORKSPACE_ROOM_ID_ONE, "/", ExcelFilename, Path.Combine("C:\Users\lenovo\Desktop\SAT", ExcelFilename), Now)
            'UploadFile(WORKSPACE_ROOM_ID_TWO, ExcelFilename, ExcelFilename, "/test1/test2/test2", Nothing, Nothing)

            If text IsNot Nothing AndAlso text.Length <> 0 Then
                For Each current As ReportFile In liReportFile
                    uploadReportFile(apiSession, current, liGroups, liDomains)
                Next
            End If
        End Sub

        Sub HappyEnd()
            Console.WriteLine("...")
            Dim r = Console.ReadLine()
            Console.WriteLine(r)
        End Sub

        Private Sub ReadExcel(ByVal fn As String, ByVal ws As String)

            'Dim fileName = String.Format("{0}\fileNameHere", Directory.GetCurrentDirectory())
            Dim fileName = String.Format("{0}\" & fn, Directory.GetCurrentDirectory())
            Dim connectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName)

            'Dim adapter = New OleDbDataAdapter("SELECT * FROM [workSheetNameHere$]", connectionString)
            Dim adapter = New OleDbDataAdapter("SELECT * FROM [" & ws & "$]", connectionString)
            Dim ds = New DataSet()

            adapter.Fill(ds, "anyNameHere")

            Dim data As DataTable = ds.Tables("anyNameHere")
        End Sub

        Private Function ReadExcelToTable(path As String) As DataTable

            'CONNECTION STRING
            Dim connstring As String = (Convert.ToString("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=") & path) + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"
            'THE SAME NAME 
            'Dim connstring As String = (Convert.ToString("Provider=Microsoft.JET.OLEDB.4.0;Data Source=") & path) + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"

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

        Private Sub CreateFolder()
            Dim workspaces As Resource.Workspaces = apiSession.GetWorkspacesResource()
            Dim x As New CreateWorkspaceFolderTreeJson
            With x
                .DeviceType = DeviceType.BROWSER
                .ExternalRepository = ExternalRepositoryType.NONE
                .ObjType = JsonObjectTypes.FOLDER
                .RoomId = WORKSPACE_ROOM_ID_ONE

            End With
            workspaces.CreateFoldersTreeV30(x)
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

            'Dim r As UploadResult = UFC.UploadDocumentToRoom(WORKSPACE_ROOM_ID_TWO, reportFile.getDstFilename(), text, VolumeSerialNumberHex, liGroups, liDomains)
            'Dim r As UploadResult = UFC.UploadDocumentToRoom(WORKSPACE_ROOM_ID_TWO, reportFile.getDstFilename, text, reportFile.getDstFolder, liGroups, liDomains)
            Dim r As UploadResult = UFC.UploadFile(WORKSPACE_ROOM_ID_TWO, text, reportFile.getDstFilename, reportFile.getDstFolder, liGroups, liDomains)
            Console.WriteLine(r.Status.ToString)

            Try
                File.Delete(text)
            Catch ex_48 As ArgumentException
            Catch ex_4B As DirectoryNotFoundException
            Catch ex_4E As IOException
            Catch ex_51 As NotSupportedException
            Catch ex_54 As UnauthorizedAccessException
            End Try
        End Sub

        Private Sub UploadFile(ByVal roomid As Integer, ByVal filename As String, ByVal destinationFileName As String, ByVal folder As String, ByVal groups As List(Of String), ByVal domains As List(Of String))
            ' Get an instance of UploadManager            
            Dim uploadManager As UploadManager = apiSession.GetUploadManager()
            ' Create a new SubmitDocumentsVdrJson JSON            
            Dim uploadInfo As SubmitDocumentsVdrJson = New SubmitDocumentsVdrJson
            With uploadInfo
                .OpenForAllRoom = False
                .Recipients = New RoomRecipientsJson()
                With .Recipients
                    .Groups = groups
                    .Domains = domains
                End With
                .Folder = folder
                .TagValueList = Nothing
                .DeviceType = DeviceType.SYNC
            End With

            ' A call to the UploadDocumentToRoom            
            Dim uploadResult As UploadResult = uploadManager.UploadDocumentToRoom(uploadInfo, roomid, destinationFileName, filename, Nothing)

            Console.WriteLine(uploadResult.Status.ToString())

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

        '//////////////////////////////////////////////////////////////////////

        Private Sub GetWorkspace()
            Dim workspaces As Resource.Workspaces = apiSession.GetWorkspacesResource()
            ' This returns a list of rooms, which can be iterated over. The other parameters            
            ' include: addExternalData, adminMode, includeSyncData, includeWorkspacePolicyData,            
            ' and workspaceTypes. Please see the javadoc documentation for details.
            Dim itemListJson As ItemListJson(Of WorkspaceInfoJson) = workspaces.ListRoomsV30(Nothing, True, True, False, False)
        End Sub

        Private Sub GetFolder()
            'Resource.Workspaces workspaces = apiSession.GetWorkspacesResource();
            Dim workspaces As Resource.Workspaces = apiSession.GetWorkspacesResource()
            ' This returns a folder object, which contains details about the current workspace,            
            ' as well as a sub folder list that can be iterated over.            
            'FolderJson folderJson = workspaces.GetFolderTreeV30(roomId);
            Dim folderJson As FolderJson = workspaces.GetFolderTreeV30(WORKSPACE_ROOM_ID_ONE)
            'List<FolderJson> subFolders = folderJson.SubFolders
            Dim subFolders As List(Of FolderJson) = folderJson.SubFolders
        End Sub

        Private Sub GetActivity(ByVal documentGuid As String)


            Dim files As Files = apiSession.GetFilesResource()
            ' Create an object to specify the documents activityLog request            
            ' The guid of a document to retrieve activity for                
            ' Indicates if only the last action for a user should be retrieved                
            ' Indicates the page number to fetch of a multipage response                
            ' The number of items to fetch per page    

            'Dim getDocumentActivityLogRequestJson As New GetDocumentActivityLogRequestJson() With { _
            '	Key .DocumentGuid = documentGuid, _
            '	Key .LastActionPerUser = False, _
            '	Key .PageNumber = 1, _
            '	Key .PageSize = 100 _
            '}

            Dim getDocumentActivityLogRequestJson As New GetDocumentActivityLogRequestJson()
            With getDocumentActivityLogRequestJson
                .DocumentGuid = documentGuid
                .LastActionPerUser = False
                .PageNumber = 1
                .PageSize = 100
            End With

            ' Call the get activity method            
            Dim result As PagingItemListJson(Of ActivityLogRecordJson) = files.GetActivityLogV30(getDocumentActivityLogRequestJson)

        End Sub

        Private Sub Add2Group(ByVal userAddresses As List(Of String), ByVal groupName As String)
            Dim workspaces As Resource.Workspaces = apiSession.GetWorkspacesResource()
            Dim memberList As New List(Of AddMemberToGroupJson)()

            For Each currentAddress As String In userAddresses
                Dim currentEntity As New PermittedEntityFromUserJson()
                With currentEntity
                    .Address = currentAddress
                    .EntityType = EntityType.USER
                End With
                'make a AddMemberToGroupJson for each user                
                Dim currentMemberJson As AddMemberToGroupJson = New AddMemberToGroupJson
                currentMemberJson.Entity = currentEntity
                memberList.Add(currentMemberJson)
            Next

            Dim groupMemberJson As AddMembersToGroupWithGroupJson = New AddMembersToGroupWithGroupJson
            With groupMemberJson
                .MembersList = memberList
                .RoomId = WORKSPACE_ROOM_ID_ONE
                .GroupName = groupName
            End With

            Dim result As String = workspaces.AddMembersToGroupV30(groupMemberJson)

        End Sub

        Private Sub DownloadFileById(ByVal docId As String, ByVal roomId As Integer, ByVal destinationPath As String, ByVal lastUpdateTime As Date)
            Try
                ' Get an instance of DownloadManager            
                Dim downloadManager As DownloadManager = apiSession.GetDownloadManager()
                ' A call to the DownloadFileById            
                downloadManager.DownloadFileById(docId, String.Empty, roomId, destinationPath, lastUpdateTime, True, True)
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Sub

        Private Sub DownloadFileByName(ByVal roomId As Integer, ByVal folderPath As String, ByVal docName As String, ByVal destinationPath As String, ByVal lastUpdateTime As Date)
            Try
                ' Get an instance of DownloadManager    
                Dim downloadManager As DownloadManager = apiSession.GetDownloadManager()
                ' A call to the DownloadFileByName            
                downloadManager.DownloadFileByName(roomId, folderPath, docName, destinationPath, lastUpdateTime)
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Sub

        Private Sub DownloadFileToBuffer(ByVal docId As String)
            Try
                ' Get an instance of DownloadManager            
                Dim downloadManager As DownloadManager = apiSession.GetDownloadManager()
                ' A call to the DownloadFileToBuffer            
                Dim buffer As Byte() = downloadManager.DownloadFileToBuffer(docId, DownloadTypes.ORIGINAL)
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Sub

        Private Sub GetFile(ByVal roomId As Integer)
            Dim workspaces As Resource.Workspaces = apiSession.GetWorkspacesResource()
            ' Create an object to specify the details of what documents to list and how            
            ' they are returned. A few options are shown here.            
            Dim selectionJson As ListDocumentsVdrJson = New ListDocumentsVdrJson
            With selectionJson
                .OrderAscending = False
                .FolderPath = "/"
            End With
            ' Call the list method            
            Dim response As PagingItemListJson(Of BaseJson) = workspaces.ListDocumentsV30(roomId, selectionJson)
        End Sub


    End Module
End Namespace