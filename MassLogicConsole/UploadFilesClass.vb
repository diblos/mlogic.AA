Imports BlackBerry.Workspaces
Imports BlackBerry.Workspaces.Enums
Imports BlackBerry.Workspaces.Json
Imports System
Imports System.Collections.Generic

Namespace MassLogicConsole

    Public Class UploadFilesClass
        Private apiSession As ApiSession

        Public Sub New(apiSession As ApiSession)
            Me.apiSession = apiSession
        End Sub

        Public Function UploadDocument(localPath As String, filename As String, userRecipients As HashSet(Of String), ADGroupsRecipients As HashSet(Of String), listRecipients As HashSet(Of String)) As UploadResult
            Dim uploadManager As UploadManager = Me.apiSession.GetUploadManager()
            Dim uploadInfo As SubmitDocumentSdsJson = New SubmitDocumentSdsJson() With {.DocumentGuids = New HashSet(Of String)() From {uploadManager.GetNewGuidForDocument()}, .Permission = New PermissionFromUserJson() With {.Copy = New Boolean?(True), .Download = New Boolean?(True), .DownloadOriginal = New Boolean?(False), .ExpirationDate = New DateTime?(DateTime.Now)}, .UserRecipients = userRecipients, .ActiveDirectoryGroupsRecipients = ADGroupsRecipients, .ListRecipients = listRecipients, .WhoCanView = New WhoCanView?(WhoCanView.RECEIPIENTS_ONLY)}
            Return uploadManager.UploadDocument(uploadInfo, localPath, Nothing, filename)
        End Function

        Public Function UploadDocumentToRoom(roomId As Integer, destinationFileName As String, filename As String, folder As String, groups As List(Of String), domains As List(Of String)) As UploadResult
            Dim arg_54_0 As UploadManager = Me.apiSession.GetUploadManager()
            Dim uploadInfo As SubmitDocumentsVdrJson = New SubmitDocumentsVdrJson() With {.OpenForAllRoom = New Boolean?(False), .Recipients = New RoomRecipientsJson() With {.Groups = groups, .Domains = domains}, .Folder = folder, .TagValueList = Nothing, .DeviceType = DeviceType.SYNC}
            Return arg_54_0.UploadDocumentToRoom(uploadInfo, roomId, destinationFileName, filename, Nothing, False)
        End Function
    End Class

End Namespace