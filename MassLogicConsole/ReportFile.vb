
Namespace MassLogicConsole
    Friend Class ReportFile
        Private VSNfromFilename As String

        Private dateTimeFromFilename As String

        Private dstFilename As String

        Public Property dateTimeFromContent() As String

        Public Property absolutePath() As String

        Public Property fileName() As String

        Public Property OISVersion() As String

        Public Property platformName() As String

        Public Property platformType() As String

        Public Property WatchdoxFileContent() As String

        Public Property WatchdoxTargetFilename() As String

        Public Sub New(absolutePath As String, fileName As String, VSNfromFilename As String, dateTimeFromFilename As String)
            Me.absolutePath = absolutePath
            Me.fileName = fileName
            Me.VSNfromFilename = VSNfromFilename
            Me.dateTimeFromFilename = dateTimeFromFilename
            Me.dstFilename = Me.getDstFilename(dateTimeFromFilename)
        End Sub

        Public Function getDstFilename() As String
            Return Me.dstFilename
        End Function

        Public Function getVSNfromFilename() As String
            Return Me.VSNfromFilename
        End Function

        Private Function getDstFilename(dateTimeFromFilename As String) As String
            Dim str As String = Nothing
            Dim str2 As String = Nothing
            Dim str3 As String = Nothing
            Try
                Dim expr_0E As String = dateTimeFromFilename.Substring(0, 8)
                str = expr_0E.Substring(0, 4)
                str2 = expr_0E.Substring(4, 2)
                str3 = expr_0E.Substring(6, 2)
            Catch ex_2A As ArgumentOutOfRangeException
            End Try
            Return str3 + str2 + str
        End Function
    End Class
End Namespace