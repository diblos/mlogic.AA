
Imports System.IO
Imports System.Reflection


Public Class LogWriter
    Private m_exePath As String = String.Empty
    Public Sub New(logMessage As String)
        LogWrite(logMessage)
    End Sub
    Public Sub LogWrite(logMessage As String)
        m_exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
        Try
            Using w As StreamWriter = File.AppendText((m_exePath & Convert.ToString("\")) + "log" & Now.ToString("yyyyMMdd") & ".txt")
                Log(logMessage, w)
            End Using
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Log(logMessage As String, txtWriter As TextWriter)
        Try
            txtWriter.Write(vbCr & vbLf & "Log Entry : ")
            'txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString())
            txtWriter.WriteLine("{0}", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"))
            'txtWriter.WriteLine("  :")
            txtWriter.WriteLine(vbTab & vbTab & "  : {0}", logMessage)
            'txtWriter.WriteLine("-------------------------------")
        Catch ex As Exception
        End Try
    End Sub
End Class

