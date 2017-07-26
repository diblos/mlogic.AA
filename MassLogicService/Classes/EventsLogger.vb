Imports System.Diagnostics

Namespace MassLogicConsole

    Public Class EventsLogger

        Dim _sSource As String
        Public Property Number() As String
            Get
                Return Me._sSource
            End Get
            Set(ByVal value As String)
                Me._sSource = value
            End Set
        End Property

        Dim _sMachine As String
        Public Property Machine() As String
            Get
                Return Me._sMachine
            End Get
            Set(ByVal value As String)
                Me._sMachine = value
            End Set
        End Property

        Dim ELog As EventLog
        Private Const _sLog As String = "Application"

        Public Sub New()

            'DEFAULT VALUES
            Me._sSource = "dotNET Sample App"
            Me._sMachine = "."

            Initiate()

        End Sub

        Public Sub New(ByVal Source As String, ByVal Machine As String)

            'SET VALUES
            Me._sSource = Source
            Me._sMachine = Machine

            Initiate()

        End Sub

        Private Sub Initiate()
            'If Not EventLog.SourceExists(Me._sSource, Me._sMachine) Then
            '    EventLog.CreateEventSource(Me._sSource, _sLog, Me._sMachine)
            'End If
            ELog = New EventLog(_sLog, Me._sMachine, Me._sSource)
        End Sub

        Public Sub WriteEvent(ByVal EventMessage As String)
            ELog.WriteEntry(EventMessage)
            'ELog.WriteEntry(EventMessage, EventLogEntryType.Warning, 234, CType(3, Short))
        End Sub

    End Class

End Namespace
