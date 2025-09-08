Public Class EventLogTester

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        GetLog(sourceName:=NameOf(EventLogTester)).LogEntry("Starting EventLog Tester", LoggingSeverity.Verbose)
    End Sub

End Class