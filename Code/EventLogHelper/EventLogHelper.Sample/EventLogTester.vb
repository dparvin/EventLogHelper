Imports System.Text
Imports EventLogHelper.Enums
Imports Microsoft.Win32

Public Class EventLogTester

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ckbCategory.Checked = True
        ckbEventID.Checked = True
        ckbSeverity.Checked = True
        ckbRawData.Checked = True
        lblStatus.Text = ""

        ' Add any initialization after the InitializeComponent() call.
        GetLog(sourceName:=NameOf(EventLogTester)).LogEntry("Starting EventLog Tester", LoggingSeverity.Verbose)

    End Sub

    Private Sub EventLogTester_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Log names
        LoadLogNames()

        ' Severity enum
        cboSeverity.DataSource = [Enum].GetValues(Of LoggingSeverity)()

        cboEventType.DataSource = [Enum].GetValues(Of EventLogEntryType)()

        ' Reasonable defaults
        cboSourceName.Text = NameOf(EventLogTester)
        numEventID.Maximum = UShort.MaxValue
        numEventID.Value = 1000
        numCategory.Maximum = Short.MaxValue
        numCategory.Value = 0
        UpdateOptionalInputs()

    End Sub

    Private Sub LoadLogNames()

        cboLogName.BeginUpdate()
        Try
            cboLogName.Items.Clear()

            For Each log As EventLog In EventLog.GetEventLogs()
                cboLogName.Items.Add(log.Log)
            Next

            ' Default to Application if present
            Dim index As Integer = cboLogName.FindStringExact("Application")
            If index >= 0 Then
                cboLogName.SelectedIndex = index
            ElseIf cboLogName.Items.Count > 0 Then
                cboLogName.SelectedIndex = 0
            End If
        Finally
            cboLogName.EndUpdate()
        End Try

    End Sub

    Private Sub LoadSourcesForSelectedLog()

        Dim logName As String = cboLogName.Text

        If String.IsNullOrWhiteSpace(logName) Then
            ClearSources()
            Return
        End If

        ' Is this an existing log?
        Dim existingLogIndex As Integer = cboLogName.FindStringExact(logName)
        If existingLogIndex < 0 Then
            ' User typed a new log name
            ClearSources()
            Return
        End If

        Dim currentText As String = cboSourceName.Text

        cboSourceName.BeginUpdate()
        Try
            cboSourceName.Items.Clear()

            Using logKey As RegistryKey = Registry.LocalMachine.OpenSubKey(
                    $"SYSTEM\CurrentControlSet\Services\EventLog\{logName}",
                    writable:=False)

                If logKey Is Nothing Then
                    ClearSources(currentText)
                    Return
                End If

                For Each SourceName As String In logKey.GetSubKeyNames()
                    cboSourceName.Items.Add(SourceName)
                Next
            End Using

            If Not String.IsNullOrWhiteSpace(currentText) Then
                cboSourceName.Text = currentText
            ElseIf cboSourceName.Items.Count > 0 Then
                cboSourceName.SelectedIndex = 0
            End If

        Catch
            ' Access denied, corrupted keys, etc.
            ClearSources(currentText)
        Finally
            cboSourceName.EndUpdate()
        End Try

    End Sub

    Private Sub ClearSources(Optional keepText As String = "")

        cboSourceName.BeginUpdate()
        Try
            cboSourceName.Items.Clear()
            cboSourceName.Text = keepText
        Finally
            cboSourceName.EndUpdate()
        End Try

    End Sub

    ''' <summary>
    ''' When the log name changes, we need to update the list of sources. If the user types in a new log 
    ''' name, we'll just clear the sources since we won't know what they are until the log is created.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub LogName_Changed(
        sender As Object,
        e As EventArgs) Handles cboLogName.SelectedIndexChanged,
                                cboLogName.TextChanged

        LoadSourcesForSelectedLog()

    End Sub

    Private Sub CheckedChanged(
        sender As Object,
        e As EventArgs) Handles ckbSeverity.CheckedChanged,
                                ckbEventID.CheckedChanged,
                                ckbCategory.CheckedChanged,
                                ckbRawData.CheckedChanged,
                                ckbEventType.CheckedChanged

        UpdateOptionalInputs()

    End Sub

    Private Sub UpdateOptionalInputs()

        cboSeverity.Enabled = ckbSeverity.Checked
        numEventID.Enabled = ckbEventID.Checked
        numCategory.Enabled = ckbCategory.Checked
        txtRawData.Enabled = ckbRawData.Checked
        cboEventType.Enabled = ckbEventType.Checked

    End Sub

    Private Sub Send_Click(
        sender As Object,
        e As EventArgs) Handles btnSend.Click

        Dim logName As String = cboLogName.Text.Trim()
        Dim sourceName As String = cboSourceName.Text.Trim()
        Dim message As String = txtMessage.Text

        If String.IsNullOrWhiteSpace(logName) Then
            logName = "Application"
        End If

        If String.IsNullOrWhiteSpace(sourceName) Then
            sourceName = NameOf(EventLogTester)
        End If

        Dim eventLog As EventLog = GetLog(logName, sourceName)

        Dim useEventType As Boolean = ckbEventType.Checked
        Dim useSeverity As Boolean = ckbSeverity.Checked
        Dim useEventId As Boolean = ckbEventID.Checked
        Dim useCategory As Boolean = ckbCategory.Checked
        Dim useRawData As Boolean = ckbRawData.Checked

        Dim eventType As EventLogEntryType = CType(cboEventType.SelectedItem, EventLogEntryType)
        Dim severity As LoggingSeverity = CType(cboSeverity.SelectedItem, LoggingSeverity)
        Dim eventId As Integer = CInt(numEventID.Value)
        Dim category As Short = CShort(numCategory.Value)
        Dim rawData As Byte() = If(useRawData AndAlso Not String.IsNullOrEmpty(txtRawData.Text),
            Encoding.UTF8.GetBytes(txtRawData.Text),
            Nothing)
        Dim resolvedEventType As EventLogEntryType = If(useEventType, eventType, EventLogEntryType.Information)
        Dim resolvedEventId As Integer = If(useEventId, eventId, 0)
        Dim resolvedCategory As Short = If(useCategory, category, CType(0, Short))

        If useEventType AndAlso useEventId AndAlso useCategory AndAlso useRawData Then
            If useSeverity Then
                eventLog.LogEntry(message, resolvedEventType, resolvedEventId, resolvedCategory, rawData, severity)
            Else
                eventLog.LogEntry(message, resolvedEventType, resolvedEventId, resolvedCategory, rawData)
            End If
        ElseIf useEventType AndAlso useEventId AndAlso useCategory Then
            If useSeverity Then
                eventLog.LogEntry(message, resolvedEventType, resolvedEventId, resolvedCategory, severity)
            Else
                eventLog.LogEntry(message, resolvedEventType, resolvedEventId, resolvedCategory)
            End If
        ElseIf useEventType AndAlso useEventId Then
            If useSeverity Then
                eventLog.LogEntry(message, resolvedEventType, resolvedEventId, severity)
            Else
                eventLog.LogEntry(message, resolvedEventType, resolvedEventId)
            End If
        ElseIf useEventType Then
            If useSeverity Then
                eventLog.LogEntry(message, resolvedEventType, severity)
            Else
                eventLog.LogEntry(message, resolvedEventType)
            End If
        Else
            If useSeverity Then
                eventLog.LogEntry(message, severity)
            Else
                eventLog.LogEntry(message)
            End If
        End If

        lblStatus.Text = $"Logged at {DateTime.Now:T}"
        tmr.Start()

    End Sub

    Private Sub Tmr_Tick(
            sender As Object,
            e As EventArgs) Handles tmr.Tick

        tmr.Stop()
        lblStatus.Text = ""

    End Sub
End Class