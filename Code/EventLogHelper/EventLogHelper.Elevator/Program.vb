Imports System.Security.Principal
#If NET35 Then
Imports EventLogHelper.Elevator.Diagnostics
#End If
#If NETCOREAPP Then
Imports System.Windows.Forms
#End If

#If NET35 Then
<ExcludeFromCodeCoverage>
Module Program
#Else
<CodeAnalysis.ExcludeFromCodeCoverage>
Module Program
#End If

    ''' <summary>
    ''' Mains the specified arguments.
    ''' </summary>
    ''' <param name="args">The arguments.</param>
    Sub Main(args As String())

        If args.Length = 0 Then
            MessageBox.Show("This helper is not meant to be run directly.",
                    "EventLog Helper",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)
            Environment.ExitCode = 1 ' run directly with no arguments
        Else
            If CheckElevation() Then
                If args.Length < 3 OrElse args.Length > 5 Then
                    MessageBox.Show("Usage: EventLogHelperCreate.exe ""<LogName>"" ""<SourceName>"" ""<MachineName>"" [maxKilobytes=64000] [retentionDays=0]",
                        "EventLog Helper",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
                    Environment.ExitCode = 3 ' wrong number of args
                Else
                    Dim logName As String = args(0)
                    Dim sourceName As String = args(1)
                    Dim machineName As String = args(2)

                    Dim maxKilobytes As Integer = 64000 ' default to 64 MB
                    Dim retentionDays As Integer = 0

                    If args.Length >= 4 AndAlso Not Integer.TryParse(args(3), maxKilobytes) Then maxKilobytes = 64000
                    If args.Length >= 5 AndAlso Not Integer.TryParse(args(4), retentionDays) Then retentionDays = 0

                    Try
                        CreateLogSource(logName, logName, machineName, maxKilobytes, retentionDays)
                        CreateLogSource(logName, sourceName, machineName, maxKilobytes, retentionDays)
                        Environment.ExitCode = 0 ' success
                    Catch ex As Exception
                        MessageBox.Show($"Error creating log/source:{vbCrLf}{ex.Message}",
                           "EventLog Helper",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error)
                        Environment.ExitCode = 4 ' error creating log/source
                    End Try
                End If
            Else
                Environment.ExitCode = 2 ' not running elevated
            End If
        End If

    End Sub

    ''' <summary>
    ''' Checks the elevation.
    ''' </summary>
    ''' <returns></returns>
    Private Function CheckElevation() As Boolean

        Dim identity = WindowsIdentity.GetCurrent()
        Dim principal = New WindowsPrincipal(identity)
        Dim isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator)

        If Not isAdmin Then
            MessageBox.Show("This action requires administrative rights. " &
                    "Please run as administrator.",
                    "EventLog Helper",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            Return False
        End If

        Return True

    End Function

    ''' <summary>
    ''' Creates the log source.
    ''' </summary>
    ''' <param name="logName">Name of the log.</param>
    ''' <param name="sourceName">Name of the source.</param>
    ''' <param name="machineName">Name of the machine.</param>
    Private Sub CreateLogSource(
            ByRef logName As String,
            ByRef sourceName As String,
            ByRef machineName As String,
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer)

        Dim LogExists As Boolean = EventLog.Exists(logName, machineName)

        ' If the log doesn’t exist, we assume the source doesn’t either and create both.
        ' Otherwise, only create the source if it is missing.
        If Not LogExists OrElse Not EventLog.SourceExists(sourceName, machineName) Then
            Dim obj As New EventSourceCreationData(sourceName, logName) With {
                .MachineName = machineName
            }
            EventLog.CreateEventSource(obj)
            If Not LogExists Then ' If the log file did not exist, assume that we just created it and change its settings
                Using evt As New EventLog(logName, machineName, sourceName)
                    evt.MaximumKilobytes = maxKilobytes
                    If retentionDays > 0 Then
                        evt.ModifyOverflowPolicy(OverflowAction.OverwriteOlder, retentionDays)
                    Else
                        evt.ModifyOverflowPolicy(OverflowAction.OverwriteAsNeeded, 0)
                    End If
                End Using
            End If
        End If

    End Sub

End Module
