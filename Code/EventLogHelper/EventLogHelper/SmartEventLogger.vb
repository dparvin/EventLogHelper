Imports Microsoft.Win32
Imports System.Reflection

''' <summary>
''' Provides high-level utilities for logging to the Windows Event Log in a simplified and customizable way.
''' </summary><remarks>
''' <para>
''' This module abstracts common patterns for writing entries to the Windows Event Log. It simplifies usage by handling:
''' </para>
''' <list type="bullet">
'''   <item>Automatic validation and creation of logs and sources.</item>
'''   <item>Falls back when <c>log</c> or <c>source</c> parameters are not provided.</item>
'''   <item>Automatic source naming based on the calling method and class context.</item>
'''   <item>Safe handling and truncation of long messages to comply with system limits.</item>
''' </list>
''' <para>
''' The default event log is <c>Application</c>. If no source is provided, a fully qualified method name
''' (including namespace and class) is used as the source identifier.
''' </para>
''' <para>
''' This module is suitable for both production use (via <see cref="RealEventLogWriter" />) and unit testing
''' (via a test implementation such as <c>EventLogHelper.Tests.TestEventLogWriter</c> in a separate test project).
''' </para>
''' </remarks>
Public Module SmartEventLogger

    ''' <summary>
    ''' The current event log writer instance used by the logger.
    ''' </summary>
    ''' <remarks>
    ''' By default, this is initialized to an instance of <see cref="RealEventLogWriter"/>, which writes to the actual
    ''' Windows Event Log. This field is typically replaced during testing with a mock or stub implementation via the
    ''' <see cref="SetWriter"/> method.
    ''' </remarks>
    Private _writer As IEventLogWriter = New RealEventLogWriter()
    ''' <summary>
    ''' The registry reader
    ''' </summary>
    Private _registryReader As IRegistryReader = New RealRegistryReader()

    ''' <summary>
    ''' Sets the <see cref="IEventLogWriter"/> implementation used for logging operations.
    ''' </summary>
    ''' <param name="writer">
    ''' The custom writer instance to use. This allows swapping out the default <see cref="RealEventLogWriter"/>
    ''' for testing or alternate logging implementations.
    ''' </param>
    ''' <remarks>
    ''' This method is primarily intended for use in unit tests or advanced scenarios where log output should be redirected
    ''' or suppressed. For example, a test implementation may capture and inspect log data without writing to the real event log.
    ''' </remarks>
    Public Sub SetWriter(writer As IEventLogWriter)
        _writer = writer
    End Sub

    ''' <summary>
    ''' Sets the registry reader.
    ''' </summary>
    ''' <param name="reader">The reader.</param>
    Public Sub SetRegistryReader(reader As IRegistryReader)

        _registryReader = reader

    End Sub

#Region " Default Value Properties ^^^^^^^^^^^^^^^^^^^^^^^^ "

    ''' <summary>
    ''' Gets or sets the name of the machine.
    ''' </summary>
    ''' <value>
    ''' The name of the machine.
    ''' </value>
    Public Property MachineName As String = "."

    ''' <summary>
    ''' Gets or sets the name of the log.
    ''' </summary>
    ''' <value>
    ''' The name of the log.
    ''' </value>
    Public Property LogName As String = "Application"

    ''' <summary>
    ''' Gets or sets the name of the source.
    ''' </summary>
    ''' <value>
    ''' The name of the source.
    ''' </value>
    Public Property SourceName As String = ""

    ''' <summary>
    ''' Gets or sets the maximum kilobytes.
    ''' </summary>
    ''' <value>
    ''' The maximum kilobytes.
    ''' </value>
    Public Property MaxKilobytes As Integer = 1024 * 1024

    ''' <summary>
    ''' Gets or sets the retention days.
    ''' </summary>
    ''' <value>
    ''' The retention days.
    ''' </value>
    Public Property RetentionDays As Integer = 7

    ''' <summary>
    ''' Gets or sets a value indicating whether to write initialize entry when creating a new log file.
    ''' </summary>
    ''' <value>
    '''   <c>true</c> if while creating a log, write initialize entry; otherwise, <c>false</c>.
    ''' </value>
    Public Property WriteInitEntry As Boolean = False

    ''' <summary>
    ''' Gets or sets the truncation marker.
    ''' </summary>
    ''' <value>
    ''' The truncation marker.
    ''' </value>
    Public Property TruncationMarker As String = "... [TRUNCATED]"

#End Region

#Region " Log Methods ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ "

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String)

        Log("", "", message)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal source As String,
            ByVal message As String)

        Log("", source, message)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType)

        Log("", message, eventType)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="sourceName">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal logName As String,
            ByVal sourceName As String,
            ByVal message As String)

        Log(logName, sourceName, message, EventLogEntryType.Information)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="sourceName">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal sourceName As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType)

        Log("", sourceName, message, eventType)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="sourceName">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal logName As String,
            ByVal sourceName As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType)

        Log(logName, sourceName, message, eventType, 0)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer)

        Log("", message, eventType, eventID)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="sourceName">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal sourceName As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer)

        Log("", sourceName, message, eventType, eventID)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="sourceName">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal logName As String,
            ByVal sourceName As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer)

        Log(logName, sourceName, message, eventType, eventID, 0)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short)

        Log("", message, eventType, eventID, category)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="sourceName">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal sourceName As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short)

        Log("", sourceName, message, eventType, eventID, category)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="sourceName">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal logName As String,
            ByVal sourceName As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short)

        Log(MachineName, logName, sourceName, message, eventType, eventID, category, Nothing)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal rawData As Byte())

        Log(message, EventLogEntryType.Information, rawData)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal rawData As Byte())

        Log(message, eventType, 0, rawData)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal rawData As Byte())

        Log(message, eventType, eventID, 0, rawData)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte())

        Log(MachineName, "", "", message, eventType, eventID, category, rawData)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="sourceName">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal sourceName As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte())

        Log(MachineName, "", sourceName, message, eventType, eventID, category, rawData)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal logName As String,
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte())

        Log(
            "",
            logName,
            source,
            message,
            eventType,
            eventID,
            category,
            rawData)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal machineName As String,
            ByVal logName As String,
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte())

        Log(
            machineName,
            logName,
            source,
            message,
            eventType,
            eventID,
            category,
            rawData,
            MaxKilobytes)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal machineName As String,
            ByVal logName As String,
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte(),
            ByVal maxKilobytes As Integer)

        Log(
            machineName,
            logName,
            source,
            message,
            eventType,
            eventID,
            category,
            rawData,
            maxKilobytes,
            RetentionDays)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer)

        Log(
            message,
            EventLogEntryType.Information,
            maxKilobytes,
            retentionDays)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer)

        Log(
            message,
            eventType,
            0,
            maxKilobytes,
            retentionDays)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer)

        Log(
            message,
            eventType,
            eventID,
            0,
            maxKilobytes,
            retentionDays)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer)

        Log(
            message,
            eventType,
            eventID,
            category,
            Nothing,
            maxKilobytes,
            retentionDays)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte(),
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer)

        Log(
            "",
            message,
            eventType,
            eventID,
            category,
            rawData,
            maxKilobytes,
            retentionDays)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte(),
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer)

        Log(
            "",
            source,
            message,
            eventType,
            eventID,
            category,
            rawData,
            maxKilobytes,
            retentionDays)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal logName As String,
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte(),
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer)

        Log(
            MachineName,
            logName,
            source,
            message,
            eventType,
            eventID,
            category,
            rawData,
            maxKilobytes,
            retentionDays)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal machineName As String,
            ByVal logName As String,
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte(),
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer)

        Log(
            machineName,
            logName,
            source,
            message,
            eventType,
            eventID,
            category,
            rawData,
            maxKilobytes,
            retentionDays,
            WriteInitEntry)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <param name="writeInitEntry">
    ''' Indicates whether an initialization entry should be written when a new log is created.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer,
            ByVal writeInitEntry As Boolean)

        Log(
            message,
            EventLogEntryType.Information,
            maxKilobytes,
            retentionDays,
            writeInitEntry)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <param name="writeInitEntry">
    ''' Indicates whether an initialization entry should be written when a new log is created.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer,
            ByVal writeInitEntry As Boolean)

        Log(
            message,
            eventType,
            0,
            maxKilobytes,
            retentionDays,
            writeInitEntry)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <param name="writeInitEntry">
    ''' Indicates whether an initialization entry should be written when a new log is created.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer,
            ByVal writeInitEntry As Boolean)

        Log(
            message,
            eventType,
            eventID,
            0,
            maxKilobytes,
            retentionDays,
            writeInitEntry)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <param name="writeInitEntry">
    ''' Indicates whether an initialization entry should be written when a new log is created.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer,
            ByVal writeInitEntry As Boolean)

        Log(
            message,
            eventType,
            eventID,
            category,
            Nothing,
            maxKilobytes,
            retentionDays,
            writeInitEntry)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <param name="writeInitEntry">
    ''' Indicates whether an initialization entry should be written when a new log is created.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte(),
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer,
            ByVal writeInitEntry As Boolean)

        Log(
            "",
            message,
            eventType,
            eventID,
            category,
            rawData,
            maxKilobytes,
            retentionDays,
            writeInitEntry)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="sourceName">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <param name="writeInitEntry">
    ''' Indicates whether an initialization entry should be written when a new log is created.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal sourceName As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte(),
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer,
            ByVal writeInitEntry As Boolean)

        Log(
            "",
            sourceName,
            message,
            eventType,
            eventID,
            category,
            rawData,
            maxKilobytes,
            retentionDays,
            writeInitEntry)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="sourceName">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <param name="writeInitEntry">
    ''' Indicates whether an initialization entry should be written when a new log is created.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal logName As String,
            ByVal sourceName As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte(),
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer,
            ByVal writeInitEntry As Boolean)

        Log(
            MachineName,
            logName,
            sourceName,
            message,
            eventType,
            eventID,
            category,
            rawData,
            maxKilobytes,
            retentionDays,
            writeInitEntry)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the specified parameters.
    ''' </summary>
    ''' <param name="_machineName">
    ''' The name of the machine where the event log resides. Use <c>"."</c> to refer to the local machine.
    ''' If <c>null</c> or empty, the local machine is used by default, or the value in <c>MachinName</c> if set.
    ''' </param>
    ''' <param name="_logName">
    ''' The name of the event log to write to (e.g., <c>"Application"</c>, <c>"System"</c>, or a custom log name).
    ''' If <c>null</c> or empty, the value of the static <c>LogName</c> property is used.
    ''' </param>
    ''' <param name="sourceName">
    ''' The source name to associate with the event log entry. If <c>null</c> or empty, a source is automatically
    ''' generated using the calling method's namespace, class, and method name.
    ''' If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to write to the event log. If <c>null</c> or empty, a default message is used.
    ''' Messages longer than 32,766 characters are truncated and suffixed with <c>"... [TRUNCATED]"</c>
    ''' unless multi-entry logging is enabled.
    ''' The message is automatically wrapped with the source name if it does not start with a square bracket.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event. Valid values are <see cref="EventLogEntryType.Error"/>,
    ''' <see cref="EventLogEntryType.Warning"/>, and <see cref="EventLogEntryType.Information"/>.
    ''' If an invalid value is passed, <c>Information</c> is used.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier. This is application-defined and can be used for grouping or filtering related events.
    ''' </param>
    ''' <param name="category">
    ''' The event category. This can be used in categorized event log views. Set to 0 if unused.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to associate with the event entry. Can be <c>null</c>.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <param name="writeInitEntry">
    ''' Indicates whether an initialization entry should be written when a new log is created.
    ''' </param>
    ''' <remarks>
    ''' <para>
    ''' If the specified log or source does not exist, they are automatically created.
    ''' </para>
    ''' <para>
    ''' This method ensures the message is formatted consistently and applies truncation if needed.
    ''' If <see cref="TruncationMarker"/> is set, it is used instead of the default <c>"... [TRUNCATED]"</c>.
    ''' </para>
    ''' <para>
    ''' This method uses the configured <see cref="_writer"/> to write the event.
    ''' </para>
    ''' </remarks>
    Public Sub Log(
            ByVal _machineName As String,
            ByVal _logName As String,
            ByVal sourceName As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte(),
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer,
            ByVal writeInitEntry As Boolean)

        Dim defaultLog As String = If(String.IsNullOrEmpty(_logName), LogName, _logName.Trim())
        Dim defaultSource As String = Source(sourceName)
        Dim defaultMachine As String = If(String.IsNullOrEmpty(_machineName), If(String.IsNullOrEmpty(MachineName), ".", MachineName), _machineName)

        If Not _writer.Exists(defaultLog, defaultMachine) OrElse Not _writer.SourceExists(defaultSource, defaultMachine) Then
            _writer.CreateEventSource(defaultSource, defaultLog, _machineName)
        End If
        Dim finalMessage As String = If(String.IsNullOrEmpty(message), "No message provided.", message.Trim())
        If Not finalMessage.EndsWith("."c) Then finalMessage &= "."
        If Not finalMessage.StartsWith("["c) Then finalMessage = $"[{defaultSource}] {finalMessage}"

        Const maxLength As Integer = 32766
        Const _truncationMarker As String = "... [TRUNCATED]"
        If String.IsNullOrEmpty(TruncationMarker) Then TruncationMarker = _truncationMarker

        If finalMessage.Length > maxLength Then
            ' Leave room for the truncation marker
            Dim maxWithoutMarker As Integer = maxLength - TruncationMarker.Length
            finalMessage = $"{finalMessage.Substring(0, maxWithoutMarker)}{TruncationMarker}"
        End If
        Select Case eventType
            Case EventLogEntryType.Error, EventLogEntryType.Warning, EventLogEntryType.Information
                ' Valid
            Case Else
                eventType = EventLogEntryType.Information
        End Select
        _writer.WriteEntry(_machineName,
                           defaultLog,
                           defaultSource,
                           finalMessage,
                           eventType,
                           eventID,
                           category,
                           rawData,
                           maxKilobytes,
                           retentionDays,
                           writeInitEntry)

    End Sub

#End Region

#Region " Support Functions ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ "

    ''' <summary>
    ''' Returns a default event source name. If none is provided, this attempts to determine the calling
    ''' method's fully qualified name as the source.
    ''' </summary>
    ''' <param name="_sourceName">Optional. A specific source name to use. If null or empty, the calling method will be used.</param>
    ''' <returns>A valid event source name derived from the input or calling method.</returns>
    Public Function Source(
            Optional ByVal _sourceName As String = Nothing) As String

        ' Final fallback if nothing else is available
        Dim strReturn As String = "SmartEventLogger"

        If String.IsNullOrEmpty(_sourceName) Then
            If Not String.IsNullOrEmpty(SourceName) Then
                strReturn = SourceName
            Else
                ' Try to get the calling assembly (e.g., the one calling your library)
                Dim thisAssembly As Assembly = Assembly.GetExecutingAssembly()

                Dim stackTrace As New StackTrace()
                For i As Integer = 1 To stackTrace.FrameCount - 1
                    Dim frame As StackFrame = stackTrace.GetFrame(i)
                    Dim method As MethodBase = frame.GetMethod()
                    Dim type As Type = method.DeclaringType
                    If type.Assembly IsNot thisAssembly Then
                        strReturn = $"{type.FullName}.{method.Name}"
                        Exit For
                    End If
                Next
            End If
        Else
            strReturn = _sourceName.Trim()
        End If

        ' Ensure the source name does not exceed the maximum length allowed by the Windows Event Log system
        Return MaxSource(strReturn)

    End Function

    ''' <summary>
    ''' Attempts to determine a default source name for the given event log by inspecting the registry.
    ''' </summary>
    ''' <param name="log">The name of the event log (e.g., "Application", "System").</param>
    ''' <returns>The name of a registered event source, or an empty string if none are found.</returns>
    Public Function DefaultSource(
            ByVal log As String) As String

        Return DefaultSource(log, MachineName)

    End Function

    ''' <summary>
    ''' Attempts to determine a default source name for the given event log by inspecting the registry.
    ''' </summary>
    ''' <param name="logName">The name of the event log (e.g., "Application", "System").</param>
    ''' <param name="machineName">The target machine name, or "." for the local machine.</param>
    ''' <returns>The name of a registered event source, or an empty string if none are found.</returns>
    Public Function DefaultSource(
            ByVal logName As String,
            ByVal machineName As String) As String

        Return _registryReader.GetDefaultEventLogSource(logName, machineName)

    End Function

    ''' <summary>
    ''' Truncates the source name to meet the maximum length allowed by the Windows Event Log system.
    ''' </summary>
    ''' <param name="sourceName">The original source name.</param>
    ''' <returns>A source name no longer than 211 characters.</returns>
    ''' <remarks>
    ''' If the input is null or empty, the method will derive a source name using the calling method.
    ''' </remarks>
    Public Function MaxSource(
        ByVal sourceName As String) As String

        If String.IsNullOrEmpty(sourceName) Then
            sourceName = Source()
        End If

        Return sourceName.Substring(0, If(sourceName.Length > 211, 211, sourceName.Length))

    End Function

    ''' <summary>
    ''' Checks whether a specific source is registered to a specific event log on the specified machine.
    ''' </summary>
    ''' <param name="sourceName">The name of the event source.</param>
    ''' <param name="logName">The name of the event log (e.g., "Application").</param>
    ''' <param name="machineName">The target machine name, or "." for the local machine.</param>
    ''' <returns><c>True</c> if the source is registered; otherwise, <c>False</c>.</returns>
    Public Function IsSourceRegisteredToLog(
            ByVal sourceName As String,
            ByVal logName As String,
            ByVal machineName As String) As Boolean

        Dim registryPath As String = $"SYSTEM\CurrentControlSet\Services\EventLog\{logName}\{sourceName}"

        Return _registryReader.SubKeyExists(RegistryHive.LocalMachine, machineName, registryPath)

    End Function

    ''' <summary>
    ''' Checks if the current user has read access to the registry key for a specific event log source.
    ''' </summary>
    ''' <param name="sourceName">The name of the event source.</param>
    ''' <param name="logName">The name of the event log.</param>
    ''' <param name="machineName">The machine to check, or "." for local machine.</param>
    ''' <returns><c>True</c> if the registry key can be read; otherwise, <c>False</c>.</returns>
    Public Function CanReadRegistryForLogAndSource(
            ByVal sourceName As String,
            ByVal logName As String,
            ByVal machineName As String) As Boolean

        Dim registryPath As String = $"SYSTEM\CurrentControlSet\Services\EventLog\{logName}\{sourceName}"

        Return _registryReader.HasRegistryAccess(RegistryHive.LocalMachine, machineName, registryPath, False)

    End Function

    ''' <summary>
    ''' Determines if the current user has access (read or write) to the registry key for a given event log.
    ''' </summary>
    ''' <param name="logName">The name of the event log.</param>
    ''' <param name="machineName">The target machine name, or "." for local machine.</param>
    ''' <param name="writeAccess">Optional. <c>True</c> to check for write access; otherwise, <c>False</c> for read-only access.</param>
    ''' <returns><c>True</c> if access is permitted; otherwise, <c>False</c>.</returns>
    Public Function CanReadRegistryForLog(
            ByVal logName As String,
            ByVal machineName As String,
            Optional ByVal writeAccess As Boolean = False) As Boolean

        Dim registryPath As String = $"SYSTEM\CurrentControlSet\Services\EventLog\{logName}"

        Return _registryReader.HasRegistryAccess(RegistryHive.LocalMachine, machineName, registryPath, writeAccess)

    End Function

#End Region

#Region " GetLog routines ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ "

    ''' <summary>
    ''' Gets the entry log class to use for writing to the event log.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to write to (e.g., <c>"Application"</c>, <c>"System"</c>, or a custom log name).
    ''' If <c>null</c> or empty, the value of the static <c>LogName</c> property is used.
    ''' </param>
    ''' <param name="sourceName">
    ''' The source name to associate with the event log entry. If <c>null</c> or empty, a source is automatically
    ''' generated using the calling method's namespace, class, and method name.
    ''' If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <returns></returns>
    Public Function GetLog(
            ByVal logName As String,
            ByVal sourceName As String) As EventLog

        Return GetLog(
            logName,
            sourceName,
            MaxKilobytes,
            RetentionDays)

    End Function

    ''' <summary>
    ''' Gets the entry log class to use for writing to the event log.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to write to (e.g., <c>"Application"</c>, <c>"System"</c>, or a custom log name).
    ''' If <c>null</c> or empty, the value of the static <c>LogName</c> property is used.
    ''' </param>
    ''' <param name="sourceName">
    ''' The source name to associate with the event log entry. If <c>null</c> or empty, a source is automatically
    ''' generated using the calling method's namespace, class, and method name.
    ''' If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <returns></returns>
    Public Function GetLog(
            ByVal logName As String,
            ByVal sourceName As String,
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer) As EventLog

        Return GetLog(
            logName,
            sourceName,
            maxKilobytes,
            retentionDays,
            WriteInitEntry)

    End Function

    ''' <summary>
    ''' Gets the entry log class to use for writing to the event log.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to write to (e.g., <c>"Application"</c>, <c>"System"</c>, or a custom log name).
    ''' If <c>null</c> or empty, the value of the static <c>LogName</c> property is used.
    ''' </param>
    ''' <param name="sourceName">
    ''' The source name to associate with the event log entry. If <c>null</c> or empty, a source is automatically
    ''' generated using the calling method's namespace, class, and method name.
    ''' If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <param name="writeInitEntry">
    ''' Indicates whether an initialization entry should be written when a new log is created.
    ''' </param>
    ''' <returns></returns>
    Public Function GetLog(
            ByVal logName As String,
            ByVal sourceName As String,
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer,
            ByVal writeInitEntry As Boolean) As EventLog

        Return GetLog(
            MachineName,
            logName,
            sourceName,
            maxKilobytes,
            retentionDays,
            writeInitEntry)

    End Function

    ''' <summary>
    ''' Gets the entry log class to use for writing to the event log.
    ''' </summary>
    ''' <param name="machineName">
    ''' The name of the machine where the event log resides. Use <c>"."</c> to refer to the local machine.
    ''' If <c>null</c> or empty, the local machine is used by default, or the value in <c>MachinName</c> if set.
    ''' </param>
    ''' <param name="logName">
    ''' The name of the event log to write to (e.g., <c>"Application"</c>, <c>"System"</c>, or a custom log name).
    ''' If <c>null</c> or empty, the value of the static <c>LogName</c> property is used.
    ''' </param>
    ''' <param name="sourceName">
    ''' The source name to associate with the event log entry. If <c>null</c> or empty, a source is automatically
    ''' generated using the calling method's namespace, class, and method name.
    ''' If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="maxKilobytes">
    ''' The maximum size of the event log in kilobytes. Only used when creating a new log.
    ''' Must be between 64 KB and 4 GB. If 0 or negative, the system default is used.
    ''' </param>
    ''' <param name="retentionDays">
    ''' The number of days to retain event log entries. If 0, events are retained indefinitely.
    ''' Only used when creating a new log.
    ''' </param>
    ''' <param name="writeInitEntry">
    ''' Indicates whether an initialization entry should be written when a new log is created.
    ''' </param>
    ''' <returns></returns>
    Public Function GetLog(
            ByVal machineName As String,
            ByVal logName As String,
            ByVal sourceName As String,
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer,
            ByVal writeInitEntry As Boolean) As EventLog

        Return _writer.GetLog(
            machineName,
            logName,
            sourceName,
            maxKilobytes,
            retentionDays,
            writeInitEntry)

    End Function

#End Region

End Module
