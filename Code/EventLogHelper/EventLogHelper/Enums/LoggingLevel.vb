''' <summary>
''' Defines the global logging level (threshold) that controls which log entries
''' are written to the Windows Event Log.
''' </summary>
''' <remarks>
''' <para>
''' This value acts as a filter: only log entries with a 
''' <see cref="LoggingSeverity"/> greater than or equal to the current
''' <see cref="LoggingLevel"/> will be written. 
''' </para>
''' 
''' <para>
''' For example, if the <c>CurrentLoggingLevel</c> is <c>Normal</c>, 
''' then entries with severities of <c>Info</c>, <c>Warning</c>, <c>Error</c>, 
''' or <c>Critical</c> are logged, while <c>Verbose</c> messages are suppressed.
''' </para>
''' 
''' <para>
''' The levels are ordered from least to most restrictive:
''' <list type="bullet">
'''   <item><term>Verbose (0)</term><description>All messages are written.</description></item>
'''   <item><term>Normal (1)</term><description>Suppresses <c>Verbose</c>, writes Info and higher.</description></item>
'''   <item><term>Minimal (2)</term><description>Suppresses <c>Verbose</c> and <c>Info</c>, writes Warning and higher.</description></item>
'''   <item><term>Error (3)</term><description>Writes only Error and Critical entries.</description></item>
'''   <item><term>Critical (4)</term><description>Writes only Critical entries.</description></item>
'''   <item><term>None (99)</term><description>No entries are written.</description></item>
''' </list>
''' </para>
''' </remarks>
Public Enum LoggingLevel
    ''' <summary>
    ''' No log entries will be written.
    ''' </summary>
    None = 99

    ''' <summary>
    ''' Write all log entries, including diagnostic Verbose entries.
    ''' </summary>
    Verbose = 0

    ''' <summary>
    ''' Write Info, Warning, Error, and Critical entries; suppress Verbose.
    ''' </summary>
    Normal = 1

    ''' <summary>
    ''' Write Warning, Error, and Critical entries; suppress Info and Verbose.
    ''' </summary>
    Minimal = 2

    ''' <summary>
    ''' Write Error and Critical entries only.
    ''' </summary>
    [Error] = 3

    ''' <summary>
    ''' Write only Critical entries.
    ''' </summary>
    Critical = 4
End Enum