''' <summary>
''' Represents the logging levels for event logging.
''' This enumeration is used to define the granularity of log messages.
''' </summary>
''' <remarks>
''' The levels range from None (no logging) to Critical (most severe). 
''' Log entries with this enumeration will only be written to the event log 
''' when the <see cref="LoggingLevel"/> is high enough for the item being logged.
''' </remarks>
Public Enum LoggingLevel
    ''' <summary>
    ''' The none
    ''' </summary>
    None = 99
    ''' <summary>
    ''' The verbose
    ''' </summary>
    Verbose = 0
    ''' <summary>
    ''' The information
    ''' </summary>
    Normal = 1
    ''' <summary>
    ''' The warning
    ''' </summary>
    Minimal = 2
    ''' <summary>
    ''' The error
    ''' </summary>
    [Error] = 3
    ''' <summary>
    ''' The critical
    ''' </summary>
    Critical = 4
End Enum