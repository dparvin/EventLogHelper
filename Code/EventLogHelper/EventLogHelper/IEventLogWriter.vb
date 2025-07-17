Public Interface IEventLogWriter

    ''' <summary>
    ''' Writes the entry.
    ''' </summary>
    ''' <param name="log">The log.</param>
    ''' <param name="source">The source.</param>
    ''' <param name="message">The message.</param>
    ''' <param name="eventType">Type of the event.</param>
    ''' <param name="eventId">The event identifier.</param>
    ''' <param name="category">The category.</param>
    ''' <param name="rawData">The raw data.</param>
    Sub WriteEntry(
            log As String,
            source As String,
            message As String,
            eventType As EventLogEntryType,
            eventId As Integer,
            category As Short,
            rawData As Byte())

End Interface
