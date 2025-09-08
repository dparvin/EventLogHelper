''' <summary>
''' Defines a contract for configuration operations, allowing for abstraction and easier testing.
''' </summary>
Friend Interface IConfigShim

    ''' <summary>
    ''' Gets the application setting.
    ''' </summary>
    ''' <param name="key">The key.</param>
    ''' <returns></returns>
    Function GetAppSetting(key As String) As String

End Interface