Imports EventLogHelper.Interfaces

''' <summary>
''' A test implementation of <see cref="IConfigShim"/> that simulates a failure by throwing an exception.
''' </summary>
''' <seealso cref="IConfigShim" />
Public Class ThrowingConfigShim

    Implements IConfigShim

    ''' <summary>
    ''' Gets the application setting.
    ''' </summary>
    ''' <param name="key">The key.</param>
    ''' <returns></returns>
    ''' <exception cref="System.Exception">Simulated failure</exception>
    Public Function GetAppSetting(key As String) As String Implements IConfigShim.GetAppSetting

        Throw New Exception("Simulated failure")

    End Function

End Class
