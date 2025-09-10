Imports EventLogHelper.Interfaces

''' <summary>
''' A real implementation of <see cref="IConfigShim"/> that uses the actual configuration manager.
''' </summary>
''' <seealso cref="IConfigShim" />
#If NET35 Then
<Diagnostics.ExcludeFromCodeCoverage>
Friend Class RealConfigShim
#Else
<CodeAnalysis.ExcludeFromCodeCoverage>
Friend Class RealConfigShim
#End If

    Implements IConfigShim

    ''' <summary>
    ''' Gets the application setting.
    ''' </summary>
    ''' <param name="key">The key.</param>
    ''' <returns></returns>
    Public Function GetAppSetting(key As String) As String Implements IConfigShim.GetAppSetting

        Return Configuration.ConfigurationManager.AppSettings(key)

    End Function

End Class