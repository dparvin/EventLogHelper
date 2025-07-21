#If NET461_OR_GREATER Or NET5_0_OR_GREATER Then
Imports XmlDocMarkdown.Core
#Else
Imports System
Imports System.Net.WebRequestMethods
#End If

Module Program

    Sub Main(args As String())

#If NET461_OR_GREATER Or NET5_0_OR_GREATER Then
        XmlDocMarkdownApp.Run(args)
        Dim ud As New UpdateDocumentation(args)
        ud.UpdateFiles()
#Else
        throw new NotImplementedException("This is a stub for the PropertyGridHelpers project.")
#End If

    End Sub

End Module
