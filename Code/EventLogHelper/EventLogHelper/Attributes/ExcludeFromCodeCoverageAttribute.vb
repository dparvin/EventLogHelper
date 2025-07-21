#If NET35 Then

Namespace Diagnostics

    ''' <summary>
    ''' Indicates that a class, struct, method, constructor, or property should be excluded from code coverage analysis.
    ''' </summary>
    ''' <seealso cref="System.Attribute" />
    <AttributeUsage(AttributeTargets.Class Or AttributeTargets.Struct Or AttributeTargets.Method Or AttributeTargets.Constructor Or AttributeTargets.Property, Inherited:=False)>
    Public NotInheritable Class ExcludeFromCodeCoverageAttribute

        Inherits Attribute

    End Class

End Namespace

#End If
