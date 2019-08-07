Imports System.Runtime.CompilerServices

Friend Module TypeExtensions

    <Extension>
    Public Function GetNullableUnderlyingType(ByVal type As ITypeSymbol) As ITypeSymbol
        If Not IsNullableType(type) Then
            Return Nothing
        End If
        Return DirectCast(type, INamedTypeSymbol).TypeArguments(0)
    End Function

    <Extension>
    Public Function IsNullableType(ByVal type As ITypeSymbol) As Boolean
        Dim original As ITypeSymbol = type.OriginalDefinition
        Return original.SpecialType = SpecialType.System_Nullable_T
    End Function

End Module