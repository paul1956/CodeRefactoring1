Option Explicit On
Option Infer Off
Option Strict On

Imports System.Runtime.CompilerServices

Public Module ITypeSymbolExtensions

    <Extension()>
    Friend Function IsEnumType(type As ITypeSymbol) As Boolean
        Debug.Assert(type IsNot Nothing)
        Return type.TypeKind = TypeKind.Enum
    End Function

    <Extension>
    Public Iterator Function GetBaseTypesAndThis(ByVal type As ITypeSymbol) As IEnumerable(Of ITypeSymbol)
        Dim current As ITypeSymbol = type
        Do While current IsNot Nothing
            Yield current
            current = current.BaseType
        Loop
    End Function

    ' Determine if "type" inherits from "baseType", ignoring constructed types, and dealing
    ' only with original types.
    <Extension>
    Public Function InheritsFromOrEqualsIgnoringConstruction(ByVal type As ITypeSymbol, ByVal baseType As ITypeSymbol) As Boolean
        Dim originalBaseType As ITypeSymbol = baseType.OriginalDefinition
        Return type.GetBaseTypesAndThis().Contains(Function(t As ITypeSymbol) SymbolEquivalenceComparer.Instance.Equals(t.OriginalDefinition, originalBaseType))
    End Function

    <Extension>
    Public Function IsErrorType(ByVal symbol As ITypeSymbol) As Boolean
        Return CBool(symbol?.TypeKind = TypeKind.Error)
    End Function

End Module