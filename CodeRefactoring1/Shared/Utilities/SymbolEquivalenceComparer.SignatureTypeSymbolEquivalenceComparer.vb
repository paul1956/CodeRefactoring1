'Namespace Microsoft.CodeAnalysis.Shared.Utilities
Partial Friend Class SymbolEquivalenceComparer
    Friend Class SignatureTypeSymbolEquivalenceComparer
        Implements IEqualityComparer(Of ITypeSymbol)

        Private ReadOnly _symbolEquivalenceComparer As SymbolEquivalenceComparer

        Public Sub New(ByVal symbolEquivalenceComparer As SymbolEquivalenceComparer)
            Me._symbolEquivalenceComparer = symbolEquivalenceComparer
        End Sub

        Public Shadows Function Equals(ByVal x As ITypeSymbol, ByVal y As ITypeSymbol) As Boolean Implements IEqualityComparer(Of ITypeSymbol).Equals
            Return Me.Equals(x, y, Nothing)
        End Function

        Public Shadows Function Equals(ByVal x As ITypeSymbol, ByVal y As ITypeSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
            Return Me._symbolEquivalenceComparer.GetEquivalenceVisitor(compareMethodTypeParametersByIndex:=True, objectAndDynamicCompareEqually:=True).AreEquivalent(x, y, equivalentTypesWithDifferingAssemblies)
        End Function

        Public Shadows Function GetHashCode(ByVal x As ITypeSymbol) As Integer Implements IEqualityComparer(Of ITypeSymbol).GetHashCode
            Return Me._symbolEquivalenceComparer.GetGetHashCodeVisitor(compareMethodTypeParametersByIndex:=True, objectAndDynamicCompareEqually:=True).GetHashCode(x, currentHash:=0)
        End Function
    End Class
End Class
'End Namespace
