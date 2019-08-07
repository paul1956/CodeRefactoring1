Option Infer On
Imports Utilities

' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'Namespace Utilities
Partial Friend Class SymbolEquivalenceComparer
    Friend Class ParameterSymbolEqualityComparer
        Implements IEqualityComparer(Of IParameterSymbol)

        Private ReadOnly _symbolEqualityComparer As SymbolEquivalenceComparer
        Private ReadOnly _distinguishRefFromOut As Boolean

        Public Sub New(ByVal symbolEqualityComparer As SymbolEquivalenceComparer, ByVal distinguishRefFromOut As Boolean)
            Me._symbolEqualityComparer = symbolEqualityComparer
            Me._distinguishRefFromOut = distinguishRefFromOut
        End Sub

        Public Shadows Function Equals(ByVal x As IParameterSymbol, ByVal y As IParameterSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol), ByVal compareParameterName As Boolean, ByVal isCaseSensitive As Boolean) As Boolean
            If ReferenceEquals(x, y) Then
                Return True
            End If

            If x Is Nothing OrElse y Is Nothing Then
                Return False
            End If

            Dim nameComparisonCheck As Boolean = True
            If compareParameterName Then
                nameComparisonCheck = If(isCaseSensitive, x.Name = y.Name, String.Equals(x.Name, y.Name, StringComparison.OrdinalIgnoreCase))
            End If

            ' See the comment in the outer type.  If we're comparing two parameters for
            ' equality, then we want to consider method type parameters by index only.

            Return AreRefKindsEquivalent(x.RefKind, y.RefKind, Me._distinguishRefFromOut) AndAlso nameComparisonCheck AndAlso Me._symbolEqualityComparer.GetEquivalenceVisitor().AreEquivalent(x.CustomModifiers, y.CustomModifiers, equivalentTypesWithDifferingAssemblies) AndAlso Me._symbolEqualityComparer.SignatureTypeEquivalenceComparer.Equals(x.Type, y.Type, equivalentTypesWithDifferingAssemblies)
        End Function

        Public Shadows Function Equals(ByVal x As IParameterSymbol, ByVal y As IParameterSymbol) As Boolean Implements IEqualityComparer(Of IParameterSymbol).Equals
            Return Me.Equals(x, y, Nothing, False, False)
        End Function

        Public Shadows Function Equals(ByVal x As IParameterSymbol, ByVal y As IParameterSymbol, ByVal compareParameterName As Boolean, ByVal isCaseSensitive As Boolean) As Boolean
            Return Me.Equals(x, y, Nothing, compareParameterName, isCaseSensitive)
        End Function

        Public Shadows Function GetHashCode(ByVal x As IParameterSymbol) As Integer Implements IEqualityComparer(Of IParameterSymbol).GetHashCode
            If x Is Nothing Then
                Return 0
            End If

            Return CodeRefactoringHash.Combine(x.IsRefOrOut(), Me._symbolEqualityComparer.SignatureTypeEquivalenceComparer.GetHashCode(x.Type))
        End Function
    End Class

    Public Shared Function AreRefKindsEquivalent(ByVal rk1 As RefKind, ByVal rk2 As RefKind, ByVal distinguishRefFromOut As Boolean) As Boolean
        Return If(distinguishRefFromOut, rk1 = rk2, (rk1 = RefKind.None) = (rk2 = RefKind.None))
    End Function
End Class
'End Namespace
