Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Runtime.CompilerServices

Public Module KindExtensions
    <Extension()>
    Public Function MatchesKind(ByVal Expression As ExpressionSyntax, ParamArray KindList() As SyntaxKind) As Boolean
        For Each kind As SyntaxKind In KindList
            If Expression.Kind = kind Then
                Return True
            End If
        Next
        Return False
    End Function
End Module
