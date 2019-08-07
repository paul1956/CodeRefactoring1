﻿Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On
Imports System.Runtime.CompilerServices
' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.


Namespace Roslyn.UnitTestFramework
    Friend Module EnumerableExtensions
        <Extension>
        Public Function OrderBy(Of T)(ByVal source As IEnumerable(Of T), ByVal comparer As IComparer(Of T)) As IEnumerable(Of T)
            Return source.OrderBy(Function(t_) t_, comparer)
        End Function

        <Extension>
        Public Function OrderBy(Of T)(ByVal source As IEnumerable(Of T), ByVal compare As Comparison(Of T)) As IEnumerable(Of T)
            Return source.OrderBy(New ComparisonComparer(Of T)(compare))
        End Function
        Private Class ComparisonComparer(Of T)
            Inherits Comparer(Of T)

            Private ReadOnly _compare As Comparison(Of T)

            Public Sub New(ByVal compare As Comparison(Of T))
                Me._compare = compare
            End Sub

            Public Overrides Function Compare(ByVal x As T, ByVal y As T) As Integer
                Return Me._compare(x, y)
            End Function
        End Class
    End Module

End Namespace