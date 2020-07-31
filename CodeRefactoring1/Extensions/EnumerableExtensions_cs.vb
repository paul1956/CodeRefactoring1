' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices

Public Module EnumerableExtensions

    <Extension>
    Public Function [Do](Of T)(ByVal source As IEnumerable(Of T), ByVal action As Action(Of T)) As IEnumerable(Of T)
        If source Is Nothing Then
            Throw New ArgumentNullException(NameOf(source))
        End If

        If action Is Nothing Then
            Throw New ArgumentNullException(NameOf(action))
        End If

        ' perf optimization. try to not use enumerator if possible
        Dim list As IList(Of T) = TryCast(source, IList(Of T))
        If list IsNot Nothing Then
            Dim i As Integer = 0
            Dim count As Integer = list.Count
            Do While i < count
                action(list(i))
                i += 1
            Loop
        Else
            For Each value As T In source
                action(value)
            Next value
        End If

        Return source
    End Function

    <Extension>
    Public Function Contains(Of T)(ByVal sequence As IEnumerable(Of T), ByVal predicate As Func(Of T, Boolean)) As Boolean
        Return sequence.Any(predicate)
    End Function

End Module
