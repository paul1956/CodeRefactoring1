﻿' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

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
