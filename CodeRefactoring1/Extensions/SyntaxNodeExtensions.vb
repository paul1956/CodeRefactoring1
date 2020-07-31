﻿' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices

Public Module SyntaxNodeExtensions

    <Extension>
    Public Function GetAncestorOrThis(Of TNode As SyntaxNode)(node As SyntaxNode) As TNode
        If node Is Nothing Then
            Return Nothing
        End If

        Return node.GetAncestorsOrThis(Of TNode)().FirstOrDefault()
    End Function

    <Extension>
    Public Iterator Function GetAncestorsOrThis(Of TNode As SyntaxNode)(node As SyntaxNode) As IEnumerable(Of TNode)
        Dim current As SyntaxNode = node
        While current IsNot Nothing
            If TypeOf current Is TNode Then
                Yield DirectCast(current, TNode)
            End If

            current = If(TypeOf current Is IStructuredTriviaSyntax, DirectCast(current, IStructuredTriviaSyntax).ParentTrivia.Token.Parent, current.Parent)
        End While
    End Function

    <Extension>
    Public Function IsKind(token As SyntaxToken, kind1 As VisualBasic.SyntaxKind, kind2 As VisualBasic.SyntaxKind) As Boolean
        Return Kind(token) = kind1 OrElse Kind(token) = kind2
    End Function

    <Extension>
    Public Function IsParentKind(node As SyntaxNode, kind As SyntaxKind) As Boolean
        Return node IsNot Nothing AndAlso node.Parent.IsKind(kind)
    End Function

    <Extension>
    Public Function WithAppendedTrailingTrivia(Of T As SyntaxNode)(node As T, ParamArray trivia As SyntaxTrivia()) As T
        If trivia.Length = 0 Then
            Return node
        End If

        Return node.WithAppendedTrailingTrivia(DirectCast(trivia, IEnumerable(Of SyntaxTrivia)))
    End Function

    <Extension>
    Public Function WithAppendedTrailingTrivia(Of T As SyntaxNode)(node As T, trivia As SyntaxTriviaList) As T
        If trivia.Count = 0 Then
            Return node
        End If

        Return node.WithTrailingTrivia(node.GetTrailingTrivia().Concat(trivia))
    End Function

    <Extension>
    Public Function WithAppendedTrailingTrivia(Of T As SyntaxNode)(node As T, trivia As IEnumerable(Of SyntaxTrivia)) As T
        Return node.WithAppendedTrailingTrivia(trivia.ToSyntaxTriviaList())
    End Function

End Module
