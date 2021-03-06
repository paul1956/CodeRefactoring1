﻿' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions

Namespace Style

    Partial Public Class AddAsClauseCodeFixProvider

        Private Class AddAsClauseCodeFixProviderCodeAction
            Inherits CodeAction

            Private ReadOnly _document As Document
            Private ReadOnly _newNode As SyntaxNode
            Private ReadOnly _node As SyntaxNode
            Private ReadOnly _title As String
            Private ReadOnly _equivalenceKey As String

            Public Sub New(key As String, document As Document, node As SyntaxNode, newNode As SyntaxNode, UniqueID As String)
                _document = document
                _newNode = newNode
                _node = node
                _title = key
                _equivalenceKey = $"{key}{UniqueID}"
            End Sub

            Public Overrides ReadOnly Property Title As String
                Get
                    Return _title
                End Get
            End Property

            Public Overrides ReadOnly Property EquivalenceKey As String
                Get
                    Return _equivalenceKey
                End Get
            End Property

            Protected Overrides Async Function GetChangedDocumentAsync(cancellationToken As CancellationToken) As Task(Of Document)
                Dim root As SyntaxNode = Await _document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
                Dim updatedRoot As SyntaxNode = root.ReplaceNode(_node, _newNode)
                Return _document.WithSyntaxRoot(updatedRoot)
            End Function

        End Class

    End Class

End Namespace
