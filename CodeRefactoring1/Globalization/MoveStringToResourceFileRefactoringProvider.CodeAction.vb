﻿' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Diagnostics.Debug

Namespace Globalization
    Partial Public Class MoveStringToResourceFileRefactoringProvider
        Private Class MoveStringToResourceFileCodeAction
            Inherits CodeAction

            Private ReadOnly _title As String
            Private ReadOnly _generateDocument As Func(Of CancellationToken, Task(Of Document))

            Public Sub New(title As String, generateDocument As Func(Of CancellationToken, Task(Of Document)))
                _title = title
                _generateDocument = generateDocument
            End Sub

            Public Overrides ReadOnly Property Title As String
                Get
                    Return _title
                End Get
            End Property

            Protected Overrides Function GetChangedDocumentAsync(cancellationToken As CancellationToken) As Task(Of Document)
                Return _generateDocument(cancellationToken)
            End Function

        End Class

    End Class

End Namespace
