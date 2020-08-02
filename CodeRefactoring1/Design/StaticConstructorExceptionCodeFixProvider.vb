' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Design

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(StaticConstructorExceptionCodeFixProvider)), [Shared]>
    Public Class StaticConstructorExceptionCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(StaticConstructorExceptionAnalyzer.Id)

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First
            context.RegisterCodeFix(CodeAction.Create("Remove this exception", Function(ct As CancellationToken) Me.RemoveThrow(context.Document, diagnostic, ct), NameOf(StaticConstructorExceptionCodeFixProvider)), diagnostic)
            Return Task.FromResult(0)
        End Function

        Private Async Function RemoveThrow(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim sourceSpan As TextSpan = diagnostic.Location.SourceSpan
            Dim throwBlock As ThrowStatementSyntax = root.FindToken(sourceSpan.Start).Parent.AncestorsAndSelf.OfType(Of ThrowStatementSyntax).First

            Return document.WithSyntaxRoot((Await document.GetSyntaxRootAsync(cancellationToken)).RemoveNode(throwBlock, SyntaxRemoveOptions.KeepNoTrivia))
        End Function

    End Class

End Namespace
