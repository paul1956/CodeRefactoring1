' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.CodeFixes

Namespace Usage

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(RemovePrivateMethodNeverUsedCodeFixProvider)), Composition.Shared>
    Public Class RemovePrivateMethodNeverUsedCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(RemovePrivateMethodNeverUsedDiagnosticId)

        Private Async Function RemoveMethodAsync(document As Document, diagnostic As Diagnostic, cancellationToken As Threading.CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim span As TextSpan = diagnostic.Location.SourceSpan
            Dim methodNotUsed As MethodStatementSyntax = root.FindToken(span.Start).Parent.FirstAncestorOrSelf(Of MethodStatementSyntax)

            Dim newRoot As SyntaxNode = root.RemoveNode(methodNotUsed.Parent, SyntaxRemoveOptions.KeepNoTrivia)
            Return document.WithSyntaxRoot(newRoot)
        End Function

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First()
            context.RegisterCodeFix(CodeAction.Create($"Remove unused private method: {diagnostic.Properties!identifier}", Function(c As CancellationToken) RemoveMethodAsync(context.Document, diagnostic, c), NameOf(RemovePrivateMethodNeverUsedCodeFixProvider)), diagnostic)
            Return Task.FromResult(0)
        End Function

    End Class

End Namespace
