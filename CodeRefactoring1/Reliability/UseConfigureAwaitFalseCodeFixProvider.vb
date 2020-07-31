' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Formatting

Namespace Reliability

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(UseConfigureAwaitFalseCodeFixProvider)), Composition.Shared>
    Public Class UseConfigureAwaitFalseCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(UseConfigureAwaitFalseDiagnosticId)

        Private Shared Async Function CreateUseConfigureAwaitAsync(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim awaitExpression As AwaitExpressionSyntax = root.FindNode(diagnostic.Location.SourceSpan).ChildNodes.OfType(Of AwaitExpressionSyntax).FirstOrDefault()
            If awaitExpression Is Nothing Then Return document

            Dim newExpression As InvocationExpressionSyntax = SyntaxFactory.InvocationExpression(
                SyntaxFactory.MemberAccessExpression(
                    SyntaxKind.SimpleMemberAccessExpression,
                    awaitExpression.Expression.WithoutTrailingTrivia(),
                    SyntaxFactory.Token(SyntaxKind.DotToken),
                    SyntaxFactory.IdentifierName("ConfigureAwait")),
                SyntaxFactory.ArgumentList().
                    AddArguments(SyntaxFactory.SimpleArgument(SyntaxFactory.FalseLiteralExpression(SyntaxFactory.Token(SyntaxKind.FalseKeyword))
                ))).
                WithTrailingTrivia(awaitExpression.Expression.GetTrailingTrivia()).
                WithAdditionalAnnotations(Formatter.Annotation)
            Dim newRoot As SyntaxNode = root.ReplaceNode(awaitExpression.Expression, newExpression)
            Dim newDocument As Document = document.WithSyntaxRoot(newRoot)
            Return newDocument
        End Function

        Public NotOverridable Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public NotOverridable Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First()
            context.RegisterCodeFix(CodeAction.Create("Use ConfigureAwait(False)", Function(c As CancellationToken) CreateUseConfigureAwaitAsync(context.Document, diagnostic, c), NameOf(UseConfigureAwaitFalseCodeFixProvider)), diagnostic)
            Return Task.FromResult(0)
        End Function

    End Class

End Namespace
