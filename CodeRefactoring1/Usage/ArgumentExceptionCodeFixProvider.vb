' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.CodeFixes

Namespace Usage

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(ArgumentExceptionCodeFixProvider)), Composition.Shared>
    Public Class ArgumentExceptionCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(ArgumentExceptionDiagnosticId)

        Private Async Function FixParamAsync(document As Document, diagnostic As Diagnostic, newParamName As String, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim span As TextSpan = diagnostic.Location.SourceSpan
            Dim objectCreation As ObjectCreationExpressionSyntax = root.FindToken(span.Start).Parent.FirstAncestorOrSelf(Of ObjectCreationExpressionSyntax)

            Dim semanticModel As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken)

            Dim argumentList As ArgumentListSyntax = objectCreation.ArgumentList
            Dim paramNameLiteral As LiteralExpressionSyntax = DirectCast(argumentList.Arguments(1).GetExpression, LiteralExpressionSyntax)
            Dim paramNameOpt As [Optional](Of Object) = semanticModel.GetConstantValue(paramNameLiteral)
            Dim currentParamName As String = paramNameOpt.Value.ToString()

            Dim newLiteral As LiteralExpressionSyntax = SyntaxFactory.LiteralExpression(SyntaxKind.StringLiteralExpression, SyntaxFactory.Literal(newParamName))
            Dim newRoot As SyntaxNode = root.ReplaceNode(paramNameLiteral, newLiteral)
            Dim newDocument As Document = document.WithSyntaxRoot(newRoot)
            Return newDocument
        End Function

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First()

            Dim parameters As IEnumerable(Of KeyValuePair(Of String, String)) = diagnostic.Properties.Where(Function(p) p.Key.StartsWith("param"))
            For Each param As KeyValuePair(Of String, String) In parameters
                Dim message As String = $"Use '{param}'"
                context.RegisterCodeFix(CodeAction.Create(message, Function(c) FixParamAsync(context.Document, diagnostic, param.Value, c), NameOf(ArgumentExceptionCodeFixProvider)), diagnostic)
            Next
            Return Task.FromResult(0)
        End Function

    End Class

End Namespace
