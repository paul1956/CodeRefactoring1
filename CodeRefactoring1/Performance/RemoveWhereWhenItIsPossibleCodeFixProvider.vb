﻿' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Performance

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(RemoveWhereWhenItIsPossibleCodeFixProvider)), [Shared]>
    Public Class RemoveWhereWhenItIsPossibleCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(RemoveWhereWhenItIsPossibleAnalyzer.Id)

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First
            Dim name As String = diagnostic.Properties!methodName
            Dim message As String = $"Remove 'Where' moving predicate to '{name}'"
            context.RegisterCodeFix(CodeAction.Create(message, Function(c As CancellationToken) Me.RemoveWhere(context.Document, diagnostic, c), NameOf(RemoveWhereWhenItIsPossibleCodeFixProvider)), diagnostic)
            Return Task.FromResult(0)
        End Function

        Private Async Function RemoveWhere(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim diagnosticSpan As TextSpan = diagnostic.Location.SourceSpan
            Dim whereInvoke As InvocationExpressionSyntax = root.FindToken(diagnosticSpan.Start).Parent.AncestorsAndSelf().OfType(Of InvocationExpressionSyntax)().First()
            Dim nextMethodInvoke As InvocationExpressionSyntax = whereInvoke.Parent.FirstAncestorOrSelf(Of InvocationExpressionSyntax)()

            Dim whereMemberAccess As MemberAccessExpressionSyntax = whereInvoke.ChildNodes.OfType(Of MemberAccessExpressionSyntax)().FirstOrDefault()
            Dim nextMethodMemberAccess As MemberAccessExpressionSyntax = nextMethodInvoke.ChildNodes.OfType(Of MemberAccessExpressionSyntax)().FirstOrDefault()

            ' We need to push the args into the next invoke's arg list instead of just replacing
            ' where with new method because next method's arg list's end paren may have the CRLF which is dropped otherwise.
            Dim whereArgs As ArgumentListSyntax = whereInvoke.ArgumentList
            Dim newArguments As ArgumentListSyntax = nextMethodInvoke.ArgumentList.WithArguments(whereArgs.Arguments)

            Dim newNextMethodInvoke As InvocationExpressionSyntax = SyntaxFactory.InvocationExpression(
            SyntaxFactory.MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression, whereMemberAccess.Expression, SyntaxFactory.Token(SyntaxKind.DotToken), nextMethodMemberAccess.Name), newArguments)

            Dim newRoot As SyntaxNode = root.ReplaceNode(nextMethodInvoke, newNextMethodInvoke)
            Dim newDocument As Document = document.WithSyntaxRoot(newRoot)
            Return newDocument
        End Function

    End Class

End Namespace
