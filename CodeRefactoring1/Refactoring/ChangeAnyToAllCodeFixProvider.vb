Imports Microsoft.CodeAnalysis.CodeFixes
Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Formatting

Namespace Refactoring
    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(ChangeAnyToAllCodeFixProvider)), [Shared]>
    Public Class ChangeAnyToAllCodeFixProvider
        Inherits CodeFixProvider

        Public Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String)
            Get
                Return ImmutableArray.Create(DiagnosticIds.ChangeAnyToAllDiagnosticId,
                                             DiagnosticIds.ChangeAllToAnyDiagnosticId)
            End Get
        End Property

        Public NotOverridable Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diag As Diagnostic = context.Diagnostics.First
            Dim message As String = If(diag.Id = DiagnosticIds.ChangeAnyToAllDiagnosticId, "Change Any to All", "Change All to Any")
            context.RegisterCodeFix(CodeAction.Create(message, Function(c As CancellationToken) ConvertAsync(context.Document, diag.Location, c), NameOf(ChangeAnyToAllCodeFixProvider)), diag)
            Return Task.FromResult(0)
        End Function

        Private Shared Async Function ConvertAsync(Document As Document, diagnosticLocation As Location, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await Document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim invocation As InvocationExpressionSyntax = root.FindNode(diagnosticLocation.SourceSpan).FirstAncestorOfType(Of InvocationExpressionSyntax)
            Dim newInvocation As ExpressionSyntax = CreateNewInvocation(invocation).
                WithAdditionalAnnotations(Formatter.Annotation)
            Dim newRoot As SyntaxNode = ReplaceInvocation(invocation, newInvocation, root)
            Dim newDocument As Document = Document.WithSyntaxRoot(newRoot)
            Return newDocument
        End Function

        Private Shared Function ReplaceInvocation(invocation As InvocationExpressionSyntax, newInvocation As ExpressionSyntax, root As SyntaxNode) As SyntaxNode
            If invocation.Parent.IsKind(SyntaxKind.NotExpression) Then
                Return root.ReplaceNode(invocation.Parent, newInvocation)
            End If
            Dim negatedInvocation As UnaryExpressionSyntax = SyntaxFactory.NotExpression(newInvocation)
            Dim newRoot As SyntaxNode = root.ReplaceNode(invocation, negatedInvocation)
            Return newRoot
        End Function

        Friend Shared Function CreateNewInvocation(invocation As InvocationExpressionSyntax) As ExpressionSyntax
            Dim methodName As String = DirectCast(invocation.Expression, MemberAccessExpressionSyntax).Name.ToString
            Dim nameToCheck As IdentifierNameSyntax = If(methodName = NameOf(Enumerable.Any), ChangeAnyToAllAnalyzer.AllName, ChangeAnyToAllAnalyzer.AnyName)
            Dim newInvocation As InvocationExpressionSyntax = invocation.WithExpression(DirectCast(invocation.Expression, MemberAccessExpressionSyntax).WithName(nameToCheck))
            Dim comparisonExpression As ExpressionSyntax = DirectCast(DirectCast(newInvocation.ArgumentList.Arguments.First().GetExpression(), SingleLineLambdaExpressionSyntax).Body, ExpressionSyntax)
            Dim newComparisonExpression As ExpressionSyntax = CreateNewComparison(comparisonExpression)
            newComparisonExpression = RemoveParenthesis(newComparisonExpression)
            newInvocation = newInvocation.ReplaceNode(comparisonExpression, newComparisonExpression)
            Return newInvocation
        End Function

        Private Shared Function CreateNewComparison(comparisonExpression As ExpressionSyntax) As ExpressionSyntax
            If comparisonExpression.IsKind(SyntaxKind.TernaryConditionalExpression) Then
                Return SyntaxFactory.EqualsExpression(comparisonExpression,
                                                      SyntaxFactory.FalseLiteralExpression(SyntaxFactory.Token(SyntaxKind.FalseKeyword)))
            End If
            If comparisonExpression.IsKind(SyntaxKind.NotExpression) Then
                Return DirectCast(comparisonExpression, UnaryExpressionSyntax).Operand
            End If

            If comparisonExpression.IsKind(SyntaxKind.EqualsExpression) Then
                Dim binaryComparison As BinaryExpressionSyntax = DirectCast(comparisonExpression, BinaryExpressionSyntax)
                If binaryComparison.Right.IsKind(SyntaxKind.TrueLiteralExpression) Then
                    Return binaryComparison.WithRight(SyntaxFactory.FalseLiteralExpression(SyntaxFactory.Token(SyntaxKind.FalseKeyword)))
                End If
                If binaryComparison.Left.IsKind(SyntaxKind.TrueLiteralExpression) Then
                    Return binaryComparison.WithLeft(SyntaxFactory.FalseLiteralExpression(SyntaxFactory.Token(SyntaxKind.FalseKeyword)))
                End If
                If binaryComparison.Right.IsKind(SyntaxKind.FalseLiteralExpression) Then
                    Return binaryComparison.Left
                End If
                If binaryComparison.Left.IsKind(SyntaxKind.FalseLiteralExpression) Then
                    Return binaryComparison.Right
                End If
                Return CreateNewBinaryExpression(comparisonExpression, SyntaxKind.NotEqualsExpression, SyntaxFactory.Token(SyntaxKind.LessThanGreaterThanToken))
            End If

            If comparisonExpression.IsKind(SyntaxKind.NotEqualsExpression) Then
                Dim binaryComparison As BinaryExpressionSyntax = DirectCast(comparisonExpression, BinaryExpressionSyntax)
                If binaryComparison.Right.IsKind(SyntaxKind.TrueLiteralExpression) Then
                    Return binaryComparison.Left
                End If
                If binaryComparison.Left.IsKind(SyntaxKind.TrueLiteralExpression) Then
                    Return binaryComparison.Right
                End If
                Return CreateNewBinaryExpression(comparisonExpression, SyntaxKind.EqualsExpression, SyntaxFactory.Token(SyntaxKind.EqualsToken))
            End If
            If comparisonExpression.IsKind(SyntaxKind.GreaterThanExpression) Then
                Return CreateNewBinaryExpression(comparisonExpression, SyntaxKind.LessThanOrEqualExpression, SyntaxFactory.Token(SyntaxKind.LessThanEqualsToken))
            End If
            If comparisonExpression.IsKind(SyntaxKind.GreaterThanOrEqualExpression) Then
                Return CreateNewBinaryExpression(comparisonExpression, SyntaxKind.LessThanExpression, SyntaxFactory.Token(SyntaxKind.LessThanToken))
            End If
            If comparisonExpression.IsKind(SyntaxKind.LessThanExpression) Then
                Return CreateNewBinaryExpression(comparisonExpression, SyntaxKind.GreaterThanOrEqualExpression, SyntaxFactory.Token(SyntaxKind.GreaterThanEqualsToken))
            End If
            If comparisonExpression.IsKind(SyntaxKind.LessThanOrEqualExpression) Then
                Return CreateNewBinaryExpression(comparisonExpression, SyntaxKind.GreaterThanExpression, SyntaxFactory.Token(SyntaxKind.GreaterThanToken))
            End If
            If comparisonExpression.IsKind(SyntaxKind.TrueLiteralExpression) Then
                Return SyntaxFactory.FalseLiteralExpression(SyntaxFactory.Token(SyntaxKind.FalseKeyword))
            End If
            If comparisonExpression.IsKind(SyntaxKind.FalseLiteralExpression) Then
                Return SyntaxFactory.TrueLiteralExpression(SyntaxFactory.Token(SyntaxKind.TrueKeyword))
            End If
            Return SyntaxFactory.EqualsExpression(comparisonExpression,
                                                  SyntaxFactory.FalseLiteralExpression(SyntaxFactory.Token(SyntaxKind.FalseKeyword)))
        End Function

        Private Shared Function RemoveParenthesis(expression As ExpressionSyntax) As ExpressionSyntax
            Return If(expression.IsKind(SyntaxKind.ParenthesizedExpression),
                DirectCast(expression, ParenthesizedExpressionSyntax).Expression,
                expression)
        End Function

        Private Shared Function CreateNewBinaryExpression(comparisonExpression As ExpressionSyntax, kind As SyntaxKind, operatorToken As SyntaxToken) As BinaryExpressionSyntax
            Dim binaryComparison As BinaryExpressionSyntax = DirectCast(comparisonExpression, BinaryExpressionSyntax)
            Dim left As ExpressionSyntax = binaryComparison.Left
            Dim newComparison As BinaryExpressionSyntax = SyntaxFactory.BinaryExpression(
                kind,
                If(left.IsKind(SyntaxKind.BinaryConditionalExpression), SyntaxFactory.ParenthesizedExpression(left), left),
                operatorToken,
                binaryComparison.Right)
            Return newComparison
        End Function
    End Class
End Namespace
