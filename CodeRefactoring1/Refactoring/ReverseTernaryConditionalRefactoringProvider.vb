' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

<ExportCodeRefactoringProvider(LanguageNames.VisualBasic, Name:=NameOf(ReverseTernaryConditionalRefactoringProvider)), [Shared]>
Friend Class ReverseTernaryConditionalRefactoringProvider
    Inherits CodeRefactoringProvider

    Public NotOverridable Overrides Async Function ComputeRefactoringsAsync(context As CodeRefactoringContext) As Task
        Dim semanticModel As SemanticModel = Await context.Document.GetSemanticModelAsync(context.CancellationToken)
        Dim root As SyntaxNode = Await context.Document.GetSyntaxRootAsync(context.CancellationToken)
        Dim invocation As TernaryConditionalExpressionSyntax = root.FindNode(context.Span, getInnermostNodeForTie:=True)?.FirstAncestorOrSelf(Of TernaryConditionalExpressionSyntax)()

        If invocation Is Nothing OrElse invocation.Kind <> SyntaxKind.TernaryConditionalExpression Then
            Exit Function
        End If
        If invocation.Condition.Kind = SyntaxKind.IsExpression OrElse invocation.Condition.Kind = SyntaxKind.IsNotExpression Then
            context.RegisterRefactoring(New ReverseIfOperatorCodeAction("Reverse Is/IsNot Nothing, If Operator",
                                                Function(c) CreateReverseTernaryOperatorAsync(invocation, context.Document, c)))
        Else
            Exit Function
        End If

        Dim Condition As BinaryExpressionSyntax = CType(invocation.Condition, BinaryExpressionSyntax)
        If Condition?.Right?.Kind = SyntaxKind.NothingLiteralExpression Then
            If invocation.WhenTrue.IsKind(SyntaxKind.SimpleMemberAccessExpression) Then
                Dim _MemberAccessExpressionSyntax As MemberAccessExpressionSyntax = CType(invocation.WhenTrue, MemberAccessExpressionSyntax)
                Dim MemberAccessText() As String = _MemberAccessExpressionSyntax.ToString.Split("."c)
                If MemberAccessText.Length <> 2 Then
                    Exit Function
                End If
            Else
                Exit Function
            End If
            context.RegisterRefactoring(New ReverseIfOperatorCodeAction("Simplify Is/IsNot Nothing, to Access Expression",
                                            Function(c) CreateConditionalAccessExpressionAsync(invocation, context.Document, c)))
        Else
            If Condition?.Left?.Kind <> SyntaxKind.NothingLiteralExpression Then
                Exit Function
            End If
            Dim MemberAccessExpressionSyntax As MemberAccessExpressionSyntax = CType(invocation.WhenFalse, MemberAccessExpressionSyntax)
            Dim MemberAccessText() As String = MemberAccessExpressionSyntax.ToString.Split(CType(".", Char()))
            If MemberAccessText.Length <> 2 Then
                Return
            End If
            context.RegisterRefactoring(New ReverseIfOperatorCodeAction("Simplify Is/IsNot Nothing, to Access Expression",
                                        Function(c) CreateConditionalAccessExpressionAsync(invocation, context.Document, c)))
        End If
    End Function

    Private Async Function CreateReverseTernaryOperatorAsync(ByVal invocation As TernaryConditionalExpressionSyntax, ByVal document As Document, ByVal cancellationToken As CancellationToken) As Task(Of Document)
        Try
            Dim semanticModel As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken)
            Dim NewCondition As BinaryExpressionSyntax
            Dim OldCondition As BinaryExpressionSyntax = CType(invocation.Condition, BinaryExpressionSyntax)
            Dim operatorToken As SyntaxToken
            If invocation.Condition.Kind = SyntaxKind.IsNotExpression Then
                operatorToken = SyntaxFactory.Token(SyntaxKind.IsKeyword).WithTriviaFrom(OldCondition.OperatorToken).WithoutAnnotations
                NewCondition = SyntaxFactory.BinaryExpression(SyntaxKind.IsExpression, OldCondition.Left, operatorToken, OldCondition.Right).WithTriviaFrom(OldCondition)
            Else
                operatorToken = SyntaxFactory.Token(SyntaxKind.IsNotKeyword).WithTriviaFrom(OldCondition.OperatorToken).WithoutAnnotations
                NewCondition = SyntaxFactory.BinaryExpression(SyntaxKind.IsNotExpression, OldCondition.Left, operatorToken, OldCondition.Right).WithTriviaFrom(OldCondition)
            End If
            Dim newRoot As SyntaxNode = root.ReplaceNode(invocation, SyntaxFactory.TernaryConditionalExpression(SyntaxFactory.Token(SyntaxKind.IfKeyword), SyntaxFactory.Token(SyntaxKind.OpenParenToken), NewCondition, SyntaxFactory.Token(SyntaxKind.CommaToken), invocation.WhenFalse, SyntaxFactory.Token(SyntaxKind.CommaToken), invocation.WhenTrue, SyntaxFactory.Token(SyntaxKind.CloseParenToken)).WithTriviaFrom(invocation))
            Return document.WithSyntaxRoot(newRoot)
        Catch ex As Exception
            Throw ex
        End Try
        Return Nothing
    End Function

    Private Async Function CreateConditionalAccessExpressionAsync(ByVal invocation As TernaryConditionalExpressionSyntax, ByVal document As Document, ByVal cancellationToken As CancellationToken) As Task(Of Document)
        Try
            Dim semanticModel As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken)
            Dim newRoot As SyntaxNode
            If invocation.Condition.Kind = SyntaxKind.IsExpression Then
                ' invocation.WhenFalse has stuff I want, but need to split it
                Dim MemberAccessExpressionSyntax As MemberAccessExpressionSyntax = CType(invocation.WhenFalse, MemberAccessExpressionSyntax)
                Dim MemberAccessText() As String = MemberAccessExpressionSyntax.ToString.Split(CType(".", Char()))
                newRoot = root.ReplaceNode(invocation, SyntaxFactory.ConditionalAccessExpression(SyntaxFactory.IdentifierName(MemberAccessText(0)), SyntaxFactory.Token(SyntaxKind.QuestionToken), SyntaxFactory.SimpleMemberAccessExpression(SyntaxFactory.IdentifierName(MemberAccessText(1)))).WithTriviaFrom(invocation))
            Else
                ' invocation.WhenTrue has stuff I want, but need to split it.
                Dim MemberAccessExpressionSyntax As MemberAccessExpressionSyntax = CType(invocation.WhenTrue, MemberAccessExpressionSyntax)
                Dim MemberAccessText() As String = MemberAccessExpressionSyntax.ToString.Split(CType(".", Char()))
                newRoot = root.ReplaceNode(invocation, SyntaxFactory.ConditionalAccessExpression(SyntaxFactory.IdentifierName(MemberAccessText(0)), SyntaxFactory.Token(SyntaxKind.QuestionToken), SyntaxFactory.SimpleMemberAccessExpression(SyntaxFactory.IdentifierName(MemberAccessText(1)))).WithTriviaFrom(invocation))
            End If
            Return document.WithSyntaxRoot(newRoot)
        Catch ex As Exception
            Throw ex
        End Try
        Return Nothing
    End Function

    Private Class ReverseIfOperatorCodeAction
        Inherits CodeAction

        Private ReadOnly _generateDocument As Func(Of CancellationToken, Task(Of Document))
        Private ReadOnly _title As String

        Public Overrides ReadOnly Property Title As String
            Get
                Return _title
            End Get
        End Property

        Public Sub New(title As String, generateDocument As Func(Of CancellationToken, Task(Of Document)))
            _title = title
            _generateDocument = generateDocument
        End Sub

        Protected Overrides Function GetChangedDocumentAsync(cancellationToken As CancellationToken) As Task(Of Document)
            Return _generateDocument(cancellationToken)
        End Function

    End Class

End Class
