Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Formatting

Namespace Style
    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(TernaryOperatorWithReturnCodeFixProvider)), Composition.Shared>
    Public Class TernaryOperatorWithReturnCodeFixProvider
        Inherits CodeFixProvider

        Public Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) =
            ImmutableArray.Create(DiagnosticIds.TernaryOperator_ReturnDiagnosticId)

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First
            context.RegisterCodeFix(CodeAction.Create("Change to ternary operator", Function(c As CancellationToken) Me.MakeTernaryAsync(context.Document, diagnostic, c), NameOf(TernaryOperatorWithReturnCodeFixProvider)), diagnostic)
            Return Task.FromResult(0)
        End Function

        Private Async Function MakeTernaryAsync(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim span As TextSpan = diagnostic.Location.SourceSpan
            Dim ifBlock As MultiLineIfBlockSyntax = root.FindToken(span.Start).Parent.FirstAncestorOrSelfOfType(Of MultiLineIfBlockSyntax)

            Dim ifReturn As ReturnStatementSyntax = TryCast(ifBlock.Statements.FirstOrDefault(), ReturnStatementSyntax)
            Dim elseReturn As ReturnStatementSyntax = TryCast(ifBlock.ElseBlock?.Statements.FirstOrDefault(), ReturnStatementSyntax)
            Dim semanticModel As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken)
            Dim type As ITypeSymbol = GetCommonBaseType(semanticModel.GetTypeInfo(ifReturn.Expression).ConvertedType, semanticModel.GetTypeInfo(elseReturn.Expression).ConvertedType)

            Dim ifType As ITypeSymbol = semanticModel.GetTypeInfo(ifReturn.Expression).Type
            Dim elseType As ITypeSymbol = semanticModel.GetTypeInfo(elseReturn.Expression).Type

            Dim typeSyntax As IdentifierNameSyntax = SyntaxFactory.IdentifierName(type.ToMinimalDisplayString(semanticModel, ifReturn.SpanStart))
            Dim trueExpression As ExpressionSyntax = ifReturn.Expression.
                ConvertToBaseType(ifType, type).
                EnsureNothingAsType(semanticModel, type, typeSyntax)

            Dim falseExpression As ExpressionSyntax = elseReturn.Expression.
                ConvertToBaseType(elseType, type).
                EnsureNothingAsType(semanticModel, type, typeSyntax)

            Dim leadingTrivia As SyntaxTriviaList = ifBlock.GetLeadingTrivia()
            leadingTrivia = leadingTrivia.InsertRange(leadingTrivia.Count - 1, ifReturn.GetLeadingTrivia())
            leadingTrivia = leadingTrivia.InsertRange(leadingTrivia.Count - 1, elseReturn.GetLeadingTrivia())

            Dim trailingTrivia As SyntaxTriviaList = ifBlock.GetTrailingTrivia.
                InsertRange(0, elseReturn.GetTrailingTrivia().Where(Function(trivia As SyntaxTrivia) Not trivia.IsKind(SyntaxKind.EndOfLineTrivia))).
                InsertRange(0, ifReturn.GetTrailingTrivia().Where(Function(trivia As SyntaxTrivia) Not trivia.IsKind(SyntaxKind.EndOfLineTrivia)))

            Dim ternary As TernaryConditionalExpressionSyntax = SyntaxFactory.TernaryConditionalExpression(ifBlock.IfStatement.Condition.WithoutTrailingTrivia(),
                                                                     trueExpression.WithoutTrailingTrivia(),
                                                                     falseExpression.WithoutTrailingTrivia())

            Dim returnStatement As ReturnStatementSyntax = SyntaxFactory.ReturnStatement(ternary).
                WithLeadingTrivia(leadingTrivia).
                WithTrailingTrivia(trailingTrivia).
                WithAdditionalAnnotations(Formatter.Annotation)

            Dim newRoot As SyntaxNode = root.ReplaceNode(ifBlock, returnStatement)
            Dim newDocument As Document = document.WithSyntaxRoot(newRoot)

            Return newDocument
        End Function
    End Class

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(TernaryOperatorWithAssignmentCodeFixProvider)), Composition.Shared>
    Public Class TernaryOperatorWithAssignmentCodeFixProvider
        Inherits CodeFixProvider

        Public Overrides ReadOnly Property FixableDiagnosticIds() As ImmutableArray(Of String) =
            ImmutableArray.Create(DiagnosticIds.TernaryOperator_AssignmentDiagnosticId)

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First
            context.RegisterCodeFix(CodeAction.Create("Change to ternary operator", Function(c As CancellationToken) Me.MakeTernaryAsync(context.Document, diagnostic, c), NameOf(TernaryOperatorWithAssignmentCodeFixProvider)), diagnostic)
            Return Task.FromResult(0)
        End Function

        Private Async Function MakeTernaryAsync(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim ifBlock As MultiLineIfBlockSyntax = root.FindToken(diagnostic.Location.SourceSpan.Start).Parent.FirstAncestorOrSelf(Of MultiLineIfBlockSyntax)

            Dim ifAssign As AssignmentStatementSyntax = TryCast(ifBlock.Statements.FirstOrDefault(), AssignmentStatementSyntax)
            Dim elseAssign As AssignmentStatementSyntax = TryCast(ifBlock.ElseBlock?.Statements.FirstOrDefault(), AssignmentStatementSyntax)
            Dim semanticModel As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken)
            Dim type As ITypeSymbol = GetCommonBaseType(semanticModel.GetTypeInfo(ifAssign.Left).ConvertedType, semanticModel.GetTypeInfo(elseAssign.Left).ConvertedType)
            Dim typeSyntax As IdentifierNameSyntax = SyntaxFactory.IdentifierName(type.ToMinimalDisplayString(semanticModel, ifAssign.SpanStart))

            Dim ifType As ITypeSymbol = semanticModel.GetTypeInfo(ifAssign.Right).Type
            Dim elseType As ITypeSymbol = semanticModel.GetTypeInfo(elseAssign.Right).Type

            Dim trueExpression As ExpressionSyntax = ifAssign.
                ExtractAssignmentAsExpressionSyntax().
                EnsureNothingAsType(semanticModel, type, typeSyntax).
                ConvertToBaseType(ifType, type)

            Dim falseExpression As ExpressionSyntax = elseAssign.
                ExtractAssignmentAsExpressionSyntax().
                EnsureNothingAsType(semanticModel, type, typeSyntax).
                ConvertToBaseType(elseType, type)

            If ifAssign.OperatorToken.Text <> "=" AndAlso ifAssign.OperatorToken.Text = elseAssign.OperatorToken.Text Then
                trueExpression = ifAssign.Right.
                EnsureNothingAsType(semanticModel, type, typeSyntax).
                ConvertToBaseType(ifType, type)

                falseExpression = elseAssign.Right.
                EnsureNothingAsType(semanticModel, type, typeSyntax).
                ConvertToBaseType(elseType, type)
            End If

            Dim leadingTrivia As SyntaxTriviaList = ifBlock.GetLeadingTrivia.
                AddRange(ifAssign.GetLeadingTrivia()).
                AddRange(trueExpression.GetLeadingTrivia()).
                AddRange(elseAssign.GetLeadingTrivia()).
                AddRange(falseExpression.GetLeadingTrivia())

            Dim trailingTrivia As SyntaxTriviaList = ifBlock.GetTrailingTrivia.
                InsertRange(0, elseAssign.GetTrailingTrivia().Where(Function(trivia As SyntaxTrivia) Not trivia.IsKind(SyntaxKind.EndOfLineTrivia))).
                InsertRange(0, ifAssign.GetTrailingTrivia().Where(Function(trivia As SyntaxTrivia) Not trivia.IsKind(SyntaxKind.EndOfLineTrivia)))

            Dim ternary As TernaryConditionalExpressionSyntax = SyntaxFactory.TernaryConditionalExpression(ifBlock.IfStatement.Condition.WithoutTrailingTrivia(),
                                                                     trueExpression.WithoutTrailingTrivia(),
                                                                     falseExpression.WithoutTrailingTrivia())

            Dim ternaryOperatorToken As SyntaxToken = If((ifAssign.OperatorToken.Text <> "=" OrElse elseAssign.OperatorToken.Text <> "=") AndAlso ifAssign.OperatorToken.Text <> elseAssign.OperatorToken.Text,
                                                          SyntaxFactory.Token(SyntaxKind.EqualsToken),
                                                          ifAssign.OperatorToken)

            Dim assignment As AssignmentStatementSyntax = SyntaxFactory.SimpleAssignmentStatement(ifAssign.Left.WithLeadingTrivia(leadingTrivia), ternaryOperatorToken, ternary).
                WithTrailingTrivia(trailingTrivia).
                WithAdditionalAnnotations(Formatter.Annotation)

            Dim newRoot As SyntaxNode = root.ReplaceNode(ifBlock, assignment)
            Dim newDocument As Document = document.WithSyntaxRoot(newRoot)
            Return newDocument
        End Function
    End Class

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(TernaryOperatorFromIifCodeFixProvider)), Composition.Shared>
    Public Class TernaryOperatorFromIifCodeFixProvider
        Inherits CodeFixProvider

        Public Overrides ReadOnly Property FixableDiagnosticIds() As ImmutableArray(Of String) =
            ImmutableArray.Create(DiagnosticIds.TernaryOperator_IifDiagnosticId)

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First
            context.RegisterCodeFix(CodeAction.Create("Change IIF to If to short circuit evaulations", Function(c As CancellationToken) Me.MakeTernaryAsync(context.Document, diagnostic, c), NameOf(TernaryOperatorFromIifCodeFixProvider)), diagnostic)
            Return Task.FromResult(0)
        End Function

        Private Async Function MakeTernaryAsync(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim iifAssignment As InvocationExpressionSyntax = root.FindToken(diagnostic.Location.SourceSpan.Start).Parent.FirstAncestorOrSelf(Of InvocationExpressionSyntax)

            Dim ternary As TernaryConditionalExpressionSyntax = SyntaxFactory.TernaryConditionalExpression(
                iifAssignment.ArgumentList.Arguments(0).GetExpression(),
                iifAssignment.ArgumentList.Arguments(1).GetExpression(),
                iifAssignment.ArgumentList.Arguments(2).GetExpression()).
                WithLeadingTrivia(iifAssignment.GetLeadingTrivia()).
                WithTrailingTrivia(iifAssignment.GetTrailingTrivia()).
                WithAdditionalAnnotations(Formatter.Annotation)

            Dim newRoot As SyntaxNode = root.ReplaceNode(iifAssignment, ternary)
            Dim newDocument As Document = document.WithSyntaxRoot(newRoot)
            Return newDocument
        End Function
    End Class
End Namespace