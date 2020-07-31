' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Formatting
Imports Microsoft.CodeAnalysis.Simplification

Namespace Performance

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(StringBuilderInLoopCodeFixProvider)), Composition.Shared>
    Public Class StringBuilderInLoopCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(StringBuilderInLoopAnalyzer.Id)

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return Nothing
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First
            context.RegisterCodeFix(CodeAction.Create($"Use StringBuilder to create a value for '{diagnostic.Properties!assignmentExpressionLeft}'", Function(c As CancellationToken) UseStringBuilder(context.Document, diagnostic, c), NameOf(StringBuilderInLoopCodeFixProvider)), diagnostic)
            Return Task.FromResult(0)
        End Function

        Private Async Function UseStringBuilder(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim diagosticSpan As TextSpan = diagnostic.Location.SourceSpan
            Dim expressionStatement As AssignmentStatementSyntax = root.FindToken(diagosticSpan.Start).Parent.AncestorsAndSelf.OfType(Of AssignmentStatementSyntax).First

            Dim expressionStatementParent As SyntaxNode = expressionStatement.Parent
            Dim semanticModel As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken)
            Dim builderName As String = FindAvailableStringBuilderVariableName(expressionStatement, semanticModel)
            Dim loopStatement As SyntaxNode = expressionStatement.FirstAncestorOrSelfOfType(
            GetType(WhileBlockSyntax),
            GetType(ForBlockSyntax),
            GetType(ForEachBlockSyntax),
            GetType(DoLoopBlockSyntax))

            Dim newExpressionStatementParent As SyntaxNode = ReplaceAddExpressionByStringBuilderAppendExpression(expressionStatement, expressionStatement, expressionStatementParent, builderName)
            Dim newLoopStatement As SyntaxNode = loopStatement.ReplaceNode(expressionStatementParent, newExpressionStatementParent)
            Dim stringBuilderType As TypeSyntax = SyntaxFactory.ParseTypeName("System.Text.StringBuilder").WithAdditionalAnnotations(Simplifier.Annotation)

            Dim declarators As New SeparatedSyntaxList(Of VariableDeclaratorSyntax)()
            declarators = declarators.Add(SyntaxFactory.VariableDeclarator(New SeparatedSyntaxList(Of ModifiedIdentifierSyntax)().Add(SyntaxFactory.ModifiedIdentifier(builderName)),
             SyntaxFactory.AsNewClause(SyntaxFactory.ObjectCreationExpression(stringBuilderType).WithArgumentList(SyntaxFactory.ArgumentList())),
             Nothing))

            Dim stringBuilderDeclaration As LocalDeclarationStatementSyntax = SyntaxFactory.LocalDeclarationStatement(SyntaxTokenList.Create(SyntaxFactory.Token(SyntaxKind.DimKeyword)),
                                                                               declarators).NormalizeWhitespace(" ").WithTrailingTrivia(SyntaxFactory.CarriageReturnLineFeed)

            Dim appendExpressionOnInitialization As StatementSyntax = SyntaxFactory.ParseExecutableStatement(builderName & ".Append(" & expressionStatement.Left.ToString() & ")").WithTrailingTrivia(SyntaxFactory.CarriageReturnLineFeed)
            Dim stringBuilderToString As StatementSyntax = SyntaxFactory.ParseExecutableStatement(expressionStatement.Left.ToString() & " = " & builderName & ".ToString()").WithTrailingTrivia(SyntaxFactory.CarriageReturnLineFeed)

            Dim loopParent As SyntaxNode = loopStatement.Parent
            Dim newLoopParent As SyntaxNode = loopParent.ReplaceNode(loopStatement,
                                                   {stringBuilderDeclaration, appendExpressionOnInitialization, newLoopStatement, stringBuilderToString}).
                                                   WithAdditionalAnnotations(Formatter.Annotation)
            Dim newroot As SyntaxNode = root.ReplaceNode(loopParent, newLoopParent)
            Dim newDocument As Document = document.WithSyntaxRoot(newroot)
            Return newDocument
        End Function

        Private Shared Function ReplaceAddExpressionByStringBuilderAppendExpression(assignment As AssignmentStatementSyntax, expressionStatement As SyntaxNode, expressionStatementParent As SyntaxNode, builderName As String) As SyntaxNode
            Dim appendExpressionOnLoop As StatementSyntax = If(assignment.IsKind(SyntaxKind.SimpleAssignmentStatement),
            SyntaxFactory.ParseExecutableStatement(builderName & ".Append(" & DirectCast(assignment.Right, BinaryExpressionSyntax).Right.ToString() & ")"),
            SyntaxFactory.ParseExecutableStatement(builderName & ".Append(" & assignment.Right.ToString() & ")")).
                WithLeadingTrivia(assignment.GetLeadingTrivia()).
                WithTrailingTrivia(assignment.GetTrailingTrivia())

            Dim newExpressionStatementParent As SyntaxNode = expressionStatementParent.ReplaceNode(expressionStatement, appendExpressionOnLoop)
            Return newExpressionStatementParent
        End Function

        Private Shared Function FindAvailableStringBuilderVariableName(assignmentStatement As AssignmentStatementSyntax, semanticModel As SemanticModel) As String
            Const builderNameBase As String = "builder"
            Dim builderName As String = builderNameBase
            Dim builderNameIncrementer As Integer = 0
            While semanticModel.LookupSymbols(assignmentStatement.GetLocation().SourceSpan.Start, name:=builderName).Any()
                builderNameIncrementer += 1
                builderName = builderNameBase & builderNameIncrementer.ToString()
            End While
            Return builderName
        End Function

    End Class

End Namespace
