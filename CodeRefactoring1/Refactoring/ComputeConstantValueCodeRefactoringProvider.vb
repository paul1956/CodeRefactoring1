' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeRefactorings
Imports Microsoft.CodeAnalysis.Formatting
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Refactoring

    <ExportCodeRefactoringProvider(LanguageNames.VisualBasic, Name:="Compute constant value")>
    Public Class ComputeConstantValueCodeRefactoringProvider
        Inherits CodeRefactoringProvider

        Friend Shared Function GetLiteralExpression(value As Object) As ExpressionSyntax
            If TypeOf value Is Boolean Then
                Return If(DirectCast(value, Boolean), SyntaxFactory.TrueLiteralExpression(SyntaxFactory.Token(SyntaxKind.TrueKeyword)), SyntaxFactory.FalseLiteralExpression(SyntaxFactory.Token(SyntaxKind.FalseKeyword)))
            End If
            If TypeOf value Is Byte Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(DirectCast(value, Byte)))
            End If
            If TypeOf value Is SByte Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(DirectCast(value, SByte)))
            End If
            If TypeOf value Is Short Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(DirectCast(value, Short)))
            End If
            If TypeOf value Is UShort Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(DirectCast(value, UShort)))
            End If
            If TypeOf value Is Integer Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(DirectCast(value, Integer)))
            End If
            If TypeOf value Is UInteger Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(DirectCast(value, UInteger)))
            End If
            If TypeOf value Is Long Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(DirectCast(value, Long)))
            End If
            If TypeOf value Is ULong Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(DirectCast(value, ULong)))
            End If

            If TypeOf value Is Single Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(DirectCast(value, Single)))
            End If
            If TypeOf value Is Double Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(DirectCast(value, Double)))
            End If
            If TypeOf value Is Decimal Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(DirectCast(value, Decimal)))
            End If

            If TypeOf value Is Char Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.CharacterLiteralExpression, SyntaxFactory.Literal(DirectCast(value, Char)))
            End If

            If TypeOf value Is String Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.StringLiteralExpression, SyntaxFactory.Literal(DirectCast(value, String)))
            End If

            If value Is Nothing Then
                Return SyntaxFactory.NothingLiteralExpression(SyntaxFactory.Token(SyntaxKind.NothingKeyword))
            End If

            Return Nothing
        End Function

        Public Overrides Async Function ComputeRefactoringsAsync(context As CodeRefactoringContext) As Task
            Dim document As Document = context.Document
            If document.Project.Solution.Workspace.Kind = WorkspaceKind.MiscellaneousFiles Then
                Return
            End If
            Dim span As TextSpan = context.Span
            If Not span.IsEmpty Then
                Return
            End If
            Dim cancellationToken As CancellationToken = context.CancellationToken
            If cancellationToken.IsCancellationRequested Then
                Return
            End If
            Dim model As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken).ConfigureAwait(False)
            If model.IsFromGeneratedCode(cancellationToken) Then
                Return
            End If
            Dim root As SyntaxNode = Await model.SyntaxTree.GetRootAsync(cancellationToken).ConfigureAwait(False)

            Dim expr As ExpressionSyntax = root.FindNode(span).FirstAncestorOrSelf(Function(n As ExpressionSyntax) TypeOf n Is BinaryExpressionSyntax OrElse TypeOf n Is UnaryExpressionSyntax)
            If expr Is Nothing Then
                Return
            End If
            If TypeOf expr Is BinaryExpressionSyntax Then
                If CType(expr, BinaryExpressionSyntax).OperatorToken.SpanStart <> span.Start AndAlso expr.SpanStart <> span.Start Then
                    Return
                End If
            Else
                If expr.SpanStart <> span.Start Then
                    Return
                End If
            End If
            Dim result As [Optional](Of Object) = model.GetConstantValue(expr, cancellationToken)
            If Not result.HasValue Then
                Return
            End If
            Dim syntaxNode As ExpressionSyntax = GetLiteralExpression(result.Value)
            If syntaxNode Is Nothing Then
                Return
            End If
            context.RegisterRefactoring(New DocumentChangeAction(root.FindNode(span).Span,
                                                                 DiagnosticSeverity.Info,
                                                                 GetString("Compute constant value"),
                                                                 Function(t2 As CancellationToken)
                                                                     Dim newRoot As SyntaxNode = root.ReplaceNode(expr, syntaxNode.WithAdditionalAnnotations(Formatter.Annotation))
                                                                     Return Task.FromResult(document.WithSyntaxRoot(newRoot))
                                                                 End Function))
        End Function

    End Class

End Namespace
