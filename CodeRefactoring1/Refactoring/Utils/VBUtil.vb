Option Explicit On
Option Infer Off
Option Strict On
Namespace Refactoring
    Public Module VBUtil
        ''' <summary>
        ''' Inverts a boolean condition. Note: The condition object can be frozen (from AST) it's cloned internally.
        ''' </summary>
        ''' <param name="condition">The condition to invert.</param>
        Friend Function InvertCondition(condition As ExpressionSyntax) As ExpressionSyntax
            If TypeOf condition Is ParenthesizedExpressionSyntax Then
                Return SyntaxFactory.ParenthesizedExpression(InvertCondition(DirectCast(condition, ParenthesizedExpressionSyntax).Expression))
            End If

            If TypeOf condition Is UnaryExpressionSyntax Then
                Dim uOp As UnaryExpressionSyntax = DirectCast(condition, UnaryExpressionSyntax)
                If uOp.IsKind(SyntaxKind.NotExpression) Then
                    Return uOp.Operand
                End If
                Return SyntaxFactory.UnaryExpression(SyntaxKind.NotExpression, uOp.OperatorToken, uOp)
            End If

            If TypeOf condition Is BinaryExpressionSyntax Then
                Dim bOp As BinaryExpressionSyntax = DirectCast(condition, BinaryExpressionSyntax)

                If bOp.IsKind(SyntaxKind.AndExpression) OrElse bOp.IsKind(SyntaxKind.AndAlsoExpression) OrElse bOp.IsKind(SyntaxKind.OrExpression) OrElse bOp.IsKind(SyntaxKind.OrElseExpression) Then
                    Dim kind As SyntaxKind = NegateConditionOperator(bOp.Kind())
                    Return SyntaxFactory.BinaryExpression(kind, InvertCondition(bOp.Left), SyntaxFactory.Token(GetBinaryExpressionOperatorTokenKind(kind)), InvertCondition(bOp.Right))
                End If

                If bOp.IsKind(SyntaxKind.EqualsExpression) OrElse bOp.IsKind(SyntaxKind.NotEqualsExpression) OrElse bOp.IsKind(SyntaxKind.GreaterThanExpression) OrElse bOp.IsKind(SyntaxKind.GreaterThanOrEqualExpression) OrElse bOp.IsKind(SyntaxKind.LessThanExpression) OrElse bOp.IsKind(SyntaxKind.LessThanOrEqualExpression) Then
                    Dim kind As SyntaxKind = NegateRelationalOperator(bOp.Kind())
                    Return SyntaxFactory.BinaryExpression(kind, bOp.Left, SyntaxFactory.Token(GetBinaryExpressionOperatorTokenKind(kind)), bOp.Right)
                End If

                Return SyntaxFactory.UnaryExpression(SyntaxKind.NotExpression, SyntaxFactory.Token(SyntaxKind.NotKeyword), SyntaxFactory.ParenthesizedExpression(condition))
            End If

            If TypeOf condition Is TernaryConditionalExpressionSyntax Then
                Dim cEx As TernaryConditionalExpressionSyntax = TryCast(condition, TernaryConditionalExpressionSyntax)
                Return cEx.WithCondition(InvertCondition(cEx.Condition))
            End If

            If TypeOf condition Is LiteralExpressionSyntax Then
                If condition.Kind() = SyntaxKind.TrueLiteralExpression Then
                    Return SyntaxFactory.LiteralExpression(SyntaxKind.FalseLiteralExpression, SyntaxFactory.Token(SyntaxKind.FalseKeyword))
                End If
                If condition.Kind() = SyntaxKind.FalseLiteralExpression Then
                    Return SyntaxFactory.LiteralExpression(SyntaxKind.TrueLiteralExpression, SyntaxFactory.Token(SyntaxKind.TrueKeyword))
                End If
            End If

            Return SyntaxFactory.UnaryExpression(SyntaxKind.NotExpression, SyntaxFactory.Token(SyntaxKind.NotKeyword), AddParensForUnaryExpressionIfRequired(condition))
        End Function
        ''' <summary>
        ''' When negating an expression this is required, otherwise you would end up with
        ''' a or b -> !a or b
        ''' </summary>
        Private Function AddParensForUnaryExpressionIfRequired(expression As ExpressionSyntax) As ExpressionSyntax
            If TypeOf expression Is BinaryExpressionSyntax OrElse TypeOf expression Is CastExpressionSyntax OrElse TypeOf expression Is LambdaExpressionSyntax Then
                Return SyntaxFactory.ParenthesizedExpression(expression)
            End If

            Return expression
        End Function
        Private Function GetBinaryExpressionOperatorTokenKind(op As SyntaxKind) As SyntaxKind
            Select Case op
                Case SyntaxKind.EqualsExpression
                    Return SyntaxKind.EqualsToken
                Case SyntaxKind.NotEqualsExpression
                    Return SyntaxKind.LessThanGreaterThanToken
                Case SyntaxKind.GreaterThanExpression
                    Return SyntaxKind.GreaterThanToken
                Case SyntaxKind.GreaterThanOrEqualExpression
                    Return SyntaxKind.GreaterThanEqualsToken
                Case SyntaxKind.LessThanExpression
                    Return SyntaxKind.LessThanToken
                Case SyntaxKind.LessThanOrEqualExpression
                    Return SyntaxKind.LessThanEqualsToken
                Case SyntaxKind.OrExpression
                    Return SyntaxKind.OrKeyword
                Case SyntaxKind.OrElseExpression
                    Return SyntaxKind.OrElseKeyword
                Case SyntaxKind.AndExpression
                    Return SyntaxKind.AndKeyword
                Case SyntaxKind.AndAlsoExpression
                    Return SyntaxKind.AndAlsoKeyword
            End Select
            Throw New ArgumentOutOfRangeException(NameOf(op))
        End Function
        ''' <summary>
        ''' Get negation of the condition operator
        ''' </summary>
        ''' <returns>
        ''' negation of the specified condition operator, or BinaryOperatorType.Any if it's not a condition operator
        ''' </returns>
        Private Function NegateConditionOperator(op As SyntaxKind) As SyntaxKind
            Select Case op
                Case SyntaxKind.OrExpression
                    Return SyntaxKind.AndExpression
                Case SyntaxKind.OrElseExpression
                    Return SyntaxKind.AndAlsoExpression
                Case SyntaxKind.AndExpression
                    Return SyntaxKind.OrExpression
                Case SyntaxKind.AndAlsoExpression
                    Return SyntaxKind.OrElseExpression
            End Select
            Throw New ArgumentOutOfRangeException(NameOf(op))
        End Function
        ''' <summary>
        ''' Get negation of the specified relational operator
        ''' </summary>
        ''' <returns>
        ''' negation of the specified relational operator, or BinaryOperatorType.Any if it's not a relational operator
        ''' </returns>
        Private Function NegateRelationalOperator(op As SyntaxKind) As SyntaxKind
            Select Case op
                Case SyntaxKind.EqualsExpression
                    Return SyntaxKind.NotEqualsExpression
                Case SyntaxKind.NotEqualsExpression
                    Return SyntaxKind.EqualsExpression
                Case SyntaxKind.GreaterThanExpression
                    Return SyntaxKind.LessThanOrEqualExpression
                Case SyntaxKind.GreaterThanOrEqualExpression
                    Return SyntaxKind.LessThanExpression
                Case SyntaxKind.LessThanExpression
                    Return SyntaxKind.GreaterThanOrEqualExpression
                Case SyntaxKind.LessThanOrEqualExpression
                    Return SyntaxKind.GreaterThanExpression
                Case SyntaxKind.OrExpression
                    Return SyntaxKind.AndExpression
                Case SyntaxKind.OrElseExpression
                    Return SyntaxKind.AndAlsoExpression
                Case SyntaxKind.AndExpression
                    Return SyntaxKind.OrExpression
                Case SyntaxKind.AndAlsoExpression
                    Return SyntaxKind.OrElseExpression
            End Select
            Throw New ArgumentOutOfRangeException(NameOf(op))
        End Function
    End Module
End Namespace
