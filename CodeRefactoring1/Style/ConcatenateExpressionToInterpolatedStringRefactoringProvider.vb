Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Collections.Immutable

Namespace Style

    <ExportCodeRefactoringProvider(LanguageNames.VisualBasic, Name:=NameOf(ConcatenateExpressionToInterpolatedStringRefactoringProvider)), [Shared]>
    Friend Class ConcatenateExpressionToInterpolatedStringRefactoringProvider
        Inherits CodeRefactoringProvider
        Private Shared ReadOnly CloseBrace As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CloseBraceToken)
        Private Shared ReadOnly OpenBrace As SyntaxToken = SyntaxFactory.Token(SyntaxKind.OpenBraceToken)
        Private Shared Function ExtractLeftTextFromExpression(invocation As BinaryExpressionSyntax) As SyntaxList(Of InterpolatedStringContentSyntax)
            Debug.Print($"Entering {Reflection.MethodBase.GetCurrentMethod().Name} {NameOf(invocation.Left)} = {invocation.Left.ToString})")
            Dim ContextArray As New SyntaxList(Of InterpolatedStringContentSyntax)()
            Try
                Select Case invocation.Left.Kind
                    Case SyntaxKind.ConcatenateExpression
                        ContextArray = MergerExpression(CType(invocation.Left, BinaryExpressionSyntax), ContextArray)
                    Case SyntaxKind.IdentifierName, SyntaxKind.InterpolatedStringExpression, SyntaxKind.InvocationExpression,
                     SyntaxKind.NameOfExpression, SyntaxKind.ObjectCreationExpression, SyntaxKind.SimpleMemberAccessExpression,
                     SyntaxKind.TernaryConditionalExpression
                        ContextArray = ContextArray.Add(SyntaxFactory.Interpolation(OpenBrace, invocation.Left.WithoutTrivia, Nothing, Nothing, CloseBrace).WithTriviaFrom(invocation.Left))
                    Case SyntaxKind.StringLiteralExpression
                        Dim InvocationLeftToken As SyntaxToken = CType(invocation.Left, LiteralExpressionSyntax).Token
                        Dim InvocationTokenValueText As String = InvocationLeftToken.ValueText.Replace("""", """""")
                        Dim TextToken As SyntaxToken = SyntaxFactory.InterpolatedStringTextToken(InvocationLeftToken.LeadingTrivia, InvocationTokenValueText, InvocationTokenValueText, InvocationLeftToken.TrailingTrivia)
                        Dim InterpolatedStringTextToken As InterpolatedStringTextSyntax = SyntaxFactory.InterpolatedStringText(TextToken)
                        ContextArray = ContextArray.Add(InterpolatedStringTextToken)
                    Case Else
                        Throw New ArgumentException($"{NameOf(invocation.Left.Kind)} missing!")
                End Select
            Catch ex As Exception
                Stop
                Throw New ArgumentException($"{NameOf(invocation.Kind)} missing!")
            End Try
            Return ContextArray
        End Function

        Private Shared Function ExtractRightTextExpression(invocation As ExpressionSyntax) As InterpolatedStringContentSyntax
            Debug.Print($"Entering {Reflection.MethodBase.GetCurrentMethod().Name} {NameOf(invocation)} As {invocation.GetType.ToString} = {invocation.ToString})")
            Try
                Select Case invocation.Kind
                    Case SyntaxKind.IdentifierName, SyntaxKind.InterpolatedStringExpression, SyntaxKind.InvocationExpression,
                     SyntaxKind.NameOfExpression, SyntaxKind.ObjectCreationExpression, SyntaxKind.SimpleMemberAccessExpression,
                     SyntaxKind.TernaryConditionalExpression
                        Return SyntaxFactory.Interpolation(OpenBrace, invocation.WithoutTrivia, Nothing, Nothing, CloseBrace).WithTriviaFrom(invocation)
                    Case SyntaxKind.StringLiteralExpression
                        Dim InvicationRightToken As SyntaxToken = CType(invocation, LiteralExpressionSyntax).Token
                        Dim InvicationRightTokenValueText As String = InvicationRightToken.ValueText.Replace("""", """""")
                        Dim TextToken As SyntaxToken = SyntaxFactory.InterpolatedStringTextToken(InvicationRightToken.LeadingTrivia, InvicationRightTokenValueText, InvicationRightTokenValueText, InvicationRightToken.TrailingTrivia)
                        Dim InterpolatedStringTextToken As InterpolatedStringTextSyntax = SyntaxFactory.InterpolatedStringText(TextToken)
                        Return InterpolatedStringTextToken
                    Case Else
                        Throw New ArgumentException($"{NameOf(invocation.Kind)} missing!")
                End Select
            Catch ex As Exception
                Stop
                Throw
            End Try
        End Function

        Private Shared Function ExtractRightTextExpression(invocation As BinaryExpressionSyntax) As InterpolatedStringContentSyntax
            Debug.Print($"Entering {Reflection.MethodBase.GetCurrentMethod().Name} {NameOf(invocation.Right)} As {invocation.GetType.ToString} = {invocation.Right.ToString})")
            Try
                Return ExtractRightTextExpression(invocation.Right)
            Catch ex As Exception
                Stop
                Throw
            End Try
        End Function

        Private Shared Function MergerExpression(invocation As BinaryExpressionSyntax) As SyntaxList(Of InterpolatedStringContentSyntax)
            Debug.Print($"Entering {Reflection.MethodBase.GetCurrentMethod().Name} {NameOf(invocation)} = {invocation.ToString})")
            Try
                Dim ContentsArray As SyntaxList(Of InterpolatedStringContentSyntax) = ExtractLeftTextFromExpression(invocation)
                Return ContentsArray.Add(ExtractRightTextExpression(invocation))
            Catch es As Exception
                Stop
                Throw
            End Try
        End Function

        Private Shared Function MergerExpression(ByVal invocation As BinaryExpressionSyntax, ByRef ContentsArray As SyntaxList(Of InterpolatedStringContentSyntax)) As SyntaxList(Of InterpolatedStringContentSyntax)
            Debug.Print($"Entering {Reflection.MethodBase.GetCurrentMethod().Name} {NameOf(invocation)} = {invocation.ToString})")
            Try
                ContentsArray = ExtractLeftTextFromExpression(invocation)
                Return ContentsArray.Add(ExtractRightTextExpression(invocation))
            Catch es As Exception
                Stop
                Throw
            End Try
        End Function

        Private Async Function CreateInterpolatedStringFromExpressionAsync(ByVal invocation As BinaryExpressionSyntax, ByVal document As Document, ByVal cancellationToken As CancellationToken) As Task(Of Document)
            Debug.Print($"Entering CreateInterpolatedStringFromExpressionAsync {NameOf(invocation)} = {invocation.ToString})")
            Try
                Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken)
                Dim semanticModel As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken)
                Dim ContextsArray As SyntaxList(Of InterpolatedStringContentSyntax) = New SyntaxList(Of InterpolatedStringContentSyntax)().AddRange(MergerExpression(invocation))
                Dim newInvocation As InterpolatedStringExpressionSyntax = SyntaxFactory.InterpolatedStringExpression(ContextsArray)
                Dim newRoot As SyntaxNode = root.ReplaceNode(invocation, newInvocation.WithLeadingTrivia(invocation.GetLeadingTrivia()).WithTrailingTrivia(invocation.GetTrailingTrivia()))
                Return document.WithSyntaxRoot(newRoot)
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Stop
                Throw
            End Try
            Return Nothing
        End Function

        Private Async Function CreateInterpolatedStringFromInterpolatedStringExpressionAsync(ByVal invocation As BinaryExpressionSyntax, ByVal document As Document, ByVal cancellationToken As CancellationToken) As Task(Of Document)
            Debug.Print($"Entering {Reflection.MethodBase.GetCurrentMethod().Name} {NameOf(invocation)} = {invocation.ToString})")
            Try
                Dim semanticModel As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken)
                Dim InterpolatedStringText As InterpolatedStringExpressionSyntax = CType(SyntaxFactory.ParseExpression($"$""{MergerExpression(invocation)}"""), InterpolatedStringExpressionSyntax)
                Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken)
                Dim newRoot As SyntaxNode = root.ReplaceNode(invocation, InterpolatedStringText.WithLeadingTrivia(invocation.GetLeadingTrivia()).WithTrailingTrivia(invocation.GetTrailingTrivia()))
                Debug.Print($"Leaving {Reflection.MethodBase.GetCurrentMethod().Name} {NameOf(InterpolatedStringText)} = {InterpolatedStringText})")
                Return document.WithSyntaxRoot(newRoot)
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Stop
                Throw
            End Try
            Return Nothing
        End Function

        Private Async Function CreateInterpolatedStringFromStringLiteralExpressionAsync(ByVal invocation As BinaryExpressionSyntax, ByVal document As Document, ByVal cancellationToken As CancellationToken) As Task(Of Document)
            Debug.Print($"Entering CreateInterpolatedStringFromStringLiteralExpressionAsync {NameOf(invocation)} = {invocation.ToString})")
            Try
                Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken)
                Dim semanticModel As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken)
                Dim ContextsArray As SyntaxList(Of InterpolatedStringContentSyntax) = New SyntaxList(Of InterpolatedStringContentSyntax)().AddRange(MergerExpression(invocation))
                Dim newInvocation As InterpolatedStringExpressionSyntax = SyntaxFactory.InterpolatedStringExpression(ContextsArray)
                Dim newRoot As SyntaxNode = root.ReplaceNode(invocation, newInvocation.WithLeadingTrivia(invocation.GetLeadingTrivia()).WithTrailingTrivia(invocation.GetTrailingTrivia()))
                Return document.WithSyntaxRoot(newRoot)
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Stop
                Throw
            End Try
            Return Nothing
        End Function

        Public NotOverridable Overrides Async Function ComputeRefactoringsAsync(context As CodeRefactoringContext) As Task
            Dim semanticModel As SemanticModel = Await context.Document.GetSemanticModelAsync(context.CancellationToken)
            Dim root As SyntaxNode = Await context.Document.GetSyntaxRootAsync(context.CancellationToken)
            Dim invocation As BinaryExpressionSyntax = root.FindNode(context.Span, getInnermostNodeForTie:=True)?.FirstAncestorOrSelf(Of BinaryExpressionSyntax)()
            If Not invocation.IsKind(SyntaxKind.ConcatenateExpression) Then
                Exit Function
            End If
            If Not invocation.Right.MatchesKind(SyntaxKind.InterpolatedStringExpression, SyntaxKind.StringLiteralExpression, SyntaxKind.IdentifierName, SyntaxKind.SimpleMemberAccessExpression, SyntaxKind.InvocationExpression) Then
                Exit Function
            End If
            Select Case invocation.Left.Kind
                Case SyntaxKind.InterpolatedStringExpression
                    context.RegisterRefactoring(New ConcatenateExpressionToInterpolatedStringCodeAction("Merge interpolated string",
                                                Function(c As CancellationToken) Me.CreateInterpolatedStringFromInterpolatedStringExpressionAsync(invocation, context.Document, c)))
                Case SyntaxKind.StringLiteralExpression
                    If invocation.Right.Kind = SyntaxKind.StringLiteralExpression Then
                        context.RegisterRefactoring(New ConcatenateExpressionToInterpolatedStringCodeAction("Concatenate 2 strings",
                                                Function(c As CancellationToken) Me.CreateInterpolatedStringFromStringLiteralExpressionAsync(invocation, context.Document, c)))
                    Else
                        context.RegisterRefactoring(New ConcatenateExpressionToInterpolatedStringCodeAction("Convert string literal to interpolated string",
                                                Function(c As CancellationToken) Me.CreateInterpolatedStringFromStringLiteralExpressionAsync(invocation, context.Document, c)))
                    End If
                Case SyntaxKind.IdentifierName, SyntaxKind.InvocationExpression, SyntaxKind.NameOfExpression, SyntaxKind.SimpleMemberAccessExpression, SyntaxKind.TernaryConditionalExpression
                    context.RegisterRefactoring(New ConcatenateExpressionToInterpolatedStringCodeAction("Convert Identifier or Expression to interpolated string",
                                                Function(c As CancellationToken) Me.CreateInterpolatedStringFromExpressionAsync(invocation, context.Document, c)))
                Case SyntaxKind.ObjectCreationExpression
                    context.RegisterRefactoring(New ConcatenateExpressionToInterpolatedStringCodeAction("Convert Identifier or Expression to interpolated string",
                                                Function(c As CancellationToken) Me.CreateInterpolatedStringFromExpressionAsync(invocation, context.Document, c)))

                Case SyntaxKind.ConcatenateExpression
                    context.RegisterRefactoring(New ConcatenateExpressionToInterpolatedStringCodeAction("Concatenate multiple strings",
                                                Function(c As CancellationToken) Me.CreateInterpolatedStringFromExpressionAsync(invocation, context.Document, c)))
                Case Else
                    Stop
            End Select
        End Function
        Private Class ConcatenateExpressionToInterpolatedStringCodeAction
            Inherits CodeAction

            Private ReadOnly _title As String
            Private ReadOnly generateDocument As Func(Of CancellationToken, Task(Of Document))
            Public Sub New(title As String, generateDocument As Func(Of CancellationToken, Task(Of Document)))
                Me._title = title
                Me.generateDocument = generateDocument
            End Sub

            Public Overrides ReadOnly Property Title As String
                Get
                    Return Me._title
                End Get
            End Property
            Protected Overrides Function GetChangedDocumentAsync(cancellationToken As CancellationToken) As Task(Of Document)
                Return Me.generateDocument(cancellationToken)
            End Function

        End Class

        Private Class InterpolatedStringRewriter
            Inherits VisualBasicSyntaxRewriter
            Private ReadOnly expandedArguments As ImmutableArray(Of ExpressionSyntax)

            Private Sub New(expandedArguments As ImmutableArray(Of ExpressionSyntax))
                Me.expandedArguments = expandedArguments
            End Sub

            Public Overloads Shared Function Visit(interpolatedString As InterpolatedStringExpressionSyntax, expandedArguments As ImmutableArray(Of ExpressionSyntax)) As InterpolatedStringExpressionSyntax
                Return DirectCast(New InterpolatedStringRewriter(expandedArguments).Visit(interpolatedString), InterpolatedStringExpressionSyntax)
            End Function

            Public Overrides Function VisitInterpolation(node As InterpolationSyntax) As SyntaxNode
                Dim literalExpression As LiteralExpressionSyntax = TryCast(node.Expression, LiteralExpressionSyntax)
                If literalExpression IsNot Nothing AndAlso literalExpression.IsKind(SyntaxKind.NumericLiteralExpression) Then
                    Dim index As Integer = CInt(literalExpression.Token.Value)
                    If index >= 0 AndAlso index < Me.expandedArguments.Length Then
                        Return node.WithExpression(Me.expandedArguments(index))
                    End If
                End If

                Return MyBase.VisitInterpolation(node)
            End Function
        End Class
    End Class

End Namespace