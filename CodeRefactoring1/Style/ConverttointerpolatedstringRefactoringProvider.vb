Option Explicit On
Option Infer Off
Option Strict On
Imports System.Collections.Immutable
Imports System.Diagnostics.Debug
Imports Microsoft.CodeAnalysis.Simplification
Namespace Style

    <ExportCodeRefactoringProvider(LanguageNames.VisualBasic, Name:=NameOf(ConverttointerpolatedstringRefactoringProvider)), [Shared]>
    Friend Class ConverttointerpolatedstringRefactoringProvider
        Inherits CodeRefactoringProvider

        Public NotOverridable Overrides Async Function ComputeRefactoringsAsync(context As CodeRefactoringContext) As Task
            Try

                Dim semanticModel As SemanticModel = Await context.Document.GetSemanticModelAsync(context.CancellationToken)
                Dim stringType As INamedTypeSymbol = semanticModel.Compilation.GetTypeByMetadataName("System.String")
                If stringType Is Nothing Then
                    Return
                End If
                Dim formatMethods As ImmutableArray(Of ISymbol) = stringType.GetMembers(NameOf(Format)).RemoveAll(AddressOf IsValidStringFormatMethod)
                If formatMethods.Length = 0 Then
                    Return
                End If
                Dim root As SyntaxNode = Await context.Document.GetSyntaxRootAsync(context.CancellationToken)
                Dim invocation As InvocationExpressionSyntax = root.FindNode(context.Span, getInnermostNodeForTie:=True)?.FirstAncestorOrSelf(Of InvocationExpressionSyntax)()

                While invocation IsNot Nothing
                    If invocation.ArgumentList IsNot Nothing Then
                        Dim arguments As SeparatedSyntaxList(Of ArgumentSyntax) = invocation.ArgumentList.Arguments
                        If (arguments.Count >= 2) Then
                            Dim firstArgument As LiteralExpressionSyntax = TryCast(arguments(0)?.GetExpression, LiteralExpressionSyntax)
                            If firstArgument Is Nothing Then
                                Return
                            End If
                            If firstArgument?.Token.IsKind(SyntaxKind.StringLiteralToken) = True Then
                                Dim invocationSymbol As ISymbol = semanticModel.GetSymbolInfo(invocation, context.CancellationToken).Symbol
                                If formatMethods.Contains(invocationSymbol) Then
                                    Exit While
                                End If
                            End If
                        End If
                    End If
                    invocation = invocation.Parent?.FirstAncestorOrSelf(Of InvocationExpressionSyntax)()
                End While
                If invocation IsNot Nothing Then
                    context.RegisterRefactoring(New ConvertToInterpolatedStringCodeAction("Convert to interpolated string",
                                                    Function(c As CancellationToken) Me.CreateInterpolatedStringAsync(invocation, context.Document, c)))
                End If
            Catch ex As Exception
                Stop
            End Try
        End Function
        Private Async Function CreateInterpolatedStringAsync(ByVal invocation As InvocationExpressionSyntax, ByVal document As Document, ByVal cancellationToken As CancellationToken) As Task(Of Document)
            Assert(Not (invocation.ArgumentList) Is Nothing)
            Assert(invocation.ArgumentList.Arguments.Count >= 2)
            Assert(invocation.ArgumentList.Arguments(0).GetExpression.IsKind(SyntaxKind.StringLiteralExpression))
            Try
                Dim semanticModel As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken)
                Dim arguments As SeparatedSyntaxList(Of ArgumentSyntax) = invocation.ArgumentList.Arguments
                Dim text As String = CType(arguments(0).GetExpression, LiteralExpressionSyntax).Token.ToString()
                Dim builder As ImmutableArray(Of ExpressionSyntax).Builder = ImmutableArray.CreateBuilder(Of ExpressionSyntax)()
                For i As Integer = 1 To arguments.Count - 1
                    builder.Add(CastAndParenthesize(arguments(i).GetExpression, semanticModel))
                Next
                Dim expandedArguments As ImmutableArray(Of ExpressionSyntax) = builder.ToImmutable
                Dim interpolatedString As InterpolatedStringExpressionSyntax = CType(SyntaxFactory.ParseExpression($"${text}"), InterpolatedStringExpressionSyntax)
                Dim newInterpolatedString As InterpolatedStringExpressionSyntax = InterpolatedStringRewriter.Visit(interpolatedString, expandedArguments)
                Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken)
                Dim newRoot As SyntaxNode = root.ReplaceNode(invocation, newInterpolatedString.WithLeadingTrivia(invocation.GetLeadingTrivia()).WithTrailingTrivia(invocation.GetTrailingTrivia()))
                Return document.WithSyntaxRoot(newRoot)
            Catch ex As Exception
                Throw ex
            End Try
            Return Nothing
        End Function
        Private Shared Function Parenthesize(expression As ExpressionSyntax) As ExpressionSyntax
            Return If(expression.IsKind(SyntaxKind.ParenthesizedExpression), expression, SyntaxFactory.ParenthesizedExpression(openParenToken:=SyntaxFactory.Token(SyntaxTriviaList.Empty, SyntaxKind.OpenParenToken, SyntaxTriviaList.Empty), expression:=expression, closeParenToken:=SyntaxFactory.Token(SyntaxTriviaList.Empty, SyntaxKind.CloseParenToken, SyntaxTriviaList.Empty)).WithAdditionalAnnotations(Simplifier.Annotation))
        End Function
        Private Shared Function Cast(ByVal expression As ExpressionSyntax, ByVal targetType As ITypeSymbol) As ExpressionSyntax
            If (targetType Is Nothing) Then
                Return expression
            End If
            Dim type As TypeSyntax = SyntaxFactory.ParseTypeName(targetType.ToDisplayString)
            Return SyntaxFactory.CTypeExpression(Parenthesize(expression), type)
        End Function
        ''' <summary>Creates a new CastExpressionSyntax instance.</summary>
        Private Shared Function CastAndParenthesize(ByVal expression As ExpressionSyntax, ByVal semanticModel As SemanticModel) As ExpressionSyntax
            Dim targetType As ITypeSymbol = semanticModel.GetTypeInfo(expression).ConvertedType
            If targetType.SpecialType = SpecialType.System_Object Then
                Return expression
            End If
            Return Parenthesize(Cast(expression, targetType))
        End Function
        Private Shared Function IsValidStringFormatMethod(ByVal symbol As ISymbol) As Boolean
            If symbol.Kind <> SymbolKind.Method OrElse Not symbol.IsStatic Then
                Return True
            End If
            Dim methodSymbol As IMethodSymbol = CType(symbol, IMethodSymbol)
            If methodSymbol.Parameters.Length = 0 Then
                Return True
            End If
            Dim firstParameter As IParameterSymbol = methodSymbol.Parameters(0)
            If firstParameter?.Name <> "format" Then
                Return True
            End If
            Return False
        End Function
        Private Class InterpolatedStringRewriter
            Inherits VisualBasicSyntaxRewriter
            Private ReadOnly expandedArguments As ImmutableArray(Of ExpressionSyntax)

            Private Sub New(expandedArguments As ImmutableArray(Of ExpressionSyntax))
                Me.expandedArguments = expandedArguments
            End Sub

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
            Public Overloads Shared Function Visit(interpolatedString As InterpolatedStringExpressionSyntax, expandedArguments As ImmutableArray(Of ExpressionSyntax)) As InterpolatedStringExpressionSyntax
                Return DirectCast(New InterpolatedStringRewriter(expandedArguments).Visit(interpolatedString), InterpolatedStringExpressionSyntax)
            End Function
        End Class
        Private Class ConvertToInterpolatedStringCodeAction
            Inherits CodeAction

            Private ReadOnly generateDocument As Func(Of CancellationToken, Task(Of Document))
            Private ReadOnly _title As String

            Public Overrides ReadOnly Property Title As String
                Get
                    Return Me._title
                End Get
            End Property

            Public Sub New(title As String, generateDocument As Func(Of CancellationToken, Task(Of Document)))
                Me._title = title
                Me.generateDocument = generateDocument
            End Sub

            Protected Overrides Function GetChangedDocumentAsync(cancellationToken As CancellationToken) As Task(Of Document)
                Return Me.generateDocument(cancellationToken)
            End Function
        End Class
    End Class
End Namespace