Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Usage
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class ArgumentExceptionAnalyzer
        Inherits DiagnosticAnalyzer

        Friend Const Title As String = "Invalid argument name"
        Friend Const MessageFormat As String = "Type argument '{0}' is not in the argument list."
        Private Const Description As String = "The string passed as the 'paramName' argument of ArgumentException constructor must be the name of one of the method arguments.
It can be either specified directly or using nameof() (VB 14 and above only)."
        Friend Shared Rule As New DiagnosticDescriptor(
        DiagnosticIds.ArgumentExceptionDiagnosticId,
        Title,
        MessageFormat,
        SupportedCategories.Naming,
        DiagnosticSeverity.Warning,
        isEnabledByDefault:=True,
        description:=Description,
        helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.ArgumentExceptionDiagnosticId))

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Me.AnalyzeNode, SyntaxKind.ObjectCreationExpression)
        End Sub

        Private Sub AnalyzeNode(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim objectCreationExpression As ObjectCreationExpressionSyntax = DirectCast(context.Node, ObjectCreationExpressionSyntax)
            Dim type As IdentifierNameSyntax = TryCast(objectCreationExpression.Type, IdentifierNameSyntax)
            If type Is Nothing Then Exit Sub
            If Not type.Identifier.ValueText.EndsWith(NameOf(ArgumentException), StringComparison.OrdinalIgnoreCase) Then Exit Sub

            Dim argumentList As ArgumentListSyntax = objectCreationExpression.ArgumentList
            If If(argumentList?.Arguments.Count, 0) < 2 Then Exit Sub

            Dim paramNameLiteral As LiteralExpressionSyntax = TryCast(argumentList.Arguments(1).GetExpression, LiteralExpressionSyntax)
            If paramNameLiteral Is Nothing Then Exit Sub

            Dim paramNameOpt As [Optional](Of Object) = context.SemanticModel.GetConstantValue(paramNameLiteral)
            If Not paramNameOpt.HasValue Then Exit Sub

            Dim paramName As String = paramNameOpt.Value.ToString()

            Dim parameters As IEnumerable(Of String) = Nothing
            If Me.IsParamNameCompatibleWithCreatingContext(objectCreationExpression, paramName, parameters) Then Exit Sub
            Dim props As ImmutableDictionary(Of String, String) = parameters.ToImmutableDictionary(Function(p) $"param{p}", Function(p) p)
            Dim diag As Diagnostic = Diagnostic.Create(Rule, paramNameLiteral.GetLocation, props.ToImmutableDictionary(), paramName)
            context.ReportDiagnostic(diag)
        End Sub

        Private Function IsParamNameCompatibleWithCreatingContext(node As SyntaxNode, paramName As String, ByRef parameters As IEnumerable(Of String)) As Boolean
            parameters = GetParameterNamesFromCreationContext(node)
            If parameters Is Nothing Then Return True
            Return parameters.Contains(paramName)
        End Function

        Friend Shared Function GetParameterNamesFromCreationContext(node As SyntaxNode) As IEnumerable(Of String)
            Dim creationContext As SyntaxNode = node.FirstAncestorOrSelfOfType(GetType(MultiLineLambdaExpressionSyntax),
                                                                 GetType(LambdaExpressionSyntax),
                                                                 GetType(AccessorBlockSyntax),
                                                                 GetType(MethodBlockSyntax),
                                                                 GetType(ConstructorBlockSyntax))
            Return GetParameterNames(creationContext)
        End Function

        Friend Shared Function GetParameterNames(node As SyntaxNode) As IEnumerable(Of String)
            If node Is Nothing Then Return Nothing

            Dim method As MethodBlockSyntax = TryCast(node, MethodBlockSyntax)
            If method IsNot Nothing Then
                Return method.SubOrFunctionStatement.ParameterList.Parameters.Select(Function(p) p.Identifier.ToString())
            End If

            Dim simpleLambda As LambdaExpressionSyntax = TryCast(node, LambdaExpressionSyntax)
            If simpleLambda IsNot Nothing Then
                Return simpleLambda.SubOrFunctionHeader.ParameterList.Parameters.Select(Function(p) p.Identifier.ToString())
            End If

            Dim accessor As AccessorBlockSyntax = TryCast(node, AccessorBlockSyntax)
            If accessor IsNot Nothing Then
                If accessor.IsKind(SyntaxKind.SetAccessorBlock) Then
                    Return {"value"}
                End If
            End If

            Dim constructor As ConstructorBlockSyntax = TryCast(node, ConstructorBlockSyntax)
            If constructor IsNot Nothing Then
                Return constructor.SubNewStatement.ParameterList.Parameters.Select(Function(p) p.Identifier.ToString())
            End If

            Return Nothing

        End Function
    End Class
End Namespace