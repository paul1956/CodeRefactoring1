Imports System.Collections.Immutable
Imports System.Reflection
Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Usage
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class JsonNetAnalyzer
        Inherits DiagnosticAnalyzer

        Friend Const Title As String = "Your JSON syntax is wrong."
        Friend Const MessageFormat As String = "{0}"
        Private Const Description As String = "This diagnostic checks the Json string and triggers if the parsing fails by throwing an exception"

        Friend Shared Rule As New DiagnosticDescriptor(
            DiagnosticIds.JsonNetDiagnosticId,
            Title,
            MessageFormat,
            SupportedCategories.Usage,
            DiagnosticSeverity.Error,
            isEnabledByDefault:=True,
            description:=Description,
            helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.JsonNetDiagnosticId))

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(Sub(c) Me.Analyze(c, "DeserializeObject", "Public Shared Overloads Function DeserializeObject(Of T)(value As String) As T"), SyntaxKind.InvocationExpression)
            context.RegisterSyntaxNodeAction(Sub(c) Me.Analyze(c, "Parse", "Public Shared Overloads Function Parse(json As String) As Newtonsoft.Json.Linq.JObject"), SyntaxKind.InvocationExpression)
        End Sub

        Private Sub Analyze(context As SyntaxNodeAnalysisContext, methodName As String, methodFullDefinition As String)
            If (context.Node.IsGenerated()) Then Return
            Dim invocationExpression As InvocationExpressionSyntax = DirectCast(context.Node, InvocationExpressionSyntax)
            Dim memberExpression As MemberAccessExpressionSyntax = TryCast(invocationExpression.Expression, MemberAccessExpressionSyntax)
            If memberExpression?.Name?.Identifier.ValueText <> methodName Then Exit Sub

            Dim memberSymbol As ISymbol = context.SemanticModel.GetSymbolInfo(memberExpression).Symbol
            If memberSymbol?.OriginalDefinition?.ToString() <> methodFullDefinition Then Exit Sub

            Dim argumentList As ArgumentListSyntax = invocationExpression.ArgumentList
            If If(argumentList?.Arguments.Count, 0) <> 1 Then Exit Sub

            Dim literalParameter As LiteralExpressionSyntax = TryCast(argumentList.Arguments(0).GetExpression(), LiteralExpressionSyntax)
            If literalParameter Is Nothing Then Exit Sub

            Dim jsonOpt As [Optional](Of Object) = context.SemanticModel.GetConstantValue(literalParameter)
            Dim json As String = jsonOpt.Value.ToString()

            CheckJsonValue(context, literalParameter, json)
        End Sub

        Private Shared Sub CheckJsonValue(context As SyntaxNodeAnalysisContext, literalParameter As LiteralExpressionSyntax, json As String)
            Try
                parseMethodInfo.Value.Invoke(Nothing, {json})
            Catch ex As Exception
                Dim diag As Diagnostic = Diagnostic.Create(Rule, literalParameter.GetLocation(), ex.InnerException.Message)
                context.ReportDiagnostic(diag)
            End Try
        End Sub

        Private Shared ReadOnly jObjectType As New Lazy(Of Type)(Function() System.Type.GetType("Newtonsoft.Json.Linq.JObject, Newtonsoft.Json"))
        Private Shared ReadOnly parseMethodInfo As New Lazy(Of MethodInfo)(Function() jObjectType.Value.GetRuntimeMethod("Parse", {GetType(String)}))
    End Class
End Namespace
