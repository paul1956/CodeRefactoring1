Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Reliability

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class UseConfigureAwaitFalseAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const Category As String = SupportedCategories.Reliability
        Private Const MessageFormat As String = "Consider using ConfigureAwait(False) on the awaited task."
        Private Const Title As String = "Use ConfigureAwait(False) on awaited task."

        Protected Shared Rule As New DiagnosticDescriptor(
                        UseConfigureAwaitFalseDiagnosticId,
                        Title,
                        MessageFormat,
                        Category,
                        DiagnosticSeverity.Hidden,
                        isEnabledByDefault:=True,
                        description:=MessageFormat,
                        helpLinkUri:=ForDiagnostic(UseConfigureAwaitFalseDiagnosticId),
                        Array.Empty(Of String)
                        )

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Private Shared Sub AnalyzeNode(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim awaitExpression As AwaitExpressionSyntax = DirectCast(context.Node, AwaitExpressionSyntax)
            Dim awaitedExpression As ExpressionSyntax = awaitExpression.Expression
            If Not IsTask(awaitedExpression, context) Then Exit Sub

            Dim diag As Diagnostic = Diagnostic.Create(Rule, awaitExpression.GetLocation())
            context.ReportDiagnostic(diag)
        End Sub

        Private Shared Function IsTask(expression As ExpressionSyntax, context As SyntaxNodeAnalysisContext) As Boolean
            Dim type As INamedTypeSymbol = TryCast(context.SemanticModel.GetTypeInfo(expression).Type, INamedTypeSymbol)
            If type Is Nothing Then Return False
            Dim taskType As INamedTypeSymbol
            If type.IsGenericType Then
                type = type.ConstructedFrom
                taskType = context.SemanticModel.Compilation.GetTypeByMetadataName("System.Threading.Tasks.Task`1")
            Else
                taskType = context.SemanticModel.Compilation.GetTypeByMetadataName("System.Threading.Tasks.Task")
            End If
            Return type.Equals(taskType)
        End Function

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf AnalyzeNode, SyntaxKind.AwaitExpression)
        End Sub

    End Class

End Namespace