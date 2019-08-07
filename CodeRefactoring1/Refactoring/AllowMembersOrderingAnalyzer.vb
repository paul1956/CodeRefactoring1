Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Refactoring
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class AllowMembersOrderingAnalyzer
        Inherits DiagnosticAnalyzer

        Public Shared ReadOnly Id As String = DiagnosticIds.AllowMembersOrderingDiagnosticId

        Friend Const Title As String = "Ordering member inside this type."
        Friend Const MessageFormat As String = "Ordering member inside this type."
        Friend Const Category As String = SupportedCategories.Refactoring
        Friend Shared Rule As New DiagnosticDescriptor(
            DiagnosticIds.AllowMembersOrderingDiagnosticId,
            Title,
            MessageFormat,
            Category,
            DiagnosticSeverity.Hidden,
            isEnabledByDefault:=True,
            helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.AllowMembersOrderingDiagnosticId))

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Me.Analyze, SyntaxKind.ClassBlock, SyntaxKind.StructureBlock, SyntaxKind.ModuleBlock)
        End Sub

        Public Sub Analyze(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim typeSyntax As TypeBlockSyntax = TryCast(context.Node, TypeBlockSyntax)
            If typeSyntax Is Nothing Then Exit Sub

            Dim currentChildNodesOrder As SyntaxList(Of StatementSyntax) = typeSyntax.Members()
            If currentChildNodesOrder.Count > 1 Then ' If there is only member, we don't need to worry about the order
                context.ReportDiagnostic(Diagnostic.Create(Rule, typeSyntax.GetLocation()))
            End If
        End Sub
    End Class
End Namespace