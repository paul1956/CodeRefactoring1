Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Refactoring

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class AllowMembersOrderingAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const Category As String = SupportedCategories.Refactoring
        Private Const MessageFormat As String = "Ordering member inside this type."
        Private Const Title As String = "Ordering member inside this type."

        Friend Shared Rule As New DiagnosticDescriptor(
                                AllowMembersOrderingDiagnosticId,
                                Title,
                                MessageFormat,
                                Category,
                                DiagnosticSeverity.Hidden,
                                isEnabledByDefault:=True,
                                description:=MessageFormat,
                                helpLinkUri:=ForDiagnostic(AllowMembersOrderingDiagnosticId),
                                Array.Empty(Of String))

        Public Shared ReadOnly Id As String = AllowMembersOrderingDiagnosticId

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Public Shared Sub Analyze(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim typeSyntax As TypeBlockSyntax = TryCast(context.Node, TypeBlockSyntax)
            If typeSyntax Is Nothing Then Exit Sub

            Dim currentChildNodesOrder As SyntaxList(Of StatementSyntax) = typeSyntax.Members()
            If currentChildNodesOrder.Count > 1 Then ' If there is only member, we don't need to worry about the order
                context.ReportDiagnostic(Diagnostic.Create(Rule, typeSyntax.GetLocation()))
            End If
        End Sub

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Analyze, SyntaxKind.ClassBlock, SyntaxKind.StructureBlock, SyntaxKind.ModuleBlock)
        End Sub

    End Class

End Namespace