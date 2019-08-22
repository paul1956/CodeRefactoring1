Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Style

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class TernaryOperatorAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const MessageFormat As String = "You can use a ternary operator."
        Private Const Title As String = "Use Ternary operator."

        Private Shared ReadOnly RuleForIfWithAssignment As New DiagnosticDescriptor(
                                TernaryOperator_AssignmentDiagnosticId,
                                Title,
                                MessageFormat,
                                SupportedCategories.Style,
                                DiagnosticSeverity.Warning,
                                isEnabledByDefault:=True,
                                description:=MessageFormat,
                                helpLinkUri:=ForDiagnostic(TernaryOperator_AssignmentDiagnosticId),
                                Array.Empty(Of String)
                                )

        Friend Shared RuleForIfWithReturn As New DiagnosticDescriptor(
                                TernaryOperator_ReturnDiagnosticId,
                                Title,
                                MessageFormat,
                                SupportedCategories.Style,
                                DiagnosticSeverity.Warning,
                                isEnabledByDefault:=True,
                                description:=MessageFormat,
                                helpLinkUri:=ForDiagnostic(TernaryOperator_ReturnDiagnosticId),
                                Array.Empty(Of String)
                                )

        Friend Shared RuleForIif As New DiagnosticDescriptor(
                                TernaryOperator_IifDiagnosticId,
                                Title,
                                MessageFormat,
                                SupportedCategories.Style,
                                DiagnosticSeverity.Warning,
                                isEnabledByDefault:=True,
                                description:=MessageFormat,
                                helpLinkUri:=ForDiagnostic(TernaryOperator_IifDiagnosticId),
                                Array.Empty(Of String)
                                )

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(RuleForIfWithAssignment, RuleForIfWithReturn, RuleForIif)
            End Get
        End Property

        Private Sub IfAnalyzer(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim ifStatement As IfStatementSyntax = TryCast(context.Node, IfStatementSyntax)
            If ifStatement Is Nothing Then Exit Sub
            Dim ifBlock As MultiLineIfBlockSyntax = TryCast(ifStatement.Parent, MultiLineIfBlockSyntax)
            If ifBlock Is Nothing Then Exit Sub

            ' Can't handle elseif clauses in ternary conditional
            If ifBlock.ElseIfBlocks.Any() Then Exit Sub

            If ifBlock.ElseBlock Is Nothing Then Exit Sub
            If ifBlock.Statements.Count <> 1 OrElse ifBlock.ElseBlock.Statements.Count <> 1 Then Exit Sub

            Dim ifClauseStatement As StatementSyntax = ifBlock.Statements(0)
            Dim elseStatement As StatementSyntax = ifBlock.ElseBlock.Statements(0)

            If TypeOf (ifClauseStatement) Is ReturnStatementSyntax AndAlso
            TypeOf (elseStatement) Is ReturnStatementSyntax Then
                If ifClauseStatement.ToString.Length > 60 Then Exit Sub
                If elseStatement.ToString.Length > 60 Then Exit Sub
                Dim diag As Diagnostic = Diagnostic.Create(RuleForIfWithReturn, ifStatement.IfKeyword.GetLocation, "You can use a ternary operator.")
                context.ReportDiagnostic(diag)
                Exit Sub
            End If

            Dim ifAssignment As AssignmentStatementSyntax = TryCast(ifClauseStatement, AssignmentStatementSyntax)
            Dim elseAssignment As AssignmentStatementSyntax = TryCast(elseStatement, AssignmentStatementSyntax)
            If ifAssignment Is Nothing OrElse elseAssignment Is Nothing Then Exit Sub

            If ifAssignment.ToString.Length + elseAssignment.ToString.Length > 80 Then Exit Sub

            Dim semanticModel As SemanticModel = context.SemanticModel
            Dim ifSymbol As ISymbol = semanticModel.GetSymbolInfo(ifAssignment.Left).Symbol
            Dim elseSymbol As ISymbol = semanticModel.GetSymbolInfo(elseAssignment.Left).Symbol
            If ifSymbol Is Nothing OrElse elseSymbol Is Nothing OrElse ifSymbol.Equals(elseSymbol) = False Then Exit Sub

            Dim assignDiag As Diagnostic = Diagnostic.Create(RuleForIfWithAssignment, ifStatement.IfKeyword.GetLocation(), "You can use a ternary operator.")
            context.ReportDiagnostic(assignDiag)
        End Sub

        Private Sub IIfAnalyzer(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim iifStatement As InvocationExpressionSyntax = TryCast(context.Node, InvocationExpressionSyntax)
            If iifStatement Is Nothing Then Exit Sub
            Dim iifExpression As IdentifierNameSyntax = TryCast(iifStatement.Expression, IdentifierNameSyntax)
            If iifExpression Is Nothing Then Exit Sub
            If String.Equals(iifExpression.Identifier.ValueText, "iif", StringComparison.OrdinalIgnoreCase) Then
                If iifStatement.ArgumentList.Arguments.Count = 3 Then
                    Dim diag As Diagnostic = Diagnostic.Create(RuleForIif, iifStatement.GetLocation())
                    context.ReportDiagnostic(diag)
                End If
            End If
        End Sub

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Me.IfAnalyzer, SyntaxKind.IfStatement)
            context.RegisterSyntaxNodeAction(AddressOf Me.IIfAnalyzer, SyntaxKind.InvocationExpression)
        End Sub

    End Class

End Namespace