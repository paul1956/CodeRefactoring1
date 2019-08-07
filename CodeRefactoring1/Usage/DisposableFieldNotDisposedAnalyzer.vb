Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Usage
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class DisposableFieldNotDisposedAnalyzer
        Inherits DiagnosticAnalyzer

        Friend Const Title As String = "Dispose Fields Properly"
        Friend Const MessageFormat As String = "Field {0} implements IDisposable and should be disposed."
        Private Const Description As String = "This class has a disposable field and is not disposing it."

        Friend Shared RuleForReturned As New DiagnosticDescriptor(
            DiagnosticIds.DisposableFieldNotDisposed_ReturnedDiagnosticId,
            Title,
            MessageFormat,
            SupportedCategories.Usage,
            DiagnosticSeverity.Info,
            isEnabledByDefault:=True,
            description:=Description,
            helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.DisposableFieldNotDisposed_ReturnedDiagnosticId))

        Friend Shared RuleForCreated As New DiagnosticDescriptor(
            DiagnosticIds.DisposableFieldNotDisposed_CreatedDiagnosticId,
            Title,
            MessageFormat,
            SupportedCategories.Usage,
            DiagnosticSeverity.Warning,
            isEnabledByDefault:=True,
            description:=Description,
            helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.DisposableFieldNotDisposed_CreatedDiagnosticId))


        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(RuleForCreated, RuleForReturned)
            End Get
        End Property

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSymbolAction(AddressOf Me.AnalyzeField, SymbolKind.Field)
        End Sub

        Private Sub AnalyzeField(context As SymbolAnalysisContext)
            If context.IsGenerated() Then Return
            Dim fieldSymbol As IFieldSymbol = DirectCast(context.Symbol, IFieldSymbol)
            If Not fieldSymbol.Type.AllInterfaces.Any(Function(i) i.ToString().EndsWith(NameOf(IDisposable))) AndAlso Not fieldSymbol.Type.ToString().EndsWith(NameOf(IDisposable)) Then Exit Sub
            Dim fieldSyntaxRef As SyntaxReference = fieldSymbol.DeclaringSyntaxReferences.FirstOrDefault
            If fieldSyntaxRef Is Nothing Then Exit Sub
            Dim variableDeclarator As VariableDeclaratorSyntax = TryCast(fieldSyntaxRef.GetSyntax().Parent, VariableDeclaratorSyntax)
            If variableDeclarator Is Nothing Then Exit Sub
            If Me.ContainingTypeImplementsIDisposableAndCallsItOnTheField(context, fieldSymbol, fieldSymbol.ContainingType) Then Exit Sub

            Dim props As ImmutableDictionary(Of String, String) = New Dictionary(Of String, String) From {{"variableIdentifier", variableDeclarator.Names.First().Identifier.ValueText}}.ToImmutableDictionary()

            If variableDeclarator.AsClause.Kind = SyntaxKind.AsNewClause Then
                context.ReportDiagnostic(Diagnostic.Create(RuleForCreated, variableDeclarator.GetLocation(), props, fieldSymbol.Name))
            ElseIf TypeOf (variableDeclarator.Initializer?.Value) Is InvocationExpressionSyntax Then
                context.ReportDiagnostic(Diagnostic.Create(RuleForReturned, variableDeclarator.GetLocation(), props, fieldSymbol.Name))
            ElseIf TypeOf (variableDeclarator.Initializer?.Value) Is ObjectCreationExpressionSyntax Then
                context.ReportDiagnostic(Diagnostic.Create(RuleForCreated, variableDeclarator.GetLocation(), props, fieldSymbol.Name))
            End If
        End Sub

        Private Function ContainingTypeImplementsIDisposableAndCallsItOnTheField(context As SymbolAnalysisContext, fieldSymbol As IFieldSymbol, typeSymbol As INamedTypeSymbol) As Boolean
            If typeSymbol Is Nothing Then Return False
            Dim disposableInterface As INamedTypeSymbol = typeSymbol.AllInterfaces.FirstOrDefault(Function(i) i.ToString().EndsWith(NameOf(IDisposable)))
            If disposableInterface Is Nothing Then Return False

            Dim disposableMethod As IMethodSymbol = disposableInterface.GetMembers("Dispose").OfType(Of IMethodSymbol).FirstOrDefault(Function(d) d.Arity = 0)
            Dim disposeMethodSymbol As IMethodSymbol = DirectCast(typeSymbol.FindImplementationForInterfaceMember(disposableMethod), IMethodSymbol)
            If disposeMethodSymbol Is Nothing Then Return False

            Dim disposeMethod As MethodBlockSyntax = TryCast(disposeMethodSymbol.DeclaringSyntaxReferences.FirstOrDefault()?.GetSyntax().Parent, MethodBlockSyntax)
            If disposeMethod Is Nothing Then Return False
            If disposeMethod.SubOrFunctionStatement.Modifiers.Any(SyntaxKind.MustInheritKeyword) Then Return True
            Dim typeDeclaration As TypeBlockSyntax = DirectCast(typeSymbol.DeclaringSyntaxReferences.FirstOrDefault()?.GetSyntax().Parent, TypeBlockSyntax)
            Dim semanticModel As SemanticModel = context.Compilation.GetSemanticModel(typeDeclaration.SyntaxTree)
            If CallsDisposeOnField(fieldSymbol, disposeMethod, semanticModel) Then Return True

            Return False
        End Function

        Private Shared Function CallsDisposeOnField(fieldSymbol As IFieldSymbol, disposeMethod As MethodBlockSyntax, semanticModel As SemanticModel) As Boolean
            Dim hasDisposeCall As Boolean = disposeMethod.Statements.OfType(Of ExpressionStatementSyntax).
                Any(Function(exp)
                        Dim invocation As InvocationExpressionSyntax = TryCast(exp.Expression, InvocationExpressionSyntax)
                        If Not If(invocation?.Expression?.IsKind(SyntaxKind.SimpleMemberAccessExpression), True) Then Return False
                        If invocation.ArgumentList.Arguments.Any() Then Return False  ' Calling the wrong dispose method
                        Dim memberAccess As MemberAccessExpressionSyntax = DirectCast(invocation.Expression, MemberAccessExpressionSyntax)
                        If memberAccess.Name.Identifier.ToString() <> "Dispose" OrElse memberAccess.Name.Arity <> 0 Then Return False
                        Dim memberAccessIndentifier As IdentifierNameSyntax = TryCast(memberAccess.Expression, IdentifierNameSyntax)
                        If memberAccessIndentifier Is Nothing Then Return False
                        Try
                            Return fieldSymbol.Equals(semanticModel.GetSymbolInfo(memberAccessIndentifier).Symbol)
                        Catch ex As Exception
                            Return True
                        End Try
                    End Function)
            Return hasDisposeCall
        End Function
    End Class
End Namespace