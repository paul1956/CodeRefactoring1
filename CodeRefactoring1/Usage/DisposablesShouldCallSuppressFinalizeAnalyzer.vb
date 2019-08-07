Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Diagnostics


Namespace Usage
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class DisposablesShouldCallSuppressFinalizeAnalyzer
        Inherits DiagnosticAnalyzer

        Friend Const Title As String = "Disposables Should Call SuppressFinalize"
        Friend Const MessageFormat As String = "'{0}' should call GC.SuppressFinalize inside the Dispose method."
        Private Const Description As String = "Classes implementing IDisposable should call the GC.SuppressFinalize method  to avoid any finalizer from being called.
This rule should be followed even if the class doesn't have a finalizer in a derived class."

        Friend Shared Rule As New DiagnosticDescriptor(
        DiagnosticIds.DisposablesShouldCallSuppressFinalizeDiagnosticId,
        Title,
        MessageFormat,
        SupportedCategories.Naming,
        DiagnosticSeverity.Warning,
        isEnabledByDefault:=True,
        description:=Description,
        helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.DisposablesShouldCallSuppressFinalizeDiagnosticId))

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSymbolAction(AddressOf Me.Analyze, SymbolKind.NamedType)
        End Sub

        Public Sub Analyze(context As SymbolAnalysisContext)
            If context.IsGenerated() Then Return
            Dim symbol As INamedTypeSymbol = DirectCast(context.Symbol, INamedTypeSymbol)
            If symbol.TypeKind <> TypeKind.Class Then Exit Sub
            If Not symbol.Interfaces.Any(Function(i) i.SpecialType = SpecialType.System_IDisposable) Then Exit Sub

            If symbol.IsSealed AndAlso Not ContainsUserDefinedFinalizer(symbol) Then Exit Sub

            If Not ContainsNonPrivateConstructors(symbol) Then Exit Sub

            Dim disposeMethod As ISymbol = FindDisposeMethod(symbol)
            If disposeMethod Is Nothing Then Exit Sub

            Dim syntaxTree As SyntaxNode = disposeMethod.DeclaringSyntaxReferences(0)?.GetSyntax(context.CancellationToken)
            Dim methodBlock As MethodBlockSyntax = TryCast(TryCast(syntaxTree, MethodStatementSyntax)?.Parent, MethodBlockSyntax)
            Dim statements As IEnumerable(Of ExpressionStatementSyntax) = methodBlock?.Statements.OfType(Of ExpressionStatementSyntax)

            If statements IsNot Nothing Then
                For Each statement As ExpressionStatementSyntax In statements
                    Dim invocation As InvocationExpressionSyntax = TryCast(statement.Expression, InvocationExpressionSyntax)
                    Dim method As MemberAccessExpressionSyntax = TryCast(invocation?.Expression, MemberAccessExpressionSyntax)
                    Dim identifier As IdentifierNameSyntax = TryCast(method?.Expression, IdentifierNameSyntax)
                    If identifier IsNot Nothing Then
                        If NameOf(GC).Equals(identifier.Identifier.ToString, StringComparison.OrdinalIgnoreCase) AndAlso "SuppressFinalize".Equals(method.Name.ToString(), StringComparison.OrdinalIgnoreCase) Then
                            Exit Sub
                        End If
                    End If
                Next
            End If
            context.ReportDiagnostic(Diagnostic.Create(Rule, disposeMethod.Locations(0), symbol.Name))
        End Sub

        Private Shared Function FindDisposeMethod(symbol As INamedTypeSymbol) As ISymbol
            Return symbol.GetMembers().
                Where(Function(x) x.ToString().IndexOf("Dispose", StringComparison.OrdinalIgnoreCase) > 0).OfType(Of IMethodSymbol).
                FirstOrDefault(Function(m) m.Parameters = Nothing Or m.Parameters.Count = 0)
        End Function

        Private Shared Function ContainsUserDefinedFinalizer(symbol As INamedTypeSymbol) As Boolean
            Return symbol.GetMembers().Any(Function(x) x.ToString().IndexOf(NameOf(Finalize), StringComparison.OrdinalIgnoreCase) > 0)
        End Function

        Private Shared Function ContainsNonPrivateConstructors(symbol As INamedTypeSymbol) As Boolean
            If IsNestedPrivateType(symbol) Then Return False

            Return symbol.GetMembers().
                Any(Function(m) m.MetadataName = ".ctor" AndAlso m.DeclaredAccessibility <> Accessibility.Private)
        End Function

        Private Shared Function IsNestedPrivateType(symbol As INamedTypeSymbol) As Boolean
            If symbol Is Nothing Then Return False
            If symbol.DeclaredAccessibility = Accessibility.Private AndAlso symbol.ContainingType IsNot Nothing Then Return True
            Return IsNestedPrivateType(symbol.ContainingType)
        End Function
    End Class
End Namespace
