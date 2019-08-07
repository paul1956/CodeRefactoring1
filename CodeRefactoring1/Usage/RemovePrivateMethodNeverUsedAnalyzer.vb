Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Usage
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class RemovePrivateMethodNeverUsedAnalyzer
        Inherits DiagnosticAnalyzer

        Friend Const Title As String = "Unused Method"
        Friend Const Message As String = "Method is not used."
        Private Const Description As String = "When a private method is declared but not used, remove it to avoid confusion."

        Friend Shared Rule As New DiagnosticDescriptor(
            DiagnosticIds.RemovePrivateMethodNeverUsedDiagnosticId,
            Title,
            Message,
            SupportedCategories.Usage,
            DiagnosticSeverity.Info,
            isEnabledByDefault:=True,
            description:=Description,
            helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.RemovePrivateMethodNeverUsedDiagnosticId))

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Me.AnalyzeNode, SyntaxKind.SubStatement, SyntaxKind.FunctionStatement)
        End Sub

        Private Sub AnalyzeNode(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim methodStatement As MethodStatementSyntax = DirectCast(context.Node, MethodStatementSyntax)
            If methodStatement.HandlesClause IsNot Nothing Then Exit Sub
            If Not methodStatement.Modifiers.Any(Function(a) a.ValueText = SyntaxFactory.Token(SyntaxKind.PrivateKeyword).ValueText) Then Exit Sub
            If (Me.IsMethodAttributeAnException(methodStatement)) Then Return
            If Me.IsMethodUsed(methodStatement, context.SemanticModel) Then Exit Sub
            Dim props As ImmutableDictionary(Of String, String) = New Dictionary(Of String, String) From {{"identifier", methodStatement.Identifier.Text}}.ToImmutableDictionary()
            Dim diag As Diagnostic = Diagnostic.Create(Rule, methodStatement.GetLocation(), props)
            context.ReportDiagnostic(diag)
        End Sub

        Private Function IsMethodAttributeAnException(methodStatement As MethodStatementSyntax) As Boolean
            For Each attributeList As AttributeListSyntax In methodStatement.AttributeLists
                For Each attribute As AttributeSyntax In attributeList.Attributes
                    Dim identifierName As IdentifierNameSyntax = TryCast(attribute.Name, IdentifierNameSyntax)
                    Dim nameText As String = Nothing
                    If (identifierName IsNot Nothing) Then
                        nameText = identifierName?.Identifier.Text
                    Else
                        Dim qualifiedName As QualifiedNameSyntax = TryCast(attribute.Name, QualifiedNameSyntax)
                        If (qualifiedName IsNot Nothing) Then
                            nameText = qualifiedName.Right?.Identifier.Text
                        End If
                    End If
                    If (nameText Is Nothing) Then Continue For
                    If (IsExcludedAttributeName(nameText)) Then Return True
                Next
            Next
            Return False
        End Function

        Private Shared ReadOnly excludedAttributeNames As String() = {"Fact", "ContractInvariantMethod", "DataMember"}

        Private Shared Function IsExcludedAttributeName(attributeName As String) As Boolean
            Return excludedAttributeNames.Contains(attributeName)
        End Function

        Private Function IsMethodUsed(methodTarget As MethodStatementSyntax, semanticModel As SemanticModel) As Boolean
            Dim typeDeclaration As ClassBlockSyntax = TryCast(methodTarget.Parent.Parent, ClassBlockSyntax)
            If typeDeclaration Is Nothing Then Return True

            Dim classStatement As ClassStatementSyntax = typeDeclaration.ClassStatement
            If classStatement Is Nothing Then Return True

            If Not classStatement.Modifiers.Any(SyntaxKind.PartialKeyword) Then
                Return Me.IsMethodUsed(methodTarget, typeDeclaration)
            End If

            Dim symbol As INamedTypeSymbol = semanticModel.GetDeclaredSymbol(typeDeclaration)

            Return symbol Is Nothing OrElse symbol.DeclaringSyntaxReferences.Any(Function(r) Me.IsMethodUsed(methodTarget, r.GetSyntax().Parent))
        End Function

        Private Function IsMethodUsed(methodTarget As MethodStatementSyntax, typeDeclaration As SyntaxNode) As Boolean
            Dim hasIdentifier As IEnumerable(Of VisualBasic.Syntax.IdentifierNameSyntax) = typeDeclaration?.DescendantNodes()?.OfType(Of IdentifierNameSyntax)()
            If (hasIdentifier Is Nothing OrElse Not hasIdentifier.Any()) Then Return False
            Return hasIdentifier.Any(Function(a) a IsNot Nothing AndAlso a.Identifier.ValueText.Equals(methodTarget?.Identifier.ValueText))
        End Function

    End Class
End Namespace


