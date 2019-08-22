Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Design

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class NameOfAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const Category As String = SupportedCategories.Design

        Private Const Description As String = "In VB14 or later the NameOf() operator should be used to specify the name of a program element instead of a string literal as it produces code that is easier to refactor."

        Private Const MessageFormat As String = "Use 'NameOf({0})' instead of specifying the program element name."

        Private Const Title As String = "You should use nameof instead of the parameter element name string"

        Protected Shared Rule As New DiagnosticDescriptor(
                                                                NameOfDiagnosticId,
                                Title,
                                MessageFormat,
                                Category,
                                DiagnosticSeverity.Warning,
                                isEnabledByDefault:=True,
                                Description,
                                helpLinkUri:=ForDiagnostic(NameOfDiagnosticId),
                                Array.Empty(Of String)
                                )

        Protected Shared RuleExtenal As New DiagnosticDescriptor(
                        NameOf_ExternalDiagnosticId,
                        Title,
                        MessageFormat,
                        Category,
                        DiagnosticSeverity.Warning,
                        isEnabledByDefault:=True,
                        Description,
                        helpLinkUri:=ForDiagnostic(NameOf_ExternalDiagnosticId),
                        Array.Empty(Of String))

        Public Overrides ReadOnly Property SupportedDiagnostics() As ImmutableArray(Of DiagnosticDescriptor) = ImmutableArray.Create(Rule, RuleExtenal)

        Private Shared Function Found(programElement As String) As Boolean
            Return Not String.IsNullOrEmpty(programElement)
        End Function

        Private Shared Function GetParameterNameThatMatchesStringLiteral(stringLiteral As LiteralExpressionSyntax) As String
            Dim ancestorThatMightHaveParameters As SyntaxNode = stringLiteral.FirstAncestorOfType(GetType(AttributeListSyntax), GetType(MethodBlockSyntax), GetType(SubNewStatementSyntax), GetType(InvocationExpressionSyntax))
            Dim parameterName As String = String.Empty
            If ancestorThatMightHaveParameters IsNot Nothing Then
                Dim parameters As SeparatedSyntaxList(Of ParameterSyntax) = New SeparatedSyntaxList(Of ParameterSyntax)()
                Select Case ancestorThatMightHaveParameters.Kind
                    Case SyntaxKind.SubBlock, SyntaxKind.FunctionBlock
                        Dim method As MethodBlockSyntax = DirectCast(ancestorThatMightHaveParameters, MethodBlockSyntax)
                        If method.SubOrFunctionStatement.ParameterList Is Nothing Then
                            Return Nothing
                        End If
                        parameters = method.SubOrFunctionStatement.ParameterList.Parameters
                    Case SyntaxKind.AttributeList
                End Select
                parameterName = GetParameterWithIdentifierEqualToStringLiteral(stringLiteral, parameters)?.Identifier.Identifier.Text

            End If
            Return parameterName
        End Function

        Private Shared Function GetParameterWithIdentifierEqualToStringLiteral(stringLiteral As LiteralExpressionSyntax, parameters As SeparatedSyntaxList(Of ParameterSyntax)) As ParameterSyntax
            Return parameters.FirstOrDefault(Function(m As ParameterSyntax) String.Equals(m.Identifier.Identifier.Text, stringLiteral.Token.ValueText, StringComparison.Ordinal))
        End Function

        Private Shared Function GetProgramElementNameThatMatchesStringLiteral(stringLiteral As LiteralExpressionSyntax, model As SemanticModel, ByRef externalSymbol As Boolean) As String
            Dim programElementName As String = GetParameterNameThatMatchesStringLiteral(stringLiteral)
            If Not Found(programElementName) Then
                Dim literalValueText As String = stringLiteral.Token.ValueText
                Dim symbol As ISymbol = model.LookupSymbols(stringLiteral.Token.SpanStart, Nothing, literalValueText).FirstOrDefault()
                If symbol Is Nothing Then Return Nothing
                externalSymbol = symbol.Locations.Any(Function(l As Location) l.IsInSource) = False
                If symbol.Kind = SymbolKind.Local Then
                    ' Only register if local variable is declared before it is used.
                    ' Don't recommend if variable is declared after string literal is used.
                    Dim symbolSpan As Text.TextSpan = symbol.Locations.Min(Function(i As Location) i.SourceSpan)
                    If symbolSpan.CompareTo(stringLiteral.Token.Span) > 0 Then
                        Return Nothing
                    End If
                End If
                programElementName = symbol?.ToDisplayParts().
                    Where(AddressOf IncludeOnlyPartsThatAreName).
                    LastOrDefault(Function(displayPart As SymbolDisplayPart) displayPart.ToString() = literalValueText).
                    ToString()
            End If
            ' special case I think this should be unnecessary
            If programElementName = "Func" Then
                Return Nothing
            End If
            Return programElementName
        End Function

        Private Sub Analyzer(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim stringLiteral As LiteralExpressionSyntax = DirectCast(context.Node, LiteralExpressionSyntax)
            If String.IsNullOrWhiteSpace(stringLiteral?.Token.ValueText) Then Return

            Dim externalSymbol As Boolean = False
            Dim programElementName As String = GetProgramElementNameThatMatchesStringLiteral(stringLiteral, context.SemanticModel, externalSymbol)
            If (Found(programElementName)) Then
                Dim diag As Diagnostic = Diagnostic.Create(If(externalSymbol, RuleExtenal, Rule), stringLiteral.GetLocation(), programElementName)
                context.ReportDiagnostic(diag)
            End If
        End Sub

        Public Shared Function IncludeOnlyPartsThatAreName(displayPart As SymbolDisplayPart) As Boolean
            Return displayPart.IsAnyKind(SymbolDisplayPartKind.ClassName, SymbolDisplayPartKind.DelegateName, SymbolDisplayPartKind.EnumName, SymbolDisplayPartKind.EventName, SymbolDisplayPartKind.FieldName, SymbolDisplayPartKind.InterfaceName, SymbolDisplayPartKind.LocalName, SymbolDisplayPartKind.MethodName, SymbolDisplayPartKind.NamespaceName, SymbolDisplayPartKind.ParameterName, SymbolDisplayPartKind.PropertyName, SymbolDisplayPartKind.StructName)
        End Function

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(LanguageVersion.VisualBasic14, AddressOf Me.Analyzer, SyntaxKind.StringLiteralExpression)
        End Sub

    End Class

End Namespace