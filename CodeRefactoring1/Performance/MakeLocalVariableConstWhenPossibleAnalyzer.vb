Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On
Imports Microsoft.CodeAnalysis.Diagnostics
Imports System.Collections.Immutable

Namespace Performance
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class MakeLocalVariableConstWhenPossibleAnalyzer
        Inherits DiagnosticAnalyzer

        Public Shared ReadOnly Id As String = DiagnosticIds.MakeLocalVariableConstWhenItIsPossibleDiagnosticId
        Public Const Title As String = "Make Local Variable Constant."
        Public Const MessageFormat As String = "This variable can be made const."
        Public Const Category As String = SupportedCategories.Performance
        Public Const Description As String = "If this variable is assigned a constant value and never changed, it can be made 'const'."
        Protected Shared Rule As DiagnosticDescriptor = New DiagnosticDescriptor(
             DiagnosticIds.MakeLocalVariableConstWhenItIsPossibleDiagnosticId,
            Title,
            MessageFormat,
            Category,
            DiagnosticSeverity.Info,
            isEnabledByDefault:=True,
            description:=Description,
            helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.MakeLocalVariableConstWhenItIsPossibleDiagnosticId))

        Public Overrides ReadOnly Property SupportedDiagnostics() As ImmutableArray(Of DiagnosticDescriptor) = ImmutableArray.Create(Rule)

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Me.AnalyzeNode, SyntaxKind.LocalDeclarationStatement)
        End Sub

        Private Sub AnalyzeNode(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim localDeclaration As LocalDeclarationStatementSyntax = DirectCast(context.Node, LocalDeclarationStatementSyntax)
            Dim semanticModel As SemanticModel = context.SemanticModel
            If Not localDeclaration.Modifiers.OfType(Of ConstDirectiveTriviaSyntax).Any() AndAlso
            IsDeclarationConstFriendly(localDeclaration, semanticModel) AndAlso
            AreVariablesOnlyWrittenInsideDeclaration(localDeclaration, semanticModel) Then

                Dim diag As Diagnostic = Diagnostic.Create(Rule, localDeclaration.GetLocation())
                context.ReportDiagnostic(diag)
            End If
        End Sub

        Private Shared Function IsDeclarationConstFriendly(declaration As LocalDeclarationStatementSyntax, semanticModel As SemanticModel) As Boolean
            For Each variable As VariableDeclaratorSyntax In declaration.Declarators
                ' In VB an initializer can either be
                ' infered with an ititializer or declared via
                ' As New ReferenceType
                If variable.Initializer Is Nothing Then Return False

                ' is constant?
                If declaration.Modifiers.Any(SyntaxKind.ConstKeyword) Then Return False

                Dim constantValue As [Optional](Of Object) = semanticModel.GetConstantValue(variable.Initializer.Value)
                Dim valueIsConstant As Boolean = constantValue.HasValue
                If Not valueIsConstant Then Return False

                Dim variableConvertedType As ITypeSymbol

                ' is declared as null reference type
                If variable.AsClause IsNot Nothing Then
                    Dim variableType As VisualBasic.Syntax.TypeSyntax = variable.AsClause.Type
                    variableConvertedType = semanticModel.GetTypeInfo(variableType).ConvertedType
                Else
                    Dim symbol As ISymbol = semanticModel.GetDeclaredSymbol(variable.Names.First())
                    variableConvertedType = DirectCast(symbol, ILocalSymbol).Type
                End If

                If variableConvertedType.IsReferenceType AndAlso
                variableConvertedType.SpecialType <> SpecialType.System_String AndAlso
                constantValue.Value IsNot Nothing Then Return False

                ' Nullable?
                If variableConvertedType.OriginalDefinition?.SpecialType = SpecialType.System_Nullable_T Then Return False
                If variable.Initializer.Value.Kind = SyntaxKind.NothingLiteralExpression Then Return True

                ' Value can be converted to variable type?
                Dim conversion As VisualBasic.Conversion = semanticModel.ClassifyConversion(variable.Initializer.Value, variableConvertedType)
                If (Not conversion.Exists OrElse conversion.IsUserDefined) Then Return False

            Next
            Return True
        End Function

        Private Shared Function AreVariablesOnlyWrittenInsideDeclaration(declaration As LocalDeclarationStatementSyntax, SemanticModel As SemanticModel) As Boolean
            Dim dfa As DataFlowAnalysis = SemanticModel.AnalyzeDataFlow(declaration)
            Dim symbols As IEnumerable(Of ISymbol) = From declarator In declaration.Declarators
                                                     From variable In declarator.Names
                                                     Select SemanticModel.GetDeclaredSymbol(variable)

            Dim result As Boolean = Not symbols.Any(Function(s As ISymbol) dfa.WrittenOutside.Contains(s))
            Return result
        End Function
    End Class
End Namespace
