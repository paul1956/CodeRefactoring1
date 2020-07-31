' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Performance

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class MakeLocalVariableConstWhenPossibleAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const Category As String = SupportedCategories.Performance
        Private Const Description As String = "If this variable is assigned a constant value and never changed, it can be made 'const'."
        Private Const MessageFormat As String = "This variable can be made const."
        Private Const Title As String = "Make Local Variable Constant."

        Protected Shared Rule As New DiagnosticDescriptor(
                            MakeLocalVariableConstWhenItIsPossibleDiagnosticId,
                            Title,
                            MessageFormat,
                            Category,
                            DiagnosticSeverity.Info,
                            isEnabledByDefault:=True,
                            Description,
                            helpLinkUri:=ForDiagnostic(MakeLocalVariableConstWhenItIsPossibleDiagnosticId),
                            Array.Empty(Of String)
                            )

        Public Shared ReadOnly Id As String = MakeLocalVariableConstWhenItIsPossibleDiagnosticId
        Public Overrides ReadOnly Property SupportedDiagnostics() As ImmutableArray(Of DiagnosticDescriptor) = ImmutableArray.Create(Rule)

        Private Shared Function AreVariablesOnlyWrittenInsideDeclaration(declaration As LocalDeclarationStatementSyntax, SemanticModel As SemanticModel) As Boolean
            Dim dfa As DataFlowAnalysis = SemanticModel.AnalyzeDataFlow(declaration)
            Dim symbols As IEnumerable(Of ISymbol) = From declarator In declaration.Declarators
                                                     From variable In declarator.Names
                                                     Select SemanticModel.GetDeclaredSymbol(variable)

            Dim result As Boolean = Not symbols.Any(Function(s As ISymbol) dfa.WrittenOutside.Contains(s))
            Return result
        End Function

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

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf AnalyzeNode, SyntaxKind.LocalDeclarationStatement)
        End Sub

    End Class

End Namespace
