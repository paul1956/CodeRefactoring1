Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Usage

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class UnusedParametersAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const Description As String = "When a method declares a parameter and does not use it might bring incorrect conclusions for anyone reading the code and also demands the parameter when the method is called unnecessarily.
You should delete the parameter in such cases."

        Private Const Message As String = "Parameter '{0}' is not used."
        Private Const Title As String = "Unused parameters."

        Protected Shared Rule As New DiagnosticDescriptor(
        UnusedParametersDiagnosticId,
        Title,
        Message,
        SupportedCategories.Usage,
        DiagnosticSeverity.Warning,
        True,
        Description,
        ForDiagnostic(UnusedParametersDiagnosticId),
        WellKnownDiagnosticTags.Unnecessary)

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Private Shared Function CreateDiagnostic(context As SyntaxNodeAnalysisContext, parameter As ParameterSyntax) As SyntaxNodeAnalysisContext
            Dim propsDic As Dictionary(Of String, String) = New Dictionary(Of String, String) From {
                {"identifier", parameter.Identifier.Identifier.Text}
            }
            Dim props As ImmutableDictionary(Of String, String) = propsDic.ToImmutableDictionary()
            Dim diag As Diagnostic = Diagnostic.Create(Rule, parameter.GetLocation(), props, parameter.Identifier.Identifier.ValueText)
            context.ReportDiagnostic(diag)
            Return context
        End Function

        Private Shared Function IsCandidateForRemoval(methodOrConstructor As MethodBlockBaseSyntax, semanticModel As SemanticModel) As Boolean
            If methodOrConstructor.BlockStatement.Modifiers.Any(Function(m As SyntaxToken) m.ValueText = "Partial" OrElse m.ValueText = "Overrides") OrElse
            Not methodOrConstructor.BlockStatement.ParameterList?.Parameters.Any() Then Return False
            If methodOrConstructor.HasAttributeOnAncestorOrSelf("DllImport") Then Return False

            Dim method As MethodBlockSyntax = TryCast(methodOrConstructor, MethodBlockSyntax)
            If method IsNot Nothing Then
                If method.SubOrFunctionStatement.ImplementsClause IsNot Nothing Then Return False
                Dim methodSymbol As IMethodSymbol = semanticModel.GetDeclaredSymbol(method)
                If methodSymbol Is Nothing Then Return False
                Dim typeSymbol As INamedTypeSymbol = methodSymbol.ContainingType
                If typeSymbol.Interfaces.SelectMany(Function(i As INamedTypeSymbol) i.GetMembers()).
                Any(Function(member As ISymbol) methodSymbol.Equals(typeSymbol.FindImplementationForInterfaceMember(member))) Then Return False

                If IsEventHandlerLike(method, semanticModel) Then Return False
            Else
                Dim constructor As ConstructorBlockSyntax = TryCast(methodOrConstructor, ConstructorBlockSyntax)
                If constructor IsNot Nothing Then
                    If IsSerializationConstructor(constructor, semanticModel) Then Return False
                Else
                    Return False
                End If
            End If
            Return True
        End Function

        Private Shared Function IsEventHandlerLike(method As MethodBlockSyntax, model As SemanticModel) As Boolean

            If method.SubOrFunctionStatement.ParameterList Is Nothing Then
                Return False
            End If

            If method.SubOrFunctionStatement.ParameterList.Parameters.Count <> 2 OrElse method.IsKind(SyntaxKind.FunctionBlock) Then
                Return False
            End If

            Dim senderType As ITypeSymbol = model.GetTypeInfo(method.SubOrFunctionStatement.ParameterList.Parameters(0).AsClause.Type).Type
            If senderType.SpecialType <> SpecialType.System_Object Then Return False
            Dim eventArgsType As INamedTypeSymbol = TryCast(model.GetTypeInfo(method.SubOrFunctionStatement.ParameterList.Parameters(1).AsClause.Type).Type, INamedTypeSymbol)
            If eventArgsType Is Nothing Then Return False
            Return eventArgsType.AllBaseTypesAndSelf().Any(Function(type As INamedTypeSymbol) type.ToString() = "System.EventArgs")
        End Function

        Private Shared Function IsSerializationConstructor(constructor As ConstructorBlockSyntax, model As SemanticModel) As Boolean
            If constructor.SubNewStatement.ParameterList.Parameters.Count <> 2 Then Return False
            Dim constructorSymbol As IMethodSymbol = model.GetDeclaredSymbol(constructor)
            Dim typeSymbol As INamedTypeSymbol = constructorSymbol?.ContainingType
            If If(Not typeSymbol?.Interfaces.Any(Function(i As INamedTypeSymbol) i.ToString() = "System.Runtime.Serialization.ISerializable"), True) Then Return False
            If Not typeSymbol.GetAttributes().Any(Function(a As AttributeData) a.AttributeClass.ToString() = "System.SerializableAttribute") Then Return False
            Dim serializationInfoType As INamedTypeSymbol = TryCast(model.GetTypeInfo(constructor.SubNewStatement.ParameterList.Parameters(0).AsClause.Type).Type, INamedTypeSymbol)
            If serializationInfoType Is Nothing Then Return False
            If Not serializationInfoType.AllBaseTypesAndSelf().Any(Function(type As INamedTypeSymbol) type.ToString() = "System.Runtime.Serialization.SerializationInfo") Then Return False

            Dim streamContextType As INamedTypeSymbol = TryCast(model.GetTypeInfo(constructor.SubNewStatement.ParameterList.Parameters(1).AsClause.Type).Type, INamedTypeSymbol)
            If streamContextType Is Nothing Then Return False
            Return streamContextType.AllBaseTypesAndSelf().Any(Function(type As INamedTypeSymbol) type.ToString() = "System.Runtime.Serialization.StreamingContext")
        End Function

        Private Sub Analyzer(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim mBlockSyntax As MethodBlockSyntax = TryCast(context.Node, MethodBlockSyntax)
            If mBlockSyntax IsNot Nothing Then
                If mBlockSyntax.SubOrFunctionStatement IsNot Nothing Then
                    If mBlockSyntax.SubOrFunctionStatement.HandlesClause IsNot Nothing Then
                        Exit Sub
                    End If
                End If
            End If
            Dim methodOrConstructor As MethodBlockBaseSyntax = TryCast(context.Node, MethodBlockBaseSyntax)
            If methodOrConstructor Is Nothing Then Exit Sub
            Dim model As SemanticModel = context.SemanticModel
            If Not IsCandidateForRemoval(methodOrConstructor, model) Then
                Exit Sub
            End If
            If methodOrConstructor.BlockStatement.ParameterList Is Nothing Then
                Exit Sub
            End If
            Dim parameters As Dictionary(Of VisualBasic.Syntax.ParameterSyntax, IParameterSymbol) = methodOrConstructor.BlockStatement.ParameterList.Parameters.ToDictionary(Function(p As ParameterSyntax) p, Function(p As ParameterSyntax) model.GetDeclaredSymbol(p))
            Dim ctor As ConstructorBlockSyntax = TryCast(methodOrConstructor, ConstructorBlockSyntax)
            ' TODO: Check if used in MyBase
            If methodOrConstructor.Statements.Any() Then
                Dim dataFlowAnalysis As DataFlowAnalysis = model.AnalyzeDataFlow(methodOrConstructor.Statements.First, methodOrConstructor.Statements.Last)
                If Not dataFlowAnalysis.Succeeded Then Exit Sub
                For Each parameter As KeyValuePair(Of VisualBasic.Syntax.ParameterSyntax, IParameterSymbol) In parameters
                    Dim parameterSymbol As IParameterSymbol = parameter.Value
                    If parameterSymbol Is Nothing Then Continue For
                    If Not dataFlowAnalysis.ReadInside.Contains(parameterSymbol) AndAlso
                        Not dataFlowAnalysis.WrittenInside.Contains(parameterSymbol) Then
                        context = CreateDiagnostic(context, parameter.Key)
                    End If
                Next
            Else
                For Each parameter As ParameterSyntax In parameters.Keys
                    context = CreateDiagnostic(context, parameter)
                Next
            End If
        End Sub

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Me.Analyzer, SyntaxKind.SubBlock, SyntaxKind.ConstructorBlock, SyntaxKind.FunctionBlock)
        End Sub

    End Class

End Namespace