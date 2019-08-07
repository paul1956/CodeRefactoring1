Imports System.Runtime.CompilerServices

Imports Microsoft.CodeAnalysis.Diagnostics

Public Module VBAnalyzerExtensions

    <Extension>
    Public Function ConvertToBaseType(source As ExpressionSyntax, sourceType As ITypeSymbol, targetType As ITypeSymbol) As ExpressionSyntax
        If (sourceType?.IsNumeric() AndAlso targetType?.IsNumeric()) OrElse
            (sourceType?.BaseType?.SpecialType = SpecialType.System_Enum AndAlso targetType?.IsNumeric()) OrElse
            (targetType?.OriginalDefinition.SpecialType = SpecialType.System_Nullable_T) Then Return source
        Return If(sourceType IsNot Nothing AndAlso sourceType.Name = targetType.Name, source, SyntaxFactory.DirectCastExpression(source.WithoutTrailingTrivia, SyntaxFactory.ParseTypeName(targetType.Name))).WithTrailingTrivia(source.GetTrailingTrivia())
    End Function

    <Extension>
    Public Function EnsureNothingAsType(expression As ExpressionSyntax, semanticModel As SemanticModel, type As ITypeSymbol, typeSyntax As TypeSyntax) As ExpressionSyntax
        If type?.OriginalDefinition.SpecialType = SpecialType.System_Nullable_T Then
            Dim constValue As [Optional](Of Object) = semanticModel.GetConstantValue(expression)
            If constValue.HasValue AndAlso constValue.Value Is Nothing Then
                Return SyntaxFactory.DirectCastExpression(expression.WithoutTrailingTrivia(), typeSyntax)
            End If
        End If

        Return expression
    End Function

    <Extension>
    Public Function ExtractAssignmentAsExpressionSyntax(expression As AssignmentStatementSyntax) As ExpressionSyntax
        Select Case expression.Kind
            Case SyntaxKind.AddAssignmentStatement
                Return SyntaxFactory.AddExpression(expression.Left, expression.Right)
            Case SyntaxKind.SubtractAssignmentStatement
                Return SyntaxFactory.SubtractExpression(expression.Left, expression.Right)
            Case SyntaxKind.ConcatenateAssignmentStatement
                Return SyntaxFactory.ConcatenateExpression(expression.Left, expression.Right)
            Case SyntaxKind.DivideAssignmentStatement
                Return SyntaxFactory.DivideExpression(expression.Left, expression.Right)
            Case SyntaxKind.ExponentiateAssignmentStatement
                Return SyntaxFactory.ExponentiateExpression(expression.Left, expression.Right)
            Case SyntaxKind.IntegerDivideAssignmentStatement
                Return SyntaxFactory.IntegerDivideExpression(expression.Left, expression.Right)
            Case SyntaxKind.LeftShiftAssignmentStatement
                Return SyntaxFactory.LeftShiftExpression(expression.Left, expression.Right)
            Case SyntaxKind.MultiplyAssignmentStatement
                Return SyntaxFactory.MultiplyExpression(expression.Left, expression.Right)
            Case SyntaxKind.RightShiftAssignmentStatement
                Return SyntaxFactory.RightShiftExpression(expression.Left, expression.Right)
            Case Else
                Return expression.Right
        End Select
    End Function

    <Extension>
    Public Function GetCommonBaseType(source As ITypeSymbol, other As ITypeSymbol) As ITypeSymbol
        If source Is Nothing AndAlso other IsNot Nothing Then
            Return other
        End If
        If source IsNot Nothing AndAlso other Is Nothing Then
            Return source
        End If

        Dim baseType As ITypeSymbol = source
        While baseType IsNot Nothing
            Dim otherBaseType As ITypeSymbol = other
            While otherBaseType IsNot Nothing
                If baseType.Equals(otherBaseType) Then Return baseType
                otherBaseType = otherBaseType.BaseType
            End While
            baseType = baseType.BaseType
        End While
        Return Nothing
    End Function

    <Extension>
    Public Function HasAttribute(attributeLists As SyntaxList(Of AttributeListSyntax), attributeName As String) As Boolean
        Return attributeLists.SelectMany(Function(a) a.Attributes).Any(Function(a) a.Name.ToString().EndsWith(attributeName, StringComparison.OrdinalIgnoreCase))
    End Function

    <Extension>
    Public Function HasAttributeOnAncestorOrSelf(node As SyntaxNode, ParamArray attributeNames As String()) As Boolean
        Dim vbNode As VisualBasicSyntaxNode = TryCast(node, VisualBasicSyntaxNode)
        If (vbNode Is Nothing) Then Throw New System.Exception("Node is not a VB node.")
        For Each attributeName As String In attributeNames
            If (vbNode.HasAttributeOnAncestorOrSelf(attributeName)) Then Return True
        Next
        Return False
    End Function

    <Extension>
    Public Function HasAttributeOnAncestorOrSelf(node As VisualBasicSyntaxNode, attributeName As String) As Boolean
        Dim parentMethod As VisualBasic.Syntax.MethodBlockBaseSyntax = DirectCast(node.FirstAncestorOrSelfOfType(GetType(MethodBlockSyntax), GetType(ConstructorBlockSyntax)), MethodBlockBaseSyntax)
        If If(parentMethod?.BlockStatement.AttributeLists.HasAttribute(attributeName), False) Then
            Return True
        End If
        Dim type As VisualBasic.Syntax.TypeBlockSyntax = DirectCast(node.FirstAncestorOrSelfOfType(GetType(ClassBlockSyntax), GetType(StructureBlockSyntax)), TypeBlockSyntax)
        While (type IsNot Nothing)
            If type.BlockStatement.AttributeLists.HasAttribute(attributeName) Then Return True
            type = DirectCast(type.FirstAncestorOfType(GetType(ClassBlockSyntax), GetType(StructureBlockSyntax)), TypeBlockSyntax)
        End While
        Dim propertyBlock As VisualBasic.Syntax.PropertyBlockSyntax = node.FirstAncestorOrSelfOfType(Of PropertyBlockSyntax)()
        If If(propertyBlock?.PropertyStatement.AttributeLists.HasAttribute(attributeName), False) Then
            Return True
        End If
        Dim accessor As VisualBasic.Syntax.AccessorBlockSyntax = node.FirstAncestorOrSelfOfType(Of AccessorBlockSyntax)()
        If If(accessor?.AccessorStatement.AttributeLists.HasAttribute(attributeName), False) Then
            Return True
        End If
        Dim anInterface As VisualBasic.Syntax.InterfaceBlockSyntax = node.FirstAncestorOrSelfOfType(Of InterfaceBlockSyntax)()
        If If(anInterface?.InterfaceStatement.AttributeLists.HasAttribute(attributeName), False) Then
            Return True
        End If
        Dim anEnum As VisualBasic.Syntax.EnumBlockSyntax = node.FirstAncestorOrSelfOfType(Of EnumBlockSyntax)()
        If If(anEnum?.EnumStatement.AttributeLists.HasAttribute(attributeName), False) Then
            Return True
        End If
        Dim theModule As VisualBasic.Syntax.ModuleBlockSyntax = node.FirstAncestorOrSelfOfType(Of ModuleBlockSyntax)()
        If If(theModule?.ModuleStatement.AttributeLists.HasAttribute(attributeName), False) Then
            Return True
        End If
        Dim eventBlock As VisualBasic.Syntax.EventBlockSyntax = node.FirstAncestorOrSelfOfType(Of EventBlockSyntax)()
        If If(eventBlock?.EventStatement.AttributeLists.HasAttribute(attributeName), False) Then
            Return True
        End If
        Dim theEvent As VisualBasic.Syntax.EventStatementSyntax = TryCast(node, EventStatementSyntax)
        If If(theEvent?.AttributeLists.HasAttribute(attributeName), False) Then
            Return True
        End If
        Dim theProperty As VisualBasic.Syntax.PropertyStatementSyntax = TryCast(node, PropertyStatementSyntax)
        If If(theProperty?.AttributeLists.HasAttribute(attributeName), False) Then
            Return True
        End If
        Dim field As VisualBasic.Syntax.FieldDeclarationSyntax = TryCast(node, FieldDeclarationSyntax)
        If If(field?.AttributeLists.HasAttribute(attributeName), False) Then
            Return True
        End If
        Dim parameter As VisualBasic.Syntax.ParameterSyntax = TryCast(node, ParameterSyntax)
        If If(parameter?.AttributeLists.HasAttribute(attributeName), False) Then
            Return True
        End If
        Dim aDelegate As VisualBasic.Syntax.DelegateStatementSyntax = TryCast(node, DelegateStatementSyntax)
        If If(aDelegate?.AttributeLists.HasAttribute(attributeName), False) Then
            Return True
        End If
        Return False
    End Function

    <Extension>
    Public Function IsNumeric(typeSymbol As ITypeSymbol) As Boolean
        Return typeSymbol.SpecialType = SpecialType.System_Byte OrElse
            typeSymbol.SpecialType = SpecialType.System_SByte OrElse
            typeSymbol.SpecialType = SpecialType.System_Int16 OrElse
            typeSymbol.SpecialType = SpecialType.System_UInt16 OrElse
            typeSymbol.SpecialType = SpecialType.System_Int16 OrElse
            typeSymbol.SpecialType = SpecialType.System_UInt32 OrElse
            typeSymbol.SpecialType = SpecialType.System_Int32 OrElse
            typeSymbol.SpecialType = SpecialType.System_UInt64 OrElse
            typeSymbol.SpecialType = SpecialType.System_Int64 OrElse
            typeSymbol.SpecialType = SpecialType.System_Decimal OrElse
            typeSymbol.SpecialType = SpecialType.System_Single OrElse
            typeSymbol.SpecialType = SpecialType.System_Double
    End Function

    <Extension>
    Public Sub RegisterCompilationStartAction(context As AnalysisContext, languageVersion As LanguageVersion, registrationAction As Action(Of CompilationStartAnalysisContext))
        context.RegisterCompilationStartAction(Sub(compilationContext) compilationContext.RunIfVBVersionOrGreater(languageVersion, Sub() registrationAction?.Invoke(compilationContext)))
    End Sub

    <Extension>
    Public Sub RegisterSyntaxNodeAction(Of TLanguageKindEnum As Structure)(context As AnalysisContext, languageVersion As LanguageVersion, action As Action(Of SyntaxNodeAnalysisContext), ParamArray syntaxKinds As TLanguageKindEnum())
        context.RegisterCompilationStartAction(languageVersion, Sub(compilationContext) compilationContext.RegisterSyntaxNodeAction(action, syntaxKinds))
    End Sub

#Disable Warning RS1012 ' Start action has no registered actions.

    <Extension>
    Public Sub RunIfVBVersionOrGreater(context As CompilationStartAnalysisContext, languageVersion As LanguageVersion, action As Action)
#Enable Warning RS1012 ' Start action has no registered actions.
        context.Compilation.RunIfVBVersionOrGreater(action, languageVersion)
    End Sub

    <Extension>
    Public Sub RunIfVBVersionOrGreater(compilation As Compilation, action As Action, languageVersion As LanguageVersion)
        Dim vbCompilation As VisualBasic.VisualBasicCompilation = TryCast(compilation, VisualBasicCompilation)
        If vbCompilation Is Nothing Then
            Return
        End If
        vbCompilation.LanguageVersion.RunWithVBVersionOrGreater(action, languageVersion)
    End Sub

    <Extension>
    Public Sub RunWithVBVersionOrGreater(languageVersion As LanguageVersion, action As Action, greaterOrEqualThanLanguageVersion As LanguageVersion)
        If languageVersion >= greaterOrEqualThanLanguageVersion Then action?.Invoke()
    End Sub

End Module