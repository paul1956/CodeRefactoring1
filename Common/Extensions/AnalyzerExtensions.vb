Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis

Namespace CodeCracker
	Public Module AnalyzerExtensions
        <Extension>
        Public Function IsName(ByVal displayPart As SymbolDisplayPart) As Boolean
            Return displayPart.IsAnyKind(SymbolDisplayPartKind.ClassName, SymbolDisplayPartKind.DelegateName, SymbolDisplayPartKind.EnumName, SymbolDisplayPartKind.EventName, SymbolDisplayPartKind.FieldName, SymbolDisplayPartKind.InterfaceName, SymbolDisplayPartKind.LocalName, SymbolDisplayPartKind.MethodName, SymbolDisplayPartKind.NamespaceName, SymbolDisplayPartKind.ParameterName, SymbolDisplayPartKind.PropertyName, SymbolDisplayPartKind.StructName)
        End Function

        <Extension>
        Public Function WithSameTriviaAs(ByVal target As SyntaxNode, ByVal source As SyntaxNode) As SyntaxNode
            If target Is Nothing Then
                Throw New ArgumentNullException(NameOf(target))
            End If
            If source Is Nothing Then
                Throw New ArgumentNullException(NameOf(source))
            End If

            Return target.WithLeadingTrivia(source.GetLeadingTrivia()).WithTrailingTrivia(source.GetTrailingTrivia())
        End Function

        <Extension>
        Public Function IsAnyKind(ByVal displayPart As SymbolDisplayPart, ParamArray ByVal kinds() As SymbolDisplayPartKind) As Boolean
            For Each kind In kinds
                If displayPart.Kind Is kind Then
                    Return True
                End If
            Next kind
            Return False
        End Function

        <Extension>
        Public Function FirstAncestorOrSelfOfType(Of T As SyntaxNode)(ByVal node As SyntaxNode) As T
            Return CType(node.FirstAncestorOrSelfOfType(GetType(T)), T)
        End Function

        <Extension>
        Public Function FirstAncestorOrSelfOfType(ByVal node As SyntaxNode, ParamArray ByVal types() As Type) As SyntaxNode
            Dim currentNode = node
            Do
                If currentNode Is Nothing Then
                    Exit Do
                End If
                For Each type In types
                    If currentNode.GetType() Is type Then
                        Return currentNode
                    End If
                Next type
                currentNode = currentNode.Parent
            Loop
            Return Nothing
        End Function

        <Extension>
        Public Function FirstAncestorOfType(Of T As SyntaxNode)(ByVal node As SyntaxNode) As T
            Dim currentNode = node
            Do
                Dim parent = currentNode.Parent
                If parent Is Nothing Then
                    Exit Do
                End If
                Dim tParent = TryCast(parent, T)
                If tParent IsNot Nothing Then
                    Return tParent
                End If
                currentNode = parent
            Loop
            Return Nothing
        End Function

        <Extension>
        Public Function FirstAncestorOfType(ByVal node As SyntaxNode, ParamArray ByVal types() As Type) As SyntaxNode
            Dim currentNode = node
            Do
                Dim parent = currentNode.Parent
                If parent Is Nothing Then
                    Exit Do
                End If
                For Each type In types
                    If parent.GetType() Is type Then
                        Return parent
                    End If
                Next type
                currentNode = parent
            Loop
            Return Nothing
        End Function

        <Extension>
        Public Function GetAllMethodsIncludingFromInnerTypes(ByVal typeSymbol As INamedTypeSymbol) As IList(Of IMethodSymbol)
            Dim methods = typeSymbol.GetMembers().OfType(Of IMethodSymbol)().ToList()
            Dim innerTypes = typeSymbol.GetMembers().OfType(Of INamedTypeSymbol)()
            For Each innerType In innerTypes
                methods.AddRange(innerType.GetAllMethodsIncludingFromInnerTypes())
            Next innerType
            Return methods
        End Function

        <Extension>
        Public Iterator Function AllBaseTypesAndSelf(ByVal typeSymbol As INamedTypeSymbol) As IEnumerable(Of INamedTypeSymbol)
            Yield typeSymbol
            For Each b In AllBaseTypes(typeSymbol)
                Yield b
            Next b
        End Function

        <Extension>
        Public Iterator Function AllBaseTypes(ByVal typeSymbol As INamedTypeSymbol) As IEnumerable(Of INamedTypeSymbol)
            Do While typeSymbol.BaseType IsNot Nothing
                Yield typeSymbol.BaseType
                typeSymbol = typeSymbol.BaseType
            Loop
        End Function

        <Extension>
        Public Function GetLastIdentifierIfQualiedTypeName(ByVal typeName As String) As String
            Dim result = typeName

            Dim parameterTypeDotIndex = typeName.LastIndexOf("."c)
            If parameterTypeDotIndex > 0 Then
                result = typeName.Substring(parameterTypeDotIndex + 1)
            End If

            Return result
        End Function
        <Extension>
        Public Function EnsureProtectedBeforeInternal(ByVal modifiers As IEnumerable(Of SyntaxToken)) As IEnumerable(Of SyntaxToken)
            Return modifiers.OrderByDescending(Function(token) token.RawKind)
        End Function

        <Extension>
        Public Function GetFullName(ByVal symbol As ISymbol, Optional ByVal addGlobal As Boolean = True) As String
            If symbol.Kind = SymbolKind.TypeParameter Then
                Return DirectCast(symbol, Object).ToString()
            End If
            Dim fullName = symbol.Name
            Dim containingSymbol = symbol.ContainingSymbol
            Do While Not (TypeOf containingSymbol Is INamespaceSymbol)
                fullName = $"{containingSymbol.Name}.{fullName}"
                containingSymbol = containingSymbol.ContainingSymbol
            Loop
            If Not DirectCast(containingSymbol, INamespaceSymbol).IsGlobalNamespace Then
                fullName = $"{containingSymbol.ToString()}.{fullName}"
            End If
            If addGlobal Then
                fullName = $"global::{fullName}"
            End If
            Return fullName
        End Function

        <Extension>
        Public Iterator Function GetAllContainingTypes(ByVal symbol As ISymbol) As IEnumerable(Of INamedTypeSymbol)
            Do While symbol.ContainingType IsNot Nothing
                Yield symbol.ContainingType
                symbol = symbol.ContainingType
            Loop
        End Function

        <Extension>
        Public Function GetMinimumCommonAccessibility(ByVal accessibility As Accessibility, ByVal otherAccessibility As Accessibility) As Accessibility
            If accessibility Is otherAccessibility OrElse otherAccessibility Is accessibility.Private Then
                Return accessibility
            End If
            If otherAccessibility Is accessibility.Public Then
                Return accessibility.Public
            End If
            Select Case accessibility
                Case accessibility.Private
                    Return otherAccessibility
                Case accessibility.ProtectedAndInternal, accessibility.Protected, accessibility.Internal
                    Return accessibility.ProtectedAndInternal
                Case accessibility.Public
                    Return accessibility.Public
                Case Else
                    Throw New NotSupportedException()
            End Select
        End Function

        <Extension>
        Public Function IsPrimitive(ByVal typeSymbol As ITypeSymbol) As Boolean
			Select Case typeSymbol.SpecialType
				Case SpecialType.System_Boolean, SpecialType.System_Byte, SpecialType.System_SByte, SpecialType.System_Int16, SpecialType.System_UInt16, SpecialType.System_Int32, SpecialType.System_UInt32, SpecialType.System_Int64, SpecialType.System_UInt64, SpecialType.System_IntPtr, SpecialType.System_UIntPtr, SpecialType.System_Char, SpecialType.System_Double, SpecialType.System_Single
					Return True
				Case Else
					Return False
			End Select
		End Function
	End Module
End Namespace