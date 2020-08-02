' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis

Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Performance

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class SealedAttributeAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const Category As String = SupportedCategories.Performance
        Private Const Description As String = "Framework methods that retrieve attributes by default search the entire inheritance hierarchy of the attribute class. Marking the type as NotInheritable eliminates this search and can improve performance."
        Private Const MessageFormat As String = "Mark '{0}' as NotInheritable."
        Private Const Title As String = "Unsealed Attribute"

        Protected Shared Rule As New DiagnosticDescriptor(
                                SealedAttributeDiagnosticId,
                                Title,
                                MessageFormat,
                                Category,
                                DiagnosticSeverity.Warning,
                                isEnabledByDefault:=True,
                                Description,
                                helpLinkUri:=ForDiagnostic(SealedAttributeDiagnosticId),
                                Array.Empty(Of String))

        Public Shared ReadOnly Id As String = SealedAttributeDiagnosticId
        Public Overrides ReadOnly Property SupportedDiagnostics() As ImmutableArray(Of DiagnosticDescriptor) = ImmutableArray.Create(Rule)

        Private Sub Analyze(context As SymbolAnalysisContext)
            If context.IsGenerated Then Return
            Dim type As INamedTypeSymbol = DirectCast(context.Symbol, INamedTypeSymbol)
            If type.TypeKind <> TypeKind.Class Then Exit Sub

            If Not IsAttribute(type) Then Exit Sub
            If type.IsAbstract OrElse type.IsSealed Then Exit Sub
            context.ReportDiagnostic(Diagnostic.Create(Rule, type.Locations(0), type.Name))
        End Sub

        Public Shared Function IsAttribute(symbol As ITypeSymbol) As Boolean
            Dim base As INamedTypeSymbol = symbol.BaseType
            Dim attributeName As String = GetType(Attribute).Name
            While base IsNot Nothing
                If base.Name = attributeName Then Return True
                base = base.BaseType
            End While
            Return False
        End Function

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSymbolAction(AddressOf Me.Analyze, SymbolKind.NamedType)
        End Sub

    End Class

End Namespace
