Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.CodeFixes

Namespace Style

    Friend Class AddAsClauseCodeFixAllProvider
        Inherits CodeFixProvider

        'Private Const title As String = "Add/Update As Clause"
        Public Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String)
            Get
                Return ImmutableArray.Create(
            DiagnosticIds.AddAsClauseDiagnosticId,
            DiagnosticIds.AddAsClauseForLambdaDiagnosticId,
            DiagnosticIds.ChangeAsObjectToMoreSpecificDiagnosticId,
            DiagnosticIds.ERR_ClassNotExpression1DiagnosticId,
            DiagnosticIds.ERR_EnumNotExpression1DiagnosticId,
            DiagnosticIds.ERR_InterfaceNotExpression1DiagnosticId,
            DiagnosticIds.ERR_NameNotDeclared1DiagnosticId,
            DiagnosticIds.ERR_NamespaceNotExpression1DiagnosticId,
            DiagnosticIds.ERR_StrictDisallowImplicitObjectDiagnosticId,
            DiagnosticIds.ERR_StrictDisallowsImplicitProcDiagnosticId,
            DiagnosticIds.ERR_StructureNotExpression1DiagnosticId,
            DiagnosticIds.ERR_TypeNotExpression1DiagnosticId)
            End Get
        End Property

        Public NotOverridable Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Throw New NotImplementedException()
        End Function

    End Class

End Namespace