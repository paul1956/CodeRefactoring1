Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.CodeFixes

Namespace Style

    Friend Class AddAsClauseCodeFixAllProvider
        Inherits CodeFixProvider

        Public Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String)
            Get
                Return ImmutableArray.Create(
            AddAsClauseDiagnosticId,
            AddAsClauseForLambdaDiagnosticId,
            ChangeAsObjectToMoreSpecificDiagnosticId,
            ERR_ClassNotExpression1DiagnosticId,
            ERR_EnumNotExpression1DiagnosticId,
            ERR_InterfaceNotExpression1DiagnosticId,
            ERR_NameNotDeclared1DiagnosticId,
            ERR_NamespaceNotExpression1DiagnosticId,
            ERR_StrictDisallowImplicitObjectDiagnosticId,
            ERR_StrictDisallowsImplicitProcDiagnosticId,
            ERR_StructureNotExpression1DiagnosticId,
            ERR_TypeNotExpression1DiagnosticId)
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