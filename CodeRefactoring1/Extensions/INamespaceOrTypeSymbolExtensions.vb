Friend Module INamespaceOrTypeSymbolExtensions

    Private Sub GetNameParts(ByVal namespaceOrTypeSymbol As INamespaceOrTypeSymbol, ByVal result As List(Of String))
        If namespaceOrTypeSymbol Is Nothing OrElse (namespaceOrTypeSymbol.IsNamespace AndAlso DirectCast(namespaceOrTypeSymbol, INamespaceSymbol).IsGlobalNamespace) Then
            Return
        End If

        GetNameParts(namespaceOrTypeSymbol.ContainingNamespace, result)
        result.Add(namespaceOrTypeSymbol.Name)
    End Sub

End Module