' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Namespace Refactoring
    Public Module UnusedImportDirectives

        Friend Async Function RemoveUnusedImportDirectivesAsync(document As Document, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync()
            Dim semanticModel As SemanticModel = Await document.GetSemanticModelAsync()

            root = RemoveUnusedImportDirectives(semanticModel, root, cancellationToken)
            document = document.WithSyntaxRoot(root)
            Return document
        End Function

        Private Function GetUnusedImportDirectives(model As SemanticModel, cancellationToken As CancellationToken) As HashSet(Of SyntaxNode)
            Dim unusedImportDirectives As New HashSet(Of SyntaxNode)()
            Dim root As SyntaxNode = model.SyntaxTree.GetRoot(cancellationToken)
            For Each diagnostic As Diagnostic In model.GetDiagnostics(Nothing, cancellationToken).Where(Function(d) d.Id = "BC50001")
                Dim usingDirectiveSyntax As ImportsStatementSyntax = TryCast(root.FindNode(diagnostic.Location.SourceSpan, False, False), ImportsStatementSyntax)
                If usingDirectiveSyntax IsNot Nothing Then
                    unusedImportDirectives.Add(usingDirectiveSyntax)
                End If
            Next

            Return unusedImportDirectives
        End Function

        Private Function RemoveUnusedImportDirectives(semanticModel As SemanticModel, root As SyntaxNode, cancellationToken As CancellationToken) As SyntaxNode
            Dim oldUsings As IEnumerable(Of SyntaxNode) = root.DescendantNodesAndSelf().Where(Function(s) TypeOf s Is ImportsStatementSyntax)
            Dim unusedUsings As HashSet(Of SyntaxNode) = GetUnusedImportDirectives(semanticModel, cancellationToken)
            Dim leadingTrivia As SyntaxTriviaList = root.GetLeadingTrivia()

            root = root.RemoveNodes(oldUsings, SyntaxRemoveOptions.KeepNoTrivia)
            Dim newUsings As SyntaxList(Of SyntaxNode) = SyntaxFactory.List(oldUsings.Except(unusedUsings))
            root = DirectCast(root, CompilationUnitSyntax).WithImports(newUsings).NormalizeWhitespace().WithLeadingTrivia(leadingTrivia)

            Return root
        End Function

    End Module
End Namespace
