Option Explicit On
Option Infer Off
Option Strict On

Namespace Refactoring
    <ExportCodeRefactoringProvider(LanguageNames.VisualBasic, Name:=NameOf(MoveClassToNewFileCodeRefactoringProvider)), [Shared]>
    Public Class MoveClassToNewFileCodeRefactoringProvider
        Inherits CodeRefactoringProvider

        Public NotOverridable Overrides Async Function ComputeRefactoringsAsync(context As CodeRefactoringContext) As Task
            Dim root As SyntaxNode = Await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(False)

            ' Find the node at the selection.
            Dim node As SyntaxNode = root.FindNode(context.Span)

            ' only for a type declaration node that doesn't match the current file name
            ' also omit all private classes
            Dim ClassStatement As ClassStatementSyntax = TryCast(node, ClassStatementSyntax)
            If ClassStatement Is Nothing Then
                Return
            End If
            Dim className As String = ClassStatement.Identifier.GetIdentifierText()
            Dim newFileName As String = $"{className}.vb"

            If context.Document.Name.ToLowerInvariant() = className.ToLowerInvariant() OrElse ClassStatement.Modifiers.Any(SyntaxKind.PrivateKeyword) Then
                Return
            End If
            Try
                If My.Computer.FileSystem.FileExists(IO.Path.Combine(context.Document.FilePath, newFileName)) Then
                    Exit Function
                End If
            Catch ex As Exception
                Exit Function
            End Try

            Dim toRemove As ClassBlockSyntax = CType(ClassStatement.Parent, ClassBlockSyntax)

            context.RegisterRefactoring(CodeAction.Create($"Move class to new file '{className}.vb'", Function(c As CancellationToken) Me.MoveClassToFile(context, toRemove, className, c)))
        End Function

        Private Async Function MoveClassToFile(lContext As CodeRefactoringContext, ClassBlock As ClassBlockSyntax, className As String, cancellationToken As CancellationToken) As Task(Of Solution)
            Dim newFileName As String = $"{className}.vb"
            Dim lDocument As Document = lContext.Document
            Dim SemanticModel As SemanticModel = Await lDocument.GetSemanticModelAsync(cancellationToken)
            Dim typeSymbol As INamedTypeSymbol = SemanticModel.GetDeclaredSymbol(ClassBlock, cancellationToken)

            ' remove type from current files
            Dim currentSyntaxTree As SyntaxTree = lDocument.GetSyntaxTreeAsync().Result
            Dim currentRoot As SyntaxNode = currentSyntaxTree.GetRootAsync().Result

            Dim replacedRoot As SyntaxNode = currentRoot.RemoveNode(ClassBlock, SyntaxRemoveOptions.KeepNoTrivia)
            Dim changedDocument As Document = lDocument.WithSyntaxRoot(replacedRoot)
            changedDocument = Await RemoveUnusedImportDirectivesAsync(changedDocument, cancellationToken)

            ' create new tree for a new file
            ' we drag all the imports because we don't know which are needed
            ' and there is no easy way to find out which
            Dim currentImports As IEnumerable(Of SyntaxNode) = currentRoot.DescendantNodesAndSelf().Where(Function(s As SyntaxNode) TypeOf s Is ImportsStatementSyntax)
            Dim currentNs As NamespaceStatementSyntax = Nothing
            For Each node As SyntaxNode In currentRoot.DescendantNodesAndSelf()
                If TypeOf node Is NamespaceStatementSyntax Then
                    currentNs = CType(node, NamespaceStatementSyntax)
                End If
            Next

            Dim c As SyntaxList(Of StatementSyntax)
            If currentNs IsNot Nothing Then
                Dim temp As SyntaxList(Of NamespaceBlockSyntax) = SyntaxFactory.SingletonList(SyntaxFactory.NamespaceBlock(SyntaxFactory.NamespaceStatement(SyntaxFactory.ParseName(currentNs.Name.ToString())), SyntaxFactory.SingletonList(ClassBlock.ClassStatement.Parent), SyntaxFactory.EndNamespaceStatement()))
                c = SyntaxFactory.List(temp.[Select](Function(i As NamespaceBlockSyntax) DirectCast(i, StatementSyntax)))
            Else
                c = SyntaxFactory.SingletonList(ClassBlock.ClassStatement.Parent)
            End If

            Dim builder As New Text.StringBuilder()
            For Each import As SyntaxNode In currentImports
                builder.Append(import.ToFullString)
            Next
            Dim NewFileText As String = builder.ToString() & c.ToFullString

            Dim newDocument As Document = Await RemoveUnusedImportDirectivesAsync(changedDocument, cancellationToken)

            'TODO: handle name conflicts
            Dim addedDocument As Document = newDocument.Project.AddDocument(newFileName, NewFileText, lDocument.Folders)

            Return addedDocument.Project.Solution

        End Function

    End Class
End Namespace

