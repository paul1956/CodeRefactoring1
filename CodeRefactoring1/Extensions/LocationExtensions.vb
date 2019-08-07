Imports System.Runtime.CompilerServices

Public Module LocationExtensions

    <Extension>
    Public Function FindNode(ByVal location As Location, ByVal getInnermostNodeForTie As Boolean, ByVal cancellationToken As CancellationToken) As SyntaxNode
        Return location.SourceTree.GetRoot(cancellationToken).FindNode(location.SourceSpan, getInnermostNodeForTie:=getInnermostNodeForTie)
    End Function

End Module