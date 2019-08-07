Imports System.Runtime.CompilerServices

Public Module SyntaxList_1_cs
    ''' <summary>
    ''' The index of the node in this list, or -1 if the node is not in the list.
    ''' </summary>
    <Extension>
    Public Function IndexOf(Of TNode)(List1 As SyntaxList(Of StatementSyntax), ByVal node As TNode) As Integer
        Dim index As Integer = 0
        For Each child As StatementSyntax In List1
            If Object.Equals(child, node) Then
                Return index
            End If

            index += 1
        Next child

        Return -1
    End Function

End Module
