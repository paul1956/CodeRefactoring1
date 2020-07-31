Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic

Module CompilationTestUtils

    ' Use the mscorlib from our current process
    Public Function CreateCompilationWithMscorlibAndVBRuntime(Text As String, Optional FileName As String = "Test") As VisualBasicCompilation
        Dim SyntaxTree As SyntaxTree = SyntaxFactory.ParseSyntaxTree(Text)
        Return VisualBasicCompilation.Create(
                      FileName,
                      syntaxTrees:={SyntaxTree},
                      SharedReferences.References)
    End Function

End Module
