Imports System.Text.RegularExpressions

Namespace CodeCracker
	Public Module GeneratedCodeAnalysisExtensions
		<System.Runtime.CompilerServices.Extension> _
		Public Function IsOnGeneratedFile(ByVal filePath As String) As Boolean
			Return Regex.IsMatch(filePath, "(\service|\TemporaryGeneratedFile_.*|\assemblyinfo|\assemblyattributes|\.(g\.i|g|designer|generated|assemblyattributes))\.(cs|vb)$", RegexOptions.IgnoreCase)
		End Function
	End Module
End Namespace