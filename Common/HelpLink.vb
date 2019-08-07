Namespace CodeCracker
	Public Module HelpLink
		Public Function ForDiagnostic(ByVal diagnosticId As DiagnosticId) As String
			Return $"https://code-cracker.github.io/diagnostics/{diagnosticId.ToDiagnosticId()}.html"
		End Function
	End Module
End Namespace