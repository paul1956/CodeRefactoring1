Public Module HelpLink
    Public Function ForDiagnostic(ByVal diagnosticId As String) As String
        Return $"https://code-cracker.github.io/diagnostics/{diagnosticId}.html"
    End Function
End Module
