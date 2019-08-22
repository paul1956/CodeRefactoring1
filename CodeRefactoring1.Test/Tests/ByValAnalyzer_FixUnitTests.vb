
Imports CodeRefactoring1
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports TestHelper

Namespace ByValAnalyzer_Fix.Test

    <TestClass>
    Public Class UnitTest
        Inherits CodeFixVerifier

        Protected Overrides Function GetBasicCodeFixProvider() As CodeFixProvider
            Return New ByValAnalyzer_FixCodeFixProvider()
        End Function

        Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
            Return New ByValAnalyzer_FixAnalyzer()
        End Function

        'Diagnostic And CodeFix both triggered And checked for
        <TestMethod>
        Public Sub TestByValNoExtraTrivia()

            Dim test As String = "
Module Module1

    Sub Main(X As String, ByVal Y As Integer)

    End Sub

End Module"
            Dim expected As New DiagnosticResult With {.Id = "ByValAnalyzer_Fix",
                .Message = String.Format("Remove ByVal from parameter {0}", "1"),
                .Severity = DiagnosticSeverity.Info,
                .Locations = New DiagnosticResultLocation() {
                        New DiagnosticResultLocation("Test0.vb", 4, 27)
                    }
            }

            Me.VerifyBasicDiagnostic(test, expected)

            Dim fixtest As String = "
Module Module1

    Sub Main(X As String, Y As Integer)

    End Sub

End Module"
            Me.VerifyBasicFix(test, fixtest)
        End Sub

        'Diagnostic And CodeFix both triggered And checked for
        <TestMethod>
        Public Sub TestByValWithLeadingTrivia()

            Dim test As String = "
Module Module1

    Sub Main(X As String, ' Comment
             ByVal Y As Integer)

    End Sub

End Module"
            Dim expected As New DiagnosticResult With {.Id = "ByValAnalyzer_Fix",
                .Message = String.Format("Remove ByVal from parameter {0}", "1"),
                .Severity = DiagnosticSeverity.Info,
                .Locations = New DiagnosticResultLocation() {
                        New DiagnosticResultLocation("Test0.vb", 5, 14)
                    }
            }

            Me.VerifyBasicDiagnostic(test, expected)

            Dim fixtest As String = "
Module Module1

    Sub Main(X As String, ' Comment
             Y As Integer)

    End Sub

End Module"
            Me.VerifyBasicFix(test, fixtest)
        End Sub

        <TestMethod>
        Public Sub TestByValWithTrailingMultiLineTrivia()

            Dim test As String = "
Module Module1

    Sub Main(X As String, ByVal _ ' Comment
 _ ' Another Comment
             Y As Integer)

    End Sub

End Module"
            Dim expected As New DiagnosticResult With {.Id = "ByValAnalyzer_Fix",
                .Message = String.Format("Remove ByVal from parameter {0}", "1"),
                .Severity = DiagnosticSeverity.Info,
                .Locations = New DiagnosticResultLocation() {
                        New DiagnosticResultLocation("Test0.vb", 4, 27)
                    }
            }

            Me.VerifyBasicDiagnostic(test, expected)

            Dim fixtest As String = "
Module Module1

    Sub Main(X As String, ' Comment
 _ ' Another Comment
             Y As Integer)

    End Sub

End Module"
            Me.VerifyBasicFix(test, fixtest)
        End Sub

        <TestMethod>
        Public Sub TestByValWithTrailingTrivia()

            Dim test As String = "
Module Module1

    Sub Main(X As String, ByVal _ ' Comment
             Y As Integer)

    End Sub

End Module"
            Dim expected As New DiagnosticResult With {.Id = "ByValAnalyzer_Fix",
                .Message = String.Format("Remove ByVal from parameter {0}", "1"),
                .Severity = DiagnosticSeverity.Info,
                .Locations = New DiagnosticResultLocation() {
                        New DiagnosticResultLocation("Test0.vb", 4, 27)
                    }
            }

            Me.VerifyBasicDiagnostic(test, expected)

            Dim fixtest As String = "
Module Module1

    Sub Main(X As String, ' Comment
             Y As Integer)

    End Sub

End Module"
            Me.VerifyBasicFix(test, fixtest)
        End Sub

        'No diagnostics expected to show up
        <TestMethod>
        Public Sub TestNoByVal()
            Dim test As String = ""
            Me.VerifyBasicDiagnostic(test)
        End Sub

    End Class

End Namespace
