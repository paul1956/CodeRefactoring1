' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports TestHelper
Imports VBRefactorings
Imports VBRefactorings.Style
Imports Xunit

Namespace ByValAnalyzer_Fix.Test

    Public Class UnitTest
        Inherits CodeFixVerifier

        Protected Overrides Function GetBasicCodeFixProvider() As CodeFixProvider
            Return New ByValAnalyzerFixCodeFixProvider()
        End Function

        Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
            Return New ByValAnalyzerFixAnalyzer()
        End Function

        'Diagnostic And CodeFix both triggered And checked for
        <Fact>
        Public Sub TestByValNoExtraTrivia()

            Dim test As String = "
Module Module1

    Sub Main(X As String, ByVal Y As Integer)

    End Sub

End Module"
            Dim expected As New DiagnosticResult With {.Id = RemoveByValDiagnosticId,
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
        <Fact>
        Public Sub TestByValWithLeadingTrivia()

            Dim test As String = "
Module Module1

    Sub Main(X As String, ' Comment
             ByVal Y As Integer)

    End Sub

End Module"
            Dim expected As New DiagnosticResult With {.Id = RemoveByValDiagnosticId,
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

        <Fact>
        Public Sub TestByValWithTrailingMultiLineTrivia()

            Dim test As String = "
Module Module1

    Sub Main(X As String, ByVal _ ' Comment
 _ ' Another Comment
             Y As Integer)

    End Sub

End Module"
            Dim expected As New DiagnosticResult With {.Id = RemoveByValDiagnosticId,
                .Message = String.Format("Remove ByVal from parameter {0}", "1"),
                .Severity = DiagnosticSeverity.Info,
                .Locations = New DiagnosticResultLocation() {
                        New DiagnosticResultLocation("Test0.vb", 4, 27)
                    }
            }

            Me.VerifyBasicDiagnostic(test, expected)

            Dim fixtest As String = "
Module Module1

    Sub Main(X As String, _ ' Comment
 _ ' Another Comment
Y As Integer)

    End Sub

End Module"
            Me.VerifyBasicFix(test, fixtest)
        End Sub

        <Fact>
        Public Sub TestByValWithTrailingTrivia()

            Dim test As String = "
Module Module1

    Sub Main(X As String, ByVal _ ' Comment
             Y As Integer)

    End Sub

End Module"
            Dim expected As New DiagnosticResult With {.Id = RemoveByValDiagnosticId,
                .Message = String.Format("Remove ByVal from parameter {0}", "1"),
                .Severity = DiagnosticSeverity.Info,
                .Locations = New DiagnosticResultLocation() {
                        New DiagnosticResultLocation("Test0.vb", 4, 27)
                    }
            }

            Me.VerifyBasicDiagnostic(test, expected)

            Dim fixtest As String = "
Module Module1

    Sub Main(X As String, _ ' Comment
Y As Integer)

    End Sub

End Module"
            Me.VerifyBasicFix(test, fixtest)
        End Sub

        'No diagnostics expected to show up
        <Fact>
        Public Sub TestNoByVal()
            Dim test As String = ""
            Me.VerifyBasicDiagnostic(test)
        End Sub

    End Class

End Namespace
