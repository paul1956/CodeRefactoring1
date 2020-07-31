Imports CodeRefactoring1
Imports CodeRefactoring1.Style

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics

Imports TestHelper
Imports Xunit

Namespace AddAsClauseForLambda.UnitTest

    <TestClass()> Public Class AddAsClauseForLambdaUnitTests
        Inherits CodeFixVerifier

        Protected Overrides Function GetBasicCodeFixProvider() As CodeFixProvider
            Return New AddAsClauseCodeFixProvider()
        End Function

        Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
            Return New AddAsClauseForLambdasAnalyzer()
        End Function

        <Fact>
        Public Sub UnitTestAddAsClauseInMultiLineLambdaFunctionSplit()
            Const OriginalSource As String =
"Class Class1
    Sub Main()
        Dim getSortColumn As System.Func(Of Integer, String)
        getSortColumn = Function(index)
                            Select Case index
                                Case 0
                                    Return ""FirstName""
                                Case Else
                                    Return ""LastName""
                            End Select
                        End Function
    End Sub
End Class"

            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseForLambdaDiagnosticId,
                    .Message = String.Format("Option Strict On requires all Lambda declarations to have an 'As' clause.", "Class1"),
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 4, 34)
                        }
                }

            VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Class Class1
    Sub Main()
        Dim getSortColumn As System.Func(Of Integer, String)
        getSortColumn = Function(index As Integer)
                            Select Case index
                                Case 0
                                    Return ""FirstName""
                                Case Else
                                    Return ""LastName""
                            End Select
                        End Function
    End Sub
End Class"
            VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub UnitTestAddAsClauseInMultiLineLambdaSubroutineSplit()
            Const OriginalSource As String =
"Class Class1
    Sub Main()
        Dim writeToLog As System.Action(Of String)
        writeToLog = Sub(msg)
                        Dim l As String = msg
                     End Sub
    End Sub
End Class"

            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseForLambdaDiagnosticId,
                    .Message = String.Format("Option Strict On requires all Lambda declarations to have an 'As' clause.", "Class1"),
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 4, 26)
                        }
                }

            VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Class Class1
    Sub Main()
        Dim writeToLog As System.Action(Of String)
        writeToLog = Sub(msg As String)
                         Dim l As String = msg
                     End Sub
    End Sub
End Class"
            VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub UnitTestAddAsClauseInSingleLineLambdaFunctionSplit()
            Const OriginalSource As String =
"Class Class1
    Sub Main()
        Dim add1 As System.Func(Of UInteger, Integer, Integer)
        add1 = Function(Num, Num1) Num + Num1
    End Sub
End Class"
            Dim expected1 As DiagnosticResult = New DiagnosticResult With
                {
                .Id = DiagnosticIds.AddAsClauseForLambdaDiagnosticId,
                .Message = String.Format("Option Strict On requires all Lambda declarations to have an 'As' clause.", "Class1"),
                .Severity = DiagnosticSeverity.Warning,
                .Locations = New DiagnosticResultLocation() {New DiagnosticResultLocation("Test0.vb", 4, 25)}
                }
            Dim expected2 As DiagnosticResult = New DiagnosticResult With
                {
                .Id = DiagnosticIds.AddAsClauseForLambdaDiagnosticId,
                .Message = String.Format("Option Strict On requires all Lambda declarations to have an 'As' clause.", "Class1"),
                .Severity = DiagnosticSeverity.Warning,
                .Locations = New DiagnosticResultLocation() {New DiagnosticResultLocation("Test0.vb", 4, 30)}
                }

            VerifyBasicDiagnostic(OriginalSource, expected1, expected2)

            Const FixedlSource As String =
"Class Class1
    Sub Main()
        Dim add1 As System.Func(Of UInteger, Integer, Integer)
        add1 = Function(Num As UInteger, Num1 As Integer) Num + Num1
    End Sub
End Class"
            VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub UnitTestAddAsClauseInSingleLineLambdaSubroutineSplit()
            Const OriginalSource As String =
"Class Class1
    Sub Main()
        Dim writeMessage As System.Action(Of String)
        writeMessage = Sub(msg) Line(msg)
    End Sub
    Private Sub Line(msg As String)
        Throw New NotImplementedException()
    End Sub
End Class"

            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseForLambdaDiagnosticId,
                    .Message = String.Format("Option Strict On requires all Lambda declarations to have an 'As' clause.", "Class1"),
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 4, 28)
                        }
                }

            VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Class Class1
    Sub Main()
        Dim writeMessage As System.Action(Of String)
        writeMessage = Sub(msg As String) Line(msg)
    End Sub
    Private Sub Line(msg As String)
        Throw New NotImplementedException()
    End Sub
End Class"
            VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

    End Class

End Namespace
