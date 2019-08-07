Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports CodeRefactoring1
Imports CodeRefactoring1.Style

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Imports Style

Imports TestHelper

Imports Xunit

Namespace AddAsClause.UnitTest

    <TestClass()> Public Class AddAsClauseUnitTests
        Inherits CodeFixVerifier

        Protected Overrides Function GetBasicCodeFixProvider() As CodeFixProvider
            Return New AddAsClauseCodeFixProvider()
        End Function

        Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
            Return New AddAsClauseAnalyzer()
        End Function

        <Fact>
        Public Sub AssertTestAddAsClauseInDiagnosticArray1()
            Dim comp As VisualBasicCompilation = CreateCompilationWithMscorlibAndVBRuntime("
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics

Partial Public MustInherit Class DiagnosticVerifier
    Protected Shared Function GetSortedDiagnosticsFromDocuments(analyzer As DiagnosticAnalyzer, documents As Document()) As Diagnostic()
        Dim diagnostics As List(Of Diagnostic) = New List(Of Diagnostic)
        Dim results = SortDiagnostics(diagnostics)
        Return results
    End Function
    Private Shared Function SortDiagnostics(diagnostics As IEnumerable(Of Diagnostic)) As Diagnostic()
        Return Nothing
    End Function
End Class")
            Dim tree As SyntaxTree = comp.SyntaxTrees(0)

            Dim model As SemanticModel = comp.GetSemanticModel(tree)
            Dim nodes As IEnumerable(Of SyntaxNode) = tree.GetCompilationUnitRoot().DescendantNodes()
            Dim node As InvocationExpressionSyntax = nodes.OfType(Of InvocationExpressionSyntax)().Single()

            Dim TypeInfo As TypeInfo = model.GetTypeInfo(node)
            Assert.Equal("Microsoft.CodeAnalysis.Diagnostic()", TypeInfo.ConvertedType.ToTestDisplayString())
        End Sub

        <Fact>
        Public Sub AssertTestForeachRepro()
            Dim comp As VisualBasicCompilation = CreateCompilationWithMscorlibAndVBRuntime(
"Module Module1
    Sub M(s As String)
        For Each c In s
        Next
    End Sub
End Module", "a.vb")

            Dim tree As SyntaxTree = comp.SyntaxTrees(0)

            Dim model As SemanticModel = comp.GetSemanticModel(tree)
            Dim nodes As IEnumerable(Of SyntaxNode) = tree.GetCompilationUnitRoot().DescendantNodes()
            Dim node As ForEachStatementSyntax = nodes.OfType(Of ForEachStatementSyntax)().Single()

            Dim foreachInfo As ForEachStatementInfo = model.GetForEachStatementInfo(node)
            Assert.Equal("System.Char", foreachInfo.ElementType.ToTestDisplayString())
        End Sub

        <Fact>
        Public Sub VerifyDiagnosticAddAsClauseInForDeclared()
            Const OriginalSource As String =
"Class Class1
    Sub Main()
            Dim I as Integer
            For I = 0 to 1
                Console.Write(i)
            Next
    End Sub
End Class"
            Me.VerifyBasicDiagnostic(OriginalSource)
        End Sub

        <Fact>
        Public Sub VerifyDiagnosticAddAsClauseInForEachWithParameter()
            Const OriginalSource As String =
"Class Class1
        Public Function C(sym As Integer) As Integer
            Dim x(10) As Integer
            For Each sym In x
                Return sym
            Next
            Return 0
        End Function
End Class"
            Me.VerifyBasicDiagnostic(OriginalSource)

        End Sub

        <Fact>
        Public Sub VerifyDiagnosticAddAsClauseInForParameter()
            Const OriginalSource As String =
"Class Class1
    Sub Main(I as Integer)
            For I = 0 to 1
                Console.Write(i)
            Next
    End Sub
End Class"
            Me.VerifyBasicDiagnostic(OriginalSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInDiagnosticArray()
            Const OriginalSource As String = "
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics

Partial Public MustInherit Class DiagnosticVerifier
    Protected Shared Function GetSortedDiagnosticsFromDocuments(analyzer As DiagnosticAnalyzer, documents As Document()) As Diagnostic()
        Dim diagnostics As List(Of Diagnostic) = New List(Of Diagnostic)
        Dim results = SortDiagnostics(diagnostics)
        Return results
    End Function
    Private Shared Function SortDiagnostics(diagnostics As IEnumerable(Of Diagnostic)) As Diagnostic()
        Return Nothing
    End Function
End Class"
            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 8, 13)
                        }
                }

            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String = "
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics

Partial Public MustInherit Class DiagnosticVerifier
    Protected Shared Function GetSortedDiagnosticsFromDocuments(analyzer As DiagnosticAnalyzer, documents As Document()) As Diagnostic()
        Dim diagnostics As List(Of Diagnostic) = New List(Of Diagnostic)
        Dim results As Diagnostic() = SortDiagnostics(diagnostics)
        Return results
    End Function
    Private Shared Function SortDiagnostics(diagnostics As IEnumerable(Of Diagnostic)) As Diagnostic()
        Return Nothing
    End Function
End Class"
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInDimNewStatement()
            Const OriginalSource As String =
"Imports  System.Collections.Generic
Module Program
    Public Sub Main()
        Dim s = New List(Of Long)
    End Sub
End Module"
            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 4, 13)
                        }
                }
            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Imports  System.Collections.Generic
Module Program
    Public Sub Main()
        Dim s As New List(Of Long)
    End Sub
End Module"
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInDimStatement()
            Const OriginalSource As String =
"Imports System
Imports AliasedType = NS.C(Of Integer)
Namespace NS
    Public Class C(Of T)
        Public Structure S(Of U)
        End Structure
    End Class
End Namespace
Module Program
    Public Sub Main()
        Dim s = New AliasedType.S(Of Long)
        Console.WriteLine(s.ToString)
    End Sub
End Module"
            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 11, 13)
                        }
                }
            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String = "
Imports System
Imports AliasedType = NS.C(Of Integer)
Namespace NS
    Public Class C(Of T)
        Public Structure S(Of U)
        End Structure
    End Class
End Namespace
Module Program
    Public Sub Main()
        Dim s As AliasedType.S(Of Long) = New AliasedType.S(Of Long)
        Console.WriteLine(s.ToString)
    End Sub
End Module"
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInDimStatementCollectionInitializer()
            Const OriginalSource As String =
"Option Explicit On
Option Infer Off
Option Strict On
Module Program
    Public Sub Main()
        Dim s = {""X""c}
    End Sub
End Module"
            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 6, 13)
                        }
                }
            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String = "
Option Explicit On
Option Infer Off
Option Strict On
Module Program
    Public Sub Main()
        Dim s As Char() = {""X""c}
    End Sub
End Module"
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInFor()
            Dim OriginalSource As String =
<text>Option Explicit Off
Option Infer Off
Option Strict On
Imports System
Class Class1
    Sub Main()
        For I = 0 to 1
            Console.Write(i)
        Next
    End Sub
End Class</text>.Value
            'ERR_StrictDisallowImplicitObject
            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 7, 13)
                        }
                }

            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Dim FixedlSource As String =
<text>Option Explicit Off
Option Infer Off
Option Strict On
Imports System
Class Class1
    Sub Main()
        For I As Integer = 0 to 1
            Console.Write(i)
        Next
    End Sub
End Class</text>.Value
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInForEachArrayElement()
            Const OriginalSource As String =
"Imports System

Class Class1
    Sub Main()
        Dim Array1(10) As String
        For Each element In Array1
            Console.Write(element)
        Next
    End Sub
End Class"
            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 6, 18)
                        }
                }

            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Imports System

Class Class1
    Sub Main()
        Dim Array1(10) As String
        For Each element As String In Array1
            Console.Write(element)
        Next
    End Sub
End Class"
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInForEachArrayElement1()
            Const OriginalSource As String =
"Module Module1
    Sub M(s As String)
        For Each c In s
        Next
    End Sub
End Module"
            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 3, 18)
                        }
                }

            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Module Module1
    Sub M(s As String)
        For Each c As Char In s
        Next
    End Sub
End Module"
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInForEachByte()
            Const OriginalSource As String =
"Import System
Class Class1
    Sub Main()
        For value = Byte.MinValue To Byte.MaxValue
           Console.Write(${value})
        Next
    End Sub
End Class"
            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 4, 13)
                        }
                }

            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Import System
Class Class1
    Sub Main()
        For value As Byte = Byte.MinValue To Byte.MaxValue
           Console.Write(${value})
        Next
    End Sub
End Class"
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInForEachIEnumerable()
            Const OriginalSource As String =
"Imports System.Collections.Generic

Class Class1
    Sub Main()
       Dim annotatedNodesOrTokens As IEnumerable(Of String) = Nothing
       For Each annotatedNodeOrToken In annotatedNodesOrTokens
            Stop
       Next
    End Sub
End Class"
            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 6, 17)
                        }
                }

            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Imports System.Collections.Generic

Class Class1
    Sub Main()
       Dim annotatedNodesOrTokens As IEnumerable(Of String) = Nothing
       For Each annotatedNodeOrToken As String In annotatedNodesOrTokens
            Stop
       Next
    End Sub
End Class"
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInForEachListElement()
            Const OriginalSource As String =
"Imports System
Imports System.Collections.Generic

Class Class1
    Sub Main()
            Dim list1 As New List(Of Integer)({20, 30, 500})
            For Each element In list1
                Console.Write(element)
            Next
    End Sub
End Class"
            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 7, 22)
                        }
                }

            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Imports System
Imports System.Collections.Generic

Class Class1
    Sub Main()
            Dim list1 As New List(Of Integer)({20, 30, 500})
            For Each element As Integer In list1
                Console.Write(element)
            Next
    End Sub
End Class"
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInForEachStructureElement()
            Const OriginalSource As String =
"Imports Microsoft.CodeAnalysis

Class Class1
    Sub Main()
            Dim triviaToMove As SyntaxTriviaList = Nothing
            For Each trivia In triviaToMove
            Next
    End Sub
End Class"
            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 6, 22)
                        }
                }

            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Imports Microsoft.CodeAnalysis

Class Class1
    Sub Main()
            Dim triviaToMove As SyntaxTriviaList = Nothing
            For Each trivia As SyntaxTrivia In triviaToMove
            Next
    End Sub
End Class"
            Me.VerifyBasicFix(OriginalSource, FixedlSource, allowNewCompilerDiagnostics:=True)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInMultiLineLambdaFunction()
            Const OriginalSource As String =
"Class Class1
    Sub Main()
        Dim getSortColumn = Function(index As Integer)
                                Select Case index
                                    Case 0
                                        Return ""FirstName""
                                    Case 1
                                        Return ""LastName""
                                    Case 2
                                        Return ""CompanyName""
                                    Case Else
                                        Return ""LastName""
                                End Select
                            End Function
    End Sub
End Class"
            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 3, 13)
                        }
                }

            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Class Class1
    Sub Main()
        Dim getSortColumn As System.Func(Of Integer, String) = Function(index As Integer)
                                                                   Select Case index
                                                                       Case 0
                                                                           Return ""FirstName""
                                                                       Case 1
                                                                           Return ""LastName""
                                                                       Case 2
                                                                           Return ""CompanyName""
                                                                       Case Else
                                                                           Return ""LastName""
                                                                   End Select
                                                               End Function
    End Sub
End Class"
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInMultiLineLambdaSubroutine()
            Const OriginalSource As String =
"Class Class1
    Sub Main()
        Dim writeToLog = Sub(msg As String)
                            Dim l As String = msg
                         End Sub
    End Sub
End Class"

            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 3, 13)
                        }
                }

            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Class Class1
    Sub Main()
        Dim writeToLog As System.Action(Of String) = Sub(msg As String)
                                                         Dim l As String = msg
                                                     End Sub
    End Sub
End Class"
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInSingleLineLambdaFunction()
            Const OriginalSource As String =
"Class Class1
    Sub Main()
        Dim add1 = Function(num As Integer) num + 1
    End Sub
End Class"
            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 3, 13)
                        }
                }

            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Class Class1
    Sub Main()
        Dim add1 As System.Func(Of Integer, Integer) = Function(num As Integer) num + 1
    End Sub
End Class"
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseInSingleLineLambdaSubroutine()
            Const OriginalSource As String =
"Imports System
Class Class1
    Sub Main()
        Dim writeMessage = Sub(msg As String) Line(msg)
    End Sub
    Private Sub Line(msg As String)
        Throw New NotImplementedException()
    End Sub
End Class"

            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 4, 13)
                        }
                }

            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Imports System
Class Class1
    Sub Main()
        Dim writeMessage As Action(Of String) = Sub(msg As String) Line(msg)
    End Sub
    Private Sub Line(msg As String)
        Throw New NotImplementedException()
    End Sub
End Class"
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

        <Fact>
        Public Sub VerifyFixAddAsClauseToDim()
            Const OriginalSource As String =
"Option Infer Off
Option Strict On
Option Explicit On

Class Class1
    Sub Main()
            Dim X = 1
    End Sub
End Class"
            Dim expected As DiagnosticResult = New DiagnosticResult With {.Id = DiagnosticIds.AddAsClauseDiagnosticId,
                    .Message = "Option Strict On requires all variable declarations to have an 'As' clause.",
                    .Severity = DiagnosticSeverity.Error,
                    .Locations = New DiagnosticResultLocation() {
                            New DiagnosticResultLocation("Test0.vb", 7, 17)
                        }
                }

            Me.VerifyBasicDiagnostic(OriginalSource, expected)

            Const FixedlSource As String =
"Option Infer Off
Option Strict On
Option Explicit On

Class Class1
    Sub Main()
            Dim X As Integer = 1
    End Sub
End Class"
            Me.VerifyBasicFix(OriginalSource, FixedlSource)
        End Sub

    End Class

End Namespace