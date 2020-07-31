Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.Formatting

Namespace TestHelper

    ''' <summary>
    ''' Superclass of all Unit tests made for diagnostics with codefixes.
    ''' Contains methods used to verify correctness of codefixes
    ''' </summary>
    Partial Public MustInherit Class CodeFixVerifier
        Inherits DiagnosticVerifier

        ''' <summary>
        ''' Returns the codefix being tested (VB) - to be implemented in non-abstract class
        ''' </summary>
        ''' <returns>The CodeFixProvider to be used for VisualBasic code</returns>
        Protected Overridable Function GetBasicCodeFixProvider() As CodeFixProvider
            Return Nothing
        End Function

        ''' <summary>
        ''' Returns the codefix being tested (C#) - to be implemented in non-abstract class
        ''' </summary>
        ''' <returns>The CodeFixProvider to be used for CSharp code</returns>
        Protected Overridable Function GetCSharpCodeFixProvider() As CodeFixProvider
            Return Nothing
        End Function

        ''' <summary>
        ''' Called to test a VB codefix when applied on the inputted string as a source
        ''' </summary>
        ''' <param name="oldSource">A class in the form of a string before the CodeFix was applied to it</param>
        ''' <param name="newSource">A class in the form of a string after the CodeFix was applied to it</param>
        ''' <param name="codeFixIndex">Index determining which codefix to apply if there are multiple</param>
        ''' <param name="allowNewCompilerDiagnostics">A bool controlling whether Or Not the test will fail if the CodeFix introduces other warnings after being applied</param>
        Protected Sub VerifyBasicFix(oldSource As String, newSource As String, Optional codeFixIndex As Integer? = Nothing, Optional allowNewCompilerDiagnostics As Boolean = False)
            VerifyFix(LanguageNames.VisualBasic, GetBasicDiagnosticAnalyzer(), GetBasicCodeFixProvider(), oldSource, newSource, codeFixIndex, allowNewCompilerDiagnostics)
        End Sub

        ''' <summary>
        ''' Called to test a C# codefix when applied on the inputted string as a source
        ''' </summary>
        ''' <param name="OldSource">A class in the form of a string before the CodeFix was applied to it</param>
        ''' <param name="NewSource">A class in the form of a string after the CodeFix was applied to it</param>
        ''' <param name="CodeFixIndex">Index determining which codefix to apply if there are multiple</param>
        ''' <param name="AllowNewCompilerDiagnostics">A bool controlling whether Or Not the test will fail if the CodeFix introduces other warnings after being applied</param>
        Protected Sub VerifyCSharpFix(OldSource As String, NewSource As String, Optional CodeFixIndex As Integer? = Nothing, Optional AllowNewCompilerDiagnostics As Boolean = False)
            VerifyFix(LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer(), GetCSharpCodeFixProvider(), OldSource, NewSource, CodeFixIndex, AllowNewCompilerDiagnostics)
        End Sub

        Private Shared Function FindFirstDifferenceColumn(DesiredLine As String, ActualLine As String) As (Integer, String)
            Dim MinLength As Integer = Math.Min(DesiredLine.Length, ActualLine.Length) - 1
            For I As Integer = 0 To MinLength
                If Not DesiredLine.Substring(I, 1).Equals(ActualLine.Substring(I, 1), StringComparison.CurrentCulture) Then
                    Return (I + 1, $"Desired Character ""{DesiredLine.Substring(I, 1)}"", Actual Character ""{ActualLine.Substring(I, 1)}""")
                End If
            Next
            If DesiredLine.Length > ActualLine.Length Then
                Return (MinLength + 1, $"Desired Character ""{DesiredLine.Substring(MinLength + 1, 1)}"", Actual Character Nothing")
            Else
                Return (MinLength + 1, $"Desired Character Nothing, Actual Character ""{ActualLine.Substring(MinLength + 1, 1)}""")
            End If
        End Function

        Private Shared Function FindFirstDifferenceLine(DesiredText As String, ActualText As String) As String
            Dim Desiredlines() As String = DesiredText.Replace(vbCr, "").Split(CType(vbLf, Char()))
            Dim ActuaLines() As String = ActualText.Replace(vbCr, "").Split(CType(vbLf, Char()))
            For I As Integer = 0 To Math.Min(Desiredlines.GetUpperBound(0), ActuaLines.GetUpperBound(0))
                Dim DesiredLine As String = Desiredlines(I)
                Dim ActualLine As String = ActuaLines(I)
                If Not DesiredLine.Equals(ActualLine, StringComparison.CurrentCulture) Then
                    Dim p As (Integer, String) = FindFirstDifferenceColumn(DesiredLine, ActualLine)
                    Return $"{vbCrLf}Difference on Line {I + 1} {DesiredLine}{vbCrLf}Difference on Line {I + 1} {ActualLine}{vbCrLf}Column {p.Item1} {p.Item2}"
                End If
            Next
            Return "Files identical"
        End Function

        ''' <summary>
        ''' General verifier for codefixes.
        ''' Creates a Document from the source string, then gets diagnostics on it And applies the relevant codefixes.
        ''' Then gets the string after the codefix Is applied And compares it with the expected result.
        ''' Note: If any codefix causes New diagnostics To show up, the test fails unless allowNewCompilerDiagnostics Is Set To True.
        ''' </summary>
        ''' <param name="Language">The language the source code Is in</param>
        ''' <param name="Analyzer">The analyzer to be applied to the source code</param>
        ''' <param name="FixProvider">The codefix to be applied to the code wherever the relevant Diagnostic Is found</param>
        ''' <param name="OldSource">A class in the form of a string before the CodeFix was applied to it</param>
        ''' <param name="NewSource">A class in the form of a string after the CodeFix was applied to it</param>
        ''' <param name="CodeFixIndex">Index determining which codefix to apply if there are multiple</param>
        ''' <param name="AllowNewCompilerDiagnostics">A bool controlling whether Or Not the test will fail if the CodeFix introduces other warnings after being applied</param>
        Private Shared Sub VerifyFix(Language As String, Analyzer As DiagnosticAnalyzer, FixProvider As CodeFixProvider, OldSource As String, NewSource As String, CodeFixIndex As Integer?, AllowNewCompilerDiagnostics As Boolean)

            Dim OldDocument As Document = CreateDocument(OldSource, Language)
            Dim NewDocument As Document = Nothing
            Dim AnalyzerDiagnostics As Diagnostic() = GetSortedDiagnosticsFromDocuments(Analyzer, New Document() {OldDocument})
            Dim CompilerDiagnostics As IEnumerable(Of Diagnostic) = GetCompilerDiagnostics(OldDocument)
            Dim Attempts As Integer = AnalyzerDiagnostics.Length

            For I As Integer = 0 To Attempts - 1
                Dim Actions As New List(Of CodeAction)()
                Dim Context As CodeFixContext = New CodeFixContext(OldDocument, AnalyzerDiagnostics(0), Sub(a As CodeAction, d As Immutable.ImmutableArray(Of Diagnostic)) Actions.Add(a), CancellationToken.None)
                FixProvider.RegisterCodeFixesAsync(Context).Wait()

                If Not Actions.Any() Then
                    Exit For
                End If

                If (CodeFixIndex IsNot Nothing) Then
                    NewDocument = ApplyFix(OldDocument, Actions.ElementAt(CInt(CodeFixIndex))).Result
                    Exit For
                End If

                NewDocument = ApplyFix(OldDocument, Actions.ElementAt(0)).Result
                AnalyzerDiagnostics = GetSortedDiagnosticsFromDocuments(Analyzer, New Document() {NewDocument})

                Dim NewCompilerDiagnostics As IEnumerable(Of Diagnostic) = GetNewDiagnostics(CompilerDiagnostics, GetCompilerDiagnostics(NewDocument))

                'check if applying the code fix introduced any New compiler diagnostics
                If Not AllowNewCompilerDiagnostics AndAlso NewCompilerDiagnostics.Any() Then
                    ' Format And get the compiler diagnostics again so that the locations make sense in the output
                    NewDocument = NewDocument.WithSyntaxRoot(Formatter.Format(NewDocument.GetSyntaxRootAsync().Result, Formatter.Annotation, NewDocument.Project.Solution.Workspace))
                    NewCompilerDiagnostics = GetNewDiagnostics(CompilerDiagnostics, GetCompilerDiagnostics(NewDocument))

                    Assert.IsTrue(False,
                        String.Format("Fix introduced new compiler diagnostics:{2}{0}{2}{2}New document:{2}{1}{2}",
                            String.Join(vbNewLine, NewCompilerDiagnostics.Select(Function(d As Diagnostic) d.ToString())),
                            NewDocument.GetSyntaxRootAsync().Result.ToFullString(), vbNewLine))
                End If

                'check if there are analyzer diagnostics left after the code fix
                If Not AnalyzerDiagnostics.Any() Then
                    Exit For
                End If
            Next

            'after applying all of the code fixes, compare the resulting string to the inputted one
            Dim Actual As String = GetStringFromDocument(NewDocument)

            Assert.AreEqual(NewSource.Trim, Actual.Trim, FindFirstDifferenceLine(NewSource.Trim, Actual.Trim))
        End Sub

    End Class

End Namespace
