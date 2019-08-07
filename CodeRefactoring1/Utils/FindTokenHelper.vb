﻿Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Friend Module FindTokenHelper

    ''' <summary>
    ''' Look inside a trivia list for a skipped token that contains the given position.
    ''' </summary>
    Public Function FindSkippedTokenBackward(ByVal skippedTokenList As IEnumerable(Of SyntaxToken), ByVal position As Integer) As SyntaxToken
        ' the given skipped token list is already in order
        Dim skippedTokenContainingPosition As SyntaxToken = skippedTokenList.LastOrDefault(Function(skipped As SyntaxToken) skipped.Span.Length > 0 AndAlso skipped.SpanStart <= position)
        If skippedTokenContainingPosition <> Nothing Then
            Return skippedTokenContainingPosition
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Look inside a trivia list for a skipped token that contains the given position.
    ''' </summary>
    Public Function FindSkippedTokenForward(ByVal skippedTokenList As IEnumerable(Of SyntaxToken), ByVal position As Integer) As SyntaxToken
        ' the given token list is already in order
        Dim skippedTokenContainingPosition As SyntaxToken = skippedTokenList.FirstOrDefault(Function(skipped As SyntaxToken) skipped.Span.Length > 0 AndAlso position <= skipped.Span.End)
        If skippedTokenContainingPosition = Nothing Then
            Return skippedTokenContainingPosition
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' If the position is inside of token, return that token; otherwise, return the token to the left.
    ''' </summary>
    Public Function FindTokenOnLeftOfPosition(Of TRoot As SyntaxNode)(ByVal root As SyntaxNode, ByVal position As Integer, ByVal skippedTokenFinder As Func(Of SyntaxTriviaList, Integer, SyntaxToken), Optional ByVal includeSkipped As Boolean = False, Optional ByVal includeDirectives As Boolean = False, Optional ByVal includeDocumentationComments As Boolean = False) As SyntaxToken
        Dim findSkippedToken As Func(Of SyntaxTriviaList, Integer, SyntaxToken) = If(skippedTokenFinder, (Function(l As SyntaxTriviaList, p As Integer) Nothing))

        Dim token As SyntaxToken = GetInitialToken(Of TRoot)(root, position, includeSkipped, includeDirectives, includeDocumentationComments)

        If position <= token.SpanStart Then
            Do
                Dim skippedToken As SyntaxToken = findSkippedToken(token.LeadingTrivia, position)
                token = If(skippedToken.RawKind <> 0, skippedToken, token.GetPreviousToken(includeZeroWidth:=False, includeSkipped:=includeSkipped, includeDirectives:=includeDirectives, includeDocumentationComments:=includeDocumentationComments))
            Loop While position <= token.SpanStart AndAlso root.FullSpan.Start < token.SpanStart
        ElseIf token.Span.End < position Then
            Dim skippedToken As SyntaxToken = findSkippedToken(token.TrailingTrivia, position)
            token = If(skippedToken.RawKind <> 0, skippedToken, token)
        End If

        If token.Span.Length = 0 Then
            token = token.GetPreviousToken()
        End If

        Return token
    End Function

    ''' <summary>
    ''' If the position is inside of token, return that token; otherwise, return the token to the right.
    ''' </summary>
    Public Function FindTokenOnRightOfPosition(Of TRoot As SyntaxNode)(ByVal root As SyntaxNode, ByVal position As Integer, ByVal skippedTokenFinder As Func(Of SyntaxTriviaList, Integer, SyntaxToken), Optional ByVal includeSkipped As Boolean = False, Optional ByVal includeDirectives As Boolean = False, Optional ByVal includeDocumentationComments As Boolean = False) As SyntaxToken
        Dim findSkippedToken As Func(Of SyntaxTriviaList, Integer, SyntaxToken) = If(skippedTokenFinder, (Function(l As SyntaxTriviaList, p As Integer) Nothing))

        Dim token As SyntaxToken = GetInitialToken(Of TRoot)(root, position, includeSkipped, includeDirectives, includeDocumentationComments)

        If position < token.SpanStart Then
            Dim skippedToken As SyntaxToken = findSkippedToken(token.LeadingTrivia, position)
            token = If(skippedToken.RawKind <> 0, skippedToken, token)
        ElseIf token.Span.End <= position Then
            Do
                Dim skippedToken As SyntaxToken = findSkippedToken(token.TrailingTrivia, position)
                token = If(skippedToken.RawKind <> 0, skippedToken, token.GetNextToken(includeZeroWidth:=False, includeSkipped:=includeSkipped, includeDirectives:=includeDirectives, includeDocumentationComments:=includeDocumentationComments))
            Loop While token.RawKind <> 0 AndAlso token.Span.End <= position AndAlso token.Span.End <= root.FullSpan.End
        End If

        If token.Span.Length = 0 Then
            token = token.GetNextToken()
        End If

        Return token
    End Function

    Private Function GetInitialToken(Of TRoot As SyntaxNode)(ByVal root As SyntaxNode, ByVal position As Integer, Optional ByVal includeSkipped As Boolean = False, Optional ByVal includeDirectives As Boolean = False, Optional ByVal includeDocumentationComments As Boolean = False) As SyntaxToken
        Dim token As SyntaxToken = If(position < root.FullSpan.End OrElse Not (TypeOf root Is TRoot), root.FindToken(position, includeSkipped OrElse includeDirectives OrElse includeDocumentationComments), root.GetLastToken(includeZeroWidth:=True, includeSkipped:=True, includeDirectives:=True, includeDocumentationComments:=True).GetPreviousToken(includeZeroWidth:=False, includeSkipped:=includeSkipped, includeDirectives:=includeDirectives, includeDocumentationComments:=includeDocumentationComments))
        Return token
    End Function

End Module