Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Friend MustInherit Class Matcher(Of T)

    ' Tries to match this matcher against the provided sequence at the given index.  If the
    ' match succeeds, 'true' is returned, and 'index' points to the location after the match
    ' ends.  If the match fails, then false it returned and index remains the same.  Note: the
    ' matcher does not need to consume to the end of the sequence to succeed.
    Public MustOverride Function TryMatch(ByVal sequence As IList(Of T), ByRef index As Integer) As Boolean

    Friend Shared Function [Single](ByVal predicate As Func(Of T, Boolean), ByVal description As String) As Matcher(Of T)
        Return New SingleMatcher(predicate, description)
    End Function

    Friend Shared Function Choice(ByVal matcher1 As Matcher(Of T), ByVal matcher2 As Matcher(Of T)) As Matcher(Of T)
        Return New ChoiceMatcher(matcher1, matcher2)
    End Function

    'INSTANT VB NOTE: The parameter matcher was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
    Friend Shared Function OneOrMore(ByVal matcher_Renamed As Matcher(Of T)) As Matcher(Of T)
        ' m+ is the same as (m m*)
        Return Sequence(matcher_Renamed, Repeat(matcher_Renamed))
    End Function

    'INSTANT VB NOTE: The parameter matcher was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
    Friend Shared Function Repeat(ByVal matcher_Renamed As Matcher(Of T)) As Matcher(Of T)
        Return New RepeatMatcher(matcher_Renamed)
    End Function

    Friend Shared Function Sequence(ParamArray ByVal matchers() As Matcher(Of T)) As Matcher(Of T)
        Return New SequenceMatcher(matchers)
    End Function

    Private Class ChoiceMatcher
        Inherits Matcher(Of T)

        Private ReadOnly _matcher1 As Matcher(Of T)
        Private ReadOnly _matcher2 As Matcher(Of T)

        Public Sub New(ByVal matcher1 As Matcher(Of T), ByVal matcher2 As Matcher(Of T))
            Me._matcher1 = matcher1
            Me._matcher2 = matcher2
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format("({0}|{1})", Me._matcher1, Me._matcher2)
        End Function

        Public Overrides Function TryMatch(ByVal sequence As IList(Of T), ByRef index As Integer) As Boolean
            Return Me._matcher1.TryMatch(sequence, index) OrElse Me._matcher2.TryMatch(sequence, index)
        End Function

    End Class

    Private Class RepeatMatcher
        Inherits Matcher(Of T)

        Private ReadOnly _matcher As Matcher(Of T)

        'INSTANT VB NOTE: The parameter matcher was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
        Public Sub New(ByVal matcher_Renamed As Matcher(Of T))
            Me._matcher = matcher_Renamed
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format("({0}*)", Me._matcher)
        End Function

        Public Overrides Function TryMatch(ByVal sequence As IList(Of T), ByRef index As Integer) As Boolean
            Do While Me._matcher.TryMatch(sequence, index)
            Loop

            Return True
        End Function

    End Class

    Private Class SequenceMatcher
        Inherits Matcher(Of T)

        Private ReadOnly _matchers() As Matcher(Of T)

        Public Sub New(ParamArray ByVal matchers() As Matcher(Of T))
            Me._matchers = matchers
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format("({0})", String.Join(",", CType(Me._matchers, Object())))
        End Function

        Public Overrides Function TryMatch(ByVal sequence As IList(Of T), ByRef index As Integer) As Boolean
            Dim currentIndex As Integer = index
            'INSTANT VB NOTE: The variable matcher was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
            For Each matcher_Renamed As Matcher(Of T) In Me._matchers
                If Not matcher_Renamed.TryMatch(sequence, currentIndex) Then
                    Return False
                End If
            Next matcher_Renamed

            index = currentIndex
            Return True
        End Function

    End Class

    Private Class SingleMatcher
        Inherits Matcher(Of T)

        Private ReadOnly _description As String
        Private ReadOnly _predicate As Func(Of T, Boolean)

        Public Sub New(ByVal predicate As Func(Of T, Boolean), ByVal description As String)
            Me._predicate = predicate
            Me._description = description
        End Sub

        Public Overrides Function ToString() As String
            Return Me._description
        End Function

        Public Overrides Function TryMatch(ByVal sequence As IList(Of T), ByRef index As Integer) As Boolean
            If index < sequence.Count AndAlso Me._predicate(sequence(index)) Then
                index += 1
                Return True
            End If

            Return False
        End Function

    End Class

End Class

Friend MustInherit Class Matcher

    ''' <summary>
    ''' Matcher that matches an element if the provide predicate returns true.
    ''' </summary>
    Public Shared Function [Single](Of T)(ByVal predicate As Func(Of T, Boolean), ByVal description As String) As Matcher(Of T)
        Return Matcher(Of T).Single(predicate, description)
    End Function

    ''' <summary>
    ''' Matcher equivalent to (m_1|m_2)
    ''' </summary>
    Public Shared Function Choice(Of T)(ByVal matcher1 As Matcher(Of T), ByVal matcher2 As Matcher(Of T)) As Matcher(Of T)
        Return Matcher(Of T).Choice(matcher1, matcher2)
    End Function

    ''' <summary>
    ''' Matcher equivalent to (m+)
    ''' </summary>
    'INSTANT VB NOTE: The parameter matcher was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
    Public Shared Function OneOrMore(Of T)(ByVal matcher_Renamed As Matcher(Of T)) As Matcher(Of T)
        Return Matcher(Of T).OneOrMore(matcher_Renamed)
    End Function

    ''' <summary>
    ''' Matcher equivalent to (m*)
    ''' </summary>
    'INSTANT VB NOTE: The parameter matcher was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
    Public Shared Function Repeat(Of T)(ByVal matcher_Renamed As Matcher(Of T)) As Matcher(Of T)
        Return Matcher(Of T).Repeat(matcher_Renamed)
    End Function

    ''' <summary>
    ''' Matcher equivalent to (m_1 ... m_n)
    ''' </summary>
    Public Shared Function Sequence(Of T)(ParamArray ByVal matchers() As Matcher(Of T)) As Matcher(Of T)
        Return Matcher(Of T).Sequence(matchers)
    End Function

End Class