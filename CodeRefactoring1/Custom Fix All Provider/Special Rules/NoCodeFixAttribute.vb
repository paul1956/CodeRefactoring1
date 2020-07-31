' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis.Diagnostics

''' <summary>
''' This attribute is applied to <see cref="DiagnosticAnalyzer"/>s for which no code fix is currently planned.
''' </summary>
''' <remarks>
''' <para>There are several reasons an analyzer does not have a code fix, including but not limited to the
''' following:</para>
''' <list type="bullet">
''' <item>Visual Studio provides a built-in code fix.</item>
''' <item>A code fix could not provide a useful solution.</item>
''' </list>
''' <para>The <see cref="Reason"/> should be provided.</para>
''' </remarks>
<AttributeUsage(AttributeTargets.Class, Inherited:=False, AllowMultiple:=False)>
Friend NotInheritable Class NoCodeFixAttribute
    Inherits System.Attribute

    ''' <summary>
    ''' Initializes a new instance of the <see cref="NoCodeFixAttribute"/> class.
    ''' </summary>
    ''' <param name="reason">The reason why the <see cref="DiagnosticAnalyzer"/> does not have a code fix.</param>
    Public Sub New(ByVal reason As String)
        Me.Reason = reason
    End Sub

    ''' <summary>
    ''' Gets the reason why the <see cref="DiagnosticAnalyzer"/> does not have a code fix.
    ''' </summary>
    ''' <value>
    ''' The reason why the <see cref="DiagnosticAnalyzer"/> does not have a code fix.
    ''' </value>
    Public ReadOnly Property Reason() As String

End Class
