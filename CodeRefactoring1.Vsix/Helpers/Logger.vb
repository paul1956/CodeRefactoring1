' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Diagnostics.CodeAnalysis
Imports System.Globalization
Imports System.Windows.Forms

Imports CodeRefactoring1.Vsix.EditorExtensions.Settings

Imports Microsoft.VisualStudio
Imports Microsoft.VisualStudio.Shell.Interop

Namespace EditorExtensions
    Public Module Logger
        Private pane As IVsOutputWindowPane
        Private ReadOnly _syncRoot As New Object()

        <SuppressMessage("Microsoft.Usage", "CA1806:DoNotIgnoreMethodResults", MessageId:="Microsoft.VisualStudio.Shell.Interop.IVsOutputWindowPane.OutputString(System.String)")>
        Public Sub Log(ByVal message As String)
            If String.IsNullOrEmpty(message) Then
                Return
            End If

            Try
                If EnsurePane() Then
                    pane.OutputString(Date.Now.ToString() & ": " & message & Environment.NewLine)
                End If
            Finally
            End Try
        End Sub

        Public Sub Log(ByVal ex As Exception)
            If ex IsNot Nothing Then
                Log(ex.ToString())
            End If
        End Sub

        Public Sub ShowMessage(ByVal message As String, Optional ByVal title As String = "Web Essentials", Optional ByVal messageBoxButtons As MessageBoxButtons = MessageBoxButtons.OK, Optional ByVal messageBoxIcon As MessageBoxIcon = MessageBoxIcon.Warning, Optional ByVal messageBoxDefaultButton As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1)
            If VBSettings.Instance.General.AllMessagesToOutputWindow Then
                Log(String.Format(CultureInfo.CurrentCulture, "{0}: {1}", title, message))
            Else
                MessageBox.Show(message, title, messageBoxButtons, messageBoxIcon, messageBoxDefaultButton)
            End If
        End Sub

        Private Function EnsurePane() As Boolean
            If pane Is Nothing Then
                SyncLock _syncRoot
                    If pane Is Nothing Then
                        pane = Command1Package.Instance.GetOutputPane(VSConstants.OutputWindowPaneGuid.BuildOutputPane_guid, "Web Essentials")
                    End If
                End SyncLock
            End If

            Return pane IsNot Nothing
        End Function

    End Module
End Namespace
