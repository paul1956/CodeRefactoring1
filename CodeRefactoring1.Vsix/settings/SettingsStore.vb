' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.IO

Imports ConfOxide

Imports EnvDTE

Imports Microsoft.VisualStudio.Settings
Imports Microsoft.VisualStudio.Shell.Settings

Namespace EditorExtensions.Settings
    Friend Module SettingsStore
        Public Const FileName As String = "VisualBasicRefactoring-Settings.json"

        Public ReadOnly Property SolutionSettingsExist() As Boolean
            Get
                Return File.Exists(GetSolutionFilePath())
            End Get
        End Property

        Public ReadOnly Property ProjectSettingsExist() As Boolean
            Get
                Return File.Exists(GetProjectFilePath())
            End Get
        End Property

        Public Property InTestMode() As Boolean

        '''<summary>Resets all settings and disables persistence for unit tests.</summary>
        ''' <param name="testSettings">Optional settings to apply for the tests.</param>
        ''' <remarks>It is completely safe to call this function multiple times.</remarks>
        Public Sub EnterTestMode(Optional ByVal testSettings As VBSettings = Nothing)
            InTestMode = True
            If testSettings IsNot Nothing Then
                VBSettings.Instance.AssignFrom(testSettings)
            Else
                VBSettings.Instance.ResetValues()
            End If
        End Sub

        '''<summary>Loads the active settings file.</summary>
        Public Sub LoadSolutionSettings()
            If InTestMode Then
                Return
            End If
            Dim jsonPath As String = GetFilePath()
            If File.Exists(jsonPath) Then
                VBSettings.Instance.ReadJsonFile(jsonPath)
                UpdateStatusBar("applied")
            End If
        End Sub

        Public Sub LoadProjectSettings()
            If InTestMode Then
                Return
            End If
            Dim ProjectList As IEnumerable(Of Project) = GetAllProjects()
            For Each lProject As Project In ProjectList
                If Path.GetDirectoryName(lProject.FileName).EndsWith("VSIX", StringComparison.InvariantCultureIgnoreCase) Then
                    Continue For
                End If
                Dim jsonPath As String = Path.Combine(Path.GetDirectoryName(lProject.FileName), "VBCodeRefactoring-Settings.json")
                If File.Exists(jsonPath) Then
                    VBSettings.Instance.ReadJsonFile(jsonPath)
                    UpdateStatusBar("applied")
                    Exit Sub
                End If
            Next
            VBSettings.Instance.General.CsvIgnoreList = "Microsoft,MS,System"
        End Sub

        Public Sub SaveProjectSettings()
            Dim ProjectList As IEnumerable(Of Project) = GetAllProjects()
            For Each lProject As Project In ProjectList
                If Not lProject.FileName.EndsWith("vbproj") Then
                    Continue For
                End If
                If Path.GetDirectoryName(lProject.FileName).EndsWith("VSIX", StringComparison.InvariantCultureIgnoreCase) Then
                    Continue For
                End If
                Dim FileFound As Boolean = False
                Dim NewFileName As String = String.Empty
                For Each item As ProjectItem In lProject.ProjectItems
                    For i As Short = 0 To item.FileCount - 1
                        If item.FileNames(i).EndsWith("VBCodeRefactoring-Settings.json") Then
                            FileFound = True
                            Exit For
                        End If
                        NewFileName = Path.Combine(Path.GetDirectoryName(lProject.FileName), "VBCodeRefactoring-Settings.json")
                    Next
                Next
                Save(NewFileName, lProject, create:=Not FileFound)
            Next
        End Sub

        Public Sub SaveSolutionSettings()
            Save(GetFilePath(), Project:=Nothing)
        End Sub

        '''<summary>Saves the current settings to the specified settings file.</summary>
        Private Sub Save(ByVal filename As String, Project As Project, Optional create As Boolean = False)
            If create Then
                VBSettings.Instance.WriteJsonFile(filename)
                Project.AddFileToProject(filename, "AdditionalFiles")
                UpdateStatusBar("created")
            Else
                CheckOutFileFromSourceControl(filename)
                VBSettings.Instance.WriteJsonFile(filename)
                UpdateStatusBar("updated")
            End If
        End Sub

        '''<summary>Creates a settings file for the active solution if one does not exist already, initialized from the current settings.</summary>
        Public Sub CreateSolutionSettings()
            Dim path As String = GetSolutionFilePath()
            If File.Exists(path) Then
                Return
            End If

            Save(path, Project:=Nothing, create:=True)
            GetSolutionItemsProject().ProjectItems.AddFromFile(path)
            UpdateStatusBar("created")
        End Sub

#Region "Modern Locater"

        Private Function GetFilePath() As String
            Dim path As String = GetSolutionFilePath()

            If Not File.Exists(path) Then
                path = GetUserFilePath()
            End If

            Return path
        End Function

        Private Function GetSolutionFilePath() As String
            Dim solution As Solution = Command1Package.DTE.Solution

            If solution Is Nothing OrElse String.IsNullOrEmpty(solution.FullName) Then
                Return Nothing
            End If

            Return Path.Combine(Path.GetDirectoryName(solution.FullName), FileName)
        End Function

        Private Function GetProjectFilePath() As String
            Dim solution As Solution = Command1Package.DTE.Solution

            If solution Is Nothing OrElse String.IsNullOrEmpty(solution.FullName) Then
                Return Nothing
            End If

            Return Path.Combine(Path.GetDirectoryName(solution.FullName), FileName)
        End Function

        Private Function GetUserFilePath() As String
            Dim ssm As New ShellSettingsManager(Command1Package.Instance)
            Return Path.Combine(ssm.GetApplicationDataFolder(ApplicationDataFolder.RoamingSettings), FileName)
        End Function

#End Region

        Public Sub UpdateStatusBar(ByVal action As String)
            Try
                Command1Package.DTE.StatusBar.Text = If(SolutionSettingsExist, "Visual Basic Refactoring: Solution settings " & action, "Visual Basic Refactoring: Global settings " & action)
            Catch ex As Exception
                Log("Error updating status bar")
            End Try
        End Sub

    End Module
End Namespace
