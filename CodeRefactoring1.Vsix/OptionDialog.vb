Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Forms
Imports CodeRefactoring1.Vsix.EditorExtensions.Settings
Imports Microsoft.CodeAnalysis
Imports Microsoft.VisualBasic
Imports Microsoft.VisualStudio.LanguageServices

Public Class OptionDialog
    Private ReadOnly MetadataReferenceList As New List(Of String)
    Private ReadOnly ProjectReferenceList As New List(Of ProjectReference)
    Private Sub OK_Button_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = DialogResult.OK
        Dim IgnoreList As New List(Of String)
        For Each li1 As Object In Me.CheckedListBox1.CheckedItems
            IgnoreList.Add(li1.ToString)
        Next li1
        VBSettings.Instance.General.CsvIgnoreList = String.Join(",", IgnoreList)
        SaveProjectSettings()
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub OptionDialog_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Command1Package.MyComponentModel IsNot Nothing Then

            Me.DataGridView1.Rows.Clear()
            Dim workspace As VisualStudioWorkspace = Command1Package.MyComponentModel.GetService(Of VisualStudioWorkspace)()
            Dim SolutionPath As String = workspace.CurrentSolution.FilePath
            Me.CheckedListBox1.Items.Clear()
            Dim IgnoreList As New List(Of String) From {
                "MS"
            }
            Me.CheckedListBox1.Items.Add("MS", VBSettings.Instance.General.CsvIgnoreList.Contains("MS"))
            For Each project As Project In workspace.CurrentSolution.Projects
                Dim builder1 As New System.Text.StringBuilder()
                For Each reference As MetadataReference In project.MetadataReferences
                    'If MetadataReferenceList.Contains(reference.Display) Then
                    '    Continue For
                    'End If
                    Me.MetadataReferenceList.Add(reference.Display)
                    Me.DataGridView1.Rows.Add()
                    Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1).Cells(0).Value = project.Name
                    Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1).Cells(1).Value = Path.GetFileNameWithoutExtension(reference.Display)
                    Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1).Cells(2).Value = Path.GetDirectoryName(reference.Display)
                    Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1).Cells(3).Value = Path.GetFileName(reference.Display)
                    Dim SplitFileNameReference() As String = Path.GetFileName(reference.Display).Split(CChar("."))
                    If SplitFileNameReference.Length <= 2 OrElse IgnoreList.Contains(SplitFileNameReference(0)) Then
                        Continue For
                    End If
                    IgnoreList.Add(SplitFileNameReference(0))
                    Dim checkboxName As String = SplitFileNameReference(0)
                    builder1.Append($"{checkboxName}{vbCrLf}")
                    Me.CheckedListBox1.Items.Add(checkboxName, VBSettings.Instance.General.CsvIgnoreList.Contains(checkboxName))
                Next
                Dim builder As New System.Text.StringBuilder()
                For Each reference As ProjectReference In project.ProjectReferences
                    Me.ProjectReferenceList.Add(reference)
                    builder.Append($"Project Name = {project.Name}, Reference = {reference.ProjectId.ToString.Replace(reference.ProjectId.Id.ToString, String.Empty).Trim}{vbCrLf}")
                Next
                Me.ProjectReferenceTextBox.Text = builder.ToString()
            Next
        End If
    End Sub
End Class
