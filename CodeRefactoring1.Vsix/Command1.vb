Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.ComponentModel.Design
Imports System.Diagnostics
Imports System.Globalization
Imports CodeRefactoring1.Vsix.EditorExtensions.Settings
Imports Microsoft.VisualStudio.Shell
Imports Microsoft.VisualStudio.Shell.Interop

''' <summary>
''' Command handler
''' </summary>
Public NotInheritable Class Command1

    ''' <summary>
    ''' Command ID.
    ''' </summary>
    Public Const CommandId As Integer = 256

    ''' <summary>
    ''' Command menu group (command set GUID).
    ''' </summary>
    Public Shared ReadOnly CommandSet As New Guid("7a657668-eed9-4916-a17f-1304b4918b4b")

    ''' <summary>
    ''' VS Package that provides this command, not null.
    ''' </summary>
    Private ReadOnly package As Package
    ''' <summary>
    ''' Initializes a new instance of the <see cref="Command1"/> class.
    ''' Adds our command handlers for menu (the commands must exist in the command table file)
    ''' </summary>
    ''' <param name="package">Owner package, not null.</param>
    Private Sub New(package As Package)
        If package Is Nothing Then
            Throw New ArgumentNullException(NameOf(package))
        End If
        Debug.Write($"In {NameOf(Command1)} Line 42")
        Stop
        Me.package = package
        Dim commandService As OleMenuCommandService = CType(Me.ServiceProvider.GetService(GetType(IMenuCommandService)), OleMenuCommandService)
        If commandService IsNot Nothing Then
            Dim menuCommandId As CommandID = New CommandID(CommandSet, CommandId)
            'Dim menuCommand As MenuCommand = New MenuCommand(AddressOf Me.MenuItemCallback, menuCommandId)
            'commandService.AddCommand(menuCommand)
        End If
    End Sub

    ''' <summary>
    ''' Gets the instance of the command.
    ''' </summary>
    Public Shared Property Instance As Command1

    ''' <summary>
    ''' Get service provider from the owner package.
    ''' </summary>
    Private ReadOnly Property ServiceProvider As IServiceProvider
        Get
            Return Me.package
        End Get
    End Property

    ''' <summary>
    ''' Initializes the singleton instance of the command.
    ''' </summary>
    ''' <param name="package">Owner package, Not null.</param>
    Public Shared Sub Initialize(package As Package)
        Debug.Write($"In {NameOf(Command1)} Line 72")
        Stop
        Instance = New Command1(package)
    End Sub

    ''' <summary>
    ''' This function is the callback used to execute the command when the menu item is clicked.
    ''' See the constructor to see how the menu item is associated with this function using
    ''' OleMenuCommandService service and MenuCommand class.
    ''' </summary>
    ''' <param name="sender">Event sender.</param>
    ''' <param name="e">Event args.</param>
    'Private Sub MenuItemCallback(sender As Object, e As EventArgs)
    '    Dim message As String = String.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", Me.[GetType]().FullName)
    '    Const title As String = "Visual Basic Refactoring"

    '    'Show a message box to prove we were here
    '    VsShellUtilities.ShowMessageBox(
    '        Me.ServiceProvider,
    '        message,
    '        title,
    '        OLEMSGICON.OLEMSGICON_INFO,
    '        OLEMSGBUTTON.OLEMSGBUTTON_OK,
    '        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST)

    '    If VBSettings.Instance.General.CsvIgnoreList Is Nothing Then
    '        LoadProjectSettings()
    '    End If

    '    Using f As New OptionDialog With {
    '        .Text = title
    '    }
    '        f.ShowDialog()
    '    End Using
    'End Sub
End Class
