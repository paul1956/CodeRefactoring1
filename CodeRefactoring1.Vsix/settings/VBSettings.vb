Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On


Imports ConfOxide
Imports System.Diagnostics.CodeAnalysis

Namespace EditorExtensions.Settings
    Public NotInheritable Class VBSettings
        Inherits SettingsBase(Of VBSettings)
        <SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")>
        Public Shared ReadOnly Instance As New VBSettings()
        Private m_General As GeneralSettings

        Public Property General() As GeneralSettings
            Get
                Return Me.m_General
            End Get
            Private Set
                Me.m_General = Value
            End Set
        End Property
    End Class

    Public NotInheritable Class GeneralSettings
        Inherits SettingsBase(Of GeneralSettings)
        Implements INameOfSettings

#Disable Warning IDE0032 ' Use auto property
        Private m_AllMessagesToOutputWindow As Boolean = True
#Enable Warning IDE0032 ' Use auto property
        Public Property AllMessagesToOutputWindow() As Boolean
            Get
                Return Me.m_AllMessagesToOutputWindow
            End Get
            Set
                Me.m_AllMessagesToOutputWindow = Value
            End Set
        End Property

#Disable Warning IDE0032 ' Use auto property
        Private _csvIgnoreList As String = "Microsoft,System,MS"
#Enable Warning IDE0032 ' Use auto property
        Public Property CsvIgnoreList As String Implements INameOfSettings.CsvIgnoreList
            Get
                Return Me._csvIgnoreList
            End Get
            Set(value As String)
                Me._csvIgnoreList = value
            End Set
        End Property

    End Class

    Public Interface INameOfSettings
        Property CsvIgnoreList As String
    End Interface

End Namespace
