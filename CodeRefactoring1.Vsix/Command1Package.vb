﻿' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Diagnostics
Imports System.Runtime.InteropServices

Imports EnvDTE

Imports EnvDTE80
Imports Microsoft.VisualStudio.ComponentModelHost
Imports Microsoft.VisualStudio.Shell
Imports Microsoft.VisualStudio.Shell.Interop

''' <summary>
''' This is the class that implements the package exposed by this assembly.
''' </summary>
''' <remarks>
''' <para>
''' The minimum requirement for a class to be considered a valid package for Visual Studio
''' Is to implement the IVsPackage interface And register itself with the shell.
''' This package uses the helper classes defined inside the Managed Package Framework (MPF)
''' to do it: it derives from the Package Class that provides the implementation Of the 
''' IVsPackage interface And uses the registration attributes defined in the framework to 
''' register itself And its components with the shell. These attributes tell the pkgdef creation
''' utility what data to put into .pkgdef file.
''' </para>
''' <para>
''' To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
''' </para>
''' </remarks>
<PackageRegistration(UseManagedResourcesOnly:=True)>
<InstalledProductRegistration("#110", "#112", "1.0", IconResourceID:=400)>
<ProvideMenuResource("Menus.ctmenu", 1)>
<Guid(Command1Package.PackageGuidString)>
Public NotInheritable Class Command1Package
    Inherits Package

    ''' <summary>
    ''' Package guid
    ''' </summary>
    Public Const PackageGuidString As String = "30854f76-d2a3-40ef-ba0f-743dfd93ef06"

    ''' <summary>
    ''' Default constructor of the package.
    ''' Inside this method you can place any initialization code that does not require 
    ''' any Visual Studio service because at this point the package object is created but 
    ''' not sited yet inside Visual Studio environment. The place to do all the other 
    ''' initialization is the Initialize method.
    ''' </summary>
    Public Sub New()
        Stop
    End Sub

#Region "Package Members"

    ''' <summary>
    ''' Initialization of the package; this method is called right after the package is sited, so this is the place
    ''' where you can put all the initialization code that rely on services provided by VisualStudio.
    ''' </summary>
    Protected Overrides Sub Initialize()
        Debug.Write($"In {NameOf(Command1Package)} Line 59")
        Stop
        Command1.Initialize(Me)
        MyBase.Initialize()
        MyComponentModel = DirectCast(Me.GetService(GetType(SComponentModel)), IComponentModel)
        Dim defaultListener As DefaultTraceListener
        defaultListener = New DefaultTraceListener() With {.LogFileName = "C:\Users\PaulM\Documents\CodeRefactoringTrace.txt", .IndentSize = 4}
        Trace.Listeners.Add(defaultListener)
        Trace.WriteLine("Text")
    End Sub

#End Region

    Public Shared Property MyComponentModel As IComponentModel

    Private Shared _dte As DTE2
    Private Shared _pct As IVsRegisterPriorityCommandTarget
    'Private _topMenu As OleMenuCommand

    Friend Shared ReadOnly Property DTE() As DTE2
        Get
            If _dte Is Nothing Then
                _dte = TryCast(ServiceProvider.GlobalProvider.GetService(GetType(DTE)), DTE2)
            End If

            Return _dte
        End Get
    End Property

    Friend Shared ReadOnly Property PriorityCommandTarget() As IVsRegisterPriorityCommandTarget
        Get
            If _pct Is Nothing Then
                _pct = TryCast(ServiceProvider.GlobalProvider.GetService(GetType(SVsRegisterPriorityCommandTarget)), IVsRegisterPriorityCommandTarget)
            End If

            Return _pct
        End Get
    End Property

    Private Shared privateInstance As Command1Package

    Public Shared Property Instance() As Command1Package
        Get
            Return privateInstance
        End Get
        Private Set(value As Command1Package)
            privateInstance = value
        End Set
    End Property

End Class
