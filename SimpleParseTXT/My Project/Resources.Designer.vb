﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'This class was auto-generated by the StronglyTypedResourceBuilder
    'class via a tool like ResGen or Visual Studio.
    'To add or remove a member, edit your .ResX file then rerun ResGen
    'with the /str option, or rebuild your VS project.
    '''<summary>
    '''  A strongly-typed resource class, for looking up localized strings, etc.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Returns the cached ResourceManager instance used by this class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("SimpleParseTXT.Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Overrides the current thread's CurrentUICulture property for all
        '''  resource lookups using this strongly typed resource class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to CREATOR.
        '''</summary>
        Friend ReadOnly Property minfoAuthor() As String
            Get
                Return ResourceManager.GetString("minfoAuthor", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to BRIEFING.
        '''</summary>
        Friend ReadOnly Property minfoBriefing() As String
            Get
                Return ResourceManager.GetString("minfoBriefing", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to BRIEFLONG.
        '''</summary>
        Friend ReadOnly Property minfoBriefingLong() As String
            Get
                Return ResourceManager.GetString("minfoBriefingLong", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to DESC.
        '''</summary>
        Friend ReadOnly Property minfoDesc() As String
            Get
                Return ResourceManager.GetString("minfoDesc", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to DEBUNKL.
        '''</summary>
        Friend ReadOnly Property minfoLose() As String
            Get
                Return ResourceManager.GetString("minfoLose", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to NAME.
        '''</summary>
        Friend ReadOnly Property minfoName() As String
            Get
                Return ResourceManager.GetString("minfoName", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to DEBUNKW.
        '''</summary>
        Friend ReadOnly Property minfoWin() As String
            Get
                Return ResourceManager.GetString("minfoWin", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to MUSIC.
        '''</summary>
        Friend ReadOnly Property MsgTextEnd() As String
            Get
                Return ResourceManager.GetString("MsgTextEnd", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to POPUP_TXT.
        '''</summary>
        Friend ReadOnly Property MsgTextStart() As String
            Get
                Return ResourceManager.GetString("MsgTextStart", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to PLAYER_1.
        '''</summary>
        Friend ReadOnly Property scenDescEnd() As String
            Get
                Return ResourceManager.GetString("scenDescEnd", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to AVCScenarioInfo.
        '''</summary>
        Friend ReadOnly Property scenDescStart() As String
            Get
                Return ResourceManager.GetString("scenDescStart", resourceCulture)
            End Get
        End Property
    End Module
End Namespace
