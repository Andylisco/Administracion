﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.ReportSource
Imports CrystalDecisions.Shared
Imports System
Imports System.ComponentModel


Public Class NuevoSACAmbos
    Inherits ReportClass
    
    Public Sub New()
        MyBase.New
    End Sub
    
    Public Overrides Property ResourceName() As String
        Get
            Return "NuevoSACAmbos.rpt"
        End Get
        Set
            'Do nothing
        End Set
    End Property
    
    Public Overrides Property NewGenerator() As Boolean
        Get
            Return true
        End Get
        Set
            'Do nothing
        End Set
    End Property
    
    Public Overrides Property FullResourceName() As String
        Get
            Return "SAC.NuevoSACAmbos.rpt"
        End Get
        Set
            'Do nothing
        End Set
    End Property
    
    <Browsable(False), _
     DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden)> _
    Public ReadOnly Property Section1() As Section
        Get
            Return Me.ReportDefinition.Sections(0)
        End Get
    End Property

    <Browsable(False), _
     DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden)> _
    Public ReadOnly Property Section2() As Section
        Get
            Return Me.ReportDefinition.Sections(1)
        End Get
    End Property

    <Browsable(False), _
     DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden)> _
    Public ReadOnly Property Section3() As Section
        Get
            Return Me.ReportDefinition.Sections(2)
        End Get
    End Property

    <Browsable(False), _
     DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden)> _
    Public ReadOnly Property DetailSection1() As Section
        Get
            Return Me.ReportDefinition.Sections(3)
        End Get
    End Property

    <Browsable(False), _
     DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden)> _
    Public ReadOnly Property DetailSection2() As Section
        Get
            Return Me.ReportDefinition.Sections(4)
        End Get
    End Property

    <Browsable(False), _
     DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden)> _
    Public ReadOnly Property DetailSection3() As Section
        Get
            Return Me.ReportDefinition.Sections(5)
        End Get
    End Property

    <Browsable(False), _
     DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden)> _
    Public ReadOnly Property Section4() As Section
        Get
            Return Me.ReportDefinition.Sections(6)
        End Get
    End Property

    <Browsable(False), _
     DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden)> _
    Public ReadOnly Property Section5() As Section
        Get
            Return Me.ReportDefinition.Sections(7)
        End Get
    End Property

    <Browsable(False), _
     DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden)> _
    Public ReadOnly Property Parameter_Comentarios() As IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(0)
        End Get
    End Property

    <Browsable(False), _
     DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden)> _
    Public ReadOnly Property Parameter_SubInfNuevoSACAccionesImpleVerifrpt_Comentarios() As IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(1)
        End Get
    End Property
End Class

<ToolboxBitmap(GetType(ExportOptions), "report.bmp")> _
Public Class CachedNuevoSACAmbos
    Inherits Component
    Implements ICachedReport

    Public Sub New()
        MyBase.New()
    End Sub

    <Browsable(False), _
     DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden)> _
    Public Overridable Property IsCacheable() As Boolean Implements ICachedReport.IsCacheable
        Get
            Return True
        End Get
        Set(ByVal value As Boolean)
            '
        End Set
    End Property

    <Browsable(False), _
     DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden)> _
    Public Overridable Property ShareDBLogonInfo() As Boolean Implements ICachedReport.ShareDBLogonInfo
        Get
            Return False
        End Get
        Set(ByVal value As Boolean)
            '
        End Set
    End Property

    <Browsable(False), _
     DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden)> _
    Public Overridable Property CacheTimeOut() As TimeSpan Implements ICachedReport.CacheTimeOut
        Get
            Return CachedReportConstants.DEFAULT_TIMEOUT
        End Get
        Set(ByVal value As TimeSpan)
            '
        End Set
    End Property

    Public Overridable Function CreateReport() As ReportDocument Implements ICachedReport.CreateReport
        Dim rpt As NuevoSACAmbos = New NuevoSACAmbos()
        rpt.Site = Me.Site
        Return rpt
    End Function

    Public Overridable Function GetCustomizedCacheKey(ByVal request As RequestContext) As String Implements ICachedReport.GetCustomizedCacheKey
        Dim key As [String] = Nothing
        '// The following is the code used to generate the default
        '// cache key for caching report jobs in the ASP.NET Cache.
        '// Feel free to modify this code to suit your needs.
        '// Returning key == null causes the default cache key to
        '// be generated.
        '
        'key = RequestContext.BuildCompleteCacheKey(
        '    request,
        '    null,       // sReportFilename
        '    this.GetType(),
        '    this.ShareDBLogonInfo );
        Return key
    End Function
End Class
