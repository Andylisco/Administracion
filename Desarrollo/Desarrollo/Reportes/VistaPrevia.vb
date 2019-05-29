Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Text.RegularExpressions

Public Class VistaPrevia
    Public Property Reporte As ReportDocument

    Public Property Formula As String
    
    Private Sub Reporte_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        With Me.CrystalReportViewer1
            .ReportSource = Me.Reporte
            '.Refresh()
        End With

    End Sub

    Public Sub EstablecerConexion(ByVal Servidor As String, ByVal BaseDatos As String)
        With Me.Reporte
            .DataSourceConnections.Item(0).SetConnection(Trim(Servidor), Trim(BaseDatos), False)
            '.Refresh()
        End With
    End Sub

    Private Sub _ReconectarBaseDatos()

        ' MANDAMOS EL PARÁMETRO DE LA EMPRESA.

        If Reporte.ParameterFields.Count > 0 Then
            Reporte.SetParameterValue(0, ClasesCompartidas.Globals.NombreEmpresa)
        End If

        ' CONECTAMOS CON LA BASE DE DATOS QUE CORRESPONDA.
        Dim cs = ""

        Try
            ' Buscamos el string de conexion.
            cs = Helper._ConectarA 'ClasesCompartidas.Globals.getConnectionString()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            Return
        End Try

        ' Extraemos los datos de conexion del string de conexion.
        Dim cnsb As New SqlClient.SqlConnectionStringBuilder(cs)

        ' Asignamos los datos al reporte.
        Reporte.SetDatabaseLogon(cnsb.UserID, cnsb.Password, cnsb.DataSource, cnsb.InitialCatalog)

        Dim conexion As New ConnectionInfo
        conexion.DatabaseName = cnsb.InitialCatalog
        conexion.ServerName = cnsb.DataSource
        conexion.UserID = cnsb.UserID
        conexion.Password = cnsb.Password
        conexion.IntegratedSecurity = True

        Dim tli As New TableLogOnInfo()
        tli.ConnectionInfo = conexion

        ' Volvemos a asignar los datos de conexion pero ahora a cada una de las tablas que tenga el reporte.
        For Each tabla As Table In Reporte.Database.Tables

            Dim _logInfo As TableLogOnInfo = tabla.LogOnInfo

            _logInfo.ConnectionInfo = conexion

            tabla.ApplyLogOnInfo(_logInfo)

        Next
    End Sub

    Public Sub Mostrar()

        _ReconectarBaseDatos()

        With Me.CrystalReportViewer1

            .ReportSource = Me.Reporte

            If Not String.IsNullOrEmpty(Me.Formula) Then

                .SelectionFormula = Me.Formula

            End If

            '.RefreshReport()

        End With

        Me.Show()

    End Sub

    Public Sub Imprimir(Optional ByVal cant As Integer = 1)
        'Me.Reporte.DataSourceConnections.Item(0).SetConnection("EMPRESA01", "SurfactanSA", False)

        _ReconectarBaseDatos()

        Me.Reporte.RecordSelectionFormula = IIf(IsNothing(Me.Formula), "", Me.Formula)
        'Me.Reporte.Refresh()
        Me.Reporte.PrintToPrinter(cant, True, 0, 0)
    End Sub

    Public Sub GuardarPDF(ByVal NombreArchivo As String, Optional ByVal ruta As String = "")
        ruta = IIf(ruta = "", Application.StartupPath & "/", ruta)

        NombreArchivo = IIf(Regex.IsMatch(NombreArchivo, "(\.pdf)$"), NombreArchivo, NombreArchivo & ".pdf")

        'Me.Reporte.DataSourceConnections.Item(0).SetConnection("EMPRESA01", "SurfactanSA", False)

        _ReconectarBaseDatos()

        Me.Reporte.RecordSelectionFormula = IIf(IsNothing(Me.Formula), "", Me.Formula)
        'Me.Reporte.Refresh()
        Me.Reporte.ExportToDisk(ExportFormatType.PortableDocFormat, ruta & NombreArchivo)
    End Sub

    Public Sub DesdeArchivo(ByVal s As String)
        Reporte = New ReportDocument
        Reporte.Load(s)

        _ReconectarBaseDatos()

    End Sub
End Class