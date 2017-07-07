﻿Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class DetallesRemitosProveedor

    Private _Consulta() As String
    Private _CodProv, _NombreProv As String

    Public Sub New(ByVal consulta As String)
        InitializeComponent()

        _Consulta = consulta.Split("$")

        _CodProv = _Consulta(0) ' Código numérico.
        _NombreProv = UCase(_Consulta(1))    ' Nombre del Proveedor.

        txtCodProveedor.Text = _CodProv
        txtNombreProveedor.Text = _NombreProv

    End Sub

    Private Sub DetallesRemitosProveedor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _AlinearCeldas()
        _TraerInfoRemitos()

    End Sub

    'Definimos las alineaciones por defecto.
    Private Sub _AlinearCeldas()
        With DGVDetalles
            .Columns("Remito").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns("Orden").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns("CantidadPedida").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns("Precio").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns("Remito").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns("Informe").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns("CantRecibida").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End With
    End Sub

    Private Sub _TraerInfoRemitos()
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand()
        Dim dr As SqlDataReader
        Dim remitos(), codProveedor, proveedor, orden, articulo, descripcion, cantPed As String
        Dim moneda, precio, condPago, informe, est, fApr, cantRecibida As String
        Dim aBuscar As New List(Of Object)
        Dim Empresas As New List(Of String) From {"SurfactanSA", "surfactan_II", "Surfactan_III", "Surfactan_IV", "Surfactan_V", "Surfactan_VI", "Surfactan_VII"}

        remitos = _Consulta(2).Replace(" ", "").Split(",") ' Eliminamos los posibles espacios en blanco y luego los separamos por las comas.

        ' Por cada remito indicado, recorro las empresas buscándolo. Extraigo el num de Orden, los articulos y el num de informe.
        For Each remitoActual In remitos

            For Each empresa In Empresas

                Try
                    cn.ConnectionString = _ConnectionString(empresa)

                    cm.CommandText = "SELECT orden, articulo, informe FROM Informe WHERE Remito = '" + remitoActual + "' AND Proveedor = '" + _CodProv + "'"
                    cm.Connection = cn

                    cn.Open()

                    dr = cm.ExecuteReader()

                    If dr.HasRows Then

                        Do While dr.Read()
                            orden = dr.Item(0)
                            articulo = dr.Item(1)
                            informe = dr.Item(2)

                            aBuscar.Add({_ConnectionString(empresa), remitoActual, _CodProv, orden, articulo, informe})
                        Loop

                        Exit For ' Salimos en la primer aparicion del remito en la primer empresa.
                    End If

                Catch ex As Exception
                    Throw New Exception("Hubo un error al querer buscar al proveedor indicado")
                Finally
                    cn.Close()
                End Try
            Next

        Next

        _TraerDetallesDeRemitos(aBuscar)

    End Sub

    Private Sub _TraerDetallesDeRemitos(ByVal RemitosABuscar As List(Of Object))
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand()
        Dim dr As SqlDataReader
        Dim _ConnectionString, remito, codProveedor, orden, articulo, descripcion, cantPed As String
        Dim moneda, precio, condPago, informe, est, fApr, cantRecibida As String

        For Each _remitoActual In RemitosABuscar

            _ConnectionString = _remitoActual(0)
            remito = _remitoActual(1)
            codProveedor = _remitoActual(2)
            orden = _remitoActual(3)
            articulo = _remitoActual(4)
            informe = _remitoActual(5)

            cn.ConnectionString = _ConnectionString

            Try ' Buscamos la decripción del Artículo
                cm.CommandText = "SELECT Descripcion FROM Articulo WHERE Codigo = '" + articulo + "'"
                cm.Connection = cn

                cn.Open()

                dr = cm.ExecuteReader()

                If dr.HasRows Then

                    dr.Read()

                    descripcion = dr.Item(0)

                End If
            Catch ex As Exception
                Throw New Exception("Hubo un error al querer buscar al proveedor indicado")
                Exit Sub
            Finally
                cn.Close()
            End Try

            Try ' Buscamos los datos de la Orden
                cm.CommandText = "SELECT cantidad, moneda, precio, condicion FROM Orden WHERE Orden = '" + orden + "' AND Articulo = '" + articulo + "'"
                cm.Connection = cn

                cn.Open()

                dr = cm.ExecuteReader()

                If dr.HasRows Then

                    dr.Read()

                    cantPed = _FormatearDecimales(dr.Item(0))
                    precio = _FormatearDecimales(CDbl(dr.Item(2)))
                    condPago = dr.Item(3)


                    Select Case dr.Item(1)
                        Case 0
                            moneda = "U$S"
                        Case 1
                            moneda = "S"
                        Case 2
                            moneda = "Euro"
                        Case Else
                            moneda = ""
                    End Select

                End If
            Catch ex As Exception
                Throw New Exception("Hubo un error al querer buscar al proveedor indicado")
                Exit Sub
            Finally
                cn.Close()
            End Try

            Try ' Buscamos la información del Laudo
                cm.CommandText = "SELECT liberada, devuelta, fecha FROM Laudo WHERE Orden = '" + orden + "' AND Informe = '" + informe + "' AND Articulo = '" + articulo + "'"
                cm.Connection = cn

                cn.Open()

                dr = cm.ExecuteReader()

                If dr.HasRows Then

                    dr.Read()

                    cantRecibida = _FormatearDecimales(dr.Item(0))


                    If cantRecibida > 0 Then
                        fApr = dr.Item(2)
                        est = "Aprob."
                    ElseIf dr.Item(1) > 0 Then
                        est = "Rech."
                    End If
                Else

                    cantRecibida = _FormatearDecimales(0)
                    fApr = ""
                    est = ""

                End If
            Catch ex As Exception
                Throw New Exception("Hubo un error al querer buscar al proveedor indicado")
                Exit Sub
            Finally
                cn.Close()
                'cm = Nothing
                'dr = Nothing
            End Try

            ' Agregamos la linea en el DGV.
            DGVDetalles.Rows.Add(remito, orden, articulo, descripcion, cantPed, moneda, precio, condPago, informe, cantRecibida, est, fApr)

        Next

    End Sub

    Private Function _FormatearDecimales(ByVal numero As Double)
        Return Format(Math.Round(numero, 2), "0.00")
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim table As New DataTable("Detalles")
        Dim row As DataRow
        Dim crdoc As ReportDocument
        crdoc = New ControlDeFacturacion

        ' Creo las Columnas
        With table
            .Columns.Add("Remito")
            .Columns.Add("Orden")
            .Columns.Add("Articulo")
            .Columns.Add("Descripcion")
            .Columns.Add("CantPedida")
            .Columns.Add("Moneda")
            .Columns.Add("Precio")
            .Columns.Add("Condicion")
            .Columns.Add("Informe")
            .Columns.Add("CantRecibida")
            .Columns.Add("Estado")
            .Columns.Add("FApr")
        End With

        ' Lleno la tabla con la informacion de los remitos.
        For i = 0 To DGVDetalles.Rows.Count - 1

            row = table.NewRow()

            row.Item("Remito") = DGVDetalles.Rows(i).Cells(0).Value.ToString()
            row.Item("Orden") = DGVDetalles.Rows(i).Cells(1).Value.ToString()
            row.Item("Articulo") = DGVDetalles.Rows(i).Cells(2).Value.ToString()
            row.Item("Descripcion") = DGVDetalles.Rows(i).Cells(3).Value.ToString()
            row.Item("CantPedida") = DGVDetalles.Rows(i).Cells(4).Value.ToString()
            row.Item("Moneda") = DGVDetalles.Rows(i).Cells(5).Value.ToString()
            row.Item("Precio") = DGVDetalles.Rows(i).Cells(6).Value.ToString()
            row.Item("Condicion") = DGVDetalles.Rows(i).Cells(7).Value.ToString()
            row.Item("Informe") = DGVDetalles.Rows(i).Cells(8).Value.ToString()
            row.Item("CantRecibida") = DGVDetalles.Rows(i).Cells(9).Value.ToString()
            row.Item("Estado") = DGVDetalles.Rows(i).Cells(10).Value.ToString()
            row.Item("FApr") = DGVDetalles.Rows(i).Cells(11).Value.ToString()

            table.Rows.Add(row)

        Next

        If table.Rows.Count = 0 Then
            Exit Sub
        End If

        crdoc.SetDataSource(table)
        ' Paso parámetros obligatorios.
        crdoc.SetParameterValue("CodProveedor", _CodProv)
        crdoc.SetParameterValue("NombreProveedor", _NombreProv)

        _Imprimir(crdoc)
        '_VistaPrevia(crdoc)
    End Sub

    Private Sub btnCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Function _ConnectionString(ByVal empresa As String) As String
        Return "Data Source=193.168.0.7;Initial Catalog=" + empresa + ";User ID=usuarioadmin; Password=usuarioadmin"
    End Function

    Private Sub _Imprimir(ByVal crdoc As ReportDocument)
        crdoc.PrintToPrinter(1, True, 0, 0)
    End Sub

    Private Sub _VistaPrevia(ByVal crdoc As ReportDocument)
        With VistaPrevia
            .CrystalReportViewer1.ReportSource = crdoc
            .ShowDialog()
            .Dispose()
        End With
    End Sub
End Class