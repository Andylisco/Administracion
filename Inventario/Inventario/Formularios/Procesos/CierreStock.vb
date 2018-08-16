﻿Imports System.Data.SqlClient
Imports Inventario.Clases

Public Class CierreStock

    Dim cn As SqlConnection = New SqlConnection()
    Dim cm As SqlCommand = New SqlCommand("")
    Dim trans As SqlTransaction = Nothing

    Dim WLaudos As New DataTable
    Dim WHojas As New DataTable
    Dim WGuias As New DataTable
    Dim WMovVarios As New DataTable
    Dim WMovLaboratorio As New DataTable
    Dim WEstadisticas As New DataTable
    Dim WDevoluciones As New DataTable

    Private Sub btnCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Close()
    End Sub

    Private Sub btnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAceptar.Click

        Try

            If MsgBox("!!! ATENCION !!!   Se actualizara los datos del sistema para mantener la ficha historica de movimientos de stock, Desea realizar el proceso", MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes Then Exit Sub

            If cn.IsOpened Then cn.Close()

            cn.ConnectionString = Helper._ConectarA(Conexion.EmpresaDeTrabajo)
            cn.Open()
            trans = cn.BeginTransaction

            cm.Connection = cn
            cm.Transaction = trans

            With ProgressBar1
                .Value = 0
                .Maximum = {WHojas, WLaudos, WGuias, WMovVarios, WMovLaboratorio, WEstadisticas, WDevoluciones}.Count
            End With

            With lblEstado
                .Text = "Obteniendo Hojas de Producción..."
                .Refresh()
            End With

            WHojas = Query.GetAll("SELECT Clave, Producto, ISNULL(Saldo, 0) Saldo, ISNULL(Real, 0) Real, ISNULL(Marca, '') Marca, ISNULL(RealAnt, 0) RealAnt FROM Hoja ORDER BY Hoja", Conexion.EmpresaDeTrabajo)

            ProgressBar1.Increment(1)

            With lblEstado
                .Text = "Obteniendo Laudos de Materia Prima..."
                .Refresh()
            End With

            WLaudos = Query.GetAll("SELECT Clave, Articulo, ISNULL(LiberadaAnt, 0) LiberadaAnt, ISNULL(DevueltaAnt, 0) DevueltaAnt, ISNULL(Marca, '') Marca, ISNULL(Liberada, 0) Liberada, ISNULL(Devuelta, 0) Devuelta, ISNULL(Saldo, 0) Saldo FROM Laudo ORDER BY Laudo", Conexion.EmpresaDeTrabajo)

            ProgressBar1.Increment(1)

            With lblEstado
                .Text = "Obteniendo Guías de Traslado Interno..."
                .Refresh()
            End With

            WGuias = Query.GetAll("SELECT Clave, Articulo, Terminado, ISNULL(CantidadAnt, 0) CantidadAnt, ISNULL(Cantidad, 0) Cantidad, ISNULL(Marca, '') Marca, ISNULL(Saldo, 0) Saldo, ISNULL(Tipo, '') Tipo FROM Guia ORDER BY Codigo", Conexion.EmpresaDeTrabajo)

            ProgressBar1.Increment(1)

            With lblEstado
                .Text = "Obteniendo Movimientos Varios..."
                .Refresh()
            End With

            WMovVarios = Query.GetAll("SELECT Clave, Articulo, Terminado, ISNULL(Tipo, '') Tipo, ISNULL(Marca, '') Marca, ISNULL(MarcaAnt, '') MarcaAnt FROM MovVar ORDER BY Codigo", Conexion.EmpresaDeTrabajo)

            ProgressBar1.Increment(1)

            With lblEstado
                .Text = "Obteniendo Movimientos de Laboratorio..."
                .Refresh()
            End With

            WMovLaboratorio = Query.GetAll("SELECT Clave, Articulo, Terminado, ISNULL(Tipo, '') Tipo, ISNULL(Marca, '') Marca, ISNULL(MarcaAnt, '') MarcaAnt FROM MovLab ORDER BY Codigo", Conexion.EmpresaDeTrabajo)

            ProgressBar1.Increment(1)

            With lblEstado
                .Text = "Obteniendo Estadísticas de Venta..."
                .Refresh()
            End With

            WEstadisticas = Query.GetAll("SELECT Clave, Ter = ISNULL(Articulo, ''), Art = (LEFT(ISNULL(Articulo, ''), 3) + RIGHT(ISNULL(Articulo, ''), 7)), ISNULL(Marca, '') Marca, ISNULL(MarcaAnt, '') MarcaAnt FROM Estadistica ORDER BY Clave", Conexion.EmpresaDeTrabajo)

            ProgressBar1.Increment(1)

            With lblEstado
                .Text = "Obteniendo Entradas de Devolucion..."
                .Refresh()
            End With

            WDevoluciones = Query.GetAll("SELECT Clave, Terminado, Articulo = (LEFT(Terminado, 3) + RIGHT(Terminado, 7)), ISNULL(Marca, '') Marca, ISNULL(MarcaAnt, '') MarcaAnt FROM EntDev ORDER BY Codigo", Conexion.EmpresaDeTrabajo)

            With lblEstado
                .Text = ""
                .Refresh()
            End With

            With ProgressBar1
                .Value = 0
                .Maximum = {WHojas, WLaudos, WGuias, WMovVarios, WMovLaboratorio, WEstadisticas, WDevoluciones}.Sum(Function(t) t.Rows.Count) + 2
            End With

            '
            ' Procesamos los laudos.
            '
            With lblEstado
                .Text = "Procesando Laudos de Materia Prima..."
                .Refresh()
            End With

            _ProcesarLaudos()

            '
            'Procesamos las Hojas de Produccion.
            '
            With lblEstado
                .Text = "Procesando Hojas de Producción..."
                .Refresh()
            End With

            _ProcesarHojas()

            '
            ' Procesamos las Guias de Traslado Interno.
            '
            With lblEstado
                .Text = "Procesando Guías de Traslado Interno..."
                .Refresh()
            End With

            _ProcesarGuias()

            '
            ' Procesamos los Movimientos Varios.
            '
            With lblEstado
                .Text = "Procesando Movimientos Varios..."
                .Refresh()
            End With

            _ProcesarMovVarios()

            '
            ' Procesamos los Movimientos de Laboratorio.
            '
            With lblEstado
                .Text = "Procesando Movimientos de Laboratorio..."
                .Refresh()
            End With

            _ProcesarMovLaboratorio()

            '
            ' Procesamos las Estadisticas.
            '
            With lblEstado
                .Text = "Procesando Estadísticas de Ventas..."
                .Refresh()
            End With

            _ProcesarEstadisticas()

            '
            ' Procesamos las Devoluciones.
            '
            With lblEstado
                .Text = "Procesando Entradas de Devolución..."
                .Refresh()
            End With

            _ProcesarDevoluciones()

            ProgressBar1.Value = 0
            With lblEstado
                .Text = "¡Finalizado con Éxito!"
                .Refresh()
            End With

            '
            ' Confirmamos los cambios.
            '
            trans.Commit()

            'MsgBox("No hay errores")

        Catch ex As Exception

            If Not IsNothing(trans) AndAlso Not IsNothing(trans.Connection) Then trans.Rollback()

            MsgBox(ex.Message)

            cn.Close()
            cm.Dispose()

        End Try

    End Sub

    Private Sub _ProcesarDevoluciones()
        Try

            For Each row As DataRow In WDevoluciones.Rows

                With row

                    If Trim(.Item("MarcaAnt")) <> Trim(.Item("Marca")) Then


                        ' Si comienza por alguno de estos tres, se lo concidera Producto Terminado, sino como Materia Prima.
                        Dim WTerm = .Item("Terminado").ToString.Substring(0, 2).ToUpper

                        If {"NK", "PT", "YQ", "YF"}.Contains(WTerm) Then

                            If Not IsNothing(Conexion.EmpresaDeTrabajo) AndAlso Conexion.EmpresaDeTrabajo.ToUpper = "SURFACTANSA" Then

                                If Helper._ProdTerminadoValidoPtaI(.Item("Terminado")) Then
                                    Continue For
                                End If

                            End If

                        Else

                            If Not IsNothing(Conexion.EmpresaDeTrabajo) AndAlso Conexion.EmpresaDeTrabajo.ToUpper = "SURFACTANSA" Then

                                If Helper._MateriaPrimaValidaPtaI(.Item("Articulo")) Then
                                    Continue For
                                End If

                            End If

                        End If

                        Dim WClave = .Item("Clave")
                        Dim WMarcaAnt = .Item("Marca")

                        cm.CommandText = String.Format("UPDATE EntDev SET MarcaAnt = '{1}' WHERE Clave = '{0}' ", WClave, WMarcaAnt)
                        cm.ExecuteNonQuery()

                    End If

                End With
                ProgressBar1.Increment(1)
            Next

        Catch ex As Exception

            Throw New Exception("Hubo un problema al querer procesar los Movimientos Varios en la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)

        End Try
    End Sub

    Private Sub _ProcesarEstadisticas()
        Try

            For Each row As DataRow In WEstadisticas.Rows

                With row

                    If Trim(.Item("MarcaAnt")) <> Trim(.Item("Marca")) Then

                        ' Si comienza por alguno de estos tres, se lo concidera Producto Terminado, sino como Materia Prima.
                        Dim WTerm  = .Item("Ter").ToString.Substring(0, 2).ToUpper
                        If {"PT", "YQ", "YF"}.Contains(WTerm) Then

                            If Not IsNothing(Conexion.EmpresaDeTrabajo) AndAlso Conexion.EmpresaDeTrabajo.ToUpper = "SURFACTANSA" Then

                                If Helper._ProdTerminadoValidoPtaI(.Item("Terminado")) Then
                                    Continue For
                                End If

                            End If

                        Else

                            If Not IsNothing(Conexion.EmpresaDeTrabajo) AndAlso Conexion.EmpresaDeTrabajo.ToUpper = "SURFACTANSA" Then

                                If Helper._MateriaPrimaValidaPtaI(.Item("Articulo")) Then
                                    Continue For
                                End If

                            End If

                        End If

                        Dim WClave = .Item("Clave")
                        Dim WMarcaAnt = .Item("Marca")

                        cm.CommandText = String.Format("UPDATE Estadistica SET MarcaAnt = '{1}' WHERE Clave = '{0}' ", WClave, WMarcaAnt)
                        cm.ExecuteNonQuery()

                    End If

                End With
                ProgressBar1.Increment(1)
            Next

        Catch ex As Exception

            Throw New Exception("Hubo un problema al querer procesar los Movimientos Varios en la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)

        End Try
    End Sub

    Private Sub _ProcesarMovLaboratorio()
        Try

            For Each row As DataRow In WMovLaboratorio.Rows

                With row

                    If Trim(.Item("MarcaAnt")) <> Trim(.Item("Marca")) Then

                        If Not IsNothing(Conexion.EmpresaDeTrabajo) AndAlso Conexion.EmpresaDeTrabajo.ToUpper = "SURFACTANSA" Then

                            If .Item("Tipo").ToString.ToUpper = "M" Then
                                If Helper._MateriaPrimaValidaPtaI(.Item("Articulo")) Then
                                    Continue For
                                End If
                            Else
                                If Helper._ProdTerminadoValidoPtaI(.Item("Terminado")) Then
                                    Continue For
                                End If
                            End If

                        End If

                        Dim WClave = .Item("Clave")
                        Dim WMarcaAnt = .Item("Marca")

                        cm.CommandText = String.Format("UPDATE MovLab SET MarcaAnt = '{1}' WHERE Clave = '{0}' ", WClave, WMarcaAnt)
                        cm.ExecuteNonQuery()

                    End If

                End With
                ProgressBar1.Increment(1)
            Next

        Catch ex As Exception

            Throw New Exception("Hubo un problema al querer procesar los Movimientos Varios en la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)

        End Try
    End Sub

    Private Sub _ProcesarMovVarios()
        Try

            For Each row As DataRow In WMovVarios.Rows

                With row

                    If Trim(.Item("MarcaAnt")) <> Trim(.Item("Marca")) Then

                        If Not IsNothing(Conexion.EmpresaDeTrabajo) AndAlso Conexion.EmpresaDeTrabajo.ToUpper = "SURFACTANSA" Then

                            If .Item("Tipo").ToString.ToUpper = "M" Then
                                If Helper._MateriaPrimaValidaPtaI(.Item("Articulo")) Then
                                    Continue For
                                End If
                            Else
                                If Helper._ProdTerminadoValidoPtaI(.Item("Terminado")) Then
                                    Continue For
                                End If
                            End If

                        End If

                        Dim WClave = .Item("Clave")
                        Dim WMarcaAnt = .Item("Marca")

                        cm.CommandText = String.Format("UPDATE MovVar SET MarcaAnt = '{1}' WHERE Clave = '{0}' ", WClave, WMarcaAnt)
                        cm.ExecuteNonQuery()

                    End If

                End With
                ProgressBar1.Increment(1)
            Next

        Catch ex As Exception

            Throw New Exception("Hubo un problema al querer procesar los Movimientos Varios en la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)

        End Try
    End Sub

    Private Sub _ProcesarGuias()
        Try

            For Each row As DataRow In WGuias.Rows

                With row

                    If .Item("CantidadAnt") = 0 And .Item("Cantidad") <> 0 Then

                        If Not IsNothing(Conexion.EmpresaDeTrabajo) AndAlso Conexion.EmpresaDeTrabajo.ToUpper = "SURFACTANSA" Then

                            If .Item("Tipo").ToString.ToUpper = "M" Then
                                If Helper._MateriaPrimaValidaPtaI(.Item("Articulo")) Then
                                    Continue For
                                End If
                            Else
                                If Helper._ProdTerminadoValidoPtaI(.Item("Terminado")) Then
                                    Continue For
                                End If
                            End If
                            
                        End If

                        Dim WClave = .Item("Clave")
                        Dim WMarcaAnt = .Item("Marca")
                        Dim WSaldoAnt = Helper.formatonumerico(.Item("Saldo"))
                        Dim WCantidadAnt = Helper.formatonumerico(.Item("Cantidad"))

                        cm.CommandText = String.Format("UPDATE Guia SET MarcaAnt = '{1}', SaldoAnt = {2}, CantidadAnt = {3} WHERE Clave = '{0}' ", WClave, WMarcaAnt, WSaldoAnt, WCantidadAnt)
                        cm.ExecuteNonQuery()

                    End If

                End With
                ProgressBar1.Increment(1)
            Next

        Catch ex As Exception

            Throw New Exception("Hubo un problema al querer procesar las Guias de Traslado Interno en la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)

        End Try
    End Sub

    Private Sub _ProcesarLaudos()

        Try

            For Each row As DataRow In WLaudos.Rows

                With row

                    If .Item("LiberadaAnt") = 0 And .Item("DevueltaAnt") = 0 Then
                        If .Item("Liberada") <> 0 Or .Item("Devuelta") <> 0 Then

                            If Not IsNothing(Conexion.EmpresaDeTrabajo) AndAlso Conexion.EmpresaDeTrabajo.ToUpper = "SURFACTANSA" Then
                                If Helper._MateriaPrimaValidaPtaI(.Item("Articulo")) Then
                                    Continue For
                                End If
                            End If

                            Dim WClave = .Item("Clave")
                            Dim WMarcaAnt = .Item("Marca")
                            Dim WSaldoAnt = Helper.formatonumerico(.Item("Saldo"))
                            Dim WLiberadaAnt = Helper.formatonumerico(.Item("Liberada"))
                            Dim WDevueltaAnt = Helper.formatonumerico(.Item("Devuelta"))

                            cm.CommandText = String.Format("UPDATE Laudo SET MarcaAnt = '{1}', SaldoAnt = {2}, LiberadaAnt = {3}, DevueltaAnt = {4} WHERE Clave = '{0}' ", WClave, WMarcaAnt, WSaldoAnt, WLiberadaAnt, WDevueltaAnt)
                            cm.ExecuteNonQuery()

                        End If
                    End If

                End With
                ProgressBar1.Increment(1)
            Next

        Catch ex As Exception

            Throw New Exception("Hubo un problema al querer procesar los laudos en la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)
            
        End Try

    End Sub

    Private Sub _ProcesarHojas()

        Try

            For Each row As DataRow In WHojas.Rows

                With row

                    If .Item("RealAnt") = 0 And .Item("Real") <> 0 Then

                        If Not IsNothing(Conexion.EmpresaDeTrabajo) AndAlso Conexion.EmpresaDeTrabajo.ToUpper = "SURFACTANSA" Then
                            If Helper._ProdTerminadoValidoPtaI(.Item("Producto")) Then
                                Continue For
                            End If
                        End If

                        Dim WClave = .Item("Clave")
                        Dim WMarcaAnt = .Item("Marca")
                        Dim WSaldoAnt = Helper.formatonumerico(.Item("Saldo"))
                        Dim WRealAnt = Helper.formatonumerico(.Item("Real"))

                        cm.CommandText = String.Format("UPDATE Hoja SET MarcaAnt = '{1}', SaldoAnt = {2}, RealAnt = {3} WHERE Clave = '{0}' ", _
                                                       WClave, WMarcaAnt, WSaldoAnt, WRealAnt)
                        cm.ExecuteNonQuery()

                    End If

                End With
                ProgressBar1.Increment(1)
            Next

        Catch ex As Exception

            Throw New Exception("Hubo un problema al querer procesar las Hojas de Producción en la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)

        End Try

    End Sub

    Private Sub CierreStock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ProgressBar1.Value = 0
        lblEstado.Text = ""
    End Sub

End Class