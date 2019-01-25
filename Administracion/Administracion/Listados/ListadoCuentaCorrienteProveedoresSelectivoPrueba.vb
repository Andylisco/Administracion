Imports ClasesCompartidas
Imports System.Data.SqlClient

Public Class ListadoCuentaCorrienteProveedoresSelectivoPrueba

    Dim varRenglon As Integer
    Dim varTotal, varSaldo, varTotalUs, varSaldoUs, varSaldoOriginal, varDife, varParidad, varParidadTotal As Double
    Dim _Claves As New List(Of Object)

    Private Sub ListadoCuentaCorrienteProveedoresSelectivo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label2.Text = Globals.NombreEmpresa()
        txtDesdeProveedor.Text = ""
        txtFechaEmision.Clear()
        varRenglon = 0
        '_CargarProveedoresPreCargados()
        _Claves.Clear()

        Proceso._PurgarSaldosCtaCtePrvs()
    End Sub

    Private Sub _CargarProveedoresPreCargados()
        Dim _Proveedores As New List(Of Object)
        'Dim _CargadosHaceMasDeUnaSemana As Integer = 0
        'Dim _FechaLimite As String = _DeterminarFechaLimite()
        Dim cn = New SqlConnection()
        Dim cm = New SqlCommand("SELECT Proveedor, FechaOrd FROM ProveedorSelectivo WHERE Fecha = '" & txtFechaEmision.Text & "'")
        Dim dr As SqlDataReader

        SQLConnector.conexionSql(cn, cm)

        Try

            GRilla.Rows.Clear()
            dr = cm.ExecuteReader()

            If dr.HasRows Then

                Do While dr.Read()

                    _Proveedores.Add({dr.Item("Proveedor"), dr.Item("FechaOrd")})

                Loop
                'Else
                'MsgBox("No hay proveedores que listar.", MsgBoxStyle.Information)
            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar los Proveedores Selectivos precargados en la Base de Datos.", MsgBoxStyle.Critical)
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

        For Each _Proveedor As Object In _Proveedores
            _CargarProveedor(DAOProveedor.buscarProveedorPorCodigo(_Proveedor(0)))
        Next

        GRilla.Sort(GRilla.Columns(1), System.ComponentModel.ListSortDirection.Ascending)

    End Sub

    Private Sub _CargarProveedor(ByVal CampoProveedor As Proveedor)
        If IsNothing(CampoProveedor) Then
            MsgBox("Proveedor incorrecto")
        Else


            If txtFechaEmision.Text.Replace(" ", "").Length < 10 Then
                MsgBox("Debe indicarse una fecha de Pago antes de cargar un Proveedor.", MsgBoxStyle.Exclamation)
                txtFechaEmision.Focus()
                Exit Sub
            End If


            varRenglon = GRilla.Rows.Add()
            GRilla.Item(0, varRenglon).Value = CampoProveedor.id
            GRilla.Item(1, varRenglon).Value = CampoProveedor.razonSocial

            GRilla.CommitEdit(DataGridViewDataErrorContexts.Commit)
            GRilla.CurrentCell = GRilla(0, 0)

            txtDesdeProveedor.Text = ""
            txtRazon.Text = ""
            txtDesdeProveedor.Focus()
        End If
    End Sub

    Private Function _ProveedorYaAgregado(ByVal _Proveedor As String) As Boolean
        Dim _YaAgregado = False

        For Each row As DataGridViewRow In GRilla.Rows

            If Trim(row.Cells(0).Value) = Trim(_Proveedor) Then
                _YaAgregado = True
                Exit For
            ElseIf Trim(row.Cells(1).Value) = Trim(_Proveedor) Then
                _YaAgregado = True
                Exit For
            End If

        Next

        Return _YaAgregado
    End Function

    Private Function _ProveedorYaAgregado(ByVal _Proveedor As String, ByVal Excepto As Integer) As Boolean
        Dim _YaAgregado = False

        For Each row As DataGridViewRow In GRilla.Rows

            If Trim(row.Cells(0).Value) = Trim(_Proveedor) And row.Index <> Excepto Then
                _YaAgregado = True
                Exit For
            ElseIf Trim(row.Cells(1).Value) = Trim(_Proveedor) And row.Index <> Excepto Then
                _YaAgregado = True
                Exit For
            End If

        Next

        Return _YaAgregado
    End Function

    Private Sub txtfechaemision_KeyPress(ByVal sender As Object, _
               ByVal e As System.Windows.Forms.KeyPressEventArgs) _
               Handles txtFechaEmision.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            If ValidaFecha(txtFechaEmision.Text) = "S" Then

                Dim CampoTipoCambio As TipoDeCambio = DAOTipoCambio.buscarTipoCambioPorFechaPago(txtFechaEmision.Text)
                If IsNothing(CampoTipoCambio) Then
                    If txtFechaEmision.Text = Date.Now.ToString("dd/MM/yyyy") Then
                        MsgBox("Paridad Inexistente")
                        txtFechaEmision.Focus()
                        Exit Sub
                    End If
                Else
                    varParidadTotal = CampoTipoCambio.paridad
                End If
                txtDesdeProveedor.Focus()
            End If
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtFechaEmision.Text = "  /  /    "
            Me.txtFechaEmision.SelectionStart = 0
        End If
    End Sub

    Private Sub btnCancela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancela.Click
        Me.Close()
        MenuPrincipal.Show()
    End Sub

    Private Sub btnConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsulta.Click

        _DeshabilitarConsulta()

        _ListarProveedores()

    End Sub

    Private Sub _HabilitarConsulta()
        txtAyuda.Text = ""
        GrupoConsultas.Visible = True
        txtAyuda.Focus()
    End Sub

    Private Sub _DeshabilitarConsulta()
        txtAyuda.Text = ""
        GrupoConsultas.Visible = False
    End Sub

    Private Sub _ListarProveedores()
        Dim _Proveedores As List(Of Proveedor) = DAOProveedor.buscarProveedoresActivoPorNombre()
        Dim item = ""

        If _Proveedores.Count > 0 Then

            lstAyuda.Items.Clear()
            _Claves.Clear()

            For Each _prv As Proveedor In _Proveedores

                item = ceros(_prv.id, 11) & "    " & _prv.razonSocial

                lstAyuda.Items.Add(item)

                _Claves.Add({item, _prv.id})

            Next

            _HabilitarConsulta()
        Else
            _DeshabilitarConsulta()
        End If

    End Sub

    Private Sub lstAyuda_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstAyuda.Click

        If Trim(lstAyuda.SelectedItem) <> "" Then
            _AgregarProveedorAListadoGeneral(lstAyuda.SelectedItem)
        End If

    End Sub

    Private Sub _AgregarProveedorAListadoGeneral(ByVal item As String)

        If Trim(item) = "" Then
            Exit Sub
        End If

        Dim proveedor As Proveedor = DAOProveedor.buscarProveedorPorCodigo(_Claves.FindLast(Function(p) p(0) = item)(1))
        If Not IsNothing(proveedor) Then
            If Not _ProveedorYaAgregado(proveedor.id) Then
                _CargarProveedor(proveedor)
            End If

            lstFiltrada.Visible = False
            lstAyuda.Visible = True
            txtAyuda.Text = ""
            txtAyuda.Focus()

        End If

    End Sub

    Private Sub _LimpiarImpCtaCtePrvNet()
        Dim cn = New SqlConnection()
        Dim cm = New SqlCommand("DELETE FROM impCtaCtePrvNet")

        SQLConnector.conexionSql(cn, cm)

        Try
            cm.ExecuteNonQuery()

        Catch ex As Exception
            Throw New Exception("Hubo un problema al querer limpiar la tabla ImpCtaCtePrvNet en la Base de Datos.")
        Finally

            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

    End Sub

    Private Function _BuscarTipoCambio(ByVal fecha As String) As Double
        Dim Paridad = 0.0
        Dim cn = New SqlConnection()
        Dim cm = New SqlCommand("SELECT CambioDivisa FROM Cambios WHERE Fecha = '" & fecha & "'")
        Dim dr As SqlDataReader

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then
                dr.Read()

                Paridad = IIf(IsDBNull(dr.Item("CambioDivisa")), 0, dr.Item("CambioDivisa"))

                If Val(Paridad) = 0 Then
                    Throw New SinParidadException("Paridad Inexistente")
                End If

            Else
                Throw New SinParidadException("Paridad Inexistente")
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

        Return Paridad
    End Function

    Private Function _BuscarCtaCtePrvSelectivo(ByVal proveedor As String) As DataTable
        Dim tabla As New DataTable
        Dim cn = New SqlConnection()
        Dim cm = "SELECT ccp.Tipo, ccp.Letra, ccp.Punto, ccp.Numero, ccp.Total, ccp.Saldo, ccp.Fecha, ccp.Vencimiento, ccp.Vencimiento1, Impre = CASE ISNULL(i.Rechazado, 0) WHEN 1 THEN 'CHR' ELSE ccp.Impre END, ccp.NroInterno, ccp.Clave, ccp.Pago, ccp.Paridad FROM CtaCtePrv ccp LEFT OUTER JOIN IvaComp i ON i.NroInterno = ccp.NroInterno WHERE ccp.Proveedor = '" & proveedor.Trim() & "' AND ccp.Saldo <> 0 ORDER BY ccp.Proveedor, ccp.OrdFecha, ccp.NroInterno"
        Dim dr As New SqlDataAdapter(cm, cn)

        Try
            cn.ConnectionString = Proceso._ConectarA
            cn.Open()

            dr.Fill(tabla)

        Catch ex As Exception
            Throw New Exception("Hubo un problema al querer consultar la Cuenta Corriente del Proveedor en la Base de Datos.")
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

        If tabla.Rows.Count > 0 Then
            Return Proceso._NormalizarFilas(tabla)
        Else
            Return Nothing
        End If
    End Function

    Private Function _BuscarProveedor(ByVal codProveedor As String) As DataRow
        Dim proveedor As New DataTable
        Dim cn = New SqlConnection()
        Dim cm = "SELECT CodIb, CodIbCaba, Iva, Tipo, PorceIb, PorceIbCaba FROM Proveedor WHERE Proveedor = '" & codProveedor & "'"
        Dim dr As New SqlDataAdapter(cm, cn)

        Try
            cn.ConnectionString = Proceso._ConectarA
            cn.Open()

            dr.Fill(proveedor)

        Catch ex As Exception
            Throw New Exception("Hubo un problema al querer consultar los datos del Proveedor en la Base de Datos.")
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

        If proveedor.Rows.Count > 0 Then
            Return Proceso._NormalizarFila(proveedor.Rows(0))
        Else
            Return Nothing
        End If

    End Function

    Private Function _BuscarCompra(ByVal NroInterno As Integer) As DataRow
        Dim compra As New DataTable
        Dim cn = New SqlConnection()
        Dim cm = "SELECT Letra, Neto, Iva21, Iva5, Iva27, Iva105, Ib, Exento, ISNULL(MarcaDifCambio, 0) As MarcaDifCambio, ISNULL(Rechazado, 0) As Rechazado FROM IvaComp WHERE NroInterno = '" & NroInterno & "'"
        Dim dr As New SqlDataAdapter(cm, cn)

        Try
            cn.ConnectionString = Proceso._ConectarA
            cn.Open()

            dr.Fill(compra)

        Catch ex As Exception
            Throw New Exception("Hubo un problema al querer consultar la Factura en la Base de Datos.")
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

        If compra.Rows.Count > 0 Then
            Return Proceso._NormalizarFila(compra.Rows(0))
        Else
            Return Nothing
        End If
    End Function

    Private Function _BuscarAcumulado(ByVal proveedor As String, ByVal fecha As String) As DataRow
        Dim acumulado As New DataTable
        Dim cn = New SqlConnection()
        Dim cm = "SELECT Neto, Retenido, Anticipo, Bruto, Iva FROM Retencion WHERE Proveedor = '" & proveedor & "' AND Fecha = '" & fecha & "'"
        Dim dr As New SqlDataAdapter(cm, cn)

        Try
            cn.ConnectionString = Proceso._ConectarA
            cn.Open()

            dr.Fill(acumulado)

        Catch ex As Exception
            Throw New Exception("Hubo un problema al querer consultar los valores acumulados correspondientes al Proveedor: " & proveedor & " en la Base de Datos.")
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

        If acumulado.Rows.Count > 0 Then
            Return Proceso._NormalizarFila(acumulado.Rows(0))
        Else
            Return Nothing
        End If
    End Function

    Enum Reporte
        Imprimir
        Pantalla
    End Enum

    Private Sub _Imprimir(ByVal TipoImpresion As Reporte)

        Dim txtUno, txtDos As String
        Dim txtFormula As String
        Dim x As Char = Chr(34)

        Dim WOrden = 0
        Dim txtEmpresa As String

        Dim varOrdFecha As String
        Dim varCiclo As Integer
        Dim varPorce As Double
        Dim varAcumuladoNetoDifCambioIBI, varDifCambioIB, varDifCambioIBII, varAcumuladoPesosOrigDifCambio As Double
        Dim varAcumulado, varAcuNeto, varAcuRetenido, varAcuIva, varAcuAnticipo, varAcuBruto, varAcumulaUs As Double
        Dim varProveedor, varLetra As String
        Dim varNeto, varIva, varIva5, varIva27, varIva105, varIb, varExento, varTotalTrabajo As Double
        Dim varRetIb, varRetIva, varRetGan As Double
        REM Dim varRetIb, varRetIva, varRetGan, varAcumulaIb, varRete As Double
        Dim varPorceIb, varPorceIbCaba As Double
        Dim varTipoIbCaba, varTipoIva, varTipoPrv, varTipoIb As Integer
        Dim varPago, varEmpresa As Integer
        Dim varAcumulaNeto, varAcumulaNetoII, varAcumulaIva, varPesosOrig, varDifCambio, AcumPesosOrig, AcumDifCambio As Double
        Dim varRetIbI, varRetIbII As Double

        Try

            _LimpiarImpCtaCtePrvNet()

        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try

        txtEmpresa = "Surfactan S.A."
        varEmpresa = 1

        If Proceso._EsPellital() Then
            txtEmpresa = "Pellital S.A"
            varEmpresa = 8
        End If

        varOrdFecha = ordenaFecha(txtFechaEmision.Text)

        Try

            varParidadTotal = _BuscarTipoCambio(txtFechaEmision.Text)

        Catch ex As SinParidadException
            MsgBox(ex.Message, MsgBoxStyle.Information)
            Exit Sub
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            txtFechaEmision.Focus()
            Exit Sub
        End Try

        GRilla.CommitEdit(DataGridViewDataErrorContexts.Commit)

        For varCiclo = 0 To GRilla.Rows.Count - 1

            varAcumulaUs = 0
            varProveedor = GRilla.Item(0, varCiclo).Value

            If Trim(varProveedor) <> "" Then

                Dim DifCambioIbI As Double = 0.0
                Dim DifCambioIbII As Double = 0.0
                Dim AcumPesosOrigII As Double = 0.0
                Dim WVieneDePesos As Boolean = False

                varAcumulado = 0
                varAcumulaIva = 0
                varAcumulaNeto = 0

                varPesosOrig = 0
                varDifCambio = 0

                AcumDifCambio = 0
                AcumPesosOrig = 0

                varTotal = 0
                varSaldo = 0
                varTotalUs = 0
                varSaldoUs = 0
                varSaldoOriginal = 0
                varDife = 0

                Dim tabla As New DataTable

                Try

                    tabla = _BuscarCtaCtePrvSelectivo(varProveedor)

                Catch ex As Exception
                    MsgBox(ex.Message)
                    Exit Sub
                End Try

                If Not IsNothing(tabla) Then
                    For Each row As DataRow In tabla.Rows

                        WOrden += 1

                        Dim CCPrv As New CtaCteProveedoresDeudaDesdeHastaII(row("Tipo"), row("Letra"), row("Punto"), row("Numero"), row("Total"), row("Saldo"), row("Fecha"), row("Vencimiento"), row("Vencimiento1"), row("Impre"), row("NroInterno"), row("Clave"), varProveedor, row("Pago"), row("Paridad"))

                        varPago = CCPrv.pago
                        varParidad = CCPrv.paridad

                        If varPago <> 2 Then
                            varTotal = CCPrv.total
                            varSaldo = CCPrv.saldo
                            varTotalUs = 0
                            varSaldoUs = 0
                            varSaldoOriginal = 0
                            varDife = 0
                        Else
                            If Val(varParidad) = 0 Then
                                MsgBox("La factura con Nro Interno " & CCPrv.nroInterno & " del Proveedor " & CCPrv.Proveedor & ", no posee Paridad cargada.", MsgBoxStyle.Information)

                                varTotal = 0
                                varSaldo = 0
                                varTotalUs = 0
                                varSaldoUs = 0
                                varSaldoOriginal = 0
                                varDife = 0
                            Else
                                varTotal = (CCPrv.total / varParidad) * varParidadTotal
                                varSaldo = (CCPrv.saldo / varParidad) * varParidadTotal
                                varTotalUs = (CCPrv.total / varParidad)
                                varSaldoUs = (CCPrv.saldo / varParidad)
                                varSaldoOriginal = CCPrv.saldo
                                varDife = varSaldo - CCPrv.saldo
                            End If


                            varAcumulaUs += varTotalUs

                        End If

                        redondeo(varTotal)
                        redondeo(varSaldo)

                        varAcumulado = varAcumulado + varSaldo

                        If varTotal = varSaldo Then
                            varPorce = 1
                        Else
                            varPorce = varSaldo / varTotal
                        End If

                        varNeto = 0
                        varIva = 0
                        varIva5 = 0
                        varIva27 = 0
                        varIva105 = 0
                        varIb = 0
                        varExento = 0
                        varTotalTrabajo = 0
                        varLetra = ""

                        Dim CampoProveedor = Nothing

                        Try
                            CampoProveedor = _BuscarProveedor(varProveedor)
                        Catch ex As Exception
                            MsgBox(ex.Message, MsgBoxStyle.Critical)
                            Exit Sub
                        End Try

                        If Not IsNothing(CampoProveedor) Then

                            varTipoIb = CampoProveedor("CodIb")
                            varTipoIbCaba = CampoProveedor("CodIbCaba")
                            varTipoIva = CampoProveedor("Iva")
                            varTipoPrv = CampoProveedor("Tipo") + 1
                            varPorceIb = CampoProveedor("PorceIb")
                            varPorceIbCaba = CampoProveedor("PorceIbCaba")

                        End If

                        Dim compra As DataRow = Nothing
                        Try
                            compra = _BuscarCompra(CCPrv.nroInterno)
                        Catch ex As Exception
                            MsgBox(ex.Message, MsgBoxStyle.Critical)
                            Exit Sub
                        End Try

                        If Not IsNothing(compra) Then

                            varLetra = compra("Letra")
                            varNeto = compra("Neto")
                            varIva = compra("iva21")
                            varIva5 = compra("Iva5")
                            varIva27 = compra("Iva27")
                            varIva105 = compra("iva105")
                            varIb = compra("Ib")
                            varExento = compra("Exento")
                            varTotalTrabajo = varNeto + varIva + varIva5 + varIva27 + varIva105 + varIb + varExento

                        End If

                        varRetIbI = 0
                        varRetIbII = 0
                        varRetIb = 0
                        varRetIva = 0
                        varRetGan = 0

                        '
                        'calcula el neto para el calculo de las retenciones
                        '
                        If varTotalTrabajo <> 0 Then
                            varAcumulaNetoII = varNeto * varPorce
                        Else
                            If varTipoIva = 2 Then
                                varAcumulaNetoII = (varSaldo / 1.21)
                            Else
                                varAcumulaNetoII = varSaldo
                            End If
                        End If

                        If varPago = 2 Then
                            varAcumulaNetoII = varAcumulaNetoII + (varDife / 1.21)
                        End If
                        varAcumulaNeto = varAcumulaNeto + varAcumulaNetoII

                        '
                        'calculo de rtencion de Ingresos brutos Pcia
                        '
                        If compra Is Nothing OrElse compra.Item("Rechazado") <> 1 Then
                            varRetIbI = CaculoRetencionIngresosBrutos(varTipoIb, varPorceIb, varAcumulaNeto)
                        End If

                        '
                        'calculo de retencion de Ingresos brutos CABA. SOLO PARA SURFACTAN.
                        '
                        If varEmpresa = 1 And (compra Is Nothing OrElse compra.Item("Rechazado") <> 1) Then

                            If Val(varPorceIbCaba) <> 0 And Val(varTipoIbCaba) <> 2 Then
                                varRetIbII = CaculoRetencionIngresosBrutosCaba(varTipoIbCaba, varPorceIbCaba, varAcumulaNeto)
                            End If

                        End If


                        '
                        'calculo de rtencion de Ganancias
                        '

                        varAcuNeto = varAcumulaNeto
                        varAcuRetenido =
                        varAcuAnticipo = 0
                        varAcuBruto = 0
                        varAcuIva = 0

                        varOrdFecha = Mid(ordenaFecha(txtFechaEmision.Text), 3, 4)

                        Dim CampoAcumulado = Nothing

                        Try
                            CampoAcumulado = _BuscarAcumulado(varProveedor, varOrdFecha)

                        Catch ex As Exception
                            MsgBox(ex.Message, MsgBoxStyle.Critical)
                            Exit Sub
                        End Try

                        Dim varAcuNetoOrig = 0

                        If Not IsNothing(CampoAcumulado) Then

                            varAcuNeto = CampoAcumulado("Neto")
                            varAcuNetoOrig = CampoAcumulado("Neto")
                            varAcuRetenido = CampoAcumulado("Retenido")
                            varAcuAnticipo = CampoAcumulado("Anticipo")
                            varAcuBruto = CampoAcumulado("Bruto")
                            varAcuIva = CampoAcumulado("Iva")

                        Else

                            varAcuNeto = 0
                            varAcuRetenido = 0
                            varAcuAnticipo = 0
                            varAcuBruto = 0
                            varAcuIva = 0

                        End If

                        If compra Is Nothing OrElse compra.Item("Rechazado") <> 1 Then
                            varRetGan = CaculoRetencionGanancia(varTipoPrv, varAcumulaNeto, varAcuNeto, varAcuRetenido, varAcuAnticipo, varAcuBruto, varAcuIva)
                        End If

                        '
                        'calculo de rtencion de IVA
                        '
                        If varLetra = "M" And (compra Is Nothing OrElse compra.Item("Rechazado") <> 1) Then
                            If varNeto >= 1000 Then
                                varAcumulaIva = varAcumulaIva + varIva
                            End If
                            varRetIva = varAcumulaIva
                        End If

                        If compra IsNot Nothing AndAlso compra.Item("MarcaDifCambio") = 1 Then
                            varParidad = 0
                            varTotalUs = 0
                            varSaldoUs = 0
                            CCPrv.Impre = "NDC"
                            varPago = 1
                        End If

                        varRetIb = varRetIbI + varRetIbII + varRetIva

                        If varParidad <> 0 Then
                            varPesosOrig = varParidad * varSaldoUs
                        Else
                            varPesosOrig = varSaldo
                        End If

                        varAcuNeto = (varAcumulado - varRetIb - varRetGan)

                        If varParidad <> 0 Then

                            If compra.Item("MarcaDifCambio") = 0 Then

                                'varDifCambio = varSaldo - varPesosOrig

                                If varTotalTrabajo <> 0 Then
                                    varAcumulaNetoII = varNeto * varPorce
                                Else
                                    If varTipoIva = 2 Then
                                        varAcumulaNetoII = (varSaldo / 1.21)
                                    Else
                                        varAcumulaNetoII = varSaldo
                                    End If
                                End If

                                If compra Is Nothing OrElse compra.Item("Rechazado") <> 1 Then
                                    DifCambioIbII = CaculoRetencionIngresosBrutos(varTipoIb, varPorceIb, varAcumulaNetoII)
                                End If

                                If varEmpresa = 1 And (compra Is Nothing OrElse compra.Item("Rechazado") <> 1) Then

                                    If Val(varPorceIbCaba) <> 0 And Val(varTipoIbCaba) <> 2 Then
                                        Dim DifCambioIbIII = CaculoRetencionIngresosBrutosCaba(varTipoIbCaba, varPorceIbCaba, varAcumulaNetoII)
                                        DifCambioIbII += DifCambioIbIII

                                    End If

                                End If

                                If compra Is Nothing OrElse compra.Item("Rechazado") <> 1 Then
                                    Dim varAcumulaNetoIII = varAcumulaNetoII
                                    If varPago = 2 Then
                                        varAcumulaNetoIII = varAcumulaNetoII + (varDife / 1.21)
                                    End If

                                    Dim DifCambioIbIV = CaculoRetencionGanancia(varTipoPrv, varAcumulaNetoIII, (varAcumuladoPesosOrigDifCambio - AcumDifCambio - varAcumuladoNetoDifCambioIBI - DifCambioIbII), varAcuRetenido, varAcuAnticipo, varAcuBruto, varAcuIva)
                                    DifCambioIbII += DifCambioIbIV
                                End If

                                DifCambioIbI += DifCambioIbII

                                varAcumuladoNetoDifCambioIBI = DifCambioIbI

                                varAcumuladoPesosOrigDifCambio += varPesosOrig

                                If WVieneDePesos Then
                                    varDifCambio = varAcuNeto - (AcumPesosOrig + AcumDifCambio + varPesosOrig - DifCambioIbII)
                                    AcumPesosOrig += (varPesosOrig - DifCambioIbII)
                                Else
                                    AcumPesosOrig = varAcumuladoPesosOrigDifCambio - varAcumuladoNetoDifCambioIBI

                                    varDifCambio = varAcuNeto - AcumPesosOrig - varDifCambio
                                End If

                                AcumDifCambio += varDifCambio
                                WVieneDePesos = False
                            Else
                                AcumPesosOrig = varAcumuladoPesosOrigDifCambio - varAcumuladoNetoDifCambioIBI
                            End If

                        Else
                            'varDifCambio = 0
                            AcumPesosOrig = varAcuNeto - AcumDifCambio
                            varAcumuladoPesosOrigDifCambio += varPesosOrig
                            WVieneDePesos = True
                        End If

                        'varPesosOrig = varParidad * varSaldoUs
                        'varDifCambio = 0

                        'varAcuNeto = (varAcumulado - varRetIb - varRetGan)

                        'If varParidad <> 0 Then

                        '    If compra.Item("MarcaDifCambio") = 0 Then

                        '        varDifCambio = varSaldo - varPesosOrig

                        '        AcumDifCambio += varDifCambio

                        '    End If

                        '    'AcumPesosOrig = IIf(varDifCambio >= 0, (varAcuNeto - AcumDifCambio), varAcuNeto)
                        '    AcumPesosOrig = varAcuNeto - AcumDifCambio

                        'Else
                        '    AcumPesosOrig = varAcuNeto
                        'End If

                        Try
                            SQLConnector.executeProcedure("alta_impCtaCtePrvNetII", CCPrv.Clave, CCPrv.Proveedor, CCPrv.Tipo, CCPrv.letra, CCPrv.punto, CCPrv.numero, varTotal, varSaldo, CCPrv.fecha, CCPrv.vencimiento, CCPrv.VencimientoII, CCPrv.Impre, CCPrv.nroInterno, txtEmpresa, varAcumulado, WOrden, txtFechaEmision.Text, "", "", "", varParidadTotal, varSaldoOriginal, varDife, 0, 0, "", varRetIb, varRetGan, varAcuNeto, varParidad, varTotalUs, varSaldoUs, varAcumulaUs, varPago, varPesosOrig, varDifCambio, AcumDifCambio, AcumPesosOrig)
                        Catch ex As Exception
                            MsgBox(ex.Message, MsgBoxStyle.Critical)
                            Exit Sub
                        End Try


                    Next
                End If

            End If

        Next

        txtUno = "{impCtaCtePrvNet.Proveedor} in " + x + "" + x + " to " + x + "ZZZZZZZZZZZ" + x
        txtDos = " and {impCtaCtePrvNet.Saldo} <> 0.00"
        txtFormula = txtUno + txtDos

        With VistaPrevia
            .Reporte = New ListadoCtaCtePrvSelectivoPrueba

            .Formula = txtFormula

            Select Case TipoImpresion
                Case Reporte.Imprimir
                    .Imprimir()
                Case Reporte.Pantalla
                    .Mostrar()
            End Select
        End With

    End Sub

    Private Sub txtDesdeProveedor_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDesdeProveedor.KeyDown
        If e.KeyData = Keys.Enter Then

            If Trim(txtDesdeProveedor.Text) = "" Then
                btnConsulta.PerformClick()
                Exit Sub
            End If

            If Not _ProveedorYaAgregado(txtDesdeProveedor.Text) Then
                ' DADA que no rompa cuando el codigo no existe y usar la funcion "ceros" para completar??
                _CargarProveedor(DAOProveedor.buscarProveedorPorCodigo(txtDesdeProveedor.Text))
            Else
                MsgBox("El proveedor indicado ya se encuentra agregado a la lista semanal.", MsgBoxStyle.Information)
                txtDesdeProveedor.Focus()
                Exit Sub
            End If

        ElseIf e.KeyData = Keys.Escape Then
            txtDesdeProveedor.Text = ""
            txtDesdeProveedor.Focus()
        End If
    End Sub

    Private Sub btnLimpiarTodo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpiarTodo.Click
        _DeshabilitarConsulta()

        Dim WFecha = txtFechaEmision.Text

        For Each _C As TextBox In Me.Panel2.Controls.OfType(Of TextBox)()
            _C.Text = ""
        Next

        For Each _C As MaskedTextBox In Me.Panel2.Controls.OfType(Of MaskedTextBox)()
            _C.Clear()
        Next

        GRilla.Rows.Clear()

        txtFechaEmision.Text = WFecha
        txtFechaEmision.Focus()

        varRenglon = 0

    End Sub

    Private Sub lstAyuda_Filtrada_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstAyuda_Filtrada.Click
        If Trim(lstFiltrada.SelectedItem) <> "" Then
            _AgregarProveedorAListadoGeneral(lstFiltrada.SelectedItem)
        End If
    End Sub

    Private Sub txtDesdeProveedor_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtDesdeProveedor.MouseDoubleClick
        btnConsulta.PerformClick()
    End Sub

    ' Rutinas de Filtrado Dinámico.
    Private Sub _FiltrarDinamicamente()
        Dim origen As ListBox = lstAyuda
        Dim final As ListBox = lstFiltrada
        Dim cadena As String = Trim(txtAyuda.Text)

        final.Items.Clear()

        If UCase(Trim(cadena)) <> "" Then

            For Each item In origen.Items

                If UCase(item.ToString()).Contains(UCase(Trim(cadena))) Then

                    final.Items.Add(item)

                End If

            Next

            final.Visible = True

        Else

            final.Visible = False

        End If
    End Sub

    Private Sub lstFiltrada_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lstFiltrada.MouseClick
        Dim origen As ListBox = lstAyuda
        Dim filtrado As ListBox = lstFiltrada
        Dim texto As TextBox = txtAyuda

        If Trim(filtrado.SelectedItem) = "" Then : Exit Sub : End If

        ' Buscamos el texto exacto del item seleccionado y seleccionamos el mismo item segun su indice en la lista de origen.
        origen.SelectedItem = filtrado.SelectedItem

        ' Llamamos al evento que tenga asosiado el control de origen.
        lstAyuda_Click(Nothing, Nothing)


        ' Sacamos de vista los resultados filtrados.
        filtrado.Visible = False
        texto.Text = ""
    End Sub

    Private Sub txtAyuda_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAyuda.TextChanged
        _FiltrarDinamicamente()
    End Sub

    Private Sub SoloNumero(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDesdeProveedor.KeyPress
        If Not Char.IsNumber(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnPantalla_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPantalla.Click
        _Imprimir(Reporte.Pantalla)
    End Sub

    Private Sub btnImprimir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImprimir.Click

        _Imprimir(Reporte.Imprimir)
        MsgBox("El reporte se ha impreso correctamente.", MsgBoxStyle.Information)

    End Sub

    Private Sub CustomButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomButton1.Click

        If txtFechaEmision.Text.Replace(" ", "").Length = 10 Then
            _CargarProveedoresPreCargados()
        Else
            MsgBox("Debe indicar una fecha de Pago.", MsgBoxStyle.Exclamation)
            txtFechaEmision.Focus()
            Exit Sub
        End If

        txtDesdeProveedor.Focus()

    End Sub

    Private Sub GRilla_CellMouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles GRilla.CellMouseDoubleClick
        If e.RowIndex < 0 Then : Exit Sub : End If
        Dim fila = GRilla.Rows(e.RowIndex)
        If Not IsNothing(fila) Then

            If Trim(fila.Cells(0).Value) <> "" Then

                If MsgBox("¿Seguro de querer eliminar el Proveedor Seleccionado?", MsgBoxStyle.YesNo, MsgBoxStyle.Information) = DialogResult.Yes Then
                    GRilla.Rows.Remove(fila)
                End If

            End If

        End If
    End Sub

    Private Sub ListadoCuentaCorrienteProveedoresSelectivo_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        txtFechaEmision.Focus()
    End Sub

    Private Function _EsNumero(ByVal keycode As Integer) As Boolean
        Return (keycode >= 48 And keycode <= 57) Or (keycode >= 96 And keycode <= 105) Or (keycode = 109)
    End Function

    Private Function _EsControl(ByVal keycode) As Boolean
        Dim valido = False

        Select Case keycode
            Case Keys.Enter, Keys.Escape, Keys.Right, Keys.Left, Keys.Back
                valido = True
            Case Else
                valido = False
        End Select

        Return valido
    End Function

    Private Function _EsNumeroOControl(ByVal keycode) As Boolean
        Dim valido = False

        If _EsNumero(CInt(keycode)) Or _EsControl(keycode) Then
            valido = True
        Else
            valido = False
        End If

        Return valido
    End Function

    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean

        If GRilla.Focused Or GRilla.IsCurrentCellInEditMode Then ' Detectamos los ENTER tanto si solo estan en foco o si estan en edición una celda.
            GRilla.CommitEdit(DataGridViewDataErrorContexts.Commit) ' Guardamos todos los datos que no hayan sido confirmados.

            Dim iCol = GRilla.CurrentCell.ColumnIndex
            Dim iRow = GRilla.CurrentCell.RowIndex

            ' Limitamos los caracteres permitidos para cada una de las columnas.
            Select Case iCol
                Case 0
                    If Not _EsNumeroOControl(keyData) Then
                        Return True
                    End If
            End Select

            If msg.WParam.ToInt32() = Keys.Enter Then

                Dim valor = GRilla.Rows(iRow).Cells(iCol).Value

                If Not IsNothing(valor) Then

                    If iCol = 0 And iRow > -1 Then

                        Dim proveedor As Proveedor = DAOProveedor.buscarProveedorPorCodigo(valor)

                        If Not IsNothing(proveedor) Then
                            If Not _ProveedorYaAgregado(proveedor.id, iRow) Then
                                GRilla.Rows(iRow).Cells(1).Value = Trim(proveedor.razonSocial)

                                If GRilla.Rows.Count < iRow + 1 Then
                                    GRilla.CurrentCell = GRilla.Rows(GRilla.Rows.Add).Cells(iCol)
                                Else
                                    GRilla.CurrentCell = GRilla.Rows(iRow + 1).Cells(iCol)
                                End If
                            Else
                                MsgBox("Proveedor ya cargado con anterioridad.", MsgBoxStyle.Information)
                                GRilla.CurrentCell = GRilla.Rows(iRow).Cells(iCol)
                            End If

                        End If

                    End If

                    Return True
                Else

                    With GRilla
                        Select Case iCol
                            Case 0, 1
                                .CurrentCell = .Rows(iRow).Cells(iCol)
                        End Select
                    End With

                    Return True

                End If
            ElseIf msg.WParam.ToInt32() = Keys.Escape Then

                With GRilla
                    .Rows(iRow).Cells(iCol).Value = ""

                    ' Solo para que pierda el foco y se refresque el contenido sino sigue quedando ahi.
                    If iCol = 1 Then
                        .CurrentCell = .Rows(iRow).Cells(iCol - 1)
                    Else
                        .CurrentCell = .Rows(iRow).Cells(iCol + 1)
                    End If

                    .CurrentCell = .Rows(iRow).Cells(iCol)
                End With


            End If
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Private Sub txtFechaEmision_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFechaEmision.KeyDown

        If e.KeyData = Keys.Enter Then
            If Trim(txtFechaEmision.Text).Replace("/", "") = "" Then : Exit Sub : End If

            If _ExistenProveedoresCargados() Then

                If MsgBox("Existen Proveedores cargados para esta fecha ¿Desea traerlos?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    CustomButton1.PerformClick()
                    Exit Sub
                End If

            End If

            txtDesdeProveedor.Focus()

        ElseIf e.KeyData = Keys.Escape Then
            txtFechaEmision.Text = ""
        End If

    End Sub

    Private Function _ExistenProveedoresCargados() As Boolean
        Dim WExistenProveedoresCargados As Boolean

        Dim cn = New SqlConnection()
        Dim cm = New SqlCommand("SELECT TOP 1 Proveedor FROM ProveedorSelectivo WHERE Fecha = '" & txtFechaEmision.Text & "'")
        Dim dr As SqlDataReader

        Try

            cn.ConnectionString = Proceso._ConectarA
            cn.Open()
            cm.Connection = cn

            dr = cm.ExecuteReader()

            WExistenProveedoresCargados = dr.HasRows

        Catch ex As Exception
            Throw New Exception("Hubo un problema al querer consultar la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try
        Return WExistenProveedoresCargados
    End Function

    Private Sub btnCerrarConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrarConsulta.Click
        _DeshabilitarConsulta()
    End Sub
End Class