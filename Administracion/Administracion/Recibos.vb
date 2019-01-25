Imports ClasesCompartidas
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports Microsoft.Office.Interop

Public Class Recibos
    ' Variables y Constantes para fecha en grilla.
    Private WRow, Wcol As Integer
    Private Const YMARGEN = 233
    Private Const XMARGEN = 471

    ' Variables para alta de nuevo recibo
    Private _ComprobanteRetIva As String = ""
    Private _ComprobanteRetGanancias As String = ""
    Private _ComprobanteRetSuss As String = ""
    Private _RetIB1, _CompIB1, _RetIB2, _CompIB2, _RetIB3, _CompIB3, _RetIB4, _CompIB4, _RetIB5, _CompIB5, _
            _RetIB6, _CompIB6, _RetIB7, _CompIB7, _RetIB8, _CompIB8, _RetIB9, _CompIB9, _RetIB10, _CompIB10, _RetIB11, _CompIB11, _RetIB12, _CompIB12, _RetIB13, _CompIB13, _RetIB14, _CompIB14 As String

    ' Variables para impresion de Recibo.
    ' Variables para impresion de Recibo.d
    Dim WRazon, WDireccion, WLocalidad, WProvincia, WPostal, _
            WRecibo, WFecha, WCliente, WEmail As String

    ' Variables Auxiliares
    Private _Provincia As Integer = 0
    Const _SEPARADOR As String = "    "
    Private _TipoConsulta As Integer = 0
    Private _ValoresMax As New List(Of Object)
    Private _CuentasContables As New List(Of Object)
    Private _ClavesCheques As New List(Of Object)
    Dim XRenglon, XRecibo, XCliente, XFecha, XFechaOrd, XTipoRec, XRetganancias, XRetIva, XRetotra, _
            XRetencion, XTiporeg, XTipo, XNumero, ClaveCtaCte, XTipo1, XLetra1, XPunto1, XNumero1, XImporte1, XTipo2, _
            XNumero2, XFecha2, XFechaOrd2, XBanco2, XImporte2, XEstado2, XObservaciones, XEmpresa, XClave, _
            XImporte, XImporteBaja, XCuenta, XMarca, XFechaDepo, XFechaDepoOrd, XImpolist, XImpo1list, XDestino, XEstado, _
            XVencimiento, XVencimiento1, XTotal, XTotalUs, XSaldo, XSaldoUs, XOrdFecha, XOrdVencimiento, _
            XOrdVencimiento1, XImpre, XNeto, XIva1, XIva2, XImpoIb, XSeguro, XFlete, XPedido, XRemito, XOrden, _
            XParidad, XProvincia, XVendedor, XRubro, XComprobante, XAceptada, XCosto, XImporte3, XImporte4, _
            XImporte5, XImporte6, XImporte7, XDate, XParam, XSql, XClaveCheque, XBancoCheque, XSucursalCheque, _
            XChequeCheque, XCuentaCheque, XCuit, XClaveCuit, XNet As String

    Dim iRow, renglon As Integer
    Dim _Cheque As Object = Nothing
    Dim _CuitExistente As Boolean = False
    Dim cn As SqlConnection = New SqlConnection()
    Dim cm As SqlCommand = New SqlCommand()
    Dim dr As SqlDataReader

    Dim WOffset As Integer = 1

    ' No tengo idea de para que son.
    Dim queryController As QueryController
    Dim commonEventsHandler As New CommonEventsHandler

    Private Sub Recibos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Label2.Text = Globals.NombreEmpresa()

        WRow = -1
        Wcol = -1
        WOffset = 1

        'commonEventsHandler.setIndexTab(Me)
        lstSeleccion.SelectedIndex = 0
        lstFiltrada.Visible = False

        renglon = 0

        '_AlinearCeldas()

        btnLimpiar.PerformClick()

        _DeterminarParidad()

        txtProvi.Focus()
    End Sub

    Private Sub lstSeleccion_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstSeleccion.Click

        Select Case lstSeleccion.SelectedIndex
            Case 0

                _TipoConsulta = 0
                _CargarClientes()

            Case 1

                _TipoConsulta = 1
                _CargarCtasCtes()

            Case Else
        End Select

    End Sub

    Private Sub _CargarCtasCtes()
        Dim cn = New SqlConnection()
        Dim cm = New SqlCommand("SELECT Impre, Numero, Fecha, Saldo FROM CtaCte WHERE Saldo <> 0 AND Cliente = '" & Trim(txtCliente.Text) & "' order by OrdFecha, Numero")
        Dim dr As SqlDataReader

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()
            lstConsulta.Items.Clear()

            If dr.HasRows Then

                Do While dr.Read()
                    lstConsulta.Items.Add(dr.Item("Impre") & _SEPARADOR & ceros(dr.Item("Numero"), 6) & _SEPARADOR & dr.Item("Fecha") & _SEPARADOR & _Redondear(dr.Item("Saldo")).ToString.PadLeft(10, "_"))
                Loop

                _HabilitarConsultas()
            Else
                MsgBox("No se pudieron hallar Cuentas Corrientes asociadas a este Cliente", MsgBoxStyle.Information)
            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar las Cuentas Corrientes del Cliente especificado.", MsgBoxStyle.Critical)
        Finally

            cn.Close()

        End Try
    End Sub

    Private Sub _CargarClientes()
        Dim cn = New SqlConnection()
        Dim cm = New SqlCommand("SELECT Cliente, Razon FROM Cliente")
        Dim dr As SqlDataReader

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()
            lstConsulta.Items.Clear()

            If dr.HasRows Then

                Do While dr.Read()
                    lstConsulta.Items.Add(dr.Item("Cliente") & _SEPARADOR & dr.Item("Razon"))
                Loop

            End If

            _HabilitarConsultas()

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar los Clientes.", MsgBoxStyle.Critical)
        Finally

            cn.Close()

        End Try
    End Sub

    Private Sub _HabilitarConsultas()
        lstSeleccion.Visible = False
        lstConsulta.Visible = True
        txtConsulta.Visible = True
        txtConsulta.Focus()
    End Sub

    Private Sub btnConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsulta.Click
        _ResetearConsultas()
        lstSeleccion.Visible = True
    End Sub

    Private Sub txtConsulta_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtConsulta.KeyDown

        If e.KeyData = Keys.Escape Then
            txtConsulta.Text = ""
            Exit Sub
        End If

        lstFiltrada.Items.Clear()

        If UCase(Trim(txtConsulta.Text)) <> "" Then

            For Each item In lstConsulta.Items

                If UCase(item.ToString()).Contains(UCase(Trim(txtConsulta.Text))) Then

                    lstFiltrada.Items.Add(item.ToString())

                End If

            Next

            lstConsulta.Visible = False
            lstFiltrada.Visible = True

        Else

            lstConsulta.Visible = True
            lstFiltrada.Visible = False

        End If
    End Sub

    Private Sub lstConsulta_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstConsulta.Click

        If Trim(lstConsulta.SelectedItem) <> "" Then
            _TraerConsulta(lstConsulta.SelectedItem)
        End If

    End Sub

    Private Sub _TraerConsulta(ByVal parametro As String)

        Select Case _TipoConsulta

            Case 0
                _AsignarCliente(parametro)
            Case 1
                _AsignarCtaCte(parametro)
        End Select

    End Sub

    Private Sub _AsignarCtaCte(ByVal stringCtaCte As String)
        Dim cuenta As String() = stringCtaCte.Replace(_SEPARADOR, "$").Split("$")
        Dim cn = New SqlConnection()
        Dim cm = New SqlCommand("SELECT Numero, Tipo, Saldo, TotalUs, Paridad FROM CtaCte WHERE Numero = '" & cuenta(1) & "' and Impre = '" & cuenta(0) & "' and Cliente = '" & Trim(txtCliente.Text) & "'")
        Dim dr As SqlDataReader
        Dim row As Integer

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then

                Do While dr.Read()

                    If Not _CuentaUtilizada(dr.Item("Numero")) Then

                        If gridPagos2.Rows.Count = 1 AndAlso (Trim(gridPagos2.Rows(0).Cells(0).Value) = "" OrElse IsNothing(gridPagos2.Rows(0).Cells(0).Value)) Then
                            row = 0
                        ElseIf gridPagos2.Rows.Count <= 38 Then
                            row = gridPagos2.Rows.Add()
                        Else
                            MsgBox("Se ha alcanzado el Número Máximo de Facturas que pueden cargarse por Recibo.")
                            Exit Sub
                        End If

                        With gridPagos2.Rows(row)
                            .Cells(0).Value = dr.Item("Tipo")
                            .Cells(3).Value = dr.Item("Numero")
                            .Cells(4).Value = _Redondear(dr.Item("Saldo"))
                            .Cells(5).Value = _Redondear(dr.Item("Saldo"))
                            '.Cells(4).Value = _NormalizarNumero(dr.Item("Saldo"))
                            '.Cells(5).Value = _NormalizarNumero(dr.Item("Saldo"))
                        End With
                        'gridPagos.Rows.Add()
                        _SumarDebitos()
                        gridPagos2.CurrentCell = gridPagos2.Rows(row).Cells(0)
                        gridPagos2.Focus()
                    End If
                Loop

            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la informacón de la Cuenta Corriente.", MsgBoxStyle.Critical)
        Finally

            cn.Close()

        End Try

    End Sub

    Private Function _SumarDebitos() As Boolean
        Dim _Paridad, _TipoCompo As String
        Dim _Error = False
        lblTotalDebitos.Text = 0
        lblDolares.Text = 0
        lblDiferencia.Text = 0

        For Each row As DataGridViewRow In gridPagos2.Rows

            _Paridad = 0
            _TipoCompo = 0

            If Val(row.Cells(4).Value) <> 0 Then

                Dim cn = New SqlConnection()
                Dim cm = New SqlCommand("SELECT Paridad, TipoCompo FROM CtaCte WHERE clave = '" & row.Cells(0).Value & ceros(row.Cells(3).Value, 8) & "01'")
                Dim dr As SqlDataReader

                SQLConnector.conexionSql(cn, cm)

                Try

                    dr = cm.ExecuteReader()

                    If dr.HasRows Then

                        With dr

                            .Read()

                            _Paridad = _NormalizarNumero(.Item("Paridad").ToString)
                            _TipoCompo = Val(IIf(IsDBNull(.Item("TipoCompo")), 0, .Item("TipoCompo")))

                        End With

                    End If

                Catch ex As Exception
                    MsgBox("Hubo un problema al querer consultar la informacón de la Cuenta Corriente.", MsgBoxStyle.Critical)
                    _Error = True
                Finally

                    cn.Close()

                End Try

                'If _Provincia = 24 And Val(_TotalUs) <> 0 Then
                If _Provincia = 24 Then

                    lblTotalDebitos.Text += (Val(_NormalizarNumero(row.Cells(4).Value)) * Val(_NormalizarNumero(_Paridad)))

                    lblDolares.Text += Val(_NormalizarNumero(row.Cells(4).Value))

                Else

                    lblTotalDebitos.Text += Val(_NormalizarNumero(row.Cells(4).Value))

                    'If _TipoCompo <> 2 And _Paridad <> 0 Then
                    If _Paridad <> 0 Then

                        lblDolares.Text += (Val(_NormalizarNumero(row.Cells(4).Value)) / Val(_NormalizarNumero(_Paridad)))

                    End If

                End If

            End If
        Next

        lblTotalDebitos.Text = _NormalizarNumero(lblTotalDebitos.Text)
        lblDiferencia.Text = Val(_NormalizarNumero(lblTotalDebitos.Text)) - Val(_NormalizarNumero(lblTotalCreditos.Text))
        lblDolares.Text = _NormalizarNumero(lblDolares.Text)

        _RecalcularDolaresfinales()

        Return _Error

    End Function

    Private Function _Redondear(ByVal numero As String) As String
        Return _NormalizarNumero(Math.Round(Val(_NormalizarNumero(numero)), 2))
    End Function

    Private Sub _RecalcularDolaresfinales()
        If Val(_NormalizarNumero(txtParidad.Text)) > 0 Then
            lblDolares.Text = Val(_NormalizarNumero(lblDolares.Text)) - (Val(_NormalizarNumero(lblTotalDebitos.Text)) / Val(_NormalizarNumero(txtParidad.Text)))
            lblDolares.Text = Val(_NormalizarNumero(lblDolares.Text)) * -1

            'lblDolares.Text = _NormalizarNumero(lblDolares.Text)
        Else
            lblDolares.Text = "0.00"
        End If

    End Sub

    Private Function _CuentaUtilizada(ByRef NumeroCuenta As String)

        For i = 0 To gridPagos2.Rows.Count - 1
            If gridPagos2.Rows(i).Cells(3).Value = NumeroCuenta Then
                Return True
            End If
        Next

        Return False

    End Function

    Private Sub _AsignarCliente(ByVal StringCliente As String)
        Dim cliente As String()

        cliente = StringCliente.Replace(_SEPARADOR, "$").Split("$")

        txtCliente.Text = cliente(0)
        txtNombre.Text = cliente(1)

        txtRetGanancias.Focus()

        _ResetearConsultas()

    End Sub

    Private Sub _ResetearConsultas()
        lstConsulta.Visible = False
        lstSeleccion.Visible = False
        lstFiltrada.Visible = False
        txtConsulta.Text = ""
        txtConsulta.Visible = False
    End Sub

    Private Sub btnLimpiar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpiar.Click
        Cleanner.clean(Me)
        setDefaults()
    End Sub

    Private Sub setDefaults()
        txtFechaAux.Visible = False
        WOffset = 1

        txtFecha.Text = Date.Today.ToString("dd/MM/yyyy")
        gridFormasPago2.Rows.Clear()
        gridPagos2.Rows.Clear()

        gridFormasPago2.Rows.Add()
        gridPagos2.Rows.Add()

        optCtaCte.Checked = True

        For Each control As Control In Panel2.Controls

            If TypeOf (control) Is TextBox Then
                control.Text = ""
            End If

        Next

        _ComprobanteRetIva = ""
        _ComprobanteRetGanancias = ""
        _ComprobanteRetSuss = ""

        _RetIB1 = ""
        _CompIB1 = ""
        _RetIB2 = ""
        _CompIB2 = ""
        _RetIB3 = ""
        _CompIB3 = ""
        _RetIB4 = ""
        _CompIB4 = ""
        _RetIB5 = ""
        _CompIB5 = ""
        _RetIB6 = ""
        _CompIB6 = ""
        _RetIB7 = ""
        _CompIB7 = ""
        _RetIB8 = ""
        _CompIB8 = ""

        _RetIB9 = ""
        _CompIB9 = ""
        _RetIB10 = ""
        _CompIB10 = ""
        _RetIB11 = ""
        _CompIB11 = ""
        _RetIB12 = ""
        _CompIB12 = ""
        _RetIB13 = ""
        _CompIB13 = ""
        _RetIB14 = ""
        _CompIB14 = ""

        WRazon = ""
        WDireccion = ""
        WLocalidad = ""
        WProvincia = ""
        WPostal = ""
        WRecibo = ""
        WFecha = ""
        WCliente = ""
        WEmail = ""

        _ClavesCheques.Clear()
        _CuentasContables.Clear()

        lblDolares.Text = "0.0"
        lblTotalCreditos.Text = "0.0"
        lblTotalDebitos.Text = "0.0"
        lblDiferencia.Text = "0.0"

        _DeterminarParidad()
        txtRecibo.Text = "0"
        txtProvi.Focus()
        _ResetearConsultas()

        _Provincia = 0
    End Sub

    Private Sub eventoSegunTipoEnFormaDePagoPara(ByVal val As Integer, ByVal rowIndex As Integer, ByVal columnIndex As Integer)
        Dim column As Integer = columnIndex
        Select Case val
            Case 1, 4
                column = 4
                With gridFormasPago2.Rows(rowIndex)
                    .Cells(1).Value = ""
                    .Cells(2).Value = ""
                    .Cells(3).Value = ""
                End With
            Case 2, 3
                column = 1
            Case Else
                Exit Sub
        End Select
        gridFormasPago2.CurrentCell.Value = ceros(val.ToString, 2)
        gridFormasPago2.CurrentCell = gridFormasPago2.Rows(rowIndex).Cells(column)
    End Sub

    ' CHECKEAR PARA QUE ESTA
    Private Sub agregarClienteABanco()
        For Each row As DataGridViewRow In gridFormasPago2.Rows
            If row.Cells(3).Value <> "" Then
                If row.Cells(3).Value.ToString.Length > 20 Or row.Cells(3).Value.ToString.Contains("/") Then
                    'row.Cells(3).Value = ""
                Else
                    row.Cells(3).Value = row.Cells(3).Value.ToString & clienteSinCeros()
                End If
            End If
        Next
    End Sub

    ' CHECKEAR PARA QUE ESTA.
    Private Function clienteSinCeros()
        Try
            Dim cliente As String = txtCliente.Text
            Dim numero As Integer = CustomConvert.toIntOrZero(cliente.Substring(1, cliente.Count - 1))
            Return "/" & cliente.First & numero
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Sub btnCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub mostrarRecibo(ByVal recibo As Recibo)
        If IsNothing(recibo) Then
            Dim temp As String = ceros(txtRecibo.Text, 6)
            btnLimpiar.PerformClick()
            txtRecibo.Text = temp
            txtRecibo.Focus()
            Exit Sub
        Else
            btnLimpiar.PerformClick()

            txtRecibo.Text = recibo.codigo
            txtFecha.Text = Proceso._Normalizarfecha(recibo.fecha)
            '_NormalizarFecha(_fecha)
            'txtFecha.Text = _fecha
            _DeterminarParidad()
            If Not IsNothing(recibo.cliente) Then mostrarCliente(recibo.cliente.id)
            txtRetGanancias.Text = _NormalizarNumero(recibo.retGanancias)
            txtRetIB.Text = _NormalizarNumero(recibo.retIB)
            txtRetIva.Text = _NormalizarNumero(recibo.retIVA)
            txtRetSuss.Text = _NormalizarNumero(recibo.retSuss)
            'txtParidad.Text = recibo.paridad
            txtObservaciones.Text = recibo.observaciones
            mostrarTipoRecibo(recibo.tipo)
            mostrarCuenta(recibo.cuenta)
            mostrarPagos(recibo.pagos)
            mostrarFormasPago(recibo.formasPago)
            _TraerDatosRestantes(txtRecibo.Text)

            _SumarCreditos()
            _SumarDebitos()
        End If
    End Sub

    Private Sub _TraerDatosRestantes(ByVal recibo As String)

        cm.CommandText = "SELECT * FROM Recibos WHERE Recibo = '" & recibo & "'"
        SQLConnector.conexionSql(cn, cm)

        _ClavesCheques.Clear()
        _CuentasContables.Clear()

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then
                Do While dr.Read()


                    If Val(dr.Item("Renglon")) = 1 Then

                        txtRetGanancias.Text = IIf(IsDBNull(dr.Item("RetGanancias")), "", dr.Item("RetGanancias"))
                        _ComprobanteRetGanancias = IIf(IsDBNull(dr.Item("ComproGanan")), "", dr.Item("ComproGanan"))
                        txtRetIva.Text = IIf(IsDBNull(dr.Item("RetIva")), "", dr.Item("RetIva"))
                        _ComprobanteRetIva = IIf(IsDBNull(dr.Item("ComproIva")), "", dr.Item("ComproIva"))
                        txtRetSuss.Text = IIf(IsDBNull(dr.Item("RetSuss")), "", dr.Item("RetSuss"))
                        _ComprobanteRetSuss = IIf(IsDBNull(dr.Item("ComproSuss")), "", dr.Item("ComproSuss"))
                        txtRetIB.Text = _NormalizarNumero(IIf(IsDBNull(dr.Item("RetOtra")), "", dr.Item("RetOtra")))
                        'txtParidad.Text = IIf(IsDBNull(dr.Item("Paridad")), "", dr.Item("Paridad"))
                        _RetIB1 = IIf(IsDBNull(dr.Item("RetIb1")), "", dr.Item("RetIb1"))
                        _CompIB1 = IIf(IsDBNull(dr.Item("NroRetIb1")), "", dr.Item("NroRetIb1"))
                        _RetIB2 = IIf(IsDBNull(dr.Item("RetIb2")), "", dr.Item("RetIb2"))
                        _CompIB2 = IIf(IsDBNull(dr.Item("NroRetIb2")), "", dr.Item("NroRetIb2"))
                        _RetIB3 = IIf(IsDBNull(dr.Item("RetIb3")), "", dr.Item("RetIb3"))
                        _CompIB3 = IIf(IsDBNull(dr.Item("NroRetIb3")), "", dr.Item("NroRetIb3"))
                        _RetIB4 = IIf(IsDBNull(dr.Item("RetIb4")), "", dr.Item("RetIb4"))
                        _CompIB4 = IIf(IsDBNull(dr.Item("NroRetIb4")), "", dr.Item("NroRetIb4"))
                        _RetIB5 = IIf(IsDBNull(dr.Item("RetIb5")), "", dr.Item("RetIb5"))
                        _CompIB5 = IIf(IsDBNull(dr.Item("NroRetIb5")), "", dr.Item("NroRetIb5"))
                        _RetIB6 = IIf(IsDBNull(dr.Item("RetIb6")), "", dr.Item("RetIb6"))
                        _CompIB6 = IIf(IsDBNull(dr.Item("NroRetIb6")), "", dr.Item("NroRetIb6"))
                        _RetIB7 = IIf(IsDBNull(dr.Item("RetIb7")), "", dr.Item("RetIb7"))
                        _CompIB7 = IIf(IsDBNull(dr.Item("NroRetIb7")), "", dr.Item("NroRetIb7"))
                        _RetIB8 = IIf(IsDBNull(dr.Item("RetIb8")), "", dr.Item("RetIb8"))
                        _CompIB8 = IIf(IsDBNull(dr.Item("NroRetIb8")), "", dr.Item("NroRetIb8"))
                        _RetIB9 = IIf(IsDBNull(dr.Item("RetIb9")), "", dr.Item("RetIb9"))
                        _CompIB9 = IIf(IsDBNull(dr.Item("NroRetIb9")), "", dr.Item("NroRetIb9"))
                        _RetIB10 = IIf(IsDBNull(dr.Item("RetIb10")), "", dr.Item("RetIb10"))
                        _CompIB10 = IIf(IsDBNull(dr.Item("NroRetIb10")), "", dr.Item("NroRetIb10"))
                        _RetIB11 = IIf(IsDBNull(dr.Item("RetIb11")), "", dr.Item("RetIb11"))
                        _CompIB11 = IIf(IsDBNull(dr.Item("NroRetIb11")), "", dr.Item("NroRetIb11"))
                        _RetIB12 = IIf(IsDBNull(dr.Item("RetIb12")), "", dr.Item("RetIb12"))
                        _CompIB12 = IIf(IsDBNull(dr.Item("NroRetIb12")), "", dr.Item("NroRetIb12"))
                        _RetIB13 = IIf(IsDBNull(dr.Item("RetIb13")), "", dr.Item("RetIb13"))
                        _CompIB13 = IIf(IsDBNull(dr.Item("NroRetIb13")), "", dr.Item("NroRetIb13"))
                        _RetIB14 = IIf(IsDBNull(dr.Item("RetIb14")), "", dr.Item("RetIb14"))
                        _CompIB14 = IIf(IsDBNull(dr.Item("NroRetIb14")), "", dr.Item("NroRetIb14"))
                        txtProvi.Text = IIf(IsDBNull(dr.Item("Provisorio")), "", dr.Item("Provisorio"))

                    End If

                    If Val(dr.Item("Tipo2")) = 2 Then
                        With dr
                            _ClavesCheques.Add({CInt(.Item("Renglon")) - 1, .Item("ClaveCheque"), .Item("BancoCheque") _
                                               , .Item("SucursalCheque"), .Item("ChequeCheque"), .Item("CuentaCheque") _
                                               , .Item("Cuit"), "", ""})
                        End With
                    End If

                    If Val(dr.Item("Tipo2")) = 4 Then
                        With dr
                            _CuentasContables.Add({Val(.Item("Renglon")) - WOffset, .Item("Cuenta")})
                        End With
                    End If

                Loop
            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally


            cn.Close()



        End Try
    End Sub

    Private Function _NormalizarFecha(ByRef fecha As String) As Boolean
        Dim _FechaValida = True
        Dim _Fecha As String() = fecha.Split("/")

        Try
            _Fecha(0) = Val(_Fecha(0)).ToString()
            _Fecha(1) = Val(_Fecha(1)).ToString()
            _Fecha(2) = Val(_Fecha(2)).ToString()

            fecha = String.Join("/", _Fecha)

            fecha = Date.ParseExact(fecha, "d/M/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo).ToString("dd/MM/yyyy")
        Catch ex As Exception
            _FechaValida = False
            MsgBox("El formato de la fecha ingresada no es válido.", MsgBoxStyle.Information)
        End Try

        Return _FechaValida
    End Function

    Private Sub mostrarTipoRecibo(ByVal tipo As Integer)
        Select Case tipo
            Case 3
                optVarios.Checked = True
            Case 2
                optAnticipos.Checked = True
            Case Else
                optCtaCte.Checked = True
        End Select
    End Sub

    Private Sub mostrarReciboProvisorio(ByVal _ReciboProvisorio As String)
        Dim cn = New SqlConnection()
        Dim cm = New SqlCommand("SELECT * FROM RecibosProvi WHERE Recibo = '" + _ReciboProvisorio + "' order by Renglon")
        Dim dr As SqlDataReader
        Dim _reciboProvi = 0

        If Trim(_ReciboProvisorio) = "" Then
            Exit Sub
        End If

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then
                btnLimpiar.PerformClick()
                _ClavesCheques.Clear()
                _CuentasContables.Clear()
                gridFormasPago2.Rows.Clear()

                Do While dr.Read()
                    Dim renglon As String = dr.Item("Renglon").ToString()

                    _reciboProvi = IIf(IsDBNull(dr.Item("ReciboDefinitivo")), 0, dr.Item("ReciboDefinitivo"))

                    If Val(_reciboProvi) <> 0 Then
                        MsgBox("El recibo Provisorio indicado ya se encuentra asociado a un Recibo Definitivo y no se encuentra disponible.", MsgBoxStyle.Information)
                        Exit Try
                    End If


                    If renglon = "01" Then
                        txtProvi.Text = IIf(IsDBNull(dr.Item("Recibo")), "", dr.Item("Recibo"))
                        txtFecha.Text = IIf(IsDBNull(dr.Item("Fecha")), "", dr.Item("Fecha"))
                        txtCliente.Text = IIf(IsDBNull(dr.Item("Cliente")), "", dr.Item("Cliente"))
                        Dim cliente As Cliente = DAOCliente.buscarClientePorCodigo(txtCliente.Text)

                        If Not IsNothing(cliente) Then
                            txtNombre.Text = cliente.razon
                            cliente = Nothing
                        End If

                        _DeterminarParidad()

                        txtRetGanancias.Text = IIf(IsDBNull(dr.Item("RetGanancias")), "0", dr.Item("RetGanancias"))
                        _ComprobanteRetGanancias = IIf(IsDBNull(dr.Item("ComproGanan")), "", dr.Item("ComproGanan"))
                        txtRetIva.Text = IIf(IsDBNull(dr.Item("RetIva")), "0", dr.Item("RetIva"))
                        _ComprobanteRetIva = IIf(IsDBNull(dr.Item("ComproIva")), "", dr.Item("ComproIva"))
                        txtRetSuss.Text = IIf(IsDBNull(dr.Item("RetSuss")), "0", dr.Item("RetSuss"))
                        _ComprobanteRetSuss = IIf(IsDBNull(dr.Item("ComproSuss")), "", dr.Item("ComproSuss"))
                        txtRetIB.Text = IIf(IsDBNull(dr.Item("RetOtra")), "0", dr.Item("RetOtra"))
                        'txtParidad.Text = IIf(IsDBNull(dr.Item("Paridad")), "0", dr.Item("Paridad"))
                        _RetIB1 = IIf(IsDBNull(dr.Item("RetIb1")), "0", dr.Item("RetIb1"))
                        _CompIB1 = IIf(IsDBNull(dr.Item("NroRetIb1")), "", dr.Item("NroRetIb1"))
                        _RetIB2 = IIf(IsDBNull(dr.Item("RetIb2")), "0", dr.Item("RetIb2"))
                        _CompIB2 = IIf(IsDBNull(dr.Item("NroRetIb2")), "", dr.Item("NroRetIb2"))
                        _RetIB3 = IIf(IsDBNull(dr.Item("RetIb3")), "0", dr.Item("RetIb3"))
                        _CompIB3 = IIf(IsDBNull(dr.Item("NroRetIb3")), "", dr.Item("NroRetIb3"))
                        _RetIB4 = IIf(IsDBNull(dr.Item("RetIb4")), "0", dr.Item("RetIb4"))
                        _CompIB4 = IIf(IsDBNull(dr.Item("NroRetIb4")), "", dr.Item("NroRetIb4"))
                        _RetIB5 = IIf(IsDBNull(dr.Item("RetIb5")), "0", dr.Item("RetIb5"))
                        _CompIB5 = IIf(IsDBNull(dr.Item("NroRetIb5")), "", dr.Item("NroRetIb5"))
                        _RetIB6 = IIf(IsDBNull(dr.Item("RetIb6")), "0", dr.Item("RetIb6"))
                        _CompIB6 = IIf(IsDBNull(dr.Item("NroRetIb6")), "", dr.Item("NroRetIb6"))
                        _RetIB7 = IIf(IsDBNull(dr.Item("RetIb7")), "0", dr.Item("RetIb7"))
                        _CompIB7 = IIf(IsDBNull(dr.Item("NroRetIb7")), "", dr.Item("NroRetIb7"))
                        _RetIB8 = IIf(IsDBNull(dr.Item("RetIb8")), "0", dr.Item("RetIb8"))
                        _CompIB8 = IIf(IsDBNull(dr.Item("NroRetIb8")), "", dr.Item("NroRetIb8"))

                        _RetIB9 = IIf(IsDBNull(dr.Item("RetIb9")), "", dr.Item("RetIb9"))
                        _CompIB9 = IIf(IsDBNull(dr.Item("NroRetIb9")), "", dr.Item("NroRetIb9"))
                        _RetIB10 = IIf(IsDBNull(dr.Item("RetIb10")), "", dr.Item("RetIb10"))
                        _CompIB10 = IIf(IsDBNull(dr.Item("NroRetIb10")), "", dr.Item("NroRetIb10"))
                        _RetIB11 = IIf(IsDBNull(dr.Item("RetIb11")), "", dr.Item("RetIb11"))
                        _CompIB11 = IIf(IsDBNull(dr.Item("NroRetIb11")), "", dr.Item("NroRetIb11"))
                        _RetIB12 = IIf(IsDBNull(dr.Item("RetIb12")), "", dr.Item("RetIb12"))
                        _CompIB12 = IIf(IsDBNull(dr.Item("NroRetIb12")), "", dr.Item("NroRetIb12"))
                        _RetIB13 = IIf(IsDBNull(dr.Item("RetIb13")), "", dr.Item("RetIb13"))
                        _CompIB13 = IIf(IsDBNull(dr.Item("NroRetIb13")), "", dr.Item("NroRetIb13"))
                        _RetIB14 = IIf(IsDBNull(dr.Item("RetIb14")), "", dr.Item("RetIb14"))
                        _CompIB14 = IIf(IsDBNull(dr.Item("NroRetIb14")), "", dr.Item("NroRetIb14"))

                        If Val(IIf(IsDBNull(dr.Item("TipoRec")), "0", dr.Item("TipoRec"))) = 1 Then
                            optCtaCte.Checked = True
                        ElseIf Val(IIf(IsDBNull(dr.Item("TipoRec")), "0", dr.Item("TipoRec"))) = 2 Then
                            optAnticipos.Checked = True
                        End If

                    End If

                    Dim _FormaPagoActual As Integer = gridFormasPago2.Rows.Add()

                    With gridFormasPago2.Rows(_FormaPagoActual)
                        .Cells(0).Value = IIf(IsDBNull(dr.Item("Tipo2")), "", dr.Item("Tipo2"))
                        .Cells(1).Value = IIf(IsDBNull(dr.Item("Numero2")), "", dr.Item("Numero2"))
                        .Cells(2).Value = IIf(IsDBNull(dr.Item("Fecha2")), "", dr.Item("Fecha2"))
                        .Cells(3).Value = IIf(IsDBNull(dr.Item("banco2")), "", dr.Item("banco2"))
                        .Cells("NroCuit").Value = IIf(IsDBNull(dr.Item("Cuit")), "00000000000", dr.Item("Cuit"))
                        .Cells(4).Value = _NormalizarNumero(Str$(IIf(IsDBNull(dr.Item("Importe2")), "0", dr.Item("Importe2"))))
                    End With

                    If Val(dr.Item("Tipo2")) = 2 Then
                        With dr
                            _ClavesCheques.Add({CInt(IIf(IsDBNull(.Item("Renglon")), "", .Item("Renglon"))) - 1, _
                                                     IIf(IsDBNull(.Item("ClaveCheque")), "", .Item("ClaveCheque")), _
                                                     IIf(IsDBNull(.Item("BancoCheque")), "", .Item("BancoCheque")), _
                                                     IIf(IsDBNull(.Item("SucursalCheque")), "", .Item("SucursalCheque")), _
                                                     IIf(IsDBNull(.Item("ChequeCheque")), "", .Item("ChequeCheque")), _
                                                     IIf(IsDBNull(.Item("CuentaCheque")), "", .Item("CuentaCheque")), _
                                                     IIf(IsDBNull(.Item("Cuit")), "00000000000", .Item("Cuit")), _
                                                     "", _
                                                     "" _
                                                })

                        End With
                    End If

                    With dr
                        If Val(IIf(IsDBNull(.Item("Tipo2")), "0", .Item("Tipo2"))) = 4 Then

                            _CuentasContables.Add({Val(.Item("Renglon")) - 1, IIf(IsDBNull(.Item("Cuenta")), "", .Item("Cuenta"))})

                        End If
                    End With

                Loop

                gridFormasPago2.Rows.Add()

                _SumarCreditos()

                txtRecibo.Text = "0"
                'txtParidad.Focus()

                If Trim(txtCliente.Text) <> "" Then
                    btnCtaCte.PerformClick()
                End If

            Else

                MsgBox("El Recibo Provisorio indicado no existe.", MsgBoxStyle.Information)

            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar el recibo provisorio.", MsgBoxStyle.Critical)
        Finally


            cn.Close()


        End Try
    End Sub

    Private Sub _DeterminarParidad(Optional ByVal fecha As String = "")
        Dim _Fecha As String = IIf(fecha = "", txtFecha.Text, fecha)
        Dim cn = New SqlConnection()
        Dim cm = New SqlCommand("SELECT Cambio FROM Cambios WHERE Fecha = '" & _Fecha & "'")
        Dim dr As SqlDataReader

        If Trim(txtFecha.Text) <> "" Then

            SQLConnector.conexionSql(cn, cm)

            Try

                dr = cm.ExecuteReader()

                If dr.HasRows Then

                    dr.Read()

                    txtParidad.Text = _NormalizarNumero(dr.Item("Cambio"))

                Else
                    txtParidad.Text = ""
                End If

            Catch ex As Exception
                MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
            Finally

                cn.Close()

            End Try
        End If

    End Sub

    Private Function _NormalizarNumero(ByVal numero As String) As String

        Return Proceso.formatonumerico(numero)

    End Function

    Private Function _SumarCreditos() As Boolean
        Dim _Error = False

        txtRetGanancias.Text = _NormalizarNumero(txtRetGanancias.Text)
        txtRetIva.Text = _NormalizarNumero(txtRetIva.Text)
        txtRetIB.Text = _NormalizarNumero(txtRetIB.Text)
        txtRetSuss.Text = _NormalizarNumero(txtRetSuss.Text)

        lblTotalCreditos.Text = Val(txtRetGanancias.Text) + Val(txtRetIva.Text) + Val(txtRetIB.Text) + Val(txtRetSuss.Text)

        For Each row As DataGridViewRow In gridFormasPago2.Rows

            If Val(row.Cells(4).Value) <> 0 Then

                If Val(row.Cells(0).Value) = 2 Then

                    If IsDate(row.Cells(2).Value) And IsDate(txtFecha.Text) And DateDiff(DateInterval.Day, CDate(row.Cells(2).Value), CDate(txtFecha.Text)) > 30 Then
                        _Error = True
                        Exit For
                    End If

                End If

                lblTotalCreditos.Text = Val(_NormalizarNumero(lblTotalCreditos.Text)) + Val(_NormalizarNumero(row.Cells(4).Value))

            End If

        Next

        lblTotalCreditos.Text = _NormalizarNumero(lblTotalCreditos.Text)
        lblDiferencia.Text = Val(_NormalizarNumero(lblTotalDebitos.Text)) - Val(_NormalizarNumero(lblTotalCreditos.Text))
        Return _Error
    End Function

    Private Sub mostrarPagos(ByVal pagos As List(Of Pago))
        gridPagos2.Rows.Clear()
        For Each pago As Pago In pagos
            gridPagos2.Rows.Add(pago.tipo, pago.letra, pago.punto, pago.numero, _NormalizarNumero(pago.importe))
            WOffset += 1
        Next
    End Sub

    Private Sub mostrarFormasPago(ByVal formasPago As List(Of FormaPago))
        gridFormasPago2.Rows.Clear()
        Dim row As Integer
        For Each forma As FormaPago In formasPago
            row = gridFormasPago2.Rows.Add

            With gridFormasPago2.Rows(row)
                .Cells(0).Value = forma.tipo
                .Cells(1).Value = forma.numero
                .Cells(2).Value = forma.fecha
                .Cells(3).Value = forma.nombre
                .Cells(4).Value = _NormalizarNumero(forma.importe)
            End With

        Next
    End Sub

    Private Sub mostrarCliente(ByVal cliente As String)
        Dim cn = New SqlConnection()
        Dim cm = New SqlCommand("SELECT c.Cliente, c.Razon, c.Direccion, c.Localidad, c.Provincia, p.Nombre, c.Postal, c.EmailFacturaII as Email " _
                                              & " FROM Cliente as c, Provincia as p WHERE c.Cliente = '" & cliente & "' AND p.Provincia = c.Provincia")
        Dim dr As SqlDataReader

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then

                dr.Read()

                txtCliente.Text = dr.Item("Cliente")
                txtNombre.Text = dr.Item("Razon")
                _Provincia = Val(dr.Item("Provincia"))
                WRazon = dr.Item("Razon")
                WDireccion = dr.Item("Direccion")
                WLocalidad = dr.Item("Localidad")
                WProvincia = dr.Item("Nombre")
                WPostal = dr.Item("Postal")
                WEmail = IIf(IsDBNull(dr.Item("Email")), "", dr.Item("Email"))
            Else
                MsgBox("El cliente especificado, no existe. Compruebe y vuelva a intentarlo.")
            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer cargar los datos del Cliente solicitado.", MsgBoxStyle.Critical)
        Finally


            cn.Close()


        End Try
    End Sub

    Private Sub mostrarCuenta(ByVal cuenta As String)
        Dim cuentaCompleta As CuentaContable = DAOCuentaContable.buscarCuentaContablePorCodigo(cuenta)

        If cuenta = "" Or IsNothing(cuentaCompleta) Then
            txtNombreCuenta.Text = ""
        Else
            txtCuenta.Text = cuentaCompleta.id
            txtNombreCuenta.Text = cuentaCompleta.descripcion
        End If
    End Sub

    Private Sub mostrarCuentaContable(ByVal cuentaCompleta As CuentaContable)

        txtCuenta.Text = cuentaCompleta.id
        txtNombreCuenta.Text = cuentaCompleta.descripcion

    End Sub


    Private Sub txtCliente_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCliente.Leave

        If Trim(txtCliente.Text) <> "" Then
            mostrarCliente(txtCliente.Text)
        End If

    End Sub

    Private Sub txtProvi_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProvi.Leave

        If Trim(txtProvi.Text) <> "" Then
            txtProvi.Text = ceros(txtProvi.Text, 6)
        End If

    End Sub

    Private Function tipoRecibo()
        If optCtaCte.Checked Then
            Return 1
        Else
            If optAnticipos.Checked Then
                Return 2
            End If
        End If
        Return 3
    End Function

    Private Sub btnAgregar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAgregar.Click
        ' Comprobamos que el numero de recibo sea cero. En caso de que no,
        ' Chequeamos si el numero se corresponde con algun Recibo existente y se avisa al usuario de que no se permite actualizar.
        Dim WMarcaDiferencia = ""


        ' Comprobamos que todos los Tipos de Pagos de Tipo 2, tengan un numero definido y distinto de 0

        For Each row In gridFormasPago2.Rows

            With row
                If Not IsNothing(.Cells(0).Value) AndAlso Val(.Cells(0).Value) = 2 Then

                    ' Controlamos que no se ingrese un numero = 0 para los cheques.
                    If IsNothing(.Cells(1).Value) OrElse Val(.Cells(1).Value) = 0 Then
                        MsgBox("Se debe informar un numero valido para el cheque.", MsgBoxStyle.Exclamation)
                        Exit Sub
                    End If

                End If

                If Not IsNothing(.Cells(0).Value) AndAlso Val(.Cells(0).Value) = 3 Then
                    MsgBox("No se puede informar 'Documentos' como Forma de Pago.", MsgBoxStyle.Exclamation)
                    Exit Sub
                End If

            End With

        Next


        If Val(txtRecibo.Text) <> 0 Then

            If _ReciboYaExistente(Trim(txtRecibo.Text)) Then
                MsgBox("El recibo ya se encuentra guardado con anterioridad y no puede ser actualizado.", MsgBoxStyle.Information)
            Else
                MsgBox("La numeracion del recibo es automatica, se debe grabar con numero 0")
            End If

            Exit Sub
        End If

        ' Controlamos que en caso de ser Varios el tipo, tenga cuenta y que sea valida.
        If optVarios.Checked Then
            If Trim(txtCuenta.Text) = "" Then
                MsgBox("Debe indicarse una cuenta contable.", MsgBoxStyle.Information)
                Exit Sub
            Else
                Dim cuenta As CuentaContable = DAOCuentaContable.buscarCuentaContablePorCodigo(Trim(txtCuenta.Text))

                If IsNothing(cuenta) Then
                    MsgBox("Debe indicarse una cuenta contable válida.", MsgBoxStyle.Information)
                    Exit Sub
                End If
            End If
        End If

        gridFormasPago2.CommitEdit(DataGridViewDataErrorContexts.Commit) ' Guardamos todos los datos que no hayan sido confirmados.
        gridPagos2.CommitEdit(DataGridViewDataErrorContexts.Commit) ' Guardamos todos los datos que no hayan sido confirmados.

        ' Calculamos nuevamente los saldo de débitos y Créditos.
        If _ErrorEnSumaDeSaldos() Then
            MsgBox("Hay errores en la carga de debitos o creditos. Por favor, verifique los datos y vuelva a intentar.", MsgBoxStyle.Information)
            Exit Sub
        End If

        ' Comprobamos que haya valores con las que realizar el recibo
        'If Val(lblTotalCreditos.Text) = 0 And Val(lblTotalDebitos.Text) = 0 Then
        '    Exit Sub
        'End If

        ' Comprobamos que los datos sean coherentes.
        If Val(lblTotalCreditos.Text) = Val(lblTotalDebitos.Text) Or optAnticipos.Checked Or optVarios.Checked Then

            ' Comprobamos que se hayan cargado tanto recibo como fecha. En caso de que no, salimos y cancelamos la grabación.
            If Trim(txtRecibo.Text) = "" Or Trim(txtFecha.Text) = "" Then
                MsgBox("Por favor verifique que tanto el número de Recibo como la fecha del mismo se encuentren cargados.", MsgBoxStyle.Information)
                Exit Sub
            End If

            ' Comprbamos la diferencia de cambio y consultamos con el usuario sobre si graba igual o no.
            If optCtaCte.Checked And _ExisteDiferenciaDeCambio() Then
                If MsgBox("Existe una diferencia de cambio de U$S " & lblDolares.Text & vbCrLf & "¿Desea grabar igualmente el recibo?", MsgBoxStyle.OkCancel, MsgBoxStyle.Information) = DialogResult.Cancel Then
                    Exit Sub
                    ' Comentado hasta que este completamente desarrollado el resto.
                    'Else
                    '    If MsgBox("¿Desea avisar a Ventas sobre esta diferencia de Cambio?", MsgBoxStyle.YesNo, MsgBoxStyle.Information) = DialogResult.Yes Then
                    '        WMarcaDiferencia = "X"
                    '    End If
                End If
            End If

            ' Comprobamos que los creditos indicados como de tipo 4 tengan su respectiva cuenta contable asignada.
            If _NoImputadosCorrectamenteValoresVarios() Then
                MsgBox("No se ha imputado correctamente ingreso de valores varios.", MsgBoxStyle.Information)
                Exit Sub
            End If

            ' Comprobamos que haya paridad cargada para la fecha
            If Val(txtParidad.Text) = 0 Then
                MsgBox("No hay paridad cargada para esta fecha")
                Exit Sub
            End If

            ' Damos de alta el recibo según el tipo de recibo definido por el usuario.

            ' Cargamos la numeración del recibo a guardar.
            txtRecibo.Text = _NumeracionDeReciboSiguiente()

            ' Normalizamos el número de recibo.
            txtRecibo.Text = ceros(txtRecibo.Text, 6)

            ' Comprobamos que ya no se haya grabado el recibo con anterioridad.
            If Not _NumeroDeReciboDisponible() Then
                MsgBox("El número de recibo ya se ha grabado con anterioridad y no se encuentra disponible.", MsgBoxStyle.Information)
                Exit Sub
            End If

            renglon = 0

            ' COMENZAR CON LAS ULTIMAS DOS QUE SON MAS CORTAS.
            If optCtaCte.Checked Then
                _AltaCtaCte(WMarcaDiferencia)
            ElseIf optAnticipos.Checked Then
                _AltaAnticipo(WMarcaDiferencia)
            ElseIf optVarios.Checked Then
                _AltaVarios(WMarcaDiferencia)
            End If

            ' Agregamos al registro del Recibo, los creditos ingresados por el usuario.
            _AltaCreditos(WMarcaDiferencia)

            ' Si se trata de un recibo de tipo CtaCte, creamos un nuevo registro de CtaCte
            If optCtaCte.Checked Then
                _AltaDeCuentaParaTipoCtaCte()
            End If

            ' Determinamos si todos los estados de los items del recibo, se encuentran listos y en caso de que si, se lo marca
            If _SeDebeMarcarRecibo() Then
                _MarcarRecibo()
            End If

            ' En caso de que se haya generado desde un recibo provisorio se lo asocia al recibo creado.
            If Val(txtProvi.Text) > 0 Then
                _AsignarReciboAProvisorio()
            End If

            ' Consultamos en caso de que desee agendar una llamada al Cliente al que se le hace el recibo.
            ' Comentado hasta que este completamente verificado y terminado.
            If MsgBox("¿Desea agregar un recordatorio en la agenda para este Cliente?", MsgBoxStyle.YesNo, MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                With AltaAgenda
                    .Cliente = Trim(txtCliente.Text)
                    .ShowDialog()
                    .Dispose()
                End With
            End If

            ' Si llegamos hasta aca se supone que todo salio bien.
            'MsgBox("Se ha generado correctamente el Recibo solicitado. Nº de Recibo: " & txtRecibo.Text, MsgBoxStyle.Information)

            ' Recargamos los datos del recibo.
            mostrarRecibo(DAORecibo.buscarRecibo(txtRecibo.Text))

            ' Imprimimos los recibos y enviamos email en caso de que corresponda.
            btnImpresion.PerformClick()

            btnLimpiar.PerformClick()
        Else

            MsgBox("Los montos de debitos y creditos no se corresponden entre si o con el Tipo de Cobro.", MsgBoxStyle.Information)
            Exit Sub

        End If

    End Sub

    Private Function _ReciboYaExistente(ByVal recibo As String) As Boolean
        Dim _existe = False
        cm.CommandText = "SELECT Recibo From Recibos WHERE Recibo = '" & recibo & "'"

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then

                _existe = True

            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally

            cn.Close()

        End Try

        Return _existe
    End Function

    Private Sub _AltaCreditos(ByVal WMarcaDiferencia As String)

        Dim iRow As Integer
        Dim _Cheque As Object = Nothing
        Dim _CuitExistente = False
        Dim cn = New SqlConnection()
        Dim cm = New SqlCommand()
        Dim dr As SqlDataReader

        ' Ahora guardamos los creditos.
        For iRow = 0 To gridFormasPago2.Rows.Count - 1

            If Not IsNothing(gridFormasPago2.Rows(iRow).Cells(4).Value) AndAlso Val(gridFormasPago2.Rows(iRow).Cells(4).Value) <> 0 Then
                renglon += 1

                XRecibo = txtRecibo.Text
                XRenglon = ceros(renglon, 2)
                XCliente = txtCliente.Text
                XFecha = txtFecha.Text
                XFechaOrd = String.Join("", txtFecha.Text.Split("/").Reverse())
                'XTipoRec = IIf(optVarios.Checked, "3", "1")

                If optVarios.Checked Then
                    XTipoRec = "3"
                ElseIf optAnticipos.Checked Then
                    XTipoRec = "2"
                Else
                    XTipoRec = "1"
                End If

                XRetganancias = Str$(Val(txtRetGanancias.Text))
                XRetIva = Str$(Val(txtRetIva.Text))
                XRetotra = Str$(Val(txtRetIB.Text))
                XRetencion = ""
                XTiporeg = "2"
                XTipo1 = ""
                XLetra1 = ""
                XPunto1 = ""
                XNumero1 = ""
                XImporte1 = ""
                XTipo2 = gridFormasPago2.Rows(iRow).Cells(0).Value
                XNumero2 = gridFormasPago2.Rows(iRow).Cells(1).Value
                XFecha2 = gridFormasPago2.Rows(iRow).Cells(2).Value
                XFechaOrd2 = String.Join("", XFecha2.Split("/").Reverse())
                XBanco2 = gridFormasPago2.Rows(iRow).Cells(3).Value
                XImporte2 = gridFormasPago2.Rows(iRow).Cells(4).Value
                XObservaciones = txtObservaciones.Text
                XEmpresa = "1"
                XClave = XRecibo + XRenglon
                XImporte = Str$(Val(lblTotalCreditos.Text))
                If XTipo2 = 4 Then
                    XCuenta = _CuentasContables.FindLast(Function(c) c(0) = iRow)(1)
                Else
                    XCuenta = ""
                End If
                XMarca = ""
                XFechaDepo = ""
                XFechaDepoOrd = ""

                XClaveCheque = ""
                XBancoCheque = ""
                XSucursalCheque = ""
                XChequeCheque = ""
                XCuentaCheque = ""
                XCuit = IIf(IsNothing(gridFormasPago2.Rows(iRow).Cells("NroCuit").Value), "00000000000", gridFormasPago2.Rows(iRow).Cells("NroCuit").Value)
                XEstado2 = ""
                XDestino = ""

                _Cheque = _ClavesCheques.FindLast(Function(c) c(0) = iRow)

                If Not IsNothing(_Cheque) Then
                    XClaveCheque = _Cheque(1)
                    XBancoCheque = _Cheque(2)
                    XSucursalCheque = _Cheque(3)
                    XChequeCheque = _Cheque(4)
                    XCuentaCheque = _Cheque(5)
                    XCuit = _Cheque(6)
                    XEstado2 = _Cheque(7)
                    XDestino = _Cheque(8)
                End If

                If Trim(XEstado2) = "" Then
                    XEstado2 = "P"
                End If

                If Trim(txtProvi.Text) <> "" Then
                    cm.CommandText = "SELECT Estado2 FROM RecibosProvi Where Recibo = '" & txtProvi.Text & "'  AND Numero2 = '" & XNumero2 & "'"
                    SQLConnector.conexionSql(cn, cm)

                    Try

                        dr = cm.ExecuteReader()

                        If dr.HasRows Then

                            With dr
                                .Read()

                                XEstado2 = Trim(UCase(.Item("Estado2")))

                            End With

                        End If

                    Catch ex As Exception
                        MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
                    Finally


                        cn.Close()


                        XSql = ""
                        XParam = ""

                    End Try
                End If

                If Val(XTipo2) = 1 Or Val(XTipo2) = 4 Then
                    XEstado2 = "X"
                End If

                XParam = "'" + XClave + "','" _
                        + XRecibo + "','" + XRenglon + "','" _
                        + XCliente + "','" _
                        + XFecha + "','" + XFechaOrd + "','" _
                        + XTipoRec + "','" _
                        + XRetganancias + "','" _
                        + XRetIva + "','" + XRetotra + "','" _
                        + XRetencion + "','" _
                        + XTiporeg + "','" _
                        + XTipo1 + "','" + XLetra1 + "','" _
                        + XPunto1 + "','" + XNumero1 + "','" _
                        + XImporte1 + "','" _
                        + XTipo2 + "','" + XNumero2 + "','" _
                        + XFecha2 + "','" + XBanco2 + "','" _
                        + XImporte2 + "','" + XEstado2 + "','" _
                        + XEmpresa + "','" _
                        + XFechaOrd2 + "','" _
                        + XImporte + "','" _
                        + XObservaciones + "','" _
                        + XImpolist + "','" + XImpo1list + "','" _
                        + XDestino + "','" _
                        + XCuenta + "','" _
                        + XMarca + "','" _
                        + XFechaDepo + "','" _
                        + XFechaDepoOrd + "','" _
                        + WMarcaDiferencia + "'"

                ' Damos de alta el nuevo renglon del nuevo recibo.
                XSql = "INSERT INTO  Recibos (Clave, Recibo, Renglon, Cliente, Fecha, Fechaord, TipoRec, RetGanancias, RetIva, RetOtra, Retencion," _
                & "TipoReg, Tipo1, Letra1, Punto1, Numero1, Importe1, Tipo2, Numero2, Fecha2, banco2, Importe2," _
                & "Estado2, Empresa, FechaOrd2, Importe, Observaciones, Impolist, Impo1list, Destino, Cuenta, Marca," _
                & "FechaDepo, FechaDepoOrd, MarcaDiferencia) " _
                & "VALUES(" & XParam & ")"

                cm.CommandText = XSql
                SQLConnector.conexionSql(cn, cm)

                Try
                    cm.ExecuteNonQuery()

                Catch ex As Exception
                    MsgBox("Hubo un problema al querer guardar el nuevo recibo.", MsgBoxStyle.Critical)
                    Exit Sub
                Finally


                    cn.Close()


                    XSql = ""
                    XParam = ""

                End Try

                XSql = "UPDATE Recibos SET " _
                 & " ClaveCheque = " + "'" + XClaveCheque + "'," _
                 & " Cuit = " + "'" + XCuit + "'," _
                 & " Provisorio = " + "'" + txtProvi.Text + "'," _
                 & " BancoCheque = " + "'" + XBancoCheque + "'," _
                 & " SucursalCheque = " + "'" + XSucursalCheque + "'," _
                 & " ChequeCheque = " + "'" + XChequeCheque + "'," _
                 & " CuentaCheque = " + "'" + XCuentaCheque + "'" _
                 & " Where Clave = " + "'" + XClave + "'"

                cm.CommandText = XSql
                SQLConnector.conexionSql(cn, cm)

                Try
                    cm.ExecuteNonQuery()

                Catch ex As Exception
                    MsgBox("Hubo un problema al querer actualizar el recibo corriente.", MsgBoxStyle.Critical)
                    Exit Sub
                Finally


                    cn.Close()


                    XSql = ""
                    XParam = ""

                End Try

                XClaveCuit = XBancoCheque + XSucursalCheque + XCuentaCheque
                XDestino = ""

                If Trim(XCuit) <> "" Then

                    cm.CommandText = "SELECT Clave FROM Cuit Where Clave = '" & XClaveCuit & "'"
                    SQLConnector.conexionSql(cn, cm)

                    Try

                        dr = cm.ExecuteReader()

                        If dr.HasRows Then

                            _CuitExistente = True

                        End If

                    Catch ex As Exception
                        MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
                    Finally


                        cn.Close()


                        XSql = ""
                        XParam = ""

                    End Try

                    If Not _CuitExistente Then

                        ' Damos de alta el nuevo Cuit.
                        XSql = "INSERT INTO Cuit (Clave, Banco, Sucursal, Cuenta, Cuit) Values ('" + XClaveCuit + "','" + XBancoCheque + "','" + XSucursalCheque + "','" + XCuentaCheque + "', '" + XCuit + "')"

                        cm.CommandText = XSql
                        SQLConnector.conexionSql(cn, cm)

                        Try
                            cm.ExecuteNonQuery()

                        Catch ex As Exception
                            MsgBox("Hubo un problema al querer guardar el nuevo Cuit.", MsgBoxStyle.Critical)
                            Exit Sub
                        Finally


                            cn.Close()


                            XSql = ""
                            XParam = ""

                        End Try

                    End If

                End If

                If Val(gridFormasPago2.Rows(iRow).Cells(0).Value) = 3 Then
                    XTipo = "50"
                    XNumero = ceros(gridFormasPago2.Rows(iRow).Cells(1).Value, 8)
                    XRenglon = "01"
                    XCliente = txtCliente.Text
                    XFecha = txtFecha.Text
                    XEstado = "1"
                    XVencimiento = gridFormasPago2.Rows(iRow).Cells(2).Value
                    XVencimiento1 = XVencimiento
                    XTotal = gridFormasPago2.Rows(iRow).Cells(4).Value
                    XTotalUs = XTotal
                    XSaldo = XTotal
                    XSaldoUs = XTotal
                    XOrdFecha = String.Join("", txtFecha.Text.Split("/").Reverse())
                    XOrdVencimiento = String.Join("", XVencimiento.Split("/").Reverse())
                    XOrdVencimiento1 = String.Join("", XVencimiento1.Split("/").Reverse())
                    XImpre = "Dc"
                    XNet = ""
                    XIva1 = ""
                    XIva2 = ""
                    XImpoIb = ""
                    XSeguro = ""
                    XFlete = ""
                    XPedido = ""
                    XRemito = ""
                    XOrden = ""
                    XParidad = txtParidad.Text
                    XProvincia = _Provincia
                    XVendedor = "0"
                    XRubro = "0"
                    XComprobante = ""
                    XAceptada = ""
                    XCosto = ""
                    XImporte1 = ""
                    XImporte2 = ""
                    XImporte3 = ""
                    XImporte4 = ""
                    XImporte5 = ""
                    XImporte6 = ""
                    XImporte7 = ""
                    XClave = "50" + XNumero + "01"
                    XDate = Date.Now.ToString("MM-dd-yyyy")

                    XParam = "'" & XClave & "','" _
                             & XTipo & "','" & XNumero & "','" _
                             & XRenglon & "','" & XCliente & "','" _
                             & XFecha & "','" & XEstado & "','" _
                             & XVencimiento & "','" & XVencimiento1 & "','" _
                             & XTotal & "','" & XTotalUs & "','" _
                             & XSaldo & "','" & XSaldoUs & "','" _
                             & XOrdFecha & "','" & XOrdVencimiento & "','" _
                             & XOrdVencimiento1 & "','" & XImpre & "','" _
                             & XEmpresa & "','" _
                             & XNet & "','" & XIva1 & "','" _
                             & XIva2 & "','" & XPedido & "','" _
                             & XRemito & "','" & XOrden & "','" _
                             & XParidad & "','" & XProvincia & "','" _
                             & XVendedor & "','" & XRubro & "','" _
                             & XComprobante & "','" & XAceptada & "','" _
                             & XCosto & "','" _
                             & XImporte1 & "','" & XImporte2 & "','" _
                             & XImporte3 & "','" & XImporte4 & "','" _
                             & XImporte5 & "','" & XImporte6 & "','" _
                             & XImporte7 & "','" & XDate & "','" _
                             & XSeguro & "','" & XFlete & "','" _
                             & XImpoIb & "'"

                    XSql = "INSERT INTO  Ctacte (Clave, Tipo, numero, Renglon, Cliente, fecha, Estado, Vencimiento, Vencimiento1," _
                            & "Total, TotalUs, Saldo, SaldoUs, Ordfecha, Ordvencimiento, Ordvencimiento1, Impre, Empresa, Neto, Iva1," _
                            & "Iva2, Pedido, Remito, Orden, Paridad, Provincia, Vendedor, Rubro, Comprobante, Aceptada, Costo, Importe1," _
                            & "Importe2, Importe3, Importe4, Importe5, Importe6, Importe7, WDate, Seguro, Flete, ImpoIb) " _
                            & "VALUES(" & XParam & ")"

                    cm.CommandText = XSql
                    SQLConnector.conexionSql(cn, cm)

                    Try
                        cm.ExecuteNonQuery()

                    Catch ex As Exception
                        MsgBox("Hubo un problema al querer guardar la nueva Cuenta Corriente.", MsgBoxStyle.Critical)
                        Exit Sub
                    Finally


                        cn.Close()


                        XSql = ""
                        XParam = ""

                    End Try
                End If
            End If
        Next

        _ActualizarComprobantes()

    End Sub

    Private Sub _AsignarReciboAProvisorio()
        XSql = ""
        XSql = XSql & "UPDATE RecibosProvi SET "
        XSql = XSql & " ReciboDefinitivo = " & "'" & txtRecibo.Text & "'"
        XSql = XSql & " Where Recibo = " & "'" & txtProvi.Text & "'"

        cm.CommandText = XSql
        SQLConnector.conexionSql(cn, cm)

        Try

            cm.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally


            cn.Close()



        End Try

    End Sub

    Private Sub _MarcarRecibo()

        XFechaDepoOrd = String.Join("", txtFecha.Text.Split("/").Reverse())

        cm.CommandText = "UPDATE Recibos SET Marca = 'X', FechaDepo = '" & txtFecha.Text & "', FechaDepoOrd = '" & XFechaDepoOrd & "' WHERE Recibo = '" & txtRecibo.Text & "'"
        SQLConnector.conexionSql(cn, cm)

        Try

            cm.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally


            cn.Close()



        End Try
    End Sub

    Private Function _SeDebeMarcarRecibo() As Boolean
        Dim _Marcar = True

        cm.CommandText = "SELECT Estado2 FROM Recibos WHERE Recibo = '" & txtRecibo.Text & "'"
        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then

                Do While dr.Read()

                    If Trim(UCase(dr.Item("Estado2"))) <> "X" Then
                        _Marcar = False
                        Exit Do
                    End If

                Loop

            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally


            cn.Close()



        End Try

        Return _Marcar
    End Function

    Private Sub _ActualizarComprobantes()

        ' Actualizamos los datos de retenciones, comprobante y diferencias de cambio.
        XSql = ""
        XSql = XSql & "UPDATE Recibos SET "
        XSql = XSql & " DifCambio = " & _NormalizarNumero(lblDolares.Text) & ","
        XSql = XSql & " RetSuss = " & _NormalizarNumero(txtRetSuss.Text) & ","
        XSql = XSql & " ComproGanan = " & "'" & Val(Trim(_ComprobanteRetGanancias)) & "',"
        XSql = XSql & " ComproIva = " & "'" & Val(Trim(_ComprobanteRetIva)) & "',"
        XSql = XSql & " ComproIb = ''," 'Comprobar si se sigue colocando o no. Me parece que no.
        XSql = XSql & " ComproSuss = " & "'" & Val(Trim(_ComprobanteRetSuss)) & "',"
        XSql = XSql & " NroRetIb1 = " & _NormalizarNumero(_CompIB1) & ","
        XSql = XSql & " NroRetIb2 = " & _NormalizarNumero(_CompIB2) & ","
        XSql = XSql & " NroRetIb3 = " & _NormalizarNumero(_CompIB3) & ","
        XSql = XSql & " NroRetIb4 = " & _NormalizarNumero(_CompIB4) & ","
        XSql = XSql & " NroRetIb5 = " & _NormalizarNumero(_CompIB5) & ","
        XSql = XSql & " NroRetIb6 = " & _NormalizarNumero(_CompIB6) & ","
        XSql = XSql & " NroRetIb7 = " & _NormalizarNumero(_CompIB7) & ","
        XSql = XSql & " NroRetIb8 = " & _NormalizarNumero(_CompIB8) & ","
        XSql = XSql & " NroRetIb9 = " & _NormalizarNumero(_CompIB9) & ","
        XSql = XSql & " NroRetIb10 = " & _NormalizarNumero(_CompIB10) & ","
        XSql = XSql & " NroRetIb11 = " & _NormalizarNumero(_CompIB11) & ","
        XSql = XSql & " NroRetIb12 = " & _NormalizarNumero(_CompIB12) & ","
        XSql = XSql & " NroRetIb13 = " & _NormalizarNumero(_CompIB13) & ","
        XSql = XSql & " NroRetIb14 = " & _NormalizarNumero(_CompIB14) & ","
        XSql = XSql & " RetIb1 = " & _NormalizarNumero(_RetIB1) & ","
        XSql = XSql & " RetIb2 = " & _NormalizarNumero(_RetIB2) & ","
        XSql = XSql & " RetIb3 = " & _NormalizarNumero(_RetIB3) & ","
        XSql = XSql & " RetIb4 = " & _NormalizarNumero(_RetIB4) & ","
        XSql = XSql & " RetIb5 = " & _NormalizarNumero(_RetIB5) & ","
        XSql = XSql & " RetIb6 = " & _NormalizarNumero(_RetIB6) & ","
        XSql = XSql & " RetIb7 = " & _NormalizarNumero(_RetIB7) & ","
        XSql = XSql & " RetIb8 = " & _NormalizarNumero(_RetIB8) & ","
        XSql = XSql & " RetIb9 = " & _NormalizarNumero(_RetIB9) & ","
        XSql = XSql & " RetIb10 = " & _NormalizarNumero(_RetIB10) & ","
        XSql = XSql & " RetIb11 = " & _NormalizarNumero(_RetIB11) & ","
        XSql = XSql & " RetIb12 = " & _NormalizarNumero(_RetIB12) & ","
        XSql = XSql & " RetIb13 = " & _NormalizarNumero(_RetIB13) & ","
        XSql = XSql & " RetIb14 = " & _NormalizarNumero(_RetIB14) & ""
        XSql = XSql & " Where Recibo = " & "'" & txtRecibo.Text & "'"

        cm.CommandText = XSql
        SQLConnector.conexionSql(cn, cm)

        Try
            cm.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Hubo un problema al querer actualizar el recibo.", MsgBoxStyle.Critical)
            Exit Sub
        Finally


            cn.Close()


            XSql = ""
            XParam = ""

        End Try

    End Sub

    Private Sub _AltaDeCuentaParaTipoCtaCte()
        ' Damos de alta la Cuenta Cte asociada al recibo por Cta Cte.
        XTipo = "06"
        XNumero = "00" + txtRecibo.Text
        ClaveCtaCte = XTipo + XNumero + "01"
        XRenglon = "01"
        XCliente = txtCliente.Text
        XFecha = txtFecha.Text
        XEstado = "1"
        XVencimiento = XFecha
        XVencimiento1 = XFecha
        If Val(_Provincia) = 24 Then
            XTotal = Str$(Val(lblTotalCreditos.Text) * -1 / Val(txtParidad.Text))
        Else
            XTotal = Str$(Val(lblTotalCreditos.Text) * -1)
        End If
        XTotalUs = Str$((Val(lblTotalCreditos.Text) * -1) / Val(txtParidad.Text))
        XSaldo = ""
        XSaldoUs = ""
        XOrdFecha = String.Join("", XFecha.Split("/").Reverse())
        XOrdVencimiento = XOrdFecha
        XOrdVencimiento1 = XOrdFecha
        XImpre = "RC"
        XNet = ""
        XIva1 = "2"
        XIva2 = ""
        XImpoIb = ""
        XSeguro = ""
        XFlete = ""
        XPedido = ""
        XRemito = ""
        XOrden = ""
        XParidad = txtParidad.Text
        XProvincia = _Provincia
        XVendedor = "0"
        XRubro = "0"
        XComprobante = ""
        XAceptada = ""
        XCosto = ""
        XImporte1 = ""
        XImporte2 = ""
        XImporte3 = ""
        XImporte4 = ""
        XImporte5 = ""
        XImporte6 = ""
        XImporte7 = ""
        XClave = XTipo & ceros(XNumero, 8) & "01"
        XDate = Date.Now.ToString("MM-dd-yyyy")

        XParam = "'" + XClave + "','" _
                + XTipo + "','" + XNumero + "','" _
                + XRenglon + "','" + XCliente + "','" _
                + XFecha + "','" + XEstado + "','" _
                + XVencimiento + "','" + XVencimiento1 + "','" _
                + XTotal + "','" + XTotalUs + "','" _
                + XSaldo + "','" + XSaldoUs + "','" _
                + XOrdFecha + "','" + XOrdVencimiento + "','" _
                + XOrdVencimiento1 + "','" + XImpre + "','" _
                + XEmpresa + "','" _
                + XNet + "','" + XIva1 + "','" _
                + XIva2 + "','" + XPedido + "','" _
                + XRemito + "','" + XOrden + "','" _
                + XParidad + "','" + XProvincia + "','" _
                + XVendedor + "','" + XRubro + "','" _
                + XComprobante + "','" + XAceptada + "','" _
                + XCosto + "','" _
                + XImporte1 + "','" + XImporte2 + "','" _
                + XImporte3 + "','" + XImporte4 + "','" _
                + XImporte5 + "','" + XImporte6 + "','" _
                + XImporte7 + "','" + XDate + "','" _
                + XSeguro + "','" + XFlete + "','" _
                + XImpoIb + "'"

        XSql = "INSERT INTO  Ctacte (Clave, Tipo, numero, Renglon, Cliente, fecha, Estado, Vencimiento, Vencimiento1," _
        & "Total, TotalUs, Saldo, SaldoUs, Ordfecha, Ordvencimiento, Ordvencimiento1, Impre, Empresa, Neto, Iva1," _
        & "Iva2, Pedido, Remito, Orden, Paridad, Provincia, Vendedor, Rubro, Comprobante, Aceptada, Costo, Importe1," _
        & "Importe2, Importe3, Importe4, Importe5, Importe6, Importe7, WDate, Seguro, Flete, ImpoIb) " _
        & "VALUES(" & XParam & ")"

        cm.CommandText = XSql
        SQLConnector.conexionSql(cn, cm)

        Try
            cm.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Hubo un problema al querer guardar la nueva Cuenta Corriente.", MsgBoxStyle.Critical)
            Exit Sub
        Finally


            cn.Close()


            XSql = ""
            XParam = ""

        End Try
    End Sub

    Private Sub _AltaCtaCte(ByVal WMarcaDiferencia As String)

        Dim iRow, WRow As Integer
        Dim _Cheque As Object = Nothing
        Dim _CuitExistente = False
        Dim cn = New SqlConnection()
        Dim cm = New SqlCommand()
        Dim dr As SqlDataReader

        For iRow = 0 To gridPagos2.Rows.Count - 1

            WRow = iRow
            If Not IsNothing(gridPagos2.Rows(iRow).Cells(4).Value) AndAlso Val(gridPagos2.Rows(iRow).Cells(4).Value) <> 0 Then

                renglon += 1

                XRecibo = txtRecibo.Text
                XRenglon = ceros(renglon, 2)
                XCliente = txtCliente.Text
                XFecha = txtFecha.Text
                XFechaOrd = String.Join("", XFecha.Split("/").Reverse())
                XTipoRec = "1"
                XRetganancias = Str$(Val(txtRetGanancias.Text))
                XRetIva = Str$(Val(txtRetIva.Text))
                XRetotra = Str$(Val(txtRetIB.Text))
                XRetencion = ""
                XTiporeg = "1"
                XTipo1 = gridPagos2.Rows(iRow).Cells(0).Value
                XLetra1 = gridPagos2.Rows(iRow).Cells(1).Value
                XPunto1 = gridPagos2.Rows(iRow).Cells(2).Value
                XNumero1 = ceros(gridPagos2.Rows(iRow).Cells(3).Value, 8)
                If Val(_Provincia) = 24 Then
                    XImporte1 = Str$(Val(gridPagos2.Rows(iRow).Cells(4).Value) * Val(txtParidad.Text))
                Else
                    XImporte1 = Str$(Val(gridPagos2.Rows(iRow).Cells(4).Value))
                End If
                XImporteBaja = Str$(Val(gridPagos2.Rows(iRow).Cells(4).Value))
                XTipo2 = ""
                XNumero2 = ""
                XFecha2 = ""
                XFechaOrd2 = ""
                XBanco2 = ""
                XImporte2 = ""
                XEstado2 = ""
                XObservaciones = txtObservaciones.Text
                XEmpresa = "1"
                XClave = XRecibo + XRenglon
                XImporte = Str$(Val(lblTotalCreditos.Text))
                XCuenta = ""
                XDestino = ""
                XImpolist = ""
                XImpo1list = ""
                XMarca = ""
                XFechaDepo = ""
                XFechaDepoOrd = ""

                XClaveCheque = ""
                XBancoCheque = ""
                XSucursalCheque = ""
                XChequeCheque = ""
                XCuentaCheque = ""
                XCuit = ""

                XParam = "'" + XClave + "','" _
                        + XRecibo + "','" + XRenglon + "','" _
                        + XCliente + "','" _
                        + XFecha + "','" + XFechaOrd + "','" _
                        + XTipoRec + "','" _
                        + XRetganancias + "','" _
                        + XRetIva + "','" + XRetotra + "','" _
                        + XRetencion + "','" _
                        + XTiporeg + "','" _
                        + XTipo1 + "','" + XLetra1 + "','" _
                        + XPunto1 + "','" + XNumero1 + "','" _
                        + XImporte1 + "','" _
                        + XTipo2 + "','" + XNumero2 + "','" _
                        + XFecha2 + "','" + XBanco2 + "','" _
                        + XImporte2 + "','" + XEstado2 + "','" _
                        + XEmpresa + "','" _
                        + XFechaOrd2 + "','" _
                        + XImporte + "','" _
                        + XObservaciones + "','" _
                        + XImpolist + "','" + XImpo1list + "','" _
                        + XDestino + "','" _
                        + XCuenta + "','" _
                        + XMarca + "','" _
                        + XFechaDepo + "','" _
                        + XFechaDepoOrd + "','" _
                        + XClaveCheque + "','" _
                        + XCuit + "','" _
                        + txtProvi.Text + "','" _
                        + XBancoCheque + "','" _
                        + XSucursalCheque + "','" _
                        + XChequeCheque + "','" _
                        + XCuentaCheque + "','" _
                        + WMarcaDiferencia + "'"


                ' Damos de alta el nuevo renglon del nuevo recibo.
                XSql = "INSERT INTO  Recibos (Clave, Recibo, Renglon, Cliente, Fecha, Fechaord, TipoRec, RetGanancias, RetIva, RetOtra, Retencion," _
                & "TipoReg, Tipo1, Letra1, Punto1, Numero1, Importe1, Tipo2, Numero2, Fecha2, banco2, Importe2," _
                & "Estado2, Empresa, FechaOrd2, Importe, Observaciones, Impolist, Impo1list, Destino, Cuenta, Marca," _
                & "FechaDepo, FechaDepoOrd, ClaveCheque, Cuit, Provisorio, BancoCheque, SucursalCheque, ChequeCheque, CuentaCheque, MarcaDiferencia) " _
                & "VALUES(" & XParam & ")"

                cm.CommandText = XSql
                SQLConnector.conexionSql(cn, cm)

                Try
                    cm.ExecuteNonQuery()

                Catch ex As Exception
                    MsgBox("Hubo un problema al querer guardar el nuevo recibo.", MsgBoxStyle.Critical)
                    Exit Sub
                Finally


                    cn.Close()


                    XSql = ""
                    XParam = ""

                End Try

                ' Calculamos el saldo para poder actualizar la cuenta cte.
                ClaveCtaCte = XTipo1 + XNumero1 + "01"
                cm.CommandText = "SELECT Saldo, Total, TotalUs, Estado FROM CtaCte Where Clave = '" & ClaveCtaCte & "' and Cliente = '" & Trim(txtCliente.Text) & "'"
                SQLConnector.conexionSql(cn, cm)

                Try

                    dr = cm.ExecuteReader()

                    If dr.HasRows Then

                        With dr
                            .Read()

                            XSaldo = Val(_NormalizarNumero(.Item("Saldo"))) - Val(XImporteBaja)

                            If Val(.Item("TotalUs")) <> 0 Then
                                XSaldoUs = Val(_NormalizarNumero(XSaldo)) / (Val(Val(_NormalizarNumero(.Item("Total"))) / Val(_NormalizarNumero(.Item("TotalUs")))))
                                'XSaldoUs = Val(_NormalizarNumero(.Item("Total"))) / Val(_NormalizarNumero(.Item("TotalUs")))
                            Else
                                XSaldoUs = ""
                            End If

                            XDate = Date.Now.ToString("MM-dd-yyyy")

                            XEstado = .Item("Estado")
                            If Val(XSaldo) = 0 And Val(XSaldoUs) = 0 Then
                                XEstado = "1"
                            End If

                        End With

                    Else
                        MsgBox("Error al querer actualizar el saldo en la Cta Cte asociada al recibo", MsgBoxStyle.Information)
                        Exit Sub
                    End If

                Catch ex As Exception
                    MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
                Finally


                    cn.Close()


                    XSql = ""
                    XParam = ""

                End Try

                'If Val(XSaldo) <> 0 Then

                XSql = "UPDATE CtaCte " _
                    & "SET " _
                    & " Saldo = " & XSaldo.Replace(",", ".") & "," _
                    & " SaldoUs = " & XSaldoUs.Replace(",", ".") & ", " _
                    & " Estado = '" & XEstado & "', " _
                    & " Wdate = '" & XDate & "'" _
                    & " WHERE Clave = '" & ClaveCtaCte & "' and Cliente = '" & Trim(txtCliente.Text) & "'"

                cm.CommandText = XSql
                SQLConnector.conexionSql(cn, cm)

                Try
                    cm.ExecuteNonQuery()

                Catch ex As Exception
                    MsgBox("Hubo un problema al querer actualizar la cuenta corriente.", MsgBoxStyle.Critical)
                    Exit Sub
                Finally


                    cn.Close()


                    XSql = ""
                    XParam = ""

                End Try

                'End If
            End If
        Next

    End Sub

    Private Sub _AltaVarios(ByVal WMarcaDiferencia As String)

        renglon += 1

        XRecibo = txtRecibo.Text
        XRenglon = ceros(renglon, 2)
        XCliente = txtCliente.Text
        XFecha = txtFecha.Text
        XFechaOrd = String.Join("", txtFecha.Text.Split("/").Reverse())
        XTipoRec = "3"
        XRetganancias = txtRetGanancias.Text
        XRetIva = txtRetIva.Text
        XRetotra = txtRetIB.Text
        XRetencion = ""
        XTiporeg = "1"
        XTipo1 = "99"
        XLetra1 = ""
        XPunto1 = ""
        XNumero1 = txtRecibo.Text
        XImporte1 = Str$(Val(lblTotalCreditos.Text))
        XTipo2 = ""
        XNumero2 = ""
        XFecha2 = ""
        XFechaOrd2 = ""
        XBanco2 = ""
        XImporte2 = ""
        XEstado2 = ""
        XObservaciones = txtObservaciones.Text
        XEmpresa = "1"
        XClave = XRecibo + XRenglon
        XImporte = Str$(Val(XImporte1))
        XCuenta = txtCuenta.Text
        XMarca = ""
        XDestino = ""
        XImpolist = ""
        XImpo1list = ""
        XFechaDepo = ""
        XFechaDepoOrd = ""

        XClaveCheque = ""
        XBancoCheque = ""
        XSucursalCheque = ""
        XChequeCheque = ""
        XCuentaCheque = ""
        XCuit = ""

        XParam = "'" + XClave + "','" _
                        + XRecibo + "','" + XRenglon + "','" _
                        + XCliente + "','" _
                        + XFecha + "','" + XFechaOrd + "','" _
                        + XTipoRec + "','" _
                        + XRetganancias + "','" _
                        + XRetIva + "','" + XRetotra + "','" _
                        + XRetencion + "','" _
                        + XTiporeg + "','" _
                        + XTipo1 + "','" + XLetra1 + "','" _
                        + XPunto1 + "','" + XNumero1 + "','" _
                        + XImporte1 + "','" _
                        + XTipo2 + "','" + XNumero2 + "','" _
                        + XFecha2 + "','" + XBanco2 + "','" _
                        + XImporte2 + "','" + XEstado2 + "','" _
                        + XEmpresa + "','" _
                        + XFechaOrd2 + "','" _
                        + XImporte + "','" _
                        + XObservaciones + "','" _
                        + XImpolist + "','" + XImpo1list + "','" _
                        + XDestino + "','" _
                        + XCuenta + "','" _
                        + XMarca + "','" _
                        + XFechaDepo + "','" _
                        + XFechaDepoOrd + "','" _
                        + XClaveCheque + "','" _
                        + XCuit + "','" _
                        + txtProvi.Text + "','" _
                        + XBancoCheque + "','" _
                        + XSucursalCheque + "','" _
                        + XChequeCheque + "','" _
                        + XCuentaCheque + "','" _
                        + WMarcaDiferencia + "'"

        XSql = "INSERT INTO  Recibos (Clave, Recibo, Renglon, Cliente, Fecha, Fechaord, TipoRec, RetGanancias, RetIva, RetOtra, Retencion," _
            & "TipoReg, Tipo1, Letra1, Punto1, Numero1, Importe1, Tipo2, Numero2, Fecha2, banco2, Importe2," _
            & "Estado2, Empresa, FechaOrd2, Importe, Observaciones, Impolist, Impo1list, Destino, Cuenta, Marca,FechaDepo, FechaDepoOrd, ClaveCheque, Cuit, Provisorio, BancoCheque, SucursalCheque, ChequeCheque, CuentaCheque, MarcaDiferencia)" _
            & " VALUES(" & XParam & ")"

        cm.CommandText = XSql
        SQLConnector.conexionSql(cn, cm)

        Try
            cm.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Hubo un problema al querer guardar el nuevo recibo.", MsgBoxStyle.Critical)
            Exit Sub
        Finally


            cn.Close()


            XSql = ""
            XParam = ""

        End Try
    End Sub

    Private Sub _AltaAnticipo(ByVal WMarcaDiferencia As String)

        renglon += 1

        ' Comienza la rutina para dar de alta la cuenta corriente
        XTipo = "07"
        XNumero = "00" + txtRecibo.Text
        ClaveCtaCte = XTipo + XNumero + "01"
        XRenglon = ceros(renglon, 2)
        XCliente = txtCliente.Text
        XFecha = txtFecha.Text
        XEstado = "1"
        XVencimiento = XFecha
        XVencimiento1 = XFecha

        If Val(_Provincia) = 24 Then
            XTotal = Str$(Val(lblTotalCreditos.Text) * -1 / Val(txtParidad.Text))
            XSaldo = Str$(Val(lblTotalCreditos.Text) * -1 / Val(txtParidad.Text))
        Else
            XTotal = _NormalizarNumero(Val(lblTotalCreditos.Text) * -1)
            XSaldo = _NormalizarNumero(Val(lblTotalCreditos.Text) * -1)
        End If

        XTotalUs = Str$(Val(lblTotalCreditos.Text) * -1 / Val(txtParidad.Text))
        XSaldoUs = Str$(Val(lblTotalCreditos.Text) * -1 / Val(txtParidad.Text))

        XOrdFecha = String.Join("", txtFecha.Text.Split("/").Reverse())
        XOrdVencimiento = XOrdFecha
        XOrdVencimiento1 = XOrdFecha

        XImpre = "AN"
        XEmpresa = "1"
        XNeto = ""
        XIva1 = ""
        XIva2 = ""
        XImpoIb = ""
        XSeguro = ""
        XFlete = ""
        XPedido = ""
        XRemito = ""
        XOrden = ""
        XParidad = txtParidad.Text
        XProvincia = _Provincia
        XVendedor = 0
        XRubro = 0
        XComprobante = ""
        XAceptada = ""
        XCosto = ""
        XImporte1 = ""
        XImporte2 = ""
        XImporte3 = ""
        XImporte4 = ""
        XImporte5 = ""
        XImporte6 = ""
        XImporte7 = ""
        XClave = XTipo + ceros(txtRecibo.Text, 8) + "01"
        XDate = Date.Now.ToString("MM-dd-yyyy")

        XParam = "'" + XClave + "','" _
                + XTipo + "','" + XNumero + "','" _
                + XRenglon + "','" + XCliente + "','" _
                + XFecha + "','" + XEstado + "','" _
                + XVencimiento + "','" + XVencimiento1 + "','" _
                + XTotal + "','" + XTotalUs + "','" _
                + XSaldo + "','" + XSaldoUs + "','" _
                + XOrdFecha + "','" + XOrdVencimiento + "','" _
                + XOrdVencimiento1 + "','" + XImpre + "','" _
                + XEmpresa + "','" _
                + XNeto + "','" + XIva1 + "','" _
                + XIva2 + "','" + XPedido + "','" _
                + XRemito + "','" + XOrden + "','" _
                + XParidad + "','" + XProvincia + "','" _
                + XVendedor + "','" + XRubro + "','" _
                + XComprobante + "','" + XAceptada + "','" _
                + XCosto + "','" _
                + XImporte1 + "','" + XImporte2 + "','" _
                + XImporte3 + "','" + XImporte4 + "','" _
                + XImporte5 + "','" + XImporte6 + "','" _
                + XImporte7 + "','" + XFlete + "','" _
                + XSeguro + "','" + XFlete + "','" _
                + XImpoIb + "'"

        XSql = "INSERT INTO  Ctacte (Clave, Tipo, numero, Renglon, Cliente, fecha, Estado, Vencimiento, Vencimiento1," _
        & "Total, TotalUs, Saldo, SaldoUs, Ordfecha, Ordvencimiento, Ordvencimiento1, Impre, Empresa, Neto, Iva1," _
        & "Iva2, Pedido, Remito, Orden, Paridad, Provincia, Vendedor, Rubro, Comprobante, Aceptada, Costo, Importe1," _
        & "Importe2, Importe3, Importe4, Importe5, Importe6, Importe7, WDate, Seguro, Flete, ImpoIb) " _
        & "VALUES(" & XParam & ")"

        cm.CommandText = XSql
        SQLConnector.conexionSql(cn, cm)

        Try
            cm.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Hubo un problema al querer guardar la nueva Cuenta Corriente.", MsgBoxStyle.Critical)
            Exit Sub
        Finally


            cn.Close()


            XSql = ""
            XParam = ""

        End Try

        ' Comienza rutina para dar de alta al recibo.

        XFechaOrd = XOrdFecha
        XTipoRec = "2"
        XRetganancias = Str$(Val(txtRetGanancias.Text))
        XRetIva = Str$(Val(txtRetIva.Text))
        XRetotra = Str$(Val(txtRetIB.Text))
        XRetencion = ""
        XTiporeg = "1"
        XTipo1 = "07"
        XLetra1 = ""
        XPunto1 = ""
        XNumero1 = txtRecibo.Text
        XImporte1 = Str$(Val(lblTotalCreditos.Text))
        XTipo2 = ""
        XNumero2 = ""
        XFecha2 = ""
        XFechaOrd2 = ""
        XBanco2 = ""
        XImporte2 = ""
        XEstado2 = ""
        XObservaciones = txtObservaciones.Text
        XEmpresa = "1"
        XClave = txtRecibo.Text + XRenglon
        XImporte = Str$(Val(XImporte1))
        XCuenta = ""
        XMarca = ""
        XFechaDepo = ""
        XFechaDepoOrd = ""
        XImpolist = ""
        XImpo1list = ""
        XDestino = ""

        XParam = "'" + XClave + "','" _
                        + txtRecibo.Text + "','" + XRenglon + "','" _
                        + XCliente + "','" _
                        + XFecha + "','" + XFechaOrd + "','" _
                        + XTipoRec + "','" _
                        + XRetganancias + "','" _
                        + XRetIva + "','" + XRetotra + "','" _
                        + XRetencion + "','" _
                        + XTiporeg + "','" _
                        + XTipo1 + "','" + XLetra1 + "','" _
                        + XPunto1 + "','" + XNumero1 + "','" _
                        + XImporte1 + "','" _
                        + XTipo2 + "','" + XNumero2 + "','" _
                        + XFecha2 + "','" + XBanco2 + "','" _
                        + XImporte2 + "','" + XEstado2 + "','" _
                        + XEmpresa + "','" _
                        + XFechaOrd2 + "','" _
                        + XImporte + "','" _
                        + XObservaciones + "','" _
                        + XImpolist + "','" + XImpo1list + "','" _
                        + XDestino + "','" _
                        + XCuenta + "','" _
                        + XMarca + "','" _
                        + XFechaDepo + "','" _
                        + XFechaDepoOrd + "','" _
                        + WMarcaDiferencia + "'"

        XSql = "INSERT INTO  Recibos (Clave, Recibo, Renglon, Cliente, Fecha, Fechaord, TipoRec, RetGanancias, RetIva, RetOtra, Retencion," _
            & "TipoReg, Tipo1, Letra1, Punto1, Numero1, Importe1, Tipo2, Numero2, Fecha2, banco2, Importe2," _
            & "Estado2, Empresa, FechaOrd2, Importe, Observaciones, Impolist, Impo1list, Destino, Cuenta, Marca," _
            & "FechaDepo, FechaDepoOrd, MarcaDiferencia) " _
        & "VALUES(" & XParam & ")"

        cm.CommandText = XSql
        SQLConnector.conexionSql(cn, cm)

        Try
            cm.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Hubo un problema al querer guardar el nuevo recibo.", MsgBoxStyle.Critical)
            Exit Sub
        Finally


            cn.Close()


            XSql = ""
            XParam = ""

        End Try


    End Sub

    Private Function _NoImputadosCorrectamenteValoresVarios() As Boolean
        Dim _Error = False

        For Each credito As DataGridViewRow In gridFormasPago2.Rows

            If Val(credito.Cells(0).Value) = 4 Then
                If Not _CuentaContableIngresada(credito.Index) Then
                    _Error = True
                    Exit For
                End If
            End If

        Next

        Return _Error
    End Function

    Private Function _CuentaContableIngresada(ByVal index As Integer)
        Dim _Ingresada = False
        Dim _CuentaContable As Object = _CuentasContables.FindLast(Function(c) c(0) = index)

        If Not IsNothing(_CuentaContable) Then
            If _CuentaContable(1) <> "" Then
                _Ingresada = True
            End If
        End If

        Return _Ingresada
    End Function

    Private Function _ExisteDiferenciaDeCambio() As Boolean
        Return Val(lblDolares.Text) < 0
    End Function

    Private Function _ErrorEnSumaDeSaldos() As Boolean
        Dim _Error = False

        If _SumarCreditos() Or _SumarDebitos() Then
            _Error = True
        End If

        Return _Error
    End Function

    Private Function _NumeroDeReciboDisponible() As Boolean
        Dim disponible = True

        cm.CommandText = "SELECT Recibo FROM Recibos WHERE Recibo = '" & Trim(txtRecibo.Text) & "01" & "'"
        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then
                disponible = False
            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally


            cn.Close()



        End Try
        Return disponible
    End Function

    Private Function _NumeracionDeReciboSiguiente()
        Dim ultimo = 0

        cm.CommandText = "SELECT Recibo FROM Recibos WHERE Recibo < 600000 ORDER BY Recibo DESC"
        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then

                dr.Read()

                ultimo = Val(dr.Item("Recibo"))

            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally


            cn.Close()



        End Try

        Return ultimo + 1
    End Function

    Private Function crearFormasPago() As List(Of FormaPago)
        Dim formasPago As New List(Of FormaPago)
        For Each row As DataGridViewRow In gridFormasPago2.Rows
            If Not row.IsNewRow Then
                formasPago.Add(New FormaPago(row.Cells(0).Value, 0, asString(row.Cells(1).Value), asString(row.Cells(2).Value), asString(row.Cells(3).Value), CustomConvert.toDoubleOrZero(row.Cells(4).Value)))
            End If
        Next
        Return formasPago
    End Function

    Private Function crearPagos() As List(Of Pago)
        Dim pagos As New List(Of Pago)
        For Each row As DataGridViewRow In gridPagos2.Rows
            If Not row.IsNewRow Then
                pagos.Add(New Pago(row.Cells(0).Value, asString(row.Cells(1).Value), asString(row.Cells(2).Value), asString(row.Cells(3).Value), "", CustomConvert.toDoubleOrZero(row.Cells(4).Value)))
            End If
        Next
        Return pagos
    End Function

    Private Function asString(ByVal value)
        If IsNothing(value) Then
            Return ""
        Else
            Return value.ToString
        End If
    End Function

    Private Sub optCtaCte_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCtaCte.CheckedChanged
        If Not optCtaCte.Checked Then
            gridPagos2.Rows.Clear()
            gridPagos2.AllowUserToAddRows = False
            'sumarValores()
        Else
            gridPagos2.AllowUserToAddRows = True
        End If
    End Sub

    Private Sub optVarios_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optVarios.CheckedChanged
        If Not optVarios.Checked Then
            txtCuenta.Enabled = False
            txtCuenta.Empty = True
            txtNombreCuenta.Empty = True
        Else
            txtCuenta.Enabled = True
            txtCuenta.Empty = False
            txtNombreCuenta.Empty = False
        End If
    End Sub

    Private Sub txtRetGanancias_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRetGanancias.KeyDown
        If e.KeyData = Keys.Enter Then

            If Val(txtRetGanancias.Text) > 0 Then
                _PedirInformacion("Ingrese el Número de Comprobante de Retención de Ganancias", txtRetGanancias, _ComprobanteRetGanancias)
            End If

            _SaltarA(txtRetIB)
        ElseIf e.KeyData = Keys.Escape Then
            txtRetGanancias.Text = 0
        End If
    End Sub

    Private Sub _PedirInformacion(ByVal msg As String, ByRef control As TextBox, ByRef VariableDestino As String)

        With InformacionRetenciones

            .TipoInformacion = msg

            .Valor = VariableDestino

            .ShowDialog(Me)

            VariableDestino = .Valor

            .Dispose()

        End With

    End Sub

    Private Sub _SaltarA(ByRef control As Control)
        control.Focus()
    End Sub

    Private Sub txtRetIva_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRetIva.KeyDown
        If e.KeyData = Keys.Enter Then

            If Val(txtRetIva.Text) > 0 Then
                _PedirInformacion("Ingrese el Número de Comprobante de Retención de IVA", txtRetIva, _ComprobanteRetIva)
            End If

            _SaltarA(txtRetSuss)
        End If
    End Sub

    Private Sub txtRetSuss_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRetSuss.KeyDown
        If e.KeyData = Keys.Enter Then

            If Val(txtRetSuss.Text) > 0 Then
                _PedirInformacion("Ingrese el Número de Comprobante de Retención de SUSS", txtRetSuss, _ComprobanteRetSuss)
            End If

            _SaltarA(txtObservaciones)
        ElseIf e.KeyData = Keys.Escape Then
            txtRetSuss.Text = 0
        End If
    End Sub

    Private Sub txtRetIB_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRetIB.KeyDown
        If e.KeyData = Keys.Enter Then

            _SaltarA(txtObservaciones)

        ElseIf e.KeyData = Keys.Escape Then
            txtRetIB.Text = 0
            _RetIB1 = ""
            _CompIB1 = ""
            _RetIB2 = ""
            _CompIB2 = ""
            _RetIB3 = ""
            _CompIB3 = ""
            _RetIB4 = ""
            _CompIB4 = ""
            _RetIB5 = ""
            _CompIB5 = ""
            _RetIB6 = ""
            _CompIB6 = ""
            _RetIB7 = ""
            _CompIB7 = ""
            _RetIB8 = ""
            _CompIB8 = ""
            _RetIB9 = ""
            _CompIB9 = ""
            _RetIB10 = ""
            _CompIB10 = ""
            _RetIB11 = ""
            _CompIB11 = ""
            _RetIB12 = ""
            _CompIB12 = ""
            _RetIB13 = ""
            _CompIB13 = ""
            _RetIB14 = ""
            _CompIB14 = ""
        End If
    End Sub

    Private Sub txtRetIB_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRetIB.Enter
        With DatosIB

            .txtRetIB1.Text = _RetIB1
            .txtCompIB1.Text = _CompIB1
            .txtRetIB2.Text = _RetIB2
            .txtCompIB2.Text = _CompIB2
            .txtRetIB3.Text = _RetIB3
            .txtCompIB3.Text = _CompIB3
            .txtRetIB4.Text = _RetIB4
            .txtCompIB4.Text = _CompIB4
            .txtRetIB5.Text = _RetIB5
            .txtCompIB5.Text = _CompIB5
            .txtRetIB6.Text = _RetIB6
            .txtCompIB6.Text = _CompIB6
            .txtRetIB7.Text = _RetIB7
            .txtCompIB7.Text = _CompIB7
            .txtRetIB8.Text = _RetIB8
            .txtCompIB8.Text = _CompIB8
            .txtRetIB9.Text = _RetIB9
            .txtCompIB9.Text = _CompIB9
            .txtRetIB10.Text = _RetIB10
            .txtCompIB10.Text = _CompIB10
            .txtRetIB11.Text = _RetIB11
            .txtCompIB11.Text = _CompIB11
            .txtRetIB12.Text = _RetIB12
            .txtCompIB12.Text = _CompIB12
            .txtRetIB13.Text = _RetIB13
            .txtCompIB13.Text = _CompIB13
            .txtRetIB14.Text = _RetIB14
            .txtCompIB14.Text = _CompIB14

            .ShowDialog(Me)

            _RetIB1 = .txtRetIB1.Text
            _CompIB1 = .txtCompIB1.Text
            _RetIB2 = .txtRetIB2.Text
            _CompIB2 = .txtCompIB2.Text
            _RetIB3 = .txtRetIB3.Text
            _CompIB3 = .txtCompIB3.Text
            _RetIB4 = .txtRetIB4.Text
            _CompIB4 = .txtCompIB4.Text
            _RetIB5 = .txtRetIB5.Text
            _CompIB5 = .txtCompIB5.Text
            _RetIB6 = .txtRetIB6.Text
            _CompIB6 = .txtCompIB6.Text
            _RetIB7 = .txtRetIB7.Text
            _CompIB7 = .txtCompIB7.Text
            _RetIB8 = .txtRetIB8.Text
            _CompIB8 = .txtCompIB8.Text
            _RetIB9 = .txtRetIB9.Text
            _CompIB9 = .txtCompIB9.Text
            _RetIB10 = .txtRetIB10.Text
            _CompIB10 = .txtCompIB10.Text
            _RetIB11 = .txtRetIB11.Text
            _CompIB11 = .txtCompIB11.Text
            _RetIB12 = .txtRetIB12.Text
            _CompIB12 = .txtCompIB12.Text
            _RetIB13 = .txtRetIB13.Text
            _CompIB13 = .txtCompIB13.Text
            _RetIB14 = .txtRetIB14.Text
            _CompIB14 = .txtCompIB14.Text

            .Dispose()

        End With

        _SaltarA(txtRetIva)

        _RecalcularRetIB()
    End Sub

    Private Sub _RecalcularRetIB()
        Dim totalIB As Double = 0

        totalIB += Val(_NormalizarNumero(_RetIB1))
        totalIB += Val(_NormalizarNumero(_RetIB2))
        totalIB += Val(_NormalizarNumero(_RetIB3))
        totalIB += Val(_NormalizarNumero(_RetIB4))
        totalIB += Val(_NormalizarNumero(_RetIB5))
        totalIB += Val(_NormalizarNumero(_RetIB6))
        totalIB += Val(_NormalizarNumero(_RetIB7))
        totalIB += Val(_NormalizarNumero(_RetIB8))
        totalIB += Val(_NormalizarNumero(_RetIB9))
        totalIB += Val(_NormalizarNumero(_RetIB10))
        totalIB += Val(_NormalizarNumero(_RetIB11))
        totalIB += Val(_NormalizarNumero(_RetIB12))
        totalIB += Val(_NormalizarNumero(_RetIB13))
        totalIB += Val(_NormalizarNumero(_RetIB14))

        txtRetIB.Text = totalIB
    End Sub

    Private Sub txtProvi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProvi.KeyDown

        If e.KeyData = Keys.Enter Then
            txtProvi.Text = ceros(txtProvi.Text, 6)
            mostrarReciboProvisorio(txtProvi.Text)

        ElseIf e.KeyData = Keys.Escape Then
            txtProvi.Text = ""
        End If

    End Sub

    Private Sub txtRetGanancias_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRetGanancias.Leave

        If txtRetGanancias.Text = "" Then
            txtRetGanancias.Text = 0
        End If

        _SumarCreditos()
    End Sub

    Private Sub txtRetIB_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRetIB.Leave
        If txtRetIB.Text = "" Then
            txtRetIB.Text = 0
        End If

        _SumarCreditos()
    End Sub

    Private Sub txtRetIva_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRetIva.Leave
        If txtRetIva.Text = "" Then
            txtRetIva.Text = 0
        End If

        _SumarCreditos()
    End Sub

    Private Sub txtRetSuss_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRetSuss.Leave
        If txtRetSuss.Text = "" Then
            txtRetSuss.Text = 0
        End If

        _SumarCreditos()
    End Sub

    Private Sub txtRecibo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRecibo.KeyDown

        If e.KeyData = Keys.Enter Then

            txtRecibo.Text = ceros(txtRecibo.Text, 6)

            mostrarRecibo(DAORecibo.buscarRecibo(txtRecibo.Text))
            _SaltarA(txtFecha)

        ElseIf e.KeyData = Keys.Escape Then
            txtRecibo.Text = "0"
        End If

    End Sub

    Private Sub txtFecha_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFecha.KeyDown
        If e.KeyData = Keys.Enter Then

            If Proceso._ValidarFecha(txtFecha.Text) Then

                _DeterminarParidad()

                If Trim(txtParidad.Text) = "" Then
                    MsgBox("No existe paridad cargada para la fecha indicada.")
                End If

                _SaltarA(txtProvi)

            End If

        ElseIf e.KeyData = Keys.Escape Then
            txtFecha.Clear()
        End If
    End Sub

    Private Sub txtCliente_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCliente.KeyDown
        If e.KeyData = Keys.Enter Then

            If Trim(txtCliente.Text) = "" Then
                lstSeleccion.SelectedIndex = 0
                lstSeleccion_Click(Nothing, Nothing)
                'btnConsulta.PerformClick()
                Exit Sub
            End If

            _SaltarA(txtRetGanancias)
        ElseIf e.KeyData = Keys.Escape Then
            txtCliente.Text = ""
        End If
    End Sub

    Private Sub txtParidad_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtParidad.KeyDown
        If e.KeyData = Keys.Enter Then

            If Val(txtParidad.Text) > 0 Then
                _SaltarA(txtObservaciones)
                _SumarDebitos()
            Else
                MsgBox("El valor de la Paridad debe ser mayor a 0", MsgBoxStyle.Information)
            End If

        ElseIf e.KeyData = Keys.Escape Then
            txtParidad.Text = ""
        End If
    End Sub

    Private Sub txtObservaciones_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtObservaciones.KeyDown
        If e.KeyValue = Keys.Enter Then
            With gridFormasPago2
                If .Rows.Count = 0 Then .Rows.Add()
                .CurrentCell = .Rows(0).Cells(0)
                .Focus()
            End With
        ElseIf e.KeyData = Keys.Escape Then
            txtObservaciones.Text = ""
        End If
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

    Private Function _EsDecimal(ByVal keycode As Integer) As Boolean
        Return (keycode >= 48 And keycode <= 57) Or (keycode >= 96 And keycode <= 105) Or (keycode = 110 Or keycode = 190 Or keycode = 109)
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

    Private Function _EsDecimalOControl(ByVal keycode) As Boolean
        Dim valido = False

        If _EsDecimal(CInt(keycode)) Or _EsControl(keycode) Then
            valido = True
        Else
            valido = False
        End If

        Return valido
    End Function

    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean

        If gridFormasPago2.Focused Or gridFormasPago2.IsCurrentCellInEditMode Then ' Detectamos los ENTER tanto si solo estan en foco o si estan en edición una celda.
            gridFormasPago2.CommitEdit(DataGridViewDataErrorContexts.Commit) ' Guardamos todos los datos que no hayan sido confirmados.

            Dim iCol = gridFormasPago2.CurrentCell.ColumnIndex
            Dim iRow = gridFormasPago2.CurrentCell.RowIndex

            ' Limitamos los caracteres permitidos para cada una de las columnas.
            Select Case iCol
                Case 1
                    If Not _EsNumeroOControl(keyData) Then
                        Return True
                    End If
                Case 4
                    If Not _EsDecimalOControl(keyData) Then
                        Return True
                    End If
            End Select

            'If iCol = 0 Then

            '    ' Por compatibilidad con lectora de Domingo que no envia enter al leer dato de cheque.
            '    If Len(gridFormasPago2.Rows(iRow).Cells(iCol).Value) = 30 Then
            '        SendKeys.Send("{ENTER}")
            '    End If

            'End If

            If msg.WParam.ToInt32() = Keys.Enter Then

                Dim valor = gridFormasPago2.Rows(iRow).Cells(iCol).Value

                If Not IsNothing(valor) Then

                    If iCol = 0 And iRow > -1 Then

                        If Trim(valor.ToString.Length) = 31 Then
                            If _ProcesarCheque(iRow, valor) Then
                                Dim _c As Object = _ClavesCheques.FindLast(Function(c) c(0) = iRow)

                                If Not IsNothing(_c) Then
                                    If Proceso.CuitValido(_c(6)) Then

                                        gridFormasPago2.Rows(iRow).Cells(5).Value = "1"

                                    Else

                                        gridFormasPago2.Rows(iRow).Cells(5).Value = "0"

                                    End If
                                Else
                                    gridFormasPago2.Rows(iRow).Cells(5).Value = "0"
                                End If

                                gridFormasPago2.CurrentCell = gridFormasPago2.Rows(iRow).Cells(2) ' Nos desplazamos para que coloque la fecha del cheque.

                            End If
                        Else
                            valor = valor.ToString().Substring(valor.ToString.Length - 1, 1)
                            If valor = "1" Or valor = "2" Or valor = "4" Then 'Or valor = "3" Or valor = "4" Then
                                eventoSegunTipoEnFormaDePagoPara(CustomConvert.toIntOrZero(valor), iRow, iCol)
                            Else ' Sólo se aceptan los valores 1 (Efectivo) , 2 (Cheque), 3 (Doc) y 4 (Varios) ?
                                gridFormasPago2.CurrentCell = gridFormasPago2.Rows(iRow).Cells(iCol)
                            End If
                        End If

                    Else

                        'Modificar el valor de la columna de ser fecha DD/MM --> DD/MM/YYYY (año actual)
                        If iCol = 2 Then

                            If _NormalizarFecha(valor) Then
                                gridFormasPago2.Rows(iRow).Cells(2).Value = valor
                                gridFormasPago2.EndEdit()
                            End If

                        End If

                        If iCol = 1 Or iCol = 2 Or iCol = 3 Then

                            If iCol = 1 Then
                                With gridFormasPago2
                                    .CurrentCell = .Rows(iRow).Cells(iCol + 1)

                                    Dim _location As Point = .GetCellDisplayRectangle(2, iRow, False).Location

                                    .ClearSelection()
                                    _location.Y += .Location.Y + (.CurrentCell.Size.Height / 4) - 1.5
                                    _location.X += .Location.X + (.CurrentCell.Size.Width - txtFechaAux.Size.Width) - 3
                                    txtFechaAux.Location = _location
                                    txtFechaAux.Text = .Rows(iRow).Cells(2).Value
                                    WRow = iRow
                                    Wcol = iCol
                                    txtFechaAux.Visible = True
                                    txtFechaAux.Focus()
                                End With

                                WRow = iRow
                                Wcol = iCol
                                txtFechaAux.Visible = True
                                txtFechaAux.Focus()
                            End If

                            If iCol = 2 Then
                                ' Completamos el año de manera automatica
                                If Len(Trim(valor)) = 6 Then
                                    Dim _mes As String = Mid(valor, 4, 2)

                                    Select Case Val(_mes)
                                        Case Is < 7
                                            valor = Mid(valor, 1, 2) & "/" & _mes & "/" & "2019"
                                        Case Else
                                            valor = Mid(valor, 1, 2) & "/" & _mes & "/" & "2018"
                                    End Select

                                End If

                                If _ChequeVencido(valor) Then
                                    MsgBox("La fecha del cheque introducida es inválida", MsgBoxStyle.Critical)
                                    gridFormasPago2.CurrentCell = gridFormasPago2.Rows(iRow).Cells(iCol)
                                Else
                                    gridFormasPago2.Rows(iRow).Cells(iCol).Value = valor
                                    gridFormasPago2.CurrentCell = gridFormasPago2.Rows(iRow).Cells(iCol + 1)
                                End If

                            ElseIf iCol = 3 Then
                                gridFormasPago2.Rows(iRow).Cells(iCol).Value = _GenerarCodigoBanco(valor)
                                gridFormasPago2.CurrentCell = gridFormasPago2.Rows(iRow).Cells(iCol + 1)
                            Else
                                gridFormasPago2.CurrentCell = gridFormasPago2.Rows(iRow).Cells(iCol + 1)
                            End If
                            Return True
                        End If

                        If iCol = 4 Then ' Avanzamos a la fila siguiente.

                            With gridFormasPago2.Rows(iRow)

                                If Not IsNothing(.Cells(0).Value) AndAlso Val(.Cells(0).Value) = 2 Then

                                    ' Controlamos que no se ingrese un numero = 0 para los cheques.
                                    If IsNothing(.Cells(1).Value) OrElse Val(.Cells(1).Value) = 0 Then
                                        MsgBox("Se debe informar un numero valido para el cheque.", MsgBoxStyle.Exclamation)
                                        Return True
                                    End If

                                End If

                            End With

                            If Val(gridFormasPago2.Rows(iRow).Cells(0).Value) = 4 Then

                                _PedirCuentaContable(iRow)

                            ElseIf Val(gridFormasPago2.Rows(iRow).Cells(0).Value) = 2 Then

                                If Val(gridFormasPago2.Rows(iRow).Cells(5).Value) <> 1 Then
                                    _PedirClaveCheque(iRow)
                                Else
                                    gridFormasPago2.Rows(iRow).Cells(5).Value = "0"
                                End If

                            End If



                            gridFormasPago2.Rows(iRow).Cells(4).Value = _NormalizarNumero(valor)

                            If gridFormasPago2.Rows.Count <= 40 Then

                                Try
                                    gridFormasPago2.CurrentCell = gridFormasPago2.Rows(iRow + 1).Cells(0)
                                Catch ex As Exception
                                    gridFormasPago2.Rows.Add()
                                    gridFormasPago2.CurrentCell = gridFormasPago2.Rows(iRow + 1).Cells(0)
                                End Try

                            Else
                                MsgBox("Se ha alcanzado el Número Máximo de Cheques/Formas de Pago permitidos en un Recibo.", MsgBoxStyle.Information)
                                gridFormasPago2.CurrentCell = gridFormasPago2.Rows(iRow).Cells(0)
                                gridFormasPago2.Focus()
                                Return True
                            End If

                            _SumarCreditos()
                            Return True
                        End If
                    End If

                    Return True
                End If
            ElseIf msg.WParam.ToInt32() = Keys.Escape Then
                With gridFormasPago2
                    .Rows(iRow).Cells(iCol).Value = ""

                    If iCol = 4 Then
                        .CurrentCell = .Rows(iRow).Cells(iCol - 1)
                    Else
                        .CurrentCell = .Rows(iRow).Cells(iCol + 1)
                    End If

                    .CurrentCell = .Rows(iRow).Cells(iCol)
                End With

                _SumarCreditos()
            End If
        End If
        
        If gridPagos2.Focused Or gridPagos2.IsCurrentCellInEditMode Then ' Detectamos los ENTER tanto si solo estan en foco o si estan en edición una celda.
            gridPagos2.CommitEdit(DataGridViewDataErrorContexts.Commit) ' Guardamos todos los datos que no hayan sido confirmados.

            Dim iCol = gridPagos2.CurrentCell.ColumnIndex
            Dim iRow = gridPagos2.CurrentCell.RowIndex

            If msg.WParam.ToInt32() = Keys.Enter Then

                Dim valor = gridPagos2.Rows(iRow).Cells(iCol).Value

                '_ComprobarDebitoPosible(iRow, iCol)
                If iCol = 4 Then
                    Dim max As Double = Val(gridPagos2.Rows(iRow).Cells(5).Value)

                    If max < 0 Then
                        If Val(gridPagos2.Rows(iRow).Cells(iCol).Value) < max Or Val(gridPagos2.Rows(iRow).Cells(iCol).Value) > 0 Then
                            MsgBox("El valor ingresado no se encuentra dentro del monto disponible en la Cuenta", MsgBoxStyle.Information)
                            gridPagos2.CurrentCell = gridPagos2.Rows(iRow).Cells(iCol)

                            Return True
                        Else
                            gridPagos2.Rows(iRow).Cells(iCol).Value = valor
                            _SumarDebitos()
                        End If
                    Else
                        If Val(gridPagos2.Rows(iRow).Cells(iCol).Value) > max Or Val(gridPagos2.Rows(iRow).Cells(iCol).Value) < 0 Then
                            MsgBox("El valor ingresado no se encuentra dentro del monto disponible en la Cuenta", MsgBoxStyle.Information)
                            gridPagos2.CurrentCell = gridPagos2.Rows(iRow).Cells(iCol)
                            Return True
                        Else
                            gridPagos2.Rows(iRow).Cells(iCol).Value = valor
                            _SumarDebitos()
                        End If
                    End If

                    Try
                        If Not IsNothing(gridPagos2.Rows(iRow).Cells(4).Value) AndAlso Val(gridPagos2.Rows(iRow).Cells(4).Value) <> 0 Then
                            gridPagos2.CurrentCell = gridPagos2.Rows(iRow + 1).Cells(0)
                        End If

                    Catch ex As Exception
                        If Not IsNothing(gridPagos2.Rows(iRow).Cells(4).Value) AndAlso Val(gridPagos2.Rows(iRow).Cells(4).Value) <> 0 Then
                            gridPagos2.Rows.Add()
                            gridPagos2.CurrentCell = gridPagos2.Rows(iRow + 1).Cells(0)
                        End If
                    End Try

                End If

            End If
            If msg.WParam.ToInt32() = Keys.Escape Then
                If iCol = 4 Then
                    With gridPagos2
                        .Rows(iRow).Cells(iCol).Value = ""

                        If iCol = 4 Then
                            .CurrentCell = .Rows(iRow).Cells(iCol - 1)
                        Else
                            .CurrentCell = .Rows(iRow).Cells(iCol + 1)
                        End If

                        .CurrentCell = .Rows(iRow).Cells(iCol)
                    End With
                    _SumarDebitos()
                End If

                Return True
            End If
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Private Function _GenerarCodigoBanco(ByVal _Banco As String) As String
        _Banco = _Banco.ToString.Split("/")(0) ' Agarramos el nombre del banco, sin el cod del cliente.

        Return _Banco & "/" & Mid(txtCliente.Text, 1, 1) & Val(Mid(txtCliente.Text, 2, 6)).ToString()
    End Function

    Private Function _ChequeYaCargado(ByVal ClaveCheque) As Boolean
        Dim _cargado = False

        If _ChequeUtilizadoEnRecibo(ClaveCheque) Then

            _cargado = True

        ElseIf _ChequeUtilizadoEnReciboProvisorio(ClaveCheque) Then

            _cargado = True

        Else

            Dim cheque As Object = _ClavesCheques.FindLast(Function(c) c(1) = ClaveCheque)
            If Not IsNothing(cheque) Then
                _cargado = True
            End If

        End If

        Return _cargado
    End Function

    Private Function _ChequeUtilizadoEnRecibo(ByVal ClaveCheque As String) As Boolean
        Dim utilizado = False
        cm.CommandText = "SELECT TOP 1 ClaveCheque FROM Recibos WHERE ClaveCheque = '" & ClaveCheque & "'"

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then
                utilizado = True
            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally

            cn.Close()

        End Try

        Return utilizado
    End Function

    Private Function _ChequeUtilizadoEnReciboProvisorio(ByVal ClaveCheque As String) As Boolean
        Dim utilizado = False
        cm.CommandText = "SELECT TOP 1 ClaveCheque FROM RecibosProvi WHERE ClaveCheque = '" & ClaveCheque & "'"

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then
                utilizado = True
            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally

            cn.Close()

        End Try

        Return utilizado
    End Function

    Private Function _FormatoValidoDeCheque(ByVal ClaveCheque As String) As Boolean
        Dim valido = False

        ClaveCheque = UCase(Trim(ClaveCheque))

        If ClaveCheque.Length = 31 And Mid(ClaveCheque, 1, 1) = "C" And Mid(ClaveCheque, ClaveCheque.Length, 1) = "E" Then
            valido = True
        End If

        Return valido
    End Function

    Private Function _ProcesarCheque(ByVal row As Integer, ByVal ClaveCheque As String) As Boolean
        Dim _ClaveBanco, _Banco, _Sucursal, _NumCheque, _NumCta, _Cuit As String
        Dim _LecturaCorrecta = True

        If Not _FormatoValidoDeCheque(ClaveCheque) Then
            MsgBox("El formato del cheque no es valido.", MsgBoxStyle.Exclamation)
            Return False
        End If

        _ClaveBanco = Mid$(ClaveCheque, 2, 3)
        _Banco = _ObtenerNombreBanco(_ClaveBanco)
        _Sucursal = Mid$(ClaveCheque, 5, 3)
        _NumCheque = Mid$(ClaveCheque, 12, 8)
        _NumCta = Mid$(ClaveCheque, 20, 11)

        ' Chequeamos que el cheque no se haya cargado.

        If _ChequeYaCargado(ClaveCheque) Then
            _LecturaCorrecta = False

            MsgBox("Cheque ya Cargado con anterioridad.", MsgBoxStyle.Exclamation)

            Return _LecturaCorrecta
        End If

        ' Extraer Datos del String.
        If _Banco = "" Then
            _LecturaCorrecta = False
            MsgBox("Error en la lectura de los datos del codigo de banco del cheque")
        End If

        With gridFormasPago2.Rows(row)
            .Cells(0).Value = "02"
            .Cells(1).Value = _NumCheque
            .Cells(2).Value = ""
            .Cells(3).Value = _Banco & "/" & Mid(txtCliente.Text, 1, 1) & Val(Mid(txtCliente.Text, 2, 6)).ToString()
        End With

        ' Buscamos si existe el cuit.
        _Cuit = _TraerNumeroCuit(_ClaveBanco & _Sucursal & _NumCta)

        ' Guardamos el nuevo Cheque.
        _GuardarNuevoCheque(row, ClaveCheque, _ClaveBanco, _Sucursal, _NumCheque, _NumCta, _Cuit)

        Return _LecturaCorrecta
    End Function

    Private Function _ObtenerNombreBanco(ByVal claveBanco As String) As String
        Dim _NombreBanco = ""
        cm.CommandText = "SELECT Nombre FROM BCRA WHERE Banco = '" & Val(claveBanco) & "'"

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then

                dr.Read()

                _NombreBanco = Trim(UCase(dr.Item("Nombre")))

            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally

            cn.Close()

        End Try

        Return _NombreBanco
    End Function

    Private Function _TraerNumeroCuit(ByVal clave As String) As String
        Dim _cuit = ""
        cm.CommandText = "SELECT Cuit FROM Cuit WHERE Clave = '" & Trim(clave) & "'"

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then
                dr.Read()

                _cuit = dr.Item("Cuit")

            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally

            cn.Close()

        End Try

        Return _cuit
    End Function


    Private Sub _GuardarNuevoCheque(ByVal row As Integer, ByVal Clave As String, _
                                    ByVal banco As String, ByVal sucursal As String, _
                                    ByVal numCheque As String, ByVal numCta As String, _
                                    ByVal _Cuit As String)

        Dim buscar As Object = _ClavesCheques.Find(Function(c) c(0) = row)

        If Not IsNothing(buscar) Then
            _ClavesCheques.Remove(buscar)
        End If

        _ClavesCheques.Add({row, Clave, banco, sucursal, numCheque, numCta, _Cuit, "", ""})

    End Sub

    Private Function _ChequeVencido(ByVal fecha_cheque As String) As Boolean
        Return IsDate(fecha_cheque) And IsDate(txtFecha.Text) And DateDiff(DateInterval.Day, CDate(fecha_cheque), CDate(txtFecha.Text)) > 30
    End Function

    Private Sub _PedirClaveCheque(ByVal row As Integer)
        Dim _clave, _banco, _sucursal, _cheque, _cuenta, _cuit, _estado, _destino As String
        Dim buscar As Object

        _cuit = ""
        buscar = _ClavesCheques.FindLast(Function(c) c(0) = row)

        If Not IsNothing(buscar) Then
            _cuit = Trim(buscar(6))
            'If _cuit.Length = 11 Then : Exit Sub : End If
        End If

        With SolicitarInformacionCuit

            .Valor = _cuit

            .ShowDialog()

            _cuit = .Valor

        End With

        _clave = ""
        _banco = ""
        _sucursal = ""
        _cheque = ""
        _cuenta = ""
        _estado = ""
        _destino = ""

        buscar = Nothing
        buscar = _ClavesCheques.Find(Function(c) c(0) = row)
        If Not IsNothing(buscar) Then

            _clave = buscar(1)
            _banco = buscar(2)
            _sucursal = buscar(3)
            _cheque = buscar(4)
            _cuenta = buscar(5)
            _estado = ""
            _destino = ""

            _ClavesCheques.Remove(buscar)
        End If

        _ClavesCheques.Add({row, _clave, _banco, _sucursal, _cheque, _cuenta, _cuit, _estado, _destino})

        gridFormasPago2.Rows(row).Cells("NroCuit").Value = _cuit

    End Sub

    Private Sub _PedirCuentaContable(ByVal row As Integer)
        Dim cuenta = ""
        Dim buscar As Object = Nothing

        buscar = _CuentasContables.FindLast(Function(c) c(0) = row)

        If Not IsNothing(buscar) Then

            cuenta = buscar(1)
            _PedirInformacion("Ingrese Cuenta Contable", New TextBox(), cuenta)

        End If

        Do While Not _CuentaContableValida(cuenta) And Trim(cuenta) <> "0"

            _PedirInformacion("Ingrese Cuenta Contable", New TextBox(), cuenta)

        Loop

        _CuentasContables.Add({row, Trim(cuenta)})

    End Sub

    Private Function _CuentaContableValida(ByRef cuenta As String)
        Dim _CuentaValida = False

        cm.CommandText = "SELECT Cuenta FROM Cuenta WHERE Cuenta= '" & Trim(cuenta) & "'"
        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then

                _CuentaValida = True

            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally


            cn.Close()



        End Try

        Return _CuentaValida
    End Function

    Private Sub _ComprobarDebitoPosible(ByVal iRow As Integer, ByVal iCol As Integer)
        Dim _Maximo As Object = _ValoresMax.FindLast(Function(m) m(0) = gridPagos2.Rows(iRow).Cells(3).Value)
        If Not IsNothing(_Maximo) Then
            If iCol = 4 And (IsNothing(gridPagos2.CurrentCell.Value) Or Val(gridPagos2.CurrentCell.Value) > _Maximo(1)) Then
                MsgBox("El valor del importe indicado no puede ser mayor a " & _Maximo(1))
                gridPagos2.CurrentCell = gridPagos2.Rows(iRow).Cells(iCol)
            Else
                gridPagos2.CurrentCell.Value = _NormalizarNumero(Val(gridPagos2.CurrentCell.Value))
                _SumarDebitos()
                gridPagos2.CurrentCell = gridPagos2.Rows(iRow + 1).Cells(iCol)
            End If
        End If
    End Sub

    Private Sub lstFiltrada_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstFiltrada.SelectedIndexChanged
        If Trim(lstFiltrada.SelectedItem) <> "" Then
            _TraerConsulta(lstFiltrada.SelectedItem)
        End If
    End Sub

    Private Sub txtCuenta_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCuenta.KeyDown

        If e.KeyData = Keys.Enter Then
            If Trim(txtCuenta.Text) <> "" Then
                mostrarCuenta(txtCuenta.Text)

                If Trim(txtNombreCuenta.Text) <> "" Then
                    _SaltarA(txtObservaciones)
                End If

            End If
        End If

    End Sub

    Private Sub txtFecha_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If Trim(txtFecha.Text).Length = 10 Then
            _DeterminarParidad()
        End If

    End Sub

    Private Sub btnImpresion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImpresion.Click

        ' Chequeamos que se haya cargado un recibo antes de mandar a imprimir.
        If Val(txtRecibo.Text) > 0 Then

            If Proceso._EsPellital Then
                _PrepararEImprimirReciboPellital()
            Else
                _PrepararEImprimirRecibo()
            End If

        End If

    End Sub

    Private Sub _PrepararEImprimirReciboPellital()
        ' Comprobamos que el Recibo exista.
        If Not _ReciboYaExistente(Trim(txtRecibo.Text)) Then
            MsgBox("El Nº de Recibo indicado no se corresponde con ninguno registrado en la Base de Datos." _
                   & vbCrLf & vbCrLf & _
                   "Se detendrá el proceso de impresión", MsgBoxStyle.Information)
            Exit Sub
        End If

        Dim Retencion, Cheque, Documento, Total2f, Pesos, Bonos, Dolares, Ajuste, Compe, Transfe, Total2, Total1 As Double
        Dim Vector(30, 10) As String
        Dim WEntra(100, 120) As String
        Dim ImpreTipo(100) As String
        Dim XLugar = 0
        Dim iRow = 0

        ' Inicializamos todas las variable a utilizar.
        If optVarios.Checked Then
            WRazon = Space$(30)
            WDireccion = Space$(30)
            WLocalidad = Space$(30)
            WProvincia = ""
            WPostal = ""
        End If

        WRecibo = Val(txtRecibo.Text)
        WFecha = txtFecha.Text
        WCliente = txtCliente.Text

        Retencion = Val(txtRetGanancias.Text) + Val(txtRetIva.Text) + Val(_NormalizarNumero(txtRetIB.Text)) + Val(txtRetSuss.Text)

        Cheque = 0
        Documento = 0
        Total2f = 0
        Pesos = 0
        Bonos = 0
        Dolares = 0
        Ajuste = 0
        Compe = 0
        Transfe = 0
        Total2 = 0

        XTipo = ""
        XNumero = ""

        ' Limpiamos los vectores para re poblarlos.
        Array.Clear(Vector, 0, Vector.Length)
        Array.Clear(WEntra, 0, WEntra.Length)

        ImpreTipo(1) = "FC"

        ' Extraemos los valores de Débitos y Créditos.
        For iRow = 0 To 19

            If iRow < gridPagos2.Rows.Count Then
                If Trim(gridPagos2.Rows(iRow).Cells(4).Value) <> "" Then
                    With gridPagos2.Rows(iRow)
                        Vector(iRow, 0) = .Cells(0).Value
                        Vector(iRow, 1) = .Cells(1).Value
                        Vector(iRow, 2) = .Cells(2).Value
                        Vector(iRow, 3) = .Cells(3).Value
                        Vector(iRow, 4) = _NormalizarNumero(.Cells(4).Value)
                    End With
                End If
            Else
                For i = 0 To 4
                    Vector(iRow, i) = ""
                Next
            End If

            If iRow < gridFormasPago2.Rows.Count Then
                If Trim(gridFormasPago2.Rows(iRow).Cells(4).Value) <> "" Then
                    With gridFormasPago2.Rows(iRow)
                        Vector(iRow, 5) = .Cells(0).Value
                        Vector(iRow, 6) = .Cells(1).Value
                        Vector(iRow, 7) = .Cells(2).Value
                        Vector(iRow, 8) = .Cells(3).Value
                        Vector(iRow, 9) = _NormalizarNumero(.Cells(4).Value)
                    End With
                End If
            Else
                For i = 5 To 9
                    Vector(iRow, i) = ""
                Next
            End If

            XTipo = ceros(Vector(iRow, 0), 2)
            XNumero = ceros(Vector(iRow, 3), 8)

            'Solo realizamos la llamada a la base de datos en caso de haber datos por los que solicitar la fecha.
            If Val(XTipo) > 0 And Val(XNumero) > 0 Then
                Vector(iRow, 10) = _ObtenerFechaCtaCte(XTipo & XNumero & "01")
            Else
                Vector(iRow, 10) = ""
            End If

        Next iRow

        ' Calculamos totales y subtotales en base a los tipos de créditos.
        For Ciclo = 0 To 19

            If Val(Vector(Ciclo, 9)) <> 0 Then
                Dim _cuenta As Object = _CuentasContables.FindLast(Function(c) c(0) = Ciclo)
                Select Case Val(Vector(Ciclo, 5))
                    Case 1
                        If IsNothing(_cuenta) Then
                            Pesos = Pesos + Val(Vector(Ciclo, 9))
                        Else
                            If Val(_cuenta(1)) <> 2 Then
                                Pesos = Pesos + Val(Vector(Ciclo, 9))
                            Else
                                Dolares = Dolares + Val(Vector(Ciclo, 9))
                            End If
                        End If
                    Case 4
                        If Not IsNothing(_cuenta) Then
                            Select Case Val(_cuenta(1))
                                Case 2
                                    Dolares = Dolares + Val(Vector(Ciclo, 9))
                                Case 5
                                    Compe = Compe + Val(Vector(Ciclo, 9))
                                Case 21, 22, 25
                                    Transfe = Transfe + Val(Vector(Ciclo, 9))
                                Case 91
                                    Ajuste = Ajuste + Val(Vector(Ciclo, 9))
                                Case 157, 7, 8
                                    Bonos = Bonos + Val(Vector(Ciclo, 9))
                                Case Else
                                    REM Pesos = Pesos + Val(Vector(Ciclo, 9))
                            End Select
                        End If
                    Case 2
                        Cheque = Cheque + Val(Vector(Ciclo, 9))
                    Case Else
                        Documento = Documento + Val(Vector(Ciclo, 9))
                End Select
            End If

            If Val(Vector(Ciclo, 4)) <> 0 Then
                Total2 = Total2 + Val(Vector(Ciclo, 4))
            End If

        Next Ciclo

        Total1 = Pesos + Cheque + Documento + Retencion + Dolares + Compe + Transfe + Ajuste + Bonos

        For Ciclo = 0 To 19
            If Val(Vector(Ciclo, 9)) <> 0 And (Val(Vector(Ciclo, 5)) = 2 Or Val(Vector(Ciclo, 5)) = 3) Then
                Vector(Ciclo, 6) = ceros(Vector(Ciclo, 6), 6)
                Vector(Ciclo, 8) = LSet(Vector(Ciclo, 8), 20)
                For Pasa = 1 To 24
                    Select Case Ciclo
                        Case 0
                            WEntra(Pasa, 17) = Vector(Ciclo, 6)
                            WEntra(Pasa, 18) = Vector(Ciclo, 7)
                            WEntra(Pasa, 19) = Vector(Ciclo, 9)
                            WEntra(Pasa, 20) = Vector(Ciclo, 8)
                        Case 1
                            WEntra(Pasa, 21) = Vector(Ciclo, 6)
                            WEntra(Pasa, 22) = Vector(Ciclo, 7)
                            WEntra(Pasa, 23) = Vector(Ciclo, 9)
                            WEntra(Pasa, 24) = Vector(Ciclo, 8)
                        Case 2
                            WEntra(Pasa, 25) = Vector(Ciclo, 6)
                            WEntra(Pasa, 26) = Vector(Ciclo, 7)
                            WEntra(Pasa, 27) = Vector(Ciclo, 9)
                            WEntra(Pasa, 28) = Vector(Ciclo, 8)
                        Case 3
                            WEntra(Pasa, 29) = Vector(Ciclo, 6)
                            WEntra(Pasa, 30) = Vector(Ciclo, 7)
                            WEntra(Pasa, 31) = Vector(Ciclo, 9)
                            WEntra(Pasa, 32) = Vector(Ciclo, 8)
                        Case 4
                            WEntra(Pasa, 33) = Vector(Ciclo, 6)
                            WEntra(Pasa, 34) = Vector(Ciclo, 7)
                            WEntra(Pasa, 35) = Vector(Ciclo, 9)
                            WEntra(Pasa, 36) = Vector(Ciclo, 8)
                        Case 5
                            WEntra(Pasa, 37) = Vector(Ciclo, 6)
                            WEntra(Pasa, 38) = Vector(Ciclo, 7)
                            WEntra(Pasa, 39) = Vector(Ciclo, 9)
                            WEntra(Pasa, 40) = Vector(Ciclo, 8)
                        Case 6
                            WEntra(Pasa, 41) = Vector(Ciclo, 6)
                            WEntra(Pasa, 42) = Vector(Ciclo, 7)
                            WEntra(Pasa, 43) = Vector(Ciclo, 9)
                            WEntra(Pasa, 44) = Vector(Ciclo, 8)
                        Case 7
                            WEntra(Pasa, 45) = Vector(Ciclo, 6)
                            WEntra(Pasa, 46) = Vector(Ciclo, 7)
                            WEntra(Pasa, 47) = Vector(Ciclo, 9)
                            WEntra(Pasa, 48) = Vector(Ciclo, 8)
                        Case 8
                            WEntra(Pasa, 49) = Vector(Ciclo, 6)
                            WEntra(Pasa, 50) = Vector(Ciclo, 7)
                            WEntra(Pasa, 51) = Vector(Ciclo, 9)
                            WEntra(Pasa, 52) = Vector(Ciclo, 8)
                        Case 9
                            WEntra(Pasa, 53) = Vector(Ciclo, 6)
                            WEntra(Pasa, 54) = Vector(Ciclo, 7)
                            WEntra(Pasa, 55) = Vector(Ciclo, 9)
                            WEntra(Pasa, 56) = Vector(Ciclo, 8)
                        Case 10
                            WEntra(Pasa, 57) = Vector(Ciclo, 6)
                            WEntra(Pasa, 58) = Vector(Ciclo, 7)
                            WEntra(Pasa, 59) = Vector(Ciclo, 9)
                            WEntra(Pasa, 60) = Vector(Ciclo, 8)
                        Case 11
                            WEntra(Pasa, 61) = Vector(Ciclo, 6)
                            WEntra(Pasa, 62) = Vector(Ciclo, 7)
                            WEntra(Pasa, 63) = Vector(Ciclo, 9)
                            WEntra(Pasa, 64) = Vector(Ciclo, 8)
                        Case 12
                            WEntra(Pasa, 65) = Vector(Ciclo, 6)
                            WEntra(Pasa, 66) = Vector(Ciclo, 7)
                            WEntra(Pasa, 67) = Vector(Ciclo, 9)
                            WEntra(Pasa, 68) = Vector(Ciclo, 8)
                        Case 13
                            WEntra(Pasa, 69) = Vector(Ciclo, 6)
                            WEntra(Pasa, 70) = Vector(Ciclo, 7)
                            WEntra(Pasa, 71) = Vector(Ciclo, 9)
                            WEntra(Pasa, 72) = Vector(Ciclo, 8)
                        Case 14
                            WEntra(Pasa, 73) = Vector(Ciclo, 6)
                            WEntra(Pasa, 74) = Vector(Ciclo, 7)
                            WEntra(Pasa, 75) = Vector(Ciclo, 9)
                            WEntra(Pasa, 76) = Vector(Ciclo, 8)
                        Case 15
                            WEntra(Pasa, 77) = Vector(Ciclo, 6)
                            WEntra(Pasa, 78) = Vector(Ciclo, 7)
                            WEntra(Pasa, 79) = Vector(Ciclo, 9)
                            WEntra(Pasa, 80) = Vector(Ciclo, 8)
                        Case 16
                            WEntra(Pasa, 81) = Vector(Ciclo, 6)
                            WEntra(Pasa, 82) = Vector(Ciclo, 7)
                            WEntra(Pasa, 83) = Vector(Ciclo, 9)
                            WEntra(Pasa, 84) = Vector(Ciclo, 8)
                        Case 17
                            WEntra(Pasa, 85) = Vector(Ciclo, 6)
                            WEntra(Pasa, 86) = Vector(Ciclo, 7)
                            WEntra(Pasa, 87) = Vector(Ciclo, 9)
                            WEntra(Pasa, 88) = Vector(Ciclo, 8)
                        Case 18
                            WEntra(Pasa, 89) = Vector(Ciclo, 6)
                            WEntra(Pasa, 90) = Vector(Ciclo, 7)
                            WEntra(Pasa, 91) = Vector(Ciclo, 9)
                            WEntra(Pasa, 92) = Vector(Ciclo, 8)
                        Case 19
                            WEntra(Pasa, 93) = Vector(Ciclo, 6)
                            WEntra(Pasa, 94) = Vector(Ciclo, 7)
                            WEntra(Pasa, 95) = Vector(Ciclo, 9)
                            WEntra(Pasa, 96) = Vector(Ciclo, 8)
                        Case Else
                    End Select
                Next Pasa
            End If
        Next Ciclo

        XLugar = 1

        WEntra(XLugar, 1) = txtRecibo.Text
        WEntra(XLugar, 2) = txtFecha.Text
        WEntra(XLugar, 3) = txtCliente.Text
        WEntra(XLugar, 4) = WRazon
        WEntra(XLugar, 5) = WDireccion
        WEntra(XLugar, 6) = WLocalidad
        WEntra(XLugar, 7) = WProvincia
        WEntra(XLugar, 8) = WPostal
        WEntra(XLugar, 9) = "Efectivo "
        WEntra(XLugar, 10) = Str$(Pesos)
        WEntra(XLugar, 11) = ""
        WEntra(XLugar, 12) = ""
        WEntra(XLugar, 13) = ""
        WEntra(XLugar, 14) = ""
        WEntra(XLugar, 15) = ""
        WEntra(XLugar, 16) = ""

        XLugar = XLugar + 1

        WEntra(XLugar, 1) = txtRecibo.Text
        WEntra(XLugar, 2) = txtFecha.Text
        WEntra(XLugar, 3) = txtCliente.Text
        WEntra(XLugar, 4) = WRazon
        WEntra(XLugar, 5) = WDireccion
        WEntra(XLugar, 6) = WLocalidad
        WEntra(XLugar, 7) = WProvincia
        WEntra(XLugar, 8) = WPostal
        WEntra(XLugar, 9) = ""
        WEntra(XLugar, 10) = ""
        WEntra(XLugar, 11) = ""
        WEntra(XLugar, 12) = ""
        WEntra(XLugar, 13) = ""
        WEntra(XLugar, 14) = ""
        WEntra(XLugar, 15) = ""
        WEntra(XLugar, 16) = ""

        XLugar = XLugar + 1

        WEntra(XLugar, 1) = txtRecibo.Text
        WEntra(XLugar, 2) = txtFecha.Text
        WEntra(XLugar, 3) = txtCliente.Text
        WEntra(XLugar, 4) = WRazon
        WEntra(XLugar, 5) = WDireccion
        WEntra(XLugar, 6) = WLocalidad
        WEntra(XLugar, 7) = WProvincia
        WEntra(XLugar, 8) = WPostal
        WEntra(XLugar, 9) = "Cheques "
        WEntra(XLugar, 10) = Str$(Cheque)
        WEntra(XLugar, 11) = ""
        WEntra(XLugar, 12) = ""
        WEntra(XLugar, 13) = ""
        WEntra(XLugar, 14) = ""
        WEntra(XLugar, 15) = ""
        WEntra(XLugar, 16) = ""

        XLugar = XLugar + 1

        WEntra(XLugar, 1) = txtRecibo.Text
        WEntra(XLugar, 2) = txtFecha.Text
        WEntra(XLugar, 3) = txtCliente.Text
        WEntra(XLugar, 4) = WRazon
        WEntra(XLugar, 5) = WDireccion
        WEntra(XLugar, 6) = WLocalidad
        WEntra(XLugar, 7) = WProvincia
        WEntra(XLugar, 8) = WPostal
        WEntra(XLugar, 9) = ""
        WEntra(XLugar, 10) = ""
        WEntra(XLugar, 11) = ""
        WEntra(XLugar, 12) = ""
        WEntra(XLugar, 13) = ""
        WEntra(XLugar, 14) = ""
        WEntra(XLugar, 15) = ""
        WEntra(XLugar, 16) = ""


        For Ciclo = 0 To 17

            XLugar = XLugar + 1

            WEntra(XLugar, 1) = txtRecibo.Text
            WEntra(XLugar, 2) = txtFecha.Text
            WEntra(XLugar, 3) = txtCliente.Text
            WEntra(XLugar, 4) = WRazon
            WEntra(XLugar, 5) = WDireccion
            WEntra(XLugar, 6) = WLocalidad
            WEntra(XLugar, 7) = WProvincia
            WEntra(XLugar, 8) = WPostal

            Select Case Ciclo
                Case 0
                    WEntra(XLugar, 9) = "Documentos "
                    WEntra(XLugar, 10) = Str$(Documento)
                    WEntra(XLugar, 11) = ""
                Case 2
                    WEntra(XLugar, 9) = "Retencion Ganancias "
                    WEntra(XLugar, 10) = txtRetGanancias.Text
                    WEntra(XLugar, 11) = ""
                Case 4
                    WEntra(XLugar, 9) = "Retencion Iva "
                    WEntra(XLugar, 10) = txtRetIva.Text
                    WEntra(XLugar, 11) = ""
                Case 6
                    WEntra(XLugar, 9) = "Retencion I.Brutos "
                    WEntra(XLugar, 10) = txtRetIB.Text
                    WEntra(XLugar, 11) = ""
                Case 8
                    WEntra(XLugar, 9) = "Moneda Ext."
                    If Val(txtParidad.Text) <> 0 Then
                        WEntra(XLugar, 10) = Str$(Val(_NormalizarNumero(Dolares)) / Val(_NormalizarNumero(txtParidad.Text)))
                    Else
                        WEntra(XLugar, 10) = Str$(Dolares)
                    End If
                    WEntra(XLugar, 11) = "U$S"
                Case 10
                    WEntra(XLugar, 9) = "Compensacion"
                    WEntra(XLugar, 10) = Str$(Compe)
                    WEntra(XLugar, 11) = ""
                Case 12
                    WEntra(XLugar, 9) = "Bonos"
                    WEntra(XLugar, 10) = Str$(Bonos)
                    WEntra(XLugar, 11) = ""
                Case 14
                    WEntra(XLugar, 9) = "Ajuste"
                    WEntra(XLugar, 10) = Str$(Ajuste)
                    WEntra(XLugar, 11) = ""
                Case 16
                    WEntra(XLugar, 9) = "Transferencia"
                    WEntra(XLugar, 10) = Str$(Transfe)
                    WEntra(XLugar, 11) = ""
                Case 17
                    WEntra(XLugar, 9) = "Ret. Suss"
                    WEntra(XLugar, 10) = txtRetSuss.Text
                    WEntra(XLugar, 11) = ""
                Case Else
                    WEntra(XLugar, 9) = ""
                    WEntra(XLugar, 10) = ""
                    WEntra(XLugar, 11) = ""
            End Select

            If Val(Vector(Ciclo, 4)) <> 0 And optCtaCte.Checked Then
                Call ceros(Vector(Ciclo, 3), 6)
                WEntra(XLugar, 12) = Vector(Ciclo, 10)
                WEntra(XLugar, 13) = ImpreTipo(Val(Vector(Ciclo, 0)))
                WEntra(XLugar, 14) = Vector(Ciclo, 3)
                If Val(_Provincia) = 24 Then
                    WEntra(XLugar, 15) = "U$S"
                    WEntra(XLugar, 16) = Vector(Ciclo, 4)
                Else
                    WEntra(XLugar, 15) = " $ "
                    WEntra(XLugar, 16) = Vector(Ciclo, 4)
                End If
            Else
                WEntra(XLugar, 12) = ""
                WEntra(XLugar, 13) = ""
                WEntra(XLugar, 14) = ""
                WEntra(XLugar, 15) = ""
            End If
        Next Ciclo

        XLugar = XLugar + 1

        WEntra(XLugar, 1) = txtRecibo.Text
        WEntra(XLugar, 2) = txtFecha.Text
        WEntra(XLugar, 3) = txtCliente.Text
        WEntra(XLugar, 4) = WRazon
        WEntra(XLugar, 5) = WDireccion
        WEntra(XLugar, 6) = WLocalidad
        WEntra(XLugar, 7) = WProvincia
        WEntra(XLugar, 8) = WPostal
        WEntra(XLugar, 9) = ""
        WEntra(XLugar, 10) = ""
        WEntra(XLugar, 11) = ""
        WEntra(XLugar, 12) = ""
        WEntra(XLugar, 13) = ""
        WEntra(XLugar, 14) = ""
        WEntra(XLugar, 15) = ""
        WEntra(XLugar, 16) = ""

        XLugar = XLugar + 1

        WEntra(XLugar, 1) = txtRecibo.Text
        WEntra(XLugar, 2) = txtFecha.Text
        WEntra(XLugar, 3) = txtCliente.Text
        WEntra(XLugar, 4) = WRazon
        WEntra(XLugar, 5) = WDireccion
        WEntra(XLugar, 6) = WLocalidad
        WEntra(XLugar, 7) = WProvincia
        WEntra(XLugar, 8) = WPostal
        WEntra(XLugar, 9) = ""
        WEntra(XLugar, 10) = Str$(Total1)
        WEntra(XLugar, 11) = ""
        WEntra(XLugar, 12) = ""
        WEntra(XLugar, 13) = ""
        WEntra(XLugar, 14) = ""
        If Val(_Provincia) = 24 Then
            WEntra(XLugar, 15) = "U$S"
            WEntra(XLugar, 16) = Str$(Total2)
        Else
            WEntra(XLugar, 15) = " $ "
            WEntra(XLugar, 16) = Str$(Total2)
        End If

        _ImprimirReciboPellital(WEntra)

    End Sub

    Private Sub _ImprimirReciboPellital(ByVal WEntra)
        Dim table As New DataTable("Detalles")
        Dim row As DataRow
        Dim crdoc As ReportDocument
        crdoc = New ReciboDefinitivo

        ' Creo las Columnas
        _PrepararTabla(table)

        ' Lleno la tabla con la informacion del Recibo.
        For i = 1 To 24

            row = table.NewRow()

            row.Item("Recibo") = WEntra(i, 1)
            row.Item("Renglon") = i
            row.Item("Fecha") = WEntra(i, 2)
            row.Item("Cliente") = WEntra(i, 3)
            row.Item("Razon") = WEntra(i, 4)
            row.Item("Direccion") = WEntra(i, 5)
            row.Item("Localidad") = WEntra(i, 6)
            row.Item("Provincia") = WEntra(i, 7)
            row.Item("Postal") = WEntra(i, 8)
            row.Item("Impre1") = WEntra(i, 9)
            row.Item("Importe1") = Val(WEntra(i, 10))
            row.Item("Signo1") = WEntra(i, 11)
            row.Item("Fecha2") = WEntra(i, 12)
            row.Item("Tipo2") = WEntra(i, 13)
            row.Item("Numero2") = WEntra(i, 14)
            row.Item("Signo2") = WEntra(i, 15)
            row.Item("Importe2") = Val(WEntra(i, 16))
            row.Item("Cheque1") = WEntra(i, 17)
            row.Item("Venci1") = WEntra(i, 18)
            row.Item("Impo1") = Val(WEntra(i, 19))
            row.Item("Banco1") = WEntra(i, 20)
            row.Item("Cheque2") = WEntra(i, 21)
            row.Item("Venci2") = WEntra(i, 22)
            row.Item("Impo2") = Val(WEntra(i, 23))
            row.Item("Banco2") = WEntra(i, 24)
            row.Item("Cheque3") = WEntra(i, 25)
            row.Item("Venci3") = WEntra(i, 26)
            row.Item("Impo3") = Val(WEntra(i, 27))
            row.Item("Banco3") = WEntra(i, 28)
            row.Item("Cheque4") = WEntra(i, 29)
            row.Item("Venci4") = WEntra(i, 30)
            row.Item("Impo4") = Val(WEntra(i, 31))
            row.Item("Banco4") = WEntra(i, 32)
            row.Item("Cheque5") = WEntra(i, 33)
            row.Item("Venci5") = WEntra(i, 34)
            row.Item("Impo5") = Val(WEntra(i, 35))
            row.Item("Banco5") = WEntra(i, 36)
            row.Item("Cheque6") = WEntra(i, 37)
            row.Item("Venci6") = WEntra(i, 38)
            row.Item("Impo6") = Val(WEntra(i, 39))
            row.Item("Banco6") = WEntra(i, 40)
            row.Item("Cheque7") = WEntra(i, 41)
            row.Item("Venci7") = WEntra(i, 42)
            row.Item("Impo7") = Val(WEntra(i, 43))
            row.Item("Banco7") = WEntra(i, 44)
            row.Item("Cheque8") = WEntra(i, 45)
            row.Item("Venci8") = WEntra(i, 46)
            row.Item("Impo8") = Val(WEntra(i, 47))
            row.Item("Banco8") = WEntra(i, 48)
            row.Item("Cheque9") = WEntra(i, 49)
            row.Item("Venci9") = WEntra(i, 50)
            row.Item("Impo9") = Val(WEntra(i, 51))
            row.Item("Banco9") = WEntra(i, 52)
            row.Item("Cheque10") = WEntra(i, 53)
            row.Item("Venci10") = WEntra(i, 54)
            row.Item("Impo10") = Val(WEntra(i, 55))
            row.Item("Banco10") = WEntra(i, 56)
            row.Item("Cheque11") = WEntra(i, 57)
            row.Item("Venci11") = WEntra(i, 58)
            row.Item("Impo11") = Val(WEntra(i, 59))
            row.Item("Banco11") = WEntra(i, 60)
            row.Item("Cheque12") = WEntra(i, 61)
            row.Item("Venci12") = WEntra(i, 62)
            row.Item("Impo12") = Val(WEntra(i, 63))
            row.Item("Banco12") = WEntra(i, 64)
            row.Item("Cheque13") = WEntra(i, 65)
            row.Item("Venci13") = WEntra(i, 66)
            row.Item("Impo13") = Val(WEntra(i, 67))
            row.Item("Banco13") = WEntra(i, 68)
            row.Item("Cheque14") = WEntra(i, 69)
            row.Item("Venci14") = WEntra(i, 70)
            row.Item("Impo14") = Val(WEntra(i, 71))
            row.Item("Banco14") = WEntra(i, 72)
            row.Item("Cheque15") = WEntra(i, 73)
            row.Item("Venci15") = WEntra(i, 74)
            row.Item("Impo15") = Val(WEntra(i, 75))
            row.Item("Banco15") = WEntra(i, 76)
            row.Item("Cheque16") = WEntra(i, 77)
            row.Item("Venci16") = WEntra(i, 78)
            row.Item("Impo16") = Val(WEntra(i, 79))
            row.Item("Banco16") = WEntra(i, 80)
            row.Item("Cheque17") = WEntra(i, 81)
            row.Item("Venci17") = WEntra(i, 82)
            row.Item("Impo17") = Val(WEntra(i, 83))
            row.Item("Banco17") = WEntra(i, 84)
            row.Item("Cheque18") = WEntra(i, 85)
            row.Item("Venci18") = WEntra(i, 86)
            row.Item("Impo18") = Val(WEntra(i, 87))
            row.Item("Banco18") = WEntra(i, 88)
            row.Item("Cheque19") = WEntra(i, 89)
            row.Item("Venci19") = WEntra(i, 90)
            row.Item("Impo19") = Val(WEntra(i, 91))
            row.Item("Banco19") = WEntra(i, 92)
            row.Item("Cheque20") = WEntra(i, 93)
            row.Item("Venci20") = WEntra(i, 94)
            row.Item("Impo20") = Val(WEntra(i, 95))
            row.Item("Banco20") = WEntra(i, 96)
            row.Item("Observaciones") = LSet(Trim(txtObservaciones.Text), 50)

            table.Rows.Add(row)

        Next

        crdoc.SetDataSource(table)

        _Imprimir(crdoc, 2)
    End Sub

    Private Sub _PrepararTabla(ByRef table As DataTable)
        ' Por defecto son de tipo String, asi que solamente defino explicitamente las de tipo Double.
        With table
            .Columns.Add("Recibo")
            .Columns.Add("Renglon")
            .Columns.Add("Fecha")
            .Columns.Add("Cliente")
            .Columns.Add("Razon")
            .Columns.Add("Direccion")
            .Columns.Add("Localidad")
            .Columns.Add("Provincia")
            .Columns.Add("Postal")
            .Columns.Add("Impre1")
            .Columns.Add("Importe1").DataType = System.Type.GetType("System.Double")

            .Columns.Add("Fecha2")
            .Columns.Add("Tipo2")
            .Columns.Add("Numero2")
            .Columns.Add("Importe2").DataType = System.Type.GetType("System.Double")

            .Columns.Add("Fecha21")
            .Columns.Add("Tipo21")
            .Columns.Add("Numero21")
            .Columns.Add("Importe21").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Fecha22")
            .Columns.Add("Tipo22")
            .Columns.Add("Numero22")
            .Columns.Add("Importe22").DataType = System.Type.GetType("System.Double")

            .Columns.Add("Cheque1")
            .Columns.Add("Venci1")
            .Columns.Add("Impo1").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco1")
            .Columns.Add("Cheque2")
            .Columns.Add("Venci2")
            .Columns.Add("Impo2").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco2")
            .Columns.Add("Cheque3")
            .Columns.Add("Venci3")
            .Columns.Add("Impo3").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco3")
            .Columns.Add("Cheque4")
            .Columns.Add("Venci4")
            .Columns.Add("Impo4").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco4")
            .Columns.Add("Cheque5")
            .Columns.Add("Venci5")
            .Columns.Add("Impo5").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco5")
            .Columns.Add("Cheque6")
            .Columns.Add("Venci6")
            .Columns.Add("Impo6").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco6")
            .Columns.Add("Cheque7")
            .Columns.Add("Venci7")
            .Columns.Add("Impo7").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco7")
            .Columns.Add("Cheque8")
            .Columns.Add("Venci8")
            .Columns.Add("Impo8").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco8")
            .Columns.Add("Cheque9")
            .Columns.Add("Venci9")
            .Columns.Add("Impo9").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco9")
            .Columns.Add("Cheque10")
            .Columns.Add("Venci10")
            .Columns.Add("Impo10").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco10")
            .Columns.Add("Cheque11")
            .Columns.Add("Venci11")
            .Columns.Add("Impo11").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco11")
            .Columns.Add("Cheque12")
            .Columns.Add("Venci12")
            .Columns.Add("Impo12").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco12")
            .Columns.Add("Cheque13")
            .Columns.Add("Venci13")
            .Columns.Add("Impo13").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco13")
            .Columns.Add("Cheque14")
            .Columns.Add("Venci14")
            .Columns.Add("Impo14").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco14")
            .Columns.Add("Cheque15")
            .Columns.Add("Venci15")
            .Columns.Add("Impo15").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco15")
            .Columns.Add("Cheque16")
            .Columns.Add("Venci16")
            .Columns.Add("Impo16").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco16")
            .Columns.Add("Cheque17")
            .Columns.Add("Venci17")
            .Columns.Add("Impo17").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco17")
            .Columns.Add("Cheque18")
            .Columns.Add("Venci18")
            .Columns.Add("Impo18").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco18")
            .Columns.Add("Cheque19")
            .Columns.Add("Venci19")
            .Columns.Add("Impo19").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco19")
            .Columns.Add("Cheque20")
            .Columns.Add("Venci20")
            .Columns.Add("Impo20").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco20")
            .Columns.Add("Signo21")
            .Columns.Add("Signo22")

            .Columns.Add("Signo1")
            .Columns.Add("Signo2")
            .Columns.Add("Observaciones")
        End With
    End Sub

    Private Sub _ImprimirRecibo(ByVal WEntra, ByVal WCheques)
        Dim table As New DataTable("Detalles")
        Dim cheques1 As New DataTable("Cheques1")
        Dim cheques2 As New DataTable("Cheques2")
        Dim detallado As New DataTable("Detallado")
        Dim row As DataRow
        Dim crdoc As ReportDocument
        Dim cantidad = 2
        Dim enviarEmail = False
        crdoc = New ReciboDefinitivoEmail

        If Trim(WEmail) <> "" Then

            If MsgBox("¿Desea enviar una copia del recibo por email a: " & Trim(WEmail) & " ?", MsgBoxStyle.YesNo) = DialogResult.Yes Then

                crdoc = New ReciboDefinitivoEmail

                cantidad = 1

                enviarEmail = True

            End If

        End If

        If Proceso._EsPellital() Then
            crdoc = New ReciboDefinitivoEmailPellital
        End If

        ' En caso de Recibos varios solamente se imprime una copia (a pedido de Domingo).
        If optVarios.Checked = True Then
            cantidad = 1
        End If

        ' Ajustes varios para acomodar la informacion en un pdf en vez de ser preimpreso.
        Dim ultimo = 1
        Dim WTemp(22, 5) As String

        For i = 5 To 22
            WTemp(ultimo, 1) = WEntra(i, 12)
            WTemp(ultimo, 2) = WEntra(i, 13)
            WTemp(ultimo, 3) = WEntra(i, 14)
            WTemp(ultimo, 4) = WEntra(i, 15)
            WTemp(ultimo, 5) = WEntra(i, 16)
            ultimo += 1
        Next

        ' Alineamos las filas de los recibos a la de las facturas.
        For u = 1 To 22
            If WEntra(u, 11) = "" And WEntra(u, 10) <> "" Then
                WEntra(u, 11) = " $ "
            End If

            WEntra(u, 12) = WTemp(u, 1)
            WEntra(u, 13) = WTemp(u, 2)
            WEntra(u, 14) = WTemp(u, 3)
            WEntra(u, 15) = WTemp(u, 4)
            WEntra(u, 16) = WTemp(u, 5)
        Next

        WEntra(24, 9) = RSet("TOTAL", 40)
        WEntra(24, 11) = " $ "

        '
        ' AJUSTE PARA NUEVO DISEÑO.
        '
        Dim WRecibo, WImpre1, WSigno, WImporte1, WIndice
        Dim WLeyendas(40) As String

        WLeyendas(1) = "Efectivo:"
        WLeyendas(3) = "Cheques:"
        WLeyendas(5) = "Documentos:"
        WLeyendas(7) = "Ret. Ganancias:"
        WLeyendas(9) = "Ret. Iva:"
        WLeyendas(11) = "Ret. I.Brutos:"
        WLeyendas(13) = "Monenda Ext:"
        WLeyendas(15) = "Compensacion:"
        WLeyendas(17) = "Bonos:"
        WLeyendas(19) = "Ajuste:"
        WLeyendas(21) = "Transferencia:"
        WLeyendas(22) = "Ret. SUSS:"

        ' Generamos la tabla.
        With detallado
            .Columns.Add("Recibo")

            For i = 1 To 14
                .Columns.Add("Impre" & i)
            Next

        End With

        WIndice = 0

        row = detallado.NewRow()

        Dim WTotalRecibo = 0.0

        For i = 0 To 40
            If Val(WEntra(i, 10)) <> 0 Then

                'WEntra(i, 10) = formatonumerico(WEntra(i, 10))

                If Trim(UCase(WEntra(i, 11))) = "U$S" Then
                    Dim W As Double = Val(WEntra(i, 10).ToString.Replace(",", ".")) * Val(formatonumerico(txtParidad.Text))
                    WTotalRecibo += Val(formatonumerico(W))
                Else
                    WTotalRecibo += Val(Proceso.formatonumerico(WEntra(i, 10)))
                End If
            End If
        Next

        With row

            For i = 1 To 39

                WRecibo = WEntra(i, 1)
                WImpre1 = WLeyendas(i) 'Trim(WEntra(i, 9))
                WImporte1 = WEntra(i, 10)
                WSigno = WEntra(i, 11)

                If Val(WImporte1) <> 0 And WIndice <= 13 Then

                    WIndice += 1

                    WImpre1 &= RSet(Trim(WSigno) & " " & Val(Proceso.formatonumerico(Trim(WImporte1))).ToString("##,###0.00"), 32 - WImpre1.ToString().Length)

                    .Item("Impre" & WIndice) = Trim(WImpre1)

                End If

                If WIndice > 13 Then
                    Exit For
                End If

            Next

            .Item("Impre14") = WEntra(24, 11) & " " & Val(WTotalRecibo).ToString("##,###0.00")

            .Item("Recibo") = WRecibo

        End With

        'If WIndice > 0 Then
        '    detallado.Rows.Add(row)
        'End If
        detallado.Rows.Add(row)

        ' Creo las Columnas
        _PrepararTabla(table)

        WIndice = 0
        ' Lleno la tabla con la informacion del Recibo.
        For i = 1 To 38 Step 2

            'If Val(WEntra(i, 16)) <> 0 Then
            If Val(WEntra(i, 16)) <> 0 Or Val(WEntra(i + 1, 16)) <> 0 Then
                WIndice += 2
                row = table.NewRow

                row.Item("Recibo") = WEntra(i, 1)
                row.Item("Renglon") = i
                row.Item("Fecha") = WEntra(i, 2)
                row.Item("Cliente") = WEntra(i, 3)
                row.Item("Razon") = WEntra(i, 4)
                row.Item("Direccion") = WEntra(i, 5)
                row.Item("Localidad") = WEntra(i, 6)
                row.Item("Provincia") = WEntra(i, 7)
                row.Item("Postal") = WEntra(i, 8)
                row.Item("Impre1") = WEntra(i, 9)
                row.Item("Importe1") = Val(WEntra(i, 10))
                row.Item("Signo1") = WEntra(i, 11)

                'If WIndice Mod 2 = 0 Then
                row.Item("Fecha22") = WEntra(i + 1, 12)
                row.Item("Tipo22") = WEntra(i + 1, 13)
                row.Item("Numero22") = WEntra(i + 1, 14)
                row.Item("Signo22") = WEntra(i + 1, 15)
                row.Item("Importe22") = Val(WEntra(i + 1, 16))
                'Else
                row.Item("Fecha21") = WEntra(i, 12)
                row.Item("Tipo21") = WEntra(i, 13)
                row.Item("Numero21") = WEntra(i, 14)
                row.Item("Signo21") = WEntra(i, 15)
                row.Item("Importe21") = Val(WEntra(i, 16))
                'End If

                row.Item("Observaciones") = LSet(Trim(txtObservaciones.Text), 50)

                table.Rows.Add(row)
            End If

            'End If

        Next

        For i = WIndice + 1 To 38 Step 2
            row = table.NewRow()

            row.Item("Recibo") = WEntra(i, 1)
            row.Item("Renglon") = i
            row.Item("Fecha") = WEntra(i, 2)
            row.Item("Cliente") = WEntra(i, 3)
            row.Item("Razon") = WEntra(i, 4)
            row.Item("Direccion") = WEntra(i, 5)
            row.Item("Localidad") = WEntra(i, 6)
            row.Item("Provincia") = WEntra(i, 7)
            row.Item("Postal") = WEntra(i, 8)
            row.Item("Impre1") = WEntra(i, 9)
            row.Item("Importe1") = Val(WEntra(i, 10))
            row.Item("Signo1") = WEntra(i, 11)

            'If WIndice Mod 2 = 0 Then
            row.Item("Fecha22") = ""
            row.Item("Tipo22") = ""
            row.Item("Numero22") = ""
            row.Item("Signo22") = ""
            row.Item("Importe22") = 0
            'Else
            row.Item("Fecha21") = ""
            row.Item("Tipo21") = ""
            row.Item("Numero21") = ""
            row.Item("Signo21") = ""
            row.Item("Importe21") = 0
            'End If

            table.Rows.Add(row)
        Next

        With cheques1
            .Columns.Add("Recibo")
            .Columns.Add("Cheque")
            .Columns.Add("Venci")
            .Columns.Add("Impo").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco")
            .Columns.Add("Renglon")
        End With

        With cheques2
            .Columns.Add("Recibo")
            .Columns.Add("Cheque")
            .Columns.Add("Venci")
            .Columns.Add("Impo").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Banco")
            .Columns.Add("Renglon")
        End With

        Dim WCRenglon = 0
        ' Ahora asignamos los datos de los cheques.
        For w = 0 To 39

            If Val(WCheques(w, 4)) <> 0 And (Val(WCheques(w, 0)) = 2 Or Val(WCheques(w, 0)) = 3) Then

                If WCRenglon < 20 Then

                    row = cheques1.NewRow
                    With row
                        .Item("Renglon") = WCRenglon
                        .Item("Recibo") = Trim(txtRecibo.Text)
                        .Item("Cheque") = WCheques(w, 1)
                        .Item("Venci") = WCheques(w, 2)
                        .Item("Impo") = Val(WCheques(w, 4))
                        .Item("Banco") = UCase(WCheques(w, 3))
                    End With
                    cheques1.Rows.Add(row)
                    WCRenglon += 1

                Else

                    row = cheques2.NewRow
                    With row
                        .Item("Renglon") = WCRenglon
                        .Item("Recibo") = Trim(txtRecibo.Text)
                        .Item("Cheque") = WCheques(w, 1)
                        .Item("Venci") = WCheques(w, 2)
                        .Item("Impo") = Val(WCheques(w, 4))
                        .Item("Banco") = UCase(WCheques(w, 3))
                    End With
                    cheques2.Rows.Add(row)
                    WCRenglon += 1

                End If

            End If

        Next

        Dim ds As New DataSet()

        With ds

            .Tables.Add(table)
            .Tables.Add(cheques1)
            .Tables.Add(cheques2)
            .Tables.Add(detallado)

        End With

        crdoc.SetDataSource(ds)

        If enviarEmail Then
            Try
                _EnviarReciboPorEmail(crdoc, WEmail)
            Catch ex As Exception ' En caso de que por alguna razoón no se haya podido enviar el email, se realizan las dos impresiones como se venia haciendo.
                MsgBox(ex.Message _
                      & vbCrLf & vbCrLf & vbCrLf & vbCrLf _
                      & "Se realizarán las dos impresiones normales.")
                cantidad = 2
            End Try
        End If

        _Imprimir(crdoc, cantidad)
        '_VistaPrevia(crdoc)

        btnLimpiar.PerformClick()

    End Sub

    Private Sub _EnviarReciboPorEmail(ByVal crdoc As ReportDocument, ByVal WEmail As String)
        Dim archivo As String = "Recibo" & Trim(txtRecibo.Text) & ".pdf"
        Dim ruta As String = Application.StartupPath & "/"
        Dim _to, _bcc, _asunto, _mensaje, _adjunto As String

        Try
            ' Guardamos el archivo.
            crdoc.ExportToDisk(ExportFormatType.PortableDocFormat, ruta & archivo)

            ' Confirmamos que se haya guardado correctamente el archivo.
            If Not System.IO.File.Exists(ruta & archivo) Then
                Throw New Exception
            End If
        Catch ex As Exception
            Throw New Exception("No se pudo generar el archivo PDF del Recibo. Se detiene el envio por email del mismo.")
            Exit Sub
        End Try

        Try

            _to = Trim(WEmail)
            _bcc = "" ' CONSULTAR SI SE AGREGA ESTO O NO.
            _asunto = "Recibo Nº " & Trim(txtRecibo.Text)
            _mensaje = _asunto
            _adjunto = ruta & archivo

            ' Enviamos por email e imprimimos una copia.
            _EnviarEmail(_to, _bcc, _asunto, _mensaje, _adjunto)

        Catch ex As Exception
            Throw New Exception("No se pudo enviar el E-Mail.")
        End Try

        Try
            ' Eliminamos el archivo.

            If System.IO.File.Exists(ruta & archivo) Then
                System.IO.File.Delete(ruta & archivo)
            End If
        Catch ex As Exception
            'Throw New Exception("No se pudo generar el archivo PDF del Recibo. Se detiene el envio por email del mismo.")
            Exit Sub
        End Try

    End Sub

    Private Sub _EnviarEmail(ByVal _to As String, ByVal _bcc As String, ByVal _subject As String, ByVal _body As String, ByVal _adjunto As String)
        Dim _Outlook As New Outlook.Application

        Try
            Dim _Mail As Outlook.MailItem = _Outlook.CreateItem(Outlook.OlItemType.olMailItem)

            With _Mail

                .To = _to
                .BCC = _bcc
                .Subject = _subject
                .Body = _body

                If Trim(_adjunto) <> "" Then
                    .Attachments.Add(_adjunto)
                End If

            End With

            _Mail.Send()

            _Mail = Nothing

            'Me.Close()

        Catch ex As Exception
            Throw New Exception("Ocurrió un problema al querer enviar el email a los proveedores.")
        Finally
            _Outlook = Nothing
        End Try

    End Sub

    Private Sub _Imprimir(ByVal crdoc As ReportDocument, Optional ByVal cant As Integer = 1)
        crdoc.PrintToPrinter(cant, True, 0, 0)
    End Sub

    Private Sub _VistaPrevia(ByVal crdoc As ReportDocument)
        With VistaPrevia
            .Reporte = crdoc
            .Mostrar()
        End With
    End Sub

    Private Sub _PrepararEImprimirRecibo()
        ' Comprobamos que el Recibo exista.
        If Not _ReciboYaExistente(Trim(txtRecibo.Text)) Then
            MsgBox("El Nº de Recibo indicado no se corresponde con ninguno registrado en la Base de Datos." _
                   & vbCrLf & vbCrLf & _
                   "Se detendrá el proceso de impresión", MsgBoxStyle.Information)
            Exit Sub
        End If

        Dim Retencion, Cheque, Documento, Total2f, Pesos, Bonos, Dolares, Ajuste, Compe, Transfe, Total2, Total1 As Double
        Dim Vector(50, 10) As String
        Dim WEntra(100, 180) As String
        Dim WCheques(100, 5) As String
        Dim WRCheques = 0
        Dim ImpreTipo(100) As String
        Dim XLugar = 0
        Dim iRow = 0

        ' Inicializamos todas las variable a utilizar.
        If optVarios.Checked Then
            WRazon = Space$(30)
            WDireccion = Space$(30)
            WLocalidad = Space$(30)
            WProvincia = ""
            WPostal = ""
        End If

        WRecibo = Val(txtRecibo.Text)
        WFecha = txtFecha.Text
        WCliente = txtCliente.Text

        Retencion = Val(txtRetGanancias.Text) + Val(txtRetIva.Text) + Val(_NormalizarNumero(txtRetIB.Text)) + Val(txtRetSuss.Text)

        Cheque = 0
        Documento = 0
        Total2f = 0
        Pesos = 0
        Bonos = 0
        Dolares = 0
        Ajuste = 0
        Compe = 0
        Transfe = 0
        Total2 = 0

        XTipo = ""
        XNumero = ""

        ' Limpiamos los vectores para re poblarlos.
        Array.Clear(Vector, 0, Vector.Length)
        Array.Clear(WEntra, 0, WEntra.Length)

        ImpreTipo(1) = "FC"

        ' Extraemos los valores de Débitos y Créditos.
        For iRow = 0 To 40

            If iRow < gridPagos2.Rows.Count Then
                If Trim(gridPagos2.Rows(iRow).Cells(4).Value) <> "" Then
                    With gridPagos2.Rows(iRow)
                        Vector(iRow, 0) = .Cells(0).Value
                        Vector(iRow, 1) = .Cells(1).Value
                        Vector(iRow, 2) = .Cells(2).Value
                        Vector(iRow, 3) = .Cells(3).Value
                        Vector(iRow, 4) = _NormalizarNumero(.Cells(4).Value)
                    End With
                End If
            Else
                For i = 0 To 4
                    Vector(iRow, i) = ""
                Next
            End If

            If iRow < gridFormasPago2.Rows.Count Then
                If Trim(gridFormasPago2.Rows(iRow).Cells(4).Value) <> "" Then
                    With gridFormasPago2.Rows(iRow)
                        Vector(iRow, 5) = .Cells(0).Value
                        Vector(iRow, 6) = .Cells(1).Value
                        Vector(iRow, 7) = .Cells(2).Value
                        Vector(iRow, 8) = .Cells(3).Value
                        Vector(iRow, 9) = _NormalizarNumero(.Cells(4).Value)
                    End With
                End If
            Else
                For i = 5 To 9
                    Vector(iRow, i) = ""
                Next
            End If

            XTipo = ceros(Vector(iRow, 0), 2)
            XNumero = ceros(Vector(iRow, 3), 8)

            'Solo realizamos la llamada a la base de datos en caso de haber datos por los que solicitar la fecha.
            If Val(XTipo) > 0 And Val(XNumero) > 0 Then
                Vector(iRow, 10) = _ObtenerFechaCtaCte(XTipo & XNumero & "01")
            Else
                Vector(iRow, 10) = ""
            End If

        Next iRow

        ' Volvemos a extraer los datos de los cheques para que se puedan generar mas filas.
        With gridFormasPago2
            If .Rows.Count Then
                For Each row As DataGridViewRow In .Rows

                    If row.Cells(4).Value <> "" Then

                        WCheques(WRCheques, 0) = row.Cells(0).Value
                        WCheques(WRCheques, 1) = row.Cells(1).Value
                        WCheques(WRCheques, 2) = row.Cells(2).Value
                        WCheques(WRCheques, 3) = row.Cells(3).Value
                        WCheques(WRCheques, 4) = _NormalizarNumero(row.Cells(4).Value)

                        If Val(WCheques(WRCheques, 4)) <> 0 And (Val(WCheques(WRCheques, 0)) = 2 Or Val(WCheques(WRCheques, 0)) = 3) Then
                            Cheque = Cheque + Val(WCheques(WRCheques, 4))
                        End If

                        WRCheques += 1

                    End If

                Next
            End If
        End With

        ' Calculamos totales y subtotales en base a los tipos de créditos.
        Dim _ciclo = 0
        For Ciclo = 0 To 40
            _ciclo = Ciclo
            If Val(Vector(Ciclo, 9)) <> 0 Then

                Dim _cuenta As Object = _CuentasContables.FindLast(Function(c) c(0) = _ciclo)
                Select Case Val(Vector(Ciclo, 5))
                    Case 1
                        If IsNothing(_cuenta) Then
                            Pesos = Pesos + Val(Vector(Ciclo, 9))
                        Else
                            If Val(_cuenta(1)) <> 2 Then
                                Pesos = Pesos + Val(Vector(Ciclo, 9))
                            Else
                                Dolares = Dolares + Val(Vector(Ciclo, 9))
                            End If
                        End If
                    Case 4
                        If Not IsNothing(_cuenta) Then
                            Select Case Val(_cuenta(1))
                                Case 2
                                    Dolares = Dolares + Val(Vector(Ciclo, 9))
                                Case 5
                                    Compe = Compe + Val(Vector(Ciclo, 9))
                                Case 21, 22, 25
                                    Transfe = Transfe + Val(Vector(Ciclo, 9))
                                Case 91
                                    Ajuste = Ajuste + Val(Vector(Ciclo, 9))
                                Case 157, 7, 8
                                    Bonos = Bonos + Val(Vector(Ciclo, 9))
                            End Select
                        End If
                    Case 2
                        'Cheque = Cheque + Val(Vector(Ciclo, 9))
                    Case Else
                        Documento = Documento + Val(Vector(Ciclo, 9))
                End Select
            End If

            If Val(Vector(Ciclo, 4)) <> 0 Then
                Total2 = Total2 + Val(Vector(Ciclo, 4))
            End If

        Next Ciclo

        Total1 = Pesos + Cheque + Documento + Retencion + Dolares + Compe + Transfe + Ajuste + Bonos

        '
        ' GUARDAMOS LA INFORMACION DE LOS CHEQUES
        '
        For Ciclo = 0 To 40
            If Val(Vector(Ciclo, 9)) <> 0 And (Val(Vector(Ciclo, 5)) = 2 Or Val(Vector(Ciclo, 5)) = 3) Then
                Vector(Ciclo, 6) = ceros(Vector(Ciclo, 6), 6)
                Vector(Ciclo, 8) = LSet(Vector(Ciclo, 8), 20)
                For Pasa = 1 To 40
                    Select Case Ciclo
                        Case 0
                            WEntra(Pasa, 17) = Vector(Ciclo, 6)
                            WEntra(Pasa, 18) = Vector(Ciclo, 7)
                            WEntra(Pasa, 19) = Vector(Ciclo, 9)
                            WEntra(Pasa, 20) = Vector(Ciclo, 8)
                        Case 1
                            WEntra(Pasa, 21) = Vector(Ciclo, 6)
                            WEntra(Pasa, 22) = Vector(Ciclo, 7)
                            WEntra(Pasa, 23) = Vector(Ciclo, 9)
                            WEntra(Pasa, 24) = Vector(Ciclo, 8)
                        Case 2
                            WEntra(Pasa, 25) = Vector(Ciclo, 6)
                            WEntra(Pasa, 26) = Vector(Ciclo, 7)
                            WEntra(Pasa, 27) = Vector(Ciclo, 9)
                            WEntra(Pasa, 28) = Vector(Ciclo, 8)
                        Case 3
                            WEntra(Pasa, 29) = Vector(Ciclo, 6)
                            WEntra(Pasa, 30) = Vector(Ciclo, 7)
                            WEntra(Pasa, 31) = Vector(Ciclo, 9)
                            WEntra(Pasa, 32) = Vector(Ciclo, 8)
                        Case 4
                            WEntra(Pasa, 33) = Vector(Ciclo, 6)
                            WEntra(Pasa, 34) = Vector(Ciclo, 7)
                            WEntra(Pasa, 35) = Vector(Ciclo, 9)
                            WEntra(Pasa, 36) = Vector(Ciclo, 8)
                        Case 5
                            WEntra(Pasa, 37) = Vector(Ciclo, 6)
                            WEntra(Pasa, 38) = Vector(Ciclo, 7)
                            WEntra(Pasa, 39) = Vector(Ciclo, 9)
                            WEntra(Pasa, 40) = Vector(Ciclo, 8)
                        Case 6
                            WEntra(Pasa, 41) = Vector(Ciclo, 6)
                            WEntra(Pasa, 42) = Vector(Ciclo, 7)
                            WEntra(Pasa, 43) = Vector(Ciclo, 9)
                            WEntra(Pasa, 44) = Vector(Ciclo, 8)
                        Case 7
                            WEntra(Pasa, 45) = Vector(Ciclo, 6)
                            WEntra(Pasa, 46) = Vector(Ciclo, 7)
                            WEntra(Pasa, 47) = Vector(Ciclo, 9)
                            WEntra(Pasa, 48) = Vector(Ciclo, 8)
                        Case 8
                            WEntra(Pasa, 49) = Vector(Ciclo, 6)
                            WEntra(Pasa, 50) = Vector(Ciclo, 7)
                            WEntra(Pasa, 51) = Vector(Ciclo, 9)
                            WEntra(Pasa, 52) = Vector(Ciclo, 8)
                        Case 9
                            WEntra(Pasa, 53) = Vector(Ciclo, 6)
                            WEntra(Pasa, 54) = Vector(Ciclo, 7)
                            WEntra(Pasa, 55) = Vector(Ciclo, 9)
                            WEntra(Pasa, 56) = Vector(Ciclo, 8)
                        Case 10
                            WEntra(Pasa, 57) = Vector(Ciclo, 6)
                            WEntra(Pasa, 58) = Vector(Ciclo, 7)
                            WEntra(Pasa, 59) = Vector(Ciclo, 9)
                            WEntra(Pasa, 60) = Vector(Ciclo, 8)
                        Case 11
                            WEntra(Pasa, 61) = Vector(Ciclo, 6)
                            WEntra(Pasa, 62) = Vector(Ciclo, 7)
                            WEntra(Pasa, 63) = Vector(Ciclo, 9)
                            WEntra(Pasa, 64) = Vector(Ciclo, 8)
                        Case 12
                            WEntra(Pasa, 65) = Vector(Ciclo, 6)
                            WEntra(Pasa, 66) = Vector(Ciclo, 7)
                            WEntra(Pasa, 67) = Vector(Ciclo, 9)
                            WEntra(Pasa, 68) = Vector(Ciclo, 8)
                        Case 13
                            WEntra(Pasa, 69) = Vector(Ciclo, 6)
                            WEntra(Pasa, 70) = Vector(Ciclo, 7)
                            WEntra(Pasa, 71) = Vector(Ciclo, 9)
                            WEntra(Pasa, 72) = Vector(Ciclo, 8)
                        Case 14
                            WEntra(Pasa, 73) = Vector(Ciclo, 6)
                            WEntra(Pasa, 74) = Vector(Ciclo, 7)
                            WEntra(Pasa, 75) = Vector(Ciclo, 9)
                            WEntra(Pasa, 76) = Vector(Ciclo, 8)
                        Case 15
                            WEntra(Pasa, 77) = Vector(Ciclo, 6)
                            WEntra(Pasa, 78) = Vector(Ciclo, 7)
                            WEntra(Pasa, 79) = Vector(Ciclo, 9)
                            WEntra(Pasa, 80) = Vector(Ciclo, 8)
                        Case 16
                            WEntra(Pasa, 81) = Vector(Ciclo, 6)
                            WEntra(Pasa, 82) = Vector(Ciclo, 7)
                            WEntra(Pasa, 83) = Vector(Ciclo, 9)
                            WEntra(Pasa, 84) = Vector(Ciclo, 8)
                        Case 17
                            WEntra(Pasa, 85) = Vector(Ciclo, 6)
                            WEntra(Pasa, 86) = Vector(Ciclo, 7)
                            WEntra(Pasa, 87) = Vector(Ciclo, 9)
                            WEntra(Pasa, 88) = Vector(Ciclo, 8)
                        Case 18
                            WEntra(Pasa, 89) = Vector(Ciclo, 6)
                            WEntra(Pasa, 90) = Vector(Ciclo, 7)
                            WEntra(Pasa, 91) = Vector(Ciclo, 9)
                            WEntra(Pasa, 92) = Vector(Ciclo, 8)
                        Case 19
                            WEntra(Pasa, 93) = Vector(Ciclo, 6)
                            WEntra(Pasa, 94) = Vector(Ciclo, 7)
                            WEntra(Pasa, 95) = Vector(Ciclo, 9)
                            WEntra(Pasa, 96) = Vector(Ciclo, 8)
                        Case 20
                            WEntra(Pasa, 97) = Vector(Ciclo, 6)
                            WEntra(Pasa, 98) = Vector(Ciclo, 7)
                            WEntra(Pasa, 99) = Vector(Ciclo, 9)
                            WEntra(Pasa, 100) = Vector(Ciclo, 8)
                        Case 21
                            WEntra(Pasa, 101) = Vector(Ciclo, 6)
                            WEntra(Pasa, 102) = Vector(Ciclo, 7)
                            WEntra(Pasa, 103) = Vector(Ciclo, 9)
                            WEntra(Pasa, 104) = Vector(Ciclo, 8)
                        Case 22
                            WEntra(Pasa, 105) = Vector(Ciclo, 6)
                            WEntra(Pasa, 106) = Vector(Ciclo, 7)
                            WEntra(Pasa, 107) = Vector(Ciclo, 9)
                            WEntra(Pasa, 108) = Vector(Ciclo, 8)
                        Case 23
                            WEntra(Pasa, 109) = Vector(Ciclo, 6)
                            WEntra(Pasa, 110) = Vector(Ciclo, 7)
                            WEntra(Pasa, 111) = Vector(Ciclo, 9)
                            WEntra(Pasa, 112) = Vector(Ciclo, 8)
                        Case 24
                            WEntra(Pasa, 113) = Vector(Ciclo, 6)
                            WEntra(Pasa, 114) = Vector(Ciclo, 7)
                            WEntra(Pasa, 115) = Vector(Ciclo, 9)
                            WEntra(Pasa, 116) = Vector(Ciclo, 8)
                        Case 25
                            WEntra(Pasa, 117) = Vector(Ciclo, 6)
                            WEntra(Pasa, 118) = Vector(Ciclo, 7)
                            WEntra(Pasa, 119) = Vector(Ciclo, 9)
                            WEntra(Pasa, 120) = Vector(Ciclo, 8)
                        Case 26
                            WEntra(Pasa, 121) = Vector(Ciclo, 6)
                            WEntra(Pasa, 122) = Vector(Ciclo, 7)
                            WEntra(Pasa, 123) = Vector(Ciclo, 9)
                            WEntra(Pasa, 124) = Vector(Ciclo, 8)
                        Case 27
                            WEntra(Pasa, 125) = Vector(Ciclo, 6)
                            WEntra(Pasa, 126) = Vector(Ciclo, 7)
                            WEntra(Pasa, 127) = Vector(Ciclo, 9)
                            WEntra(Pasa, 128) = Vector(Ciclo, 8)
                        Case 28
                            WEntra(Pasa, 129) = Vector(Ciclo, 6)
                            WEntra(Pasa, 130) = Vector(Ciclo, 7)
                            WEntra(Pasa, 131) = Vector(Ciclo, 9)
                            WEntra(Pasa, 132) = Vector(Ciclo, 8)
                        Case 29
                            WEntra(Pasa, 133) = Vector(Ciclo, 6)
                            WEntra(Pasa, 134) = Vector(Ciclo, 7)
                            WEntra(Pasa, 135) = Vector(Ciclo, 9)
                            WEntra(Pasa, 136) = Vector(Ciclo, 8)
                        Case 30
                            WEntra(Pasa, 137) = Vector(Ciclo, 6)
                            WEntra(Pasa, 138) = Vector(Ciclo, 7)
                            WEntra(Pasa, 139) = Vector(Ciclo, 9)
                            WEntra(Pasa, 140) = Vector(Ciclo, 8)
                        Case 31
                            WEntra(Pasa, 141) = Vector(Ciclo, 6)
                            WEntra(Pasa, 142) = Vector(Ciclo, 7)
                            WEntra(Pasa, 143) = Vector(Ciclo, 9)
                            WEntra(Pasa, 144) = Vector(Ciclo, 8)
                        Case 32
                            WEntra(Pasa, 145) = Vector(Ciclo, 6)
                            WEntra(Pasa, 146) = Vector(Ciclo, 7)
                            WEntra(Pasa, 147) = Vector(Ciclo, 9)
                            WEntra(Pasa, 148) = Vector(Ciclo, 8)
                        Case 33
                            WEntra(Pasa, 149) = Vector(Ciclo, 6)
                            WEntra(Pasa, 150) = Vector(Ciclo, 7)
                            WEntra(Pasa, 151) = Vector(Ciclo, 9)
                            WEntra(Pasa, 152) = Vector(Ciclo, 8)
                        Case 34
                            WEntra(Pasa, 153) = Vector(Ciclo, 6)
                            WEntra(Pasa, 154) = Vector(Ciclo, 7)
                            WEntra(Pasa, 155) = Vector(Ciclo, 9)
                            WEntra(Pasa, 156) = Vector(Ciclo, 8)
                        Case 35
                            WEntra(Pasa, 157) = Vector(Ciclo, 6)
                            WEntra(Pasa, 158) = Vector(Ciclo, 7)
                            WEntra(Pasa, 159) = Vector(Ciclo, 9)
                            WEntra(Pasa, 160) = Vector(Ciclo, 8)
                        Case 36
                            WEntra(Pasa, 161) = Vector(Ciclo, 6)
                            WEntra(Pasa, 162) = Vector(Ciclo, 7)
                            WEntra(Pasa, 163) = Vector(Ciclo, 9)
                            WEntra(Pasa, 164) = Vector(Ciclo, 8)
                        Case 37
                            WEntra(Pasa, 165) = Vector(Ciclo, 6)
                            WEntra(Pasa, 166) = Vector(Ciclo, 7)
                            WEntra(Pasa, 167) = Vector(Ciclo, 9)
                            WEntra(Pasa, 168) = Vector(Ciclo, 8)
                        Case 38
                            WEntra(Pasa, 169) = Vector(Ciclo, 6)
                            WEntra(Pasa, 170) = Vector(Ciclo, 7)
                            WEntra(Pasa, 171) = Vector(Ciclo, 9)
                            WEntra(Pasa, 172) = Vector(Ciclo, 8)
                        Case 39
                            WEntra(Pasa, 173) = Vector(Ciclo, 6)
                            WEntra(Pasa, 174) = Vector(Ciclo, 7)
                            WEntra(Pasa, 175) = Vector(Ciclo, 9)
                            WEntra(Pasa, 176) = Vector(Ciclo, 8)
                        Case 40
                            WEntra(Pasa, 177) = Vector(Ciclo, 6)
                            WEntra(Pasa, 178) = Vector(Ciclo, 7)
                            WEntra(Pasa, 179) = Vector(Ciclo, 9)
                            WEntra(Pasa, 180) = Vector(Ciclo, 8)
                    End Select
                Next Pasa
            End If
        Next Ciclo

        XLugar = 1

        WEntra(XLugar, 1) = txtRecibo.Text
        WEntra(XLugar, 2) = txtFecha.Text
        WEntra(XLugar, 3) = txtCliente.Text
        WEntra(XLugar, 4) = WRazon
        WEntra(XLugar, 5) = WDireccion
        WEntra(XLugar, 6) = WLocalidad
        WEntra(XLugar, 7) = WProvincia
        WEntra(XLugar, 8) = WPostal
        WEntra(XLugar, 9) = "Efectivo "
        WEntra(XLugar, 10) = Str$(Pesos)
        WEntra(XLugar, 11) = ""
        WEntra(XLugar, 12) = ""
        WEntra(XLugar, 13) = ""
        WEntra(XLugar, 14) = ""
        WEntra(XLugar, 15) = ""
        WEntra(XLugar, 16) = ""

        XLugar = XLugar + 1

        WEntra(XLugar, 1) = txtRecibo.Text
        WEntra(XLugar, 2) = txtFecha.Text
        WEntra(XLugar, 3) = txtCliente.Text
        WEntra(XLugar, 4) = WRazon
        WEntra(XLugar, 5) = WDireccion
        WEntra(XLugar, 6) = WLocalidad
        WEntra(XLugar, 7) = WProvincia
        WEntra(XLugar, 8) = WPostal
        WEntra(XLugar, 9) = ""
        WEntra(XLugar, 10) = ""
        WEntra(XLugar, 11) = ""
        WEntra(XLugar, 12) = ""
        WEntra(XLugar, 13) = ""
        WEntra(XLugar, 14) = ""
        WEntra(XLugar, 15) = ""
        WEntra(XLugar, 16) = ""

        XLugar = XLugar + 1

        WEntra(XLugar, 1) = txtRecibo.Text
        WEntra(XLugar, 2) = txtFecha.Text
        WEntra(XLugar, 3) = txtCliente.Text
        WEntra(XLugar, 4) = WRazon
        WEntra(XLugar, 5) = WDireccion
        WEntra(XLugar, 6) = WLocalidad
        WEntra(XLugar, 7) = WProvincia
        WEntra(XLugar, 8) = WPostal
        WEntra(XLugar, 9) = "Cheques "
        WEntra(XLugar, 10) = Str$(Cheque)
        WEntra(XLugar, 11) = ""
        WEntra(XLugar, 12) = ""
        WEntra(XLugar, 13) = ""
        WEntra(XLugar, 14) = ""
        WEntra(XLugar, 15) = ""
        WEntra(XLugar, 16) = ""

        XLugar = XLugar + 1

        WEntra(XLugar, 1) = txtRecibo.Text
        WEntra(XLugar, 2) = txtFecha.Text
        WEntra(XLugar, 3) = txtCliente.Text
        WEntra(XLugar, 4) = WRazon
        WEntra(XLugar, 5) = WDireccion
        WEntra(XLugar, 6) = WLocalidad
        WEntra(XLugar, 7) = WProvincia
        WEntra(XLugar, 8) = WPostal
        WEntra(XLugar, 9) = ""
        WEntra(XLugar, 10) = ""
        WEntra(XLugar, 11) = ""
        WEntra(XLugar, 12) = ""
        WEntra(XLugar, 13) = ""
        WEntra(XLugar, 14) = ""
        WEntra(XLugar, 15) = ""
        WEntra(XLugar, 16) = ""


        For Ciclo = 0 To 37

            XLugar = XLugar + 1

            WEntra(XLugar, 1) = txtRecibo.Text
            WEntra(XLugar, 2) = txtFecha.Text
            WEntra(XLugar, 3) = txtCliente.Text
            WEntra(XLugar, 4) = WRazon
            WEntra(XLugar, 5) = WDireccion
            WEntra(XLugar, 6) = WLocalidad
            WEntra(XLugar, 7) = WProvincia
            WEntra(XLugar, 8) = WPostal

            Select Case Ciclo
                Case 0
                    WEntra(XLugar, 9) = "Documentos "
                    WEntra(XLugar, 10) = Str$(Documento)
                    WEntra(XLugar, 11) = ""
                Case 2
                    WEntra(XLugar, 9) = "Retencion Ganancias "
                    WEntra(XLugar, 10) = txtRetGanancias.Text
                    WEntra(XLugar, 11) = ""
                Case 4
                    WEntra(XLugar, 9) = "Retencion Iva "
                    WEntra(XLugar, 10) = txtRetIva.Text
                    WEntra(XLugar, 11) = ""
                Case 6
                    WEntra(XLugar, 9) = "Retencion I.Brutos "
                    WEntra(XLugar, 10) = txtRetIB.Text
                    WEntra(XLugar, 11) = ""
                Case 8
                    WEntra(XLugar, 9) = "Moneda Ext."
                    If Val(txtParidad.Text) <> 0 And Val(Dolares) <> 0 Then
                        WEntra(XLugar, 10) = Str$(Val(_NormalizarNumero(Dolares)) / Val(_NormalizarNumero(txtParidad.Text)))
                        WEntra(XLugar, 11) = "U$S"
                    Else
                        WEntra(XLugar, 10) = Str$(Dolares)
                        WEntra(XLugar, 11) = ""
                    End If

                Case 10
                    WEntra(XLugar, 9) = "Compensacion"
                    WEntra(XLugar, 10) = Str$(Compe)
                    WEntra(XLugar, 11) = ""
                Case 12
                    WEntra(XLugar, 9) = "Bonos"
                    WEntra(XLugar, 10) = Str$(Bonos)
                    WEntra(XLugar, 11) = ""
                Case 14
                    WEntra(XLugar, 9) = "Ajuste"
                    WEntra(XLugar, 10) = Str$(Ajuste)
                    WEntra(XLugar, 11) = ""
                Case 16
                    WEntra(XLugar, 9) = "Transferencia"
                    WEntra(XLugar, 10) = Str$(Transfe)
                    WEntra(XLugar, 11) = ""
                Case 17
                    WEntra(XLugar, 9) = "Ret. Suss"
                    WEntra(XLugar, 10) = txtRetSuss.Text
                    WEntra(XLugar, 11) = ""
                Case Else
                    WEntra(XLugar, 9) = ""
                    WEntra(XLugar, 10) = ""
                    WEntra(XLugar, 11) = ""
            End Select

            If Val(Vector(Ciclo, 4)) <> 0 And optCtaCte.Checked Then
                Call ceros(Vector(Ciclo, 3), 6)
                WEntra(XLugar, 12) = Vector(Ciclo, 10)
                WEntra(XLugar, 13) = ImpreTipo(Val(Vector(Ciclo, 0)))
                WEntra(XLugar, 14) = Vector(Ciclo, 3)
                If Val(_Provincia) = 24 Then
                    WEntra(XLugar, 15) = "U$S"
                    WEntra(XLugar, 16) = Vector(Ciclo, 4)
                Else
                    WEntra(XLugar, 15) = " $ "
                    WEntra(XLugar, 16) = Vector(Ciclo, 4)
                End If
            Else
                WEntra(XLugar, 12) = ""
                WEntra(XLugar, 13) = ""
                WEntra(XLugar, 14) = ""
                WEntra(XLugar, 15) = ""
            End If
        Next Ciclo

        XLugar = XLugar + 1

        WEntra(XLugar, 1) = txtRecibo.Text
        WEntra(XLugar, 2) = txtFecha.Text
        WEntra(XLugar, 3) = txtCliente.Text
        WEntra(XLugar, 4) = WRazon
        WEntra(XLugar, 5) = WDireccion
        WEntra(XLugar, 6) = WLocalidad
        WEntra(XLugar, 7) = WProvincia
        WEntra(XLugar, 8) = WPostal
        WEntra(XLugar, 9) = ""
        WEntra(XLugar, 10) = ""
        WEntra(XLugar, 11) = ""
        WEntra(XLugar, 12) = ""
        WEntra(XLugar, 13) = ""
        WEntra(XLugar, 14) = ""
        WEntra(XLugar, 15) = ""
        WEntra(XLugar, 16) = ""

        XLugar = XLugar + 1

        WEntra(XLugar, 1) = txtRecibo.Text
        WEntra(XLugar, 2) = txtFecha.Text
        WEntra(XLugar, 3) = txtCliente.Text
        WEntra(XLugar, 4) = WRazon
        WEntra(XLugar, 5) = WDireccion
        WEntra(XLugar, 6) = WLocalidad
        WEntra(XLugar, 7) = WProvincia
        WEntra(XLugar, 8) = WPostal
        WEntra(XLugar, 9) = ""
        WEntra(XLugar, 10) = Str$(Total1)
        WEntra(XLugar, 11) = ""
        WEntra(XLugar, 12) = ""
        WEntra(XLugar, 13) = ""
        WEntra(XLugar, 14) = ""
        If Val(_Provincia) = 24 Then
            WEntra(XLugar, 15) = "U$S"
            WEntra(XLugar, 16) = Str$(Total2)
        Else
            WEntra(XLugar, 15) = " $ "
            WEntra(XLugar, 16) = Str$(Total2)
        End If

        _ImprimirRecibo(WEntra, WCheques)

    End Sub

    Private Function _ObtenerFechaCtaCte(ByVal clave As String) As String
        Dim fecha = ""

        cm.CommandText = "SELECT fecha FROM CtaCte WHERE Clave = '" & clave & "'"

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then
                dr.Read()

                fecha = dr.Item("fecha")

            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally

            cn.Close()

        End Try

        Return fecha
    End Function

    Private Sub btnDias_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDias.Click
        Dim Suma1, Suma2, Suma3, Suma4 As Double
        Dim FechaBase = "01/01/2006"
        Dim XFecha1 = ""
        Dim iRow = 0

        Suma1 = 0
        Suma2 = 0
        Suma3 = 0
        Suma4 = 0

        For iRow = 0 To gridPagos2.Rows.Count - 1
            With gridPagos2.Rows(iRow)
                XTipo1 = ceros(.Cells(0).Value, 2)
                XLetra1 = .Cells(1).Value
                XPunto1 = .Cells(2).Value
                XNumero1 = ceros(.Cells(3).Value, 8)
                XImporte1 = .Cells(4).Value
            End With

            ClaveCtaCte = XTipo1 + XNumero1 + "01"

            XFecha1 = _ObtenerFechaCtaCte(ClaveCtaCte)

            If XFecha1 <> "" Then
                Dim XDias1 As Integer = DateDiff("d", FechaBase, XFecha1)
                Suma1 = Suma1 + Val(XImporte1)
                Suma2 = Suma2 + (Val(XImporte1) * XDias1)
            End If
        Next

        For iRow = 0 To gridFormasPago2.Rows.Count - 1

            With gridFormasPago2.Rows(iRow)
                XTipo2 = .Cells(0).Value
                XFecha2 = .Cells(2).Value
                XImporte2 = .Cells(4).Value
            End With

            If Val(XImporte2) <> 0 Then
                If XFecha2 = "" Then
                    XFecha2 = txtFecha.Text
                End If
                Dim WFechaCheque As String = String.Join("", XFecha2.Split("/").Reverse()) 'Right$(XFecha2, 4) + Mid$(XFecha2, 4, 2) + Left$(XFecha2, 2)
                Dim WFechaRecibo As String = String.Join("", txtFecha.Text.Split("/").Reverse()) 'Right$(fecha.Text, 4) + Mid$(fecha.Text, 4, 2) + Left$(fecha.Text, 2)
                If WFechaCheque < WFechaRecibo Then
                    XFecha2 = txtFecha.Text
                End If
                Dim XDias2 As Integer = DateDiff("d", FechaBase, XFecha2)
                Suma3 = Suma3 + Val(XImporte2)
                Suma4 = Suma4 + (Val(XImporte2) * XDias2)
            End If

        Next iRow

        Dim ZImpo1 As Double = 0
        Dim ZImpo2 As Double = 0

        If Suma1 <> 0 Then
            ZImpo1 = Suma2 / Suma1
        End If
        If Suma3 <> 0 Then
            ZImpo2 = Suma4 / Suma3
        End If

        Dim ZDife As Double = ZImpo2 - ZImpo1

        MsgBox("Se esta cancelando la deuda a " + Str$(Int(ZDife)) + " Dias", MsgBoxStyle.Information, "Emision de Recibos")

    End Sub

    Private Sub _PrepararTablaIntereses(ByRef tabla As DataTable)
        ' Por defecto son de tipo String, asi que solamente defino explicitamente las de tipo Double.
        With tabla
            .Columns.Add("CodigoCliente")
            .Columns.Add("Cliente")
            .Columns.Add("Fecha")
            .Columns.Add("Numero")
            .Columns.Add("Fecha2")
            .Columns.Add("Banco")
            .Columns.Add("Importe").DataType = System.Type.GetType("System.Double")
            .Columns.Add("Dias")
            .Columns.Add("Tasa")
            .Columns.Add("Intereses").DataType = System.Type.GetType("System.Double")
        End With
    End Sub

    Private Sub btnIntereses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIntereses.Click
        Dim tabla = New DataTable("Detalles")
        Dim DiasTasa = ""
        _PedirInformacion("Informe la tasa Mensual", New TextBox, DiasTasa)

        Dim ZZSuma As Double = 0
        Dim ZZCodigo As Double = 0

        Dim FechaBase As String = txtFecha.Text

        Dim row As DataRow
        Dim crdoc As ReportDocument
        crdoc = New InformeIntereses

        ' Creo las Columnas
        _PrepararTablaIntereses(tabla)

        ' Lleno la tabla con la informacion del Recibo.
        For i = 0 To gridFormasPago2.Rows.Count - 1

            With gridFormasPago2.Rows(i)

                XTipo2 = .Cells(0).Value
                XNumero2 = ceros(.Cells(1).Value, 8)
                If Val(XNumero2) = 0 Then
                    XNumero2 = ""
                End If
                XFecha2 = .Cells(2).Value
                XBanco2 = .Cells(3).Value
                XImporte2 = .Cells(4).Value

                If Val(XImporte2) <> 0 Then

                    XImporte2 = _NormalizarNumero(XImporte2)
                    DiasTasa = _NormalizarNumero(DiasTasa)

                    If XFecha2 = "" Then
                        XFecha2 = txtFecha.Text
                    End If
                    Dim WFechaCheque As String = String.Join("", XFecha2.Split("/").Reverse())
                    Dim WFechaRecibo As String = String.Join("", txtFecha.Text.Split("/").Reverse())
                    If Val(WFechaCheque) < Val(WFechaRecibo) Then
                        XFecha2 = txtFecha.Text
                    End If
                    Dim XDias2 As Integer = DateDiff("d", FechaBase, XFecha2)
                    Dim ZZInteres As Double = ((Val(XImporte2) * XDias2 * (Val(DiasTasa) / 100)) / 365)
                    ZZInteres = Val(_NormalizarNumero(ZZInteres))
                    ZZSuma = ZZSuma + ZZInteres

                    row = tabla.NewRow()

                    row.Item("CodigoCliente") = txtCliente.Text
                    row.Item("Cliente") = txtNombre.Text
                    row.Item("Fecha") = txtFecha.Text
                    row.Item("Numero") = XNumero2
                    row.Item("Fecha2") = XFecha2
                    row.Item("Banco") = XBanco2
                    row.Item("Importe") = Val(_NormalizarNumero(XImporte2))
                    row.Item("Dias") = XDias2
                    row.Item("Tasa") = _NormalizarNumero(DiasTasa)
                    row.Item("Intereses") = Val(_NormalizarNumero(ZZInteres))

                    tabla.Rows.Add(row)
                End If

            End With

        Next

        crdoc.SetDataSource(tabla)

        _Imprimir(crdoc)
        '_VistaPrevia(crdoc)

        MsgBox("El interes a pagar es de " + Str$(ZZSuma), MsgBoxStyle.Information, "Emision de Recibos")

    End Sub

    Private Sub txtFecha_TypeValidationCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TypeValidationEventArgs)
        e.Cancel = Proceso._ValidarFecha(txtFecha.Text, e.IsValidInput)
    End Sub

    Private Sub txtFechaAux_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFechaAux.KeyDown

        If e.KeyData = Keys.Enter Then
            If Trim(txtFechaAux.Text.Replace("/", "")) = "" Then : Exit Sub : End If


            ' Completamos el año de manera automatica
            If Len(Trim(txtFechaAux.Text)) = 6 Then
                Dim _mes As String = Mid(txtFechaAux.Text, 4, 2)

                Select Case Val(_mes)
                    Case Is < 7
                        txtFechaAux.Text = Mid(txtFechaAux.Text, 1, 2) & "/" & _mes & "/" & "2019"
                    Case Else
                        txtFechaAux.Text = Mid(txtFechaAux.Text, 1, 2) & "/" & _mes & "/" & "2018"
                End Select

            End If

            'Debug.Print(Proceso._ValidarFecha(Trim(txtFechaAux.Text)))

            If Proceso._ValidarFecha(Trim(txtFechaAux.Text)) And WRow >= 0 And Wcol >= 0 Then

                If _ChequeVencido(Trim(txtFechaAux.Text)) Then
                    MsgBox("La fecha del cheque introducida es inválida", MsgBoxStyle.Critical)
                    'gridRecibos.CurrentCell = gridRecibos.Rows(WRow).Cells(Wcol)
                    Exit Sub
                End If


                With gridFormasPago2
                    .Rows(WRow).Cells(2).Value = txtFechaAux.Text

                    If Trim(.Rows(WRow).Cells(3).Value) <> "" Then
                        .CurrentCell = .Rows(WRow).Cells(4)
                    Else
                        .CurrentCell = .Rows(WRow).Cells(3)
                    End If

                    .Focus()
                End With

                txtFechaAux.Visible = False
                txtFechaAux.Location = New Point(680, 390) ' Lo reubicamos lejos de la grilla.

            End If

        ElseIf e.KeyData = Keys.Escape Then
            txtFechaAux.Text = ""
        End If

    End Sub

    Private Sub txtCliente_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtCliente.MouseDoubleClick
        lstSeleccion.SelectedIndex = 0

        lstSeleccion_Click(Nothing, Nothing)
    End Sub

    Private Sub btnCtaCte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCtaCte.Click
        lstSeleccion.SelectedIndex = 1

        lstSeleccion_Click(Nothing, Nothing)
    End Sub

    Private Sub gridRecibos_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridFormasPago2.CellClick
        With gridFormasPago2
            If e.ColumnIndex = 2 Then

                Dim _location As Point = .GetCellDisplayRectangle(2, e.RowIndex, False).Location

                .ClearSelection()
                _location.Y += .Location.Y + (.CurrentCell.Size.Height / 4) - 1.5
                _location.X += .Location.X + (.CurrentCell.Size.Width - txtFechaAux.Size.Width) - 3
                txtFechaAux.Location = _location
                txtFechaAux.Text = .Rows(e.RowIndex).Cells(2).Value
                WRow = e.RowIndex
                Wcol = e.ColumnIndex
                txtFechaAux.Visible = True
                txtFechaAux.Focus()
            End If
        End With

    End Sub

    Private Sub gridRecibos_CellEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridFormasPago2.CellEnter
        With gridFormasPago2
            If e.ColumnIndex = 2 Then

                Dim _location As Point = .GetCellDisplayRectangle(2, e.RowIndex, False).Location

                .ClearSelection()
                _location.Y += .Location.Y + (.CurrentCell.Size.Height / 4) - 1.5
                _location.X += .Location.X + (.CurrentCell.Size.Width - txtFechaAux.Size.Width) - 3
                txtFechaAux.Location = _location
                txtFechaAux.Text = .Rows(e.RowIndex).Cells(2).Value
                WRow = e.RowIndex
                Wcol = e.ColumnIndex
                txtFechaAux.Visible = True
                txtFechaAux.Focus()
            End If
        End With
    End Sub

    Private Sub Recibos_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        txtProvi.Focus()
    End Sub

    Private Sub lblDolares_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblDolares.TextChanged
        CustomLabel14.Text = _NormalizarNumero(Val(_NormalizarNumero(lblDolares.Text)) * Val(txtParidad.Text))
    End Sub

End Class