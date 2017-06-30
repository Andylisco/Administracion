﻿Imports ClasesCompartidas
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class Recibos

    ' Variables para alta de nuevo recibo
    Private _ComprobanteRetIva As String = ""
    Private _ComprobanteRetGanancias As String = ""
    Private _ComprobanteRetSuss As String = ""
    Private _RetIB1, _CompIB1, _RetIB2, _CompIB2, _RetIB3, _CompIB3, _RetIB4, _CompIB4, _RetIB5, _CompIB5, _
            _RetIB6, _CompIB6, _RetIB7, _CompIB7 As String

    ' Variables para impresion de Recibo.
    Dim WRazon, WDireccion, WLocalidad, WProvincia, WPostal, _
            WRecibo, WFecha, WCliente As String

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

    Dim iRow, WRow, renglon As Integer
    Dim _Cheque As Object = Nothing
    Dim _CuitExistente As Boolean = False
    Dim cn As SqlConnection = New SqlConnection()
    Dim cm As SqlCommand = New SqlCommand()
    Dim dr As SqlDataReader

    ' No tengo idea de para que son.
    Dim queryController As QueryController
    Dim commonEventsHandler As New CommonEventsHandler

    Private Sub Recibos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        commonEventsHandler.setIndexTab(Me)
        lstSeleccion.SelectedIndex = 0
        lstFiltrada.Visible = False

        renglon = 0

        _AlinearCeldas()

        txtProvi.Focus()

        btnLimpiar.PerformClick()

        _DeterminarParidad()
    End Sub

    Private Sub _DeterminarParidad(Optional ByVal fecha As String = "")

        If Trim(txtFecha.Text) <> "" Then

            Dim _Fecha As String = IIf(fecha = "", txtFecha.Text, fecha)

            cm.CommandText = "SELECT Cambio FROM Cambios WHERE Fecha = '" & _Fecha & "'"

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

    'Definimos las alineaciones por defecto.
    Private Sub _AlinearCeldas()
        With gridFormasPago
            .Columns("Tipo").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns("numero").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns("fecha").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns("importe").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End With

        With gridPagos
            .Columns("TipoCC").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns("NumeroCC").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns("ImporteCC").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End With
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
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("SELECT Impre, Numero, Fecha, Saldo FROM CtaCte WHERE Saldo > 0 AND Cliente = '" & Trim(txtCliente.Text) & "'")
        Dim dr As SqlDataReader

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()
            lstConsulta.Items.Clear()

            If dr.HasRows Then

                Do While dr.Read()
                    lstConsulta.Items.Add(dr.Item("Impre") & _SEPARADOR & ceros(dr.Item("Numero"), 6) & _SEPARADOR & dr.Item("Fecha") & _SEPARADOR & _Redondear(dr.Item("Saldo")))
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
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("SELECT Cliente, Razon FROM Cliente")
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
            Case Else

        End Select

    End Sub

    Private Sub _AsignarCtaCte(ByVal stringCtaCte As String)
        Dim cuenta As String() = stringCtaCte.Replace(_SEPARADOR, "$").Split("$")
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("SELECT Numero, Tipo, Saldo, TotalUs, Paridad FROM CtaCte WHERE Numero = '" & cuenta(1) & "'")
        Dim dr As SqlDataReader

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then

                Do While dr.Read()

                    If Not _CuentaUtilizada(dr.Item("Numero")) Then
                        Dim row = gridPagos.Rows.Add()

                        With gridPagos.Rows(row)
                            .Cells(0).Value = dr.Item("Tipo")
                            .Cells(3).Value = dr.Item("Numero")
                            .Cells(4).Value = _Redondear(dr.Item("Saldo"))
                        End With

                        _SumarDebitos()

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
        Dim _Paridad, _TipoCompo, _TotalUs As Double
        Dim _Error As Boolean = False
        lblTotalDebitos.Text = 0
        lblDolares.Text = 0

        For Each row As DataGridViewRow In gridPagos.Rows

            _Paridad = 0
            _TipoCompo = 0
            _TotalUs = 0

            If Val(row.Cells(4).Value) <> 0 Then

                Dim cn As SqlConnection = New SqlConnection()
                Dim cm As SqlCommand = New SqlCommand("SELECT TotalUs, Paridad, TipoCompo FROM CtaCte WHERE clave = '" & row.Cells(0).Value & ceros(row.Cells(3).Value, 8) & "01'")
                Dim dr As SqlDataReader

                SQLConnector.conexionSql(cn, cm)

                Try

                    dr = cm.ExecuteReader()

                    If dr.HasRows Then

                        With dr

                            .Read()

                            _Paridad = Val(_NormalizarNumero(.Item("Paridad")))
                            _TipoCompo = Val(IIf(IsDBNull(.Item("TipoCompo")), 0, .Item("TipoCompo")))
                            _TotalUs = Val(_NormalizarNumero(.Item("TotalUs")))

                        End With

                    End If

                Catch ex As Exception
                    MsgBox("Hubo un problema al querer consultar la informacón de la Cuenta Corriente.", MsgBoxStyle.Critical)
                    _Error = True
                Finally

                    cn.Close()

                End Try

                If _Provincia = 24 And _TotalUs <> 0 Then

                    lblTotalDebitos.Text = Val(_NormalizarNumero(lblTotalDebitos.Text)) + (Val(_NormalizarNumero(row.Cells(4).Value)) * Val(_NormalizarNumero(_Paridad)))

                    lblDolares.Text = Val(_NormalizarNumero(lblDolares.Text)) + Val(_NormalizarNumero(row.Cells(4).Value))

                Else

                    lblTotalDebitos.Text = Val(_NormalizarNumero(lblTotalDebitos.Text)) + Val(_NormalizarNumero(row.Cells(4).Value))

                    If _TipoCompo <> 2 And _Paridad <> 0 Then

                        lblDolares.Text = Val(_NormalizarNumero(lblDolares.Text)) + (Val(_NormalizarNumero(row.Cells(4).Value)) / Val(_NormalizarNumero(_Paridad)))

                    End If

                End If

            End If
        Next

        lblTotalDebitos.Text = _Redondear(_NormalizarNumero(lblTotalDebitos.Text))
        lblDolares.Text = _Redondear(_NormalizarNumero(lblDolares.Text))

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
            lblDolares.Text = "0"
        End If

    End Sub

    Private Function _CuentaUtilizada(ByRef NumeroCuenta As String)

        For i = 0 To gridPagos.Rows.Count - 1
            If gridPagos.Rows(i).Cells(3).Value = NumeroCuenta Then
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
        txtFecha.Text = Date.Today.ToString("dd/MM/yyyy")
        gridFormasPago.Rows.Clear()
        gridPagos.Rows.Clear()
        optCtaCte.Checked = True

        For Each control As Control In Panel2.Controls

            If TypeOf (control) Is TextBox Then
                control.Text = ""
            End If

        Next

        _ClavesCheques.Clear()
        _CuentasContables.Clear()

        lblDolares.Text = "0.0"
        lblTotalCreditos.Text = "0.0"
        lblTotalDebitos.Text = "0.0"
        txtProvi.Focus()
        _ResetearConsultas()
    End Sub

    Private Sub eventoSegunTipoEnFormaDePagoPara(ByVal val As Integer, ByVal rowIndex As Integer, ByVal columnIndex As Integer)
        Dim column As Integer = columnIndex
        Select Case val
            Case 1, 4
                column = 4
                With gridFormasPago.Rows(rowIndex)
                    .Cells(1).Value = ""
                    .Cells(2).Value = ""
                    .Cells(3).Value = ""
                End With
            Case 2, 3
                column = 1
            Case Else
                Exit Sub
        End Select
        gridFormasPago.CurrentCell.Value = ceros(val.ToString, 2)
        gridFormasPago.CurrentCell = gridFormasPago.Rows(rowIndex).Cells(column)
    End Sub

    ' CHECKEAR PARA QUE ESTA
    Private Sub agregarClienteABanco()
        For Each row As DataGridViewRow In gridFormasPago.Rows
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

    Private Sub txtRecibo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRecibo.Leave

        If Trim(txtRecibo.Text) <> "" Then
            txtRecibo.Text = ceros(txtRecibo.Text, 6)
            mostrarRecibo(DAORecibo.buscarRecibo(txtRecibo.Text))
        End If

    End Sub

    Private Sub mostrarRecibo(ByVal recibo As Recibo)
        Dim _fecha As String
        If IsNothing(recibo) Then
            Dim temp As String = ceros(txtRecibo.Text, 6)
            btnLimpiar.PerformClick()
            txtRecibo.Text = temp
            txtRecibo.Focus()
            Exit Sub
        Else
            txtRecibo.Text = recibo.codigo
            _fecha = recibo.fecha
            _NormalizarFecha(_fecha)
            txtFecha.Text = _fecha
            _DeterminarParidad()
            mostrarCliente(recibo.cliente.id)
            txtRetGanancias.Text = recibo.retGanancias
            txtRetIB.Text = _NormalizarNumero(recibo.retIB)
            txtRetIva.Text = recibo.retIVA
            txtRetSuss.Text = recibo.retSuss
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

                        txtRetGanancias.Text = dr.Item("RetGanancias")
                        _ComprobanteRetGanancias = dr.Item("ComproGanan")
                        txtRetIva.Text = dr.Item("RetIva")
                        _ComprobanteRetIva = dr.Item("ComproIva")
                        txtRetSuss.Text = dr.Item("RetSuss")
                        _ComprobanteRetSuss = dr.Item("ComproSuss")
                        txtRetIB.Text = _NormalizarNumero(dr.Item("RetOtra"))
                        'txtParidad.Text = IIf(IsDBNull(dr.Item("Paridad")), "", dr.Item("Paridad"))
                        _RetIB1 = dr.Item("RetIb1")
                        _CompIB1 = dr.Item("NroRetIb1")
                        _RetIB2 = dr.Item("RetIb2")
                        _CompIB2 = dr.Item("NroRetIb2")
                        _RetIB3 = dr.Item("RetIb3")
                        _CompIB3 = dr.Item("NroRetIb3")
                        _RetIB4 = dr.Item("RetIb4")
                        _CompIB4 = dr.Item("NroRetIb4")
                        _RetIB5 = dr.Item("RetIb5")
                        _CompIB5 = dr.Item("NroRetIb5")
                        _RetIB6 = dr.Item("RetIb6")
                        _CompIB6 = dr.Item("NroRetIb6")
                        _RetIB7 = dr.Item("RetIb7")
                        _CompIB7 = dr.Item("NroRetIb7")
                        txtProvi.Text = dr.Item("Provisorio")

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
                            _CuentasContables.Add({Val(.Item("Renglon")) - 1, .Item("Cuenta")})
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
        Dim _FechaValida As Boolean = True
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
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("SELECT * FROM RecibosProvi WHERE Recibo = '" + _ReciboProvisorio + "'")
        Dim dr As SqlDataReader


        If Trim(_ReciboProvisorio) = "" Then
            Exit Sub
        End If


        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then
                gridFormasPago.Rows.Clear()

                Do While dr.Read()
                    Dim renglon As String = dr.Item("Renglon").ToString()


                    If Val(dr.Item("ReciboDefinitivo")) <> 0 Then
                        MsgBox("El recibo Provisorio indicado ya se encuentra asociado a un Recibo Definitivo y no se encuentra disponible.", MsgBoxStyle.Information)
                        Exit Try
                    End If


                    If renglon = "01" Then
                        txtProvi.Text = dr.Item("Recibo")
                        txtFecha.Text = dr.Item("Fecha")
                        txtCliente.Text = dr.Item("Cliente")
                        Dim cliente As Cliente = DAOCliente.buscarClientePorCodigo(txtCliente.Text)

                        If Not IsNothing(cliente) Then
                            txtNombre.Text = cliente.razon
                            cliente = Nothing
                        End If

                        txtRetGanancias.Text = dr.Item("RetGanancias")
                        _ComprobanteRetGanancias = dr.Item("ComproGanan")
                        txtRetIva.Text = dr.Item("RetIva")
                        _ComprobanteRetIva = dr.Item("ComproIva")
                        txtRetSuss.Text = dr.Item("RetSuss")
                        _ComprobanteRetSuss = dr.Item("ComproSuss")
                        txtRetIB.Text = -_NormalizarNumero(dr.Item("RetOtra"))
                        'txtParidad.Text = IIf(IsDBNull(dr.Item("Paridad")), "", dr.Item("Paridad"))
                        _RetIB1 = dr.Item("RetIb1")
                        _CompIB1 = dr.Item("NroRetIb1")
                        _RetIB2 = dr.Item("RetIb2")
                        _CompIB2 = dr.Item("NroRetIb2")
                        _RetIB3 = dr.Item("RetIb3")
                        _CompIB3 = dr.Item("NroRetIb3")
                        _RetIB4 = dr.Item("RetIb4")
                        _CompIB4 = dr.Item("NroRetIb4")
                        _RetIB5 = dr.Item("RetIb5")
                        _CompIB5 = dr.Item("NroRetIb5")
                        _RetIB6 = dr.Item("RetIb6")
                        _CompIB6 = dr.Item("NroRetIb6")
                        _RetIB7 = dr.Item("RetIb7")
                        _CompIB7 = dr.Item("NroRetIb7")

                        If Val(dr.Item("TipoRec")) = 1 Then
                            optCtaCte.Checked = True
                        ElseIf Val(dr.Item("TipoRec")) = 2 Then
                            optAnticipos.Checked = True
                        End If

                    End If

                    Dim _FormaPagoActual As Integer = gridFormasPago.Rows.Add()

                    With gridFormasPago.Rows(_FormaPagoActual)
                        .Cells(0).Value = dr.Item("Tipo2")
                        .Cells(1).Value = dr.Item("Numero2")
                        .Cells(2).Value = dr.Item("Fecha2")
                        .Cells(3).Value = dr.Item("banco2")
                        .Cells(4).Value = _NormalizarNumero(Str$(dr.Item("Importe2")))
                    End With

                    If Val(dr.Item("Tipo2")) = 2 Then
                        With dr
                            _ClavesCheques.Add({CInt(.Item("Renglon")) - 1, .Item("ClaveCheque"), .Item("BancoCheque") _
                                               , .Item("SucursalCheque"), .Item("ChequeCheque"), .Item("CuentaCheque") _
                                               , .Item("Cuit"), "", ""})
                        End With
                    End If

                    If Val(dr.Item("Tipo2")) = 4 Then
                        With dr
                            _CuentasContables.Add({Val(.Item("Renglon")) - 1, .Item("Cuenta")})
                        End With
                    End If

                Loop

                _SumarCreditos()

                txtRecibo.Text = "0"
                txtParidad.Focus()

            Else

                MsgBox("El Recibo Provisorio indicado no existe.", MsgBoxStyle.Information)

            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar el recibo provisorio.", MsgBoxStyle.Critical)
        Finally


            cn.Close()


        End Try
    End Sub

    Private Function _NormalizarNumero(ByVal numero As String) As String

        If numero.Contains(",") Then
            numero = String.Format("{0:F2}", CDbl(numero)).Replace(",", ".")
        ElseIf Not numero.Contains(".") Then
            numero &= ".00"
        End If

        Return numero
    End Function

    Private Function _SumarCreditos() As Boolean
        Dim _Error As Boolean = False
        lblTotalCreditos.Text = Val(txtRetGanancias.Text) + Val(txtRetIva.Text) + Val(txtRetIB.Text) + Val(txtRetSuss.Text)

        For Each row As DataGridViewRow In gridFormasPago.Rows

            If Val(row.Cells(4).Value) <> 0 Then

                If Val(row.Cells(0).Value) = 2 Then

                    If IsDate(row.Cells(2).Value) And IsDate(txtFecha.Text) And DateDiff(DateInterval.Day, CDate(row.Cells(2).Value), CDate(txtFecha.Text)) > 30 Then
                        _Error = True
                        Exit For
                    End If

                End If

                lblTotalCreditos.Text = Val(_NormalizarNumero(lblTotalCreditos.Text)) + Val(row.Cells(4).Value)

            End If

        Next

        lblTotalCreditos.Text = _NormalizarNumero(lblTotalCreditos.Text)
        Return _Error
    End Function

    Private Sub mostrarPagos(ByVal pagos As List(Of Pago))
        gridPagos.Rows.Clear()
        For Each pago As Pago In pagos
            gridPagos.Rows.Add(pago.tipo, pago.letra, pago.punto, pago.numero, _NormalizarNumero(pago.importe))
        Next
    End Sub

    Private Sub mostrarFormasPago(ByVal formasPago As List(Of FormaPago))
        gridFormasPago.Rows.Clear()
        For Each forma As FormaPago In formasPago
            gridFormasPago.Rows.Add(forma.tipo, forma.numero, forma.fecha, forma.nombre, _NormalizarNumero(forma.importe))
        Next
    End Sub

    Private Sub mostrarCliente(ByVal cliente As String)
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("SELECT c.Cliente, c.Razon, c.Direccion, c.Localidad, c.Provincia, p.Nombre, c.Postal " _
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
        If Val(txtRecibo.Text) <> 0 Then

            If _ReciboYaExistente(Trim(txtRecibo.Text)) Then
                MsgBox("El recibo ya se encuentra guardado con anterioridad y no puede ser actualizado.", MsgBoxStyle.Information)
            Else
                MsgBox("La numeracion del recibo es automatica, se debe grabar con numero 0")
            End If

            Exit Sub
        End If

        ' Calculamos nuevamente los saldo de débitos y Créditos.
        If _ErrorEnSumaDeSaldos() Then
            MsgBox("Hay errores en la carga de debitos o creditos. Por favor, verifique los datos y vuelva a intentar.", MsgBoxStyle.Information)
            Exit Sub
        End If

        ' Comprobamos que haya valores con las que realizar el recibo
        If Val(lblTotalCreditos.Text) = 0 And Val(lblTotalDebitos.Text) = 0 Then
            Exit Sub
        End If

        ' Comprobamos que los datos sean coherentes.
        If Val(lblTotalCreditos.Text) = Val(lblTotalDebitos.Text) Or optAnticipos.Checked Or optVarios.Checked Then
            ' Comprbamos la diferencia de cambio y consultamos con el usuario sobre si graba igual o no.
            If optCtaCte.Checked And _ExisteDiferenciaDeCambio() Then
                If MsgBox("Existe una diferencia de cambio de U$S " & lblDolares.Text & vbCrLf & "¿Desea grabar igualmente el recibo?", MsgBoxStyle.OkCancel, MsgBoxStyle.Information) = DialogResult.Cancel Then
                    Exit Sub
                Else
                    MsgBox("Recuerde emitir la diferencia de Cambio.", MsgBoxStyle.Information)
                End If
            End If

            ' Comprobamos que los creditos indicados como de tipo 4 tengan su respectiva cuenta contable asignada.
            If Not _ImputadosCorrectamenteValoresVarios() Then
                MsgBox("No se ha imputado correctamente ingreso de valores varios.", MsgBoxStyle.Information)
                Exit Sub
            End If

            ' Cargamos la numeración del recibo a guardar.
            txtRecibo.Text = _NumeracionDeReciboSiguiente()

            ' Normalizamos el número de recibo.
            txtRecibo.Text = ceros(txtRecibo.Text, 6)

            ' Comprobamos que ya no se haya grabado el recibo con anterioridad.
            If Not _NumeroDeReciboDisponible() Then
                MsgBox("El número de recibo ya se ha grabado con anterioridad y no se encuentra disponible.", MsgBoxStyle.Information)
                Exit Sub
            End If

            ' Comprobamos que se hayan cargado tanto recibo como fecha. En caso de que no, salimos y cancelamos la grabación.
            If Trim(txtRecibo.Text) = "" Or Trim(txtFecha.Text) = "" Then
                MsgBox("Por favor verifique que tanto el número de Recibo como la fecha del mismo se encuentren cargados.", MsgBoxStyle.Information)
                Exit Sub
            End If

            ' Comprobamos que haya paridad cargada para la fecha
            If Val(txtParidad.Text) = 0 Then
                MsgBox("No hay paridad cargada para esta fecha")
                Exit Sub
            End If

            ' Damos de alta el recibo según el tipo de recibo definido por el usuario.

            ' COMENZAR CON LAS ULTIMAS DOS QUE SON MAS CORTAS.
            If optCtaCte.Checked Then
                _AltaCtaCte()
            ElseIf optAnticipos.Checked Then
                _AltaAnticipo()
            ElseIf optVarios.Checked Then
                _AltaVarios()
            End If

            ' Agregamos al registro del Recibo, los creditos ingresados por el usuario.
            _AltaCreditos()

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

            ' Si llegamos hasta aca se supone que todo salio bien.
            MsgBox("Se ha generado correctamente el Recibo solicitado. Nº de Recibo: " & txtRecibo.Text, MsgBoxStyle.Information)
            btnLimpiar.PerformClick()
        Else

            MsgBox("Los montos de debitos y creditos no se corresponden entre si o con el Tipo de Cobro.", MsgBoxStyle.Information)
            Exit Sub

        End If

    End Sub

    Private Function _ReciboYaExistente(ByVal recibo As String) As Boolean
        Dim _existe As Boolean = False
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

    Private Sub _AltaCreditos()

        Dim iRow As Integer
        Dim _Cheque As Object = Nothing
        Dim _CuitExistente As Boolean = False
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand()
        Dim dr As SqlDataReader

        ' Ahora guardamos los creditos.
        For iRow = 0 To gridFormasPago.Rows.Count - 1

            If Val(gridFormasPago.Rows(iRow).Cells(4).Value) <> 0 Then
                renglon += 1

                XRecibo = txtRecibo.Text
                XRenglon = ceros(renglon, 2)
                XCliente = txtCliente.Text
                XFecha = txtFecha.Text
                XFechaOrd = String.Join("", txtFecha.Text.Split("/").Reverse())
                XTipoRec = "1"
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
                XTipo2 = gridFormasPago.Rows(iRow).Cells(0).Value
                XNumero2 = gridFormasPago.Rows(iRow).Cells(1).Value
                XFecha2 = gridFormasPago.Rows(iRow).Cells(2).Value
                XFechaOrd2 = String.Join("", XFecha2.Split("/").Reverse())
                XBanco2 = gridFormasPago.Rows(iRow).Cells(3).Value
                XImporte2 = gridFormasPago.Rows(iRow).Cells(4).Value
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
                XCuit = ""
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
                        + XFechaDepoOrd + "'"

                ' Damos de alta el nuevo renglon del nuevo recibo.
                XSql = "INSERT INTO  Recibos (Clave, Recibo, Renglon, Cliente, Fecha, Fechaord, TipoRec, RetGanancias, RetIva, RetOtra, Retencion," _
                & "TipoReg, Tipo1, Letra1, Punto1, Numero1, Importe1, Tipo2, Numero2, Fecha2, banco2, Importe2," _
                & "Estado2, Empresa, FechaOrd2, Importe, Observaciones, Impolist, Impo1list, Destino, Cuenta, Marca," _
                & "FechaDepo, FechaDepoOrd) " _
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

                If Val(gridFormasPago.Rows(iRow).Cells(0).Value) = 3 Then
                    XTipo = "50"
                    XNumero = ceros(gridFormasPago.Rows(iRow).Cells(1).Value, 8)
                    XRenglon = "01"
                    XCliente = txtCliente.Text
                    XFecha = txtFecha.Text
                    XEstado = "1"
                    XVencimiento = gridFormasPago.Rows(iRow).Cells(2).Value
                    XVencimiento1 = XVencimiento
                    XTotal = gridFormasPago.Rows(iRow).Cells(4).Value
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
        Dim _Marcar As Boolean = True

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
        XSql = XSql & " DifCambio = " & "'" & lblDolares.Text & "',"
        XSql = XSql & " RetSuss = " & "'" & txtRetSuss.Text & "',"
        XSql = XSql & " ComproGanan = " & "'" & _ComprobanteRetGanancias & "',"
        XSql = XSql & " ComproIva = " & "'" & _ComprobanteRetIva & "',"
        XSql = XSql & " ComproIb = ''," 'Comprobar si se sigue colocando o no. Me parece que no.
        XSql = XSql & " ComproSuss = " & "'" & _ComprobanteRetSuss & "',"
        XSql = XSql & " NroRetIb1 = " & "'" & _CompIB1 & "',"
        XSql = XSql & " NroRetIb2 = " & "'" & _CompIB2 & "',"
        XSql = XSql & " NroRetIb3 = " & "'" & _CompIB3 & "',"
        XSql = XSql & " NroRetIb4 = " & "'" & _CompIB4 & "',"
        XSql = XSql & " NroRetIb5 = " & "'" & _CompIB5 & "',"
        XSql = XSql & " NroRetIb6 = " & "'" & _CompIB6 & "',"
        XSql = XSql & " NroRetIb7 = " & "'" & _CompIB7 & "',"
        XSql = XSql & " RetIb1 = " & "'" & _RetIB1 & "',"
        XSql = XSql & " RetIb2 = " & "'" & _RetIB2 & "',"
        XSql = XSql & " RetIb3 = " & "'" & _RetIB3 & "',"
        XSql = XSql & " RetIb4 = " & "'" & _RetIB4 & "',"
        XSql = XSql & " RetIb5 = " & "'" & _RetIB5 & "',"
        XSql = XSql & " RetIb6 = " & "'" & _RetIB6 & "',"
        XSql = XSql & " RetIb7 = " & "'" & _RetIB7 & "'"
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

    Private Sub _AltaCtaCte()

        Dim iRow, WRow As Integer
        Dim _Cheque As Object = Nothing
        Dim _CuitExistente As Boolean = False
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand()
        Dim dr As SqlDataReader

        For iRow = 0 To gridPagos.Rows.Count - 1

            WRow = iRow
            If Val(gridPagos.Rows(iRow).Cells(4).Value) <> 0 Then

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
                XTipo1 = gridPagos.Rows(iRow).Cells(0).Value
                XLetra1 = gridPagos.Rows(iRow).Cells(1).Value
                XPunto1 = gridPagos.Rows(iRow).Cells(2).Value
                XNumero1 = ceros(gridPagos.Rows(iRow).Cells(3).Value, 8)
                If Val(_Provincia) = 24 Then
                    XImporte1 = Str$(Val(gridPagos.Rows(iRow).Cells(4).Value) * Val(txtParidad.Text))
                Else
                    XImporte1 = Str$(Val(gridPagos.Rows(iRow).Cells(4).Value))
                End If
                XImporteBaja = Str$(Val(gridPagos.Rows(iRow).Cells(4).Value))
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
                        + XFechaDepoOrd + "'"

                ' Damos de alta el nuevo renglon del nuevo recibo.
                XSql = "INSERT INTO  Recibos (Clave, Recibo, Renglon, Cliente, Fecha, Fechaord, TipoRec, RetGanancias, RetIva, RetOtra, Retencion," _
                & "TipoReg, Tipo1, Letra1, Punto1, Numero1, Importe1, Tipo2, Numero2, Fecha2, banco2, Importe2," _
                & "Estado2, Empresa, FechaOrd2, Importe, Observaciones, Impolist, Impo1list, Destino, Cuenta, Marca," _
                & "FechaDepo, FechaDepoOrd) " _
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

                ' Actualizamos los datos del cheque.
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
                    MsgBox("Hubo un problema al querer actualizar el recibo.", MsgBoxStyle.Critical)
                    Exit Sub
                Finally


                    cn.Close()


                    XSql = ""
                    XParam = ""

                End Try

                ' Calculamos el saldo para poder actualizar la cuenta cte.
                ClaveCtaCte = XTipo1 + XNumero1 + "01"
                cm.CommandText = "SELECT Saldo, Total, TotalUs, Estado FROM CtaCte Where Clave = '" & ClaveCtaCte & "'"
                SQLConnector.conexionSql(cn, cm)

                Try

                    dr = cm.ExecuteReader()

                    If dr.HasRows Then

                        With dr
                            .Read()

                            XSaldo = _Redondear((Val(.Item("Saldo")) - Val(XImporteBaja)).ToString())

                            If Val(.Item("TotalUs")) <> 0 Then
                                XSaldoUs = _Redondear((Val(.Item("Total")) / Val(.Item("TotalUs"))).ToString())
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
                    End If

                Catch ex As Exception
                    MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
                Finally


                    cn.Close()


                    XSql = ""
                    XParam = ""

                End Try

                If Val(XSaldo) <> 0 Then

                    XSql = "UPDATE CtaCte " _
                        & "SET " _
                        & " Saldo = '" & XSaldo & "'," _
                        & " SaldoUs = '" & XSaldoUs & "', " _
                        & " Estado = '" & XEstado & "', " _
                        & " Wdate = '" & XDate & "'" _
                        & " WHERE Clave = '" & ClaveCtaCte & "'"

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

                End If
            Else
                Exit Sub
            End If
        Next

    End Sub

    Private Sub _AltaVarios()

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
        XImporte1 = Str$(lblTotalCreditos.Text)
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
        XImporte = Str$(XImporte1)
        XCuenta = txtCuenta.Text
        XMarca = ""
        XFechaDepo = ""
        XFechaDepoOrd = ""
        XImpolist = ""
        XImpo1list = ""
        XDestino = ""

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
                        + XFechaDepoOrd + "'"

        XSql = "INSERT INTO  Recibos (Clave, Recibo, Renglon, Cliente, Fecha, Fechaord, TipoRec, RetGanancias, RetIva, RetOtra, Retencion," _
            & "TipoReg, Tipo1, Letra1, Punto1, Numero1, Importe1, Tipo2, Numero2, Fecha2, banco2, Importe2," _
            & "Estado2, Empresa, FechaOrd2, Importe, Observaciones, Impolist, Impo1list, Destino, Cuenta, Marca," _
            & "FechaDepo, FechaDepoOrd) " _
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

    Private Sub _AltaAnticipo()

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
            XTotal = Str$(lblTotalCreditos.Text * -1 / Val(txtParidad.Text))
            XSaldo = Str$(lblTotalCreditos.Text * -1 / Val(txtParidad.Text))
        Else
            XTotal = Str$(lblTotalCreditos.Text * -1)
            XSaldo = Str$(lblTotalCreditos.Text * -1)
        End If

        XTotalUs = Str$(lblTotalCreditos.Text * -1 / Val(txtParidad.Text))
        XSaldoUs = Str$(lblTotalCreditos.Text * -1 / Val(txtParidad.Text))

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
        XImporte1 = Str$(lblTotalCreditos.Text)
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
        XImporte = XImporte1
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
                        + XFechaDepoOrd + "'"

        XSql = "INSERT INTO  Recibos (Clave, Recibo, Renglon, Cliente, Fecha, Fechaord, TipoRec, RetGanancias, RetIva, RetOtra, Retencion," _
            & "TipoReg, Tipo1, Letra1, Punto1, Numero1, Importe1, Tipo2, Numero2, Fecha2, banco2, Importe2," _
            & "Estado2, Empresa, FechaOrd2, Importe, Observaciones, Impolist, Impo1list, Destino, Cuenta, Marca," _
            & "FechaDepo, FechaDepoOrd) " _
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

    Private Function _ImputadosCorrectamenteValoresVarios() As Boolean
        Dim _Error As Boolean = False

        For Each credito As DataGridViewRow In gridFormasPago.Rows

            If Val(credito.Cells(0).Value) = 4 Then
                If Not _CuentaContableIngresada(credito.Index) Then
                    _Error = True
                End If
            End If

        Next

        Return _Error
    End Function

    Private Function _CuentaContableIngresada(ByVal index As Integer)
        Dim _Ingresada As Boolean = False
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
        Dim _Error As Boolean = False

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
        Dim ultimo As Integer = 0

        cm.CommandText = "SELECT Recibo FROM Recibos WHERE Recibo < 6000 ORDER BY Recibo DESC"
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
        For Each row As DataGridViewRow In gridFormasPago.Rows
            If Not row.IsNewRow Then
                formasPago.Add(New FormaPago(row.Cells(0).Value, 0, asString(row.Cells(1).Value), asString(row.Cells(2).Value), asString(row.Cells(3).Value), CustomConvert.toDoubleOrZero(row.Cells(4).Value)))
            End If
        Next
        Return formasPago
    End Function

    Private Function crearPagos() As List(Of Pago)
        Dim pagos As New List(Of Pago)
        For Each row As DataGridViewRow In gridPagos.Rows
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
            gridPagos.Rows.Clear()
            gridPagos.AllowUserToAddRows = False
            'sumarValores()
        Else
            gridPagos.AllowUserToAddRows = True
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

            _SaltarA(txtParidad)

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

            .Dispose()

        End With

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
            _SaltarA(txtFecha)
        ElseIf e.KeyData = Keys.Escape Then
            txtRecibo.Text = "0"
        End If

    End Sub

    Private Sub txtFecha_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFecha.KeyDown
        If e.KeyData = Keys.Enter Then

            If Trim(txtFecha.Text) <> "" Then
                _DeterminarParidad()
            End If

            _SaltarA(txtProvi)
        ElseIf e.KeyData = Keys.Escape Then
            txtRecibo.Text = Date.Now.ToString("dd/MM/yyyy")
        End If
    End Sub

    Private Sub txtCliente_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCliente.KeyDown
        If e.KeyData = Keys.Enter Then
            _SaltarA(txtRetGanancias)
        ElseIf e.KeyData = Keys.Escape Then
            txtCliente.Text = ""
        End If
    End Sub

    Private Sub txtParidad_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtParidad.KeyDown
        If e.KeyData = Keys.Enter Then

            If Val(txtParidad.Text) > 0 Then
                _SaltarA(txtRetIva)
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
            With gridFormasPago
                .Rows(0).Selected = True
                .CurrentCell = .Rows(0).Cells(0)
                .Focus()
            End With
        ElseIf e.KeyData = Keys.Escape Then
            txtObservaciones.Text = ""
        End If
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean

        If gridFormasPago.Focused Or gridFormasPago.IsCurrentCellInEditMode Then ' Detectamos los ENTER tanto si solo estan en foco o si estan en edición una celda.
            gridFormasPago.CommitEdit(DataGridViewDataErrorContexts.Commit) ' Guardamos todos los datos que no hayan sido confirmados.

            Dim iCol = gridFormasPago.CurrentCell.ColumnIndex
            Dim iRow = gridFormasPago.CurrentCell.RowIndex

            If msg.WParam.ToInt32() = Keys.Enter Then

                Dim valor = gridFormasPago.Rows(iRow).Cells(iCol).Value

                If Not IsNothing(valor) Then

                    If Trim(valor.ToString.Length) = 31 Then
                        If _ProcesarCheque(iRow, valor) Then
                            gridFormasPago.CurrentCell = gridFormasPago.Rows(iRow).Cells(iCol + 2) ' Nos desplazamos para que coloque la fecha del cheque.
                        End If
                    Else
                        valor = valor.ToString().Substring(valor.ToString.Length - 1, 1)
                        If valor = "1" Or valor = "2" Or valor = "3" Or valor = "4" Then
                            eventoSegunTipoEnFormaDePagoPara(CustomConvert.toIntOrZero(valor), iRow, iCol)
                        Else ' Sólo se aceptan los valores 1 (Efectivo) , 2 (Cheque), 3 (Doc) y 4 (Varios) ?
                            gridFormasPago.CurrentCell = gridFormasPago.Rows(iRow).Cells(iCol)
                        End If
                    End If

                    If iCol = 0 And iRow > -1 Then
                        

                    Else

                        'Modificar el valor de la columna de ser fecha DD/MM --> DD/MM/YYYY (año actual)
                        If iCol = 2 Then

                            If _NormalizarFecha(valor) Then
                                gridFormasPago.Rows(iRow).Cells(2).Value = valor
                                gridFormasPago.EndEdit()
                            End If

                        End If

                        If iCol = 1 Or iCol = 2 Or iCol = 3 Then
                            If iCol = 2 Then
                                If _ChequeVencido(valor) Then
                                    MsgBox("La fecha del cheque introducida es inválida", MsgBoxStyle.Critical)
                                    gridFormasPago.CurrentCell = gridFormasPago.Rows(iRow).Cells(iCol)
                                ElseIf iCol = 3 Then
                                    gridFormasPago.Rows(iRow).Cells(iCol).Value = _GenerarCodigoBanco(valor)
                                    gridFormasPago.CurrentCell = gridFormasPago.Rows(iRow).Cells(iCol + 1)
                                Else
                                    gridFormasPago.Rows(iRow).Cells(iCol).Value = valor
                                    gridFormasPago.CurrentCell = gridFormasPago.Rows(iRow).Cells(iCol + 1)
                                End If
                            Else
                                gridFormasPago.CurrentCell = gridFormasPago.Rows(iRow).Cells(iCol + 1)
                            End If
                        End If

                        If iCol = 4 Then ' Avanzamos a la fila siguiente.
                            If Val(gridFormasPago.Rows(iRow).Cells(0).Value) = 4 Then
                                _PedirCuentaContable(iRow)
                            ElseIf Val(gridFormasPago.Rows(iRow).Cells(0).Value) = 2 Then
                                _PedirClaveCheque(iRow)
                            End If

                            gridFormasPago.Rows(iRow).Cells(4).Value = _NormalizarNumero(valor)
                            gridFormasPago.CurrentCell = gridFormasPago.Rows(iRow + 1).Cells(0)
                            _SumarCreditos()
                        End If
                    End If

                    Return True
                End If
            ElseIf msg.WParam.ToInt32() = Keys.Escape Then
                gridFormasPago.Rows.Remove(gridFormasPago.Rows(iRow))
                _SumarCreditos()
            End If
        End If

        If gridPagos.Focused Or gridPagos.IsCurrentCellInEditMode Then ' Detectamos los ENTER tanto si solo estan en foco o si estan en edición una celda.
            gridPagos.CommitEdit(DataGridViewDataErrorContexts.Commit) ' Guardamos todos los datos que no hayan sido confirmados.

            Dim iCol = gridPagos.CurrentCell.ColumnIndex
            Dim iRow = gridPagos.CurrentCell.RowIndex

            If msg.WParam.ToInt32() = Keys.Enter Then

                _ComprobarDebitoPosible(iRow, iCol)

                Return True
            End If
            If msg.WParam.ToInt32() = Keys.Escape Then

                gridPagos.Rows.Remove(gridPagos.Rows(iRow))
                _SumarDebitos()

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
        Dim _cargado As Boolean = False

        If _ChequeUtilizadoEnRecibo(ClaveCheque) Then

            _cargado = True

        ElseIf Not _ChequeUtilizadoEnReciboProvisorio(ClaveCheque) Then

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
        Dim utilizado As Boolean = False
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
        Dim utilizado As Boolean = False
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
        Dim _LecturaCorrecta As Boolean = True

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

        With gridFormasPago.Rows(row)
            .Cells(0).Value = "02"
            .Cells(1).Value = _NumCheque
            .Cells(2).Value = ""
            .Cells(3).Value = _Banco & "/" & Mid(txtCliente.Text, 1, 1) & Val(Mid(txtCliente.Text, 2, 6)).ToString()
        End With
        ' Buscamos si existe el cuit.
        _Cuit = _TraerNumeroCuit(_ClaveBanco & _Sucursal & _NumCta)

        ' Guardamos el nuevo Cheque.
        _GuardarNuevoCheque(row, ClaveCheque, _Banco, _Sucursal, _NumCheque, _NumCta, _Cuit)

        Return _LecturaCorrecta
    End Function

    Private Function _ObtenerNombreBanco(ByVal claveBanco As String) As String
        Dim _NombreBanco As String = ""
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
        Dim _cuit As String = ""
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
            _cuit = buscar(6)
            _PedirInformacion("Ingrese Cuit del Firmante", New TextBox(), _cuit)
            'If _cuit.Length = 11 Then : Exit Sub : End If
        Else
            Do While Not _cuit.Length = 11

                _PedirInformacion("Ingrese Cuit del Firmante", New TextBox(), _cuit)

            Loop
        End If

        _clave = ""
        _banco = ""
        _sucursal = ""
        _cheque = ""
        _cuenta = ""
        _estado = ""
        _destino = ""

        _ClavesCheques.Add({row, _clave, _banco, _sucursal, _cheque, _cuenta, _cuit, _estado, _destino})

    End Sub

    Private Sub _PedirCuentaContable(ByVal row As Integer)
        Dim cuenta As String
        Dim buscar As Object

        buscar = _CuentasContables.FindLast(Function(c) c(0) = row)

        If Not IsNothing(buscar) Then
            cuenta = buscar(1)
            _PedirInformacion("Ingrese Cuenta Contable", New TextBox(), cuenta)
            If cuenta = "" Then : Exit Sub : End If
        End If

        Do While Not _CuentaContableValida(cuenta)

            _PedirInformacion("Ingrese Cuenta Contable", New TextBox(), cuenta)

            If cuenta = "" Then
                Exit Do
            End If
        Loop

        _CuentasContables.Add({row, Trim(cuenta)})

    End Sub

    Private Function _CuentaContableValida(ByRef cuenta As String)
        Dim _CuentaValida As Boolean = False

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
        Dim _Maximo As Object = _ValoresMax.FindLast(Function(m) m(0) = gridPagos.Rows(iRow).Cells(3).Value)
        If Not IsNothing(_Maximo) Then
            If iCol = 4 And (IsNothing(gridPagos.CurrentCell.Value) Or Val(gridPagos.CurrentCell.Value) > _Maximo(1)) Then
                MsgBox("El valor del importe indicado no puede ser mayor a " & _Maximo(1))
                gridPagos.CurrentCell = gridPagos.Rows(iRow).Cells(iCol)
            Else
                gridPagos.CurrentCell.Value = _NormalizarNumero(Val(gridPagos.CurrentCell.Value))
                _SumarDebitos()
                gridPagos.CurrentCell = gridPagos.Rows(iRow + 1).Cells(iCol)
            End If
        End If
    End Sub

    Private Sub lstFiltrada_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstFiltrada.SelectedIndexChanged
        If Trim(lstFiltrada.SelectedItem) <> "" Then
            _TraerConsulta(lstFiltrada.SelectedItem)
        End If
    End Sub

    Private Sub txtCliente_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCliente.TextChanged
        gridPagos.Rows.Clear()
    End Sub

    Private Sub gridPagos_CellBeginEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles gridPagos.CellBeginEdit

        If gridPagos.CurrentCell.ColumnIndex = 4 And Not IsNothing(gridPagos.CurrentCell.Value) And Val(gridPagos.Rows(e.RowIndex).Cells(3).Value) > 0 Then

            Dim valor As Object = _ValoresMax.Find(Function(f As Object)
                                                       Return f(0) = gridPagos.Rows(e.RowIndex).Cells(3).Value
                                                   End Function)

            If IsNothing(valor) Then
                _ValoresMax.Add({gridPagos.Rows(e.RowIndex).Cells(3).Value, gridPagos.Rows(e.RowIndex).Cells(4).Value})
            End If

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

    Private Sub txtFecha_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFecha.TextChanged

        If Trim(txtFecha.Text).Length = 10 Then
            _DeterminarParidad()
        End If

    End Sub

    Private Sub btnImpresion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImpresion.Click

        ' Chequeamos que se haya cargado un recibo antes de mandar a imprimir.
        If Val(txtRecibo.Text) > 0 Then
            _PrepararEImprimirRecibo()
        End If

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
            .Columns.Add("Signo1")
            .Columns.Add("Signo2")
            .Columns.Add("Observaciones")
        End With
    End Sub

    Private Sub _ImprimirRecibo(ByVal WEntra)
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

        _Imprimir(crdoc)
        '_VistaPrevia(crdoc)
    End Sub

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

    Private Sub _PrepararEImprimirRecibo()
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
        Dim XLugar As Integer = 0
        Dim iRow As Integer = 0

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

            If iRow < gridPagos.Rows.Count Then
                If Trim(gridPagos.Rows(iRow).Cells(4).Value) <> "" Then
                    With gridPagos.Rows(iRow)
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

            If iRow < gridFormasPago.Rows.Count Then
                If Trim(gridFormasPago.Rows(iRow).Cells(4).Value) <> "" Then
                    With gridFormasPago.Rows(iRow)
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

        _ImprimirRecibo(WEntra)

    End Sub

    Private Function _ObtenerFechaCtaCte(ByVal clave As String) As String
        Dim fecha As String = ""

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
        Dim FechaBase As String = "01/01/2006"
        Dim XFecha1 = ""
        Dim iRow As Integer = 0

        Suma1 = 0
        Suma2 = 0
        Suma3 = 0
        Suma4 = 0

        For iRow = 0 To gridPagos.Rows.Count - 1
            With gridPagos.Rows(iRow)
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

        For iRow = 0 To gridFormasPago.Rows.Count - 1

            With gridFormasPago.Rows(iRow)
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
        Dim tabla As DataTable = New DataTable("Detalles")
        Dim DiasTasa As String = ""
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
        For i = 0 To gridFormasPago.Rows.Count - 1

            With gridFormasPago.Rows(i)

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
End Class