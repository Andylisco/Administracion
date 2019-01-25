Imports System.Data.SqlClient
Imports ClasesCompartidas

Public Class ListadoImputacionesContable

    Dim txtVectorBanco(1000) As String

    Private Sub ListadoImputacionesContable_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label2.Text = Globals.NombreEmpresa()
        txtDesdeFecha.Text = "  /  /    "
        txthastafecha.Text = "  /  /    "

        txtDesdeCuenta.Text = ""
        txtHastaCuenta.Text = "99999999999"


        TipoListado.Items.Clear()
        TipoListado.Items.Add("Completo")
        TipoListado.Items.Add("Resumido")
        TipoListado.SelectedIndex = 0

        chkDepositos.Checked = False
        chkPagos.Checked = False
        chkRecibos.Checked = False
    End Sub

    Private Sub txtdesdefecha_KeyPress(ByVal sender As Object, _
               ByVal e As System.Windows.Forms.KeyPressEventArgs) _
               Handles txtDesdeFecha.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            If ValidaFecha(txtDesdeFecha.Text) = "S" Then
                txthastafecha.Focus()
            End If
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtDesdeFecha.Text = "  /  /    "
            Me.txtDesdeFecha.SelectionStart = 0
        End If
    End Sub

    Private Sub txthastafecha_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txthastafecha.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            If ValidaFecha(txthastafecha.Text) = "S" Then
                txtDesdeCuenta.Focus()
            End If
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txthastafecha.Text = "  /  /    "
            Me.txthastafecha.SelectionStart = 0
        End If
    End Sub

    Private Sub txtdesdecuenta_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtDesdeCuenta.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            txtHastaCuenta.Text = txtDesdeCuenta.Text
            txtHastaCuenta.Focus()
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtDesdeCuenta.Text = ""
        End If
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txthastacuenta_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtHastaCuenta.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            txtDesdeFecha.Focus()
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtHastaCuenta.Text = ""
        End If
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnCancela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancela.Click
        Me.Close()
        MenuPrincipal.Show()
    End Sub


    Private Sub btnConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsulta.Click

        Me.Size = New System.Drawing.Size(556, 557)

        lstAyuda.DataSource = DAOCuentaContable.buscarCuentaContablePorDescripcion("")

        txtAyuda.Text = ""
        txtAyuda.Visible = True
        lstAyuda.Visible = True

        txtAyuda.Focus()

    End Sub

    '    Private Sub txtAyuda_KeyPress(ByVal sender As Object, _
    '                   ByVal e As System.Windows.Forms.KeyPressEventArgs)
    '
    '        If e.KeyChar = Convert.ToChar(Keys.Return) Then
    '            e.Handled = True
    '            lstAyuda.DataSource = DAOCuentaContable.buscarCuentaContablePorDescripcion(txtAyuda.Text)
    '        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
    '            e.Handled = True
    '            txtAyuda.Text = ""
    '            lstAyuda.DataSource = DAOCuentaContable.buscarCuentaContablePorDescripcion(txtAyuda.Text)
    '        End If
    '    End Sub

    Private Sub mostrarcuenta(ByVal cuenta As CuentaContable)
        txtDesdeCuenta.Text = cuenta.id
        txtHastaCuenta.Text = cuenta.id
        txtDesdeCuenta.Focus()
    End Sub

    Private Sub lstAyuda_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstAyuda.Click
        mostrarcuenta(lstAyuda.SelectedValue)
        REM txtDesdeProveedor.Text = lstAyuda.SelectedValue.id
    End Sub

    Enum Reporte
        Pantalla
        Imprimir
    End Enum

    Private Sub _Imprimir(ByVal TipoImpresion As Reporte)

        Dim txtUno As String

        Dim txtEmpresa = "", WEmpresa = "1"
        Dim txtFormula As String
        Dim x As Char = Chr(34)
        Dim txtDesdefechaOrd, txtHastafechaOrd
        Dim txtCorte As String
        Dim txtRenglonII As Integer

        Dim txtClave, txtClaveOrd, txtCuenta As String
        Dim txtLetra, txtTipo, txtPunto, txtNumero, txtRenglon As String
        Dim txtDebito, txtCredito As Double
        Dim txtProveedor, txtFecha, txtObservaciones As String
        Dim txtTipomovi, txtFechaOrd, txtTitulo, txtTituloList, txtVarios As String
        Dim txtNroInterno As Integer
        Dim txtAuxiliar As Double

        Dim varProvincia, varClaveRecibo As String
        Dim varParidadRecibo As Double

        Dim txtBancoCodigo As Integer
        Dim txtBancoCuenta As String

        Dim viewer As ReportViewer = Nothing

        'Dim banco As New Banco(0, "", New CuentaContable(0, ""))

        SQLConnector.retrieveDataTable("limpiar_impCyb")

        txtEmpresa = "Surfactan S.A."

        If Proceso._EsPellital() Then
            txtEmpresa = "Pellital S.A."
            WEmpresa = "8"
        End If

        txtDesdefechaOrd = ordenaFecha(txtDesdeFecha.Text)
        txtHastafechaOrd = ordenaFecha(txthastafecha.Text)

        Dim tablaII As DataTable
        tablaII = SQLConnector.retrieveDataTable("buscar_banco_por_nombre", "")

        For Each row As DataRow In tablaII.Rows

            Dim CamposBanco As New LeeBanco(row.Item(0), row.Item(1), row.Item(2))

            txtBancoCuenta = CamposBanco.Cuenta
            txtBancoCodigo = CamposBanco.banco

            txtVectorBanco(txtBancoCodigo) = txtBancoCuenta

        Next



        txtCorte = ""
        txtCuenta = ""
        txtRenglonII = 0

        If chkPagos.Checked = True Then

            Dim tabla As DataTable
            tabla = SQLConnector.retrieveDataTable("buscar_pagos_fecha", txtDesdefechaOrd, txtHastafechaOrd)

            For Each row As DataRow In tabla.Rows



                Dim CampoPagos As New LeePagosII(row.Item(0).ToString, row.Item(1).ToString, row.Item(2).ToString,
                                               row.Item(3), row.Item(4).ToString, row.Item(5).ToString,
                                               row.Item(6).ToString, row.Item(7).ToString, row.Item(8).ToString,
                                               row.Item(9), row.Item(10), row.Item(11), row.Item(12),
                                               row.Item(13), row.Item(14), row.Item(15), row.Item(16), row.Item(17),
                                               row.Item(18), row.Item(19), row.Item(20))

                'Select Case CampoPagos.orden
                '    Case 117336, 117534, 117539, 117381, 117551
                '        Stop
                '    Case Else
                'End Select

                If txtCorte <> CampoPagos.orden Then
                    txtCorte = CampoPagos.orden
                    txtRenglonII = 0
                End If

                Select Case CampoPagos.tiporeg
                    Case 1
                        If CampoPagos.tipoOrd = "3" Or CampoPagos.tipoOrd = "4" Or CampoPagos.tipoOrd = "5" Then

                            Select Case CampoPagos.tipoOrd
                                Case "4"
                                    txtCuenta = txtVectorBanco(CampoPagos.banco2)

                                Case "5"
                                    txtCuenta = "111"
                                Case Else
                                    txtCuenta = CampoPagos.cuenta
                            End Select

                        Else

                            Dim CampoProveedor As Proveedor = DAOProveedor.buscarProveedorPorCodigo(CampoPagos.proveedor)

                            If Not IsNothing(CampoProveedor) Then

                                txtCuenta = "2001"

                                If Proceso._EsPellital() AndAlso Val(CampoProveedor.id) = 10077777777 Then
                                    txtCuenta = "2046"
                                End If

                                If Val(CampoProveedor.provincia) = 24 Then
                                    txtCuenta = "2010"
                                End If
                            End If

                        End If

                        If Trim(CampoPagos.letra1) = "" And Val(CampoPagos.tipo1) = 0 And Val(CampoPagos.punto1) = 0 And Val(CampoPagos.numero1) = 0 Then
                            If Trim(CampoPagos.cuenta) <> "" And CampoPagos.cuenta <> "999999" Then
                                txtCuenta = CampoPagos.cuenta
                            End If
                        End If

                        txtRenglonII = txtRenglonII + 1

                        txtTipomovi = "1"
                        txtNroInterno = CampoPagos.orden
                        txtProveedor = CampoPagos.proveedor
                        txtTipo = CampoPagos.tipo1
                        txtLetra = CampoPagos.letra1
                        txtPunto = CampoPagos.punto1
                        txtNumero = CampoPagos.numero1
                        txtRenglon = txtRenglonII
                        txtFecha = CampoPagos.fecha
                        txtObservaciones = CampoPagos.observaciones
                        txtDebito = CampoPagos.importe1
                        txtCredito = 0
                        txtFechaOrd = CampoPagos.fechaord
                        txtTitulo = "Pagos"
                        txtEmpresa = 1
                        txtTituloList = "Surfactan S.A."

                        If Proceso._EsPellital() Then
                            txtTituloList = "Pellital S.A."
                        End If

                        txtVarios = "Desde el " + txtDesdeFecha.Text + " hasta el " + txthastafecha.Text

                        txtClave = txtTipomovi + txtNroInterno + txtRenglon
                        txtClaveOrd = txtTipomovi + txtNroInterno

                        SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                      txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)


                        REM graba las retenciones
                        If CampoPagos.renglon = 1 And CampoPagos.retotra <> 0 Then

                            txtRenglonII = txtRenglonII + 1

                            txtTipo = ""
                            txtLetra = ""
                            txtPunto = ""
                            txtNumero = ""
                            txtRenglon = txtRenglonII
                            txtCuenta = "2108"
                            txtDebito = 0
                            txtCredito = CampoPagos.retotra

                            txtClave = txtTipomovi + txtNroInterno + txtRenglon
                            txtClaveOrd = txtTipomovi + txtNroInterno

                            SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                          txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)

                        End If

                        If CampoPagos.renglon = 1 And CampoPagos.retencion <> 0 Then

                            txtRenglonII = txtRenglonII + 1

                            txtTipo = ""
                            txtLetra = ""
                            txtPunto = ""
                            txtNumero = ""
                            txtRenglon = txtRenglonII
                            txtCuenta = "2101"
                            txtDebito = 0
                            txtCredito = CampoPagos.retencion

                            txtClave = txtTipomovi + txtNroInterno + txtRenglon
                            txtClaveOrd = txtTipomovi + txtNroInterno

                            SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                          txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)

                        End If


                        If CampoPagos.renglon = 1 And CampoPagos.retiva <> 0 Then

                            txtRenglonII = txtRenglonII + 1

                            txtTipo = ""
                            txtLetra = ""
                            txtPunto = ""
                            txtNumero = ""
                            txtRenglon = txtRenglonII
                            txtCuenta = "2111"
                            txtDebito = 0
                            txtCredito = CampoPagos.retiva

                            txtClave = txtTipomovi + txtNroInterno + txtRenglon
                            txtClaveOrd = txtTipomovi + txtNroInterno

                            SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                          txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)

                        End If


                        If CampoPagos.renglon = 1 And CampoPagos.retibciudad <> 0 Then

                            txtRenglonII = txtRenglonII + 1

                            txtTipo = ""
                            txtLetra = ""
                            txtPunto = ""
                            txtNumero = ""
                            txtRenglon = txtRenglonII
                            txtCuenta = "2113"
                            txtDebito = 0
                            txtCredito = CampoPagos.retibciudad

                            txtClave = txtTipomovi + txtNroInterno + txtRenglon
                            txtClaveOrd = txtTipomovi + txtNroInterno

                            SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                          txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)

                        End If



                    Case Else
                        Select Case Val(CampoPagos.tipo2)
                            Case 1
                                txtCuenta = "1"

                            Case 2
                                'REM banco
                                'txtCuenta = "999999"
                                'Banco = DAOBanco.buscarBancoPorCodigo(CampoPagos.banco2)
                                'If Not IsNothing(Banco) Then
                                '    txtCuenta = Banco.cuenta.id.ToString()
                                'End If
                                txtCuenta = txtVectorBanco(CampoPagos.banco2)

                            Case 3
                                txtCuenta = "40"

                            Case 5
                                txtCuenta = "2"

                            Case 6
                                txtCuenta = CampoPagos.cuenta

                            Case Else
                                txtCuenta = "101"

                        End Select

                        txtRenglonII = txtRenglonII + 1

                        txtTipomovi = "1"
                        txtNroInterno = CampoPagos.orden
                        txtProveedor = CampoPagos.proveedor
                        txtTipo = CampoPagos.tipo2
                        txtLetra = ""
                        txtPunto = ""
                        txtNumero = CampoPagos.numero2
                        txtRenglon = txtRenglonII
                        txtFecha = CampoPagos.fecha
                        txtObservaciones = CampoPagos.observaciones
                        txtDebito = 0
                        txtCredito = CampoPagos.importe2
                        txtFechaOrd = CampoPagos.fechaord
                        txtTitulo = "Pagos"
                        txtEmpresa = 1
                        txtTituloList = "Surfactan S.A."

                        If Proceso._EsPellital() Then
                            txtTituloList = "Pellital S.A."
                        End If

                        txtVarios = "Desde el " + txtDesdeFecha.Text + " hasta el " + txthastafecha.Text

                        txtClave = txtTipomovi + txtNroInterno + txtRenglon
                        txtClaveOrd = txtTipomovi + txtNroInterno

                        SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                      txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)


                        If CampoPagos.renglon = 1 And CampoPagos.retotra <> 0 Then

                            txtRenglonII = txtRenglonII + 1

                            txtTipo = ""
                            txtLetra = ""
                            txtPunto = ""
                            txtNumero = ""
                            txtRenglon = txtRenglonII
                            txtCuenta = "2108"
                            txtDebito = 0
                            txtCredito = CampoPagos.retotra

                            txtClave = txtTipomovi + txtNroInterno + txtRenglon
                            txtClaveOrd = txtTipomovi + txtNroInterno

                            SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                          txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)

                        End If

                        If CampoPagos.renglon = 1 And CampoPagos.retencion <> 0 Then

                            txtRenglonII = txtRenglonII + 1

                            txtTipo = ""
                            txtLetra = ""
                            txtPunto = ""
                            txtNumero = ""
                            txtRenglon = txtRenglonII
                            txtCuenta = "2101"
                            txtDebito = 0
                            txtCredito = CampoPagos.retencion

                            txtClave = txtTipomovi + txtNroInterno + txtRenglon
                            txtClaveOrd = txtTipomovi + txtNroInterno

                            SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                          txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)

                        End If


                        If CampoPagos.renglon = 1 And CampoPagos.retiva <> 0 Then

                            txtRenglonII = txtRenglonII + 1

                            txtTipo = ""
                            txtLetra = ""
                            txtPunto = ""
                            txtNumero = ""
                            txtRenglon = txtRenglonII
                            txtCuenta = "2111"
                            txtDebito = 0
                            txtCredito = CampoPagos.retiva

                            txtClave = txtTipomovi + txtNroInterno + txtRenglon
                            txtClaveOrd = txtTipomovi + txtNroInterno

                            SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                          txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)

                        End If


                        If CampoPagos.renglon = 1 And CampoPagos.retibciudad <> 0 Then

                            txtRenglonII = txtRenglonII + 1

                            txtTipo = ""
                            txtLetra = ""
                            txtPunto = ""
                            txtNumero = ""
                            txtRenglon = txtRenglonII
                            txtCuenta = "2113"
                            txtDebito = 0
                            txtCredito = CampoPagos.retibciudad

                            txtClave = txtTipomovi + txtNroInterno + txtRenglon
                            txtClaveOrd = txtTipomovi + txtNroInterno

                            SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                          txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)

                        End If

                End Select

            Next

        End If
















        Dim tablaCtaCte As DataTable
        Dim ZRetIb(14, 2) As String
        ZRetIb(1, 2) = "162"
        ZRetIb(2, 2) = "161"
        ZRetIb(3, 2) = "190"
        ZRetIb(4, 2) = "191"
        ZRetIb(5, 2) = "192"
        ZRetIb(6, 2) = "193"
        ZRetIb(7, 2) = "194"
        ZRetIb(8, 2) = "195"

        ZRetIb(9, 2) = "176"
        ZRetIb(10, 2) = "170"
        ZRetIb(11, 2) = "171"
        ZRetIb(12, 2) = "167"
        ZRetIb(13, 2) = "172"
        ZRetIb(14, 2) = "179"

        Dim ZExcluir(1000) As String
        Dim Zex = 0

        txtCorte = ""
        txtRenglonII = 0

        If chkRecibos.Checked = True Then

            Dim tabla As DataTable
            tabla = SQLConnector.retrieveDataTable("buscar_recibos_fecha", txtDesdefechaOrd, txtHastafechaOrd)

            For Each row As DataRow In tabla.Rows

                Dim CampoRecibos As New LeeRecibos(row.Item(0), row.Item(1), row.Item(2),
                                               row.Item(3), row.Item(4), row.Item(5),
                                               row.Item(6), row.Item(7), row.Item(8),
                                               row.Item(9), row.Item(10), row.Item(11), row.Item(12),
                                               row.Item(13), row.Item(14), row.Item(15), row.Item(16), row.Item(17),
                                               row.Item(18), row.Item(19), row.Item(20))

                REM If Val(CampoRecibos.cuenta) = 162 Then : Stop : End If

                REM If Val(CampoRecibos.recibo) = 88760 Then Stop
                REM If Val(CampoRecibos.recibo) = 88564 Then Stop


                If txtCorte <> CampoRecibos.recibo Then
                    txtCorte = CampoRecibos.recibo
                    txtRenglonII = 0
                End If

                varProvincia = CampoRecibos.provincia

                Select Case Val(CampoRecibos.tiporeg)
                    Case 1

                        If CampoRecibos.tiporec = "3" Then
                            txtCuenta = CampoRecibos.cuenta
                        Else
                            If Val(CampoRecibos.tipo1) > 49 Then
                                txtCuenta = "101"
                            Else
                                txtCuenta = "91"
                                If Val(WEmpresa) <> 1 Then
                                    If Val(CampoRecibos.provincia) = 24 Then
                                        txtCuenta = "92"
                                    End If
                                End If
                            End If
                        End If

                        txtAuxiliar = CampoRecibos.importe1
                        varParidadRecibo = 0

                        If Val(varProvincia) = 24 Then

                            varClaveRecibo = "06" + ceros(CampoRecibos.recibo, 8) + "01"
                            REM Dim tablaCtaCte As DataTable
                            tablaCtaCte = SQLConnector.retrieveDataTable("lee_cuenta_corriente_clave", varClaveRecibo)
                            For Each rowctacte As DataRow In tablaCtaCte.Rows
                                Dim CampoCtacte As New LeeCtaCte(rowctacte.Item(0), rowctacte.Item(1), rowctacte.Item(2), rowctacte.Item(3), rowctacte.Item(4), rowctacte.Item(5), rowctacte.Item(6), rowctacte.Item(7), rowctacte.Item(8), rowctacte.Item(9), rowctacte.Item(10), rowctacte.Item(11), rowctacte.Item(12), rowctacte.Item(13), rowctacte.Item(14))
                                varParidadRecibo = CampoCtacte.paridad
                            Next

                            If varParidadRecibo = 0 Then
                                varClaveRecibo = "07" + ceros(CampoRecibos.recibo, 8) + "01"
                                REM Dim tablaCtaCte As DataTable
                                tablaCtaCte = SQLConnector.retrieveDataTable("lee_cuenta_corriente_clave", varClaveRecibo)
                                For Each rowctacte As DataRow In tablaCtaCte.Rows
                                    Dim CampoCtacte As New LeeCtaCte(rowctacte.Item(0), rowctacte.Item(1), rowctacte.Item(2), rowctacte.Item(3), rowctacte.Item(4), rowctacte.Item(5), rowctacte.Item(6), rowctacte.Item(7), rowctacte.Item(8), rowctacte.Item(9), rowctacte.Item(10), rowctacte.Item(11), rowctacte.Item(12), rowctacte.Item(13), rowctacte.Item(14))
                                    varParidadRecibo = CampoCtacte.paridad
                                Next
                            End If

                            If Val(CampoRecibos.tipo1) <> 7 And varParidadRecibo <> 0 Then
                                txtAuxiliar = txtAuxiliar / varParidadRecibo
                            Else
                                txtAuxiliar = CampoRecibos.importe1
                            End If

                            If Val(CampoRecibos.tipo1) <> 7 Then
                                varClaveRecibo = CampoRecibos.tipo1 + CampoRecibos.numero1 + "01"
                                REM Dim tablaCtaCte As DataTable
                                tablaCtaCte = SQLConnector.retrieveDataTable("lee_cuenta_corriente_clave", varClaveRecibo)
                                For Each rowctacte As DataRow In tablaCtaCte.Rows
                                    Dim CampoCtacte As New LeeCtaCte(rowctacte.Item(0), rowctacte.Item(1), rowctacte.Item(2), rowctacte.Item(3), rowctacte.Item(4), rowctacte.Item(5), rowctacte.Item(6), rowctacte.Item(7), rowctacte.Item(8), rowctacte.Item(9), rowctacte.Item(10), rowctacte.Item(11), rowctacte.Item(12), rowctacte.Item(13), rowctacte.Item(14))
                                    If CampoCtacte.totalus <> 0 Then
                                        txtAuxiliar = txtAuxiliar * CampoCtacte.paridad
                                    End If
                                Next
                            End If

                        End If







                        txtRenglonII = txtRenglonII + 1

                        txtTipomovi = "3"
                        txtNroInterno = CampoRecibos.recibo
                        txtProveedor = ""
                        txtTipo = CampoRecibos.tipo1
                        txtLetra = CampoRecibos.letra1
                        txtPunto = CampoRecibos.punto1
                        txtNumero = CampoRecibos.numero1
                        txtRenglon = txtRenglonII
                        txtFecha = CampoRecibos.fecha
                        txtObservaciones = ""
                        txtDebito = 0
                        txtCredito = txtAuxiliar
                        txtFechaOrd = CampoRecibos.fechaord
                        txtTitulo = "Recibos"
                        txtEmpresa = 1
                        txtTituloList = "Surfactan S.A."

                        If Proceso._EsPellital() Then
                            txtTituloList = "Pellital S.A."
                        End If

                        txtVarios = "Desde el " + txtDesdeFecha.Text + " hasta el " + txthastafecha.Text

                        txtClave = txtTipomovi + txtNroInterno + txtRenglon
                        txtClaveOrd = txtTipomovi + txtNroInterno

                        SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                      txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)


                    Case Else
                        Select Case Val(CampoRecibos.tipo2)
                            Case 1
                                txtCuenta = "1"
                            Case 2
                                txtCuenta = "40"
                            Case 4
                                txtCuenta = CampoRecibos.cuenta
                            Case Else
                                txtCuenta = "101"
                        End Select

                        txtRenglonII = txtRenglonII + 1

                        txtTipomovi = "3"
                        txtNroInterno = CampoRecibos.recibo
                        txtProveedor = ""
                        txtTipo = CampoRecibos.tipo2
                        txtLetra = ""
                        txtPunto = ""
                        txtNumero = CampoRecibos.numero2
                        txtRenglon = txtRenglonII
                        txtFecha = CampoRecibos.fecha
                        txtObservaciones = ""
                        txtDebito = CampoRecibos.importe2
                        txtCredito = 0
                        txtFechaOrd = CampoRecibos.fechaord
                        txtTitulo = "Recibos"
                        txtEmpresa = 1
                        txtTituloList = "Surfactan S.A."

                        If Proceso._EsPellital() Then
                            txtTituloList = "Pellital S.A."
                        End If

                        txtVarios = "Desde el " + txtDesdeFecha.Text + " hasta el " + txthastafecha.Text

                        txtClave = txtTipomovi + txtNroInterno + txtRenglon
                        txtClaveOrd = txtTipomovi + txtNroInterno

                        SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                      txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)

                End Select

                If Val(CampoRecibos.renglon) = 1 And Val(CampoRecibos.retganancias) <> 0 Then

                    txtRenglonII = txtRenglonII + 1
                    txtCuenta = "142"

                    txtTipomovi = "3"
                    txtNroInterno = CampoRecibos.recibo
                    txtProveedor = ""
                    txtTipo = ""
                    txtLetra = ""
                    txtPunto = ""
                    txtNumero = ""
                    txtRenglon = txtRenglonII
                    txtFecha = CampoRecibos.fecha
                    txtObservaciones = ""
                    txtDebito = CampoRecibos.retganancias
                    txtCredito = 0
                    txtFechaOrd = CampoRecibos.fechaord
                    txtTitulo = "Recibos"
                    txtEmpresa = 1
                    txtTituloList = "Surfactan S.A."

                    If Proceso._EsPellital() Then
                        txtTituloList = "Pellital S.A."
                    End If

                    txtVarios = "Desde el " + txtDesdeFecha.Text + " hasta el " + txthastafecha.Text

                    txtClave = txtTipomovi + txtNroInterno + txtRenglon
                    txtClaveOrd = txtTipomovi + txtNroInterno

                    SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                  txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)

                End If


                If Val(CampoRecibos.renglon) = 1 And Val(CampoRecibos.retiva) <> 0 Then

                    txtRenglonII = txtRenglonII + 1
                    txtCuenta = "153"

                    txtTipomovi = "3"
                    txtNroInterno = CampoRecibos.recibo
                    txtProveedor = ""
                    txtTipo = ""
                    txtLetra = ""
                    txtPunto = ""
                    txtNumero = ""
                    txtRenglon = txtRenglonII
                    txtFecha = CampoRecibos.fecha
                    txtObservaciones = ""
                    txtDebito = CampoRecibos.retiva
                    txtCredito = 0
                    txtFechaOrd = CampoRecibos.fechaord
                    txtTitulo = "Recibos"
                    txtEmpresa = 1
                    txtTituloList = "Surfactan S.A."

                    If Proceso._EsPellital() Then
                        txtTituloList = "Pellital S.A."
                    End If

                    txtVarios = "Desde el " + txtDesdeFecha.Text + " hasta el " + txthastafecha.Text

                    txtClave = txtTipomovi + txtNroInterno + txtRenglon
                    txtClaveOrd = txtTipomovi + txtNroInterno

                    SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                  txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)

                End If

                If Val(CampoRecibos.renglon) = 1 And Val(CampoRecibos.retsuss) <> 0 Then

                    txtRenglonII = txtRenglonII + 1
                    txtCuenta = "145"

                    txtTipomovi = "3"
                    txtNroInterno = CampoRecibos.recibo
                    txtProveedor = ""
                    txtTipo = ""
                    txtLetra = ""
                    txtPunto = ""
                    txtNumero = ""
                    txtRenglon = txtRenglonII
                    txtFecha = CampoRecibos.fecha
                    txtObservaciones = ""
                    txtDebito = CampoRecibos.retsuss
                    txtCredito = 0
                    txtFechaOrd = CampoRecibos.fechaord
                    txtTitulo = "Recibos"
                    txtEmpresa = 1
                    txtTituloList = "Surfactan S.A."

                    If Proceso._EsPellital() Then
                        txtTituloList = "Pellital S.A."
                    End If

                    txtVarios = "Desde el " + txtDesdeFecha.Text + " hasta el " + txthastafecha.Text

                    txtClave = txtTipomovi + txtNroInterno + txtRenglon
                    txtClaveOrd = txtTipomovi + txtNroInterno

                    SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                  txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)

                End If

            Next

            '
            ' Buscamos los datos de las retenciones que no estaban implementadas y las discrimamos por cuentas.
            '
            Try
                Dim WRetIBs As DataTable = _BuscarRetencionesIBRecibos(Proceso.ordenaFecha(txtDesdeFecha.Text), Proceso.ordenaFecha(txthastafecha.Text))
                Dim WRecibo = "", WFecha = "", WImporte = ""

                If Not IsNothing(WRetIBs) Then
                    For Each ret As DataRow In WRetIBs.Rows

                        With ret

                            WRecibo = IIf(IsDBNull(.Item("Recibo")), "0", Trim(.Item("Recibo")))
                            WFecha = IIf(IsDBNull(.Item("Fecha")), "0", Trim(.Item("Fecha")))

                            For i = 1 To 14
                                If IsDBNull(.Item("RetIb" & i)) Then
                                    ZRetIb(i, 1) = "0"
                                Else
                                    ZRetIb(i, 1) = Proceso.formatonumerico(.Item("RetIb" & i))
                                End If
                            Next

                        End With

                        For i = 1 To 14

                            If Val(ZRetIb(i, 1)) <> 0 AndAlso Val(WRecibo) <> 0 Then

                                txtRenglonII = txtRenglonII + 1
                                txtCuenta = ZRetIb(i, 2)

                                WImporte = Proceso.formatonumerico(ZRetIb(i, 1))

                                txtTipomovi = "3"
                                txtNroInterno = WRecibo
                                txtProveedor = ""
                                txtTipo = ""
                                txtLetra = ""
                                txtPunto = ""
                                txtNumero = ""
                                txtRenglon = txtRenglonII
                                txtFecha = WFecha
                                txtObservaciones = ""
                                txtDebito = Val(WImporte)
                                txtCredito = 0
                                txtFechaOrd = Proceso.ordenaFecha(WFecha)
                                txtTitulo = "Recibos"
                                txtEmpresa = 1
                                txtTituloList = "Surfactan S.A."

                                If Proceso._EsPellital() Then
                                    txtTituloList = "Pellital S.A."
                                End If

                                txtVarios = "Desde el " + txtDesdeFecha.Text + " hasta el " + txthastafecha.Text

                                txtClave = txtTipomovi + txtNroInterno + txtRenglon
                                txtClaveOrd = txtTipomovi + txtNroInterno

                                SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                                              txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)


                            End If

                        Next

                    Next
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Sub
            End Try

        End If



        txtCorte = ""
        txtRenglonII = 0

        If chkDepositos.Checked = True Then

            Dim tabla As DataTable
            tabla = SQLConnector.retrieveDataTable("buscar_depositos_fecha", txtDesdefechaOrd, txtHastafechaOrd)

            For Each row As DataRow In tabla.Rows

                Dim CampoDepositos As New LeeDepositos(row.Item(0), row.Item(1), row.Item(2),
                                               row.Item(3), row.Item(4), row.Item(5),
                                               row.Item(6))


                txtRenglonII = txtRenglonII + 1

                txtTipomovi = "2"
                txtNroInterno = CampoDepositos.deposito
                txtProveedor = ""
                txtTipo = CampoDepositos.tipo2
                txtLetra = ""
                txtPunto = ""
                txtNumero = CampoDepositos.numero2
                txtRenglon = txtRenglonII
                txtFecha = CampoDepositos.Fecha
                txtObservaciones = ""
                txtCuenta = txtVectorBanco(CampoDepositos.Banco)
                txtDebito = CampoDepositos.importe2
                txtCredito = 0
                txtFechaOrd = CampoDepositos.fechaord
                txtTitulo = "Deposito"
                txtEmpresa = 1
                txtTituloList = "Surfactan S.A."

                If Proceso._EsPellital() Then
                    txtTituloList = "Pellital S.A."
                End If

                txtVarios = "Desde el " + txtDesdeFecha.Text + " hasta el " + txthastafecha.Text

                txtClave = txtTipomovi + txtNroInterno + txtRenglon
                txtClaveOrd = txtTipomovi + txtNroInterno

                SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                              txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)



                Select Case Val(CampoDepositos.tipo2)
                    Case 2
                        txtCuenta = "2"
                    Case 3
                        txtCuenta = "40"
                    Case Else
                        txtCuenta = "1"
                End Select


                txtRenglonII = txtRenglonII + 1

                txtTipomovi = "2"
                txtNroInterno = CampoDepositos.deposito
                txtProveedor = ""
                txtTipo = CampoDepositos.tipo2
                txtLetra = ""
                txtPunto = ""
                txtNumero = CampoDepositos.numero2
                txtRenglon = txtRenglonII
                txtFecha = CampoDepositos.Fecha
                txtObservaciones = ""
                txtDebito = 0
                txtCredito = CampoDepositos.importe2
                txtFechaOrd = CampoDepositos.fechaord
                txtTitulo = "Deposito"
                txtEmpresa = 1
                txtTituloList = "Surfactan S.A."

                If Proceso._EsPellital() Then
                    txtTituloList = "Pellital S.A."
                End If

                txtVarios = "Desde el " + txtDesdeFecha.Text + " hasta el " + txthastafecha.Text

                txtClave = txtTipomovi + txtNroInterno + txtRenglon
                txtClaveOrd = txtTipomovi + txtNroInterno

                SQLConnector.executeProcedure("alta_impcyb", txtClave, txtTipomovi, txtNroInterno, txtProveedor, txtTipo, txtLetra, txtPunto, txtNumero,
                                              txtRenglon, txtFecha, txtObservaciones, txtCuenta, txtCredito, txtDebito, txtFechaOrd, txtTitulo, txtEmpresa, txtTituloList, txtVarios, txtClaveOrd)

            Next



        End If


        txtUno = "{Impcyb.Cuenta} in " + x + txtDesdeCuenta.Text + x + " to " + x + txtHastaCuenta.Text + x
        txtFormula = txtUno


        Select Case TipoListado.SelectedIndex
            Case 0
                viewer = New ReportViewer("Imputaciones Contables", Globals.reportPathWithName("wImpCybnet.rpt"), txtFormula)

            Case Else
                viewer = New ReportViewer("Imputaciones Contables", Globals.reportPathWithName("wImpCybResunet.rpt"), txtFormula)

        End Select

        If IsNothing(viewer) Then : Exit Sub : End If

        Select Case TipoImpresion
            Case Reporte.Imprimir
                viewer.imprimirReporte()
            Case Reporte.Pantalla
                viewer.Show()
        End Select

    End Sub

    Private Function _BuscarRetencionesIBRecibos(ByVal WDesde As String, ByVal WHasta As String) As DataTable

        Dim cn = New SqlConnection()
        Dim cm = New SqlCommand("SELECT Recibo, Fecha, RetIb1, RetIb2, RetIb3, RetIb4, RetIb5, RetIb6, RetIb7, RetIb8, RetIb9, RetIb10, RetIb11, RetIb12, RetIb13, RetIb14 FROM Recibos WHERE (Renglon = '01' OR Renglon = '1') AND FechaOrd BETWEEN '" & WDesde & "' AND '" & WHasta & "'")
        Dim dr As SqlDataReader
        Dim tabla As New DataTable

        Try

            cn.ConnectionString = Proceso._ConectarA
            cn.Open()
            cm.Connection = cn

            dr = cm.ExecuteReader()

            If dr.HasRows Then

                tabla.Load(dr)

                Return tabla
            Else
                Return Nothing

            End If

        Catch ex As Exception
            Throw New Exception("Hubo un problema al querer consultar los montos y cuentas de IB desde la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

    End Function


    Private Sub txtDesdeCuenta_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtDesdeCuenta.MouseDoubleClick, txtHastaCuenta.MouseDoubleClick
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

    '    Private Sub lstFiltrada_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
    '        Dim origen As ListBox = lstAyuda
    '        Dim filtrado As ListBox = lstFiltrada
    '        Dim texto As TextBox = txtAyuda
    '
    '        ' Buscamos el texto exacto del item seleccionado y seleccionamos el mismo item segun su indice en la lista de origen.
    '        origen.SelectedIndex = origen.FindStringExact(filtrado.SelectedItem.ToString)
    '
    '        ' Llamamos al evento que tenga asosiado el control de origen.
    '        lstAyuda_Click(Nothing, Nothing)
    '
    '
    '        ' Sacamos de vista los resultados filtrados.
    '        filtrado.Visible = False
    '        texto.Text = ""
    '    End Sub

    Private Sub txtAyuda_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAyuda.TextChanged
        _FiltrarDinamicamente()
    End Sub

    Private Sub SoloNumero(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDesdeCuenta.KeyPress, txtHastaCuenta.KeyPress
        If Not Char.IsNumber(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnPantalla_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPantalla.Click
        _Imprimir(Reporte.Pantalla)
    End Sub

    Private Sub btnImprimir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImprimir.Click
        _Imprimir(Reporte.Imprimir)
    End Sub

    Private Sub ListadoImputacionesContable_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        txtDesdeFecha.Focus()
    End Sub
End Class