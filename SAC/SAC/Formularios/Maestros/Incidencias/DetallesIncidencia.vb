﻿Imports System.Configuration
Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO

Public Class DetallesIncidencia : Implements IAuxiNuevaSACDesdeINC, IAyudaListadoSACs

    Const TextoBtnGenerarSac = "Generar/Asociar SAC"
    Const TextoBtnVerSac = "Ver SAC Asociado"
    Private WControlRetornoError As Control = Nothing
    Private WEmpresaProd As Integer = 0

    Private Const EXTENSIONES_PERMITIDAS = "*.bmp|*.png|*.jpg|*.jpeg|*.pdf|*.doc|*.docx|*.xls|*.xlsx"

    Public Property MostrarBotonVerSac As Boolean = True

    Sub New(Optional ByVal WNroIncidencia As Object = "")

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        txtIncidencia.Text = WNroIncidencia
    End Sub

    Private Sub DetallesIncidencia_Shown(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Shown
        txtIncidencia.Focus()
    End Sub

    Private Sub DetallesIncidencia_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        txtDescProducto.BackColor = WBackColorTerciario

        If txtIncidencia.Text.Trim <> "" Then
            txtIncidencia_KeyDown(Nothing, New KeyEventArgs(Keys.Enter))
        Else
            _Limpiar()
        End If
    End Sub

    Private Sub _Limpiar()

        For Each c As Control In {txtDescProducto, txtFecha, txtIncidencia, txtLotePartida, txtPosiblesUsos, txtProducto, txtReferencia, txtTitulo, txtMotivos}
            c.Text = ""
        Next

        WEmpresaProd = 0

        _CargarEstados()

        txtFecha.Text = Date.Now.ToString("dd/MM/yyyy")

        btnControles.Enabled = True
        btnHojaProduccion.Enabled = True

        rbMatPrima.Checked = True
        rbProdTerminado_Click(Nothing, Nothing)
        cmbEstado.SelectedIndex = 0
        btnSac.Text = TextoBtnGenerarSac '"Generar/Asociar SAC"

        Dim WIncidencia As Integer = 0

        Dim WIncid As DataRow = GetSingle("SELECT Max(Incidencia) Ultimo FROM CargaIncidencias")

        If Not IsNothing(WIncid) Then WIncidencia = OrDefault(WIncid.Item("Ultimo"), 0)

        WIncidencia += 1

        txtIncidencia.Text = WIncidencia

        TabControl1.SelectTab(0)

        btnSac.Enabled = MostrarBotonVerSac
        txtIncidencia.Enabled = MostrarBotonVerSac

    End Sub

    Private Sub _CargarEstados()
        Dim WEstados As DataTable = GetAll("SELECT Estado, LTRIM(Descripcion) Descripcion FROM EstadosINC ORDER BY Estado")

        With cmbEstado
            .DataSource = WEstados
            .DisplayMember = "Descripcion"
            .ValueMember = "Estado"
            If .Items.Count > 0 Then .SelectedIndex = 0
        End With

    End Sub

    Private Sub rbProdTerminado_Click(ByVal sender As Object, ByVal e As EventArgs) Handles rbProdTerminado.Click, rbMatPrima.Click
        With txtProducto
            Dim WProducto = .Text
            .Mask = IIf(rbMatPrima.Checked, ">LL-000-000", ">LL-00000-000")
            .Text = WProducto
            .Focus()
            .SelectionStart = 0
            .SelectionLength = .Text.Length

            btnControles.Enabled = Val(txtLotePartida.Text) <> 0
            btnHojaProduccion.Enabled = rbProdTerminado.Checked

        End With
    End Sub

    Private Sub txtIncidencia_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtIncidencia.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtIncidencia.Text) = "" Then : Exit Sub : End If

            Dim WIncid = txtIncidencia.Text

            _Limpiar()

            txtIncidencia.Text = WIncid

            ' Busco la incidencia en caso de que se indique una.
            If Val(WIncid) <> 0 Then

                Dim WIncidencia As DataRow = GetSingle("SELECT * FROM CargaIncidencias WHERE Incidencia = '" & txtIncidencia.Text & "'")

                If WIncidencia IsNot Nothing Then
                    With WIncidencia

                        Dim WTipo = OrDefault(.Item("Tipo"), TipoIncidencias.General)

                        If WTipo = TipoIncidencias.RechazoMP Then
                            Dim frm As New DetallesIncidenciaRechazoMP(.Item("Incidencia"))

                            frm.Show(Owner)

                            Close()
                        End If

                        txtFecha.Text = OrDefault(.Item("Fecha"), "")

                        Dim WTempEstado = OrDefault(.Item("Estado"), 0)

                        For Each rowView As DataRowView In cmbEstado.Items
                            If rowView.Item("Estado") = WTempEstado Then
                                cmbEstado.SelectedItem = rowView
                                Exit For
                            End If
                        Next

                        txtLotePartida.Text = OrDefault(.Item("Lote"), "")
                        rbProdTerminado.Checked = OrDefault(.Item("TipoProd"), "M") = "T"
                        rbProdTerminado_Click(Nothing, Nothing)
                        txtProducto.Text = OrDefault(.Item("Producto"), "")
                        txtProducto_KeyDown(Nothing, New KeyEventArgs(Keys.Enter))
                        txtTitulo.Text = OrDefault(.Item("Titulo"), "")
                        txtReferencia.Text = OrDefault(.Item("Referencia"), "")
                        txtPosiblesUsos.Text = OrDefault(.Item("PosiblesUsos"), "")
                        txtMotivos.Text = OrDefault(.Item("Motivos"), "")
                        btnSac.Text = IIf(Val(OrDefault(.Item("ClaveSac"), "")) = 0, TextoBtnGenerarSac, TextoBtnVerSac)
                    End With

                    _CargarArchivosRelacionados()

                End If

            End If

            For Each c As Control In {txtDescProducto, txtFecha, txtIncidencia, txtLotePartida, txtPosiblesUsos, txtProducto, txtReferencia, txtTitulo}
                c.Text = c.Text.Trim
            Next

            txtFecha.Focus()

        ElseIf e.KeyData = Keys.Escape Then
            txtIncidencia.Text = ""
        End If

    End Sub

    Private Sub txtFecha_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtFecha.KeyDown

        If e.KeyData = Keys.Enter Then
            If txtFecha.Text.Replace(" ", "").Length < 10 Then : Exit Sub : End If

            If _ValidarFecha(txtFecha.Text) Then

                With cmbEstado
                    .Focus()
                    .DroppedDown = True
                End With

            End If

        ElseIf e.KeyData = Keys.Escape Then
            txtFecha.Text = ""
        End If

    End Sub

    Private Sub cmbEstado_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles cmbEstado.KeyDown

        If e.KeyData = Keys.Enter Then
            If Trim(cmbEstado.Text) = "" Then : Exit Sub : End If

            txtProducto.Focus()

        ElseIf e.KeyData = Keys.Escape Then
            cmbEstado.SelectedIndex = 0
        End If

    End Sub

    Private Sub cmbEstado_DropDownClosed(ByVal sender As Object, ByVal e As EventArgs) Handles cmbEstado.DropDownClosed
        cmbEstado_KeyDown(Nothing, New KeyEventArgs(Keys.Enter))
    End Sub

    Private Sub txtProducto_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtProducto.KeyDown

        If e.KeyData = Keys.Enter Then
            If rbMatPrima.Checked And txtProducto.Text.Replace(" ", "").Length < 10 Then : Exit Sub : End If
            If rbProdTerminado.Checked And txtProducto.Text.Replace(" ", "").Length < 12 Then : Exit Sub : End If

            txtDescProducto.Text = ""

            Dim WProd As DataRow = Nothing

            If rbMatPrima.Checked Then
                WProd = GetSingle("SELECT Descripcion FROM Articulo WHERE Codigo = '" & txtProducto.Text & "'")
            Else
                WProd = GetSingle("SELECT Descripcion FROM Terminado WHERE Codigo = '" & txtProducto.Text & "'")
            End If

            If IsNothing(WProd) Then Exit Sub

            txtDescProducto.Text = OrDefault(WProd.Item("Descripcion"), "").ToString.Trim

            btnControles.Enabled = Val(txtLotePartida.Text) <> 0
            btnHojaProduccion.Enabled = rbProdTerminado.Checked

            txtLotePartida.Focus()

        ElseIf e.KeyData = Keys.Escape Then
            txtProducto.Text = ""
        End If

    End Sub

    Private Sub txtLotePartida_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtLotePartida.KeyDown

        If e.KeyData = Keys.Enter Then
            If Val(txtLotePartida.Text) = 0 Then : Exit Sub : End If

            Try
                _CorroborarCorrespondenciaProductoLotePartida()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation)
                txtLotePartida.Focus()
            End Try

        ElseIf e.KeyData = Keys.Escape Then
            txtLotePartida.Text = ""
        End If

    End Sub

    Private Sub _CorroborarCorrespondenciaProductoLotePartida()

        Dim WProd As DataRow = _ObtenerProductoAsociado()

        If IsNothing(WProd) Then Throw New Exception("No se encuentra Laudo/Hoja indicada.")

        Dim WProductoInfLength As String = txtProducto.Text.Replace(" ", "").Length

        If (rbMatPrima.Checked And WProductoInfLength = 10) _
           Or (rbProdTerminado.Checked And WProductoInfLength = 12) Then

            If txtProducto.Text.ToUpper <> OrDefault(WProd.Item("Producto"), "") Then
                If MsgBox("El Lote/Partida que se indicó, tiene asociado un Código de Producto distinto al informado en el Campo Producto de este formulario." _
                          & vbCrLf _
                          & vbCrLf _
                          & "Producto Indicado: " & txtProducto.Text _
                          & vbCrLf _
                          & "Producto Asociado Lote/Partida: " & WProd.Item("Producto") _
                          & vbCrLf _
                          & vbCrLf _
                          & "¿Desea modificarlo por el asociado al Lote/Partida?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                    txtProducto.Text = OrDefault(WProd.Item("Producto"), "")
                Else
                    txtLotePartida.Focus()
                    Exit Sub
                End If
            End If

        Else
            txtProducto.Text = OrDefault(WProd.Item("Producto"), "")
        End If

        txtProducto_KeyDown(Nothing, New KeyEventArgs(Keys.Enter))

        txtTitulo.Focus()

    End Sub

    Private Function _ObtenerProductoAsociado() As DataRow

        Dim WProd As DataRow = Nothing

        Dim WEmpresas As New List(Of String) From {"SurfactanSa", "Surfactan_II", "Surfactan_III", "Surfactan_IV" _
                , "Surfactan_V", "Surfactan_VI", "Surfactan_VII"}

        If _EsPellital() Then
            WEmpresas = New List(Of String) From {"PellitalSa", "Pelitall_II", "Pellital_III", "Pellital_V"}
        End If

        For Each emp As String In WEmpresas

            If rbMatPrima.Checked Then
                WProd = GetSingle("SELECT Articulo As Producto FROM Laudo WHERE Laudo = '" & txtLotePartida.Text & "' And Renglon IN ('1', '01')", emp)
            Else
                WProd = GetSingle("SELECT Producto FROM Hoja WHERE Hoja = '" & txtLotePartida.Text & "' And Renglon IN ('1', '01')", emp)
            End If

            If WProd IsNot Nothing Then
                WEmpresaProd = _IdEmpresaSegunBase(emp)
                Exit For
            End If

        Next

        Return WProd
    End Function

    Private Sub txtTitulo_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtTitulo.KeyDown

        If e.KeyData = Keys.Enter Then
            If Trim(txtTitulo.Text) = "" Then : Exit Sub : End If

            txtReferencia.Focus()

        ElseIf e.KeyData = Keys.Escape Then
            txtTitulo.Text = ""
        End If

    End Sub

    Private Sub txtReferencia_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtReferencia.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtReferencia.Text) = "" Then : Exit Sub : End If

            With txtPosiblesUsos
                .Focus()
                .SelectionStart = .Text.Length
                .SelectionLength = .Text.Length
            End With

        ElseIf e.KeyData = Keys.Escape Then
            txtReferencia.Text = ""
        End If

    End Sub

    Private Sub btnCerrar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCerrar.Click

        Dim WOwner As INuevaIncidencia = CType(Owner, INuevaIncidencia)

        If WOwner IsNot Nothing Then WOwner._ProcesarNuevaIncidencia(txtIncidencia.Text)

        Close()
    End Sub

    Private Sub btnGrabar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnGrabar.Click
        Try
            WEmpresaProd = 0

            _ValidarDatosIngresados()

            Dim WSqls As New List(Of String)

            Dim WIncid As DataRow = GetSingle("SELECT Incidencia, Tipo, ClaveSac FROM CargaIncidencias WHERE Incidencia = '" & txtIncidencia.Text & "'")

            Dim WTipo = 1
            Dim WClaveSAC = ""

            If IsNothing(WIncid) Then
                txtIncidencia.Text = ""
            Else
                WTipo = OrDefault(WIncid.Item("Tipo"), 1)
                WClaveSAC = OrDefault(WIncid.Item("ClaveSac"), "")
            End If

            If Val(txtIncidencia.Text) = 0 Then
                WIncid = GetSingle("SELECT Max(Incidencia) As Ultimo FROM CargaIncidencias")
                If Not IsNothing(WIncid) Then
                    txtIncidencia.Text = OrDefault(WIncid.Item("Ultimo"), 0)
                End If
                txtIncidencia.Text = Val(txtIncidencia.Text) + 1
            End If

            Dim WEstado As Integer = CType(cmbEstado.SelectedItem, DataRowView).Item("Estado")

            WSqls.Add("DELETE CargaIncidencias WHERE Incidencia = '" & txtIncidencia.Text & "'")

            Dim ZSql = String.Format("INSERT INTO CargaIncidencias " _
                       & "(Incidencia, Renglon, Tipo, Fecha, FechaOrd, Estado, Titulo, Referencia, Producto, Lote, ClaveSac, TipoProd, Posiblesusos, Motivos, Empresa) " _
                       & "VALUES " _
                       & " ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}') ", _
                       txtIncidencia.Text, 1, WTipo, txtFecha.Text, ordenaFecha(txtFecha.Text), WEstado, txtTitulo.Text, txtReferencia.Text, txtProducto.Text, txtLotePartida.Text, WClaveSAC, IIf(rbMatPrima.Checked, "M", "T"), txtPosiblesUsos.Text, txtMotivos.Text, WEmpresaProd)
            WSqls.Add(ZSql)

            ExecuteNonQueries(WSqls.ToArray)

            If ContinuarSalirMsgBox.Show("Actualización se ha realizado con Éxito" & vbCrLf _
                                         & "Indique como quiere proseguir.", "Continuar editando Informe de No Conformidad", "Volver a Listado") = DialogResult.OK Then
                txtTitulo.Focus()
                Exit Sub
            End If

            btnCerrar.PerformClick()

        Catch ex As Exception
            If Trim(ex.Message) <> "" Then MsgBox(ex.Message, MsgBoxStyle.Exclamation)
            If WControlRetornoError IsNot Nothing Then WControlRetornoError.Focus()
        End Try
    End Sub

    Private Sub _ValidarDatosIngresados()

        WControlRetornoError = Nothing

        If txtFecha.Text.Replace(" ", "").Length < 10 OrElse Not _ValidarFecha(txtFecha.Text) Then
            WControlRetornoError = txtFecha
            Throw New Exception("Debe indicarse una fecha válida.")
        End If

        If cmbEstado.SelectedIndex < 0 Then
            WControlRetornoError = cmbEstado
            Throw New Exception("Debe indicarse un estado válido para este Informe de No Conformidad.")
        End If

        If txtProducto.Text.Replace(" ", "").Length > 2 Then

            If rbMatPrima.Checked And txtProducto.Text.Replace(" ", "").Length < 10 Then Throw New Exception("Debe cargar una Materia Prima válida")
            If rbProdTerminado.Checked And txtProducto.Text.Replace(" ", "").Length < 12 Then Throw New Exception("Debe cargar un Producto Terminado válido")

            If Val(txtLotePartida.Text) = 0 Then Throw New Exception("Se debe indicar un Lote/Partida válida para el Producto indicado.")

            Dim WProd As DataRow = _ObtenerProductoAsociado()

            If WProd Is Nothing Then
                WControlRetornoError = txtProducto
                Throw New Exception("No se ha podido validar la correspondencia entre el Producto y la Partida informados.")
            End If

            If txtProducto.Text.ToUpper <> OrDefault(WProd.Item("Producto"), "") Then
                If MsgBox("El Lote/Partida que se indicó, tiene asociado un Código de Producto distinto al informado en el Campo Producto de este formulario." _
                          & vbCrLf _
                          & vbCrLf _
                          & "Producto Indicado: " & txtProducto.Text _
                          & vbCrLf _
                          & "Producto Asociado Lote/Partida: " & WProd.Item("Producto") _
                          & vbCrLf _
                          & vbCrLf _
                          & "¿Desea modificarlo por el asociado al Lote/Partida?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                    txtProducto.Text = OrDefault(WProd.Item("Producto"), "")
                Else
                    WControlRetornoError = txtProducto
                    Throw New Exception("El Producto indicado no se corresponde con el informado en la Hoja/Laudo.")
                End If

            End If
        Else

            If Val(txtLotePartida.Text) <> 0 Then
                Dim WProd As DataRow = _ObtenerProductoAsociado()

                If WProd IsNot Nothing Then txtProducto.Text = OrDefault(WProd.Item("Producto"), "")
            End If

        End If
        
        If txtTitulo.Text.Trim = "" Then
            WControlRetornoError = txtTitulo
            If MsgBox("No ha agregado nigún Título para este Informe de No Conformidad ¿Quiere proseguir con la grabación del mismo?", MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes Then
                Throw New Exception("")
            End If

        End If
        If txtReferencia.Text.Trim = "" Then
            WControlRetornoError = txtReferencia
            If MsgBox("No ha agregado niguna Referencia para este Informe de No Conformidad ¿Quiere proseguir con la grabación del mismo?", MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes Then
                Throw New Exception("")
            End If
        End If

    End Sub

    Private Sub SoloNumero(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles txtIncidencia.KeyPress
        If Not Char.IsNumber(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnControles_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnControles.Click

        If Val(txtLotePartida.Text) = 0 Then Exit Sub

        '
        ' Buscamos los ensayos guardados según corresponda.
        '
        If rbMatPrima.Checked Then
            _MostrarEnsayosMP()
        Else
            _MostrarEnsayosPT()
        End If

    End Sub

    Private Sub _MostrarEnsayosPT()
        If Val(txtLotePartida.Text) = 0 Then Exit Sub

        With New DetallesEnsayosPT(txtLotePartida.Text)
            .Show(Me)
        End With

    End Sub

    Private Sub _MostrarEnsayosMP()
        If Val(txtLotePartida.Text) = 0 Then Exit Sub

        With New DetallesEnsayosMP(txtLotePartida.Text)
            .Show(Me)
        End With

    End Sub

    Private Sub txtLotePartida_Enter(ByVal sender As Object, ByVal e As EventArgs) Handles txtLotePartida.Enter
        With txtLotePartida
            .Text = .Text.Trim
            .SelectionStart = 0
            .SelectionLength = .Text.Length
        End With
    End Sub

    Private Sub btnHojaProduccion_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnHojaProduccion.Click
        With New DetallesHojaProduccion(txtLotePartida.Text)
            .Show(Me)
        End With
    End Sub

    Private Sub btnMovimientos_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnMovimientos.Click
        If rbMatPrima.Checked Then
            With New DetalleMovimientosMP(txtLotePartida.Text)
                .Show(Me)
            End With
        Else
            With New DetalleMovimientosPT(txtLotePartida.Text)
                .Show(Me)
            End With
        End If
    End Sub

    Private Sub btnSac_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSac.Click

        Try
            Dim WNC As DataRow = GetSingle("SELECT ClaveSac FROM CargaIncidencias WHERE Incidencia = '" & txtIncidencia.Text & "' And Renglon = 1")

            If WNC Is Nothing Then Exit Sub

            Dim WClaveSAC As String = Trim(OrDefault(WNC.Item("ClaveSac"), ""))

            If WClaveSAC = "" Then
                '
                ' Consultamos si quieren abrir nueva Sac o asignar una ya abierta.
                '
                Dim WPregunta = New ContinuarSalirMsgBox("¿Como desea asignarle una SAC a este Informe de No Conformidad?",
                                              "Abrir Nueva SAC", "Asignar SAC ya Abierta")
                WPregunta._DrBtn1 = Windows.Forms.DialogResult.Yes
                WPregunta._DrBtn2 = Windows.Forms.DialogResult.No

                Dim Resp As DialogResult = WPregunta.ShowDialog(Me)

                Select Case Resp
                    Case DialogResult.Yes
                        '
                        ' Creamos nueva SAC, colocando como datos por Default los provistos por el actual
                        ' Informe de No Conformidad.
                        '
                        Dim frm As New AuxiNuevaSACDesdeINC(txtIncidencia.Text, txtFecha.Text, txtTitulo.Text, txtReferencia.Text, txtMotivos.Text)
                        frm.Show(Me)

                    Case DialogResult.No
                        '
                        ' Mostramos Listado con todas las SAC's Disponibles.
                        '
                        Dim frm As New AyudaListadoSACs()
                        frm.Show(Me)

                End Select

            Else
                _AbrirSACPorClave(WClaveSAC)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub _AbrirSACPorClave(ByVal WClaveSAC As String)

        Dim pattern = "(?<Tipo>[0-9]{4})(?<Anio>[0-9]{4})(?<Numero>[0-9]{6})"

        If Regex.IsMatch(WClaveSAC, pattern) Then
            Dim matches As MatchCollection = Regex.Matches(WClaveSAC, pattern)

            Dim WTipo = matches.Item(0).Groups("Tipo").Value
            Dim WAnio = matches.Item(0).Groups("Anio").Value
            Dim WNum = matches.Item(0).Groups("Numero").Value

            Dim frm As New NuevoSac(WTipo, WNum, WAnio, True)
            frm.Show(Me)
        End If

    End Sub

    Public Sub _ProcesarNuevaSACDesdeINC(ByVal WClaveSAC As String) Implements IAuxiNuevaSACDesdeINC._ProcesarNuevaSACDesdeINC
        If WClaveSAC.Trim = "" Then Exit Sub

        btnSac.Text = TextoBtnVerSac

        WClaveSAC = WClaveSAC.PadLeft(14, "0")

        Dim WAnio = WClaveSAC.Substring(4, 4)
        Dim WNum = WClaveSAC.Substring(8, 6)

        _CopiarArchivosINCASAC(WClaveSAC)

        MsgBox(String.Format("Se ha generado SAC Bajo la siguiente referencia {0} {2} / {1}", "", WAnio, WNum), MsgBoxStyle.Information)

        If MsgBox("¿Desea Abrir la SAC generada?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            _AbrirSACPorClave(WClaveSAC)
            Exit Sub
        End If



        Close()

    End Sub

    Private Sub _CopiarArchivosINCASAC(ByVal WClaveSAC As String)

        Dim WDirectorioRaiz = ConfigurationManager.AppSettings("ARCHIVOS_RELACIONADOS")

        Dim WDIrRaizINC As String = WDirectorioRaiz & "Informes No Conformidad\INC_" & txtIncidencia.Text
        Dim WDIrRaizSAC As String = WDirectorioRaiz & "SAC_" & WClaveSAC

        If Not Directory.Exists(WDIrRaizINC) Then Directory.CreateDirectory(WDIrRaizINC)
        If Not Directory.Exists(WDIrRaizSAC) Then Directory.CreateDirectory(WDIrRaizSAC)

        For Each f As String In Directory.GetFiles(WDIrRaizINC)
            If Not File.Exists(WDIrRaizSAC & "\" & Path.GetFileName(f)) Then
                File.Copy(f, WDIrRaizSAC & "\" & Path.GetFileName(f))
            End If

        Next
    End Sub

    Public Sub _ProcesarAyudaListadoSACs(ByVal WClaveSAC As String) Implements IAyudaListadoSACs._ProcesarAyudaListadoSACs
        If WClaveSAC.Trim = "" Then Exit Sub

        btnSac.Text = TextoBtnVerSac

        WClaveSAC = WClaveSAC.PadLeft(14, "0")

        Dim WAnio = WClaveSAC.Substring(4, 4)
        Dim WNum = WClaveSAC.Substring(8, 6)

        ExecuteNonQueries("UPDATE CargaIncidencias SET ClaveSac = '" & WClaveSAC & "' WHERE Incidencia = '" & txtIncidencia.Text & "'")

        _CopiarArchivosINCASAC(WClaveSAC)

        MsgBox(String.Format("Se ha Asociado la siguiente SAC {0} {2} / {1} al Informe de No Conformidad Actual.", "", WAnio, WNum), MsgBoxStyle.Information)

        If MsgBox("¿Desea Abrir la SAC asociada?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            _AbrirSACPorClave(WClaveSAC)
            Exit Sub
        End If

        Close()
    End Sub

#Region "Rutina Archivos"


    '
    ' Rutinas de Archivos.
    '
    Private Sub dgvArchivos_CellMouseDoubleClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles dgvArchivos.CellMouseDoubleClick
        If e.RowIndex < 0 Then Exit Sub
        With dgvArchivos.Rows(e.RowIndex)
            If Not IsNothing(.Cells("PathArchivo").Value) Then

                Try
                    Process.Start(.Cells("PathArchivo").Value, "f")
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

            End If
        End With
    End Sub


    Private Sub dgvArchivos_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) Handles dgvArchivos.DragEnter
        _PermitirDrag(e)
    End Sub

    Private Sub _PermitirDrag(ByVal e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub _ProcesarDragDeArchivo(ByVal e As DragEventArgs)
        Dim archivos() As String = e.Data.GetData(DataFormats.FileDrop)
        _SubirArchvios(archivos)
    End Sub

    Private Sub _SubirArchvios(ByVal archivos As String())
        Dim WNombreCarpetaArchivos As String = "INC_" & Trim(Str$(txtIncidencia.Text))
        Dim WRutaArchivosRelacionados = _RutaCarpetaArchivos() & "\" & WNombreCarpetaArchivos

        Dim WDestino = ""
        Dim WCantCorrectas = 0

        If archivos.Length = 0 Then : Return : End If

        If Not Directory.Exists(WRutaArchivosRelacionados) Then
            Try
                Directory.CreateDirectory(WRutaArchivosRelacionados)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation)
                Return
            End Try
        End If

        For Each archivo In archivos

            If File.Exists(archivo) Then

                If EXTENSIONES_PERMITIDAS.Contains(Path.GetExtension(archivo).ToLower()) Then

                    WDestino = WRutaArchivosRelacionados & "\" & Path.GetFileName(archivo)

                    Try
                        If Not File.Exists(WDestino) Then
                            File.Copy(archivo, WDestino)
                            WCantCorrectas += 1
                        Else
                            If MsgBox("El Archivo " & Path.GetFileName(archivo) & ", ya existe en la carpeta. ¿Desea sobreescribir el archivo existente?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                                File.Copy(archivo, WDestino, True)
                                WCantCorrectas += 1
                            End If
                        End If

                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Critical)
                        Return
                    End Try

                End If

            End If

        Next

        If WCantCorrectas > 0 Then
            MsgBox("Se subieron correctamente " & WCantCorrectas & " Archivo(s)", MsgBoxStyle.Information)
        End If

        _CargarArchivosRelacionados()
    End Sub

    Private Sub dgvArchivos_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles dgvArchivos.DragDrop
        _ProcesarDragDeArchivo(e)
    End Sub

    Private Function _RutaCarpetaArchivos()
        Return ConfigurationManager.AppSettings("ARCHIVOS_RELACIONADOSII")
    End Function

    Private Sub _CargarArchivosRelacionados()
        Dim WRutaArchivosRelacionados = ""
        Dim WNombreCarpetaArchivos As String = "INC_" & Trim(Str$(txtIncidencia.Text))

        If Not Directory.Exists(_RutaCarpetaArchivos) Then
            Throw New Exception("No se ha logrado tener acceso a la Carpeta Compartida de Archivos Relacionados.")
        End If

        WRutaArchivosRelacionados = _RutaCarpetaArchivos() & "\" & WNombreCarpetaArchivos

        ' Creamos la Carpeta en caso de que no exista aún.
        If Not Directory.Exists(WRutaArchivosRelacionados) Then
            Try
                Directory.CreateDirectory(WRutaArchivosRelacionados)
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End If

        Dim InfoArchivo As FileInfo

        dgvArchivos.Rows.Clear()

        ' Recorremos unicamente aquellos archivos que tengan una extensión que esté entre las permitidas por la aplicación.
        For Each WNombreArchivo As String In Directory.GetFiles(WRutaArchivosRelacionados).Where(Function(s) EXTENSIONES_PERMITIDAS.Contains(Path.GetExtension(s).ToLower()))

            InfoArchivo = FileSystem.GetFileInfo(WNombreArchivo)

            With InfoArchivo
                dgvArchivos.Rows.Add(.CreationTime.ToString("dd/MM/yyyy"), UCase(.Name), _ObtenerIconoSegunTipoArchivo(.Extension), .FullName)
            End With

        Next

    End Sub

    Private Function _ObtenerIconoSegunTipoArchivo(ByVal extension As String)
        Dim Wicono = Nothing

        Select Case UCase(extension)

            Case ".DOC", ".DOCX"
                Wicono = My.Resources.Word_icon
            Case ".XLS", ".XLSX", ".XLSM"
                Wicono = My.Resources.Excel_icon
            Case ".PDF"
                Wicono = My.Resources.pdf_icono
            Case ".JPG", ".JPEG", ".BMP", ".ICO", ".PNG"
                Wicono = My.Resources.imagen_icono
            Case ".TXT"
                Wicono = My.Resources.txt_icono
            Case Else
                Wicono = My.Resources.archivo_default
        End Select

        Return Wicono
    End Function

    Private Sub EliminarArchivoToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles EliminarArchivoToolStripMenuItem.Click
        If dgvArchivos.SelectedRows.Count = 0 Then Exit Sub

        If MsgBox("¿Está seguro de querer eliminar todos los archivos indicados?", MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes Then Exit Sub

        For Each row As DataGridViewRow In dgvArchivos.SelectedRows
            With row

                If OrDefault(.Cells("PathArchivo").Value, "") = "" Then Continue For

                File.Delete(.Cells("PathArchivo").Value)

            End With
        Next

        _CargarArchivosRelacionados()

    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        Try
            With OpenFileDialog1
                .Filter = "Imágenes (bmp, jpg, png) | " & String.Join(";", EXTENSIONES_PERMITIDAS.Split("|"))
                If .ShowDialog() = DialogResult.OK Then
                    Dim WArchivos() = .FileNames

                    If WArchivos.Length > 0 Then
                        _SubirArchvios(WArchivos)
                    End If

                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button2.Click
        Try
            If dgvArchivos.SelectedRows.Count > 0 Then
                EliminarArchivoToolStripMenuItem_Click(Nothing, Nothing)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

#End Region

End Class