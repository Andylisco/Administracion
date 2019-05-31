Imports System.Configuration
Imports System.IO
Imports System.Text.RegularExpressions
Imports ConsultasVarias
Imports CrystalDecisions.Shared
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.VisualBasic.FileIO

Public Class DetallesIncidenciaRechazoMP : Implements IAuxiNuevaSACDesdeINC, IAyudaListadoSACs, IAyudaMPAsociadasOC, IExportarINC, IModifNumeracionINC, IIngresoClaveSeguridad

    Const TextoBtnGenerarSac = "Asociar SAC"
    Const TextoBtnVerSac = "Ver SAC Asociado"
    Private WControlRetornoError As Control = Nothing

    Private WAutorizadoEliminar As Boolean = False

    Private Const EXTENSIONES_PERMITIDAS = "*.bmp|*.png|*.jpg|*.jpeg|*.pdf|*.doc|*.docx|*.xls|*.xlsx"

    Public Property MostrarBotonVerSac As Boolean = True
    Public Property WPrimeraCarga As Boolean = False

    Sub New(Optional ByVal WNroIncidencia As Object = "")

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        txtIncidencia.Text = WNroIncidencia
    End Sub

    Private Sub DetallesIncidencia_Shown(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Shown
        txtNumero.Focus()
    End Sub

    Private Sub DetallesIncidencia_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

        For Each l As TextBox In {txtDescProducto, txtDescProv, txtDescMP, txtDescMP, txtDescTipo}
            l.BackColor = WBackColorTerciario
        Next

        btnEliminar.Enabled = Val(Operador.CodigoResponsableSac) = 1
        btnModifNumINC.Enabled = Val(Operador.CodigoResponsableSac) = 1

        cmbEmpresa.Items.Clear()

        Dim WEmpresas As New DataTable

        With WEmpresas
            With .Columns
                .Add("Nombre")
                .Add("Id", GetType(Integer))
                .Add("Base")
            End With
            With .Rows
                .Add("SURFACTAN I", 1, "SurfactanSa")
                .Add("SURFACTAN II", 3, "Surfactan_II")
                .Add("SURFACTAN III", 5, "Surfactan_III")
                .Add("SURFACTAN IV", 6, "Surfactan_IV")
                .Add("SURFACTAN V", 7, "Surfactan_V")
                .Add("SURFACTAN VI", 10, "Surfactan_VI")
                .Add("SURFACTAN VII", 11, "Surfactan_VII")
                .Add("PELLITAL I", 2, "PellitalSa")
                .Add("PELLITAL II", 4, "Pelitall_II")
                .Add("PELLITAL III", 8, "Pellital_III")
                .Add("PELLITAL V", 9, "Pellital_V")
            End With
        End With

        With cmbEmpresa
            .DataSource = WEmpresas
            .DisplayMember = "Nombre"
            .ValueMember = "Id"
        End With

        _CargarEstados()

        If txtIncidencia.Text.Trim <> "" Then
            txtIncidencia_KeyDown(Nothing, New KeyEventArgs(Keys.Enter))
        Else
            _Limpiar()
        End If
    End Sub

    Private Sub _CargarEstados()
        Dim WEstados As DataTable = GetAll("SELECT Estado, LTrim(Descripcion) Descripcion FROM EstadosINC Order By Estado")

        With cmbEstado
            .DataSource = WEstados
            .DisplayMember = "Descripcion"
            .ValueMember = "Estado"
            If .Items.Count > 0 Then .SelectedIndex = 0
        End With
    End Sub

    Private Sub _Limpiar()

        For Each c As Control In {txtFecha, txtIncidencia, txtLotePartida, txtMotivos, txtProducto, txtReferencia, txtTitulo, txtOrden, txtProveedor, txtDescProducto, txtDescTipo, txtTipo, txtAnio, txtNumero, txtCantidadMP, txtDescProducto, txtDescMP, txtDescProv, txtDescTipo}
            c.Text = ""
        Next

        txtFecha.Text = Date.Now.ToString("dd/MM/yyyy")

        btnControles.Enabled = True
        btnHojaProduccion.Enabled = True

        rbMatPrima.Checked = True
        rbProdTerminado_Click(Nothing, Nothing)

        _CargarEstados()

        WAutorizadoEliminar = False

        cmbEstado.SelectedIndex = 0

        cmbEmpresa.SelectedIndex = 0

        btnSac.Text = TextoBtnGenerarSac '"Generar/Asociar SAC"

        Dim WIncidencia As Integer = 0

        Dim WIncid As DataRow = GetSingle("SELECT Max(Incidencia) Ultimo FROM CargaIncidencias")

        If Not IsNothing(WIncid) Then WIncidencia = OrDefault(WIncid.Item("Ultimo"), 0)

        WIncidencia += 1

        txtIncidencia.Text = WIncidencia

        TabControl1.SelectTab(0)

        btnSac.Enabled = MostrarBotonVerSac
        txtIncidencia.Enabled = MostrarBotonVerSac

        cmbEmpresaIncidencia.SelectedIndex = 0

        btnDesvincularSAC.Enabled = btnSac.Text = TextoBtnVerSac

        txtTipo.Text = "2"
        txtTipo_KeyDown(Nothing, New KeyEventArgs(Keys.Enter))

        txtAnio.Text = Date.Now.ToString("yyyy")

        WPrimeraCarga = True

    End Sub

    Private Sub rbProdTerminado_Click(ByVal sender As Object, ByVal e As EventArgs) Handles rbProdTerminado.Click
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

                        Dim WTipo = OrDefault(.Item("Tipo"), 2)

                        If WTipo = 1 Then
                            Dim frm As New DetallesIncidencia(.Item("Incidencia"))

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

                        'rbProdTerminado.Checked = OrDefault(.Item("TipoProd"), "M") = "T"
                        'rbProdTerminado_Click(Nothing, Nothing)
                        txtProducto.Text = OrDefault(.Item("Producto"), "")
                        txtAnio.Text = OrDefault(.Item("Ano"), "")
                        txtTipo.Text = OrDefault(.Item("TipoINC"), "")
                        txtNumero.Text = OrDefault(.Item("Numero"), "")

                        txtTipo_KeyDown(Nothing, New KeyEventArgs(Keys.Enter))

                        txtCantidadMP.Text = formatonumerico(OrDefault(.Item("CantOrden"), "0"))
                        'txtLotePartida.Text = OrDefault(.Item("Lote"), "")
                        'txtProducto_KeyDown(Nothing, New KeyEventArgs(Keys.Enter))
                        cmbEmpresaIncidencia.SelectedIndex = OrDefault(.Item("EmpresaIncidencia"), 0)
                        txtTitulo.Text = OrDefault(.Item("Titulo"), "")
                        txtReferencia.Text = OrDefault(.Item("Referencia"), "")
                        txtMotivos.Text = OrDefault(.Item("Motivos"), "")
                        txtAcciones.Text = OrDefault(.Item("Acciones"), "")
                        txtOrden.Text = OrDefault(.Item("Orden"), "")

                        Dim Auxi = OrDefault(.Item("Empresa"), 1)

                        For Each rowView As DataRowView In cmbEmpresa.Items
                            With rowView
                                If rowView.Item("Id") = Auxi Then
                                    cmbEmpresa.SelectedItem = rowView
                                    Exit For
                                End If
                            End With
                        Next

                        If Val(txtOrden.Text) <> 0 Then
                            txtOrden_KeyDown(Nothing, New KeyEventArgs(Keys.Enter))
                        End If

                        btnSac.Text = IIf(Val(OrDefault(.Item("ClaveSac"), "")) = 0, TextoBtnGenerarSac, TextoBtnVerSac)
                        btnDesvincularSAC.Enabled = Val(OrDefault(.Item("EsSacAsociada"), "0")) = 1
                    End With

                    _CargarArchivosRelacionados()

                End If

            End If

            For Each c As Control In {txtDescProducto, txtFecha, txtIncidencia, txtLotePartida, txtMotivos, txtProducto, txtReferencia, txtTitulo}
                c.Text = c.Text.Trim
            Next

            WPrimeraCarga = False

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

            With cmbEmpresa
                .Focus()
                .DroppedDown = True
            End With

        ElseIf e.KeyData = Keys.Escape Then
            cmbEstado.SelectedIndex = 0
        End If

    End Sub

    Private Sub cmbempresa_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles cmbEmpresa.KeyDown

        If e.KeyData = Keys.Enter Then
            If Trim(cmbEmpresa.Text) = "" Then : Exit Sub : End If

            txtOrden.Focus()

            If cmbEmpresaIncidencia.SelectedIndex < 1 Then cmbEmpresaIncidencia.SelectedIndex = cmbEmpresa.SelectedIndex

        ElseIf e.KeyData = Keys.Escape Then
            cmbEmpresa.SelectedIndex = 0
        End If

    End Sub

    Private Sub cmbEstado_DropDownClosed(ByVal sender As Object, ByVal e As EventArgs) Handles cmbEstado.DropDownClosed
        cmbEstado_KeyDown(Nothing, New KeyEventArgs(Keys.Enter))
    End Sub

    Private Sub cmbEmpresa_DropDownClosed(ByVal sender As Object, ByVal e As EventArgs) Handles cmbEmpresa.DropDownClosed
        cmbempresa_KeyDown(Nothing, New KeyEventArgs(Keys.Enter))
    End Sub

    Private Sub txtProducto_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtProducto.KeyDown

        If e.KeyData = Keys.Enter Then
            If rbMatPrima.Checked And txtProducto.Text.Replace(" ", "").Length < 10 Then : Exit Sub : End If
            If rbProdTerminado.Checked And txtProducto.Text.Replace(" ", "").Length < 12 Then : Exit Sub : End If

            txtDescProducto.Text = ""

            Dim WProd As DataRow = Nothing

            WProd = GetSingle("SELECT TOP 1 o.Cantidad, LTRIM(a.Descripcion) Descripcion FROM Orden o LEFT OUTER JOIN Articulo a ON a.Codigo = o.Articulo WHERE o.Orden = '" & txtOrden.Text & "' And o.Articulo = '" & txtProducto.Text & "'", CType(cmbEmpresa.SelectedItem, DataRowView).Item("Base"))

            If IsNothing(WProd) Then Exit Sub

            txtDescMP.Text = OrDefault(WProd.Item("Descripcion"), "").ToString.Trim
            If Val(txtCantidadMP.Text) = 0 Then txtCantidadMP.Text = formatonumerico(OrDefault(WProd.Item("Cantidad"), 0))

            btnControles.Enabled = Val(txtLotePartida.Text) <> 0
            btnHojaProduccion.Enabled = rbProdTerminado.Checked

            txtTitulo.Focus()

        ElseIf e.KeyData = Keys.Escape Then
            txtProducto.Text = ""
        End If

    End Sub

    Private Sub txtLotePartida_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtLotePartida.KeyDown

        If e.KeyData = Keys.Enter Then
            If Val(txtLotePartida.Text) = 0 Then : Exit Sub : End If

            Try
                _CorroborarCorrespondenciaProductoLotePartida()

            Catch ex As System.Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation)
                txtLotePartida.Focus()
            End Try

        ElseIf e.KeyData = Keys.Escape Then
            txtLotePartida.Text = ""
        End If

    End Sub

    Private Sub _CorroborarCorrespondenciaProductoLotePartida()

        Dim WProd As DataRow = _ObtenerProductoAsociado()

        If IsNothing(WProd) Then Throw New System.Exception("No se encuentra Laudo/Hoja indicada.")

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

            If WProd IsNot Nothing Then Exit For

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

            With txtMotivos
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
            _ValidarDatosIngresados()

            Dim WEnviarMail = False

            Dim WSqls As New List(Of String)

            Dim WIncid As DataRow = GetSingle("SELECT Incidencia, Tipo, ClaveSac FROM CargaIncidencias WHERE Incidencia = '" & txtIncidencia.Text & "'")

            Dim WTipo = 2
            Dim WClaveSAC = ""

            If IsNothing(WIncid) Then
                txtIncidencia.Text = ""
                WEnviarMail = True
            Else
                WTipo = OrDefault(WIncid.Item("Tipo"), 2)
                WClaveSAC = OrDefault(WIncid.Item("ClaveSac"), "")
            End If

            If Val(txtIncidencia.Text) = 0 Then
                WIncid = GetSingle("SELECT Max(Incidencia) As Ultimo FROM CargaIncidencias")
                If Not IsNothing(WIncid) Then
                    txtIncidencia.Text = OrDefault(WIncid.Item("Ultimo"), 0)
                End If
                txtIncidencia.Text = Val(txtIncidencia.Text) + 1

                Dim WProxNumero As DataRow = GetSingle("SELECT Max(Numero) As Ultimo FROM CargaIncidencias WHERE Tipo ='" & txtTipo.Text & "' And Ano = '" & txtAnio.Text & "'")

                txtNumero.Text = Val(OrDefault(WProxNumero.Item("Ultimo"), 0))

                txtNumero.Text = Val(txtNumero.Text) + 1

            End If

            Dim WEstado As Integer = CType(cmbEstado.SelectedItem, DataRowView).Item("Estado")

            Dim WEmpresa = CType(cmbEmpresa.SelectedItem, DataRowView).Item("Id")
            Dim WOrden = txtOrden.Text

            WSqls.Add("DELETE CargaIncidencias WHERE Incidencia = '" & txtIncidencia.Text & "'")

            Dim ZSql = String.Format("INSERT INTO CargaIncidencias " _
                       & "(Incidencia, Renglon, Tipo, Fecha, FechaOrd, Estado, Titulo, Referencia, Producto, Lote, ClaveSac, TipoProd, Motivos, Empresa, Orden, Acciones, Proveedor, DescProveedor, ImpreDescProd, CantOrden, EmpresaIncidencia, TipoInc, Ano, Numero) " _
                       & "VALUES " _
                       & " ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}', '{19}', '{20}', '{21}', '{22}', '{23}') ", _
                       txtIncidencia.Text, 1, WTipo, txtFecha.Text, ordenaFecha(txtFecha.Text), WEstado, txtTitulo.Text, txtReferencia.Text, txtProducto.Text, txtLotePartida.Text, WClaveSAC, IIf(rbMatPrima.Checked, "M", "T"), txtMotivos.Text, WEmpresa, WOrden, txtAcciones.Text, txtProveedor.Text, txtDescProv.Text.SliceLeft(100), txtDescProducto.Text.SliceLeft(100), formatonumerico(txtCantidadMP.Text), cmbEmpresaIncidencia.SelectedIndex, txtTipo.Text, txtAnio.Text, txtNumero.Text)
            WSqls.Add(ZSql)

            ExecuteNonQueries(WSqls.ToArray)

            If WEnviarMail Then

                If MsgBox("¿Desea enviar el aviso al Responsable de Calidad?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes _
                    Then

                    _EnviarEmail("ebiglieri@surfactan.com.ar; calidad@surfactan.com.ar; wbarosio@surfactan.com.ar; calidad2@surfactan.com.ar; isocalidad@surfactan.com.ar;juanfs@surfactan.com.ar; lsantos@surfactan.com.ar; drodriguez@surfactan.com.ar",
                                 "Carga de Informe de No Conformidad - Nro.:" + txtIncidencia.Text.PadLeft(4, "0") +
                                 " - " + Microsoft.VisualBasic.Left(txtReferencia.Text, 50),
                                 "Se inició un Informe de No Conformidad : " & txtIncidencia.Text.PadLeft(4, "0") &
                                 ". Referencia : " & txtReferencia.Text.Trim & " Título : " & txtTitulo.Text.Trim)

                End If
            End If

            If ContinuarSalirMsgBox.Show("Actualización se ha realizado con Éxito" & vbCrLf _
                                         & "Indique como quiere proseguir.", "Continuar editando Informe de NC", "Volver a Listado") = DialogResult.OK Then
                txtTitulo.Focus()
                Exit Sub
            End If

            btnCerrar.PerformClick()

        Catch ex As System.Exception
            If Trim(ex.Message) <> "" Then MsgBox(ex.Message, MsgBoxStyle.Exclamation)
            If WControlRetornoError IsNot Nothing Then WControlRetornoError.Focus()
        End Try
    End Sub

    Private Sub _EnviarEmail(ByVal Direccion As String, ByVal Subject As String, ByVal Body As String, Optional ByVal EnvioAutomatico As Boolean = False)
        Dim oApp As _Application
        Dim oMsg As _MailItem

        Try
            oApp = New Application()

            oMsg = oApp.CreateItem(OlItemType.olMailItem)
            oMsg.Subject = Subject
            oMsg.Body = Body

            ' Modificar por los E-Mails que correspondan.
            oMsg.To = Direccion

            '
            ' Generamos el PDf para poder adjuntarlo.
            '
            Dim frm As New ConsultasVarias.VistaPrevia

            With frm

                .Reporte = New ReporteINCIndividual
                .Formula = "{CargaIncidencias.Incidencia} = " & txtIncidencia.Text
                .Reporte.SetParameterValue("MostrarPosiblesUsos", 1)
                .Reporte.SetParameterValue("MostrarAcciones", 0)

                Dim WNombreArchivo = String.Format("INC {0} - {1}", txtIncidencia.Text.PadLeft(4, "0"), Date.Now.ToString("dd-MM-yyyy"))

                Dim WRuta = "C:/tempIndice/"

                WNombreArchivo &= ".pdf"

                If Directory.Exists(WRuta) Then Directory.Delete(WRuta, True)

                Directory.CreateDirectory(WRuta)

                ConsultasVarias.Clases.Conexion.EmpresaDeTrabajo = "SurfactanSa"
                ConsultasVarias.Clases.Helper._ExportarReporte(frm, ConsultasVarias.Clases.Enumeraciones.FormatoExportacion.PDF, WNombreArchivo, WRuta)

                If File.Exists(WRuta & WNombreArchivo) Then oMsg.Attachments.Add(WRuta & WNombreArchivo)

            End With

            If EnvioAutomatico Then
                oMsg.Send()
            Else
                oMsg.Display()
            End If

        Catch ex As System.Exception
            Throw New System.Exception("No se pudo crear el E-Mail solicitado." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)
        End Try
    End Sub

    Private Sub _ValidarDatosIngresados()

        WControlRetornoError = Nothing

        If txtFecha.Text.Replace(" ", "").Length < 10 OrElse Not _ValidarFecha(txtFecha.Text) Then
            WControlRetornoError = txtFecha
            Throw New System.Exception("Debe indicarse una fecha válida.")
        End If

        If cmbEstado.SelectedIndex < 0 Then
            WControlRetornoError = cmbEstado
            Throw New System.Exception("Debe indicarse un estado válido para este Informe de No Conformidad.")
        End If

        If Val(txtOrden.Text) <> 0 Then
            Dim WProd As DataRow = GetSingle("SELECT Orden FROM Orden WHERE Orden = '" & txtOrden.Text & "' And Articulo = '" & txtProducto.Text & "'", CType(cmbEmpresa.SelectedItem, DataRowView).Item("Base"))

            If WProd Is Nothing Then
                WControlRetornoError = txtProducto
                Throw New System.Exception("La Materia Prima indicada no corresponde a la Orden de Compra informada.")
            End If
        End If

        If txtTitulo.Text.Trim = "" Then
            WControlRetornoError = txtTitulo
            If MsgBox("No ha agregado nigún Título para este Informe de No Conformidad ¿Quiere proseguir con la grabación del mismo?", MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes Then
                Throw New System.Exception("")
            End If

        End If

        If Val(txtAnio.Text) = 0 Then

            WControlRetornoError = txtAnio

            MsgBox("No ha indicado el año al que pertenece en este Informe de No Conformidad.", MsgBoxStyle.Exclamation)

            Throw New System.Exception("")

        End If

        If txtReferencia.Text.Trim = "" Then
            WControlRetornoError = txtReferencia
            If MsgBox("No ha agregado niguna Referencia para este Informe de No Conformidad ¿Quiere proseguir con la grabación del mismo?", MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes Then
                Throw New System.Exception("")
            End If
        End If

    End Sub

    Private Sub SoloNumero(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles txtIncidencia.KeyPress, txtNumero.KeyPress, txtTipo.KeyPress, txtAnio.KeyPress
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

    Private Sub txtOrden_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtOrden.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtOrden.Text) = "" Then : Exit Sub : End If

            Dim WEmpresa As String = CType(cmbEmpresa.SelectedItem, DataRowView).Item("Base")

            If WEmpresa <> "" And Val(txtOrden.Text) <> 0 Then
                Dim WOrden As DataRow = GetSingle("SELECT o.Proveedor, DescProv = p.Nombre FROM Orden o LEFT OUTER JOIN Proveedor p ON p.Proveedor = o.Proveedor WHERE Orden = '" & txtOrden.Text & "' Order by o.Renglon", WEmpresa)

                If WOrden IsNot Nothing Then

                    With WOrden
                        txtProveedor.Text = OrDefault(.Item("Proveedor"), "")
                        txtDescProv.Text = Trim(OrDefault(.Item("DescProv"), ""))
                    End With

                    If Not WPrimeraCarga Then
                        btnMPAsociadasOC_Click(Nothing, Nothing)
                    ElseIf txtProducto.Text.Replace(" ", "").Length = 10 Then
                        txtProducto_KeyDown(Nothing, New KeyEventArgs(Keys.Enter))
                    End If

                    txtTitulo.Focus()
                Else
                    MsgBox("No se encontró la Orden de Compra en la Empresa indicada.", MsgBoxStyle.Exclamation)
                End If

            End If

        ElseIf e.KeyData = Keys.Escape Then
            txtOrden.Text = ""
        End If

    End Sub

    Private Sub btnSac_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSac.Click

        Try
            Dim WNC As DataRow = GetSingle("SELECT ClaveSac FROM CargaIncidencias WHERE Incidencia = '" & txtIncidencia.Text & "' And Renglon = 1")

            If WNC Is Nothing Then Exit Sub

            Dim WClaveSAC As String = Trim(OrDefault(WNC.Item("ClaveSac"), ""))

            If WClaveSAC = "" Then
                btnDesvincularSAC.Enabled = False
                '
                ' Consultamos si quieren abrir nueva Sac o asignar una ya abierta.
                '
                Dim WPregunta = New ContinuarSalirMsgBox("¿Como desea asignarle una SAC a este Informe de No Conformidad?",
                                              "Abrir Nueva SAC", "Asignar SAC ya Abierta")
                WPregunta._DrBtn1 = DialogResult.Yes
                WPregunta._DrBtn2 = DialogResult.No

                Dim Resp As DialogResult = WPregunta.ShowDialog(Me)


                Select Case Resp
                    Case DialogResult.Yes

                        '
                        ' Creamos nueva SAC, colocando como datos por Default los provistos por el actual
                        ' Informe de No Conformidad.
                        '

                        Dim WMotivos As String = ""

                        If Val(txtOrden.Text) <> 0 Then
                            WMotivos &= vbCrLf & "======================================================================="

                            WMotivos &= vbCrLf & vbCrLf & String.Format("Orden de Compra: {0} {4}Articulo: ({1}) {2} {4}Cantidad: {3}", txtOrden.Text, txtProducto.Text, txtDescMP.Text.Trim, txtCantidadMP.Text, vbTab)

                            WMotivos &= vbCrLf & vbCrLf & "=======================================================================" & vbCrLf & vbCrLf
                        End If

                        WMotivos &= txtMotivos.Text.Trim


                        Dim frm As New AuxiNuevaSACDesdeINC(txtIncidencia.Text, txtFecha.Text, txtTitulo.Text, txtReferencia.Text, WMotivos)

                        frm.Show(Me)

                    Case DialogResult.No
                        '
                        ' Desplegamos el listado con las SAC's disponibles.
                        '
                        Dim frm As New AyudaListadoSACs
                        frm.Show(Me)

                End Select

            Else
                _AbrirSACPorClave(WClaveSAC)
            End If

        Catch ex As System.Exception
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

    Public Sub _ProcesarAyudaListadoSACs(ByVal WClaveSac As String) Implements IAyudaListadoSACs._ProcesarAyudaListadoSACs
        If WClaveSac.Trim = "" Then Exit Sub

        btnSac.Text = TextoBtnVerSac

        WClaveSac = WClaveSac.PadLeft(14, "0")

        Dim WAnio = WClaveSac.Substring(4, 4)
        Dim WNum = WClaveSac.Substring(8, 6)

        ExecuteNonQueries("UPDATE CargaIncidencias SET ClaveSac = '" & WClaveSac & "' WHERE Incidencia = '" & txtIncidencia.Text & "'")

        _CopiarArchivosINCASAC(WClaveSac)

        btnDesvincularSAC.Enabled = True

        MsgBox(String.Format("Se ha Asociado la siguiente SAC {0} {2} / {1} al Informe de No Conformidad Actual.", "", WAnio, WNum), MsgBoxStyle.Information)

        If MsgBox("¿Desea Abrir la SAC asociada?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            _AbrirSACPorClave(WClaveSac)
            Exit Sub
        End If
    End Sub

    Private Sub txtOrden_KeyUp(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtOrden.KeyUp
        btnMPAsociadasOC.Enabled = Trim(txtOrden.Text) <> ""
    End Sub

    Private Sub btnMPAsociadasOC_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnMPAsociadasOC.Click
        Dim WEmp As DataRowView = cmbEmpresa.SelectedItem

        If WEmp Is Nothing Then Exit Sub

        Dim frm As New MPAsociadasOC(txtOrden.Text, WEmp.Item("Nombre"), WEmp.Item("Base"))
        frm.Show(Me)
    End Sub

    Public Sub _ProcesarMPAsociadasOC(ByVal WCodigo As Object, ByVal WDescripcion As Object, ByVal WCantidad As Object) Implements IAyudaMPAsociadasOC._ProcesarMPAsociadasOC
        txtProducto.Text = WCodigo
        txtDescMP.Text = WDescripcion
        txtCantidadMP.Text = formatonumerico(WCantidad)
        txtTitulo.Focus()
    End Sub

    Private Sub txtProducto_DoubleClick(ByVal sender As Object, ByVal e As EventArgs) Handles txtProducto.DoubleClick
        btnMPAsociadasOC_Click(Nothing, Nothing)
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
                        Catch ex As System.Exception
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
            Catch ex As System.Exception
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

                    Catch ex As System.Exception
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
            Throw New System.Exception("No se ha logrado tener acceso a la Carpeta Compartida de Archivos Relacionados.")
        End If

        WRutaArchivosRelacionados = _RutaCarpetaArchivos() & "\" & WNombreCarpetaArchivos

        ' Creamos la Carpeta en caso de que no exista aún.
        If Not Directory.Exists(WRutaArchivosRelacionados) Then
            Try
                Directory.CreateDirectory(WRutaArchivosRelacionados)
            Catch ex As System.Exception
                Throw New System.Exception(ex.Message)
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
        Catch ex As System.Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button2.Click
        Try
            If dgvArchivos.SelectedRows.Count > 0 Then
                EliminarArchivoToolStripMenuItem_Click(Nothing, Nothing)
            End If
        Catch ex As System.Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

#End Region

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            ExecuteNonQueries("UPDATE CargaIncidencias SET ImpreDescProd = CASE inc.TipoProd WHEN 'T' THEN LEFT(t.Descripcion, 100) WHEN 'M' THEN LEFT(a.Descripcion, 100) ELSE '' END FROM CargaIncidencias inc LEFT OUTER JOIN Terminado t ON t.Codigo = inc.Producto LEFT OUTER JOIN Articulo a ON a.Codigo = inc.Producto")

            Dim frm As New ExportarINC
            frm.Show(Me)

        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub _ProcesarExportarINC(ByVal TipoFormato As Integer) Implements IExportarINC._ProcesarExportarINC
        With New VistaPrevia
            .Reporte = New ReporteINCIndividual
            .Formula = "{CargaIncidencias.Incidencia} = " & txtIncidencia.Text
            .Reporte.SetParameterValue("MostrarPosiblesUsos", 0)
            .Reporte.SetParameterValue("MostrarAcciones", 1)

            Dim WNombreArchivo = String.Format("INC {0} - {1}", txtIncidencia.Text.PadLeft(4, "0"), Date.Now.ToString("dd-MM-yyyy"))

            Select Case TipoFormato

                Case 0 ' PDF
                    WNombreArchivo &= ".pdf"
                    .Exportar(WNombreArchivo, ExportFormatType.PortableDocFormat)
                Case 1 ' Excel
                    .Exportar(WNombreArchivo, ExportFormatType.Excel)
                Case 2 ' Word
                    .Exportar(WNombreArchivo, ExportFormatType.WordForWindows)
                Case 3
                    Dim WRuta = "C:/tempIndice/"
                    WNombreArchivo &= ".pdf"

                    If Directory.Exists(WRuta) Then Directory.Delete(WRuta, True)

                    Directory.CreateDirectory(WRuta)

                    .Exportar(WNombreArchivo, ExportFormatType.PortableDocFormat, WRuta)

                    .EnviarPorEmail(WRuta & WNombreArchivo)

            End Select

        End With
    End Sub

    Private Sub btnDesvincularSAC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDesvincularSAC.Click

        Try

            If MsgBox("¿Está seguro de querer desvincular la SAC asociada a este Informe de No Conformidad?", vbYesNo) <> MsgBoxResult.Yes Then Exit Sub

            ExecuteNonQueries("UPDATE CargaIncidencias SET ClaveSAC = '', EsSacAsociada = '0' WHERE Incidencia = '" & txtIncidencia.Text & "'")

            MsgBox("SAC desvinculada de manera exitosa", MsgBoxStyle.Information)

            btnSac.Text = TextoBtnGenerarSac
            btnDesvincularSAC.Enabled = False

        Catch ex As System.Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub btnEliminar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEliminar.Click

        If Not WAutorizadoEliminar Then

            If MsgBox("¿Está seguro de querer eliminar este Informe de no Conformidad?", MsgBoxStyle.YesNo) <> MsgBoxResult.Yes Then Exit Sub

            With New IngresoClaveSeguridad
                .Show(Me)
            End With

            Exit Sub
        End If

        ExecuteNonQueries("DELETE FROM CargaIncidencias WHERE Incidencia = '" & txtIncidencia.Text & "'")

        btnCerrar.PerformClick()

    End Sub

    Private Sub btnModifNumINC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModifNumINC.Click

        Dim WActual As DataRow = GetSingle("SELECT Incidencia FROM CargaIncidencias WHERE Incidencia = '" & txtIncidencia.Text & "'")

        If WActual Is Nothing Then Exit Sub

        With New ModifNumeracionINC(txtTipo.Text, txtAnio.Text, txtNumero.Text)
            .Show(Me)
        End With

    End Sub

    Public Sub _ProcesarModifNumeracionINC(ByVal NuevaNumeracion As String) Implements IModifNumeracionINC._ProcesarModifNumeracionINC
        txtIncidencia.Text = NuevaNumeracion
        txtIncidencia_KeyDown(Nothing, New KeyEventArgs(Keys.Enter))
    End Sub

    Public Sub _ProcesarIngresoClaveSeguridad(ByVal WClave As Object) Implements IIngresoClaveSeguridad._ProcesarIngresoClaveSeguridad
        WAutorizadoEliminar = UCase(Trim(WClave)) = UCase(Trim(Operador.Clave))

        If Not WAutorizadoEliminar Then
            MsgBox("Clave erronea", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        btnEliminar.PerformClick()
    End Sub

    Private Sub txtCantidadMP_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCantidadMP.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtCantidadMP.Text) = "" Then : Exit Sub : End If

            txtCantidadMP.Text = formatonumerico(txtCantidadMP.Text)

            txtTitulo.Focus()

        ElseIf e.KeyData = Keys.Escape Then
            txtCantidadMP.Text = ""
        End If

    End Sub

    Private Sub txtAnio_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAnio.KeyDown

        If e.KeyData = Keys.Enter Then
            If Trim(txtAnio.Text) = "" Then txtAnio.Text = Date.Now.ToString("yyyy")

            txtNumero.Focus()

        ElseIf e.KeyData = Keys.Escape Then
        txtAnio.Text = ""
        End If

    End Sub

    Private Sub txtNumero_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNumero.KeyDown

        If e.KeyData = Keys.Enter Then

            If Val(txtNumero.Text) <> 0 Then

                Dim WNumero = txtNumero.Text
                Dim WAnio = txtAnio.Text
                Dim WTipo = txtTipo.Text

                _Limpiar()

                txtAnio.Text = WAnio
                txtTipo.Text = WTipo
                txtNumero.Text = WNumero

                Dim WInc As DataRow = GetSingle("SELECT Incidencia FROM CargaIncidencias WHERE Ano = '" & txtAnio.Text & "' And Tipo = '" & txtTipo.Text & "' And Numero = '" & txtNumero.Text & "'")

                If WInc IsNot Nothing Then
                    txtIncidencia.Text = OrDefault(WInc.Item("Incidencia"), "")

                    txtIncidencia_KeyDown(Nothing, New KeyEventArgs(Keys.Enter))

                    txtFecha.Focus()

                End If

            End If

        ElseIf e.KeyData = Keys.Escape Then
            txtNumero.Text = ""
        End If

    End Sub

    Private Sub txtTipo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTipo.KeyDown

        If e.KeyData = Keys.Enter Then
            txtDescTipo.Text = ""

            If Trim(txtTipo.Text) = "" Then : Exit Sub : End If


            Dim WTipo As DataRow = GetSingle("SELECT Descripcion FROM TiposINC WHERE Tipo = '" & txtTipo.Text & "'")

            If WTipo IsNot Nothing Then
                txtDescTipo.Text = Trim(OrDefault(WTipo.Item("Descripcion"), ""))
            End If

            txtNumero.Focus()
            
        ElseIf e.KeyData = Keys.Escape Then
            txtTipo.Text = ""
        End If

    End Sub
End Class