﻿Imports System.Data.SqlClient

Public Class ConsultaInformacionPersonal
    Private Sub IngresoOrdenTrabajo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        btnLimpiar.PerformClick()
    End Sub

    Private Sub btnCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Close()
    End Sub

    Private Sub btnLimpiar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpiar.Click
        ' Limpiamos los campos.
        _LimpiarCampos()

        TabControl1.SelectedIndex = 0
        txtDni.Focus()
    End Sub

    Private Sub _LimpiarCampos()

        For Each _txt As TextBox In _CamposDeTexto()
            _txt.Text = ""
        Next

        For Each _m As MaskedTextBox In _
            {txtFechaCasamiento, txtFechaEgreso, txtFechaIngreso, txtFechaNacimiento, txtFechaNacimientoConyugue}
            _m.Clear()
        Next

        For Each _c As ComboBox In {cmbCategoria, cmbEstado, cmbUbicacion}
            _c.SelectedIndex = 0
        Next

        With dgvHijos.Rows
            .Clear
            .Add
        End With

        With dgvEducacion.Rows
            .Clear
            '.Add
        End With

        With dgvIndumentaria.Rows
            .Clear
            '.Add
        End With

        pnlConsulta.Visible = False
    End Sub

    Private Function _CamposDeTexto() As TextBox()
        Return _
            {txtDni, txtCalle, txtNumero, txtDpto, txtCodPostal, txtLocalidad, txtAclaracion, txtNombreCompletoConyugue,
             txtEdadConyugue, txtDniConyugue, txtSueldoBruto, txtAyuda, txtNombreCompleto, txtLegajo}
    End Function

    Private Sub IngresoOrdenTrabajo_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles MyBase.Shown
        txtDni.Focus
    End Sub

    Private Sub txtDni_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtDni.KeyDown

        If e.KeyData = Keys.Enter Then
            If Trim(txtDni.Text) = "" Then : Exit Sub :
            End If

            ' Rutina para cargar los datos.
            Try

                _CargarDatos()

                txtFechaNacimiento.Focus

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation)
            End Try
        ElseIf e.KeyData = Keys.Escape Then
            txtDni.Text = ""
        End If
    End Sub

    Private Sub _CargarDatos()

        Try
            ' Busco los datos provenientes del Legajo.
            _CargarDatosLegajo()

            ' Busco los Datos Personales.
            _CargarDatosPersonales()

            ' Cargo los datos de los Hijos
            _CargarDatosHijos()


        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub _CargarDatosHijos()

        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("SELECT * FROM PersonalHijos WHERE Dni = '" & txtDni.Text & "'")
        Dim dr As SqlDataReader
        Dim WNombre, WApellido, WDni, WEdad, WFechaNac As String
        Dim WFila As Short

        Try

            cn.ConnectionString = Helper._ConectarA
            cn.Open()
            cm.Connection = cn

            dr = cm.ExecuteReader()

            dgvHijos.Rows.Clear
            WFila = 0

            With dr
                If .HasRows Then

                    Do While .Read()

                        WNombre = IIf(IsDBNull(.Item("Nombre")), "", .Item("Nombre"))
                        WApellido = IIf(IsDBNull(.Item("Apellido")), "", .Item("Apellido"))
                        WDni = IIf(IsDBNull(.Item("DniHijo")), "", .Item("DniHijo"))
                        WEdad = IIf(IsDBNull(.Item("Edad")), "", .Item("Edad"))
                        WFechaNac = IIf(IsDBNull(.Item("FechaNac")), "", .Item("FechaNac"))

                        WFila = dgvHijos.Rows.Add

                        dgvHijos.Rows(WFila).Cells("NombreHijo").Value = Trim(WNombre)
                        dgvHijos.Rows(WFila).Cells("ApellidoHijo").Value = Trim(WApellido)
                        dgvHijos.Rows(WFila).Cells("DniHijo").Value = Trim(WDni)
                        dgvHijos.Rows(WFila).Cells("EdadHijo").Value = Trim(WEdad)
                        dgvHijos.Rows(WFila).Cells("FechaNacimientoHijo").Value = Trim(WFechaNac)

                    Loop

                End If
            End With

        Catch ex As Exception
            Throw _
                New Exception(
                    "Hubo un problema al querer consultar la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try
    End Sub

    Private Sub _CargarDatosPersonales()

        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("SELECT * FROM Personal WHERE Dni = '" & txtDni.Text & "'")
        Dim dr As SqlDataReader
        Dim WFila As Short
        Dim WObsPrimaria, WObsSecundaria, WObsTerciaria As String

        Try

            cn.ConnectionString = Helper._ConectarA
            cn.Open()
            cm.Connection = cn

            dr = cm.ExecuteReader()

            dgvIndumentaria.Rows.Clear

            If dr.HasRows Then
                With dr
                    .Read

                    txtFechaNacimiento.Text = IIf(IsDBNull(.Item("FechaNac")), "", .Item("FechaNac"))
                    txtCalle.Text = IIf(IsDBNull(.Item("Calle")), "", .Item("Calle"))
                    txtNumero.Text = IIf(IsDBNull(.Item("Numero")), "", .Item("Numero"))
                    txtDpto.Text = IIf(IsDBNull(.Item("Dpto")), "", .Item("Dpto"))
                    txtCodPostal.Text = IIf(IsDBNull(.Item("Postal")), "", .Item("Postal"))
                    txtLocalidad.Text = IIf(IsDBNull(.Item("Localidad")), "", .Item("Localidad"))
                    txtAclaracion.Text = IIf(IsDBNull(.Item("Aclaracion")), "", .Item("Aclaracion"))
                    txtNombreCompletoConyugue.Text = IIf(IsDBNull(.Item("ConyugeNombre")), "", .Item("ConyugeNombre"))
                    txtEdadConyugue.Text = IIf(IsDBNull(.Item("ConyugeEdad")), "", .Item("ConyugeEdad"))
                    txtDniConyugue.Text = IIf(IsDBNull(.Item("ConyugeDni")), "", .Item("ConyugeDni"))
                    txtFechaNacimientoConyugue.Text = IIf(IsDBNull(.Item("ConyugeFechaNac")), "",
                                                          .Item("ConyugeFechaNac"))
                    txtFechaCasamiento.Text = IIf(IsDBNull(.Item("FechaCasamiento")), "", .Item("FechaCasamiento"))
                    cmbEstado.SelectedIndex = IIf(IsDBNull(.Item("Estado")), 0, .Item("Estado"))
                    cmbCategoria.SelectedIndex = IIf(IsDBNull(.Item("Categoria")), 0, .Item("Categoria"))
                    cmbUbicacion.SelectedIndex = IIf(IsDBNull(.Item("Ubicacion")), 0, .Item("Ubicacion"))
                    txtSueldoBruto.Text = IIf(IsDBNull(.Item("SueldoBruto")), "0", .Item("SueldoBruto"))
                    txtSueldoBruto.Text = Helper.formatonumerico(txtSueldoBruto.Text)

                    ' Cargo Informacion de la Indumentaria.
                    WFila = 0

                    ' Buzo: 1, Camisa: 2, Campera: 3, Pantalón: 4, Remera: 5, Zapato: 6
                    Dim WIndumentaria As String() = {"", "Buzo", "Camisa", "Campera", "Pantalon", "Remera", "Zapato"}
                    Dim WItem, WTalle, WObs As String

                    For i = 1 to 6

                        WItem = WIndumentaria(i)

                        WFila = dgvIndumentaria.Rows.Add
                        dgvIndumentaria.Rows(WFila).Cells("Indumentaria").Value = WItem

                        WTalle = IIf(IsDBNull(.Item(WItem)), "", .Item(WItem))
                        dgvIndumentaria.Rows(wfila).Cells("Talle").Value = Trim(WTalle)

                        WObs = IIf(IsDBNull(.Item("Obs" & WItem)), "", .Item("Obs" & WItem))
                        dgvIndumentaria.Rows(wfila).Cells("ObservacionesIndumentaria").Value = Trim(WObs)
                        dgvIndumentaria.Rows(WFila).Cells("TipoInd").Value = Str$(i)

                    Next

                    ' Obtenemos los datos de los Estudios.
                    WObsPrimaria = IIf(IsDBNull(.Item("ObsPrimaria")), "", .Item("ObsPrimaria"))
                    WObsSecundaria = IIf(IsDBNull(.Item("ObsSecundaria")), "", .Item("ObsSecundaria"))
                    WObsTerciaria = IIf(IsDBNull(.Item("ObsTerciaria")), "", .Item("ObsTerciaria"))


                    For Each row As DataGridViewRow In dgvEducacion.Rows

                        Select Case val(row.Cells("IdFormacion").Value)
                            Case 1 ' Primaria
                                row.Cells("ObservacionesFormacion").Value = Trim(WObsPrimaria)
                            Case 2 ' Secundaria
                                row.Cells("ObservacionesFormacion").Value = Trim(WObsSecundaria)
                            Case 3 ' Terciaria
                                row.Cells("ObservacionesFormacion").Value = Trim(WObsTerciaria)
                            Case Else
                                row.Cells("ObservacionesFormacion").Value = ""
                        End Select

                    Next

                    for Each txt As TextBox In _CamposDeTexto
                        txt.Text = Trim(txt.Text)
                    Next

                End With
            End If

        Catch ex As Exception
            Throw _
                New Exception(
                    "Hubo un problema al querer consultar la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try
    End Sub

    Private Sub _CargarDatosLegajo()

        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("SELECT Codigo, FechaVersion, FIngreso, FEgreso, Descripcion, " _
                                              & " EstaI, EstaII, EstaIII, EstadoI, EstadoII, EstadoIII " _
                                              & " FROM Legajo WHERE Dni = '" & txtDni.Text & "' AND Renglon = 1")
        Dim dr As SqlDataReader
        Dim _
            WLegajoActual,
            WLegajos,
            WAux,
            WFechaVersion,
            WFechaVersionOrd,
            WNombreCompleto,
            WFechaIngreso,
            WFechaEgreso,
            WEstado As String
        Dim WFilaFormacion As Short

        Try

            cn.ConnectionString = Helper._ConectarA
            cn.Open()
            cm.Connection = cn

            dr = cm.ExecuteReader()

            If dr.HasRows Then

                WLegajos = ""

                Do While dr.Read

                    With dr

                        WFechaVersion = IIf(IsDBNull(.Item("FechaVersion")), "", .Item("FechaVersion"))
                        WAux = Helper.ordenaFecha(WFechaVersion)
                        WLegajoActual = IIf(IsDBNull(.Item("Codigo")), "", .Item("Codigo"))

                        WLegajos &= IIf(Trim(wlegajos) <> "", " - " & WLegajoActual, WLegajoActual)

                        If Val(WAux) > Val(WFechaVersionOrd) then

                            WFechaVersionOrd = WAux

                            WNombreCompleto = IIf(IsDBNull(.Item("Descripcion")), "", .Item("Descripcion"))
                            WFechaIngreso = IIf(IsDBNull(.Item("FIngreso")), "", .Item("FIngreso"))
                            WFechaEgreso = IIf(IsDBNull(.Item("FEgreso")), "", .Item("FEgreso"))

                            dgvEducacion.Rows.Clear

                            ' Grabamos Primaria.
                            WEstado = IIf(IsDBNull(.Item("EstaI")), "", .Item("EstaI"))

                            Select Case Val(WEstado)
                                Case 1, 2, 8

                                    WFilaFormacion = dgvEducacion.Rows.Add

                                    dgvEducacion.Rows(WFilaFormacion).Cells("TipoFormacion").Value = "Primaria"
                                    dgvEducacion.Rows(WFilaFormacion).Cells("TituloFormacion").Value = "" _
                                    'IIf(IsDBNull(.Item("EstadoI")), "", .Item("EstadoI"))
                                    dgvEducacion.Rows(WFilaFormacion).Cells("IdFormacion").Value = "1" _
                                    'IIf(IsDBNull(.Item("EstadoI")), "", .Item("EstadoI"))

                                Case Else
                                    WEstado = ""
                            End Select

                            ' Grabamos Secundaria
                            WEstado = IIf(IsDBNull(.Item("EstaII")), "", .Item("EstaII"))

                            Select Case Val(WEstado)
                                Case 1, 2, 8

                                    WFilaFormacion = dgvEducacion.Rows.Add

                                    dgvEducacion.Rows(WFilaFormacion).Cells("TipoFormacion").Value = "Secundaria"
                                    dgvEducacion.Rows(WFilaFormacion).Cells("TituloFormacion").Value =
                                        IIf(IsDBNull(.Item("EstadoII")), "", .Item("EstadoII"))
                                    dgvEducacion.Rows(WFilaFormacion).Cells("IdFormacion").Value = "2"

                                Case Else
                                    WEstado = ""
                            End Select

                            ' Grabamos Terciaria/Universitaria
                            WEstado = IIf(IsDBNull(.Item("EstaIII")), "", .Item("EstaIII"))

                            Select Case Val(WEstado)
                                Case 1, 2, 8

                                    WFilaFormacion = dgvEducacion.Rows.Add

                                    dgvEducacion.Rows(WFilaFormacion).Cells("TipoFormacion").Value =
                                        "Terciaria/Universitaria"
                                    dgvEducacion.Rows(WFilaFormacion).Cells("TituloFormacion").Value =
                                        IIf(IsDBNull(.Item("EstadoIII")), "", .Item("EstadoIII"))
                                    dgvEducacion.Rows(WFilaFormacion).Cells("IdFormacion").Value = "3"

                                Case Else
                                    WEstado = ""
                            End Select

                        End If

                    End With

                Loop

                txtLegajo.Text = WLegajos
                txtFechaIngreso.Text = WFechaIngreso
                txtFechaEgreso.Text = IIf(WFechaEgreso = "00/00/0000", "  /  /    ", WFechaEgreso)
                txtNombreCompleto.Text = trim(WNombreCompleto)

            End If

        Catch ex As Exception
            Throw _
                New Exception(
                    "Hubo un problema al querer consultar la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try
    End Sub

    Private Sub txtNombreCompleto_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtNombreCompleto.KeyDown


        If e.KeyData = Keys.Enter Then
            If Trim(txtNombreCompleto.Text) = "" Then : Exit Sub :
            End If

            txtFechaNacimiento.Focus

        ElseIf e.KeyData = Keys.Escape Then
            txtNombreCompleto.Text = ""
        End If
    End Sub

    Private Sub txtFechaNacimiento_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtFechaNacimiento.KeyDown

        If e.KeyData = Keys.Enter Then
            'If txtFechaNacimiento.Text.estaVacia Then : Exit Sub : End If

            If Helper._ValidarFecha(txtFechaNacimiento.Text) then
                txtCalle.Focus
            End If

        ElseIf e.KeyData = Keys.Escape Then
            txtFechaNacimiento.Text = ""
        End If
    End Sub

    Private Sub txtCalle_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtCalle.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtCalle.Text) = "" Then : Exit Sub : End If

            txtNumero.Focus

        ElseIf e.KeyData = Keys.Escape Then
            txtCalle.Text = ""
        End If
    End Sub

    Private Sub txtNumero_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtNumero.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtNumero.Text) = "" Then : Exit Sub : End If

            txtDpto.Focus

        ElseIf e.KeyData = Keys.Escape Then
            txtNumero.Text = ""
        End If
    End Sub

    Private Sub txtDpto_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtDpto.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtDpto.Text) = "" Then : Exit Sub : End If

            txtCodPostal.Focus

        ElseIf e.KeyData = Keys.Escape Then
            txtDpto.Text = ""
        End If
    End Sub

    Private Sub txtCodPostal_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtCodPostal.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtCodPostal.Text) = "" Then : Exit Sub : End If

            txtLocalidad.Focus

        ElseIf e.KeyData = Keys.Escape Then
            txtCodPostal.Text = ""
        End If
    End Sub

    Private Sub txtLocalidad_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtLocalidad.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtLocalidad.Text) = "" Then : Exit Sub : End If

            txtAclaracion.Focus

        ElseIf e.KeyData = Keys.Escape Then
            txtLocalidad.Text = ""
        End If
    End Sub

    Private Sub txtAclaracion_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtAclaracion.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtAclaracion.Text) = "" Then : Exit Sub : End If

            txtFechaIngreso.Focus

        ElseIf e.KeyData = Keys.Escape Then
            txtAclaracion.Text = ""
        End If
    End Sub

    Private Sub txtFechaIngreso_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtFechaIngreso.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtFechaEgreso.Text) = "" Then : Exit Sub : End If

            txtFechaEgreso.Focus

        ElseIf e.KeyData = Keys.Escape Then
            txtFechaEgreso.Clear
        End If
    End Sub

    Private Sub txtFechaEgreso_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtFechaEgreso.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtFechaEgreso.Text) = "" Then : Exit Sub : End If

            TabControl1.SelectTab(1)

            txtNombreCompletoConyugue.Focus

        ElseIf e.KeyData = Keys.Escape Then
            txtFechaEgreso.Text = ""
        End If
    End Sub

    Private Sub txtNombreCompletoConyugue_KeyDown(ByVal sender As System.Object,
                                                  ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtNombreCompletoConyugue.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtNombreCompletoConyugue.Text) = "" Then : Exit Sub : End If

            txtEdadConyugue.Focus

        ElseIf e.KeyData = Keys.Escape Then
            txtNombreCompletoConyugue.Text = ""
        End If
    End Sub

    Private Sub txtEdadConyugue_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtEdadConyugue.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtEdadConyugue.Text) = "" Then : Exit Sub : End If

            txtDniConyugue.Focus

        ElseIf e.KeyData = Keys.Escape Then
            txtEdadConyugue.Text = ""
        End If
    End Sub

    Private Sub txtDniConyugue_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtDniConyugue.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtDniConyugue.Text) = "" Then : Exit Sub : End If

            txtFechaNacimientoConyugue.Focus

        ElseIf e.KeyData = Keys.Escape Then
            txtDniConyugue.Text = ""
        End If
    End Sub

    Private Sub txtFechaNacimientoConyugue_KeyDown(ByVal sender As System.Object,
                                                   ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtFechaNacimientoConyugue.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtFechaNacimientoConyugue.Text) = "" Then : Exit Sub : End If

            ' Validamos la fecha introducida sólo en los casos en que haya colocado alguna.
            If Not txtFechaNacimientoConyugue.Text.estaVacia then
                If Helper._ValidarFecha(txtFechaNacimientoConyugue.Text) then
                    txtFechaCasamiento.Focus
                Else
                    Exit Sub
                End If
            Else
                txtFechaCasamiento.Focus
            End If

        ElseIf e.KeyData = Keys.Escape Then
            txtFechaNacimientoConyugue.Clear
        End If
    End Sub

    Private Sub txtFechaCasamiento_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtFechaCasamiento.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtFechaCasamiento.Text) = "" Then : Exit Sub : End If

            If Not txtFechaCasamiento.Text.estaVacia then

                If Helper._ValidarFecha(txtFechaCasamiento.Text) then

                    TabControl1.SelectTab(2)

                    With cmbEstado
                        .DroppedDown = True
                        .Focus
                    End With

                Else
                    Exit Sub
                End If
            Else

                TabControl1.SelectTab(2)

                With cmbEstado
                    .DroppedDown = True
                    .Focus
                End With

            End If

        ElseIf e.KeyData = Keys.Escape Then
            txtFechaCasamiento.Clear
        End If
    End Sub

    'Private Sub txtLegajo_KeyDown( ByVal sender As System.Object,  ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLegajo.KeyDown

    '    If e.KeyData = Keys.Enter Then
    '     If Trim(txtLegajo.Text) = "" Then : Exit Sub : End If


    '    ElseIf e.KeyData = Keys.Escape Then
    '        txtLegajo.Text = ""
    '    End If

    'End Sub

    Private Sub txtSueldoBruto_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles txtSueldoBruto.KeyDown

        If e.KeyData = Keys.Enter Then
            'If Trim(txtSueldoBruto.Text) = "" Then : Exit Sub : End If

            txtSueldoBruto.Text = Helper.formatonumerico(txtSueldoBruto.Text)

            With cmbUbicacion
                .DroppedDown = True
                .Focus
            End With

        ElseIf e.KeyData = Keys.Escape Then
            txtSueldoBruto.Text = ""
        End If
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles TabControl1.SelectedIndexChanged
        Select Case TabControl1.SelectedIndex
            Case 0
                txtDni.Focus
            Case 1
                txtNombreCompletoConyugue.Focus
            Case 2
                txtSueldoBruto.Focus
            Case 3
                dgvEducacion.Focus
        End Select
    End Sub

    Private Sub cmbEstado_DropDownClosed(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles cmbEstado.DropDownClosed

        With cmbCategoria
            .DroppedDown = True
            .Focus
        End With
    End Sub

    Private Sub cmbEstado_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles cmbEstado.SelectedIndexChanged

        Select Case cmbEstado.SelectedIndex
            Case 1
                cmbCategoria.DataSource =
                    Configuration.ConfigurationManager.AppSettings("CAT_EN_CCT").Split(",").ToArray()
            Case 2
                cmbCategoria.DataSource =
                    Configuration.ConfigurationManager.AppSettings("CAT_FUERA_CCT").Split(",").ToArray()
            Case Else
                cmbCategoria.DataSource = {"", "", "", ""}
        End Select
    End Sub

    Private Sub cmbCategoria_DropDownClosed(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles cmbCategoria.DropDownClosed
        txtSueldoBruto.Focus
    End Sub
End Class