﻿Imports System.Data.SqlClient

Public Class HistorialProforma

    ' Para controles de grilla.
    Private Const YMARGEN = 1.5
    Private Const XMARGEN = 4.9
    Private WRow, Wcol As Integer

    ' Constantes
    Private Const PRODUCTOS_MAX = 6

    Private _NroProforma As String

    Public Property NroProforma() As String
        Get
            Return _NroProforma
        End Get
        Set(ByVal value As String)
            _NroProforma = value
        End Set
    End Property

    Private Sub HistorialProforma_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not IsNothing(Me.NroProforma) Then
            Try
                txtNroProforma.Text = Helper.ceros(Me.NroProforma, 6)

                _MostrarHistorial()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
                Me.Close()
            End Try
        End If

        txtFechaAux.Visible = False

        txtFecha.Text = Date.Now.ToString("dd-MM-yyyy")

        WRow = -1
        Wcol = -1

    End Sub

    Private Sub HistorialProforma_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        txtNroProforma.Focus()
    End Sub

    Private Function _CS()
        Return Helper._ConectarA()
    End Function

    Private Sub _MostrarHistorial()
        'If IsNothing(Me.Proforma) Then : Exit Sub : End If

        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand()
        Dim dr As SqlDataReader
        Dim WClave, WFecha, WFechaOrd, WNroObservacion, WNroObsAnt, WObservacion, WUsuario, XRenglon, ZSql, WRefPrimeraFilaObs

        ZSql = ""
        ZSql &= "SELECT DISTINCT h.Clave, h.NroObservacion, h.Renglon, h.Fecha, h.FechaOrd, h.Usuario, h.Observaciones, p.Cliente, c.Razon"
        ZSql &= " FROM ProformaExportacionHistorial as h, ProformaExportacion as p, Cliente as c"
        ZSql &= " WHERE h.Proforma = '" & txtNroProforma.Text & "' AND p.Proforma = h.Proforma AND c.Cliente = p.Cliente ORDER BY h.FechaOrd, h.Renglon"

        Try
            cn.ConnectionString = _CS()
            cn.Open()
            cm.CommandText = ZSql
            cm.Connection = cn

            dr = cm.ExecuteReader()

            If dr.HasRows Then

                btnLimpiar.PerformClick()

                dgvProductos.Rows.Clear()

                WRefPrimeraFilaObs = 0
                WNroObsAnt = -1

                Do While dr.Read()

                    WClave = ""
                    WFecha = ""
                    WFechaOrd = 0
                    WObservacion = ""
                    WUsuario = ""
                    WNroObservacion = 0

                    txtCliente.Text = IIf(IsDBNull(dr.Item("Cliente")), "", dr.Item("Cliente"))
                    txtDescripcionCliente.Text = IIf(IsDBNull(dr.Item("Razon")), "", dr.Item("Razon"))

                    XRenglon = dgvProductos.Rows.Add

                    WFecha = IIf(IsDBNull(dr.Item("Fecha")), "", dr.Item("Fecha"))
                    WNroObservacion = IIf(IsDBNull(dr.Item("NroObservacion")), 0, dr.Item("NroObservacion"))
                    WFechaOrd = IIf(IsDBNull(dr.Item("FechaOrd")), "", dr.Item("FechaOrd"))
                    WObservacion = IIf(IsDBNull(dr.Item("Observaciones")), "", dr.Item("Observaciones"))
                    WClave = IIf(IsDBNull(dr.Item("Clave")), "", dr.Item("Clave"))
                    WUsuario = IIf(IsDBNull(dr.Item("Usuario")), "", dr.Item("Usuario"))

                    With dgvProductos.Rows(XRenglon)

                        .Cells("Fecha").Value = WFecha
                        .Cells("FechaOrd").Value = WFechaOrd
                        .Cells("Clave").Value = WClave
                        .Cells("Observacion").Value = Trim(WObservacion)
                        .Cells("Usuario").Value = Trim(WUsuario)
                        '.Cells("NroObservacion").Value = WNroObservacion

                        If WNroObservacion = WNroObsAnt Then

                            If WNroObservacion & WRefPrimeraFilaObs = dgvProductos.Rows(XRenglon - 1).Cells("NroObservacion").Value Then
                                .Cells("Fecha").Value = ""
                            End If

                            .Cells("NroObservacion").Value = WNroObservacion & WRefPrimeraFilaObs
                        Else
                            WNroObsAnt = WNroObservacion
                            WRefPrimeraFilaObs = Helper.ceros(XRenglon, 4)
                            .Cells("NroObservacion").Value = WNroObservacion & WRefPrimeraFilaObs
                        End If

                    End With

                Loop

            End If

        Catch ex As Exception
            Throw New Exception("Hubo un problema al querer consultar la Base de Datos.")
            Exit Sub
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

    End Sub

    Private Sub txtNroProforma_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNroProforma.KeyDown

        If e.KeyData = Keys.Enter Then
            If Trim(txtNroProforma.Text) = "" Then : Exit Sub : End If

            txtNroProforma.Text = Helper.ceros(txtNroProforma.Text, 6)

            Try
                '_TraerProforma(txtNroProforma.Text)
                _MostrarHistorial()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
                Exit Sub
            End Try

            txtFechaAux.Visible = False
            txtFecha.Focus()

        ElseIf e.KeyData = Keys.Escape Then
            txtNroProforma.Text = ""
        End If

    End Sub

    'Private Sub txtFecha_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFecha.KeyDown

    '    If e.KeyData = Keys.Enter Then
    '        If Trim(txtFecha.Text.Replace("/", "")) = "" Then : Exit Sub : End If

    '        If Helper._ValidarFecha(txtFecha.Text) Then
    '            txtCliente.Focus()
    '        End If

    '    ElseIf e.KeyData = Keys.Escape Then
    '        txtFecha.Text = ""
    '    End If

    'End Sub

    'Private Sub txtFechaAux_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFechaAux.KeyDown

    '    If e.KeyData = Keys.Enter Then
    '        If Trim(txtFechaAux.Text.Replace("-", "")) = "" Then : Exit Sub : End If

    '        If WRow >= 0 And Wcol >= 0 Then

    '            'With dgvProductos
    '            '    .Rows(WRow).Cells(0).Value = txtFechaAux.Text

    '            '    Dim terminado = _BuscarTerminado(txtFechaAux.Text)

    '            '    If Not IsNothing(terminado) Then
    '            '        .Rows(WRow).Cells(0).Value = terminado("Codigo")
    '            '        .Rows(WRow).Cells(1).Value = terminado("Descripcion")

    '            '        .CurrentCell = .Rows(WRow).Cells(2)
    '            '        .Focus()

    '            '        txtFechaAux.Visible = False
    '            '        txtFechaAux.Location = New Point(680, 390) ' Lo reubicamos lejos de la grilla.
    '            '    Else
    '            '        txtFechaAux.Focus()
    '            '    End If
    '            'End With

    '        End If

    '    ElseIf e.KeyData = Keys.Escape Then
    '        txtFechaAux.Clear()
    '    End If

    'End Sub

    Private Function _EsNumero(ByVal keycode As Integer) As Boolean
        Return (keycode >= 48 And keycode <= 57) Or (keycode >= 96 And keycode <= 105)
    End Function

    Private Function _EsControl(ByVal keycode) As Boolean
        Dim valido As Boolean = False

        Select Case keycode
            Case Keys.Enter, Keys.Escape, Keys.Right, Keys.Left, Keys.Back
                valido = True
            Case Else
                valido = False
        End Select

        Return valido
    End Function

    Private Function _EsDecimal(ByVal keycode As Integer) As Boolean
        Return (keycode >= 48 And keycode <= 57) Or (keycode >= 96 And keycode <= 105) Or (keycode = 110 Or keycode = 190)
    End Function

    Private Function _EsNumeroOControl(ByVal keycode) As Boolean
        Dim valido As Boolean = False

        If _EsNumero(CInt(keycode)) Or _EsControl(keycode) Then
            valido = True
        Else
            valido = False
        End If

        Return valido
    End Function

    Private Function _EsDecimalOControl(ByVal keycode) As Boolean
        Dim valido As Boolean = False

        If _EsDecimal(CInt(keycode)) Or _EsControl(keycode) Then
            valido = True
        Else
            valido = False
        End If

        Return valido
    End Function

    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean

        With dgvProductos
            If .Focused Or .IsCurrentCellInEditMode Then ' Detectamos los ENTER tanto si solo estan en foco o si estan en edición una celda.
                .CommitEdit(DataGridViewDataErrorContexts.Commit) ' Guardamos todos los datos que no hayan sido confirmados.

                Dim iCol = .CurrentCell.ColumnIndex
                Dim iRow = .CurrentCell.RowIndex
                Dim valor = .CurrentCell.Value

                ' Limitamos los caracteres permitidos para cada una de las columnas.
                Select Case iCol
                    'Case 1
                    'If Not _EsNumeroOControl(keyData) Then
                    '    Return True
                    'End If
                    'Case 1
                    '    If Not _EsDecimalOControl(keyData) Then
                    '        Return True
                    '    End If
                    Case Else

                End Select

                If msg.WParam.ToInt32() = Keys.Enter Then

                    If valor <> "" Then

                        Select Case iCol
                            'Case 2, 3
                            '    _RecalcularTotalFila(iRow)
                            Case Else

                        End Select

                        '_NormalizarNumerosGrilla()

                    End If

                    Select Case iCol
                        'Case 3, 4
                        '    If iRow = PRODUCTOS_MAX - 1 Then
                        '        .CurrentCell = .Rows(iRow).Cells(iCol)
                        '    Else
                        '        Try
                        '            .CurrentCell = .Rows(iRow + 1).Cells(0)
                        '        Catch ex As Exception
                        '            .CurrentCell = .Rows(iRow).Cells(iCol)
                        '        End Try
                        '    End If

                        'Case Else
                        '    .CurrentCell = .Rows(iRow).Cells(iCol + 1)
                    End Select

                    Return True

                ElseIf msg.WParam.ToInt32() = Keys.Escape Then


                    Select Case iCol
                        'Case 0, 2, 3

                        '    .Rows(iRow).Cells(iCol).Value = ""

                        '    If iCol = 4 Then
                        '        .CurrentCell = .Rows(iRow).Cells(iCol - 1)
                        '    Else
                        '        .CurrentCell = .Rows(iRow).Cells(iCol + 1)
                        '    End If

                        '    .CurrentCell = .Rows(iRow).Cells(iCol)

                        Case Else

                    End Select

                End If
            End If

        End With

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    'Private Sub dgvProductos_CellEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvProductos.CellEnter
    '    With dgvProductos
    '        If e.ColumnIndex = 0 Then
    '            .ClearSelection()
    '            .CurrentCell.Style.SelectionBackColor = Color.White ' Evitamos que se vea la seleccion de la celda.
    '            Dim _location As Point = .GetCellDisplayRectangle(0, e.RowIndex, False).Location

    '            _location.Y += .Location.Y + (.CurrentCell.Size.Height / 4) - YMARGEN
    '            _location.X += .Location.X + (.CurrentCell.Size.Width - txtFechaAux.Size.Width) - XMARGEN
    '            txtFechaAux.Location = _location
    '            txtFechaAux.Text = .Rows(e.RowIndex).Cells(0).Value
    '            WRow = e.RowIndex
    '            Wcol = e.ColumnIndex
    '            txtFechaAux.Visible = True
    '            txtFechaAux.Focus()
    '        End If
    '    End With
    'End Sub

    Private Sub SoloNumero(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNroProforma.KeyPress
        If Not Char.IsNumber(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Function _Left(ByVal texto, ByVal largo) As String
        Return Microsoft.VisualBasic.Left(Trim(texto), largo)
    End Function

    Private Function _TraerProximoRenglon() As Integer
        Dim _ProximoRenglon As Integer = 0
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("")
        Dim dr As SqlDataReader

        Try

            cn.ConnectionString = _CS()
            cn.Open()
            cm.Connection = cn
            cm.CommandText = "SELECT MAX(Renglon) as Ultimo FROM ProformaExportacionHistorial WHERE Proforma = '" & txtNroProforma.Text & "'"

            dr = cm.ExecuteReader()

            If dr.HasRows Then
                dr.Read()

                _ProximoRenglon = IIf(IsDBNull(dr.Item("Ultimo")), 0, Val(dr.Item("Ultimo")) + 1)

            Else
                MsgBox("Error consulta sql", MsgBoxStyle.Information)
            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try
        Return _ProximoRenglon
    End Function

    Private Function _ProximoNroObservacion()
        Dim actual = 0
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("SELECT MAX(NroObservacion) as NroActual FROM ProformaExportacionHistorial")
        Dim dr As SqlDataReader

        Try

            cn.ConnectionString = _CS()
            cn.Open()
            cm.Connection = cn

            dr = cm.ExecuteReader()

            If dr.HasRows Then
                dr.Read()

                actual = dr.Item("NroActual")

            End If

        Catch ex As Exception
            Throw New Exception("Hubo un problema al querer traer el próximo número de Observación de la Base de Datos.")
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

        Return actual + 1

    End Function

    Private Sub btnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAceptar.Click

        Dim cn As New SqlConnection()
        Dim cm As New SqlCommand()
        Dim trans As SqlTransaction = Nothing

        Dim WClave, WRenglon, XRenglon, WNroProforma, WFecha, WFechaOrd, WCliente, WObservacion, WUsuario, WNroObservacion, WClaveObs, ZSql
        Dim WFila = "", WFilasAGrabar() As String
        Dim WParrafosAGrabar() As String

        cn.ConnectionString = _CS()
        cn.Open()

        trans = cn.BeginTransaction
        cm.Connection = cn
        cm.Transaction = trans

        ZSql = ""
        WNroProforma = Helper.ceros(txtNroProforma.Text, 6)
        WFecha = txtFecha.Text
        WFechaOrd = Helper.ordenaFecha(WFecha)
        WCliente = txtCliente.Text
        WObservacion = Trim(txtObservacion.Text)
        WClaveObs = Trim(WClaveObservacion.Text)

        WUsuario = txtUsuario.Text

        ' Chequeamos si se trata de la actualización de un renglon existente.
        ' Si tiene clave, es una actualizacion y recuperamos el nro de referencia.
        If WClaveObs = "" Then

            Try
                WNroObservacion = _ProximoNroObservacion()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
                Exit Sub
            End Try

        Else

            WNroObservacion = Mid(WClaveObs, 1, WClaveObs.ToString.Length - 4)

        End If

        Try

            ZSql = ""
            ZSql = "DELETE FROM ProformaExportacionHistorial WHERE NroObservacion = '" & WNroObservacion & "'"

            cm.CommandText = ZSql

            cm.ExecuteNonQuery()

            ' Aca calculamos y grabamos todas las filas correspondientes.

            WObservacion = Trim(WObservacion)

            ' Separamos los párrafos para conservar el formato del comentario.
            WParrafosAGrabar = WObservacion.ToString.Split(vbCrLf)

            WRenglon = 0

            For i = 0 To WParrafosAGrabar.Length - 1

                If WParrafosAGrabar(i) <> "" Then

                    ' Redondeamos el párrafo a un multiplo de 100 (max de caracteres por linea.)
                    WParrafosAGrabar(i) = LSet(WParrafosAGrabar(i), 100 - (WParrafosAGrabar(i).Length Mod 100))

                    ' Calculamos el numero de filas a grabar.
                    ReDim WFilasAGrabar((WParrafosAGrabar(i).Length / 100))

                    ' Por cada fila, cortamos trozos del parrafo cada 100 caracteres.
                    For x = 0 To WFilasAGrabar.Length - 1

                        WFilasAGrabar(x) = Trim(Mid(WParrafosAGrabar(i), (x * 100) + 1, 100))

                        ' Grabamos unicamente aquellos que no sean nueva fila y tengan contenido.
                        If Trim(WFilasAGrabar(x)) <> "" And WFilasAGrabar(x) <> vbLf Then
                            WRenglon += 1

                            XRenglon = Helper.ceros(WRenglon, 2)

                            WClave = Helper.ceros(WNroObservacion, 6) + XRenglon

                            ZSql = ""
                            ZSql = "INSERT INTO ProformaExportacionHistorial (Clave, NroObservacion, Renglon, Proforma, Usuario, Fecha, FechaOrd, Observaciones) "
                            ZSql &= "VALUES "
                            ZSql &= "('" & WClave & "'," & WNroObservacion & ",'" & XRenglon & "','" & WNroProforma & "','" & WUsuario & "','" & WFecha & "','" & WFechaOrd & "','" & WFilasAGrabar(x) & "')"

                            cm.CommandText = ZSql

                            cm.ExecuteNonQuery()
                        End If

                    Next

                End If

            Next

        Catch ex As Exception
            If Not IsNothing(trans) Then
                trans.Rollback()
            End If
            cn.Close()
            MsgBox("No se pudo grabar la observación en la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message, MsgBoxStyle.Critical)
            Exit Sub
        End Try

        trans.Commit()

        MsgBox("La observación ha sido grabada con exito.", MsgBoxStyle.Information)

        _MostrarHistorial()

        ' Despues de cargar todo el historial, no posicionamos en la primer fila de ese comentario creado/actualizado.
        For Each row As DataGridViewRow In dgvProductos.Rows
            If row.Cells("Clave").Value = WClaveObs Then
                row.Cells(0).Selected = True
                Exit For
            End If
        Next

        'btnVistaPrevia.PerformClick()

        'btnLimpiar.PerformClick()
        GrupoNuevaObs.Visible = False

    End Sub

    Private Sub btnLimpiar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpiar.Click
        txtFecha.Text = Date.Now.ToString("dd-MM-yyyy")

        txtUsuario.Text = ""
        txtObservacion.Text = ""
        WClaveObservacion.Text = ""

        txtFecha.Focus()
    End Sub

    Private Sub _CargarRenglon(ByVal rowIndex As Integer)
        Dim WRow = rowIndex
        If rowIndex + 1 <> Val(Microsoft.VisualBasic.Right(WClaveObservacion.Text, 2)) Or Trim(WClaveObservacion.Text) = "" Then
            WClaveObservacion.Text = ""

            btnLimpiar.PerformClick()

            ' Calculamos la primera fila de la observacion, segun su clave.
            rowIndex = Val(Microsoft.VisualBasic.Right(dgvProductos.Rows(rowIndex).Cells("NroObservacion").Value, 4))

            With dgvProductos.Rows(rowIndex)
                txtFecha.Text = Trim(.Cells("Fecha").Value)

                txtUsuario.Text = Trim(.Cells("Usuario").Value)
                WClaveObservacion.Text = Trim(.Cells("NroObservacion").Value)

                txtObservacion.Text = ""

                ' Recorremos hasta la ultima fila de la observacion.
                For i = rowIndex To dgvProductos.Rows.Count - 1

                    ' Salimos en la primera en la que ya no coincidan las claves de la observacion.
                    If Val(dgvProductos.Rows(i).Cells("NroObservacion").Value) <> Val(WClaveObservacion.Text) Then
                        Exit For
                    End If

                    txtObservacion.Text &= dgvProductos.Rows(i).Cells("Observacion").Value.ToString.Trim() & vbCrLf

                Next

            End With

            GrupoNuevaObs.Visible = True
        End If

    End Sub

    Private Sub txtFecha_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFecha.KeyDown

        If e.KeyData = Keys.Enter Then
            If Trim(txtFecha.Text.Replace("/", "")) = "" Then : Exit Sub : End If

            If Helper._ValidarFecha(txtFecha.Text) Then
                txtUsuario.Focus()
            End If

        ElseIf e.KeyData = Keys.Escape Then
            txtFecha.Clear()
        End If

    End Sub

    Private Sub txtUsuario_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUsuario.KeyDown

        If e.KeyData = Keys.Enter Then
            If Trim(txtUsuario.Text) = "" Then : Exit Sub : End If

            txtObservacion.Focus()

        ElseIf e.KeyData = Keys.Escape Then
            txtUsuario.Text = ""
        End If

    End Sub

    Private Sub btnEliminar_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEliminar.Click

        If Trim(WClaveObservacion.Text) = "" Or Trim(WClaveObservacion.Text).Length < 8 Then : Exit Sub : End If

        If MsgBox("¿Está seguro de que quiere eliminar la observación actual?", MsgBoxStyle.YesNo, MsgBoxStyle.Question) = MsgBoxResult.No Then
            Exit Sub
        End If

        Try
            _EliminarObservacion(WClaveObservacion.Text)
        Catch ex As Exception
            MsgBox("No se pudo eliminar la observación de la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message, MsgBoxStyle.Critical)
            Exit Sub
        End Try

    End Sub

    Private Sub _EliminarObservacion(ByVal WClave As String)
        WClave = Trim(WClave)
        If WClave = "" Or WClave.Length < 8 Then : Exit Sub : End If

        Dim cn As New SqlConnection()
        Dim cm As New SqlCommand()

        Try
            cn.ConnectionString = _CS()
            cn.Open()
            cm.Connection = cn
            cm.CommandText = "DELETE FROM ProformaExportacionHistorial WHERE Clave = '" & WClave & "'"

            cm.ExecuteNonQuery()

        Catch ex As Exception
            Throw New Exception(ex.Message)
            Exit Sub
        Finally

            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

        btnLimpiar.PerformClick()

        _MostrarHistorial()

    End Sub

    Private Sub btnArchivos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnArchivos.Click
        ' Abrimos ventana de archivos Relacionados.
        With ArchivosRelacionados
            .NroProforma = txtNroProforma.Text
            .Show()
        End With
    End Sub

    Private Sub btnNuevaObservacion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNuevaObservacion.Click
        GrupoNuevaObs.Visible = True
        btnLimpiar.PerformClick()
        txtFecha.Focus()
    End Sub

    Private Sub btnCerrarFormObs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrarFormObs.Click
        btnLimpiar.PerformClick()
        GrupoNuevaObs.Visible = False
    End Sub

    Private Sub dgvProductos_RowHeaderMouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvProductos.RowHeaderMouseDoubleClick
        ' Eliminar la observacion?
    End Sub

    Private Sub dgvProductos_CellContentDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvProductos.CellContentDoubleClick
        _CargarRenglon(e.RowIndex)
    End Sub
End Class