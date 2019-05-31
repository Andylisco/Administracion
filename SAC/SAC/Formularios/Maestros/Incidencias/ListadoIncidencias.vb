Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class ListadoIncidencias : Implements INuevaIncidencia, ISeleccionNuevaIncidencia, IExportarIndice

    Private Sub btnCerrar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCerrar.Click
        Close()
    End Sub

    Private Sub ListadoIncidencias_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        txtDesdeFecha.Text = ""
        txtHastaFecha.Text = ""
        txtDesdeAño.Text = Date.Now.ToString("yyyy")
        txtHastaAño.Text = txtDesdeAño.Text

        rbAñosCompletos.Checked = True
        rbAñosCompletos_Click(Nothing, Nothing)

        _CargarEstados()
        _CargarTipos()
        _CargarPlantas()

        cmbOrdenI.SelectedIndex = 1
        cmbOrdenII.SelectedIndex = 0
        cmbOrdenIII.SelectedIndex = 2

        For Each o As CheckedListBox In {clbEstados, clbTipos, clbPlantas}
            For i = 0 To o.Items.Count - 1
                o.SetItemCheckState(i, CheckState.Checked)
            Next
        Next

        btnFiltrar_Click(Nothing, Nothing)
    End Sub

    Private Sub CopiarConCabecerasToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CopiarConCabecerasToolStripMenuItem.Click
        dgvIncidencias.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        _CopiarSeleccion()
    End Sub

    Private Sub CopiarSóloDatosToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CopiarSóloDatosToolStripMenuItem.Click
        dgvIncidencias.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        _CopiarSeleccion()
    End Sub

    Private Sub _CopiarSeleccion()
        If dgvIncidencias.GetCellCount(DataGridViewElementStates.Selected) > 0 Then

            dgvIncidencias.RowHeadersVisible = False

            Dim WData = dgvIncidencias.GetClipboardContent()

            If Not IsNothing(WData) Then
                Clipboard.SetDataObject(WData)
            End If

            dgvIncidencias.RowHeadersVisible = True

        End If
    End Sub

    Private Sub _CargarPlantas()
        Dim WPlantas As New DataTable

        With WPlantas
            .Columns.Add("Codigo")
            .Columns.Add("Descripcion")
            .Rows.Add(0, "TODOS")
            If _EsPellital() Then
                .Rows.Add(2, "PELLITAL I")
                .Rows.Add(4, "PELLITAL II")
                .Rows.Add(8, "PELLITAL III")
                .Rows.Add(9, "PELLITAL V")
            Else
                .Rows.Add(1, "SURFACTAN I")
                .Rows.Add(3, "SURFACTAN II")
                .Rows.Add(5, "SURFACTAN III")
                .Rows.Add(6, "SURFACTAN IV")
                .Rows.Add(7, "SURFACTAN V")
                .Rows.Add(10, "SURFACTAN VI")
                .Rows.Add(11, "SURFACTAN VII")
            End If

            '.DefaultView.Sort = "Codigo"
        End With

        With CType(clbPlantas, ListBox)
            .DataSource = WPlantas
            .DisplayMember = "Descripcion"
            .ValueMember = "Codigo"
        End With
    End Sub

    Private Sub _CargarTipos()
        Dim WTipos As DataTable = GetAll("SELECT Tipo As Codigo, Descripcion FROM TiposINC ORDER BY Tipo")

        WTipos.Rows.Add(-1, "TODOS")

        WTipos.DefaultView.Sort = "Codigo"

        With CType(clbTipos, ListBox)
            .DataSource = WTipos
            .DisplayMember = "Descripcion"
            .ValueMember = "Codigo"
        End With
    End Sub

    Private Sub _CargarEstados()
        Dim WEstados As DataTable = GetAll("SELECT Estado As Codigo, Descripcion FROM EstadosINC ORDER BY Estado")

        WEstados.Rows.Add(-1, "TODOS")

        WEstados.DefaultView.Sort = "Codigo"

        With CType(clbEstados, ListBox)
            .DataSource = WEstados
            .DisplayMember = "Descripcion"
            .ValueMember = "Codigo"
        End With
    End Sub

    Private Function _GenerarStringOrdenamiento(ByVal cmbOrd As ComboBox) As String

        Select Case cmbOrd.SelectedIndex
            Case 0
                Return "Numero"
            Case 1
                Return "TipoINC"
            Case 2
                Return "Estado"
            Case 3
                Return "Empresa"
            Case Else
                Return ""
        End Select

    End Function

    Private Sub btnFiltrar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFiltrar.Click

        ExecuteNonQueries("UPDATE CargaIncidencias SET EmpresaIncidencia = Empresa WHERE EmpresaIncidencia IS NULL")

        Dim WDesdeFecha As String = ordenaFecha(txtDesdeFecha.Text)
        Dim WHastaFecha As String = ordenaFecha(txtHastaFecha.Text)

        If rbAñosCompletos.Checked Then

            If Val(txtDesdeAño.Text) <> 0 And Val(txtHastaAño.Text) <> 0 Then
                WDesdeFecha = ordenaFecha("01/01/" & txtDesdeAño.Text)
                WHastaFecha = ordenaFecha("31/12/" & txtHastaAño.Text)
            Else
                WDesdeFecha = 0
                WHastaFecha = 0
            End If

        End If

        '
        ' Armamos el filtro de los Estados en caso que corresponda.
        '
        Dim WFiltroEstados As String = ""

        If clbEstados.GetItemCheckState(0) <> CheckState.Checked Then
            For Each o As Object In clbEstados.CheckedItems
                WFiltroEstados &= CType(o, DataRowView).Item("Codigo") & ","
            Next
            WFiltroEstados = WFiltroEstados.TrimEnd(",")
        End If

        '
        ' Armamos el filtro de las Plantas en caso que corresponda.
        '
        Dim WFiltroPlantas As String = ""

        If clbPlantas.GetItemCheckState(0) <> CheckState.Checked Then
            For Each o As Object In clbPlantas.CheckedItems
                WFiltroPlantas &= CType(o, DataRowView).Item("Codigo") & ","
            Next
            WFiltroPlantas = WFiltroPlantas.TrimEnd(",")
        End If

        '
        ' Armamos el filtro de los Tipos en caso que corresponda.
        '
        Dim WFiltroTipos As String = ""

        If clbTipos.GetItemCheckState(0) <> CheckState.Checked Then
            For Each o As Object In clbTipos.CheckedItems
                WFiltroTipos &= CType(o, DataRowView).Item("Codigo") & ","
            Next
            WFiltroTipos = WFiltroTipos.TrimEnd(",")
        End If

        Dim ZSql = ""
        ZSql = "SELECT Incidencia, Ano As Anio, TipoInc As Tipo, Numero, Fecha, Estado, Titulo, Referencia, " _
            & " DescEstado = CASE ISNULL(Estado, 0) WHEN 1 THEN 'Genera SAC' WHEN 2 THEN 'No Genera SAC' ELSE 'Pend. Análisis' END, ISNULL(EmpresaIncidencia, Empresa) Empresa " _
            & " FROM CargaIncidencias WHERE Renglon = 1 "

        If Val(WDesdeFecha) <> 0 And Val(WHastaFecha) <> 0 Then
            ZSql &= " And FechaOrd BETWEEN '" & WDesdeFecha & "' And '" & WHastaFecha & "' "
        End If

        If WFiltroEstados.Trim <> "" Then ZSql &= " And Estado IN (" & WFiltroEstados & ") "
        If WFiltroTipos.Trim <> "" Then ZSql &= " And Tipo IN (" & WFiltroTipos & ") "
        If WFiltroPlantas.Trim <> "" Then ZSql &= " And EmpresaIncidencia IN (" & WFiltroPlantas & ") "

        ZSql &= " Order by Ano, "
        ZSql &= _GenerarStringOrdenamiento(cmbOrdenI)
        If cmbOrdenII.SelectedIndex <> cmbOrdenI.SelectedIndex Then ZSql &= ", " & _GenerarStringOrdenamiento(cmbOrdenII)
        If Not {cmbOrdenI.SelectedIndex, cmbOrdenII.SelectedIndex}.Contains(cmbOrdenII.SelectedIndex) Then ZSql &= ", " & _GenerarStringOrdenamiento(cmbOrdenIII)

        Dim WIncidencias As DataTable = GetAll(ZSql)

        WIncidencias.Columns.Add("Planta")
        WIncidencias.Columns.Add("DescTipo")

        For Each row As DataRow In WIncidencias.Rows
            With row
                .Item("Planta") = _GenerarDescEmpresa(.Item("Empresa"))

                Dim WTipo As DataRow = GetSingle("SELECT Descripcion FROM TiposInc WHERE Tipo = '" & OrDefault(.Item("Tipo"), "") & "'")

                .Item("DescTipo") = ""

                If WTipo IsNot Nothing Then .Item("DescTipo") = Microsoft.VisualBasic.Left(OrDefault(WTipo.Item("Descripcion"), ""), 20)

            End With
        Next

        dgvIncidencias.DataSource = WIncidencias

        txtDesdeFecha.Focus()

    End Sub

    Private Function _GenerarDescEmpresa(ByVal emp As Integer) As String

        Select Case emp
            Case 1
                Return "Surfactan I"
            Case 2
                Return "Pellital I"
            Case 3
                Return "Surfactan II"
            Case 4
                Return "Pellital II"
            Case 5
                Return "Surfactan III"
            Case 6
                Return "Surfactan IV"
            Case 7
                Return "Surfactan V"
            Case 8
                Return "Pellital III"
            Case 9
                Return "Pellital V"
            Case 10
                Return "Surfactan VI"
            Case 11
                Return "Surfactan VII"
            Case Else
                Return "No Definida"
        End Select

    End Function

    Private Sub ListadoIncidencias_Shown(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Shown
        txtDesdeFecha.Focus()
    End Sub

    Private Sub txtDesde_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtDesdeFecha.KeyDown

        If e.KeyData = Keys.Enter Then
            If txtDesdeFecha.Text.Replace(" ", "").Length < 10 Then : Exit Sub : End If

            If Not _ValidarFecha(txtDesdeFecha.Text) Then Exit Sub

            txtHastaFecha.Focus()

        ElseIf e.KeyData = Keys.Escape Then
            txtDesdeFecha.Text = ""
        End If

    End Sub

    Private Sub dgvIncidencias_CellDoubleClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles dgvIncidencias.CellDoubleClick
        If e.RowIndex < 0 Then Exit Sub

        Dim WIncidencia As Integer = OrDefault(dgvIncidencias.CurrentRow.Cells("Incidencia").Value, 0)
        Dim WTipo As Integer = OrDefault(dgvIncidencias.CurrentRow.Cells("Tipo").Value, -1)

        If WIncidencia = 0 Then Exit Sub
        If WTipo = -1 Then Exit Sub

        Select Case WTipo
            Case TipoIncidencias.General
                With New DetallesIncidencia(WIncidencia)
                    .Show(Me)
                End With
            Case TipoIncidencias.RechazoMP
                With New DetallesIncidenciaRechazoMP(WIncidencia)
                    .Show(Me)
                End With
        End Select

    End Sub

    Public Sub _ProcesarNuevaIncidencia(ByVal WIncidencia As Object) Implements INuevaIncidencia._ProcesarNuevaIncidencia
        btnFiltrar.PerformClick()
    End Sub

    Private Sub btnNuevaIncidencia_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnNuevaIncidencia.Click
        With New SeleccionTipoIncidencia
            .Show(Me)
        End With
    End Sub

    Public Sub ISeleccionNuevaIncidencia__ProcesarNuevaIncidencia(ByVal opt As Object) Implements ISeleccionNuevaIncidencia._ProcesarNuevaIncidencia
        Select Case CType(opt, TipoIncidencias)
            Case TipoIncidencias.General
                With New DetallesIncidencia
                    .Show(Me)
                End With
            Case TipoIncidencias.RechazoMP
                With New DetallesIncidenciaRechazoMP
                    .Show(Me)
                End With
        End Select
    End Sub


    Private Sub clbEstados_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles clbEstados.MouseUp
        With CType(sender, CheckedListBox)
            If .SelectedIndex = 0 Then
                For i = 1 To .Items.Count - 1

                    If .GetItemCheckState(0) = CheckState.Checked Then
                        .SetItemChecked(i, True)
                    Else
                        .SetItemChecked(i, False)
                    End If

                Next
            Else
                .SetItemChecked(0, False)
            End If
        End With
    End Sub

    Private Sub clbTipos_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles clbTipos.MouseUp, clbPlantas.MouseUp
        With CType(sender, CheckedListBox)
            If .SelectedIndex = 0 Then
                For i = 1 To .Items.Count - 1

                    If .GetItemCheckState(0) = CheckState.Checked Then
                        .SetItemChecked(i, True)
                    Else
                        .SetItemChecked(i, False)
                    End If

                Next
            Else
                .SetItemChecked(0, False)
            End If
        End With
    End Sub

    Private Sub txtDesdeAño_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtDesdeAño.KeyDown

        If e.KeyData = Keys.Enter Then
            If Val(txtDesdeAño.Text) = 0 Then : Exit Sub : End If

            txtHastaAño.Focus()

        ElseIf e.KeyData = Keys.Escape Then
            txtDesdeAño.Text = ""
        End If

    End Sub

    Private Sub txtHastaFecha_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtHastaFecha.KeyDown

        If e.KeyData = Keys.Enter Then
            If txtHastaFecha.Text.Replace(" ", "").Length < 10 Then : Exit Sub : End If

            btnFiltrar.PerformClick()

        ElseIf e.KeyData = Keys.Escape Then
            txtHastaFecha.Text = ""
        End If

    End Sub

    Private Sub txtHastaAño_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtHastaAño.KeyDown

        If e.KeyData = Keys.Enter Then
            If Val(txtHastaAño.Text) = 0 Then : Exit Sub : End If

            btnFiltrar.PerformClick()

        ElseIf e.KeyData = Keys.Escape Then
            txtHastaAño.Text = ""
        End If

    End Sub

    Private Sub rbAñosCompletos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEntreFechas.Click, rbAñosCompletos.Click
        txtDesdeAño.Enabled = True
        txtHastaAño.Enabled = True
        txtDesdeFecha.Enabled = True
        txtHastaFecha.Enabled = True

        If rbEntreFechas.Checked Then
            txtDesdeAño.Enabled = False
            txtHastaAño.Enabled = False
            txtDesdeFecha.Focus()
        End If

        If rbAñosCompletos.Checked Then
            txtDesdeFecha.Enabled = False
            txtHastaFecha.Enabled = False
            txtDesdeAño.Focus()
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        With ExportarIndice
            .Show(Me)
        End With
    End Sub


    Public Sub _ProcesarExportarIndice(ByVal WAgrupar As Boolean, ByVal WFormato As Object) Implements IExportarIndice._ProcesarExportarIndice

        Try

            If dgvIncidencias.Rows.Count = 0 Then Exit Sub

            Dim frm As New VistaPrevia
            Dim rpt As reportdocument

            If WAgrupar Then
                rpt = New ReporteIndiceINCsAgrupado
            Else
                rpt = New ReporteIndiceINCs
            End If

            Dim ds As DataSet = New DetallesIndiceGral

            For Each row As DataGridViewRow In dgvIncidencias.Rows
                Dim r As DataRow = ds.Tables("Detalles").NewRow

                With r
                    .Item("Tipo") = If(row.Cells("DescTipo").Value, "")
                    .Item("Nro") = If(row.Cells("Incidencia").Value, 0)
                    .Item("Fecha") = If(row.Cells("Fecha").Value, "")
                    .Item("Estado") = If(row.Cells("DescEstado").Value, "")
                    .Item("Titulo") = If(row.Cells("Titulo").Value, "")
                    .Item("Referencia") = If(row.Cells("Referencia").Value, "")
                    .Item("Emisor") = If(row.Cells("Planta").Value, "")
                End With
                ds.Tables("Detalles").Rows.Add(r)
            Next

            rpt.SetDataSource(ds)

            With frm
                .Reporte = rpt
            End With

            If WFormato = 5 Then

                Dim WRuta = "C:/tempIndice/"

                If Directory.Exists(WRuta) Then Directory.Delete(WRuta, True)

                Directory.CreateDirectory(WRuta)

                Dim WNombreArchivo = String.Format("Indice Gral de INC's al {0}.pdf", Date.Now.ToString("dd-MM-yyyy"))

                frm.Exportar(WNombreArchivo, ExportFormatType.PortableDocFormat, WRuta)

                With VistaPrevia
                    .EnviarPorEmail(WRuta & WNombreArchivo)
                End With

            Else

                _ExportarReporte(frm, WFormato)

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub _ExportarReporte(ByVal frm2 As VistaPrevia, ByVal WFormato As Object)

        With frm2

            Dim WNombreArchivo = String.Format("Indice Gral de INC's al {0}", Date.Now.ToString("dd-MM-yyyy"))

            Select Case WFormato
                Case 0 ' Imprimir
                    .Imprimir()
                Case 1 ' Mostrar por Pantalla
                    .Mostrar()
                Case 2 ' PDF
                    .Exportar(WNombreArchivo, ExportFormatType.PortableDocFormat)
                Case 3 ' Excel
                    .Exportar(WNombreArchivo, ExportFormatType.Excel)
                Case 4 ' Word
                    .Exportar(WNombreArchivo, ExportFormatType.WordForWindows)
            End Select

        End With
    End Sub

End Class