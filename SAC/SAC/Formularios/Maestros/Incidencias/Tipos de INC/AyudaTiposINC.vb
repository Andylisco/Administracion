Public Class AyudaTiposINC

    Private WINCRechazo As Boolean

    Sub New(Optional ByVal INCRechazo As Boolean = False)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        WINCRechazo = INCRechazo

    End Sub

    Private Sub ListadoTiposSolicitud_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Label1.Text = NombreEmpresa()

        _CargarTipos()
    End Sub

    Private Sub _CargarTipos()

        Dim WFiltro As String = " WHERE Tipo = '2' "

        If Not WINCRechazo Then
            WFiltro = " WHERE Tipo <> '2' "
        End If

        Dim WTipos As DataTable = GetAll("SELECT Tipo As Codigo, LTRIM(RTRIM(Descripcion)) Descripcion FROM TiposINC " & WFiltro & " Order by Tipo")

        dgvResponsables.DataSource = WTipos

        txtFiltrar.Text = ""

    End Sub

    Private Sub btnCerrar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCerrar.Click
        Close()
    End Sub

    Private Sub txtFiltrar_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles txtFiltrar.KeyPress
        Try
            Dim tabla = CType(dgvResponsables.DataSource, DataTable)

            If Not IsNothing(tabla) Then
                tabla.DefaultView.RowFilter = String.Format("Descripcion LIKE '%{0}%' " & " Or Convert(Codigo, System.String) LIKE '%{0}%'", txtFiltrar.Text)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub txtFiltrar_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtFiltrar.KeyDown

        If e.KeyData = Keys.Escape Then
            txtFiltrar.Text = ""
        End If
    End Sub

    Private Sub dgvTIpos_CellMouseDoubleClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles dgvResponsables.CellMouseDoubleClick

        If e.ColumnIndex < 0 Then Exit Sub

        dgvTIpos_RowHeaderMouseDoubleClick(sender, e)

    End Sub

    Private Sub dgvTIpos_RowHeaderMouseDoubleClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles dgvResponsables.RowHeaderMouseDoubleClick

        If e.RowIndex < 0 Then Exit Sub

        Dim WOWner As IAyudaTipoINC = TryCast(Owner, iayudatipoinc)

        With dgvResponsables.Rows(e.RowIndex)
            If WOWner IsNot Nothing Then WOWner._ProcesarAyudaTipoINC(OrDefault(.Cells("Codigo").Value, 0), Trim(OrDefault(.Cells("Descripcion").Value, "")))
        End With

        btnCerrar_Click(Nothing, Nothing)

    End Sub

    Private Sub AyudaTiposINC_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        txtFiltrar.Focus()
    End Sub

End Class