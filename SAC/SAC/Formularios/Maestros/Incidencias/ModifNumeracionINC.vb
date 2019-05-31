Imports ConsultasVarias

Public Class ModifNumeracionINC : Implements IIngresoClaveSeguridad

    Private WNumActual As String
    Private WAutorizado As Boolean = False

    Private WTipo As String = ""
    Private WAnio As String = ""

    Sub New(ByVal TipoINC As String, ByVal Anio As String, ByVal Num As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        txtNumActual.Text = Trim(Num)

        WTipo = TipoINC
        WAnio = Anio

    End Sub

    Private Sub btnCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelar.Click
        Close()
    End Sub

    Private Sub ModifNumeracionINC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        WAutorizado = False
    End Sub

    Private Sub ModifNumeracionINC_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        txtNuevoNum.Focus()
    End Sub

    Private Sub txtNuevoNum_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNuevoNum.KeyDown

        If e.KeyData = Keys.Enter Then
            If Val(txtNuevoNum.Text) = 0 Then : Exit Sub : End If

            btnModif.PerformClick()

        ElseIf e.KeyData = Keys.Escape Then
            txtNuevoNum.Text = ""
        End If

    End Sub

    Private Sub btnModif_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModif.Click

        If Val(txtNumActual.Text) = 0 Or Val(txtNuevoNum.Text) = 0 Then Exit Sub

        Dim WActual As DataRow = GetSingle("SELECT Incidencia FROM CargaIncidencias WHERE Numero = '" & txtNumActual.Text & "' And Tipo = '" & WTipo & "' And Ano = '" & WAnio & "'")

        If WActual Is Nothing Then
            MsgBox("No se encuentra el Informe de No Conformidad al que se quiere modificar la numeración.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim WNuevo As DataRow = GetSingle("SELECT Incidencia FROM CargaIncidencias WHERE Incidencia = '" & txtNuevoNum.Text & "' And Tipo = '" & WTipo & "' And Ano = '" & WAnio & "'")

        If WNuevo IsNot Nothing Then
            MsgBox("Ya existe un Informe de no Conformidad con el número que indicó en 'NUEVA NUMERACIÓN'.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If Not WAutorizado Then

            With New IngresoClaveSeguridad
                .Show(Me)
            End With

            Exit Sub
        End If

        ExecuteNonQueries("UPDATE CargaIncidencias SET Numero = '" & txtNuevoNum.Text & "' WHERE Numero = '" & txtNumActual.Text & "' And Tipo = '" & WTipo & "' And Ano = '" & WAnio & "'")

        Dim WOwner As IModifNumeracionINC = TryCast(Owner, IModifNumeracionINC)

        If WOwner IsNot Nothing Then WOwner._ProcesarModifNumeracionINC(txtNuevoNum.Text)

        Close()

    End Sub

    Public Sub _ProcesarIngresoClaveSeguridad(ByVal WClave As Object) Implements IIngresoClaveSeguridad._ProcesarIngresoClaveSeguridad

        WAutorizado = UCase(Trim(WClave)) = UCase(Trim(Operador.Clave))

        btnModif.PerformClick()

    End Sub
End Class