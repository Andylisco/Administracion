Imports System.Configuration
Imports System.IO
Imports CrystalDecisions.Shared
Imports SAC.Clases

Public Class Login

    Private WAbiertoporComando As Boolean = False

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click
        Close()
    End Sub

    Private Sub Login_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

        Dim tabla As New DataTable
        With tabla
            .Columns.Add("Empresa")
            .Columns.Add("Base")
            .Rows.Add("SURFACTAN S.A.", "SurfactanSa")
            .Rows.Add("PELLITAL S.A.", "Pellital_III")
        End With

        With cmbEntity
            .DataSource = tabla
            .DisplayMember = "Empresa"
            .ValueMember = "Base"
            .SelectedIndex = 0
        End With

        txtPsw.Text = ""

        Try
            '
            ' Chequeamos que se haya abierto por linea de comandos.
            '
            If Environment.GetCommandLineArgs.Length > 1 Then

                Dim WTipoLlamada As Integer = Environment.GetCommandLineArgs(1)

                
                Select Case WTipoLlamada
                    Case 1

                        Dim WTipo As String = Environment.GetCommandLineArgs(2)
                        Dim WNumero As String = Environment.GetCommandLineArgs(3)
                        Dim WAnio As String = Environment.GetCommandLineArgs(4)

                        With New NuevoSac(WTipo, WNumero, WAnio, True)
                            .Show()
                        End With

                    Case 2 ' Envío de SAC por Mail

                        Dim WTipo As String = Environment.GetCommandLineArgs(2)
                        Dim WNumero As String = Environment.GetCommandLineArgs(3)
                        Dim WAnio As String = Environment.GetCommandLineArgs(4)
                        Dim WDirecciones As String = Environment.GetCommandLineArgs(5)
                        Dim WAsunto As String = Environment.GetCommandLineArgs(6)
                        Dim WCuerpoMsj As String = Environment.GetCommandLineArgs(7)

                        Dim frm As New ConsultasVarias.VistaPrevia

                        With frm

                            .Reporte = New NuevoSACAmbos

                            .Formula = "{CargaSac.Tipo} = " & WTipo & " And {CargaSac.Numero} = " & WNumero & " And {CargaSac.Ano} = " & WAnio & ""

                        End With

                        ConsultasVarias.Clases.Conexion.EmpresaDeTrabajo = "SurfactanSa"
                        ConsultasVarias.Clases.Helper._ExportarReporte(frm, ConsultasVarias.Clases.Enumeraciones.FormatoExportacion.PDF, WTipo & WNumero & WAnio & ".pdf", "C:\TempReclamos\")

                        If IO.File.Exists("C:\TempReclamos\" & WTipo & WNumero & WAnio & ".pdf") Then
                            ConsultasVarias.Clases.Helper._EnviarEmail(WDirecciones, WAsunto, WCuerpoMsj, {"C:\TempReclamos\" & WTipo & WNumero & WAnio & ".pdf"}, False)
                        Else
                            MsgBox("No se encontró el archivo " & "C:\TempReclamos\" & WTipo & WNumero & WAnio & ".pdf")
                        End If

                    Case 3

                        Dim WNumero As String = Environment.GetCommandLineArgs(2)
                        Dim WDirecciones As String = Environment.GetCommandLineArgs(3)
                        Dim WAsunto As String = Environment.GetCommandLineArgs(4)
                        Dim WCuerpoMsj As String = Environment.GetCommandLineArgs(5)

                        ConsultasVarias.Clases.Conexion.EmpresaDeTrabajo = "SurfactanSa"

                        Dim frm As New ConsultasVarias.VistaPrevia

                        With frm

                            .Reporte = New ReclamoClienteAvisoMail

                            .Formula = "{CentroReclamos.Numero} = " & WNumero & ""

                        End With

                        '
                        ' De aca, van a venir los mails de los responsables. Agregamos las que faltan.
                        '
                        WDirecciones = _DefinirMailsExtras(WNumero, WDirecciones)

                        ConsultasVarias.Clases.Helper._ExportarReporte(frm, ConsultasVarias.Clases.Enumeraciones.FormatoExportacion.PDF, WNumero & "Reclamo.pdf", "C:\TempReclamos\")

                        If IO.File.Exists("C:\TempReclamos\" & WNumero & "Reclamo.pdf") Then
                            ConsultasVarias.Clases.Helper._EnviarEmail(WDirecciones, WAsunto, WCuerpoMsj, {"C:\TempReclamos\" & WNumero & "Reclamo.pdf"}, True)
                        Else
                            MsgBox("No se encontró el archivo " & "C:\TempReclamos\" & WNumero & "Reclamo.pdf")
                        End If
                    Case 4

                        Dim WNumero As String = Environment.GetCommandLineArgs(2)
                        Dim WDirecciones As String = Environment.GetCommandLineArgs(3)
                        Dim WAsunto As String = Environment.GetCommandLineArgs(4)
                        Dim WCuerpoMsj As String = Environment.GetCommandLineArgs(5)

                        Dim frm As New ConsultasVarias.VistaPrevia

                        With frm
                            .Reporte = New ReporteINCIndividual
                            .Formula = "{CargaIncidencias.Incidencia} = " & WNumero

                            .Reporte.SetParameterValue("MostrarPosiblesUsos", 1)
                            .Reporte.SetParameterValue("MostrarAcciones", 0)

                            Dim WNombreArchivo = String.Format("INC {0} - {1}", WNumero.PadLeft(4, "0"), Date.Now.ToString("dd-MM-yyyy"))

                            Dim WRuta = "C:/tempIndice/"

                            WNombreArchivo &= ".pdf"

                            If Directory.Exists(WRuta) Then Directory.Delete(WRuta, True)

                            Directory.CreateDirectory(WRuta)

                            ConsultasVarias.Clases.Conexion.EmpresaDeTrabajo = "SurfactanSa"
                            ConsultasVarias.Clases.Helper._ExportarReporte(frm, ConsultasVarias.Clases.Enumeraciones.FormatoExportacion.PDF, WNombreArchivo, WRuta)

                            If File.Exists(WRuta & WNombreArchivo) Then
                                ConsultasVarias.Clases.Helper._EnviarEmail(WDirecciones, WAsunto, WCuerpoMsj, {WRuta & WNombreArchivo}, False)
                            Else
                                MsgBox("No se encontró el archivo " & WRuta & WNombreArchivo)
                            End If

                        End With

                End Select

                btnCancel_Click(Nothing, Nothing)


            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            btnCancel_Click(Nothing, Nothing)
        End Try

    End Sub

    Private Function _DefinirMailsExtras(ByVal WNumero As String, ByVal WDirecciones As String) As String

        '
        ' Completamos los datos de las direcciones cuando lleguen.
        '
        WDirecciones = ""
        '
        ' Definimos los mails obligatorios, según tipo Producto Terminado indicado en Reclamo.
        '
        Dim WReclamo As DataRow = GetSingle("SELECT * FROM CentroReclamos WHERE Numero = '" & WNumero & "'")

        If WReclamo Is Nothing Then
            Return WDirecciones
        End If

        Dim WTerminado As String = OrDefault(WReclamo.Item("Producto"), "")

        Dim WTipoProd As String = _DeterminarTipoProd(WTerminado)

        Select Case WTipoProd
            Case "PG"
                For Each d As String In {"drodriguez@surfactan.com.ar", "ctomaszek@surfactan.com.ar"}
                    If Not WDirecciones.Contains(d) Then WDirecciones &= d & ";"
                Next
            Case "CO"
                For Each d As String In {"hfondevielle@surfactan.com.ar", "textil@surfactan.com.ar", "hsuarez@surfactan.com.ar"}
                    If Not WDirecciones.Contains(d) Then WDirecciones &= d & ";"
                Next
            Case "FA"
                For Each d As String In {"drodriguez@surfactan.com.ar", "dsuarez@surfactan.com.ar", "msosa@surfactan.com.ar", "scoppiello@surfactan.com.ar", "svarela@surfactan.com.ar", "supcc@surfactan.com.ar"}
                    If Not WDirecciones.Contains(d) Then WDirecciones &= d & ";"
                Next
            Case Else ' El resto de los Productos los tomamos como Químicos.
                For Each d As String In {"drodriguez@surfactan.com.ar", "dsuarez@surfactan.com.ar", "mescames@surfactan.com.ar", "hmuller@surfactan.com.ar", "svarela@surfactan.com.ar", "supcc@surfactan.com.ar"}
                    If Not WDirecciones.Contains(d) Then WDirecciones &= d & ";"
                Next
        End Select

        For Each d As String In ("ebiglieri@surfactan.com.ar; calidad2@surfactan.com.ar; calidad@surfactan.com.ar; isocalidad@surfactan.com.ar; wbarosio@surfactan.com.ar; wbarosio@gmail.com; lsantos@surfactan.com.ar; juanfs@surfactan.com.ar").Split(";")
            If Not WDirecciones.Contains(Trim(d)) Then WDirecciones &= Trim(d) & ";"
        Next

        '
        ' Busco el vendedor definido para el Cliente.
        '
        Dim WCliente As DataRow = GetSingle("SELECT * FROM Cliente WHERE Cliente = '" & OrDefault(WReclamo.Item("Cliente"), "") & "'")

        If WCliente IsNot Nothing Then
            Dim WVendedor As DataRow = GetSingle("SELECT * FROM Vendedor WHERE Vendedor  = '" & OrDefault(WCliente.Item("Vendedor"), 99) & "'")

            If WVendedor IsNot Nothing Then
                If WTipoProd = "FA" Then
                    WDirecciones &= "hsein@surfactan.com.ar; grodriguez@surfactan.com.ar"
                Else

                    If Trim(OrDefault(WVendedor.Item("Email1"), "")) <> "" Then
                        If Not WDirecciones.Contains(OrDefault(WVendedor.Item("Email1"), "")) Then WDirecciones &= OrDefault(WVendedor.Item("Email1"), "")
                    End If

                    If Trim(OrDefault(WVendedor.Item("Email2"), "")) <> "" Then
                        If Not WDirecciones.Contains(OrDefault(WVendedor.Item("Email2"), "")) Then WDirecciones &= OrDefault(WVendedor.Item("Email2"), "")
                    End If

                End If
            End If

        End If

        Return WDirecciones

    End Function

    Private Function _DeterminarTipoProd(ByVal WTerminado As String) As String

        WTerminado = UCase(WTerminado)

        Dim XTipoPro As String = "PT"

        Dim XCodigo As Integer = Val(Mid(WTerminado, 4, 5))

        If Microsoft.VisualBasic.Left(WTerminado, 2) <> "PT" Then
            Select Case Microsoft.VisualBasic.Left(WTerminado, 2)
                Case "DY", "DS"
                    XTipoPro = "CO"
                Case "QC", "YF"
                    XTipoPro = "FA"
                Case "YQ", "YH", "YP"
                    XTipoPro = "PT"
                Case Else
                    XTipoPro = "PT"
            End Select
        Else
            If XCodigo >= 0 And XCodigo <= 999 Then
                XTipoPro = "CO"
            Else
                If XCodigo >= 11000 And XCodigo <= 12999 Then
                    XTipoPro = "CO"
                Else
                    If XCodigo >= 25000 And XCodigo <= 25999 Then
                        XTipoPro = "FA"
                    Else
                        If XCodigo >= 2300 And XCodigo <= 2399 Then
                            XTipoPro = "BI"
                        Else
                            XTipoPro = "PT"
                        End If
                    End If
                End If
            End If
        End If

        Dim ZLinea As Integer = 0
        Dim WTerm As DataRow = GetSingle("SELECT Linea FROM Terminado WHERE Codigo = '" & WTerminado & "'")

        If WTerm IsNot Nothing Then ZLinea = OrDefault(WTerm.Item("Linea"), 0)

        Select Case ZLinea
            Case 8
                XTipoPro = "PG"
            Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                XTipoPro = "FA"
        End Select

        Return XTipoPro
    End Function

    Private Sub btnAccept_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAccept.Click

        'Globals.empresa = cmbEntity.SelectedItem

        Conexion.EmpresaDeTrabajo = CType(cmbEntity.SelectedItem, DataRowView).Item("Base")

        ' Validamos la contraseña.

        If Not Conexion._Login(txtPsw.Text, 4) Then
            MsgBox("La contraseña no es correcta.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If _EsPellital Then
            ' En caso de ser PELLITAL, validamos que la conexion se haga desde una pc con Permisos. Los mismos se definen segun nombre de PC.

            If Not _PermisosPellitalValidos() Then

                MsgBox("No tiene los permisos necesarios para poder ingresar a esta Empresa.", MsgBoxStyle.Exclamation)

                Exit Sub

            End If

        End If

        MenuPrincipal.Show()

        Close()

    End Sub

    Private Function _PermisosPellitalValidos() As Boolean

        Dim WPermitidos() = ConfigurationManager.AppSettings("PERMISOS_PELLITAL").ToString.Split(",")
        Dim WNombrePC = getNombrePC

        Return (From N In WPermitidos Where UCase(Trim(N)) = UCase(Trim(WNombrePC))).Any()

    End Function

    Private Sub Login_Shown(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Shown
        If WAbiertoporComando Then
            btnCancel_Click(Nothing, Nothing)
        End If
        txtPsw.Focus()
    End Sub

    Private Sub txtPsw_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtPsw.KeyDown

        If e.KeyData = Keys.Enter Then
            If Trim(txtPsw.Text) = "" Then : Exit Sub : End If

            btnAccept.PerformClick()

        ElseIf e.KeyData = Keys.Escape Then
            txtPsw.Text = ""
        End If

    End Sub

End Class