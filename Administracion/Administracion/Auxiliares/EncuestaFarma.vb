﻿Imports System.Configuration
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop.Outlook

Public Class EncuestaFarma

    Private WProveedor As String

    Sub New(ByVal Proveedor As String, ByVal Email As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        WProveedor = Proveedor.PadLeft(11, "0")
        txtDestinatarios.Text = Email.Trim

    End Sub

    Private Sub EncuestaFarma_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        rbEsp.Checked = True
        _CargarProductosFarma()
        dgvProductos.Focus()
    End Sub

    Private Sub _CargarProductosFarma()

        dgvProductos.Rows.Clear()

        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("SELECT DISTINCT ct.Articulo Producto, a.Descripcion, Enviar=1 FROM Cotiza ct INNER JOIN Articulo a ON a.Codigo = ct.Articulo And a.ClasificacionFarma > 0 WHERE ct.Proveedor = '" & WProveedor & "' Order by ct.Articulo")
        Dim dr As SqlDataReader
        Dim dt As New DataTable

        With dt
            .Columns.Add("Producto")
            .Columns.Add("Descripcion")
            .Columns.Add("Enviar", GetType(Boolean))
        End With

        Try

            cn.ConnectionString = _ConectarA()
            cn.Open()
            cm.Connection = cn

            dr = cm.ExecuteReader()

            If dr.HasRows Then

                dt.Load(dr)

            End If

        Catch ex As System.Exception
            Throw New System.Exception("Hubo un problema al querer consultar la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

        dgvProductos.DataSource = dt

    End Sub

    Private Sub btnCancelar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancelar.Click
        Close()
    End Sub

    Private Sub btnEnviar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnEnviar.Click

        '
        ' Validamos que el Proveedor Tenga Cargado por lo menos un mail.
        '
        If txtDestinatarios.Text.Trim = "" Then Exit Sub

        '
        ' Validamos que al menos un articulo haya sido seleccionado.
        '
        Dim Pasa As Boolean = False

        For Each row As DataGridViewRow In From row1 As DataGridViewRow In dgvProductos.Rows Where row1.Cells("Enviar").Value
            Pasa = True
        Next

        If Not Pasa Then
            MsgBox("No se ha seleccionado ningún Producto para ser Procesado.", MsgBoxStyle.Information)
            Exit Sub
        End If

        '
        ' Recorremos las materias primas y para aquellas que se encuentren marcadas, copiamos el PDF segun idioma indicado,
        ' lo renombramos como "Encuesta {Nom MP}.pdf" y lo enviamos por mail.
        '
        Dim PATH_ENCUESTAS_FARMA As String = ConfigurationManager.AppSettings("PATH_ENCUESTAS_FARMA")

        If PATH_ENCUESTAS_FARMA.Trim = "" Then
            MsgBox("No se encuentra definido la ruta en la que se debe buscar la Encuesta para enviar al Cliente.", MsgBoxStyle.Information)
            Exit Sub
        End If

        Dim WPathTemp As String = "c:\EncuestasTemp\"
        Dim WNombreArchivoOrigenEspañol As String = "Encuesta.pdf"
        Dim WNombreArchivoOrigenIngles As String = "EncuestaIngles.pdf"
        Dim WNombreArchivoOrigen As String = ""
        Dim WMail As String = txtDestinatarios.Text.Trim
        Dim WSubject As String = "Evaluación de Proveedores - Cuestionario Inicial - SURFACTAN S.A. "

        Dim WBody As String = "Estimados:  <br/>" & _
                                "Agradeceríamos completaran el formulario adjunto, a la brevedad posible,  el mismo es de suma importancia para nosotros y está confeccionado por nuestro departamento de calidad. Las compras de los insumos de referencia están condicionados a la presentación de este formulario junto con todas las certificaciones que puedan aportar.  <br/>" & _
                                "Asimismo solicitamos que  lo confeccione una persona autorizada donde figure su firma, aclaración y el cargo que ocupa. <br/>"


        If rbIng.Checked Then

            WNombreArchivoOrigen = WNombreArchivoOrigenIngles
            WSubject = "SUPPLIER EVALUATION - INITIAL QUESTIONNAIRE - SURFACTAN S.A. "

            WBody = "Dear <br/>" & _
                    "We would appreciate your completing the attached form, as soon as possible.  <br/>" & _
                    "It is of the utmost importance to us and it has been made by our quality department as per the guidelines available to be performed by any new supplier being added to our preferred/valuable supplying points.  <br/>" & _
                    "Mainly, referred to supplies/raw materials to be used in our Pharmaceutical’s production site .  <br/>" & _
                    "As from now onwards, the purchases of the reference supplies are conditioned to the presentation of this form along with, all the certifications that you can provide.  <br/>" & _
                    "We also request, that you are having all the sent/given documents, duly signed by an authorized person at your Production or Quality Control room and adding their name and responsibility level there. <br/>"

        End If

        If rbEsp.Checked Then WNombreArchivoOrigen = WNombreArchivoOrigenEspañol

        If Not File.Exists(PATH_ENCUESTAS_FARMA & WNombreArchivoOrigen) Then
            MsgBox("No existe el PDF con la Encuesta para ser enviada al Proveedor.")
            Exit Sub
        End If

        '
        ' Generamos el directorio en caso de que no exista.
        '
        Directory.CreateDirectory(WPathTemp)

        ProgressBar1.Value = 0
        ProgressBar1.Maximum = dgvProductos.Rows.Count + 1

        For Each row As DataGridViewRow In dgvProductos.Rows
            With row

                Dim WDescripcion As String = Trim(OrDefault(.Cells("Descripcion").Value, ""))
                Dim WEnviar As Boolean = OrDefault(.Cells("Enviar").Value, False)
                Dim WRutaCompletaAdjunto As String = WPathTemp & WDescripcion & ".pdf"

                If WEnviar Then

                    '
                    ' Copiamos el archivo a local y lo renombramos para luego adjuntarlo.
                    '
                    File.Copy(PATH_ENCUESTAS_FARMA & WNombreArchivoOrigen, WRutaCompletaAdjunto, True)

                    If Not File.Exists(WRutaCompletaAdjunto) Then
                        MsgBox("No se pudo copiar el archivo PDF que contiene la encuensta. No se enviará el Mail.", MsgBoxStyle.Information)
                        Continue For
                    End If

                    '
                    ' Envío el Correo, adjuntando el archivo renombrado.
                    '
                    _EnviarEmail(WMail, "", WSubject & WDescripcion, WBody, {WRutaCompletaAdjunto})

                End If

            End With
            ProgressBar1.Increment(1)
        Next

        ProgressBar1.Value = 0

        Directory.Delete(WPathTemp, True)

        MsgBox("El Proceso finalizó correctamente.", MsgBoxStyle.Information)
        
    End Sub

    Private Sub _EnviarEmail(ByVal _to As String, ByVal _bcc As String, ByVal _subject As String, ByVal _body As String, ByVal _adjuntos() As String)

        Dim _Outlook As New Application

        Try

            Dim appNameSpace As _NameSpace = _Outlook.GetNamespace("MAPI")
            appNameSpace.Logon(Nothing, Nothing, False, False)

            Dim _Mail As MailItem = _Outlook.CreateItem(OlItemType.olMailItem)

            With _Mail

                '
                ' (NO BORRAR) Obtenemos la Instancia de Inspector para que nos agrege la firma que se encuentra definida por defecto.
                '
                Dim WInspector = .GetInspector

                .To = _to
                .BCC = _bcc
                .Subject = _subject
                .HTMLBody = "<p>" & _body & "</p>" & .HTMLBody

                For Each adjunto As String In _adjuntos
                    If Trim(adjunto) <> "" Then
                        .Attachments.Add(adjunto)
                    End If
                Next

                If Val(OrDefault(ConfigurationManager.AppSettings("TESTING"), "0")) = 1 Then
                    .Display()
                Else
                    .Send()
                End If

            End With

            _Mail = Nothing

        Catch ex As System.Exception
            MsgBox("Ocurrió un problema al querer enviar la Encuesta al Proveedor '" & WProveedor & "' " & vbCrLf & vbCrLf & " Motivo: " & ex.Message, MsgBoxStyle.Exclamation)
        Finally
            _Outlook = Nothing
        End Try

    End Sub

    Private Sub txtFiltrar_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtFiltrar.KeyDown
        If e.KeyCode = Keys.Escape Then txtFiltrar.Text = ""
    End Sub

    Private Sub txtFiltrar_KeyUp(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtFiltrar.KeyUp

        Dim dt As DataTable

        dt = TryCast(dgvProductos.DataSource, DataTable)

        If dt IsNot Nothing Then dt.DefaultView.RowFilter = String.Format("Producto LIKE '%{0}%' Or Descripcion LIKE '%{0}%'", txtFiltrar.Text)

    End Sub

End Class