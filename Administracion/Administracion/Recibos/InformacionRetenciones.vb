﻿Public Class InformacionRetenciones

    Private _valor As String = ""
    Protected _TipoInformacion As String = ""

    Public Property TipoInformacion As String
        Get
            Return _TipoInformacion
        End Get
        Set(ByVal value As String)
            _TipoInformacion = Trim(value)
        End Set
    End Property


    Public Property Valor As String

        Get
            Return _valor
        End Get
        Set(ByVal value As String)
            _valor = Trim(value)
        End Set

    End Property

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        _AsignarNuevoValor()

    End Sub

    Private Sub _AsignarNuevoValor()

        With Me

            .Valor = txtInfo.Text

            .Close()

        End With

    End Sub

    Private Sub InformacionRetenciones_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LBLTipoInformacion.Text = Me.TipoInformacion

        txtInfo.Text = Me.Valor

    End Sub

    Private Sub InformacionRetenciones_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        txtInfo.Focus()
    End Sub

    Private Sub txtInfo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInfo.KeyDown

        If e.KeyData = Keys.Enter Then
            _AsignarNuevoValor()
        ElseIf e.KeyData = Keys.Escape Then
            txtInfo.Text = ""
        End If

    End Sub

    Private Sub InformacionRetenciones_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If

    End Sub
End Class