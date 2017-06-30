﻿Imports ClasesCompartidas
Imports System.IO
Imports System.Data.SqlClient

Public Class ListadoCuentaCorrienteProveedoresSelectivoPreparacion

    Dim varRenglon As Integer
    Dim varTotal, varSaldo, varTotalUs, varSaldoUs, varSaldoOriginal, varDife, varParidad, varParidadTotal As Double
    Dim varPago As Integer

    Private Sub ListadoCuentaCorrienteProveedoresSelectivoPreparacion_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtCodProveedor.Text = ""
        varRenglon = 0
        _CargarProveedoresPreCargados()
    End Sub

    Private Sub _CargarProveedoresPreCargados()
        Dim _Proveedores As New List(Of Object)
        Dim _CargadosHaceMasDeUnaSemana As Integer = 0
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("SELECT ps.Proveedor, ps.FechaOrd, p.Nombre FROM ProveedorSelectivo as ps, Proveedor as p WHERE ps.Proveedor = p.Proveedor")
        Dim dr As SqlDataReader

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then

                Do While dr.Read()

                    GRilla.Rows.Add()
                    GRilla.Item(0, varRenglon).Value = dr.Item("Proveedor")
                    GRilla.Item(1, varRenglon).Value = dr.Item("Nombre")
                    GRilla.CommitEdit(DataGridViewDataErrorContexts.Commit)
                    varRenglon = varRenglon + 1
                    GRilla.CurrentCell = GRilla(0, 0)

                Loop

            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

    End Sub

    Private Sub btnConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsulta.Click

        Me.Height = 690

        lstAyuda.DataSource = DAOProveedor.buscarProveedorPorNombre("")

        txtAyuda.Text = ""
        txtAyuda.Visible = True
        lstAyuda.Visible = True

        txtAyuda.Focus()

    End Sub

    Private Sub txtAyuda_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtAyuda.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            lstAyuda.DataSource = DAOProveedor.buscarProveedorPorNombre(txtAyuda.Text)
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtAyuda.Text = ""
        End If
    End Sub

    Private Function _BuscarProveedor(ByVal proveedor As String) As Object
        Dim _Proveedor As Object = Nothing
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("SELECT Proveedor, Nombre FROM Proveedor WHERE Proveedor = '" & proveedor & "' OR Nombre = '" & proveedor & "'")
        Dim dr As SqlDataReader

        SQLConnector.conexionSql(cn, cm)

        Try

            dr = cm.ExecuteReader()

            If dr.HasRows Then

                dr.Read()
                _Proveedor = {dr.Item("Proveedor"), dr.Item("Nombre")}
            End If

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally

            dr = Nothing
            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

        Return _Proveedor
    End Function

    Private Function _ProveedorYaAgregado(ByVal _CodProveedor As String) As Boolean
        Dim _YaAgregado As Boolean = False

        For Each row As DataGridViewRow In GRilla.Rows

            If Trim(row.Cells(0).Value) = Trim(_CodProveedor) Then
                _YaAgregado = True
                Exit For
            End If

        Next

        Return _YaAgregado
    End Function

    Private Sub mostrarProveedor(ByVal proveedor As String)
        Dim _Proveedor As Object = _BuscarProveedor(proveedor)
        If IsNothing(_Proveedor) Then
            MsgBox("Proveedor incorrecto")
            txtCodProveedor.Focus()
        Else
            If Not _ProveedorYaAgregado(_Proveedor(0)) Then
                GRilla.Rows.Add()
                GRilla.Item(0, varRenglon).Value = _Proveedor(0)
                GRilla.Item(1, varRenglon).Value = _Proveedor(1)
                GRilla.CommitEdit(DataGridViewDataErrorContexts.Commit)
                varRenglon = varRenglon + 1
                GRilla.CurrentCell = GRilla(0, 0)

                _ActualizarProveedoresInscriptos(_Proveedor(0))

                txtCodProveedor.Text = ""
                txtAyuda.Text = ""
            Else
                MsgBox("El Proveedor ya se encuentra agregado en el listado semanal.", MsgBoxStyle.Information)
                txtCodProveedor.Focus()
            End If
            
        End If
    End Sub

    Private Sub _ActualizarProveedoresInscriptos(ByVal CodProveedor As String)
        Dim _Fecha As String = Date.Now.ToString("dd/MM/yyyy")
        Dim _FechaOrd As String = String.Join("", _Fecha.Split("/").Reverse())
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("INSERT INTO ProveedorSelectivo (Proveedor, Fecha, FechaOrd) Values ('" & CodProveedor & "', '" & _Fecha & "', '" & _FechaOrd & "')")

        SQLConnector.conexionSql(cn, cm)

        Try

            cm.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally

            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try
    End Sub

    Private Sub lstAyuda_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstAyuda.Click
        mostrarProveedor(lstAyuda.SelectedItem.ToString())
    End Sub

    Private Sub txtCodProveedor_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCodProveedor.KeyDown

        If e.KeyData = Keys.Enter And Trim(txtCodProveedor.Text) <> "" Then
            mostrarProveedor(txtCodProveedor.Text)
        End If

    End Sub

    Private Sub btnAcepta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAcepta.Click
        Me.Close()
        MenuPrincipal.Show()
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean

        If GRilla.Focused Or GRilla.IsCurrentCellInEditMode Then

            Dim key As Integer = msg.WParam.ToInt32()

            GRilla.CommitEdit(DataGridViewDataErrorContexts.Commit)

            If key = Keys.Escape Then

                If GRilla.SelectedRows.Count > 0 Then

                    _EliminarProveedoresSeleccionados()

                End If

            End If

            Return True
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Private Sub _EliminarProveedoresSeleccionados()
        If MsgBox("¿Esta seguro que desea eliminar los proveedores seleccionados?", MsgBoxStyle.OkCancel) = DialogResult.OK Then
            For Each row As DataGridViewRow In GRilla.SelectedRows

                If Not IsNothing(row.Cells(0).Value) And Trim(row.Cells(0).Value) <> "" Then

                    If _EliminarProveedorSelectivo(row.Cells(0).Value) Then
                        GRilla.Rows.Remove(row)
                        varRenglon -= 1
                    End If

                End If

            Next
        End If
    End Sub

    Private Function _EliminarProveedorSelectivo(ByVal codProv As String) As Boolean
        Dim exito As Boolean = False
        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("DELETE FROM ProveedorSelectivo WHERE Proveedor = '" & Trim(codProv) & "'")

        SQLConnector.conexionSql(cn, cm)

        Try

            cm.ExecuteNonQuery()
            exito = True

        Catch ex As Exception
            MsgBox("Hubo un problema al querer eliminar el Proveedor del periodo actual.", MsgBoxStyle.Critical)
        Finally

            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try

        Return exito
    End Function

    Private Sub btnLimpiarTodo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpiarTodo.Click

        _LimpiarProveedoresSelectivos()

    End Sub

    Private Sub _LimpiarProveedoresSelectivos()

        If GRilla.SelectedRows.Count > 0 Then ' Eliminamos solamente los seleccionados.

            _EliminarProveedoresSeleccionados()

        Else ' Si no hay seleccionados preguntamos si quiere en realidad eliminar todos, sino no se hace nada.

            If MsgBox("¿Está seguro que quiere eliminar todos los proveedores listados?", MsgBoxStyle.OkCancel) = DialogResult.OK Then
                _LimpiarTodosLosProveedoresSelectivos()
            End If

        End If

    End Sub

    Private Sub _LimpiarTodosLosProveedoresSelectivos()

        Dim cn As SqlConnection = New SqlConnection()
        Dim cm As SqlCommand = New SqlCommand("DELETE FROM ProveedorSelectivo")

        SQLConnector.conexionSql(cn, cm)

        Try

            cm.ExecuteNonQuery()

            MsgBox("Se han eliminado todos los proveedores correspondientes a este periodo.", MsgBoxStyle.Information)

            GRilla.Rows.Clear()

            varRenglon = 0

        Catch ex As Exception
            MsgBox("Hubo un problema al querer consultar la Base de Datos.", MsgBoxStyle.Critical)
        Finally

            cn.Close()
            cn = Nothing
            cm = Nothing

        End Try
    End Sub
End Class