﻿Imports System.Data.SqlClient

Namespace Clases

    Public Class Conexion

        Public Shared EsPellital As Boolean = False
        Private Shared _EmpresaTrabajo As String = ""
        Private Shared _Operador As String = ""
        Public Shared WAtributosOperador(10, 100) As String

        Public Shared Property Operador As String

            Get

                Return IIf(_Operador = "", 0, _Operador)

            End Get

            Private Set(ByVal value As String)
                _Operador = value
            End Set

        End Property

        Public Shared Property EmpresaDeTrabajo As String
            Get
                Return _EmpresaTrabajo

            End Get
            Set(ByVal value As String)

                _EmpresaTrabajo = DeterminarBasePara(value)

                If UCase(_EmpresaTrabajo).StartsWith("SURFACTAN") Then

                    EsPellital = False

                ElseIf UCase(_EmpresaTrabajo).StartsWith("PELLITAL") Or UCase(_EmpresaTrabajo).StartsWith("PELITALL") Then

                    EsPellital = True

                End If

            End Set
        End Property


        Public Shared ReadOnly Property Empresas As List(Of String)
            Get
                ' DETERMINO LAS EMPRESAS CON LAS QUE TRABAJAR.

                Dim WEmpresa = "SURFACTAN"

                If EsPellital Then
                    WEmpresa = "PELLITAL"
                End If

                Select Case UCase(WEmpresa)

                    Case "SURFACTAN", "LOCAL"

                        Return New List(Of String) From {"SurfactanSA", "surfactan_II", "Surfactan_III", "Surfactan_IV", "Surfactan_V", "Surfactan_VI", "Surfactan_VII"}

                    Case "PELLITAL"
                        '
                        ' PELLITAL_II ESTA MAL ESCRITA PORQUE LA BASE DE DATOS LO TIENE ASI AL NOMBRE.
                        '
                        Return New List(Of String) From {"PellitalSA", "Pelitall_II", "Pellital_III", "Pellital_V"}

                    Case Else

                        Return New List(Of String) From {}

                End Select
            End Get
        End Property

        Public Shared Function _ConectarA(ByVal empresa As String) As String
            If String.IsNullOrEmpty(empresa) Then

                If EmpresaDeTrabajo = "" Then Throw New Exception("Empresa no definida para realizar la conexión a la Base de Datos.")

                empresa = EmpresaDeTrabajo

            End If

            If EsPellital AndAlso empresa = "SurfactanSA" Then
                empresa = "Pellital_III"
            End If

            Dim _empresa = Nothing

            _empresa = DeterminarBasePara(empresa)

            Return _ArmarCS(_empresa)

        End Function

        Private Shared Function _ArmarCS(ByVal WEmpresa As String)

            If Not String.IsNullOrEmpty(WEmpresa) Then

                Dim cs = Configuration.ConfigurationManager.ConnectionStrings("CS").ToString

                Return cs.Replace("#EMPRESA#", WEmpresa)

            Else
                Throw New Exception("No se pudo encontrar la empresa a la que se quiere conectar.")
            End If

        End Function

        Public Shared Function _ConectarASegunID(ByVal empresa As Integer) As String

            Dim _empresa = Nothing

            _empresa = DeterminarSegunIDIDBasePara(empresa)

            If Trim(_empresa) = "" Then Throw New Exception("No se pudo encontrar la empresa " & empresa)

            Return _ArmarCS(_empresa)

        End Function

        Private Shared Function DeterminarSegunIDIDBasePara(ByVal empresa As Integer) As String

            Select Case empresa
                Case 1
                    Return "SurfactanSa"
                Case 2
                    Return "PellitalSa"
                Case 3
                    Return "Surfactan_II"
                Case 4
                    Return "Pelitall_II" ' Asi está escrita la Base de Datos en el SQL.
                Case 5
                    Return "Surfactan_III"
                Case 6
                    Return "Surfactan_IV"
                Case 7
                    Return "Surfactan_V"
                Case 8
                    Return "Pellital_III"
                Case 9
                    Return "Pellital_V"
                Case 10
                    Return "Surfactan_VI"
                Case 11
                    Return "Surfactan_VII"
                Case Else
                    Return ""
            End Select

        End Function

        Public Shared Function DeterminarBasePara(ByVal empresa As String) As String

            Select Case UCase(empresa)
                Case "SURFACTAN QUIMICOS"

                    Return "Surfactan_II"

                Case "SURFACTAN FARMA"

                    Return "Surfactan_III"

                Case "PELLITAL"

                    Return "Pelitall_II" ' Se encuentra escrito asi el nombre en SQL.

                Case "SURFACTAN PIGMENTOS"

                    Return "SurfactanSA"
                Case Else

                    Return Empresas.Find(Function(e) UCase(e) = UCase(empresa))

            End Select
        End Function

        ''' <summary>
        ''' Chequea que la Clave exista en la Base de Datos que se está trabajando.
        ''' </summary>
        ''' <param name="Clave">Clave ingresada por el Usuario.</param>
        ''' <returns>Boolean</returns>
        ''' <remarks>True en caso de existir. Se almacena en la Propiedad Operador de la Clase Conexion, el número correspondiente para esa Clave.</remarks>
        Public Shared Function _Login(ByVal Clave As String, Optional ByVal WProceso As Integer = 0, Optional ByVal WEmpresa As String = "SurfactanSa") As Boolean

            Dim cn As SqlConnection = New SqlConnection()
            Dim cm As SqlCommand = New SqlCommand("SELECT Clave, Operador FROM Operador WHERE (Clave = '" & Clave & "' Or Clave = '" & UCase(Clave) & "')")
            Dim dr As SqlDataReader
            Dim WAtributos(10) As String

            ' Inicializamos los valores de Atributos del Operador.
            For i = 0 To 10
                For j = 0 To 100
                    WAtributosOperador(i, j) = "0"
                Next
            Next

            Try

                cn.ConnectionString = Helper._ConectarA(WEmpresa)
                cn.Open()
                cm.Connection = cn

                dr = cm.ExecuteReader()

                If dr.HasRows Then
                    dr.Read()
                    Operador = IIf(IsDBNull(dr.Item("Operador")), "", dr.Item("Operador"))
                    dr.Close()

                    ' Armamos los datos de los Atributos.
                    cm.CommandText = "SELECT Atributo1, Atributo2, Atributo3, Atributo4, Atributo5, Atributo6, Atributo7, Atributo8, Atributo9, Atributo10 " _
                                    & " FROM Atributos WHERE Operador = '" & Operador & "' AND Proceso = '" & WProceso & "'"

                    dr = cm.ExecuteReader()

                    If dr.HasRows Then

                        dr.Read()

                        For i = 1 To 10

                            WAtributos(i) = IIf(IsDBNull(dr.Item("Atributo" & i)), "", dr.Item("Atributo" & i))
                            WAtributos(i) = WAtributos(i).PadRight(33, "0")

                        Next

                        ' Parseamos los valores de atributos.
                        Dim Auxiliar = ""
                        For Ciclo = 1 To 10
                            Auxiliar = WAtributos(Ciclo)
                            For Ciclo1 = 1 To 30
                                WAtributosOperador(Ciclo, Ciclo1) = Val(Mid$(Auxiliar, Ciclo1, 1))
                            Next Ciclo1
                        Next Ciclo

                    End If

                    Return True
                Else
                    Operador = ""
                    Return False
                End If

            Catch ex As Exception
                Throw New Exception("Hubo un problema al querer consultar la Base de Datos." & vbCrLf & vbCrLf & "Motivo: " & ex.Message)
            Finally

                dr = Nothing
                cn.Close()
                cn = Nothing
                cm = Nothing

            End Try

        End Function
    End Class
End Namespace