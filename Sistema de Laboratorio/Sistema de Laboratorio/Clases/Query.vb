﻿Imports System.Data.SqlClient
Imports Laboratorio.Clases

Module Query

    Public Function GetSingle(ByVal q As String, Optional ByVal WBase As String = "") As DataRow

        If WBase.Trim = "" Then WBase = Operador.Base

        Dim tabla As New DataTable

        Using cn As New SqlConnection

            cn.ConnectionString = _ConectarA(WBase) 'ConfigurationManager.ConnectionStrings("CS").ToString
            cn.Open()

            Using cm As New SqlCommand(q)

                cm.Connection = cn

                Using dr As SqlDataReader = cm.ExecuteReader(CommandBehavior.SingleResult)
                    tabla.Load(dr)
                End Using

            End Using

        End Using

        If tabla.Rows.Count > 0 Then Return tabla.Rows(0)

        Return Nothing

    End Function

    Public Function GetAll(ByVal q As String, Optional ByVal WBase As String = "") As DataTable

        If WBase.Trim = "" Then WBase = Operador.Base

        Dim tabla As New DataTable

        Using cn As New SqlConnection

            cn.ConnectionString = _ConectarA(WBase) 'ConfigurationManager.ConnectionStrings("CS").ToString
            cn.Open()

            Using cm As New SqlCommand(q)

                cm.Connection = cn

                Using dr As SqlDataReader = cm.ExecuteReader
                    tabla.Load(dr)
                End Using

            End Using

        End Using

        Return tabla

    End Function

    Public Function ExecuteQueryRead(ByVal q As String, Optional ByVal WBase As String = "SurfactanSa") As SqlDataReader

        Dim cn As New SqlConnection
        Dim cm As New SqlCommand
        Dim dr As SqlDataReader

        cn.ConnectionString = _ConectarA(WBase) 'ConfigurationManager.ConnectionStrings("CS").ToString
        cn.Open()

        cm.Connection = cn
        cm.CommandText = q

        dr = cm.ExecuteReader

        Return dr

    End Function

    Public Sub ExecuteNonQueries(ByVal ParamArray q As String())

        Dim trans As SqlTransaction = Nothing
        Try
            If q.Length = 0 Then Throw New Exception("No se han pasado consultas para ejecutar.")

            Using cn As New SqlConnection
                cn.ConnectionString = _ConectarA(Operador.Base) 'ConfigurationManager.ConnectionStrings("CS").ToString
                cn.Open()
                trans = cn.BeginTransaction

                Using cm As New SqlCommand()

                    cm.Connection = cn
                    cm.Transaction = trans

                    For Each _q As String In q
                        Debug.Print(_q)
                        cm.CommandText = _q
                        cm.ExecuteNonQuery()
                    Next

                    trans.Commit()
                End Using
            End Using

        Catch ex As Exception
            If Not IsNothing(trans) AndAlso Not IsNothing(trans.Connection) Then trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub ExecuteNonQueries(ByVal q As String)
        ExecuteNonQueries({q})
    End Sub

    Public Sub ExecuteNonQueries(ByVal empresa As String, ByVal ParamArray q As String())

        Dim trans As SqlTransaction = Nothing
        Try
            If q.Length = 0 Then Throw New Exception("No se han pasado consultas para ejecutar.")

            Using cn As New SqlConnection
                cn.ConnectionString = _ConectarA(empresa) 'ConfigurationManager.ConnectionStrings("CS").ToString
                cn.Open()
                trans = cn.BeginTransaction

                Using cm As New SqlCommand()

                    cm.Connection = cn
                    cm.Transaction = trans

                    For Each _q As String In q
                        Debug.Print(_q)
                        cm.CommandText = _q
                        cm.ExecuteNonQuery()
                    Next

                    trans.Commit()
                End Using
            End Using

        Catch ex As Exception
            If Not IsNothing(trans) AndAlso Not IsNothing(trans.Connection) Then trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function CallProcedureWithReturns(ByVal proc As String, ByVal params As Dictionary(Of String, Object), Optional ByVal Base As String = "") As DataTable
        If Base.Trim = "" Then Base = Operador.Base
        Dim tabla As New DataTable
        Try

            Using cn As New SqlConnection
                cn.ConnectionString = _ConectarA(Base) 'ConfigurationManager.ConnectionStrings("CS").ToString
                cn.Open()

                Using cm As New SqlCommand()

                    cm.CommandType = CommandType.StoredProcedure
                    cm.CommandText = proc
                    cm.Connection = cn

                    For Each v As KeyValuePair(Of String, Object) In params
                        Dim p As New SqlParameter
                        p.DbType = SqlDbType.VarChar
                        p.Value = v.Value
                        p.ParameterName = v.Key
                        cm.Parameters.Add(p)
                    Next

                    Using dr As SqlDataReader = cm.ExecuteReader()
                        tabla.Load(dr)
                    End Using

                End Using
            End Using

            Return tabla

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

End Module
