Imports System.Data
Imports System.Data.Common

Public Class DatabaseConnection
    Implements IDisposable

    Private _connectionString As String
    Private _providerName As String
    Private _factory As DbProviderFactory
    Private _connection As DbConnection

    Public Sub New(connectionString As String, providerName As String)
        _connectionString = connectionString
        _providerName = providerName
        _factory = DbProviderFactories.GetFactory(_providerName)
        _connection = _factory.CreateConnection()
        _connection.ConnectionString = _connectionString
    End Sub

    Public Sub OpenConnection()
        If _connection.State <> ConnectionState.Open Then
            _connection.Open()
        End If
    End Sub

    Public Sub CloseConnection()
        If _connection.State <> ConnectionState.Closed Then
            _connection.Close()
        End If
    End Sub

    Public Function ExecuteQuery(query As String) As DataTable
        Dim command As DbCommand = _connection.CreateCommand()
        command.CommandText = query

        Dim adapter As DbDataAdapter = _factory.CreateDataAdapter()
        adapter.SelectCommand = command

        Dim result As New DataTable()
        adapter.Fill(result)

        Return result
    End Function

    Public Sub ExecuteNonQuery(query As String)
        Dim command As DbCommand = _connection.CreateCommand()
        command.CommandText = query
        command.ExecuteNonQuery()
    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        If disposing Then
            If _connection IsNot Nothing Then
                _connection.Dispose()
                _connection = Nothing
            End If
        End If
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
End Class
