Imports System.Data.SqlClient


Public Class DBConnector

    Public Shared MAINCONNECTION = My.Settings.connection

    Public Shared Function ExecuteSQLScalarComandInteger(sqlCommand As String) As Integer
        Dim connection = New SqlConnection(MAINCONNECTION)
        connection.Open()
        Dim returnString As Integer = 0

        Try
            Dim cmd = New SqlCommand(sqlCommand, connection)
            returnString = cmd.ExecuteScalar
        Catch ex As Exception
        End Try


        connection.Close()
        Return returnString

    End Function

    Public Shared Function FillDataset(sqlString As String, datasetName As String) As DataSet
        Dim connection = New SqlConnection(DBConnector.MAINCONNECTION)
        connection.Open()
        Dim adapter = New SqlDataAdapter(sqlString, connection)
        Dim resultDS = New DataSet()
        adapter.Fill(resultDS, datasetName)
        connection.Close()
        Return resultDS
    End Function


    Public Shared Function FillDatasetStoreProcedure(procedureName As String, flowID As Integer, version As Integer) As DataSet

        Dim ds As New DataSet
        Dim cn As New SqlClient.SqlConnection(DBConnector.MAINCONNECTION)
        Dim Cmd As New SqlCommand(procedureName, cn)
        Cmd.CommandTimeout = 10000
        Cmd.CommandType = CommandType.StoredProcedure


        If flowID > 0 Then
            Cmd.Parameters.Add("@flowID", SqlDbType.Int)
            Cmd.Parameters("@flowID").Value = flowID
        End If


        If version > 0 Then
            Cmd.Parameters.Add("@version", SqlDbType.Int)
            Cmd.Parameters("@version").Value = flowID
        End If

        Dim sa As New SqlDataAdapter(Cmd)
        cn.Open()

        Try
            sa.Fill(ds)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return ds

        cn.Close()

    End Function


    Public Shared Function ExecuteSQLScalarComandIntegerProcedureWithParameters(procedureName As String, flowID As Integer, version As Integer) As Integer

        Dim returnValue As Integer = 0
        Dim cn As New SqlClient.SqlConnection(DBConnector.MAINCONNECTION)
        Dim Cmd As New SqlCommand(procedureName, cn)
        Cmd.CommandType = CommandType.StoredProcedure

        If flowID > 0 Then
            Cmd.Parameters.Add("@flowID", SqlDbType.Int)
            Cmd.Parameters("@flowID").Value = flowID
        End If


        If version > 0 Then
            Cmd.Parameters.Add("@version", SqlDbType.Int)
            Cmd.Parameters("@version").Value = version
        End If

        cn.Open()

        Try
            returnValue = Cmd.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        cn.Close()

        Return returnValue
    End Function


    Public Shared Function ExecuteSQLScalarComandString(sqlCommand As String) As String
        Dim connection = New SqlConnection(DBConnector.MAINCONNECTION)
        connection.Open()
        Dim returnString As String = ""

        Try
            Dim cmd = New SqlCommand(sqlCommand, connection)
            returnString = cmd.ExecuteScalar
        Catch ex As Exception
        End Try


        connection.Close()
        Return returnString
    End Function

    Public Shared Sub ExecuteSQLComand(sqlCommand As String)
        Dim connection = New SqlConnection(DBConnector.MAINCONNECTION)
        connection.Open()

        Try
            Dim cmd = New SqlCommand(sqlCommand, connection)
            Dim reuslt As Integer = cmd.ExecuteNonQuery
        Catch ex As Exception
        End Try


        connection.Close()
    End Sub


    Public Shared Sub SaveToEventLog(command As String, parameters As String, answerFromAPI As String)
        DBConnector.ExecuteSQLComand("INSERT INTO [dbo].[eventsLog] ([command],[parameters],[answerFromAPI]) VALUES ('" & command & "','" & parameters & "','" & answerFromAPI & "')")
    End Sub

    Public Shared Sub SaveConnCheck(answerText As String)
        DBConnector.ExecuteSQLComand("INSERT INTO [dbo].[applicationsConnectionCheck] (answer) VALUES ('" & answerText & "')")
    End Sub


    Public Shared Sub wait(ByVal seconds As Integer)
        For i As Integer = 0 To seconds * 100
            System.Threading.Thread.Sleep(10)
        Next
    End Sub

End Class
