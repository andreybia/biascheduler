Imports BIAScheduler.SchedulerFunctions
Imports BIAScheduler.DBConnector


Public Class Comunications


    Public Shared Sub SendProcessEmail(fromstring As String, flowID As Integer, flowName As String, subject As String, body As String)
        SendTriggeredEmail(fromstring, flowID, flowName, subject, body)
    End Sub

    Public Shared Sub PostActionProcessedActions(flowName As String, flowID As String, testID As Integer, actionParameter As String, version As Integer, triggerTableID As Integer, actionID As Integer, actionName As String)

        '''send message if need
        SendnotificationIfNeed("email", "Flow is completed", "Flow" & flowName & " is completed", flowID, 1)

        ''save to reports
        Try
            ExecuteSQLComand("INSERT INTO [dbo].[testFlowReports]( flowID, actionID, actionName, actionDateTime, testID, actionValue,version) VALUES (" & flowID & "," & actionID & ",'" & actionName & "','" & Date.Now & "'," & testID & ",'" & actionParameter & "'," & version & ")")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        ''show in report tab
        Try
            ExecuteSQLComand("UPDATE reports SET showFlag=1 where flowID=" & flowID)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


        '''change status in history
        Try
            ExecuteSQLComand("UPDATE testFlowHistory SET [testStatus] = ' " & GetTestStatusByIDNew(1) & " ', testStatusID=1, executionDateTime=getdate()  WHERE id=" & triggerTableID)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


        ''''update flow
        Try
            ExecuteSQLComand("UPDATE flows SET lastRun = '" & Format(Date.Now, "yyyy-MM-dd HH:mm:ss") & "',nextRun = getdate() WHERE flowID=" & flowID)
        Catch ex As Exception
            MsgBox(ex)
        End Try

    End Sub

End Class
