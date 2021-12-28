Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail
Imports System.Security.Cryptography
Imports System.Text
Imports BIAScheduler.DBConnector
Imports BIAScheduler.SchedulerFunctions
Imports Microsoft.VisualBasic.FileIO
Imports BIAScheduler.SeonAPI
Imports BIAScheduler.Comunications
Imports System.Net.Http
Imports System.Net.Http.Headers

Public Class Form1

    Sub CheckIfSomethingNeedToRun()

        Dim workDataset As New DataSet
        workDataset = FillDatasetStoreProcedure("GetIfNeedSomethingToAddToQueue", 0, 0)

        For i As Integer = 0 To workDataset.Tables(0).Rows.Count - 1

            Dim flowID As Double = workDataset.Tables(0).Rows(i).Item("flowID")
            Dim version As Integer = workDataset.Tables(0).Rows(i).Item("version")
            Dim id As Integer = workDataset.Tables(0).Rows(i).Item("id")
            Dim testID As Integer = workDataset.Tables(0).Rows(i).Item("testID")


            Dim reusltProductna As Integer = ExecuteSQLScalarComandIntegerProcedureWithParameters("CheckScheduledTasksPerFlowIDAndVersion", flowID, version)
            Dim schOrEventTrigger As Integer = 0 ' 0 - scheduler, 1 - event


            If reusltProductna = 0 Then

                '''''check if scheduled or event
                schOrEventTrigger = ExecuteSQLScalarComandIntegerProcedureWithParameters("CheckEventTriggerForFlow", flowID, 0)


                Select Case schOrEventTrigger

                    Case 0

                        '''get parameters
                        Dim parametersDataset As New DataSet
                        parametersDataset = FillDataset("SELECT id, flowID, scheduleOptionID, scheduleOptionName, repeatEveryHorMinute, specificTimeBetween, timeFrom, timeTo, whichDayOfWeek, whichMonth, whichDates, repeatEveryHorMinuteValue, version, whichDate, whichMonth, whichDates FROM flowcreatorchoosescheduleconf WHERE  flowID = " & flowID, "results")

                        '''''calculate next run time
                        Dim scheduleOptionName As String = parametersDataset.Tables(0).Rows(0).Item("scheduleOptionName")
                        Dim repeatEveryHorMinute As String = parametersDataset.Tables(0).Rows(0).Item("repeatEveryHorMinute")
                        Dim repeatEveryHorMinuteValue As Integer = parametersDataset.Tables(0).Rows(0).Item("repeatEveryHorMinuteValue")
                        Dim timeFrom As String = parametersDataset.Tables(0).Rows(0).Item("timeFrom")
                        Dim timeTo As String = parametersDataset.Tables(0).Rows(0).Item("timeTo")
                        Dim specificTimeBetween As String = parametersDataset.Tables(0).Rows(0).Item("specificTimeBetween")
                        Dim whichMonth As String = parametersDataset.Tables(0).Rows(0).Item("whichMonth")
                        Dim whichDates As String = parametersDataset.Tables(0).Rows(0).Item("whichDates")
                        Dim nextRunDateTime As DateTime = CalculateNecxtRuntime(scheduleOptionName, repeatEveryHorMinute, repeatEveryHorMinuteValue, timeFrom, specificTimeBetween, whichMonth, whichDates)

                        ''''insert new task
                        Try
                            ExecuteSQLComand("INSERT INTO schedulerQueue (flowID,version,scheduleOptionID,scheduleOptionName,repeatEveryHorMinute,repeatEveryHorMinuteValue,nextRunDateTime,specificTimeBetween,timeFrom,timeTo,triggerTableID,testID,whichMonth,whichDates) VALUES (" & flowID & "," & version & "," & parametersDataset.Tables(0).Rows(0).Item("scheduleOptionID") & ",'" & scheduleOptionName & "','" & repeatEveryHorMinute & "'," & repeatEveryHorMinuteValue & ",'" & Format(nextRunDateTime, "yyyy-MM-dd HH:mm:ss") & "','" & specificTimeBetween & "','" & timeFrom & "','" & timeTo & "'," & id & "," & testID & ",'" & whichMonth & "','" & whichDates & "')")
                        Catch ex As Exception
                            'MsgBox(ex.ToString)
                        End Try

                    Case 1

                        Try
                            ExecuteSQLComand("UPDATE flowcreatoreventbasedtriggers set active=1,triggerTableID=" & id & ",testID=" & testID & " where flowID=" & flowID)
                        Catch ex As Exception
                            MsgBox(ex.ToString)
                        End Try


                End Select


            End If


        Next


    End Sub


    Sub CheckIfSomethingStoppedScheduler()

        '''get all tasks from queue
        Dim workDataset As New DataSet
        workDataset = FillDataset("SELECT id, insertDateTime, flowID, version, scheduleOptionID, scheduleOptionName, repeatEveryHorMinute, repeatEveryHorMinuteValue, lastRunDateTime, nextRunDateTime, specificTimeBetween, timeFrom,timeTo,triggerTableID FROM schedulerQueue", "results")

        For i As Integer = 0 To workDataset.Tables(0).Rows.Count - 1

            Dim resultCheck As Integer = ExecuteSQLScalarComandInteger("SELECT  COUNT(*) As Expr1 from  testFlowHistory where  id=" & workDataset.Tables(0).Rows(i).Item("triggerTableID") & " and testStatusID in (2,1)")

            If resultCheck = 0 Then
                ExecuteSQLComand("delete from schedulerQueue where id=" & workDataset.Tables(0).Rows(i).Item("id"))
            End If

        Next


    End Sub

    Sub CheckIfSomethingStoppedEvent()

        '''get all tasks from queue
        Dim workDataset As New DataSet
        Dim sqlStringVersion As String = "SELECT id, insertDateTime, flowID, eventBasedTriggerID, eventBasedTriggerName, emilFromOption, emailFromOperatorID, emailFromOperatorText, emailFromValue, emailSubjectOperatorID, emailSubjectOperatorText, " &
                                         " emailSubjectValue, hasAttachment, encryptFileFlag, encryptFilePassword, attachmentFileID, version, applicationID, applicationName, active,triggerTableID FROM  flowcreatoreventbasedtriggers where active=1"
        workDataset = FillDataset(sqlStringVersion, "results")

        For i As Integer = 0 To workDataset.Tables(0).Rows.Count - 1

            Dim resultCheck As Integer = ExecuteSQLScalarComandInteger("SELECT  COUNT(*) As Expr1 from  testFlowHistory where  id=" & workDataset.Tables(0).Rows(i).Item("triggerTableID") & " and testStatusID in (2,1)") 'cmdCheck.ExecuteScalar

            If resultCheck = 0 Then
                ExecuteSQLComand("UPDATE flowcreatoreventbasedtriggers set active=0 where id=" & workDataset.Tables(0).Rows(i).Item("id"))
            End If

        Next

    End Sub


    Sub RunScheduledTasks()
        '   Dim connection = New SqlConnection("Data Source=LeoTest\SQLEXPRESS;Initial Catalog=robots;Persist Security Info=True;User ID=webSericeUser;Password=H96yN40Wp2jZZDaaVSnF")
        '   connection.Open()

        '''get all tasks from queue
        Dim workDataset As New DataSet
        workDataset = FillDataset("SELECT id, insertDateTime, flowID, version, scheduleOptionID, scheduleOptionName, repeatEveryHorMinute, repeatEveryHorMinuteValue, lastRunDateTime, nextRunDateTime, specificTimeBetween, timeFrom,timeTo,triggerTableID,testID,whichMonth,whichDates FROM schedulerQueue where nextRunDateTime<='" & Format(Date.Now, "yyyy-MM-dd HH:mm:ss") & "'", "results")

        For i As Integer = 0 To workDataset.Tables(0).Rows.Count - 1

            ''get actions from DB
            Dim actionsDataset As New DataSet
            actionsDataset = FillDataset("SELECT id, flowID, emailTo, subject, emailBody, encriptEmail, encriptPassword, applicationID, applicationName, actionID, actionName, actionRowID, version FROM flowcreatorActions WHERE flowID = " & workDataset.Tables(0).Rows(i).Item("flowID") & " order by id", "results")


            ''''ecute actions
            For ii As Integer = 0 To actionsDataset.Tables(0).Rows.Count - 1

                If actionsDataset.Tables(0).Rows(ii).Item("actionName") = My.Settings.actionNameSendEmailText Then
                    SendAnyEmailNew(actionsDataset.Tables(0).Rows(ii).Item("emailTo"), actionsDataset.Tables(0).Rows(ii).Item("subject"), actionsDataset.Tables(0).Rows(ii).Item("emailBody"))
                    Dim actionParameter = actionsDataset.Tables(0).Rows(ii).Item("emailTo")
                    Dim flowID As Integer = actionsDataset.Tables(0).Rows(i).Item("flowID")
                    Dim flowName As String = ExecuteSQLScalarComandString("select flowName from flows where flowID=" & actionsDataset.Tables(0).Rows(i).Item("flowID"))

                    '''send message if need
                    SendnotificationIfNeed("email", "Flow is completed", "Flow" & flowName & " is completed", flowID, 1)

                    ''save to reports
                    Try
                        ExecuteSQLComand("INSERT INTO [dbo].[testFlowReports]( flowID, actionID, actionName, actionDateTime, testID, actionValue,version) VALUES (" & actionsDataset.Tables(0).Rows(ii).Item("flowID") & ",1,'" & "Send Email" & "','" & Date.Now & "'," & workDataset.Tables(0).Rows(i).Item("testID") & ",'" & "Users email" & "'," & workDataset.Tables(0).Rows(i).Item("version") & ")")
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                    ''show in report tab
                    Try
                        ExecuteSQLComand("UPDATE [dbo].[reports] SET showFlag=1 where flowID=" & workDataset.Tables(0).Rows(i).Item("flowID"))
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                    '''change status in history
                    Try
                        ExecuteSQLComand("UPDATE testFlowHistory SET [testStatus] = ' " & GetTestStatusByIDNew(1) & " ', testStatusID=1, executionDateTime=getdate()  WHERE id=" & workDataset.Tables(0).Rows(i).Item("triggerTableID"))
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                    ''calculate next run time and save last runtime
                    Dim nextRunDateTime As DateTime = CalculateNecxtRuntime(workDataset.Tables(0).Rows(i).Item("scheduleOptionName"), workDataset.Tables(0).Rows(i).Item("repeatEveryHorMinute"), workDataset.Tables(0).Rows(i).Item("repeatEveryHorMinuteValue"), workDataset.Tables(0).Rows(i).Item("timeFrom"), workDataset.Tables(0).Rows(i).Item("specificTimeBetween"), workDataset.Tables(0).Rows(i).Item("whichMonth"), workDataset.Tables(0).Rows(i).Item("whichDates"))

                    ''''update queue
                    Try
                        ExecuteSQLComand("UPDATE schedulerQueue SET lastRunDateTime = '" & Format(Date.Now, "yyyy-MM-dd HH:mm:ss") & "',nextRunDateTime = '" & Format(nextRunDateTime, "yyyy-MM-dd HH:mm:ss") & "' WHERE id=" & workDataset.Tables(0).Rows(i).Item("id"))
                    Catch ex As Exception
                        MsgBox(ex)
                    End Try

                    ''''update flow
                    Try
                        ExecuteSQLComand("UPDATE flows SET lastRun = '" & Format(Date.Now, "yyyy-MM-dd HH:mm:ss") & "',nextRun = '" & Format(nextRunDateTime, "yyyy-MM-dd HH:mm:ss") & "' WHERE flowID=" & workDataset.Tables(0).Rows(i).Item("flowID"))
                    Catch ex As Exception
                        MsgBox(ex)
                    End Try


                End If

            Next

        Next

    End Sub


    Sub CheckedEventsTrigger()

        '''get all tasks from queue
        Dim triggerDataset As New DataSet
        Dim sqlStringVersion As String = "SELECT        id, insertDateTime, flowID, eventBasedTriggerID, eventBasedTriggerName, emilFromOption, emailFromOperatorID, emailFromOperatorText, emailFromValue, emailSubjectOperatorID, emailSubjectOperatorText, " &
                                         "emailSubjectValue,emailSubjectPlaceholder, hasAttachment, encryptFileFlag, encryptFilePassword, attachmentFileID, version, applicationID, applicationName, active, triggerTableID,testID FROM            flowcreatoreventbasedtriggers where active=1"
        triggerDataset = FillDataset(sqlStringVersion, "results")

        For i As Integer = 0 To triggerDataset.Tables(0).Rows.Count - 1



            '''check if some email triggered

            Dim fromSettings = triggerDataset.Tables(0).Rows(i).Item("emilFromOption")
            Dim fromOperator = triggerDataset.Tables(0).Rows(i).Item("emailFromOperatorText")
            Dim fromValue = triggerDataset.Tables(0).Rows(i).Item("emailFromValue")

            Dim emailSubjectOperatorText = triggerDataset.Tables(0).Rows(i).Item("emailSubjectOperatorText")
            Dim emailSubjectValue = triggerDataset.Tables(0).Rows(i).Item("emailSubjectValue")
            Dim hasAttachment As Boolean = triggerDataset.Tables(0).Rows(i).Item("hasAttachment")


            Dim emailSubjectPlaceholder = ""
            Try
                emailSubjectPlaceholder = triggerDataset.Tables(0).Rows(i).Item("emailSubjectPlaceholder")
            Catch ex As Exception
            End Try

            ''''''''remove symbols from place holders
            emailSubjectPlaceholder = Replace(emailSubjectPlaceholder, "][", "],[")
            emailSubjectPlaceholder = Replace(emailSubjectPlaceholder, "[", "")
            emailSubjectPlaceholder = Replace(emailSubjectPlaceholder, "]", "")
            emailSubjectPlaceholder = Replace(emailSubjectPlaceholder, ",", ";")


            ''''check emails queue
            Dim emailsDataset As New DataSet
            emailsDataset = FillDataset("SELECT id, insertDateTime, emailFromString, subjectString, triggeredDateTime,fileName FROM eventTriggerQueue where triggeredDateTime is null", "results")

            For y As Integer = 0 To emailsDataset.Tables(0).Rows.Count - 1

                Dim fromTriggered As Boolean = False
                Dim subjectTriggered As Boolean = False
                Dim attachmentTriggered As Boolean = True
                Dim fromString = emailsDataset.Tables(0).Rows(y).Item("emailFromString")
                Dim subjectString = emailsDataset.Tables(0).Rows(y).Item("subjectString")
                Dim emailFromStringEmail = emailsDataset.Tables(0).Rows(y).Item("emailFromString")
                Dim fileName = emailsDataset.Tables(0).Rows(y).Item("fileName")

                subjectString = Replace(subjectString, "[", "")
                subjectString = Replace(subjectString, "]", "")

                ''''check from 
                Select Case fromSettings
                    Case "Any"
                        fromTriggered = True
                    Case Else

                        If fromOperator = "is" Then
                            If fromString = fromValue Then
                                fromTriggered = True
                            End If
                        End If

                        If fromOperator = "is not" Then
                            If fromString <> fromValue Then
                                fromTriggered = True
                            End If
                        End If

                        If fromOperator = "contains any of" Then
                            If InStr(fromString, fromValue) > 0 Then
                                fromTriggered = True
                            End If
                        End If

                End Select



                ''''check subject
                Select Case emailSubjectOperatorText
                    Case "Any"
                        subjectTriggered = True
                    Case Else

                        If emailSubjectOperatorText = "is" Then



                            If emailSubjectPlaceholder = "" Then
                                If subjectString = emailSubjectValue Then
                                    subjectTriggered = True
                                End If
                            End If

                            If emailSubjectPlaceholder <> "" Then

                                If InStr(subjectString, emailSubjectValue) > 0 Then
                                    Dim placeHoldersFromSettings As String() = Split(emailSubjectPlaceholder, ";")
                                    Dim clearSubjectPlaceHolders = Trim(Replace(subjectString, emailSubjectValue, ""))
                                    Dim plaveholdersClearSubjectPlaceHolders As String() = Split(clearSubjectPlaceHolders, ";")



                                    subjectTriggered = True
                                    For j As Integer = 0 To placeHoldersFromSettings.Count - 1
                                        Dim settingPlHolder As String = placeHoldersFromSettings(j)
                                        If settingPlHolder <> Trim(Split(plaveholdersClearSubjectPlaceHolders(j), ":")(0)) Then
                                            subjectTriggered = False
                                        End If


                                    Next


                                    ''''''''check values
                                    Dim errorText As String = "Hello,<br><br>Placeholder value is missing. Please resend the request including the missing information.<br><br>﻿﻿﻿This email was sent from an email address that can't receive emails. Please don't reply to this email.<br><br><br>﻿Regards,<br>﻿Bia team"
                                    For j As Integer = 0 To placeHoldersFromSettings.Count - 1

                                        If InStr(plaveholdersClearSubjectPlaceHolders(j), ":") = 0 Then
                                            SendAnyEmailNew(fromString, "Error in subject", errorText)
                                            subjectTriggered = False
                                            ExecuteSQLComand("UPDATE eventTriggerQueue SET triggeredDateTime=getdate() WHERE id=" & emailsDataset.Tables(0).Rows(y).Item("id"))
                                        End If

                                        If InStr(plaveholdersClearSubjectPlaceHolders(j), ":") > 0 Then
                                            If Trim(Split(plaveholdersClearSubjectPlaceHolders(j), ":")(1)) = "" Then
                                                SendAnyEmailNew(fromString, "Error in subject", errorText)
                                                subjectTriggered = False
                                                ExecuteSQLComand("UPDATE eventTriggerQueue SET triggeredDateTime=getdate() WHERE id=" & emailsDataset.Tables(0).Rows(y).Item("id"))
                                            End If
                                        End If

                                    Next


                                End If
                            End If


                        End If

                        If emailSubjectOperatorText = "starts with" Or emailSubjectOperatorText = "contains" Then
                            If InStr(subjectString, emailSubjectValue) > 0 Then
                                subjectTriggered = True
                            End If

                        End If

                        If emailSubjectOperatorText = "is not" Then
                            If subjectString <> emailSubjectValue Then
                                subjectTriggered = True
                            End If
                        End If

                End Select




                If fromTriggered = True And subjectTriggered = True Then

                    Dim flowName As String = ExecuteSQLScalarComandString("select flowName from flows where flowID=" & triggerDataset.Tables(0).Rows(i).Item("flowID"))
                    Dim subject As String = ""
                    Dim body As String = ""


                    '''check attachment
                    If hasAttachment = True Then

                        ''''check if atatched file
                        If fileName = "" Then
                            Dim subjectError As String = "ERROR-flow " & flowName & ". CSV file is missing"
                            Dim bodyError As String = "Hello.<br>CSV file is missing. Please resend the email with a csv file as configured in the flow.<br>Correct file format can be viewed in the trigger step of the flow.<br><br>Best Regards,<br>Bia team"
                            attachmentTriggered = False
                            ExecuteSQLComand("UPDATE eventTriggerQueue SET triggeredDateTime=getdate(), errorExist=1, errorReason='No attachment',errorEmailSent=getdate() WHERE id=" & emailsDataset.Tables(0).Rows(y).Item("id"))
                            SendnotificationIfNeed("email", subjectError, bodyError, triggerDataset.Tables(0).Rows(i).Item("flowID"), 3)
                            SendTriggeredEmail(fromString, triggerDataset.Tables(0).Rows(i).Item("flowID"), flowName, subjectError, bodyError)
                        End If


                        ''if extension is correct
                        If fileName <> "" And attachmentTriggered = True Then
                            Dim file As String = Replace(fileName, "C:\EmailAttachments\", "")
                            file = Replace(file, vbCrLf, "")
                            If Path.GetExtension(file).ToString <> ".csv" Then
                                Dim subjectError As String = "ERROR-flow " & flowName & ". Only csv files supported"
                                Dim bodyError As String = "Hello.<br>Only csv files supported.  Please resend the email with a csv file.<br><br>Best Regards,<br>Bia team"
                                attachmentTriggered = False
                                ExecuteSQLComand("UPDATE eventTriggerQueue SET triggeredDateTime=getdate(), errorExist=1, errorReason='Not CSV',errorEmailSent=getdate() WHERE id=" & emailsDataset.Tables(0).Rows(y).Item("id"))
                                SendnotificationIfNeed("email", subjectError, bodyError, triggerDataset.Tables(0).Rows(i).Item("flowID"), 3)
                                SendTriggeredEmail(fromString, triggerDataset.Tables(0).Rows(i).Item("flowID"), flowName, subjectError, bodyError)
                            End If
                        End If


                        '''check columns count in file
                        If fileName <> "" And attachmentTriggered = True Then

                            Dim placeholdersIDString As String = ExecuteSQLScalarComandString("select placeHoldersID from flowcreatorShowPlaceHolders where flowID=" & triggerDataset.Tables(0).Rows(i).Item("flowID"))
                            Dim flowstatesDS As New DataSet
                            flowstatesDS = FillDataset("Select PlaceHolder FROM  placeholders where id in (" & placeholdersIDString & ")  ", "results")
                            Dim tfp As New TextFieldParser(fileName.ToString)
                            tfp.Delimiters = New String() {","}
                            tfp.TextFieldType = FieldType.Delimited


                            For ii As Integer = 0 To 0 'check only first row

                                ''''check columns count
                                Dim fields = tfp.ReadFields()
                                If fields.Count <> flowstatesDS.Tables(0).Rows.Count Then
                                    Dim subjectError As String = "ERROR-flow " & flowName & ". Column not correct"
                                    Dim bodyError As String = "Hello.<br>Column names in the received  file do not correspond to the example file format configured in the flow.<br>Please resend the email with a csv file as configured in the flow.<br>Correct file format can be viewed in the trigger step of the flow.<br><br>Best Regards,<br>Bia team"
                                    attachmentTriggered = False
                                    ExecuteSQLComand("UPDATE eventTriggerQueue SET triggeredDateTime=getdate(), errorExist=1, errorReason='Columns count',errorEmailSent=getdate() WHERE id=" & emailsDataset.Tables(0).Rows(y).Item("id"))
                                    SendnotificationIfNeed("email", subjectError, bodyError, triggerDataset.Tables(0).Rows(i).Item("flowID"), 3)
                                    SendTriggeredEmail(fromString, triggerDataset.Tables(0).Rows(i).Item("flowID"), flowName, subjectError, bodyError)
                                End If


                                ''''check if columns in file exist in dataset
                                If attachmentTriggered = True Then

                                    For Each fd As String In fields
                                        Dim row As DataRow = flowstatesDS.Tables(0).Select("PlaceHolder = '" + fd.Trim() + "'").FirstOrDefault()
                                        If row Is Nothing Then
                                            Dim subjectError As String = "ERROR-flow " & flowName & ". Column not correct"
                                            Dim bodyError As String = "Hello.<br>Column names in the received  file do not correspond to the example file format configured in the flow.Please resend the email with a csv file as configured in the flow.Correct file format can be viewed in the trigger step of the flow.<br><br>Best Regards,<br>Bia team"
                                            attachmentTriggered = False
                                            ExecuteSQLComand("UPDATE eventTriggerQueue SET triggeredDateTime=getdate(), errorExist=1, errorReason='Columns names',errorEmailSent=getdate() WHERE id=" & emailsDataset.Tables(0).Rows(y).Item("id"))
                                            SendnotificationIfNeed("email", subjectError, bodyError, triggerDataset.Tables(0).Rows(i).Item("flowID"), 3)
                                            SendTriggeredEmail(fromString, triggerDataset.Tables(0).Rows(i).Item("flowID"), flowName, subjectError, bodyError)
                                            Exit For
                                        End If
                                    Next
                                End If



                                ''''check if columns in file exist in dataset
                                If attachmentTriggered = True Then

                                    Dim counter = 0
                                    For Each fd As String In fields
                                        If fd.Trim().ToLower() <> flowstatesDS.Tables(0).Rows(counter).Item("PlaceHolder").ToString().ToLower() Then
                                            Dim subjectError As String = "ERROR-flow " & flowName & ". Column order not correct"
                                            Dim bodyError As String = "Hello.<br>Column order in the received  file does not correspond to the example file format configured in the flow. Please resend the email with a csv file as configured in the flow. Correct file format can be viewed in the trigger step of the flow.<br><br>Best Regards,<br>Bia team"
                                            attachmentTriggered = False
                                            ExecuteSQLComand("UPDATE eventTriggerQueue SET triggeredDateTime=getdate(), errorExist=1, errorReason='Columns order',errorEmailSent=getdate() WHERE id=" & emailsDataset.Tables(0).Rows(y).Item("id"))
                                            SendnotificationIfNeed("email", subjectError, bodyError, triggerDataset.Tables(0).Rows(i).Item("flowID"), 3)
                                            SendTriggeredEmail(fromString, triggerDataset.Tables(0).Rows(i).Item("flowID"), flowName, subjectError, bodyError)
                                            Exit For
                                        End If
                                        counter = counter + 1
                                    Next
                                End If

                            Next ''rows

                        End If

                    End If



                    If attachmentTriggered = True Then

                        ''''send email about email process
                        Dim flowIDForEmail As Integer = triggerDataset.Tables(0).Rows(i).Item("flowID")
                        flowName = ExecuteSQLScalarComandString("select flowName from flows where flowID=" & flowIDForEmail)
                        subject = flowName & " Flow has been triggered"
                        body = "Hello.<br><br>This is a confirmation message informing you that you have successfully triggered " & flowName & " Flow.<br><br>﻿﻿﻿This email was sent from an email address that can't receive emails. Please don't reply to this email.<br><br>Best Regards,<br>Bia team"
                        SendProcessEmail(fromString, flowIDForEmail, flowName, subject, body)

                        '''update emails queue
                        ExecuteSQLComand("UPDATE eventTriggerQueue SET triggeredDateTime=getdate() WHERE id=" & emailsDataset.Tables(0).Rows(y).Item("id"))

                        ''get actions from DB
                        Dim actionsDataset As New DataSet
                        actionsDataset = FillDataset("SELECT  id, flowID, emailTo, subject, emailBody, encriptEmail, encriptPassword, applicationID, applicationName, actionID, actionName, actionRowID, version, actionPlaceholder, emailToPlaceholder, subjectPlaceholder, emailBodyPlaceholder FROM flowcreatorActions WHERE flowID = " & triggerDataset.Tables(0).Rows(i).Item("flowID") & " order by id", "results")

                        ''''execute actions
                        For ii As Integer = 0 To actionsDataset.Tables(0).Rows.Count - 1


                            Select Case actionsDataset.Tables(0).Rows(ii).Item("actionID")

                                Case 1 '"Send Email"


                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''check if some place holders in subject 
                                    Dim subjectForSending As String = actionsDataset.Tables(0).Rows(ii).Item("subject")

                                    ''''''check if sent place holders with email
                                    If actionsDataset.Tables(0).Rows(ii).Item("subjectPlaceholder") <> "" Then
                                        Dim allPlaceHoldersFromSettings As String = Replace(Replace(actionsDataset.Tables(0).Rows(ii).Item("subjectPlaceholder"), "[", ""), "]", "") ''''get place holders from settings
                                        Dim subjectText As String = triggerDataset.Tables(0).Rows(i).Item("emailSubjectValue") ' get usual subject from trigger 
                                        Dim emailSubjectText As String = Replace(emailsDataset.Tables(0).Rows(y).Item("subjectString"), subjectText, "") 'get subject from email

                                        Dim placeHoldersFromSettingsArray As String() = Split(allPlaceHoldersFromSettings, ";")
                                        For Each word As String In placeHoldersFromSettingsArray
                                            Dim placeHoldersEmailArray As String() = Split(emailSubjectText, " ")
                                            For Each word1 As String In placeHoldersEmailArray
                                                If Trim(word) = Trim(Split(word1, ":")(0)) Then
                                                    subjectForSending = subjectForSending & " " & Trim(Split(word1, ":")(1))
                                                End If
                                            Next
                                        Next
                                    End If

                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''check if some place holders in body 
                                    Dim bodyForSending As String = actionsDataset.Tables(0).Rows(ii).Item("emailBody")

                                    If actionsDataset.Tables(0).Rows(ii).Item("emailBodyPlaceholder") <> "" Then
                                        Dim allPlaceHoldersFromSettings As String = Replace(Replace(actionsDataset.Tables(0).Rows(ii).Item("emailBodyPlaceholder"), "[", ""), "]", "") ''''get place holders from settings
                                        Dim subjectText As String = triggerDataset.Tables(0).Rows(i).Item("emailSubjectValue") ' get usual subject from trigger 
                                        Dim emailSubjectText As String = Replace(emailsDataset.Tables(0).Rows(y).Item("subjectString"), subjectText, "") 'get subject from email

                                        Dim placeHoldersFromSettingsArray As String() = Split(allPlaceHoldersFromSettings, ";")
                                        For Each word As String In placeHoldersFromSettingsArray
                                            Dim placeHoldersEmailArray As String() = Split(emailSubjectText, " ")
                                            For Each word1 As String In placeHoldersEmailArray
                                                If Trim(word) = Trim(Split(word1, ":")(0)) Then
                                                    bodyForSending = bodyForSending & " " & Trim(Split(word1, ":")(1))
                                                End If
                                            Next
                                        Next
                                    End If


                                    '''''execute action
                                    Dim actionParameter = actionsDataset.Tables(0).Rows(ii).Item("emailTo")
                                    SendAnyEmailNew(actionParameter, subjectForSending, bodyForSending)
                                    PostActionProcessedActions(flowName, triggerDataset.Tables(0).Rows(i).Item("flowID"), triggerDataset.Tables(0).Rows(i).Item("testID"), "Users email", triggerDataset.Tables(0).Rows(i).Item("version"), triggerDataset.Tables(0).Rows(i).Item("triggerTableID"), actionsDataset.Tables(0).Rows(ii).Item("actionID"), actionsDataset.Tables(0).Rows(ii).Item("actionName"))




                                Case 88 '"Run Email Check"

                                    Dim subjectText As String = triggerDataset.Tables(0).Rows(i).Item("emailSubjectValue")
                                    Dim emilForSeon = GetValueFromSubjectPerPlaceHolder(triggerDataset.Tables(0).Rows(i).Item("emailSubjectValue"), Replace(emailsDataset.Tables(0).Rows(y).Item("subjectString"), subjectText, ""), "userEmail")
                                    Dim idForSeon = GetValueFromSubjectPerPlaceHolder(triggerDataset.Tables(0).Rows(i).Item("emailSubjectValue"), Replace(emailsDataset.Tables(0).Rows(y).Item("subjectString"), subjectText, ""), "userID")
                                    Dim answer = SendSeonFraudAPIRequest(2, idForSeon, "", emilForSeon, "")
                                    flowName = ExecuteSQLScalarComandString("select flowName from flows where flowID=" & triggerDataset.Tables(0).Rows(i).Item("flowID"))
                                    subject = flowName & " Answer fromSeon"

                                    Dim actionParameter = "Request [" & idForSeon & "],[" & emilForSeon & "]"
                                    SendAnyEmailNew(emailFromStringEmail, subject, Mid(answer, 1, 100))
                                    PostActionProcessedActions(flowName, triggerDataset.Tables(0).Rows(i).Item("flowID"), triggerDataset.Tables(0).Rows(i).Item("testID"), "userID=" & idForSeon, triggerDataset.Tables(0).Rows(i).Item("version"), triggerDataset.Tables(0).Rows(i).Item("triggerTableID"), actionsDataset.Tables(0).Rows(ii).Item("actionID"), actionsDataset.Tables(0).Rows(ii).Item("actionName"))


                                Case 90 '"Run Phone Check"

                                    Dim subjectText As String = triggerDataset.Tables(0).Rows(i).Item("emailSubjectValue")
                                    Dim PhoneForSeon = GetValueFromSubjectPerPlaceHolder(triggerDataset.Tables(0).Rows(i).Item("emailSubjectValue"), Replace(emailsDataset.Tables(0).Rows(y).Item("subjectString"), subjectText, ""), "userPhone")
                                    Dim idForSeon = GetValueFromSubjectPerPlaceHolder(triggerDataset.Tables(0).Rows(i).Item("emailSubjectValue"), Replace(emailsDataset.Tables(0).Rows(y).Item("subjectString"), subjectText, ""), "userID")
                                    Dim answer = SendSeonFraudAPIRequest(2, idForSeon, PhoneForSeon, "", "")
                                    flowName = ExecuteSQLScalarComandString("select flowName from flows where flowID=" & triggerDataset.Tables(0).Rows(i).Item("flowID"))
                                    subject = flowName & " Answer fromSeon"

                                    Dim actionParameter = "Request [" & idForSeon & "],[" & PhoneForSeon & "]"
                                    SendAnyEmailNew(emailFromStringEmail, subject, Mid(answer, 1, 100))
                                    PostActionProcessedActions(flowName, triggerDataset.Tables(0).Rows(i).Item("flowID"), triggerDataset.Tables(0).Rows(i).Item("testID"), "UserID=" & idForSeon, triggerDataset.Tables(0).Rows(i).Item("version"), triggerDataset.Tables(0).Rows(i).Item("triggerTableID"), actionsDataset.Tables(0).Rows(ii).Item("actionID"), actionsDataset.Tables(0).Rows(ii).Item("actionName"))

                                Case 91 '"Run Fraud Check"

                                    Dim subjectText As String = triggerDataset.Tables(0).Rows(i).Item("emailSubjectValue")
                                    Dim idForSeon = GetValueFromSubjectPerPlaceHolder(triggerDataset.Tables(0).Rows(i).Item("emailSubjectValue"), Replace(emailsDataset.Tables(0).Rows(y).Item("subjectString"), subjectText, ""), "userID")
                                    Dim emailForSeon = GetValueFromSubjectPerPlaceHolder(triggerDataset.Tables(0).Rows(i).Item("emailSubjectValue"), Replace(emailsDataset.Tables(0).Rows(y).Item("subjectString"), subjectText, ""), "userEmail")
                                    Dim PhoneForSeon = GetValueFromSubjectPerPlaceHolder(triggerDataset.Tables(0).Rows(i).Item("emailSubjectValue"), Replace(emailsDataset.Tables(0).Rows(y).Item("subjectString"), subjectText, ""), "userPhone")
                                    Dim ipForSeon = GetValueFromSubjectPerPlaceHolder(triggerDataset.Tables(0).Rows(i).Item("emailSubjectValue"), Replace(emailsDataset.Tables(0).Rows(y).Item("subjectString"), subjectText, ""), "userIP")


                                    Dim answer = SendSeonFraudAPIRequest(2, idForSeon, ipForSeon, emailForSeon, PhoneForSeon)
                                    flowName = ExecuteSQLScalarComandString("select flowName from flows where flowID=" & triggerDataset.Tables(0).Rows(i).Item("flowID"))
                                    subject = flowName & " Answer fromSeon"

                                    Dim actionParameter = "Request [" & idForSeon & "],[" & ipForSeon & "],[" & PhoneForSeon & "],[" & emailForSeon & "]"
                                    SendAnyEmailNew(emailFromStringEmail, subject, Mid(answer, 1, 100))
                                    PostActionProcessedActions(flowName, triggerDataset.Tables(0).Rows(i).Item("flowID"), triggerDataset.Tables(0).Rows(i).Item("testID"), "UserID=" & idForSeon, triggerDataset.Tables(0).Rows(i).Item("version"), triggerDataset.Tables(0).Rows(i).Item("triggerTableID"), actionsDataset.Tables(0).Rows(ii).Item("actionID"), actionsDataset.Tables(0).Rows(ii).Item("actionName"))




                            End Select




                        Next

                    End If

                End If

            Next

        Next

    End Sub

    Sub CheckInboxGetAllEmailsDecideWhatToDo()
        Dim glob As New Chilkat.Global
        Dim success As Boolean = glob.UnlockBundle(My.Settings.unb)
        If (success <> True) Then
            Debug.WriteLine(glob.LastErrorText)
            Exit Sub
        End If
        Dim status As Integer = glob.UnlockStatus
        If (status = 2) Then
            Debug.WriteLine("Unlocked using purchased unlock code.")
        Else
            Debug.WriteLine("Unlocked in trial mode.")
        End If
        Debug.WriteLine(glob.LastErrorText)

    End Sub




    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ''''''''''  CheckInboxGetAllEmailsDecideWhatToDo()

        ''run Emails Checker
        Shell("C:\auto\GetEmails.exe", AppWinStyle.MinimizedNoFocus, True)

        CheckIfSomethingNeedToRun()
        CheckIfSomethingStoppedScheduler()
        CheckIfSomethingStoppedEvent()
        RunScheduledTasks()
        CheckedEventsTrigger()
        CloseOpenMoreThanFourHoursFlows()




        Close()
    End Sub


    Sub CloseOpenMoreThanFourHoursFlows()

        Try
            ExecuteSQLComand("delete from flowcreatorFlowEdit where insertDateTime<='" & DateAdd(DateInterval.Hour, -9, Date.Now) & "'")
        Catch ex As Exception
        End Try

    End Sub


End Class
