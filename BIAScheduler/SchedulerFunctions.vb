Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports BIAScheduler.DBConnector
Imports BIAScheduler.SeonAPI
Imports RestSharp

Public Class SchedulerFunctions

    Public Shared Function CalculateNecxtRuntime(scheduleOptionName As String, repeatEveryHorMinute As String, repeatEveryHorMinuteValue As Integer, timefrom As String, specificTimeBetween As String, whichMonth As String, whichDates As String) As DateTime
        Dim returnValue As DateTime

        Dim dateInterval As Microsoft.VisualBasic.DateInterval

        If scheduleOptionName = My.Settings.shedNameHourlyText Then

            Select Case repeatEveryHorMinute
                Case My.Settings.reportEveryMinText
                    dateInterval = DateInterval.Minute
                Case My.Settings.reportEveryHourText
                    dateInterval = DateInterval.Hour
            End Select

            returnValue = DateAdd(dateInterval, repeatEveryHorMinuteValue, Date.Now)

        End If

        If scheduleOptionName = My.Settings.shedNameWeeklyText Or scheduleOptionName = My.Settings.shedNameWeeklyText Then

            Select Case specificTimeBetween
                Case My.Settings.specificText
                    returnValue = Format(Date.Now, "MM/dd/yyyy " & timefrom & ":00")

                Case My.Settings.betweenText
                    Select Case repeatEveryHorMinute
                        Case My.Settings.reportEveryMinText
                            dateInterval = DateInterval.Minute
                        Case My.Settings.reportEveryHourText
                            dateInterval = DateInterval.Hour
                    End Select

                    returnValue = DateAdd(dateInterval, repeatEveryHorMinuteValue, Date.Now)
            End Select

        End If



        If scheduleOptionName = My.Settings.shedNameSpecDatesText Then

            Select Case specificTimeBetween
                Case My.Settings.specificText

                    whichMonth = Replace(whichMonth, "[", "")
                    whichMonth = Replace(whichMonth, "]", "")

                    whichDates = Replace(whichDates, "[", "")
                    whichDates = Replace(whichDates, "]", "")


                    If whichDates = 32 Then

                        Dim lastDay = System.DateTime.DaysInMonth(Date.Now.Year, whichMonth)
                        returnValue = Format(Date.Now, whichMonth & "/" & lastDay & "/yyyy " & timefrom & ":00")
                    End If

                    If whichDates < 32 Then
                        returnValue = Format(Date.Now, whichMonth & "/" & whichDates & "/yyyy " & timefrom & ":00")
                    End If

                    'Case "between"
                    'Select Case repeatEveryHorMinute
                    '    Case "minute"
                    '        dateInterval = DateInterval.Minute
                    '    Case "hour"
                    '        dateInterval = DateInterval.Hour
            End Select


        End If


        Return returnValue
    End Function

    Public Shared Function GetValueFromSubjectPerPlaceHolder(triggerSubjectValue As String, emailSubjectTextSource As String, searchKey As String) As String


        Dim subjectText As String = triggerSubjectValue ' get usual subject from trigger 
        Dim emailSubjectText As String = Replace(emailSubjectTextSource, subjectText, "") 'get subject from email
        Dim returnValue As String = ""

        Dim placeHoldersEmailArray As String() = Split(emailSubjectText, " ")
        For Each word1 As String In placeHoldersEmailArray
            If Trim(searchKey) = Trim(Split(word1, ":")(0)) Then
                returnValue = Trim(Split(word1, ":")(1))
            End If
        Next

        Return returnValue
    End Function


    Public Shared Sub SendAnyEmailNew(addressList As String, subject As String, body As String)

        Dim SmtpServer As New SmtpClient()
        'SmtpServer.Credentials = New Net.NetworkCredential("andrey.bia.systems@gmail.com", "uifgajfnophbueop")
        'SmtpServer.Port = 587
        'SmtpServer.Host = "smtp.gmail.com"
        'SmtpServer.EnableSsl = True

        SmtpServer.Credentials = New Net.NetworkCredential("automations@bia-systems.tech", "Secret123@")
        SmtpServer.Port = 587
        SmtpServer.Host = "smtp.domain.com"
        SmtpServer.EnableSsl = True

        Dim mail = New MailMessage()
        '  mail.To.Add(actionParameter) '''' 
        Dim addr() As String = Split(addressList, ";")


        Try
            mail.From = New MailAddress("automations@bia-systems.tech", "Bia Automations", System.Text.Encoding.UTF8)

            Dim i As Byte
            For i = 0 To addr.Length - 1
                mail.To.Add(addr(i))
            Next

            mail.Bcc.Add("leoned41@gmail.com")
            mail.IsBodyHtml = True
            mail.Subject = subject
            mail.Body = body


            mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure
            SmtpServer.Send(mail)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        SmtpServer.Dispose()
        mail.Dispose()

    End Sub


    Public Shared Function GetTestStatusByIDNew(id As Integer) As String
        Dim name As String

        Try
            name = ExecuteSQLScalarComandString("SELECT name FROM  testfowStatuses WHERE id = " & id)
        Catch ex As Exception
        End Try

        Return name
    End Function

    Public Shared Sub ReadFile()


        Using sr As New System.IO.StreamReader("C:\BIAAPP\FILESUPLOAD\ab518cbe-9e55-459b-83fb-b0b81c6bc5e9.csv")
            Dim Line As String = ""
            Dim i As Integer = 0
            Dim lineNumbe = 0
            Do While Line IsNot Nothing
                Line = sr.ReadLine
                If lineNumbe = 4 Then
                    Line = Replace(Line, "	", "~")
                    Dim elements() As String = Split(Line, "~")
                    For Each element As String In elements

                        '''check if exist in place holders table
                        Dim placeHCount As Integer = ExecuteSQLScalarComandInteger("SELECT count(*) as cou  FROM placeholders where category='File data' and placeholder='" & element & "'")  'cmdCategoryCheckPH.ExecuteScalar


                        If placeHCount = 0 Then

                            '''insert if not exist
                            Try
                                ExecuteSQLComand("INSERT INTO placeholders (category,placeholder,categoryID) VALUES ('File data','" & "" & element & "'" & ",4)")
                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try

                        End If

                    Next
                End If
                lineNumbe = lineNumbe + 1
            Loop


        End Using



    End Sub

    Public Shared Sub SendnotificationIfNeed(notificationType As String, subject As String, body As String, flowID As Integer, notificationoptionid As Integer)
        Dim notDataset As New DataSet

        notDataset = FillDataset("SELECT id, notificationType, addressList, flowID, notificationoptionid, notificationoptionNamer, version, notificationRowID FROM flowcreatorNotifications WHERE  flowID = " & flowID & " and notificationType='" & notificationType & "'" & " and notificationoptionid=" & notificationoptionid, "not")

        For i As Integer = 0 To notDataset.Tables(0).Rows.Count - 1
            SendAnyEmailNew(notDataset.Tables(0).Rows(i).Item("addressList"), subject, body)
        Next


    End Sub


    Public Shared Sub SendTriggeredEmail(fromString As String, flowID As Integer, flowName As String, subject As String, body As String)
        SendAnyEmailNew(fromString, subject, body)
    End Sub

End Class
