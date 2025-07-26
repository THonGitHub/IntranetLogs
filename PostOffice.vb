Imports System.Configuration
Imports System.Net.Http
Imports System.Net.Mail
Imports System.Text
Imports Newtonsoft.Json

Public Class PostOffice

    ''' <summary>
    ''' send email message in HTML format
    ''' </summary>
    ''' <param name="recipients"></param>
    ''' <param name="CCrecipients"></param>
    ''' <param name="message"></param>
    ''' <param name="logFile"></param>
    ''' <param name="environment"></param>
    ''' <param name="subjectOptional"></param>
    ''' <returns></returns>
    Public Shared Function SendMessageToMailbox(
        recipients As String,
        CCrecipients As String,
        message As String,
        logFile As String,
        environment As String,
        Optional subjectOptional As String = Nothing) As Boolean

        Dim subject As String = $"{subjectOptional} Message from HVD SERVERS MONITORING application - {environment}."
        Dim Host As String = ConfigurationManager.AppSettings("host")
        Dim From As String = ConfigurationManager.AppSettings("from")
        Dim recipientAuthor As String = ConfigurationManager.AppSettings("recipientAuthor")
        Dim recipientBcc As String = ConfigurationManager.AppSettings("recipientBcc")
        Try
            Dim mMailMessage As New MailMessage With {
                .From = New MailAddress(From),
                .IsBodyHtml = True
            }
            mMailMessage.To.Add(New MailAddress(recipients))

            If CCrecipients.Length > 0 Then
                mMailMessage.CC.Add(New MailAddress(CCrecipients))
            End If

            'mMailMessage.Bcc.Add(New MailAddress(recipientBcc))
            mMailMessage.Subject = subject
            mMailMessage.Body = message
            Dim client As New SmtpClient With {
                .Host = Host,
                .Port = "587",
                .EnableSsl = True,
                .UseDefaultCredentials = True
            }
            client.Send(mMailMessage)

            Return True

        Catch ex As System.Exception
            ErrorHandler.HandleError(ex, environment, "SendMessageToMailbox", logFile)
            Return False
        End Try
    End Function


    ''' <summary>
    ''' method to send a message to MS TEAMS team channel using 'Incoming Webhook' app
    ''' </summary>
    ''' <param name="payload"></param>
    ''' <param name="logFile"></param>
    ''' <param name="environment"></param>
    Shared Async Sub SendMessageToMsTeams(
        payload As String,
        logFile As String,
        Optional environment As String = Nothing)

        ' TSI ASSET channel
        Dim webhookUrl As String = ConfigurationManager.AppSettings("webhookUrl")

        ' Modify the payload to include environment
        Try
            Dim messageObj = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(payload)
            If messageObj.ContainsKey("text") Then
                Dim currentText As String = messageObj("text")

                ' Find the first line break
                Dim firstLineEnd = currentText.IndexOf("<br>")
                If firstLineEnd > -1 Then
                    ' Insert environment before the line break
                    Dim firstLine = currentText.Substring(0, firstLineEnd)
                    Dim restOfText = currentText.Substring(firstLineEnd)
                    messageObj("text") = $"{firstLine} - {environment}{restOfText}"
                    payload = JsonConvert.SerializeObject(messageObj)
                End If
            End If
        Catch ex As Exception
            ErrorHandler.HandleError(ex, environment, "SendMessageToMsTeams - Payload Modification", logFile)
        End Try

        ' Create a new HttpClient instance
        Using client As New HttpClient()
            ' Set the content of the request
            Dim content As New StringContent(payload, Encoding.UTF8, "application/json")

            Try
                ' Send the POST request
                Dim response As HttpResponseMessage = Await client.PostAsync(webhookUrl, content)

                ' Check the response status
                If Not response.IsSuccessStatusCode Then
                    Console.WriteLine("Error sending message to MS TEAMS channel: " & response.ReasonPhrase & "; " & response.StatusCode)
                    FileOperations.WriteToFileNewLine(logFile, "Error sending message to MS TEAMS channel:", response.ReasonPhrase & " " & response.StatusCode, " ")
                End If
            Catch ex As Exception
                ErrorHandler.HandleError(ex, environment, "SendMessageToMsTeams - HTTP Request", logFile)
            End Try
        End Using
    End Sub


    Public Shared Sub SendNotification(
        recipient As String,
        message As String,
        environment As String,
        logFile As String,
        subject As String)

        Dim payloadObj = New With {.text = message}
        Dim payload = JsonConvert.SerializeObject(payloadObj)
        PostOffice.SendMessageToMsTeams(payload, logFile, environment)
        PostOffice.SendMessageToMailbox(recipient, "jptestdl@intl.att.com", message, logFile, environment, subject)
    End Sub

End Class