Imports System.Net.Mail
Public Class Email
    Public Property sendto As String
    Public Property sender As String
    Public Property subject As String
    Public Property body As String
    Public Property cc As String
    Public Property bcc As String
    Public Property isBodyHtml As Boolean
    Public Property attachmentlist As List(Of String)
    Public Property htmlView As AlternateView = Nothing
    Public Property plainView As AlternateView = Nothing


    Dim mailsent As Boolean
    Public Sub New()

    End Sub

    Public Function send(ByRef message As String) As Boolean
        Dim myresult As Boolean = False
        Try
            Using mail As New MailMessage
                mail.ReplyToList.Add("sebasiahd@groupeseb.com")
                'mail.ReplyToList.Add("dwijanto@yahoo.com")
                Dim mysendtos() = sendto.Split(";")
                For Each mysendto In mysendtos
                    If mysendto.Length <> 0 Then
                        mail.To.Add(Trim(mysendto))
                    End If

                Next
                mail.From = New MailAddress(Trim(sender))

                If cc <> "" Then
                    Dim myccs() = cc.Split(";")

                    For Each mycc In myccs
                        If mycc.Length <> 0 Then
                            mail.CC.Add(Trim(mycc))
                        End If

                    Next
                End If
                If bcc <> "" Then
                    Dim mybccs() = bcc.Split(";")
                    For Each mybcc In mybccs
                        If mybcc.Length <> 0 Then
                            mail.Bcc.Add(Trim(mybcc))
                        End If
                    Next
                End If
                ' End If

                mail.Subject = subject
                mail.Body = body
                mail.IsBodyHtml = isBodyHtml

                If Not IsNothing(htmlView) Then
                    mail.AlternateViews.Add(htmlView)
                End If
                If Not IsNothing(plainView) Then
                    mail.AlternateViews.Add(plainView)
                End If

                If Not IsNothing(attachmentlist) Then
                    For Each mystr In attachmentlist
                        mail.Attachments.Add(New Attachment(mystr))
                    Next
                End If

                Using smtp = New SmtpClient(My.Settings.smtpclient)
                    'MessageBox.Show(My.Settings.smtpclient)
                    smtp.Send(mail)
                End Using
            End Using


            myresult = True
        Catch ex As Exception
            message = ex.Message
        Finally

        End Try
        Return myresult
    End Function

    Public Function sendAsync(ByRef message As String) As Boolean
        Dim myresult As Boolean = False
        Try
            Using smtp = New SmtpClient(My.Settings.smtpclient)
                AddHandler smtp.SendCompleted, AddressOf SendCompletedCallBack
                Using mail As New MailMessage
                    Dim mysendtos() = sendto.Split(";")
                    For Each mysendto In mysendtos
                        mail.To.Add(Trim(mysendto))
                    Next
                    mail.From = New MailAddress(Trim(sender))

                    If cc <> "" Then
                        Dim myccs() = cc.Split(";")
                        For Each mycc In myccs
                            mail.CC.Add(Trim(mycc))
                        Next
                    End If
                    If bcc <> "" Then
                        Dim mybccs() = bcc.Split(";")
                        For Each mybcc In mybccs
                            mail.Bcc.Add(Trim(mybcc))
                        Next
                    End If

                    mail.Subject = subject
                    mail.Body = body
                    mail.IsBodyHtml = isBodyHtml
                    If Not IsNothing(attachmentlist) Then
                        For Each mystr In attachmentlist
                            mail.Attachments.Add(New Attachment(mystr))
                        Next
                    End If
                    Dim userstate As String = "testMessage1"
                    smtp.SendAsync(mail, userstate)
                End Using
            End Using



            myresult = True
        Catch ex As Exception
            message = ex.Message
        Finally

        End Try
        Return myresult
    End Function

    Private Sub SendCompletedCallBack(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        Dim token As String = CStr(e.UserState)

        If e.Cancelled Then
            Debug.WriteLine("[{0}] Send canceled.", token)
        End If
        If e.Error IsNot Nothing Then
            Debug.WriteLine("[{0}] {1}", token, e.Error.ToString())
        Else
            Debug.WriteLine("Message sent.")
        End If
        mailSent = True
    End Sub
End Class
