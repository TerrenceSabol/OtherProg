Imports System.Data.Odbc
Imports outlook = Microsoft.Office.Interop.Outlook
Imports System.Web
Imports System.Net.Mail

Public Class Form1
    Dim timeover As String = "N"
    Dim emailin As DateTime
    Dim emailout As DateTime

    Dim strstudyseconds As String
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label3.Text = " "
        Label5.Text = " "
        Label7.Visible = False
        Label8.Visible = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        timeover = "N"
        If TextBox1.Text.Length > 6 Then
            TextBox1.Text = TextBox1.Text.Substring(0, 6)
        End If
        If TextBox1.Text.Length = 0 Then
            MsgBox("The ID Number is missing")
            Exit Sub
        End If
        Dim theid As String
        theid = TextBox1.Text
        Dim theseq As Integer
        Dim lastType As String = "I"
        theseq = GetSeq(TextBox1.Text)
        If theseq = -1 Then
            timeover = "N"
        Else
            timeover = " "
            lastType = GetlastType(TextBox1.Text, theseq)
        End If
        If lastType <> "O" And timeover <> "N" And theseq <> 0 And theseq <> -1 Then
            timeover = "Y"
        End If
        If timeover = "Y" Then
            Label3.Text = "Type is Incorrect--Button Disabled "
            Button1.Enabled = False
            Exit Sub
        End If
        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select   Coalesce(first_name + ' '  + middle_name + '  ' +  Last_name, first_name + ' '  + Last_name, last_name)  from name_master where id_num='" & theid & "'"


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)


            Button1.Enabled = False
            Dim thecount As Integer = 0
            Dim thetime As DateTime
            Dim myReader As OdbcDataReader
            Dim okr As Boolean
            Dim thetype As String = "I"

            myReader = catCMD.ExecuteReader
            While myReader.Read()
                thecount = thecount + 1
                Label3.Text = myReader(0)

                thetime = Now()
                Label5.Text = thetime
                thetype = "I"
                okr = WritetheRecord(thetype, TextBox1.Text, thetime)
                If okr = True Then
                    ShowlastFive()
                End If
            End While
            If thecount = 0 Then
                Label3.Text = "Name Not Found"
                Button1.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(" You can retry the entry or manually record the information. If the retry fails, please contact OIT and provide them with this message." & ex.Message)
        End Try
        TextBox1.Clear()
    End Sub

    Private Sub TextBox1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Enter
        Button1.Enabled = True
        Button2.Enabled = True
        Label7.Visible = False
        Label8.Visible = False
    End Sub
    Function WritetheRecord(ByVal thetype As String, ByVal theid As String, ByVal thetimestamp As DateTime) As Boolean
        Dim wtype As String
        Dim wid As String
        Dim wtime As DateTime
        Dim theseq As Integer
        Dim flagconstant As String = " "
        theseq = GetSeq(theid)
        If theseq = -1 Then
            theseq = 10
        Else
            theseq = theseq + 10
        End If

        wtype = thetype
        wid = theid
        wtime = thetimestamp
        Dim writesql As String
        writesql = "Insert into dbo.NC_StudyHall (ID_NUM,SEQNO,TYPEINFO,ENTRYTIME,FLAGENTRY) values('" & wid & "'," & theseq & ",'" & wtype & "' ,'" & wtime & "','" & flagconstant & "')"
        Try

            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()







            Dim catCMD As OdbcCommand = New OdbcCommand(writesql, oODBCConnection)



            Dim okcount As Integer = 0

            okcount = catCMD.ExecuteNonQuery
            oODBCConnection.Close()


            If okcount > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(" The sequence for this student did not update. Please record the information manually and contact OIT and provide them with this message." & ex.Message)
        Finally

        End Try
        Return False
    End Function
    Function GetSeq(ByVal theid As String) As Integer
        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select   max(seqno) from NC_StudyHall where id_num='" & TextBox1.Text & "' group by id_num"


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



            Dim theseq As Integer = 0

            theseq = catCMD.ExecuteScalar
            oODBCConnection.Close()

            Return theseq
        Catch ex As Exception
            MsgBox(" The sequence for this student did not update. Please record the information manually and contact OIT and provide them with this message." & ex.Message)
        Finally

        End Try
        Return -1
    End Function
    Function GetlastType(ByVal theid As String, ByVal theseq As Integer) As String
        Dim lastseq As Integer = theseq
        Try

            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select   TypeInfo from NC_StudyHall where id_num='" & theid & "'  and seqno=" & lastseq


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



            Dim thetype As String = "N"


            thetype = catCMD.ExecuteScalar
            oODBCConnection.Close()

            Return thetype

        Catch ex As Exception
            MsgBox(" The sequence for this student did not update. Please record the information manually and contact OIT and provide them with this message." & ex.Message)
        Finally

        End Try
        Return "N"
    End Function


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text.Length > 6 Then
            TextBox1.Text = TextBox1.Text.Substring(0, 6)
        End If
        TextBox1.Enabled = False

        Dim theseq As Integer
        Dim theid As String
        theid = TextBox1.Text
        Dim lastType As String = "O"
        theseq = GetSeq(TextBox1.Text)
        If theseq = -1 Then
            timeover = "N"
        Else
            timeover = " "
            lastType = GetlastType(TextBox1.Text, theseq)
        End If
        If lastType <> "I" And timeover <> "N" And theseq <> 0 And theseq <> -1 Then
            timeover = "Y"
        End If
        If timeover = "Y" Then
            Label3.Text = "Type is Incorrect--Button Disabled "
            Button2.Enabled = False
            Exit Sub
        End If
        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select   Coalesce(first_name + ' '  + middle_name + '  ' +  Last_name, first_name + ' '  + Last_name, last_name) from name_master where id_num='" & theid & "'"


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)


            Button2.Enabled = False
            Dim thecount As Integer = 0
            Dim thetime As DateTime
            Dim myReader As OdbcDataReader
            Dim okr As Boolean
            Dim thetype As String = "O"

            myReader = catCMD.ExecuteReader
            While myReader.Read()
                thecount = thecount + 1
                Label3.Text = myReader(0)
                thetime = Now()
                Label5.Text = thetime
                thetype = "O"
                okr = WritetheRecord(thetype, theid, thetime)
                If okr = True Then
                    ShowlastFive()
                End If
            End While
            If thecount = 0 Then
                Label3.Text = "Name Not Found"
                Button2.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(" You can retry the entry or manually record the information. If the retry fails, please contact OIT and provide them with this message." & ex.Message)
        End Try
        Label8.Text = Getlastperiodtime(TextBox1.Text)
       
        Dim emailsent As Boolean
        Dim subjecttime As String = Now().ToString
        Dim subjectline As String
        Dim theemail As String
        theemail = Getemailaddress(theid)
        subjectline = "Your  study time recorded at " & subjecttime
        'emailsent = SendMail(subjectline, "tsabol@n3t.com", "Y)
        Dim themessage As String
        themessage = "Your study time started at " & emailin.ToLongTimeString & " on " & emailin.ToLongDateString & " and ended at " & emailout.ToLongTimeString & " on " & emailout.ToLongDateString & ControlChars.CrLf & " . Your total time for this session is " & strstudyseconds & " seconds which is " & Label8.Text & " minutes."

        emailsent = SendMail("Studyhall@newberry.edu", "Study Hall", theemail, theemail, subjectline, themessage)
        TextBox1.Clear()

        Label7.Visible = True
        Label8.Visible = True
        TextBox1.Enabled = True
    End Sub
    Sub ShowlastFive()
        ListBox1.Items.Clear()

        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select   id_num,typeinfo,entrytime from NC_STUDYHALL where (FLAGENTRY <> 'D' or FLAGENTRY is NULL) and id_num='" & TextBox1.Text & "' order by SEQNO desc"


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



            Dim thecount As Integer = 0

            Dim myReader As OdbcDataReader

            Dim thetype As String = "I"
            Dim theseq As Integer = 0
            myReader = catCMD.ExecuteReader
            While myReader.Read()
                thecount = thecount + 1
                If thecount > 5 Then
                    oODBCConnection.Close()
                    Exit Sub
                End If

                ListBox1.Items.Add(myReader(1) & ControlChars.Tab & myReader(2))
            End While
            oODBCConnection.Close()
        Catch ex As Exception
            MsgBox(" Study hall history failed, please contact OIT and provide them with this message." & ex.Message)
        Finally

        End Try


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim curItem As String = ListBox1.SelectedItem.ToString()
        Dim thetime As String
        thetime = curItem.Substring(2)
        Dim okdelete As Boolean
        okdelete = DeletetheRecord(TextBox1.Text, thetime)
        If okdelete = True Then
            ShowlastFive()
            Button1.Enabled = True
            Button2.Enabled = True
        End If
    End Sub
    Function DeletetheRecord(ByVal theid As String, ByVal thetime As String) As Boolean
        Dim wtype As String
        Dim wid As String
        Dim wtime As DateTime
        wtime = Convert.ToDateTime(thetime)

        Dim deletesql As String
        deletesql = "Delete dbo.NC_StudyHall  where id_num='" & theid & "' and EntryTime='" & wtime & "'"
        Try

            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()







            Dim catCMD As OdbcCommand = New OdbcCommand(deletesql, oODBCConnection)



            Dim okcount As Integer = 0

            okcount = catCMD.ExecuteNonQuery
            oODBCConnection.Close()

            If okcount > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(" The sequence was not deleted. Please record the information manually and contact OIT and provide them with this message." & ex.Message)
        Finally

        End Try
        Return False
    End Function

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Function Getlastperiodtime(ByVal id As String) As String
        Try
            Dim minutestring As String = " "
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()


            strSQL = "Select cast( StudySeconds /60.0 as decimal(8,2)),OutTime,intime,studyseconds  from dbo.NC_StudyTime where id_num ='" & id & "' order by seqno"




            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



            Dim thecount As Integer = 0

            Dim myReader As OdbcDataReader

            Dim thetype As String = "I"
            Dim theseq As Integer = 0
            myReader = catCMD.ExecuteReader
            While myReader.Read()
                thecount = thecount + 1
                minutestring = myReader(0).ToString
                emailout = myReader(1)
                emailin = myReader(2)
                strstudyseconds = myReader(3).ToString
            End While
            oODBCConnection.Close()
            If thecount = 0 Then
                Return "No Study Time record found. Write down ID and notify OIT"
            Else
                Return minutestring

            End If
        Catch ex As Exception
            MsgBox(" Study hall Time calculation failed, please contact OIT and provide them with this message." & ex.Message)
        Finally

        End Try
    End Function


    Private Function SendMail(ByVal strMailFromEmail As String, ByVal strMailFromname As String, _
                          ByVal strMailToEmail As String, ByVal strMailToEmailName As String, _
                          ByVal strMailSubject As String, ByVal strMailBody As String, _
                          Optional ByVal strAttachment As String = Nothing) As String

        Dim strErrorMessage As String = "-NONE-"

         

        If InStr(Trim(strMailToEmail), " ") > 0 Then
            Return "EMAIL ERROR: SPACE IN ADDRESS"
            Exit Function
        End If

        Dim myMessage As New MailMessage
        myMessage.From = New MailAddress(strMailFromEmail, strMailFromname)
        myMessage.To.Add(New MailAddress(strMailToEmail, strMailToEmailName))
        myMessage.Subject = strMailSubject
        myMessage.Body = strMailBody
        myMessage.IsBodyHtml = True

        If strAttachment Is Nothing Then
            'do nothing        
        Else
            Dim myattachment As Attachment = New Net.Mail.Attachment(strAttachment)
            myMessage.Attachments.Add(myattachment)
        End If

        Dim server As New SmtpClient
        server.Port = 25
        server.Host = "mail.newberry.edu"
        server.DeliveryMethod = SmtpDeliveryMethod.Network
        server.UseDefaultCredentials = True


        Try
            server.Send(myMessage)
        Catch ex As Exception
            System.Threading.Thread.Sleep(2000) 'If there is a timing error, sleep for 2 seconds and try again

            Try
                server.Send(myMessage)
            Catch ex1 As Exception
                strErrorMessage = Replace(ex1.Message, ",", ";")
                Return False
            End Try




        End Try

        Return True
    End Function
    Function Getemailaddress(ByVal id_num As String) As String
        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select   email_address from name_master where id_num='" & id_num & "'"


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)


            Button1.Enabled = False
            Dim thecount As Integer = 0
            Dim thetime As DateTime
            Dim myReader As OdbcDataReader
            Dim okr As Boolean
            Dim thetype As String = "I"
            Dim theemailaddr As String
            myReader = catCMD.ExecuteReader
            While myReader.Read()
                thecount = thecount + 1

                theemailaddr = myReader(0)
                Return theemailaddr
            End While
            If thecount = 0 Then
                Return "tsabol@n3t.com"
             
            End If
        Catch ex As Exception
            MsgBox(" You can retry the entry or manually record the information. If the retry fails, please contact OIT and provide them with this message." & ex.Message)
        End Try

    End Function
End Class