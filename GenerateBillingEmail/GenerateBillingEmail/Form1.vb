Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Web
Imports System.Net.Mail
Imports System.IO


Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim theend As Date = Date.Today

        Dim theendstr As String
        theendstr = theend.ToString("yyyy-MM-dd")
        Dim thestart As Date = theend.AddDays(-7)
        Dim thestartstr As String
        thestartstr = thestart.ToString("yyyy-MM-dd")
        TextBox6.Text = thestartstr
        TextBox7.Text = theendstr
        TextBox5.Text = "Your Newberry College Account Balance has changed"
    End Sub
    Sub Getthehtml()

        Dim thehtmlinfo As String = ""



        Try

            ' Create an instance of StreamReader to read from a file. 
            ' The using statement also closes the StreamReader. 
            Using sr As StreamReader = New StreamReader("H:\emailhtmlsrc.txt")
                Dim line As String
                ' Read and display lines from the file until the end of  
                ' the file is reached. 
                line = sr.ReadLine()

                TextBox1.Text = line & ControlChars.CrLf

                While sr.Peek() >= 0

                    TextBox1.Text += line & ControlChars.CrLf
                    line = sr.ReadLine()
                End While
            End Using
        Catch ex As Exception
            MsgBox("ErrorToString:" & ex.Message)
        End Try
    End Sub

    Function GetEmail(ByVal lineinfo As String) As String
        Dim outemail As String
        Dim thestart As Integer = 0
        Dim thelen As Integer = 0
        thestart = InStr(lineinfo, "|")
        thelen = lineinfo.Length
        outemail = lineinfo.Substring(thestart, thelen - thestart)
        Return outemail
    End Function
    Function GetAmount(ByVal lineinfo As String) As String
        Dim outamount As String
        Dim thestart As Integer = 0
        Dim theend As Integer = 0
        thestart = InStr(lineinfo, "*")
        theend = InStr(lineinfo, "~")
        theend = theend - 1
        outamount = lineinfo.Substring(thestart, theend - thestart)
        Return outamount
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
                Return "Error"
            End Try




        End Try

        Return "Good"
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim i As Integer
        Dim emailsent As String
        Dim subjecttime As String = Now().ToString
        Dim subjectline As String
        Dim theemail As String
        Dim theid As String
        Dim theamount As Decimal
        Dim parsemessage As String
        Dim holdparsemessage As String
        Dim sentemails As Integer = 0
        parsemessage = TextBox1.Text.Replace("|", "<br>")
        holdparsemessage = parsemessage
        If CheckBox1.Checked = True Then
            parsemessage = parsemessage & "<br><h4>Do not reply to this message via email</h4>"
        End If
        For i = 0 To ListBox2.Items.Count - 1
            'theemail = Getemailaddress(theid)
            theemail = Trim(GetEmail(ListBox2.Items(i).ToString))
            theid = Getid(theemail)
            theamount = CDec(Trim(GetAmount(ListBox2.Items(i).ToString)))
            subjectline = TextBox5.Text
            'emailsent = SendMail(subjectline, "tsabol@n3t.com", "Y)
            Dim themessage As String
            If CheckBox4.Checked = True Then
                parsemessage = holdparsemessage
                parsemessage = parsemessage & "<br><h4>" & theemail & " Do not reply to this message via email</h4>"
                theemail = TextBox8.Text
            End If
            themessage = "<HTML><BODY>" & parsemessage & "</BODY></HTML>"

            emailsent = SendMail(TextBox4.Text, TextBox3.Text, theemail, theemail, subjectline, themessage)
            If emailsent = "Good" Then
                sentemails = sentemails + 1
                updategoodemail(theid, theemail, theamount)
            Else
                updateerroremail(theid, theemail, theamount)
            End If
        Next i
        Label2.Text = sentemails & " emails sent"
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim theitem As String
        theitem = ListBox1.SelectedItem.ToString
        Dim i As Integer
        Dim pos As Integer
        pos = InStr(theitem, ".....................")
        i = pos + 20
        ListBox2.Items.Add(theitem.Substring(i))
        Label2.Text = " "
    End Sub



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim emailsent As String
        Dim subjecttime As String = Now().ToString
        Dim subjectline As String
        Dim theemail As String
        Dim parsemessage As String
        parsemessage = TextBox1.Text.Replace("|", "<br>")
        'theemail = Getemailaddress(theid)

        If CheckBox1.Checked = True Then
            parsemessage = parsemessage & "<br><h4>Do not reply to this message via email</h4>"
        End If
        theemail = TextBox2.Text
        subjectline = TextBox5.Text
        'emailsent = SendMail(subjectline, "tsabol@n3t.com", "Y)
        Dim themessage As String
        themessage = "<HTML><BODY>" & parsemessage & "</BODY></HTML>"

        emailsent = SendMail(TextBox4.Text, TextBox3.Text, theemail, theemail, subjectline, themessage)
        Label2.Text = "Email Sent"
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ListBox1.Items.Add("000000" & ".....................*" & 0.0 & "~....................+" & "Manual" & ".....................|" & TextBox2.Text)

    End Sub
    Public Function updateerroremail(id_num As String, stremail As String, balance As Decimal) As Boolean
        Dim strSQL As String
        Dim soSQLConnection As SqlConnection
        Dim sConnString As String = "Server=ncsql1;Database=Tmseprd;Uid=sa;pwd=jenzadmin;"

        soSQLConnection = New SqlClient.SqlConnection(sConnString)
        soSQLConnection.Open()


        Dim cmdx As New SqlCommand("[dbo].[NC_Bill_Email_ErrorInsert]", soSQLConnection)
        cmdx.CommandType = CommandType.StoredProcedure
        cmdx.Parameters.AddWithValue("@id_num", id_num)
        cmdx.Parameters.AddWithValue("@email", stremail)
        cmdx.Parameters.AddWithValue("@balance_amt", balance)



        Try
            cmdx.ExecuteNonQuery()
            soSQLConnection.Close()
            Return True

        Catch o As OdbcException

            MsgBox(" log  failed for " & id_num)
            soSQLConnection.Close()
            Return False
        End Try




    End Function
    Public Function updategoodemail(id_num As String, stremail As String, balance As Decimal) As Boolean
        Dim strSQL As String
        Dim soSQLConnection As SqlConnection
        Dim sConnString As String = "Server=ncsql1;Database=Tmseprd;Uid=sa;pwd=jenzadmin;"

        soSQLConnection = New SqlClient.SqlConnection(sConnString)
        soSQLConnection.Open()


        Dim cmdx As New SqlCommand("[dbo].[NC_Bill_Email_GoodInsert]", soSQLConnection)
        cmdx.CommandType = CommandType.StoredProcedure
        cmdx.Parameters.AddWithValue("@id_num", id_num)
        cmdx.Parameters.AddWithValue("@email", stremail)
        cmdx.Parameters.AddWithValue("@balance_amt", balance)



        Try
            cmdx.ExecuteNonQuery()
            soSQLConnection.Close()
            Return True

        Catch o As OdbcException

            MsgBox(" log  failed for " & id_num)
            soSQLConnection.Close()
            Return False
        End Try




    End Function
    Function GetSender_name() As String
        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select   CtlFld1 from dbo.NC_CONTROL_INFO   where CTLNAME='EMAIL_NAME' and CTLFLD2='BILLEMAIL' and CTLFLD3=0"


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



            Dim thename As String

            thename = catCMD.ExecuteScalar
            oODBCConnection.Close()

            Return thename
        Catch ex As Exception
            MsgBox(" The Could not retrieve the default sender name." & ex.Message)
        Finally

        End Try
        Return "Enter Manually"
    End Function
    Function Getid(ByVal email As String) As String
        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select   id_num from dbo.name_master   where email_address = '" & email & "'"


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



            Dim theid As String

            theid = catCMD.ExecuteScalar
            oODBCConnection.Close()

            Return theid
        Catch ex As Exception
            MsgBox(" The Could not retrieve the id for " & email & "Error:" & ex.Message)
        Finally

        End Try

    End Function


    Function GetSender_email() As String
        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select   CtlFld1 from dbo.NC_CONTROL_INFO   where CTLNAME='EMAIL_Address' and CTLFLD2='BILLEMAIL' and CTLFLD3=0"


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



            Dim theaddress As String

            theaddress = catCMD.ExecuteScalar
            oODBCConnection.Close()

            Return theaddress
        Catch ex As Exception
            MsgBox(" The Could not retrieve the default sender name." & ex.Message)
        Finally

        End Try
        Return "Enter Manually"
    End Function

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Try
                Dim strSQL As String
                Dim oODBCConnection As OdbcConnection
                'Dim sConnString As String = "Driver={SQL Server};   


                Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;  Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
                oODBCConnection = New Odbc.OdbcConnection(sConnString)
                oODBCConnection.Open()
                strSQL = "update  dbo.NC_CONTROL_INFO set CTLFLD1='" & Trim(TextBox3.Text) & "'   where CTLNAME='EMAIL_NAME' and CTLFLD2='BillEMAIL' and CTLFLD3='0'"


                Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



                Dim okcount As Integer = 0

                okcount = catCMD.ExecuteNonQuery
                oODBCConnection.Close()

                If okcount > 0 Then
                    MsgBox("Sender Name updated")
                Else
                    MsgBox("Sender Name update failed---previous stored name still valid---please contact OIT")
                End If
            Catch ex As Exception
                MsgBox(" Base update failed---contact OIT and provide them with this message." & ex.Message)
            Finally

            End Try

        End If
    End Sub
    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            Try
                Dim strSQL As String
                Dim oODBCConnection As OdbcConnection
                'Dim sConnString As String = "Driver={SQL Server};   


                Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;  Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
                oODBCConnection = New Odbc.OdbcConnection(sConnString)
                oODBCConnection.Open()
                strSQL = "update  dbo.NC_CONTROL_INFO set Ctlfld1='" & Trim(TextBox4.Text) & "'   where CTLNAME='EMAIL_ADDRESS' and CTLFLD2='BillEMAIL' and CTLFLD3='0'"


                Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



                Dim okcount As Integer = 0

                okcount = catCMD.ExecuteNonQuery
                oODBCConnection.Close()

                If okcount > 0 Then
                    MsgBox("Sender email address updated")
                Else
                    MsgBox("Sender email address update failed---previous stored address still valid---please contact OIT")
                End If
            Catch ex As Exception
                MsgBox(" Base update failed---contact OIT and provide them with this message." & ex.Message)
            Finally

            End Try

        End If
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        ListBox2.Items.Remove(ListBox2.SelectedItem)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim ix As Integer
        For ix = 0 To ListBox1.Items.Count - 1
            Dim theitem As String
            theitem = ListBox1.Items.Item(ix).ToString
            Dim i As Integer
            Dim pos As Integer
            pos = InStr(theitem, ".....................")
            i = pos + 20
            ListBox2.Items.Add(theitem.Substring(i))
            Label2.Text = " "
        Next

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()


            strSQL = "Select a.ID_NUM, sum(a.trans_amt)   , last_name +   ', ' +  first_name ,  email_address   from trans_hist a left join name_master nm on a.id_num=nm.id_num   where SUBSID_CDE='AR'   and                         a.job_time"
            strSQL += " between '" & TextBox6.Text & "' and '" & TextBox7.Text & "'   group by a.id_num,last_name +   ', ' +  first_name  ,  email_address having sum(trans_amt) < ( select ctlfld1 from NC_CONTROL_INFO where ctlname='BILLING_DEBIT_AMOUNT')"
            strSQL += " OR sum(trans_amt) > ( select ctlfld1 from NC_CONTROL_INFO where ctlname='BILLING_CREDIT_AMOUNT')"
            strSQL += "  order by last_name +   ', ' +  first_name,a.id_num"


            'strSQL = "Select   Coalesce(last_name + ' '  + middle_name + '  ' +  first_name, last_name  + ' '  + first_name, last_name),  email_address   from name_master where id_num in (select id_num from NC_CURR_STUDENTS) order by last_name "


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



            Dim thecount As Integer = 0
            Dim thetime As DateTime
            Dim myReader As OdbcDataReader
            Dim okr As Boolean
            Dim thetype As String = "I"

            myReader = catCMD.ExecuteReader
            While myReader.Read()
                thecount = thecount + 1
                ListBox1.Items.Add(myReader(0) & ".....................*" & myReader(1) & "~....................+" & myReader(2) & ".....................|" & myReader(3))
            End While

        Catch ex As Exception
            MsgBox(" You can retry the entry or manually record the information. If the retry fails, please contact OIT and provide them with this message." & ex.Message)
        End Try
        Dim theename As String
        theename = GetSender_name()
        TextBox3.Text = theename
        Dim theeaddress As String
        theeaddress = GetSender_email()
        TextBox4.Text = theeaddress
        Call Getthehtml()
    End Sub
End Class
