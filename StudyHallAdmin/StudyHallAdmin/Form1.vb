Imports System.Data.Odbc
Public Class Form1

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select   passwordx from NC_StudyHall_password where passwordx='" & TextBox2.Text & "'"


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



            Dim thepass As String = "N'"

            thepass = catCMD.ExecuteScalar
            oODBCConnection.Close()
            If thepass <> TextBox2.Text Then
                Label3.Text = "Invalid Password"
            Else
                Button1.Enabled = True
                Button2.Enabled = True
                RadioButton1.Enabled = True
                RadioButton2.Enabled = True
                TextBox3.Enabled = True
                Label10.Visible = True
                Label11.Visible = True
                TextBox4.Visible = True
                TextBox5.Visible = True
                Button4.Visible = True
                Label12.Text = getthename(TextBox1.Text)
                Label12.Visible = True
                Call Showlast()
            End If

        Catch ex As Exception
            MsgBox(" The password system has malfunctioned--please contact OIT." & ex.Message)
        Finally

        End Try
    End Sub
    Sub Showlast()
        ListBox1.Items.Clear()

        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select   seqno ,    id_num ,typeinfo,entrytime from NC_STUDYHALL where (FLAGENTRY <> 'D' or FLAGENTRY is NULL) and id_num='" & TextBox1.Text & "' order by SEQNO desc"


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



            Dim thecount As Integer = 0

            Dim myReader As OdbcDataReader

            Dim thetype As String = "I"
            Dim theseq As Integer = 0
            myReader = catCMD.ExecuteReader
            While myReader.Read()
                thecount = thecount + 1
                 

                ListBox1.Items.Add(myReader(0) & "..." & myReader(1) & ControlChars.Tab & myReader(2) & ControlChars.Tab & myReader(3))
            End While
            oODBCConnection.Close()
        Catch ex As Exception
            MsgBox(" Study hall history failed, please contact OIT and provide them with this message." & ex.Message)
        Finally

        End Try

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button1.Enabled = False
        Button2.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        TextBox3.Enabled = False
        Label10.Visible = False
        Label11.Visible = False
        TextBox4.Visible = False
        TextBox5.Visible = False
        Button4.Visible = False
        Label12.Visible = False




        ListBox2.SelectedIndex = DateTime.Now.Month - 1
        ListBox3.SelectedIndex = DateTime.Now.Day - 1
        ListBox4.SelectedIndex = 0
        ListBox5.SelectedIndex = 16
        ListBox6.SelectedIndex = 0
    End Sub
    Function DeletetheRecord(ByVal theid As String, ByVal theseq As Integer) As Boolean


        Dim deletesql As String
        deletesql = "Delete dbo.NC_StudyHall  where id_num='" & theid & "' and seqno=" & theseq
        Try

            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New Odbc.OdbcConnection(sConnString)
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim okd As Boolean

        Dim theselectedval As String
        Dim thestrseq As String
        Dim instrloc As Integer
        instrloc = InStr(ListBox1.SelectedItem.ToString, "...")
        thestrseq = ListBox1.SelectedItem.ToString.Substring(0, instrloc - 1)
        okd = DeletetheRecord(TextBox1.Text, CInt(thestrseq))
        If okd = True Then
            MsgBox("Record Removed")
            Call Showlast()
        Else
            MsgBox("Delete Failed---Record not removed")
        End If
    End Sub
    Function WritetheRecord(ByVal thetype As String, ByVal theid As String, ByVal thetimestamp As DateTime) As Boolean
        Dim wtype As String
        Dim wid As String
        Dim wtime As DateTime
        Dim theseq As Integer
        Dim flagconstant As String = " "

         


        wtype = thetype
        wid = theid
        wtime = thetimestamp
        Dim writesql As String
        writesql = "Insert into dbo.NC_StudyHall (ID_NUM,SEQNO,TYPEINFO,ENTRYTIME,FLAGENTRY) values('" & wid & "'," & theseq & ",'" & wtype & "' ,'" & wtime & "','" & flagconstant & "')"
        Try

            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New Odbc.OdbcConnection(sConnString)
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
    Function WritetheRecordManual(ByVal newseq As Integer, ByVal thetype As String, ByVal theid As String, ByVal thetimestamp As DateTime) As Boolean
        Dim wtype As String
        Dim wid As String
        Dim wtime As DateTime
        Dim theseq As Integer
        Dim flagconstant As String = " "




        wtype = thetype
        wid = theid
        wtime = thetimestamp
        Dim writesql As String
        writesql = "Insert into dbo.NC_StudyHall (ID_NUM,SEQNO,TYPEINFO,ENTRYTIME,FLAGENTRY) values('" & theid & "'," & newseq & ",'" & thetype & "' ,'" & thetimestamp & "','" & flagconstant & "')"
        Try

            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New Odbc.OdbcConnection(sConnString)
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
            Showlast()
        End Try
        Return False
    End Function
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim nodup As Boolean
        Dim okcyc As Boolean
        Dim verifytype As String
        If TextBox3.Text.Length > 0 Then
            Dim i As Integer = TextBox3.Text.Length
            Dim iloop As Integer
            For iloop = 0 To (i - 1)
                If Not Char.IsDigit(TextBox3.Text(iloop)) Then
                    MessageBox.Show("Sequence Value must be numeric ")
                    Exit Sub
                End If
            Next
        Else
            MessageBox.Show("Sequence must have a value ")
            Exit Sub
        End If
        If RadioButton1.Checked = False And RadioButton2.Checked = False Then
            MessageBox.Show("Select either IN or OUT ")
            Exit Sub
        End If
        nodup = CheckforDUP(TextBox1.Text, CInt(TextBox3.Text))
        If nodup = False Then
            MessageBox.Show("The sequence number you wish to use already exists ")
            Exit Sub
        End If
        If RadioButton1.Checked = True Then
            verifytype = "I"
        Else
            verifytype = "O"
        End If
        okcyc = Checkfortype(TextBox1.Text, CInt(TextBox3.Text), verifytype)
        If okcyc = False Then
            MessageBox.Show("Inserting a Type of " & verifytype & " at Sequence " & TextBox3.Text & " would result in an invalid type chain")
            Exit Sub
        End If
        Dim newseq As Integer
        newseq = CInt(TextBox3.Text)
        Dim date1 As New Date(ListBox4.SelectedItem, ListBox2.SelectedItem, ListBox3.SelectedItem, ListBox5.SelectedItem, ListBox6.SelectedItem, 0)
        Dim datetime1 As New DateTime
        datetime1 = date1
        Dim successwrite As Boolean
        successwrite = WritetheRecordManual(CInt(TextBox3.Text), verifytype, TextBox1.Text, datetime1)
        If successwrite = False Then
            MessageBox.Show("The attempt to insert Type of " & verifytype & " at Sequence " & TextBox3.Text & " failed. Please contact OIT.")
        Else
            MessageBox.Show("The Type of " & verifytype & " at Sequence " & TextBox3.Text & " was successful. ")
        End If
    End Sub
    Function CheckforDUP(ByVal id_num As String, ByVal ckseqno As Integer) As Boolean
        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select   seqno ,    id_num ,typeinfo,entrytime from NC_STUDYHALL where (FLAGENTRY <> 'D' or FLAGENTRY is NULL) and id_num='" & id_num & "' and seqno =" & ckseqno

            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)

            Dim thecount As Integer = 0

            Dim myReader As OdbcDataReader

            Dim thetype As String = "I"
            Dim theseq As Integer = 0
            myReader = catCMD.ExecuteReader
            While myReader.Read()
                thecount = thecount + 1
            End While
            If thecount = 0 Then
                Return True
            Else
                Return False
            End If
            oODBCConnection.Close()

        Catch ex As Exception
            MsgBox(" Study hall duplicate check failed, please contact OIT and provide them with this message." & ex.Message)
        Finally

        End Try
    End Function
    Function getthename(ByVal id_num As String) As String
        If id_num <> "" Then
            Try
                Dim strSQL As String
                Dim oODBCConnection As OdbcConnection
                'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
                Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
                oODBCConnection = New Odbc.OdbcConnection(sConnString)
                oODBCConnection.Open()




                strSQL = "Select   First_name + ' ' + last_Name from name_master where  id_num='" & id_num & "'  "

                Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)

                Dim thecount As Integer = 0

                Dim myReader As OdbcDataReader

                Dim thename As String = "Not Found"
                Dim theseq As Integer = 0
                myReader = catCMD.ExecuteReader
                While myReader.Read()
                    thecount = thecount + 1
                    thename = myReader(0)
                End While
                If thecount = 0 Then
                    Return thename
                Else
                    Return thename
                End If
                oODBCConnection.Close()

            Catch ex As Exception
                MsgBox(" Name check failed, please contact OIT and provide them with this message." & ex.Message)
            Finally

            End Try
        End If
    End Function
    Function Checkfortype(ByVal id_num As String, ByVal ckseqno As Integer, ByVal cktype As String) As Boolean
        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()

            Dim Thecurrentvaluesseq As New ArrayList
            Dim thecurrentvaluestype As New ArrayList


            strSQL = "Select   seqno,     id_num ,typeinfo,entrytime from NC_STUDYHALL where (FLAGENTRY <> 'D' or FLAGENTRY is NULL) and id_num='" & id_num & "'  order by seqno"

            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)

            Dim thecount As Integer = 0

            Dim myReader As OdbcDataReader
            Dim holdlow As Integer = 0
            Dim thetype As String
            Dim theseq As Integer = 0
            myReader = catCMD.ExecuteReader
            While myReader.Read()
                thecount = thecount + 1
                Thecurrentvaluesseq.Add(myReader(0))
                thecurrentvaluestype.Add(myReader(2))
            End While
            Dim i As Integer
            Dim holdarrcount As Integer
            Dim holdprevvalue As String
            Dim holdnextvalue As String
            For i = 0 To Thecurrentvaluesseq.Count - 1
                If ckseqno > Thecurrentvaluesseq(i) Then
                    holdlow = Thecurrentvaluesseq(i)
                    holdarrcount = i
                End If
            Next
            If holdlow = 0 Then
                holdprevvalue = "B"
                holdnextvalue = thecurrentvaluestype(0)
            Else
                If holdarrcount = thecurrentvaluestype.Count Then
                    holdnextvalue = "E"
                    holdprevvalue = thecurrentvaluestype(thecurrentvaluestype.Count - 1)
                Else
                    holdprevvalue = thecurrentvaluestype(holdarrcount)

                    holdnextvalue = thecurrentvaluestype(holdarrcount + 1)
                End If
            End If
            Dim threevalues As String
            threevalues = holdprevvalue & cktype & holdnextvalue
            If threevalues = "BOI" Or threevalues = "BIO" Or threevalues = "OIE" Or threevalues = "IOE" Or threevalues = "OIO" Or threevalues = "IOI" Or threevalues = "OOE" Then

                Return True
            Else
                Return False
            End If
            oODBCConnection.Close()

        Catch ex As Exception
            MsgBox(" Study hall type sequencing check  failed, please contact OIT and provide them with this message." & ex.Message)
        Finally

        End Try
    End Function

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If TextBox4.Text <> TextBox5.Text Then
            MsgBox("Password and Confirmation entries do not match")
        Else



            Try
                Dim strSQL As String
                Dim oODBCConnection As OdbcConnection
                'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
                Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPRD;   Uid=sa;   Pwd=jenzadmin"
                oODBCConnection = New Odbc.OdbcConnection(sConnString)
                oODBCConnection.Open()




                strSQL = "update  NC_StudyHall_password set passwordx = '" & TextBox5.Text & "' where passwordx='" & TextBox2.Text & "'"


                Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



                Dim okcount As Integer = 0

                okcount = catCMD.ExecuteNonQuery
                oODBCConnection.Close()

                If okcount > 0 Then
                    MsgBox("Password updated")
                Else
                    MsgBox("Password update failed---previous password still valid---please contact OIT")
                End If
            Catch ex As Exception
                MsgBox(" The sequence was not deleted. Please record the information manually and contact OIT and provide them with this message." & ex.Message)
            Finally

            End Try

               End if
    End Sub
End Class
