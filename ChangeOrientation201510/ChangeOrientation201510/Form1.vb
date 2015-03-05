Imports System
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.SqlClient

    Public Class Form1
        Public theto As String
        Public thefrom As String
        Public thesess As String
        Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
            Label8.Text = " "
            Try
                Dim ok As Boolean
                Dim updatedt As DateTime
                updatedt = DateTime.Now()
                Dim updateuser As String = "PROG"
                Dim updatejob As String = "CHGORI"
                Dim strSQL2 As String
                Dim oODBCConnection2 As OdbcConnection

                Dim sConnString2 As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEprd;   Uid=sa;   Pwd=jenzadmin"
                oODBCConnection2 = New Odbc.OdbcConnection(sConnString2)
                oODBCConnection2.Open()

                strSQL2 = "Update nc_orientation set UDEF_DTE_1 = '" & theto & "', JS_SESSION = '" & thesess & "' , Job_time = '" & updatedt & "'" & " , USER_Name = '" & updateuser & "'" & " , Job_Name = '" & updatejob & "' where id_num = '" & TextBox1.Text & "'"

                Dim catCMD2 As OdbcCommand = New OdbcCommand(strSQL2, oODBCConnection2)







                Dim okcount As Integer = 0

                okcount = catCMD2.ExecuteNonQuery
                oODBCConnection2.Close()

                If okcount > 0 Then
                    Label8.Text = "Update was Successful"
                Else
                    Label8.Text = "Update was NOT Successful"
                End If
                ok = RefreshtheBoxes()
            Catch ex As Exception
                MsgBox(" The update  did not occur. Please record the information manually and contact Campus Technology and provide them with this message." & ex.Message)
            Finally

            End Try


            Button1.Text = "Done"
            Button1.Enabled = False
        End Sub










        Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            ListBox1.Items.Clear()
            ListBox2.Items.Clear()
            ListBox3.Items.Clear()
            ListBox4.Items.Clear()
            Button1.Enabled = False
        ListBox1.Items.Add("June 26")
        ListBox1.Items.Add("July 24")
        ListBox1.Items.Add("August 20")
        Label8.Text = " "
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            Dim sConnString As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEprd;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()

        strSQL = "Select last_name + ', ' + first_name + '...' + Cast(a.id_num as varchar(10)),b.UDEF_DTE_1  from dbo.name_master a left join nc_orientation b on a.id_num=b.id_num  and b.yr_cde='2015' and b.trm_cde='10' where a.id_num in (Select id_num from  dbo.NC_Orientation where yr_cde='2015' and trm_cde='10')"

            strSQL += " order by last_name + ', ' + first_name "
            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)

            Dim strins As String



            Dim myReader As OdbcDataReader = catCMD.ExecuteReader()

            Dim thename As String
            Dim thedate As String
            Dim thecount As Integer = 0
            Dim inscount As Integer = 0


            Dim agroup As String

            If myReader.HasRows Then


                Dim isok As Boolean
                Do While myReader.Read()

                    thecount = thecount + 1


                    Try

                        ' define connection string and stored procedure name 





                        ' define as stored procedure


                        If myReader(1) = "2015-06-26" Then
                            isok = AddtoJune(myReader(0))
                        End If
                        If myReader(1) = "2015-07-24" Then
                            isok = AddtoJuly(myReader(0))
                        End If
                        If myReader(1) = "2015-08-20" Then
                            isok = AddtoAugust(myReader(0))
                        End If



                    Catch o As OdbcException

                        MsgBox(o.Message.ToString)

                    Finally



                    End Try

                Loop
                oODBCConnection.Close()



            End If

        End Sub
        Function AddtoJune(thename As String) As Boolean
            ListBox4.Items.Add(thename)
            Return True
        End Function
        Function AddtoJuly(thename As String) As Boolean
            ListBox2.Items.Add(thename)
            Return True
        End Function
        Function AddtoAugust(thename As String) As Boolean
            ListBox3.Items.Add(thename)
            Return True
        End Function

        Private Sub ListBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox4.SelectedIndexChanged
            Label5.Text = " from June 26"
            Dim thetext As String
            Dim theacn As String
            Dim thelabelname As String
            Dim i As Integer
            thetext = ListBox4.SelectedItem.ToString
            i = InStr(thetext, "...")
            theacn = thetext.Substring(i + 2)
            thelabelname = thetext.Substring(0, i - 1)
            Label7.Text = thelabelname
            TextBox1.Text = theacn
            thefrom = "2015-06-26"

        End Sub

        Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
            Label5.Text = " from August 20"
            Dim thetext As String
            Dim theacn As String
            Dim thelabelname As String
            Dim i As Integer
            thetext = ListBox3.SelectedItem.ToString
            i = InStr(thetext, "...")
            theacn = thetext.Substring(i + 2)
            thelabelname = thetext.Substring(0, i - 1)
            Label7.Text = thelabelname
            TextBox1.Text = theacn
            thefrom = "2015-08-20"

        End Sub

        Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
            Label5.Text = " from July 24"
            Dim thetext As String
            Dim theacn As String
            Dim thelabelname As String
            Dim i As Integer
            thetext = ListBox2.SelectedItem.ToString
            i = InStr(thetext, "...")
            theacn = thetext.Substring(i + 2)
            thelabelname = thetext.Substring(0, i - 1)
            Label7.Text = thelabelname
            TextBox1.Text = theacn
            thefrom = "2015-07-24"
        End Sub

        Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
            Button1.Enabled = True
            Label6.Text = " to " & ListBox1.SelectedItem.ToString
            If ListBox1.SelectedItem.ToString = "June 26" Then
            theto = "2015-06-26"
                thesess = "Jun 26 20"
            End If
            If ListBox1.SelectedItem.ToString = "July 24" Then
                theto = "2015-07-24"
                thesess = "Jul 24 20"
            End If
            If ListBox1.SelectedItem.ToString = "August 20" Then
                thesess = "Aug 20 20"
                theto = "2015-08-20"
            End If


        End Sub
        Function RefreshtheBoxes() As Boolean
            ListBox1.Items.Clear()
            ListBox2.Items.Clear()
            ListBox3.Items.Clear()
            ListBox4.Items.Clear()



            ListBox1.Items.Add("June 26")
            ListBox1.Items.Add("July 24")
            ListBox1.Items.Add("August 20")
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            Dim sConnString As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEprd;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()

            strSQL = "Select last_name + ', ' + first_name + '...' + Cast(a.id_num as varchar(10)),b.UDEF_DTE_1  from dbo.name_master a left join nc_orientation b on a.id_num=b.id_num where a.id_num in (Select id_num from  dbo.NC_Orientation where yr_cde='2015' and trm_cde='10')"

            strSQL += " order by last_name + ', ' + first_name "
            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)

            Dim strins As String



            Dim myReader As OdbcDataReader = catCMD.ExecuteReader()

            Dim thename As String
            Dim thedate As String
            Dim thecount As Integer = 0
            Dim inscount As Integer = 0


            Dim agroup As String

            If myReader.HasRows Then


                Dim isok As Boolean
                Do While myReader.Read()

                    thecount = thecount + 1


                    Try

                        ' define connection string and stored procedure name 





                        ' define as stored procedure


                        If myReader(1) = "2015-06-26" Then
                            isok = AddtoJune(myReader(0))
                        End If
                        If myReader(1) = "2015-07-24" Then
                            isok = AddtoJuly(myReader(0))
                        End If
                        If myReader(1) = "2015-08-20" Then
                            isok = AddtoAugust(myReader(0))
                        End If



                    Catch o As OdbcException

                        MsgBox(o.Message.ToString)
                        Return False
                    Finally



                    End Try

                Loop
                oODBCConnection.Close()
                Return True


            End If
        End Function

    End Class
