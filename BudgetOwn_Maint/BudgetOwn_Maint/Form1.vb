Imports System
Imports System.Data
Imports System.Data.Odbc


Public Class Form1
    Public x As String
    Public idinfo As String
    Dim nameinfo As String
    Dim Nameinfo2 As String

    Private Function FetchData() As DataTable


        Dim sqlstr As String = "Select last_name,first_name,a.id_num, a.PrimaryGroup, a.SecondaryGroup, b.PrimaryGroupDesc, b.SecondaryGroupDesc from dbo.NC_Budget_OWN a left join NC_budget_secondary  b on a.primaryGroup = b.primarygroup and a.SecondaryGroup=b.SecondaryGroup left join name_master nm on a.id_num=nm.id_num order by last_name,first_name,PrimaryGroupDesc,SecondaryGroupDesc "


        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseply;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand(sqlstr, Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)


            Try
                Dim Table As New DataTable
                sAdapter.Fill(Table)
                Return Table
            Catch o As Exception

                MsgBox(o.Message.ToString)
            End Try
        End Using
        Label7.Text = " "
    End Function
     
     
    Private Function FetchDataselective() As DataTable

        Dim sqlstr As String = "Select last_name,first_name,a.id_num, a.PrimaryGroup, a.SecondaryGroup, b.PrimaryGroupDesc, b.SecondaryGroupDesc from dbo.NC_Budget_OWN a left join NC_budget_secondary  b on a.primaryGroup = b.primarygroup and a.SecondaryGroup=b.SecondaryGroup left join name_master nm on a.id_num=nm.id_num order by PrimaryGroupDesc,SecondaryGroupDesc,a.id_num where a.id_num='" & idinfo & "'"
        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseply;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand("SELECT distinct * from NC_control_info  where ctlname='" & nameinfo & "'   order by ctlname,ctlfld2,ctlfld3    ", Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)


            Try
                Dim Table As New DataTable
                sAdapter.Fill(Table)
                Return Table
            Catch o As Exception

                MsgBox(o.Message.ToString)
            End Try
        End Using
        Label7.Text = " "
    End Function
     
    Private Function FetchNames() As DataTable


        Dim sqlstr As String = "Select  distinct a.PrimaryGroup, a.SecondaryGroup, b.PrimaryGroupDesc, b.SecondaryGroupDesc from dbo.NC_Budget_OWN a left join NC_budget_secondary  b on a.primaryGroup = b.primarygroup and a.SecondaryGroup=b.SecondaryGroup      order by PrimaryGroupDesc,SecondaryGroupDesc "


        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseply;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand(sqlstr, Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)


            Try
                Dim Table As New DataTable
                sAdapter.Fill(Table)
                Return Table
            Catch o As Exception

                MsgBox(o.Message.ToString)
            End Try
        End Using

    End Function
     
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = FetchData()
        DataGridView2.DataSource = FetchNames()
        Button6.Visible = False
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        Dim cellroomstatus As String
        TextBox1.Text = DataGridView1.CurrentRow.Cells(2).Value.ToString()
        TextBox2.Text = DataGridView1.CurrentRow.Cells(3).Value.ToString()
        TextBox3.Text = DataGridView1.CurrentRow.Cells(4).Value.ToString()
         
        nameinfo = DataGridView1.CurrentRow.Cells(0).Value.ToString()
        Nameinfo2 = DataGridView1.CurrentRow.Cells(1).Value.ToString()
        RichTextBox1.Text = " "
        Dim foundtext As String
        foundtext = GetName(TextBox1.Text)
        RichTextBox1.Text = foundtext
        Button6.Visible = False
        'DataGridView2.DataSource = "null"
        'DataGridView3.DataSource = "null"

        'DataGridView2.DataSource = FetchRoommateInfo()
        'DataGridView3.DataSource = FetchRoomPrefs()
        'If Me.DataGridView3.Rows.Count > 0 Then
        'For i As Integer = 0 To Me.DataGridView3.Rows.Count - 1
        'If IsDBNull(Me.DataGridView3.Rows(i).Cells(5).Value) = False Then
        'If Me.DataGridView3.Rows(i).Cells(5).Value = "A" Then
        'Me.DataGridView3.Rows(i).DefaultCellStyle.ForeColor = Color.Red
        'End If
        'End If
        'Next

        'End If
    End Sub
    Private Sub DataGridView2_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick

        TextBox2.Text = DataGridView2.CurrentRow.Cells(0).Value.ToString()
        TextBox3.Text = DataGridView2.CurrentRow.Cells(1).Value.ToString()
        Button6.Visible = False
         
    End Sub

    Function GetName(ByVal id_num As String) As String
        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPLY;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select  first_name + ', ' + last_name   from name_master where id_num='" & id_num & "'"


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



            Dim thecount As Integer = 0

            Dim myReader As OdbcDataReader
            Dim thedescription As String

            myReader = catCMD.ExecuteReader
            While myReader.Read()
                thecount = thecount + 1
                If IsDBNull(myReader(0)) = True Then
                    thedescription = "No name was found"
                Else

                    thedescription = myReader(0)
                End If
            End While
            If thecount = 0 Then

                Return "No name was found"
            Else
                Return thedescription
            End If

        Catch ex As Exception
            MsgBox("   Please contact OIT and provide them with this message." & ex.Message)
        End Try
    End Function



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        DataGridView1.DataSource = FetchData()
        'DataGridView2.DataSource = "null"
        'DataGridView3.DataSource = "null"

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim issuccess As Boolean
        issuccess = WritetheRecord(TextBox1.Text, TextBox2.Text, TextBox3.Text)
        If issuccess = True Then
            DataGridView1.DataSource = FetchData()
            DataGridView2.DataSource = FetchNames()
        End If
    End Sub


    Function WritetheRecord(ByVal id_num As String, ByVal PrimaryGroup As String, ByVal SecondaryGroup As String) As Boolean
        Dim outid As String
        Dim Pgroup As String
        Dim Sgroup As String
         
        Dim flagconstant As String = " "

        outid = id_num
        Pgroup = PrimaryGroup
        Sgroup = SecondaryGroup
         


        Dim writesql As String
        writesql = "Insert into dbo.NC_Budget_OWN (PrimaryGroup,SecondaryGroup,id_num  ) values('" & Trim(Pgroup) & "','" & Trim(Sgroup) & "' ,'" & Trim(outid) & "')"
        Try

            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPly;   Uid=sa;   Pwd=jenzadmin"
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
            MsgBox(" The  insert did not occur. Please record the information manually and contact OIT and provide them with this message." & ex.Message)
        Finally

        End Try
        Return False
    End Function

     






             
           
            

    Function DeleteallRecords(ByVal id_num As String) As Boolean
        Dim wname As String
        Dim Pgroup As String
        Dim SGroup As String

        Dim flagconstant As String = " "

        wname = id_num
         


        Dim deletesql As String
        deletesql = "delete dbo.NC_Budget_OWN "

        deletesql += "where  id_num = '" & Trim(wname) & "'"


        Try

            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPly;   Uid=sa;   Pwd=jenzadmin"
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
            MsgBox(" The  deletion did not occur. Please record the information manually and contact OIT and provide them with this message." & ex.Message)
        Finally

        End Try
        Return False
    End Function
    Function DeletetheRecord(ByVal id_num As String, ByVal PrimaryG As String, ByVal SecondaryG As String) As Boolean
        Dim wname As String
        Dim Pgroup As String
        Dim SGroup As String

        Dim flagconstant As String = " "

        wname = id_num
        Pgroup = PrimaryG
        SGroup = SecondaryG



        Dim deletesql As String
        deletesql = "delete dbo.NC_Budget_OWN "

        deletesql += "where  id_num = '" & Trim(wname) & "'  and "
        deletesql += "  PrimaryGroup = '" & Trim(Pgroup) & "'  and "
        deletesql += "  SecondaryGroup = '" & Trim(SGroup) & "'"


        Try

            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPly;   Uid=sa;   Pwd=jenzadmin"
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
            MsgBox(" The  deletion did not occur. Please record the information manually and contact OIT and provide them with this message." & ex.Message)
        Finally

        End Try
        Return False
    End Function


     








    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
       
        Dim issuccess As Boolean
        issuccess = DeletetheRecord(TextBox1.Text, TextBox2.Text, TextBox3.Text)
        If issuccess = True Then
            DataGridView1.DataSource = FetchDataselective()
            DataGridView2.DataSource = FetchNames()
        End If



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim strSQL2 As String
        Dim oODBCConnection2 As OdbcConnection
        Dim sConnString As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEply;   Uid=sa;   Pwd=jenzadmin"
        Try
            oODBCConnection2 = New Odbc.OdbcConnection(sConnString)
            oODBCConnection2.Open()
            strSQL2 = "Select PrimaryGroup,SecondaryGroup    from dbo.NC_Budget_Own    where    id_num ='" & TextBox4.Text & "'"


            Dim catCMD2 As OdbcCommand = New OdbcCommand(strSQL2, oODBCConnection2)






            Dim myReader As OdbcDataReader = catCMD2.ExecuteReader()



            Dim thecount As Integer = 0




            If myReader.HasRows Then



                Do While myReader.Read()

                    thecount = thecount + 1





                    Try






                        Dim createok As Boolean


                        createok = WritetheRecord(TextBox1.Text, myReader(0), myReader(1))






                    Catch o As OdbcException

                        MsgBox(o.Message.ToString)

                    Finally



                    End Try

                Loop
                Label7.Text = CStr(thecount) & " records added"
            End If
                Catch o1 As OdbcException

            MsgBox(o1.Message.ToString)
        End Try
        oODBCConnection2.Close()











    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        RichTextBox1.Text = "All records from " & GetName(TextBox4.Text) & " will be duplicated for " & GetName(TextBox1.Text)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim okdelete As Boolean
        RichTextBox1.Text = "All records for " & GetName(TextBox1.Text) & " will be deleted "

        Button6.Visible = True
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim okdelete As Boolean

        okdelete = DeleteallRecords(TextBox1.Text)
        RichTextBox1.Text = "All records for " & GetName(TextBox1.Text) & " HAVE BEEN DELETED  "
        DataGridView1.DataSource = FetchData()
        Button6.Visible = False
    End Sub
End Class

