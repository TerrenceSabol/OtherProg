Imports System
Imports System.Data
Imports System.Data.Odbc


Public Class Form1
    Public x As String
    Public nameinfo As String

    Private Function FetchData() As DataTable





        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseply;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand("SELECT a.id_num, last_name,first_name,Sport,Active_mbr,Study_Hall_Watch,ID from NC_ROSTERS a left join name_master nm on a.id_num=nm.id_num    order by last_name,first_name,a.id_num ", Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)


            Try
                Dim Table As New DataTable
                sAdapter.Fill(Table)
                Return Table
            Catch o As Exception

                MsgBox(o.Message.ToString)
            End Try
        End Using

    End Function
     
     
    Private Function FetchDataselective() As DataTable


         
        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseply;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand("SELECT a.id_num,last_name,first_name,Sport,Active_mbr,Study_Hall_Watch,ID from NC_ROSTERS a left join name_master nm on a.id_num=nm.id_num where sport='" & nameinfo & "'   order by last_name,first_name,a.id_num ", Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)
            Try
                Dim Table As New DataTable
                sAdapter.Fill(Table)
                Return Table
            Catch o As Exception

                MsgBox(o.Message.ToString)
            End Try
        End Using

    End Function
     
    Private Function FetchNames() As DataTable





        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseply;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand("SELECT distinct Sport,description from NC_ROSTERS_DESC     order by description     ", Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)


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
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        Dim cellroomstatus As String
        TextBox1.Text = DataGridView1.CurrentRow.Cells(0).Value.ToString()
        TextBox2.Text = DataGridView1.CurrentRow.Cells(3).Value.ToString()
        TextBox3.Text = DataGridView1.CurrentRow.Cells(4).Value.ToString()
        TextBox4.Text = DataGridView1.CurrentRow.Cells(5).Value.ToString()

        nameinfo = DataGridView1.CurrentRow.Cells(3).Value.ToString()
        RichTextBox1.Text = " "
        Dim foundtext As String
        foundtext = GetDesc(nameinfo)
        RichTextBox1.Text = foundtext
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

        Dim cellroomstatus As String
        nameinfo = DataGridView2.CurrentRow.Cells(0).Value.ToString()
        RichTextBox1.Text = " "
        Dim foundtext As String
        foundtext = GetDesc(nameinfo)

        RichTextBox1.Text = foundtext
        DataGridView1.DataSource = "null"
        DataGridView1.DataSource = FetchDataselective()
        TextBox1.Text = " "
        TextBox2.Text = " "
        TextBox3.Text = " "
        TextBox4.Text = " "

    End Sub

    Function GetDesc(ByVal nametofind As String) As String
        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPLY;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select  description   from NC_ROSTERS_DESC where SPORT='" & Trim(nametofind) & "'"


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



            Dim thecount As Integer = 0

            Dim myReader As OdbcDataReader
            Dim thedescription As String

            myReader = catCMD.ExecuteReader
            While myReader.Read()
                thecount = thecount + 1
                If IsDBNull(myReader(0)) = True Then
                    thedescription = "No description is available"
                Else

                    thedescription = myReader(0)
                End If
            End While
            If thecount = 0 Then

                Return "No description is available"
            Else
                Return thedescription
            End If

        Catch ex As Exception
            MsgBox(" You can retry the entry or manually record the information. If the retry fails, please contact OIT and provide them with this message." & ex.Message)
        End Try
    End Function



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        DataGridView1.DataSource = FetchData()
        'DataGridView2.DataSource = "null"
        'DataGridView3.DataSource = "null"

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim issuccess As Boolean
        Dim thehold As String
        issuccess = WritetheRecord(TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text)
        If issuccess = True Then
            DataGridView1.DataSource = FetchData()
            DataGridView2.DataSource = FetchNames()
        End If
    End Sub


    Function WritetheRecord(ByVal id_num As String, ByVal Sport As String, ByVal activembr As String, ByVal StudyHall As String) As Boolean
        Dim widnum As String
        Dim wtype As String
        Dim wfld1 As String
        Dim wfld2 As String
        Dim wfld3 As String
        Dim flagconstant As String = " "

        widnum = id_num
        wtype = Sport
        wfld1 = activembr
        wfld2 = StudyHall



        Dim writesql As String
        writesql = "Insert into dbo.NC_ROSTERS (ID_NUM,SPORT,ACTIVE_MBR,STUDY_HALL_WATCH) values('" & Trim(widnum) & "','" & Trim(wtype) & "' ,'" & Trim(wfld1) & "' ,'" & Trim(wfld2) & "')"
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

    Function updatetheRecord(ByVal id_num As String, ByVal Sport As String, ByVal activembr As String, ByVal StudyHall As String) As Boolean
        Dim widNum As String
        Dim wtype As String
        Dim wfld1 As String
        Dim wfld2 As String
        Dim wfld3 As String
        Dim flagconstant As String = " "

        widnum = id_num
        wtype = Sport
        wfld1 = activembr
        wfld2 = StudyHall



        Dim updatesql As String
        updatesql = "update dbo.NC_Rosters "
        updatesql += "set Sport = '" & Trim(wtype) & "' ,"
        updatesql += " active_mbr = '" & Trim(wfld1) & "' ,"
        updatesql += " Study_Hall_watch = '" & Trim(wfld2) & "'"

        updatesql += " where  id_num = '" & Trim(widNum) & "'"
        updatesql += " and Sport = '" & Trim(wtype) & "'"


        Try

            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPly;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()







            Dim catCMD As OdbcCommand = New OdbcCommand(updatesql, oODBCConnection)



            Dim okcount As Integer = 0

            okcount = catCMD.ExecuteNonQuery
            oODBCConnection.Close()


            If okcount > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(" The  update did not occur. Please record the information manually and contact OIT and provide them with this message." & ex.Message)
        Finally

        End Try
        Return False
    End Function

    Function DeletetheRecord(ByVal id_num As String, ByVal Sport As String, ByVal activembr As String, ByVal StudyHall As String) As Boolean
         Dim widNum As String
        Dim wtype As String
        Dim wfld1 As String
        Dim wfld2 As String
        Dim wfld3 As String
        Dim flagconstant As String = " "

        widNum = id_num
        wtype = Sport
        wfld1 = activembr
        wfld2 = StudyHall


        Dim deletesql As String
        deletesql = "delete dbo.NC_ROSTERS "

        deletesql += "where  id_num = '" & Trim(widNum) & "'  and "
        deletesql += "  Sport = '" & Trim(wtype) & "'"
         


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


    Private Sub UpdateStars(ByVal row As DataGridViewRow, _
        ByVal stars As String)
        Dim ratingColumn As Integer
        row.Cells(ratingColumn).Value = stars

        ' Resize the column width to account for the new value.
        row.DataGridView.AutoResizeColumn(ratingColumn, _
            DataGridViewAutoSizeColumnMode.DisplayedCells)

    End Sub




    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim issuccess As Boolean
        issuccess = updatetheRecord(TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text)
        If issuccess = True Then
            DataGridView1.DataSource = FetchDataselective()
            DataGridView2.DataSource = FetchNames()
        End If





    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
       
        Dim issuccess As Boolean
        issuccess = DeletetheRecord(TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text)
        If issuccess = True Then
            DataGridView1.DataSource = FetchDataselective()
            DataGridView2.DataSource = FetchNames()
        End If



    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Label5.Text = GetStudentname(TextBox1.Text)
    End Sub
    Function GetStudentname(ByVal id_num As String)
        Try
            Dim strSQL As String
            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=TMSEPLY;   Uid=sa;   Pwd=jenzadmin"
            oODBCConnection = New System.Data.Odbc.OdbcConnection(sConnString)
            oODBCConnection.Open()




            strSQL = "Select  first_name + ' ' + last_name  from name_master where id_num='" & Trim(id_num) & "'"


            Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



            Dim thecount As Integer = 0

            Dim myReader As OdbcDataReader
            Dim thedescription As String

            myReader = catCMD.ExecuteReader
            While myReader.Read()
                thecount = thecount + 1
                If IsDBNull(myReader(0)) = True Then
                    thedescription = "No description is available"
                Else

                    thedescription = myReader(0)
                End If
            End While
            If thecount = 0 Then

                Return "No description is available"
            Else
                Return thedescription
            End If

        Catch ex As Exception
            MsgBox(" You can retry the entry or manually record the information. If the retry fails, please contact OIT and provide them with this message." & ex.Message)
        End Try
    End Function
End Class

