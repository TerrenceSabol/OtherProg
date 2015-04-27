Imports System
Imports System.Data
Imports System.Data.Odbc


Public Class Form1
    Public x As String
    Public nameinfo As String

    Private Function FetchData() As DataTable





        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseprd;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand("SELECT distinct * from NC_control_info_def     order by ctlname     ", Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)


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


        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseprd;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand("SELECT distinct * from NC_control_info_def  where ctlname='" & nameinfo & "'   order by ctlname       ", Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)


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





        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseprd;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand("SELECT distinct ctlname from NC_control_info_def     order by ctlname     ", Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)


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
        RichTextBox1.Text = DataGridView1.CurrentRow.Cells(1).Value.ToString()

        nameinfo = DataGridView1.CurrentRow.Cells(0).Value.ToString()


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

        DataGridView1.DataSource = "null"
        DataGridView1.DataSource = FetchDataselective()
        TextBox1.Text = " "

    End Sub





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        DataGridView1.DataSource = FetchData()
        'DataGridView2.DataSource = "null"
        'DataGridView3.DataSource = "null"

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim issuccess As Boolean
        issuccess = WritetheRecord(TextBox1.Text, RichTextBox1.Text)
        If issuccess = True Then
            DataGridView1.DataSource = FetchData()
            DataGridView2.DataSource = FetchNames()
        End If
    End Sub


    Function WritetheRecord(ByVal ctlname As String, ByVal ctldef As String) As Boolean
        Dim wname As String
        Dim wdef As String
        Dim wfld1 As String
        Dim wfld2 As String
        Dim wfld3 As String
        Dim flagconstant As String = " "

        wname = ctlname
        wdef = ctldef



        Dim writesql As String
        writesql = "Insert into dbo.NC_CONTROL_INFO_def (ctlname,ctlname_desc  ) values('" & Trim(wname) & "','" & Trim(wdef) & "')"
        Try

            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=tmseprd;   Uid=sa;   Pwd=jenzadmin"
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


    Function UpdatetheRecord(ByVal ctlname As String, ByVal ctldef As String) As Boolean
        Dim wname As String
        Dim wdef As String

        Dim flagconstant As String = " "

        wname = ctlname
        wdef = ctldef


        Dim updatesql As String
        updatesql = "update dbo.NC_CONTROL_INFO_DEF "
        updatesql += "set ctlname_desc = '" & Trim(wdef) & "'"

        updatesql += " where  ctlname = '" & Trim(wname) & "'"


        Try

            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=tmseprd;   Uid=sa;   Pwd=jenzadmin"
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

    Function DeletetheRecord(ByVal ctlname As String, ByVal ctldef As String) As Boolean
        Dim wname As String
        Dim wdef As String

        Dim flagconstant As String = " "

        wname = ctlname
        wdef = ctldef


        Dim deletesql As String
        deletesql = "delete dbo.NC_CONTROL_INFO_DEF "

        deletesql += "where  ctlname = '" & Trim(wname) & "'  and "
        deletesql += "  ctlname_desc = '" & Trim(wdef) & "'"


        Try

            Dim oODBCConnection As OdbcConnection
            'Dim sConnString As String = "Driver={SQL Server};   Server=NCfutureSQL;   Database=FUTURE01_ICS_NET;   Uid=sa;   Pwd=Jenzadmin!"
            Dim sConnString As String = "Driver={SQL Server};   Server=NCsql1;   Database=tmseprd;   Uid=sa;   Pwd=jenzadmin"
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
        issuccess = UpdatetheRecord(TextBox1.Text, RichTextBox1.Text)
        If issuccess = True Then
            DataGridView1.DataSource = FetchDataselective()
            DataGridView2.DataSource = FetchNames()
        End If





    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
       
        Dim issuccess As Boolean
        issuccess = DeletetheRecord(TextBox1.Text, RichTextBox1.Text)
        If issuccess = True Then
            DataGridView1.DataSource = FetchDataselective()
            DataGridView2.DataSource = FetchNames()
        End If



    End Sub
End Class

