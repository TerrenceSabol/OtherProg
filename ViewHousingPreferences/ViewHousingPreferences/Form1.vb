Imports System
Imports System.Data

Public Class Form1
    Public x As String
     
    Private Function FetchDataEligible() As DataTable


        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseprd;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand("SELECT distinct nm.last_name, nm.first_name, nm.id_num,     b.EnableGroup,b.EnableDate, b.EnableTime, ns.BLDG_CDE, ns.ROOM_CDE as AssignedRoom, ns.ROOM_SLOT_NUM, ns.ASSIGN_DTE,full_status as EligibleStatus, cast(b.EnableSeq as Int) as Controlseq , b.ID     FROM dbo.nc_housing_select_ctl b left join dbo.STUD_ROOM_PREFS a on a.id_num=b.id and a.sess_cde =(select ctlfld1 from NC_Control_Info where ctlname='ReturnSESS') left join name_master nm on b.id=nm.id_num left join room_assign ns on b.id=ns.id_num and ns.sess_cde=(select ctlfld1 from NC_Control_Info where ctlname='ReturnSESS') where Full_status='Eligible'   ORDER BY   cast(b.EnableSeq as Int) , b.ID ", Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)
            Try
                Dim Table As New DataTable
                sAdapter.Fill(Table)
                Return Table
            Catch o As Exception

                MsgBox(o.Message.ToString)
            End Try
        End Using

    End Function
    Private Function FetchDataAssigned() As DataTable


        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseprd;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand("SELECT distinct nm.last_name, nm.first_name, nm.id_num,      b.EnableGroup,b.EnableDate, b.EnableTime, ns.BLDG_CDE, ns.ROOM_CDE as AssignedRoom, ns.ROOM_SLOT_NUM, ns.ASSIGN_DTE,full_status as EligibleStatus, cast(b.EnableSeq as Int) as Controlseq , b.ID   FROM dbo.nc_housing_select_ctl b left join dbo.STUD_ROOM_PREFS a on a.id_num=b.id and a.sess_cde =(select ctlfld1 from NC_Control_Info where ctlname='ReturnSESS') left join name_master nm on b.id=nm.id_num left join room_assign ns on b.id=ns.id_num and ns.sess_cde=(select ctlfld1 from NC_Control_Info where ctlname='ReturnSESS') where  ns.room_cde is not null ORDER BY  last_name,first_name, cast(b.EnableSeq as Int) , b.ID ", Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)
            Try
                Dim Table As New DataTable
                sAdapter.Fill(Table)
                Return Table
            Catch o As Exception

                MsgBox(o.Message.ToString)
            End Try
        End Using

    End Function
    Private Function FetchDataUnAssigned() As DataTable


        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseprd;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand("SELECT distinct nm.last_name, nm.first_name, nm.id_num,     b.EnableGroup,b.EnableDate, b.EnableTime, ns.BLDG_CDE, ns.ROOM_CDE as AssignedRoom, ns.ROOM_SLOT_NUM, ns.ASSIGN_DTE,full_status as EligibleStatus,  cast(b.EnableSeq as Int) as Controlseq , b.ID   as Controlseq FROM dbo.nc_housing_select_ctl b left join dbo.STUD_ROOM_PREFS a on a.id_num=b.id and a.sess_cde =(select ctlfld1 from NC_Control_Info where ctlname='ReturnSESS') left join name_master nm on b.id=nm.id_num left join room_assign ns on b.id=ns.id_num and ns.sess_cde=(select ctlfld1 from NC_Control_Info where ctlname='ReturnSESS') where  ns.room_cde is null and full_status ='Eligible'  ORDER BY   cast(b.EnableSeq as Int) , b.ID ", Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)
            Try
                Dim Table As New DataTable
                sAdapter.Fill(Table)
                Return Table
            Catch o As Exception

                MsgBox(o.Message.ToString)
            End Try
        End Using

    End Function
    Private Function FetchData() As DataTable


        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseprd;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand("SELECT distinct nm.last_name, nm.first_name, nm.id_num,     b.EnableGroup,b.EnableDate, b.EnableTime, ns.BLDG_CDE, ns.ROOM_CDE as AssignedRoom, ns.ROOM_SLOT_NUM, ns.ASSIGN_DTE,full_status as EligibleStatus, cast(b.EnableSeq as Int) as Controlseq , b.ID FROM dbo.nc_housing_select_ctl b left join dbo.STUD_ROOM_PREFS a on a.id_num=b.id and a.sess_cde =(select ctlfld1 from NC_Control_Info where ctlname='ReturnSESS') left join name_master nm on b.id=nm.id_num left join room_assign ns on b.id=ns.id_num and ns.sess_cde=(select ctlfld1 from NC_Control_Info where ctlname='ReturnSESS')     ORDER BY   cast(b.EnableSeq as Int) , b.ID   ", Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)
            Try
                Dim Table As New DataTable
                sAdapter.Fill(Table)
                Return Table
            Catch o As Exception

                MsgBox(o.Message.ToString)
            End Try
        End Using

    End Function
    Private Function FetchRoommateInfo() As DataTable


        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseprd;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand("SELECT  case when Req_Actual_flag ='R' then 'Requested' when Req_Actual_flag ='A' Then 'Actual' else ' ' end as Request_Actual, nm.last_name + ',' + nm.first_name as Roommate_Name, Roommate_ID,Bldg_cde,Room_Cde,a.User_name as Created_UpdatedBy,a.Job_Name as Job,a.Job_Time   FROM dbo.Stud_roommates a left join name_master nm on a.roommate_id=nm.id_num   where a.sess_cde =(select ctlfld1 from NC_Control_Info where ctlname='ReturnSESS')  and a.id_num=" & x & "  ORDER BY   last_name,first_name,req_actual_flag", Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)
            Try
                Dim Table As New DataTable
                sAdapter.Fill(Table)
                Return Table
            Catch o As Exception

                MsgBox(o.Message.ToString)
            End Try
        End Using

    End Function
    Function FetchRoomPrefs() As DataTable
        Using Conn As New Data.SqlClient.SqlConnection("Data Source=NCSQL1;Initial Catalog=tmseprd;timeout=60; User Id=sa;Password=jenzadmin;"), Command As New Data.SqlClient.SqlCommand(" SELECT ROOM_PREF_BLDG, ROOM_PREF_ROOM, ROOM_PREF_PTY, ROOM_CDE, ROOM_SLOT_NUM, ROOM_ASSIGN_STS FROM dbo.STUD_ROOM_PREFS a left join room_assign b on a.SESS_CDE=b.sess_cde and ROOM_PREF_BLDG= BLDG_CDE and ROOM_PREF_ROOM=b.room_cde WHERE a.SESS_CDE=(Select ctlfld1 from NC_control_info where ctlname='Returnsess') and a.id_num=" & x & "  ORDER BY Room_assign_sts,room_PREF_PTY   ", Conn), sAdapter As New Data.SqlClient.SqlDataAdapter(Command)
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
        DataGridView1.DataSource = FetchDataEligible()
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        Dim cellroomstatus As String
        x = DataGridView1.CurrentRow.Cells(2).Value.ToString()

        DataGridView2.DataSource = FetchRoommateInfo()
        DataGridView3.DataSource = FetchRoomPrefs()
        If Me.DataGridView3.Rows.Count > 0 Then
            For i As Integer = 0 To Me.DataGridView3.Rows.Count - 1
                If IsDBNull(Me.DataGridView3.Rows(i).Cells(5).Value) = False Then
                    If Me.DataGridView3.Rows(i).Cells(5).Value = "A" Then
                        Me.DataGridView3.Rows(i).DefaultCellStyle.ForeColor = Color.Red
                    End If
                End If
            Next

        End If
    End Sub

     

     
      
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        DataGridView1.DataSource = FetchDataAssigned()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.DataSource = FetchDataUnAssigned()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        DataGridView1.DataSource = FetchData()



        For i As Integer = 0 To Me.DataGridView1.Rows.Count - 1
            If Me.DataGridView1.Rows(i).Cells(10).Value = "Not Yet Eligible" Then
                Me.DataGridView1.Rows(i).Cells(10).Style.ForeColor = Color.Red
            Else
                Me.DataGridView1.Rows(i).Cells(10).Style.ForeColor = Color.Green
            End If
        Next





    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class
