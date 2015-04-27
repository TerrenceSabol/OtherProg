Imports System
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.SqlClient

Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ok As Boolean
        ok = Processh1()

        Button1.Text = "Done"
        Button1.Enabled = "False"

    End Sub


    Function Processh1() As Boolean


        Dim strSQL2 As String
        Dim oODBCConnection2 As OdbcConnection
        Dim sConnString As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEprd;   Uid=sa;   Pwd=jenzadmin"
        oODBCConnection2 = New Odbc.OdbcConnection(sConnString)
        oODBCConnection2.Open()
        strSQL2 = "Select advisorctl,"
        strSQL2 += "studentctl,"
        strSQL2 += "majorcde,"
        strSQL2 += "adv1seq"

        strSQL2 += " from dbo.NC_Generate_advisor_ctl    where    "

        strSQL2 += "  (divmastadv is Null or advstud is null or advmast is null or majorcde is null) and studentctl ='" & TextBox1.Text & "'"

        Dim catCMD2 As OdbcCommand = New OdbcCommand(strSQL2, oODBCConnection2)



        ' put all disposable items in using() blocks - this applies to 
        ' StreamReader, SqlConnection, SqlCommand (and many more!)


        Dim myReader As OdbcDataReader = catCMD2.ExecuteReader()



        Dim thecount As Integer = 0
        Dim inscount As Integer = 0
        Dim stradvisorctl As String
        Dim strstudentctl As String
        Dim strmajorcde As String
        Dim adv1seq As Integer

        If myReader.HasRows Then



            Do While myReader.Read()

                thecount = thecount + 1


                Try

                    ' define connection string and stored procedure name 

                    stradvisorctl = myReader(0)
                    strstudentctl = myReader(1)
                    strmajorcde = myReader(2)
                    adv1seq = myReader(3)



                    Dim ok As Boolean
                    ' 

                    ok = updateadvisor(stradvisorctl, strstudentctl, strmajorcde, adv1seq)

                    ' define as stored procedure

                    If ok = False Then
                        Dim errorlogok As Boolean
                        errorlogok = updateerroradvgen(stradvisorctl, strstudentctl, strmajorcde, adv1seq)
                    Else
                        Dim logok = updategoodadvgen(stradvisorctl, strstudentctl, strmajorcde, adv1seq)
                    End If




                Catch o As OdbcException

                    MsgBox(o.Message.ToString)

                Finally



                End Try

            Loop
            oODBCConnection2.Close()


        End If






        Label1.Text = "Advisor  was successfully generated for '" & TextBox1.Text & "'"

        Return True


    End Function






    Public Function updateadvisor(stradvisorctl As String, strstudentctl As String, strmajorcde As String, adv1seq As Integer) As Boolean
        Dim strSQL As String
        Dim soSQLConnection As SqlConnection
        Dim sConnString As String = "Server=ncsql1;Database=Tmseprd;Uid=sa;pwd=jenzadmin;"

        soSQLConnection = New SqlClient.SqlConnection(sConnString)
        soSQLConnection.Open()


        Dim cmdx As New SqlCommand("[dbo].[NC_BUILD_ADVISOR]", soSQLConnection)
        cmdx.CommandType = CommandType.StoredProcedure
        cmdx.Parameters.AddWithValue("@id_num", strstudentctl)
        cmdx.Parameters.AddWithValue("@advisor_id", stradvisorctl)
        cmdx.Parameters.AddWithValue("@advisor_seq", adv1seq)
        cmdx.Parameters.AddWithValue("@degree_code", strmajorcde)


        Try
            cmdx.ExecuteNonQuery()
            soSQLConnection.Close()
            Return True

        Catch o As OdbcException

            MsgBox(o.Message.ToString)
            soSQLConnection.Close()
            Return False
        End Try




    End Function
    Public Function updateerroradvgen(stradvisorctl As String, strstudentctl As String, strmajorcde As String, adv1seq As Integer) As Boolean
        Dim strSQL As String
        Dim soSQLConnection As SqlConnection
        Dim sConnString As String = "Server=ncsql1;Database=Tmseprd;Uid=sa;pwd=jenzadmin;"

        soSQLConnection = New SqlClient.SqlConnection(sConnString)
        soSQLConnection.Open()


        Dim cmdx As New SqlCommand("[dbo].[NC_BUILD_ADVISOR_ERROR]", soSQLConnection)
        cmdx.CommandType = CommandType.StoredProcedure
        cmdx.Parameters.AddWithValue("@id_num", strstudentctl)
        cmdx.Parameters.AddWithValue("@advisor_id", stradvisorctl)
        cmdx.Parameters.AddWithValue("@advisor_seq", CStr(adv1seq))
        cmdx.Parameters.AddWithValue("@degree_code", strmajorcde)


        Try
            cmdx.ExecuteNonQuery()
            soSQLConnection.Close()
            Return True

        Catch o As OdbcException

            MsgBox(MsgBox("error log  failed for" & stradvisorctl & " -  " & strstudentctl & "audit trail failed for" & strmajorcde & " -  " & CStr(adv1seq)))
            soSQLConnection.Close()
            Return False
        End Try




    End Function
    Public Function updategoodadvgen(stradvisorctl As String, strstudentctl As String, strmajorcde As String, adv1seq As Integer) As Boolean
        Dim strSQL As String
        Dim soSQLConnection As SqlConnection
        Dim sConnString As String = "Server=ncsql1;Database=Tmseprd;Uid=sa;pwd=jenzadmin;"

        soSQLConnection = New SqlClient.SqlConnection(sConnString)
        soSQLConnection.Open()


        Dim cmdx As New SqlCommand("[dbo].[NC_BUILD_ADVISOR_GOOD]", soSQLConnection)
        cmdx.CommandType = CommandType.StoredProcedure
        cmdx.Parameters.AddWithValue("@id_num", strstudentctl)
        cmdx.Parameters.AddWithValue("@advisor_id", stradvisorctl)
        cmdx.Parameters.AddWithValue("@advisor_seq", CStr(adv1seq))
        cmdx.Parameters.AddWithValue("@degree_code", strmajorcde)


        Try
            cmdx.ExecuteNonQuery()
            soSQLConnection.Close()
            Return True

        Catch o As OdbcException

            MsgBox("audit trail failed for" & stradvisorctl & " -  " & strstudentctl & "audit trail failed for" & strmajorcde & " -  " & CStr(adv1seq))
            soSQLConnection.Close()
            Return False
        End Try




    End Function


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Dim strSQL As String
        'Dim oODBCConnection As OdbcConnection
        'Dim sConnString As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEprd;   Uid=sa;   Pwd=jenzadmin"
        'oODBCConnection = New Odbc.OdbcConnection(sConnString)
        'oODBCConnection.Open()
        'strSQL = "Select cast(a.id_num as varchar(10)) +'...' + last_name  + ',' + first_name    from dbo.candidacy    a left join name_master nm on a.id_num=nm.id_num where  a.YR_CDE=(Select CTLFLD1 from dbo.NC_CONTROL_INFO where CTLNAME='ORIYR') and a.TRM_CDE=(Select CTLFLD1 from dbo.NC_CONTROL_INFO where CTLNAME='ORITRM') and a.candidacy_type='R' and a.cur_candidacy='Y' and a.stage in ('22','23') order by last_name "


        'Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)



        ' put all disposable items in using() blocks - this applies to 
        ' StreamReader, SqlConnection, SqlCommand (and many more!)


        'Dim myReader As OdbcDataReader = catCMD.ExecuteReader()



        'Dim thecount As Integer = 0
        'Dim inscount As Integer = 0
        'Dim acn As String

        'Dim agroup As String

        'If myReader.HasRows Then



        'Do While myReader.Read()

        'thecount = thecount + 1


        'Try

        ' define connection string and stored procedure name 

        'ListBox1.Items.Add(myReader(0))



        ' define as stored procedure





        'Catch o As OdbcException

        'MsgBox(o.Message.ToString)

        'Finally



        'End Try

        'Loop
        'oODBCConnection.Close()


        'End If


    End Sub

    '

    '

End Class

