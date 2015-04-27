Imports System
Imports System.Data
Imports System.Data.Odbc
Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ok As Boolean
        ok = Processr1()
        ok = Processr2()
        ok = Processr3()
        ok = Processr4()
         
        Button1.Text = "Done"
    End Sub


    Function Processr1() As Boolean
        Dim strSQL As String
        Dim oODBCConnection As OdbcConnection
        Dim sConnString As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEprd;   Uid=sa;   Pwd=jenzadmin"
        oODBCConnection = New Odbc.OdbcConnection(sConnString)
        oODBCConnection.Open()
        strSQL = "Select id, 'student_r1'"

        strSQL += " from NC_Register_SELECT_CTL WHERE FULL_STATUS='Eligible' and ENABLEGROUP='RS1' and id not in (select id_num from TW_GRP_MEMBERSHIP where "

        strSQL += "GROUP_ID='student_r1')"
        Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)

        Dim strins As String
        Dim oODBCConnection2 As OdbcConnection
        Dim sConnString2 As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEprd;   Uid=sa;   Pwd=jenzadmin"
        oODBCConnection2 = New Odbc.OdbcConnection(sConnString2)
        Dim CMD2 As OdbcCommand = New OdbcCommand(strins, oODBCConnection2)
        ' put all disposable items in using() blocks - this applies to 
        ' StreamReader, SqlConnection, SqlCommand (and many more!)


        Dim myReader As OdbcDataReader = catCMD.ExecuteReader()



        Dim thecount As Integer = 0
        Dim inscount As Integer = 0
        Dim acn As String
        Dim agroup As String
        If myReader.HasRows Then

            oODBCConnection2.Open()

            Do While myReader.Read()

                thecount = thecount + 1


                Try

                    ' define connection string and stored procedure name 





                    ' define as stored procedure


                    '  




                    acn = myReader(0)
                    agroup = myReader(1)
                    strins = "INSERT INTO dbo.TW_GRP_MEMBERSHIP VALUES('" & acn & "','" & agroup & "','SYSTEM','RegisSELINIT',GetDate())"
                    CMD2.CommandType = CommandType.Text
                    CMD2.CommandText = strins
                    If acn <> "477993" Then

                        inscount = CMD2.ExecuteNonQuery()
                        ' execute procedure
                    End If

                Catch o As OdbcException

                    MsgBox(o.Message.ToString)

                Finally



                End Try

            Loop
            oODBCConnection.Close()
            oODBCConnection2.Close()


        End If
        Return True
    End Function
    Function Processr2() As Boolean
        Dim strSQL As String
        Dim oODBCConnection As OdbcConnection
        Dim sConnString As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEprd;   Uid=sa;   Pwd=jenzadmin"
        oODBCConnection = New Odbc.OdbcConnection(sConnString)
        oODBCConnection.Open()
        strSQL = "Select id, 'student_r2'"

        strSQL += " from NC_Register_SELECT_CTL WHERE FULL_STATUS='Eligible' and ENABLEGROUP='RS2' and id not in (select id_num from TW_GRP_MEMBERSHIP where "

        strSQL += "GROUP_ID='student_r2')"
        Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)

        Dim strins As String
        Dim oODBCConnection2 As OdbcConnection
        Dim sConnString2 As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEprd;   Uid=sa;   Pwd=jenzadmin"
        oODBCConnection2 = New Odbc.OdbcConnection(sConnString2)
        Dim CMD2 As OdbcCommand = New OdbcCommand(strins, oODBCConnection2)
        ' put all disposable items in using() blocks - this applies to 
        ' StreamReader, SqlConnection, SqlCommand (and many more!)


        Dim myReader As OdbcDataReader = catCMD.ExecuteReader()



        Dim thecount As Integer = 0
        Dim inscount As Integer = 0
        Dim acn As String
        Dim agroup As String
        If myReader.HasRows Then

            oODBCConnection2.Open()

            Do While myReader.Read()

                thecount = thecount + 1


                Try

                    ' define connection string and stored procedure name 





                    ' define as stored procedure


                    '  




                    acn = myReader(0)
                    agroup = myReader(1)
                    strins = "INSERT INTO dbo.TW_GRP_MEMBERSHIP VALUES('" & acn & "','" & agroup & "','SYSTEM','RegisSELINIT',GetDate())"
                    CMD2.CommandType = CommandType.Text
                    CMD2.CommandText = strins
                    If acn <> "477993" Then

                        inscount = CMD2.ExecuteNonQuery()
                        ' execute procedure
                    End If

                Catch o As OdbcException

                    MsgBox(o.Message.ToString)

                Finally



                End Try

            Loop
            oODBCConnection.Close()
            oODBCConnection2.Close()


        End If
        Return True
    End Function
    Function Processr3() As Boolean
        Dim strSQL As String
        Dim oODBCConnection As OdbcConnection
        Dim sConnString As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEprd;   Uid=sa;   Pwd=jenzadmin"
        oODBCConnection = New Odbc.OdbcConnection(sConnString)
        oODBCConnection.Open()
        strSQL = "Select id, 'student_r3'"

        strSQL += " from NC_Register_SELECT_CTL WHERE FULL_STATUS='Eligible' and ENABLEGROUP='RS3' and id not in (select id_num from TW_GRP_MEMBERSHIP where "

        strSQL += "GROUP_ID='student_r3')"
        Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)

        Dim strins As String
        Dim oODBCConnection2 As OdbcConnection
        Dim sConnString2 As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEprd;   Uid=sa;   Pwd=jenzadmin"
        oODBCConnection2 = New Odbc.OdbcConnection(sConnString2)
        Dim CMD2 As OdbcCommand = New OdbcCommand(strins, oODBCConnection2)
        ' put all disposable items in using() blocks - this applies to 
        ' StreamReader, SqlConnection, SqlCommand (and many more!)


        Dim myReader As OdbcDataReader = catCMD.ExecuteReader()



        Dim thecount As Integer = 0
        Dim inscount As Integer = 0
        Dim acn As String
        Dim agroup As String
        If myReader.HasRows Then

            oODBCConnection2.Open()

            Do While myReader.Read()

                thecount = thecount + 1


                Try

                    ' define connection string and stored procedure name 





                    ' define as stored procedure


                    '  




                    acn = myReader(0)
                    agroup = myReader(1)
                    strins = "INSERT INTO dbo.TW_GRP_MEMBERSHIP VALUES('" & acn & "','" & agroup & "','SYSTEM','RegisSELINIT',GetDate())"
                    CMD2.CommandType = CommandType.Text
                    CMD2.CommandText = strins
                    If acn <> "477993" Then

                        inscount = CMD2.ExecuteNonQuery()
                        ' execute procedure
                    End If

                Catch o As OdbcException

                    MsgBox(o.Message.ToString)

                Finally



                End Try

            Loop
            oODBCConnection.Close()
            oODBCConnection2.Close()


        End If
        Return True
    End Function
    Function Processr4() As Boolean
        Dim strSQL As String
        Dim oODBCConnection As OdbcConnection
        Dim sConnString As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEprd;   Uid=sa;   Pwd=jenzadmin"
        oODBCConnection = New Odbc.OdbcConnection(sConnString)
        oODBCConnection.Open()
        strSQL = "Select id, 'student_r4'"

        strSQL += " from NC_Register_SELECT_CTL WHERE FULL_STATUS='Eligible' and ENABLEGROUP='RS4' and id not in (select id_num from TW_GRP_MEMBERSHIP where "

        strSQL += "GROUP_ID='student_r4')"
        Dim catCMD As OdbcCommand = New OdbcCommand(strSQL, oODBCConnection)

        Dim strins As String
        Dim oODBCConnection2 As OdbcConnection
        Dim sConnString2 As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEprd;   Uid=sa;   Pwd=jenzadmin"
        oODBCConnection2 = New Odbc.OdbcConnection(sConnString2)
        Dim CMD2 As OdbcCommand = New OdbcCommand(strins, oODBCConnection2)
        ' put all disposable items in using() blocks - this applies to 
        ' StreamReader, SqlConnection, SqlCommand (and many more!)


        Dim myReader As OdbcDataReader = catCMD.ExecuteReader()



        Dim thecount As Integer = 0
        Dim inscount As Integer = 0
        Dim acn As String
        Dim agroup As String
        If myReader.HasRows Then

            oODBCConnection2.Open()

            Do While myReader.Read()

                thecount = thecount + 1


                Try

                    ' define connection string and stored procedure name 





                    ' define as stored procedure


                    '  




                    acn = myReader(0)
                    agroup = myReader(1)
                    strins = "INSERT INTO dbo.TW_GRP_MEMBERSHIP VALUES('" & acn & "','" & agroup & "','SYSTEM','RegisSELINIT',GetDate())"
                    CMD2.CommandType = CommandType.Text
                    CMD2.CommandText = strins
                    If acn <> "477993" Then

                        inscount = CMD2.ExecuteNonQuery()
                        ' execute procedure
                    End If

                Catch o As OdbcException

                    MsgBox(o.Message.ToString)

                Finally



                End Try

            Loop
            oODBCConnection.Close()
            oODBCConnection2.Close()


        End If
        Return True
    End Function

   
    

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class