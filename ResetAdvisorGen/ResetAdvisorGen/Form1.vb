Imports System.Data.SqlClient
Imports System.Data.Odbc

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Text = " "
        Label3.Text = " "
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ok As Boolean
        Label2.Text = " "
        Label3.Text = " "
        Dim validateresult As String
        validateresult = Validateentry(TextBox1.Text)
        Label3.Text = validateresult
        If validateresult.Substring(0, 2) = "Va" Then
            ok = updateadvisorinfo(TextBox1.Text)
            If ok = True Then
                Label2.Text = "Advisor Clear was successful"
            Else
                Label2.Text = " AdvisorClear Failed"
            End If
        Else
        End If

    End Sub
    Public Function updateadvisorinfo(id_num As String) As Boolean
        Dim strSQL As String
        Dim soSQLConnection As SqlConnection
        Dim sConnString As String = "Server=ncsql1;Database=Tmseprd;Uid=sa;pwd=jenzadmin;"

        soSQLConnection = New SqlClient.SqlConnection(sConnString)
        soSQLConnection.Open()


        Dim cmdx As New SqlCommand("[dbo].[NC_CLEAR_BUILD_ADVISOR]", soSQLConnection)
        cmdx.CommandType = CommandType.StoredProcedure
        cmdx.Parameters.AddWithValue("@id_num", id_num)
         



        Try
            cmdx.ExecuteNonQuery()
            soSQLConnection.Close()
            Return True

        Catch o As Exception

            MsgBox(" Clear Failed   for " & id_num)
            soSQLConnection.Close()
            Return False
        End Try




    End Function
    Function Validateentry(id_num As String) As String
        Dim strSQL2 As String
        Dim oODBCConnection2 As OdbcConnection
        Dim sConnString As String = "Driver={SQL Server};   Server=NCSQL1;   Database=TmsEprd;   Uid=sa;   Pwd=jenzadmin"
        oODBCConnection2 = New Odbc.OdbcConnection(sConnString)
        oODBCConnection2.Open()
        strSQL2 = "Select first_name + '  ' + last_name,c.id_num from name_master nm left join candidacy c on nm.id_Num=c.id_num  and c.yr_cde =(select ctlfld1 from NC_CONTROL_INFO where ctlname='INCOMINGYR') and "
        strSQL2 += " c.trm_cde =(select ctlfld1 from NC_CONTROL_INFO where ctlname='INCOMINGTRM') where nm.id_num ='" & id_num & "'"



        Dim catCMD2 As OdbcCommand = New OdbcCommand(strSQL2, oODBCConnection2)



        ' put all disposable items in using() blocks - this applies to 
        ' StreamReader, SqlConnection, SqlCommand (and many more!)


        Dim myReader As OdbcDataReader = catCMD2.ExecuteReader()



        Dim thecount As Integer = 0


        If myReader.HasRows Then



            Do While myReader.Read()

                thecount = thecount + 1


                Try
                    If IsDBNull(myReader(0)) = True Then
                        Return "Invalid--- No Name in Name Master"
                    End If

                    If IsDBNull(myReader(0)) = False And IsDBNull(myReader(1)) = True Then
                        Return "Invalid No Candidacy Record --- " & CStr(myReader(0))
                    Else
                        Return "Valid " & CStr(myReader(0))
                    End If







                Catch o As OdbcException

                    MsgBox(o.Message.ToString)

                Finally



                End Try
                If thecount = 0 Then
                    Return "Name not found for Student ID"
                End If
            Loop
            oODBCConnection2.Close()


        End If









    End Function
End Class
