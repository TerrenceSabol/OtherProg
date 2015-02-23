Imports System.Data.SqlClient

Public Class frmStages


    Private Sub frmStages_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call Load_Table()
    End Sub

    Private Sub Load_Table()

        Dim strSQL As String = "SELECT RTRIM(Stage) AS STAGE, RTRIM(Category_Main) AS CATEGORY_MAIN, "
        strSQL &= "RTRIM(Category_Detail) AS CATEGORY_DETAIL, RTRIM(BUCKET) AS BUCKET FROM NC_Admissions_Stage_Category "
        strSQL &= "ORDER BY CATEGORY_DETAIL"
        Dim myConnectionEX As New SqlConnection(frmAdmiss.sConnString)
        Dim dataadapter As New SqlDataAdapter(strSQL, myConnectionEX)
        Dim ds As New DataSet()
        myConnectionEX.Open()
        dataadapter.Fill(ds, "Stages")
        myConnectionEX.Close()
        dgvStages.DataSource = ds
        dgvStages.DataMember = "Stages"
        dgvStages.AllowUserToAddRows = False

        dgvStages.Columns(0).Width = 50
        dgvStages.Columns(1).Width = 120
        dgvStages.Columns(2).Width = 180
        dgvStages.Columns(3).Width = 120
        dgvStages.Width = 490
        Me.Width = dgvStages.Width + 40
        dgvStages.Height = Me.Height - 50

        dgvStages.RowHeadersVisible = False

        ds.Dispose()


    End Sub
End Class