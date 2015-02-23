Imports System.Data.SqlClient
Imports Microsoft.Reporting.WinForms

Public Class frmSSRS
    Dim sConnStringREPORTSERVER As String = "Data Source=NCFUTURESQL; Database=REPORTSERVER; Uid=sa; Pwd=Jenzadmin!"


    Private Sub frmSSRS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height - 50 'taskbar at the bottom
        Me.Top = 1
        Me.Left = 1

        Me.ReportViewer1.Width = Me.Width - 30
        Me.ReportViewer1.Height = Me.Height - 100
        Me.ReportViewer1.Top = 50

        pboxLogo.Top = 1
        pboxLogo.Left = Me.Width - pboxLogo.Width - 30

        ReportViewer1.Visible = False
        Call Get_List_Of_Reports("ADMISSIONS")

        If Len(frmAdmiss.ssrs_reportname) > 0 Then
            Call GET_REPORT_NUMBER_AND_RUN(frmAdmiss.ssrs_reportname)
        End If

    End Sub
    Private Sub Get_List_Of_Reports(ByVal strMainCategory As String)
        Dim strSQL As String = "SELECT DISTINCT PATH, NAME FROM CATALOG "
        strSQL &= "WHERE Path LIKE '/~MAIN_CATEGORY~/%' "
        strSQL &= "AND Path NOT LIKE '%test%' AND Path NOT LIKE '%subreports%' "
        strSQL &= "ORDER BY NAME "

        strSQL = Replace(strSQL, "~MAIN_CATEGORY~", strMainCategory)
        Dim myConnectionRS As New SqlConnection(sConnStringREPORTSERVER)
        myConnectionRS.Open()

        Dim myCommand As New SqlCommand(strSQL, myConnectionRS)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        cboReports.Items.Clear()
        cboPath.Items.Clear()

        While reader.Read()
            cboReports.Items.Add(Trim(reader("NAME")))
            cboPath.Items.Add(Trim(reader("PATH")))
        End While

        reader.Close()
        myCommand.Dispose()
        myConnectionRS.Close()

        cboReports.Text = cboReports.Items.Count & " REPORTS"
    End Sub
    Public Sub GET_REPORT_NUMBER_AND_RUN(ByVal strReportName As String)
        Dim kount As Int16 = 0
        Dim bolReportFound As Boolean = False

        For kount = 0 To cboReports.Items.Count - 1
            If cboReports.Items(kount) = strReportName Then
                Call RUN_REPORT(kount)
                bolReportFound = True
                Exit For
            End If
        Next

        If Not bolReportFound Then
            MsgBox("Cannot find: " & strReportName)
            Stop
        End If


    End Sub


    Private Sub cboReports_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboReports.SelectedIndexChanged

        Call RUN_REPORT(cboReports.SelectedIndex)

        '   Me.Cursor = Cursors.WaitCursor
        '  cboReports.Enabled = False
        ' cboReports.Refresh()

        '   ReportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Remote

        ' Dim serverReport As Microsoft.Reporting.WinForms.ServerReport
        ' serverReport = ReportViewer1.ServerReport

        '       'Get a reference to the default credentials
        '        Dim credentials As System.Net.ICredentials
        '        credentials = System.Net.CredentialCache.DefaultCredentials

        '      'Get a reference to the report server credentials
        '      Dim rsCredentials As Microsoft.Reporting.WinForms.ReportServerCredentials
        '      rsCredentials = serverReport.ReportServerCredentials
        '
        '        'Set the credentials for the server report
        '        rsCredentials.NetworkCredentials = credentials
        '
        '        'Set the report server URL and report path
        '        serverReport.ReportServerUrl = New Uri("http://ncfuturesql/reportserver")
        '        serverReport.ReportPath = cboPath.Items(cboReports.SelectedIndex)
        '
        '        ReportViewer1.Visible = True
        '        ReportViewer1.RefreshReport()
        '
        '
        '        Me.Cursor = Cursors.Default
        '
        '       cboReports.Enabled = True


    End Sub
    Public Sub RUN_REPORT(ByVal intReportNumber As String)

        Me.Cursor = Cursors.WaitCursor
        cboReports.Enabled = False
        cboReports.Refresh()

        ReportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Remote

        Dim serverReport As Microsoft.Reporting.WinForms.ServerReport
        serverReport = ReportViewer1.ServerReport

        'Get a reference to the default credentials
        Dim credentials As System.Net.ICredentials
        credentials = System.Net.CredentialCache.DefaultCredentials

        'Get a reference to the report server credentials
        Dim rsCredentials As Microsoft.Reporting.WinForms.ReportServerCredentials
        rsCredentials = serverReport.ReportServerCredentials

        'Set the credentials for the server report
        rsCredentials.NetworkCredentials = credentials

        'Set the report server URL and report path
        serverReport.ReportServerUrl = New Uri("http://ncfuturesql/reportserver")
        serverReport.ReportPath = cboPath.Items(intReportNumber)

        If Len(frmAdmiss.ssrs_parameter_1) > 0 Then
            Dim ssrs_prmUserName As New ReportParameter()
            ssrs_prmUserName.Name = "prmUserName"
            ssrs_prmUserName.Values.Add(Environment.UserName)

            Dim ssrs_prmAsOf As New ReportParameter()
            ssrs_prmAsOf.Name = frmAdmiss.ssrs_parameter_name
            ssrs_prmAsOf.Values.Add(frmAdmiss.ssrs_parameter_1)


            Dim ssrs_RegionYear As New ReportParameter()
            ssrs_RegionYear.Name = "prmBreakdownByRegionYear"
            ssrs_RegionYear.Values.Add(frmAdmiss.cboTermEnd.Text)

            'Set the report parameters for the report
            Dim parameters() As ReportParameter = {ssrs_prmAsOf, ssrs_prmUserName, ssrs_RegionYear}
            serverReport.SetParameters(parameters)
        End If


        'Refresh the report
        ReportViewer1.Visible = True
        ReportViewer1.RefreshReport()

        txtprmAsOf.Text = ""

        Me.Cursor = Cursors.Default

        cboReports.Enabled = True


    End Sub

    ' Public Sub SET_PARAMETERS()
    'Dim Param1 As New Microsoft.Reporting.WinForms.ReportParameter()
    '    Param1.Name = "prmAsOf" '*** The actual name of the parameter in the report ***
    '   Param1.Values.Add(prmAsOf.Text)
    '  ReportViewer1.ServerReport.SetParameters(New Microsoft.Reporting.WinForms.ReportParameter() {Param1})
    'End Sub



    Private Sub btnReturn_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub RETURNTOHOMEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RETURNTOHOMEToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub RETURNHOMEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RETURNHOMEToolStripMenuItem.Click
        Me.Close()
    End Sub
End Class