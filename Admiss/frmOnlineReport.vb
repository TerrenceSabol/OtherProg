Public Class frmOnlineReport

    Private Sub frmOnlineReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Initialize()
    End Sub

    Private Sub Initialize()
        Dim kountRows As Int16 = 0
        Dim KountCols As Int16 = 0

        dgvReport.ColumnCount = frmAdmiss.tblOnlineReportColumns + 1

        If dgvReport.Rows.Count > frmAdmiss.tblOnlineReportRows Then
            While dgvReport.Rows.Count > frmAdmiss.tblOnlineReportRows
                dgvReport.Rows.RemoveAt(dgvReport.Rows.Count)
            End While
        Else
            dgvReport.Rows.Add(frmAdmiss.tblOnlineReportRows)
        End If


        For KountCols = 1 To frmAdmiss.tblOnlineReportColumns
            dgvReport.Columns(KountCols).Name = frmAdmiss.tblOnlineReportHeaders(KountCols)
        Next

        For kountRows = 0 To frmAdmiss.tblOnlineReportRows - 1
            For KountCols = 0 To frmAdmiss.tblOnlineReportColumns
                dgvReport.Rows(kountRows).Cells(KountCols).Value = frmAdmiss.tblOnlineReport(1, kountRows, KountCols)
            Next
        Next

        dgvReport.RowHeadersVisible = False

        For KountCols = 0 To dgvReport.Columns.Count
            dgvReport.AutoResizeColumn(KountCols)
        Next


        '   dgvReport.Columns(0).Width = 

    End Sub
End Class