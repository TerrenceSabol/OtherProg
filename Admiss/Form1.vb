Imports System.Data.SqlClient

Public Class frmAdmiss

    Public Const cstStageType As String = "BUCKET" ' CHOICES ARE CATEGORY_MAIN & CATEGORY_DETAIL (FROM NC_ADMISSIONS_STAGE_CATEGORY)
    ' Public Const cstStageType As String = "CATEGORY_DETAIL" ' CHOICES ARE CATEGORY_MAIN & CATEGORY_DETAIL (FROM NC_ADMISSIONS_STAGE_CATEGORY)


    Dim CommandLineArgs As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs
    Public sConnString As String = "Data Source=NCSQL1;   Database=TmsEPrd;   Uid=sa;   Pwd=jenzadmin"


    Dim conTempTable As New System.Data.SqlClient.SqlConnection(Replace(sConnString, "TmsEPrd", "TmsEPly"))

    Const strReportOutputDirectory As String = "C:\NC_Report"
    Public strOutputFileName As String = strReportOutputDirectory & "\Admiss.csv"
    Public bolStopRun As Boolean = False
    Public bolReportStarted As Boolean = False
    Public aryStages(50) As String
    Public aryStagesCount As Int16
    Public aryStageCode(100) As String  'Stage Desc / Stage Number
    Public inxStageCode As Int16 = 0
    Public aryRegions(1000) As RECORD_REGIONS
    Public aryRegionsIndex As Int16 = 0
    Public aryTypes(5) As String

    Public aryBigBucketsNAME(10) As String
    Public aryBigBuckets(10) As Int16
    Public inxBigBuckets As Int16 = 0
    'PROSPECT/APPLIED/ADMIT/ENROLLED  -  CURRENT BUCKETS

    Public tblOnlineReport(20, 20, 20) As String
    Public tblOnlineReportColumns As Int16 = 0
    Public tblOnlineReportRows As Int16 = 0
    Public tblOnlineReportHeaders(20) As String

    Public bolRunningSummary As Boolean = False

    Public myConnectionGET_STUDENT_ADDRESS As New SqlConnection(sConnString)
    Public myConnectionGET_HIGH_SCHOOL_ADDRESS As New SqlConnection(sConnString)
    Public myConnectionGET_STAGE As New SqlConnection(sConnString)

    Public ssrs_parameter_1 As String = ""
    Public ssrs_parameter_name As String = ""
    Public ssrs_reportname As String = ""

    Public aryCounselorInitials(20) As String
    Public aryHighSchools(5000) As String

    Structure RECORD_DATA
        Dim ID_NUM As String
        Dim LAST_NAME As String
        Dim FIRST_NAME As String
        Dim MIDDLE_NAME As String
        Dim PREFERRED_NAME As String
        Dim YEAR As String
        Dim TERM As String
        Dim GENDER As String
        Dim BIRTH_DATE As String
        Dim REGISTRAR_ENROLLED As String
        Dim PROG_CODE As String
        Dim CANDIDACY_TYPE As String
        Dim CURRENT_STAGE As String
        Dim HOME_PHONE As String
        Dim CELL_PHONE As String
        Dim EMAIL As String
        Dim COUNSELOR As String
        Dim ATHLETIC_PROSPECT As String
        Dim HIGH_SCHOOL As String
        Dim LAST_ORGANIZATION As String

        Dim PROSPECT_ADDRESS_1 As String
        Dim PROSPECT_ADDRESS_2 As String
        Dim PROSPECT_CITY As String
        Dim PROSPECT_STATE As String
        Dim PROSPECT_ZIP As String
        Dim PROSPECT_COUNTRY As String

        Dim LAST_ORGANIZATION_ADDRESS_1 As String
        Dim LAST_ORGANIZATION_ADDRESS_2 As String
        Dim LAST_ORGANIZATION_CITY As String
        Dim LAST_ORGANIZATION_STATE As String
        Dim LAST_ORGANIZATION_ZIP As String
        Dim LAST_ORGANIZATION_COUNTRY As String
        Dim LAST_ORGANIZATION_REGION As String

        Dim PROSPECT_BUCKET As String
        Dim APPLIED_BUCKET As String
        Dim DENY_BUCKET As String
        Dim ADMITTED_BUCKET As String
        Dim WITHDRAWN_BUCKET As String
        Dim ENROLLED_BUCKET As String
        Dim DEPOSIT_BUCKET As String
    End Structure

    Public recData As RECORD_DATA

    Structure RECORD_REGIONS
        Dim ZIP_CODE As String
        Dim REGION As String
    End Structure

    Private Sub frmAdmiss_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = Chr(13) Then btnRun.PerformClick()
    End Sub
    Private Sub frmAdmiss_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call Initialize()
    End Sub
    Private Sub Initialize()

        If CAN_USER_RUN_PROGRAM() Then
            'OK TO RUN PROGRAM
        Else
            Application.Exit()
        End If

        chkDataItems.Items.Add("STUDENT ADDRESS")
        chkDataItems.Items.Add("LAST ORG ADDRESS")
        chkDataItems.Items.Add("STAGES")

        aryTypes(0) = "T O T A L"
        aryTypes(1) = "FRESHMAN"
        aryTypes(2) = "TRANSFER"
        aryTypes(3) = "RE-ADMIT"
        aryTypes(4) = "INTERNATIONAL"
        aryTypes(5) = "SPECIAL"

        chkType.Items.Add("FRESHMAN", True)
        chkType.Items.Add("TRANSFER", True)
        chkType.Items.Add("RE-ADMIT", True)
        chkType.Items.Add("INTERNATIONAL", True)
        chkType.Items.Add("SPECIAL", True)

        Call Fill_in_Terms()
        Call Fill_in_COUNSELORS()
        Call Fill_in_HIGH_SCHOOLS()
        Call Load_Stages()
        Call Initialize_Build_Stage_Array()
        Call Initialize_Build_Regions()
        Call Initialize_Big_Buckets()
        lblPercentDone.Visible = False

        If Not System.IO.Directory.Exists(strReportOutputDirectory) Then
            System.IO.Directory.CreateDirectory(strReportOutputDirectory)
        End If

        cboSummary.Items.Add("STAGES")

        rdoFall.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Bold)
        rdoFall.Checked = True
        dteTargetDate.Text = Today
        btnRun.Focus()


    End Sub
    Private Sub SSRS_Report()
        Dim tmpForm As Form = frmSSRS

        tmpForm.ShowDialog()

    End Sub
    Private Sub Initialize_Big_Buckets()
        Dim intBucketNumber As Int16 = 0
        Dim strSQL As String = "SELECT DISTINCT BUCKET FROM NC_ADMISSIONS_STAGE_CATEGORY ORDER BY BUCKET "

        Dim myConnectionEX As New SqlConnection(sConnString)
        myConnectionEX.Open()
        Dim myCommand As New SqlCommand(strSQL, myConnectionEX)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        While reader.Read()
            intBucketNumber = Val(Mid(reader("BUCKET"), 1, 2))  'BUCKET IS IN DB IN THE FORMAT ##_<BUCKETNAME> SUCH AS 01_PROSPECT
            aryBigBucketsNAME(intBucketNumber) = Trim(Mid(reader("BUCKET"), 4)) 'BUCKET 00 WILL BE THE BUCKET NAMED: OTHER
            aryBigBuckets(intBucketNumber) = 0
        End While

        reader.Close()
        myCommand.Dispose()
        myConnectionEX.Close()

        inxBigBuckets = intBucketNumber 'SAVE THE LAST BUCKETNUMBER

    End Sub
    Private Function CAN_USER_RUN_PROGRAM() As Boolean
        Dim strReturn As Boolean = False
        Dim strUserName As String = Trim(Environment.UserName)
        Dim strPermissionLevel As String = "0"

        Dim User_OK As Boolean = False
        Dim User_Not_OK_But_In_DB As Boolean = False
        Dim User_Not_In_DB As Boolean = False
        Dim UPDATE_LAST_USED As Boolean = False

        Dim strSQL As String = "SELECT ID_NUM FROM NC_REPORT_PERMISSION WHERE PROGRAM_NAME = 'ADMISS' AND NETWORK_NAME = '~USERNAME~' "
        strSQL = Replace(strSQL, "~USERNAME~", strUserName)
        Dim myConnectionEX As New SqlConnection(sConnString)
        myConnectionEX.Open()

        Dim myCommand As New SqlCommand(strSQL, myConnectionEX)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        If reader.HasRows Then

            UPDATE_LAST_USED = True

            reader.Read()
            If reader("ID_NUM") > 0 Then
                User_OK = True
                strReturn = True
            Else
                User_Not_OK_But_In_DB = True
                strReturn = False
            End If
        Else 'user not in db
            UPDATE_LAST_USED = False
            strReturn = False
            User_Not_In_DB = True
        End If

        reader.Close()


        If UPDATE_LAST_USED Then 'UPDATE LAST_USED
            Try
                Dim myUpdateCommand As SqlCommand = myConnectionEX.CreateCommand()
                '   myConnectionEX.Open()
                strSQL = "UPDATE NC_REPORT_PERMISSION SET LAST_USED = '~NOW~' WHERE PROGRAM_NAME = 'ADMISS' AND NETWORK_NAME = '~USERNAME~' "
                strSQL = Replace(strSQL, "~USERNAME~", strUserName)
                strSQL = Replace(strSQL, "~NOW~", Now)
                myCommand.CommandText = strSQL
                myCommand.ExecuteNonQuery()

            Catch ex As SqlException
                reader.Close()
                MsgBox(ex.Message)
            End Try
        End If

        If User_Not_In_DB Then 'ADD RECORD OF USER TRYING TO RUN PROGRAM WITHOUT HAVING A PERMISSION RECORD
            strReturn = False
            Try
                Dim myUpdateCommand As SqlCommand = myConnectionEX.CreateCommand()
                '   myConnectionEX.Open()
                strSQL = "INSERT INTO NC_REPORT_PERMISSION (ID_NUM, PROGRAM_NAME, NETWORK_NAME, PERMISSION_LEVEL, LAST_USED) VALUES (0,'ADMISS','~USERNAME~',0, '~NOW~')  "
                strSQL = Replace(strSQL, "~USERNAME~", strUserName)
                strSQL = Replace(strSQL, "~NOW~", Now)
                myCommand.CommandText = strSQL
                myCommand.ExecuteNonQuery()

            Catch ex As SqlException
                reader.Close()
                MsgBox(ex.Message)
            End Try

            MsgBox(strUserName & " DOES NOT HAVE PERMISSIONS TO RUN THIS PROGRAM", MsgBoxStyle.Critical, "MISSING RECORD IN NC_REPORT_PERMISSION")
        End If

        If User_Not_OK_But_In_DB Then 'USER ALREADY HAS TRIED TO RUN PROGRAM WITHOUT PERMISSION
            MsgBox(strUserName & " DOES NOT HAVE PERMISSIONS TO RUN THIS PROGRAM", MsgBoxStyle.Critical, "MISSING RECORD IN NC_REPORT_PERMISSION")
        End If

        reader.Close()
        myCommand.Dispose()
        myConnectionEX.Close()

        Return strReturn
    End Function
    Private Sub Fill_in_Terms()
        Dim strSQL As String = "SELECT DISTINCT RTRIM(YR_CDE) from CANDIDACY ORDER BY RTRIM(YR_CDE) DESC"
        Dim myConnectionEX As New SqlConnection(sConnString)
        myConnectionEX.Open()

        Dim myCommand As New SqlCommand(strSQL, myConnectionEX)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        While reader.Read()

            'If InStr(reader(0), "-10") > 1 Then
            'cboTermStart.Items.Add("----------------------")
            'cboTermEnd.Items.Add("----------------------")
            'End If

            cboTermStart.Items.Add(reader(0))
            cboTermEnd.Items.Add(reader(0))
            cboTargetTerm.Items.Add(reader(0))
        End While

        reader.Close()
        myCommand.Dispose()
        myConnectionEX.Close()

        If Month(Now) <= 9 Then
            cboTermStart.Text = Year(Now) ' & "-10"
            cboTermEnd.Text = Year(Now) '& "-10"
            cboTargetTerm.Text = Year(Now)
        Else
            cboTermStart.Text = Year(Now) + 1 '& "-10"
            cboTermEnd.Text = Year(Now) + 1 '& "-10"
            cboTargetTerm.Text = Year(Now) + 1
        End If

    End Sub
    Private Sub Fill_in_COUNSELORS()
        Dim strSQL As String = "SELECT DISTINCT COUNSELOR_TITLE, CANDIDATE.COUNSELOR_INITIALS from CANDIDATE "
        strSQL &= "INNER JOIN CANDIDACY ON CANDIDACY.ID_NUM = CANDIDATE.ID_NUM "
        strSQL &= "INNER JOIN COUNSELOR_RESPONSI ON CANDIDATE.COUNSELOR_INITIALS = COUNSELOR_RESPONSI.COUNSELOR_INITIALS "
        strSQL &= "WHERE CANDIDACY.YR_CDE >= Year(GETDATE()) "
        strSQL &= "ORDER BY COUNSELOR_TITLE "
        Dim myConnectionEX As New SqlConnection(sConnString)
        myConnectionEX.Open()

        Dim myCommand As New SqlCommand(strSQL, myConnectionEX)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        Dim kount As Int16 = 0

        While reader.Read()
            cboCounselor.Items.Add(reader(0) & " [" & Trim(reader(1)) & "]")
            aryCounselorInitials(kount) = reader(1)
            kount += 1

        End While

        reader.Close()
        myCommand.Dispose()
        myConnectionEX.Close()

    End Sub
    Private Sub Fill_in_HIGH_SCHOOLS()
        Dim strSQL As String = "SELECT DISTINCT RTRIM(REPLACE(REPLACE (ISNULL((SELECT RTRIM(ISNULL(LAST_NAME,' ')) + RTRIM(ISNULL(FIRST_NAME,' ')) "
        strSQL &= "FROM NAME_MASTER WHERE NAME_MASTER.ID_NUM = HIGH_SCHOOL),' '),'*',''),'HIGH SCHOOL','')) AS HS, HIGH_SCHOOL  "
        strSQL &= "FROM CANDIDATE WHERE CUR_YR IS NOT NULL AND HIGH_SCHOOL IS NOT NULL "
        ' strSQL &= "AND HIGH_SCHOOL IN (419393,523320,970000168) "
        strSQL &= "ORDER BY HS "

        Dim kount As Int16 = 0
        Dim bolUpdateThisTime As Boolean = False
        Dim strHighSchoolIDNums = ""
        Dim strLastHighSchool As String = ""
        Dim myConnectionEX As New SqlConnection(sConnString)
        myConnectionEX.Open()

        Dim myCommand As New SqlCommand(strSQL, myConnectionEX)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        While reader.Read()
            If Len(Trim(reader(0))) = 0 Then
                'DO NOTHING
            Else
                If reader(0) = strLastHighSchool Then 'SAME NAME / DIFFERENT ID_NUM
                    aryHighSchools(kount) &= reader(1) & ","
                Else 'NEW NAME
                    cboHS.Items.Add(reader(0))
                    strLastHighSchool = reader(0)

                    If bolUpdateThisTime Then
                        kount += 1
                    End If
                    aryHighSchools(kount) = reader(1) & ","

                    bolUpdateThisTime = True

                End If
            End If

        End While

        reader.Close()
        myCommand.Dispose()
        myConnectionEX.Close()

    End Sub
    Private Function GET_MATCHING_HIGH_SCHOOLS(ByVal strZips As String) As String
        Dim strSQL As String = "SELECT DISTINCT HIGH_SCHOOL, ZIP FROM dbo.CANDIDATE "
        strSQL &= "INNER JOIN ADDRESS_MASTER AM ON AM.ID_NUM = CANDIDATE.HIGH_SCHOOL "
        strSQL &= "WHERE (HIGH_SCHOOL IS NOT NULL) AND ZIP IS NOT NULL " & strZips

        Dim strReturn As String = "AND HIGH_SCHOOL IN ("
        Dim myConnectionEX As New SqlConnection(sConnString)
        myConnectionEX.Open()

        Dim myCommand As New SqlCommand(strSQL, myConnectionEX)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        While reader.Read()
            strReturn &= reader(0) & ","
        End While

        strReturn &= ")"
        strReturn = Replace(strReturn, ",)", ")")

        reader.Close()
        myCommand.Dispose()
        myConnectionEX.Close()

        Return strReturn

    End Function
    Private Function GET_MATCHING_HOME_ZIP_ID_NUMS() As String
        Dim strReturn As String = ""
        Dim strZIP As String = ""
        Dim aryZips() As String = Split(txtHomeZip.Text, ",")

        strZIP = "WHERE ("
        For Kount = 0 To aryZips.GetUpperBound(0)
            strZIP &= "ZIP LIKE '" & aryZips(Kount) & "%' OR "
        Next
        strZIP = Mid(strZIP, 1, Len(strZIP) - 4)
        strZIP &= ") "

        Dim strSQL As String = "SELECT DISTINCT CANDIDACY.ID_NUM FROM CANDIDACY "
        strSQL &= "INNER JOIN ADDRESS_MASTER AM ON AM.ID_NUM = CANDIDACY.ID_NUM "
        strSQL &= strZIP & "AND CUR_CANDIDACY = 'Y' "

        If chkAllYears.Checked Then
            'DON'T CHECK FOR YEARS
        Else
            strSQL &= "AND (YR_CDE >= " & cboTermStart.Text & " AND YR_CDE <= " & cboTermEnd.Text & ") "
        End If

        Dim myConnectionEX As New SqlConnection(sConnString)
        myConnectionEX.Open()

        Dim myCommand As New SqlCommand(strSQL, myConnectionEX)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        strReturn = "AND CANDIDACY.ID_NUM IN ("
        While reader.Read()
            strReturn &= reader(0) & ","
        End While

        strReturn &= ")"
        strReturn = Replace(strReturn, ",)", ")")

        reader.Close()
        myCommand.Dispose()
        myConnectionEX.Close()

        Return strReturn

    End Function
    Private Sub Load_Stages()
        ' Dim strSQL As String = "SELECT RTRIM(Stage) AS STAGE, RTRIM(Category_Main) AS CATEGORY_MAIN, "
        ' strSQL &= "RTRIM(Category_Detail) AS CATEGORY_DETAIL, RTRIM(BUCKET) AS BUCKET FROM NC_Admissions_Stage_Category "
        ' strSQL &= "ORDER BY CATEGORY_DETAIL"

        Dim strLastStage As String = ""
        Dim strStage As String = ""
        Dim strSQL As String = "SELECT DISTINCT STAGE, RTRIM(~cstStageType~),SORT_FIELD FROM NC_Admissions_Stage_Category "
        ' strSQL &= "ORDER BY SORT_FIELD"
        strSQL &= "ORDER BY RTRIM(~cstStageType~)" 'CHANGED 2014/01/31

        strSQL = Replace(strSQL, "~cstStageType~", cstStageType)

        Dim myConnectionEX As New SqlConnection(sConnString)
        myConnectionEX.Open()

        Dim myCommand As New SqlCommand(strSQL, myConnectionEX)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader
        inxStageCode = 0
        Dim intChecked As Int16 = 0

        While reader.Read()
            strStage = Trim(reader(1))
            If Mid(strStage, 3, 1) = "_" Then
                strStage = Mid(strStage, 4)
            End If

            If strLastStage = strStage Then
                'DON'T DUPLICATE
            Else
                chkStage.Items.Add(strStage)
                chkStage.SetItemChecked(intChecked, True)
                intChecked += 1
                strLastStage = strStage
            End If

            aryStageCode(inxStageCode) = "~" & strStage & "~?" & Trim(reader("STAGE"))
            inxStageCode += 1
        End While

        reader.Close()
        myCommand.Dispose()
        myConnectionEX.Close()

    End Sub
    Private Sub Initialize_Build_Stage_Array()
        Dim tmpIndex As Int16 = 0
        '   Dim strSQL As String = "SELECT DISTINCT ~STAGETYPE~,SORT_FIELD FROM NC_ADMISSIONS_STAGE_CATEGORY ORDER BY SORT_FIELD "
        Dim strSQL As String = "SELECT DISTINCT BUCKET FROM NC_ADMISSIONS_STAGE_CATEGORY "
        Dim myConnectionEX As New SqlConnection(sConnString)
        myConnectionEX.Open()

        strSQL = Replace(strSQL, "~STAGETYPE~", cstStageType)

        Dim myCommand As New SqlCommand(strSQL, myConnectionEX)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        While reader.Read()
            tmpIndex += 1
            aryStages(tmpIndex) = Mid(Trim(reader(0)), 4)
        End While

        aryStagesCount = tmpIndex

        reader.Close()
        myCommand.Dispose()
        myConnectionEX.Close()

    End Sub
    Private Sub Initialize_Build_Regions()

        Dim strReturn As String = ""
        Dim strSQL As String = "SELECT ZIP_CODE,RTRIM(REGION) AS REGION FROM NC_ADMISSIONS_REGIONS "
        Dim myConnectionEX As New SqlConnection(sConnString)
        myConnectionEX.Open()
        Dim myCommand As New SqlCommand(strSQL, myConnectionEX)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader
        Dim recRegions As New RECORD_REGIONS

        While reader.Read()
            aryRegionsIndex += 1

            recRegions.ZIP_CODE = "'" & reader("ZIP_CODE") 'PUT SINGLE QUOTE TO PROTECT LEADING ZEROS
            recRegions.REGION = reader("REGION")

            aryRegions(aryRegionsIndex) = recRegions
        End While

        reader.Close()
        myCommand.Dispose()
        myConnectionEX.Close()

    End Sub
    Private Sub Write_to_File(ByVal tmpObjwriter As System.IO.StreamWriter, ByVal tmpLineToWrite As String)

        tmpObjwriter.WriteLine(tmpLineToWrite)

    End Sub
    Private Function Format_Number(ByVal tmpNumber As Object) As String
        Format_Number = """" & FormatNumber(tmpNumber, 2, TriState.True, TriState.True, TriState.True) & """"
    End Function
    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        Dim strOKMessage As String = "OK"
        Dim strReportMessage As String = IS_OK_TO_RUN_REPORT(strOKMessage)

        If strReportMessage = strOKMessage Then
            'PROCEED WITH REPORT
        Else
            strReportMessage = Replace(strReportMessage, strOKMessage, "")
            MsgBox(strReportMessage)
            Exit Sub
        End If

        If btnRun.Text = "S T O P" Then
            Beep()
            bolStopRun = True
            btnRun.Text = "STOPPING.."
            btnRun.Refresh()
            btnRun.Enabled = False
            btnRun.Top = ProgressBar1.Top - btnRun.Height
            Exit Sub

        Else 'RUN REPORT
            btnRun.Enabled = False
            Me.Cursor = Cursors.WaitCursor
            Application.DoEvents()

            bolRunningSummary = chkSummary.Checked
            Call INITIALIZE_TEMP_TABLE()

            If Not bolReportStarted Then
                If System.IO.File.Exists(strOutputFileName) Then
                    Try
                        System.IO.File.Delete(strOutputFileName)
                    Catch EX As Exception
                        MsgBox("Please close Excel before running report", MsgBoxStyle.Critical, "EXCEL OPEN ERROR")
                        btnRun.Enabled = True
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End Try

                End If

                bolReportStarted = True
            End If

            If chkAllDates.Checked Then
                dteTargetDate.Value = DateAdd(DateInterval.Year, 5, dteTargetDate.Value)
            End If

            If bolRunningSummary Then
                Call GET_DATA_SUMMARY("STAGES")
            Else
                Call GET_DATA()
            End If

        End If ' RUN REPORT


        'CLEAN UP AND REFRESH
        Call Clean_Up_After_Run()


    End Sub
    Private Sub Clean_Up_After_Run()

        Me.Cursor = Cursors.Default
        bolStopRun = False
        bolReportStarted = False

        If ProgressBar1.Value > 0 Then
            If chkOnlineReport.Checked Then
                'Call Handle_Summary_Report()

                Dim tmpForm As Form = frmSSRS

                ssrs_parameter_name = "prmAsOf"
                ssrs_parameter_1 = dteTargetDate.Text
                If Len(ssrs_parameter_1) = 0 Then
                    ssrs_parameter_1 = Format(Today, "MMM dd, yyyy")
                End If
                ssrs_reportname = "Admiss"

                tmpForm.ShowDialog()


                ssrs_parameter_name = ""
                ssrs_parameter_1 = ""
                ssrs_reportname = ""
                tmpForm.Dispose()

            Else
                System.Diagnostics.Process.Start(strOutputFileName)
            End If
        Else
            MsgBox("NO DATA FOUND")
        End If

        ProgressBar1.Value = ProgressBar1.Maximum
        lblPercentDone.Visible = False
        btnRun.Text = "&RUN"
        btnRun.Refresh()
        ProgressBar1.Value = 0
        btnRun.Enabled = True
        chkAllYears.Checked = False
        lblPercentDone.Text = ""


        ' txtPhoneMatch.Text = ""
        ' txtIDMatch.Text = ""
        ' txtNameMatch.Text = ""

        Call CLOSE_DATABASE_CONNECTION(myConnectionGET_STAGE)
        Call CLOSE_DATABASE_CONNECTION(myConnectionGET_STUDENT_ADDRESS)
        Call CLOSE_DATABASE_CONNECTION(myConnectionGET_HIGH_SCHOOL_ADDRESS)
        Call CLOSE_DATABASE_CONNECTION(conTempTable)

    End Sub
    Private Sub Handle_Summary_Report()
        Dim kountRows As Int16 = 0
        Dim kountCols As Int16 = 0
        Dim kountType As Int16 = 0

        Dim strDateData As String = ""

        Dim strLine As String = ""

        Dim objWriter As New System.IO.StreamWriter(strOutputFileName, True)

        strDateData = "FOR TARGET TERM: " & cboTargetTerm.Text & " AS OF: " & Format(CDate(dteTargetDate.Text))
        strDateData &= " ( " & DateDiff(DateInterval.Month, CDate(dteTargetDate.Text), CDate("08/15/" & cboTargetTerm.Text))
        strDateData &= " MONTHS BEFORE TERM BEGINS)"

        For kountType = 0 To 5 'freshman / transfer / international / special / readmit

            'HEADER
            strLine = aryTypes(kountType) & ","
            For Kount = 1 To inxBigBuckets 'STAGES
                strLine &= aryBigBucketsNAME(Kount) & ","
            Next

            If kountType = 0 Then
                strLine &= ",," & strDateData
            End If
            Call Write_to_File(objWriter, strLine)
            strLine = ""


            For kountRows = 0 To tblOnlineReportRows - 1
                For kountCols = 0 To tblOnlineReportColumns
                    strLine &= tblOnlineReport(kountType, kountRows, kountCols) & ","
                Next

                Call Write_to_File(objWriter, strLine)
                strLine = ""
            Next

            Call Write_to_File(objWriter, " ")
            Call Write_to_File(objWriter, " ")

        Next

        objWriter.Close()

        System.Diagnostics.Process.Start(strOutputFileName)
    End Sub
    Private Sub INITIALIZE_TEMP_TABLE()
        Dim strTemp As String = ""

        Dim cmd As New System.Data.SqlClient.SqlCommand
        cmd.CommandType = System.Data.CommandType.Text

        If conTempTable.State = ConnectionState.Closed Then
            conTempTable.Open()
        End If


        cmd.CommandText = "DELETE [NC_ADMISS_TEMP_REPORT] WHERE USERNAME = '" & Environment.UserName & "'"
        cmd.Connection = conTempTable
        cmd.ExecuteNonQuery()

    End Sub
    Private Sub INITIALIZE_RECDATA()
        With recData
            .ID_NUM = "' ',"
            .LAST_NAME = "' ',"
            .FIRST_NAME = "' ',"
            .MIDDLE_NAME = "' ',"
            .PREFERRED_NAME = "' ',"
            .YEAR = "' ',"
            .TERM = "' ',"
            .GENDER = "' ',"
            .BIRTH_DATE = "' ',"
            .REGISTRAR_ENROLLED = "' ',"
            .PROG_CODE = "' ',"
            .CANDIDACY_TYPE = "' ',"
            .CURRENT_STAGE = "' ',"
            .HOME_PHONE = "' ',"
            .CELL_PHONE = "' ',"
            .EMAIL = "' ',"
            .COUNSELOR = "' ',"
            .ATHLETIC_PROSPECT = "' ',"
            .HIGH_SCHOOL = "' ',"
            .LAST_ORGANIZATION = "' ',"
            .PROSPECT_BUCKET = "' ',"
            .APPLIED_BUCKET = "' ',"
            .DENY_BUCKET = "' ',"
            .ADMITTED_BUCKET = "' ',"
            .WITHDRAWN_BUCKET = "' ',"
            .ENROLLED_BUCKET = "' ',"
            .DEPOSIT_BUCKET = "' ',"
            .PROSPECT_ADDRESS_1 = "' ',"
            .PROSPECT_ADDRESS_2 = "' ',"
            .PROSPECT_CITY = "' ',"
            .PROSPECT_STATE = "' ',"
            .PROSPECT_ZIP = "' ',"
            .LAST_ORGANIZATION_ADDRESS_1 = "' ',"
            .LAST_ORGANIZATION_ADDRESS_2 = "' ',"
            .LAST_ORGANIZATION_CITY = "' ',"
            .LAST_ORGANIZATION_STATE = "' ',"
            .LAST_ORGANIZATION_ZIP = "' ',"
            .LAST_ORGANIZATION_REGION = "' ',"
        End With

    End Sub
    Private Sub WRITE_TO_TEMP_TABLE()
        Dim strTemp As String = ""

        Dim cmd As New System.Data.SqlClient.SqlCommand
        cmd.CommandType = System.Data.CommandType.Text

        cmd.CommandText = "INSERT INTO [NC_ADMISS_TEMP_REPORT] (Username,ID_NUM,LAST_NAME,FIRST_NAME,MIDDLE_NAME,PREFERRED_NAME,YEAR,TERM,GENDER,BIRTH_DATE,"
        cmd.CommandText &= "REGISTRAR_ENROLLED,PROG_CODE,CANDIDACY_TYPE,CURRENT_STAGE,HOME_PHONE,CELL_PHONE,EMAIL,COUNSELOR,ATHLETIC_PROSPECT,HIGH_SCHOOL,"
        cmd.CommandText &= "LAST_ORGANIZATION,PROSPECT_BUCKET,APPLIED_BUCKET,DENY_BUCKET,ADMITTED_BUCKET,WITHDRAWN_BUCKET,ENROLLED_BUCKET,PROSPECT_ADDRESS_1,"
        cmd.CommandText &= "PROSPECT_ADDRESS_2,PROSPECT_CITY,PROSPECT_STATE,PROSPECT_ZIP,LAST_ORGANIZATION_ADDRESS_1,LAST_ORGANIZATION_ADDRESS_2,"
        cmd.CommandText &= "LAST_ORGANIZATION_CITY,LAST_ORGANIZATION_STATE,LAST_ORGANIZATION_ZIP,LAST_ORGANIZATION_REGION) VALUES ("

        With recData
            cmd.CommandText &= "'" & Environment.UserName & "'," & .ID_NUM & .LAST_NAME & .FIRST_NAME & .MIDDLE_NAME & .PREFERRED_NAME & .YEAR & .TERM & .GENDER & .BIRTH_DATE
            cmd.CommandText &= .REGISTRAR_ENROLLED & .PROG_CODE & .CANDIDACY_TYPE & .CURRENT_STAGE & .HOME_PHONE & .CELL_PHONE & .EMAIL & .COUNSELOR & .ATHLETIC_PROSPECT & .HIGH_SCHOOL
            cmd.CommandText &= .LAST_ORGANIZATION & .PROSPECT_BUCKET & .APPLIED_BUCKET & .DENY_BUCKET & .ADMITTED_BUCKET & .WITHDRAWN_BUCKET & .ENROLLED_BUCKET & .PROSPECT_ADDRESS_1
            cmd.CommandText &= .PROSPECT_ADDRESS_2 & .PROSPECT_CITY & .PROSPECT_STATE & .PROSPECT_ZIP & .LAST_ORGANIZATION_ADDRESS_1 & .LAST_ORGANIZATION_ADDRESS_2
            cmd.CommandText &= .LAST_ORGANIZATION_CITY & .LAST_ORGANIZATION_STATE & .LAST_ORGANIZATION_ZIP & .LAST_ORGANIZATION_REGION
        End With

        cmd.CommandText = Trim(cmd.CommandText)
        cmd.CommandText = Mid(cmd.CommandText, 1, Len(cmd.CommandText) - 1) & ")" ';STRIP OFF LAST COMMA


        cmd.Connection = conTempTable
        cmd.ExecuteNonQuery()


    End Sub
    Private Function BUILD_STAGES_HEADER() As String
        Dim tmpReturnString As String = ""
        Dim tmpIndex As Int16 = 0

        For tmpIndex = 1 To aryStagesCount
            tmpReturnString &= aryStages(tmpIndex) & ","
        Next

        tmpReturnString &= "OTHER,"

        For tmpIndex = 1 To inxBigBuckets
            tmpReturnString &= aryBigBucketsNAME(tmpIndex) & " BUCKET,"
        Next

        Return tmpReturnString
    End Function
    Private Sub Build_Header()
        Dim objWriter As New System.IO.StreamWriter(strOutputFileName, False)

        Dim strHeader As String = "'ID_NUM,LAST_NAME,FIRST_NAME,MIDDLE_NAME, PREFERRED_NAME, YEAR, TERM, GENDER, BIRTH_DATE, REGISTRAR_ENROLLED, PROG_CODE, "

        strHeader &= "CANDIDACY_TYPE, CURRENT_STAGE, HOME PHONE, CELL PHONE, EMAIL, COUNSELOR, ATHLETIC_PROSPECT, HIGH_SCHOOL, LAST_ORGANIZATION, "

        'ADD THE ADDITIONAL INFO SELECTED BY USER
        Dim Kount As Int16 = 0
        For Kount = 0 To chkDataItems.Items.Count - 1
            If chkDataItems.GetItemChecked(Kount) Then
                Select Case Trim(chkDataItems.Items(Kount))

                    Case "STUDENT ADDRESS"
                        strHeader &= "STUDENT ADDR_LINE_1, STUDENT ADDR_LINE_2,STUDENT CITY, STUDENT STATE, STUDENT ZIP,"
                        myConnectionGET_STUDENT_ADDRESS.Open()
                    Case "LAST ORG ADDRESS"
                        strHeader &= "LO ADDR_LINE_1, LO ADDR_LINE_2,LO CITY, LO STATE, LO ZIP, LO REGION,"
                        myConnectionGET_HIGH_SCHOOL_ADDRESS.Open()
                    Case "STAGES"
                        strHeader &= BUILD_STAGES_HEADER()
                        myConnectionGET_STAGE.Open()
                        Call Initialize_Report_Table("STAGE")

                End Select
            End If
        Next

        If Not bolRunningSummary Then
            Call Write_to_File(objWriter, strHeader)
        End If

        objWriter.Close()

    End Sub
    Private Sub Initialize_Report_Table(ByVal tmpReportType As String)
        Dim kountIndex As Int16 = 0
        Dim kountType As Int16 = 0

        Select Case tmpReportType
            Case "STAGE"

                For kountType = 0 To 5
                    For kountIndex = 0 To (cboTermEnd.Text - cboTermStart.Text) 'YEARS
                        tblOnlineReport(kountType, kountIndex, 0) = (cboTargetTerm.Text) - kountIndex
                    Next
                Next


                For kountIndex = 1 To inxBigBuckets 'STAGES
                    tblOnlineReportHeaders(kountIndex) = aryBigBucketsNAME(kountIndex)
                Next

                tblOnlineReportColumns = inxBigBuckets
                tblOnlineReportRows = (cboTermEnd.Text - cboTermStart.Text) + 1  'NUMBER OF YEARS
        End Select


    End Sub
    Private Function GET_MATCHING_ID_NUMS(ByVal strSQL As String) As String

        Dim KOUNT As Integer = 0
        Dim bolHasRows As Boolean = False
        Dim strReturn As String = ""
        Dim myConnectionEX As New SqlConnection(sConnString)
        myConnectionEX.Open()

        Dim myCommand As New SqlCommand(strSQL, myConnectionEX)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        bolHasRows = reader.HasRows

        If bolHasRows Then
            strReturn = "("
        End If

        While reader.Read()
            strReturn &= reader(0) & ","
        End While

        If bolHasRows Then
            strReturn &= ") "
            strReturn = Replace(strReturn, ",)", ")")
        End If

        reader.Close()
        myCommand.Dispose()
        myConnectionEX.Close()

        Return strReturn

    End Function
    Private Function FILTERS() As String
        Dim strReturn As String = ""
        Dim Kount As Int16 = 0
        Dim tmpPhone As String = ""
        Dim tmpString As String = ""
        Dim tmpIDNums As String = ""
        Dim aryZips(10) As String
        Dim strZip As String = ""

        If Val(txtIDMatch.Text) > 0 Then 'IF THE USER IS DOING A ID_MATCH, DON'T MATCH ON ANYTHING ELSE
            strReturn &= "AND (CANDIDACY.ID_NUM = " & Val(txtIDMatch.Text) & ") "
            'Select all data for a single ID_NUM
            For i As Integer = 0 To chkDataItems.Items.Count - 1
                chkDataItems.SetItemChecked(i, True)
            Next

        Else
            strReturn = "AND CANDIDACY_TYPE IN ("
            For Kount = 0 To chkType.Items.Count - 1
                If chkType.GetItemChecked(Kount) Then
                    strReturn &= "'" & Mid(chkType.Items(Kount), 1, 1) & "'"
                End If
            Next

            strReturn = Replace(strReturn, "''", "','")
            strReturn &= ") "


            If Len(txtNameMatch.Text) > 0 Then
                tmpString = "SELECT ID_NUM FROM NAME_MASTER WHERE "
                tmpString &= "(LAST_NAME LIKE '%" & Trim(txtNameMatch.Text) & "%' OR FIRST_NAME LIKE '%" & Trim(txtNameMatch.Text) & "%' "
                tmpString &= "OR   MIDDLE_NAME LIKE '%" & Trim(txtNameMatch.Text) & "%' OR PREFERRED_NAME LIKE '%" & Trim(txtNameMatch.Text) & "%') "

                tmpIDNums = GET_MATCHING_ID_NUMS(tmpString)
                If Len(tmpIDNums) = 0 Then
                    strReturn &= " AND 0 = 1 "
                Else
                    strReturn &= " AND CANDIDACY.ID_NUM IN " & tmpIDNums
                End If

                'strReturn &= "AND (LAST_NAME LIKE '%" & Trim(txtNameMatch.Text) & "%' OR FIRST_NAME LIKE '%" & Trim(txtNameMatch.Text) & "%' "
                'strReturn &= "OR   MIDDLE_NAME LIKE '%" & Trim(txtNameMatch.Text) & "%' OR PREFERRED_NAME LIKE '%" & Trim(txtNameMatch.Text) & "%') "
            End If

            If Len(txtPhoneMatch.Text) > 0 Then
                tmpPhone = Replace(Replace(txtPhoneMatch.Text, "-", ""), "(", "")
                tmpPhone = Replace(txtPhoneMatch.Text, ")", "")

                tmpString = "SELECT ID_NUM FROM ADDRESS_MASTER WHERE PHONE LIKE '%~PHONE~%' UNION SELECT ID_NUM FROM NAME_MASTER WHERE MOBILE_PHONE LIKE '%~PHONE~%' "
                tmpString = Replace(tmpString, "~PHONE~", tmpPhone)

                tmpIDNums = GET_MATCHING_ID_NUMS(tmpString)
                If Len(tmpIDNums) = 0 Then
                    strReturn &= " AND 0 = 1 "
                Else
                    strReturn &= " AND CANDIDACY.ID_NUM IN " & tmpIDNums
                    chkAllYears.Checked = True
                End If


                ' strReturn &= "AND ((SELECT PHONE FROM ADDRESS_MASTER WHERE ADDRESS_MASTER.ID_NUM = CANDIDACY.ID_NUM AND ADDRESS_MASTER.ADDR_CDE = '*LHP') LIKE '%~PHONE~%' OR NAME_MASTER.MOBILE_PHONE LIKE '%~PHONE~%') "

                ' strReturn &= "AND ( (SELECT PHONE FROM ADDRESS_MASTER WHERE ADDRESS_MASTER.ID_NUM = CANDIDACY.ID_NUM AND ADDRESS_MASTER.PHONE LIKE '%~PHONE~%') "
                ' strReturn &= "OR (SELECT COUNT(PHONE) FROM NAME_MASTER) ) " 


            End If


            '  If chkAllYears.Checked Then
            ''DON'T CHECK FOR YEARS
            'Else
            strReturn &= "AND (YR_CDE >= " & cboTermStart.Text & " AND YR_CDE <= " & cboTermEnd.Text & ") "
            'End If

            If rdoFall.Checked Then
                strReturn &= "AND CANDIDACY.TRM_CDE = '10' "
            End If

            If rdoSpring.Checked Then
                strReturn &= "AND CANDIDACY.TRM_CDE = '20' "
            End If

            Dim bolAllChecked As Boolean = True
            Dim intStageDetail As Int16 = 0

            tmpString = "AND CANDIDACY.STAGE IN ("
            For Kount = 0 To chkStage.Items.Count - 1

                If chkStage.GetItemChecked(Kount) = True Then 'get all stages that match stage description
                    For intStageDetail = 0 To inxStageCode
                        If InStr(aryStageCode(intStageDetail), chkStage.Items(Kount) & "~") > 0 Then
                            tmpString &= "'" & Mid(aryStageCode(intStageDetail), InStr(aryStageCode(intStageDetail), "?") + 1) & "',"
                        End If
                    Next

                Else
                    bolAllChecked = False
                End If
            Next

            If bolAllChecked Then
                'don't add this check since all stages are selected
            Else
                tmpString = Mid(tmpString, 1, Len(tmpString) - 1) & ") "
                strReturn &= tmpString
            End If

            If chkOnlineReport.Checked Then
                For i As Integer = 0 To chkDataItems.Items.Count - 1
                    chkDataItems.SetItemChecked(i, True)
                Next
            End If


            If Len(Trim(cboCounselor.Text)) > 0 Then
                strReturn &= " AND CANDIDATE.COUNSELOR_INITIALS = '" & aryCounselorInitials(cboCounselor.SelectedIndex) & "' "
            End If

            If Len(Trim(cboHS.Text)) > 0 Then
                If cboHS.SelectedIndex = -1 Then 'if user typed in school instead of selecting one in combobox
                    Dim index As Integer
                    index = cboHS.FindString(cboHS.Text)
                    cboHS.SelectedIndex = index

                    If index = -1 Then
                        strReturn = "CANNOT FIND HIGH SCHOOL"
                        cboHS.Focus()
                        Return strReturn
                        Exit Function
                    Else
                        cboHS.Text = cboHS.Items(index).ToString
                    End If
                End If

                strReturn &= " AND HIGH_SCHOOL IN (" & aryHighSchools(cboHS.SelectedIndex) & ") "
                strReturn = Replace(strReturn, ",)", ")")
                strReturn &= "AND CANDIDATE.HIGH_SCHOOL = ISNULL(CANDIDATE.LAST_ORG_ATTEND,CANDIDATE.HIGH_SCHOOL) "
            End If ' HIGH SCHOOL

            If Len(Trim(txtHSZIP.Text)) > 0 Then
                strZip = ""
                aryZips = Split(txtHSZIP.Text, ",")
                For Kount = 0 To aryZips.GetUpperBound(0)
                    strZip &= "ZIP LIKE '" & aryZips(Kount) & "%' OR "
                Next

                strReturn &= GET_MATCHING_HIGH_SCHOOLS("AND (" & Mid(strZip, 1, Len(strZip) - 4) & ") ")

                For Kount = 0 To chkDataItems.Items.Count - 1
                    If chkDataItems.Items(Kount).ToString = "LAST ORG ADDRESS" Then
                        chkDataItems.SetItemChecked(Kount, True)
                    End If
                Next
            End If 'HS ZIP CODE


            If Len(Trim(txtHomeZip.Text)) > 0 Then

                strReturn &= GET_MATCHING_HOME_ZIP_ID_NUMS()

                For Kount = 0 To chkDataItems.Items.Count - 1
                    If chkDataItems.Items(Kount).ToString = "STUDENT ADDRESS" Then
                        chkDataItems.SetItemChecked(Kount, True)
                    End If
                Next
            End If



        End If 'ID MATCH

        Return strReturn
    End Function
    Private Function IS_OK_TO_RUN_REPORT(ByVal strOKMessage As String) As String
        Dim strReportMessage As String = strOKMessage
        Dim Kount As Int16 = 0

        'verify a type is selected
        Dim bol_type_OK As Boolean = False
        For Kount = 0 To chkType.Items.Count - 1
            If chkType.GetItemChecked(Kount) Then
                bol_type_OK = True
            End If
        Next
        If Not bol_type_OK Then
            strReportMessage &= "SELECT A CANDIDACY TYPE" & vbCrLf
        End If


        If cboTermStart.Text > cboTermEnd.Text Then
            strReportMessage &= "TERM START CANNOT BE LATER THAN TERM END" & vbCrLf
        End If


        Return strReportMessage
    End Function
    Private Sub GET_DATA()

        Dim percent As Integer = 0
        Dim strAllData As String = ""
        Dim tmpReturnString As String = ""
        Dim bolShowCandidate As Boolean = True
        Dim strFilters As String = ""

        Dim strSQL As String = "SELECT DISTINCT CANDIDACY.ID_NUM,LAST_NAME, FIRST_NAME, MIDDLE_NAME,ISNULL(PREFERRED_NAME,FIRST_NAME) AS PREFERRED_NAME, YR_CDE,TRM_CDE, "
        strSQL &= "ISNULL((SELECT PROG_DESC FROM dbo.PROGRAM_DEF WHERE (PROG_CDE = dbo.CANDIDACY.PROG_CDE)), ' ') AS PROG_CDE,CANDIDACY_TYPE,STAGE, "
        strSQL &= "CASE WHEN (SELECT COUNT(HRS_ATTEMPTED) FROM STUDENT_CRS_HIST WHERE TRANSACTION_STS IN ('H','C') AND STUDENT_CRS_HIST.ID_NUM = CANDIDACY.ID_NUM "
        strSQL &= "AND STUDENT_CRS_HIST.YR_CDE >= CANDIDACY.YR_CDE) > 0 THEN '1' ELSE '0' END AS REGISTRAR_ENROLLED, GENDER, BIRTH_DTE, ISNULL(DECEASED,'N') AS DECEASED, "
        strSQL &= "(SELECT PHONE FROM ADDRESS_MASTER WHERE ADDRESS_MASTER.ID_NUM = CANDIDACY.ID_NUM AND ADDRESS_MASTER.ADDR_CDE = '*LHP') AS PHONE,  NAME_MASTER.MOBILE_PHONE AS MOBILE_PHONE,"
        strSQL &= "(SELECT COUNSELOR_TITLE FROM COUNSELOR_RESPONSI WHERE COUNSELOR_RESPONSI.COUNSELOR_INITIALS = CANDIDATE.COUNSELOR_INITIALS) AS COUNSELOR,"
        strSQL &= "(SELECT  STAGE_DESC FROM STAGE_CONFIG WHERE STAGE_CONFIG.STAGE = CANDIDACY.STAGE) AS CURRENT_STAGE, "
        strSQL &= "ISNULL(HIGH_SCHOOL,0) AS HIGH_SCHOOL_ID, "
        strSQL &= "ISNULL(ISNULL(CANDIDATE.LAST_ORG_ATTEND,HIGH_SCHOOL),'  ') AS LAST_ORGANIZATION_ID,"
        strSQL &= "ISNULL((SELECT RTRIM(ISNULL(LAST_NAME,'') + ISNULL(FIRST_NAME,'')) FROM NAME_MASTER WHERE NAME_MASTER.ID_NUM = CANDIDATE.HIGH_SCHOOL),' ') AS HIGH_SCHOOL_NAME, "
        strSQL &= "ISNULL((SELECT RTRIM(ISNULL(LAST_NAME,'') + ISNULL(FIRST_NAME,'')) FROM NAME_MASTER WHERE NAME_MASTER.ID_NUM = ISNULL(CANDIDATE.LAST_ORG_ATTEND,CANDIDATE.HIGH_SCHOOL)),' ') AS LAST_ORGANIZATION_NAME, "
        strSQL &= "ATHLETIC_PROSPECT, ISNULL(EMAIL_ADDRESS,' ') AS EMAIL_ADDRESS FROM CANDIDACY " 'DON'T CHANGE "FROM CANDIDACY" BECAUSE IT'S USED IN GET_ROW_COUNT
        strSQL &= "INNER JOIN CANDIDATE ON CANDIDATE.ID_NUM = CANDIDACY.ID_NUM "
        strSQL &= "INNER JOIN NAME_MASTER ON NAME_MASTER.ID_NUM = CANDIDACY.ID_NUM "
        strSQL &= "LEFT OUTER JOIN BIOGRAPH_MASTER ON BIOGRAPH_MASTER.ID_NUM = CANDIDACY.ID_NUM "
        strSQL &= "WHERE CUR_CANDIDACY = 'Y' "

        '   strSQL &= " and candidacy.id_num = 532234 "

        strFilters = FILTERS() 'SETUP THE FILTERS FOR THE SQL CALL
        If Mid(strFilters, 1, 3) <> "AND" Then
            ProgressBar1.Value = 0
            MsgBox(strFilters, MsgBoxStyle.Critical)
            Exit Sub
        End If

        strSQL &= strFilters

        Call Build_Header()


        Dim objWriter As New System.IO.StreamWriter(strOutputFileName, True)

        lblPercentDone.Visible = True
        lblPercentDone.Text = "COUNTING RECORDS..."
        Me.Cursor = Cursors.WaitCursor
        lblPercentDone.Refresh()

        ProgressBar1.Maximum = GET_ROW_COUNT(strSQL)
        ProgressBar1.Value = 0

        lblPercentDone.Text = "RETRIEVING " & Format(ProgressBar1.Maximum, "###,###") & " RECORDS..."
        lblPercentDone.Refresh()

        Dim myConnectionEX As New SqlConnection(sConnString)
        myConnectionEX.Open()

        strSQL &= "ORDER BY LAST_NAME, FIRST_NAME, MIDDLE_NAME "
        Dim myCommand As New SqlCommand(strSQL, myConnectionEX)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        btnRun.Text = "S T O P"
        btnRun.Cursor = Cursors.Arrow
        btnRun.Enabled = True
        btnRun.Refresh()

        If bolRunningSummary Then
            objWriter.Close()
        End If


        While reader.Read()
            If bolStopRun Then
                Application.DoEvents()
                Exit While
            End If
            strAllData = ""
            If reader("DECEASED") = "Y" Then
                strAllData &= "* DECEASED * "
            End If

            Call INITIALIZE_RECDATA()
            recData.ID_NUM = HANDLE_INSERT_DATA(reader("ID_NUM"))
            recData.LAST_NAME = HANDLE_INSERT_DATA(reader("LAST_NAME"))
            recData.FIRST_NAME = HANDLE_INSERT_DATA(reader("FIRST_NAME"))
            recData.MIDDLE_NAME = HANDLE_INSERT_DATA(reader("MIDDLE_NAME"))
            recData.PREFERRED_NAME = HANDLE_INSERT_DATA(reader("PREFERRED_NAME"))
            recData.YEAR = HANDLE_INSERT_DATA(reader("YR_CDE"))
            recData.TERM = HANDLE_INSERT_DATA(reader("TRM_CDE"))
            recData.GENDER = HANDLE_INSERT_DATA(reader("GENDER"))
            recData.BIRTH_DATE = HANDLE_INSERT_DATA(reader("BIRTH_DTE"))
            recData.REGISTRAR_ENROLLED = HANDLE_INSERT_DATA(reader("REGISTRAR_ENROLLED"))
            recData.PROG_CODE = HANDLE_INSERT_DATA(reader("PROG_CDE"))
            recData.CANDIDACY_TYPE = HANDLE_INSERT_DATA(reader("CANDIDACY_TYPE"))
            recData.CURRENT_STAGE = HANDLE_INSERT_DATA(reader("CURRENT_STAGE"))
            recData.HOME_PHONE = HANDLE_INSERT_DATA(reader("PHONE"), 10)
            recData.CELL_PHONE = HANDLE_INSERT_DATA(reader("MOBILE_PHONE"), 10)
            recData.EMAIL = HANDLE_INSERT_DATA(reader("EMAIL_ADDRESS"))
            recData.COUNSELOR = HANDLE_INSERT_DATA(reader("COUNSELOR"))
            recData.ATHLETIC_PROSPECT = HANDLE_INSERT_DATA(reader("ATHLETIC_PROSPECT"))
            recData.HIGH_SCHOOL = HANDLE_INSERT_DATA(reader("HIGH_SCHOOL_NAME"))
            recData.LAST_ORGANIZATION = HANDLE_INSERT_DATA(reader("LAST_ORGANIZATION_NAME"))

            strAllData &= reader("ID_NUM") & ","
            strAllData &= HANDLE_COMMAS(reader("LAST_NAME"))
            strAllData &= HANDLE_COMMAS(reader("FIRST_NAME"))
            strAllData &= HANDLE_COMMAS(reader("MIDDLE_NAME"))
            strAllData &= HANDLE_COMMAS(reader("PREFERRED_NAME"))
            strAllData &= HANDLE_COMMAS(reader("YR_CDE"))
            strAllData &= HANDLE_COMMAS(reader("TRM_CDE"))
            strAllData &= HANDLE_COMMAS(reader("GENDER"))
            strAllData &= HANDLE_COMMAS(reader("BIRTH_DTE"))
            strAllData &= HANDLE_COMMAS(reader("REGISTRAR_ENROLLED"))
            strAllData &= HANDLE_COMMAS(reader("PROG_CDE"))
            strAllData &= HANDLE_COMMAS(reader("CANDIDACY_TYPE"))
            strAllData &= HANDLE_COMMAS(reader("CURRENT_STAGE"))

            strAllData &= HANDLE_PHONE(reader("PHONE"))
            strAllData &= HANDLE_PHONE(reader("MOBILE_PHONE"))
            strAllData &= HANDLE_PHONE(reader("EMAIL_ADDRESS"))

            strAllData &= HANDLE_COMMAS(reader("COUNSELOR"))
            strAllData &= HANDLE_COMMAS(reader("ATHLETIC_PROSPECT"))
            strAllData &= HANDLE_COMMAS(reader("HIGH_SCHOOL_NAME"))
            strAllData &= HANDLE_COMMAS(reader("LAST_ORGANIZATION_NAME"))

            bolShowCandidate = True

            'ADD THE ADDITIONAL INFO SELECTED BY USER
            Dim Kount As Int16 = 0

            For Kount = 0 To chkDataItems.Items.Count - 1
                If chkDataItems.GetItemChecked(Kount) Then
                    Select Case Trim(chkDataItems.Items(Kount))

                        Case "STUDENT ADDRESS"
                            strAllData &= GET_STUDENT_ADDRESS(reader("ID_NUM"))
                        Case "LAST ORG ADDRESS" 'CHANGED TO HOLD LAST ORGANIZATION'S ADDRESS
                            strAllData &= GET_LAST_ORGANIZATION_ADDRESS(reader("LAST_ORGANIZATION_ID"), reader("ID_NUM"))
                        Case "STAGES"
                            '  If chkShowCandidatesWithNoStages.Checked Then
                            'strAllData &= GET_STAGES(reader("ID_NUM"), reader("YR_CDE"), reader("CANDIDACY_TYPE"), "DO_NOT_SHOW_CANDIDATE")
                            'Else
                            tmpReturnString = GET_STAGES(reader("ID_NUM"), reader("YR_CDE"), reader("CANDIDACY_TYPE"), "DO_NOT_SHOW_CANDIDATE")
                            If tmpReturnString = "DO_NOT_SHOW_CANDIDATE" Then
                                bolShowCandidate = False
                            Else
                                strAllData &= tmpReturnString
                            End If
                            'End If

                    End Select
                End If
            Next

            If bolShowCandidate And Not bolRunningSummary Then
                Call Write_to_File(objWriter, strAllData)
                If chkOnlineReport.Checked Then
                    Call WRITE_TO_TEMP_TABLE()
                End If

            End If


            If ProgressBar1.Value < ProgressBar1.Maximum Then
                ProgressBar1.Value += 1
            End If

            If ProgressBar1.Value Mod 10 = 0 Then

                lblPercentDone.Text = "[" & Format((ProgressBar1.Value) / ProgressBar1.Maximum, "###.0%") & "]" & vbCrLf & _
                                        "     " & Format(ProgressBar1.Value, "###,###") & "/" & Format(ProgressBar1.Maximum, "###,###")

                Application.DoEvents()
            End If

        End While

        lblPercentDone.Text = Format((ProgressBar1.Value) / ProgressBar1.Maximum, "###.0%")

        reader.Close()
        myCommand.Dispose()
        myConnectionEX.Close()
        objWriter.Close()

    End Sub
    Private Sub GET_DATA_SUMMARY(ByVal strSummaryType As String)

        ' Dim objWriter As New System.IO.StreamWriter(strOutputFileName, True)

        Dim percent As Integer = 0
        Dim strAllData As String = ""
        Dim tmpReturnString As String = ""
        Dim bolShowCandidate As Boolean = True
        Dim strSQL As String = "SELECT DISTINCT CANDIDACY.ID_NUM, CANDIDACY_TYPE, YR_CDE,TRM_CDE FROM CANDIDACY "
        strSQL &= "WHERE CUR_CANDIDACY = 'Y' "

        strSQL &= FILTERS() 'SETUP THE FILTERS FOR THE SQL CALL

        lblPercentDone.Visible = True
        lblPercentDone.Text = "COUNTING RECORDS..."
        Me.Cursor = Cursors.WaitCursor
        lblPercentDone.Refresh()

        ProgressBar1.Maximum = GET_ROW_COUNT(strSQL)
        ProgressBar1.Value = 0

        lblPercentDone.Text = "RETRIEVING " & Format(ProgressBar1.Maximum, "###,###") & " RECORDS..."
        lblPercentDone.Refresh()

        Dim myConnectionEX As New SqlConnection(sConnString)
        myConnectionEX.Open()

        strSQL &= "ORDER BY YR_CDE, CANDIDACY_TYPE "
        Dim myCommand As New SqlCommand(strSQL, myConnectionEX)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        btnRun.Text = "S T O P"
        btnRun.Cursor = Cursors.Arrow
        btnRun.Enabled = True
        btnRun.Refresh()

        Select Case strSummaryType
            Case "STAGES"
                myConnectionGET_STAGE.Open()
                Call Initialize_Report_Table("STAGE")
        End Select


        While reader.Read()
            If bolStopRun Then
                Application.DoEvents()
                Exit While
            End If
            strAllData = ""

            Select Case strSummaryType

                Case "STAGES"
                    strAllData &= GET_STAGES(reader("ID_NUM"), reader("YR_CDE"), reader("CANDIDACY_TYPE"), "DO_NOT_SHOW_CANDIDATE")

            End Select


            If ProgressBar1.Value < ProgressBar1.Maximum Then
                ProgressBar1.Value += 1
            End If

            If ProgressBar1.Value Mod 10 = 0 Then

                lblPercentDone.Text = "[" & Format((ProgressBar1.Value) / ProgressBar1.Maximum, "###.0%") & "]" & vbCrLf & _
                                        "     " & Format(ProgressBar1.Value, "###,###") & "/" & Format(ProgressBar1.Maximum, "###,###")

                Application.DoEvents()
            End If

        End While

        lblPercentDone.Text = Format((ProgressBar1.Value) / ProgressBar1.Maximum, "###.0%")

        reader.Close()
        myCommand.Dispose()
        myConnectionEX.Close()

    End Sub
    Private Function GET_ROW_COUNT(ByVal tmpSQL As String) As Long
        Dim intRowCount As Long = 0
        Dim Kount As Int16 = 0
        Dim conn As New SqlConnection(sConnString)
        Dim cmd As New SqlCommand(tmpSQL, conn)
        cmd.CommandType = CommandType.Text

        Kount = InStr(tmpSQL, "FROM CANDIDACY")
        tmpSQL = "SELECT ID_NUM " & Mid(tmpSQL, Kount - 1)

        Dim da As New SqlDataAdapter(cmd)

        Dim results As New DataTable("SearchResults")

        Try
            conn.Open()
            da.Fill(results)
            intRowCount = results.Rows.Count
        Catch ex As SqlException
            da.Dispose()
        Finally
            conn.Close()
        End Try

        Return intRowCount
    End Function
    Private Function GET_STUDENT_ADDRESS(ByVal ID_NUM As String) As String
        Dim strSQL As String = "SELECT ADDR_LINE_1, ADDR_LINE_2, CITY, STATE, '''' + SUBSTRING(ISNULL(ZIP,'      '),1,5) AS ZIP FROM ADDRESS_MASTER WHERE ADDR_CDE = '*LHP' AND ID_NUM = " & ID_NUM
        Dim KOUNT As Integer = 0
        Dim strReturn As String = ""
        Dim bolSpecialCall As Boolean = False

        bolSpecialCall = (myConnectionGET_STUDENT_ADDRESS.State = ConnectionState.Closed)

        If bolSpecialCall Then
            myConnectionGET_STUDENT_ADDRESS.Open()
        End If

        Dim myCommand As New SqlCommand(strSQL, myConnectionGET_STUDENT_ADDRESS)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        If reader.HasRows Then
            reader.Read()
            For KOUNT = 0 To reader.VisibleFieldCount - 1
                strReturn &= HANDLE_COMMAS(reader(KOUNT))
            Next

            recData.PROSPECT_ADDRESS_1 = HANDLE_INSERT_DATA(reader("ADDR_LINE_1"))
            recData.PROSPECT_ADDRESS_2 = HANDLE_INSERT_DATA(reader("ADDR_LINE_2"))
            recData.PROSPECT_CITY = HANDLE_INSERT_DATA(reader("CITY"))
            recData.PROSPECT_STATE = HANDLE_INSERT_DATA(reader("STATE"))
            recData.PROSPECT_ZIP = HANDLE_INSERT_DATA(reader("ZIP"))
        Else
            For KOUNT = 0 To reader.VisibleFieldCount - 1
                strReturn &= ","
            Next
        End If

        reader.Close()
        myCommand.Dispose()


        If bolSpecialCall Then
            myConnectionGET_STUDENT_ADDRESS.Close()
        End If

        Return strReturn
    End Function
    Private Function GET_STAGES(ByVal ID_NUM As String, ByVal ENTRY_TERM_YEAR As Int16, ByVal tmpCandidateType As String, ByVal strNoStages As String) As String
        Dim strSQL As String = "SELECT HIST_STAGE, HIST_STAGE_DTE, ISNULL(~STAGETYPE~,'OTHER') AS ~STAGETYPE~, SORT_FIELD, "
        strSQL &= "RTRIM(SUBSTRING(ISNULL(BUCKET,'00_OTHER'),4,50)) AS BUCKET FROM STAGE_HISTORY_TRAN "
        strSQL &= "LEFT OUTER JOIN NC_Admissions_Stage_Category ON NC_Admissions_Stage_Category.Stage = STAGE_HISTORY_TRAN.HIST_STAGE "
        strSQL &= "WHERE ID_NUM = " & ID_NUM & " "
        strSQL &= "AND STAGE_HISTORY_TRAN.HIST_STAGE_DTE <= '~SAME_MONTH_DAY~' "

        strSQL &= "AND ADD_TO_COUNT_DUP = 'Y' "

        strSQL &= "AND (YR_CDE >= " & cboTermStart.Text & " AND YR_CDE <= " & cboTermEnd.Text & ") "
        If rdoFall.Checked Then
            strSQL &= "AND STAGE_HISTORY_TRAN.TRM_CDE = '10' "
        End If

        If rdoSpring.Checked Then
            strSQL &= "AND STAGE_HISTORY_TRAN.TRM_CDE = '20' "
        End If


        strSQL &= "ORDER BY SORT_FIELD,HIST_STAGE_DTE DESC "

        Dim bolAtLeastOneStage As Boolean = False
        Dim kountType As Int16 = 0
        Dim intTypeIndex As Int16 = 0

        Dim dteSameMonthDay As Date = Nothing
        Dim years_to_subtract As Int16 = 0


        years_to_subtract = ENTRY_TERM_YEAR - cboTargetTerm.Text
        dteSameMonthDay = DateAdd(DateInterval.Year, years_to_subtract, dteTargetDate.Value)
        strSQL = Replace(strSQL, "~SAME_MONTH_DAY~", dteSameMonthDay)

        Dim KOUNT As Integer = 0
        Dim SUBKOUNT As Int16 = 0
        Dim bolStageFound As Boolean = False

        Dim aryBuckets(50) As String

        Dim intStageArrayLocation As Int16 = 0
        Dim strReturn As String = ""

        For KOUNT = 1 To UBound(aryBuckets)
            aryBuckets(KOUNT) = "0"
        Next

        For KOUNT = 1 To inxBigBuckets
            aryBigBuckets(KOUNT) = "0"
        Next

        strSQL = Replace(strSQL, "~STAGETYPE~", cstStageType)

        Dim myCommand As New SqlCommand(strSQL, myConnectionGET_STAGE)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        While reader.Read
            For KOUNT = 1 To aryStagesCount

                If Mid(Trim(reader(cstStageType)), 4) = aryStages(KOUNT) Then
                    aryBuckets(KOUNT) = Format(reader("HIST_STAGE_DTE"), "yyyy/MM/dd")
                    Exit For
                End If

            Next

            For KOUNT = 1 To inxBigBuckets 'PROSPECT / APPLIED / DENY / ACCEPTED / WITHDRAWN / ENROLLED

                If Mid(Trim(reader("BUCKET")), 4) = aryBigBucketsNAME(KOUNT) Then
                    aryBigBuckets(KOUNT) = 1
                    If KOUNT > 0 Then

                        For kountType = 1 To 5 'CANDIDATE TYPES
                            If Mid(aryTypes(kountType), 1, 1) = tmpCandidateType Then
                                intTypeIndex = kountType
                                Exit For
                            End If
                        Next

                        tblOnlineReport(intTypeIndex, cboTargetTerm.Text - ENTRY_TERM_YEAR, KOUNT) += 1  'CANDIDATE TYPE / YEAR / BUCKET
                        tblOnlineReport(0, cboTargetTerm.Text - ENTRY_TERM_YEAR, KOUNT) += 1 ' 0 IS THE GRAND TOTAL

                        Select Case aryBigBucketsNAME(KOUNT)
                            Case "PROSPECT"
                                recData.PROSPECT_BUCKET = "1,"
                            Case "APPLIED"
                                recData.PROSPECT_BUCKET = "1,"
                                recData.APPLIED_BUCKET = "1,"
                            Case "DENY"
                                recData.PROSPECT_BUCKET = "1,"
                                recData.APPLIED_BUCKET = "1,"
                            Case "ADMITTED"
                                recData.PROSPECT_BUCKET = "1,"
                                recData.APPLIED_BUCKET = "1,"
                                recData.ADMITTED_BUCKET = "1,"
                            Case "WITHDRAWN"
                                recData.ADMITTED_BUCKET = "1,"
                                recData.APPLIED_BUCKET = "1,"
                                recData.PROSPECT_BUCKET = "1,"
                            Case "ENROLLED"
                                recData.PROSPECT_BUCKET = "1,"
                                recData.ADMITTED_BUCKET = "1,"
                                recData.APPLIED_BUCKET = "1,"
                                recData.PROSPECT_BUCKET = "1,"                    '
                        End Select


                    End If

                    Exit For
                End If

            Next

        End While

        'CLEANUP
        'IF A BIG BUCKET HAS A 1, THEN MAKE SURE THAT ALL OF THE BUCKETS BELOW IT ALSO HAVE A 1
        'THIS IS TO HANDLE THE ISSUE OF CANDIDATES THAT HAVE JUST AN ENROLL STAGE 
        'this is a catchall for 

        For KOUNT = inxBigBuckets To 2 Step -1

            If aryBigBuckets(KOUNT) = "1" Then
                For SUBKOUNT = 1 To KOUNT

                    If (InStr(aryBigBucketsNAME(SUBKOUNT), "WITHDRAWN") + InStr(aryBigBucketsNAME(SUBKOUNT), "DENY")) = 0 Then 'DON'T ADD TO DENY/WITHDRAWN
                        aryBigBuckets(SUBKOUNT) = 1
                    End If

                Next
                Exit For

            End If


        Next


        bolAtLeastOneStage = False

        For KOUNT = 1 To aryStagesCount
            If aryBuckets(KOUNT) = "0" Then
                strReturn &= ","
            Else
                strReturn &= HANDLE_COMMAS(aryBuckets(KOUNT))
                bolAtLeastOneStage = True
            End If
        Next

        strReturn &= "," 'ADD THE DATA FOR THE OTHER COLUMN

        For KOUNT = 1 To inxBigBuckets

            If aryBigBuckets(KOUNT) = "0" Then
                strReturn &= ","
            Else
                strReturn &= HANDLE_COMMAS(aryBigBuckets(KOUNT))
            End If

        Next

        reader.Close()
        myCommand.Dispose()

        If chkShowCandidatesWithNoStages.Checked Then
            'ok to show
        Else
            If Not bolAtLeastOneStage Then
                strReturn = "DO_NOT_SHOW_CANDIDATE"
                '  strReturn = strNoStages     GO AHEAD AND SHOW THOSE WITH NO STAGES
            End If
        End If

        Return strReturn
    End Function
    Private Sub SHOW_ONLINE_REPORT()
        Dim tmpForm As Form = frmOnlineReport

        tmpForm.ShowDialog()
    End Sub
    Private Function GET_LAST_ORGANIZATION_ADDRESS(ByVal ID_NUM As String, ByVal STUDENT_ID_NUM As String) As String
        Dim strSQL As String = "SELECT TOP 1 ADDR_LINE_1, ADDR_LINE_2, CITY, ISNULL(STATE,' ') AS STATE, '''' + SUBSTRING(ISNULL(ZIP,'      '),1,5) AS ZIP "
        strSQL &= " FROM ADDRESS_MASTER WHERE ADDR_CDE IN ('OCMP','*LHP') AND ID_NUM = " & ID_NUM
        Dim KOUNT As Integer = 0
        Dim strReturn As String = ""
        Dim tmpRegion As String = ""

        Dim tmpAddrLine1 As String = ""
        Dim tmpAddrLine2 As String = ""

        Dim aryStudentAddress(5) As String
        Dim tmpZip As String = ""
        Dim tmpState As String = ""


        '   If Val(ID_NUM) = 0 Then
        ' strReturn = ",,,,,"
        ' Else

        Dim myCommand As New SqlCommand(strSQL, myConnectionGET_HIGH_SCHOOL_ADDRESS)
        myCommand.CommandType = CommandType.Text
        Dim reader As SqlDataReader = myCommand.ExecuteReader

        If reader.HasRows Then
            reader.Read()

            tmpAddrLine1 = HANDLE_COMMAS(reader("ADDR_LINE_1"))
            tmpAddrLine2 = HANDLE_COMMAS(reader("ADDR_LINE_2"))

            If tmpAddrLine1 = Chr(34) & "ATTN:  Accounts Receivable" & Chr(34) & "," Then
                tmpAddrLine1 = tmpAddrLine2
                tmpAddrLine2 = ","
            End If

            If Val(ID_NUM) = 0 Then
                strReturn = ",,,,,"
            Else
                strReturn &= tmpAddrLine1 & tmpAddrLine2
                For KOUNT = 2 To reader.VisibleFieldCount - 1
                    strReturn &= HANDLE_COMMAS(reader(KOUNT))
                Next
            End If


            If Trim(reader("ZIP")) = "'" Then 'if zip is blank, get home address zip. If that's blank, get LO state. If that's blank, get home state. if that's blank, <unknown>
                aryStudentAddress = Split(GET_STUDENT_ADDRESS(STUDENT_ID_NUM), ",")
                tmpZip = Replace(aryStudentAddress(4), Chr(34), "")
                tmpRegion = Trim(reader("STATE"))
                If Len(Trim(tmpRegion)) = 0 Or ID_NUM = 0 Then
                    tmpRegion = Trim(Replace(aryStudentAddress(3), Chr(34), ""))
                End If
            Else
                tmpZip = reader("ZIP")
                tmpRegion = Trim(reader("STATE"))
            End If


            If tmpRegion = "SC" Then
                For KOUNT = 1 To aryRegionsIndex
                    If tmpZip = aryRegions(KOUNT).ZIP_CODE Then
                        tmpRegion = "[" & aryRegions(KOUNT).REGION & "]"
                        Exit For
                    End If
                Next
            End If

            If Len(Trim(tmpRegion)) = 0 Then
                tmpRegion = "<UNKNOWN>"
            End If

            If ID_NUM <> 0 Then
                recData.LAST_ORGANIZATION_ADDRESS_1 = HANDLE_INSERT_DATA(tmpAddrLine1)
                recData.LAST_ORGANIZATION_ADDRESS_2 = HANDLE_INSERT_DATA(Replace(tmpAddrLine2, ",", ""))
                recData.LAST_ORGANIZATION_CITY = HANDLE_INSERT_DATA(reader("CITY"))
                recData.LAST_ORGANIZATION_STATE = HANDLE_INSERT_DATA(reader("STATE"))
                recData.LAST_ORGANIZATION_ZIP = HANDLE_INSERT_DATA(reader("ZIP"))
            Else
                recData.LAST_ORGANIZATION_ADDRESS_1 = HANDLE_INSERT_DATA(" ")
                recData.LAST_ORGANIZATION_ADDRESS_2 = HANDLE_INSERT_DATA(" ")
                recData.LAST_ORGANIZATION_CITY = HANDLE_INSERT_DATA(" ")
                recData.LAST_ORGANIZATION_STATE = HANDLE_INSERT_DATA(" ")
                recData.LAST_ORGANIZATION_ZIP = HANDLE_INSERT_DATA(" ")
            End If

            recData.LAST_ORGANIZATION_REGION = HANDLE_INSERT_DATA(tmpRegion)


        Else 'NO ADDRESS FOUND
            For KOUNT = 0 To reader.VisibleFieldCount - 1
                strReturn &= ","
            Next
        End If

        reader.Close()
        myCommand.Dispose()

        '  End If

        strReturn &= tmpRegion & ","
        Return strReturn
    End Function
    Private Function HANDLE_COMMAS(ByVal tmpString As Object) As String
        If tmpString Is DBNull.Value Then
            tmpString = " "
        End If

        tmpString = Replace(tmpString, "*", "") 'REMOVE ASTERISK ON SOME OF THE HIGH SCHOOL NAMES

        tmpString = Chr(34) & Replace(Trim(tmpString), Chr(34), "") & Chr(34) & "," 'HAD TO REPLACE DOUBLE QUOTES FIRST BECAUSE OF ID:535943 FIRST NAME 
        Return tmpString
    End Function
    Private Function HANDLE_INSERT_DATA(ByVal tmpString As Object, Optional ByVal maxsize As Int16 = 99) As String
        If tmpString Is DBNull.Value Then
            tmpString = " "
        Else
            tmpString = Mid(tmpString, 1, maxsize)
        End If

        tmpString = "'" & Trim(Replace(tmpString, "'", "''")) & "',"



        Return tmpString
    End Function
    Private Function HANDLE_PHONE(ByVal tmpString As Object) As String
        If tmpString Is DBNull.Value Then
            tmpString = " "
        End If

        tmpString = Trim(tmpString)

        If Len(tmpString) = 10 Then
            tmpString = "(" & Mid(tmpString, 1, 3) & ") " & Mid(tmpString, 4, 3) & "-" & Mid(tmpString, 7, 4)
        End If


        tmpString = Chr(34) & Replace(Trim(tmpString), Chr(34), "") & Chr(34) & "," 'HAD TO REPLACE DOUBLE QUOTES FIRST BECAUSE OF ID:535943 FIRST NAME 
        Return tmpString
    End Function
    Private Sub chkAllYears_CheckedChanged(sender As Object, e As EventArgs)

        cboTermStart.Enabled = Not chkAllYears.Checked
        cboTermEnd.Enabled = Not chkAllYears.Checked

    End Sub
    Private Sub txtIDMatch_TextChanged(sender As Object, e As EventArgs)

    End Sub
    Private Sub CLOSE_DATABASE_CONNECTION(ByVal tmpDatabaseConnection As SqlConnection)

        Try
            tmpDatabaseConnection.Close()
        Catch ex As Exception
            'ignore errors
        End Try


    End Sub
    Private Sub HELPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HELPToolStripMenuItem.Click
        Call SSRS_Report()
    End Sub
    Private Sub mnuSTAGES()
        Dim tmpForm As Form = frmStages

        tmpForm.ShowDialog()

    End Sub
    Private Sub STAGESToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles STAGESToolStripMenuItem.Click
        Call mnuSTAGES()
    End Sub
    Private Sub chkDataItems_SelectedIndexChanged(sender As Object, e As EventArgs) Handles chkDataItems.SelectedIndexChanged

        If chkDataItems.SelectedItem = "STAGES" Then
            grpSameMonthDay.Visible = chkDataItems.GetItemCheckState(chkDataItems.SelectedIndex)
            Call MonthsBeforeTerm(Nothing, Nothing)
        End If

    End Sub
    Public Sub Term_Checked_Changed()
        If rdoFall.Checked Then
            rdoFall.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Bold)
            rdoSpring.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Regular)
            rdoAll.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Regular)
        End If

        If rdoSpring.Checked Then
            rdoFall.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Regular)
            rdoSpring.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Bold)
            rdoAll.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Regular)
        End If

        If rdoAll.Checked Then
            rdoFall.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Regular)
            rdoSpring.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Regular)
            rdoAll.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Bold)
        End If

    End Sub
    Private Sub CheckedChanged(sender As Object, e As EventArgs) Handles rdoFall.CheckedChanged, rdoSpring.CheckedChanged, rdoAll.CheckedChanged
        Call Term_Checked_Changed()
    End Sub
    Private Sub chkSummary_CheckedChanged(sender As Object, e As EventArgs) Handles chkSummary.CheckedChanged, cboSummary.SelectedIndexChanged
        If chkSummary.Checked And Len(cboSummary.Text) = 0 Then
            btnRun.Enabled = False
        Else
            btnRun.Enabled = True
        End If
    End Sub
    Private Sub btnSelectStage_Click(sender As Object, e As EventArgs) Handles btnSelectStage.Click
        If btnSelectStage.Text = "UNCHECK ALL STAGES" Then
            btnSelectStage.Text = "CHECK ALL STAGES"
            For i As Integer = 0 To chkStage.Items.Count - 1
                chkStage.SetItemChecked(i, False)
            Next
        Else
            btnSelectStage.Text = "UNCHECK ALL STAGES"
            For i As Integer = 0 To chkStage.Items.Count - 1
                chkStage.SetItemChecked(i, True)
            Next
        End If
    End Sub
    Private Sub chkOnlineReport_CheckedChanged(sender As Object, e As EventArgs) Handles chkOnlineReport.CheckedChanged

    End Sub

    Private Sub btnClearAllFilters_Click(sender As Object, e As EventArgs) Handles btnClearAllFilters.Click
        txtIDMatch.Text = ""
        txtPhoneMatch.Text = ""
        txtNameMatch.Text = ""
        cboCounselor.Text = ""
        cboHS.Text = ""
        txtHomeZip.Text = ""
        txtHSZIP.Text = ""
    End Sub

    Private Sub chkAllDates_CheckedChanged(sender As Object, e As EventArgs) Handles chkAllDates.CheckedChanged
        cboTargetTerm.Enabled = Not chkAllDates.Checked
        lblTargetDate.Enabled = Not chkAllDates.Checked
        lblTargetTerm.Enabled = Not chkAllDates.Checked
        dteTargetDate.Enabled = Not chkAllDates.Checked
        chkShowCandidatesWithNoStages.Enabled = Not chkAllDates.Checked

        If chkAllDates.Checked Then
            lblMonthsBeforeTerm.Text = ""
        Else
            Call MonthsBeforeTerm(Nothing, Nothing)
        End If

    End Sub

    Private Sub MonthsBeforeTerm(sender As Object, e As EventArgs) Handles dteTargetDate.ValueChanged, cboTargetTerm.SelectedIndexChanged
        lblMonthsBeforeTerm.Text = "MONTHS BEFORE TERM: " & DateDiff(DateInterval.Month, dteTargetDate.Value, CDate("08/01/" & cboTargetTerm.Text))
    End Sub

End Class
