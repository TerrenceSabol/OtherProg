<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAdmiss
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAdmiss))
        Me.btnRun = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FILEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HELPToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.STAGESToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.lblPercentDone = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.txtProgress = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnClearAllFilters = New System.Windows.Forms.Button()
        Me.chkOnlineReport = New System.Windows.Forms.CheckBox()
        Me.txtHomeZip = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtHSZIP = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.cboHS = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboCounselor = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtPhoneMatch = New System.Windows.Forms.TextBox()
        Me.txtNameMatch = New System.Windows.Forms.TextBox()
        Me.txtIDMatch = New System.Windows.Forms.TextBox()
        Me.grpSameMonthDay = New System.Windows.Forms.GroupBox()
        Me.lblMonthsBeforeTerm = New System.Windows.Forms.Label()
        Me.chkAllDates = New System.Windows.Forms.CheckBox()
        Me.chkShowCandidatesWithNoStages = New System.Windows.Forms.CheckBox()
        Me.dteTargetDate = New System.Windows.Forms.DateTimePicker()
        Me.lblTargetDate = New System.Windows.Forms.Label()
        Me.lblTargetTerm = New System.Windows.Forms.Label()
        Me.cboTargetTerm = New System.Windows.Forms.ComboBox()
        Me.grpTermRange = New System.Windows.Forms.GroupBox()
        Me.rdoAll = New System.Windows.Forms.RadioButton()
        Me.rdoSpring = New System.Windows.Forms.RadioButton()
        Me.rdoFall = New System.Windows.Forms.RadioButton()
        Me.chkAllYears = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboTermEnd = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboTermStart = New System.Windows.Forms.ComboBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkDataItems = New System.Windows.Forms.CheckedListBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.chkType = New System.Windows.Forms.CheckedListBox()
        Me.grpSummary = New System.Windows.Forms.GroupBox()
        Me.cboSummary = New System.Windows.Forms.ComboBox()
        Me.chkSummary = New System.Windows.Forms.CheckBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.btnSelectStage = New System.Windows.Forms.Button()
        Me.chkStage = New System.Windows.Forms.CheckedListBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MenuStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.grpSameMonthDay.SuspendLayout()
        Me.grpTermRange.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.grpSummary.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(485, 451)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(93, 31)
        Me.btnRun.TabIndex = 0
        Me.btnRun.Text = "&Run"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.Gainsboro
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FILEToolStripMenuItem, Me.HELPToolStripMenuItem, Me.STAGESToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(579, 24)
        Me.MenuStrip1.TabIndex = 9
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FILEToolStripMenuItem
        '
        Me.FILEToolStripMenuItem.Name = "FILEToolStripMenuItem"
        Me.FILEToolStripMenuItem.Size = New System.Drawing.Size(40, 20)
        Me.FILEToolStripMenuItem.Text = "FILE"
        '
        'HELPToolStripMenuItem
        '
        Me.HELPToolStripMenuItem.Name = "HELPToolStripMenuItem"
        Me.HELPToolStripMenuItem.Size = New System.Drawing.Size(99, 20)
        Me.HELPToolStripMenuItem.Text = "REPORT MENU"
        '
        'STAGESToolStripMenuItem
        '
        Me.STAGESToolStripMenuItem.Name = "STAGESToolStripMenuItem"
        Me.STAGESToolStripMenuItem.Size = New System.Drawing.Size(60, 20)
        Me.STAGESToolStripMenuItem.Text = "STAGES"
        '
        'lblPercentDone
        '
        Me.lblPercentDone.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.lblPercentDone.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblPercentDone.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPercentDone.Location = New System.Drawing.Point(-1, 451)
        Me.lblPercentDone.Name = "lblPercentDone"
        Me.lblPercentDone.Size = New System.Drawing.Size(495, 31)
        Me.lblPercentDone.TabIndex = 12
        Me.lblPercentDone.Text = "lblPercentDone"
        Me.lblPercentDone.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(6, 27)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(21, 13)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "ID "
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(6, 53)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(38, 13)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "NAME"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(6, 79)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(51, 13)
        Me.Label7.TabIndex = 17
        Me.Label7.Text = "PHONE  "
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(2, 481)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(577, 22)
        Me.ProgressBar1.TabIndex = 21
        '
        'txtProgress
        '
        Me.txtProgress.Location = New System.Drawing.Point(170, 298)
        Me.txtProgress.Name = "txtProgress"
        Me.txtProgress.Size = New System.Drawing.Size(93, 20)
        Me.txtProgress.TabIndex = 22
        Me.txtProgress.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtProgress.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnClearAllFilters)
        Me.GroupBox1.Controls.Add(Me.chkOnlineReport)
        Me.GroupBox1.Controls.Add(Me.txtHomeZip)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.txtHSZIP)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.cboHS)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.cboCounselor)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtPhoneMatch)
        Me.GroupBox1.Controls.Add(Me.txtNameMatch)
        Me.GroupBox1.Controls.Add(Me.txtIDMatch)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(3, 346)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(576, 105)
        Me.GroupBox1.TabIndex = 23
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "F I L T E R"
        '
        'btnClearAllFilters
        '
        Me.btnClearAllFilters.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClearAllFilters.Location = New System.Drawing.Point(101, -1)
        Me.btnClearAllFilters.Name = "btnClearAllFilters"
        Me.btnClearAllFilters.Size = New System.Drawing.Size(128, 20)
        Me.btnClearAllFilters.TabIndex = 37
        Me.btnClearAllFilters.Text = "CLEAR ALL FILTERS"
        Me.btnClearAllFilters.UseVisualStyleBackColor = True
        '
        'chkOnlineReport
        '
        Me.chkOnlineReport.AutoSize = True
        Me.chkOnlineReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOnlineReport.Location = New System.Drawing.Point(450, 88)
        Me.chkOnlineReport.Name = "chkOnlineReport"
        Me.chkOnlineReport.Size = New System.Drawing.Size(121, 17)
        Me.chkOnlineReport.TabIndex = 35
        Me.chkOnlineReport.Text = "ONLINE REPORTS"
        Me.chkOnlineReport.UseVisualStyleBackColor = True
        '
        'txtHomeZip
        '
        Me.txtHomeZip.Location = New System.Drawing.Point(489, 49)
        Me.txtHomeZip.Name = "txtHomeZip"
        Me.txtHomeZip.Size = New System.Drawing.Size(80, 20)
        Me.txtHomeZip.TabIndex = 36
        Me.ToolTip1.SetToolTip(Me.txtHomeZip, "Enter Zipcodes (full or partial) separated by commas" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "For Example: 293,29201")
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(390, 53)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(59, 13)
        Me.Label10.TabIndex = 35
        Me.Label10.Text = "HOME ZIP"
        '
        'txtHSZIP
        '
        Me.txtHSZIP.Location = New System.Drawing.Point(489, 24)
        Me.txtHSZIP.Name = "txtHSZIP"
        Me.txtHSZIP.Size = New System.Drawing.Size(80, 20)
        Me.txtHSZIP.TabIndex = 34
        Me.ToolTip1.SetToolTip(Me.txtHSZIP, "Enter Zipcodes (full or partial) separated by commas" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "For Example: 293,29201")
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(390, 27)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(101, 13)
        Me.Label8.TabIndex = 33
        Me.Label8.Text = "HIGH SCHOOL ZIP"
        '
        'cboHS
        '
        Me.cboHS.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboHS.FormattingEnabled = True
        Me.cboHS.Location = New System.Drawing.Point(237, 50)
        Me.cboHS.Name = "cboHS"
        Me.cboHS.Size = New System.Drawing.Size(142, 21)
        Me.cboHS.TabIndex = 32
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(152, 53)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(81, 13)
        Me.Label4.TabIndex = 31
        Me.Label4.Text = "HIGH SCHOOL"
        '
        'cboCounselor
        '
        Me.cboCounselor.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCounselor.FormattingEnabled = True
        Me.cboCounselor.Location = New System.Drawing.Point(237, 24)
        Me.cboCounselor.Name = "cboCounselor"
        Me.cboCounselor.Size = New System.Drawing.Size(142, 21)
        Me.cboCounselor.TabIndex = 30
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(152, 27)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 13)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "COUNSELOR"
        '
        'txtPhoneMatch
        '
        Me.txtPhoneMatch.Location = New System.Drawing.Point(55, 77)
        Me.txtPhoneMatch.Name = "txtPhoneMatch"
        Me.txtPhoneMatch.Size = New System.Drawing.Size(91, 20)
        Me.txtPhoneMatch.TabIndex = 19
        '
        'txtNameMatch
        '
        Me.txtNameMatch.Location = New System.Drawing.Point(55, 52)
        Me.txtNameMatch.Name = "txtNameMatch"
        Me.txtNameMatch.Size = New System.Drawing.Size(91, 20)
        Me.txtNameMatch.TabIndex = 17
        '
        'txtIDMatch
        '
        Me.txtIDMatch.Location = New System.Drawing.Point(55, 27)
        Me.txtIDMatch.Name = "txtIDMatch"
        Me.txtIDMatch.Size = New System.Drawing.Size(91, 20)
        Me.txtIDMatch.TabIndex = 15
        '
        'grpSameMonthDay
        '
        Me.grpSameMonthDay.Controls.Add(Me.lblMonthsBeforeTerm)
        Me.grpSameMonthDay.Controls.Add(Me.chkAllDates)
        Me.grpSameMonthDay.Controls.Add(Me.chkShowCandidatesWithNoStages)
        Me.grpSameMonthDay.Controls.Add(Me.dteTargetDate)
        Me.grpSameMonthDay.Controls.Add(Me.lblTargetDate)
        Me.grpSameMonthDay.Controls.Add(Me.lblTargetTerm)
        Me.grpSameMonthDay.Controls.Add(Me.cboTargetTerm)
        Me.grpSameMonthDay.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpSameMonthDay.Location = New System.Drawing.Point(246, 27)
        Me.grpSameMonthDay.Name = "grpSameMonthDay"
        Me.grpSameMonthDay.Size = New System.Drawing.Size(230, 139)
        Me.grpSameMonthDay.TabIndex = 29
        Me.grpSameMonthDay.TabStop = False
        Me.grpSameMonthDay.Text = "STAGE COMPARISON"
        Me.grpSameMonthDay.Visible = False
        '
        'lblMonthsBeforeTerm
        '
        Me.lblMonthsBeforeTerm.AutoSize = True
        Me.lblMonthsBeforeTerm.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMonthsBeforeTerm.Location = New System.Drawing.Point(20, 81)
        Me.lblMonthsBeforeTerm.Name = "lblMonthsBeforeTerm"
        Me.lblMonthsBeforeTerm.Size = New System.Drawing.Size(140, 13)
        Me.lblMonthsBeforeTerm.TabIndex = 37
        Me.lblMonthsBeforeTerm.Text = "MONTHS BEFORE TERM: "
        '
        'chkAllDates
        '
        Me.chkAllDates.AutoSize = True
        Me.chkAllDates.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAllDates.Location = New System.Drawing.Point(150, 0)
        Me.chkAllDates.Name = "chkAllDates"
        Me.chkAllDates.Size = New System.Drawing.Size(84, 17)
        Me.chkAllDates.TabIndex = 36
        Me.chkAllDates.Text = "ALL DATES"
        Me.chkAllDates.UseVisualStyleBackColor = True
        '
        'chkShowCandidatesWithNoStages
        '
        Me.chkShowCandidatesWithNoStages.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkShowCandidatesWithNoStages.Location = New System.Drawing.Point(23, 100)
        Me.chkShowCandidatesWithNoStages.Name = "chkShowCandidatesWithNoStages"
        Me.chkShowCandidatesWithNoStages.Size = New System.Drawing.Size(170, 33)
        Me.chkShowCandidatesWithNoStages.TabIndex = 34
        Me.chkShowCandidatesWithNoStages.Text = "SHOW CANDIDATES WITH NO STAGES IN RANGE"
        Me.chkShowCandidatesWithNoStages.UseVisualStyleBackColor = True
        '
        'dteTargetDate
        '
        Me.dteTargetDate.CustomFormat = "MM/dd/yyyy"
        Me.dteTargetDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dteTargetDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dteTargetDate.Location = New System.Drawing.Point(115, 50)
        Me.dteTargetDate.Name = "dteTargetDate"
        Me.dteTargetDate.Size = New System.Drawing.Size(78, 20)
        Me.dteTargetDate.TabIndex = 32
        '
        'lblTargetDate
        '
        Me.lblTargetDate.AutoSize = True
        Me.lblTargetDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTargetDate.Location = New System.Drawing.Point(20, 53)
        Me.lblTargetDate.Name = "lblTargetDate"
        Me.lblTargetDate.Size = New System.Drawing.Size(81, 13)
        Me.lblTargetDate.TabIndex = 31
        Me.lblTargetDate.Text = "CUTOFF DATE"
        '
        'lblTargetTerm
        '
        Me.lblTargetTerm.AutoSize = True
        Me.lblTargetTerm.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTargetTerm.Location = New System.Drawing.Point(20, 25)
        Me.lblTargetTerm.Name = "lblTargetTerm"
        Me.lblTargetTerm.Size = New System.Drawing.Size(85, 13)
        Me.lblTargetTerm.TabIndex = 30
        Me.lblTargetTerm.Text = "TARGET TERM"
        '
        'cboTargetTerm
        '
        Me.cboTargetTerm.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTargetTerm.FormattingEnabled = True
        Me.cboTargetTerm.Location = New System.Drawing.Point(115, 21)
        Me.cboTargetTerm.Name = "cboTargetTerm"
        Me.cboTargetTerm.Size = New System.Drawing.Size(65, 21)
        Me.cboTargetTerm.TabIndex = 29
        '
        'grpTermRange
        '
        Me.grpTermRange.Controls.Add(Me.rdoAll)
        Me.grpTermRange.Controls.Add(Me.rdoSpring)
        Me.grpTermRange.Controls.Add(Me.rdoFall)
        Me.grpTermRange.Controls.Add(Me.chkAllYears)
        Me.grpTermRange.Controls.Add(Me.Label2)
        Me.grpTermRange.Controls.Add(Me.cboTermEnd)
        Me.grpTermRange.Controls.Add(Me.Label1)
        Me.grpTermRange.Controls.Add(Me.cboTermStart)
        Me.grpTermRange.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpTermRange.Location = New System.Drawing.Point(0, 27)
        Me.grpTermRange.Name = "grpTermRange"
        Me.grpTermRange.Size = New System.Drawing.Size(245, 115)
        Me.grpTermRange.TabIndex = 30
        Me.grpTermRange.TabStop = False
        Me.grpTermRange.Text = "YEAR RANGE"
        '
        'rdoAll
        '
        Me.rdoAll.AutoSize = True
        Me.rdoAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoAll.Location = New System.Drawing.Point(170, 53)
        Me.rdoAll.Name = "rdoAll"
        Me.rdoAll.Size = New System.Drawing.Size(44, 17)
        Me.rdoAll.TabIndex = 39
        Me.rdoAll.TabStop = True
        Me.rdoAll.Text = "ALL"
        Me.rdoAll.UseVisualStyleBackColor = True
        '
        'rdoSpring
        '
        Me.rdoSpring.AutoSize = True
        Me.rdoSpring.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoSpring.Location = New System.Drawing.Point(170, 32)
        Me.rdoSpring.Name = "rdoSpring"
        Me.rdoSpring.Size = New System.Drawing.Size(66, 17)
        Me.rdoSpring.TabIndex = 38
        Me.rdoSpring.TabStop = True
        Me.rdoSpring.Text = "SPRING"
        Me.rdoSpring.UseVisualStyleBackColor = True
        '
        'rdoFall
        '
        Me.rdoFall.AutoSize = True
        Me.rdoFall.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoFall.Location = New System.Drawing.Point(170, 11)
        Me.rdoFall.Name = "rdoFall"
        Me.rdoFall.Size = New System.Drawing.Size(50, 17)
        Me.rdoFall.TabIndex = 37
        Me.rdoFall.TabStop = True
        Me.rdoFall.Text = "FALL"
        Me.rdoFall.UseVisualStyleBackColor = True
        '
        'chkAllYears
        '
        Me.chkAllYears.AutoSize = True
        Me.chkAllYears.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAllYears.Location = New System.Drawing.Point(15, 81)
        Me.chkAllYears.Name = "chkAllYears"
        Me.chkAllYears.Size = New System.Drawing.Size(84, 17)
        Me.chkAllYears.TabIndex = 24
        Me.chkAllYears.Text = "ALL YEARS"
        Me.chkAllYears.UseVisualStyleBackColor = True
        Me.chkAllYears.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(35, 53)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "END"
        '
        'cboTermEnd
        '
        Me.cboTermEnd.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTermEnd.FormattingEnabled = True
        Me.cboTermEnd.Location = New System.Drawing.Point(84, 49)
        Me.cboTermEnd.Name = "cboTermEnd"
        Me.cboTermEnd.Size = New System.Drawing.Size(65, 21)
        Me.cboTermEnd.TabIndex = 22
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(35, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "START"
        '
        'cboTermStart
        '
        Me.cboTermStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTermStart.FormattingEnabled = True
        Me.cboTermStart.Location = New System.Drawing.Point(84, 21)
        Me.cboTermStart.Name = "cboTermStart"
        Me.cboTermStart.Size = New System.Drawing.Size(65, 21)
        Me.cboTermStart.TabIndex = 20
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkDataItems)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(2, 172)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(190, 168)
        Me.GroupBox2.TabIndex = 31
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "A D D I T I O N A L   D A T A"
        '
        'chkDataItems
        '
        Me.chkDataItems.CheckOnClick = True
        Me.chkDataItems.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDataItems.FormattingEnabled = True
        Me.chkDataItems.Location = New System.Drawing.Point(15, 19)
        Me.chkDataItems.Name = "chkDataItems"
        Me.chkDataItems.Size = New System.Drawing.Size(157, 94)
        Me.chkDataItems.TabIndex = 2
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.chkType)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(192, 172)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(158, 168)
        Me.GroupBox3.TabIndex = 32
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "T Y P E"
        '
        'chkType
        '
        Me.chkType.CheckOnClick = True
        Me.chkType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkType.FormattingEnabled = True
        Me.chkType.Location = New System.Drawing.Point(22, 19)
        Me.chkType.Name = "chkType"
        Me.chkType.Size = New System.Drawing.Size(120, 94)
        Me.chkType.TabIndex = 7
        '
        'grpSummary
        '
        Me.grpSummary.Controls.Add(Me.cboSummary)
        Me.grpSummary.Controls.Add(Me.chkSummary)
        Me.grpSummary.Location = New System.Drawing.Point(774, 31)
        Me.grpSummary.Name = "grpSummary"
        Me.grpSummary.Size = New System.Drawing.Size(152, 45)
        Me.grpSummary.TabIndex = 33
        Me.grpSummary.TabStop = False
        Me.grpSummary.Visible = False
        '
        'cboSummary
        '
        Me.cboSummary.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSummary.FormattingEnabled = True
        Me.cboSummary.Location = New System.Drawing.Point(33, 19)
        Me.cboSummary.Name = "cboSummary"
        Me.cboSummary.Size = New System.Drawing.Size(104, 21)
        Me.cboSummary.TabIndex = 35
        '
        'chkSummary
        '
        Me.chkSummary.AutoSize = True
        Me.chkSummary.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSummary.Location = New System.Drawing.Point(10, 0)
        Me.chkSummary.Name = "chkSummary"
        Me.chkSummary.Size = New System.Drawing.Size(112, 17)
        Me.chkSummary.TabIndex = 34
        Me.chkSummary.Text = "S U M M A R Y"
        Me.chkSummary.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.btnSelectStage)
        Me.GroupBox4.Controls.Add(Me.chkStage)
        Me.GroupBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(350, 172)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(186, 168)
        Me.GroupBox4.TabIndex = 34
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "C U R R E N T   S T A G E"
        '
        'btnSelectStage
        '
        Me.btnSelectStage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelectStage.Location = New System.Drawing.Point(23, 146)
        Me.btnSelectStage.Name = "btnSelectStage"
        Me.btnSelectStage.Size = New System.Drawing.Size(157, 20)
        Me.btnSelectStage.TabIndex = 8
        Me.btnSelectStage.Text = "UNCHECK ALL STAGES"
        Me.btnSelectStage.UseVisualStyleBackColor = True
        '
        'chkStage
        '
        Me.chkStage.CheckOnClick = True
        Me.chkStage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkStage.FormattingEnabled = True
        Me.chkStage.Location = New System.Drawing.Point(22, 19)
        Me.chkStage.Name = "chkStage"
        Me.chkStage.Size = New System.Drawing.Size(158, 124)
        Me.chkStage.TabIndex = 7
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.Logo
        Me.PictureBox1.Location = New System.Drawing.Point(453, 27)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(150, 62)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 35
        Me.PictureBox1.TabStop = False
        '
        'frmAdmiss
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(579, 505)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.grpSummary)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.grpTermRange)
        Me.Controls.Add(Me.grpSameMonthDay)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.txtProgress)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.lblPercentDone)
        Me.Controls.Add(Me.PictureBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmAdmiss"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ADMISS"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.grpSameMonthDay.ResumeLayout(False)
        Me.grpSameMonthDay.PerformLayout()
        Me.grpTermRange.ResumeLayout(False)
        Me.grpTermRange.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.grpSummary.ResumeLayout(False)
        Me.grpSummary.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnRun As System.Windows.Forms.Button
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FILEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HELPToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblPercentDone As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents txtProgress As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtNameMatch As System.Windows.Forms.TextBox
    Friend WithEvents txtIDMatch As System.Windows.Forms.TextBox
    Friend WithEvents txtPhoneMatch As System.Windows.Forms.TextBox
    Friend WithEvents STAGESToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents grpSameMonthDay As System.Windows.Forms.GroupBox
    Friend WithEvents lblTargetTerm As System.Windows.Forms.Label
    Friend WithEvents cboTargetTerm As System.Windows.Forms.ComboBox
    Friend WithEvents dteTargetDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblTargetDate As System.Windows.Forms.Label
    Friend WithEvents chkShowCandidatesWithNoStages As System.Windows.Forms.CheckBox
    Friend WithEvents grpTermRange As System.Windows.Forms.GroupBox
    Friend WithEvents chkAllYears As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboTermEnd As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboTermStart As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents chkDataItems As System.Windows.Forms.CheckedListBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents chkType As System.Windows.Forms.CheckedListBox
    Friend WithEvents rdoAll As System.Windows.Forms.RadioButton
    Friend WithEvents rdoSpring As System.Windows.Forms.RadioButton
    Friend WithEvents rdoFall As System.Windows.Forms.RadioButton
    Friend WithEvents grpSummary As System.Windows.Forms.GroupBox
    Friend WithEvents cboSummary As System.Windows.Forms.ComboBox
    Friend WithEvents chkSummary As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents chkStage As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnSelectStage As System.Windows.Forms.Button
    Friend WithEvents chkOnlineReport As System.Windows.Forms.CheckBox
    Friend WithEvents cboCounselor As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtHSZIP As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cboHS As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtHomeZip As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnClearAllFilters As System.Windows.Forms.Button
    Friend WithEvents chkAllDates As System.Windows.Forms.CheckBox
    Friend WithEvents lblMonthsBeforeTerm As System.Windows.Forms.Label

End Class
