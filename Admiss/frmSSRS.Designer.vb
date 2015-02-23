<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSSRS
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
        Me.ReportViewer1 = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.cboReports = New System.Windows.Forms.ComboBox()
        Me.cboPath = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pboxLogo = New System.Windows.Forms.PictureBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FILEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RETURNTOHOMEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RETURNHOMEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.txtAutoReport = New System.Windows.Forms.TextBox()
        Me.txtprmAsOf = New System.Windows.Forms.TextBox()
        CType(Me.pboxLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ReportViewer1
        '
        Me.ReportViewer1.Location = New System.Drawing.Point(4, 57)
        Me.ReportViewer1.Name = "ReportViewer1"
        Me.ReportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Remote
        Me.ReportViewer1.ServerReport.ReportPath = "/Admissions/Regions"
        Me.ReportViewer1.ServerReport.ReportServerUrl = New System.Uri("http://ncfuturesql/reportserver", System.UriKind.Absolute)
        Me.ReportViewer1.Size = New System.Drawing.Size(832, 515)
        Me.ReportViewer1.TabIndex = 0
        '
        'cboReports
        '
        Me.cboReports.FormattingEnabled = True
        Me.cboReports.Location = New System.Drawing.Point(66, 25)
        Me.cboReports.Name = "cboReports"
        Me.cboReports.Size = New System.Drawing.Size(188, 21)
        Me.cboReports.TabIndex = 1
        '
        'cboPath
        '
        Me.cboPath.FormattingEnabled = True
        Me.cboPath.Location = New System.Drawing.Point(549, 25)
        Me.cboPath.Name = "cboPath"
        Me.cboPath.Size = New System.Drawing.Size(79, 21)
        Me.cboPath.TabIndex = 2
        Me.cboPath.Text = "* HIDDEN *"
        Me.cboPath.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(1, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "REPORTS"
        '
        'pboxLogo
        '
        Me.pboxLogo.Image = Global.WindowsApplication1.My.Resources.Resources.Logo
        Me.pboxLogo.InitialImage = Global.WindowsApplication1.My.Resources.Resources.Logo
        Me.pboxLogo.Location = New System.Drawing.Point(740, -2)
        Me.pboxLogo.Name = "pboxLogo"
        Me.pboxLogo.Size = New System.Drawing.Size(96, 48)
        Me.pboxLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pboxLogo.TabIndex = 26
        Me.pboxLogo.TabStop = False
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.Gainsboro
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FILEToolStripMenuItem, Me.RETURNHOMEToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(839, 24)
        Me.MenuStrip1.TabIndex = 27
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FILEToolStripMenuItem
        '
        Me.FILEToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RETURNTOHOMEToolStripMenuItem})
        Me.FILEToolStripMenuItem.Name = "FILEToolStripMenuItem"
        Me.FILEToolStripMenuItem.Size = New System.Drawing.Size(40, 20)
        Me.FILEToolStripMenuItem.Text = "FILE"
        '
        'RETURNTOHOMEToolStripMenuItem
        '
        Me.RETURNTOHOMEToolStripMenuItem.Name = "RETURNTOHOMEToolStripMenuItem"
        Me.RETURNTOHOMEToolStripMenuItem.Size = New System.Drawing.Size(175, 22)
        Me.RETURNTOHOMEToolStripMenuItem.Text = "RETURN TO HOME"
        '
        'RETURNHOMEToolStripMenuItem
        '
        Me.RETURNHOMEToolStripMenuItem.Name = "RETURNHOMEToolStripMenuItem"
        Me.RETURNHOMEToolStripMenuItem.Size = New System.Drawing.Size(101, 20)
        Me.RETURNHOMEToolStripMenuItem.Text = "RETURN HOME"
        '
        'txtAutoReport
        '
        Me.txtAutoReport.Location = New System.Drawing.Point(329, 28)
        Me.txtAutoReport.Name = "txtAutoReport"
        Me.txtAutoReport.Size = New System.Drawing.Size(97, 20)
        Me.txtAutoReport.TabIndex = 28
        Me.txtAutoReport.Visible = False
        '
        'txtprmAsOf
        '
        Me.txtprmAsOf.Location = New System.Drawing.Point(432, 28)
        Me.txtprmAsOf.Name = "txtprmAsOf"
        Me.txtprmAsOf.Size = New System.Drawing.Size(97, 20)
        Me.txtprmAsOf.TabIndex = 29
        '
        'frmSSRS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(839, 575)
        Me.Controls.Add(Me.txtprmAsOf)
        Me.Controls.Add(Me.txtAutoReport)
        Me.Controls.Add(Me.pboxLogo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboPath)
        Me.Controls.Add(Me.cboReports)
        Me.Controls.Add(Me.ReportViewer1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmSSRS"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SSRS REPORTING"
        CType(Me.pboxLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ReportViewer1 As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents cboReports As System.Windows.Forms.ComboBox
    Friend WithEvents cboPath As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pboxLogo As System.Windows.Forms.PictureBox
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FILEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RETURNTOHOMEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RETURNHOMEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents txtAutoReport As System.Windows.Forms.TextBox
    Public WithEvents txtprmAsOf As System.Windows.Forms.TextBox
End Class
