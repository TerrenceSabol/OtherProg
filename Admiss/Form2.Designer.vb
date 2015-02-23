<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStages
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmStages))
        Me.NCAdmissionsStageCategoryBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.TmsEPrdDataSet = New WindowsApplication1.TmsEPrdDataSet()
        Me.NC_Admissions_Stage_CategoryTableAdapter = New WindowsApplication1.TmsEPrdDataSetTableAdapters.NC_Admissions_Stage_CategoryTableAdapter()
        Me.dgvStages = New System.Windows.Forms.DataGridView()
        CType(Me.NCAdmissionsStageCategoryBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TmsEPrdDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvStages, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'NCAdmissionsStageCategoryBindingSource
        '
        Me.NCAdmissionsStageCategoryBindingSource.DataMember = "NC_Admissions_Stage_Category"
        Me.NCAdmissionsStageCategoryBindingSource.DataSource = Me.TmsEPrdDataSet
        '
        'TmsEPrdDataSet
        '
        Me.TmsEPrdDataSet.DataSetName = "TmsEPrdDataSet"
        Me.TmsEPrdDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'NC_Admissions_Stage_CategoryTableAdapter
        '
        Me.NC_Admissions_Stage_CategoryTableAdapter.ClearBeforeFill = True
        '
        'dgvStages
        '
        Me.dgvStages.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvStages.Location = New System.Drawing.Point(12, 12)
        Me.dgvStages.Name = "dgvStages"
        Me.dgvStages.Size = New System.Drawing.Size(347, 568)
        Me.dgvStages.TabIndex = 0
        '
        'frmStages
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(394, 762)
        Me.Controls.Add(Me.dgvStages)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmStages"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "STAGES"
        CType(Me.NCAdmissionsStageCategoryBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TmsEPrdDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvStages, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TmsEPrdDataSet As WindowsApplication1.TmsEPrdDataSet
    Friend WithEvents NCAdmissionsStageCategoryBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents NC_Admissions_Stage_CategoryTableAdapter As WindowsApplication1.TmsEPrdDataSetTableAdapters.NC_Admissions_Stage_CategoryTableAdapter
    Friend WithEvents dgvStages As System.Windows.Forms.DataGridView
End Class
