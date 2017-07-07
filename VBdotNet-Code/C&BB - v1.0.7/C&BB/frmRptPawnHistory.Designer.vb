<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRptPawnHistory
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
        Me.dgPawnHistory = New System.Windows.Forms.DataGridView()
        Me.CBBDataSetBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.CBB_DataSet = New CashBuyBack.CBB_DataSet()
        Me.ListView1 = New System.Windows.Forms.ListView()
        CType(Me.dgPawnHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CBBDataSetBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CBB_DataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgPawnHistory
        '
        Me.dgPawnHistory.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgPawnHistory.Location = New System.Drawing.Point(36, 12)
        Me.dgPawnHistory.Name = "dgPawnHistory"
        Me.dgPawnHistory.Size = New System.Drawing.Size(270, 94)
        Me.dgPawnHistory.TabIndex = 0
        '
        'CBBDataSetBindingSource
        '
        Me.CBBDataSetBindingSource.DataSource = Me.CBB_DataSet
        Me.CBBDataSetBindingSource.Position = 0
        '
        'CBB_DataSet
        '
        Me.CBB_DataSet.DataSetName = "CBB_DataSet"
        Me.CBB_DataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ListView1
        '
        Me.ListView1.FullRowSelect = True
        Me.ListView1.GridLines = True
        Me.ListView1.Location = New System.Drawing.Point(325, 19)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(270, 87)
        Me.ListView1.TabIndex = 1
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'frmRptPawnHistory
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1128, 703)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.dgPawnHistory)
        Me.Name = "frmRptPawnHistory"
        Me.Text = "frmRptPawnHistory"
        CType(Me.dgPawnHistory, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CBBDataSetBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CBB_DataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgPawnHistory As System.Windows.Forms.DataGridView
    Friend WithEvents CBBDataSetBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents CBB_DataSet As CashBuyBack.CBB_DataSet
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
End Class
