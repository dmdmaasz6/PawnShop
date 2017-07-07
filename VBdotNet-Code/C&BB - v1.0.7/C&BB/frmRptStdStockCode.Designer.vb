<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRptStdStockCode
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRptStdStockCode))
        Me.PawnStockBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.CBB_DataSet = New CashBuyBack.CBB_DataSet()
        Me.grpAdd = New System.Windows.Forms.GroupBox()
        Me.dgStockCode = New System.Windows.Forms.DataGridView()
        Me.Column11 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        CType(Me.PawnStockBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CBB_DataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpAdd.SuspendLayout()
        CType(Me.dgStockCode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PawnStockBindingSource
        '
        Me.PawnStockBindingSource.DataMember = "PawnStock"
        Me.PawnStockBindingSource.DataSource = Me.CBB_DataSet
        '
        'CBB_DataSet
        '
        Me.CBB_DataSet.DataSetName = "CBB_DataSet"
        Me.CBB_DataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'grpAdd
        '
        Me.grpAdd.Controls.Add(Me.dgStockCode)
        Me.grpAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpAdd.Location = New System.Drawing.Point(12, 12)
        Me.grpAdd.Name = "grpAdd"
        Me.grpAdd.Size = New System.Drawing.Size(370, 491)
        Me.grpAdd.TabIndex = 14
        Me.grpAdd.TabStop = False
        Me.grpAdd.Text = " Standard Stock Code List: "
        '
        'dgStockCode
        '
        Me.dgStockCode.AllowUserToAddRows = False
        Me.dgStockCode.AllowUserToDeleteRows = False
        Me.dgStockCode.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgStockCode.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column11, Me.Column3})
        Me.dgStockCode.Location = New System.Drawing.Point(9, 22)
        Me.dgStockCode.Name = "dgStockCode"
        Me.dgStockCode.ReadOnly = True
        Me.dgStockCode.RowHeadersVisible = False
        Me.dgStockCode.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgStockCode.Size = New System.Drawing.Size(353, 457)
        Me.dgStockCode.TabIndex = 1
        '
        'Column11
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Column11.DefaultCellStyle = DataGridViewCellStyle1
        Me.Column11.HeaderText = "NS ID"
        Me.Column11.Name = "Column11"
        Me.Column11.ReadOnly = True
        Me.Column11.Width = 75
        '
        'Column3
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Column3.DefaultCellStyle = DataGridViewCellStyle2
        Me.Column3.HeaderText = "Description"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        Me.Column3.Width = 250
        '
        'cmdPrint
        '
        Me.cmdPrint.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPrint.Location = New System.Drawing.Point(80, 523)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(86, 40)
        Me.cmdPrint.TabIndex = 13
        Me.cmdPrint.Text = "&Export ..."
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.Location = New System.Drawing.Point(225, 523)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(86, 40)
        Me.cmdExit.TabIndex = 12
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'frmRptStdStockCode
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdExit
        Me.ClientSize = New System.Drawing.Size(395, 581)
        Me.Controls.Add(Me.grpAdd)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.cmdExit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmRptStdStockCode"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Cash & BuyBack - Standard Stock"
        CType(Me.PawnStockBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CBB_DataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpAdd.ResumeLayout(False)
        CType(Me.dgStockCode, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PawnStockBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents CBB_DataSet As CashBuyBack.CBB_DataSet
    Friend WithEvents grpAdd As System.Windows.Forms.GroupBox
    Friend WithEvents dgStockCode As System.Windows.Forms.DataGridView
    Friend WithEvents cmdPrint As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents Column11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
