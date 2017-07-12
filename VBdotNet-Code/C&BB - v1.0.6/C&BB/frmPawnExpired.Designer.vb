<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPawnExpired
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPawnExpired))
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dgExpiredPawningTransactions = New System.Windows.Forms.DataGridView()
        Me.dgExpiredPawningStock = New System.Windows.Forms.DataGridView()
        Me.lblExpiredPawningStock = New System.Windows.Forms.Label()
        Me.txtPnCID = New System.Windows.Forms.TextBox()
        Me.cmdReportAndMove = New System.Windows.Forms.Button()
        Me.cmdExit2 = New System.Windows.Forms.Button()
        Me.cmdUpdate = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.SS1 = New System.Windows.Forms.ToolStripStatusLabel()
        CType(Me.dgExpiredPawningTransactions, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgExpiredPawningStock, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdExit
        '
        Me.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.Location = New System.Drawing.Point(503, 564)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(86, 40)
        Me.cmdExit.TabIndex = 10
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = True
        Me.cmdExit.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(711, 16)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Click in the grid below to select each transaction.  Remember to go through each " & _
    "transaction, specifying prices for items."
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'dgExpiredPawningTransactions
        '
        Me.dgExpiredPawningTransactions.AllowUserToAddRows = False
        Me.dgExpiredPawningTransactions.AllowUserToDeleteRows = False
        Me.dgExpiredPawningTransactions.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgExpiredPawningTransactions.Location = New System.Drawing.Point(15, 47)
        Me.dgExpiredPawningTransactions.MultiSelect = False
        Me.dgExpiredPawningTransactions.Name = "dgExpiredPawningTransactions"
        Me.dgExpiredPawningTransactions.ReadOnly = True
        Me.dgExpiredPawningTransactions.RowHeadersVisible = False
        Me.dgExpiredPawningTransactions.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgExpiredPawningTransactions.Size = New System.Drawing.Size(711, 261)
        Me.dgExpiredPawningTransactions.TabIndex = 14
        '
        'dgExpiredPawningStock
        '
        Me.dgExpiredPawningStock.AllowUserToAddRows = False
        Me.dgExpiredPawningStock.AllowUserToDeleteRows = False
        Me.dgExpiredPawningStock.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgExpiredPawningStock.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dgExpiredPawningStock.Location = New System.Drawing.Point(60, 366)
        Me.dgExpiredPawningStock.MultiSelect = False
        Me.dgExpiredPawningStock.Name = "dgExpiredPawningStock"
        Me.dgExpiredPawningStock.RowHeadersVisible = False
        Me.dgExpiredPawningStock.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgExpiredPawningStock.Size = New System.Drawing.Size(631, 166)
        Me.dgExpiredPawningStock.TabIndex = 15
        '
        'lblExpiredPawningStock
        '
        Me.lblExpiredPawningStock.AutoSize = True
        Me.lblExpiredPawningStock.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExpiredPawningStock.Location = New System.Drawing.Point(205, 344)
        Me.lblExpiredPawningStock.Name = "lblExpiredPawningStock"
        Me.lblExpiredPawningStock.Size = New System.Drawing.Size(266, 16)
        Me.lblExpiredPawningStock.TabIndex = 16
        Me.lblExpiredPawningStock.Text = "Specify a Selling Price for each of the items."
        Me.lblExpiredPawningStock.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtPnCID
        '
        Me.txtPnCID.Location = New System.Drawing.Point(532, 340)
        Me.txtPnCID.Name = "txtPnCID"
        Me.txtPnCID.Size = New System.Drawing.Size(100, 20)
        Me.txtPnCID.TabIndex = 17
        Me.txtPnCID.Visible = False
        '
        'cmdReportAndMove
        '
        Me.cmdReportAndMove.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReportAndMove.Location = New System.Drawing.Point(276, 564)
        Me.cmdReportAndMove.Name = "cmdReportAndMove"
        Me.cmdReportAndMove.Size = New System.Drawing.Size(131, 40)
        Me.cmdReportAndMove.TabIndex = 18
        Me.cmdReportAndMove.Text = "Generate Report and Move Stock"
        Me.cmdReportAndMove.UseVisualStyleBackColor = True
        Me.cmdReportAndMove.Visible = False
        '
        'cmdExit2
        '
        Me.cmdExit2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit2.Location = New System.Drawing.Point(321, 564)
        Me.cmdExit2.Name = "cmdExit2"
        Me.cmdExit2.Size = New System.Drawing.Size(86, 40)
        Me.cmdExit2.TabIndex = 19
        Me.cmdExit2.Text = "E&xit"
        Me.cmdExit2.UseVisualStyleBackColor = True
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdate.Location = New System.Drawing.Point(111, 564)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(86, 40)
        Me.cmdUpdate.TabIndex = 20
        Me.cmdUpdate.Text = "Update PawnStock Values"
        Me.cmdUpdate.UseVisualStyleBackColor = True
        Me.cmdUpdate.Visible = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SS1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 629)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(746, 22)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 21
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'SS1
        '
        Me.SS1.Name = "SS1"
        Me.SS1.Size = New System.Drawing.Size(0, 17)
        '
        'frmPawnExpired
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdExit
        Me.ClientSize = New System.Drawing.Size(746, 651)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.cmdUpdate)
        Me.Controls.Add(Me.cmdExit2)
        Me.Controls.Add(Me.cmdReportAndMove)
        Me.Controls.Add(Me.txtPnCID)
        Me.Controls.Add(Me.lblExpiredPawningStock)
        Me.Controls.Add(Me.dgExpiredPawningStock)
        Me.Controls.Add(Me.dgExpiredPawningTransactions)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdExit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPawnExpired"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Expired Pawning Transactions"
        CType(Me.dgExpiredPawningTransactions, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgExpiredPawningStock, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgExpiredPawningTransactions As System.Windows.Forms.DataGridView
    Friend WithEvents dgExpiredPawningStock As System.Windows.Forms.DataGridView
    Friend WithEvents lblExpiredPawningStock As System.Windows.Forms.Label
    Friend WithEvents txtPnCID As System.Windows.Forms.TextBox
    Friend WithEvents cmdReportAndMove As System.Windows.Forms.Button
    Friend WithEvents cmdExit2 As System.Windows.Forms.Button
    Friend WithEvents cmdUpdate As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents SS1 As System.Windows.Forms.ToolStripStatusLabel
End Class
