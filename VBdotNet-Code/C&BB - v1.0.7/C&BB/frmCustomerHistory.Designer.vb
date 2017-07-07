<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCustomerHistory
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCustomerHistory))
        Me.grpCustomer = New System.Windows.Forms.GroupBox()
        Me.txtCustName = New System.Windows.Forms.TextBox()
        Me.txtCustID = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.grpHistory = New System.Windows.Forms.GroupBox()
        Me.dgCustHistory = New System.Windows.Forms.DataGridView()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.PawningTransactionsPawnStockBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.PawningTransactionsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.CBB_DataSet = New CashBuyBack.CBB_DataSet()
        Me.PawningTransactionsTableAdapter = New CashBuyBack.CBB_DataSetTableAdapters.PawningTransactionsTableAdapter()
        Me.PawnStockTableAdapter = New CashBuyBack.CBB_DataSetTableAdapters.PawnStockTableAdapter()
        Me.grpCustomer.SuspendLayout()
        Me.grpHistory.SuspendLayout()
        CType(Me.dgCustHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PawningTransactionsPawnStockBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PawningTransactionsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CBB_DataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpCustomer
        '
        Me.grpCustomer.Controls.Add(Me.txtCustName)
        Me.grpCustomer.Controls.Add(Me.txtCustID)
        Me.grpCustomer.Controls.Add(Me.Label2)
        Me.grpCustomer.Controls.Add(Me.Label1)
        Me.grpCustomer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpCustomer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpCustomer.Location = New System.Drawing.Point(12, 3)
        Me.grpCustomer.Name = "grpCustomer"
        Me.grpCustomer.Size = New System.Drawing.Size(701, 64)
        Me.grpCustomer.TabIndex = 0
        Me.grpCustomer.TabStop = False
        Me.grpCustomer.Text = " Customer "
        '
        'txtCustName
        '
        Me.txtCustName.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustName.Location = New System.Drawing.Point(434, 24)
        Me.txtCustName.Name = "txtCustName"
        Me.txtCustName.ReadOnly = True
        Me.txtCustName.Size = New System.Drawing.Size(245, 23)
        Me.txtCustName.TabIndex = 9
        Me.txtCustName.TabStop = False
        '
        'txtCustID
        '
        Me.txtCustID.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustID.Location = New System.Drawing.Point(139, 24)
        Me.txtCustID.Name = "txtCustID"
        Me.txtCustID.ReadOnly = True
        Me.txtCustID.Size = New System.Drawing.Size(95, 23)
        Me.txtCustID.TabIndex = 8
        Me.txtCustID.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(315, 27)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(113, 17)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Customer Name:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(44, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(89, 17)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Customer ID:"
        '
        'grpHistory
        '
        Me.grpHistory.Controls.Add(Me.dgCustHistory)
        Me.grpHistory.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpHistory.Location = New System.Drawing.Point(12, 73)
        Me.grpHistory.Name = "grpHistory"
        Me.grpHistory.Size = New System.Drawing.Size(701, 343)
        Me.grpHistory.TabIndex = 1
        Me.grpHistory.TabStop = False
        Me.grpHistory.Text = " History "
        '
        'dgCustHistory
        '
        Me.dgCustHistory.AllowUserToAddRows = False
        Me.dgCustHistory.AllowUserToDeleteRows = False
        Me.dgCustHistory.AllowUserToOrderColumns = True
        Me.dgCustHistory.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgCustHistory.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgCustHistory.Location = New System.Drawing.Point(3, 18)
        Me.dgCustHistory.MultiSelect = False
        Me.dgCustHistory.Name = "dgCustHistory"
        Me.dgCustHistory.ReadOnly = True
        Me.dgCustHistory.RowHeadersVisible = False
        Me.dgCustHistory.Size = New System.Drawing.Size(695, 322)
        Me.dgCustHistory.TabIndex = 0
        '
        'cmdOK
        '
        Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.Location = New System.Drawing.Point(308, 437)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(100, 40)
        Me.cmdOK.TabIndex = 2
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'PawningTransactionsPawnStockBindingSource
        '
        Me.PawningTransactionsPawnStockBindingSource.DataMember = "PawningTransactionsPawnStock"
        Me.PawningTransactionsPawnStockBindingSource.DataSource = Me.PawningTransactionsBindingSource
        '
        'PawningTransactionsBindingSource
        '
        Me.PawningTransactionsBindingSource.DataMember = "PawningTransactions"
        Me.PawningTransactionsBindingSource.DataSource = Me.CBB_DataSet
        '
        'CBB_DataSet
        '
        Me.CBB_DataSet.DataSetName = "CBB_DataSet"
        Me.CBB_DataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'PawningTransactionsTableAdapter
        '
        Me.PawningTransactionsTableAdapter.ClearBeforeFill = True
        '
        'PawnStockTableAdapter
        '
        Me.PawnStockTableAdapter.ClearBeforeFill = True
        '
        'frmCustomerHistory
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdOK
        Me.ClientSize = New System.Drawing.Size(731, 489)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.grpHistory)
        Me.Controls.Add(Me.grpCustomer)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCustomerHistory"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Customer History"
        Me.grpCustomer.ResumeLayout(False)
        Me.grpCustomer.PerformLayout()
        Me.grpHistory.ResumeLayout(False)
        CType(Me.dgCustHistory, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PawningTransactionsPawnStockBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PawningTransactionsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CBB_DataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grpCustomer As System.Windows.Forms.GroupBox
    Friend WithEvents txtCustName As System.Windows.Forms.TextBox
    Friend WithEvents txtCustID As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents grpHistory As System.Windows.Forms.GroupBox
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents CBB_DataSet As CashBuyBack.CBB_DataSet
    Friend WithEvents PawningTransactionsBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents PawningTransactionsTableAdapter As CashBuyBack.CBB_DataSetTableAdapters.PawningTransactionsTableAdapter
    Friend WithEvents dgCustHistory As System.Windows.Forms.DataGridView
    Friend WithEvents PawningTransactionsPawnStockBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents PawnStockTableAdapter As CashBuyBack.CBB_DataSetTableAdapters.PawnStockTableAdapter
End Class
