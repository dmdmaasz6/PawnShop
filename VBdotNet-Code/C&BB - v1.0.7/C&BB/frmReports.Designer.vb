<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReports
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmReports))
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdDailyRpt = New System.Windows.Forms.Button()
        Me.CBB_DataSet = New CashBuyBack.CBB_DataSet()
        Me.CustomersBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.CustomersTableAdapter = New CashBuyBack.CBB_DataSetTableAdapters.CustomersTableAdapter()
        Me.TableAdapterManager = New CashBuyBack.CBB_DataSetTableAdapters.TableAdapterManager()
        Me.cmdCustomRpt = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdPawnHistRpt = New System.Windows.Forms.Button()
        Me.cmdStdStockList = New System.Windows.Forms.Button()
        Me.cmdTangoStockLst = New System.Windows.Forms.Button()
        Me.cmdStockReport = New System.Windows.Forms.Button()
        Me.cmdPawnReport = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdRealIncome = New System.Windows.Forms.Button()
        Me.cmdStockSale = New System.Windows.Forms.Button()
        Me.cmdMoveExpiredStock = New System.Windows.Forms.Button()
        CType(Me.CBB_DataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CustomersBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdExit
        '
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.Location = New System.Drawing.Point(65, 331)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(86, 40)
        Me.cmdExit.TabIndex = 2
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'cmdDailyRpt
        '
        Me.cmdDailyRpt.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDailyRpt.Location = New System.Drawing.Point(22, 78)
        Me.cmdDailyRpt.Name = "cmdDailyRpt"
        Me.cmdDailyRpt.Size = New System.Drawing.Size(86, 40)
        Me.cmdDailyRpt.TabIndex = 3
        Me.cmdDailyRpt.Text = "Daily Rpt"
        Me.cmdDailyRpt.UseVisualStyleBackColor = True
        '
        'CBB_DataSet
        '
        Me.CBB_DataSet.DataSetName = "CBB_DataSet"
        Me.CBB_DataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'CustomersBindingSource
        '
        Me.CustomersBindingSource.DataMember = "Customers"
        Me.CustomersBindingSource.DataSource = Me.CBB_DataSet
        '
        'CustomersTableAdapter
        '
        Me.CustomersTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager
        '
        Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager.CashTransactionsTableAdapter = Nothing
        Me.TableAdapterManager.CustomersTableAdapter = Me.CustomersTableAdapter
        Me.TableAdapterManager.LayBuyHistoryTableAdapter = Nothing
        Me.TableAdapterManager.LayBuyTableAdapter = Nothing
        Me.TableAdapterManager.NormalStockTableAdapter = Nothing
        Me.TableAdapterManager.PawningTransactionsTableAdapter = Nothing
        Me.TableAdapterManager.PawnStockTableAdapter = Nothing
        Me.TableAdapterManager.StockSaleTableAdapter = Nothing
        Me.TableAdapterManager.UpdateOrder = CashBuyBack.CBB_DataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'cmdCustomRpt
        '
        Me.cmdCustomRpt.Location = New System.Drawing.Point(22, 124)
        Me.cmdCustomRpt.Name = "cmdCustomRpt"
        Me.cmdCustomRpt.Size = New System.Drawing.Size(86, 40)
        Me.cmdCustomRpt.TabIndex = 5
        Me.cmdCustomRpt.Text = "Custom Rpt"
        Me.cmdCustomRpt.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(77, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 20)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Label1"
        '
        'cmdPawnHistRpt
        '
        Me.cmdPawnHistRpt.Location = New System.Drawing.Point(22, 170)
        Me.cmdPawnHistRpt.Name = "cmdPawnHistRpt"
        Me.cmdPawnHistRpt.Size = New System.Drawing.Size(86, 40)
        Me.cmdPawnHistRpt.TabIndex = 7
        Me.cmdPawnHistRpt.Text = "PawnHist Rpt"
        Me.cmdPawnHistRpt.UseVisualStyleBackColor = True
        '
        'cmdStdStockList
        '
        Me.cmdStdStockList.Location = New System.Drawing.Point(114, 78)
        Me.cmdStdStockList.Name = "cmdStdStockList"
        Me.cmdStdStockList.Size = New System.Drawing.Size(86, 40)
        Me.cmdStdStockList.TabIndex = 8
        Me.cmdStdStockList.Text = "StdStock"
        Me.cmdStdStockList.UseVisualStyleBackColor = True
        '
        'cmdTangoStockLst
        '
        Me.cmdTangoStockLst.Location = New System.Drawing.Point(114, 124)
        Me.cmdTangoStockLst.Name = "cmdTangoStockLst"
        Me.cmdTangoStockLst.Size = New System.Drawing.Size(86, 40)
        Me.cmdTangoStockLst.TabIndex = 9
        Me.cmdTangoStockLst.Text = "TangoStock"
        Me.cmdTangoStockLst.UseVisualStyleBackColor = True
        '
        'cmdStockReport
        '
        Me.cmdStockReport.Location = New System.Drawing.Point(114, 170)
        Me.cmdStockReport.Name = "cmdStockReport"
        Me.cmdStockReport.Size = New System.Drawing.Size(86, 40)
        Me.cmdStockReport.TabIndex = 10
        Me.cmdStockReport.Text = "Stock Report"
        Me.cmdStockReport.UseVisualStyleBackColor = True
        '
        'cmdPawnReport
        '
        Me.cmdPawnReport.Location = New System.Drawing.Point(22, 216)
        Me.cmdPawnReport.Name = "cmdPawnReport"
        Me.cmdPawnReport.Size = New System.Drawing.Size(86, 40)
        Me.cmdPawnReport.TabIndex = 11
        Me.cmdPawnReport.Text = "Pawn Report"
        Me.cmdPawnReport.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(22, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(49, 42)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 12
        Me.PictureBox1.TabStop = False
        '
        'cmdRealIncome
        '
        Me.cmdRealIncome.Location = New System.Drawing.Point(114, 216)
        Me.cmdRealIncome.Name = "cmdRealIncome"
        Me.cmdRealIncome.Size = New System.Drawing.Size(86, 40)
        Me.cmdRealIncome.TabIndex = 13
        Me.cmdRealIncome.Text = "Real. Income"
        Me.cmdRealIncome.UseVisualStyleBackColor = True
        '
        'cmdStockSale
        '
        Me.cmdStockSale.Location = New System.Drawing.Point(22, 262)
        Me.cmdStockSale.Name = "cmdStockSale"
        Me.cmdStockSale.Size = New System.Drawing.Size(86, 40)
        Me.cmdStockSale.TabIndex = 14
        Me.cmdStockSale.Text = "Stock Sale"
        Me.cmdStockSale.UseVisualStyleBackColor = True
        '
        'cmdMoveExpiredStock
        '
        Me.cmdMoveExpiredStock.Location = New System.Drawing.Point(114, 262)
        Me.cmdMoveExpiredStock.Name = "cmdMoveExpiredStock"
        Me.cmdMoveExpiredStock.Size = New System.Drawing.Size(86, 40)
        Me.cmdMoveExpiredStock.TabIndex = 15
        Me.cmdMoveExpiredStock.Text = "Move Expired Stock"
        Me.cmdMoveExpiredStock.UseVisualStyleBackColor = True
        '
        'frmReports
        '
        Me.AcceptButton = Me.cmdExit
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(220, 411)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdMoveExpiredStock)
        Me.Controls.Add(Me.cmdStockSale)
        Me.Controls.Add(Me.cmdRealIncome)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.cmdPawnReport)
        Me.Controls.Add(Me.cmdStockReport)
        Me.Controls.Add(Me.cmdTangoStockLst)
        Me.Controls.Add(Me.cmdStdStockList)
        Me.Controls.Add(Me.cmdPawnHistRpt)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdCustomRpt)
        Me.Controls.Add(Me.cmdDailyRpt)
        Me.Controls.Add(Me.cmdExit)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmReports"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        CType(Me.CBB_DataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CustomersBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents cmdDailyRpt As System.Windows.Forms.Button
    Friend WithEvents CBB_DataSet As CashBuyBack.CBB_DataSet
    Friend WithEvents CustomersBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents CustomersTableAdapter As CashBuyBack.CBB_DataSetTableAdapters.CustomersTableAdapter
    Friend WithEvents TableAdapterManager As CashBuyBack.CBB_DataSetTableAdapters.TableAdapterManager
    Friend WithEvents cmdCustomRpt As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdPawnHistRpt As System.Windows.Forms.Button
    Friend WithEvents cmdStdStockList As System.Windows.Forms.Button
    Friend WithEvents cmdTangoStockLst As System.Windows.Forms.Button
    Friend WithEvents cmdStockReport As System.Windows.Forms.Button
    Friend WithEvents cmdPawnReport As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdRealIncome As System.Windows.Forms.Button
    Friend WithEvents cmdStockSale As System.Windows.Forms.Button
    Friend WithEvents cmdMoveExpiredStock As System.Windows.Forms.Button
End Class
