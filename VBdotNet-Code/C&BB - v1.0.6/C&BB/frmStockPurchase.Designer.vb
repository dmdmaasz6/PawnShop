<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStockPurchase
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmStockPurchase))
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdPurchase = New System.Windows.Forms.Button()
        Me.rbNormal = New System.Windows.Forms.RadioButton()
        Me.rbStandard = New System.Windows.Forms.RadioButton()
        Me.rbTango = New System.Windows.Forms.RadioButton()
        Me.grpNormalStock = New System.Windows.Forms.GroupBox()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.txtAddress = New System.Windows.Forms.TextBox()
        Me.txtIDNr = New System.Windows.Forms.TextBox()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.grpStandardStock = New System.Windows.Forms.GroupBox()
        Me.dgStdStockManagment = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.grpTangoStock = New System.Windows.Forms.GroupBox()
        Me.dgTangoStockManagment = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.grpTransaction = New System.Windows.Forms.GroupBox()
        Me.txtNSID = New System.Windows.Forms.TextBox()
        Me.dtpDate = New System.Windows.Forms.DateTimePicker()
        Me.txtRecommendedSellingPrice = New System.Windows.Forms.TextBox()
        Me.txtPurchasePrice = New System.Windows.Forms.TextBox()
        Me.txtQuantity = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.NormalStockTableAdapter1 = New CashBuyBack.CBB_DataSetTableAdapters.NormalStockTableAdapter()
        Me.CashTransactionsTableAdapter1 = New CashBuyBack.CBB_DataSetTableAdapters.CashTransactionsTableAdapter()
        Me.grpNormalStock.SuspendLayout()
        Me.grpStandardStock.SuspendLayout()
        CType(Me.dgStdStockManagment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpTangoStock.SuspendLayout()
        CType(Me.dgTangoStockManagment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpTransaction.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdExit
        '
        Me.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdExit.Location = New System.Drawing.Point(288, 546)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(120, 40)
        Me.cmdExit.TabIndex = 13
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'cmdPurchase
        '
        Me.cmdPurchase.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPurchase.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPurchase.Location = New System.Drawing.Point(60, 546)
        Me.cmdPurchase.Name = "cmdPurchase"
        Me.cmdPurchase.Size = New System.Drawing.Size(120, 40)
        Me.cmdPurchase.TabIndex = 12
        Me.cmdPurchase.Text = "Purchase"
        Me.cmdPurchase.UseVisualStyleBackColor = True
        '
        'rbNormal
        '
        Me.rbNormal.AutoSize = True
        Me.rbNormal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbNormal.Location = New System.Drawing.Point(15, 26)
        Me.rbNormal.Name = "rbNormal"
        Me.rbNormal.Size = New System.Drawing.Size(107, 20)
        Me.rbNormal.TabIndex = 14
        Me.rbNormal.TabStop = True
        Me.rbNormal.Text = "Normal Stock"
        Me.rbNormal.UseVisualStyleBackColor = True
        '
        'rbStandard
        '
        Me.rbStandard.AutoSize = True
        Me.rbStandard.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbStandard.Location = New System.Drawing.Point(176, 26)
        Me.rbStandard.Name = "rbStandard"
        Me.rbStandard.Size = New System.Drawing.Size(114, 20)
        Me.rbStandard.TabIndex = 15
        Me.rbStandard.TabStop = True
        Me.rbStandard.Text = "Standad Stock"
        Me.rbStandard.UseVisualStyleBackColor = True
        '
        'rbTango
        '
        Me.rbTango.AutoSize = True
        Me.rbTango.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbTango.Location = New System.Drawing.Point(355, 26)
        Me.rbTango.Name = "rbTango"
        Me.rbTango.Size = New System.Drawing.Size(103, 20)
        Me.rbTango.TabIndex = 16
        Me.rbTango.TabStop = True
        Me.rbTango.Text = "Tango Stock"
        Me.rbTango.UseVisualStyleBackColor = True
        '
        'grpNormalStock
        '
        Me.grpNormalStock.Controls.Add(Me.txtDescription)
        Me.grpNormalStock.Controls.Add(Me.txtAddress)
        Me.grpNormalStock.Controls.Add(Me.txtIDNr)
        Me.grpNormalStock.Controls.Add(Me.txtName)
        Me.grpNormalStock.Controls.Add(Me.Label4)
        Me.grpNormalStock.Controls.Add(Me.Label3)
        Me.grpNormalStock.Controls.Add(Me.Label2)
        Me.grpNormalStock.Controls.Add(Me.Label1)
        Me.grpNormalStock.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpNormalStock.ForeColor = System.Drawing.Color.Green
        Me.grpNormalStock.Location = New System.Drawing.Point(15, 52)
        Me.grpNormalStock.Name = "grpNormalStock"
        Me.grpNormalStock.Size = New System.Drawing.Size(443, 269)
        Me.grpNormalStock.TabIndex = 17
        Me.grpNormalStock.TabStop = False
        Me.grpNormalStock.Text = " Normal Stock "
        '
        'txtDescription
        '
        Me.txtDescription.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDescription.Location = New System.Drawing.Point(172, 160)
        Me.txtDescription.MaxLength = 50
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(250, 103)
        Me.txtDescription.TabIndex = 7
        '
        'txtAddress
        '
        Me.txtAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAddress.Location = New System.Drawing.Point(172, 84)
        Me.txtAddress.MaxLength = 100
        Me.txtAddress.Multiline = True
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(250, 68)
        Me.txtAddress.TabIndex = 6
        '
        'txtIDNr
        '
        Me.txtIDNr.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIDNr.Location = New System.Drawing.Point(172, 53)
        Me.txtIDNr.MaxLength = 30
        Me.txtIDNr.Name = "txtIDNr"
        Me.txtIDNr.Size = New System.Drawing.Size(250, 22)
        Me.txtIDNr.TabIndex = 5
        '
        'txtName
        '
        Me.txtName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.Location = New System.Drawing.Point(172, 23)
        Me.txtName.MaxLength = 30
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(250, 22)
        Me.txtName.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(22, 163)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(113, 16)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Stock Description"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(22, 87)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Supplier Address"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(22, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Supplier ID:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(22, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Supplier:"
        '
        'grpStandardStock
        '
        Me.grpStandardStock.Controls.Add(Me.dgStdStockManagment)
        Me.grpStandardStock.Controls.Add(Me.Label9)
        Me.grpStandardStock.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpStandardStock.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.grpStandardStock.Location = New System.Drawing.Point(482, 52)
        Me.grpStandardStock.Name = "grpStandardStock"
        Me.grpStandardStock.Size = New System.Drawing.Size(443, 269)
        Me.grpStandardStock.TabIndex = 18
        Me.grpStandardStock.TabStop = False
        Me.grpStandardStock.Text = " Standard Stock "
        '
        'dgStdStockManagment
        '
        Me.dgStdStockManagment.AllowUserToAddRows = False
        Me.dgStdStockManagment.AllowUserToDeleteRows = False
        Me.dgStdStockManagment.AllowUserToResizeRows = False
        Me.dgStdStockManagment.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgStdStockManagment.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column3})
        Me.dgStdStockManagment.Location = New System.Drawing.Point(6, 33)
        Me.dgStdStockManagment.MultiSelect = False
        Me.dgStdStockManagment.Name = "dgStdStockManagment"
        Me.dgStdStockManagment.RowHeadersVisible = False
        Me.dgStdStockManagment.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgStdStockManagment.Size = New System.Drawing.Size(431, 230)
        Me.dgStdStockManagment.TabIndex = 9
        '
        'Column1
        '
        Me.Column1.HeaderText = "NSID"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Column1.Width = 75
        '
        'Column2
        '
        Me.Column2.HeaderText = "Description"
        Me.Column2.Name = "Column2"
        Me.Column2.Width = 200
        '
        'Column3
        '
        Me.Column3.HeaderText = "Selling Price"
        Me.Column3.Name = "Column3"
        Me.Column3.Width = 125
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(84, 18)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(179, 13)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "Click in row to select purchased item"
        '
        'grpTangoStock
        '
        Me.grpTangoStock.Controls.Add(Me.dgTangoStockManagment)
        Me.grpTangoStock.Controls.Add(Me.Label10)
        Me.grpTangoStock.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpTangoStock.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.grpTangoStock.Location = New System.Drawing.Point(945, 52)
        Me.grpTangoStock.Name = "grpTangoStock"
        Me.grpTangoStock.Size = New System.Drawing.Size(443, 269)
        Me.grpTangoStock.TabIndex = 19
        Me.grpTangoStock.TabStop = False
        Me.grpTangoStock.Text = " Tango Stock "
        '
        'dgTangoStockManagment
        '
        Me.dgTangoStockManagment.AllowUserToAddRows = False
        Me.dgTangoStockManagment.AllowUserToDeleteRows = False
        Me.dgTangoStockManagment.AllowUserToResizeRows = False
        Me.dgTangoStockManagment.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgTangoStockManagment.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3})
        Me.dgTangoStockManagment.Location = New System.Drawing.Point(6, 33)
        Me.dgTangoStockManagment.MultiSelect = False
        Me.dgTangoStockManagment.Name = "dgTangoStockManagment"
        Me.dgTangoStockManagment.RowHeadersVisible = False
        Me.dgTangoStockManagment.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgTangoStockManagment.Size = New System.Drawing.Size(431, 230)
        Me.dgTangoStockManagment.TabIndex = 4
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "NSID"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        Me.DataGridViewTextBoxColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridViewTextBoxColumn1.Width = 75
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "Description"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.Width = 200
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "Selling Price"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Width = 125
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(84, 18)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(179, 13)
        Me.Label10.TabIndex = 1
        Me.Label10.Text = "Click in row to select purchased item"
        '
        'grpTransaction
        '
        Me.grpTransaction.Controls.Add(Me.txtNSID)
        Me.grpTransaction.Controls.Add(Me.dtpDate)
        Me.grpTransaction.Controls.Add(Me.txtRecommendedSellingPrice)
        Me.grpTransaction.Controls.Add(Me.txtPurchasePrice)
        Me.grpTransaction.Controls.Add(Me.txtQuantity)
        Me.grpTransaction.Controls.Add(Me.Label8)
        Me.grpTransaction.Controls.Add(Me.Label7)
        Me.grpTransaction.Controls.Add(Me.Label6)
        Me.grpTransaction.Controls.Add(Me.Label5)
        Me.grpTransaction.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpTransaction.Location = New System.Drawing.Point(51, 327)
        Me.grpTransaction.Name = "grpTransaction"
        Me.grpTransaction.Size = New System.Drawing.Size(376, 182)
        Me.grpTransaction.TabIndex = 20
        Me.grpTransaction.TabStop = False
        Me.grpTransaction.Text = " Transaction "
        '
        'txtNSID
        '
        Me.txtNSID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNSID.Location = New System.Drawing.Point(254, 30)
        Me.txtNSID.Name = "txtNSID"
        Me.txtNSID.Size = New System.Drawing.Size(73, 22)
        Me.txtNSID.TabIndex = 9
        Me.txtNSID.TabStop = False
        Me.txtNSID.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtNSID.Visible = False
        '
        'dtpDate
        '
        Me.dtpDate.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpDate.Location = New System.Drawing.Point(172, 63)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(187, 22)
        Me.dtpDate.TabIndex = 8
        '
        'txtRecommendedSellingPrice
        '
        Me.txtRecommendedSellingPrice.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRecommendedSellingPrice.Location = New System.Drawing.Point(172, 134)
        Me.txtRecommendedSellingPrice.Name = "txtRecommendedSellingPrice"
        Me.txtRecommendedSellingPrice.Size = New System.Drawing.Size(125, 22)
        Me.txtRecommendedSellingPrice.TabIndex = 7
        Me.txtRecommendedSellingPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtPurchasePrice
        '
        Me.txtPurchasePrice.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPurchasePrice.Location = New System.Drawing.Point(172, 98)
        Me.txtPurchasePrice.Name = "txtPurchasePrice"
        Me.txtPurchasePrice.Size = New System.Drawing.Size(125, 22)
        Me.txtPurchasePrice.TabIndex = 6
        Me.txtPurchasePrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtQuantity
        '
        Me.txtQuantity.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQuantity.Location = New System.Drawing.Point(172, 30)
        Me.txtQuantity.Name = "txtQuantity"
        Me.txtQuantity.Size = New System.Drawing.Size(44, 22)
        Me.txtQuantity.TabIndex = 5
        Me.txtQuantity.Text = "1"
        Me.txtQuantity.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(22, 137)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(86, 16)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "Selling Price:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(22, 101)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(102, 16)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "Purchase Price:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(22, 68)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 16)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "Date:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(22, 33)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(59, 16)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Quantity:"
        '
        'NormalStockTableAdapter1
        '
        Me.NormalStockTableAdapter1.ClearBeforeFill = True
        '
        'CashTransactionsTableAdapter1
        '
        Me.CashTransactionsTableAdapter1.ClearBeforeFill = True
        '
        'frmStockPurchase
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdExit
        Me.ClientSize = New System.Drawing.Size(1394, 625)
        Me.Controls.Add(Me.grpTransaction)
        Me.Controls.Add(Me.grpTangoStock)
        Me.Controls.Add(Me.grpStandardStock)
        Me.Controls.Add(Me.grpNormalStock)
        Me.Controls.Add(Me.rbTango)
        Me.Controls.Add(Me.rbStandard)
        Me.Controls.Add(Me.rbNormal)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdPurchase)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmStockPurchase"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Stock Purchase"
        Me.grpNormalStock.ResumeLayout(False)
        Me.grpNormalStock.PerformLayout()
        Me.grpStandardStock.ResumeLayout(False)
        Me.grpStandardStock.PerformLayout()
        CType(Me.dgStdStockManagment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpTangoStock.ResumeLayout(False)
        Me.grpTangoStock.PerformLayout()
        CType(Me.dgTangoStockManagment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpTransaction.ResumeLayout(False)
        Me.grpTransaction.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents cmdPurchase As System.Windows.Forms.Button
    Friend WithEvents rbNormal As System.Windows.Forms.RadioButton
    Friend WithEvents rbStandard As System.Windows.Forms.RadioButton
    Friend WithEvents rbTango As System.Windows.Forms.RadioButton
    Friend WithEvents grpNormalStock As System.Windows.Forms.GroupBox
    Friend WithEvents grpStandardStock As System.Windows.Forms.GroupBox
    Friend WithEvents grpTangoStock As System.Windows.Forms.GroupBox
    Friend WithEvents txtIDNr As System.Windows.Forms.TextBox
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents grpTransaction As System.Windows.Forms.GroupBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtRecommendedSellingPrice As System.Windows.Forms.TextBox
    Friend WithEvents txtPurchasePrice As System.Windows.Forms.TextBox
    Friend WithEvents txtQuantity As System.Windows.Forms.TextBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents NormalStockTableAdapter1 As CashBuyBack.CBB_DataSetTableAdapters.NormalStockTableAdapter
    Friend WithEvents CashTransactionsTableAdapter1 As CashBuyBack.CBB_DataSetTableAdapters.CashTransactionsTableAdapter
    Friend WithEvents dgStdStockManagment As System.Windows.Forms.DataGridView
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgTangoStockManagment As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents txtNSID As System.Windows.Forms.TextBox
End Class
