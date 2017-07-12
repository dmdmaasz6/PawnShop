<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPawnIn
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPawnIn))
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.dtpBuyBackDate = New System.Windows.Forms.DateTimePicker()
        Me.txtPNCID = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdDetail = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblCustomerInformation = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lblPawningTransaction = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lblGoods = New System.Windows.Forms.Label()
        Me.txtActualBuyBackAmount = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmdPawnIn = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.dgPawnInList = New System.Windows.Forms.DataGridView()
        Me.CBB_DataSet = New CashBuyBack.CBB_DataSet()
        Me.txtCustID = New System.Windows.Forms.TextBox()
        Me.dgCustomerList = New System.Windows.Forms.DataGridView()
        Me.CustIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NameDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IDNrDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TelWDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TelHDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AddressDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CustomersBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.cmdSearchForPNC = New System.Windows.Forms.Button()
        Me.txtIDnr = New System.Windows.Forms.TextBox()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CustomersTableAdapter = New CashBuyBack.CBB_DataSetTableAdapters.CustomersTableAdapter()
        Me.lblNoTransactions = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.dgPawnInList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CBB_DataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgCustomerList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CustomersBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdExit
        '
        Me.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.Location = New System.Drawing.Point(639, 641)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(86, 40)
        Me.cmdExit.TabIndex = 10
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(13, 28)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(46, 17)
        Me.Label8.TabIndex = 46
        Me.Label8.Text = "Date: "
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'dtpBuyBackDate
        '
        Me.dtpBuyBackDate.Location = New System.Drawing.Point(77, 28)
        Me.dtpBuyBackDate.Name = "dtpBuyBackDate"
        Me.dtpBuyBackDate.Size = New System.Drawing.Size(160, 20)
        Me.dtpBuyBackDate.TabIndex = 45
        '
        'txtPNCID
        '
        Me.txtPNCID.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPNCID.Location = New System.Drawing.Point(77, 53)
        Me.txtPNCID.Name = "txtPNCID"
        Me.txtPNCID.Size = New System.Drawing.Size(111, 23)
        Me.txtPNCID.TabIndex = 48
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(13, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 17)
        Me.Label1.TabIndex = 49
        Me.Label1.Text = "CBB ID:"
        '
        'cmdDetail
        '
        Me.cmdDetail.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDetail.Location = New System.Drawing.Point(253, 36)
        Me.cmdDetail.Name = "cmdDetail"
        Me.cmdDetail.Size = New System.Drawing.Size(86, 40)
        Me.cmdDetail.TabIndex = 47
        Me.cmdDetail.Text = "Get Detail"
        Me.cmdDetail.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblCustomerInformation)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Blue
        Me.GroupBox1.Location = New System.Drawing.Point(12, 82)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(327, 238)
        Me.GroupBox1.TabIndex = 50
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = " Customer Information "
        '
        'lblCustomerInformation
        '
        Me.lblCustomerInformation.AutoSize = True
        Me.lblCustomerInformation.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCustomerInformation.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCustomerInformation.Location = New System.Drawing.Point(20, 27)
        Me.lblCustomerInformation.Name = "lblCustomerInformation"
        Me.lblCustomerInformation.Size = New System.Drawing.Size(46, 16)
        Me.lblCustomerInformation.TabIndex = 0
        Me.lblCustomerInformation.Text = "lblcust"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lblPawningTransaction)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.Blue
        Me.GroupBox2.Location = New System.Drawing.Point(16, 326)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(327, 124)
        Me.GroupBox2.TabIndex = 51
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = " Pawning Transaction "
        '
        'lblPawningTransaction
        '
        Me.lblPawningTransaction.AutoSize = True
        Me.lblPawningTransaction.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPawningTransaction.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPawningTransaction.Location = New System.Drawing.Point(20, 27)
        Me.lblPawningTransaction.Name = "lblPawningTransaction"
        Me.lblPawningTransaction.Size = New System.Drawing.Size(55, 16)
        Me.lblPawningTransaction.TabIndex = 0
        Me.lblPawningTransaction.Text = "lblPawn"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.lblGoods)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.Blue
        Me.GroupBox3.Location = New System.Drawing.Point(16, 456)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(327, 143)
        Me.GroupBox3.TabIndex = 52
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = " Goods "
        '
        'lblGoods
        '
        Me.lblGoods.AutoSize = True
        Me.lblGoods.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGoods.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblGoods.Location = New System.Drawing.Point(20, 27)
        Me.lblGoods.Name = "lblGoods"
        Me.lblGoods.Size = New System.Drawing.Size(63, 16)
        Me.lblGoods.TabIndex = 0
        Me.lblGoods.Text = "lblGoods"
        '
        'txtActualBuyBackAmount
        '
        Me.txtActualBuyBackAmount.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtActualBuyBackAmount.Location = New System.Drawing.Point(338, 650)
        Me.txtActualBuyBackAmount.Name = "txtActualBuyBackAmount"
        Me.txtActualBuyBackAmount.Size = New System.Drawing.Size(111, 23)
        Me.txtActualBuyBackAmount.TabIndex = 53
        Me.txtActualBuyBackAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label2.Location = New System.Drawing.Point(154, 653)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(178, 16)
        Me.Label2.TabIndex = 54
        Me.Label2.Text = "Actual Buyback Amount: "
        '
        'cmdPawnIn
        '
        Me.cmdPawnIn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPawnIn.Location = New System.Drawing.Point(507, 641)
        Me.cmdPawnIn.Name = "cmdPawnIn"
        Me.cmdPawnIn.Size = New System.Drawing.Size(86, 40)
        Me.cmdPawnIn.TabIndex = 55
        Me.cmdPawnIn.Text = "Pawn In"
        Me.cmdPawnIn.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.lblNoTransactions)
        Me.GroupBox4.Controls.Add(Me.dgPawnInList)
        Me.GroupBox4.Controls.Add(Me.txtCustID)
        Me.GroupBox4.Controls.Add(Me.dgCustomerList)
        Me.GroupBox4.Controls.Add(Me.GroupBox5)
        Me.GroupBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.ForeColor = System.Drawing.Color.Blue
        Me.GroupBox4.Location = New System.Drawing.Point(349, 12)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(660, 587)
        Me.GroupBox4.TabIndex = 56
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = " Search for CBB ID: "
        '
        'dgPawnInList
        '
        Me.dgPawnInList.AllowUserToAddRows = False
        Me.dgPawnInList.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgPawnInList.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgPawnInList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgPawnInList.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgPawnInList.Location = New System.Drawing.Point(6, 325)
        Me.dgPawnInList.MultiSelect = False
        Me.dgPawnInList.Name = "dgPawnInList"
        Me.dgPawnInList.ReadOnly = True
        Me.dgPawnInList.RowHeadersVisible = False
        Me.dgPawnInList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgPawnInList.Size = New System.Drawing.Size(643, 189)
        Me.dgPawnInList.TabIndex = 19
        Me.dgPawnInList.Visible = False
        '
        'CBB_DataSet
        '
        Me.CBB_DataSet.DataSetName = "CBB_DataSet"
        Me.CBB_DataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'txtCustID
        '
        Me.txtCustID.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustID.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtCustID.Location = New System.Drawing.Point(510, 41)
        Me.txtCustID.Name = "txtCustID"
        Me.txtCustID.Size = New System.Drawing.Size(90, 23)
        Me.txtCustID.TabIndex = 18
        Me.txtCustID.Visible = False
        '
        'dgCustomerList
        '
        Me.dgCustomerList.AllowUserToAddRows = False
        Me.dgCustomerList.AllowUserToDeleteRows = False
        Me.dgCustomerList.AutoGenerateColumns = False
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgCustomerList.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgCustomerList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgCustomerList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CustIDDataGridViewTextBoxColumn, Me.NameDataGridViewTextBoxColumn, Me.IDNrDataGridViewTextBoxColumn, Me.TelWDataGridViewTextBoxColumn, Me.TelHDataGridViewTextBoxColumn, Me.AddressDataGridViewTextBoxColumn})
        Me.dgCustomerList.DataSource = Me.CustomersBindingSource
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgCustomerList.DefaultCellStyle = DataGridViewCellStyle4
        Me.dgCustomerList.Location = New System.Drawing.Point(6, 119)
        Me.dgCustomerList.MultiSelect = False
        Me.dgCustomerList.Name = "dgCustomerList"
        Me.dgCustomerList.ReadOnly = True
        Me.dgCustomerList.RowHeadersVisible = False
        Me.dgCustomerList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgCustomerList.Size = New System.Drawing.Size(643, 189)
        Me.dgCustomerList.TabIndex = 14
        '
        'CustIDDataGridViewTextBoxColumn
        '
        Me.CustIDDataGridViewTextBoxColumn.DataPropertyName = "CustID"
        Me.CustIDDataGridViewTextBoxColumn.HeaderText = "CustID"
        Me.CustIDDataGridViewTextBoxColumn.Name = "CustIDDataGridViewTextBoxColumn"
        Me.CustIDDataGridViewTextBoxColumn.ReadOnly = True
        '
        'NameDataGridViewTextBoxColumn
        '
        Me.NameDataGridViewTextBoxColumn.DataPropertyName = "Name"
        Me.NameDataGridViewTextBoxColumn.HeaderText = "Name"
        Me.NameDataGridViewTextBoxColumn.Name = "NameDataGridViewTextBoxColumn"
        Me.NameDataGridViewTextBoxColumn.ReadOnly = True
        '
        'IDNrDataGridViewTextBoxColumn
        '
        Me.IDNrDataGridViewTextBoxColumn.DataPropertyName = "IDNr"
        Me.IDNrDataGridViewTextBoxColumn.HeaderText = "IDNr"
        Me.IDNrDataGridViewTextBoxColumn.Name = "IDNrDataGridViewTextBoxColumn"
        Me.IDNrDataGridViewTextBoxColumn.ReadOnly = True
        '
        'TelWDataGridViewTextBoxColumn
        '
        Me.TelWDataGridViewTextBoxColumn.DataPropertyName = "TelW"
        Me.TelWDataGridViewTextBoxColumn.HeaderText = "TelW"
        Me.TelWDataGridViewTextBoxColumn.Name = "TelWDataGridViewTextBoxColumn"
        Me.TelWDataGridViewTextBoxColumn.ReadOnly = True
        '
        'TelHDataGridViewTextBoxColumn
        '
        Me.TelHDataGridViewTextBoxColumn.DataPropertyName = "TelH"
        Me.TelHDataGridViewTextBoxColumn.HeaderText = "TelH"
        Me.TelHDataGridViewTextBoxColumn.Name = "TelHDataGridViewTextBoxColumn"
        Me.TelHDataGridViewTextBoxColumn.ReadOnly = True
        '
        'AddressDataGridViewTextBoxColumn
        '
        Me.AddressDataGridViewTextBoxColumn.DataPropertyName = "Address"
        Me.AddressDataGridViewTextBoxColumn.HeaderText = "Address"
        Me.AddressDataGridViewTextBoxColumn.Name = "AddressDataGridViewTextBoxColumn"
        Me.AddressDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CustomersBindingSource
        '
        Me.CustomersBindingSource.DataMember = "Customers"
        Me.CustomersBindingSource.DataSource = Me.CBB_DataSet
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.cmdSearchForPNC)
        Me.GroupBox5.Controls.Add(Me.txtIDnr)
        Me.GroupBox5.Controls.Add(Me.txtName)
        Me.GroupBox5.Controls.Add(Me.Label3)
        Me.GroupBox5.Controls.Add(Me.Label4)
        Me.GroupBox5.Location = New System.Drawing.Point(6, 16)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(487, 97)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = " Customer Information "
        '
        'cmdSearchForPNC
        '
        Me.cmdSearchForPNC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSearchForPNC.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSearchForPNC.Location = New System.Drawing.Point(329, 31)
        Me.cmdSearchForPNC.Name = "cmdSearchForPNC"
        Me.cmdSearchForPNC.Size = New System.Drawing.Size(86, 40)
        Me.cmdSearchForPNC.TabIndex = 48
        Me.cmdSearchForPNC.Text = "Search"
        Me.cmdSearchForPNC.UseVisualStyleBackColor = True
        '
        'txtIDnr
        '
        Me.txtIDnr.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIDnr.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtIDnr.Location = New System.Drawing.Point(69, 57)
        Me.txtIDnr.Name = "txtIDnr"
        Me.txtIDnr.Size = New System.Drawing.Size(196, 23)
        Me.txtIDnr.TabIndex = 18
        '
        'txtName
        '
        Me.txtName.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtName.Location = New System.Drawing.Point(69, 28)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(196, 23)
        Me.txtName.TabIndex = 17
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(6, 60)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(25, 17)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "ID:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(6, 31)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 17)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Name:"
        '
        'CustomersTableAdapter
        '
        Me.CustomersTableAdapter.ClearBeforeFill = True
        '
        'lblNoTransactions
        '
        Me.lblNoTransactions.AutoSize = True
        Me.lblNoTransactions.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNoTransactions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNoTransactions.Location = New System.Drawing.Point(115, 325)
        Me.lblNoTransactions.Name = "lblNoTransactions"
        Me.lblNoTransactions.Size = New System.Drawing.Size(436, 24)
        Me.lblNoTransactions.TabIndex = 20
        Me.lblNoTransactions.Text = "No pawning transactions available for this customer"
        Me.lblNoTransactions.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblNoTransactions.Visible = False
        '
        'frmPawnIn
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdExit
        Me.ClientSize = New System.Drawing.Size(1020, 709)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.cmdPawnIn)
        Me.Controls.Add(Me.txtActualBuyBackAmount)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtPNCID)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdDetail)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.dtpBuyBackDate)
        Me.Controls.Add(Me.cmdExit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPawnIn"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Pawning In"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.dgPawnInList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CBB_DataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgCustomerList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CustomersBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dtpBuyBackDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtPNCID As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdDetail As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblCustomerInformation As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblPawningTransaction As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents lblGoods As System.Windows.Forms.Label
    Friend WithEvents txtActualBuyBackAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdPawnIn As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdSearchForPNC As System.Windows.Forms.Button
    Friend WithEvents txtIDnr As System.Windows.Forms.TextBox
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CBB_DataSet As CashBuyBack.CBB_DataSet
    Friend WithEvents CustomersBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents CustomersTableAdapter As CashBuyBack.CBB_DataSetTableAdapters.CustomersTableAdapter
    Friend WithEvents dgCustomerList As System.Windows.Forms.DataGridView
    Friend WithEvents CustIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NameDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IDNrDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TelWDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TelHDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AddressDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents txtCustID As System.Windows.Forms.TextBox
    Friend WithEvents dgPawnInList As System.Windows.Forms.DataGridView
    Friend WithEvents QuantityDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DescriptionDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SerialNrDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QuantitySoldDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents lblNoTransactions As System.Windows.Forms.Label
End Class
