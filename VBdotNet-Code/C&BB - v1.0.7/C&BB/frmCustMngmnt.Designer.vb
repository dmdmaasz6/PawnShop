<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCustMngmnt
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCustMngmnt))
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.grpDetails = New System.Windows.Forms.GroupBox()
        Me.txtAddress = New System.Windows.Forms.RichTextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtTelW = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtTelH = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtIDnr = New System.Windows.Forms.TextBox()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.txtCustID = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdSearchCustomer = New System.Windows.Forms.Button()
        Me.cmdUpdateCustomerInformation = New System.Windows.Forms.Button()
        Me.cmdAddCustomer = New System.Windows.Forms.Button()
        Me.cmdClearFields = New System.Windows.Forms.Button()
        Me.cmdShowHistory = New System.Windows.Forms.Button()
        Me.cmdUseThisCustomer = New System.Windows.Forms.Button()
        Me.dgCustomerList = New System.Windows.Forms.DataGridView()
        Me.CustIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NameDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IDNrDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TelWDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TelHDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AddressDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CustomersBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.CBB_DataSet = New CashBuyBack.CBB_DataSet()
        Me.CustomersTableAdapter = New CashBuyBack.CBB_DataSetTableAdapters.CustomersTableAdapter()
        Me.TableAdapterManager = New CashBuyBack.CBB_DataSetTableAdapters.TableAdapterManager()
        Me.grpDetails.SuspendLayout()
        CType(Me.dgCustomerList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CustomersBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CBB_DataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdExit
        '
        Me.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.Location = New System.Drawing.Point(473, 508)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(150, 40)
        Me.cmdExit.TabIndex = 1
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'grpDetails
        '
        Me.grpDetails.Controls.Add(Me.txtAddress)
        Me.grpDetails.Controls.Add(Me.Label6)
        Me.grpDetails.Controls.Add(Me.txtTelW)
        Me.grpDetails.Controls.Add(Me.Label5)
        Me.grpDetails.Controls.Add(Me.txtTelH)
        Me.grpDetails.Controls.Add(Me.Label4)
        Me.grpDetails.Controls.Add(Me.txtIDnr)
        Me.grpDetails.Controls.Add(Me.txtName)
        Me.grpDetails.Controls.Add(Me.txtCustID)
        Me.grpDetails.Controls.Add(Me.Label3)
        Me.grpDetails.Controls.Add(Me.Label2)
        Me.grpDetails.Controls.Add(Me.Label1)
        Me.grpDetails.Controls.Add(Me.cmdSearchCustomer)
        Me.grpDetails.Controls.Add(Me.cmdUpdateCustomerInformation)
        Me.grpDetails.Controls.Add(Me.cmdAddCustomer)
        Me.grpDetails.Controls.Add(Me.cmdClearFields)
        Me.grpDetails.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.grpDetails.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpDetails.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpDetails.Location = New System.Drawing.Point(12, 12)
        Me.grpDetails.Name = "grpDetails"
        Me.grpDetails.Size = New System.Drawing.Size(663, 218)
        Me.grpDetails.TabIndex = 6
        Me.grpDetails.TabStop = False
        Me.grpDetails.Text = " Details: "
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(309, 22)
        Me.txtAddress.MaxLength = 100
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(180, 178)
        Me.txtAddress.TabIndex = 6
        Me.txtAddress.Text = ""
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(213, 25)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(90, 17)
        Me.Label6.TabIndex = 24
        Me.Label6.Text = "Full Address:"
        '
        'txtTelW
        '
        Me.txtTelW.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTelW.Location = New System.Drawing.Point(76, 139)
        Me.txtTelW.MaxLength = 30
        Me.txtTelW.Name = "txtTelW"
        Me.txtTelW.Size = New System.Drawing.Size(196, 23)
        Me.txtTelW.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(13, 142)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(51, 17)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Tel(w):"
        '
        'txtTelH
        '
        Me.txtTelH.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTelH.Location = New System.Drawing.Point(76, 177)
        Me.txtTelH.MaxLength = 30
        Me.txtTelH.Name = "txtTelH"
        Me.txtTelH.Size = New System.Drawing.Size(196, 23)
        Me.txtTelH.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(13, 180)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(50, 17)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "Tel(h):"
        '
        'txtIDnr
        '
        Me.txtIDnr.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIDnr.Location = New System.Drawing.Point(76, 97)
        Me.txtIDnr.MaxLength = 30
        Me.txtIDnr.Name = "txtIDnr"
        Me.txtIDnr.Size = New System.Drawing.Size(196, 23)
        Me.txtIDnr.TabIndex = 3
        '
        'txtName
        '
        Me.txtName.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.Location = New System.Drawing.Point(76, 60)
        Me.txtName.MaxLength = 30
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(196, 23)
        Me.txtName.TabIndex = 2
        '
        'txtCustID
        '
        Me.txtCustID.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustID.Location = New System.Drawing.Point(76, 22)
        Me.txtCustID.Name = "txtCustID"
        Me.txtCustID.Size = New System.Drawing.Size(111, 23)
        Me.txtCustID.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(13, 100)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(25, 17)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "ID:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(13, 63)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 17)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Name:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(13, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 17)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Cust ID:"
        '
        'cmdSearchCustomer
        '
        Me.cmdSearchCustomer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSearchCustomer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSearchCustomer.Location = New System.Drawing.Point(507, 160)
        Me.cmdSearchCustomer.Name = "cmdSearchCustomer"
        Me.cmdSearchCustomer.Size = New System.Drawing.Size(150, 40)
        Me.cmdSearchCustomer.TabIndex = 10
        Me.cmdSearchCustomer.Text = "Search"
        Me.cmdSearchCustomer.UseVisualStyleBackColor = True
        '
        'cmdUpdateCustomerInformation
        '
        Me.cmdUpdateCustomerInformation.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdateCustomerInformation.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUpdateCustomerInformation.Location = New System.Drawing.Point(507, 114)
        Me.cmdUpdateCustomerInformation.Name = "cmdUpdateCustomerInformation"
        Me.cmdUpdateCustomerInformation.Size = New System.Drawing.Size(150, 40)
        Me.cmdUpdateCustomerInformation.TabIndex = 9
        Me.cmdUpdateCustomerInformation.Text = "Update Information"
        Me.cmdUpdateCustomerInformation.UseVisualStyleBackColor = True
        '
        'cmdAddCustomer
        '
        Me.cmdAddCustomer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddCustomer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAddCustomer.Location = New System.Drawing.Point(507, 68)
        Me.cmdAddCustomer.Name = "cmdAddCustomer"
        Me.cmdAddCustomer.Size = New System.Drawing.Size(150, 40)
        Me.cmdAddCustomer.TabIndex = 8
        Me.cmdAddCustomer.Text = "Add New"
        Me.cmdAddCustomer.UseVisualStyleBackColor = True
        '
        'cmdClearFields
        '
        Me.cmdClearFields.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClearFields.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClearFields.Location = New System.Drawing.Point(507, 22)
        Me.cmdClearFields.Name = "cmdClearFields"
        Me.cmdClearFields.Size = New System.Drawing.Size(150, 40)
        Me.cmdClearFields.TabIndex = 7
        Me.cmdClearFields.Text = "Clear Fields"
        Me.cmdClearFields.UseVisualStyleBackColor = True
        '
        'cmdShowHistory
        '
        Me.cmdShowHistory.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShowHistory.Location = New System.Drawing.Point(49, 508)
        Me.cmdShowHistory.Name = "cmdShowHistory"
        Me.cmdShowHistory.Size = New System.Drawing.Size(150, 40)
        Me.cmdShowHistory.TabIndex = 11
        Me.cmdShowHistory.Text = "Show History ..."
        Me.cmdShowHistory.UseVisualStyleBackColor = True
        Me.cmdShowHistory.Visible = False
        '
        'cmdUseThisCustomer
        '
        Me.cmdUseThisCustomer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUseThisCustomer.Location = New System.Drawing.Point(205, 508)
        Me.cmdUseThisCustomer.Name = "cmdUseThisCustomer"
        Me.cmdUseThisCustomer.Size = New System.Drawing.Size(150, 40)
        Me.cmdUseThisCustomer.TabIndex = 12
        Me.cmdUseThisCustomer.Text = "Use this Customer"
        Me.cmdUseThisCustomer.UseVisualStyleBackColor = True
        Me.cmdUseThisCustomer.Visible = False
        '
        'dgCustomerList
        '
        Me.dgCustomerList.AllowUserToAddRows = False
        Me.dgCustomerList.AllowUserToDeleteRows = False
        Me.dgCustomerList.AutoGenerateColumns = False
        Me.dgCustomerList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgCustomerList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CustIDDataGridViewTextBoxColumn, Me.NameDataGridViewTextBoxColumn, Me.IDNrDataGridViewTextBoxColumn, Me.TelWDataGridViewTextBoxColumn, Me.TelHDataGridViewTextBoxColumn, Me.AddressDataGridViewTextBoxColumn})
        Me.dgCustomerList.DataSource = Me.CustomersBindingSource
        Me.dgCustomerList.Location = New System.Drawing.Point(32, 236)
        Me.dgCustomerList.MultiSelect = False
        Me.dgCustomerList.Name = "dgCustomerList"
        Me.dgCustomerList.ReadOnly = True
        Me.dgCustomerList.RowHeadersVisible = False
        Me.dgCustomerList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgCustomerList.Size = New System.Drawing.Size(643, 255)
        Me.dgCustomerList.TabIndex = 13
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
        'CBB_DataSet
        '
        Me.CBB_DataSet.DataSetName = "CBB_DataSet"
        Me.CBB_DataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
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
        'frmCustMngmnt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdExit
        Me.ClientSize = New System.Drawing.Size(692, 561)
        Me.Controls.Add(Me.dgCustomerList)
        Me.Controls.Add(Me.cmdUseThisCustomer)
        Me.Controls.Add(Me.cmdShowHistory)
        Me.Controls.Add(Me.grpDetails)
        Me.Controls.Add(Me.cmdExit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCustMngmnt"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Customer Management"
        Me.grpDetails.ResumeLayout(False)
        Me.grpDetails.PerformLayout()
        CType(Me.dgCustomerList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CustomersBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CBB_DataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents grpDetails As System.Windows.Forms.GroupBox
    Friend WithEvents cmdUpdateCustomerInformation As System.Windows.Forms.Button
    Friend WithEvents cmdAddCustomer As System.Windows.Forms.Button
    Friend WithEvents cmdClearFields As System.Windows.Forms.Button
    Friend WithEvents cmdSearchCustomer As System.Windows.Forms.Button
    Friend WithEvents txtAddress As System.Windows.Forms.RichTextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtTelW As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtTelH As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtIDnr As System.Windows.Forms.TextBox
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents txtCustID As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdShowHistory As System.Windows.Forms.Button
    Friend WithEvents cmdUseThisCustomer As System.Windows.Forms.Button
    Friend WithEvents dgCustomerList As System.Windows.Forms.DataGridView
    Friend WithEvents CBB_DataSet As CashBuyBack.CBB_DataSet
    Friend WithEvents CustomersBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents CustomersTableAdapter As CashBuyBack.CBB_DataSetTableAdapters.CustomersTableAdapter
    Friend WithEvents TableAdapterManager As CashBuyBack.CBB_DataSetTableAdapters.TableAdapterManager
    Friend WithEvents CustIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NameDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IDNrDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TelWDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TelHDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AddressDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
