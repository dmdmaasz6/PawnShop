<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.SS1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.SS2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.SS3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Settings = New System.Windows.Forms.MenuStrip()
        Me.SettingsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DBLoactionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CashRegisterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MaintenanceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DatabaseMaintenanceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ClearHistoryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ChangeTangoEntriesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.grpCustomer = New System.Windows.Forms.GroupBox()
        Me.cmdCustUpdate = New System.Windows.Forms.Button()
        Me.cmdCustsearch = New System.Windows.Forms.Button()
        Me.cmdCustAdd = New System.Windows.Forms.Button()
        Me.grpPawning = New System.Windows.Forms.GroupBox()
        Me.cmdPawnExpired = New System.Windows.Forms.Button()
        Me.cmdPawnIn = New System.Windows.Forms.Button()
        Me.cmdPawnOut = New System.Windows.Forms.Button()
        Me.grpCashtrans = New System.Windows.Forms.GroupBox()
        Me.cmdCashOther = New System.Windows.Forms.Button()
        Me.cmdCashSale = New System.Windows.Forms.Button()
        Me.cmdCashStock = New System.Windows.Forms.Button()
        Me.grpReports = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtReportTo = New System.Windows.Forms.DateTimePicker()
        Me.dtReportFrom = New System.Windows.Forms.DateTimePicker()
        Me.cmdReportCustom = New System.Windows.Forms.Button()
        Me.cmdReportDaily = New System.Windows.Forms.Button()
        Me.grpPawnHistory = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dtPawnTo = New System.Windows.Forms.DateTimePicker()
        Me.dtPawnFrom = New System.Windows.Forms.DateTimePicker()
        Me.cmdPawnHistory = New System.Windows.Forms.Button()
        Me.grpStockMngmnt = New System.Windows.Forms.GroupBox()
        Me.cmdStockTango = New System.Windows.Forms.Button()
        Me.cmdStockStnd = New System.Windows.Forms.Button()
        Me.grpStock = New System.Windows.Forms.GroupBox()
        Me.cmdStackPawnRpt = New System.Windows.Forms.Button()
        Me.cmdStockIncome = New System.Windows.Forms.Button()
        Me.cmdStockReport = New System.Windows.Forms.Button()
        Me.cmdStockSearch = New System.Windows.Forms.Button()
        Me.grpStockCode = New System.Windows.Forms.GroupBox()
        Me.cmdStockCdTango = New System.Windows.Forms.Button()
        Me.cmdStockCdStnd = New System.Windows.Forms.Button()
        Me.StatusStrip1.SuspendLayout()
        Me.Settings.SuspendLayout()
        Me.grpCustomer.SuspendLayout()
        Me.grpPawning.SuspendLayout()
        Me.grpCashtrans.SuspendLayout()
        Me.grpReports.SuspendLayout()
        Me.grpPawnHistory.SuspendLayout()
        Me.grpStockMngmnt.SuspendLayout()
        Me.grpStock.SuspendLayout()
        Me.grpStockCode.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SS1, Me.SS2, Me.SS3})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 614)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(627, 28)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 0
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'SS1
        '
        Me.SS1.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Right
        Me.SS1.Name = "SS1"
        Me.SS1.Size = New System.Drawing.Size(287, 23)
        Me.SS1.Spring = True
        Me.SS1.Text = "Cash Register: "
        Me.SS1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'SS2
        '
        Me.SS2.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Right
        Me.SS2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.SS2.Name = "SS2"
        Me.SS2.Size = New System.Drawing.Size(287, 23)
        Me.SS2.Spring = True
        Me.SS2.Text = "Shop: "
        Me.SS2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'SS3
        '
        Me.SS3.Name = "SS3"
        Me.SS3.Size = New System.Drawing.Size(38, 23)
        Me.SS3.Text = "Time"
        '
        'cmdExit
        '
        Me.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.Location = New System.Drawing.Point(385, 528)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(150, 40)
        Me.cmdExit.TabIndex = 1
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'Settings
        '
        Me.Settings.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Settings.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SettingsToolStripMenuItem, Me.MaintenanceToolStripMenuItem, Me.AboutToolStripMenuItem})
        Me.Settings.Location = New System.Drawing.Point(0, 0)
        Me.Settings.Name = "Settings"
        Me.Settings.Size = New System.Drawing.Size(627, 24)
        Me.Settings.TabIndex = 2
        Me.Settings.Text = "Menu1"
        '
        'SettingsToolStripMenuItem
        '
        Me.SettingsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DBLoactionToolStripMenuItem, Me.CashRegisterToolStripMenuItem, Me.ToolStripSeparator1, Me.ExitToolStripMenuItem})
        Me.SettingsToolStripMenuItem.Name = "SettingsToolStripMenuItem"
        Me.SettingsToolStripMenuItem.Size = New System.Drawing.Size(68, 20)
        Me.SettingsToolStripMenuItem.Text = "Settings"
        '
        'DBLoactionToolStripMenuItem
        '
        Me.DBLoactionToolStripMenuItem.Name = "DBLoactionToolStripMenuItem"
        Me.DBLoactionToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.DBLoactionToolStripMenuItem.Text = "View"
        '
        'CashRegisterToolStripMenuItem
        '
        Me.CashRegisterToolStripMenuItem.Name = "CashRegisterToolStripMenuItem"
        Me.CashRegisterToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.CashRegisterToolStripMenuItem.Text = "Configure"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(130, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'MaintenanceToolStripMenuItem
        '
        Me.MaintenanceToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DatabaseMaintenanceToolStripMenuItem})
        Me.MaintenanceToolStripMenuItem.Name = "MaintenanceToolStripMenuItem"
        Me.MaintenanceToolStripMenuItem.Size = New System.Drawing.Size(97, 20)
        Me.MaintenanceToolStripMenuItem.Text = "Maintenance"
        '
        'DatabaseMaintenanceToolStripMenuItem
        '
        Me.DatabaseMaintenanceToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClearHistoryToolStripMenuItem, Me.ChangeTangoEntriesToolStripMenuItem})
        Me.DatabaseMaintenanceToolStripMenuItem.Name = "DatabaseMaintenanceToolStripMenuItem"
        Me.DatabaseMaintenanceToolStripMenuItem.Size = New System.Drawing.Size(216, 22)
        Me.DatabaseMaintenanceToolStripMenuItem.Text = "Database Maintenance"
        '
        'ClearHistoryToolStripMenuItem
        '
        Me.ClearHistoryToolStripMenuItem.Name = "ClearHistoryToolStripMenuItem"
        Me.ClearHistoryToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.ClearHistoryToolStripMenuItem.Text = "Clear History"
        '
        'ChangeTangoEntriesToolStripMenuItem
        '
        Me.ChangeTangoEntriesToolStripMenuItem.Name = "ChangeTangoEntriesToolStripMenuItem"
        Me.ChangeTangoEntriesToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.ChangeTangoEntriesToolStripMenuItem.Text = "Change Tango Entries"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'grpCustomer
        '
        Me.grpCustomer.Controls.Add(Me.cmdCustUpdate)
        Me.grpCustomer.Controls.Add(Me.cmdCustsearch)
        Me.grpCustomer.Controls.Add(Me.cmdCustAdd)
        Me.grpCustomer.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.grpCustomer.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpCustomer.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.grpCustomer.Location = New System.Drawing.Point(12, 27)
        Me.grpCustomer.Name = "grpCustomer"
        Me.grpCustomer.Size = New System.Drawing.Size(197, 166)
        Me.grpCustomer.TabIndex = 3
        Me.grpCustomer.TabStop = False
        Me.grpCustomer.Text = " Customer Maintenance: "
        '
        'cmdCustUpdate
        '
        Me.cmdCustUpdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCustUpdate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCustUpdate.Location = New System.Drawing.Point(25, 114)
        Me.cmdCustUpdate.Name = "cmdCustUpdate"
        Me.cmdCustUpdate.Size = New System.Drawing.Size(150, 40)
        Me.cmdCustUpdate.TabIndex = 4
        Me.cmdCustUpdate.Text = "Update"
        Me.cmdCustUpdate.UseVisualStyleBackColor = True
        '
        'cmdCustsearch
        '
        Me.cmdCustsearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCustsearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCustsearch.Location = New System.Drawing.Point(25, 68)
        Me.cmdCustsearch.Name = "cmdCustsearch"
        Me.cmdCustsearch.Size = New System.Drawing.Size(150, 40)
        Me.cmdCustsearch.TabIndex = 3
        Me.cmdCustsearch.Text = "Search"
        Me.cmdCustsearch.UseVisualStyleBackColor = True
        '
        'cmdCustAdd
        '
        Me.cmdCustAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCustAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCustAdd.Location = New System.Drawing.Point(25, 22)
        Me.cmdCustAdd.Name = "cmdCustAdd"
        Me.cmdCustAdd.Size = New System.Drawing.Size(150, 40)
        Me.cmdCustAdd.TabIndex = 2
        Me.cmdCustAdd.Text = "Add"
        Me.cmdCustAdd.UseVisualStyleBackColor = True
        '
        'grpPawning
        '
        Me.grpPawning.Controls.Add(Me.cmdPawnExpired)
        Me.grpPawning.Controls.Add(Me.cmdPawnIn)
        Me.grpPawning.Controls.Add(Me.cmdPawnOut)
        Me.grpPawning.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.grpPawning.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpPawning.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.grpPawning.Location = New System.Drawing.Point(215, 27)
        Me.grpPawning.Name = "grpPawning"
        Me.grpPawning.Size = New System.Drawing.Size(197, 166)
        Me.grpPawning.TabIndex = 4
        Me.grpPawning.TabStop = False
        Me.grpPawning.Text = " Pawning: "
        '
        'cmdPawnExpired
        '
        Me.cmdPawnExpired.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPawnExpired.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPawnExpired.Location = New System.Drawing.Point(24, 114)
        Me.cmdPawnExpired.Name = "cmdPawnExpired"
        Me.cmdPawnExpired.Size = New System.Drawing.Size(150, 40)
        Me.cmdPawnExpired.TabIndex = 4
        Me.cmdPawnExpired.Text = "Expired"
        Me.cmdPawnExpired.UseVisualStyleBackColor = True
        '
        'cmdPawnIn
        '
        Me.cmdPawnIn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPawnIn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPawnIn.Location = New System.Drawing.Point(24, 68)
        Me.cmdPawnIn.Name = "cmdPawnIn"
        Me.cmdPawnIn.Size = New System.Drawing.Size(150, 40)
        Me.cmdPawnIn.TabIndex = 3
        Me.cmdPawnIn.Text = "In"
        Me.cmdPawnIn.UseVisualStyleBackColor = True
        '
        'cmdPawnOut
        '
        Me.cmdPawnOut.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPawnOut.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPawnOut.Location = New System.Drawing.Point(24, 22)
        Me.cmdPawnOut.Name = "cmdPawnOut"
        Me.cmdPawnOut.Size = New System.Drawing.Size(150, 40)
        Me.cmdPawnOut.TabIndex = 2
        Me.cmdPawnOut.Text = "Out"
        Me.cmdPawnOut.UseVisualStyleBackColor = True
        '
        'grpCashtrans
        '
        Me.grpCashtrans.Controls.Add(Me.cmdCashOther)
        Me.grpCashtrans.Controls.Add(Me.cmdCashSale)
        Me.grpCashtrans.Controls.Add(Me.cmdCashStock)
        Me.grpCashtrans.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.grpCashtrans.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpCashtrans.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.grpCashtrans.Location = New System.Drawing.Point(418, 27)
        Me.grpCashtrans.Name = "grpCashtrans"
        Me.grpCashtrans.Size = New System.Drawing.Size(197, 166)
        Me.grpCashtrans.TabIndex = 5
        Me.grpCashtrans.TabStop = False
        Me.grpCashtrans.Text = " Cash Transactions: "
        '
        'cmdCashOther
        '
        Me.cmdCashOther.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCashOther.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCashOther.Location = New System.Drawing.Point(20, 114)
        Me.cmdCashOther.Name = "cmdCashOther"
        Me.cmdCashOther.Size = New System.Drawing.Size(150, 40)
        Me.cmdCashOther.TabIndex = 4
        Me.cmdCashOther.Text = "Other Expenses"
        Me.cmdCashOther.UseVisualStyleBackColor = True
        '
        'cmdCashSale
        '
        Me.cmdCashSale.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCashSale.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCashSale.Location = New System.Drawing.Point(20, 68)
        Me.cmdCashSale.Name = "cmdCashSale"
        Me.cmdCashSale.Size = New System.Drawing.Size(150, 40)
        Me.cmdCashSale.TabIndex = 3
        Me.cmdCashSale.Text = "Sale (Pawn && Stock)"
        Me.cmdCashSale.UseVisualStyleBackColor = True
        '
        'cmdCashStock
        '
        Me.cmdCashStock.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCashStock.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCashStock.Location = New System.Drawing.Point(20, 22)
        Me.cmdCashStock.Name = "cmdCashStock"
        Me.cmdCashStock.Size = New System.Drawing.Size(150, 40)
        Me.cmdCashStock.TabIndex = 2
        Me.cmdCashStock.Text = "Stock Purchase"
        Me.cmdCashStock.UseVisualStyleBackColor = True
        '
        'grpReports
        '
        Me.grpReports.Controls.Add(Me.Label2)
        Me.grpReports.Controls.Add(Me.Label1)
        Me.grpReports.Controls.Add(Me.dtReportTo)
        Me.grpReports.Controls.Add(Me.dtReportFrom)
        Me.grpReports.Controls.Add(Me.cmdReportCustom)
        Me.grpReports.Controls.Add(Me.cmdReportDaily)
        Me.grpReports.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.grpReports.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpReports.ForeColor = System.Drawing.Color.DarkRed
        Me.grpReports.Location = New System.Drawing.Point(12, 199)
        Me.grpReports.Name = "grpReports"
        Me.grpReports.Size = New System.Drawing.Size(299, 155)
        Me.grpReports.TabIndex = 6
        Me.grpReports.TabStop = False
        Me.grpReports.Text = " Reports: "
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(22, 67)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 17)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "End Date:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(22, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 17)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Start Date:"
        '
        'dtReportTo
        '
        Me.dtReportTo.Location = New System.Drawing.Point(124, 62)
        Me.dtReportTo.Name = "dtReportTo"
        Me.dtReportTo.Size = New System.Drawing.Size(160, 23)
        Me.dtReportTo.TabIndex = 3
        '
        'dtReportFrom
        '
        Me.dtReportFrom.Location = New System.Drawing.Point(124, 33)
        Me.dtReportFrom.Name = "dtReportFrom"
        Me.dtReportFrom.Size = New System.Drawing.Size(160, 23)
        Me.dtReportFrom.TabIndex = 2
        '
        'cmdReportCustom
        '
        Me.cmdReportCustom.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReportCustom.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdReportCustom.Location = New System.Drawing.Point(151, 102)
        Me.cmdReportCustom.Name = "cmdReportCustom"
        Me.cmdReportCustom.Size = New System.Drawing.Size(120, 40)
        Me.cmdReportCustom.TabIndex = 5
        Me.cmdReportCustom.Text = "Custom Report"
        Me.cmdReportCustom.UseVisualStyleBackColor = True
        '
        'cmdReportDaily
        '
        Me.cmdReportDaily.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReportDaily.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdReportDaily.Location = New System.Drawing.Point(25, 102)
        Me.cmdReportDaily.Name = "cmdReportDaily"
        Me.cmdReportDaily.Size = New System.Drawing.Size(120, 40)
        Me.cmdReportDaily.TabIndex = 4
        Me.cmdReportDaily.Text = "Daily Report"
        Me.cmdReportDaily.UseVisualStyleBackColor = True
        '
        'grpPawnHistory
        '
        Me.grpPawnHistory.Controls.Add(Me.Label3)
        Me.grpPawnHistory.Controls.Add(Me.Label4)
        Me.grpPawnHistory.Controls.Add(Me.dtPawnTo)
        Me.grpPawnHistory.Controls.Add(Me.dtPawnFrom)
        Me.grpPawnHistory.Controls.Add(Me.cmdPawnHistory)
        Me.grpPawnHistory.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.grpPawnHistory.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpPawnHistory.ForeColor = System.Drawing.Color.DarkRed
        Me.grpPawnHistory.Location = New System.Drawing.Point(317, 199)
        Me.grpPawnHistory.Name = "grpPawnHistory"
        Me.grpPawnHistory.Size = New System.Drawing.Size(299, 155)
        Me.grpPawnHistory.TabIndex = 7
        Me.grpPawnHistory.TabStop = False
        Me.grpPawnHistory.Text = " Pawning History:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(21, 67)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 17)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "End Date:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(19, 38)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(76, 17)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Start Date:"
        '
        'dtPawnTo
        '
        Me.dtPawnTo.Location = New System.Drawing.Point(121, 62)
        Me.dtPawnTo.Name = "dtPawnTo"
        Me.dtPawnTo.Size = New System.Drawing.Size(160, 23)
        Me.dtPawnTo.TabIndex = 3
        '
        'dtPawnFrom
        '
        Me.dtPawnFrom.Location = New System.Drawing.Point(121, 33)
        Me.dtPawnFrom.Name = "dtPawnFrom"
        Me.dtPawnFrom.Size = New System.Drawing.Size(160, 23)
        Me.dtPawnFrom.TabIndex = 2
        '
        'cmdPawnHistory
        '
        Me.cmdPawnHistory.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPawnHistory.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPawnHistory.Location = New System.Drawing.Point(68, 102)
        Me.cmdPawnHistory.Name = "cmdPawnHistory"
        Me.cmdPawnHistory.Size = New System.Drawing.Size(150, 40)
        Me.cmdPawnHistory.TabIndex = 4
        Me.cmdPawnHistory.Text = "Pawning History"
        Me.cmdPawnHistory.UseVisualStyleBackColor = True
        '
        'grpStockMngmnt
        '
        Me.grpStockMngmnt.Controls.Add(Me.cmdStockTango)
        Me.grpStockMngmnt.Controls.Add(Me.cmdStockStnd)
        Me.grpStockMngmnt.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.grpStockMngmnt.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpStockMngmnt.ForeColor = System.Drawing.Color.Green
        Me.grpStockMngmnt.Location = New System.Drawing.Point(12, 360)
        Me.grpStockMngmnt.Name = "grpStockMngmnt"
        Me.grpStockMngmnt.Size = New System.Drawing.Size(299, 119)
        Me.grpStockMngmnt.TabIndex = 8
        Me.grpStockMngmnt.TabStop = False
        Me.grpStockMngmnt.Text = " Stock Management: "
        '
        'cmdStockTango
        '
        Me.cmdStockTango.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStockTango.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStockTango.Location = New System.Drawing.Point(25, 68)
        Me.cmdStockTango.Name = "cmdStockTango"
        Me.cmdStockTango.Size = New System.Drawing.Size(246, 40)
        Me.cmdStockTango.TabIndex = 3
        Me.cmdStockTango.Text = "Tango Stock"
        Me.cmdStockTango.UseVisualStyleBackColor = True
        '
        'cmdStockStnd
        '
        Me.cmdStockStnd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStockStnd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStockStnd.Location = New System.Drawing.Point(25, 22)
        Me.cmdStockStnd.Name = "cmdStockStnd"
        Me.cmdStockStnd.Size = New System.Drawing.Size(246, 40)
        Me.cmdStockStnd.TabIndex = 2
        Me.cmdStockStnd.Text = "Standard Stock"
        Me.cmdStockStnd.UseVisualStyleBackColor = True
        '
        'grpStock
        '
        Me.grpStock.Controls.Add(Me.cmdStackPawnRpt)
        Me.grpStock.Controls.Add(Me.cmdStockIncome)
        Me.grpStock.Controls.Add(Me.cmdStockReport)
        Me.grpStock.Controls.Add(Me.cmdStockSearch)
        Me.grpStock.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.grpStock.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpStock.ForeColor = System.Drawing.Color.Green
        Me.grpStock.Location = New System.Drawing.Point(316, 360)
        Me.grpStock.Name = "grpStock"
        Me.grpStock.Size = New System.Drawing.Size(299, 119)
        Me.grpStock.TabIndex = 9
        Me.grpStock.TabStop = False
        Me.grpStock.Text = " Stock: "
        '
        'cmdStackPawnRpt
        '
        Me.cmdStackPawnRpt.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStackPawnRpt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStackPawnRpt.Location = New System.Drawing.Point(151, 68)
        Me.cmdStackPawnRpt.Name = "cmdStackPawnRpt"
        Me.cmdStackPawnRpt.Size = New System.Drawing.Size(120, 40)
        Me.cmdStackPawnRpt.TabIndex = 5
        Me.cmdStackPawnRpt.Text = "Pawn Report"
        Me.cmdStackPawnRpt.UseVisualStyleBackColor = True
        '
        'cmdStockIncome
        '
        Me.cmdStockIncome.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStockIncome.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStockIncome.Location = New System.Drawing.Point(151, 22)
        Me.cmdStockIncome.Name = "cmdStockIncome"
        Me.cmdStockIncome.Size = New System.Drawing.Size(120, 40)
        Me.cmdStockIncome.TabIndex = 3
        Me.cmdStockIncome.Text = "Real. Income"
        Me.cmdStockIncome.UseVisualStyleBackColor = True
        '
        'cmdStockReport
        '
        Me.cmdStockReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStockReport.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStockReport.Location = New System.Drawing.Point(25, 68)
        Me.cmdStockReport.Name = "cmdStockReport"
        Me.cmdStockReport.Size = New System.Drawing.Size(120, 40)
        Me.cmdStockReport.TabIndex = 4
        Me.cmdStockReport.Text = "Stock Report"
        Me.cmdStockReport.UseVisualStyleBackColor = True
        '
        'cmdStockSearch
        '
        Me.cmdStockSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStockSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStockSearch.Location = New System.Drawing.Point(25, 22)
        Me.cmdStockSearch.Name = "cmdStockSearch"
        Me.cmdStockSearch.Size = New System.Drawing.Size(120, 40)
        Me.cmdStockSearch.TabIndex = 2
        Me.cmdStockSearch.Text = "Search"
        Me.cmdStockSearch.UseVisualStyleBackColor = True
        '
        'grpStockCode
        '
        Me.grpStockCode.Controls.Add(Me.cmdStockCdTango)
        Me.grpStockCode.Controls.Add(Me.cmdStockCdStnd)
        Me.grpStockCode.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.grpStockCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpStockCode.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.grpStockCode.Location = New System.Drawing.Point(12, 485)
        Me.grpStockCode.Name = "grpStockCode"
        Me.grpStockCode.Size = New System.Drawing.Size(299, 119)
        Me.grpStockCode.TabIndex = 10
        Me.grpStockCode.TabStop = False
        Me.grpStockCode.Text = " Stock Code List: "
        '
        'cmdStockCdTango
        '
        Me.cmdStockCdTango.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStockCdTango.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStockCdTango.Location = New System.Drawing.Point(25, 68)
        Me.cmdStockCdTango.Name = "cmdStockCdTango"
        Me.cmdStockCdTango.Size = New System.Drawing.Size(246, 40)
        Me.cmdStockCdTango.TabIndex = 3
        Me.cmdStockCdTango.Text = "Tango Stock"
        Me.cmdStockCdTango.UseVisualStyleBackColor = True
        '
        'cmdStockCdStnd
        '
        Me.cmdStockCdStnd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStockCdStnd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStockCdStnd.Location = New System.Drawing.Point(25, 22)
        Me.cmdStockCdStnd.Name = "cmdStockCdStnd"
        Me.cmdStockCdStnd.Size = New System.Drawing.Size(246, 40)
        Me.cmdStockCdStnd.TabIndex = 2
        Me.cmdStockCdStnd.Text = "Standard Stock"
        Me.cmdStockCdStnd.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdExit
        Me.ClientSize = New System.Drawing.Size(627, 642)
        Me.Controls.Add(Me.grpStockCode)
        Me.Controls.Add(Me.grpStock)
        Me.Controls.Add(Me.grpStockMngmnt)
        Me.Controls.Add(Me.grpPawnHistory)
        Me.Controls.Add(Me.grpReports)
        Me.Controls.Add(Me.grpCashtrans)
        Me.Controls.Add(Me.grpPawning)
        Me.Controls.Add(Me.grpCustomer)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.Settings)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.Settings
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cash&BuyBack"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.Settings.ResumeLayout(False)
        Me.Settings.PerformLayout()
        Me.grpCustomer.ResumeLayout(False)
        Me.grpPawning.ResumeLayout(False)
        Me.grpCashtrans.ResumeLayout(False)
        Me.grpReports.ResumeLayout(False)
        Me.grpReports.PerformLayout()
        Me.grpPawnHistory.ResumeLayout(False)
        Me.grpPawnHistory.PerformLayout()
        Me.grpStockMngmnt.ResumeLayout(False)
        Me.grpStock.ResumeLayout(False)
        Me.grpStockCode.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents SS1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents SS2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents SS3 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Settings As System.Windows.Forms.MenuStrip
    Friend WithEvents SettingsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DBLoactionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CashRegisterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents MaintenanceToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DatabaseMaintenanceToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ClearHistoryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ChangeTangoEntriesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents grpCustomer As System.Windows.Forms.GroupBox
    Friend WithEvents cmdCustUpdate As System.Windows.Forms.Button
    Friend WithEvents cmdCustsearch As System.Windows.Forms.Button
    Friend WithEvents cmdCustAdd As System.Windows.Forms.Button
    Friend WithEvents grpPawning As System.Windows.Forms.GroupBox
    Friend WithEvents cmdPawnExpired As System.Windows.Forms.Button
    Friend WithEvents cmdPawnIn As System.Windows.Forms.Button
    Friend WithEvents cmdPawnOut As System.Windows.Forms.Button
    Friend WithEvents grpCashtrans As System.Windows.Forms.GroupBox
    Friend WithEvents cmdCashOther As System.Windows.Forms.Button
    Friend WithEvents cmdCashSale As System.Windows.Forms.Button
    Friend WithEvents cmdCashStock As System.Windows.Forms.Button
    Friend WithEvents grpReports As System.Windows.Forms.GroupBox
    Friend WithEvents cmdReportCustom As System.Windows.Forms.Button
    Friend WithEvents cmdReportDaily As System.Windows.Forms.Button
    Friend WithEvents grpPawnHistory As System.Windows.Forms.GroupBox
    Friend WithEvents cmdPawnHistory As System.Windows.Forms.Button
    Friend WithEvents dtReportFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtReportTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtPawnTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPawnFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents grpStockMngmnt As System.Windows.Forms.GroupBox
    Friend WithEvents cmdStockTango As System.Windows.Forms.Button
    Friend WithEvents cmdStockStnd As System.Windows.Forms.Button
    Friend WithEvents grpStock As System.Windows.Forms.GroupBox
    Friend WithEvents cmdStockReport As System.Windows.Forms.Button
    Friend WithEvents cmdStockSearch As System.Windows.Forms.Button
    Friend WithEvents cmdStackPawnRpt As System.Windows.Forms.Button
    Friend WithEvents cmdStockIncome As System.Windows.Forms.Button
    Friend WithEvents grpStockCode As System.Windows.Forms.GroupBox
    Friend WithEvents cmdStockCdTango As System.Windows.Forms.Button
    Friend WithEvents cmdStockCdStnd As System.Windows.Forms.Button

End Class
