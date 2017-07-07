<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTangoStockManagment
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
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.dgTangoStockManagment = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CBB_DataSet = New CashBuyBack.CBB_DataSet()
        Me.CBBDataSetBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.grpAdd = New System.Windows.Forms.GroupBox()
        Me.txtPrice = New System.Windows.Forms.TextBox()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.grpUpdate = New System.Windows.Forms.GroupBox()
        Me.txtPriceUpd = New System.Windows.Forms.TextBox()
        Me.txtDescriptionUpd = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmdUpdate = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.grpDelete = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.NormalStockBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        CType(Me.dgTangoStockManagment, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CBB_DataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CBBDataSetBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpAdd.SuspendLayout()
        Me.grpUpdate.SuspendLayout()
        Me.grpDelete.SuspendLayout()
        CType(Me.NormalStockBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdExit
        '
        Me.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.Location = New System.Drawing.Point(254, 597)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(150, 40)
        Me.cmdExit.TabIndex = 1
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'dgTangoStockManagment
        '
        Me.dgTangoStockManagment.AllowUserToAddRows = False
        Me.dgTangoStockManagment.AllowUserToDeleteRows = False
        Me.dgTangoStockManagment.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgTangoStockManagment.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgTangoStockManagment.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgTangoStockManagment.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column3, Me.Column4, Me.Column5})
        Me.dgTangoStockManagment.Location = New System.Drawing.Point(7, 9)
        Me.dgTangoStockManagment.MultiSelect = False
        Me.dgTangoStockManagment.Name = "dgTangoStockManagment"
        Me.dgTangoStockManagment.RowHeadersVisible = False
        Me.dgTangoStockManagment.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgTangoStockManagment.Size = New System.Drawing.Size(641, 203)
        Me.dgTangoStockManagment.TabIndex = 3
        '
        'Column1
        '
        Me.Column1.HeaderText = "NSID"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
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
        '
        'Column4
        '
        Me.Column4.HeaderText = "Qty Available"
        Me.Column4.Name = "Column4"
        Me.Column4.ReadOnly = True
        '
        'Column5
        '
        Me.Column5.HeaderText = "Qty on LayBuy"
        Me.Column5.Name = "Column5"
        Me.Column5.ReadOnly = True
        '
        'CBB_DataSet
        '
        Me.CBB_DataSet.DataSetName = "CBB_DataSet"
        Me.CBB_DataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'CBBDataSetBindingSource
        '
        Me.CBBDataSetBindingSource.DataMember = "NormalStock"
        Me.CBBDataSetBindingSource.DataSource = Me.CBB_DataSet
        '
        'grpAdd
        '
        Me.grpAdd.Controls.Add(Me.txtPrice)
        Me.grpAdd.Controls.Add(Me.txtDescription)
        Me.grpAdd.Controls.Add(Me.Label4)
        Me.grpAdd.Controls.Add(Me.Label3)
        Me.grpAdd.Controls.Add(Me.cmdAdd)
        Me.grpAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpAdd.Location = New System.Drawing.Point(7, 237)
        Me.grpAdd.Name = "grpAdd"
        Me.grpAdd.Size = New System.Drawing.Size(645, 90)
        Me.grpAdd.TabIndex = 4
        Me.grpAdd.TabStop = False
        Me.grpAdd.Text = " Add: "
        '
        'txtPrice
        '
        Me.txtPrice.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPrice.Location = New System.Drawing.Point(137, 48)
        Me.txtPrice.Name = "txtPrice"
        Me.txtPrice.Size = New System.Drawing.Size(106, 23)
        Me.txtPrice.TabIndex = 3
        '
        'txtDescription
        '
        Me.txtDescription.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDescription.Location = New System.Drawing.Point(137, 19)
        Me.txtDescription.MaxLength = 50
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(358, 23)
        Me.txtDescription.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label4.Location = New System.Drawing.Point(44, 51)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(94, 17)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Selling Price: "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label3.Location = New System.Drawing.Point(44, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(87, 17)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Description: "
        '
        'cmdAdd
        '
        Me.cmdAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdd.Location = New System.Drawing.Point(519, 28)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(105, 40)
        Me.cmdAdd.TabIndex = 4
        Me.cmdAdd.Text = "Add"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'grpUpdate
        '
        Me.grpUpdate.Controls.Add(Me.txtPriceUpd)
        Me.grpUpdate.Controls.Add(Me.txtDescriptionUpd)
        Me.grpUpdate.Controls.Add(Me.Label5)
        Me.grpUpdate.Controls.Add(Me.Label6)
        Me.grpUpdate.Controls.Add(Me.cmdUpdate)
        Me.grpUpdate.Controls.Add(Me.Label2)
        Me.grpUpdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpUpdate.Location = New System.Drawing.Point(7, 333)
        Me.grpUpdate.Name = "grpUpdate"
        Me.grpUpdate.Size = New System.Drawing.Size(645, 158)
        Me.grpUpdate.TabIndex = 5
        Me.grpUpdate.TabStop = False
        Me.grpUpdate.Text = " Update: "
        '
        'txtPriceUpd
        '
        Me.txtPriceUpd.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPriceUpd.Location = New System.Drawing.Point(137, 115)
        Me.txtPriceUpd.Name = "txtPriceUpd"
        Me.txtPriceUpd.Size = New System.Drawing.Size(106, 23)
        Me.txtPriceUpd.TabIndex = 6
        '
        'txtDescriptionUpd
        '
        Me.txtDescriptionUpd.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDescriptionUpd.Location = New System.Drawing.Point(137, 86)
        Me.txtDescriptionUpd.MaxLength = 50
        Me.txtDescriptionUpd.Name = "txtDescriptionUpd"
        Me.txtDescriptionUpd.Size = New System.Drawing.Size(358, 23)
        Me.txtDescriptionUpd.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label5.Location = New System.Drawing.Point(44, 118)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(94, 17)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Selling Price: "
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label6.Location = New System.Drawing.Point(44, 89)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(87, 17)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Description: "
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUpdate.Location = New System.Drawing.Point(519, 95)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(105, 40)
        Me.cmdUpdate.TabIndex = 7
        Me.cmdUpdate.Text = "Update"
        Me.cmdUpdate.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label2.Location = New System.Drawing.Point(44, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(553, 38)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "To change existing information, select it in the grid, and change the information" & _
    " below.  Only Description and Selling Price can be changed."
        '
        'grpDelete
        '
        Me.grpDelete.Controls.Add(Me.Label1)
        Me.grpDelete.Controls.Add(Me.cmdDelete)
        Me.grpDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grpDelete.Location = New System.Drawing.Point(7, 497)
        Me.grpDelete.Name = "grpDelete"
        Me.grpDelete.Size = New System.Drawing.Size(645, 82)
        Me.grpDelete.TabIndex = 6
        Me.grpDelete.TabStop = False
        Me.grpDelete.Text = " Delete: "
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label1.Location = New System.Drawing.Point(44, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(367, 17)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "To delete one of the entries, select it and click on Delete."
        '
        'cmdDelete
        '
        Me.cmdDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDelete.Location = New System.Drawing.Point(519, 22)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(105, 40)
        Me.cmdDelete.TabIndex = 8
        Me.cmdDelete.Text = "Delete"
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'NormalStockBindingSource
        '
        Me.NormalStockBindingSource.DataMember = "NormalStock"
        Me.NormalStockBindingSource.DataSource = Me.CBB_DataSet
        '
        'frmTangoStockManagment
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdExit
        Me.ClientSize = New System.Drawing.Size(660, 649)
        Me.Controls.Add(Me.grpDelete)
        Me.Controls.Add(Me.grpUpdate)
        Me.Controls.Add(Me.grpAdd)
        Me.Controls.Add(Me.dgTangoStockManagment)
        Me.Controls.Add(Me.cmdExit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmTangoStockManagment"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Tango Stock Management"
        CType(Me.dgTangoStockManagment, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CBB_DataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CBBDataSetBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpAdd.ResumeLayout(False)
        Me.grpAdd.PerformLayout()
        Me.grpUpdate.ResumeLayout(False)
        Me.grpUpdate.PerformLayout()
        Me.grpDelete.ResumeLayout(False)
        Me.grpDelete.PerformLayout()
        CType(Me.NormalStockBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents dgTangoStockManagment As System.Windows.Forms.DataGridView
    Friend WithEvents grpAdd As System.Windows.Forms.GroupBox
    Friend WithEvents grpUpdate As System.Windows.Forms.GroupBox
    Friend WithEvents grpDelete As System.Windows.Forms.GroupBox
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents CBBDataSetBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents CBB_DataSet As CashBuyBack.CBB_DataSet
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NormalStockBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtPrice As System.Windows.Forms.TextBox
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents txtPriceUpd As System.Windows.Forms.TextBox
    Friend WithEvents txtDescriptionUpd As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmdUpdate As System.Windows.Forms.Button
End Class
