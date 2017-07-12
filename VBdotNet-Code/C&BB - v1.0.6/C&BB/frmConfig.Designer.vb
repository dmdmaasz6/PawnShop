<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConfig
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConfig))
        Me.txtShop = New System.Windows.Forms.TextBox()
        Me.txtCashRegister = New System.Windows.Forms.TextBox()
        Me.txtDBLocation = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.cmdFile = New System.Windows.Forms.Button()
        Me.txtDate = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.rbnAuditY = New System.Windows.Forms.RadioButton()
        Me.rbnAuditN = New System.Windows.Forms.RadioButton()
        Me.cmdFile2 = New System.Windows.Forms.Button()
        Me.txtAgreeLocation = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.SuspendLayout()
        '
        'txtShop
        '
        Me.txtShop.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShop.Location = New System.Drawing.Point(170, 183)
        Me.txtShop.Name = "txtShop"
        Me.txtShop.Size = New System.Drawing.Size(361, 23)
        Me.txtShop.TabIndex = 13
        '
        'txtCashRegister
        '
        Me.txtCashRegister.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCashRegister.Location = New System.Drawing.Point(170, 139)
        Me.txtCashRegister.Name = "txtCashRegister"
        Me.txtCashRegister.Size = New System.Drawing.Size(361, 23)
        Me.txtCashRegister.TabIndex = 12
        '
        'txtDBLocation
        '
        Me.txtDBLocation.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDBLocation.Location = New System.Drawing.Point(170, 55)
        Me.txtDBLocation.Name = "txtDBLocation"
        Me.txtDBLocation.Size = New System.Drawing.Size(361, 23)
        Me.txtDBLocation.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(65, 186)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 17)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Shop:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(65, 142)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(101, 17)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Cash Register:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(65, 55)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(89, 17)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "DB Location:"
        '
        'cmdOK
        '
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.Location = New System.Drawing.Point(154, 361)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(100, 40)
        Me.cmdOK.TabIndex = 7
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.Location = New System.Drawing.Point(350, 361)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(100, 40)
        Me.cmdCancel.TabIndex = 14
        Me.cmdCancel.Text = "C&ancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(65, 19)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(448, 17)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "This is the current values. To change, edit the values, and press 'OK'."
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'cmdFile
        '
        Me.cmdFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFile.Location = New System.Drawing.Point(537, 48)
        Me.cmdFile.Name = "cmdFile"
        Me.cmdFile.Size = New System.Drawing.Size(27, 27)
        Me.cmdFile.TabIndex = 16
        Me.cmdFile.Text = "..."
        Me.cmdFile.UseVisualStyleBackColor = True
        '
        'txtDate
        '
        Me.txtDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDate.Location = New System.Drawing.Point(170, 316)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.ReadOnly = True
        Me.txtDate.Size = New System.Drawing.Size(361, 23)
        Me.txtDate.TabIndex = 18
        Me.txtDate.TabStop = False
        Me.txtDate.Text = "Change this in the PC's Control Panel, Region settings..."
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(64, 319)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(90, 17)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Date Setting:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(64, 273)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(69, 17)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Currency:"
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(170, 270)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(361, 23)
        Me.TextBox1.TabIndex = 21
        Me.TextBox1.TabStop = False
        Me.TextBox1.Text = "Change this in the PC's Control Panel, Region settings..."
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(65, 228)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(123, 17)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "Write to Audit File:"
        '
        'rbnAuditY
        '
        Me.rbnAuditY.AutoSize = True
        Me.rbnAuditY.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.rbnAuditY.Location = New System.Drawing.Point(242, 225)
        Me.rbnAuditY.Name = "rbnAuditY"
        Me.rbnAuditY.Size = New System.Drawing.Size(50, 21)
        Me.rbnAuditY.TabIndex = 24
        Me.rbnAuditY.TabStop = True
        Me.rbnAuditY.Text = "Yes"
        Me.rbnAuditY.UseVisualStyleBackColor = True
        '
        'rbnAuditN
        '
        Me.rbnAuditN.AutoSize = True
        Me.rbnAuditN.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.rbnAuditN.Location = New System.Drawing.Point(328, 226)
        Me.rbnAuditN.Name = "rbnAuditN"
        Me.rbnAuditN.Size = New System.Drawing.Size(44, 21)
        Me.rbnAuditN.TabIndex = 25
        Me.rbnAuditN.TabStop = True
        Me.rbnAuditN.Text = "No"
        Me.rbnAuditN.UseVisualStyleBackColor = True
        '
        'cmdFile2
        '
        Me.cmdFile2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFile2.Location = New System.Drawing.Point(537, 90)
        Me.cmdFile2.Name = "cmdFile2"
        Me.cmdFile2.Size = New System.Drawing.Size(27, 27)
        Me.cmdFile2.TabIndex = 28
        Me.cmdFile2.Text = "..."
        Me.cmdFile2.UseVisualStyleBackColor = True
        '
        'txtAgreeLocation
        '
        Me.txtAgreeLocation.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAgreeLocation.Location = New System.Drawing.Point(170, 94)
        Me.txtAgreeLocation.Name = "txtAgreeLocation"
        Me.txtAgreeLocation.Size = New System.Drawing.Size(361, 23)
        Me.txtAgreeLocation.TabIndex = 27
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(65, 90)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(89, 45)
        Me.Label8.TabIndex = 26
        Me.Label8.Text = "Agreement Location:"
        '
        'frmConfig
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(591, 433)
        Me.Controls.Add(Me.cmdFile2)
        Me.Controls.Add(Me.txtAgreeLocation)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.rbnAuditN)
        Me.Controls.Add(Me.rbnAuditY)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtDate)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.cmdFile)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.txtShop)
        Me.Controls.Add(Me.txtCashRegister)
        Me.Controls.Add(Me.txtDBLocation)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdOK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmConfig"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Change Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtShop As System.Windows.Forms.TextBox
    Friend WithEvents txtCashRegister As System.Windows.Forms.TextBox
    Friend WithEvents txtDBLocation As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cmdFile As System.Windows.Forms.Button
    Friend WithEvents txtDate As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents rbnAuditY As System.Windows.Forms.RadioButton
    Friend WithEvents rbnAuditN As System.Windows.Forms.RadioButton
    Friend WithEvents cmdFile2 As System.Windows.Forms.Button
    Friend WithEvents txtAgreeLocation As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
End Class
