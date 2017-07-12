<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSettings
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSettings))
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDBLocation = New System.Windows.Forms.TextBox()
        Me.txtCashRegister = New System.Windows.Forms.TextBox()
        Me.txtShop = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtDate = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtCurrency = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.rbnAuditN = New System.Windows.Forms.RadioButton()
        Me.rbnAuditY = New System.Windows.Forms.RadioButton()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtAgreeLocation = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdOK
        '
        Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.Location = New System.Drawing.Point(212, 361)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(100, 40)
        Me.cmdOK.TabIndex = 0
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(35, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(89, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "DB Location:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(35, 120)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(101, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Cash Register:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(35, 164)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 17)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Shop:"
        '
        'txtDBLocation
        '
        Me.txtDBLocation.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDBLocation.Location = New System.Drawing.Point(141, 24)
        Me.txtDBLocation.Name = "txtDBLocation"
        Me.txtDBLocation.ReadOnly = True
        Me.txtDBLocation.Size = New System.Drawing.Size(361, 23)
        Me.txtDBLocation.TabIndex = 4
        Me.txtDBLocation.TabStop = False
        '
        'txtCashRegister
        '
        Me.txtCashRegister.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCashRegister.Location = New System.Drawing.Point(141, 117)
        Me.txtCashRegister.Name = "txtCashRegister"
        Me.txtCashRegister.ReadOnly = True
        Me.txtCashRegister.Size = New System.Drawing.Size(361, 23)
        Me.txtCashRegister.TabIndex = 5
        Me.txtCashRegister.TabStop = False
        '
        'txtShop
        '
        Me.txtShop.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShop.Location = New System.Drawing.Point(141, 161)
        Me.txtShop.Name = "txtShop"
        Me.txtShop.ReadOnly = True
        Me.txtShop.Size = New System.Drawing.Size(361, 23)
        Me.txtShop.TabIndex = 6
        Me.txtShop.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Red
        Me.Label4.Location = New System.Drawing.Point(39, 321)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(375, 16)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "The PC Short Date format, must be set to MM/dd/yyyy"
        '
        'txtDate
        '
        Me.txtDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDate.Location = New System.Drawing.Point(141, 295)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.ReadOnly = True
        Me.txtDate.Size = New System.Drawing.Size(361, 23)
        Me.txtDate.TabIndex = 9
        Me.txtDate.TabStop = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(35, 298)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(90, 17)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Date Setting:"
        '
        'txtCurrency
        '
        Me.txtCurrency.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCurrency.Location = New System.Drawing.Point(141, 247)
        Me.txtCurrency.Name = "txtCurrency"
        Me.txtCurrency.ReadOnly = True
        Me.txtCurrency.Size = New System.Drawing.Size(361, 23)
        Me.txtCurrency.TabIndex = 11
        Me.txtCurrency.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(35, 250)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(69, 17)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Currency:"
        '
        'rbnAuditN
        '
        Me.rbnAuditN.AutoSize = True
        Me.rbnAuditN.Enabled = False
        Me.rbnAuditN.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.rbnAuditN.Location = New System.Drawing.Point(298, 204)
        Me.rbnAuditN.Name = "rbnAuditN"
        Me.rbnAuditN.Size = New System.Drawing.Size(44, 21)
        Me.rbnAuditN.TabIndex = 28
        Me.rbnAuditN.TabStop = True
        Me.rbnAuditN.Text = "No"
        Me.rbnAuditN.UseVisualStyleBackColor = True
        '
        'rbnAuditY
        '
        Me.rbnAuditY.AutoSize = True
        Me.rbnAuditY.Enabled = False
        Me.rbnAuditY.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.rbnAuditY.Location = New System.Drawing.Point(212, 205)
        Me.rbnAuditY.Name = "rbnAuditY"
        Me.rbnAuditY.Size = New System.Drawing.Size(50, 21)
        Me.rbnAuditY.TabIndex = 27
        Me.rbnAuditY.TabStop = True
        Me.rbnAuditY.Text = "Yes"
        Me.rbnAuditY.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(35, 208)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(123, 17)
        Me.Label7.TabIndex = 26
        Me.Label7.Text = "Write to Audit File:"
        '
        'txtAgreeLocation
        '
        Me.txtAgreeLocation.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAgreeLocation.Location = New System.Drawing.Point(141, 68)
        Me.txtAgreeLocation.Name = "txtAgreeLocation"
        Me.txtAgreeLocation.ReadOnly = True
        Me.txtAgreeLocation.Size = New System.Drawing.Size(361, 23)
        Me.txtAgreeLocation.TabIndex = 30
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(36, 64)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(89, 45)
        Me.Label8.TabIndex = 29
        Me.Label8.Text = "Agreement Location:"
        '
        'frmSettings
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(522, 414)
        Me.Controls.Add(Me.txtAgreeLocation)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.rbnAuditN)
        Me.Controls.Add(Me.rbnAuditY)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtCurrency)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtDate)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
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
        Me.Name = "frmSettings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDBLocation As System.Windows.Forms.TextBox
    Friend WithEvents txtCashRegister As System.Windows.Forms.TextBox
    Friend WithEvents txtShop As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtDate As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtCurrency As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents rbnAuditN As System.Windows.Forms.RadioButton
    Friend WithEvents rbnAuditY As System.Windows.Forms.RadioButton
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtAgreeLocation As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
End Class
