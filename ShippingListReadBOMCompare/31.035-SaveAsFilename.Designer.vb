<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class SaveAsFilename
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents TextBox1 As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents SaveButton As System.Windows.Forms.Button
	Public WithEvents CancelButton_Renamed As System.Windows.Forms.Button
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.SaveButton = New System.Windows.Forms.Button
        Me.CancelButton_Renamed = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.AcceptsReturn = True
        Me.TextBox1.BackColor = System.Drawing.SystemColors.Window
        Me.TextBox1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TextBox1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TextBox1.Location = New System.Drawing.Point(128, 8)
        Me.TextBox1.MaxLength = 0
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TextBox1.Size = New System.Drawing.Size(128, 24)
        Me.TextBox1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(112, 33)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Shipping List Filename:"
        '
        'SaveButton
        '
        Me.SaveButton.BackColor = System.Drawing.SystemColors.Control
        Me.SaveButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.SaveButton.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SaveButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SaveButton.Location = New System.Drawing.Point(160, 40)
        Me.SaveButton.Name = "SaveButton"
        Me.SaveButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SaveButton.Size = New System.Drawing.Size(96, 32)
        Me.SaveButton.TabIndex = 2
        Me.SaveButton.Text = "&Save"
        Me.SaveButton.UseVisualStyleBackColor = False
        '
        'CancelButton_Renamed
        '
        Me.CancelButton_Renamed.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton_Renamed.Cursor = System.Windows.Forms.Cursors.Default
        Me.CancelButton_Renamed.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CancelButton_Renamed.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CancelButton_Renamed.Location = New System.Drawing.Point(160, 80)
        Me.CancelButton_Renamed.Name = "CancelButton_Renamed"
        Me.CancelButton_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CancelButton_Renamed.Size = New System.Drawing.Size(96, 32)
        Me.CancelButton_Renamed.TabIndex = 3
        Me.CancelButton_Renamed.Text = "&Cancel"
        Me.CancelButton_Renamed.UseVisualStyleBackColor = False
        '
        'SaveAsFilename
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(264, 119)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.SaveButton)
        Me.Controls.Add(Me.CancelButton_Renamed)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(3, 15)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SaveAsFilename"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class