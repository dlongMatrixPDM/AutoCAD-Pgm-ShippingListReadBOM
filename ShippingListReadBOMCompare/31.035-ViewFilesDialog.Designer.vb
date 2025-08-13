<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class ViewFilesDialog
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
	Public WithEvents Filelist As System.Windows.Forms.ListBox
	Public WithEvents CommandButton1 As System.Windows.Forms.Button
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Filelist = New System.Windows.Forms.ListBox()
        Me.CommandButton1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Filelist
        '
        Me.Filelist.BackColor = System.Drawing.SystemColors.Window
        Me.Filelist.Cursor = System.Windows.Forms.Cursors.Default
        Me.Filelist.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Filelist.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Filelist.ItemHeight = 16
        Me.Filelist.Location = New System.Drawing.Point(8, 8)
        Me.Filelist.Name = "Filelist"
        Me.Filelist.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Filelist.Size = New System.Drawing.Size(279, 116)
        Me.Filelist.TabIndex = 0
        '
        'CommandButton1
        '
        Me.CommandButton1.BackColor = System.Drawing.SystemColors.Control
        Me.CommandButton1.Cursor = System.Windows.Forms.Cursors.Default
        Me.CommandButton1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CommandButton1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CommandButton1.Location = New System.Drawing.Point(224, 128)
        Me.CommandButton1.Name = "CommandButton1"
        Me.CommandButton1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CommandButton1.Size = New System.Drawing.Size(64, 32)
        Me.CommandButton1.TabIndex = 1
        Me.CommandButton1.Text = "&Ok"
        Me.CommandButton1.UseVisualStyleBackColor = False
        '
        'ViewFilesDialog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(291, 165)
        Me.Controls.Add(Me.Filelist)
        Me.Controls.Add(Me.CommandButton1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(3, 15)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ViewFilesDialog"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class