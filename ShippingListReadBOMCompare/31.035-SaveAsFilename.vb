Option Strict On           'Option Strict Off
Option Explicit On
Friend Class SaveAsFilename
	Inherits System.Windows.Forms.Form
	
	Private Sub UserForm_Initialize()
		Me.TextBox1.Focus()
	End Sub
	
	Private Sub SaveButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SaveButton.Click
        If TextBox1.Text <> vbNullString Then
            PassFilename = TextBox1.Text
            ReadyToContinue = True
        Else
            ReadyToContinue = False
        End If
		Me.Close()
	End Sub
	
	Private Sub CancelButton_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CancelButton_Renamed.Click
		PassFilename = "CancelProgram"
		ReadyToContinue = False
		Me.Close()
	End Sub
End Class