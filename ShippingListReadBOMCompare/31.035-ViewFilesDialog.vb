Option Strict On           'Option Strict Off
Option Explicit On
Friend Class ViewFilesDialog
	Inherits System.Windows.Forms.Form
	
    Private Sub UserForm_Initialize()
        Dim BOMMnu As ShippingList_Menu
        Dim BOMorShip As String
        BOMMnu = ShippingList_Menu

        Me.Text = "Bulk BOM Files in: " & BOMMnu.File1.Text
        BOMorShip = "*BOM*.XLS*"
        Call ViewFiles(BOMorShip)
    End Sub
	
	Private Sub CommandButton1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CommandButton1.Click
		Me.Close()
	End Sub
End Class