Option Strict Off
Option Explicit On

Imports Excel = Microsoft.Office.Interop.Excel.Application

<System.Runtime.InteropServices.ProgId("ThisWorkbook_NET.ThisWorkbook")> Public Class ThisWorkbook

    Private Sub Workbook_Open()
        Dim Excel As Object

        If UCase(Left(Excel.ActiveWorkbook.Name, 8)) = "BULK BOM" Then
            Excel.Application.Visible = False
            ShippingList_Menu.BulkBom()             'Comparison.InputType3.BulkBom()
        End If

    End Sub
End Class