Imports DevExpress.Spreadsheet
Public Class frmExcel
    Private Sub SpreadsheetControl1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SpreadsheetControl1.Click
        Dim workbook As IWorkbook = SpreadsheetControl1.Document
        Dim worksheet As Worksheet = workbook.Worksheets(0)
        frmWord.Show()
    End Sub

    Private Sub frmExcel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class