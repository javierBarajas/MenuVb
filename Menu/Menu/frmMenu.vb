Public Class frmMenu
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        If radioGroup1.SelectedIndex = 0 Then
            Dim sourceTable As New DataTable("Products")
            sourceTable.Columns.Add("Product", GetType([String]))
            sourceTable.Columns.Add("Price", GetType([Single]))
            sourceTable.Columns.Add("Quantity", GetType(Int32))
            sourceTable.Columns.Add("Discount", GetType([Single]))

            sourceTable.Rows.Add("Chocolade", 5, 15, 0.03)
            sourceTable.Rows.Add("Konbu", 9, 55, 0.1)
            sourceTable.Rows.Add("Geitost", 15, 70, 0.07)

            Dim frmExcel As New frmExcel(sourceTable)
            frmExcel.AddHeader = True
            frmExcel.FirstColIndex = 0
            frmExcel.FirstRowIndex = 0
            frmExcel.Create()
            frmExcel.Show()
        ElseIf radioGroup1.SelectedIndex = 1 Then
            Dim sourceTable As New DataTable("Products")
            sourceTable.Columns.Add("Product", GetType([String]))
            sourceTable.Columns.Add("Price", GetType([Single]))
            sourceTable.Columns.Add("Quantity", GetType(Int32))
            sourceTable.Columns.Add("Discount", GetType([Single]))

            sourceTable.Rows.Add("Chocolade", 5, 15, 0.03)
            sourceTable.Rows.Add("Konbu", 9, 55, 0.1)
            sourceTable.Rows.Add("Geitost", 15, 70, 0.07)

            Dim ds As New DataSet()
            ds.Tables.Add(sourceTable)

            sourceTable = New DataTable("Names")
            sourceTable.Columns.Add("Name", GetType([String]))
            sourceTable.Columns.Add("Age", GetType(Int32))

            sourceTable.Rows.Add("Eduardo", 24)
            sourceTable.Rows.Add("Omar", 24)
            sourceTable.Rows.Add("Lupe", 23)
            ds.Tables.Add(sourceTable)

            Dim frmExcel As New frmExcel(ds)
            frmExcel.AddHeader = True
            frmExcel.FirstColIndex = 0
            frmExcel.FirstRowIndex = 0
            frmExcel.Create()
            frmExcel.Worksheet(0).Name = "Funciona"
            frmExcel.Show()
        ElseIf radioGroup1.SelectedIndex = 2 Then
            If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Dim frmExcel As New frmExcel(OpenFileDialog1.FileName)
                frmExcel.AddHeader = True
                frmExcel.FirstColIndex = 0
                frmExcel.FirstRowIndex = 0
                frmExcel.Create()
                Dim dt As DataTable = frmExcel.ExportToDataTable(0)

                frmExcel = New frmExcel(dt)
                frmExcel.Create()
                frmExcel.Show()
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim frmExcel As New frmExcel(OpenFileDialog1.FileName)
            frmExcel.Create() 
            Dim frmWord As New frmWord(frmExcel.ExportToDataTable())
            frmWord.Show()
        End If
    End Sub

    Private Sub radioGroup1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radioGroup1.SelectedIndexChanged
        Me.btnExcel.Enabled = True
    End Sub
End Class
