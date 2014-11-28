Imports DevExpress.Spreadsheet

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
            ' ShipAllocations
        ElseIf radioGroup1.SelectedIndex = 3 Then
            Dim frmExcel As New frmExcel()
            frmExcel.AddHeader = True
            frmExcel.FirstColIndex = 0
            frmExcel.FirstRowIndex = 0
            frmExcel.Create()

            Dim sheet As Worksheet = frmExcel.Worksheet(0)
            Dim rowIndex As Integer = 0

            ' Nombre de la compañia
            sheet(rowIndex, 0).SetValue("Fresh Software Concepts, L.L.C")
            sheet.Cells(rowIndex, 0).Font.Bold = True
            sheet.Cells(rowIndex, 0).Font.Size = 14
            rowIndex += 1

            ' Load Data
            sheet(rowIndex, 0).SetValue("Load Date: 18/10/2014")
            rowIndex += 2

            ' Leemos y obtenemos el DataTable
            Dim allocations As New DataTable("Allocations")
            Dim frmExcel2 As New frmExcel("D:\Escritorio\clsAllocations.xlsx")
            frmExcel2.Create()
            allocations = frmExcel2.ExportToDataTable(0)

            Dim linq = (From row In allocations.AsEnumerable()
                        Group row By OrderKey = Long.Parse(row.Field(Of String)("OrderKey")) Into OrderKeyGroup = Group
                        Select OrderKey).ToList()

            Dim dt As DataTable = New DataTable()
            With dt.Columns
                .Add("ProductDesc")
                .Add("PackerGrowerDesc")
                .Add("GrowerLot")
                .Add("PalletTagNum")
                .Add("Truck")
                .Add("ArvDate", Type.GetType("System.DateTime"))
                .Add("PkgsShp")
            End With

            sheet.Columns(5).NumberFormat = "mm/dd/yyyy"
            sheet.Columns(6).NumberFormat = "#,##0;(#,##0)"
            sheet.Columns(6).Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center

            ' Colocamos los encabezados
            sheet(rowIndex, 0).SetValue("Description")
            sheet(rowIndex, 1).SetValue("Packer/Grower")
            sheet(rowIndex, 2).SetValue("Gwr.Lot")
            sheet(rowIndex, 3).SetValue("Pallet-Tag")
            sheet(rowIndex, 4).SetValue("Truck")
            sheet(rowIndex, 5).SetValue("Arv.Date")
            sheet(rowIndex, 6).SetValue("Pkgs")
            rowIndex += 1

            ' Colocamos la referencia
            sheet(rowIndex, 0).SetValue("B/L " & allocations.AsEnumerable().FirstOrDefault()("Reference"))
            rowIndex += 1

            For Each OrderKey As Long In linq.ToArray()
                Dim linq2 = (From row In allocations.AsEnumerable()
                             Where Long.Parse(row.Field(Of String)("OrderKey")) = OrderKey
                             Select dt.LoadDataRow(New [Object]() {row.Field(Of String)("ProductDesc"), row.Field(Of String)("PackerGrowerDesc"), row.Field(Of String)("GrowerLot"), row.Field(Of String)("PalletTagNum"), row.Field(Of String)("Truck"), row.Field(Of String)("ArvDate"), row.Field(Of String)("PkgsShp")}, False)).CopyToDataTable()

                sheet.Import(linq2, False, rowIndex, 0)
                rowIndex += linq2.Rows.Count
                sheet(rowIndex, 0).SetValue("TOTAL LINE")
                sheet(rowIndex, 6).SetValue(linq2.AsEnumerable().Sum(Function(order) Int32.Parse(order.Field(Of String)("PkgsShp"))))

                rowIndex += 1
            Next

            sheet(rowIndex, 0).SetValue("TOTAL B/L")
            sheet(rowIndex, 6).SetValue(allocations.AsEnumerable().Sum(Function(order) Int32.Parse(order.Field(Of String)("PkgsShp"))))

            sheet.GetUsedRange().AutoFitRows()
            sheet.GetUsedRange().AutoFitColumns()

            frmExcel.Show()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim frmExcel As New frmExcel(OpenFileDialog1.FileName)
            frmExcel.Create()
            Dim frmWord As New frmWord(frmExcel.ExportToDataTable(0))
            frmWord.Show()
        End If
    End Sub

    Private Sub radioGroup1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radioGroup1.SelectedIndexChanged
        Me.btnExcel.Enabled = True
    End Sub
End Class