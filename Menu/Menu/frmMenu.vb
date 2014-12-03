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
            Dim iRow As Integer = 0

            ' Nombre de la compañia
            sheet(iRow, 0).SetValue("Fresh Software Concepts, L.L.C")
            sheet.Cells(iRow, 0).Font.Bold = True
            sheet.Cells(iRow, 0).Font.Size = 14
            iRow += 1

            ' Load Data
            sheet(iRow, 0).SetValue("Load Date: 18/10/2014")
            iRow += 2

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
            sheet(iRow, 0).SetValue("Description")
            sheet(iRow, 1).SetValue("Packer/Grower")
            sheet(iRow, 2).SetValue("Gwr.Lot")
            sheet(iRow, 3).SetValue("Pallet-Tag")
            sheet(iRow, 4).SetValue("Truck")
            sheet(iRow, 5).SetValue("Arv.Date")
            sheet(iRow, 6).SetValue("Pkgs")
            iRow += 1

            ' Colocamos la referencia
            sheet(iRow, 0).SetValue("B/L " & allocations.AsEnumerable().FirstOrDefault()("Reference"))
            iRow += 1

            For Each OrderKey As Long In linq.ToArray()
                Dim linq2 = (From row In allocations.AsEnumerable()
                             Where Long.Parse(row.Field(Of String)("OrderKey")) = OrderKey
                             Select dt.LoadDataRow(New [Object]() {row.Field(Of String)("ProductDesc"), row.Field(Of String)("PackerGrowerDesc"), row.Field(Of String)("GrowerLot"), row.Field(Of String)("PalletTagNum"), row.Field(Of String)("Truck"), row.Field(Of String)("ArvDate"), row.Field(Of String)("PkgsShp")}, False)).CopyToDataTable()

                sheet.Import(linq2, False, iRow, 0)
                iRow += linq2.Rows.Count
                sheet(iRow, 0).SetValue("TOTAL LINE")
                sheet(iRow, 6).SetValue(linq2.AsEnumerable().Sum(Function(order) Int32.Parse(order.Field(Of String)("PkgsShp"))))

                iRow += 1
            Next

            sheet(iRow, 0).SetValue("TOTAL B/L")
            sheet(iRow, 6).SetValue(allocations.AsEnumerable().Sum(Function(order) Int32.Parse(order.Field(Of String)("PkgsShp"))))

            sheet.GetUsedRange().AutoFitRows()
            sheet.GetUsedRange().AutoFitColumns()

            frmExcel.Show()
        ElseIf radioGroup1.SelectedIndex = 4 Then
            Dim frmExcel As New frmExcel()
            frmExcel.AddHeader = True
            frmExcel.FirstColIndex = 0
            frmExcel.FirstRowIndex = 0
            frmExcel.Create()

            Dim iRow As Integer = 0
            Dim iCol As Integer = 0
            Dim ShowAudit As Boolean = True
            Dim _IsPageOrientationPortrait As Boolean = False

            Dim sheet As Worksheet = frmExcel.Worksheet(0)

            ' Leemos y obtenemos el DataTable
            Dim Aging1 As New DataTable("Aging1")
            Dim frmExcel2 As New frmExcel("D:\Escritorio\clsAging1.xlsx")
            frmExcel2.Create()
            Aging1 = frmExcel2.ExportToDataTable(0)

            ' Agregar opciones de impresion

            With sheet
                .Name = "AR Aging By Grower"
                .FreezeRows(6, .Range("A1"))

                ' Hacemos una celda combinada
                .MergeCells(.Range("A1:B1"))
                .Cells(iRow, 0).SetValue("Fresh Software Concepts, L.L.C")
                .Cells(iRow, 0).Font.Bold = True
                .Cells(iRow, 0).Font.Size = 14
                .Cells(iRow, 0).Style.Alignment.Vertical = SpreadsheetVerticalAlignment.Center
                iRow += 1

                ' Hacemos una celda combinada
                .MergeCells(.Range("A2:B2"))
                .Cells(iRow, 0).SetValue("Grower-Aging Report")
                .Cells(iRow, 0).Font.Bold = True
                .Cells(iRow, 0).Font.Size = 12
                .Cells(iRow, 0).Style.Alignment.Vertical = SpreadsheetVerticalAlignment.Center
                iRow += 1

                ' Hacemos una celda combinada
                .MergeCells(.Range("A3:B3"))
                .Cells(iRow, 0).SetValue("Report Date: 13/06/2013 (Packers Format)")
                .Cells(iRow, 0).Font.Bold = True
                .Cells(iRow, 0).Font.Size = 12
                .Cells(iRow, 0).Style.Alignment.Vertical = SpreadsheetVerticalAlignment.Center
                iRow += 2

                Dim headers As String() = New [String]() {"Grower/Customer", "Invoice#", "Ld.Date", "Inv.Date", "Due Date", "Invoiced$", "Adj$", "Balance", "00-21", "22-35", "36-45", "46-60", "61- OVERGrower/Customer", "Invoice#", "Ld.Date", "Inv.Date", "Due Date", "Invoiced$", "Adj$", "Balance", "00-21", "22-35", "36-45", "46-60", "61- OVER"}

                For Each header As String In headers
                    .Cells(iRow, iCol).SetValue(header)
                    .Cells(iRow, iCol).Style.Font.Bold = True
                    iCol += 1
                Next
                .Cells(iRow, 0).Style.Font.Size = 12
                iRow += 1

                With .ActiveView
                    .ShowGridlines = False

                    If _IsPageOrientationPortrait Then
                        .Orientation = PageOrientation.Portrait
                    Else
                        .Orientation = PageOrientation.Landscape
                    End If

                End With

                .Import(Aging1, False, 7, 0)
            End With


            'sheet = frmExcel.Worksheet(1, True)

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