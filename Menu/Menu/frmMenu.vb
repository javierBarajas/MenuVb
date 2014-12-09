Imports DevExpress.Spreadsheet
'Imports System.Globalization
'Imports System.Threading

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
        ElseIf radioGroup1.SelectedIndex = 4 Then ' Grower's Aging
            'Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US")
            Dim frmExcel As New frmExcel()
            frmExcel.AddHeader = True
            frmExcel.FirstColIndex = 0
            frmExcel.FirstRowIndex = 0
            frmExcel.Create()

            Dim iRow As Integer = 0
            Dim iRowFilter As Integer = 0
            Dim iCol As Integer = 0
            Dim ShowAudit As Boolean = True
            Dim _IsPageOrientationPortrait As Boolean = False
            Dim ARBalance As Double = 2660136.27
            Dim PackerDesc As String
            Dim FormulaTotal As String = ""

            Dim sheet As Worksheet = frmExcel.Worksheet(0, True)

            ' Leemos y obtenemos el DataTable
            Dim Aging As New DataTable("Aging")
            Dim frmExcel2 As New frmExcel("D:\Escritorio\clsAging.xlsx")
            frmExcel2.Create()
            Aging = frmExcel2.ExportToDataTable(0)

            Dim Aging1 As New DataTable("Aging1")
            frmExcel2 = New frmExcel("D:\Escritorio\clsAging1.xlsx")
            frmExcel2.Create()
            Aging1 = frmExcel2.ExportToDataTable(0)

            ' Obtenemos todos los packers
            Dim Packers = (From row In Aging.AsEnumerable()
                           Group row By PackerID = row.Field(Of String)("PackerID") Into PackerIDGroup = Group
                           Select PackerID).ToList()

            ' DataTable para agregar la info a la hoja
            Dim dt As DataTable = New DataTable()
            With dt.Columns
                .Add("Customer")
                .Add("InvoiceNumSq")
                '.Add("LoadDate", Type.GetType("System.DateTime"))
                .Add("LoadDate")
                '.Add("InvoiceDate", Type.GetType("System.DateTime"))
                .Add("InvoiceDate")
                '.Add("DueDate", Type.GetType("System.DateTime"))
                .Add("DueDate")
                .Add("InvoiceAmt", Type.GetType("System.Decimal"))
                .Add("AdjustmentsAmt", Type.GetType("System.Decimal"))
                .Add("Balance", Type.GetType("System.Decimal"))
                .Add("Pd0Amount", Type.GetType("System.Decimal"))
                .Add("Pd1Amount", Type.GetType("System.Decimal"))
                .Add("Pd2Amount", Type.GetType("System.Decimal"))
                .Add("Pd3Amount", Type.GetType("System.Decimal"))
                .Add("Pd4Amount", Type.GetType("System.Decimal"))
            End With

            ' Agregar opciones de impresion

            With sheet
                ' Nombre de la hoja
                .Name = "AR Aging By Grower"

                ' Inmovilizamos 7 filas
                .FreezeRows(6, .Range("A1"))

                ' Hacemos una celda combinada
                .MergeCells(.Range("A1:B1"))
                .Cells(iRow, 0).SetValue("Fresh Software Concepts, L.L.C")
                .Cells(iRow, 0).Font.Bold = True
                .Cells(iRow, 0).Font.Size = 14
                .Cells(iRow, 0).Alignment.Vertical = SpreadsheetVerticalAlignment.Center
                iRow += 1

                ' Hacemos una celda combinada
                .MergeCells(.Range("A2:B2"))
                .Cells(iRow, 0).SetValue("Grower-Aging Report")
                .Cells(iRow, 0).Font.Bold = True
                .Cells(iRow, 0).Font.Size = 12
                .Cells(iRow, 0).Alignment.Vertical = SpreadsheetVerticalAlignment.Center
                iRow += 1

                ' Hacemos una celda combinada
                .MergeCells(.Range("A3:B3"))
                .Cells(iRow, 0).SetValue("Report Date: 13/06/2013 (Packers Format)")
                .Cells(iRow, 0).Font.Bold = True
                .Cells(iRow, 0).Font.Size = 12
                .Cells(iRow, 0).Alignment.Vertical = SpreadsheetVerticalAlignment.Center
                iRow += 2

                Dim headers As String() = New [String]() {"Grower/Customer", "Invoice#", "Ld.Date", "Inv.Date", "Due Date", "Invoiced$", "Adj$", "Balance", "00-21", "22-35", "36-45", "46-60", "61- OVER"}

                For Each header As String In headers
                    .Cells(iRow, iCol).SetValue(header)
                    .Cells(iRow, iCol).Font.Bold = True
                    iCol += 1
                Next
                .Cells(iRow, 0).Font.Size = 12
                iRow += 1

                With .ActiveView
                    ' Ocultamos las lineas del grid
                    .ShowGridlines = False

                    If _IsPageOrientationPortrait Then
                        .Orientation = PageOrientation.Portrait
                    Else
                        .Orientation = PageOrientation.Landscape
                    End If

                End With

                ' Espacio en blanco
                iRow += 1

                ' Fila para los filtros
                iRowFilter = iRow
                iRow += 1

                For Each PackerID As String In Packers
                    PackerDesc = (From row In Aging.AsEnumerable()
                                  Where (row.Field(Of String)("PackerID").Equals(PackerID))
                                  Group row By Packer = row.Field(Of String)("Packer") Into PackerGroup = Group
                                  Select Packer).FirstOrDefault()

                    .Cells(iRow, 0).SetValue(PackerDesc)
                    .Cells(iRow, 0).Font.Bold = True
                    iRow += 1

                    Dim data = (From row In Aging.AsEnumerable()
                                Where row.Field(Of String)("PackerID").Equals(PackerID)
                                Select dt.LoadDataRow(New [Object]() {row.Field(Of String)("Customer"), row.Field(Of String)("InvoiceNumSq"), row.Field(Of String)("LoadDate"), row.Field(Of String)("InvoiceDate"), row.Field(Of String)("DueDate"), row.Field(Of String)("InvoiceAmt"), row.Field(Of String)("AdjustmentsAmt"), row.Field(Of String)("Balance"), row.Field(Of String)("Pd0Amount"), row.Field(Of String)("Pd1Amount"), row.Field(Of String)("Pd2Amount"), row.Field(Of String)("Pd3Amount"), row.Field(Of String)("Pd4Amount")}, False)).CopyToDataTable()

                    sheet.Import(data, False, iRow, 0)
                    .Cells(iRow + data.Rows.Count, 0).SetValue(PackerDesc & " Total: ")
                    .Cells(iRow + data.Rows.Count, 0).Font.Bold = True
                    .Cells(iRow + data.Rows.Count, 5).Formula = String.Format("=SUMA(F{0}:F{1})", (iRow + 1), (iRow + data.Rows.Count))
                    .Cells(iRow + data.Rows.Count, 6).Formula = String.Format("=SUMA(G{0}:G{1})", (iRow + 1), (iRow + data.Rows.Count))
                    .Cells(iRow + data.Rows.Count, 7).Formula = String.Format("=SUMA(H{0}:H{1})", (iRow + 1), (iRow + data.Rows.Count))
                    .Cells(iRow + data.Rows.Count, 8).Formula = String.Format("=SUMA(I{0}:I{1})", (iRow + 1), (iRow + data.Rows.Count))
                    .Cells(iRow + data.Rows.Count, 9).Formula = String.Format("=SUMA(J{0}:J{1})", (iRow + 1), (iRow + data.Rows.Count))
                    .Cells(iRow + data.Rows.Count, 10).Formula = String.Format("=SUMA(K{0}:K{1})", (iRow + 1), (iRow + data.Rows.Count))
                    .Cells(iRow + data.Rows.Count, 11).Formula = String.Format("=SUMA(L{0}:L{1})", (iRow + 1), (iRow + data.Rows.Count))
                    .Cells(iRow + data.Rows.Count, 12).Formula = String.Format("=SUMA(M{0}:M{1})", (iRow + 1), (iRow + data.Rows.Count))

                    FormulaTotal += "{0}" & (iRow + data.Rows.Count + 1) & "+"

                    iRow += data.Rows.Count + 1
                Next

                ' Quitamos el ultimo signo
                FormulaTotal = FormulaTotal.Substring(0, FormulaTotal.Length - 1)

                .Cells(iRow, 0).SetValue("GRAND TOTAL")
                .Cells(iRow, 0).Font.Bold = True
                .Cells(iRow, 5).Formula = String.Format(FormulaTotal, "F")
                .Cells(iRow, 6).Formula = String.Format(FormulaTotal, "G")
                .Cells(iRow, 7).Formula = String.Format(FormulaTotal, "H")
                .Cells(iRow, 8).Formula = String.Format(FormulaTotal, "I")
                .Cells(iRow, 9).Formula = String.Format(FormulaTotal, "J")
                .Cells(iRow, 10).Formula = String.Format(FormulaTotal, "K")
                .Cells(iRow, 11).Formula = String.Format(FormulaTotal, "L")
                .Cells(iRow, 12).Formula = String.Format(FormulaTotal, "M")

                ' Hacemos los filtros
                .AutoFilter.Apply(.Range.FromLTRB(0, iRowFilter, (headers.Length - 1), iRow))

                '' LoadDate
                '.Columns(2).NumberFormat = "mm/dd/yyyy"
                '.Columns(2).Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left

                '' InvoiceDate
                '.Columns(3).NumberFormat = "mm/dd/yyyy"
                '.Columns(3).Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left

                '' DueDate
                '.Columns(4).NumberFormat = "mm/dd/yyyy"
                '.Columns(4).Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left

                ' InvoiceAmt, AdjustmentsAmt, Balance, Pd0Amount, Pd1Amount, Pd2Amount, Pd3Amount, Pd4Amount
                .Range("F:M").NumberFormat = "#,##0.00;[Red]\(#,##0.00)"
                .Range("F:M").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right

                .GetUsedRange().AutoFitRows()
                .GetUsedRange().AutoFitColumns()


            End With

            If ShowAudit Then
                Dim sheet2 As Worksheet = frmExcel.Worksheet(1, True)
                Dim FormulaSubTotal As String = Aging.AsEnumerable().Sum(Function(row) Double.Parse(row.Field(Of String)("Balance"))).ToString()
                Dim FormulaDiference As String = ""

                With sheet2
                    .Name = "Audit"

                    .Cells(0, 0).SetValue("Grower's Balance")
                    .Cells(0, 0).Font.Bold = True
                    .Cells(0, 1).Formula = "='" & sheet.Name & "'!H" & (iRow + 1)
                    .Cells(0, 1).Font.Bold = True
                    iRow = 1

                    .Cells(iRow, 0).SetValue("UpChargesPrice")
                    .Cells(iRow, 1).SetValue(Aging.AsEnumerable().Sum(Function(row) Double.Parse(row.Field(Of String)("UpChargesAmount"))))
                    iRow += 1
                    FormulaSubTotal += "+B" & iRow

                    .Cells(iRow, 0).SetValue("FreightPrice")
                    .Cells(iRow, 1).SetValue(Aging.AsEnumerable().Sum(Function(row) Double.Parse(row.Field(Of String)("FreightAmount"))))
                    iRow += 1
                    FormulaSubTotal += "+B" & iRow

                    .Cells(iRow, 0).SetValue("MiscePrice")
                    .Cells(iRow, 1).SetValue(Aging.AsEnumerable().Sum(Function(row) Double.Parse(row.Field(Of String)("MisceChgAmount"))))
                    iRow += 1
                    FormulaSubTotal += "+B" & iRow

                    .Cells(iRow, 0).SetValue("Pallets")
                    .Cells(iRow, 1).SetValue(Aging.AsEnumerable().Sum(Function(row) Double.Parse(row.Field(Of String)("PalletAmount"))))
                    iRow += 1
                    FormulaSubTotal += "+B" & iRow

                    .Cells(iRow, 0).SetValue("Temp. Recorder")
                    .Cells(iRow, 1).SetValue(Aging1.AsEnumerable().Sum(Function(row) Double.Parse(row.Field(Of String)("TempRec"))))
                    iRow += 1
                    FormulaSubTotal += "+B" & iRow

                    .Cells(iRow, 0).SetValue("Miscellaneous")
                    .Cells(iRow, 1).SetValue(Aging1.AsEnumerable().Sum(Function(row) Double.Parse(row.Field(Of String)("Misce"))))
                    iRow += 1
                    FormulaSubTotal += "+B" & iRow

                    .Cells(iRow, 0).SetValue("Temp. Recorder Adj.")
                    .Cells(iRow, 1).SetValue(Aging1.AsEnumerable().Sum(Function(row) Double.Parse(row.Field(Of String)("TempRecAdjustment"))) * -1)
                    iRow += 1
                    FormulaSubTotal += "+B" & iRow

                    .Cells(iRow, 0).SetValue("MisceLlaneous Adj.")
                    .Cells(iRow, 1).SetValue(Aging1.AsEnumerable().Sum(Function(row) Double.Parse(row.Field(Of String)("MisceAdjustment"))) * -1)
                    iRow += 1
                    FormulaSubTotal += "+B" & iRow

                    .Cells(iRow, 0).SetValue("TOTAL Wt. OTHERS")
                    .Cells(iRow, 0).Font.Bold = True
                    .Cells(iRow, 1).Formula = "=SUMA(B1:B" & iRow & ")"
                    .Cells(iRow, 1).Font.Bold = True
                    iRow += 1

                    .Cells(iRow, 0).SetValue("Partial Payments")
                    .Cells(iRow, 1).SetValue(Aging1.AsEnumerable().Sum(Function(row) Double.Parse(row.Field(Of String)("PartialPay"))) * -1)
                    iRow += 1
                    FormulaSubTotal += "+B" & iRow

                    .Cells(iRow, 0).SetValue("Sub Total")
                    .Cells(iRow, 0).Font.Bold = True
                    .Cells(iRow, 1).Formula = FormulaSubTotal
                    .Cells(iRow, 1).Font.Bold = True
                    iRow += 1
                    FormulaDiference = "=B" & iRow

                    .Cells(iRow, 0).SetValue("TOTAL A/R Aging")
                    .Cells(iRow, 1).SetValue(ARBalance)
                    iRow += 1
                    FormulaDiference += "-B" & iRow

                    .Cells(iRow, 0).SetValue("DIFERENCE")
                    .Cells(iRow, 0).Font.Bold = True
                    .Cells(iRow, 1).Formula = FormulaDiference
                    .Cells(iRow, 1).Font.Bold = True
                    iRow += 1

                    .Columns(1).NumberFormat = "#,##0.00;[Red]\(#,##0.00)"
                    .Columns(1).Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right

                    .GetUsedRange().AutoFitRows()
                    .GetUsedRange().AutoFitColumns()
                End With

                frmExcel.ActiveWorksheet(0)
            End If

            frmExcel.Show()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim frmWord As New frmWord()
        frmWord.Show()
    End Sub

    Private Sub radioGroup1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radioGroup1.SelectedIndexChanged
        Me.btnExcel.Enabled = True
    End Sub
End Class