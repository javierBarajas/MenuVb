﻿Imports System.IO
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Export

Public Class frmExcel

#Region "Variables"
    Private path As String = Nothing
    Private dt As DataTable = Nothing
    Private ds As DataSet = Nothing
    Private _AddHeader As Boolean = Nothing
    Private _FirstRowIndex As Integer = Nothing
    Private _FirstColIndex As Integer = Nothing
#End Region

#Region "Properties"
    Public WriteOnly Property AddHeader() As Boolean
        Set(ByVal Value As Boolean)
            _AddHeader = Value
        End Set
    End Property

    Public WriteOnly Property FirstRowIndex() As Integer
        Set(ByVal Value As Integer)
            _FirstRowIndex = Value
        End Set
    End Property

    Public WriteOnly Property FirstColIndex() As Integer
        Set(ByVal Value As Integer)
            _FirstColIndex = Value
        End Set
    End Property
#End Region

#Region "Constructor"
    Private Sub Initialize(ByVal path As [String], ByVal dt As DataTable, ByVal ds As DataSet)
        InitializeComponent()

        Me.path = path
        Me.dt = dt
        Me.ds = ds

        Me.AddHeader = True
        Me.FirstColIndex = 0
        Me.FirstRowIndex = 0
    End Sub

    Public Sub New(ByVal path As String)
        Initialize(path, Nothing, Nothing)
    End Sub

    Public Sub New(ByVal dt As DataTable)
        Initialize(Nothing, dt, Nothing)
    End Sub

    Public Sub New(ByVal ds As DataSet)
        Initialize(Nothing, Nothing, ds)
    End Sub

    Public Sub New()
        Initialize(Nothing, Nothing, Nothing)
    End Sub
#End Region

    Public Function Worksheet(ByVal index As Integer, Optional ByVal AddNew As Boolean = False) As Worksheet
        If index >= 0 AndAlso index < SpreadsheetControl1.Document.Worksheets.Count Then
            SpreadsheetControl1.Document.Worksheets.ActiveWorksheet = SpreadsheetControl1.Document.Worksheets(index)
            Return SpreadsheetControl1.Document.Worksheets(index)
        ElseIf AddNew Then
            Dim NewWorksheet As Worksheet = SpreadsheetControl1.Document.Worksheets.Add()
            SpreadsheetControl1.Document.Worksheets.ActiveWorksheet = NewWorksheet
            Return NewWorksheet
        End If

        Return Nothing
    End Function

    Public Sub ActiveWorksheet(ByVal index As Integer)
        If index >= 0 AndAlso index < SpreadsheetControl1.Document.Worksheets.Count Then
            SpreadsheetControl1.Document.Worksheets.ActiveWorksheet = SpreadsheetControl1.Document.Worksheets(index)
        End If
    End Sub

    Public Function ExportToDataTable(ByVal index As Integer) As DataTable
        If index >= 0 AndAlso index < SpreadsheetControl1.Document.Worksheets.Count Then
            Dim worksheet As Worksheet = SpreadsheetControl1.Document.Worksheets(index)
            Dim dataTable As DataTable = worksheet.CreateDataTable(worksheet.GetUsedRange(), True, True)
            Dim exporter As DataTableExporter = worksheet.CreateDataTableExporter(worksheet.GetUsedRange(), dataTable, True)
            'exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = True
            AddHandler exporter.CellValueConversionError, AddressOf exporter_CellValueConversionError
            exporter.Export()

            dataTable.TableName = worksheet.Name
            Return dataTable
        End If

        Return Nothing
    End Function

    Private Sub exporter_CellValueConversionError(ByVal sender As Object, ByVal e As CellValueConversionErrorEventArgs)
        'MessageBox.Show("Error in cell " & e.Cell.GetReferenceA1())
        e.DataTableValue = Nothing
        e.Action = DataTableExporterAction.Continue
    End Sub

    Public Sub Create()
        If dt IsNot Nothing Then
            SpreadsheetControl1.Document.Styles.Add("HeaderStyle")
            Dim worksheet As Worksheet = SpreadsheetControl1.Document.Worksheets(0)


            'TableStyleCollection tableStyle = spreadsheetControl1.Document.TableStyles;

            worksheet.Name = dt.TableName

            'Table table = worksheet.Tables.Add(worksheet.Range.FromLTRB(FirstColIndex, FirstRowIndex, dt.Columns.Count - 1, dt.Rows.Count), true);
            '                table.Style = spreadsheetControl1.Document.TableStyles[BuiltInTableStyleId.TableStyleLight9];

            worksheet.Import(dt, _AddHeader, _FirstRowIndex, _FirstColIndex)
            worksheet.GetUsedRange().AutoFitRows()
            worksheet.GetUsedRange().AutoFitColumns()
        ElseIf ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            Dim worksheetCollection As WorksheetCollection = SpreadsheetControl1.Document.Worksheets
            Dim worksheet As Worksheet = Nothing

            For i As Integer = 0 To ds.Tables.Count - 1
                worksheet = Nothing

                If i < worksheetCollection.Count Then
                    worksheet = worksheetCollection(i)
                Else
                    worksheet = worksheetCollection.Add()
                End If

                worksheet.Name = ds.Tables(i).TableName

                'Table table = worksheet.Tables.Add(worksheet.Range.FromLTRB(FirstColIndex, FirstRowIndex, ds.Tables[i].Columns.Count - 1, ds.Tables[i].Rows.Count), true);
                '                    table.Style = spreadsheetControl1.Document.TableStyles[BuiltInTableStyleId.TableStyleLight9];

                worksheet.Import(ds.Tables(i), _AddHeader, _FirstRowIndex, _FirstColIndex)
                worksheet.GetUsedRange().AutoFitRows()
                worksheet.GetUsedRange().AutoFitColumns()
            Next
        ElseIf path IsNot Nothing AndAlso File.Exists(path) Then
            Dim workbook As IWorkbook = SpreadsheetControl1.Document

            Using stream As New FileStream(path, FileMode.Open)
                workbook.LoadDocument(stream, DocumentFormat.OpenXml)

                For Each worksheet As Worksheet In workbook.Worksheets
                    worksheet.GetUsedRange().AutoFitRows()
                    worksheet.GetUsedRange().AutoFitColumns()
                Next
            End Using
        End If

        If SpreadsheetControl1.Document.Worksheets.Count > 0 Then
            SpreadsheetControl1.Document.Worksheets.ActiveWorksheet = SpreadsheetControl1.Document.Worksheets(0)
        End If
    End Sub
End Class