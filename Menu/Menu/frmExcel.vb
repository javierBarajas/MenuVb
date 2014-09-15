Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Export
Imports System.IO

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
#End Region

    Public Function Worksheet(ByVal index As Integer) As Worksheet
        If index >= 0 AndAlso index < SpreadsheetControl1.Document.Worksheets.Count Then
            Return SpreadsheetControl1.Document.Worksheets(index)
        End If

        Return Nothing
    End Function

    Public Function ExportToDataTable() As DataTable
        Dim worksheet As Worksheet = SpreadsheetControl1.Document.Worksheets(0)
        Dim dataTable As DataTable = worksheet.CreateDataTable(worksheet.GetUsedRange(), True)
        Dim exporter As DataTableExporter = worksheet.CreateDataTableExporter(worksheet.GetUsedRange(), dataTable, True)
        exporter.Export()
    End Function

    Public Sub Create()
        If dt IsNot Nothing Then
            spreadsheetControl1.Document.Styles.Add("HeaderStyle")
            Dim worksheet As Worksheet = spreadsheetControl1.Document.Worksheets(0)


            'TableStyleCollection tableStyle = spreadsheetControl1.Document.TableStyles;

            worksheet.Name = dt.TableName

            'Table table = worksheet.Tables.Add(worksheet.Range.FromLTRB(FirstColIndex, FirstRowIndex, dt.Columns.Count - 1, dt.Rows.Count), true);
            '                table.Style = spreadsheetControl1.Document.TableStyles[BuiltInTableStyleId.TableStyleLight9];

            worksheet.Import(dt, _AddHeader, _FirstRowIndex, _FirstColIndex)
        ElseIf ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            Dim worksheetCollection As WorksheetCollection = spreadsheetControl1.Document.Worksheets
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
            Next
        ElseIf path IsNot Nothing AndAlso File.Exists(path) Then
            Dim workbook As IWorkbook = spreadsheetControl1.Document

            Using stream As New FileStream(path, FileMode.Open)
                workbook.LoadDocument(stream, DocumentFormat.OpenXml)

                For Each worksheet As Worksheet In workbook.Worksheets
                    worksheet.GetUsedRange().AutoFitRows()
                    worksheet.GetUsedRange().AutoFitColumns()
                Next
            End Using
        End If

        If spreadsheetControl1.Document.Worksheets.Count > 0 Then
            spreadsheetControl1.Document.Worksheets.ActiveWorksheet = spreadsheetControl1.Document.Worksheets(0)
        End If
    End Sub
End Class