Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit
Imports System.IO
Public Class frmWord
#Region "Variables"
    Private dt As DataTable = Nothing
#End Region

#Region "Constructor"
    Private Sub Initialize(ByVal dt As DataTable)
        InitializeComponent()
        Me.dt = dt
    End Sub

    Public Sub New(ByVal dt As DataTable)
        Initialize(dt)

        RibbonControl1.SelectedPage = MailingsRibbonPage1
        DataNavigator1.DataSource = dt
        RichEditControl1.Options.MailMerge.DataSource = dt
        RichEditControl1.Options.MailMerge.ViewMergedData = True
        RichEditControl1.LoadDocument("Pantilla.rtf")

    End Sub
#End Region

    Private Sub frmWord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '    RichEditControl1.CreateNewDocument()
        '    RichEditControl1.Options.MailMerge.DataSource = dt
        '    DataNavigator1.DataSource = dt
    End Sub

    Private Sub BarButtonItem1_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem1.ItemClick

        Dim count As Integer = dt.Rows.Count

        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*"
        saveFileDialog1.FilterIndex = 1
        saveFileDialog1.RestoreDirectory = True

        If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Using docServer As New RichEditDocumentServer()
                Dim options As MailMergeOptions = RichEditControl1.CreateMailMergeOptions()

                For i As Integer = 0 To count - 1
                    Dim filename As String = String.Format("{0}{1}.pdf", saveFileDialog1.FileName, (i + 1).ToString())
                    options.LastRecordIndex = i
                    options.FirstRecordIndex = options.LastRecordIndex

                    Using fs As New FileStream(filename, FileMode.Create, System.IO.FileAccess.Write)
                        RichEditControl1.MailMerge(options, docServer.Document)
                        docServer.ExportToPdf(fs)
                    End Using
                Next i
            End Using
        End If
    End Sub

    Private Sub BarButtonItem2_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem2.ItemClick
        Dim count As Integer = dt.Rows.Count

        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*"
        saveFileDialog1.FilterIndex = 1
        saveFileDialog1.RestoreDirectory = True

        If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Using docServer As New RichEditDocumentServer()
                Dim options As MailMergeOptions = RichEditControl1.CreateMailMergeOptions()

                options.LastRecordIndex = count
                options.FirstRecordIndex = 1

                Using fs As New FileStream(saveFileDialog1.FileName, FileMode.Create, System.IO.FileAccess.Write)
                    RichEditControl1.MailMerge(options, docServer.Document)
                    docServer.ExportToPdf(fs)
                End Using
            End Using
        End If
    End Sub

    Private Sub BarButtonItem4_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem4.ItemClick
        Dim myMergeOptions As MailMergeOptions = RichEditControl1.Document.CreateMailMergeOptions()
        myMergeOptions.DataSource = dt
        myMergeOptions.FirstRecordIndex = 1
        myMergeOptions.LastRecordIndex = dt.Rows.Count
        myMergeOptions.MergeMode = MergeMode.NewSection

        Dim fileDialog As New SaveFileDialog()
        fileDialog.Filter = "MS Word 2007 documents (*.docx)|*.docx|All files (*.*)|*.*"
        fileDialog.FilterIndex = 1
        fileDialog.RestoreDirectory = True

        fileDialog.ShowDialog()
        Dim fName As String = fileDialog.FileName
        If fName <> "" Then
            RichEditControl1.Document.MailMerge(myMergeOptions, fileDialog.FileName, DocumentFormat.OpenXml)
        End If
        'System.Diagnostics.Process.Start(fileDialog.FileName)
    End Sub
End Class