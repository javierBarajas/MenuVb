Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit
Imports System.IO
Public Class frmWord
#Region "Variables"
    Private dt As DataTable = Nothing
#End Region

    Private Sub frmWord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RichEditControl1.CreateNewDocument()
    End Sub

    Private Sub BarButtonItem1_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem1.ItemClick

        Dim document As Document = RichEditControl1.Document
        Dim fields As FieldCollection = document.Fields
        If (fields.Count = 0 Or Not (InsertMergeFieldItem1.Enabled)) Then

            Dim saveFileDialog1 As New SaveFileDialog()
            saveFileDialog1.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*"
            If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                RichEditControl1.ExportToPdf(saveFileDialog1.FileName)
            End If

        Else

            Dim frmFiltros As New frmFiltros(dt)
            frmFiltros.ShowDialog(Me)
            If (frmFiltros.linq2 IsNot Nothing) Then

                DataNavigator1.DataSource = frmFiltros.linq2
                RichEditControl1.Options.MailMerge.DataSource = frmFiltros.linq2
                RichEditControl1.Options.MailMerge.ViewMergedData = True

                Dim count As Integer = frmFiltros.linq2.Rows.Count

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
                            frmFiltros.linq2.Rows(i)("FileName") = IO.Path.GetDirectoryName(saveFileDialog1.FileName.ToString()) + "\" + filename
                        Next i
                    End Using
                End If

            End If

            DataNavigator1.DataSource = dt
            RichEditControl1.Options.MailMerge.DataSource = dt
            RichEditControl1.Options.MailMerge.ViewMergedData = True

        End If
    End Sub
    Private Sub BarButtonItem4_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem4.ItemClick

        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "MS Word 2007 documents (*.docx)|*.docx|All files (*.*)|*.*"

        Dim document As Document = RichEditControl1.Document
        Dim fields As FieldCollection = document.Fields
        If (fields.Count = 0 Or Not (InsertMergeFieldItem1.Enabled)) Then

            If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                RichEditControl1.SaveDocument(saveFileDialog1.FileName, DocumentFormat.OpenXml)
            End If

        Else

            Dim myMergeOptions As MailMergeOptions = RichEditControl1.Document.CreateMailMergeOptions()
            myMergeOptions.DataSource = dt
            myMergeOptions.FirstRecordIndex = 0
            myMergeOptions.LastRecordIndex = dt.Rows.Count
            myMergeOptions.MergeMode = MergeMode.NewSection

            saveFileDialog1.FilterIndex = 1
            saveFileDialog1.RestoreDirectory = True

            saveFileDialog1.ShowDialog()
            Dim fName As String = saveFileDialog1.FileName
            If fName <> "" Then
                RichEditControl1.Document.MailMerge(myMergeOptions, saveFileDialog1.FileName, DocumentFormat.OpenXml)
            End If

        End If
    End Sub

    Private Sub BarButtonItem2_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem2.ItemClick
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim frmExcel As New frmExcel(OpenFileDialog1.FileName)
            frmExcel.Create()
            dt = frmExcel.ExportToDataTable(0)
            DataNavigator1.DataSource = dt
            RichEditControl1.Options.MailMerge.DataSource = dt
            RichEditControl1.Options.MailMerge.ViewMergedData = True
        End If
    End Sub
End Class