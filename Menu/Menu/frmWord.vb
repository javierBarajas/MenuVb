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
    End Sub
#End Region

    Private Sub frmWord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RichEditControl1.CreateNewDocument()
        RichEditControl1.Options.MailMerge.DataSource = dt
        DataNavigator1.DataSource = dt
    End Sub
End Class