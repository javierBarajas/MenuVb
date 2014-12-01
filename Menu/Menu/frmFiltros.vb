Public Class frmFiltros
#Region "Variables"
    Private dt As DataTable = Nothing
    Private linq
    Private linq2 As DataTable
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

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case (ComboBox1.SelectedItem)
            Case "City"
                linq = (From row In dt.AsEnumerable()
                        Group row By City = row.Field(Of String)("City") Into CityGroup = Group
                        Select City).ToList()
                ComboBox2.DataSource = linq
            Case "St"
                linq = (From row In dt.AsEnumerable()
                        Group row By St = row.Field(Of String)("St.") Into StGroup = Group
                        Select St).ToList()
                ComboBox2.DataSource = linq
            Case "ZipCode"
                linq = (From row In dt.AsEnumerable()
                        Group row By ZipCode = row.Field(Of String)("ZipCode") Into ZipCodeGroup = Group
                        Select ZipCode).ToList()
                ComboBox2.DataSource = linq
                ComboBox2.SelectedIndex = 0
            Case "Country"
                linq = (From row In dt.AsEnumerable()
                        Group row By Country = row.Field(Of String)("Country") Into CountryGroup = Group
                        Select Country).ToList()
                ComboBox2.DataSource = linq
        End Select
        'linq = ""
    End Sub

    Private Sub frmFiltros_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.ComboBox1.Items.Add("City")
        Me.ComboBox1.Items.Add("St")
        Me.ComboBox1.Items.Add("ZipCode")
        Me.ComboBox1.Items.Add("Country")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (ComboBox1.SelectedIndex = -1) Then
            Dim frmWord As New frmWord(dt)
            frmWord.Show()
        End If

        If (ComboBox1.SelectedIndex >= 0 And ComboBox2.SelectedIndex >= 0) Then

            Dim dt2 As DataTable = New DataTable()
            With dt2.Columns
                .Add("Customer")
                .Add("Address1")
                .Add("Address2")
                .Add("City")
                .Add("St.")
                .Add("ZipCode")
                .Add("Country")
            End With

            Select Case (ComboBox1.SelectedItem)
                Case "City"

                    linq2 = (From row In dt.AsEnumerable()
                                     Where row.Field(Of String)("City") = ComboBox2.SelectedValue
                                     Select dt2.LoadDataRow(New [Object]() {row.Field(Of String)("Customer"), row.Field(Of String)("Address1"), row.Field(Of String)("Address2"), row.Field(Of String)("City"), row.Field(Of String)("St."), row.Field(Of String)("ZipCode"), row.Field(Of String)("Country")}, False)).CopyToDataTable()
                Case "St."

                    linq2 = (From row In dt.AsEnumerable()
                                     Where row.Field(Of String)("St.") = ComboBox2.SelectedValue
                                     Select dt2.LoadDataRow(New [Object]() {row.Field(Of String)("Customer"), row.Field(Of String)("Address1"), row.Field(Of String)("Address2"), row.Field(Of String)("City"), row.Field(Of String)("St."), row.Field(Of String)("ZipCode"), row.Field(Of String)("Country")}, False)).CopyToDataTable()
                Case "ZipCode"

                    linq2 = (From row In dt.AsEnumerable()
                                     Where row.Field(Of String)("ZipCode") = ComboBox2.SelectedValue
                                     Select dt2.LoadDataRow(New [Object]() {row.Field(Of String)("Customer"), row.Field(Of String)("Address1"), row.Field(Of String)("Address2"), row.Field(Of String)("City"), row.Field(Of String)("St."), row.Field(Of String)("ZipCode"), row.Field(Of String)("Country")}, False)).CopyToDataTable()
                Case "Country"

                    linq2 = (From row In dt.AsEnumerable()
                                     Where row.Field(Of String)("Country") = ComboBox2.SelectedValue
                                     Select dt2.LoadDataRow(New [Object]() {row.Field(Of String)("Customer"), row.Field(Of String)("Address1"), row.Field(Of String)("Address2"), row.Field(Of String)("City"), row.Field(Of String)("St."), row.Field(Of String)("ZipCode"), row.Field(Of String)("Country")}, False)).CopyToDataTable()
            End Select
            Dim frmWord As New frmWord(linq2)
            frmWord.Show()
        End If
    End Sub
End Class