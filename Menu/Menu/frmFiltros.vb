Public Class frmFiltros
#Region "Variables"
    Private dt As DataTable
    Private linq
    Public linq2 As DataTable
#End Region

#Region "Constructor"
    Private Sub Initialize(ByRef dt As DataTable)
        InitializeComponent()
        Me.dt = dt
    End Sub

    Public Sub New(ByRef dt As DataTable)
        Initialize(dt)
    End Sub
#End Region

    Private Sub frmFiltros_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       
        linq = (From row In dt.AsEnumerable()
                Group row By City = row.Field(Of String)("City") Into CityGroup = Group
                Select City).ToList()
        cbCity.DataSource = linq
        cbCity.SelectedIndex = -1

        linq = (From row In dt.AsEnumerable()
                Group row By St = row.Field(Of String)("St.") Into StGroup = Group
                Select St).ToList()
        cbSt.DataSource = linq
        cbSt.SelectedIndex = -1

        linq = (From row In dt.AsEnumerable()
                Group row By ZipCode = row.Field(Of String)("ZipCode") Into ZipCodeGroup = Group
                Select ZipCode).ToList()
        cbZipCode.DataSource = linq
        cbZipCode.SelectedIndex = -1

        linq = (From row In dt.AsEnumerable()
                Group row By Country = row.Field(Of String)("Country") Into CountryGroup = Group
                Select Country).ToList()
        cbCountry.DataSource = linq
        cbCountry.SelectedIndex = -1
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (cbCity.SelectedIndex = -1 And cbSt.SelectedIndex = -1 And cbZipCode.SelectedIndex = -1 And cbCountry.SelectedIndex = -1) Then
            linq2 = dt
        End If

        If (cbCity.SelectedIndex >= 0 Or cbSt.SelectedIndex >= 0 Or cbZipCode.SelectedIndex >= 0 Or cbCountry.SelectedIndex >= 0) Then
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

            Dim filtros(0 To 3, 0 To 1) As String

            Dim valor As Integer = 0
            If (cbCity.SelectedIndex >= 0) Then
                filtros(valor, 0) = "City"
                filtros(valor, 1) = cbCity.SelectedValue
                valor = valor + 1
            End If
            If (cbSt.SelectedIndex >= 0) Then
                filtros(valor, 0) = "St."
                filtros(1, 1) = cbSt.SelectedValue
                valor = valor + 1
            End If
            If (cbZipCode.SelectedIndex >= 0) Then
                filtros(valor, 0) = "ZipCode"
                filtros(2, 1) = cbZipCode.SelectedValue
                valor = valor + 1
            End If
            If (cbCountry.SelectedIndex >= 0) Then
                filtros(valor, 0) = "Country"
                filtros(0, 1) = cbCity.SelectedValue
                valor = valor + 1
            End If
            Try
                If (valor = 1) Then
                    linq2 = (From row In dt.AsEnumerable()
                                        Where row.Field(Of String)(filtros(0, 0)) = filtros(0, 1)
                                        Select dt2.LoadDataRow(New [Object]() {row.Field(Of String)("Customer"), row.Field(Of String)("Address1"), row.Field(Of String)("Address2"), row.Field(Of String)("City"), row.Field(Of String)("St."), row.Field(Of String)("ZipCode"), row.Field(Of String)("Country")}, False)).CopyToDataTable()
                End If
                If (valor = 2) Then
                    linq2 = (From row In dt.AsEnumerable()
                                        Where row.Field(Of String)(filtros(0, 0)) = filtros(0, 1) And row.Field(Of String)(filtros(1, 0)) = filtros(1, 1)
                                        Select dt2.LoadDataRow(New [Object]() {row.Field(Of String)("Customer"), row.Field(Of String)("Address1"), row.Field(Of String)("Address2"), row.Field(Of String)("City"), row.Field(Of String)("St."), row.Field(Of String)("ZipCode"), row.Field(Of String)("Country")}, False)).CopyToDataTable()
                End If
                If (valor = 3) Then
                    linq2 = (From row In dt.AsEnumerable()
                                        Where row.Field(Of String)(filtros(0, 0)) = filtros(0, 1) And row.Field(Of String)(filtros(1, 0)) = filtros(1, 1) And row.Field(Of String)(filtros(2, 0)) = filtros(2, 1)
                                        Select dt2.LoadDataRow(New [Object]() {row.Field(Of String)("Customer"), row.Field(Of String)("Address1"), row.Field(Of String)("Address2"), row.Field(Of String)("City"), row.Field(Of String)("St."), row.Field(Of String)("ZipCode"), row.Field(Of String)("Country")}, False)).CopyToDataTable()
                End If
                If (valor = 4) Then
                    linq2 = (From row In dt.AsEnumerable()
                                        Where row.Field(Of String)(filtros(0, 0)) = filtros(0, 1) And row.Field(Of String)(filtros(1, 0)) = filtros(1, 1) And row.Field(Of String)(filtros(2, 0)) = filtros(2, 1) And row.Field(Of String)(filtros(3, 0)) = filtros(3, 1)
                                        Select dt2.LoadDataRow(New [Object]() {row.Field(Of String)("Customer"), row.Field(Of String)("Address1"), row.Field(Of String)("Address2"), row.Field(Of String)("City"), row.Field(Of String)("St."), row.Field(Of String)("ZipCode"), row.Field(Of String)("Country")}, False)).CopyToDataTable()
                End If
            Catch
                Dim dt3 As DataTable = New DataTable()
                With dt3.Columns
                    .Add("Customer")
                    .Add("Address1")
                    .Add("Address2")
                    .Add("City")
                    .Add("St.")
                    .Add("ZipCode")
                    .Add("Country")
                End With
                linq2 = dt3
            End Try
            valor = 0
            Me.Close()
        End If
    End Sub
End Class