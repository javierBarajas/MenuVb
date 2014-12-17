<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFiltros
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.cbCity = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cbSt = New System.Windows.Forms.ComboBox()
        Me.gcFiltros = New System.Windows.Forms.GroupBox()
        Me.gcCountry = New System.Windows.Forms.GroupBox()
        Me.cbCountry = New System.Windows.Forms.ComboBox()
        Me.gcZipCode = New System.Windows.Forms.GroupBox()
        Me.cbZipCode = New System.Windows.Forms.ComboBox()
        Me.gcSt = New System.Windows.Forms.GroupBox()
        Me.gcCity = New System.Windows.Forms.GroupBox()
        Me.cgSelected = New System.Windows.Forms.GroupBox()
        Me.cbSelected = New System.Windows.Forms.ComboBox()
        Me.gcFiltros.SuspendLayout()
        Me.gcCountry.SuspendLayout()
        Me.gcZipCode.SuspendLayout()
        Me.gcSt.SuspendLayout()
        Me.gcCity.SuspendLayout()
        Me.cgSelected.SuspendLayout()
        Me.SuspendLayout()
        '
        'cbCity
        '
        Me.cbCity.FormattingEnabled = True
        Me.cbCity.Location = New System.Drawing.Point(9, 19)
        Me.cbCity.Name = "cbCity"
        Me.cbCity.Size = New System.Drawing.Size(168, 21)
        Me.cbCity.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(153, 357)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(79, 22)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Filtrar"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'cbSt
        '
        Me.cbSt.FormattingEnabled = True
        Me.cbSt.Location = New System.Drawing.Point(9, 19)
        Me.cbSt.Name = "cbSt"
        Me.cbSt.Size = New System.Drawing.Size(168, 21)
        Me.cbSt.TabIndex = 2
        '
        'gcFiltros
        '
        Me.gcFiltros.Controls.Add(Me.gcCountry)
        Me.gcFiltros.Controls.Add(Me.gcZipCode)
        Me.gcFiltros.Controls.Add(Me.gcSt)
        Me.gcFiltros.Controls.Add(Me.gcCity)
        Me.gcFiltros.Location = New System.Drawing.Point(12, 12)
        Me.gcFiltros.Name = "gcFiltros"
        Me.gcFiltros.Size = New System.Drawing.Size(219, 339)
        Me.gcFiltros.TabIndex = 3
        Me.gcFiltros.TabStop = False
        Me.gcFiltros.Text = "Filtrar por:"
        '
        'gcCountry
        '
        Me.gcCountry.Controls.Add(Me.cbCountry)
        Me.gcCountry.Location = New System.Drawing.Point(15, 199)
        Me.gcCountry.Name = "gcCountry"
        Me.gcCountry.Size = New System.Drawing.Size(187, 53)
        Me.gcCountry.TabIndex = 6
        Me.gcCountry.TabStop = False
        Me.gcCountry.Text = "Country"
        '
        'cbCountry
        '
        Me.cbCountry.FormattingEnabled = True
        Me.cbCountry.Location = New System.Drawing.Point(9, 19)
        Me.cbCountry.Name = "cbCountry"
        Me.cbCountry.Size = New System.Drawing.Size(168, 21)
        Me.cbCountry.TabIndex = 0
        '
        'gcZipCode
        '
        Me.gcZipCode.Controls.Add(Me.cbZipCode)
        Me.gcZipCode.Location = New System.Drawing.Point(15, 140)
        Me.gcZipCode.Name = "gcZipCode"
        Me.gcZipCode.Size = New System.Drawing.Size(187, 53)
        Me.gcZipCode.TabIndex = 5
        Me.gcZipCode.TabStop = False
        Me.gcZipCode.Text = "ZipCode"
        '
        'cbZipCode
        '
        Me.cbZipCode.FormattingEnabled = True
        Me.cbZipCode.Location = New System.Drawing.Point(9, 19)
        Me.cbZipCode.Name = "cbZipCode"
        Me.cbZipCode.Size = New System.Drawing.Size(168, 21)
        Me.cbZipCode.TabIndex = 0
        '
        'gcSt
        '
        Me.gcSt.Controls.Add(Me.cbSt)
        Me.gcSt.Location = New System.Drawing.Point(15, 79)
        Me.gcSt.Name = "gcSt"
        Me.gcSt.Size = New System.Drawing.Size(187, 55)
        Me.gcSt.TabIndex = 4
        Me.gcSt.TabStop = False
        Me.gcSt.Text = "St."
        '
        'gcCity
        '
        Me.gcCity.Controls.Add(Me.cbCity)
        Me.gcCity.Location = New System.Drawing.Point(15, 19)
        Me.gcCity.Name = "gcCity"
        Me.gcCity.Size = New System.Drawing.Size(187, 54)
        Me.gcCity.TabIndex = 3
        Me.gcCity.TabStop = False
        Me.gcCity.Text = "City"
        '
        'cgSelected
        '
        Me.cgSelected.Controls.Add(Me.cbSelected)
        Me.cgSelected.Location = New System.Drawing.Point(27, 283)
        Me.cgSelected.Name = "cgSelected"
        Me.cgSelected.Size = New System.Drawing.Size(187, 62)
        Me.cgSelected.TabIndex = 4
        Me.cgSelected.TabStop = False
        Me.cgSelected.Text = "Selected"
        '
        'cbSelected
        '
        Me.cbSelected.CausesValidation = False
        Me.cbSelected.FormattingEnabled = True
        Me.cbSelected.Location = New System.Drawing.Point(10, 20)
        Me.cbSelected.Name = "cbSelected"
        Me.cbSelected.Size = New System.Drawing.Size(166, 21)
        Me.cbSelected.TabIndex = 0
        '
        'frmFiltros
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(244, 391)
        Me.Controls.Add(Me.cgSelected)
        Me.Controls.Add(Me.gcFiltros)
        Me.Controls.Add(Me.Button1)
        Me.Name = "frmFiltros"
        Me.Text = "frmFiltros"
        Me.gcFiltros.ResumeLayout(False)
        Me.gcCountry.ResumeLayout(False)
        Me.gcZipCode.ResumeLayout(False)
        Me.gcSt.ResumeLayout(False)
        Me.gcCity.ResumeLayout(False)
        Me.cgSelected.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cbCity As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cbSt As System.Windows.Forms.ComboBox
    Friend WithEvents gcFiltros As System.Windows.Forms.GroupBox
    Friend WithEvents gcCountry As System.Windows.Forms.GroupBox
    Friend WithEvents cbCountry As System.Windows.Forms.ComboBox
    Friend WithEvents gcZipCode As System.Windows.Forms.GroupBox
    Friend WithEvents cbZipCode As System.Windows.Forms.ComboBox
    Friend WithEvents gcSt As System.Windows.Forms.GroupBox
    Friend WithEvents gcCity As System.Windows.Forms.GroupBox
    Friend WithEvents cgSelected As System.Windows.Forms.GroupBox
    Friend WithEvents cbSelected As System.Windows.Forms.ComboBox
End Class
