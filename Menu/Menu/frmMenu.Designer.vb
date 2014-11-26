<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMenu
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
        Me.btnExcel = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.gcExcel = New DevExpress.XtraEditors.GroupControl()
        Me.radioGroup1 = New DevExpress.XtraEditors.RadioGroup()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        CType(Me.gcExcel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gcExcel.SuspendLayout()
        CType(Me.radioGroup1.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnExcel
        '
        Me.btnExcel.Enabled = False
        Me.btnExcel.Location = New System.Drawing.Point(120, 43)
        Me.btnExcel.Name = "btnExcel"
        Me.btnExcel.Size = New System.Drawing.Size(75, 23)
        Me.btnExcel.TabIndex = 0
        Me.btnExcel.Text = "Run"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(12, 125)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Word"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'gcExcel
        '
        Me.gcExcel.Controls.Add(Me.radioGroup1)
        Me.gcExcel.Controls.Add(Me.btnExcel)
        Me.gcExcel.Location = New System.Drawing.Point(12, 12)
        Me.gcExcel.Name = "gcExcel"
        Me.gcExcel.Size = New System.Drawing.Size(200, 100)
        Me.gcExcel.TabIndex = 5
        Me.gcExcel.Text = "Excel"
        '
        'radioGroup1
        '
        Me.radioGroup1.Location = New System.Drawing.Point(6, 25)
        Me.radioGroup1.Name = "radioGroup1"
        Me.radioGroup1.Properties.Items.AddRange(New DevExpress.XtraEditors.Controls.RadioGroupItem() {New DevExpress.XtraEditors.Controls.RadioGroupItem(0, "DataTable"), New DevExpress.XtraEditors.Controls.RadioGroupItem(1, "DataSet"), New DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Open XLS")})
        Me.radioGroup1.Size = New System.Drawing.Size(96, 55)
        Me.radioGroup1.TabIndex = 1
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'frmMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(230, 160)
        Me.Controls.Add(Me.gcExcel)
        Me.Controls.Add(Me.Button2)
        Me.Name = "frmMenu"
        Me.Text = "Menu"
        CType(Me.gcExcel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gcExcel.ResumeLayout(False)
        CType(Me.radioGroup1.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnExcel As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents gcExcel As DevExpress.XtraEditors.GroupControl
    Friend WithEvents radioGroup1 As DevExpress.XtraEditors.RadioGroup
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog

End Class
