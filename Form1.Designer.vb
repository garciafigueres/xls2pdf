<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPrincipal
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
        Me.btnCargar = New System.Windows.Forms.Button()
        Me.ofd = New System.Windows.Forms.OpenFileDialog()
        Me.lblEstado = New System.Windows.Forms.Label()
        Me.btnLeer = New System.Windows.Forms.Button()
        Me.lstFacturas = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.sfd = New System.Windows.Forms.SaveFileDialog()
        Me.SuspendLayout()
        '
        'btnCargar
        '
        Me.btnCargar.Location = New System.Drawing.Point(0, 0)
        Me.btnCargar.Name = "btnCargar"
        Me.btnCargar.Size = New System.Drawing.Size(75, 23)
        Me.btnCargar.TabIndex = 0
        Me.btnCargar.Text = "Cargar XLSX"
        Me.btnCargar.UseVisualStyleBackColor = True
        '
        'ofd
        '
        Me.ofd.FileName = "OpenFileDialog1"
        '
        'lblEstado
        '
        Me.lblEstado.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lblEstado.Location = New System.Drawing.Point(0, 147)
        Me.lblEstado.Name = "lblEstado"
        Me.lblEstado.Size = New System.Drawing.Size(324, 42)
        Me.lblEstado.TabIndex = 1
        '
        'btnLeer
        '
        Me.btnLeer.Location = New System.Drawing.Point(81, 0)
        Me.btnLeer.Name = "btnLeer"
        Me.btnLeer.Size = New System.Drawing.Size(75, 23)
        Me.btnLeer.TabIndex = 2
        Me.btnLeer.Text = "Leer XLSX"
        Me.btnLeer.UseVisualStyleBackColor = True
        '
        'lstFacturas
        '
        Me.lstFacturas.FormattingEnabled = True
        Me.lstFacturas.Location = New System.Drawing.Point(0, 49)
        Me.lstFacturas.Name = "lstFacturas"
        Me.lstFacturas.Size = New System.Drawing.Size(323, 95)
        Me.lstFacturas.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(0, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(240, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Resumen facturas (doble click para generar PDF)"
        '
        'btnSalir
        '
        Me.btnSalir.Location = New System.Drawing.Point(248, 0)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(75, 23)
        Me.btnSalir.TabIndex = 5
        Me.btnSalir.Text = "Salir"
        Me.btnSalir.UseVisualStyleBackColor = True
        '
        'frmPrincipal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(324, 189)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnSalir)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lstFacturas)
        Me.Controls.Add(Me.btnLeer)
        Me.Controls.Add(Me.lblEstado)
        Me.Controls.Add(Me.btnCargar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPrincipal"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "xls2pdf"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCargar As System.Windows.Forms.Button
    Friend WithEvents ofd As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lblEstado As System.Windows.Forms.Label
    Friend WithEvents btnLeer As System.Windows.Forms.Button
    Friend WithEvents lstFacturas As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnSalir As System.Windows.Forms.Button
    Friend WithEvents sfd As System.Windows.Forms.SaveFileDialog

End Class
