﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ListadoIvaCompras
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.CustomLabel1 = New Administracion.CustomLabel()
        Me.CustomLabel2 = New Administracion.CustomLabel()
        Me.CustomLabel3 = New Administracion.CustomLabel()
        Me.txtDesdeFecha = New System.Windows.Forms.MaskedTextBox()
        Me.txthastafecha = New System.Windows.Forms.MaskedTextBox()
        Me.TipoListado = New Administracion.CustomComboBox()
        Me.txtDesde = New System.Windows.Forms.MaskedTextBox()
        Me.btnCancela = New System.Windows.Forms.Button()
        Me.btnPantalla = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'CustomLabel1
        '
        Me.CustomLabel1.AutoSize = True
        Me.CustomLabel1.ControlAssociationKey = -1
        Me.CustomLabel1.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold)
        Me.CustomLabel1.ForeColor = System.Drawing.SystemColors.Control
        Me.CustomLabel1.Location = New System.Drawing.Point(26, 37)
        Me.CustomLabel1.Name = "CustomLabel1"
        Me.CustomLabel1.Size = New System.Drawing.Size(86, 18)
        Me.CustomLabel1.TabIndex = 19
        Me.CustomLabel1.Text = "Desde Fecha"
        '
        'CustomLabel2
        '
        Me.CustomLabel2.AutoSize = True
        Me.CustomLabel2.ControlAssociationKey = -1
        Me.CustomLabel2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold)
        Me.CustomLabel2.ForeColor = System.Drawing.SystemColors.Control
        Me.CustomLabel2.Location = New System.Drawing.Point(254, 38)
        Me.CustomLabel2.Name = "CustomLabel2"
        Me.CustomLabel2.Size = New System.Drawing.Size(81, 18)
        Me.CustomLabel2.TabIndex = 20
        Me.CustomLabel2.Text = "Hasta Fecha"
        '
        'CustomLabel3
        '
        Me.CustomLabel3.AutoSize = True
        Me.CustomLabel3.ControlAssociationKey = -1
        Me.CustomLabel3.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold)
        Me.CustomLabel3.ForeColor = System.Drawing.SystemColors.Control
        Me.CustomLabel3.Location = New System.Drawing.Point(143, 80)
        Me.CustomLabel3.Name = "CustomLabel3"
        Me.CustomLabel3.Size = New System.Drawing.Size(82, 18)
        Me.CustomLabel3.TabIndex = 21
        Me.CustomLabel3.Text = "Tipo Listado"
        '
        'txtDesdeFecha
        '
        Me.txtDesdeFecha.Location = New System.Drawing.Point(124, 36)
        Me.txtDesdeFecha.Mask = "##/##/####"
        Me.txtDesdeFecha.Name = "txtDesdeFecha"
        Me.txtDesdeFecha.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtDesdeFecha.Size = New System.Drawing.Size(106, 20)
        Me.txtDesdeFecha.TabIndex = 0
        Me.txtDesdeFecha.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txthastafecha
        '
        Me.txthastafecha.Location = New System.Drawing.Point(347, 37)
        Me.txthastafecha.Mask = "##/##/####"
        Me.txthastafecha.Name = "txthastafecha"
        Me.txthastafecha.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txthastafecha.Size = New System.Drawing.Size(106, 20)
        Me.txthastafecha.TabIndex = 1
        Me.txthastafecha.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TipoListado
        '
        Me.TipoListado.Cleanable = False
        Me.TipoListado.Empty = False
        Me.TipoListado.EnterIndex = -1
        Me.TipoListado.FormattingEnabled = True
        Me.TipoListado.LabelAssociationKey = -1
        Me.TipoListado.Location = New System.Drawing.Point(237, 79)
        Me.TipoListado.Name = "TipoListado"
        Me.TipoListado.Size = New System.Drawing.Size(107, 21)
        Me.TipoListado.TabIndex = 2
        Me.TipoListado.Validator = Administracion.ValidatorType.None
        '
        'txtDesde
        '
        Me.txtDesde.Location = New System.Drawing.Point(214, 51)
        Me.txtDesde.Mask = "##/##/####"
        Me.txtDesde.Name = "txtDesde"
        Me.txtDesde.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtDesde.Size = New System.Drawing.Size(106, 20)
        Me.txtDesde.TabIndex = 22
        '
        'btnCancela
        '
        Me.btnCancela.BackgroundImage = Global.Administracion.My.Resources.Resources.Salir2
        Me.btnCancela.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnCancela.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCancela.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control
        Me.btnCancela.FlatAppearance.BorderSize = 0
        Me.btnCancela.FlatAppearance.CheckedBackColor = System.Drawing.SystemColors.Control
        Me.btnCancela.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.Control
        Me.btnCancela.FlatAppearance.MouseOverBackColor = System.Drawing.SystemColors.Control
        Me.btnCancela.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCancela.Location = New System.Drawing.Point(313, 190)
        Me.btnCancela.Name = "btnCancela"
        Me.btnCancela.Size = New System.Drawing.Size(63, 34)
        Me.btnCancela.TabIndex = 37
        Me.ToolTip1.SetToolTip(Me.btnCancela, "Cerrar Formulario")
        Me.btnCancela.UseVisualStyleBackColor = True
        '
        'btnPantalla
        '
        Me.btnPantalla.BackgroundImage = Global.Administracion.My.Resources.Resources.Screen_preview
        Me.btnPantalla.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnPantalla.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnPantalla.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control
        Me.btnPantalla.FlatAppearance.BorderSize = 0
        Me.btnPantalla.FlatAppearance.CheckedBackColor = System.Drawing.SystemColors.Control
        Me.btnPantalla.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.Control
        Me.btnPantalla.FlatAppearance.MouseOverBackColor = System.Drawing.SystemColors.Control
        Me.btnPantalla.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnPantalla.Location = New System.Drawing.Point(123, 190)
        Me.btnPantalla.Name = "btnPantalla"
        Me.btnPantalla.Size = New System.Drawing.Size(63, 34)
        Me.btnPantalla.TabIndex = 36
        Me.ToolTip1.SetToolTip(Me.btnPantalla, "Reporte por Pantalla")
        Me.btnPantalla.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(74, Byte), Integer), CType(CType(95, Byte), Integer))
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(498, 50)
        Me.Panel1.TabIndex = 38
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Control
        Me.Label2.Location = New System.Drawing.Point(318, 10)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(156, 26)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "SURFACTAN S.A."
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Control
        Me.Label1.Location = New System.Drawing.Point(27, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(166, 19)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Listado de Iva Compras"
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(139, Byte), Integer), CType(CType(82, Byte), Integer))
        Me.Panel2.Controls.Add(Me.CustomLabel1)
        Me.Panel2.Controls.Add(Me.CustomLabel2)
        Me.Panel2.Controls.Add(Me.TipoListado)
        Me.Panel2.Controls.Add(Me.CustomLabel3)
        Me.Panel2.Controls.Add(Me.txthastafecha)
        Me.Panel2.Controls.Add(Me.txtDesdeFecha)
        Me.Panel2.Location = New System.Drawing.Point(0, 49)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(498, 126)
        Me.Panel2.TabIndex = 40
        '
        'btnImprimir
        '
        Me.btnImprimir.BackgroundImage = Global.Administracion.My.Resources.Resources.imprimir
        Me.btnImprimir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnImprimir.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.FlatAppearance.BorderSize = 0
        Me.btnImprimir.FlatAppearance.CheckedBackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.FlatAppearance.MouseOverBackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnImprimir.Location = New System.Drawing.Point(218, 190)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.Size = New System.Drawing.Size(63, 34)
        Me.btnImprimir.TabIndex = 37
        Me.ToolTip1.SetToolTip(Me.btnImprimir, "Imprimir Reporte")
        Me.btnImprimir.UseVisualStyleBackColor = True
        '
        'ListadoIvaCompras
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(498, 232)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnCancela)
        Me.Controls.Add(Me.btnPantalla)
        Me.Name = "ListadoIvaCompras"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CustomLabel1 As Administracion.CustomLabel
    Friend WithEvents CustomLabel2 As Administracion.CustomLabel
    Friend WithEvents CustomLabel3 As Administracion.CustomLabel
    Friend WithEvents txtDesdeFecha As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txthastafecha As System.Windows.Forms.MaskedTextBox
    Friend WithEvents TipoListado As Administracion.CustomComboBox
    Friend WithEvents txtDesde As System.Windows.Forms.MaskedTextBox
    Friend WithEvents btnCancela As System.Windows.Forms.Button
    Friend WithEvents btnPantalla As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnImprimir As System.Windows.Forms.Button
End Class
