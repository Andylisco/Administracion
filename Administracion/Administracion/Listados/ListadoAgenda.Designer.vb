﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ListadoAgenda
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
        Me.btnCancela = New System.Windows.Forms.Button()
        Me.btnAcepta = New System.Windows.Forms.Button()
        Me.TipoListado = New Administracion.CustomComboBox()
        Me.txthastafecha = New System.Windows.Forms.MaskedTextBox()
        Me.txtDesdeFecha = New System.Windows.Forms.MaskedTextBox()
        Me.CustomLabel3 = New Administracion.CustomLabel()
        Me.CustomLabel2 = New Administracion.CustomLabel()
        Me.CustomLabel1 = New Administracion.CustomLabel()
        Me.opcImpesora = New System.Windows.Forms.RadioButton()
        Me.opcPantalla = New System.Windows.Forms.RadioButton()
        Me.TipoLetras = New Administracion.CustomComboBox()
        Me.CustomLabel4 = New Administracion.CustomLabel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
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
        Me.btnCancela.Location = New System.Drawing.Point(305, 192)
        Me.btnCancela.Name = "btnCancela"
        Me.btnCancela.Size = New System.Drawing.Size(125, 42)
        Me.btnCancela.TabIndex = 47
        Me.ToolTip1.SetToolTip(Me.btnCancela, "Cerrar")
        Me.btnCancela.UseVisualStyleBackColor = True
        '
        'btnAcepta
        '
        Me.btnAcepta.BackgroundImage = Global.Administracion.My.Resources.Resources.Aceptar_N2
        Me.btnAcepta.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnAcepta.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnAcepta.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control
        Me.btnAcepta.FlatAppearance.BorderSize = 0
        Me.btnAcepta.FlatAppearance.CheckedBackColor = System.Drawing.SystemColors.Control
        Me.btnAcepta.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.Control
        Me.btnAcepta.FlatAppearance.MouseOverBackColor = System.Drawing.SystemColors.Control
        Me.btnAcepta.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAcepta.Location = New System.Drawing.Point(149, 192)
        Me.btnAcepta.Name = "btnAcepta"
        Me.btnAcepta.Size = New System.Drawing.Size(125, 42)
        Me.btnAcepta.TabIndex = 46
        Me.ToolTip1.SetToolTip(Me.btnAcepta, "Aceptar")
        Me.btnAcepta.UseVisualStyleBackColor = True
        '
        'TipoListado
        '
        Me.TipoListado.Cleanable = False
        Me.TipoListado.Empty = False
        Me.TipoListado.EnterIndex = -1
        Me.TipoListado.FormattingEnabled = True
        Me.TipoListado.LabelAssociationKey = -1
        Me.TipoListado.Location = New System.Drawing.Point(134, 58)
        Me.TipoListado.Name = "TipoListado"
        Me.TipoListado.Size = New System.Drawing.Size(146, 21)
        Me.TipoListado.TabIndex = 40
        Me.TipoListado.Validator = Administracion.ValidatorType.None
        '
        'txthastafecha
        '
        Me.txthastafecha.Location = New System.Drawing.Point(388, 17)
        Me.txthastafecha.Mask = "##/##/####"
        Me.txthastafecha.Name = "txthastafecha"
        Me.txthastafecha.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txthastafecha.Size = New System.Drawing.Size(146, 20)
        Me.txthastafecha.TabIndex = 39
        '
        'txtDesdeFecha
        '
        Me.txtDesdeFecha.Location = New System.Drawing.Point(138, 17)
        Me.txtDesdeFecha.Mask = "##/##/####"
        Me.txtDesdeFecha.Name = "txtDesdeFecha"
        Me.txtDesdeFecha.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtDesdeFecha.Size = New System.Drawing.Size(142, 20)
        Me.txtDesdeFecha.TabIndex = 38
        '
        'CustomLabel3
        '
        Me.CustomLabel3.AutoSize = True
        Me.CustomLabel3.ControlAssociationKey = -1
        Me.CustomLabel3.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold)
        Me.CustomLabel3.ForeColor = System.Drawing.SystemColors.Control
        Me.CustomLabel3.Location = New System.Drawing.Point(45, 59)
        Me.CustomLabel3.Name = "CustomLabel3"
        Me.CustomLabel3.Size = New System.Drawing.Size(82, 18)
        Me.CustomLabel3.TabIndex = 45
        Me.CustomLabel3.Text = "Tipo Listado"
        '
        'CustomLabel2
        '
        Me.CustomLabel2.AutoSize = True
        Me.CustomLabel2.ControlAssociationKey = -1
        Me.CustomLabel2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold)
        Me.CustomLabel2.ForeColor = System.Drawing.SystemColors.Control
        Me.CustomLabel2.Location = New System.Drawing.Point(292, 18)
        Me.CustomLabel2.Name = "CustomLabel2"
        Me.CustomLabel2.Size = New System.Drawing.Size(81, 18)
        Me.CustomLabel2.TabIndex = 44
        Me.CustomLabel2.Text = "Hasta Fecha"
        '
        'CustomLabel1
        '
        Me.CustomLabel1.AutoSize = True
        Me.CustomLabel1.ControlAssociationKey = -1
        Me.CustomLabel1.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold)
        Me.CustomLabel1.ForeColor = System.Drawing.SystemColors.Control
        Me.CustomLabel1.Location = New System.Drawing.Point(41, 18)
        Me.CustomLabel1.Name = "CustomLabel1"
        Me.CustomLabel1.Size = New System.Drawing.Size(86, 18)
        Me.CustomLabel1.TabIndex = 43
        Me.CustomLabel1.Text = "Desde Fecha"
        '
        'opcImpesora
        '
        Me.opcImpesora.AutoSize = True
        Me.opcImpesora.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold)
        Me.opcImpesora.ForeColor = System.Drawing.SystemColors.Control
        Me.opcImpesora.Location = New System.Drawing.Point(323, 99)
        Me.opcImpesora.Name = "opcImpesora"
        Me.opcImpesora.Size = New System.Drawing.Size(89, 22)
        Me.opcImpesora.TabIndex = 42
        Me.opcImpesora.TabStop = True
        Me.opcImpesora.Text = "Impresora"
        Me.opcImpesora.UseVisualStyleBackColor = True
        '
        'opcPantalla
        '
        Me.opcPantalla.AutoSize = True
        Me.opcPantalla.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold)
        Me.opcPantalla.ForeColor = System.Drawing.SystemColors.Control
        Me.opcPantalla.Location = New System.Drawing.Point(189, 99)
        Me.opcPantalla.Name = "opcPantalla"
        Me.opcPantalla.Size = New System.Drawing.Size(76, 22)
        Me.opcPantalla.TabIndex = 41
        Me.opcPantalla.TabStop = True
        Me.opcPantalla.Text = "Pantalla"
        Me.opcPantalla.UseVisualStyleBackColor = True
        '
        'TipoLetras
        '
        Me.TipoLetras.Cleanable = False
        Me.TipoLetras.Empty = False
        Me.TipoLetras.EnterIndex = -1
        Me.TipoLetras.FormattingEnabled = True
        Me.TipoLetras.LabelAssociationKey = -1
        Me.TipoLetras.Location = New System.Drawing.Point(388, 58)
        Me.TipoLetras.Name = "TipoLetras"
        Me.TipoLetras.Size = New System.Drawing.Size(146, 21)
        Me.TipoLetras.TabIndex = 48
        Me.TipoLetras.Validator = Administracion.ValidatorType.None
        '
        'CustomLabel4
        '
        Me.CustomLabel4.AutoSize = True
        Me.CustomLabel4.ControlAssociationKey = -1
        Me.CustomLabel4.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold)
        Me.CustomLabel4.ForeColor = System.Drawing.SystemColors.Control
        Me.CustomLabel4.Location = New System.Drawing.Point(298, 59)
        Me.CustomLabel4.Name = "CustomLabel4"
        Me.CustomLabel4.Size = New System.Drawing.Size(75, 18)
        Me.CustomLabel4.TabIndex = 49
        Me.CustomLabel4.Text = "Tipo Letras"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(74, Byte), Integer), CType(CType(95, Byte), Integer))
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(-1, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(582, 50)
        Me.Panel1.TabIndex = 50
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Control
        Me.Label2.Location = New System.Drawing.Point(400, 10)
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
        Me.Label1.Size = New System.Drawing.Size(251, 19)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Listado de Agenda de Vencimientos"
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(139, Byte), Integer), CType(CType(82, Byte), Integer))
        Me.Panel2.Controls.Add(Me.CustomLabel1)
        Me.Panel2.Controls.Add(Me.opcPantalla)
        Me.Panel2.Controls.Add(Me.TipoLetras)
        Me.Panel2.Controls.Add(Me.opcImpesora)
        Me.Panel2.Controls.Add(Me.CustomLabel4)
        Me.Panel2.Controls.Add(Me.CustomLabel2)
        Me.Panel2.Controls.Add(Me.CustomLabel3)
        Me.Panel2.Controls.Add(Me.txtDesdeFecha)
        Me.Panel2.Controls.Add(Me.TipoListado)
        Me.Panel2.Controls.Add(Me.txthastafecha)
        Me.Panel2.Location = New System.Drawing.Point(-1, 50)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(582, 136)
        Me.Panel2.TabIndex = 51
        '
        'ListadoAgenda
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(579, 240)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.btnCancela)
        Me.Controls.Add(Me.btnAcepta)
        Me.Name = "ListadoAgenda"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCancela As System.Windows.Forms.Button
    Friend WithEvents btnAcepta As System.Windows.Forms.Button
    Friend WithEvents TipoListado As Administracion.CustomComboBox
    Friend WithEvents txthastafecha As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtDesdeFecha As System.Windows.Forms.MaskedTextBox
    Friend WithEvents CustomLabel3 As Administracion.CustomLabel
    Friend WithEvents CustomLabel2 As Administracion.CustomLabel
    Friend WithEvents CustomLabel1 As Administracion.CustomLabel
    Friend WithEvents opcImpesora As System.Windows.Forms.RadioButton
    Friend WithEvents opcPantalla As System.Windows.Forms.RadioButton
    Friend WithEvents TipoLetras As Administracion.CustomComboBox
    Friend WithEvents CustomLabel4 As Administracion.CustomLabel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
End Class
