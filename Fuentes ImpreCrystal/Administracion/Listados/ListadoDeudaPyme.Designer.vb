﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ListadoDeudaPyme
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
        Me.btnCancela = New System.Windows.Forms.Button()
        Me.btnAcepta = New System.Windows.Forms.Button()
        Me.txthastafecha = New System.Windows.Forms.MaskedTextBox()
        Me.txtDesdeFecha = New System.Windows.Forms.MaskedTextBox()
        Me.CustomLabel2 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel1 = New WindowsApplication1.CustomLabel()
        Me.opcImpesora = New System.Windows.Forms.RadioButton()
        Me.opcPantalla = New System.Windows.Forms.RadioButton()
        Me.SuspendLayout()
        '
        'btnCancela
        '
        Me.btnCancela.Location = New System.Drawing.Point(222, 154)
        Me.btnCancela.Name = "btnCancela"
        Me.btnCancela.Size = New System.Drawing.Size(125, 34)
        Me.btnCancela.TabIndex = 45
        Me.btnCancela.Text = "Cancela"
        Me.btnCancela.UseVisualStyleBackColor = True
        '
        'btnAcepta
        '
        Me.btnAcepta.Location = New System.Drawing.Point(66, 154)
        Me.btnAcepta.Name = "btnAcepta"
        Me.btnAcepta.Size = New System.Drawing.Size(125, 34)
        Me.btnAcepta.TabIndex = 44
        Me.btnAcepta.Text = "Acepta"
        Me.btnAcepta.UseVisualStyleBackColor = True
        '
        'txthastafecha
        '
        Me.txthastafecha.Location = New System.Drawing.Point(200, 61)
        Me.txthastafecha.Mask = "##/##/####"
        Me.txthastafecha.Name = "txthastafecha"
        Me.txthastafecha.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txthastafecha.Size = New System.Drawing.Size(106, 20)
        Me.txthastafecha.TabIndex = 39
        '
        'txtDesdeFecha
        '
        Me.txtDesdeFecha.Location = New System.Drawing.Point(200, 25)
        Me.txtDesdeFecha.Mask = "##/##/####"
        Me.txtDesdeFecha.Name = "txtDesdeFecha"
        Me.txtDesdeFecha.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtDesdeFecha.Size = New System.Drawing.Size(106, 20)
        Me.txtDesdeFecha.TabIndex = 38
        '
        'CustomLabel2
        '
        Me.CustomLabel2.AutoSize = True
        Me.CustomLabel2.ControlAssociationKey = -1
        Me.CustomLabel2.Location = New System.Drawing.Point(78, 64)
        Me.CustomLabel2.Name = "CustomLabel2"
        Me.CustomLabel2.Size = New System.Drawing.Size(68, 13)
        Me.CustomLabel2.TabIndex = 43
        Me.CustomLabel2.Text = "Hasta Fecha"
        '
        'CustomLabel1
        '
        Me.CustomLabel1.AutoSize = True
        Me.CustomLabel1.ControlAssociationKey = -1
        Me.CustomLabel1.Location = New System.Drawing.Point(78, 28)
        Me.CustomLabel1.Name = "CustomLabel1"
        Me.CustomLabel1.Size = New System.Drawing.Size(71, 13)
        Me.CustomLabel1.TabIndex = 42
        Me.CustomLabel1.Text = "Desde Fecha"
        '
        'opcImpesora
        '
        Me.opcImpesora.AutoSize = True
        Me.opcImpesora.Location = New System.Drawing.Point(235, 110)
        Me.opcImpesora.Name = "opcImpesora"
        Me.opcImpesora.Size = New System.Drawing.Size(71, 17)
        Me.opcImpesora.TabIndex = 41
        Me.opcImpesora.TabStop = True
        Me.opcImpesora.Text = "Impresora"
        Me.opcImpesora.UseVisualStyleBackColor = True
        '
        'opcPantalla
        '
        Me.opcPantalla.AutoSize = True
        Me.opcPantalla.Location = New System.Drawing.Point(101, 110)
        Me.opcPantalla.Name = "opcPantalla"
        Me.opcPantalla.Size = New System.Drawing.Size(63, 17)
        Me.opcPantalla.TabIndex = 40
        Me.opcPantalla.TabStop = True
        Me.opcPantalla.Text = "Pantalla"
        Me.opcPantalla.UseVisualStyleBackColor = True
        '
        'ListadoDeudaPyme
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(412, 223)
        Me.Controls.Add(Me.btnCancela)
        Me.Controls.Add(Me.btnAcepta)
        Me.Controls.Add(Me.txthastafecha)
        Me.Controls.Add(Me.txtDesdeFecha)
        Me.Controls.Add(Me.CustomLabel2)
        Me.Controls.Add(Me.CustomLabel1)
        Me.Controls.Add(Me.opcImpesora)
        Me.Controls.Add(Me.opcPantalla)
        Me.Name = "ListadoDeudaPyme"
        Me.Text = "Listado de Deuda de Pyme Nacion"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancela As System.Windows.Forms.Button
    Friend WithEvents btnAcepta As System.Windows.Forms.Button
    Friend WithEvents txthastafecha As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtDesdeFecha As System.Windows.Forms.MaskedTextBox
    Friend WithEvents CustomLabel2 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel1 As WindowsApplication1.CustomLabel
    Friend WithEvents opcImpesora As System.Windows.Forms.RadioButton
    Friend WithEvents opcPantalla As System.Windows.Forms.RadioButton
End Class