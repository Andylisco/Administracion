﻿Imports System.ComponentModel
Imports Microsoft.VisualBasic.CompilerServices

<DesignerGenerated()> _
Partial Class MenuPrincipal
    Inherits Form

    'Form overrides dispose to clean up the component list.
    <DebuggerNonUserCode()> _
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
    Private components As IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.LayoutPrincipal = New TableLayoutPanel()
        Me.LayoutMenu = New TableLayoutPanel()
        Me.Panel4 = New Panel()
        Me.Label2 = New Label()
        Me.Panel2 = New Panel()
        Me.Label1 = New Label()
        Me.Panel1 = New Panel()
        Me.GroupBox1 = New GroupBox()
        Me.DataGridView1 = New DataGridView()
        Me.ComboBox1 = New ComboBox()
        Me.ProgressBar1 = New ProgressBar()
        Me.Button2 = New Button()
        Me.Button1 = New Button()
        Me.MenuStrip1 = New MenuStrip()
        Me.MaestrosToolStripMenuItem = New ToolStripMenuItem()
        Me.TiposDeSolicitudToolStripMenuItem = New ToolStripMenuItem()
        Me.ResponsablesToolStripMenuItem = New ToolStripMenuItem()
        Me.CentrosToolStripMenuItem = New ToolStripMenuItem()
        Me.IndiceGeneralDeAccionesToolStripMenuItem = New ToolStripMenuItem()
        Me.ProcesosToolStripMenuItem = New ToolStripMenuItem()
        Me.CerrarSistemaToolStripMenuItem = New ToolStripMenuItem()
        Me.LayoutPrincipal.SuspendLayout()
        Me.LayoutMenu.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'LayoutPrincipal
        '
        Me.LayoutPrincipal.ColumnCount = 1
        Me.LayoutPrincipal.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.LayoutPrincipal.Controls.Add(Me.LayoutMenu, 0, 0)
        Me.LayoutPrincipal.Controls.Add(Me.Panel1, 0, 1)
        Me.LayoutPrincipal.Dock = DockStyle.Fill
        Me.LayoutPrincipal.Location = New Point(0, 0)
        Me.LayoutPrincipal.Name = "LayoutPrincipal"
        Me.LayoutPrincipal.RowCount = 2
        Me.LayoutPrincipal.RowStyles.Add(New RowStyle(SizeType.Absolute, 46.0!))
        Me.LayoutPrincipal.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.LayoutPrincipal.Size = New Size(1041, 463)
        Me.LayoutPrincipal.TabIndex = 0
        '
        'LayoutMenu
        '
        Me.LayoutMenu.BackColor = Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(74, Byte), Integer), CType(CType(95, Byte), Integer))
        Me.LayoutMenu.ColumnCount = 3
        Me.LayoutMenu.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 294.0!))
        Me.LayoutMenu.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.LayoutMenu.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 247.0!))
        Me.LayoutMenu.Controls.Add(Me.Panel4, 2, 0)
        Me.LayoutMenu.Controls.Add(Me.Panel2, 0, 0)
        Me.LayoutMenu.Dock = DockStyle.Fill
        Me.LayoutMenu.Location = New Point(0, 0)
        Me.LayoutMenu.Margin = New Padding(0)
        Me.LayoutMenu.Name = "LayoutMenu"
        Me.LayoutMenu.RowCount = 1
        Me.LayoutMenu.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.LayoutMenu.Size = New Size(1041, 46)
        Me.LayoutMenu.TabIndex = 0
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.Label2)
        Me.Panel4.Dock = DockStyle.Fill
        Me.Panel4.Location = New Point(794, 0)
        Me.Panel4.Margin = New Padding(0)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New Size(247, 46)
        Me.Panel4.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New Font("Calibri", 18.0!, FontStyle.Bold)
        Me.Label2.ForeColor = SystemColors.Control
        Me.Label2.Location = New Point(35, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New Size(176, 29)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "SURFACTAN S.A."
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Dock = DockStyle.Fill
        Me.Panel2.Location = New Point(0, 0)
        Me.Panel2.Margin = New Padding(0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New Size(294, 46)
        Me.Panel2.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New Font("Calibri", 14.0!, FontStyle.Bold)
        Me.Label1.ForeColor = SystemColors.Control
        Me.Label1.Location = New Point(50, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New Size(115, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "SISTEMA SAC"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.MenuStrip1)
        Me.Panel1.Dock = DockStyle.Fill
        Me.Panel1.Location = New Point(0, 46)
        Me.Panel1.Margin = New Padding(0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New Size(1041, 417)
        Me.Panel1.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.ComboBox1)
        Me.GroupBox1.Controls.Add(Me.ProgressBar1)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Location = New Point(985, 252)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New Size(31, 149)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "GroupBox1"
        Me.GroupBox1.Visible = False
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New Point(152, 40)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New Size(831, 241)
        Me.DataGridView1.TabIndex = 2
        Me.DataGridView1.Visible = False
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"SurfactanSa", "Surfactan_II", "Surfactan_III", "Surfactan_IV", "Surfactan_V", "Surfactan_VI", "Surfactan_VII"})
        Me.ComboBox1.Location = New Point(6, 40)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New Size(130, 21)
        Me.ComboBox1.TabIndex = 4
        Me.ComboBox1.Visible = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New Point(7, 122)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New Size(128, 17)
        Me.ProgressBar1.TabIndex = 3
        Me.ProgressBar1.Visible = False
        '
        'Button2
        '
        Me.Button2.Location = New Point(3, 69)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New Size(133, 46)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Comprobar Guias"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New Point(12, 252)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New Size(133, 46)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Cambiar de Empresa"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New ToolStripItem() {Me.MaestrosToolStripMenuItem, Me.ProcesosToolStripMenuItem})
        Me.MenuStrip1.Location = New Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New Size(1041, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'MaestrosToolStripMenuItem
        '
        Me.MaestrosToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {Me.TiposDeSolicitudToolStripMenuItem, Me.ResponsablesToolStripMenuItem, Me.CentrosToolStripMenuItem, Me.IndiceGeneralDeAccionesToolStripMenuItem})
        Me.MaestrosToolStripMenuItem.Name = "MaestrosToolStripMenuItem"
        Me.MaestrosToolStripMenuItem.Size = New Size(67, 20)
        Me.MaestrosToolStripMenuItem.Text = "Maestros"
        '
        'TiposDeSolicitudToolStripMenuItem
        '
        Me.TiposDeSolicitudToolStripMenuItem.Name = "TiposDeSolicitudToolStripMenuItem"
        Me.TiposDeSolicitudToolStripMenuItem.Size = New Size(216, 22)
        Me.TiposDeSolicitudToolStripMenuItem.Text = "Tipos de Solicitud"
        '
        'ResponsablesToolStripMenuItem
        '
        Me.ResponsablesToolStripMenuItem.Name = "ResponsablesToolStripMenuItem"
        Me.ResponsablesToolStripMenuItem.Size = New Size(216, 22)
        Me.ResponsablesToolStripMenuItem.Text = "Responsables"
        '
        'CentrosToolStripMenuItem
        '
        Me.CentrosToolStripMenuItem.Name = "CentrosToolStripMenuItem"
        Me.CentrosToolStripMenuItem.Size = New Size(216, 22)
        Me.CentrosToolStripMenuItem.Text = "Centros"
        '
        'IndiceGeneralDeAccionesToolStripMenuItem
        '
        Me.IndiceGeneralDeAccionesToolStripMenuItem.Name = "IndiceGeneralDeAccionesToolStripMenuItem"
        Me.IndiceGeneralDeAccionesToolStripMenuItem.Size = New Size(216, 22)
        Me.IndiceGeneralDeAccionesToolStripMenuItem.Text = "Indice General de Acciones"
        '
        'ProcesosToolStripMenuItem
        '
        Me.ProcesosToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {Me.CerrarSistemaToolStripMenuItem})
        Me.ProcesosToolStripMenuItem.Name = "ProcesosToolStripMenuItem"
        Me.ProcesosToolStripMenuItem.Size = New Size(66, 20)
        Me.ProcesosToolStripMenuItem.Text = "Procesos"
        '
        'CerrarSistemaToolStripMenuItem
        '
        Me.CerrarSistemaToolStripMenuItem.Name = "CerrarSistemaToolStripMenuItem"
        Me.CerrarSistemaToolStripMenuItem.Size = New Size(150, 22)
        Me.CerrarSistemaToolStripMenuItem.Text = "Cerrar Sistema"
        '
        'MenuPrincipal
        '
        Me.AutoScaleDimensions = New SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = AutoScaleMode.Font
        Me.ClientSize = New Size(1041, 463)
        Me.Controls.Add(Me.LayoutPrincipal)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "MenuPrincipal"
        Me.WindowState = FormWindowState.Maximized
        Me.LayoutPrincipal.ResumeLayout(False)
        Me.LayoutMenu.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridView1, ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LayoutPrincipal As TableLayoutPanel
    Friend WithEvents LayoutMenu As TableLayoutPanel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Button1 As Button
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents MaestrosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProcesosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CerrarSistemaToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents TiposDeSolicitudToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ResponsablesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CentrosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents IndiceGeneralDeAccionesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Button2 As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents GroupBox1 As GroupBox

End Class
