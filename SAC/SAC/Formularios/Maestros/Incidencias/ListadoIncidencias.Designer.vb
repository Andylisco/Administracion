<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ListadoIncidencias
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.cmbOrdenIII = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmbOrdenII = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmbOrdenI = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnNuevaIncidencia = New System.Windows.Forms.Button()
        Me.btnCerrar = New System.Windows.Forms.Button()
        Me.btnFiltrar = New System.Windows.Forms.Button()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.clbPlantas = New System.Windows.Forms.CheckedListBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.clbTipos = New System.Windows.Forms.CheckedListBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.clbEstados = New System.Windows.Forms.CheckedListBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rbEntreFechas = New System.Windows.Forms.RadioButton()
        Me.rbAñosCompletos = New System.Windows.Forms.RadioButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtHastaAño = New System.Windows.Forms.MaskedTextBox()
        Me.txtDesdeAño = New System.Windows.Forms.MaskedTextBox()
        Me.txtHastaFecha = New System.Windows.Forms.MaskedTextBox()
        Me.txtDesdeFecha = New System.Windows.Forms.MaskedTextBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.dgvIncidencias = New System.Windows.Forms.DataGridView()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.CopiarConCabecerasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopiarSóloDatosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Anio = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Numero = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Incidencia = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Tipo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DescTipo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Renglon = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Fecha = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FechaOrd = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Estado = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DescEstado = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Titulo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Referencia = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Producto = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Lote = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ClaveSAC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Planta = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Empresa = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.dgvIncidencias, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(74, Byte), Integer), CType(CType(95, Byte), Integer))
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1106, 50)
        Me.Panel1.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Control
        Me.Label2.Location = New System.Drawing.Point(25, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(303, 18)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "LISTADO DE INFORMES DE NO CONFORMIDADES"
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 15.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Control
        Me.Label1.Location = New System.Drawing.Point(953, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(143, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "SURFACTAN S.A."
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox5)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.btnNuevaIncidencia)
        Me.GroupBox1.Controls.Add(Me.btnCerrar)
        Me.GroupBox1.Controls.Add(Me.btnFiltrar)
        Me.GroupBox1.Controls.Add(Me.GroupBox6)
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(10, 10)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(10)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1086, 215)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Parámetros"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.cmbOrdenIII)
        Me.GroupBox5.Controls.Add(Me.Label5)
        Me.GroupBox5.Controls.Add(Me.cmbOrdenII)
        Me.GroupBox5.Controls.Add(Me.Label6)
        Me.GroupBox5.Controls.Add(Me.cmbOrdenI)
        Me.GroupBox5.Controls.Add(Me.Label7)
        Me.GroupBox5.Location = New System.Drawing.Point(218, 25)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(554, 59)
        Me.GroupBox5.TabIndex = 4
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Orden"
        '
        'cmbOrdenIII
        '
        Me.cmbOrdenIII.FormattingEnabled = True
        Me.cmbOrdenIII.Items.AddRange(New Object() {"Numero", "Tipo", "Estado", "Planta"})
        Me.cmbOrdenIII.Location = New System.Drawing.Point(419, 23)
        Me.cmbOrdenIII.Name = "cmbOrdenIII"
        Me.cmbOrdenIII.Size = New System.Drawing.Size(92, 21)
        Me.cmbOrdenIII.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(369, 27)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 13)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "Terciario"
        '
        'cmbOrdenII
        '
        Me.cmbOrdenII.FormattingEnabled = True
        Me.cmbOrdenII.Items.AddRange(New Object() {"Numero", "Tipo", "Estado", "Planta"})
        Me.cmbOrdenII.Location = New System.Drawing.Point(261, 23)
        Me.cmbOrdenII.Name = "cmbOrdenII"
        Me.cmbOrdenII.Size = New System.Drawing.Size(92, 21)
        Me.cmbOrdenII.TabIndex = 7
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(198, 27)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(61, 13)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "Secundario"
        '
        'cmbOrdenI
        '
        Me.cmbOrdenI.FormattingEnabled = True
        Me.cmbOrdenI.Items.AddRange(New Object() {"Numero", "Tipo", "Estado", "Planta"})
        Me.cmbOrdenI.Location = New System.Drawing.Point(90, 23)
        Me.cmbOrdenI.Name = "cmbOrdenI"
        Me.cmbOrdenI.Size = New System.Drawing.Size(92, 21)
        Me.cmbOrdenI.TabIndex = 6
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(43, 27)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(44, 13)
        Me.Label7.TabIndex = 4
        Me.Label7.Text = "Primario"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(790, 51)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(142, 33)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "EXPORTAR INDICE"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnNuevaIncidencia
        '
        Me.btnNuevaIncidencia.Location = New System.Drawing.Point(926, 14)
        Me.btnNuevaIncidencia.Name = "btnNuevaIncidencia"
        Me.btnNuevaIncidencia.Size = New System.Drawing.Size(142, 33)
        Me.btnNuevaIncidencia.TabIndex = 3
        Me.btnNuevaIncidencia.Text = "NUEVO IINFORME"
        Me.btnNuevaIncidencia.UseVisualStyleBackColor = True
        '
        'btnCerrar
        '
        Me.btnCerrar.Location = New System.Drawing.Point(936, 51)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(132, 33)
        Me.btnCerrar.TabIndex = 3
        Me.btnCerrar.Text = "CERRAR"
        Me.btnCerrar.UseVisualStyleBackColor = True
        '
        'btnFiltrar
        '
        Me.btnFiltrar.Location = New System.Drawing.Point(790, 14)
        Me.btnFiltrar.Name = "btnFiltrar"
        Me.btnFiltrar.Size = New System.Drawing.Size(132, 33)
        Me.btnFiltrar.TabIndex = 3
        Me.btnFiltrar.Text = "FILTRAR"
        Me.btnFiltrar.UseVisualStyleBackColor = True
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.clbPlantas)
        Me.GroupBox6.Location = New System.Drawing.Point(654, 86)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(196, 122)
        Me.GroupBox6.TabIndex = 1
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Planta"
        '
        'clbPlantas
        '
        Me.clbPlantas.CheckOnClick = True
        Me.clbPlantas.FormattingEnabled = True
        Me.clbPlantas.Items.AddRange(New Object() {"TODAS", "INCIDENCIA GENERAL MP/PT", "INCIDENCIA POR RECHAZO DE MP"})
        Me.clbPlantas.Location = New System.Drawing.Point(16, 18)
        Me.clbPlantas.Name = "clbPlantas"
        Me.clbPlantas.Size = New System.Drawing.Size(164, 94)
        Me.clbPlantas.TabIndex = 2
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.clbTipos)
        Me.GroupBox4.Location = New System.Drawing.Point(436, 86)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(196, 122)
        Me.GroupBox4.TabIndex = 1
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Tipo de Informe de NC"
        '
        'clbTipos
        '
        Me.clbTipos.CheckOnClick = True
        Me.clbTipos.FormattingEnabled = True
        Me.clbTipos.Items.AddRange(New Object() {"TODAS", "INCIDENCIA GENERAL MP/PT", "INCIDENCIA POR RECHAZO EN RECEPCIÓN"})
        Me.clbTipos.Location = New System.Drawing.Point(16, 18)
        Me.clbTipos.Name = "clbTipos"
        Me.clbTipos.Size = New System.Drawing.Size(164, 94)
        Me.clbTipos.TabIndex = 2
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.clbEstados)
        Me.GroupBox3.Location = New System.Drawing.Point(218, 86)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(196, 122)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Estado de Informe de NC"
        '
        'clbEstados
        '
        Me.clbEstados.CheckOnClick = True
        Me.clbEstados.FormattingEnabled = True
        Me.clbEstados.Location = New System.Drawing.Point(19, 18)
        Me.clbEstados.Name = "clbEstados"
        Me.clbEstados.Size = New System.Drawing.Size(159, 94)
        Me.clbEstados.TabIndex = 1
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rbEntreFechas)
        Me.GroupBox2.Controls.Add(Me.rbAñosCompletos)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.txtHastaAño)
        Me.GroupBox2.Controls.Add(Me.txtDesdeAño)
        Me.GroupBox2.Controls.Add(Me.txtHastaFecha)
        Me.GroupBox2.Controls.Add(Me.txtDesdeFecha)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 24)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox2.Size = New System.Drawing.Size(205, 184)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Período"
        '
        'rbEntreFechas
        '
        Me.rbEntreFechas.AutoSize = True
        Me.rbEntreFechas.Location = New System.Drawing.Point(113, 46)
        Me.rbEntreFechas.Name = "rbEntreFechas"
        Me.rbEntreFechas.Size = New System.Drawing.Size(88, 17)
        Me.rbEntreFechas.TabIndex = 2
        Me.rbEntreFechas.Text = "Entre Fechas"
        Me.rbEntreFechas.UseVisualStyleBackColor = True
        '
        'rbAñosCompletos
        '
        Me.rbAñosCompletos.AutoSize = True
        Me.rbAñosCompletos.Checked = True
        Me.rbAñosCompletos.Location = New System.Drawing.Point(7, 46)
        Me.rbAñosCompletos.Name = "rbAñosCompletos"
        Me.rbAñosCompletos.Size = New System.Drawing.Size(100, 17)
        Me.rbAñosCompletos.TabIndex = 2
        Me.rbAñosCompletos.TabStop = True
        Me.rbAñosCompletos.Text = "Años completos"
        Me.rbAñosCompletos.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(94, 86)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(16, 13)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Al"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(94, 123)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(16, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Al"
        '
        'txtHastaAño
        '
        Me.txtHastaAño.Location = New System.Drawing.Point(119, 82)
        Me.txtHastaAño.Mask = "0000"
        Me.txtHastaAño.Name = "txtHastaAño"
        Me.txtHastaAño.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtHastaAño.Size = New System.Drawing.Size(67, 20)
        Me.txtHastaAño.TabIndex = 0
        Me.txtHastaAño.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtDesdeAño
        '
        Me.txtDesdeAño.Location = New System.Drawing.Point(18, 82)
        Me.txtDesdeAño.Mask = "0000"
        Me.txtDesdeAño.Name = "txtDesdeAño"
        Me.txtDesdeAño.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtDesdeAño.Size = New System.Drawing.Size(67, 20)
        Me.txtDesdeAño.TabIndex = 0
        Me.txtDesdeAño.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtHastaFecha
        '
        Me.txtHastaFecha.Location = New System.Drawing.Point(119, 119)
        Me.txtHastaFecha.Mask = "00/00/0000"
        Me.txtHastaFecha.Name = "txtHastaFecha"
        Me.txtHastaFecha.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtHastaFecha.Size = New System.Drawing.Size(67, 20)
        Me.txtHastaFecha.TabIndex = 0
        Me.txtHastaFecha.Text = "00000000"
        Me.txtHastaFecha.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtDesdeFecha
        '
        Me.txtDesdeFecha.Location = New System.Drawing.Point(18, 119)
        Me.txtDesdeFecha.Mask = "00/00/0000"
        Me.txtDesdeFecha.Name = "txtDesdeFecha"
        Me.txtDesdeFecha.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtDesdeFecha.Size = New System.Drawing.Size(67, 20)
        Me.txtDesdeFecha.TabIndex = 0
        Me.txtDesdeFecha.Text = "00000000"
        Me.txtDesdeFecha.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.GroupBox1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.dgvIncidencias, 0, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 50)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 235.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1106, 447)
        Me.TableLayoutPanel1.TabIndex = 5
        '
        'dgvIncidencias
        '
        Me.dgvIncidencias.AllowUserToAddRows = False
        Me.dgvIncidencias.AllowUserToDeleteRows = False
        Me.dgvIncidencias.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvIncidencias.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Anio, Me.Numero, Me.Incidencia, Me.Tipo, Me.DescTipo, Me.Renglon, Me.Fecha, Me.FechaOrd, Me.Estado, Me.DescEstado, Me.Titulo, Me.Referencia, Me.Producto, Me.Lote, Me.ClaveSAC, Me.Planta, Me.Empresa})
        Me.dgvIncidencias.ContextMenuStrip = Me.ContextMenuStrip1
        Me.dgvIncidencias.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvIncidencias.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dgvIncidencias.Location = New System.Drawing.Point(3, 238)
        Me.dgvIncidencias.Name = "dgvIncidencias"
        Me.dgvIncidencias.ReadOnly = True
        Me.dgvIncidencias.RowHeadersWidth = 15
        Me.dgvIncidencias.RowTemplate.Height = 20
        Me.dgvIncidencias.ShowCellToolTips = False
        Me.dgvIncidencias.Size = New System.Drawing.Size(1100, 206)
        Me.dgvIncidencias.TabIndex = 3
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CopiarConCabecerasToolStripMenuItem, Me.CopiarSóloDatosToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(229, 48)
        '
        'CopiarConCabecerasToolStripMenuItem
        '
        Me.CopiarConCabecerasToolStripMenuItem.Name = "CopiarConCabecerasToolStripMenuItem"
        Me.CopiarConCabecerasToolStripMenuItem.Size = New System.Drawing.Size(228, 22)
        Me.CopiarConCabecerasToolStripMenuItem.Text = "Copiar Incluyendo Cabeceras"
        '
        'CopiarSóloDatosToolStripMenuItem
        '
        Me.CopiarSóloDatosToolStripMenuItem.Name = "CopiarSóloDatosToolStripMenuItem"
        Me.CopiarSóloDatosToolStripMenuItem.Size = New System.Drawing.Size(228, 22)
        Me.CopiarSóloDatosToolStripMenuItem.Text = "Copiar Sólo Datos"
        '
        'Anio
        '
        Me.Anio.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.Anio.DataPropertyName = "Anio"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Anio.DefaultCellStyle = DataGridViewCellStyle1
        Me.Anio.HeaderText = "Año"
        Me.Anio.MaxInputLength = 4
        Me.Anio.Name = "Anio"
        Me.Anio.ReadOnly = True
        Me.Anio.Width = 51
        '
        'Numero
        '
        Me.Numero.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.Numero.DataPropertyName = "Numero"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Numero.DefaultCellStyle = DataGridViewCellStyle2
        Me.Numero.HeaderText = "Numero"
        Me.Numero.Name = "Numero"
        Me.Numero.ReadOnly = True
        Me.Numero.Width = 69
        '
        'Incidencia
        '
        Me.Incidencia.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.Incidencia.DataPropertyName = "Incidencia"
        Me.Incidencia.HeaderText = "Nro"
        Me.Incidencia.Name = "Incidencia"
        Me.Incidencia.ReadOnly = True
        Me.Incidencia.Visible = False
        Me.Incidencia.Width = 49
        '
        'Tipo
        '
        Me.Tipo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.Tipo.DataPropertyName = "Tipo"
        Me.Tipo.HeaderText = "Tipo"
        Me.Tipo.Name = "Tipo"
        Me.Tipo.ReadOnly = True
        Me.Tipo.Visible = False
        Me.Tipo.Width = 53
        '
        'DescTipo
        '
        Me.DescTipo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.DescTipo.DataPropertyName = "DescTipo"
        Me.DescTipo.HeaderText = "Tipo"
        Me.DescTipo.Name = "DescTipo"
        Me.DescTipo.ReadOnly = True
        Me.DescTipo.Width = 53
        '
        'Renglon
        '
        Me.Renglon.DataPropertyName = "Renglon"
        Me.Renglon.HeaderText = "Renglon"
        Me.Renglon.Name = "Renglon"
        Me.Renglon.ReadOnly = True
        Me.Renglon.Visible = False
        '
        'Fecha
        '
        Me.Fecha.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.Fecha.DataPropertyName = "Fecha"
        Me.Fecha.HeaderText = "Fecha"
        Me.Fecha.Name = "Fecha"
        Me.Fecha.ReadOnly = True
        Me.Fecha.Width = 62
        '
        'FechaOrd
        '
        Me.FechaOrd.DataPropertyName = "FechaOrd"
        Me.FechaOrd.HeaderText = "FechaOrd"
        Me.FechaOrd.Name = "FechaOrd"
        Me.FechaOrd.ReadOnly = True
        Me.FechaOrd.Visible = False
        '
        'Estado
        '
        Me.Estado.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.Estado.DataPropertyName = "Estado"
        Me.Estado.HeaderText = "Estado"
        Me.Estado.Name = "Estado"
        Me.Estado.ReadOnly = True
        Me.Estado.Visible = False
        Me.Estado.Width = 65
        '
        'DescEstado
        '
        Me.DescEstado.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.DescEstado.DataPropertyName = "DescEstado"
        Me.DescEstado.HeaderText = "Estado"
        Me.DescEstado.Name = "DescEstado"
        Me.DescEstado.ReadOnly = True
        Me.DescEstado.Width = 65
        '
        'Titulo
        '
        Me.Titulo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Titulo.DataPropertyName = "Titulo"
        Me.Titulo.HeaderText = "Título"
        Me.Titulo.MinimumWidth = 350
        Me.Titulo.Name = "Titulo"
        Me.Titulo.ReadOnly = True
        '
        'Referencia
        '
        Me.Referencia.DataPropertyName = "Referencia"
        Me.Referencia.HeaderText = "Referencia"
        Me.Referencia.MinimumWidth = 350
        Me.Referencia.Name = "Referencia"
        Me.Referencia.ReadOnly = True
        Me.Referencia.Width = 350
        '
        'Producto
        '
        Me.Producto.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.Producto.DataPropertyName = "Producto"
        Me.Producto.HeaderText = "Producto"
        Me.Producto.Name = "Producto"
        Me.Producto.ReadOnly = True
        Me.Producto.Visible = False
        Me.Producto.Width = 75
        '
        'Lote
        '
        Me.Lote.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.Lote.DataPropertyName = "Lote"
        Me.Lote.HeaderText = "Lote/Partida"
        Me.Lote.Name = "Lote"
        Me.Lote.ReadOnly = True
        Me.Lote.Visible = False
        Me.Lote.Width = 91
        '
        'ClaveSAC
        '
        Me.ClaveSAC.DataPropertyName = "ClaveSAC"
        Me.ClaveSAC.HeaderText = "ClaveSAC"
        Me.ClaveSAC.Name = "ClaveSAC"
        Me.ClaveSAC.ReadOnly = True
        Me.ClaveSAC.Visible = False
        '
        'Planta
        '
        Me.Planta.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.Planta.DataPropertyName = "Planta"
        Me.Planta.HeaderText = "Planta"
        Me.Planta.Name = "Planta"
        Me.Planta.ReadOnly = True
        Me.Planta.Width = 62
        '
        'Empresa
        '
        Me.Empresa.DataPropertyName = "Empresa"
        Me.Empresa.HeaderText = "Empresa"
        Me.Empresa.Name = "Empresa"
        Me.Empresa.ReadOnly = True
        Me.Empresa.Visible = False
        '
        'ListadoIncidencias
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1106, 497)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "ListadoIncidencias"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.dgvIncidencias, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtHastaFecha As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtDesdeFecha As System.Windows.Forms.MaskedTextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnFiltrar As System.Windows.Forms.Button
    Friend WithEvents btnCerrar As System.Windows.Forms.Button
    Friend WithEvents dgvIncidencias As System.Windows.Forms.DataGridView
    Friend WithEvents btnNuevaIncidencia As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents clbTipos As System.Windows.Forms.CheckedListBox
    Friend WithEvents clbEstados As System.Windows.Forms.CheckedListBox
    Friend WithEvents rbEntreFechas As System.Windows.Forms.RadioButton
    Friend WithEvents rbAñosCompletos As System.Windows.Forms.RadioButton
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtHastaAño As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtDesdeAño As System.Windows.Forms.MaskedTextBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbOrdenIII As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmbOrdenII As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmbOrdenI As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents clbPlantas As System.Windows.Forms.CheckedListBox
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents CopiarConCabecerasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CopiarSóloDatosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Anio As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Numero As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Incidencia As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Tipo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DescTipo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Renglon As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Fecha As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FechaOrd As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Estado As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DescEstado As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Titulo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Referencia As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Producto As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Lote As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ClaveSAC As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Planta As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Empresa As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
