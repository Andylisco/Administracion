<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DetallesIncidencia
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtAnio = New System.Windows.Forms.TextBox()
        Me.txtNumero = New System.Windows.Forms.TextBox()
        Me.txtTipo = New System.Windows.Forms.TextBox()
        Me.txtLotePartida = New System.Windows.Forms.TextBox()
        Me.btnDesvincularSAC = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.btnSac = New System.Windows.Forms.Button()
        Me.cmbEmpresaIncidencia = New System.Windows.Forms.ComboBox()
        Me.cmbEstado = New System.Windows.Forms.ComboBox()
        Me.txtFecha = New System.Windows.Forms.MaskedTextBox()
        Me.txtProducto = New System.Windows.Forms.MaskedTextBox()
        Me.rbVario = New System.Windows.Forms.RadioButton()
        Me.rbProdTerminado = New System.Windows.Forms.RadioButton()
        Me.rbMatPrima = New System.Windows.Forms.RadioButton()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtReferencia = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtIncidencia = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtDescTipo = New System.Windows.Forms.TextBox()
        Me.txtDescProducto = New System.Windows.Forms.TextBox()
        Me.txtTitulo = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnMovimientos = New System.Windows.Forms.Button()
        Me.btnHojaProduccion = New System.Windows.Forms.Button()
        Me.btnControles = New System.Windows.Forms.Button()
        Me.btnCerrar = New System.Windows.Forms.Button()
        Me.btnGrabar = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.txtMotivos = New System.Windows.Forms.TextBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtPosiblesUsos = New System.Windows.Forms.TextBox()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.dgvArchivos = New System.Windows.Forms.DataGridView()
        Me.FechaArchivo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DescArchivo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Icono = New System.Windows.Forms.DataGridViewImageColumn()
        Me.PathArchivo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.EliminarArchivoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnModificarNumeracion = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvArchivos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip2.SuspendLayout()
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
        Me.Panel1.Size = New System.Drawing.Size(831, 50)
        Me.Panel1.TabIndex = 3
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
        Me.Label2.Size = New System.Drawing.Size(269, 18)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "INFORME DE NO CONFORMIDAD GENERAL"
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Control
        Me.Label1.Location = New System.Drawing.Point(711, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(108, 18)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "SURFACTAN S.A."
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtAnio)
        Me.GroupBox1.Controls.Add(Me.txtNumero)
        Me.GroupBox1.Controls.Add(Me.txtTipo)
        Me.GroupBox1.Controls.Add(Me.txtLotePartida)
        Me.GroupBox1.Controls.Add(Me.btnDesvincularSAC)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.btnSac)
        Me.GroupBox1.Controls.Add(Me.cmbEmpresaIncidencia)
        Me.GroupBox1.Controls.Add(Me.cmbEstado)
        Me.GroupBox1.Controls.Add(Me.txtFecha)
        Me.GroupBox1.Controls.Add(Me.txtProducto)
        Me.GroupBox1.Controls.Add(Me.rbVario)
        Me.GroupBox1.Controls.Add(Me.rbProdTerminado)
        Me.GroupBox1.Controls.Add(Me.rbMatPrima)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.txtReferencia)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.txtIncidencia)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtDescTipo)
        Me.GroupBox1.Controls.Add(Me.txtDescProducto)
        Me.GroupBox1.Controls.Add(Me.txtTitulo)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 61)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(818, 168)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "DATOS GENERALES"
        '
        'txtAnio
        '
        Me.txtAnio.Location = New System.Drawing.Point(72, 21)
        Me.txtAnio.MaxLength = 6
        Me.txtAnio.Name = "txtAnio"
        Me.txtAnio.Size = New System.Drawing.Size(53, 20)
        Me.txtAnio.TabIndex = 7
        Me.txtAnio.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtNumero
        '
        Me.txtNumero.Location = New System.Drawing.Point(499, 21)
        Me.txtNumero.MaxLength = 6
        Me.txtNumero.Name = "txtNumero"
        Me.txtNumero.Size = New System.Drawing.Size(53, 20)
        Me.txtNumero.TabIndex = 7
        Me.txtNumero.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtTipo
        '
        Me.txtTipo.Location = New System.Drawing.Point(171, 21)
        Me.txtTipo.MaxLength = 6
        Me.txtTipo.Name = "txtTipo"
        Me.txtTipo.Size = New System.Drawing.Size(53, 20)
        Me.txtTipo.TabIndex = 7
        Me.txtTipo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtLotePartida
        '
        Me.txtLotePartida.Location = New System.Drawing.Point(444, 79)
        Me.txtLotePartida.MaxLength = 6
        Me.txtLotePartida.Name = "txtLotePartida"
        Me.txtLotePartida.Size = New System.Drawing.Size(53, 20)
        Me.txtLotePartida.TabIndex = 7
        Me.txtLotePartida.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnDesvincularSAC
        '
        Me.btnDesvincularSAC.Location = New System.Drawing.Point(699, 47)
        Me.btnDesvincularSAC.Name = "btnDesvincularSAC"
        Me.btnDesvincularSAC.Size = New System.Drawing.Size(109, 25)
        Me.btnDesvincularSAC.TabIndex = 12
        Me.btnDesvincularSAC.Text = "Desvincular SAC"
        Me.btnDesvincularSAC.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(584, 47)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(109, 25)
        Me.Button3.TabIndex = 12
        Me.Button3.Text = "Exportar INC"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'btnSac
        '
        Me.btnSac.Location = New System.Drawing.Point(469, 47)
        Me.btnSac.Name = "btnSac"
        Me.btnSac.Size = New System.Drawing.Size(109, 25)
        Me.btnSac.TabIndex = 12
        Me.btnSac.Text = "Ver SAC Asociada"
        Me.btnSac.UseVisualStyleBackColor = True
        '
        'cmbEmpresaIncidencia
        '
        Me.cmbEmpresaIncidencia.FormattingEnabled = True
        Me.cmbEmpresaIncidencia.Items.AddRange(New Object() {"", "S I", "P I", "S II", "P II", "S III", "S IV", "S V", "P III", "P V", "S VI", "S VII"})
        Me.cmbEmpresaIncidencia.Location = New System.Drawing.Point(754, 79)
        Me.cmbEmpresaIncidencia.Name = "cmbEmpresaIncidencia"
        Me.cmbEmpresaIncidencia.Size = New System.Drawing.Size(58, 21)
        Me.cmbEmpresaIncidencia.TabIndex = 3
        '
        'cmbEstado
        '
        Me.cmbEstado.FormattingEnabled = True
        Me.cmbEstado.Items.AddRange(New Object() {"PENDIENTE DE ANÁLISIS", "GENERA SAC", "NO GENERA SAC"})
        Me.cmbEstado.Location = New System.Drawing.Point(205, 49)
        Me.cmbEstado.Name = "cmbEstado"
        Me.cmbEstado.Size = New System.Drawing.Size(164, 21)
        Me.cmbEstado.TabIndex = 3
        '
        'txtFecha
        '
        Me.txtFecha.Location = New System.Drawing.Point(72, 49)
        Me.txtFecha.Mask = "00/00/0000"
        Me.txtFecha.Name = "txtFecha"
        Me.txtFecha.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtFecha.Size = New System.Drawing.Size(67, 20)
        Me.txtFecha.TabIndex = 2
        Me.txtFecha.Text = "99999999"
        Me.txtFecha.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtProducto
        '
        Me.txtProducto.Location = New System.Drawing.Point(291, 79)
        Me.txtProducto.Mask = ">LL-00000-000"
        Me.txtProducto.Name = "txtProducto"
        Me.txtProducto.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtProducto.Size = New System.Drawing.Size(78, 20)
        Me.txtProducto.TabIndex = 6
        Me.txtProducto.Text = "PT00000000"
        Me.txtProducto.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'rbVario
        '
        Me.rbVario.AutoSize = True
        Me.rbVario.Location = New System.Drawing.Point(186, 81)
        Me.rbVario.Name = "rbVario"
        Me.rbVario.Size = New System.Drawing.Size(49, 17)
        Me.rbVario.TabIndex = 5
        Me.rbVario.Text = "Vario"
        Me.rbVario.UseVisualStyleBackColor = True
        '
        'rbProdTerminado
        '
        Me.rbProdTerminado.AutoSize = True
        Me.rbProdTerminado.Location = New System.Drawing.Point(103, 81)
        Me.rbProdTerminado.Name = "rbProdTerminado"
        Me.rbProdTerminado.Size = New System.Drawing.Size(80, 17)
        Me.rbProdTerminado.TabIndex = 5
        Me.rbProdTerminado.Text = "Prod. Term."
        Me.rbProdTerminado.UseVisualStyleBackColor = True
        '
        'rbMatPrima
        '
        Me.rbMatPrima.AutoSize = True
        Me.rbMatPrima.Checked = True
        Me.rbMatPrima.Location = New System.Drawing.Point(24, 81)
        Me.rbMatPrima.Name = "rbMatPrima"
        Me.rbMatPrima.Size = New System.Drawing.Size(75, 17)
        Me.rbMatPrima.TabIndex = 4
        Me.rbMatPrima.TabStop = True
        Me.rbMatPrima.Text = "Mat. Prima"
        Me.rbMatPrima.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(446, 25)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(47, 13)
        Me.Label13.TabIndex = 0
        Me.Label13.Text = "Número:"
        '
        'txtReferencia
        '
        Me.txtReferencia.Location = New System.Drawing.Point(72, 136)
        Me.txtReferencia.MaxLength = 100
        Me.txtReferencia.Name = "txtReferencia"
        Me.txtReferencia.Size = New System.Drawing.Size(740, 20)
        Me.txtReferencia.TabIndex = 9
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(37, 25)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(29, 13)
        Me.Label12.TabIndex = 0
        Me.Label12.Text = "Año:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(29, 53)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(40, 13)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Fecha:"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(134, 25)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(31, 13)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "Tipo:"
        '
        'txtIncidencia
        '
        Me.txtIncidencia.Location = New System.Drawing.Point(754, 21)
        Me.txtIncidencia.Name = "txtIncidencia"
        Me.txtIncidencia.Size = New System.Drawing.Size(53, 20)
        Me.txtIncidencia.TabIndex = 1
        Me.txtIncidencia.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(376, 83)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(69, 13)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "Lote/Partida:"
        '
        'txtDescTipo
        '
        Me.txtDescTipo.BackColor = System.Drawing.Color.Cyan
        Me.txtDescTipo.Location = New System.Drawing.Point(230, 21)
        Me.txtDescTipo.Name = "txtDescTipo"
        Me.txtDescTipo.ReadOnly = True
        Me.txtDescTipo.Size = New System.Drawing.Size(202, 20)
        Me.txtDescTipo.TabIndex = 1
        Me.txtDescTipo.TabStop = False
        '
        'txtDescProducto
        '
        Me.txtDescProducto.BackColor = System.Drawing.Color.Cyan
        Me.txtDescProducto.Location = New System.Drawing.Point(500, 79)
        Me.txtDescProducto.Name = "txtDescProducto"
        Me.txtDescProducto.ReadOnly = True
        Me.txtDescProducto.Size = New System.Drawing.Size(184, 20)
        Me.txtDescProducto.TabIndex = 1
        Me.txtDescProducto.TabStop = False
        '
        'txtTitulo
        '
        Me.txtTitulo.Location = New System.Drawing.Point(72, 108)
        Me.txtTitulo.MaxLength = 100
        Me.txtTitulo.Name = "txtTitulo"
        Me.txtTitulo.Size = New System.Drawing.Size(740, 20)
        Me.txtTitulo.TabIndex = 8
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(690, 83)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(62, 13)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "Origen INC:"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(158, 53)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(43, 13)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "Estado:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(237, 83)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(53, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Producto:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(696, 25)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(52, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "ID Único:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 137)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(62, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Referencia:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(31, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Título:"
        '
        'btnMovimientos
        '
        Me.btnMovimientos.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMovimientos.Location = New System.Drawing.Point(585, 513)
        Me.btnMovimientos.Name = "btnMovimientos"
        Me.btnMovimientos.Size = New System.Drawing.Size(120, 41)
        Me.btnMovimientos.TabIndex = 17
        Me.btnMovimientos.Text = "VER MOVIMIENTOS"
        Me.btnMovimientos.UseVisualStyleBackColor = True
        '
        'btnHojaProduccion
        '
        Me.btnHojaProduccion.Location = New System.Drawing.Point(466, 513)
        Me.btnHojaProduccion.Name = "btnHojaProduccion"
        Me.btnHojaProduccion.Size = New System.Drawing.Size(115, 41)
        Me.btnHojaProduccion.TabIndex = 16
        Me.btnHojaProduccion.Text = "VER HOJA PROD."
        Me.btnHojaProduccion.UseVisualStyleBackColor = True
        '
        'btnControles
        '
        Me.btnControles.Location = New System.Drawing.Point(347, 513)
        Me.btnControles.Name = "btnControles"
        Me.btnControles.Size = New System.Drawing.Size(115, 41)
        Me.btnControles.TabIndex = 15
        Me.btnControles.Text = "VER CONTROLES"
        Me.btnControles.UseVisualStyleBackColor = True
        '
        'btnCerrar
        '
        Me.btnCerrar.Location = New System.Drawing.Point(709, 513)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(107, 41)
        Me.btnCerrar.TabIndex = 14
        Me.btnCerrar.Text = "CERRAR"
        Me.btnCerrar.UseVisualStyleBackColor = True
        '
        'btnGrabar
        '
        Me.btnGrabar.Location = New System.Drawing.Point(14, 513)
        Me.btnGrabar.Name = "btnGrabar"
        Me.btnGrabar.Size = New System.Drawing.Size(107, 41)
        Me.btnGrabar.TabIndex = 13
        Me.btnGrabar.Text = "GRABAR"
        Me.btnGrabar.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl1.ItemSize = New System.Drawing.Size(265, 30)
        Me.TabControl1.Location = New System.Drawing.Point(12, 235)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(804, 272)
        Me.TabControl1.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
        Me.TabControl1.TabIndex = 6
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.GroupBox3)
        Me.TabPage1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabPage1.Location = New System.Drawing.Point(4, 34)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(796, 234)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "MOTIVOS"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.txtMotivos)
        Me.GroupBox3.Location = New System.Drawing.Point(9, 3)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(778, 223)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        '
        'txtMotivos
        '
        Me.txtMotivos.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMotivos.Location = New System.Drawing.Point(13, 16)
        Me.txtMotivos.Margin = New System.Windows.Forms.Padding(10)
        Me.txtMotivos.Multiline = True
        Me.txtMotivos.Name = "txtMotivos"
        Me.txtMotivos.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtMotivos.Size = New System.Drawing.Size(753, 197)
        Me.txtMotivos.TabIndex = 10
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage2.Controls.Add(Me.GroupBox2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 34)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(796, 234)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "POSIBLES USOS / ACCIONES A TOMAR"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtPosiblesUsos)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(9, 5)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(778, 223)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        '
        'txtPosiblesUsos
        '
        Me.txtPosiblesUsos.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPosiblesUsos.Location = New System.Drawing.Point(13, 16)
        Me.txtPosiblesUsos.Margin = New System.Windows.Forms.Padding(10)
        Me.txtPosiblesUsos.Multiline = True
        Me.txtPosiblesUsos.Name = "txtPosiblesUsos"
        Me.txtPosiblesUsos.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtPosiblesUsos.Size = New System.Drawing.Size(753, 197)
        Me.txtPosiblesUsos.TabIndex = 11
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage3.Controls.Add(Me.Button2)
        Me.TabPage3.Controls.Add(Me.Button1)
        Me.TabPage3.Controls.Add(Me.dgvArchivos)
        Me.TabPage3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabPage3.Location = New System.Drawing.Point(4, 34)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(796, 234)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "ARCHIVOS"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(173, 6)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(164, 23)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Eliminar Archivo(s)"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(3, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(164, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Agregar Archivo(s)"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'dgvArchivos
        '
        Me.dgvArchivos.AllowDrop = True
        Me.dgvArchivos.AllowUserToAddRows = False
        Me.dgvArchivos.AllowUserToDeleteRows = False
        Me.dgvArchivos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvArchivos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.FechaArchivo, Me.DescArchivo, Me.Icono, Me.PathArchivo})
        Me.dgvArchivos.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.dgvArchivos.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dgvArchivos.Location = New System.Drawing.Point(3, 5)
        Me.dgvArchivos.Name = "dgvArchivos"
        Me.dgvArchivos.ReadOnly = True
        Me.dgvArchivos.RowHeadersWidth = 15
        Me.dgvArchivos.RowTemplate.Height = 30
        Me.dgvArchivos.ShowCellToolTips = False
        Me.dgvArchivos.Size = New System.Drawing.Size(790, 226)
        Me.dgvArchivos.TabIndex = 1
        '
        'FechaArchivo
        '
        Me.FechaArchivo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.FechaArchivo.HeaderText = "Fecha"
        Me.FechaArchivo.Name = "FechaArchivo"
        Me.FechaArchivo.ReadOnly = True
        Me.FechaArchivo.Width = 62
        '
        'DescArchivo
        '
        Me.DescArchivo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.DescArchivo.HeaderText = "Descripción"
        Me.DescArchivo.Name = "DescArchivo"
        Me.DescArchivo.ReadOnly = True
        '
        'Icono
        '
        Me.Icono.HeaderText = ""
        Me.Icono.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Zoom
        Me.Icono.Name = "Icono"
        Me.Icono.ReadOnly = True
        Me.Icono.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Icono.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Icono.Width = 50
        '
        'PathArchivo
        '
        Me.PathArchivo.HeaderText = "Path"
        Me.PathArchivo.Name = "PathArchivo"
        Me.PathArchivo.ReadOnly = True
        Me.PathArchivo.Visible = False
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EliminarArchivoToolStripMenuItem})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(162, 26)
        Me.ContextMenuStrip2.Text = "Eliminar Archivo"
        '
        'EliminarArchivoToolStripMenuItem
        '
        Me.EliminarArchivoToolStripMenuItem.Name = "EliminarArchivoToolStripMenuItem"
        Me.EliminarArchivoToolStripMenuItem.Size = New System.Drawing.Size(161, 22)
        Me.EliminarArchivoToolStripMenuItem.Text = "Eliminar Archivo"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'btnEliminar
        '
        Me.btnEliminar.Location = New System.Drawing.Point(125, 513)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(107, 41)
        Me.btnEliminar.TabIndex = 14
        Me.btnEliminar.Text = "ELIMINAR"
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'btnModificarNumeracion
        '
        Me.btnModificarNumeracion.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnModificarNumeracion.Location = New System.Drawing.Point(236, 513)
        Me.btnModificarNumeracion.Name = "btnModificarNumeracion"
        Me.btnModificarNumeracion.Size = New System.Drawing.Size(107, 41)
        Me.btnModificarNumeracion.TabIndex = 14
        Me.btnModificarNumeracion.Text = "MODIF. NUMERACIÓN INC"
        Me.btnModificarNumeracion.UseVisualStyleBackColor = True
        '
        'DetallesIncidencia
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(831, 557)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnMovimientos)
        Me.Controls.Add(Me.btnHojaProduccion)
        Me.Controls.Add(Me.btnControles)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnModificarNumeracion)
        Me.Controls.Add(Me.btnEliminar)
        Me.Controls.Add(Me.btnCerrar)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnGrabar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "DetallesIncidencia"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.dgvArchivos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtTitulo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtReferencia As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtIncidencia As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents rbProdTerminado As System.Windows.Forms.RadioButton
    Friend WithEvents rbMatPrima As System.Windows.Forms.RadioButton
    Friend WithEvents txtProducto As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtFecha As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtDescProducto As System.Windows.Forms.TextBox
    Friend WithEvents cmbEstado As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnSac As System.Windows.Forms.Button
    Friend WithEvents btnCerrar As System.Windows.Forms.Button
    Friend WithEvents btnGrabar As System.Windows.Forms.Button
    Friend WithEvents btnHojaProduccion As System.Windows.Forms.Button
    Friend WithEvents btnControles As System.Windows.Forms.Button
    Friend WithEvents btnMovimientos As System.Windows.Forms.Button
    Friend WithEvents txtLotePartida As System.Windows.Forms.TextBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtPosiblesUsos As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtMotivos As System.Windows.Forms.TextBox
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents dgvArchivos As System.Windows.Forms.DataGridView
    Friend WithEvents FechaArchivo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DescArchivo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Icono As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents PathArchivo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ContextMenuStrip2 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents EliminarArchivoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents btnDesvincularSAC As System.Windows.Forms.Button
    Friend WithEvents rbVario As System.Windows.Forms.RadioButton
    Friend WithEvents btnEliminar As System.Windows.Forms.Button
    Friend WithEvents btnModificarNumeracion As System.Windows.Forms.Button
    Friend WithEvents cmbEmpresaIncidencia As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtNumero As System.Windows.Forms.TextBox
    Friend WithEvents txtTipo As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtDescTipo As System.Windows.Forms.TextBox
    Friend WithEvents txtAnio As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
End Class
