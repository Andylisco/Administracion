﻿using System.ComponentModel;
using System.Windows.Forms;

namespace Modulo_Capacitacion.Listados.CronogramaPorSectorYTema
{
    partial class Inicio
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Inicio));
            this.panel1 = new System.Windows.Forms.Panel();
            this.LBChofer = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.TB_Año = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.TB_HastaSector = new System.Windows.Forms.TextBox();
            this.TB_DesdeSector = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.BT_Pantalla = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.BT_Imprimir = new System.Windows.Forms.Button();
            this.BT_Salir = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(95)))), ((int)(((byte)(139)))), ((int)(((byte)(82)))));
            this.panel1.Controls.Add(this.LBChofer);
            this.panel1.Location = new System.Drawing.Point(-3, -1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(477, 39);
            this.panel1.TabIndex = 10;
            // 
            // LBChofer
            // 
            this.LBChofer.AutoSize = true;
            this.LBChofer.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LBChofer.ForeColor = System.Drawing.Color.White;
            this.LBChofer.Location = new System.Drawing.Point(10, 10);
            this.LBChofer.Name = "LBChofer";
            this.LBChofer.Size = new System.Drawing.Size(344, 19);
            this.LBChofer.TabIndex = 0;
            this.LBChofer.Text = "LISTADO DE CRONOGRAMA POR  SECTOR Y TEMA";
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(32)))), ((int)(((byte)(74)))), ((int)(((byte)(95)))));
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Location = new System.Drawing.Point(0, 37);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(476, 302);
            this.panel2.TabIndex = 13;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.SystemColors.Control;
            this.panel3.Controls.Add(this.TB_Año);
            this.panel3.Controls.Add(this.label3);
            this.panel3.Controls.Add(this.TB_HastaSector);
            this.panel3.Controls.Add(this.TB_DesdeSector);
            this.panel3.Controls.Add(this.label2);
            this.panel3.Controls.Add(this.BT_Pantalla);
            this.panel3.Controls.Add(this.button2);
            this.panel3.Controls.Add(this.button1);
            this.panel3.Controls.Add(this.BT_Imprimir);
            this.panel3.Controls.Add(this.BT_Salir);
            this.panel3.Controls.Add(this.label1);
            this.panel3.Location = new System.Drawing.Point(7, 7);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(397, 249);
            this.panel3.TabIndex = 0;
            // 
            // TB_Año
            // 
            this.TB_Año.Location = new System.Drawing.Point(231, 116);
            this.TB_Año.MaxLength = 4;
            this.TB_Año.Name = "TB_Año";
            this.TB_Año.Size = new System.Drawing.Size(51, 20);
            this.TB_Año.TabIndex = 87;
            this.TB_Año.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(83, 116);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(37, 18);
            this.label3.TabIndex = 86;
            this.label3.Text = "Año:";
            // 
            // TB_HastaSector
            // 
            this.TB_HastaSector.Location = new System.Drawing.Point(231, 77);
            this.TB_HastaSector.MaxLength = 4;
            this.TB_HastaSector.Name = "TB_HastaSector";
            this.TB_HastaSector.Size = new System.Drawing.Size(51, 20);
            this.TB_HastaSector.TabIndex = 85;
            this.TB_HastaSector.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.TB_HastaSector.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TB_HastaSector_KeyDown);
            this.TB_HastaSector.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.Ayuda_MouseDoubleClick);
            // 
            // TB_DesdeSector
            // 
            this.TB_DesdeSector.Location = new System.Drawing.Point(231, 36);
            this.TB_DesdeSector.MaxLength = 4;
            this.TB_DesdeSector.Name = "TB_DesdeSector";
            this.TB_DesdeSector.Size = new System.Drawing.Size(51, 20);
            this.TB_DesdeSector.TabIndex = 84;
            this.TB_DesdeSector.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.TB_DesdeSector.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TB_DesdeSector_KeyDown);
            this.TB_DesdeSector.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.Ayuda_MouseDoubleClick);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(83, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(88, 18);
            this.label2.TabIndex = 83;
            this.label2.Text = "Sector Hasta:";
            // 
            // BT_Pantalla
            // 
            this.BT_Pantalla.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("BT_Pantalla.BackgroundImage")));
            this.BT_Pantalla.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.BT_Pantalla.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BT_Pantalla.ForeColor = System.Drawing.SystemColors.Control;
            this.BT_Pantalla.Location = new System.Drawing.Point(83, 176);
            this.BT_Pantalla.Name = "BT_Pantalla";
            this.BT_Pantalla.Size = new System.Drawing.Size(40, 40);
            this.BT_Pantalla.TabIndex = 81;
            this.BT_Pantalla.UseVisualStyleBackColor = true;
            this.BT_Pantalla.Click += new System.EventHandler(this.BT_Pantalla_Click);
            // 
            // button2
            // 
            this.button2.BackgroundImage = global::Modulo_Capacitacion.Properties.Resources.Consulta_Dat_N1;
            this.button2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.ForeColor = System.Drawing.SystemColors.Control;
            this.button2.Location = new System.Drawing.Point(288, 74);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(25, 25);
            this.button2.TabIndex = 80;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.BackgroundImage = global::Modulo_Capacitacion.Properties.Resources.Consulta_Dat_N1;
            this.button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.ForeColor = System.Drawing.SystemColors.Control;
            this.button1.Location = new System.Drawing.Point(288, 34);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(25, 25);
            this.button1.TabIndex = 80;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // BT_Imprimir
            // 
            this.BT_Imprimir.BackgroundImage = global::Modulo_Capacitacion.Properties.Resources.imprimir;
            this.BT_Imprimir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.BT_Imprimir.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BT_Imprimir.ForeColor = System.Drawing.SystemColors.Control;
            this.BT_Imprimir.Location = new System.Drawing.Point(178, 176);
            this.BT_Imprimir.Name = "BT_Imprimir";
            this.BT_Imprimir.Size = new System.Drawing.Size(40, 40);
            this.BT_Imprimir.TabIndex = 80;
            this.BT_Imprimir.UseVisualStyleBackColor = true;
            this.BT_Imprimir.Click += new System.EventHandler(this.BT_Imprimir_Click);
            // 
            // BT_Salir
            // 
            this.BT_Salir.BackgroundImage = global::Modulo_Capacitacion.Properties.Resources.apagar;
            this.BT_Salir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.BT_Salir.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BT_Salir.ForeColor = System.Drawing.SystemColors.Control;
            this.BT_Salir.Location = new System.Drawing.Point(273, 176);
            this.BT_Salir.Name = "BT_Salir";
            this.BT_Salir.Size = new System.Drawing.Size(40, 40);
            this.BT_Salir.TabIndex = 79;
            this.BT_Salir.UseVisualStyleBackColor = true;
            this.BT_Salir.Click += new System.EventHandler(this.BT_Salir_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(83, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 18);
            this.label1.TabIndex = 1;
            this.label1.Text = "Sector Desde:";
            // 
            // Inicio
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(410, 300);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "Inicio";
            this.Load += new System.EventHandler(this.Inicio_Load);
            this.Shown += new System.EventHandler(this.Inicio_Shown);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Panel panel1;
        private Label LBChofer;
        private Panel panel2;
        private Panel panel3;
        private TextBox TB_Año;
        private Label label3;
        private TextBox TB_HastaSector;
        private TextBox TB_DesdeSector;
        private Label label2;
        private Button BT_Pantalla;
        private Button BT_Imprimir;
        private Button BT_Salir;
        private Label label1;
        private Button button1;
        private Button button2;
    }
}