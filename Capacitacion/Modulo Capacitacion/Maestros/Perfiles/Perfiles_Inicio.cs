﻿using System;
using System.Data;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using Modulo_Capacitacion.Listados;
using Modulo_Capacitacion.Listados.Perfiles;
using Negocio;

namespace Modulo_Capacitacion.Maestros.Perfiles
{
    public partial class Perfiles_Inicio : Form
    {
        Perfil P = new Perfil();            
        public Perfiles_Inicio()
        {
            InitializeComponent();
            
            CargarPerfiles();
        }

        private void CargarPerfiles()
        {
            DGV_Perfiles.DataSource = P.ListarTodosInicio();
            if (DGV_Perfiles.Columns["Descripcion"] != null)
                DGV_Perfiles.Columns["Descripcion"].Visible = false;
            if (DGV_Perfiles.Columns["Perfil"] != null)
                DGV_Perfiles.Columns["Perfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        private void BTAgregarPerfil_Click(object sender, EventArgs e)
        {
            AgModPerfil agregarModPerfil = new AgModPerfil();
            agregarModPerfil.StartPosition = FormStartPosition.CenterScreen;
            agregarModPerfil.ShowDialog();

            ActualizarGrilla();
        }

        private void ActualizarGrilla()
        {

            DGV_Perfiles.DataSource = P.ListarTodosInicio();
        }

        private void Bt_Fin_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void BT_Eliminar_Click(object sender, EventArgs e)
        {
            try
            {
                if (DGV_Perfiles.SelectedRows.Count != 1) throw new Exception("No hay filas seleccionadas o se selecciono mas de una");
                string IdAEliminar = DGV_Perfiles.SelectedRows[0].Cells[0].Value.ToString();

                if (MessageBox.Show("¿Está seguro de querer eliminar el perfil seleccionado?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    P.Eliminar(IdAEliminar);
                }

                ActualizarGrilla();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Error");
            }
        }

        private void BTModifSector_Click(object sender, EventArgs e)
        {
            try
            {
                if (DGV_Perfiles.SelectedRows.Count != 1) throw new Exception("Se debe seleccionar una fila a modificar");

                string IdAModificar = DGV_Perfiles.SelectedRows[0].Cells[0].Value.ToString();
                Perfil PerfilAModificar = new Perfil();
                PerfilAModificar = P.BuscarUno(IdAModificar);

                AgModPerfil modificarPerfil = new AgModPerfil(PerfilAModificar);
                modificarPerfil.StartPosition = FormStartPosition.CenterScreen;
                modificarPerfil.ShowDialog();

                ActualizarGrilla();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Error");
            }
        }


        private void TBFiltro_KeyUp(object sender, KeyEventArgs e)
        {
            
            DataTable dataTable = DGV_Perfiles.DataSource as DataTable;
            if (dataTable != null)
                dataTable.DefaultView.RowFilter = string.Format("CONVERT(Codigo, System.String) like '%{0}%' "
                                                + " OR CONVERT(Perfil, System.String) like '%{0}%'"
                                                + " OR CONVERT(Sector, System.String) like '%{0}%'"
                                                + " OR CONVERT(Vigencia, System.String) like '%{0}%'"
                                                + " OR CONVERT(Version, System.String) like '%{0}%'"
                                                + " OR CONVERT(Descripcion, System.String) like '%{0}%'", TBFiltro.Text);
            
        }

        private void DGV_Perfiles_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            BTModifSector.PerformClick();
        }

        private void Periles_Inicio_Shown(object sender, EventArgs e)
        {
            txtCodigo.Focus();
        }

        private void TBFiltro_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                TBFiltro.Text = "";
            }
        }

        private void btnImprimir_Click(object sender, EventArgs e)
        {
            ReportDocument reporte = new ListadoPerfiles();

            VistaPrevia rp = new VistaPrevia();
            rp.CargarReporte(reporte);
            rp.ShowDialog();
        }

        private void txtCodigo_KeyDown(object sender, KeyEventArgs e)
        {
            
            if (e.KeyData == Keys.Enter)
            {
                if (txtCodigo.Text.Trim() == "") return;

                foreach (DataGridViewRow row in DGV_Perfiles.Rows)
                {
                    var WCodigo = row.Cells["Codigo"].Value ?? "";

                    if (WCodigo.ToString().Trim() != "")
                    {
                        if (WCodigo.ToString().Trim() == txtCodigo.Text.Trim())
                        {
                            row.Selected = true;
                            DGV_Perfiles_RowHeaderMouseDoubleClick(null, null);
                            txtCodigo.Text = "";
                            return;
                        }
                    }
                }

            }
            else if (e.KeyData == Keys.Escape)
            {
                txtCodigo.Text = "";
            }
	        
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TBFiltro.Text = "";
            TBFiltro.Focus();
            TBFiltro_KeyUp(null, null);
        }

        private void DGV_Perfiles_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Perfiles_Inicio_Load(object sender, EventArgs e)
        {

        }
        
    }
}
