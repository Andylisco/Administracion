﻿using System;
using System.Data;
using System.Windows.Forms;

namespace Modulo_Capacitacion.Listados.CronogramaPorSectorYTema
{
    public partial class ImpreInforme : Form
    {
        private DataTable dtInforme;
        private string Tipo;
        private int Año;

        public ImpreInforme()
        {
            InitializeComponent();
        }

        

        public ImpreInforme(DataTable dtInforme, string Tipo, int Año)
        {
            // TODO: Complete member initialization
            this.dtInforme = dtInforme;
            this.Tipo = Tipo;
            this.Año = Año;
            InitializeComponent();
        }

        private void ImpreInforme_Load(object sender, EventArgs e)
        {
            DSInforme Ds = new DSInforme();

            int filas = dtInforme.Rows.Count;

            for (int i =0; i < filas ; i++)
            {
                DataRow dr = dtInforme.Rows[i];
                Ds.Tables[0].Rows.Add
                (Año.ToString(), dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6]);
            }

            if (Tipo == "Pantalla")
            {
                //Hago visible el report viewer
                CRVInforme.Visible = true;


                //Instancio el ReportImpre creado
                Reporte RImpre = new Reporte();


                //Cargo el reportImpre con el dataset DS
                RImpre.SetDataSource(Ds);

                //Cargo el report viewer con el ReportImp
                CRVInforme.ReportSource = RImpre;
            }
            else
            {
                //Instancio el ReportImpre creado
                Reporte RImpre = new Reporte();

                //Cargo el reportImpre con el dataset DS
                RImpre.SetDataSource(Ds);

                RImpre.PrintToPrinter(1, true, 1, 999);
                Close();
            }
        }
    }
}
