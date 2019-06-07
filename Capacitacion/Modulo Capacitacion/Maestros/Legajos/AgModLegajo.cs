using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using Negocio;

namespace Modulo_Capacitacion.Maestros.Legajos
{
    public partial class AgModLegajo : Form
    {
        //private Legajo LegajoAModificar;
        private Legajo LegajoViejo = new Legajo();
        private Legajo L = new Legajo();
        private Perfil Per = new Perfil();
        private bool AModificar;
        Cursada C = new Cursada();
        LegajoVersion LegajoVer = new LegajoVersion();
        Sector S = new Sector();
        DataTable dtSectores;
        DataTable dtPerfil;
        bool Cargado;
        string ObservExtra1 = "";
        string ObservExtra2 = "";
        string ObservExtra3 = "";
        string ObservExtra4 = "";
        string ObservExtra5 = "";

        //private int indexRow;

        public AgModLegajo()
        {
            InitializeComponent();
            
            Legajo _L = new Legajo();
            TB_Codigo.Text = _L.ObtenerUltimoId().ToString();
            tabControl1.TabPages.Remove(tabPage3);
            CargarSectores();
            CargarPerfil();
            Cargado = true;
            

            //CargarCursadas();
        }

        private void CargarCursadas()
        {            
            DGV_CursosRealiz.DataSource = C.BuscarUnoGrilla(TB_Codigo.Text);
        }

        public AgModLegajo(Legajo LegajoAModificar)
        {
            InitializeComponent();
            _CargarLegajo(LegajoAModificar);
        }

        private void _CargarLegajo(Legajo LegajoAModificar)
        {
            L = LegajoAModificar;
            LegajoViejo = LegajoAModificar;

            CargarSectores();
            CargarPerfil();
            Cargado = true;

            AModificar = L.Codigo != 0;
            CargarDatosABM();
            CargarDatosPefil();
            //CargarTemas(L.Temas);
            CargarTemas(Per.Temas);
            CargarDatosLegajoVer();
            CargarObservaciones();

            TB_DescPerfil.Enabled = false;
            TB_CodSector.Enabled = false;
            TB_DescSec.Enabled = false;

            //Cargo cursos realizados de cursadas
            CargarCursadas();

            ComprobarModifPerfil();

            //TB_FechaIng.Enabled = false;

            if (TB_FechaEgreso.BackColor == Color.Red)
            {
                InhabilitarModifiaciones();
            }
        }

        private void CargarObservaciones()
        {
            if (TB_Codigo.Text.Trim() != "")
            {
                using (SqlConnection conn = new SqlConnection())
                {
                    conn.ConnectionString = ConfigurationManager.ConnectionStrings["Surfactan"].ConnectionString;
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "SELECT ObservaI1, ObservaI2, ObservaI3, ObservaI4, ObservaI5 FROM Legajo WHERE Codigo = '" + TB_Codigo.Text + "' AND Renglon = 1";

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                dr.Read();

                                ObservExtra1 = dr["ObservaI1"] == null ? "" : dr["ObservaI1"].ToString().Trim();
                                ObservExtra2 = dr["ObservaI2"] == null ? "" : dr["ObservaI2"].ToString().Trim();
                                ObservExtra3 = dr["ObservaI3"] == null ? "" : dr["ObservaI3"].ToString().Trim();
                                ObservExtra4 = dr["ObservaI4"] == null ? "" : dr["ObservaI4"].ToString().Trim();
                                ObservExtra5 = dr["ObservaI5"] == null ? "" : dr["ObservaI5"].ToString().Trim();

                            }
                        }
                    }

                }
        
            }
        }

        private void ComprobarModifPerfil()
        {
            if (L.Actualizado == "N")
            {
                TB_PerfilAct.Visible = true;
            }
        }

        private void InhabilitarModifiaciones()
        {
            TB_DescPerfil.Enabled = false;
            //DTP_Fecha.Text = DateTime.Now.ToShortDateString();
            TB_Version.ReadOnly = true;
            TB_CodSector.Enabled = false;
            TB_DescSec.Enabled = false;
            TB_CodPerfil.Enabled = false;
            TB_DescPerfil.Enabled = false;
            TB_VersPer.ReadOnly = true;
            TB_Tareas1.ReadOnly = true;
            TB_Tareas2.ReadOnly = true;
            TB_Tareas3.ReadOnly = true;
            TB_Primaria.ReadOnly = true;
            TB_Secundaria.ReadOnly = true;
            TB_Terciaria.ReadOnly = true;
            TB_Idioma.ReadOnly = true;
            TB_Exp.ReadOnly = true;
            TB_CondFisica.ReadOnly = true;
            TB_Otros1.ReadOnly = true;
            TB_Otros2.ReadOnly = true;
            TB_Equiv1.ReadOnly = true;
            TB_Equiv2.ReadOnly = true;
            TB_ObservCondFisica.ReadOnly = true;
            TB_ObservPrimaria.ReadOnly = true;
            TB_ObservSecundaria.ReadOnly = true;
            TB_ObservTerciaria.ReadOnly = true;
            TB_ObservIdioma.ReadOnly = true;
            TB_ObservExp.ReadOnly = true;
            CB_DesPrim.Enabled = false;
            CB_DesSec.Enabled = false;
            CB_DesTerc.Enabled = false;
            CB_DesIdioma.Enabled = false;
            CB_DesExp.Enabled = false;
            CB_DesCondFisica.Enabled = false;
            CB_DesOtros1.Enabled = false;
            CB_DesOtros2.Enabled = false;
            CB_NecPrim.Enabled = false;

            CB_NecSec.Enabled = false;
            CB_NecTerc.Enabled = false;
            CB_NecIdioma.Enabled = false;
            CB_NecExp.Enabled = false;
            CB_NecCondFisica.Enabled = false;
            CB_NecOtros1.Enabled = false;
            CB_NecOtros2.Enabled = false;

            BT_Guardar.Visible = false;
            
            
        }

        private void CargarTemas(List<Tema> list)
        {
            DataTable dtTemasGuardados = new DataTable();
            dtTemasGuardados.Rows.Clear();
            dtTemasGuardados.Columns.Add("Tema", typeof(string));
            dtTemasGuardados.Columns.Add("Descripcion", typeof(string));
            dtTemasGuardados.Columns.Add("Necesaria", typeof(string));
            dtTemasGuardados.Columns.Add("Deseable", typeof(string));
            dtTemasGuardados.Columns.Add("Estado", typeof(string));
            dtTemasGuardados.Columns.Add("Observaciones", typeof(string));

            if (list != null)
            {
                foreach (var item in list)
                {
                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(DGV_Temas);
                    row.Cells[0].Value = item.Codigo;
                    row.Cells[1].Value = item.Descripcion;
                    //necesario
                    row.Cells[2].Value = item.Necesaria == 1 ? "X" : "";
                    //deseable
                    row.Cells[3].Value = item.Deseable == 1 ? "X" : "";
                    //combo
                    // DataGridViewComboBoxColumn Estado = new DataGridViewComboBoxColumn();
                    // Estado.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton;
                    // Estado.HeaderText = "Estado";
                    //Estado. = 1;
                    row.Cells[4].Value = _BuscarEstadoCurso(item.Codigo);
                    //ObtenerValor(item.EstaCurso);
                    //observacion
                    //row.Cells[5].Value = item.Observacion;
                    row.Cells[5].Value = item.EstadoCurso;
                    //row.Cells[5].Value = item.Estado;

                    DGV_Temas.Rows.Add(row);
                }
            }
        }

        private string _BuscarEstadoCurso(int WCodigoCurso)
        {

            try
            {
                using (SqlConnection conn = new SqlConnection())
                {
                    conn.ConnectionString = ConfigurationManager.ConnectionStrings["Surfactan"].ConnectionString;
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "SELECT ISNULL(EstaCurso, 0) EstaCurso FROM Legajo WHERE Codigo = '" + TB_Codigo.Text + "' AND Curso = '" + WCodigoCurso + "' ";

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                dr.Read();

                                return ObtenerValor(int.Parse(dr["EstaCurso"].ToString()));
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al procesar la consulta a la Base de Datos. Motivo: " + ex.Message);
            }

            return "";
        }

        private string ObtenerValor(int Esta)
        {
            //estados d ela primera pantalla
            switch (Esta)
            {
                case 1:
                    return "Exced";
                case 2:
                    return "Cumple";
                case 3:
                    return "Reforzar";
                case 4:
                    return "En Entren";
                case 5:
                    return "No Cumple";
                case 6:
                    return "No Aplica";
                case 7:
                    return "No Evalua";
                case 8:
                    return "Cumple Act";

                default:
                    return "";
            }
        }

        private void CargarDatosABM()
        {
            if (L.Codigo != 0)
            {
                TB_Codigo.Text = L.Codigo.ToString().Trim();
                TB_DescLegajo.Text = L.Descripcion.Trim();
                TB_DNI.Text = L.DNI;
                TB_CUIL.Text = L.CUIL;
                if (L.FIngreso.Trim() != "")
                {
                    TB_FechaIng.Text = L.FIngreso;
                }

                DateTime FEgresoParse;
                DateTime.TryParse(L.FEgreso, out FEgresoParse);

                if (FEgresoParse.ToString("d/M/yyyy") == "1/1/0001")
                {
                    TB_FechaEgreso.Text = "00/00/0000";
                }
                else
                {
                    TB_FechaEgreso.Text = FEgresoParse.ToShortDateString();
                    TB_FechaEgreso.BackColor = Color.Red;
                }

                int valorVersion = int.Parse(L.Version);
                TB_Version.Text = valorVersion.ToString().Trim();
                DTP_Fecha.Text = L.FechaVersion;
                TB_CodPerfil.Text = L.Perfil.Codigo.ToString();

                //Falta los combos y observaciones
                TB_ObservPrimariaLeg.Text = L.EstadoI.Trim();
                TB_ObservSecundariaLeg.Text = L.EstadoII.Trim();
                TB_ObservTerciariaLeg.Text = L.EstadoIII.Trim();
                TB_ObservIdiomaLeg.Text = L.EstadoIV.Trim();
                TB_ObservExpLeg.Text = L.EstadoV.Trim();
                TB_ObservCondFisicaLeg.Text = L.EstadoVI.Trim();
                TB_Otros1Leg.Text = L.EstadoVII.Trim();
                TB_Otros2Leg.Text = L.EstadoVIII.Trim();
                TB_Equiv1Leg.Text = L.EstadoIX.Trim();
                TB_Equiv2Leg.Text = L.EstadoX.Trim();

                L.EstaX = string.IsNullOrEmpty(L.EstaX.Trim()) ? "0" : L.EstaX;

                CB_EstPrim.SelectedIndex = int.Parse(L.EstaI);
                CB_EstSec.SelectedIndex = int.Parse(L.EstaII);
                CB_EstTerc.SelectedIndex = int.Parse(L.EstaIII);
                CB_EstIdioma.SelectedIndex = int.Parse(L.EstaIV);
                CB_EstExp.SelectedIndex = int.Parse(L.EstaV);
                CB_EstCondFisic.SelectedIndex = int.Parse(L.EstaVI);
                CB_EstOtros1.SelectedIndex = int.Parse(L.EstaVII);
                CB_EstOtros2.SelectedIndex = int.Parse(L.EstaVIII);
                CB_EstEquiv1.SelectedIndex = int.Parse(L.EstaIX);
                CB_Estequiv2.SelectedIndex = int.Parse(L.EstaX);
            }
        }

        private void CargarTemas(Perfil PerfilTemas)
        {

            DGV_Temas.Rows.Clear();

            if (PerfilTemas.Temas == null) return;

            foreach (var item in PerfilTemas.Temas)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(DGV_Temas);
                row.Cells[0].Value = item.Codigo;
                row.Cells[1].Value = item.Descripcion;
                row.Cells[2].Value = item.Necesaria == 1 ? "X" : "";
                row.Cells[3].Value = item.Deseable == 1 ? "X" : "";
                row.Cells[4].Value = "";

                DGV_Temas.Rows.Add(row);
            }
        }

        private void CargarDatosPefil()
        {
            BuscarCodperfil();

            Per = _BuscarPerfilPorVersion() ? Per.BuscarUno(TB_CodPerfil.Text, TB_VersPer.Text) : Per.BuscarUno(TB_CodPerfil.Text);

            if (Per.Codigo == 0 && TB_CodPerfil.Text != "") throw new Exception("No se encontro elemento con el codigo ingresado");

            if (TB_CodPerfil.Text == "") return;

            TB_DescPerfil.Text = Per.Descripcion.Trim();
            TB_CodSector.Text = Per.sector.Codigo.ToString().Trim();
            TB_DescSec.Text = Per.sector.Descripcion.Trim();
            TB_VersPer.Text = Per.Version.ToString().Trim();
            TB_Tareas1.Text = Per.TareasI.Trim();
            TB_Tareas2.Text = Per.TareasII.Trim();
            TB_Tareas3.Text = Per.TareasIII.Trim();

            TB_Primaria.Text = Per.DescriI.Trim();
            TB_ObservPrimaria.Text = Per.ObservaI.Trim();
            CB_NecPrim.Checked = Per.NecesariaI == 1;
            CB_DesPrim.Checked = Per.DeseableI == 1;

            TB_Secundaria.Text = Per.DescriII.Trim();
            TB_ObservSecundaria.Text = Per.ObservaII.Trim();
            CB_NecSec.Checked = Per.NecesariaII == 1;
            CB_DesSec.Checked = Per.DeseableII == 1;

            TB_Terciaria.Text = Per.DescriIII.Trim();
            TB_ObservTerciaria.Text = Per.ObservaIII.Trim();
            CB_NecTerc.Checked = Per.NecesariaIII == 1;
            CB_DesTerc.Checked = Per.DeseableIII == 1;

            TB_Idioma.Text = Per.DescriIV.Trim();
            TB_ObservIdioma.Text = Per.ObservaIV.Trim();
            CB_NecIdioma.Checked = Per.NecesariaIV == 1;
            CB_DesIdioma.Checked = Per.DeseableIV == 1;

            TB_Exp.Text = Per.DescriV.Trim();
            TB_ObservExp.Text = Per.ObservaV.Trim();
            CB_NecExp.Checked = Per.NecesariaV == 1;
            CB_DesExp.Checked = Per.DeseableV == 1;

            TB_CondFisica.Text = Per.Fisica.Trim();
            CB_NecCondFisica.Checked = Per.NecesariaVI == 1;
            CB_DesCondFisica.Checked = Per.DeseableVI == 1;

            TB_Otros1.Text = Per.OtrosI.Trim();
            CB_DesOtros1.Checked = Per.DeseableVII == 1;
            CB_NecOtros1.Checked = Per.NecesariaVII == 1;

            TB_Otros2.Text = Per.OtrosII.Trim();
            CB_DesOtros2.Checked = Per.DeseableVIII == 1;
            CB_NecOtros2.Checked = Per.NecesariaVIII == 1;

            TB_Equiv1.Text = Per.EquivalenciasI.Trim();
            TB_Equiv2.Text = Per.EquivalenciasII.Trim();
        }

        private bool _BuscarPerfilPorVersion()
        {
            bool WBuscarPorVersion = false;

            try
            {
                int WVersionPerfilLegajo = 0;
                int WVersionTarea = 0;

                using (SqlConnection conn = new SqlConnection())
                {
                    conn.ConnectionString = ConfigurationManager.ConnectionStrings["Surfactan"].ConnectionString;
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "SELECT VersionDePerfil = CASE WHEN Fegreso IS NULL OR Fegreso IN ('00/00/0000', '  /  /    ') THEN '0' ELSE PerfilVersion END FROM Legajo WHERE Codigo = '" + TB_Codigo.Text + "' AND Renglon = '1'";

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                dr.Read();
                                WVersionPerfilLegajo = int.Parse(dr["VersionDePerfil"].ToString());
                            }
                        }

                        cmd.CommandText = "SELECT Version FROM Tarea WHERE Codigo = '" + TB_CodPerfil.Text + "' AND Renglon = '1'";

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                dr.Read();
                                WVersionTarea = int.Parse(dr["Version"].ToString());
                            }
                        }

                        TB_VersPer.Text = WVersionTarea.ToString();

                        if (!(WVersionPerfilLegajo == 0 || WVersionTarea == 0 || WVersionTarea == WVersionPerfilLegajo))
                        {
                            TB_VersPer.Text = WVersionPerfilLegajo.ToString();
                            WBuscarPorVersion = true;
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al procesar la consulta a la Base de Datos. Motivo: " + ex.Message);
            }

            return WBuscarPorVersion;
        }

        private void BT_Salir_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void BT_LimpiarPant_Click(object sender, EventArgs e)
        {
            TB_FechaIng.Enabled = true;
            if (!AModificar)
            {
                TB_DescPerfil.Text = "";
                DTP_Fecha.Text = DateTime.Now.ToShortDateString();
                TB_Version.Text = "1";
                TB_CodSector.Text = "";
                TB_DescSec.Text = "";
                TB_CodPerfil.Text = "";
                TB_DescPerfil.Text = "";
                TB_VersPer.Text = "";
                TB_Tareas1.Text = "";
                TB_Tareas2.Text = "";
                TB_Tareas3.Text = "";
                TB_Primaria.Text = "";
                TB_Secundaria.Text = "";
                TB_Terciaria.Text = "";
                TB_Idioma.Text = "";
                TB_Exp.Text = "";
                TB_CondFisica.Text = "";
                TB_Otros1.Text = "";
                TB_Otros2.Text = "";
                TB_Equiv1.Text = "";
                TB_Equiv2.Text = "";
                TB_ObservCondFisica.Text = "";
                TB_ObservPrimaria.Text = "";
                TB_ObservSecundaria.Text = "";
                TB_ObservTerciaria.Text = "";
                TB_ObservIdioma.Text = "";
                TB_ObservExp.Text = "";
                CB_DesPrim.Checked = false;
                CB_DesSec.Checked = false;
                CB_DesTerc.Checked = false;
                CB_DesIdioma.Checked = false;
                CB_DesExp.Checked = false;
                CB_DesCondFisica.Checked = false;
                CB_DesOtros1.Checked = false;
                CB_DesOtros2.Checked = false;
                CB_NecPrim.Checked = false;

                CB_NecSec.Checked = false;
                CB_NecTerc.Checked = false;
                CB_NecIdioma.Checked = false;
                CB_NecExp.Checked = false;
                CB_NecCondFisica.Checked = false;
                CB_NecOtros1.Checked = false;
                CB_NecOtros2.Checked = false;

                DGV_Temas.DataSource = null;
                DGV_Temas.Rows.Clear();
            }
        }

        private void BT_Guardar_Click(object sender, EventArgs e)
        {
            try
            {
                //VALIDO QUE ESTEN TODOS LOS DATOS
                ValidarDatosCargados();

                //Cargo los datos al objeto L
                CargarDatosLegajo();

                if (AModificar)
                {
                    DialogResult Result = MessageBox.Show("Desea modificar la versión?", "Modificar Versión", MessageBoxButtons.YesNo);

                    if (Result == DialogResult.Yes)
                    {
                        CargarDatosLegajoVer();

                        LegajoVer.Agregar();

                        L.Eliminar(L.Codigo.ToString());

                        //COLOCO LA FECHA DEL DIA 
                        DateTime Hoy = DateTime.Today;
                        L.FechaVersion = Hoy.ToString("dd/MM/yyyy");

                        //SUMO UNO A VERSION

                        L.Version = (int.Parse(TB_Version.Text) + 1).ToString();

                        L.Agregar();

                        MessageBox.Show("Se ha modificado el legajo", "Modificar Legajo",
                        MessageBoxButtons.OK, MessageBoxIcon.None);

                    }

                    else if (Result == DialogResult.No)
                    {
                        //COLOCO LA FECHA DEL DIA 
                        DateTime Hoy = DateTime.Today;
                        L.FechaVersion = Hoy.ToString("dd/MM/yyyy");
                        
                        
                        L.ModificarN(L);

                        MessageBox.Show("Se ha modificado el legajo sin modificar la versión", "Modificar Legajo",
                        MessageBoxButtons.OK, MessageBoxIcon.None);
                    }


                    ////Cargo el legajo de version (el viejo)
                    //CargarDatosLegajoVer();
                    ////Lo guardo en legajoversion
                    //LegajoVer.Agregar();
                    //
                    //L.Eliminar(L.Codigo.ToString());
                    //L.Agregar();
                }
                else
                {
                    L.Agregar();
                    MessageBox.Show("Se ha agregado el legajo", "Agregar Legajo",
                        MessageBoxButtons.OK, MessageBoxIcon.None);
                }


                Close();

            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "ERROR");
            }
        }

        private void CargarDatosLegajoVer()
        {
            
            LegajoVer.Codigo = LegajoViejo.Codigo;
            
            LegajoVer.Version = LegajoViejo.Version;
            LegajoVer.Descripcion = LegajoViejo.Descripcion;
            LegajoVer.FIngreso = LegajoViejo.FIngreso;
            LegajoVer.Perfil = new Perfil();
            LegajoVer.Perfil = LegajoViejo.Perfil;
            LegajoVer.EstadoI = LegajoViejo.EstadoI;
            LegajoVer.EstadoII = LegajoViejo.EstadoII;
            LegajoVer.EstadoIII = LegajoViejo.EstadoIII;
            LegajoVer.EstadoIV = LegajoViejo.EstadoIV;
            LegajoVer.EstadoV = LegajoViejo.EstadoV;
            LegajoVer.EstadoVI = LegajoViejo.EstadoVI;
            LegajoVer.EstadoVII = LegajoViejo.EstadoVII;
            LegajoVer.EstadoVIII = LegajoViejo.EstadoVIII;
            LegajoVer.EstadoIX = LegajoViejo.EstadoIX;
            LegajoVer.EstadoX = LegajoViejo.EstadoX;

            LegajoVer.EstaI = LegajoViejo.EstaI;
            LegajoVer.EstaII = LegajoViejo.EstaII;
            LegajoVer.EstaIII = LegajoViejo.EstaIII;
            LegajoVer.EstaIV = LegajoViejo.EstaIV;
            LegajoVer.EstaV = LegajoViejo.EstaV;
            LegajoVer.EstaVI = LegajoViejo.EstaVI;
            LegajoVer.EstaVII = LegajoViejo.EstaVII;
            LegajoVer.EstaVIII = LegajoViejo.EstaVIII;
            LegajoVer.EstaIX = LegajoViejo.EstaIX;
            LegajoVer.EstaX = LegajoViejo.EstaX;

            LegajoVer.FechaVersionI = LegajoViejo.FechaVersion;
            LegajoVer.FechaVersionII = DTP_Fecha.Text;
            
            
            LegajoVer.FEgreso = LegajoViejo.FEgreso;
            LegajoVer.PerfilVersion = LegajoViejo.PerfilVersion;
            
            //LegajoVer.Temas = new Tema();
            LegajoVer.Temas = LegajoViejo.Temas;

            LegajoVer.DNI = LegajoViejo.DNI;

        }

        private void CargarDatosLegajo()
        {

            //L = new Legajo();
            L.Codigo = int.Parse(TB_Codigo.Text);
            L.Descripcion = TB_DescLegajo.Text;
            L.DNI = TB_DNI.Text;
            L.CUIL = TB_CUIL.Text;
            L.FIngreso = TB_FechaIng.Text;
            L.EstadoI = TB_ObservPrimariaLeg.Text;
            L.EstadoII = TB_ObservSecundariaLeg.Text;
            L.EstadoIII = TB_ObservTerciariaLeg.Text;
            L.EstadoIV = TB_ObservIdiomaLeg.Text;
            L.EstadoV = TB_ObservExpLeg.Text;
            L.EstadoVI = TB_ObservCondFisicaLeg.Text;
            L.EstadoVII = TB_Otros1Leg.Text;
            L.EstadoVIII = TB_Otros2Leg.Text;
            L.EstadoIX = TB_Equiv1Leg.Text;
            L.EstadoX = TB_Equiv2Leg.Text;



            L.EstaI = (CB_EstPrim.SelectedIndex).ToString();
            L.EstaII = (CB_EstSec.SelectedIndex).ToString();
            L.EstaIII = (CB_EstTerc.SelectedIndex).ToString();
            L.EstaIV = (CB_EstIdioma.SelectedIndex).ToString();
            L.EstaV = (CB_EstExp.SelectedIndex).ToString();
            L.EstaVI = (CB_EstCondFisic.SelectedIndex).ToString();
            L.EstaVII = (CB_EstOtros1.SelectedIndex).ToString();
            L.EstaVIII = (CB_EstOtros2.SelectedIndex).ToString();
            L.EstaIX = (CB_EstEquiv1.SelectedIndex).ToString();
            L.EstaX = (CB_Estequiv2.SelectedIndex).ToString();
            L.ObservExtI = ObservExtra1;
            L.ObservExtII = ObservExtra2;
            L.ObservExtIII = ObservExtra3;
            L.ObservExtIV = ObservExtra4;
            L.ObservExtV = ObservExtra5;

            List<Tema> Temas = new List<Tema>();
            foreach (DataGridViewRow row in DGV_Temas.Rows)
            {
                Tema T = new Tema
                {
                    Codigo = int.Parse(row.Cells[0].Value.ToString()),
                    Descripcion = row.Cells[1].Value.ToString(),
                    Necesaria = row.Cells[2].Value.ToString() == "X" ? 1 : 0,
                    Deseable = row.Cells[3].Value.ToString() == "X" ? 1 : 0
                };
                string EstaCurs = row.Cells[4].Value == null ? "" : row.Cells[4].Value.ToString();
                string result = BuscarEstado(EstaCurs);

                T.Estado = result;
                T.Observacion = row.Cells[5].Value;

                Temas.Add(T);
            }
            L.Temas = Temas;

            L.FechaVersion = DTP_Fecha.Text;
            L.Version = TB_Version.Text;

            L.Perfil = new Perfil
            {
                Codigo = int.Parse(TB_CodPerfil.Text),
                Descripcion = TB_DescPerfil.Text
            };
            //L.ImprePerfil = TB_DescPerfil.Text;

            L.FEgreso = TB_FechaEgreso.Text;
            //L.PerfilVersion = TB_VersPer.Text;
            L.Perfil.Version = int.Parse(TB_VersPer.Text);

            L.Sector = new Sector
            {
                Codigo = int.Parse(TB_CodSector.Text),
                Descripcion = TB_DescSec.Text
            };
        }

        private string BuscarEstado(string EstaCurs)
        {
            switch (EstaCurs)
            {
                case "":
                    return "0";

                case "Exced":
                    return "1";

                case "Cumple":
                    return "2";

                case "Reforzar":
                    return "3";

                case "En Entren":
                    return "4";

                case "No Cumple":
                    return "5";

                case "No Aplica":
                    return "6";

                case "No Evalua":
                    return "7";

                case "Cumple Act":
                    return "8";

                default:
                    return "";
            }
        }

        private void ValidarDatosCargados()
        {
            if (TB_DescLegajo.Text == "")
                throw new Exception("Se debe cargar la descripcion");
            if (TB_DescPerfil.Text == "" || TB_DescSec.Text == "")
                throw new Exception("Se deben completar los datos de perfil, sector");
            //if (DGV_Temas.Rows.Count == 0)
            //    throw new Exception("Se deben cargar temas");
        }

        private void AgModLegajo_Load(object sender, EventArgs e)
        {
            //CargarSectores();
            //CargarPerfil();
            //Cargado = true;
        }

        

        private void CargarSectores()
        {
            dtSectores = S.ListarTodos();

            var fila = dtSectores.NewRow();
            dtSectores.Rows.InsertAt(fila, 0);


            CargarDescSector();
            CargarCodigosSec();
        }

        private void CargarDescSector()
        {
            TB_DescSec.DataSource = dtSectores;
            TB_DescSec.DisplayMember = "Descripcion";
            TB_DescSec.ValueMember = "Codigo";
            TB_DescSec.Text = "";

            AutoCompleteStringCollection stringCodArti = new AutoCompleteStringCollection();
            foreach (DataRow row in dtSectores.Rows)
            {
                stringCodArti.Add(Convert.ToString(row["Descripcion"]));

            }

            TB_DescSec.AutoCompleteCustomSource = stringCodArti;
            TB_DescSec.AutoCompleteMode = AutoCompleteMode.Suggest;
            TB_DescSec.AutoCompleteSource = AutoCompleteSource.CustomSource;
        }

        private void CargarCodigosSec()
        {
            TB_CodSector.DataSource = dtSectores;
            TB_CodSector.DisplayMember = "Codigo";
            TB_CodSector.ValueMember = "Codigo";
            TB_CodSector.Text = "";

            AutoCompleteStringCollection stringCodArti = new AutoCompleteStringCollection();
            foreach (DataRow row in dtSectores.Rows)
            {
                stringCodArti.Add(Convert.ToString(row["Codigo"]));

            }

            TB_CodSector.AutoCompleteCustomSource = stringCodArti;
            TB_CodSector.AutoCompleteMode = AutoCompleteMode.Suggest;
            TB_CodSector.AutoCompleteSource = AutoCompleteSource.CustomSource;
        }


        private void CargarPerfil()
        {
            dtPerfil = Per.ListarTodos();

            var fila = dtPerfil.NewRow();
            dtPerfil.Rows.InsertAt(fila, 0);


            CargarDescPerfil();
            CargarCodigosPerfil();
        }

        private void CargarDescPerfil()
        {
            TB_DescPerfil.DataSource = dtPerfil;
            TB_DescPerfil.DisplayMember = "Descripcion";
            TB_DescPerfil.ValueMember = "Codigo";
            TB_DescPerfil.Text = "";

            AutoCompleteStringCollection stringCodArti = new AutoCompleteStringCollection();
            foreach (DataRow row in dtPerfil.Rows)
            {
                stringCodArti.Add(Convert.ToString(row["Descripcion"]));

            }

            TB_DescPerfil.AutoCompleteCustomSource = stringCodArti;
            TB_DescPerfil.AutoCompleteMode = AutoCompleteMode.Suggest;
            TB_DescPerfil.AutoCompleteSource = AutoCompleteSource.CustomSource;
        }

        private void CargarCodigosPerfil()
        {
            TB_CodPerfil.DataSource = dtPerfil;
            TB_CodPerfil.DisplayMember = "Codigo";
            TB_CodPerfil.ValueMember = "Codigo";
            TB_CodPerfil.Text = "";

            AutoCompleteStringCollection stringCodArti = new AutoCompleteStringCollection();
            foreach (DataRow row in dtPerfil.Rows)
            {
                stringCodArti.Add(Convert.ToString(row["Codigo"]));

            }

            TB_CodPerfil.AutoCompleteCustomSource = stringCodArti;
            TB_CodPerfil.AutoCompleteMode = AutoCompleteMode.Suggest;
            TB_CodPerfil.AutoCompleteSource = AutoCompleteSource.CustomSource;
        }

        private void TB_CodPerfil_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Cargado)
            {
                if (AModificar == false)
                {
                    TB_DescPerfil.SelectedValue = TB_CodPerfil.SelectedValue;
                    //BuscarCodperfil();
                    LlenarPerfil();
                    TB_CodSector.Focus();
                }
            }
            
        }

        private void BuscarCodperfil()
        {
            foreach (DataRow fila in dtPerfil.Rows)
            {
                if (fila[1].ToString().Trim() == TB_DescPerfil.Text.Trim())
                {
                    TB_CodPerfil.Enabled = true;
                    TB_DescPerfil.Enabled = true;
                    if (fila[0].ToString() != "")
                    {
                        TB_CodPerfil.SelectedValue = fila[0].ToString();
                        TB_DescPerfil.SelectedValue = TB_CodPerfil.SelectedValue ?? 0;
                    }
                    
                    break;
                }
            }
        }

       

        private void TB_CodPerfil_KeyDown_1(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (TB_CodPerfil.Text == "")
                {
                    TB_DescPerfil.Focus();
                }
                else
                {
                    LlenarPerfil();
                    TB_CodSector.Focus();
                }
            }
            
        }

        private void TB_DescPerfil_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (TB_CodPerfil.Text == "")
                {
                    TB_DescPerfil.Focus();
                }
                else
                {
                    LlenarPerfil();
                    TB_CodSector.Focus();
                }
                
            }
            
        }

        private void TB_DescPerfil_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Cargado)
            {
                if(AModificar == false)
                {
                     LlenarPerfil();
                TB_CodSector.Focus();
                }
               
            }
            
        }

        private void LlenarPerfil()
        {
            try
            {
                CargarDatosPefil();
                CargarTemas(Per);
                if (Per.Codigo == 0) return;
                CargarCursadas();
                //e.SuppressKeyPress = true;
                //e.Handled = true;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Error");
            }
        }

        private void TB_DescLegajo_KeyDown(object sender, KeyEventArgs e)
        {
            
            if (e.KeyData == Keys.Enter)
            {
                if (TB_DescLegajo.Text.Trim() == "") return;

                TB_FechaIng.Focus();

            }
            else if (e.KeyData == Keys.Escape)
            {
                TB_DescLegajo.Text = "";
            }
	        
            
        }

        private void BT_MasObserv_Click(object sender, EventArgs e)
        {
            Observaciones Ini = new Observaciones(L.ObservExtI, L.ObservExtII, L.ObservExtIII, L.ObservExtIII, L.ObservExtIV, L.ObservExtV);
            Ini.ShowDialog();
            ObservExtra1 = Ini.ObservExt1;
            ObservExtra2 = Ini.ObservExt2;
            ObservExtra3 = Ini.ObservExt3;
            ObservExtra4 = Ini.ObservExt4;
            ObservExtra5 = Ini.ObservExt5;

            L.ObservExtI = Ini.ObservExt1;
            L.ObservExtII = Ini.ObservExt2;
            L.ObservExtIII = Ini.ObservExt3;
            L.ObservExtIV = Ini.ObservExt4;
            L.ObservExtV = Ini.ObservExt5;
        }

        public object Estacurso { get; set; }

        private void AgModLegajo_Shown(object sender, EventArgs e)
        {
            TB_DescLegajo.Focus();
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            BT_MasObserv.PerformClick();
        }

        private void TB_FechaEgreso_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void TB_FechaIng_KeyDown(object sender, KeyEventArgs e)
        {
            
            if (e.KeyData == Keys.Enter)
            {
                if (TB_FechaIng.Text.Replace("/", "").Trim() == "") return;

                if (!Helper._FechaValida(TB_FechaIng.Text)) return;

                TB_DNI.Focus();

            }
            else if (e.KeyData == Keys.Escape)
            {
                TB_FechaIng.Text = "";
            }
	        
        }

        private void TB_DNI_KeyDown(object sender, KeyEventArgs e)
        {
            
            if (e.KeyData == Keys.Enter)
            {
                //if (TB_DNI.Text.Trim() == "") return;

                TB_DescPerfil.Focus();

            }
            else if (e.KeyData == Keys.Escape)
            {
                TB_DNI.Text = "";
            }
	        
        }

        private void TB_Codigo_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyData == Keys.Enter)
            {
                if (TB_Codigo.Text.Trim() == "") return;

                Legajo LegajoAModificar = new Legajo();
                LegajoAModificar = L.BuscarUno(TB_Codigo.Text);

                _CargarLegajo(LegajoAModificar);

                TB_DescLegajo.Focus();

            }
            else if (e.KeyData == Keys.Escape)
            {
                TB_Codigo.Text = "";
            }
	        
        }
    }
}
