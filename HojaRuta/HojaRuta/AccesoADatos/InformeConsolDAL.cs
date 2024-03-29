﻿using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace AccesoADatos
{
    public class InformeConsolDAL
    {

        public DataTable Lista()
        {
            DataTable dtInforme = new DataTable();
            
            using (SqlConnection cnx = new SqlConnection(ConfigurationManager.ConnectionStrings["SurfactanSA"].ConnectionString))
            {

                cnx.Open();
                string sqlQuery = "select * from InformeConsol order by Informe asc";
                using (SqlCommand cmd = new SqlCommand(sqlQuery, cnx))
                {
                    SqlDataReader dataReader = cmd.ExecuteReader();

                    dtInforme.Load(dataReader);

                }
                return dtInforme;
            }
        }

        private void _ConsolidarInformes(string WDesde, string WHasta)
        {

            try
            {
                DataTable tabla = new DataTable();

                using (SqlConnection conn = new SqlConnection())
                {  
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        foreach (string WEmpresa in new[] { "SurfactanSa", "surfactan_II", "surfactan_III", "surfactan_IV", "surfactan_V", "surfactan_VI", "surfactan_VII" })
                        {
                            if (conn.State == ConnectionState.Open) conn.Close();

                            conn.ConnectionString = ConfigurationManager.ConnectionStrings[WEmpresa].ConnectionString;
                            conn.Open();

                            cmd.Connection = conn;
                            cmd.CommandText = "SELECT Informe, Fecha, FechaOrd, Expreso, Chapa, Chofer, Item1, Item2, Item3, Item4, Item5, Item6, Item7, Item8, Placa, Rombo, ISNULL(Observaciones, '') Observaciones, ISNULL(DesExpreso, '') DesExpreso "
                                                + " From Informe "
                                                + " WHERE Expreso > 0 AND Renglon = 1 AND FechaOrd BETWEEN '" + WDesde + "' AND '" + WHasta + "'";

                            using (SqlDataReader dr = cmd.ExecuteReader())
                            {
                                if (dr.HasRows)
                                {
                                    tabla.Load(dr);
                                }
                            }
                        }

                        if (conn.State == ConnectionState.Open) conn.Close();

                        conn.ConnectionString = ConfigurationManager.ConnectionStrings["SurfactanSa"].ConnectionString;
                        conn.Open();

                        cmd.CommandText = "DELETE FROM InformeConsol";
                        cmd.ExecuteNonQuery();

                        foreach (DataRow row in tabla.Rows)
                        {
                            string WInforme = row["Informe"].ToString();
                            string WFecha = row["Fecha"].ToString();
                            string WFechaOrd = row["FechaOrd"].ToString();
                            string WExpreso = row["Expreso"].ToString();
                            string WChapa = row["Chapa"].ToString();
                            string WChofer = row["Chofer"].ToString();
                            string WItem1 = row["Item1"].ToString();
                            string WItem2 = row["Item2"].ToString();
                            string WItem3 = row["Item3"].ToString();
                            string WItem4 = row["Item4"].ToString();
                            string WItem5 = row["Item5"].ToString();
                            string WItem6 = row["Item6"].ToString();
                            string WItem7 = row["Item7"].ToString();
                            string WItem8 = row["Item8"].ToString();
                            string WPlaca = row["Placa"].ToString();
                            string WRombo = row["Rombo"].ToString();
                            string WObservaciones = row["Observaciones"].ToString();
                            string WDesExpreso = row["DesExpreso"].ToString();

                            cmd.CommandText = "INSERT INTO InformeConsol " +
                                              " (Informe, Fecha, OrdFecha, Expreso, Chapa, Chofer, Item1, Item2, Item3, Item4, Item5, Item6, Item7, Item8, " +
                                              " Placa, Rombo, Observaciones, DesExpreso) VALUES " +
                                              " ('" + WInforme + "', '" + WFecha + "', '" + WFechaOrd + "', '" + WExpreso + "', '" + WChapa + "'," +
                                              " '" + WChofer + "', '" + WItem1 + "', '" + WItem2 + "', '" + WItem3 + "', '" + WItem4 + "', '" + WItem5 + "', '" + WItem6 + "', '" + WItem7 + "', '" + WItem8 + "'," +
                                              " '" + WPlaca + "', '" + WRombo + "', '" + WObservaciones + "', '" + WDesExpreso + "')";
                            cmd.ExecuteNonQuery();
                        }
                    }
                }

                
            }
            catch (Exception ex)
            {
                throw new Exception("Error al Consolidar los Informes. Motivo: " + ex.Message);
            }
        
        }

        public DataTable Lista(string wDesde, string wHasta)
        {
            _ConsolidarInformes(wDesde, wHasta);

            return Lista();
        }
    }
}
