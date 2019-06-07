using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClassConexion;

namespace Negocio
{
    public class LegajoVersion: IRepoMetodos<LegajoVersion>
    {
        public int Codigo;
        public string Descripcion;
        public string FIngreso;
        public string EstadoI;
        public string EstadoII;
        public string EstadoIII;
        public string EstadoX;
        public string EstadoIV;
        public string EstadoV;
        public string EstadoVI;
        public string EstadoVII;
        public string EstadoVIII;
        public string EstadoIX;
        public string FechaVersionI;
        public string FechaVersionII;
        public string Version;
        public Perfil Perfil;
        public string FEgreso;
        public string PerfilVersion;
        //public Sector Sector;
        public string EstaI;
        public string EstaII;
        public string EstaIII;
        public string EstaIV;
        public string EstaV;
        public string EstaVI;
        public string EstaVII;
        public string EstaVIII;
        public string EstaIX;
        public string EstaX;
        public string DNI;
        public List<Tema> Temas { get; set; }



        public System.Data.DataTable ListarTodos()
        {
            throw new NotImplementedException();
        }

        public int ObtenerUltimoId()
        {
            throw new NotImplementedException();
        }

        public void Agregar()
        {
            Conexion Repo = new Conexion();

            int renglon = 1;

            foreach (var t in this.Temas)
            {
                var clave1 = this.Codigo.ToString().PadLeft(6, '0');
                var clave2 = this.Version.ToString().PadLeft(4, '0');
                var clave3 = renglon.ToString().PadLeft(2, '0');

                var claper1 = Perfil.Codigo.ToString().PadLeft(6, '0');
                var claper2 = "1".ToString().PadLeft(2, '0');

                string nece = t.Necesaria == 1 ? "X" : "";
                string dese = t.Deseable == 1 ? "X" : "";

                //int estado = ObtenerEstado(t.Estado);


                string consulta = "insert into LegajoVersion "
                + " (Clave, Codigo,Renglon, Descripcion,FIngreso,EstadoI,"
                + " EstadoII,EstadoIII,EstadoIV,EstadoV,EstadoVI,EstadoVII,"
                + " EstadoVIII,EstadoIX,EstadoX, Curso, EstadoCurso, EstaI,"
                + " EstaII,EstaIII,EstaIV,EstaV,EstaVI,EstaVII,EstaVIII,EstaIX,EstaX,"
                + " EstaCurso,Perfil,ClavePerfil, NecesariaCurso,DeseableCurso,FechaVersionI, FechaVersionII,"
                + " Version, Fegreso, PerfilVersion)"
                + " values " +
                "('" + clave1 + clave2 + clave3 + "'," + Codigo + "," + renglon + ",'" + Descripcion + "','" + FIngreso + "','" + EstadoI + "',"
                + " '" + EstadoII + "','" + EstadoIII + "','" + EstadoIV + "','" + EstadoV + "','" + EstadoVI + "','" + EstadoVII + "',"
                + " '" + EstadoVIII + "','" + EstadoIX + "','" + EstadoX + "'," + t.Codigo + ",'" + t.Observacion + "', '" + ObtenerValor(EstaI) + "',"
                + " '" + ObtenerValor(EstaII) + "','" + ObtenerValor(EstaIII) + "','" + ObtenerValor(EstaIV) + "','" + ObtenerValor(EstaV) + "','"
                + ObtenerValor(EstaVI) + "','" + ObtenerValor(EstaVII) + "','" + ObtenerValor(EstaVIII) + "','" + ObtenerValor(EstaIX) + "','" + ObtenerValor(EstaX) + "',"
                + " " + t.EstaCurso + "," + Perfil.Codigo + ",'" + claper1 + claper2 + "','" + nece + "','" + dese + "','" + FechaVersionI + "','" + FechaVersionII + "',"
                + " " + Version + ",'" + FEgreso + "'," + Perfil.Version + ")";

                Repo.Agregar(consulta);

                renglon++;
            }
        }

        private object ObtenerValor(string Esta)
        {
            switch (Esta)
            {
                case "Cumple":
                    return "0";
                case "Cumple Act.":
                    return "1";
                case "Reforzar":
                    return "2";
                case "No Cumple":
                    return "3";
                default:
                    return "";
            }
        }

        public LegajoVersion BuscarUno(string IdAModificar)
        {
            throw new NotImplementedException();  
        }

        public void Modificar(LegajoVersion t)
        {
            throw new NotImplementedException();
        }

        public void Eliminar(string IdAEliminar)
        {
            throw new NotImplementedException();
        }

        public LegajoVersion BuscarUnaVersion(string cod,string ver,string fecha)
        {
            Conexion repo = new Conexion();

            string consulta = "select lv.*, l.FIngreso as FechaIngreso from legajoversion lv INNER JOIN Legajo l ON l.Codigo = lv.Codigo And l.Renglon = 1 where lv.codigo = " + cod + " and lv.version = " + ver;
            //" and convert(datetime,FIngreso,103) >= convert(datetime,'" + fecha + "',103)";

            System.Data.DataTable DT = repo.BuscarUno(consulta);

            LegajoVersion obj = new LegajoVersion();

            if (DT.Rows.Count > 0)
            {
                obj.Codigo = int.Parse(DT.Rows[0]["Codigo"].ToString());

                obj.FechaVersionI = DT.Rows[0]["FechaVersionI"].ToString();
                obj.FechaVersionII = DT.Rows[0]["FechaVersionII"].ToString();
                obj.Descripcion = DT.Rows[0]["Descripcion"].ToString();
                obj.Version = DT.Rows[0]["Version"].ToString();
                obj.FIngreso = DT.Rows[0]["FechaIngreso"].ToString();
                obj.FEgreso = DT.Rows[0]["Fegreso"].ToString();

                //DateTime FEgresoParse;
                //DateTime.TryParse(L.FEgreso, out FEgresoParse);

                //if (FEgresoParse.ToString() == "01/01/0001 0:00:00")
                //{
                //    TB_FechaEgreso.Text = "01/01/1890 12:00:00 AM";
                //}
                //else { TB_FechaEgreso.Text = FEgresoParse.ToShortDateString(); }

                obj.EstadoI = DT.Rows[0]["EstadoI"].ToString();
                obj.EstadoII = DT.Rows[0]["EstadoII"].ToString();
                obj.EstadoIII = DT.Rows[0]["EstadoIII"].ToString();
                obj.EstadoIV = DT.Rows[0]["EstadoIV"].ToString();
                obj.EstadoV = DT.Rows[0]["EstadoV"].ToString();
                obj.EstadoVI = DT.Rows[0]["EstadoVI"].ToString();
                obj.EstadoVII = DT.Rows[0]["EstadoVII"].ToString();
                obj.EstadoVIII = DT.Rows[0]["EstadoVIII"].ToString();
                obj.EstadoIX = DT.Rows[0]["EstadoIX"].ToString();
                obj.EstadoX = DT.Rows[0]["EstadoX"].ToString();
                obj.EstaI = DT.Rows[0]["EstaI"].ToString();
                obj.EstaII = DT.Rows[0]["EstaII"].ToString();
                obj.EstaIII = DT.Rows[0]["EstaIII"].ToString();
                obj.EstaIV = DT.Rows[0]["EstaIV"].ToString();
                obj.EstaV = DT.Rows[0]["EstaV"].ToString();
                obj.EstaVI = DT.Rows[0]["EstaVI"].ToString();
                obj.EstaVII = DT.Rows[0]["EstaVII"].ToString();
                obj.EstaVIII = DT.Rows[0]["EstaVIII"].ToString();
                obj.EstaIX = DT.Rows[0]["EstaIX"].ToString();
                obj.EstaX = DT.Rows[0]["EstaX"].ToString();
                obj.Perfil = new Perfil();
                obj.Perfil.Codigo = int.Parse((DT.Rows[0]["Perfil"].ToString()));
                //obj.Perfil.Descripcion = DT.Rows[0]["ImprePerfil"].ToString();
                obj.Perfil.Version = int.Parse(DT.Rows[0]["PerfilVersion"].ToString());

                //Sector S = new Sector();
                //S.Codigo = int.Parse(DT.Rows[0]["Sector"].ToString());
                //S = S.BuscarUno(S.Codigo.ToString());
                //obj.Sector = new Sector();
                //obj.Sector = S;

                obj.Temas = new List<Tema>();
                int renglon = 1;
                foreach (System.Data.DataRow item in DT.Rows)
                {
                    Tema T = new Tema();
                    T.Codigo = int.Parse(item["Curso"].ToString());
                    T = T.BuscarUno_Tema(T.Codigo.ToString());
                    T.Necesaria = item["NecesariaCurso"].ToString() == "X" ? 1 : 0;
                    T.Deseable = item["DeseableCurso"].ToString() == "X" ? 1 : 0;
                    T.EstadoCurso = item["EstadoCurso"].ToString();
                    T.EstaCurso = item["EstaCurso"].ToString() == "1" ? 1 : 0;
                    obj.Temas.Add(T);
                    renglon++;
                }

            }
            return obj;
        }

        public string MaxVersiones(string codigo)
        {
            Conexion repo = new Conexion();
            string consulta = "select max(version) from legajoversion where codigo = " + codigo;
            return repo.TraerVersionMax(consulta);
        }
    }
}
