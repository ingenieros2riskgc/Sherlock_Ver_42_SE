using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.OleDb;
using ListasSarlaft.Classes;
using System.Data;
using System.Data.SqlClient;

namespace ListasSarlaft.Classes
{ 
    public class cEvento : cPropiedades 
    {
        private cDataBase cDataBase = new cDataBase();
        private cError cError = new cError();
        private Object thisLock = new Object();
        private OleDbConnection oleDbCnn;
        private SqlConnection sqlCnn;
        //private OleDbParameter[] parameters;
        //private OleDbParameter parameter;

        public void agregarArchivoPlanAccion(String IdRegistro, String UrlArchivo)
        {
            try
            {
                cDataBase.conectar();
                cDataBase.ejecutarQuery("INSERT INTO Riesgos.Archivos (IdControlUsuario, IdRegistro, NombreUsuario, FechaRegistro, UrlArchivo) VALUES (7, " + IdRegistro + ", '" + NombreUsuario + "', GETDATE(), N'" + UrlArchivo + "')");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public void agregarComentarioPlanAccion(String Comentario, String IdRegistro)
        {
            try
            {
                cDataBase.conectar();
                cDataBase.ejecutarQuery("INSERT INTO Riesgos.Comentarios (IdControlUsuario, IdRegistro, NombreUsuario, FechaRegistro, Comentario) VALUES (7, " + IdRegistro + ", '" + NombreUsuario + "', GETDATE(), N'" + Comentario + "')");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public void actualizarPlanAccionEvento(String IdPlanAccion, String DescripcionAccion, String Responsable, String IdTipoRecursoPlanAccion, String ValorRecurso, String IdEstadoPlanAccion, String FechaCompromiso)
        {
            try
            {
                cDataBase.conectar();

                cDataBase.ejecutarQuery("UPDATE Riesgos.PlanesAccion SET DescripcionAccion = N'" + DescripcionAccion +
                    "', Responsable = " + Responsable + ", IdTipoRecursoPlanAccion = " + IdTipoRecursoPlanAccion +
                    ", ValorRecurso = N'" + ValorRecurso + "', IdEstadoPlanAccion = " + IdEstadoPlanAccion +
                    ", FechaCompromiso = CONVERT(datetime, '" + FechaCompromiso + "', 120) WHERE (IdPlanAccion = " + IdPlanAccion + ")");

                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        //public void registrarPlanAccionEvento(String IdRegistro, String DescripcionAccion, 
        //    String Responsable, String IdTipoRecursoPlanAccion, String ValorRecurso, 
        //    String IdEstadoPlanAccion, String FechaCompromiso)
        //{
        //    string strConsulta = string.Empty;

        //    try
        //    {
        //        strConsulta = string.Format("INSERT INTO Riesgos.PlanesAccion (IdControlUsuario, IdRegistro, DescripcionAccion, Responsable, " + 
        //            "IdTipoRecursoPlanAccion, ValorRecurso, IdEstadoPlanAccion, FechaCompromiso) " + 
        //            "VALUES (6, {0}, N'{1}', {2}, {3}, N'{4}', CONVERT(datetime, '{5}', 120)) ", 
        //            IdRegistro, DescripcionAccion, Responsable,
        //            IdTipoRecursoPlanAccion, ValorRecurso, IdEstadoPlanAccion, FechaCompromiso);

        //        cDataBase.conectar();
        //        cDataBase.ejecutarQuery(strConsulta);
        //        cDataBase.desconectar();
        //    }
        //    catch (Exception ex)
        //    {
        //        cDataBase.desconectar();
        //        cError.errorMessage(ex.Message + ", " + ex.StackTrace);
        //        throw new Exception(ex.Message);
        //    }
        //}

        public DataTable mtdRegistrarPlanAccionEvento(string IdRegistro, string DescripcionAccion,
           string Responsable, string IdTipoRecursoPlanAccion, string ValorRecurso,
           string IdEstadoPlanAccion, string FechaCompromiso)
        {
            string strConsulta = string.Empty;
            DataTable dtInfo = new DataTable();

            try
            {
                strConsulta = string.Format("INSERT INTO Riesgos.PlanesAccion (IdControlUsuario, IdRegistro, DescripcionAccion, Responsable, " +
                    "IdTipoRecursoPlanAccion, ValorRecurso, IdEstadoPlanAccion, FechaCompromiso) " +
                    "VALUES (6, {0}, N'{1}', {2}, {3}, N'{4}', {5}, CONVERT(datetime, '{6}', 120)) SELECT SCOPE_IDENTITY()",
                    IdRegistro, DescripcionAccion, Responsable,
                    IdTipoRecursoPlanAccion, ValorRecurso, IdEstadoPlanAccion, FechaCompromiso);

                cDataBase.conectar();
                dtInfo = cDataBase.mtdEjecutarConsultaSQL(strConsulta);
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }

            return dtInfo;
        }
        #region Validacion Total Riesgos vs Eventos
        public DataTable TotalRiesgosxEvento(int IdRiesgo)
        {
            string strConsulta = string.Empty;
            DataTable dtInfo = new DataTable();

            try
            {
                strConsulta = string.Format("SELECT COUNT(IdRiesgo) as CantRiesgo FROM [Riesgos].[EventoRiesgo] where IdRiesgo = {0}",
                    IdRiesgo);

                cDataBase.conectar();
                dtInfo = cDataBase.mtdEjecutarConsultaSQL(strConsulta);
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }

            return dtInfo;
        }

        //yoendy
        public DataTable mtdTotalRiesgosxEventoFrecuencia(int IdRiesgo)
        {
            string strConsulta = string.Empty;
            DataTable dtInfo = new DataTable();

            try
            {
                strConsulta = string.Format("SELECT COUNT(IdRiesgo) as CantRiesgo FROM [Riesgos].[ParamFrecEventos] where IdRiesgo = {0}",
                    IdRiesgo);

                cDataBase.conectar();
                dtInfo = cDataBase.mtdEjecutarConsultaSQL(strConsulta);
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }

            return dtInfo;
        }

        public DataTable LimiteRiesgosxEvento(int IdFrecuencia)
        {
            string strConsulta = string.Empty;
            DataTable dtInfo = new DataTable();

            try
            {
                strConsulta = string.Format("SELECT [EventosMaximos] FROM [Parametrizacion].[FrecuenciavsEventos] where CodigoFrecuencia = {0}",
                    IdFrecuencia);

                cDataBase.conectar();
                dtInfo = cDataBase.mtdEjecutarConsultaSQL(strConsulta);
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }

            return dtInfo;
        }

        //yoendy
        public DataTable MtdLimiteRiesgosxEventoNew(int idRiesgo)
        {
            string strConsulta = string.Empty;
            DataTable dtInfo = new DataTable();

            try
            {
                strConsulta = string.Format("SELECT MAX(EventosMaximos) as cantFrecuenciaNew FROM [Riesgos].[ParamFrecEventos] where idRiesgo = {0}",
                    idRiesgo);

                cDataBase.conectar();
                dtInfo = cDataBase.mtdEjecutarConsultaSQL(strConsulta);
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }

            return dtInfo;
        }

        //yoendy
        public DataTable mtdTotalRiesgosFrecuencia(int IdRiesgo, string NombreFrecuencia)
        {
            string strConsulta = string.Empty;
            DataTable dtInfo = new DataTable();

            try
            {
                strConsulta = string.Format("SELECT COUNT(IdRiesgo) as ExisteFreq FROM [Riesgos].[ParamFrecEventos] where (IdRiesgo = " + IdRiesgo + " and NombreFrecuencia= '" + NombreFrecuencia + "')");

                cDataBase.conectar();
                dtInfo = cDataBase.mtdEjecutarConsultaSQL(strConsulta);
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }

            return dtInfo;
        }
        public void actualizarFrecuenciaRiesgo(String IdRiesgo, String IdFrecuencia, String IdImpacto)
        {
            string query = string.Empty;
            query = string.Format("UPDATE [Riesgos].[Riesgo] SET [IdProbabilidad] = {1},[IdImpacto] = {2}"
                    + " WHERE (IdRiesgo = {0})", IdRiesgo, IdFrecuencia, IdImpacto);
            try
            {
                cDataBase.conectar();

                cDataBase.ejecutarQuery(query);

                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public int actualizarFrecuenciaMax(int idRiesgo, int codFrecuencia, int eventosMaximos, string nombreFrecuencia, int idUsuario)
        {
            int resultado = 0;
            string query = string.Empty;
            //query = string.Format("UPDATE [Riesgos].[Riesgo] SET [IdProbabilidad] = {1},[IdImpacto] = {2}"
            //        + " WHERE (IdRiesgo = {0})", IdRiesgo, IdFrecuencia, IdImpacto);
            try
            {
                cDataBase.conectar();

                cDataBase.ejecutarQuery(query);

                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return resultado;
        }

        public int EjecutarSPParametrosReturnInteger(string NombreSp, List<SqlParameter> Parametros)
        {
            SqlConnection cnn = sqlCnn;
            try
            {
                cnn.Open();
                SqlCommand sqlCmd = new SqlCommand(NombreSp, cnn);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.Parameters.AddRange(Parametros.ToArray());
                sqlCmd.CommandTimeout = 3600;
                sqlCmd.ExecuteNonQuery();
                int resultado = (int)sqlCmd.Parameters["@Resultado"].Value;

                return resultado;
            }
            catch (Exception ex)
            {//capturar error                            
                cError ce = new cError();
                ce.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            finally
            {
                cnn.Close();
            }
        }

        #endregion Validacion Total Riesgos vs Eventos
        public DataTable loadInfoArchivoPlanAccion(String IdRegistro) 
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT Riesgos.Archivos.IdArchivo, Riesgos.Archivos.NombreUsuario, Riesgos.Archivos.FechaRegistro, Riesgos.Archivos.UrlArchivo FROM Riesgos.Archivos INNER JOIN Parametrizacion.ControlesUsuario ON Riesgos.Archivos.IdControlUsuario = Parametrizacion.ControlesUsuario.IdControlUsuario WHERE (Riesgos.Archivos.IdRegistro = " + IdRegistro + ") AND (Parametrizacion.ControlesUsuario.IdControlUsuario = 7)");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public DataTable ConsultaPlanesAsociados(string CodigoEvento)
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT pea.CodigoPlan, \n"
           + "       pea.CodigoEvento, \n"
           + "       pea.FechaRegistro, \n"
           + "       pea.Usuario\n"
           + "       FROM riesgos.PlanesEventosAsociados pea\n"
           + "       WHERE CodigoEvento = ('" + CodigoEvento + "') ");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public DataTable loadInfoComentarioPlanAccion(String IdRegistro)
        {
            DataTable dtInformacion = new DataTable(); 
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT Riesgos.Comentarios.IdComentario, Riesgos.Comentarios.NombreUsuario, Riesgos.Comentarios.FechaRegistro, LTRIM(RTRIM(SUBSTRING(Riesgos.Comentarios.Comentario, 1, 20))) + '...' AS ComentarioCorto, Riesgos.Comentarios.Comentario FROM Riesgos.Comentarios INNER JOIN Parametrizacion.ControlesUsuario ON Riesgos.Comentarios.IdControlUsuario = Parametrizacion.ControlesUsuario.IdControlUsuario WHERE (Parametrizacion.ControlesUsuario.IdControlUsuario = 7) AND (Riesgos.Comentarios.IdRegistro = " + IdRegistro + ")");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public DataTable loadInfoPlanAccionEvento(String IdRegistro)
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT Riesgos.PlanesAccion.IdPlanAccion, Riesgos.PlanesAccion.DescripcionAccion, Riesgos.PlanesAccion.Responsable, Riesgos.PlanesAccion.IdTipoRecursoPlanAccion, Riesgos.PlanesAccion.ValorRecurso, Parametrizacion.EstadoPlanAccion.IdEstadoPlanAccion, Parametrizacion.EstadoPlanAccion.NombreEstadoPlanAccion, REPLACE(CONVERT(varchar, Riesgos.PlanesAccion.FechaCompromiso, 102), '.', '-') AS FechaCompromiso, Parametrizacion.JerarquiaOrganizacional.NombreHijo FROM Riesgos.PlanesAccion INNER JOIN Parametrizacion.EstadoPlanAccion ON Riesgos.PlanesAccion.IdEstadoPlanAccion = Parametrizacion.EstadoPlanAccion.IdEstadoPlanAccion INNER JOIN Parametrizacion.JerarquiaOrganizacional ON Riesgos.PlanesAccion.Responsable = Parametrizacion.JerarquiaOrganizacional.idHijo WHERE (Riesgos.PlanesAccion.IdControlUsuario = 6) AND (Riesgos.PlanesAccion.IdRegistro = " + IdRegistro + ")");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public DataTable loadInfoArchivoEvento(String IdRegistro)
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT Riesgos.Archivos.IdArchivo, Riesgos.Archivos.NombreUsuario, Riesgos.Archivos.FechaRegistro, Riesgos.Archivos.UrlArchivo FROM Riesgos.Archivos INNER JOIN Parametrizacion.ControlesUsuario ON Riesgos.Archivos.IdControlUsuario = Parametrizacion.ControlesUsuario.IdControlUsuario WHERE (Riesgos.Archivos.IdRegistro = " + IdRegistro + ") AND (Parametrizacion.ControlesUsuario.IdControlUsuario = 6)");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public void agregarComentarioEvento(String Comentario, String IdRegistro)
        {
            try
            {
                cDataBase.conectar();
                cDataBase.ejecutarQuery("INSERT INTO Riesgos.Comentarios (IdControlUsuario, IdRegistro, NombreUsuario, FechaRegistro, Comentario) VALUES (6, " + IdRegistro + ", '" + NombreUsuario + "', GETDATE(), N'" + Comentario + "')");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public DataTable BuscarIdEvento(String CodEvento)
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT MAX(IdEvento) AS IdEvento FROM riesgos.eventos WHERE Riesgos.eventos.CodigoEvento = ('" + CodEvento+"') ");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public DataTable loadInfoComentarioEvento(String IdRegistro)
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT Riesgos.Comentarios.IdComentario, Riesgos.Comentarios.NombreUsuario, Riesgos.Comentarios.FechaRegistro, LTRIM(RTRIM(SUBSTRING(Riesgos.Comentarios.Comentario, 1, 20))) + '...' AS ComentarioCorto, Riesgos.Comentarios.Comentario FROM Riesgos.Comentarios INNER JOIN Parametrizacion.ControlesUsuario ON Riesgos.Comentarios.IdControlUsuario = Parametrizacion.ControlesUsuario.IdControlUsuario WHERE (Parametrizacion.ControlesUsuario.IdControlUsuario = 6) AND (Riesgos.Comentarios.IdRegistro = " + IdRegistro + ") ORDER BY idComentario DESC ");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public DataTable IdNotificacion()
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT TOP 1 IdEvento  from  notificaciones.Evento WHERE nombreEvento = 'Eventos (Máximos Superados)'");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public void relacionarRiesgoEvento(String IdRiesgo, String IdEvento)
        {
            try
            {
                cDataBase.conectar();
                cDataBase.ejecutarQuery("INSERT INTO Riesgos.EventoRiesgo (IdRiesgo, IdEvento, IdUsuario, FechaRegistro) VALUES (" + IdRiesgo + ", " + IdEvento + ", " + IdUsuario + ", GETDATE())");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public void RelacionarRiesgoEventoCausas(int IdRiesgo, int IdEvento, int IdCausa)
        {
            try
            {
                cDataBase.conectar();
                cDataBase.ejecutarQuery("INSERT INTO [Riesgos].[EventosVsRiesgosCausas] ([IdEvento],[IdRiesgo],[IdCausa]) VALUES (" + IdEvento + ", " + IdRiesgo + ", " + IdCausa + ")");
                cDataBase.desconectar(); 
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        //yoendy - inserta table frecuencia
        public void relacionarRiesgoFrecuencia(int IdRiesgo, int IdFrecuencia, int EventosMaximos, string NombreFrecuencia)
        {
            try
            {
                DateTime today = Convert.ToDateTime(DateTime.Now);                
                string hoy = string.Format("{0:dd-MM-yyyy hh:mm:ss}", today);

                cDataBase.conectar();
                cDataBase.ejecutarQuery("INSERT INTO [Riesgos].[ParamFrecEventos] VALUES (" + IdRiesgo + ", " + IdFrecuencia + ", " + EventosMaximos + ", '" + NombreFrecuencia + "', " + idUsuario + ", '" + hoy + "')");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public DataTable loadInfoRiesgoEvento(String IdEvento)
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT Riesgos.EventoRiesgo.IdEventoRiesgo, Riesgos.EventoRiesgo.IdRiesgo, Riesgos.Riesgo.Codigo, Riesgos.Riesgo.Nombre, Riesgos.EventoRiesgo.FechaRegistro FROM Riesgos.EventoRiesgo INNER JOIN Riesgos.Riesgo ON Riesgos.EventoRiesgo.IdRiesgo = Riesgos.Riesgo.IdRiesgo WHERE (Riesgos.EventoRiesgo.IdEvento = " + IdEvento + ")");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public string mtdGetTipoArea(String IdJerarquia)
        {
            DataTable dtInformacion = new DataTable();
            string TipoArea = string.Empty;
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT isnull(JO.TipoArea,0) as TipoArea FROM [Parametrizacion].[JerarquiaOrganizacional] AS JO " +
                    "INNER JOIN [Parametrizacion].[DetalleJerarquiaOrg] AS DJ ON DJ.idHijo = JO.idHijo" +
                    " WHERE JO.idHijo ='" + IdJerarquia + "'");
                cDataBase.desconectar();
                TipoArea = dtInformacion.Rows[0]["TipoArea"].ToString().Trim();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return TipoArea;
        }
        public string mtdGetIdJerarquia(string idEvento)
        {
            DataTable dtInformacion = new DataTable();
            string TipoArea = string.Empty;
            try
            {
                cDataBase.conectar();
                string query = "SELECT GeneraEvento FROM [Riesgos].[Eventos] where IdEvento = "+idEvento;
                dtInformacion = cDataBase.ejecutarConsulta(query);
                cDataBase.desconectar();
                TipoArea = dtInformacion.Rows[0]["GeneraEvento"].ToString().Trim();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return TipoArea;
        }
        public string mtdGetCodEvento(string idEvento)
        {
            DataTable dtInformacion = new DataTable();
            string TipoArea = string.Empty;
            try
            {
                cDataBase.conectar();
                string query = "SELECT CodigoEvento FROM [Riesgos].[Eventos] where IdEvento = " + idEvento;
                dtInformacion = cDataBase.ejecutarConsulta(query);
                cDataBase.desconectar();
                TipoArea = dtInformacion.Rows[0]["CodigoEvento"].ToString().Trim();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return TipoArea;
        }
        public string mtdGetDesEvento(string idEvento)
        {
            DataTable dtInformacion = new DataTable();
            string TipoArea = string.Empty;
            try
            {
                cDataBase.conectar();
                string query = "SELECT DescripcionEvento FROM [Riesgos].[Eventos] where IdEvento = " + idEvento;
                dtInformacion = cDataBase.ejecutarConsulta(query);
                cDataBase.desconectar();
                TipoArea = dtInformacion.Rows[0]["DescripcionEvento"].ToString().Trim();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return TipoArea;
        }
        //public void modificarEvento(String FechaEvento, String CodigoEvento, String IdRegion, String IdPais, String IdDepartamento, String IdCiudad, String IdOficinaSucursal, String IdCadenaValor, String IdMacroproceso, String IdProceso, String IdSubProceso, String IdActividad, String ResponsableEvento, String ProcesoInvolucrado, String AplicativoInvolucrado, String DescripcionEvento, String IdClaseEvento, String IdTipoPerdidaEvento, String ServicioProductoAfectado, String FechaInicio, String HoraInicio, String FechaFinalizacion, String HoraFinalizacion, String FechaDescubrimiento, String HoraDescubrimiento, String ResponsableContabilidad, String CuentaPUC, String CuentaOrden, String CuentaPerdida, String Moneda1, String TasaCambio1, String ValorPesos1, String ValorRecuperadoTotal, String Moneda2, String TasaCambio2, String ValorPesos2, String ValorRecuperadoSeguro, String ValorPesos3, String Observaciones, String FuenteRecuperacion, String IdEvento)
        //{
        //    try
        //    {
        //        cDataBase.conectar();
        //        cDataBase.ejecutarQuery("UPDATE Riesgos.Eventos SET FechaEvento = CONVERT(datetime, '" + FechaEvento + "', 120), CodigoEvento = N'" + CodigoEvento + "', IdRegion = " + IdRegion + ", IdPais = " + IdPais + ", IdDepartamento = " + IdDepartamento + ", IdCiudad = " + IdCiudad + ", IdOficinaSucursal = " + IdOficinaSucursal + ", IdCadenaValor = " + IdCadenaValor + ", IdMacroproceso = " + IdMacroproceso + ", IdProceso = " + IdProceso + ", IdSubProceso = " + IdSubProceso + ", IdActividad = " + IdActividad + ", ResponsableEvento = " + ResponsableEvento + ", ProcesoInvolucrado = N'" + ProcesoInvolucrado + "', AplicativoInvolucrado = N'" + AplicativoInvolucrado + "', DescripcionEvento = N'" + DescripcionEvento + "', IdClaseEvento = " + IdClaseEvento + ", IdTipoPerdidaEvento = " + IdTipoPerdidaEvento + ", ServicioProductoAfectado = N'" + ServicioProductoAfectado + "', FechaInicio = CONVERT(datetime, '" + FechaInicio + "', 120), HoraInicio = N'" + HoraInicio + "', FechaFinalizacion = CONVERT(datetime, '" + FechaFinalizacion + "', 120), HoraFinalizacion = N'" + HoraFinalizacion + "', FechaDescubrimiento = CONVERT(datetime, '" + FechaDescubrimiento + "', 120), HoraDescubrimiento = N'" + HoraDescubrimiento + "', ResponsableContabilidad = " + ResponsableContabilidad + ", CuentaPUC = N'" + CuentaPUC + "', CuentaOrden = N'" + CuentaOrden + "', CuentaPerdida = N'" + CuentaPerdida + "', Moneda1 = N'" + Moneda1 + "', TasaCambio1 = N'" + TasaCambio1 + "', ValorPesos1 = N'" + ValorPesos1 + "', ValorRecuperadoTotal = N'" + ValorRecuperadoTotal + "', Moneda2 = N'" + Moneda2 + "', TasaCambio2 = N'" + TasaCambio2 + "', ValorPesos2 = N'" + ValorPesos2 + "', ValorRecuperadoSeguro = N'" + ValorRecuperadoSeguro + "', ValorPesos3 = N'" + ValorPesos3 + "', Observaciones = N'" + Observaciones + "', FuenteRecuperacion = N'" + FuenteRecuperacion + "' WHERE (IdEvento = " + IdEvento + ")");
        //        cDataBase.desconectar();
        //    }
        //    catch (Exception ex)
        //    {
        //        cDataBase.desconectar();
        //        cError.errorMessage(ex.Message + ", " + ex.StackTrace);
        //        throw new Exception(ex.Message);
        //    }
        //}

        public void modificarEvento1(String ResponsableContabilidad, String CuentaPUC, String CuentaOrden, String CuentaPerdida, String Moneda1, String TasaCambio1, String ValorPesos1, String ValorRecuperadoTotal, String Moneda2, String TasaCambio2, String ValorPesos2, String ValorRecuperadoSeguro, String ValorPesos3, String Recuperacion, String ValorRecuperacion, String IdEvento)
        {
            try
            {
                cDataBase.conectar();
                cDataBase.ejecutarQuery("update Riesgos.Eventos set ResponsableContabilidad='" + ResponsableContabilidad + "',CuentaPUC = '" + CuentaPUC + "',CuentaOrden='" + CuentaOrden + "',CuentaPerdida='" + CuentaPerdida + "',Moneda1='" + Moneda1 + "',TasaCambio1='" + TasaCambio1 + "',ValorPesos1='" + ValorPesos1 + "',ValorRecuperadoTotal='" + ValorRecuperadoTotal + "',Moneda2='" + Moneda2 + "',TasaCambio2='" + TasaCambio2 + "',ValorPesos2='" + ValorPesos2 + "',ValorRecuperadoSeguro='" + ValorRecuperadoSeguro + "',ValorPesos3='" + ValorPesos3 + "',Recuperacion='" + Recuperacion + "',ValorRecuperacion='" + ValorRecuperacion + "' WHERE (CodigoEvento = '" + IdEvento + "')");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public void mtdModificarEvento2(string ResponsableContabilidad, string CuentaPUC, string CuentaOrden, string CuentaPerdida,
            string Moneda1, string TasaCambio1, string ValorPesos1, string ValorRecuperadoTotal, string Moneda2, string TasaCambio2,
            string ValorPesos2, string ValorRecuperadoSeguro, string ValorPesos3, string Recuperacion, string ValorRecuperacion, string IdEvento,
            bool booIsFecha, string strFechaContab, string strHoraContab, string strFechaRecup, string strHoraRecup, string CuantiaRecup, 
            string CuantiaOtraRecup, string CuantiaNeta)
        {
            string strConsulta = string.Empty;
            //string strFechaContab = string.Empty, strHoraContab = string.Empty;

            try
            {
                if (booIsFecha)
                {
                    //strFechaContab = DateTime.Today.ToString("yyyy-MM-dd") + " 00:00:00";

                    //if (Convert.ToInt32(DateTime.Now.ToString("HH")) > 13)
                    //    strHoraContab = DateTime.Now.ToString("hh:mm") + " p.m";
                    //else
                    //    strHoraContab = DateTime.Now.ToString("hh:mm") + " a.m";

                    strConsulta = string.Format("UPDATE Riesgos.Eventos SET ResponsableContabilidad='{0}',CuentaPUC = '{1}',CuentaOrden='{2}'," +
                        "CuentaPerdida='{3}',Moneda1='{4}',TasaCambio1='{5}',ValorPesos1='{6}',ValorRecuperadoTotal='{7}',Moneda2='{8}'," +
                        "TasaCambio2='{9}',ValorPesos2='{10}',ValorRecuperadoSeguro='{11}',ValorPesos3='{12}',Recuperacion='{13}'," +
                        "ValorRecuperacion='{14}', FechaContabilidad = CONVERT(datetime, '" + strFechaContab + "', 120), HoraContabilidad = '{16}'," +
                        //"ValorRecuperacion='{14}', FechaContabilidad = '{15}', HoraContabilidad = '{16}'," +
                        "FechaRecuperacion = CONVERT(datetime, '" + strFechaRecup + "', 120), HoraRecuperacion='{18}', CuantiaRecup='{19}', CuantiaOtraRecup='{20}'," +
                        //"FechaRecuperacion='{17}', HoraRecuperacion='{18}', CuantiaRecup='{19}', CuantiaOtraRecup='{20}'," +
                        "CuantiaNeta = '{21}' WHERE (CodigoEvento = '{22}')",
                        ResponsableContabilidad, CuentaPUC, CuentaOrden, CuentaPerdida, Moneda1, TasaCambio1, ValorPesos1, ValorRecuperadoTotal,
                        Moneda2, TasaCambio2, ValorPesos2, ValorRecuperadoSeguro, ValorPesos3, Recuperacion, ValorRecuperacion, strFechaContab, 
                        strHoraContab, strFechaRecup, strHoraRecup, CuantiaRecup, CuantiaOtraRecup, CuantiaNeta, IdEvento);
                }
                else
                    strConsulta = string.Format("UPDATE Riesgos.Eventos SET ResponsableContabilidad='{0}',CuentaPUC = '{1}',CuentaOrden='{2}'," +
                        "CuentaPerdida='{3}',Moneda1='{4}',TasaCambio1='{5}',ValorPesos1='{6}',ValorRecuperadoTotal='{7}',Moneda2='{8}'," +
                        "TasaCambio2='{9}',ValorPesos2='{10}',ValorRecuperadoSeguro='{11}',ValorPesos3='{12}',Recuperacion='{13}'," +
                        "ValorRecuperacion='{14}', FechaContabilidad = CONVERT(datetime, '" + strFechaContab + "', 120), HoraRecuperacion = '{16}'," +
                        "CuantiaRecup='{17}', CuantiaOtraRecup='{18}', CuantiaNeta = '{19}' WHERE (CodigoEvento = '{20}')",
                        ResponsableContabilidad, CuentaPUC, CuentaOrden, CuentaPerdida, Moneda1, TasaCambio1, ValorPesos1, ValorRecuperadoTotal, 
                        Moneda2, TasaCambio2, ValorPesos2, ValorRecuperadoSeguro, ValorPesos3, Recuperacion, ValorRecuperacion,
                        strFechaRecup, strHoraRecup, CuantiaRecup, CuantiaOtraRecup, CuantiaNeta, IdEvento);

                cDataBase.conectar();
                cDataBase.ejecutarQuery(strConsulta);
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public DataTable loadCodigoEvento()
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT TOP (1) IdEvento+1 AS NumRegistros FROM Riesgos.Eventos ORDER BY IdEvento DESC");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public DataTable loadCodigoNHEvento()
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT TOP (1) IdEvento+1 AS NumRegistros FROM Riesgos.NoHuboEventos ORDER BY IdEvento DESC");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public void ModificaEventoTab1(String IdEmpresa, String IdRegion, String IdPais, String IdDepartamento, String IdCiudad, String IdOficinaSucursal, String DetalleUbicacion, String DescripcionEvento, String IdServicio, String IdSubServicio, String FechaInicio, String HoraInicio, String FechaFinalizacion, String HoraFinalizacion, String FechaDescubrimiento, String HoraDescubrimiento, String IdCanal, String IdGeneraEvento, String GeneraEvento, String cuantiaperdida, String CodigoEvento, String NomGeneradorEvento, String NombreImpCual)
        {
            try
            {
                cDataBase.conectar();               
                cDataBase.ejecutarQuery("UPDATE Riesgos.Eventos SET IdEmpresa='" + IdEmpresa + "',IdRegion='" + IdRegion + "',IdPais='" + IdPais + "',IdDepartamento='" + IdDepartamento + "',IdCiudad='" + IdCiudad + "',IdOficinaSucursal='" + IdOficinaSucursal + "',DetalleUbicacion='" + DetalleUbicacion + "',DescripcionEvento='" + DescripcionEvento + "',IdServicio='" + IdServicio + "',IdSubServicio='" + IdSubServicio + "',FechaInicio=CONVERT(datetime, '" + FechaInicio + "', 120),HoraInicio='" + HoraInicio + "',FechaFinalizacion=" + FechaFinalizacion + ", HoraFinalizacion='" + HoraFinalizacion + "',FechaDescubrimiento = CONVERT(datetime, '" + FechaDescubrimiento + "', 120), HoraDescubrimiento='" + HoraDescubrimiento + "',IdCanal='" + IdCanal + "',IdGeneraEvento='" + IdGeneraEvento + "',GeneraEvento='" + GeneraEvento + "',cuantiaperdida='" + cuantiaperdida + "',NomGeneradorEvento=" + NomGeneradorEvento + ", NombreImpactoCualitativo='" + NombreImpCual + "' WHERE CodigoEvento ='" + CodigoEvento + "'");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }
        public void ModificaEventoMasivo(string cadenaValor, string macroProceso, string proceso, string subproceso, string actividad, 
            string IdClase, string IdSubClase, string IdTipoPerdidaEvento, string IdLineaProceso, string IdSubLineaProceso, 
            string AfectaContinudad, string IdEstado, string Observaciones, string CuentaPUC, string CuentaOrden, string TasaCambio1, 
            string ValorPesos1, string ValorRecuperadoTotal, string Moneda2, string TasaCambio2, string ValorPesos2, string Recuperacion, 
            string FechaContabilidad, string HoraContabilidad, string IdEvento
            , string FechaRecuperacion, string HoraRecuperacion, string CuantiaRecup, string CuantiaOtraRecup, string CuantiaNeta)
        {
            try
            {

                string query = "UPDATE Riesgos.Eventos SET IdCadenaValor='" + cadenaValor + "',IdMacroproceso='" + macroProceso + 
                    "',IdProceso='" + proceso + "',IdSubProceso='" + subproceso + "',IdActividad='" + actividad + "',IdClase='" + IdClase + "'" +
                    ",IdSubClase='" + IdSubClase +
                    "',IdTipoPerdidaEvento='" + IdTipoPerdidaEvento + "',IdLineaProceso='" + IdLineaProceso + "'" +
                    ",IdSubLineaProceso='" + IdSubLineaProceso + "',AfectaContinudad='" + AfectaContinudad +
                    "',IdEstado='" + IdEstado + "',Observaciones='" + Observaciones + "',CuentaPUC='" + CuentaPUC + "'" +
                    ",CuentaOrden='" + CuentaOrden + "',TasaCambio1='" + TasaCambio1 + "',ValorPesos1='" + ValorPesos1 +
                    "',ValorRecuperadoTotal='" + ValorRecuperadoTotal + "',Moneda2='" + Moneda2 + "',TasaCambio2='" + TasaCambio2 + "'" +
                    ",ValorPesos2='" + ValorPesos2 + "',Recuperacion='" + Recuperacion +
                    "',FechaContabilidad= CONVERT(datetime,'" + FechaContabilidad + "', 120),HoraContabilidad='" + HoraContabilidad +
                    "', FechaRecuperacion=CONVERT(datetime,'" + FechaRecuperacion+ "', 120),HoraRecuperacion='" + HoraRecuperacion+ "',CuantiaRecup='"+CuantiaRecup+"'" +
                    ",CuantiaOtraRecup='"+CuantiaOtraRecup+ "',CuantiaNeta='"+CuantiaNeta+"'" +
                    " WHERE IdEvento ='" + IdEvento + "'";
                cDataBase.conectar();
                cDataBase.ejecutarQuery(query);
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }
        public DataTable loadDDLClaseEvento()
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT IdClaseEvento, NombreClaseEvento FROM Parametrizacion.ClaseEvento");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public DataTable loadDDLTipoPerdidaEvento()
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT IdTipoPerdidaEvento, NombreTipoPerdidaEvento " +
                    "FROM Parametrizacion.TipoPerdidaEvento where Estado = 1");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public DataTable ReporteEventos(string IdCadenaValor, string IdMacroProceso,
            string IdProceso, string IdClasificacionRiesgo, string IdClasificacionGeneralRiesgo,
            string NombreRiesgoInherente, string NombreRiesgoResidual, string IdEmpresa,
            string numeroQuery, string IdRiesgo, string FechaIni, string FechaFin)
        {
            #region Variables
            DataTable dtInformacion = new DataTable();
            string condicion = string.Empty, strConsulta = string.Empty, strQry1 = string.Empty, strQry2 = string.Empty,
                strCondRE = string.Empty, strCondRNHE = string.Empty;
            #endregion Variables

            try
            {
                if (numeroQuery == "3")
                {
                    #region Filtros para Reporte Eventos Vs Plan Accion
                    #region Cadena Valor
                    if (IdCadenaValor != "---")
                        condicion = "AND (a.IdCadenaValor = " + IdCadenaValor + ") ";
                    #endregion

                    #region MacroProceso
                    if (IdMacroProceso != "---")
                    {
                        if (string.IsNullOrEmpty(condicion.Trim()))
                            condicion = "AND (a.IdMacroProceso = " + IdMacroProceso + ") ";
                        else
                            condicion += "AND (a.IdMacroProceso = " + IdMacroProceso + ") ";
                    }
                    #endregion

                    #region Proceso
                    if (IdProceso != "---")
                    {
                        if (string.IsNullOrEmpty(condicion.Trim()))
                            condicion = "AND (a.IdProceso = " + IdProceso + ") ";
                        else
                            condicion += "AND (a.IdProceso = " + IdProceso + ") ";
                    }
                    #endregion

                    #region Empresa
                    if (IdEmpresa != "---")
                    {
                        if (string.IsNullOrEmpty(condicion.Trim()))
                            condicion = " AND a.IdEmpresa = '" + IdEmpresa + "'";
                        else
                            condicion += " AND a.IdEmpresa = '" + IdEmpresa + "'";
                    }
                    #endregion

                    #region FechaIni
                    if (FechaIni != "")
                    {
                        FechaIni += " 00:00:00:000";
                        if (string.IsNullOrEmpty(condicion.Trim()))
                            condicion = " AND (a.FechaEvento >= CONVERT(datetime, '" + FechaIni + "', 120)) ";
                        else
                            condicion += " AND (a.FechaEvento >= CONVERT(datetime, '" + FechaIni + "', 120)) ";
                    }
                    #endregion

                    #region FechaFin
                    if (FechaFin != "")
                    {
                        FechaFin += " 23:59:59:998";
                        if (string.IsNullOrEmpty(condicion.Trim()))
                            condicion = " AND (a.FechaEvento <= CONVERT(datetime, '" + FechaFin + "', 120)) ";
                        else
                            condicion += "AND (a.FechaEvento <= CONVERT(datetime, '" + FechaFin + "', 120)) ";
                    }
                    #endregion
                    #endregion Filtros
                }
                else
                {
                    if (numeroQuery == "4")
                    {
                        #region Filtros Reporte Sin Reporte
                        if (!string.IsNullOrEmpty(FechaIni))
                        {
                            //if (string.IsNullOrEmpty(condicion))
                            //    condicion = "WHERE (PA.FechaEvento >= CONVERT(DATETIME, '" + FechaIni + " 00:00:00:000', 120)) ";
                            //else
                            strCondRNHE += "AND (RNHE.FechaEvento >= CONVERT(DATETIME, '" + FechaIni + " 00:00:00:000', 120)) ";
                            strCondRE += "AND (RE.FechaEvento >= CONVERT(DATETIME, '" + FechaIni + " 00:00:00:000', 120)) ";
                        }

                        if (!string.IsNullOrEmpty(FechaFin))
                        {
                            //if (string.IsNullOrEmpty(condicion))
                            //    condicion = "WHERE (PA.FechaEvento <= CONVERT(DATETIME, '" + FechaFin + " 23:59:59:998', 120)) ";
                            //else
                            strCondRNHE += "AND (RNHE.FechaEvento <= CONVERT(DATETIME, '" + FechaFin + " 23:59:59:998', 120)) ";
                            strCondRE += "AND (RE.FechaEvento <= CONVERT(DATETIME, '" + FechaFin + " 23:59:59:998', 120)) ";
                        }
                        #endregion Filtros Reporte Sin Reporte
                    }
                    else
                    {
                        #region Filtros otros reportes
                        #region Cadena Valor
                        if (IdCadenaValor != "---")
                            condicion = "WHERE (a.IdCadenaValor = " + IdCadenaValor + ") ";
                        #endregion

                        #region MacroProceso
                        if (IdMacroProceso != "---")
                        {
                            if (condicion.Trim() == "")
                                condicion = "WHERE (a.IdMacroProceso = " + IdMacroProceso + ") ";
                            else
                                condicion += "AND (a.IdMacroProceso = " + IdMacroProceso + ") ";
                        }
                        #endregion

                        #region Proceso
                        if (IdProceso != "---")
                        {
                            if (condicion.Trim() == "")
                                condicion = "WHERE (a.IdProceso = " + IdProceso + ") ";
                            else
                                condicion += "AND (a.IdProceso = " + IdProceso + ") ";
                        }
                        #endregion

                        #region Empresa
                        if (IdEmpresa != "---")
                        {
                            if (condicion.Trim() == "")
                                condicion = " WHERE a.IdEmpresa = '" + IdEmpresa + "'";
                            else
                                condicion += " AND a.IdEmpresa = '" + IdEmpresa + "'";
                        }
                        #endregion

                        #region FechaIni
                        if (FechaIni != "")
                        {
                            FechaIni += " 00:00:00:000";
                            if (condicion == "")
                                condicion = "WHERE (a.FechaEvento >= CONVERT(datetime, '" + FechaIni + "', 120)) ";
                            else
                                condicion += "AND (a.FechaEvento >= CONVERT(datetime, '" + FechaIni + "', 120)) ";
                        }
                        #endregion

                        #region FechaFin
                        if (FechaFin != "")
                        {
                            FechaFin += " 23:59:59:998";
                            if (condicion == "")
                                condicion = "WHERE (a.FechaEvento <= CONVERT(datetime, '" + FechaFin + "', 120)) ";
                            else
                                condicion += "AND (a.FechaEvento <= CONVERT(datetime, '" + FechaFin + "', 120)) ";
                        }
                        #endregion
                        #endregion Filtros
                    }
                }

                switch (numeroQuery)
                {
                    case "1":
                        #region No Hubo Eventos
                        //dtInformacion = cDataBase.ejecutarConsulta("SELECT a.CodigoEvento,b.Descripcion Empresa, SUBSTRING(convert(varchar,a.FechaEvento,120),1,10) as FechaNoHuboEvento,j.NombreResponsable,jo.NombreHijo as Cargo,jj.NombreHijo as Area from Riesgos.NoHuboEventos a left join Listas.Usuarios l on a.IdUsuario = l.IdUsuario left join Parametrizacion.DetalleJerarquiaOrg j on l.IdJerarquia = j.idHijo left join Parametrizacion.JerarquiaOrganizacional jo on j.idHijo = jo.idHijo left join Parametrizacion.JerarquiaOrganizacional jj on jo.idPadre = jj.idHijo left join Eventos.Empresa b on a.IdEmpresa = b.IdEmpresa " + condicion + "ORDER BY a.CodigoEvento");

                        strQry1 = "SELECT a.CodigoEvento, LTRIM(RTRIM(ISNULL(b.Descripcion, ''))) Empresa, SUBSTRING(CONVERT(VARCHAR, a.FechaEvento, 120), 1, 10)  FechaNoHuboEvento, j.NombreResponsable, jo.NombreHijo Cargo, PAr.NombreArea Area";
                        strQry2 = "FROM Riesgos.NoHuboEventos a LEFT JOIN Listas.Usuarios l ON a.IdUsuario = l.IdUsuario LEFT JOIN Parametrizacion.DetalleJerarquiaOrg j ON l.IdJerarquia = j.idHijo LEFT JOIN Parametrizacion.JerarquiaOrganizacional jo ON j.idHijo = jo.idHijo LEFT JOIN Eventos.Empresa b ON a.IdEmpresa = b.IdEmpresa LEFT JOIN [Parametrizacion].[JerarquiaOrganizacional] PJO ON l.IdJerarquia  = PJO.idHijo LEFT JOIN [Parametrizacion].[DetalleJerarquiaOrg] PDJO ON PJO.idHijo = PDJO.idHijo LEFT JOIN [Parametrizacion].[Area] PAr ON PDJO.IdArea = PAr.IdArea";
                        strConsulta = string.Format("{0} {1} {2} ORDER BY SUBSTRING(a.CodigoEvento, 1, 3) ASC , CONVERT(INT,SUBSTRING(a.CodigoEvento, 4, 3)) ASC", strQry1, strQry2, condicion);
                        #endregion

                        break;
                    case "2":
                        #region Comentado
                        //strSelect = "SELECT a.CodigoEvento 'Código', LTRIM(RTRIM(ISNULL(em.Descripcion, ''))) Empresa, reg.NombreRegion 'Región', pai.NombrePais 'Pais', LTRIM(RTRIM(ISNULL(dep.NombreDepartamento, ''))) Departamento, ciu.NombreCiudad Ciudad, LTRIM(RTRIM(ISNULL(ofi.NombreOficinaSucursal, ''))) 'Oficina/Sucursal', LTRIM(RTRIM(ISNULL(a.DetalleUbicacion, ''))) 'Detalle Ubicación', LTRIM(RTRIM(ISNULL(a.DescripcionEvento, ''))) 'Descripción Evento', LTRIM(RTRIM(ISNULL(ser.Descripcion, ''))) 'Servicio/Producto', LTRIM(RTRIM(ISNULL(subs.SubDescripcion, ''))) 'SubServicio/SubProducto', SUBSTRING(CONVERT(VARCHAR, a.FechaInicio, 120), 1, 10) 'Fecha Inicio', a.HoraInicio 'Hora Inicio', LTRIM(RTRIM(ISNULL(SUBSTRING(CONVERT(VARCHAR, a.FechaFinalizacion, 120), 1, 10), ''))) 'Fecha Finalización', LTRIM(RTRIM(ISNULL(a.HoraFinalizacion, ''))) 'Hora Finalización', SUBSTRING(convert(varchar,a.FechaDescubrimiento,120),1,10) 'Fecha Descubrimiento', a.HoraDescubrimiento 'Hora Descubrimiento', LTRIM(RTRIM(ISNULL(can.Descripcion, ''))) Canal, LTRIM(RTRIM(ISNULL(gen.Descripcion, ''))) 'Generador del Evento', CASE a.IdGeneraEvento WHEN 2 THEN LTRIM(RTRIM(ISNULL(a.NomGeneradorEvento, ''))) ELSE LTRIM(RTRIM(ISNULL(jer.NombreResponsable, ''))) END AS 'Responsable Evento', LTRIM(RTRIM(ISNULL(a.CuantiaPerdida, ''))) 'Posible Cuantía de Pérdida ', SUBSTRING(convert(varchar,a.FechaEvento,120),1,10) 'Fecha Creacion', LTRIM(RTRIM(ISNULL(usu.Usuario, ''))) Usuario, LTRIM(RTRIM(ISNULL(cv.NombreCadenaValor, ''))) 'Cadena Valor', LTRIM(RTRIM(ISNULL(mp.Nombre, ''))) MacroProceso, LTRIM(RTRIM(ISNULL(pp.Nombre, ''))) Proceso, LTRIM(RTRIM(ISNULL(sp.Nombre, ''))) SubProceso, LTRIM(RTRIM(ISNULL(pa.Nombre, ''))) Actividad, LTRIM(RTRIM(ISNULL(jer1.NombreResponsable, ''))) 'Responsable Solución', LTRIM(RTRIM(ISNULL(ec.Descripcion, ''))) 'Clase de Riesgo', LTRIM(RTRIM(ISNULL(esc.SubDescripcion, ''))) 'SubClase de Riesgo', LTRIM(RTRIM(ISNULL(ptp.NombreTipoPerdidaEvento, ''))) 'Tipo de Pérdida', LTRIM(RTRIM(ISNULL(eln.Descripcion, ''))) 'Línea Operativa ', LTRIM(RTRIM(ISNULL(esln.SubDescripcion, ''))) 'SubLínea Operativa', LTRIM(RTRIM(ISNULL(a.MasLineas, ''))) 'Más Líneas Operativas', CASE a.AfectaContinudad WHEN 1 THEN 'Si' WHEN 0 THEN 'No' ELSE '' END 'Afecta Continuidad', LTRIM(RTRIM(ISNULL(ee.Descripcion, ''))) Estado, LTRIM(RTRIM(ISNULL(a.Observaciones, ''))) Observaciones, LTRIM(RTRIM(ISNULL(jer2.NombreResponsable, ''))) 'Responsable Contabilidad', LTRIM(RTRIM(ISNULL(a.CuentaPUC, ''))) 'Cuenta PUC', LTRIM(RTRIM(ISNULL(a.CuentaOrden, ''))) 'Cuenta de Orden', LTRIM(RTRIM(ISNULL(a.TasaCambio1, ''))) 'Tasa de Cambio', LTRIM(RTRIM(ISNULL(a.ValorPesos1, ''))) 'Valor en Pesos', LTRIM(RTRIM(ISNULL(a.ValorRecuperadoTotal, ''))) 'Valor Recuperado Total', LTRIM(RTRIM(ISNULL(a.TasaCambio2, ''))) 'Tasa de Cambio 2', LTRIM(RTRIM(ISNULL(a.ValorPesos2, ''))) 'Valor en Pesos 2', CASE a.Recuperacion WHEN 1 THEN 'Si' WHEN 0 THEN 'No' ELSE '' END 'Recuperación ', LTRIM(RTRIM(ISNULL(a.ValorRecuperacion, ''))) 'Fuente de la Recuperación', LTRIM(RTRIM(ISNULL(a.Moneda2, ''))) Moneda, LTRIM(RTRIM(ISNULL(PAr.NombreArea, ''))) Area";
                        //strFrom = "FROM Riesgos.Eventos a LEFT JOIN Eventos.Empresa em ON a.IdEmpresa = em.IdEmpresa LEFT JOIN Parametrizacion.Regiones reg ON a.IdRegion = reg.IdRegion LEFT JOIN Parametrizacion.Paises pai ON a.IdRegion = pai.IdRegion AND a.IdPais = pai.IdPais LEFT JOIN Parametrizacion.Departamentos dep ON a.IdPais = dep.IdPais AND a.IdDepartamento = dep.IdDepartamento LEFT JOIN Parametrizacion.Ciudades ciu ON a.IdDepartamento = ciu.IdDepartamento AND a.IdCiudad = ciu.IdCiudad LEFT JOIN Parametrizacion.OficinaSucursal ofi ON a.IdCiudad = ofi.IdCiudad AND a.IdOficinaSucursal = ofi.IdOficinaSucursal LEFT JOIN Eventos.Servicio ser ON a.IdServicio = ser.IdServicio LEFT JOIN Eventos.SubServicio subs ON a.IdSubServicio = subs.IdSubServicio LEFT JOIN Eventos.Canal can ON a.IdCanal = can.IdCanal LEFT JOIN Eventos.Generador gen ON a.IdGeneraEvento = gen.IdGenerador LEFT JOIN Parametrizacion.DetalleJerarquiaOrg jer ON a.GeneraEvento = jer.idHijo LEFT JOIN Listas.Usuarios usu ON a.IdUsuario = usu.IdUsuario LEFT JOIN Procesos.CadenaValor cv ON a.IdCadenaValor = cv.IdCadenaValor LEFT JOIN Procesos.Macroproceso mp ON a.IdCadenaValor = mp.IdCadenaValor AND a.IdMacroproceso = mp.IdMacroProceso LEFT JOIN Procesos.Proceso pp ON a.IdMacroproceso = pp.IdMacroProceso AND a.IdProceso = pp.IdProceso LEFT JOIN Procesos.Subproceso sp ON a.IdProceso = sp.IdProceso AND a.IdSubProceso = sp.IdSubproceso LEFT JOIN Procesos.Actividad pa ON a.IdSubProceso = pa.IdSubproceso AND a.IdActividad = pa.IdActividad LEFT JOIN Parametrizacion.DetalleJerarquiaOrg jer1 ON a.ResponsableEvento = jer1.idHijo LEFT JOIN Eventos.Clase ec ON a.IdClase = ec.IdClase LEFT JOIN Eventos.SubClase esc ON a.IdClase = esc.IdClase AND a.IdSubClase = esc.IdSubClase LEFT JOIN Parametrizacion.TipoPerdidaEvento ptp ON a.IdTipoPerdidaEvento = ptp.IdTipoPerdidaEvento LEFT JOIN Eventos.LineaNegocio eln ON a.IdLineaProceso = eln.IdLineaNegocio LEFT JOIN Eventos.SubLineaNegocio esln ON a.IdLineaProceso = esln.IdLineaNegocio AND a.IdSubLineaProceso = esln.IdSubLineaNegocio LEFT JOIN Eventos.Estado ee ON a.IdEstado = ee.IdEstado LEFT JOIN Parametrizacion.DetalleJerarquiaOrg jer2 ON a.ResponsableContabilidad = jer2.idHijo LEFT JOIN Parametrizacion.DetalleJerarquiaOrg jer3 ON usu.IdJerarquia = jer3.idHijo LEFT JOIN Parametrizacion.Area PAr ON jer3.IdArea = PAr.IdArea";
                        //strConsulta = string.Format("{0} {1} {2} ORDER BY a.CodigoEvento", strSelect, strFrom, condicion);

                        //dtInformacion = cDataBase.ejecutarConsulta(strConsulta);
                        #endregion Comentado

                        #region Eventos
                        //dtInformacion = cDataBase.ejecutarConsulta("select a.CodigoEvento, b.Descripcion as Empresa,c.NombreRegion,d.NombrePais,e.NombreDepartamento,f.NombreCiudad,g.NombreOficinaSucursal,a.DetalleUbicacion,a.DescripcionEvento,h.Descripcion as Servicio,i.SubDescripcion as Subservicio,substring(convert(varchar,a.FechaInicio,120),1,10) as FechaInicio,a.HoraInicio,substring(convert(varchar,a.FechaFinalizacion,120),1,10) as FechaFinalizacion,a.HoraFinalizacion,substring(convert(varchar,a.FechaDescubrimiento,120),1,10) as FechaDescubrimiento,a.HoraDescubrimiento,j.Descripcion as Canal,k.Descripcion as GeneradorEvento,l2.NombreResponsable as ResponsableEvento,l.NombreResponsable as ResponsableSolucion,l1.NombreResponsable as ResponsableContabilizacion,a.CuantiaPerdida,m.NombreCadenaValor as CadenaValor,n.Nombre as Macroproceso,o.Nombre as Proceso,p.Nombre as Subproceso,q.Nombre as Actividad,r.Descripcion as ClaseRiesgo,s.SubDescripcion as SubClaseRiesgo,t.NombreTipoPerdidaEvento,a.AfectaContinudad,u.Descripcion as Estado,a.Observaciones,SUBSTRING(convert(varchar,a.FechaEvento,120),1,10) as FechaEvento,j1.NombreResponsable as Originador,j2.NombreHijo as Cargo,j3.NombreHijo as Area from Riesgos.Eventos a left join Eventos.Empresa b on a.IdEmpresa = b.IdEmpresa left join Parametrizacion.Regiones c on a.IdRegion = c.IdRegion left join Parametrizacion.Paises d on a.IdPais = d.IdPais left join Parametrizacion.Departamentos e on a.IdDepartamento = e.IdDepartamento left join Parametrizacion.Ciudades f on a.IdCiudad = f.IdCiudad left join Parametrizacion.OficinaSucursal g on a.IdOficinaSucursal = g.IdOficinaSucursal left join Eventos.Servicio h on a.IdServicio = h.IdServicio left join Eventos.SubServicio i on a.IdSubServicio = i.IdSubServicio left join Eventos.Canal j on a.IdCanal = j.IdCanal left join Eventos.Generador k on a.IdGeneraEvento = k.IdGenerador left join Parametrizacion.DetalleJerarquiaOrg l2 on a.GeneraEvento = l2.idHijo left join Parametrizacion.DetalleJerarquiaOrg l1 on a.ResponsableContabilidad = l1.idHijo left join Parametrizacion.DetalleJerarquiaOrg l on a.ResponsableEvento = l.idHijo left join Procesos.CadenaValor m on a.IdCadenaValor = m.IdCadenaValor left join Procesos.Macroproceso n on a.IdMacroproceso = n.IdMacroProceso left join Procesos.Proceso o on a.IdProceso = o.IdProceso left join Procesos.Subproceso p on a.IdSubProceso = p.IdSubproceso left join Procesos.Actividad q on a.IdActividad = q.IdActividad left join Eventos.Clase r on a.IdClase = r.IdClase left join Eventos.SubClase s on a.IdSubClase = s.IdSubClase left join Parametrizacion.TipoPerdidaEvento t on a.IdTipoPerdidaEvento = t.IdTipoPerdidaEvento left join Eventos.Estado u on a.IdEstado = u.IdEstado left join Listas.Usuarios v on a.IdUsuario = v.IdUsuario left join Parametrizacion.DetalleJerarquiaOrg j1 on v.IdJerarquia = j1.idHijo left join Parametrizacion.JerarquiaOrganizacional j2 on j1.idHijo = j2.idHijo left join Parametrizacion.JerarquiaOrganizacional j3 on j2.idPadre = j3.idHijo " + condicion + " ORDER BY a.CodigoEvento");
                        //dtInformacion = cDataBase.ejecutarConsulta(
                        //    "select a.CodigoEvento 'Código',em.Descripcion Empresa,reg.NombreRegion 'Región',pai.NombrePais 'Pais',dep.NombreDepartamento Departamento,ciu.NombreCiudad Ciudad,ofi.NombreOficinaSucursal 'Oficina/Sucursal' ,a.DetalleUbicacion 'Detalle Ubicación',a.DescripcionEvento 'Descripción Evento',ser.Descripcion 'Servicio/Producto',subs.SubDescripcion 'SubServicio/SubProducto',SUBSTRING(convert(varchar,a.FechaInicio,120),1,10) 'Fecha Inicio', a.HoraInicio 'Hora Inicio',SUBSTRING(convert(varchar,a.FechaFinalizacion,120),1,10) 'Fecha Finalización',a.HoraFinalizacion 'Hora Finalización',SUBSTRING(convert(varchar,a.FechaDescubrimiento,120),1,10) 'Fecha Descubrimiento',a.HoraDescubrimiento 'Hora Descubrimiento',can.Descripcion Canal,gen.Descripcion 'Generador del Evento',case a.IdGeneraEvento when 2 then a.NomGeneradorEvento else jer.NombreResponsable end as 'Responsable Evento', a.CuantiaPerdida 'Posible Cuantía de Pérdida ',SUBSTRING(convert(varchar,a.FechaEvento,120),1,10) 'Fecha Registro',usu.Usuario Usuario,cv.NombreCadenaValor 'Cadena Valor',mp.Nombre MacroProceso, pp.Nombre Proceso,sp.Nombre SubProceso,pa.Nombre Actividad,jer1.NombreResponsable 'Responsable Solución',ec.Descripcion 'Clase de Riesgo',esc.SubDescripcion 'SubClase de Riesgo',ptp.NombreTipoPerdidaEvento 'Tipo de Pérdida',eln.Descripcion 'Línea Operativa ',esln.SubDescripcion 'SubLínea Operativa',a.MasLineas 'Más Líneas Operativas',case a.AfectaContinudad when 1 then 'Si' when 0 then 'No' end 'Afecta Continuidad',ee.Descripcion Estado,a.Observaciones,jer2.NombreResponsable 'Responsable Contabilidad',a.CuentaPUC 'Cuenta PUC',a.CuentaOrden 'Cuenta de Orden',a.TasaCambio1 'Tasa de Cambio',a.ValorPesos1 'Valor en Pesos',a.ValorRecuperadoTotal 'Valor Recuperado Total',a.TasaCambio2 'Tasa de Cambio 2',a.ValorPesos2 'Valor en Pesos 2',case a.Recuperacion when 1 then 'Si' when 0 then 'No' end 'Recuperación ',a.ValorRecuperacion 'Fuente de la Recuperación' from Riesgos.Eventos a left join Eventos.Empresa em on a.IdEmpresa = em.IdEmpresa left join Parametrizacion.Regiones reg on a.IdRegion = reg.IdRegion left join Parametrizacion.Paises pai on a.IdRegion = pai.IdRegion and a.IdPais = pai.IdPais left join Parametrizacion.Departamentos dep on a.IdPais = dep.IdPais and a.IdDepartamento = dep.IdDepartamento left join Parametrizacion.Ciudades ciu on a.IdDepartamento = ciu.IdDepartamento and a.IdCiudad = ciu.IdCiudad left join Parametrizacion.OficinaSucursal ofi on a.IdCiudad = ofi.IdCiudad and a.IdOficinaSucursal = ofi.IdOficinaSucursal left join Eventos.Servicio ser on a.IdServicio = ser.IdServicio left join Eventos.SubServicio subs on a.IdSubServicio = subs.IdSubServicio left join Eventos.Canal can on a.IdCanal = can.IdCanal left join Eventos.Generador gen on a.IdGeneraEvento = gen.IdGenerador left join Parametrizacion.DetalleJerarquiaOrg jer on a.GeneraEvento = jer.idHijo left join Listas.Usuarios usu on a.IdUsuario = usu.IdUsuario left join Procesos.CadenaValor cv on a.IdCadenaValor = cv.IdCadenaValor left join Procesos.Macroproceso mp on a.IdCadenaValor = mp.IdCadenaValor and a.IdMacroproceso = mp.IdMacroProceso left join Procesos.Proceso pp on a.IdMacroproceso = pp.IdMacroProceso and a.IdProceso = pp.IdProceso left join Procesos.Subproceso sp on a.IdProceso = sp.IdProceso and a.IdSubProceso = sp.IdSubproceso left join Procesos.Actividad pa on a.IdSubProceso = pa.IdSubproceso and a.IdActividad = pa.IdActividad left join Parametrizacion.DetalleJerarquiaOrg jer1 on a.ResponsableEvento = jer1.idHijo left join Eventos.Clase ec on a.IdClase = ec.IdClase left join Eventos.SubClase esc on a.IdClase = esc.IdClase and a.IdSubClase = esc.IdSubClase left join Parametrizacion.TipoPerdidaEvento ptp on a.IdTipoPerdidaEvento = ptp.IdTipoPerdidaEvento left join Eventos.LineaNegocio eln on a.IdLineaProceso = eln.IdLineaNegocio left join Eventos.SubLineaNegocio esln on a.IdLineaProceso = esln.IdLineaNegocio and a.IdSubLineaProceso = esln.IdSubLineaNegocio left join Eventos.Estado ee on a.IdEstado = ee.IdEstado left join Parametrizacion.DetalleJerarquiaOrg jer2 on a.ResponsableContabilidad = jer2.idHijo " + condicion + " ORDER BY a.CodigoEvento");

                        strQry1 = "SELECT a.CodigoEvento 'Código', LTRIM(RTRIM(ISNULL(em.Descripcion, ''))) Empresa, PAr.NombreArea Area, " +
                            "usu.Usuario Usuario, LTRIM(RTRIM(USU.Nombres)) + ' ' + LTRIM(RTRIM(USU.Apellidos)) NombreUsuarioRegistro," + "\n" +
                            "SUBSTRING(CONVERT(VARCHAR,a.FechaEvento, 120), 1, 10) 'Fecha Registro', reg.NombreRegion 'Región', pai.NombrePais 'Pais', " +
                            "dep.NombreDepartamento Departamento, ciu.NombreCiudad Ciudad, ofi.NombreOficinaSucursal 'Oficina/Sucursal'" + "\n" +
                            " , a.DetalleUbicacion 'Detalle Ubicación', a.DescripcionEvento 'Descripción Evento', ser.Descripcion 'Servicio/Producto'," +
                            " subs.SubDescripcion 'SubServicio/SubProducto'," + "\n" +
                            " SUBSTRING(CONVERT(VARCHAR, a.FechaInicio, 120), 1, 10) 'Fecha Inicio', a.HoraInicio 'Hora Inicio', " +
                            "SUBSTRING(CONVERT(VARCHAR, a.FechaFinalizacion, 120), 1, 10) 'Fecha Finalización', " + "\n" + "" +
                            "a.HoraFinalizacion 'Hora Finalización', " +
                            "SUBSTRING(CONVERT(VARCHAR, a.FechaDescubrimiento, 120), 1, 10) 'Fecha Descubrimiento', " +
                            "a.HoraDescubrimiento 'Hora Descubrimiento',can.Descripcion Canal," + "\n" +
                            "gen.Descripcion 'Generador del Evento', a.NomGeneradorEvento 'Responsable Evento', " +
                            "a.CuantiaPerdida 'Posible Cuantía de Pérdida', cv.NombreCadenaValor 'Cadena Valor', " +
                            "mp.Nombre MacroProceso, pp.Nombre Proceso," + "\n" +
                            " sp.Nombre SubProceso, pa.Nombre Actividad, jer1.NombreResponsable 'Responsable Solución', " +
                            "ec.Descripcion 'Clase de Riesgo',esc.SubDescripcion 'SubClase de Riesgo', " +
                            "ptp.NombreTipoPerdidaEvento 'Tipo de Pérdida'," + "\n" +
                            " eln.Descripcion 'Línea Operativa', esln.SubDescripcion 'SubLínea Operativa', a.MasLineas 'Más Líneas Operativas'," +
                            " CASE a.AfectaContinudad WHEN 1 THEN 'Si' WHEN 0 THEN 'No' END 'Afecta Continuidad', " + "\n" +
                            "ee.Descripcion Estado, a.Observaciones, jer2.NombreResponsable 'Responsable Contabilidad', a.CuentaPUC 'Cuenta PUC', " +
                            "a.CuentaOrden 'Cuenta de Orden', a.TasaCambio1 'Tasa de Cambio', a.ValorPesos1 'Valor en Pesos'," + "\n" +
                            " a.ValorRecuperadoTotal 'Valor Recuperado Total', a.TasaCambio2 'Tasa de Cambio 2', a.ValorPesos2 'Valor en Pesos 2', " +
                            "CASE a.Recuperacion WHEN 1 THEN 'Si' WHEN 0 THEN 'No' END 'Recuperación', a.ValorRecuperacion 'Fuente de la Recuperación'" + "\n" +
                            ",SUBSTRING(CONVERT(VARCHAR,a.FechaContabilidad, 120), 1, 10) 'Fecha Contabilización', [NombreImpactoCualitativo], " +
                            "a.IdEvento, a.FechaRecuperacion, a.HoraRecuperacion, a.CuantiaRecup, a.CuantiaOtraRecup, " +
                            "a.CuantiaNeta " + "\n";
                        strQry2 = "FROM Riesgos.Eventos a " + "\n" +
"left JOIN Eventos.Empresa em ON a.IdEmpresa = em.IdEmpresa" + "\n" +
"left JOIN Listas.Usuarios usu ON a.IdUsuario = usu.IdUsuario" + "\n" +
"left JOIN [Parametrizacion].[JerarquiaOrganizacional] PJO ON usu.IdJerarquia  = PJO.idHijo" + "\n" +
"left JOIN [Parametrizacion].[DetalleJerarquiaOrg] PDJO ON PJO.idHijo = PDJO.idHijo " + "\n" +
"left JOIN [Parametrizacion].[Area] PAr ON PDJO.IdArea = PAr.IdArea " + "\n" +
"left JOIN Parametrizacion.Regiones reg ON a.IdRegion = reg.IdRegion " + "\n" +
"left JOIN Parametrizacion.Paises pai ON a.IdRegion = pai.IdRegion AND a.IdPais = pai.IdPais  " + "\n" +
"left JOIN Parametrizacion.Departamentos dep ON a.IdPais = dep.IdPais AND a.IdDepartamento = dep.IdDepartamento " + "\n" +
"left JOIN Parametrizacion.Ciudades ciu ON a.IdDepartamento = ciu.IdDepartamento AND a.IdCiudad = ciu.IdCiudad " + "\n" +
"left JOIN Parametrizacion.OficinaSucursal ofi ON a.IdCiudad = ofi.IdCiudad AND a.IdOficinaSucursal = ofi.IdOficinaSucursal " + "\n" +
"left JOIN Eventos.Servicio ser ON a.IdServicio = ser.IdServicio " + "\n" +
"left JOIN Eventos.SubServicio subs on a.IdSubServicio = subs.IdSubServicio  " + "\n" +
"left JOIN Eventos.Canal can ON a.IdCanal = can.IdCanal " + "\n" +
"left JOIN Eventos.Generador gen ON a.IdGeneraEvento = gen.IdGenerador " + "\n" +
"left JOIN Procesos.CadenaValor cv ON a.IdCadenaValor = cv.IdCadenaValor  " + "\n" +
"left JOIN Procesos.Macroproceso mp ON a.IdCadenaValor = mp.IdCadenaValor AND a.IdMacroproceso = mp.IdMacroProceso " + "\n" +
"left JOIN Procesos.Proceso pp ON a.IdMacroproceso = pp.IdMacroProceso AND a.IdProceso = pp.IdProceso  " + "\n" +
"left JOIN Procesos.Subproceso sp ON a.IdProceso = sp.IdProceso AND a.IdSubProceso = sp.IdSubproceso " + "\n" +
"left JOIN Procesos.Actividad pa ON a.IdSubProceso = pa.IdSubproceso AND a.IdActividad = pa.IdActividad " + "\n" +
"left JOIN Parametrizacion.DetalleJerarquiaOrg jer1 ON a.ResponsableEvento = jer1.idHijo  " + "\n" +
"left JOIN Eventos.Clase ec ON a.IdClase = ec.IdClase " + "\n" +
"left JOIN Eventos.SubClase esc ON a.IdClase = esc.IdClase AND a.IdSubClase = esc.IdSubClase " + "\n" +
"left JOIN Parametrizacion.TipoPerdidaEvento ptp ON a.IdTipoPerdidaEvento = ptp.IdTipoPerdidaEvento " + "\n" +
"left JOIN Eventos.LineaNegocio eln ON a.IdLineaProceso = eln.IdLineaNegocio " + "\n" +
"left JOIN Eventos.SubLineaNegocio esln ON a.IdLineaProceso = esln.IdLineaNegocio AND a.IdSubLineaProceso = esln.IdSubLineaNegocio  " + "\n" +
"left JOIN Eventos.Estado ee ON a.IdEstado = ee.IdEstado  " + "\n" +
"left JOIN Parametrizacion.DetalleJerarquiaOrg jer2 ON a.ResponsableContabilidad = jer2.idHijo ";

                        strConsulta = string.Format("{0} {1} {2} ORDER BY a.CodigoEvento", strQry1, strQry2, condicion);
                        #endregion

                        break;
                    case "3":
                        #region Eventos Vs Plan Accion
                        //dtInformacion = cDataBase.ejecutarConsulta("select a.CodigoEvento, b.Descripcion as Empresa,c.NombreRegion,d.NombrePais,e.NombreDepartamento,f.NombreCiudad,g.NombreOficinaSucursal,a.DetalleUbicacion,a.DescripcionEvento,h.Descripcion as Servicio,i.SubDescripcion as Subservicio,substring(convert(varchar,a.FechaInicio,120),1,10) as FechaInicio,a.HoraInicio,substring(convert(varchar,a.FechaFinalizacion,120),1,10) as FechaFinalizacion,a.HoraFinalizacion,substring(convert(varchar,a.FechaDescubrimiento,120),1,10) as FechaDescubrimiento,a.HoraDescubrimiento,j.Descripcion as Canal,k.Descripcion as GeneradorEvento,l.NombreResponsable as Responsable,a.CuantiaPerdida,m.NombreCadenaValor as CadenaValor,n.Nombre as Macroproceso,o.Nombre as Proceso,p.Nombre as Subproceso,q.Nombre as Actividad,r.Descripcion as ClaseRiesgo,s.SubDescripcion as SubClaseRiesgo,t.NombreTipoPerdidaEvento,a.AfectaContinudad,u.Descripcion as Estado,a.Observaciones,SUBSTRING(convert(varchar,a.FechaEvento,120),1,10) as FechaEvento,j1.NombreResponsable as Originador,j2.NombreHijo as Cargo,j3.NombreHijo as Area, Paa.IdPlanAccion, Paa.DescripcionAccion as PlanAccion, Paa.Responsable as ResponsablePlaAccion, Pad.NombreTipoRecursoPlanAccion,  Paa.ValorRecurso as ValorRecursoPlanAccion, Pab.NombreEstadoPlanAccion as EstadoPlanAccion, substring(CONVERT(varchar, Paa.FechaCompromiso, 102),1,10) AS FechaCompromisoPlanAccion,   Pac.NombreHijo as ResponsablePlanAccion from Riesgos.Eventos a left join Eventos.Empresa b on a.IdEmpresa = b.IdEmpresa left join Parametrizacion.Regiones c on a.IdRegion = c.IdRegion left join Parametrizacion.Paises d on a.IdPais = d.IdPais left join Parametrizacion.Departamentos e on a.IdDepartamento = e.IdDepartamento left join Parametrizacion.Ciudades f on a.IdCiudad = f.IdCiudad left join Parametrizacion.OficinaSucursal g on a.IdOficinaSucursal = g.IdOficinaSucursal left join Eventos.Servicio h on a.IdServicio = h.IdServicio left join Eventos.SubServicio i on a.IdSubServicio = i.IdSubServicio left join Eventos.Canal j on a.IdCanal = j.IdCanal left join Eventos.Generador k on a.IdGeneraEvento = k.IdGenerador left join Parametrizacion.DetalleJerarquiaOrg l on a.ResponsableEvento = l.idHijo left join Procesos.CadenaValor m on a.IdCadenaValor = m.IdCadenaValor left join Procesos.Macroproceso n on a.IdMacroproceso = n.IdMacroProceso left join Procesos.Proceso o on a.IdProceso = o.IdProceso left join Procesos.Subproceso p on a.IdSubProceso = p.IdSubproceso left join Procesos.Actividad q on a.IdActividad = q.IdActividad left join Eventos.Clase r on a.IdClase = r.IdClase left join Eventos.SubClase s on a.IdSubClase = s.IdSubClase left join Parametrizacion.TipoPerdidaEvento t on a.IdTipoPerdidaEvento = t.IdTipoPerdidaEvento left join Eventos.Estado u on a.IdEstado = u.IdEstado left join Listas.Usuarios v on a.IdUsuario = v.IdUsuario left join Parametrizacion.DetalleJerarquiaOrg j1 on v.IdJerarquia = j1.idHijo left join Parametrizacion.JerarquiaOrganizacional j2 on j1.idHijo = j2.idHijo left join Parametrizacion.JerarquiaOrganizacional j3 on j2.idPadre = j3.idHijo right join Riesgos.PlanesAccion Paa on Paa.IdRegistro = a.IdEvento left JOIN Parametrizacion.EstadoPlanAccion Pab ON Paa.IdEstadoPlanAccion = Pab.IdEstadoPlanAccion left JOIN Parametrizacion.JerarquiaOrganizacional Pac ON Paa.Responsable = Pac.idHijo left join Parametrizacion.TipoRecursoPlanAccion Pad on Paa.IdTipoRecursoPlanAccion = Pad.IdTipoRecursoPlanAccion where (Paa.IdControlUsuario = 6) " + condicion + " ORDER BY a.CodigoEvento");
                        //dtInformacion = cDataBase.ejecutarConsulta("SELECT a.CodigoEvento 'Código',em.Descripcion Empresa,reg.NombreRegion 'Región',pai.NombrePais 'Pais',dep.NombreDepartamento Departamento,ciu.NombreCiudad Ciudad,ofi.NombreOficinaSucursal 'Oficina/Sucursal' ,a.DetalleUbicacion 'Detalle Ubicación',a.DescripcionEvento 'Descripción Evento',ser.Descripcion 'Servicio/Producto',subs.SubDescripcion 'SubServicio/SubProducto',SUBSTRING(convert(varchar,a.FechaInicio,120),1,10) 'Fecha Inicio', a.HoraInicio 'Hora Inicio',SUBSTRING(convert(varchar,a.FechaFinalizacion,120),1,10) 'Fecha Finalización',a.HoraFinalizacion 'Hora Finalización',SUBSTRING(convert(varchar,a.FechaDescubrimiento,120),1,10) 'Fecha Descubrimiento',a.HoraDescubrimiento 'Hora Descubrimiento',can.Descripcion Canal,gen.Descripcion 'Generador del Evento',case a.IdGeneraEvento when 2 then a.NomGeneradorEvento else jer.NombreResponsable end as 'Responsable Evento', a.CuantiaPerdida 'Posible Cuantía de Pérdida ',SUBSTRING(convert(varchar,a.FechaEvento,120),1,10) 'Fecha Registro',usu.Usuario Usuario,cv.NombreCadenaValor 'Cadena Valor',mp.Nombre MacroProceso, pp.Nombre Proceso,sp.Nombre SubProceso,pa.Nombre Actividad,jer1.NombreResponsable 'Responsable Solución',ec.Descripcion 'Clase de Riesgo',esc.SubDescripcion 'SubClase de Riesgo',ptp.NombreTipoPerdidaEvento 'Tipo de Pérdida',eln.Descripcion 'Línea Operativa ',esln.SubDescripcion 'SubLínea Operativa',a.MasLineas 'Más Líneas Operativas',case a.AfectaContinudad when 1 then 'Si' when 0 then 'No' end 'Afecta Continuidad',ee.Descripcion Estado,a.Observaciones,jer2.NombreResponsable 'Responsable Contabilidad',a.CuentaPUC 'Cuenta PUC',a.CuentaOrden 'Cuenta de Orden',a.TasaCambio1 'Tasa de Cambio',a.ValorPesos1 'Valor en Pesos',a.ValorRecuperadoTotal 'Valor Recuperado Total',a.TasaCambio2 'Tasa de Cambio 2',a.ValorPesos2 'Valor en Pesos 2',case a.Recuperacion when 1 then 'Si' when 0 then 'No' end 'Recuperación ',a.ValorRecuperacion 'Fuente de la Recuperación',Paa.IdPlanAccion 'Código Plan Acción', Paa.DescripcionAccion as 'Plan Acción', Pac.NombreResponsable 'Responsable Plan Acción', Pad.NombreTipoRecursoPlanAccion 'Tipo Recurso Plan Accion', Paa.ValorRecurso 'Valor Recurso Plan Acción', Pab.NombreEstadoPlanAccion 'Estado Plan Acción', substring(CONVERT(varchar, Paa.FechaCompromiso, 120),1,10) 'Fecha Compromiso Plan Acción' from Riesgos.Eventos a left join Eventos.Empresa em on a.IdEmpresa = em.IdEmpresa left join Parametrizacion.Regiones reg on a.IdRegion = reg.IdRegion left join Parametrizacion.Paises pai on a.IdRegion = pai.IdRegion and a.IdPais = pai.IdPais left join Parametrizacion.Departamentos dep on a.IdPais = dep.IdPais and a.IdDepartamento = dep.IdDepartamento left join Parametrizacion.Ciudades ciu on a.IdDepartamento = ciu.IdDepartamento and a.IdCiudad = ciu.IdCiudad left join Parametrizacion.OficinaSucursal ofi on a.IdCiudad = ofi.IdCiudad and a.IdOficinaSucursal = ofi.IdOficinaSucursal left join Eventos.Servicio ser on a.IdServicio = ser.IdServicio left join Eventos.SubServicio subs on a.IdSubServicio = subs.IdSubServicio left join Eventos.Canal can on a.IdCanal = can.IdCanal left join Eventos.Generador gen on a.IdGeneraEvento = gen.IdGenerador left join Parametrizacion.DetalleJerarquiaOrg jer on a.GeneraEvento = jer.idHijo left join Listas.Usuarios usu on a.IdUsuario = usu.IdUsuario left join Procesos.CadenaValor cv on a.IdCadenaValor = cv.IdCadenaValor left join Procesos.Macroproceso mp on a.IdCadenaValor = mp.IdCadenaValor and a.IdMacroproceso = mp.IdMacroProceso left join Procesos.Proceso pp on a.IdMacroproceso = pp.IdMacroProceso and a.IdProceso = pp.IdProceso left join Procesos.Subproceso sp on a.IdProceso = sp.IdProceso and a.IdSubProceso = sp.IdSubproceso left join Procesos.Actividad pa on a.IdSubProceso = pa.IdSubproceso and a.IdActividad = pa.IdActividad left join Parametrizacion.DetalleJerarquiaOrg jer1 on a.ResponsableEvento = jer1.idHijo left join Eventos.Clase ec on a.IdClase = ec.IdClase left join Eventos.SubClase esc on a.IdClase = esc.IdClase and a.IdSubClase = esc.IdSubClase left join Parametrizacion.TipoPerdidaEvento ptp on a.IdTipoPerdidaEvento = ptp.IdTipoPerdidaEvento left join Eventos.LineaNegocio eln on a.IdLineaProceso = eln.IdLineaNegocio left join Eventos.SubLineaNegocio esln on a.IdLineaProceso = esln.IdLineaNegocio and a.IdSubLineaProceso = esln.IdSubLineaNegocio left join Eventos.Estado ee on a.IdEstado = ee.IdEstado left join Parametrizacion.DetalleJerarquiaOrg jer2 on a.ResponsableContabilidad = jer2.idHijo right join Riesgos.PlanesAccion Paa on Paa.IdRegistro = a.IdEvento  left JOIN Parametrizacion.EstadoPlanAccion Pab ON Paa.IdEstadoPlanAccion = Pab.IdEstadoPlanAccion  left JOIN Parametrizacion.DetalleJerarquiaOrg Pac ON Paa.Responsable = Pac.idHijo  left join Parametrizacion.TipoRecursoPlanAccion Pad on Paa.IdTipoRecursoPlanAccion = Pad.IdTipoRecursoPlanAccion where (Paa.IdControlUsuario = 6) " + condicion + " ORDER BY a.CodigoEvento");

                        strQry1 = "SELECT a.CodigoEvento 'Código', em.Descripcion Empresa, PAr.NombreArea Area, usu.Usuario Usuario, " +
                            "LTRIM(RTRIM(USU.Nombres)) + ' ' + LTRIM(RTRIM(USU.Apellidos)) NombreUsuarioRegistro, " +
                            "SUBSTRING(CONVERT(VARCHAR,a.FechaEvento, 120), 1, 10) 'Fecha Registro', reg.NombreRegion 'Región', pai.NombrePais 'Pais', " +
                            "dep.NombreDepartamento Departamento, ciu.NombreCiudad Ciudad, ofi.NombreOficinaSucursal 'Oficina/Sucursal' , " +
                            "a.DetalleUbicacion 'Detalle Ubicación', a.DescripcionEvento 'Descripción Evento', ser.Descripcion 'Servicio/Producto', " +
                            "subs.SubDescripcion 'SubServicio/SubProducto', SUBSTRING(CONVERT(VARCHAR,a.FechaInicio, 120), 1, 10) 'Fecha Inicio', " +
                            "a.HoraInicio 'Hora Inicio', SUBSTRING(CONVERT(VARCHAR,a.FechaFinalizacion, 120), 1, 10) 'Fecha Finalización', " +
                            "a.HoraFinalizacion 'Hora Finalización', " +
                            "SUBSTRING(CONVERT(VARCHAR,a.FechaDescubrimiento, 120), 1, 10) 'Fecha Descubrimiento', " +
                            "a.HoraDescubrimiento 'Hora Descubrimiento', can.Descripcion Canal,gen.Descripcion 'Generador del Evento', " +
                            "a.NomGeneradorEvento 'Responsable Evento', a.CuantiaPerdida 'Posible Cuantía de Pérdida', " +
                            "cv.NombreCadenaValor 'Cadena Valor',mp.Nombre MacroProceso, pp.Nombre Proceso, sp.Nombre SubProceso, " +
                            "pa.Nombre Actividad, jer1.NombreResponsable 'Responsable Solución', ec.Descripcion 'Clase de Riesgo', " +
                            "esc.SubDescripcion 'SubClase de Riesgo', ptp.NombreTipoPerdidaEvento 'Tipo de Pérdida', " +
                            "eln.Descripcion 'Línea Operativa', esln.SubDescripcion 'SubLínea Operativa', a.MasLineas 'Más Líneas Operativas', " +
                            "CASE a.AfectaContinudad WHEN 1 THEN 'Si' WHEN 0 THEN 'No' END 'Afecta Continuidad', ee.Descripcion Estado, " +
                            "a.Observaciones, jer2.NombreResponsable 'Responsable Contabilidad', a.CuentaPUC 'Cuenta PUC', " +
                            "a.CuentaOrden 'Cuenta de Orden', a.TasaCambio1 'Tasa de Cambio', a.ValorPesos1 'Valor en Pesos', " +
                            "a.ValorRecuperadoTotal 'Valor Recuperado Total', a.TasaCambio2 'Tasa de Cambio 2', a.ValorPesos2 'Valor en Pesos 2', " +
                            "CASE a.Recuperacion WHEN 1 THEN 'Si' WHEN 0 THEN 'No' END 'Recuperación', " +
                            "a.ValorRecuperacion 'Fuente de la Recuperación', Paa.IdPlanAccion 'Código Plan Acción', " +
                            "Paa.DescripcionAccion 'Plan Acción', Pac.NombreResponsable 'Responsable Plan Acción', " +
                            "Pad.NombreTipoRecursoPlanAccion 'Tipo Recurso Plan Accion', Paa.ValorRecurso 'Valor Recurso Plan Acción', " +
                            "Pab.NombreEstadoPlanAccion 'Estado Plan Acción', " +
                            "SUBSTRING(CONVERT(VARCHAR, Paa.FechaCompromiso, 120), 1, 10) 'Fecha Compromiso Plan Acción' , " +
                            "[NombreImpactoCualitativo], a.FechaRecuperacion, a.HoraRecuperacion, a.CuantiaRecup, a.CuantiaOtraRecup, " +
                            "a.CuantiaNeta ";
                        strQry2 = "FROM Riesgos.Eventos a LEFT JOIN Eventos.Empresa em ON a.IdEmpresa = em.IdEmpresa LEFT JOIN Parametrizacion.Regiones reg ON a.IdRegion = reg.IdRegion LEFT JOIN Parametrizacion.Paises pai ON a.IdRegion = pai.IdRegion AND a.IdPais = pai.IdPais LEFT JOIN Parametrizacion.Departamentos dep ON a.IdPais = dep.IdPais AND a.IdDepartamento = dep.IdDepartamento LEFT JOIN Parametrizacion.Ciudades ciu on a.IdDepartamento = ciu.IdDepartamento AND a.IdCiudad = ciu.IdCiudad LEFT JOIN Parametrizacion.OficinaSucursal ofi ON a.IdCiudad = ofi.IdCiudad AND a.IdOficinaSucursal = ofi.IdOficinaSucursal LEFT JOIN Eventos.Servicio ser ON a.IdServicio = ser.IdServicio LEFT JOIN Eventos.SubServicio subs ON a.IdSubServicio = subs.IdSubServicio LEFT JOIN Eventos.Canal can ON a.IdCanal = can.IdCanal LEFT JOIN Eventos.Generador gen ON a.IdGeneraEvento = gen.IdGenerador LEFT JOIN Parametrizacion.DetalleJerarquiaOrg jer ON a.GeneraEvento = jer.idHijo LEFT JOIN Listas.Usuarios usu ON a.IdUsuario = usu.IdUsuario LEFT JOIN Procesos.CadenaValor cv ON a.IdCadenaValor = cv.IdCadenaValor LEFT JOIN Procesos.Macroproceso mp ON a.IdCadenaValor = mp.IdCadenaValor AND a.IdMacroproceso = mp.IdMacroProceso LEFT JOIN Procesos.Proceso pp ON a.IdMacroproceso = pp.IdMacroProceso AND a.IdProceso = pp.IdProceso LEFT JOIN Procesos.Subproceso sp ON a.IdProceso = sp.IdProceso AND a.IdSubProceso = sp.IdSubproceso LEFT JOIN Procesos.Actividad pa ON a.IdSubProceso = pa.IdSubproceso AND a.IdActividad = pa.IdActividad LEFT JOIN Parametrizacion.DetalleJerarquiaOrg jer1 on a.ResponsableEvento = jer1.idHijo LEFT JOIN Eventos.Clase ec ON a.IdClase = ec.IdClase LEFT JOIN Eventos.SubClase esc ON a.IdClase = esc.IdClase AND a.IdSubClase = esc.IdSubClase LEFT JOIN Parametrizacion.TipoPerdidaEvento ptp ON a.IdTipoPerdidaEvento = ptp.IdTipoPerdidaEvento LEFT JOIN Eventos.LineaNegocio eln ON a.IdLineaProceso = eln.IdLineaNegocio LEFT JOIN Eventos.SubLineaNegocio esln ON a.IdLineaProceso = esln.IdLineaNegocio AND a.IdSubLineaProceso = esln.IdSubLineaNegocio LEFT JOIN Eventos.Estado ee ON a.IdEstado = ee.IdEstado LEFT JOIN Parametrizacion.DetalleJerarquiaOrg jer2 ON a.ResponsableContabilidad = jer2.idHijo RIGHT JOIN Riesgos.PlanesAccion Paa ON Paa.IdRegistro = a.IdEvento LEFT JOIN Parametrizacion.EstadoPlanAccion Pab ON Paa.IdEstadoPlanAccion = Pab.IdEstadoPlanAccion LEFT JOIN Parametrizacion.DetalleJerarquiaOrg Pac ON Paa.Responsable = Pac.idHijo LEFT JOIN Parametrizacion.TipoRecursoPlanAccion Pad ON Paa.IdTipoRecursoPlanAccion = Pad.IdTipoRecursoPlanAccion LEFT JOIN [Parametrizacion].[JerarquiaOrganizacional] PJO ON usu.IdJerarquia  = PJO.idHijo LEFT JOIN [Parametrizacion].[DetalleJerarquiaOrg] PDJO ON PJO.idHijo = PDJO.idHijo LEFT JOIN [Parametrizacion].[Area] PAr ON PDJO.IdArea = PAr.IdArea";
                        strConsulta = string.Format("{0} {1} WHERE (Paa.IdControlUsuario = 6) {2} ORDER BY a.CodigoEvento", strQry1, strQry2, condicion);
                        #endregion

                        break;
                    case "4":
                        #region Sin Reporte
                        //dtInformacion = cDataBase.ejecutarConsulta("select a.IdUsuario,a.IdJerarquia,b.NombreResponsable,c.NombreHijo  as AreaSinReporte from Listas.Usuarios a,Parametrizacion.DetalleJerarquiaOrg b,Parametrizacion.JerarquiaOrganizacional c where a.IdJerarquia = b.idHijo and b.idHijo = c.idHijo and a.IdUsuario not in ( select distinct a.IdUsuario from Riesgos.NoHuboEventos a " + condicion + " union select distinct a.idusuario from Riesgos.Eventos a " + condicion + " ) order by a.IdJerarquia");

                        strQry1 = "SELECT PAP.NombreArea Area FROM [Parametrizacion].[Area] PAP WHERE NOT EXISTS";
                        strQry2 = string.Format("(SELECT DISTINCT PA.IdArea FROM [Riesgos].[NoHuboEventos] RNHE LEFT JOIN Listas.Usuarios LU ON RNHE.IdUsuario = LU.IdUsuario LEFT JOIN [Parametrizacion].[JerarquiaOrganizacional] PJO ON LU.IdJerarquia  = PJO.idHijo LEFT JOIN [Parametrizacion].[DetalleJerarquiaOrg] PDJO ON PJO.idHijo = PDJO.idHijo LEFT JOIN [Parametrizacion].[Area] PA ON PDJO.IdArea = PA.IdArea WHERE PDJO.IdArea IS NOT NULL AND PA.IdArea = PAP.IdArea {0} {2} UNION SELECT DISTINCT PA.IdArea FROM [Riesgos].[Eventos] RE LEFT JOIN Listas.Usuarios LU ON RE.IdUsuario = LU.IdUsuario LEFT JOIN [Parametrizacion].[JerarquiaOrganizacional] PJO ON LU.IdJerarquia  = PJO.idHijo LEFT JOIN [Parametrizacion].[DetalleJerarquiaOrg] PDJO ON PJO.idHijo = PDJO.idHijo LEFT JOIN [Parametrizacion].[Area] PA ON PDJO.IdArea = PA.IdArea WHERE PDJO.IdArea IS NOT NULL AND PA.IdArea = PAP.IdArea {1} {2})", strCondRNHE, strCondRE, condicion);

                        strConsulta = string.Format("{0} {1}", strQry1, strQry2);
                        #endregion

                        break;
                }

                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta(strConsulta);
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public void agregarNHEvento(ref string CodigoEvento, String IdEmpresa)
        {
            string strCodigoEvento = string.Empty, strPrefixEvento = string.Empty;
            string strConsulta = string.Empty, strCamposInsert = string.Empty, strValoresInsert = string.Empty;

            try
            {
                switch (IdEmpresa.Trim())
                {
                    case "1":
                        strPrefixEvento = "NEV";
                        break;
                    case "2":
                        strPrefixEvento = "NEG";
                        break;
                    case "3":
                        strPrefixEvento = "NEE";
                        break;
                }

                strCodigoEvento = string.Format("(CASE WHEN (SELECT MAX(CAST(SUBSTRING(CodigoEvento, 4, 10) AS INT)) + 1 FROM Riesgos.NoHuboEventos WHERE CodigoEvento LIKE '{0}%')IS NULL THEN '{0}1' ELSE (SELECT '{0}'+ CAST ((SELECT MAX(CAST(SUBSTRING(CodigoEvento, 4, 10) AS INT)) + 1 FROM Riesgos.NoHuboEventos WHERE CodigoEvento LIKE '{0}%') AS NVARCHAR(50))) END )", strPrefixEvento);
                strCamposInsert = "(CodigoEvento, IdEmpresa, FechaEvento, IdUsuario)";
                strValoresInsert = "(" + strCodigoEvento + ", '" + IdEmpresa + "', GETDATE(), " + IdUsuario + ")";
                strConsulta = string.Format("INSERT INTO Riesgos.NoHuboEventos {0} VALUES {1}", strCamposInsert, strValoresInsert);

                cDataBase.conectar();
                cDataBase.ejecutarQuery(strConsulta);
                cDataBase.desconectar();

                #region Traer Valor Evento
                DataTable dtInformacion = new DataTable();
                strConsulta = string.Format("SELECT MAX(CAST(SUBSTRING(CodigoEvento, 4, 10) AS INT)) AS LastEvent FROM Riesgos.NoHuboEventos WHERE CodigoEvento LIKE '{0}%'", strPrefixEvento);
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta(strConsulta);
                cDataBase.desconectar();

                CodigoEvento = strPrefixEvento + dtInformacion.Rows[0]["LastEvent"].ToString().Trim();
                #endregion Traer Valor Evento
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public void agregarEvento(ref string CodigoEvento, String IdEmpresa, String IdRegion, String IdPais,
            String IdDepartamento, String IdCiudad, String IdOficinaSucursal, String DetalleUbicacion,
            String DescripcionEvento, String IdServicio, String IdSubServicio, String FechaInicio,
            String HoraInicio, String FechaFinalizacion, String HoraFinalizacion, String FechaDescubrimiento,
            String HoraDescubrimiento, String IdCanal, String IdGeneraEvento, String GeneraEvento,
            String cuantiaperdida, String FechaEvento, String NomGeneradorEvento, int IdUsuarioEventos, String NombreImpCual,
            string FechaRecuperacion, double cuantiaRecuperadaSeguros, double cuantiaOtrasRecuperaciones, double cuantiaNetaRecuperaciones)
        {
            #region Variables

            DataTable dtInformacion = new DataTable();
            string strCodigoEvento = string.Empty, strPrefixEvento = string.Empty;
            string strConsulta = string.Empty, strCamposInsert = string.Empty, strValoresInsert = string.Empty;
            #endregion Variables

            try
            {
                switch (IdEmpresa.Trim())
                {
                    case "1":
                        strPrefixEvento = "EV";
                        break;
                    case "2":
                        strPrefixEvento = "EG";
                        break;
                    case "3":
                        strPrefixEvento = "EE";
                        break;
                }

                lock (thisLock)
                {
                    strCodigoEvento = string.Format("(CASE WHEN (SELECT MAX(CAST(SUBSTRING(CodigoEvento, 3, 10)AS INT)) + 1 FROM Riesgos.Eventos WHERE CodigoEvento LIKE '{0}%')IS NULL THEN '{0}1' ELSE (SELECT '{0}'+ CAST ((SELECT MAX(CAST(SUBSTRING(CodigoEvento, 3, 10)AS INT)) + 1 FROM Riesgos.Eventos WHERE CodigoEvento LIKE '{0}%') AS NVARCHAR(50))) END )", strPrefixEvento);
                    strCamposInsert = "(CodigoEvento, IdEmpresa, IdRegion, IdPais, IdDepartamento, IdCiudad, IdOficinaSucursal, DetalleUbicacion, DescripcionEvento, IdServicio, IdSubServicio, FechaInicio, HoraInicio, FechaFinalizacion, HoraFinalizacion, FechaDescubrimiento, HoraDescubrimiento, IdCanal, IdGeneraEvento, GeneraEvento, cuantiaperdida, FechaEvento, NomGeneradorEvento, IdUsuario, NombreImpactoCualitativo,fechaRecuperacion,cuantiaRecuperadaSeguros,cuantiaOtrasRecuperaciones,cuantiaNetaRecuperaciones)";
                    strValoresInsert = "(" + strCodigoEvento + ",'" + IdEmpresa + "','" + IdRegion + "','" + IdPais + "','" + IdDepartamento + "','" + IdCiudad + "','" + IdOficinaSucursal + "','" + DetalleUbicacion + "','" + DescripcionEvento + "','" + IdServicio + "','" + IdSubServicio + "',CONVERT(datetime, '" + FechaInicio + "', 120),'" + HoraInicio + "'," + FechaFinalizacion + "," + HoraFinalizacion + ",CONVERT(datetime, '" + FechaDescubrimiento + "', 120),'" + HoraDescubrimiento + "','" + IdCanal + "','" + IdGeneraEvento + "','" + GeneraEvento + "','" + cuantiaperdida + "',CONVERT(datetime, '" + FechaEvento + "', 120),'" + NomGeneradorEvento + "','" + IdUsuarioEventos + "','" + NombreImpCual + ",CONVERT(datetime, '" + FechaRecuperacion + "', 120),'"+cuantiaRecuperadaSeguros+","+cuantiaOtrasRecuperaciones+","+cuantiaNetaRecuperaciones + "')";
                    strConsulta = string.Format("INSERT INTO Riesgos.Eventos {0} VALUES {1}", strCamposInsert, strValoresInsert);

                    cDataBase.conectar();
                    cDataBase.ejecutarQuery(strConsulta);
                    cDataBase.desconectar();

                    #region Traer Valor Evento
                    strConsulta = string.Format("SELECT MAX(CAST(SUBSTRING(CodigoEvento, 3, 10) AS INT)) AS LastEvent FROM Riesgos.Eventos WHERE CodigoEvento LIKE '{0}%'", strPrefixEvento);
                    cDataBase.conectar();
                    dtInformacion = cDataBase.ejecutarConsulta(strConsulta);
                    cDataBase.desconectar();

                    CodigoEvento = strPrefixEvento + dtInformacion.Rows[0]["LastEvent"].ToString().Trim();
                    #endregion Traer Valor Evento
                }
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public void modificarEvento(String IdCadenaValor, String IdMacroproceso, String IdProceso, String IdSubProceso, String IdActividad, String ResponsableEvento, String IdClase, String IdSubClase, String IdTipoPerdidaEvento, String AfectaContinudad, String IdEstado, String Observaciones, String IdEvento, String IdLineaProceso, String IdSubLineaProceso, String MasLineas)
        {
            try
            {
                cDataBase.conectar();
                cDataBase.ejecutarQuery("UPDATE Riesgos.Eventos SET IdCadenaValor='" + IdCadenaValor + "',IdMacroproceso = '" + IdMacroproceso + "',IdProceso = '" + IdProceso + "', IdSubProceso  ='" + IdSubProceso + "',IdActividad  ='" + IdActividad + "',ResponsableEvento='" + ResponsableEvento + "',IdClase ='" + IdClase + "',IdSubClase='" + IdSubClase + "', IdTipoPerdidaEvento='" + IdTipoPerdidaEvento + "',AfectaContinudad='" + AfectaContinudad + "',IdEstado='" + IdEstado + "',Observaciones='" + Observaciones + "',IdLineaProceso=" + IdLineaProceso + ",IdSubLineaProceso=" + IdSubLineaProceso + ",MasLineas='" + MasLineas + "' WHERE (CodigoEvento = '" + IdEvento + "')");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public DataTable loadInfoEventos(String CodigoEvento, String DescripcionEvento, String IdCadenaValor,
            String IdMacroproceso, String IdProceso, String IdSubProceso, int IdUsuarioJerarquia, int IdSesion, string TipoArea)
        {
            DataTable dtInformacion = new DataTable();
            string condicion = string.Empty;
            string CondicionOtroFiltro = string.Empty;
            try
            {                        
                    #region Filtros
                    if (CodigoEvento != "")
                        condicion = "AND (a.CodigoEvento = '" + CodigoEvento + "') ";
                    if (DescripcionEvento != "")
                        condicion += "AND (a.DescripcionEvento LIKE '%" + DescripcionEvento + "%') ";
                    if (IdCadenaValor != "---")
                        condicion += "AND (a.IdCadenaValor = " + IdCadenaValor + ") ";
                    if (IdMacroproceso != "---")
                        condicion += "AND ( a.IdMacroproceso = " + IdMacroproceso + " ) ";
                    if (IdProceso != "---")
                        condicion += "AND (a.IdProceso = " + IdProceso + ") ";
                    if (IdSubProceso != "---")
                        condicion += "AND (a.IdSubProceso = " + IdSubProceso + ") ";
               
                    #endregion Filtros
                    CondicionOtroFiltro= "SELECT a.IdEvento, \n"
           + "       a.CodigoEvento, \n"
           + "       a.IdEmpresa, \n"
           + "       a.IdRegion, \n"
           + "       a.IdPais, \n"
           + "       a.IdDepartamento, \n"
           + "       a.IdCiudad, \n"
           + "       a.IdOficinaSucursal, \n"
           + "       a.DetalleUbicacion, \n"
           + "       a.DescripcionEvento, \n"
           + "       a.IdServicio, \n"
           + "       a.IdSubServicio, \n"
           + "       SUBSTRING(CONVERT(VARCHAR, a.FechaInicio, 120), 1, 10) AS FechaInicio, \n"
           + "       SUBSTRING(a.HoraInicio, 1, 2) AS HorI, \n"
           + "       SUBSTRING(a.HoraInicio, 4, 2) AS MinI, \n"
           + "       SUBSTRING(a.HoraInicio, 7, 3) AS amI, \n"
           + "       SUBSTRING(CONVERT(VARCHAR, a.FechaFinalizacion, 120), 1, 10) AS FechaFinalizacion, \n"
           + "       SUBSTRING(a.HoraFinalizacion, 1, 2) AS HorF, \n"
           + "       SUBSTRING(a.HoraFinalizacion, 4, 2) AS MinF, \n"
           + "       SUBSTRING(a.HoraFinalizacion, 7, 3) AS amF, \n"
           + "       SUBSTRING(CONVERT(VARCHAR, a.FechaDescubrimiento, 120), 1, 10) AS FechaDescubrimiento, \n"
           + "       SUBSTRING(a.HoraDescubrimiento, 1, 2) AS HorD, \n"
           + "       SUBSTRING(a.HoraDescubrimiento, 4, 2) AS MinD, \n"
           + "       SUBSTRING(a.HoraDescubrimiento, 7, 3) AS amD, \n"
           + "       a.IdCanal, \n"
           + "       a.IdGeneraEvento, \n"
           + "       a.GeneraEvento, \n"
           + "       j1.NombreResponsable AS GeneradorEvento, \n"
           + "       a.cuantiaperdida, \n"
           + "       SUBSTRING(CONVERT(VARCHAR, a.FechaEvento, 120), 1, 10) AS FechaEvento, \n"
           + "       u.Usuario, \n"
           + "       a.IdCadenaValor, \n"
           + "       a.IdMacroproceso, \n"
           + "       a.IdProceso, \n"
           + "       a.IdSubProceso, \n"
           + "       a.IdActividad, \n"
           + "       a.ResponsableEvento, \n"
           + "       j2.NombreResponsable AS ResponsableSolucion, \n"
           + "       a.IdClase, \n"
           + "       b.Descripcion AS NombreClaseEvento, \n"
           + "       a.IdSubClase, \n"
           + "       c.NombreTipoPerdidaEvento, \n"
           + "       a.AfectaContinudad, \n"
           + "       a.IdEstado, \n"
           + "       a.Observaciones, \n"
           + "       a.ResponsableContabilidad, \n"
           + "       j3.NombreResponsable AS NombreResContabilidad, \n"
           + "       a.CuentaPUC, \n"
           + "       a.CuentaOrden, \n"
           + "       a.CuentaPerdida, \n"
           + "       a.Moneda1, \n"
           + "       a.TasaCambio1, \n"
           + "       a.ValorPesos1, \n"
           + "       a.ValorRecuperadoTotal, \n"
           + "       a.Moneda2, \n"
           + "       a.TasaCambio2, \n"
           + "       a.ValorPesos2, \n"
           + "       a.ValorRecuperadoSeguro, \n"
           + "       a.ValorPesos3, \n"
           + "       a.Recuperacion, \n"
           + "       a.ValorRecuperacion, \n"
           + "       a.IdLineaProceso, \n"
           + "       a.IdSubLineaProceso, \n"
           + "       a.MasLineas, \n"
           + "       a.NomGeneradorEvento, \n"
           + "       SUBSTRING(CONVERT(VARCHAR, a.FechaContabilidad, 120), 1, 10) AS FechaContab, \n"
           + "       SUBSTRING(a.HoraContabilidad, 1, 2) AS HoraContab, \n"
           + "       SUBSTRING(a.HoraContabilidad, 4, 2) AS MinContab, \n"
           + "       SUBSTRING(a.HoraContabilidad, 7, 3) AS amContab,\n"
           + "       ISNULL(a.NombreImpactoCualitativo,'') as NombreImpactoCualitativo, \n"
           + "       SUBSTRING(CONVERT(VARCHAR, a.FechaRecuperacion, 120), 1, 10) AS FechaRecup, \n"
           + "       SUBSTRING(a.HoraRecuperacion, 1, 2) AS HoraRecup, \n"
           + "       SUBSTRING(a.HoraRecuperacion, 4, 2) AS MinRecup, \n"
           + "       SUBSTRING(a.HoraRecuperacion, 7, 3) AS amRecup, \n"
           + "       a.CuantiaRecup, \n"
           + "       a.CuantiaOtraRecup, \n" 
           + "       a.CuantiaNeta \n"
          // + "a.fechaRecuperacion,a.cuantiaRecuperadaSeguros,a.cuantiaOtrasRecuperaciones,a.cuantiaNetaRecuperaciones \n"
           + "FROM Riesgos.Eventos a WITH	(NOLOCK)\n"
           + "     LEFT JOIN Eventos.Clase b ON a.IdClase = b.IdClase\n"
           + "     LEFT JOIN Parametrizacion.TipoPerdidaEvento c WITH	(NOLOCK) ON a.IdTipoPerdidaEvento = c.IdTipoPerdidaEvento\n"
           + "     LEFT JOIN Listas.Usuarios u WITH	(NOLOCK) ON a.IdUsuario = u.IdUsuario\n"
           + "     LEFT JOIN Parametrizacion.DetalleJerarquiaOrg j1 WITH	(NOLOCK) ON j1.idHijo = a.GeneraEvento\n"
           + "     LEFT JOIN Parametrizacion.DetalleJerarquiaOrg j2 WITH	(NOLOCK) ON j2.idHijo = a.ResponsableEvento\n"
           + "     LEFT JOIN Parametrizacion.DetalleJerarquiaOrg j3 WITH	(NOLOCK) ON a.ResponsableContabilidad = j3.idHijo  " +
           "       LEFT JOIN Parametrizacion.JerarquiaOrganizacional jo ON a.GeneraEvento = jo.idHijo ";

                if (TipoArea == "R")
                { CondicionOtroFiltro += "WHERE(a.CodigoEvento <> '')" + condicion; }

                else
                { CondicionOtroFiltro += "WHERE  (a.CodigoEvento <> '') AND ((a.IdUsuario = " + IdUsuario + ") OR (a.GeneraEvento  = " + IdUsuarioJerarquia + "))" + condicion; }


                #region consulta
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta(CondicionOtroFiltro);
                #endregion Consulta

                cDataBase.desconectar();
                return dtInformacion;
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public DataTable loadInfoEventosPlan(String CodigoEvento, String DescripcionEvento, String IdCadenaValor,
          String IdMacroproceso, String IdProceso, String IdSubProceso, int IdUsuarioJerarquia)
        {
            DataTable dtInformacion = new DataTable();

            try
            {
                cDataBase.conectar();
                List<SqlParameter> parametros = new List<SqlParameter>()
                {
                    new SqlParameter() { ParameterName = "@CodigoEvento", SqlDbType = SqlDbType.VarChar, Value =  CodigoEvento }
                };
                dtInformacion = cDataBase.EjecutarSPParametrosReturnDatatable("[Eventos].[SeleccionarEventoPlan]", parametros);
                cDataBase.desconectar();
                return dtInformacion;
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        // Yoendy

        public System.Data.DataSet VerificaParametrizacion(int idRiesgo, int codFrecuencia)
        {
            System.Data.DataSet dsInformacion = new System.Data.DataSet();

            try
            {
                cDataBase.conectar();
                List<SqlParameter> parametros = new List<SqlParameter>()
                {
                    new SqlParameter() { ParameterName = "@idRiesgo", SqlDbType = SqlDbType.Int, Value =  idRiesgo },
                    new SqlParameter() { ParameterName = "@codFrecuencia", SqlDbType = SqlDbType.Int, Value =  codFrecuencia },
                };
                dsInformacion = cDataBase.EjecutarSp("[dbo].[pa_VerificarParametrizacion]", parametros);
                cDataBase.desconectar();
                return dsInformacion;
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public DataTable loadInfoEventosvsUsuario(int IdUsuarioJerarquia)
        {
            DataTable dtInformacion = new DataTable();
            string condicion = string.Empty;
            string strConsulta = string.Empty, strSelect = string.Empty, strFrom = string.Empty, strRiesgo = string.Empty;

            try
            {
                #region ValidacionAreaRiesgo
                cDataBase.conectar();
                string Area = string.Empty;
                strRiesgo = string.Format("SELECT [TipoArea] FROM [Parametrizacion].[JerarquiaOrganizacional] where idHijo = {0}", IdUsuarioJerarquia);
                dtInformacion = cDataBase.ejecutarConsulta(strRiesgo);
                if (dtInformacion.Rows.Count > 0)
                {
                    for (int rows = 0; rows < dtInformacion.Rows.Count; rows++)
                    {
                        Area = dtInformacion.Rows[rows]["TipoArea"].ToString().Trim();
                    }
                }
                if (Area != "R")
                {
                    #region JerarquiaOrganizacional y UsuarioRegistro
                    //string condicion2 = condicion.Replace("WHERE", "AND");
                    condicion += strConsulta = string.Format(" WHERE ([ResponsableEvento] = {0} or a.IdUsuario = {1})", IdUsuarioJerarquia, Session["IdUsuario"].ToString());
                    #endregion JerarquiaOrganizacional y UsuarioRegistro
                }
                cDataBase.desconectar();
                #endregion ValidacionAreaRiesgo



                cDataBase.conectar();
                //23-01-2014
                //dtInformacion = cDataBase.ejecutarConsulta("select a.IdEvento,a.CodigoEvento,a.IdEmpresa,a.IdRegion,a.IdPais,a.IdDepartamento,a.IdCiudad,a.IdOficinaSucursal,a.DetalleUbicacion,a.DescripcionEvento,a.IdServicio,a.IdSubServicio,substring(convert(varchar,a.FechaInicio,120),1,10) as FechaInicio,a.HoraInicio,substring(convert(varchar,a.FechaFinalizacion,120),1,10) as FechaFinalizacion,a.HoraFinalizacion,substring(convert(varchar,a.FechaDescubrimiento,120),1,10) as FechaDescubrimiento,a.HoraDescubrimiento,a.IdCanal,a.IdGeneraEvento,a.GeneraEvento, j1.NombreResponsable as GeneradorEvento,a.cuantiaperdida,substring(convert(varchar,a.FechaEvento,120),1,10) as FechaEvento,u.Usuario,a.IdCadenaValor,a.IdMacroproceso,a.IdProceso,a.IdSubProceso,a.IdActividad,a.ResponsableEvento,j2.NombreResponsable as ResponsableSolucion,a.IdClase,b.Descripcion as NombreClaseEvento,a.IdSubClase,c.NombreTipoPerdidaEvento,a.AfectaContinudad,a.IdEstado,a.Observaciones,a.ResponsableContabilidad,j3.NombreResponsable as NombreResContabilidad,a.CuentaPUC,a.CuentaOrden,a.CuentaPerdida,a.Moneda1,a.TasaCambio1,a.ValorPesos1,a.ValorRecuperadoTotal,a.Moneda2,a.TasaCambio2,a.ValorPesos2,a.ValorRecuperadoSeguro,a.ValorPesos3,a.Recuperacion,a.ValorRecuperacion,a.IdLineaProceso,a.IdSubLineaProceso,a.MasLineas FROM Riesgos.Eventos a left join Eventos.Clase b on a.IdClase = b.IdClase left join Parametrizacion.TipoPerdidaEvento c on a.IdTipoPerdidaEvento = c.IdTipoPerdidaEvento left join Listas.Usuarios u on a.IdUsuario = u.IdUsuario left join Parametrizacion.DetalleJerarquiaOrg j1 on j1.idHijo = a.GeneraEvento left join Parametrizacion.DetalleJerarquiaOrg j2 on j2.idHijo = a.ResponsableEvento left join Parametrizacion.DetalleJerarquiaOrg j3 on a.ResponsableContabilidad = j3.idHijo " + condicion + " ORDER BY a.IdEvento");
                strSelect = "SELECT a.IdEvento,a.CodigoEvento,a.IdEmpresa,a.IdRegion,a.IdPais,a.IdDepartamento,a.IdCiudad,a.IdOficinaSucursal,a.DetalleUbicacion,a.DescripcionEvento,a.IdServicio,a.IdSubServicio,substring(convert(varchar,a.FechaInicio,120),1,10) as FechaInicio,substring(a.HoraInicio,1,2) as HorI,substring(a.HoraInicio,4,2) as MinI,substring(a.HoraInicio,7,3) as amI,substring(convert(varchar,a.FechaFinalizacion,120),1,10) as FechaFinalizacion,substring(a.HoraFinalizacion,1,2) as HorF,substring(a.HoraFinalizacion,4,2) as MinF,substring(a.HoraFinalizacion,7,3) as amF,substring(convert(varchar,a.FechaDescubrimiento,120),1,10) as FechaDescubrimiento,substring(a.HoraDescubrimiento,1,2) as HorD,substring(a.HoraDescubrimiento,4,2) as MinD,substring(a.HoraDescubrimiento,7,3) as amD,a.IdCanal,a.IdGeneraEvento,a.GeneraEvento, j1.NombreResponsable as GeneradorEvento,a.cuantiaperdida,substring(convert(varchar,a.FechaEvento,120),1,10) as FechaEvento,u.Usuario,a.IdCadenaValor,a.IdMacroproceso,a.IdProceso,a.IdSubProceso,a.IdActividad,a.ResponsableEvento,j2.NombreResponsable as ResponsableSolucion,a.IdClase,b.Descripcion as NombreClaseEvento,a.IdSubClase,c.NombreTipoPerdidaEvento,a.AfectaContinudad,a.IdEstado,a.Observaciones,a.ResponsableContabilidad,j3.NombreResponsable as NombreResContabilidad,a.CuentaPUC,a.CuentaOrden,a.CuentaPerdida,a.Moneda1,a.TasaCambio1,a.ValorPesos1,a.ValorRecuperadoTotal,a.Moneda2,a.TasaCambio2,a.ValorPesos2,a.ValorRecuperadoSeguro,a.ValorPesos3,a.Recuperacion,a.ValorRecuperacion,a.IdLineaProceso,a.IdSubLineaProceso,a.MasLineas,a.NomGeneradorEvento, SUBSTRING(CONVERT(VARCHAR,a.FechaContabilidad,120),1,10) AS FechaContab,SUBSTRING(a.HoraContabilidad,1,2) as HoraContab,SUBSTRING(a.HoraContabilidad,4,2) AS MinContab,SUBSTRING(a.HoraContabilidad,7,3) AS amContab";
                strFrom = "FROM Riesgos.Eventos a LEFT JOIN Eventos.Clase b on a.IdClase = b.IdClase LEFT JOIN Parametrizacion.TipoPerdidaEvento c on a.IdTipoPerdidaEvento = c.IdTipoPerdidaEvento LEFT JOIN Listas.Usuarios u on a.IdUsuario = u.IdUsuario LEFT JOIN Parametrizacion.DetalleJerarquiaOrg j1 on j1.idHijo = a.GeneraEvento LEFT JOIN Parametrizacion.DetalleJerarquiaOrg j2 on j2.idHijo = a.ResponsableEvento LEFT JOIN Parametrizacion.DetalleJerarquiaOrg j3 on a.ResponsableContabilidad = j3.idHijo";
                //strConsulta = string.Format("{0} {1} {2} ORDER BY a.IdEvento", strSelect, strFrom, condicion);
                strConsulta = string.Format("{0} {1} {2} ORDER BY a.IdEvento", strSelect, strFrom, condicion);
                //dtInformacion = cDataBase.ejecutarConsulta("select a.IdEvento,a.CodigoEvento,a.IdEmpresa,a.IdRegion,a.IdPais,a.IdDepartamento,a.IdCiudad,a.IdOficinaSucursal,a.DetalleUbicacion,a.DescripcionEvento,a.IdServicio,a.IdSubServicio,substring(convert(varchar,a.FechaInicio,120),1,10) as FechaInicio,substring(a.HoraInicio,1,2) as HorI,substring(a.HoraInicio,4,2) as MinI,substring(a.HoraInicio,7,3) as amI,substring(convert(varchar,a.FechaFinalizacion,120),1,10) as FechaFinalizacion,substring(a.HoraFinalizacion,1,2) as HorF,substring(a.HoraFinalizacion,4,2) as MinF,substring(a.HoraFinalizacion,7,3) as amF,substring(convert(varchar,a.FechaDescubrimiento,120),1,10) as FechaDescubrimiento,substring(a.HoraDescubrimiento,1,2) as HorD,substring(a.HoraDescubrimiento,4,2) as MinD,substring(a.HoraDescubrimiento,7,3) as amD,a.IdCanal,a.IdGeneraEvento,a.GeneraEvento, j1.NombreResponsable as GeneradorEvento,a.cuantiaperdida,substring(convert(varchar,a.FechaEvento,120),1,10) as FechaEvento,u.Usuario,a.IdCadenaValor,a.IdMacroproceso,a.IdProceso,a.IdSubProceso,a.IdActividad,a.ResponsableEvento,j2.NombreResponsable as ResponsableSolucion,a.IdClase,b.Descripcion as NombreClaseEvento,a.IdSubClase,c.NombreTipoPerdidaEvento,a.AfectaContinudad,a.IdEstado,a.Observaciones,a.ResponsableContabilidad,j3.NombreResponsable as NombreResContabilidad,a.CuentaPUC,a.CuentaOrden,a.CuentaPerdida,a.Moneda1,a.TasaCambio1,a.ValorPesos1,a.ValorRecuperadoTotal,a.Moneda2,a.TasaCambio2,a.ValorPesos2,a.ValorRecuperadoSeguro,a.ValorPesos3,a.Recuperacion,a.ValorRecuperacion,a.IdLineaProceso,a.IdSubLineaProceso,a.MasLineas,a.NomGeneradorEvento FROM Riesgos.Eventos a left join Eventos.Clase b on a.IdClase = b.IdClase left join Parametrizacion.TipoPerdidaEvento c on a.IdTipoPerdidaEvento = c.IdTipoPerdidaEvento left join Listas.Usuarios u on a.IdUsuario = u.IdUsuario left join Parametrizacion.DetalleJerarquiaOrg j1 on j1.idHijo = a.GeneraEvento left join Parametrizacion.DetalleJerarquiaOrg j2 on j2.idHijo = a.ResponsableEvento left join Parametrizacion.DetalleJerarquiaOrg j3 on a.ResponsableContabilidad = j3.idHijo " + condicion + " ORDER BY a.IdEvento");
                dtInformacion = cDataBase.ejecutarConsulta(strConsulta);
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }
        #region PDF Evento

        #region [Viejo] Agregar Archivo Evento
        public void agregarArchivoEvento(String IdRegistro, String UrlArchivo)
        {
            try
            {
                cDataBase.conectar();
                cDataBase.ejecutarQuery("INSERT INTO Riesgos.Archivos (IdControlUsuario, IdRegistro, NombreUsuario, FechaRegistro, UrlArchivo) VALUES (6, " + IdRegistro + ", '" + NombreUsuario + "', GETDATE(), N'" + UrlArchivo + "')");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }
        #endregion [Viejo] Agregar Archivo Evento

        public void mtdAgregarArchivoPdf(string strIdControl, string strIdRegistro, string strUrlArchivo, byte[] bPdfData)
        {
            string strConsulta = string.Empty;

            try
            {
                strConsulta = string.Format("INSERT INTO Riesgos.Archivos ([IdControlUsuario], [IdRegistro], [NombreUsuario], [FechaRegistro], [UrlArchivo], [ArchivoPDF]) VALUES ({3}, {0}, '{1}', GETDATE(),N'{2}', @PdfData)",
                    strIdRegistro, NombreUsuario, strUrlArchivo, strIdControl);

                cDataBase.mtdConectarSql();
                cDataBase.mtdEjecutarConsultaSQL(strConsulta, bPdfData);
                cDataBase.mtdDesconectarSql();
            }
            catch (Exception ex)
            {
                cDataBase.mtdDesconectarSql();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public byte[] mtdDescargarArchivoPdf(string strNombreArchivo)
        {
            #region Vars
            byte[] bInfo = null;
            string strConsulta = string.Empty;
            #endregion Vars

            try
            {
                strConsulta = string.Format("SELECT [UrlArchivo], [ArchivoPDF] FROM [Riesgos].[Archivos] WHERE [UrlArchivo] = N'{0}'", strNombreArchivo);

                cDataBase.mtdConectarSql();
                bInfo = cDataBase.mtdEjecutarConsultaSqlPdf(strConsulta);
                cDataBase.mtdDesconectarSql();
            }
            catch (Exception ex)
            {
                cDataBase.mtdDesconectarSql();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }

            return bInfo;
        }
        #endregion PDF Evento

        #region MailData
        public void mtdJerarquiaUsuario(string idUsuario, ref string NombreJerarquia, ref string NombreArea)
        {
            DataTable dtInformacion = new DataTable();
            string consulta = string.Empty;
            consulta = string.Format("SELECT Users.[IdUsuario],PJO.NombreHijo,PA.NombreArea"
                + " FROM [Listas].[Usuarios] as Users"
                + " inner join Parametrizacion.JerarquiaOrganizacional as PJO on PJO.idHijo = Users.IdJerarquia"
                + " inner join Parametrizacion.DetalleJerarquiaOrg as PDJO on PDJO.idHijo = PJO.idHijo"
                + " inner join [Parametrizacion].[Area] as PA on PA.IdArea = PDJO.IdArea"
                + " WHERE Users.[IdUsuario] = {0}", idUsuario);
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta(consulta);
                cDataBase.desconectar();

                NombreJerarquia = dtInformacion.Rows[0]["NombreHijo"].ToString().Trim();
                NombreArea = dtInformacion.Rows[0]["NombreArea"].ToString().Trim();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }
        public string mtdLastNoHuboEvento(ref int IdEvento)
        {
            DataTable dtInformacion = new DataTable();
            string consulta = string.Empty;
            string CodNHE = string.Empty;
            consulta = string.Format("select max(IdEvento) as LastId from [Riesgos].[NoHuboEventos]");
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta(consulta);
                cDataBase.desconectar();
                IdEvento = Convert.ToInt32(dtInformacion.Rows[0]["LastId"].ToString().Trim());

                consulta = string.Format("select CodigoEvento from [Riesgos].[NoHuboEventos] where IdEvento = {0}", IdEvento);
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta(consulta);
                cDataBase.desconectar();
                CodNHE = dtInformacion.Rows[0]["CodigoEvento"].ToString().Trim();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return CodNHE;
        }
        public DataTable loadMailEvents()
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT [IdEvento],[NombreEvento] FROM [Notificaciones].[Evento]");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }
        #endregion MailData

        public DataTable ConsultaNombreJerarquia(string CodigoEvento)
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT IdUsuario  FROM riesgos.Eventos   WHERE CodigoEvento = ('" + CodigoEvento + "')");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public DataTable ConsultaGeneraEvento(string CodigoEvento)
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT top 1 e.GeneraEvento, rtrim(Ltrim(e.NomGeneradorEvento)) NomGeneradorEvento  FROM riesgos.Eventos e  WHERE e.CodigoEvento = ('" + CodigoEvento+"')");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public DataTable NombreGeneraEvento(string IdJerarquia)
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT rtrim(Ltrim(idPadre)) idPadre,  rtrim(Ltrim(TipoArea)) TipoArea FROM Parametrizacion.JerarquiaOrganizacional WHERE idHijo =  (" + IdJerarquia + ")");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public DataTable PadreArea(string IdPadre)
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT rtrim(Ltrim(idHijo)) idHijo FROM Parametrizacion.JerarquiaOrganizacional WHERE idHijo =  (" + IdPadre + ") " );
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public DataTable TipoArea(string IdHijo)
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT rtrim(Ltrim(TipoArea)) TipoArea, IdPadre FROM Parametrizacion.JerarquiaOrganizacional WHERE idHijo =  (" + IdHijo + ")");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }


        public DataTable ConsultaIdPadre(int IdHijo)
        {
            DataTable dtInformacion = new DataTable();
            try
            {
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta("SELECT rtrim(Ltrim(idPadre)) IdPadre FROM Parametrizacion.JerarquiaOrganizacional WHERE IdHijo =  (" + IdHijo + ")");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public void CrearAsociacion(String IdEventoPrincipal, String CodigoEventoPrincipal, String IdEventoAsociado, String CodigoEventoAsociado, int IdUserAsociado)
        {
            try
            {
                cDataBase.conectar();
                cDataBase.ejecutarQuery(
                    "INSERT INTO Riesgos.EventosAsociados ([IdEventoPrincipal], [CodigoEventoPrincipal], [IdEventoAsociado], [CodigoEventoAsociado], [IdUsuario], [FechaRegistro]) " +
                    "VALUES (" + IdEventoPrincipal + ", '" + CodigoEventoPrincipal + "', " + IdEventoAsociado + ", '" + CodigoEventoAsociado + "', " + IdUserAsociado + ", GETDATE())");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public DataTable loadInfoEventosAsociados(String CodigoEvento)
        {
            DataTable dtInformacion = new DataTable();
            string CondicionOtroFiltro = string.Empty;
            try
            {
                CondicionOtroFiltro = "SELECT a.IdEventoAsociado as IdEvento, a.CodigoEventoAsociado as CodigoEvento, d.DescripcionEvento, b.Descripcion AS NombreClaseEvento, c.NombreTipoPerdidaEvento " +
                    "FROM Riesgos.EventosAsociados a WITH(NOLOCK) " +
                    "LEFT JOIN Riesgos.Eventos d WITH(NOLOCK) ON d.CodigoEvento = a.CodigoEventoAsociado " +
                    "LEFT JOIN Eventos.Clase b ON d.IdClase = b.IdClase " +
                    "LEFT JOIN Parametrizacion.TipoPerdidaEvento c WITH(NOLOCK) ON d.IdTipoPerdidaEvento = c.IdTipoPerdidaEvento " +
                    "WHERE(a.CodigoEventoPrincipal = '" + CodigoEvento + "')";
                

                #region consulta
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta(CondicionOtroFiltro);
                #endregion Consulta

                cDataBase.desconectar();
                return dtInformacion;
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        public string mtdGetCodEventAsociados(int IdEvento)
        {
            DataTable dtInformacion = new DataTable();
            string result = string.Empty;
            try
            {
                string query = "SELECT CodigoEventoAsociado FROM Riesgos.EventosAsociados WHERE IdEventoPrincipal = "+IdEvento;
                
                cDataBase.conectar();
                dtInformacion = cDataBase.ejecutarConsulta(query);
                cDataBase.desconectar();
                for(int i = 0; i < dtInformacion.Rows.Count; i++)
                {
                    result += dtInformacion.Rows[i]["CodigoEventoAsociado"].ToString().Trim() + "|";
                }
                //TipoArea = dtInformacion.Rows[0]["TipoArea"].ToString().Trim();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
            return result;
        }

    }
}