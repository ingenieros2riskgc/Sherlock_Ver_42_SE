using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using ListasSarlaft.Classes.BLL.Riesgos.Parametrizacion;

namespace ListasSarlaft.Classes.DAL.Riesgos.Parametrizacion.Planes
{
    public class clsDALPlanes
    {
        private cDataBase cDataBase = new cDataBase();

        public DataTable ConsultarPlanes(ref clsDTOPlanes listaPlanes, int IdTransaccion)
        {
            #region Vars
            DataTable resultado = new DataTable(); 

            object[] objetoPeriodo; 
            objetoPeriodo = new object[1];

            object[] objetoMeta;
            objetoMeta = new object[1];

            object[] objetoCumplimiento;
            objetoCumplimiento = new object[1];
            string error = listaPlanes.FechaCompromiso.ToString();

            #endregion Vars
            //  var objetoPeriodo = DBNull.Value;
            try
            {
                if (string.IsNullOrEmpty (listaPlanes.Responsable ) )
                {
                    listaPlanes.Responsable = string.Empty;
                }
                if (string.IsNullOrEmpty(listaPlanes.Codigo))
                {
                    listaPlanes.Codigo = string.Empty;
                }
                if (string.IsNullOrEmpty(listaPlanes.NombrePlan))
                {
                    listaPlanes.NombrePlan = string.Empty;
                }
                if (string.IsNullOrEmpty(listaPlanes.DescripcionPlan))
                {
                    listaPlanes.DescripcionPlan = string.Empty;
                }
                if (string.IsNullOrEmpty(listaPlanes.Estado))
                {
                    listaPlanes.Estado = string.Empty;
                }
                if (string.IsNullOrEmpty(listaPlanes.CodigoRiesgo))
                {
                    listaPlanes.CodigoRiesgo = string.Empty;
                }
                if (string.IsNullOrEmpty(listaPlanes.Justificacion))
                {
                    listaPlanes.Justificacion = string.Empty;
                }
                if (string.IsNullOrEmpty(listaPlanes.CodigoEvento))
                {
                    listaPlanes.CodigoEvento = string.Empty;
                }                

                if (listaPlanes.Periodo.ToString() == "01/01/0001 12:00:00 a. m." || listaPlanes.Periodo.ToString() == "1/01/0001 12:00:00 a. m." ||
                    listaPlanes.Periodo.ToString() == "1/1/0001 12:00:00 a. m." )
                {
                    listaPlanes.Periodo = DateTime.Now;
                    objetoPeriodo[0] = DateTime.Now;                    
                }
                else if (listaPlanes.Periodo == null)
                {
                    objetoPeriodo[0] = DBNull.Value;
                }
                else
                {
                    objetoPeriodo[0] = listaPlanes.Periodo;
                }
                if (listaPlanes.Meta == null)
                {
                    objetoMeta[0] = DBNull.Value;
                }
                else
                {
                    objetoMeta[0] = listaPlanes.Meta;
                }
                if (listaPlanes.Cumplimiento == null)
                {
                    objetoCumplimiento[0] = DBNull.Value;
                }
                else
                {
                    objetoCumplimiento[0] = listaPlanes.Cumplimiento;
                }
                if (string.IsNullOrEmpty( listaPlanes.Seguimiento))
                {
                    listaPlanes.Seguimiento = string.Empty;
                }

                if (listaPlanes.FechaCompromiso.ToString() == "1/01/0001 12:00:00 a.m." || listaPlanes.FechaCompromiso.ToString() == "01/01/0001 12:00:00 a.m." ||
                    listaPlanes.FechaCompromiso.ToString() == "1/1/0001 12:00:00 a.m." || listaPlanes.FechaCompromiso.ToString() == "1/01/0001 12:00:00 a. m." || 
                    listaPlanes.FechaCompromiso.ToString() == "01/01/0001 12:00:00 a. m." || listaPlanes.FechaCompromiso.ToString() == "1/1/0001 12:00:00 a. m." ||
                    listaPlanes.FechaCompromiso.ToString() == "01/01/0001 0:00:00")
                {
                    listaPlanes.FechaCompromiso = DateTime.Now;
                }               

                if (string.IsNullOrEmpty(listaPlanes.Adjuntos))
                {
                    listaPlanes.Adjuntos = string.Empty;
                }
                if (string.IsNullOrEmpty(listaPlanes.Usuario))
                {
                    listaPlanes.Usuario = string.Empty;
                }
                if (string.IsNullOrEmpty(listaPlanes.Resultado))
                {
                    listaPlanes.Resultado = string.Empty;
                }
                if (IdTransaccion == 0)
                {
                   // listaPlanes.Periodo = DateTime.Now;
                    listaPlanes.Meta = 0;
                    listaPlanes.Cumplimiento = 0;
                }     
                
                List<SqlParameter> parametros = new List<SqlParameter>()
                        {
                            new SqlParameter() { ParameterName = "@IdTransaccion", SqlDbType = SqlDbType.Int, Value = IdTransaccion },
                            new SqlParameter() { ParameterName = "@CodigoPlan", SqlDbType = SqlDbType.VarChar, Value =  listaPlanes.Codigo},
                            new SqlParameter() { ParameterName = "@NombrePlan", SqlDbType = SqlDbType.VarChar, Value =  listaPlanes.NombrePlan },
                            new SqlParameter() { ParameterName = "@DescripcionPlan", SqlDbType = SqlDbType.VarChar, Value =  listaPlanes.DescripcionPlan },
                            new SqlParameter() { ParameterName = "@Responsable", SqlDbType = SqlDbType.VarChar, Value =  listaPlanes.Responsable },
                            new SqlParameter() { ParameterName = "@Estado", SqlDbType = SqlDbType.VarChar, Value =  listaPlanes.Estado },
                            new SqlParameter() { ParameterName = "@FechaCompromiso", SqlDbType = SqlDbType.DateTime, Value =  listaPlanes.FechaCompromiso},
                            new SqlParameter() { ParameterName = "@Adjuntos ", SqlDbType = SqlDbType.VarChar, Value =  listaPlanes.Adjuntos },
                            new SqlParameter() { ParameterName = "@Justificacion ", SqlDbType = SqlDbType.VarChar, Value =  listaPlanes.Justificacion },
                            new SqlParameter() { ParameterName = "@CodigoRiesgo ", SqlDbType = SqlDbType.VarChar, Value =  listaPlanes.CodigoRiesgo },
                            new SqlParameter() { ParameterName = "@CodigoEvento ", SqlDbType = SqlDbType.VarChar, Value =  listaPlanes.CodigoEvento },
                            new SqlParameter() { ParameterName = "@Periodo ", SqlDbType = SqlDbType.DateTime, Value = objetoPeriodo[0] },                            
                            new SqlParameter() { ParameterName = "@Meta ", SqlDbType = SqlDbType.Int, Value =  objetoMeta[0] }, 
                            new SqlParameter() { ParameterName = "@Cumplimiento ", SqlDbType = SqlDbType.Int, Value =   objetoCumplimiento[0] },
                            new SqlParameter() { ParameterName = "@Seguimiento ", SqlDbType = SqlDbType.VarChar, Value =  listaPlanes.Seguimiento },
                            new SqlParameter() { ParameterName = "@Usuario", SqlDbType = SqlDbType.VarChar, Value =  listaPlanes.Usuario },
                            new SqlParameter() { ParameterName = "@Resultado", SqlDbType = SqlDbType.Int, Direction = ParameterDirection.Output }
                        };            

                  resultado  =   cDataBase.EjecutarSPParametrosReturnDatatable("[dbo].[pa_CargarPlanes]", parametros);                
            }
            catch (Exception ex)
            {
                throw new Exception("Error al generar la consulta: " + ex.Message.ToString() + "- Verifica formato fecha:" + listaPlanes.FechaCompromiso.ToString() + error);               
            }
            finally
            {
                cDataBase.desconectar();
            }

            return resultado;

        }
    }
}