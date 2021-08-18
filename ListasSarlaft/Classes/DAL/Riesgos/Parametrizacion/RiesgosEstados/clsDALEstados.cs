using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace ListasSarlaft.Classes.DAL.Riesgos.Parametrizacion.RiesgosEstados
{
    public class clsDALEstados  
    {
        //consultar estado
        public bool mtdConsultarEstados(ref DataTable dtCaracOut, ref string strErrMsg)
        {
            #region Vars
            bool booResult = false;
            string strConsulta = string.Empty;
            cDataBase cDataBase = new cDataBase();
            #endregion Vars

            try
            {
                strConsulta = string.Format("SELECT PT.[IdEstado],PT.[NombreEstado], PT.[Estado], PT.[UsuarioCreacion], Users.Usuario,PT.[FechaCreacion]"
                    + " FROM [Parametrizacion].[EstadosRiesgo] as PT"
                    + " inner join Listas.Usuarios as Users on Users.IdUsuario = PT.UsuarioCreacion");

                cDataBase.conectar();
                dtCaracOut = cDataBase.ejecutarConsulta(strConsulta);
                booResult = true;
            }
            catch (Exception ex)
            {
                strErrMsg = string.Format("Error al consultar los tratamientos. [{0}]", ex.Message);
                booResult = false;
            }
            finally
            {
                cDataBase.desconectar();
            }

            return booResult;
        }

        //insertar estado 
        public int mtdInsertarEstado(clsDTORiesgos objEstado, ref string strErrMsg, int evento)
        {
            #region Vars
            int booResult = 0;
            string strConsulta = string.Empty;
            cDataBase cDatabase = new cDataBase();
            #endregion Vars

            if (objEstado.strEstado == "Activo")
            {
                objEstado.strEstado = "1";
            }
            else
            {
                objEstado.strEstado = "0";
            }
            try
            {
                    List<SqlParameter> parametros = new List<SqlParameter>()
                {
                    new SqlParameter() { ParameterName = "@NombreEstado", SqlDbType = SqlDbType.VarChar, Value =  objEstado.strNombreEstado },
                    new SqlParameter() { ParameterName = "@Estado", SqlDbType = SqlDbType.Int, Value =  objEstado.strEstado },
                    new SqlParameter() { ParameterName = "@UsuarioCreacion", SqlDbType = SqlDbType.Int, Value =  objEstado.intIdUsuario },
                    new SqlParameter() { ParameterName = "@IdEstado", SqlDbType = SqlDbType.Int, Value =  objEstado.intIdEstado },
                    new SqlParameter() { ParameterName = "@TipoEvento", SqlDbType = SqlDbType.Int, Value =  evento }
                };
                if (evento == 2)
                {
                    DataTable dtInfo = new DataTable();
                    dtInfo = cDatabase.EjecutarSPParametrosReturnDatatable("[Riesgos].[pa_ParametrizacionEstados]", parametros);
                    string  cto =dtInfo.Rows[0]["cuantos"].ToString();
                    booResult = Convert.ToInt32(cto);
                }
                else
                {
                    booResult = cDatabase.EjecutarSPParametros("[Riesgos].[pa_ParametrizacionEstados]", parametros);

                }

                return booResult;
            }
            catch (Exception ex)
            {
                strErrMsg = string.Format("Error al crear el tratamiento. [{0}]", ex.Message);
                booResult = 0;
            }
            finally
            {
                cDatabase.desconectar();
            }

            return booResult;
        }
    }
}