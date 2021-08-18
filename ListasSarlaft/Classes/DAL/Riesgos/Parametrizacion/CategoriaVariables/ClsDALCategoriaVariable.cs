using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using ListasSarlaft.Classes.BLL.Riesgos.Parametrizacion;

namespace ListasSarlaft.Classes.DAL.Riesgos.Parametrizacion.CategoriaVariables
{
    public class ClsDALCategoriaVariable
    {
        private cDataBase cDataBase = new cDataBase();

        public DataTable GestionarVariables(ref CCalificacionExperta objVariables, int IdTransaccion)
        {
            DataTable Resultado = new DataTable();

            try
            {
                if (string.IsNullOrEmpty(objVariables.NombreVariable))
                {
                    objVariables.NombreVariable = string.Empty;
                }

                if (string.IsNullOrEmpty(objVariables.UsuarioRegistro))
                {
                    objVariables.UsuarioRegistro = string.Empty;
                }

                if (string.IsNullOrEmpty(objVariables.EstadoVariable))
                {
                    objVariables.EstadoVariable = string.Empty;
                }                

                /*********************CATEGORIAS*********************************/
                if (string.IsNullOrEmpty(objVariables.NombreCategoria))
                {
                    objVariables.NombreCategoria = string.Empty;
                }
                 
                /*********************PUNTO DE CORTE*****************************/
                if (string.IsNullOrEmpty(objVariables.NombreFrecuencia))
                {
                    objVariables.NombreFrecuencia = string.Empty;
                }


                List<SqlParameter> parametros = new List<SqlParameter>()
                {
                     new SqlParameter() { ParameterName = "@Transaccion", SqlDbType = SqlDbType.Int, Value = IdTransaccion },
                     new SqlParameter() { ParameterName = "@IdVariable", SqlDbType = SqlDbType.Int, Value = objVariables.IdVariable },
                     new SqlParameter() { ParameterName = "@NombreVar", SqlDbType = SqlDbType.VarChar, Value = objVariables.NombreVariable },
                     new SqlParameter() { ParameterName = "@PonderacionVar", SqlDbType = SqlDbType.Int, Value = objVariables.Ponderacion },
                     new SqlParameter() { ParameterName = "@PuntuacionVar", SqlDbType = SqlDbType.Int, Value = objVariables.Ponderacion },
                     new SqlParameter() { ParameterName = "@EstadoVar", SqlDbType = SqlDbType.VarChar, Value = objVariables.EstadoVariable },
                     new SqlParameter() { ParameterName = "@NombreCategoria", SqlDbType = SqlDbType.VarChar, Value = objVariables.NombreCategoria },
                     new SqlParameter() { ParameterName = "@IdFrecuenciaEventos", SqlDbType = SqlDbType.Int, Value = objVariables.IdFrecuenciaEventos },
                     new SqlParameter() { ParameterName = "@NombreFrecuencia", SqlDbType = SqlDbType.VarChar, Value = objVariables.NombreFrecuencia },
                     new SqlParameter() { ParameterName = "@Min", SqlDbType = SqlDbType.Int, Value = objVariables.Min },
                     new SqlParameter() { ParameterName = "@Max", SqlDbType = SqlDbType.Int, Value = objVariables.Max },
                     new SqlParameter() { ParameterName = "@FechaRegistro", SqlDbType = SqlDbType.DateTime, Value = objVariables.FechaRegistro },
                     new SqlParameter() { ParameterName = "@UsuarioRegistro", SqlDbType = SqlDbType.VarChar, Value = objVariables.UsuarioRegistro },
                     new SqlParameter() { ParameterName = "@Resultado", SqlDbType = SqlDbType.Int, Direction = ParameterDirection.Output }
                };

                Resultado = cDataBase.EjecutarSPParametrosReturnDatatable("[Riesgos].[pa_GestionCategoriaVariable]", parametros);
            }
            catch (Exception ex)
            {
                throw new Exception("Error en procedimiento [Riesgos].[pa_GestionCategoriaVariable] : " + ex.Message.ToString());
            }

            return Resultado;
        }

        public DataTable GestionarCategorias (ref CCalificacionExperta objCategorias, int IdTransaccion)
        {
            DataTable Resultado = new DataTable();

            try
            {

                if (string.IsNullOrEmpty(objCategorias.NombreVariable))
                {
                    objCategorias.NombreVariable = string.Empty;
                }

                if (string.IsNullOrEmpty(objCategorias.UsuarioRegistro))
                {
                    objCategorias.UsuarioRegistro = string.Empty;
                }

                if (string.IsNullOrEmpty(objCategorias.EstadoVariable))
                {
                    objCategorias.EstadoVariable = string.Empty;
                }

                if (string.IsNullOrEmpty(objCategorias.EstadoVariable))
                {
                    objCategorias.EstadoVariable = string.Empty;
                }

            }
            catch (Exception ex)
            {

                throw new Exception("Error en procedimiento [Riesgos].[pa_GestionCategoriaVariable] : " + ex.Message.ToString());
            }

            return Resultado;
        }

    } //Fin espacio de nombres
}