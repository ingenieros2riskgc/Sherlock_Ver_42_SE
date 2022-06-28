using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using clsDTO;

namespace clsDatos
{
    public class clsDtImpactoCualitativo
    {
        public clsDtImpactoCualitativo() { }

        public DataTable mtdConsultaImpactoCualitativo(ref string strErrMsg)
        {
            #region Vars
            clsDatabase cDatabase = new clsDatabase();
            DataTable dtInformacion = new DataTable();
            string strConsulta = string.Empty;
            #endregion Vars

            try
            {
                strConsulta = string.Format("SELECT [idImpactoCualitativo], [Nombre], [IdUsuario], [FechaRegistro] FROM [Riesgos].[tblImpactoCualitativo]");
                cDatabase.conectar();
                dtInformacion = cDatabase.ejecutarConsulta(strConsulta);
            }
            catch (Exception ex)
            {
                strErrMsg = string.Format("Error al consultar los impactos cualitativos. [{0}]", ex.Message);
                dtInformacion = null;
            }
            finally
            {
                cDatabase.desconectar();
            }

            return dtInformacion;
        }

        public DataTable mtdInsertarImpactoCualitativo(clsDTOImpactoCualitativo objImpCual, ref string strErrMsg)
        {
            #region Vars
            clsDatabase cDatabase = new clsDatabase();
            DataTable dtInformacion = new DataTable();
            string strConsulta = string.Empty;
            #endregion Vars

            try
            {
                strConsulta = string.Format("INSERT INTO [Riesgos].[tblImpactoCualitativo] ([Nombre], [IdUsuario], [FechaRegistro]) " +
                    "VALUES ('{0}', {1}, GETDATE())", objImpCual.Nombre, objImpCual.IdUsuario, "GETDATE()");

                cDatabase.conectar();
                dtInformacion = cDatabase.ejecutarConsulta(strConsulta);
            }
            catch (Exception ex)
            {
                strErrMsg = string.Format("Error al crear el requerimiento. [{0}]", ex.Message);
                dtInformacion = null;
            }
            finally
            {
                cDatabase.desconectar();
            }
            return dtInformacion;
        }

        public DataTable mtdActualizarImpactoCualitativo(clsDTOImpactoCualitativo objImpCual, ref string strErrMsg)
        {
            #region Vars
            clsDatabase cDatabase = new clsDatabase();
            DataTable dtInformacion = new DataTable();
            string strConsulta = string.Empty;
            #endregion Vars

            try
            {
                strConsulta = string.Format("update [Riesgos].[tblImpactoCualitativo] set [Nombre] = '{1}', [IdUsuario] = {2} " +
                    "where idImpactoCualitativo = {0}", objImpCual.idImpactoCualitativo, objImpCual.Nombre, objImpCual.IdUsuario);

                cDatabase.conectar();
                dtInformacion = cDatabase.ejecutarConsulta(strConsulta);
            }
            catch (Exception ex)
            {
                strErrMsg = string.Format("Error al actualizar el impacto cualitativo. [{0}]", ex.Message);
                dtInformacion = null;
            }
            finally
            {
                cDatabase.desconectar();
            }
            return dtInformacion;
        }

        public DataTable ConsultaEliminar()
        {
            DataTable dtInformacion = new DataTable();
            clsDatabase cDatabase = new clsDatabase();

            try
            {
                cDatabase.conectar();
                dtInformacion = cDatabase.ejecutarConsulta("select IdImpactoCualitativo from [Riesgos].[tblImpactoCualitativo]");
                cDatabase.desconectar();
            }
            catch (Exception ex)
            {
                cDatabase.desconectar();
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        public DataTable ConsultaImpactoCualitativo()
        {
            DataTable dtInformacion = new DataTable();
            clsDatabase cDatabase = new clsDatabase();

            try
            {
                cDatabase.conectar();
                dtInformacion = cDatabase.ejecutarConsulta("SELECT [idImpactoCualitativo], [Nombre], [IdUsuario], [FechaRegistro] " +
                    "FROM [Riesgos].[tblImpactoCualitativo] order by idImpactoCualitativo DESC");
                cDatabase.desconectar();
            }
            catch (Exception ex)
            {
                cDatabase.desconectar();
                throw new Exception(ex.Message);
            }
            return dtInformacion;
        }

        //public DataTable loadNombreImpCual()
        //{
        //    DataTable dtInformacion = new DataTable();
        //    clsDatabase cDatabase = new clsDatabase();

        //    try
        //    {
        //        cDatabase.conectar();
        //        dtInformacion = cDatabase.ejecutarConsulta("select distinct Nombre from Riesgos.tblImpactoCualitativo");
        //        cDatabase.desconectar();
        //    }
        //    catch (Exception ex)
        //    {
        //        cDatabase.desconectar();
        //        throw new Exception(ex.Message);
        //    }
        //    return dtInformacion;
        //}
    }
}
