using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace ListasSarlaft.Classes
{
    public class clsDALTratamientos
    {
        public bool mtdConsultarTratamientos(ref DataTable dtCaracOut, ref string strErrMsg)
        {
            #region Vars
            bool booResult = false;
            string strConsulta = string.Empty;
            cDataBase cDataBase = new cDataBase();
            #endregion Vars

            try
            {
                strConsulta = string.Format("SELECT PT.[IdTratamiento], PT.[NombreTratamiento], PT.[Estado], PT.[UsuarioCreacion], Users.Usuario,  PT.[FechaCreacion] FROM[Parametrizacion].[Tratamiento] AS PT INNER JOIN Listas.Usuarios AS Users ON Users.IdUsuario = PT.UsuarioCreacion;");

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
        public bool mtdInsertarTratamiento(clsDTOTratamientos objTratamiento, ref string strErrMsg)
        {
            #region Vars
            bool booResult = false;
            string strConsulta = string.Empty;
            cDataBase cDatabase = new cDataBase();
            string hoy = string.Format("{0:yyyy-MM-dd hh:mm:ss}", objTratamiento.dtFechaRegistro);
            #endregion Vars

            if (objTratamiento.strEstado == "Activo")
            {
                objTratamiento.strEstado = "1";
            }
            else
            {
                objTratamiento.strEstado = "0";
            }
            try
            {//modificar formato fecha 
                strConsulta = string.Format("INSERT INTO [Parametrizacion].[Tratamiento] ([NombreTratamiento], [Estado], [UsuarioCreacion],[FechaCreacion])" +
                    "VALUES('{0}',{1},{2},'{3}') ",
                    objTratamiento.strTratamiento, objTratamiento.strEstado, objTratamiento.intIdUsuario, hoy);

                cDatabase.conectar();
                cDatabase.ejecutarQuery(strConsulta);
                booResult = true;
            }
            catch (Exception ex)
            {
                strErrMsg = string.Format("Error al crear el tratamiento. [{0}]", ex.Message);
                booResult = false;
            }
            finally
            {
                cDatabase.desconectar();
            }

            return booResult;
        }
        public bool mtdUpdateTratamiento(clsDTOTratamientos objTratamiento, ref string strErrMsg)
        {
            #region Vars
            bool booResult = false;
            string strConsulta = string.Empty;
            cDataBase cDatabase = new cDataBase();
            #endregion Vars
            if (objTratamiento.strEstado == "Activo")
            { objTratamiento.strEstado = "1"; }
            else
            { objTratamiento.strEstado = "0"; }
            try
            {
                strConsulta = string.Format("UPDATE [Parametrizacion].[Tratamiento] SET [NombreTratamiento] = '{0}', [Estado] = {1}" +
                    " WHERE IdTratamiento = {2}",
                    objTratamiento.strTratamiento, objTratamiento.strEstado, objTratamiento.intIdTratamiento);

                cDatabase.conectar();
                cDatabase.ejecutarQuery(strConsulta);
                booResult = true;
            }
            catch (Exception ex)
            {
                strErrMsg = string.Format("Error al actualizar el tratamiento. [{0}]", ex.Message);
                booResult = false;
            }
            finally
            {
                cDatabase.desconectar();
            }

            return booResult;
        }

        //verifico relacion 
        public int mtdUpdateVerificar(clsDTOTratamientos objTratamiento, ref string strErrMsg)
        {
            #region Vars
            int cuantos = 0;
            string strConsulta = string.Empty;
            cDataBase cDatabase = new cDataBase();
            #endregion Vars
            if (objTratamiento.strEstado == "Activo")
            { objTratamiento.strEstado = "1"; }
            else
            { objTratamiento.strEstado = "0"; }
            try
            {
                strConsulta = string.Format("SELECT TOP 1 ListaTratamiento FROM [Riesgos].[Riesgo] WHERE  ListaTratamiento LIKE '%{0}%' ",
                    objTratamiento.intIdTratamiento);

                cDatabase.conectar();
                cuantos = Convert.ToInt32(cDatabase.ejecutarTratamiento(strConsulta));

            }
            catch (Exception ex)
            {
                strErrMsg = string.Format("Error al actualizar el tratamiento. [{0}]", ex.Message);
                cuantos = 0;
            }
            finally
            {
                cDatabase.desconectar();
            }

            return cuantos;
        }

    }
}