using ListasSarlaft.Classes.DAL.Riesgos.Parametrizacion.RiesgosEstados;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace ListasSarlaft.Classes.BLL.Riesgos.Parametrizacion.Riesgos
{
    public class clsBLLEstados 
    {
        /// <summary>
        /// Metodo para consultar y visualizar los Tratamientos del Riesgo
        /// </summary>
        /// <param name="strErrMsg">Mensaje de error</param>
        /// <returns>Retorna si el proceso fue exitoso o no</returns> 
        public List<clsDTORiesgos> mtdConsultarEstados(ref List<clsDTORiesgos> lstEstado, ref string strErrMsg)
        {
            #region Vars
            bool booResult = false;
            DataTable dtInfo = new DataTable();
            clsDALEstados cDtEstados = new clsDALEstados();
            #endregion Vars

            booResult = cDtEstados.mtdConsultarEstados(ref dtInfo, ref strErrMsg);

            if (booResult)
            {
                if (dtInfo != null)
                {
                    if (dtInfo.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtInfo.Rows)
                        {
                            clsDTORiesgos objEstado = new clsDTORiesgos();
                            objEstado.intIdEstado = Convert.ToInt32(dr["IdEstado"].ToString().Trim());
                            objEstado.strNombreEstado = dr["NombreEstado"].ToString().Trim();
                            objEstado.strEstado = dr["Estado"].ToString().Trim();
                            objEstado.intIdUsuario = Convert.ToInt32(dr["UsuarioCreacion"].ToString().Trim());
                            objEstado.strUsuario = dr["Usuario"].ToString().Trim();
                            objEstado.dtFechaRegistro = Convert.ToDateTime(dr["FechaCreacion"].ToString().Trim());

                            lstEstado.Add(objEstado);
                        }
                    }
                    else
                        lstEstado = null;
                }
                else
                    lstEstado = null;
            }

            return lstEstado;
        }

        public int mtdInsertarEstado(clsDTORiesgos objEstado, ref string strErrMsg, int evento)
        { 
            int booResult = 0;
            clsDALEstados cDALEstados = new clsDALEstados();
            if (evento == 1)
            {
                booResult = cDALEstados.mtdInsertarEstado(objEstado, ref strErrMsg, 1);
            }
            else
            {
                booResult = cDALEstados.mtdInsertarEstado(objEstado, ref strErrMsg, 0);
            }
            
            return booResult;
        }

        public int verificaRelacionEstado(clsDTORiesgos objEstado, ref string strErrMsg, int evento)
        {
            int booResult = 0;
            clsDALEstados cDALEstados = new clsDALEstados();
            if (evento == 1)
            {
                booResult = cDALEstados.mtdInsertarEstado(objEstado, ref strErrMsg, 1);
            }
            else if(evento == 0)
            {
                booResult = cDALEstados.mtdInsertarEstado(objEstado, ref strErrMsg, 0);
            }
            else
            {
                booResult = cDALEstados.mtdInsertarEstado(objEstado, ref strErrMsg, 2);
            }

            return booResult;
        }
    }
}