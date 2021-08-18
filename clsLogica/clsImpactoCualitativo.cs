using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using clsDatos;
using clsDTO;

namespace clsLogica
{
    public class clsImpactoCualitativo
    {
        public clsImpactoCualitativo()
        {
        }

        public List<clsDTOImpactoCualitativo> mtdCargarInfoImpactoCualitativo(ref string strErrMsg)
        {
            #region Vars
            DataTable dtInfo = new DataTable();
            clsDtImpactoCualitativo cDtImpCual = new clsDtImpactoCualitativo();
            clsDTOImpactoCualitativo objImpCual = new clsDTOImpactoCualitativo();
            List<clsDTOImpactoCualitativo> lstImpCual = new List<clsDTOImpactoCualitativo>();
            #endregion Vars

            dtInfo = cDtImpCual.mtdConsultaImpactoCualitativo(ref strErrMsg);

            if (dtInfo != null)
            {
                if (dtInfo.Rows.Count > 0)
                {
                    foreach (DataRow dr in dtInfo.Rows)
                    {
                        objImpCual = new clsDTOImpactoCualitativo(
                            dr["idImpactoCualitativo"].ToString().Trim(),
                            dr["Nombre"].ToString().Trim(),
                            dr["IdUsuario"].ToString().Trim(),
                            dr["FechaRegistro"].ToString().Trim()
                            );

                        lstImpCual.Add(objImpCual);
                    }
                }
                else
                {
                    lstImpCual = null;
                    strErrMsg = "No hay información de impactos cualitativos registrados.";
                }
            }
            else
                lstImpCual = null;

            return lstImpCual;
        }

        public void mtdInsertarImpactoCualitativo(clsDTOImpactoCualitativo objImpCual, ref string strErrMsg)
        {
            clsDtImpactoCualitativo cDtImpCual = new clsDtImpactoCualitativo();

            cDtImpCual.mtdInsertarImpactoCualitativo(objImpCual, ref strErrMsg);
        }
        public void mtdActualizarImpactoCualitativo(clsDTOImpactoCualitativo objImpCual, ref string strErrMsg)
        {
            clsDtImpactoCualitativo cDtImpCual = new clsDtImpactoCualitativo();

            cDtImpCual.mtdActualizarImpactoCualitativo(objImpCual, ref strErrMsg);
        }

    }
}
