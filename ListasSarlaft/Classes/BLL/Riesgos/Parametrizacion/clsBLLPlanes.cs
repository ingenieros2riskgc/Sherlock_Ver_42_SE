using ListasSarlaft.Classes.DAL.Riesgos.Parametrizacion.Planes;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace ListasSarlaft.Classes.BLL.Riesgos.Parametrizacion
{
    public class clsBLLPlanes
    {
        public List<clsDTOPlanes> ConsultarPlanes(ref List<clsDTOPlanes> listaPlanes,  clsDTOPlanes TodosPlanes, int Transaccion)
        { 
            #region var
            bool resultado = false;
            DataTable dtCarga = new DataTable();
            clsDALPlanes planes = new clsDALPlanes();
            #endregion var
             
            clsDTOPlanes ptPlanes = new clsDTOPlanes();           
                      
            dtCarga = planes.ConsultarPlanes(ref TodosPlanes, Transaccion);
            try
            {
                if (Transaccion == 0)
                {
                    if (dtCarga != null)
                    {
                        if (dtCarga.Rows.Count > 0)
                        {
                            foreach (DataRow dr in dtCarga.Rows)
                            {
                                clsDTOPlanes objPlanes = new clsDTOPlanes();
                                objPlanes.Codigo = dr["CodigoPlan"].ToString().Trim();
                                objPlanes.NombrePlan = dr["NombrePlan"].ToString().Trim();
                                objPlanes.Usuario = dr["Usuario"].ToString().Trim();
                                objPlanes.FechaRegistro = dr["FechaRegistro"].ToString().Trim();

                                listaPlanes.Add(objPlanes);
                            }
                        }
                        else
                        { listaPlanes = null; }
                    }
                    else
                    { listaPlanes = null; }
                }
                if (Transaccion == 1 || Transaccion == 3)
                {
                    TodosPlanes.Resultado = "OK";
                }
                if (Transaccion == 2)
                {
                    if (dtCarga.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtCarga.Rows)
                        {
                            clsDTOPlanes objPlanes = new clsDTOPlanes();
                            objPlanes.Codigo = dr["CodigoPlan"].ToString().Trim();
                            objPlanes.NombrePlan = dr["NombrePlan"].ToString().Trim();
                            objPlanes.DescripcionPlan = dr["DescripcionPlan"].ToString().Trim();
                            objPlanes.Responsable = dr["Responsable"].ToString().Trim();
                            objPlanes.Estado = dr["Estado"].ToString().Trim();
                            objPlanes.FechaCompromiso = Convert.ToDateTime(dr["FechaCompromiso"].ToString());                                                  
                            objPlanes.FechaRegistro = dr["FechaRegistro"].ToString().Trim();
                            objPlanes.Usuario = dr["Usuario"].ToString().Trim();

                            listaPlanes.Add(objPlanes);
                        }
                    }                                       
                }
            }
            catch (Exception ex)
            {
              throw new Exception("Error al mostrar los valores: " + ex.Message.ToString());
           
            }
            return listaPlanes;

        }

        
    }
}