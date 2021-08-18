using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using ListasSarlaft.Classes.DAL.Riesgos.Parametrizacion.CategoriaVariables;

namespace ListasSarlaft.Classes.BLL.Riesgos.Parametrizacion
{
    public class ClsBLLCategoriaVariable
    {
        public List<CCalificacionExperta> GestionCategoriaVariable(ref List<CCalificacionExperta> ListaVariables, CCalificacionExperta objVariables, int Transaccion)
        {

            DataTable dtCarga = new DataTable();
            ClsDALCategoriaVariable CV = new ClsDALCategoriaVariable();
            
            try
            {
                dtCarga = CV.GestionarVariables(ref objVariables, Transaccion);

                if (Transaccion == 0)
                {
                    if (dtCarga.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtCarga.Rows)
                        {
                            CCalificacionExperta CargaVariables = new CCalificacionExperta();

                            CargaVariables.IdVariable = Convert.ToInt32(dr["IdVariable"]);
                            CargaVariables.NombreVariable = dr["NombreVariable"].ToString().Trim();
                            CargaVariables.Ponderacion = Convert.ToInt32 ( dr["Ponderacion"]);
                            CargaVariables.EstadoVariable = dr["EstadoVariable"].ToString();
                            CargaVariables.UsuarioRegistro = dr["UsuarioRegistro"].ToString().Trim();
                            CargaVariables.FechaRegistro = Convert.ToDateTime (dr["FechaRegistro"]);

                            ListaVariables.Add(CargaVariables);
                        }
                    }
                    else
                    {
                        ListaVariables = null;
                    }
                }
                if (Transaccion == 5 || Transaccion == 13)
                {
                    if (dtCarga.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtCarga.Rows)
                        {
                            CCalificacionExperta CargaCategorias = new CCalificacionExperta();

                            CargaCategorias.IdCategoria = Convert.ToInt32(dr["IdCategoria"]);
                            CargaCategorias.IdVariable = Convert.ToInt32(dr["IdVariable"]);                            
                            CargaCategorias.NombreVariable = dr["NombreVariable"].ToString().Trim();
                            CargaCategorias.NombreCategoria = dr["NombreCategoria"].ToString().Trim();
                            CargaCategorias.Ponderacion = Convert.ToInt32(dr["Ponderacion"]);                            
                            CargaCategorias.UsuarioRegistro = dr["UsuarioRegistro"].ToString().Trim();
                            CargaCategorias.FechaRegistro = Convert.ToDateTime(dr["FechaRegistro"]);

                            ListaVariables.Add(CargaCategorias);
                        }
                    }
                }

                if (Transaccion == 7)
                {
                    if (dtCarga.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtCarga.Rows)
                        {
                            CCalificacionExperta CargaPuntoCorte = new CCalificacionExperta();
                            CargaPuntoCorte.IdPuntoCorte = Convert.ToInt32(dr["IdPuntoCorte"]);
                            CargaPuntoCorte.IdFrecuenciaEventos = Convert.ToInt32(dr["IdFrecuenciaEventos"]);
                            CargaPuntoCorte.NombreFrecuencia = dr["NombreFrecuencia"].ToString().Trim();
                            CargaPuntoCorte.Min = Convert.ToInt32(dr["Min"]);
                            CargaPuntoCorte.Max = Convert.ToInt32(dr["Max"]);

                            ListaVariables.Add(CargaPuntoCorte);
                        }
                    }
                }
                 
                if (Transaccion == 9)
                {
                    if (ListaVariables.Count > 0)
                    {
                        ListaVariables.RemoveAt(0);
                    }
                    
                    if (dtCarga.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtCarga.Rows)
                        {
                            CCalificacionExperta CargaVariablesImpacto = new CCalificacionExperta();
                            
                            CargaVariablesImpacto.IdVariable = Convert.ToInt32(dr["IdVariable"]);
                            CargaVariablesImpacto.NombreVariable = dr["NombreVariable"].ToString().Trim();
                           // CargaVariablesImpacto.Ponderacion = Convert.ToInt32(dr["Ponderacion"]);
                            CargaVariablesImpacto.EstadoVariable = dr["EstadoVariable"].ToString();
                            CargaVariablesImpacto.UsuarioRegistro = dr["UsuarioRegistro"].ToString().Trim();
                            CargaVariablesImpacto.FechaRegistro = Convert.ToDateTime(dr["FechaRegistro"]);

                            ListaVariables.Add(CargaVariablesImpacto);
                        }
                    }
                }

                if (Transaccion == 16)
                {
                    if (dtCarga.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtCarga.Rows)
                        {
                            CCalificacionExperta CargaPuntoCorte = new CCalificacionExperta();
                            CargaPuntoCorte.IdPuntoCorte = Convert.ToInt32(dr["IdPuntoCorte"]);
                            CargaPuntoCorte.IdFrecuenciaEventos = Convert.ToInt32(dr["IdImpacto"]);
                            CargaPuntoCorte.NombreFrecuencia = dr["NombreImpacto"].ToString().Trim();
                            CargaPuntoCorte.Min = Convert.ToInt32(dr["Min"]);
                            CargaPuntoCorte.Max = Convert.ToInt32(dr["Max"]);

                            ListaVariables.Add(CargaPuntoCorte);
                        }
                    }
                }

                if (Transaccion == 18)
                {
                    if (dtCarga.Rows.Count > 0)
                    {
                        CCalificacionExperta Sum = new CCalificacionExperta();
                        Sum.Max = Convert.ToInt32(dtCarga.Rows[0]["sumatoria"]);
                        ListaVariables.Add(Sum);
                    }
                }               
            }
            catch (Exception ex)
            {
                throw new Exception("Error al gestionar pa_GestionCategoriaVariable: " + ex.Message.ToString());                
            }
          
            return ListaVariables;
        }
    } // Fin espacios de nombres
}