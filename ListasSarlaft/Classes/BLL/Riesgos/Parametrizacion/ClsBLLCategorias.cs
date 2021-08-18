using ListasSarlaft.Classes.DAL.Riesgos.Parametrizacion.CategoriaVariables;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace ListasSarlaft.Classes.BLL.Riesgos.Parametrizacion
{
    public class ClsBLLCategorias
    {
        public List<CCalificacionExperta> GestionarVariables(ref List<CCalificacionExperta> ListaVariables, CCalificacionExperta objVariables, int Transaccion)
        {
            DataTable dtCarga = new DataTable();
            ClsDALCategoriaVariable CV = new ClsDALCategoriaVariable();

            try
            {
                dtCarga = CV.GestionarVariables(ref objVariables, Transaccion);                
            }
            catch (Exception ex)
            {

                throw new Exception("Error al procesar categorías: " + ex.Message.ToString());
            }

            return ListaVariables;
        }
    }
}