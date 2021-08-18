using ListasSarlaft.Classes;
using ListasSarlaft.Classes.BLL.Riesgos.Parametrizacion;
using System;
using System.Collections.Generic;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ListasSarlaft.UserControls.Riesgos
{
    public partial class CalificacionExpPuntosDeCorteImpacto : System.Web.UI.UserControl
    {
        private cCuenta cCuenta = new cCuenta();
        private CCalificacionExperta objPuntosCorte = new CCalificacionExperta();
        private List<CCalificacionExperta> ListaPuntosCorte = new List<CCalificacionExperta>();
        private ClsBLLCategoriaVariable CV = new ClsBLLCategoriaVariable();
        private string IdFormulario = "11009";

        #region Variables

        private DataTable todosLosPuntosCorte;
        private DataTable TodosLosPuntosCorte
        {
            get
            {
                todosLosPuntosCorte = (DataTable)ViewState["todosLosPuntosCorte"];
                return todosLosPuntosCorte;
            }
            set
            {
                todosLosPuntosCorte = value;
                ViewState["todosLosPuntosCorte"] = todosLosPuntosCorte;
            }
        }

        private int PagIndexGvPuntosCorte;

        private int rowGridPuntosCorte;
        private int RowGridPuntosCorte
        {
            get
            {
                rowGridPuntosCorte = (int)ViewState["rowGridPuntosCorte"];
                return rowGridPuntosCorte;
            }
            set
            {
                rowGridPuntosCorte = value;
                ViewState["rowGridPuntosCorte"] = rowGridPuntosCorte;
            }
        }

        private int idPuntosCorteGlobal;
        private int IdPuntosCorteGlobal
        {
            get
            {
                idPuntosCorteGlobal = (int)ViewState["idPuntosCorteGlobal"];
                return idPuntosCorteGlobal;
            }
            set
            {
                idPuntosCorteGlobal = value;
                ViewState["idPuntosCorteGlobal"] = idPuntosCorteGlobal;
            }
        }

        private int idEventoFrecuenciasGlobal;
        private int IdEventoFrecuenciasGlobal
        {
            get
            {
                idEventoFrecuenciasGlobal = (int)ViewState["idEventoFrecuenciasGlobal"];
                return idEventoFrecuenciasGlobal;
            }
            set
            {
                idEventoFrecuenciasGlobal = value;
                ViewState["idEventoFrecuenciasGlobal"] = idEventoFrecuenciasGlobal;
            }
        }

        #endregion Variables

        //*********************************************************************
        #region Eventos
        protected void Page_Load(object sender, EventArgs e)
        {
            ScriptManager scriptManager = ScriptManager.GetCurrent(this.Page);
            Page.Form.Attributes.Add("enctype", "multipart/form-data");
            objPuntosCorte.UsuarioRegistro = Session["NombreUsuario"].ToString();
            objPuntosCorte.FechaRegistro = DateTime.Now;

            if (cCuenta.permisosConsulta(IdFormulario) == "False")
            {
                Response.Redirect("~/Formularios/Sarlaft/Admin/HomeAdmin.aspx?Denegar=1");
            }
            else
            {
                if (!Page.IsPostBack)
                {
                    GrillaCategorias();
                }
            }
        }

        protected void GvPuntosCorte_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            PagIndexGvPuntosCorte = e.NewPageIndex;
            GvPuntosCorte.PageIndex = PagIndexGvPuntosCorte;
            GvPuntosCorte.DataSource = TodosLosPuntosCorte;
            GvPuntosCorte.DataBind();
        }

        protected void GvPuntosCorte_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            switch (e.CommandName)
            {
                case "Editar":
                    RowGridPuntosCorte = (Convert.ToInt16(GvPuntosCorte.PageSize) * PagIndexGvPuntosCorte) + Convert.ToInt16(e.CommandArgument);
                    GridViewRow row = GvPuntosCorte.Rows[Convert.ToInt16(e.CommandArgument)];
                    System.Collections.Specialized.IOrderedDictionary colsNoVisible = GvPuntosCorte.DataKeys[Convert.ToInt16(e.CommandArgument)].Values;

                    IdPuntosCorteGlobal = Convert.ToInt32(colsNoVisible[0]);
                    IdEventoFrecuenciasGlobal = Convert.ToInt32(colsNoVisible[1]);
                    ModalNombreFrecuencia.Text = colsNoVisible[2].ToString();
                    ModalMin.Text = colsNoVisible[3].ToString();
                    ModalMax.Text = colsNoVisible[4].ToString();

                    int trans = 13;
                    string script = @"<script type='text/javascript'>FocusPeriodo(" + trans + ");" + "</script>";
                    ScriptManager.RegisterStartupScript(this, typeof(Page), "invocarfuncion", script, false);
                    ModalMin.Focus();
                    ModalEditarPuntosCorte.Show();
                    break;
            }
        }

        protected void Actualizar_Click(object sender, EventArgs e)
        {
            int Transaccion = 17;
            objPuntosCorte.IdVariable = IdPuntosCorteGlobal;
            objPuntosCorte.IdFrecuenciaEventos = IdEventoFrecuenciasGlobal;
            objPuntosCorte.NombreFrecuencia = ModalNombreFrecuencia.Text;
            objPuntosCorte.Min = Convert.ToInt32(ModalMin.Text);
            objPuntosCorte.Max = Convert.ToInt32(ModalMax.Text);

            CV.GestionCategoriaVariable(ref ListaPuntosCorte, objPuntosCorte, Transaccion);

            omb.ShowMessage("Se actualizaron los valores satisfactoriamente! ", 3, "Atención");
            GrillaCategorias();

        }

        #endregion Eventos

        //************************************************************************
        #region Metodos
        public void GrillaCategorias()
        {
            objPuntosCorte.FechaRegistro = DateTime.Now;
            int Transaccion = 16;
            try
            {
                ListaPuntosCorte.Clear();
                ListaPuntosCorte = CV.GestionCategoriaVariable(ref ListaPuntosCorte, objPuntosCorte, Transaccion);

                if (ListaPuntosCorte != null)
                {
                    InicializaGrillaPuntosCorte();
                    CargaPuntosCorte(ListaPuntosCorte);
                    GvPuntosCorte.DataSource = ListaPuntosCorte;
                    GvPuntosCorte.PageIndex = PagIndexGvPuntosCorte;
                    GvPuntosCorte.DataBind();
                }
            }
            catch (Exception ex)
            {
                omb.ShowMessage("Error al cargar los Puntos de Corte: " + ex.Message.ToString(), 1, "Error");
            }
        }

        private void InicializaGrillaPuntosCorte()
        {
            DataTable grid = new DataTable();

            grid.Columns.Add("IdPuntoCorte", typeof(string));
            grid.Columns.Add("IdFrecuenciaEventos", typeof(string));
            grid.Columns.Add("NombreFrecuencia", typeof(string));
            grid.Columns.Add("Min", typeof(string));
            grid.Columns.Add("Max", typeof(string));

            GvPuntosCorte.DataSource = grid;
            GvPuntosCorte.DataBind();
            TodosLosPuntosCorte = grid;
        }

        private void CargaPuntosCorte(List<CCalificacionExperta> TodasCategorias)
        {

            foreach (CCalificacionExperta objPuntosCorte in TodasCategorias)
            {

                TodosLosPuntosCorte.Rows.Add(new object[] {
                    objPuntosCorte.IdPuntoCorte.ToString(),
                    objPuntosCorte.IdFrecuenciaEventos.ToString(),
                    objPuntosCorte.NombreFrecuencia,
                    objPuntosCorte.Min,
                    objPuntosCorte.Max
            });
            }
        }

        #endregion Metodos


    }
}