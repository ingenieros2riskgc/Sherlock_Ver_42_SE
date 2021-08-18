using ListasSarlaft.Classes;
using ListasSarlaft.Classes.BLL.Riesgos.Parametrizacion;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ListasSarlaft.UserControls.Riesgos
{
    public partial class CalificacionExpCategoriaImpacto : System.Web.UI.UserControl
    {
        private cCuenta cCuenta = new cCuenta();
        private CCalificacionExperta objCategorias = new CCalificacionExperta();
        private List<CCalificacionExperta> ListaCategorias = new List<CCalificacionExperta>();
        private ClsBLLCategoriaVariable CV = new ClsBLLCategoriaVariable();
        private string IdFormulario = "11009";
      
        #region Variables
        private int PagIndexGvCategorias;

        private int idVariableGlobal;
        private int IdVariableGlobal
        {
            get
            {
                idVariableGlobal = (int)ViewState["idVariableGlobal"];
                return idVariableGlobal;
            }
            set
            {
                idVariableGlobal = value;
                ViewState["idVariableGlobal"] = idVariableGlobal;
            }
        }

        private int idCategoriaGlobal;
        private int IdCategoriaGlobal
        {
            get
            {
                idCategoriaGlobal = (int)ViewState["idCategoriaGlobal"];
                return idCategoriaGlobal;
            }
            set
            {
                idCategoriaGlobal = value;
                ViewState["idCategoriaGlobal"] = idCategoriaGlobal;
            }
        }

        private int rowGridCategorias;
        private int RowGridCategorias
        {
            get
            {
                rowGridCategorias = (int)ViewState["rowGridCategorias"];
                return rowGridCategorias;
            }
            set
            {
                rowGridCategorias = value;
                ViewState["rowGridCategorias"] = rowGridCategorias;
            }
        }

        DataTable todasLasCategorias;
        private DataTable TodasLasCategorias
        {
            get
            {
                todasLasCategorias = (DataTable)ViewState["todasLasCategorias"];
                return todasLasCategorias;
            }
            set
            {
                todasLasCategorias = value;
                ViewState["todasLasCategorias"] = todasLasCategorias;
            }
        }
        #endregion Variables

        //******************************************************************************************
        #region Eventos
        protected void Page_Load(object sender, EventArgs e)
        {
            ScriptManager scriptManager = ScriptManager.GetCurrent(this.Page);
            Page.Form.Attributes.Add("enctype", "multipart/form-data");
            objCategorias.UsuarioRegistro = Session["NombreUsuario"].ToString();
            objCategorias.FechaRegistro = DateTime.Now;

            if (cCuenta.permisosConsulta(IdFormulario) == "False")
            {
                Response.Redirect("~/Formularios/Sarlaft/Admin/HomeAdmin.aspx?Denegar=1");
            }
            else
            {
                if (!Page.IsPostBack)
                {
                     ComboNombresVariables();
                     GrillaCategorias();
                }
            }
        }

        protected void Aceptar_Click(object sender, ImageClickEventArgs e)
        {
            try
            {
                int
                Transaccion = 14;
                objCategorias.IdVariable = Convert.ToInt32(NombreVariable.SelectedItem.Value);
                objCategorias.NombreVariable = NombreVariable.SelectedItem.Text;
                objCategorias.NombreCategoria = NombreCategoria.Text;
                objCategorias.Ponderacion = Convert.ToInt32(PonderacionCategoria.Text);
                ListaCategorias = CV.GestionCategoriaVariable(ref ListaCategorias, objCategorias, Transaccion);

                omb.ShowMessage("Se ha asociado la Categoría a la Variable Satisfactoriamente!", 3, "Atención");
                GrillaCategorias();
                NombreCategoria.Text = string.Empty;
                PonderacionCategoria.Text = string.Empty;
            }
            catch (Exception ex)
            {
                omb.ShowMessage("No se pudo asociar la categoría: " + ex.Message.ToString(), 1, "Error");
            }

        }

        protected void Limpiar_Click(object sender, ImageClickEventArgs e)
        {
            GrillaCategorias();
            NombreVariable.SelectedIndex = 0;
            NombreCategoria.Text = string.Empty;
            PonderacionCategoria.Text = string.Empty;
        }

        protected void GvCategorias_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            PagIndexGvCategorias = e.NewPageIndex;
            GvCategorias.PageIndex = PagIndexGvCategorias;
            GvCategorias.DataSource = TodasLasCategorias;
            GvCategorias.DataBind();
        }

        protected void GvCategorias_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            switch (e.CommandName)
            {
                case "Editar":
                    RowGridCategorias = (Convert.ToInt16(GvCategorias.PageSize) * PagIndexGvCategorias) + Convert.ToInt16(e.CommandArgument);
                    GridViewRow row = GvCategorias.Rows[Convert.ToInt16(e.CommandArgument)];
                    System.Collections.Specialized.IOrderedDictionary colsNoVisible = GvCategorias.DataKeys[Convert.ToInt16(e.CommandArgument)].Values;

                    objCategorias.IdVariable = Convert.ToInt32(colsNoVisible[0]);
                    objCategorias.NombreVariable = colsNoVisible[1].ToString();
                    objCategorias.IdCategoria = Convert.ToInt32(colsNoVisible[2]);
                    objCategorias.NombreCategoria = colsNoVisible[3].ToString();
                    objCategorias.PonderacionCategoria = Convert.ToInt32(colsNoVisible[4]);
                    ModalNombreCategorias.Text = objCategorias.NombreCategoria;
                    ModalPonderacion.Text = objCategorias.PonderacionCategoria.ToString();
                    IdCategoriaGlobal = objCategorias.IdCategoria;
                    IdVariableGlobal = objCategorias.IdVariable;
                    ModalEditarCategorias.Show();

                    int trans = 12;
                    string script = @"<script type='text/javascript'>FocusPeriodo(" + trans + ");" + "</script>";
                    ScriptManager.RegisterStartupScript(this, typeof(Page), "invocarfuncion", script, false);
                    ModalNombreCategorias.Focus();


                    break;
            }
        }

        protected void Actualizar_Click(object sender, EventArgs e)
        {
            int Transaccion = 15;

            objCategorias.IdVariable = IdVariableGlobal;
            objCategorias.IdFrecuenciaEventos = IdCategoriaGlobal;
            objCategorias.NombreCategoria = ModalNombreCategorias.Text;
            objCategorias.Ponderacion = Convert.ToInt32(ModalPonderacion.Text);

            CV.GestionCategoriaVariable(ref ListaCategorias, objCategorias, Transaccion);

            omb.ShowMessage("Se actualizaron los valores satisfactoriamente! ", 3, "Atención");

            GrillaCategorias();
            NombreVariable.SelectedIndex = 0;
            NombreCategoria.Text = string.Empty;
            PonderacionCategoria.Text = string.Empty;
        }

        #endregion Eventos

        //******************************************************************************************
        #region Metodos
        public void ComboNombresVariables()
        {
            int Transaccion = 9;
            ListaCategorias = CV.GestionCategoriaVariable(ref ListaCategorias, objCategorias, Transaccion);

            foreach (var item in ListaCategorias)
            {
                int i = 0;
                int Id = item.IdVariable;
                string Nombre = item.NombreVariable;
                NombreVariable.Items.Insert(i + 1, new ListItem(Nombre, Id.ToString()));
                i++;
            }
        }

        public void GrillaCategorias()
        {
            objCategorias.FechaRegistro = DateTime.Now;
            int Transaccion = 13;
            try
            {
                ListaCategorias.Clear();
                ListaCategorias = CV.GestionCategoriaVariable(ref ListaCategorias, objCategorias, Transaccion);

                if (ListaCategorias != null)
                {
                    InicializaGrillaCategorias();
                    CargaCategorias(ListaCategorias);
                    GvCategorias.DataSource = ListaCategorias;
                    GvCategorias.PageIndex = PagIndexGvCategorias;
                    GvCategorias.DataBind();
                }
            }
            catch (Exception ex)
            {
                omb.ShowMessage("Error al cargar las categorías: " + ex.Message.ToString(), 1, "Error");

            }
        }

        private void InicializaGrillaCategorias()
        {
            DataTable grid = new DataTable();

            grid.Columns.Add("IdCategoria", typeof(string));
            grid.Columns.Add("IdVariable", typeof(string));
            grid.Columns.Add("NombreVariable", typeof(string));
            grid.Columns.Add("NombreCategoria", typeof(string));
            grid.Columns.Add("Ponderacion", typeof(string));
            grid.Columns.Add("FechaRegistro", typeof(string));
            grid.Columns.Add("UsuarioRegistro", typeof(string));
            grid.Columns.Add("EditarRegistro", typeof(string));

            GvCategorias.DataSource = grid;
            GvCategorias.DataBind();
            TodasLasCategorias = grid;
        }

        private void CargaCategorias(List<CCalificacionExperta> TodasCategorias)
        {
            foreach (CCalificacionExperta objCategorias in TodasCategorias)
            {

                TodasLasCategorias.Rows.Add(new object[] {
                    objCategorias.IdCategoria.ToString(),
                    objCategorias.IdVariable.ToString(),
                    objCategorias.NombreVariable,
                    objCategorias.NombreCategoria,
                    objCategorias.Ponderacion.ToString(),
                    objCategorias.UsuarioRegistro,
                    objCategorias.FechaRegistro.ToString()
                    });
            }
        }

        #endregion Metodos

        
    }
}