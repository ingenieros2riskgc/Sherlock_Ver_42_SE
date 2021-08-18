using ListasSarlaft.Classes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ListasSarlaft.Classes.DAL.Riesgos.Parametrizacion.CategoriaVariables;
using ListasSarlaft.Classes.BLL.Riesgos.Parametrizacion;
using System.Data;

namespace ListasSarlaft.UserControls.Riesgos
{
    public partial class CalificacionExpVarCat : System.Web.UI.UserControl
    {
        private string IdFormulario = "11009";        
        private cCuenta cCuenta = new cCuenta();
        private CCalificacionExperta objVariables= new CCalificacionExperta();
        private List<CCalificacionExperta> listaVariables = new List<CCalificacionExperta>();
        private ClsBLLCategoriaVariable CV = new ClsBLLCategoriaVariable();

        private int PagIndexGvVariables;
        
        DataTable todasLasVariables;
        private DataTable TodasLasVariables
        {
            get
            {
                todasLasVariables = (DataTable)ViewState["todasLasVariables"];
                return todasLasVariables;
            }
            set
            {
                todasLasVariables = value;
                ViewState["todasLasVariables"] = todasLasVariables;
            }
        }

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

        private int rowGridVariables;
        private int RowGridVariables
        {
            get
            {
                rowGridVariables = (int)ViewState["rowGridVariables"];
                return rowGridVariables;
            }
            set
            {
                rowGridVariables = value;
                ViewState["rowGridVariables"] = rowGridVariables;
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            ScriptManager scriptManager = ScriptManager.GetCurrent(this.Page);
            Page.Form.Attributes.Add("enctype", "multipart/form-data");
            objVariables.UsuarioRegistro = Session["NombreUsuario"].ToString();
            objVariables.FechaRegistro = DateTime.Now;

            if (cCuenta.permisosConsulta(IdFormulario) == "False")
            {
                Response.Redirect("~/Formularios/Sarlaft/Admin/HomeAdmin.aspx?Denegar=1");
            }
            else
            {
                if (!Page.IsPostBack)
                {
                    CargarGrillaVariables();
                }                    
            }
        }

        #region Eventos
        protected void Limpiar_Click(object sender, ImageClickEventArgs e)
        {
            NombreVariable.Text = string.Empty;
            PonderacionVariable.Text = string.Empty;
            EstadoVariable.SelectedIndex = 0;
            NombreVariable.Focus();
        }

        protected void Aceptar_Click(object sender, ImageClickEventArgs e)
        {
            int Transaccion = 1;
            objVariables.NombreVariable = NombreVariable.Text;
            objVariables.Ponderacion = Convert.ToInt32( PonderacionVariable.Text);

            if (EstadoVariable.SelectedValue.ToString() == "Activo")
            {objVariables.EstadoVariable = "1";}
            else{objVariables.EstadoVariable = "0";}                        
            
            if (cCuenta.permisosAgregar(IdFormulario) == "False")
            {                
                    omb.ShowMessage("No tiene los permisos suficientes para llevar a cabo esta acción.", 1, "Atención");                
            }
            else
            {
                try
                {
                    CV.GestionCategoriaVariable(ref listaVariables, objVariables, Transaccion);
                    omb.ShowMessage("Se ha cargado la variable satisfactoriamente! ", 3, "Atención");
                    CargarGrillaVariables();
                    NombreVariable.Text = string.Empty;
                    PonderacionVariable.Text = string.Empty;
                    EstadoVariable.SelectedIndex = 0;
                    NombreVariable.Focus();
                }
                catch (Exception ex)
                {
                    omb.ShowMessage("Error al registar la variable:" +ex.Message.ToString(),1,"Error");
                }                               
            }
        }

        protected void GvVariables_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            
            PagIndexGvVariables = e.NewPageIndex;
            GvVariables.PageIndex = PagIndexGvVariables;
            GvVariables.DataSource = TodasLasVariables;
            GvVariables.DataBind();
            int paginaGrid = (Convert.ToInt16(GvVariables.PageSize) * PagIndexGvVariables);
            cambiaEstado(paginaGrid);
        }

        protected void GvVariables_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            switch (e.CommandName)
            {
                case "Editar":
                    RowGridVariables = (Convert.ToInt16(GvVariables.PageSize) * PagIndexGvVariables) + Convert.ToInt16(e.CommandArgument);
                    GridViewRow row = GvVariables.Rows[Convert.ToInt16(e.CommandArgument)];
                    System.Collections.Specialized.IOrderedDictionary colsNoVisible = GvVariables.DataKeys[Convert.ToInt16(e.CommandArgument)].Values;
                    objVariables.IdVariable = Convert.ToInt32 (colsNoVisible[0]);
                    objVariables.NombreVariable = colsNoVisible[1].ToString();
                    objVariables.Ponderacion = Convert.ToInt32(colsNoVisible[2]);
                    IdVariableGlobal = objVariables.IdVariable;
                    ModalNombreVariable.Text = objVariables.NombreVariable;
                    ModalPonderacion.Text = objVariables.Ponderacion.ToString();
                    ModalEditarVariable.Show();

                    int trans = 11;                    
                    string script = @"<script type='text/javascript'>FocusPeriodo(" + trans + ");" + "</script>";
                    ScriptManager.RegisterStartupScript(this, typeof(Page), "invocarfuncion", script, false);
                    ModalNombreVariable.Focus();

                    break;

                case "Estado":
                    RowGridVariables = (Convert.ToInt16(GvVariables.PageSize) * PagIndexGvVariables) + Convert.ToInt16(e.CommandArgument);
                    GridViewRow rowm = GvVariables.Rows[Convert.ToInt16(e.CommandArgument)];
                    System.Collections.Specialized.IOrderedDictionary colsNoVisibles = GvVariables.DataKeys[Convert.ToInt16(e.CommandArgument)].Values;
                    objVariables.IdVariable = Convert.ToInt32(colsNoVisibles[0]);
                    IdVariableGlobal = objVariables.IdVariable;

                    if (RowGridVariables >= 10)
                    {
                        RowGridVariables = Convert.ToInt16(e.CommandArgument);
                    }
                    GridViewRow rows = GvVariables.Rows[RowGridVariables];
                    Label lbl = ((Label)rows.FindControl("booActivo"));
                    ImageButton ImgBnt = ((ImageButton)rows.FindControl("ImgBtnInact"));

                    if (lbl.Text == "Activo")
                    {
                        lblMsgBox1.Text = "¿Está seguro que desea inactivar la variable?";
                    }
                    else
                    {
                        lblMsgBox1.Text = "¿Está seguro que desea activar la variable?";
                    }
                    mpeMsgBox1.Show();
                    break;
            }
        }

        protected void Guardar_Click(object sender, EventArgs e)
        {
            try
            {
                int Transaccion = 2;
                objVariables.IdVariable = IdVariableGlobal;
                objVariables.NombreVariable = ModalNombreVariable.Text;
                objVariables.Ponderacion = Convert.ToInt32(ModalPonderacion.Text);
                CV.GestionCategoriaVariable(ref listaVariables, objVariables, Transaccion);
                omb.ShowMessage("Se actualizaron los valores satisfactoriamente! ", 3, "Atención");
                CargarGrillaVariables();
                NombreVariable.Text = string.Empty;
                PonderacionVariable.Text = string.Empty;
                EstadoVariable.SelectedIndex = 0;
                NombreVariable.Focus();
            }
            catch (Exception ex)
            {
                omb.ShowMessage("Error al actualizar la variable: " + ex.Message.ToString(),1,"Error");
                
            }
            
        }

        private void CargaVariables(List<CCalificacionExperta> TodasVariables)
        {

            foreach (CCalificacionExperta objVariables in TodasVariables)
            {

                TodasLasVariables.Rows.Add(new object[] {
                    objVariables.IdVariable.ToString(),
                    objVariables.NombreVariable,
                    objVariables.Ponderacion.ToString(),
                    objVariables.EstadoVariable,
                    objVariables.UsuarioRegistro,
                    objVariables.FechaRegistro.ToString()
                    });
            }
        }

        protected void btnAceptar_Click(object sender, EventArgs e)
        {
            objVariables.IdVariable = IdVariableGlobal;
            GridViewRow row = GvVariables.Rows[RowGridVariables];
            string nombreVariable = objVariables.NombreVariable;
            int booActivo = 0;
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                if (cellIndex == 4)
                {
                    Label lbl = ((Label)row.FindControl("booActivo"));
                    ImageButton ImgBnt = ((ImageButton)row.FindControl("ImgBtnInact"));
                    if (booActivo == 0)
                    {
                        if (lbl.Text == "Activo")
                        {
                            //continua - inactivar
                            objVariables.EstadoVariable = "0";
                            int Transaccion = 3;
                            listaVariables = CV.GestionCategoriaVariable(ref listaVariables, objVariables, Transaccion);
                            lbl.Text = "Inactivo";
                            ImgBnt.ToolTip = "Inactivo";
                            ImgBnt.ImageUrl = "~/Imagenes/Icons/switch-off-icon.png";
                            omb.ShowMessage("Se inactivó la variable satisfactoriamente!", 3, "Atención");
                        }
                        else
                        {
                            //continua - activar
                            objVariables.EstadoVariable = "1";
                            int Transaccion = 3;
                            listaVariables = CV.GestionCategoriaVariable(ref listaVariables, objVariables, Transaccion);
                            lbl.Text = "Activo";
                            ImgBnt.ToolTip = "Activo";
                            ImgBnt.ImageUrl = "~/Imagenes/Icons/switch-on-icon.png";
                            omb.ShowMessage("Se activó la variable satisfactoriamente!", 3, "Atención");
                        }
                    }
                }
            }           
        }
        #endregion Eventos

        /********************************************************************************/
        #region Metodos
        public bool CargarGrillaVariables()
        {
            bool resultado = false;
            try
            {
                int Transaccion = 0;
                objVariables.FechaRegistro = DateTime.Now;
                listaVariables = CV.GestionCategoriaVariable(ref listaVariables, objVariables, Transaccion);
                if (listaVariables != null)
                {
                    GrillaVariables();
                    CargaVariables(listaVariables);
                    GvVariables.DataSource = listaVariables;
                    GvVariables.PageIndex = PagIndexGvVariables;
                    GvVariables.DataBind();                    
                    resultado = true;

                    //Para los estados
                    if (listaVariables.Count > 10)
                    {
                        for (int rowIndex = 0; rowIndex < 10; rowIndex++)
                        {
                            string estado = TodasLasVariables.Rows[rowIndex]["EstadoVariable"].ToString().Trim();

                            GridViewRow row = GvVariables.Rows[rowIndex];
                            Label lbl = ((Label)row.FindControl("booActivo"));
                            ImageButton ImgBnt = ((ImageButton)row.FindControl("ImgBtnInact"));
                            if (estado == "False")
                            {
                                lbl.Text = "Inactivo";
                                ImgBnt.ToolTip = "Inactivo";
                                ImgBnt.ImageUrl = "~/Imagenes/Icons/switch-off-icon.png";
                            }
                            else
                            {
                                lbl.Text = "Activo";
                                ImgBnt.ToolTip = "Activo";
                                ImgBnt.ImageUrl = "~/Imagenes/Icons/switch-on-icon.png";
                            }
                        }
                    }
                    else
                    {
                        for (int rowIndex = 0; rowIndex < listaVariables.Count; rowIndex++)
                        {
                            string estado = TodasLasVariables.Rows[rowIndex]["EstadoVariable"].ToString().Trim();

                            GridViewRow row = GvVariables.Rows[rowIndex];
                            Label lbl = ((Label)row.FindControl("booActivo"));
                            ImageButton ImgBnt = ((ImageButton)row.FindControl("ImgBtnInact"));
                            if (estado == "False")
                            {
                                lbl.Text = "Inactivo";
                                ImgBnt.ToolTip = "Inactivo";
                                ImgBnt.ImageUrl = "~/Imagenes/Icons/switch-off-icon.png";
                            }
                            else
                            {
                                lbl.Text = "Activo";
                                ImgBnt.ToolTip = "Activo";
                                ImgBnt.ImageUrl = "~/Imagenes/Icons/switch-on-icon.png";
                            }
                        }
                    }                    
                }
                else
                {
                    omb.ShowMessage("No existen variables cargadas", 2, "Atención");
                }
            }
            catch (Exception ex)
            {
                omb.ShowMessage(ex.Message.ToString(), 1, "Error");
            }
            return resultado;
        }

        private void GrillaVariables()
        {
            DataTable grid = new DataTable();

            grid.Columns.Add("IdVariable", typeof(string));
            grid.Columns.Add("NombreVariable", typeof(string));
            grid.Columns.Add("Ponderacion", typeof(string));
            grid.Columns.Add("EstadoVariable", typeof(string));
            grid.Columns.Add("FechaRegistro", typeof(string));
            grid.Columns.Add("UsuarioRegistro", typeof(string));
            grid.Columns.Add("EditarRegistro", typeof(string));

            GvVariables.DataSource = grid;
            GvVariables.DataBind();
            TodasLasVariables = grid;
        }

        private void cambiaEstado(int paginaGrid)
        {
            int cantidad = GvVariables.Rows.Count;

            for (int rowIndex = 0; rowIndex < GvVariables.Rows.Count; rowIndex++)
            {
                string estado = TodasLasVariables.Rows[paginaGrid]["EstadoVariable"].ToString().Trim();

                GridViewRow row = GvVariables.Rows[rowIndex];
                Label lbl = ((Label)row.FindControl("booActivo"));
                ImageButton ImgBnt = ((ImageButton)row.FindControl("ImgBtnInact"));
                if (estado == "False")
                {
                    lbl.Text = "Inactivo";
                    ImgBnt.ToolTip = "Inactivo";
                    ImgBnt.ImageUrl = "~/Imagenes/Icons/switch-off-icon.png";
                    paginaGrid++;
                }
                else
                {
                    lbl.Text = "Activo";
                    ImgBnt.ToolTip = "Activo";
                    ImgBnt.ImageUrl = "~/Imagenes/Icons/switch-on-icon.png";
                    paginaGrid++;
                }
            }
        }


        #endregion Metodos

       
    }
}