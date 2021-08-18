using ListasSarlaft.Classes;
using ListasSarlaft.Classes.BLL.Riesgos.Parametrizacion.Riesgos;
using Microsoft.Security.Application;
using System;
using System.Collections.Generic;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ListasSarlaft.UserControls.Riesgos
{
    public partial class EstadoRiesgo : System.Web.UI.UserControl
    {
        private string IdFormulario = "5023";
        private cCuenta cCuenta = new cCuenta();
        private cRiesgo cRiesgo = new cRiesgo();

        protected void Page_Load(object sender, EventArgs e)
        {
            ScriptManager scriptManager = ScriptManager.GetCurrent(this.Page);
            scriptManager.RegisterPostBackControl(this.GVRiesgos);
            scriptManager.RegisterPostBackControl(this.IBinsertGVC);
            scriptManager.RegisterPostBackControl(this.IBupdateGVC);
            if (cCuenta.permisosConsulta(IdFormulario) == "False")
            {
                Response.Redirect("~/Formularios/Sarlaft/Admin/HomeAdmin.aspx?Denegar=1");
            }
            else
            {
                if (!Page.IsPostBack)
                {
                    mtdStard();
                    mtdInicializarValores();
                }
            }
        }

        private DataTable infoGridEstado;
        private int rowGridEstado;
        private int pagIndexEstado;

        #region SetVariables
        private DataTable InfoGridEstado
        {
            get
            {
                infoGridEstado = (DataTable)ViewState["infoGridEstado"];
                return infoGridEstado;
            }
            set
            {
                infoGridEstado = value;
                ViewState["infoGridEstado"] = infoGridEstado;
            }
        }

        private int RowGridEstado
        {
            get
            {
                rowGridEstado = (int)ViewState["rowGridEstado"];
                return rowGridEstado;
            }
            set
            {
                rowGridEstado = value;
                ViewState["rowGridEstado"] = rowGridEstado;
            }
        }

        private int PagIndexEstado
        {
            get
            {
                pagIndexEstado = (int)ViewState["pagIndexEstado"];
                return pagIndexEstado;
            }
            set
            {
                pagIndexEstado = value;
                ViewState["pagIndex"] = pagIndexEstado;
            }
        }

        private bool mtdLoadEstados(ref string strErrMsg)
        {
            #region Vars
            bool booResult = false;
            clsDTORiesgos objEstados = new clsDTORiesgos();
            List<clsDTORiesgos> lstEstados = new List<clsDTORiesgos>();
            clsBLLEstados cEstados = new clsBLLEstados();
            #endregion Vars

            lstEstados = cEstados.mtdConsultarEstados(ref lstEstados, ref strErrMsg);

            if (lstEstados != null)
            {
                mtdLoadGridEstados();
                mtdLoadGridEstados(lstEstados);
                GVRiesgos.DataSource = lstEstados;
                GVRiesgos.PageIndex = pagIndexEstado;
                GVRiesgos.DataBind();
                booResult = true;
            }
            else
            {
                strErrMsg = "No hay Estados registrados";
            }

            return booResult;
        }

        private void mtdLoadGridEstados()
        {
            DataTable grid = new DataTable();

            grid.Columns.Add("intIdEstado", typeof(string));
            grid.Columns.Add("strNombreEstado", typeof(string));
            grid.Columns.Add("strEstado", typeof(string));
            grid.Columns.Add("dtFechaRegistro", typeof(string));
            grid.Columns.Add("intIdUsuario", typeof(string));
            grid.Columns.Add("strUsuario", typeof(string));

            GVRiesgos.DataSource = grid;
            GVRiesgos.DataBind();
            InfoGridEstado = grid;
        }
        #endregion SetVariables

        protected void mtdStard()
        {
            string strErrMsg = string.Empty;

            if (!mtdLoadEstados(ref strErrMsg))
            {
                omb.ShowMessage(strErrMsg, 2, "Atención");
            }
        }

        private void mtdInicializarValores()
        {
            pagIndexEstado = 0;
        }

        private void mtdLoadGridEstados(List<clsDTORiesgos> lstTratamiento)
        {
            string strErrMsg = string.Empty;

            foreach (clsDTORiesgos objTratamiento in lstTratamiento)
            {

                InfoGridEstado.Rows.Add(new object[] {
                    objTratamiento.intIdEstado.ToString().Trim(),
                    objTratamiento.strNombreEstado.ToString().Trim(),
                    objTratamiento.strEstado.ToString().Trim(),
                    objTratamiento.dtFechaRegistro.ToString().Trim(),
                    objTratamiento.intIdUsuario.ToString().Trim(),
                    objTratamiento.strUsuario.ToString().Trim()
                    });
            }
        }

        // Eventos - Yoendy
        protected void IBinsertGVC_Click(object sender, ImageClickEventArgs e)
        {
            string strErrMsg = string.Empty;
            if (mtdInsertarEstado(ref strErrMsg, 1) == true)
            {
                omb.ShowMessage(strErrMsg, 3, "Atención");
                mtdResetFields();
                mtdStard();
            }
        }
        protected void btnInsertarNuevo_Click(object sender, ImageClickEventArgs e)
        {
            BodyFormT.Visible = true;
            BodyGridT.Visible = false;
            IBinsertGVC.Visible = true;
            IBupdateGVC.Visible = false;
            TXNombreEstado.Focus();
        }

        protected void btnImgCancelar_Click(object sender, ImageClickEventArgs e)
        {
            mtdResetFields();
        }

        protected void GVtratamientos_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            RowGridEstado = Convert.ToInt16(e.CommandArgument);
            switch (e.CommandName)
            {
                case "Seleccionar":
                    mtdShowUpdate(RowGridEstado);
                    IBinsertGVC.Visible = false;
                    IBupdateGVC.Visible = true;
                    BodyFormT.Visible = true;
                    BodyGridT.Visible = false;
                    break;
            }
        }
        protected void GVtratamientos_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            pagIndexEstado = e.NewPageIndex;
            GVRiesgos.PageIndex = pagIndexEstado;
            GVRiesgos.DataBind();
            string strErrMsg = "";
            mtdLoadEstados(ref strErrMsg);
        }

        protected void IBupdateGVC_Click(object sender, ImageClickEventArgs e)
        {
            string strErrMsg = string.Empty;
            if (mtdInsertarEstado(ref strErrMsg, 0) == true)
            {
                omb.ShowMessage(strErrMsg, 3, "Atención");
                mtdResetFields();
                mtdStard();
            }
            else
            {
                omb.ShowMessage(strErrMsg, 1, "Atención");
            }
        }

        // Métodos - Yoendy
        private bool mtdInsertarEstado(ref string strErrMsg, int evento)
        {
            #region Vars
            bool booResult = false;
            clsDTORiesgos objEstados = new clsDTORiesgos();
            clsBLLEstados cEstados = new clsBLLEstados();
           
            objEstados.strNombreEstado = Sanitizer.GetSafeHtmlFragment(TXNombreEstado.Text.Trim());
            objEstados.intIdUsuario = Convert.ToInt32(Session["IdUsuario"].ToString());
            objEstados.strEstado = Sanitizer.GetSafeHtmlFragment(cbEstado.Text.Trim());
            objEstados.dtFechaRegistro = DateTime.Now;
            #endregion Vars

            if (evento == 1)
            {
                objEstados.intIdEstado = 1;
                int result = cEstados.mtdInsertarEstado(objEstados, ref strErrMsg, 1);
                if (result == 1)
                {
                    strErrMsg = "Nombre de Estado " + TXNombreEstado.Text.ToUpper() + " registrado exitosamente!";
                    booResult = true;
                }
                else
                {
                    strErrMsg = "Error al registrar el tratamiento";
                    booResult = false;
                }
            }
            else
            {
                objEstados.intIdEstado = Convert.ToInt32(txtId.Text);
                int relacion = 0;
                if (cbEstado.SelectedIndex == 2)
                {
                    // Inhabilitado - Verifico primero relacion  
                    relacion = cEstados.verificaRelacionEstado(objEstados, ref strErrMsg, 2);
                }

                if (relacion == 0)
                {
                    int result = cEstados.mtdInsertarEstado(objEstados, ref strErrMsg, 0);
                    if (result == 1)
                    {
                        strErrMsg = "Nombre de Estado " + TXNombreEstado.Text.ToUpper() + " actualizado exitosamente!";
                        booResult = true;
                    }
                    else
                    {
                        strErrMsg = "Error al registrar el estado";
                        booResult = false;
                    }
                }
                else
                {
                    strErrMsg = "No es posible inhabilitar el estado porque tiene riesgos asociados";
                    cbEstado.SelectedIndex = 0;
                    booResult = false;
                }
            }
            return booResult;
        }

        protected void mtdResetFields()
        {
            BodyFormT.Visible = false;
            BodyGridT.Visible = true;

            txtId.Text = string.Empty;
            TXNombreEstado.Text = string.Empty;
            tbxUsuarioCreacion.Text = string.Empty;
            txtFecha.Text = string.Empty;
        }

        protected void mtdShowUpdate(int Rowgrid)
        {
            GridViewRow row = GVRiesgos.Rows[Rowgrid];
            System.Collections.Specialized.IOrderedDictionary colsNoVisible = GVRiesgos.DataKeys[Rowgrid].Values;
            txtId.Text = row.Cells[0].Text;
            TXNombreEstado.Text = ((Label)row.FindControl("strNombreEstado")).Text;
            tbxUsuarioCreacion.Text = colsNoVisible[0].ToString();
            txtFecha.Text = colsNoVisible[1].ToString();
            string estado = string.Empty;

            estado = colsNoVisible[2].ToString();
            if (estado == "False")
            {
                cbEstado.SelectedIndex = 2;
            }
            else if (estado == "True")
            {
                cbEstado.SelectedIndex = 1;
            }
            else
            {
                cbEstado.SelectedIndex = 0;
            }
        }

    } // Fin espacio de nombres
}