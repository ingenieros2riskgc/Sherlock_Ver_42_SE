using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Text.RegularExpressions;
using ListasSarlaft.Classes;
using Microsoft.Security.Application;
using System.Data;
using clsLogica;
using clsDTO;
using clsDatos;

namespace ListasSarlaft.UserControls.Eventos
{
    public partial class ImpactoCualitativo : System.Web.UI.UserControl
    {
        string IdFormulario = "5010";
        cCuenta cCuenta = new cCuenta();
        private cError cError = new cError();
        private cDataBase cDataBase = new cDataBase();
        private int pagIndex;
        private int rowGrid;
        private DataTable infoGrid;
        clsDtImpactoCualitativo clsDt = new clsDtImpactoCualitativo();

        #region Page_Load
        protected void Page_Load(object sender, EventArgs e)
        {
            //int? IdUsuario = Convert.ToInt32(Session["IdUsuario"]);
            //if (IdUsuario == 0)
            //{
            //    IdUsuario = null;
            //}
            //if (string.IsNullOrEmpty(IdUsuario.ToString().Trim()))
            //{
            //    Response.Redirect("~/Formularios/Sarlaft/Admin/HomeAdmin.aspx?Denegar=1");
            //}
            //else
            //{
            //    if (cCuenta.permisosConsulta(IdFormulario) == "NOPERMISO")
            //    {
            //        Response.Redirect("~/Formularios/Sarlaft/Admin/HomeAdmin.aspx?NP=2");
            //    }
            //}
            if (!Page.IsPostBack)
            {
                Page.Form.Attributes.Add("enctype", "multipart/form-data");
                ScriptManager scriptManager = ScriptManager.GetCurrent(this.Page);
                mtdInicializarValoresImpactoCualitativo();
                mtdLoadGridViewImpactoCualitativo();
                
                lblNombreImpacto.Visible = false;
                txtNombreImpacto.Visible = false;
                btnImgInsertar.Visible = false;
                btnImgCancelar.Visible = false;

                btnImgActualizar.Visible = false;
                imgBtnInsertar.Visible = true;
                
            }
            }
        #endregion

        #region Grid Properties
        private int PagIndex
        {
            get
            {
                pagIndex = (int)ViewState["pagIndex"];
                return pagIndex;
            }
            set
            {
                pagIndex = value;
                ViewState["pagIndex"] = pagIndex;
            }
        }

        private DataTable InfoGrid
        {
            get
            {
                infoGrid = (DataTable)ViewState["infGrid2"];
                return infoGrid;
            }
            set
            {
                infoGrid = value;
                ViewState["infGrid2"] = infoGrid;
            }
        }

        private DataTable infoGridRequerimientos;
        private DataTable InfoGridRequerimientos
        {
            get
            {
                infoGridRequerimientos = (DataTable)ViewState["infGrid2"];
                return infoGridRequerimientos;
            }
            set
            {
                infoGridRequerimientos = value;
                ViewState["infGrid2"] = infoGridRequerimientos;
            }
        }

        private int RowGrid
        {
            get
            {
                rowGrid = (int)ViewState["rowGrid"];
                return rowGrid;
            }
            set
            {
                rowGrid = value;
                ViewState["rowGrid"] = rowGrid;
            }
        }
        private void mtdMensaje(string Mensaje)
        {
            lblMsgBox.Text = Mensaje;
            mpeMsgBox.Show();
        }
        #endregion

        #region gvImpactoCualitativo
        protected void gvImpactoCualitativo_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            PagIndex = e.NewPageIndex;
            gvImpactoCualitativo.PageIndex = PagIndex;
            gvImpactoCualitativo.DataSource = infoGrid;
            gvImpactoCualitativo.DataBind();

            mtdLoadGridViewImpactoCualitativo();
        }

        private void mtdLoadGridViewImpactoCualitativo()
        {
            mtdLoadGridImpactoCualitativo();
            mtdLoadInfoGridImpactoCualitativo();
        }

        private void mtdLoadGridImpactoCualitativo()
        {
            DataTable grid = new DataTable();
            grid.Columns.Add("idImpactoCualitativo", typeof(string));
            grid.Columns.Add("Nombre", typeof(string));
            grid.Columns.Add("IdUsuario", typeof(string));
            grid.Columns.Add("FechaRegistro", typeof(string));
            gvImpactoCualitativo.DataSource = grid;
            gvImpactoCualitativo.DataBind();
            InfoGridRequerimientos = grid;
        }

        private void mtdLoadInfoGridImpactoCualitativo()
        {
            string strErrMsg = string.Empty;
            clsImpactoCualitativo cImpCual = new clsImpactoCualitativo();
            List<clsDTOImpactoCualitativo> lstImpCual = new List<clsDTOImpactoCualitativo>();

            lstImpCual = cImpCual.mtdCargarInfoImpactoCualitativo(ref strErrMsg);

            if (lstImpCual != null)
            {
                mtdLoadGridImpactoCualitativo();
                mtdLoadInfoGrid(lstImpCual);
                gvImpactoCualitativo.DataSource = lstImpCual;
            }
        }

        private void mtdLoadInfoGrid(List<clsDTOImpactoCualitativo> lstImpCual)
        {
            DataTable dtInfo = new DataTable();

            string strId = string.Empty;

            dtInfo = clsDt.ConsultaImpactoCualitativo();
            if (dtInfo.Rows.Count > 0)
            {
                for (int rows = 0; rows < dtInfo.Rows.Count; rows++)
                {
                    InfoGridRequerimientos.Rows.Add(new Object[]
                        {
                        dtInfo.Rows[rows]["idImpactoCualitativo"].ToString().Trim(),
                        dtInfo.Rows[rows]["Nombre"].ToString().Trim(),
                        dtInfo.Rows[rows]["IdUsuario"].ToString().Trim(),
                        dtInfo.Rows[rows]["FechaRegistro"].ToString().Trim()
                }
                        );
                }
                gvImpactoCualitativo.DataSource = InfoGridRequerimientos;
                gvImpactoCualitativo.DataBind();
            }
        }
        
        #endregion Loads

        #region propiedades
        private void mtdInicializarValores()
        {
            PagIndex = 0;
        }

        private void mtdMensajeImpactoCualitativo(string Mensaje)
        {
            lblMsgBox.Text = Mensaje;
            mpeMsgBox.Show();
        }

        private DataTable InfoGridImpactoCualitativo
        {
            get
            {
                infoGrid = (DataTable)ViewState["infGrid2"];
                return infoGrid;
            }
            set
            {
                infoGrid = value;
                ViewState["infGrid2"] = infoGrid;
            }
        }

        private int pagIndexInfoGridImpactoCualitativo;
        private int PagIndexInfoGridImpactocualitativo
        {
            get
            {
                pagIndexInfoGridImpactoCualitativo = (int)ViewState["pagIndexInfoGridImpactoCualitativo"];
                return pagIndexInfoGridImpactoCualitativo;
            }
            set
            {
                pagIndexInfoGridImpactoCualitativo = value;
                ViewState["pagIndexInfoGridImpactoCualitativo"] = pagIndexInfoGridImpactoCualitativo;
            }
        }
        #endregion

        protected void gvImpactoCualitativo_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            string strErrMsg = string.Empty;

            //RowGrid = Convert.ToInt16(e.CommandArgument);
            switch (e.CommandName)
            {
                case "Modificar":
                    BtnShowModificar(ref strErrMsg);
                    break;
                case "Eliminar":
                    RowGrid = (Convert.ToInt16(gvImpactoCualitativo.PageSize) * PagIndex) + Convert.ToInt16(e.CommandArgument);
                    string strIdUserSoporte = InfoGridRequerimientos.Rows[RowGrid]["idImpactoCualitativo"].ToString().Trim();
                    
                    break;
            }
        }

        #region Buttons Crear Impacto
        protected void imgBtnInsertar_Click(object sender, ImageClickEventArgs e)
        {
            filaDetalle.Visible = true;
            filaGrid.Visible = false;
            lblNombreImpacto.Visible = true;
            txtNombreImpacto.Visible = true;
            btnImgInsertar.Visible = true;
            btnImgCancelar.Visible = true;

            btnImgActualizar.Visible = false;
            imgBtnInsertar.Visible = false;
        }

        protected void btnImgInsertar_Click(object sender, ImageClickEventArgs e)
        {

            string strErrMsg = string.Empty;

            try
            {
                try
                {
                    strErrMsg = "Impacto cualitativo registrado exitosamente.";

                    mtdInsertarImpCual(
                            string.Empty,
                            Session["IdUsuario"].ToString().Trim(),
                            Sanitizer.GetSafeHtmlFragment(txtNombreImpacto.Text.Trim()),
                            ref strErrMsg);
                    
                    gvImpactoCualitativo.Visible = true;
                    //mtdMensaje("Impacto cualitativo registrado exitosamente.");
                    string strMensaje = string.Format("Impacto cualitativo registrado exitosamente.");
                    Mensaje(strMensaje);
                }
                catch (Exception except)
                {
                    strErrMsg = "Atención." + except.Message.ToString();
                }
                mtdLoadGridViewImpactoCualitativo();
                lblNombreImpacto.Visible = false;
                txtNombreImpacto.Visible = false;
                btnImgInsertar.Visible = false;
                btnImgCancelar.Visible = false;

                btnImgActualizar.Visible = false;
                imgBtnInsertar.Visible = true;

                filaDetalle.Visible = false;
                filaGrid.Visible = true;
            }
            catch (Exception except)
            {
                //strErrMsg = "Error al registrar el impacto cualitativo." + except.Message.ToString();
                string strMensaje = string.Format("Error al registrar el impacto cualitativo.");
                Mensaje(strMensaje);

            }
        }
        private void Mensaje(string Mensaje)
        {
            lblMsgBox.Text = Mensaje;
            mpeMsgBox.Show();
        }

        private void mtdInsertarImpCual(string idImpactoCualitativo, string IdUsuario, string Nombre,  ref string strErrMsg)
        {
            clsImpactoCualitativo cImpCual = new clsImpactoCualitativo();
            clsDTOImpactoCualitativo objImpCual = new clsDTOImpactoCualitativo(idImpactoCualitativo, Nombre, IdUsuario, string.Empty);

            cImpCual.mtdInsertarImpactoCualitativo(objImpCual, ref strErrMsg);
        }
        #endregion

        #region Actualizar Impacto
        private void BtnShowModificar(ref string strErrMsg)
        {
            
            
        }
        private void mtdInicializarValoresImpactoCualitativo()
        {
            PagIndex = 0;
        }
        protected void btnImgActualizar_Click(object sender, ImageClickEventArgs e)
        {
            string strErrMsg = string.Empty;


            try
            {
                string idImpactoCualitativo = Session["idImpactoCualitativo"].ToString();
                mtdActualizarImpCual(
                //InfoGrid.Rows[RowGrid]["idImpactoCualitativo"].ToString().Trim(),
                idImpactoCualitativo,
                Sanitizer.GetSafeHtmlFragment(txtNombreImpacto.Text.Trim()),
                Session["IdUsuario"].ToString().Trim(),
                ref strErrMsg
                );
                //mtdMensaje("Impacto cualitativo actualizado exitosamente.");
                string strMensaje = string.Format("Impacto cualitativo actualizado exitosamente.");
                Mensaje(strMensaje);
            }
            catch (Exception except)
            {
                strErrMsg = "Atención." + except.Message.ToString();
            }
            try { 
            
            mtdLoadGridViewImpactoCualitativo();
                lblNombreImpacto.Visible = false;
                txtNombreImpacto.Visible = false;
                btnImgInsertar.Visible = false;
                btnImgActualizar.Visible = false;
                imgBtnInsertar.Visible = true;
                filaDetalle.Visible = false;
                filaGrid.Visible = true;
                mtdLoadGridImpactoCualitativo();
            mtdLoadInfoGridImpactoCualitativo();
        }
            catch (Exception except)
            {
                //strErrMsg = "Error al actualizar el impacto cualitativo." + except.Message.ToString();
                string strMensaje = string.Format("Error al actualizar el impacto cualitativo.");
                Mensaje(strMensaje);
            }
        }

        private void mtdActualizarImpCual(string idImpactoCualitativo, string Nombre, string IdUsuario, ref string strErrMsg)
        {
            clsImpactoCualitativo cImpCual = new clsImpactoCualitativo();
            clsDTOImpactoCualitativo objImpCual = new clsDTOImpactoCualitativo(idImpactoCualitativo, Nombre, IdUsuario, string.Empty);

            cImpCual.mtdActualizarImpactoCualitativo(objImpCual, ref strErrMsg);
        }
        #endregion

        #region Eliminar impacto
        private void BtnDelete_Click(ref string strErrMsg)
        {
            

            try
            {
                borrarUsuario(
                    InfoGrid.Rows[RowGrid]["idImpactoCualitativo"].ToString().Trim()
                    );
                //mtdMensaje("Impacto cualitativo eliminado con éxito.");
                btnImgokEliminar.Visible = false;
                string strMensaje = string.Format("Impacto cualitativo eliminado con éxito.");
                Mensaje(strMensaje);
            }
            catch (Exception ex)
            {
                //mtdMensaje("Error eliminando el impacto cualitativo." + ex.Message);
                string strMensaje = string.Format("Error eliminando el impacto cualitativo.");
                Mensaje(strMensaje);
            }
            mtdLoadGridViewImpactoCualitativo();
        }

        private string ConsultaDelete()
        {
            DataTable dtInfo = new DataTable();
            dtInfo = clsDt.ConsultaEliminar();
            string strId = string.Format(dtInfo.Rows[RowGrid]["idImpactoCualitativo"].ToString().Trim());
            return strId;
        }

        public void borrarUsuario(String IdImpactoCualitativo)
        {
            try
            {
                cDataBase.conectar();
                cDataBase.ejecutarQuery("DELETE FROM [Riesgos].[tblImpactoCualitativo] WHERE (idImpactoCualitativo = " + IdImpactoCualitativo + ")");
                cDataBase.desconectar();
            }
            catch (Exception ex)
            {
                cDataBase.desconectar();
                cError.errorMessage(ex.Message + ", " + ex.StackTrace);
                throw new Exception(ex.Message);
            }
        }

        #endregion

        protected void btnImgCancelar_Click(object sender, ImageClickEventArgs e)
        {
            lblNombreImpacto.Visible = false;
            txtNombreImpacto.Visible = false;
            btnImgInsertar.Visible = false;
            btnImgActualizar.Visible = false;
            imgBtnInsertar.Visible = false;
        }

        protected void btnImgEliminar_Click(object sender, ImageClickEventArgs e)
        {
            if (cCuenta.permisosBorrar(IdFormulario) == "False")
            {
                //omb.ShowMessage("No tiene los permisos suficientes para llevar a cabo esta acción.", 2, "Atención");
                lblMsgBox.Text = "No tiene los permisos suficientes para llevar a cabo esta acción.";
                
                mpeMsgBox.Show();
            }
            else
            {
                lblMsgBox.Text = "Desea eliminar la información de la Base de Datos?";
                btnImgokEliminar.Visible = true;
                mpeMsgBox.Show();
            }
        }

        protected void btnImgokEliminar_Click(object sender, EventArgs e)
        {
            string strErrMsg = string.Empty;

            BtnDelete_Click(ref strErrMsg);
        }

        protected void gvImpactoCualitativo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (gvImpactoCualitativo.SelectedRow.RowType == DataControlRowType.DataRow)
            {
                try
                {
                    
                    filaDetalle.Visible = true;
                    filaGrid.Visible = false;
                    lblNombreImpacto.Visible = true;
                    txtNombreImpacto.Visible = true;
                    btnImgInsertar.Visible = false;
                    btnImgActualizar.Visible = true;
                    imgBtnInsertar.Visible = false;
                    txtNombreImpacto.Text = gvImpactoCualitativo.SelectedRow.Cells[1].Text.Trim();//InfoGrid.Rows[RowGrid]["Nombre"].ToString().Trim();
                    Session["idImpactoCualitativo"] = gvImpactoCualitativo.SelectedDataKey[0].ToString().Trim();
                }
                catch(Exception ex)
                {
                    lblMsgBox.Text = "Error! Mostrando los datos: "+ex.Message;
                    mpeMsgBox.Show();
                }
                
            }
        }
    }
}
