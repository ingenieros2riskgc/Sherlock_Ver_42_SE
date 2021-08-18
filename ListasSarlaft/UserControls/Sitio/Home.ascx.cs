using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;


using System.Web.Configuration;



namespace ListasSarlaft.UserControls
{
    public partial class Home : System.Web.UI.UserControl
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                string valor = Request.QueryString["Denegar"];
                
                if (Convert.ToInt32(Request.QueryString["Denegar"]) == 1)
                {
                     Mensaje("Expiró el tiempo de inactividad de la aplicación, presione el botón 'OK' e ingrese nuevamente!");
                     imgInfo.ImageUrl = "~/Imagenes/Icons/RelojArena.gif";                    
                }
                else if (Convert.ToInt32(Request.QueryString["NP"]) == 2)
                {                                        
                    Mensaje("No tiene los permisos suficientes para llevar a cabo esta acción!");
                    imgInfo.ImageUrl = "~/Imagenes/Icons/Alerta.png";                     
                }
            }
        }

        private void Mensaje(String Mensaje)
        {
            lblMsgBox.Text = Mensaje;
            mpeMsgBox.Show();
        }       
    }
}