using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web;

namespace ListasSarlaft.Classes
{
    public class cambiaCultura
    {
        public void CambiaCultura()
        {
            //Cambio de la Cultura del Regional Settings del panel de control para la aplciación que use esta clase.
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("es-CO", true);

            //'Cambia la configuración regional de Fecha
            Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
            Thread.CurrentThread.CurrentCulture.DateTimeFormat.DateSeparator = "/";

            Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator = ",";
            Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator = ".";

            Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencySymbol = "$";
            Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyGroupSeparator = ",";
            Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator = ".";

            Thread.CurrentThread.CurrentCulture.DateTimeFormat.AMDesignator = "AM";
            Thread.CurrentThread.CurrentCulture.DateTimeFormat.PMDesignator = "PM";
            Thread.CurrentThread.CurrentCulture.DateTimeFormat.TimeSeparator = ":";

            //Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyNegativePattern = 12; //-$ n 
            Thread.CurrentThread.CurrentCulture.NumberFormat.PositiveSign = "+";

        }
    }
}