using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace clsDTO
{
    public class clsDTOImpactoCualitativo
    {
        private string IdImpactoCualitativo;
        private string nombre;
        private string idUsuario;
        private string fechaRegistro;

        public string idImpactoCualitativo
        {
            get { return IdImpactoCualitativo; }
            set { IdImpactoCualitativo = value; }
        }

        public string Nombre
        {
            get { return nombre; }
            set { nombre = value; }
        }
        public string IdUsuario
        {
            get { return idUsuario; }
            set { idUsuario = value; }
        }
        public string FechaRegistro
        {
            get { return fechaRegistro; }
            set { fechaRegistro = value; }
        }

        public clsDTOImpactoCualitativo() { }

        public clsDTOImpactoCualitativo(string IdImpactoCualitativo, string nombre, string idUsuario, string fechaRegistro)
        {
            this.idImpactoCualitativo = IdImpactoCualitativo;
            this.Nombre = nombre;
            this.IdUsuario = idUsuario;
            this.FechaRegistro = fechaRegistro;
        }

    }
}
