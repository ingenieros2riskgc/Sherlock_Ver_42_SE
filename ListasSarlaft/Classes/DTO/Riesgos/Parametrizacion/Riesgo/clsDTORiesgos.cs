using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ListasSarlaft.Classes
{
    public class clsDTORiesgos
    {
        #region Variables
        private int _IdEstado;
        private string _NombreEstado;
        private string _Estado;
        private int _IdUsuario;
        private string _Usuario;
        private DateTime _FechaRegistro;
        #endregion Variables
        #region Get/Set
        public int intIdEstado
        {
            get { return _IdEstado; }
            set { _IdEstado = value; }
        }
        public string strNombreEstado
        {
            get { return _NombreEstado; }
            set { _NombreEstado = value; }
        }
        public string strEstado
        {
            get { return _Estado; }
            set { _Estado = value; }
        }
        public int intIdUsuario
        {
            get { return _IdUsuario; }
            set { _IdUsuario = value; }
        }
        public string strUsuario
        {
            get { return _Usuario; }
            set { _Usuario = value; }
        }
        public DateTime dtFechaRegistro
        {
            get { return _FechaRegistro; }
            set { _FechaRegistro = value; }
        }
        #endregion
        #region Constructor
        public clsDTORiesgos() { }
        #endregion Constructor
    }
}