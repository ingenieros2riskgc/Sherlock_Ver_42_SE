using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ListasSarlaft.Classes
{
    public class clsDTOPlanes
    {
        #region Variables
        private string _Codigo;
        private string _NombrePlan;
        private string _DescripcionPlan;
        private string _Responsable;
        private string _Estado;
        private DateTime _FechaCompromiso;
        private string _Adjuntos;
        private string _Justificacion;
        private string _CodigoRiesgo;
        private string _CodigoEvento;
        private DateTime ? _Periodo;
        private int ? _Meta;
        private int ?  _Cumplimiento;
        private int _Gestion;
        private string _Seguimiento;
        private string _Usuario;                
        private string _FechaRegistro;
        private string _Resultado;
        private int _IdGestion;
        #endregion Variables

        #region GET/SET
        public string Codigo
        {
            get { return _Codigo; }
            set { _Codigo = value; }
        }
        public string NombrePlan
        {
            get { return _NombrePlan; }
            set { _NombrePlan = value; }
        }

        public string DescripcionPlan
        {
            get { return _DescripcionPlan; }
            set { _DescripcionPlan = value; }
        }

        public string Responsable
        {
            get { return _Responsable; }
            set { _Responsable = value; }
        }

        public string Estado
        {
            get { return _Estado; }
            set { _Estado = value; }
        }
        public DateTime FechaCompromiso
        {
            get { return _FechaCompromiso; }
            set { _FechaCompromiso = value; }
        }

        public string Adjuntos
        {
            get { return _Adjuntos; }
            set { _Adjuntos = value; }
        }

        public string Justificacion
        {
            get { return _Justificacion; }
            set { _Justificacion = value; }
        }

        public string CodigoRiesgo
        {
            get { return _CodigoRiesgo; }
            set { _CodigoRiesgo = value; }
        }

        public string CodigoEvento
        {
            get { return _CodigoEvento; }
            set { _CodigoEvento = value; }
        }

        public DateTime ? Periodo
        {
            get { return _Periodo; }
            set { _Periodo = value; }
        }

        public int ?  Meta
        {
            get { return _Meta; }
            set { _Meta = value; }
        }

        public int ?  Cumplimiento
        {
            get { return _Cumplimiento; }
            set { _Cumplimiento = value; }
        }
        public int Gestion
        {
            get { return _Gestion; }
            set { _Gestion = value; }
        }

        public int IdGestion
        {
            get { return _IdGestion; }
            set { _IdGestion = value; }
        }

        public string Seguimiento
        {
            get { return _Seguimiento; }
            set { _Seguimiento = value; }
        }

        public string Resultado
        {
            get { return _Resultado; }
            set { _Resultado = value; }
        }

        public string Usuario
        {
            get { return _Usuario; }
            set { _Usuario = value; }
        } 
        public string FechaRegistro
        {
            get { return _FechaRegistro; }
            set { _FechaRegistro = value; }
        }
        #endregion GET/SET

        #region Constructor
        public clsDTOPlanes() { }

        #endregion Constructor
    }
}