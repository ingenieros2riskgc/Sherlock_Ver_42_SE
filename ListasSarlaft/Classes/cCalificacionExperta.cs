using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ListasSarlaft.Classes
{
    public class CCalificacionExperta
    {
        #region Variables
        private int _IdVariable;
        private string _NombreVar;
        private int _PonderacionVar;
        private string _EstadoVar;
        private string _NombreUsuario;
        private DateTime _Hoy;
        /*********CATEGORIAS*******************/
        private int _IdCategoria;
        private string _NombreVariableCategoria;
        private string _NombreCategoria;
        private int _PonderacionCategoria;
        /**************************************/

        /*********PUNTOS DE CORTE**************/
        private int _Min;
        private int _Max;
        private int _IdPuntoCorte;
        private int _IdFrecuenciaEventos;
        private string _NombreFrecuencia;

        /**************************************/


        #endregion Variables

        #region GET/SET
        public int IdVariable
        {
            get { return _IdVariable; }
            set { _IdVariable = value; }
        }
        public string NombreVariable
        {
            get { return _NombreVar; }
            set { _NombreVar = value; }
        }

        public int Ponderacion
        {
            get { return _PonderacionVar; }
            set { _PonderacionVar = value; }
        }

        public string EstadoVariable
        {
            get { return _EstadoVar; }
            set { _EstadoVar = value; }
        }

        public string UsuarioRegistro
        {
            get { return _NombreUsuario; }
            set { _NombreUsuario = value; }
        }

        public DateTime FechaRegistro
        {
            get { return _Hoy; }
            set { _Hoy = value; }
        }

        /*********CATEGORIAS*******************/
        public int IdCategoria
        {
            get { return _IdCategoria; }
            set { _IdCategoria = value; }
        }

        public string NombreVariableCategoria
        {
            get { return _NombreVariableCategoria; }
            set { _NombreVariableCategoria = value; }
        }

        public string NombreCategoria
        {
            get { return _NombreCategoria; }
            set { _NombreCategoria = value; }
        }

        public int PonderacionCategoria
        {
            get { return _PonderacionCategoria; }
            set { _PonderacionCategoria = value; }
        }

        /**************************************/

        /*********PUNTOS DE CORTE**************/
        public int Min
        {
            get { return _Min; }
            set { _Min = value; }
        }

        public int Max
        {
            get { return _Max; }
            set { _Max = value; }
        }

        public int IdPuntoCorte
        {
            get { return _IdPuntoCorte; }
            set { _IdPuntoCorte = value; }
        }

        public int IdFrecuenciaEventos
        {
            get { return _IdFrecuenciaEventos; }
            set { _IdFrecuenciaEventos = value; }
        }

        public string NombreFrecuencia
        {
            get { return _NombreFrecuencia; }
            set { _NombreFrecuencia = value; }
        }

        /**************************************/
        #endregion GET/SET

        #region Constructor
        public CCalificacionExperta()
        {

        }
        #endregion Contructor
    }
}