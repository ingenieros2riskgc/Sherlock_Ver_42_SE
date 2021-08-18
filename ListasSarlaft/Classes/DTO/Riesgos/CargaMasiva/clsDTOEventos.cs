using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ListasSarlaft.Classes.DTO.Riesgos.CargaMasiva
{
    public class clsDTOEventos
    {
        #region Definicion Parametros
        private int _IdEvento;
        private string _PrefixEvento;
        private int _IdEmpresa;
        private int _IdRegion;
        private int _IdPais;
        private int _IdDepartamento;
        private int _IdCiudad;
        private int _IdOficinaSucursal;
        private string _DetalleUbicacion;
        private string _DescripcionEvento;
        private int _IdServicio;
        private int _IdSubServicio;
        private DateTime _FechaInicio;
        private string _HoraInicio;
        private DateTime _FechaFinalizacion;
        private string _HoraFinalizacion;
        private DateTime _FechaDescubrimiento;
        private string _HoraDescubrimiento;
        private int _IdCanal;
        private int _IdGeneraEvento;
        private int _GeneraEvento;
        private string _NomGeneradorEvento;
        private string _cuantiaperdida;
        private int _IdUsuario;
        private DateTime _fechaEvento;
        private int _idCadenaValor;

        private int _idMacroProceso;

        private int _idProceso;

        private int _idSubproceso;

        private string _idActividad;

        private int _IdClase;

        private int _IdSubClase;

        private int _IdTipoPerdidaEvento;

        private int _IdLineaProceso;

        private int _IdSubLineaProceso;

        private int _AfectaContinudad;

        private int _IdEstado;

        private string _Observaciones;

        private string _CuentaPUC;

        private string _CuentaOrden;

        private string _TasaCambio1;

        private string _ValorPesos1;

        private string _ValorRecuperadoTotal;

        private string _Moneda2;

        private string _TasaCambio2;

        private string _ValorPesos2;

        private string _Recuperacion;

        private DateTime? _FechaContabilidad;

        private string _HoraContabilidad;
        private string _ImpactoCualitativo;

        //******CAMPOS NUEVOS******
        private DateTime? _FechaRecuperacion;
        private string _HoraRecuperacion;
        private string _CuantiaRecuperacion;
        private string _CuantiaOtraRecuperacion;
        private string _CuantiaNeta;
        //******CAMPOS NUEVOS******

        #endregion Definicion Parametros
        #region Get/Set
        public int intIdEvento
        {
            get { return _IdEvento; }
            set { _IdEvento = value; }
        }
        public string strPrefixEvento
        {
            get { return _PrefixEvento; }
            set { _PrefixEvento = value; }
        }
        public int intIdEmpresa
        {
            get { return _IdEmpresa; }
            set { _IdEmpresa = value; }
        }
        public int intIdRegion
        {
            get { return _IdRegion; }
            set { _IdRegion = value; }
        }
        public int intIdPais
        {
            get { return _IdPais; }
            set { _IdPais = value; }
        }
        public int intIdDepartamento
        {
            get { return _IdDepartamento; }
            set { _IdDepartamento = value; }
        }
        public int intIdCiudad
        {
            get { return _IdCiudad; }
            set { _IdCiudad = value; }
        }
        public int intIdOficinaSucursal
        {
            get { return _IdOficinaSucursal; }
            set { _IdOficinaSucursal = value; }
        }
        public string strDetalleUbicacion
        {
            get { return _DetalleUbicacion; }
            set { _DetalleUbicacion = value; }
        }
        public string strDescripcionEvento
        {
            get { return _DescripcionEvento; }
            set { _DescripcionEvento = value; }
        }
        public int intIdServicio
        {
            get { return _IdServicio; }
            set { _IdServicio = value; }
        }
        public int intIdSubServicio
        {
            get { return _IdSubServicio; }
            set { _IdSubServicio = value; }
        }
        public DateTime dtFechaInicio
        {
            get { return _FechaInicio; }
            set { _FechaInicio = value; }
        }
        public string strHoraInicio
        {
            get { return _HoraInicio; }
            set { _HoraInicio = value; }
        }
        public DateTime dtFechaFinalizacion
        {
            get { return _FechaFinalizacion; }
            set { _FechaFinalizacion = value; }
        }
        public string strHoraFinalizacion
        {
            get { return _HoraFinalizacion; }
            set { _HoraFinalizacion = value; }
        }
        public DateTime dtFechaDescubrimiento
        {
            get { return _FechaDescubrimiento; }
            set { _FechaDescubrimiento = value; }
        }
        public string strHoraDescubrimiento
        {
            get { return _HoraDescubrimiento; }
            set { _HoraDescubrimiento = value; }
        }
        public int intIdCanal
        {
            get { return _IdCanal; }
            set { _IdCanal = value; }
        }
        public int intIdGeneraEvento
        {
            get { return _IdGeneraEvento; }
            set { _IdGeneraEvento = value; }
        }
        public int intGeneraEvento
        {
            get { return _GeneraEvento; }
            set { _GeneraEvento = value; }
        }
        public string strNomGeneradorEvento
        {
            get { return _NomGeneradorEvento; }
            set { _NomGeneradorEvento = value; }
        }
        public string strCuantiaperdida
        {
            get { return _cuantiaperdida; }
            set { _cuantiaperdida = value; }
        }
        public int intIdUsuario
        {
            get { return _IdUsuario; }
            set { _IdUsuario = value; }
        }
        public DateTime dtfechaEvento
        {
            get { return _fechaEvento; }
            set { _fechaEvento = value; }
        }
        public int intIdCadenaValor
        {
            get { return _idCadenaValor; }
            set { _idCadenaValor = value; }
        }
        public int intIdMacroProceso
        {
            get { return _idMacroProceso; }
            set { _idMacroProceso = value; }
        }
        public int intIdProceso
        {
            get { return _idProceso; }
            set { _idProceso = value; }
        }
        public int intIdSubproceso
        {
            get { return _idSubproceso; }
            set { _idSubproceso = value; }
        }
        public string intIdActividad
        {
            get { return _idActividad; }
            set { _idActividad = value; }
        }
        public int intIdClase
        {
            get { return _IdClase; }
            set { _IdClase = value; }
        }
        public int intIdSubClase
        {
            get { return _IdSubClase; }
            set { _IdSubClase = value; }
        }
        public int intIdTipoPerdidaEvento
        {
            get { return _IdTipoPerdidaEvento; }
            set { _IdTipoPerdidaEvento = value; }
        }
        public int intIdLineaProceso
        {
            get { return _IdLineaProceso; }
            set { _IdLineaProceso = value; }
        }
        public int intIdSubLineaProceso
        {
            get { return _IdSubLineaProceso; }
            set { _IdSubLineaProceso = value; }
        }
        public int intAfectaContinudad
        {
            get { return _AfectaContinudad; }
            set { _AfectaContinudad = value; }
        }
        public int intIdEstado
        {
            get { return _IdEstado; }
            set { _IdEstado = value; }
        }
        public string strObservaciones
        {
            get { return _Observaciones; }
            set { _Observaciones = value; }
        }
        public string strCuentaPUC
        {
            get { return _CuentaPUC; }
            set { _CuentaPUC = value; }
        }
        public string strCuentaOrden
        {
            get { return _CuentaOrden; }
            set { _CuentaOrden = value; }
        }
        public string strTasaCambio1
        {
            get { return _TasaCambio1; }
            set { _TasaCambio1 = value; }
        }
        public string strValorPesos1
        {
            get { return _ValorPesos1; }
            set { _ValorPesos1 = value; }
        }
        public string strValorRecuperadoTotal
        {
            get { return _ValorRecuperadoTotal; }
            set { _ValorRecuperadoTotal = value; }
        }
        public string strMoneda2
        {
            get { return _Moneda2; }
            set { _Moneda2 = value; }
        }
        public string strTasaCambio2
        {
            get { return _TasaCambio2; }
            set { _TasaCambio2 = value; }
        }
        public string strValorPesos2
        {
            get { return _ValorPesos2; }
            set { _ValorPesos2 = value; }
        }
        public string intRecuperacion
        {
            get { return _Recuperacion; }
            set { _Recuperacion = value; }
        }
        public DateTime? dtFechaContabilidad
        {
            get { return _FechaContabilidad; }
            set { _FechaContabilidad = value; }
        }
        public string strHoraContabilidad
        {
            get { return _HoraContabilidad; }
            set { _HoraContabilidad = value; }
        }
        public string strImpactoCualitativo
        {
            get { return _ImpactoCualitativo; }
            set { _ImpactoCualitativo = value; }
        }

        //******CAMPOS NUEVOS******
        public DateTime? dtFechaRecuperacion
        {
            get { return _FechaRecuperacion; }
            set { _FechaRecuperacion = value; }
        }
        public string strHoraRecuperacion
        {
            get { return _HoraRecuperacion; }
            set { _HoraRecuperacion = value; }
        }
        public string strCuantiaRecup
        {
            get { return _CuantiaRecuperacion; }
            set { _CuantiaRecuperacion = value; }
        }
        public string strCuantiaOtraRecup
        {
            get { return _CuantiaOtraRecuperacion; }
            set { _CuantiaOtraRecuperacion = value; }
        }
        public string strCuantiaNeta
        {
            get { return _CuantiaNeta; }
            set { _CuantiaNeta = value; }
        }
        //******CAMPOS NUEVOS******
        #endregion Get/Set
        #region Constructor
        public clsDTOEventos() { }
        #endregion Constructor
    }
}