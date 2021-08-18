using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace ListasSarlaft.Classes
{
    public class clsBLLFrecuenciavsEventos
    {
        private cDataBase cDataBase = new cDataBase();
        /// <summary>
        /// Metodo para insertar el registro de Frecuencia de los eventos
        /// </summary>
        /// <param name="strErrMsg">Mensaje de error</param>
        /// <returns>Retorna si el proceso fue exitoso o no</returns>
        public bool mtdInsertarFrecuenciavsEventos(clsDTOFrecuenciavsEventos frequencyEvents, ref string strErrMsg)
        {
            bool booResult = false;
            clsDALFrecuenciavsEventos cDALFrequencyEvents = new clsDALFrecuenciavsEventos();

            booResult = cDALFrequencyEvents.mtdInsertarFrecuenciavsEventos(frequencyEvents, ref strErrMsg);

            return booResult;
        }
        /// <summary>
        /// Metodo para validar antes de insertar el registro de Frecuencia de los eventos
        /// </summary>
        /// <param name="strErrMsg">Mensaje de error</param>
        /// <returns>Retorna si el proceso fue exitoso o no</returns>
        public bool mtdValidaInsertarFrecuenciavsEventos(clsDTOFrecuenciavsEventos frequencyEvents, ref string strErrMsg)
        {
            bool booResult = false;
            clsDALFrecuenciavsEventos cDALFrequencyEvents = new clsDALFrecuenciavsEventos();

            booResult = cDALFrequencyEvents.mtdValidarInsertarFrecuenciavsEventos(frequencyEvents, ref strErrMsg);

            return booResult;
        }
        /// <summary>
        /// Metodo para consultar y visualizar las Frecuencia de los eventos
        /// </summary>
        /// <param name="strErrMsg">Mensaje de error</param>
        /// <returns>Retorna si el proceso fue exitoso o no</returns>
        public List<clsDTOFrecuenciavsEventos> mtdConsultarFrecuenciavsEventos(ref List<clsDTOFrecuenciavsEventos> lstFrequencyEvents, ref string strErrMsg)
        {
            #region Vars
            bool booResult = false;
            DataTable dtInfo = new DataTable();
            clsDALFrecuenciavsEventos cDtFrequencyEvents = new clsDALFrecuenciavsEventos();
            #endregion Vars

            booResult = cDtFrequencyEvents.mtdConsultarFrecuenciavsEventos(ref dtInfo, ref strErrMsg);

            if (booResult)
            {
                if (dtInfo != null)
                {
                    if (dtInfo.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtInfo.Rows)
                        {
                            clsDTOFrecuenciavsEventos objFrequencyEvents = new clsDTOFrecuenciavsEventos();
                            objFrequencyEvents.intIdFrecuenciaEventos = Convert.ToInt32(dr["IdFrecuenciaEventos"].ToString().Trim());
                            objFrequencyEvents.intEventosMaximos = Convert.ToInt32(dr["EventosMaximos"].ToString().Trim());
                            objFrequencyEvents.intCodigoFrecuencia = Convert.ToInt32(dr["CodigoFrecuencia"].ToString().Trim());
                            objFrequencyEvents.strNombreFrecuencia = dr["NombreFrecuencia"].ToString().Trim();
                            objFrequencyEvents.dtFechaRegistro = Convert.ToDateTime(dr["FechaRegistro"].ToString().Trim());
                            objFrequencyEvents.intIdUsuario = Convert.ToInt32(dr["UsuarioCreacion"].ToString().Trim());
                            objFrequencyEvents.strUsuario = dr["Usuario"].ToString().Trim();

                            lstFrequencyEvents.Add(objFrequencyEvents);
                        }
                    }
                    else
                    {
                        lstFrequencyEvents = null;
                    }
                }
                else
                {
                    lstFrequencyEvents = null;
                }
            }

            return lstFrequencyEvents;
        }
        /// <summary>
        /// Metodo que permite la actualizacion de las frecuencias de eventos
        /// </summary>
        /// <param name="objFrequencyEvents">objeto con la informacion de las frecuencias de eventos</param>
        /// <param name="strErrMsg">Mensaje de error</param>
        /// <returns>Retorna si el proceso fue exitoso o no</returns>
        public bool mtdUpdateFrecuenciavsEventos(ref clsDTOFrecuenciavsEventos objFrequencyEvents, ref string strErrMsg)
        {
            bool booResult = false;
            clsDALFrecuenciavsEventos cDtFrequencyEvents = new clsDALFrecuenciavsEventos();

            booResult = cDtFrequencyEvents.mtdUpdateFrecuenciavsEventos(ref objFrequencyEvents, ref strErrMsg);

            return booResult;
        }

        //Yoendy - Pestaña Frecuencia vs Eventos 
        public List<clsDTOFrecuenciavsEventos> consultaFrecuenciaxRiesgo(ref List<clsDTOFrecuenciavsEventos> lstFrequencyEvents, ref int idRiesgo, ref string strErrMsg)
        {
            #region Vars
            bool todasFreqEventos = false;
            bool freqEvenEncontrado = false;
            bool freqTablaNueva = false;
            DataTable dtTablaBase = new DataTable();
            DataTable dtInfoGeneral = new DataTable();
            DataTable dtTablaNueva = new DataTable();
            clsDALFrecuenciavsEventos cDtFrequencyEvents = new clsDALFrecuenciavsEventos();

            todasFreqEventos = cDtFrequencyEvents.mtdConsultarFrecuenciavsEventos(ref dtInfoGeneral, ref strErrMsg); // Todos
            freqEvenEncontrado = cDtFrequencyEvents.mtdConsultarFrecuenciavsRiesgos(ref dtTablaBase, ref idRiesgo, ref strErrMsg); // Base
            string codFrecuencia = dtTablaBase.Rows[0]["codFrecuencia"].ToString(); // Obligo a traer el codFrecuencia - es obligatorio, siempre lo trae
            freqTablaNueva = cDtFrequencyEvents.mtdConsultarFrecuenciaRelacionada(ref dtTablaNueva, ref idRiesgo, ref codFrecuencia, ref strErrMsg); // Personalizada

            int pos = 0, cont = 0, posB = 0, contB = 0, contC = 0;
            string[] arr = new string[dtInfoGeneral.Rows.Count];
            string[] encontrado = new string[dtInfoGeneral.Rows.Count];
            string codigoFrecuencia = string.Empty;
            #endregion Vars                        

            for (int i = 0; i < dtTablaBase.Rows.Count; i++)
            {
                arr[i] = dtTablaBase.Rows[i]["NombreFrecuencia"].ToString();
                contC++;
            }

            if (freqEvenEncontrado == true)
            {
                DataTable dtRelacion = dtTablaBase;
                string valor = string.Empty, encontro = string.Empty;
                //Comparar             

                for (int i = 0; i < dtInfoGeneral.Rows.Count; i++)
                {
                    string nombreFrecuencia = dtInfoGeneral.Rows[i]["NombreFrecuencia"].ToString();
                    if (nombreFrecuencia == arr[pos])
                    {
                        encontrado[cont] = "<span style= 'font-weight: bold; color: red;'>" + nombreFrecuencia + "</span>";
                        pos++;
                        cont++;
                    }
                    else
                    {
                        encontrado[cont] = nombreFrecuencia;
                        cont++;
                    }
                }
            }
            else
            {
                for (int i = 0; i < dtInfoGeneral.Rows.Count; i++)
                {
                    encontrado[i] = dtInfoGeneral.Rows[i]["NombreFrecuencia"].ToString();
                }
            }
            if (todasFreqEventos)
            {
                if (dtInfoGeneral != null)
                {
                    if (dtInfoGeneral.Rows.Count > 0)
                    {
                        int i = 0;
                        foreach (DataRow dr in dtInfoGeneral.Rows)
                        {
                            clsDTOFrecuenciavsEventos objFrequencyEvents = new clsDTOFrecuenciavsEventos();
                            string nombreFrecuencia = dtInfoGeneral.Rows[i]["NombreFrecuencia"].ToString();
                            if (nombreFrecuencia == arr[posB])
                            {
                                if (dtTablaNueva.Rows.Count > 0)
                                {                                   
                                    objFrequencyEvents.intEventosMaximos = Convert.ToInt32(dtTablaNueva.Rows[0]["EventosMaximos"]);                                                                       
                                    posB++;
                                    contB++;
                                }
                                else
                                {
                                    objFrequencyEvents.intEventosMaximos = Convert.ToInt32(dr["EventosMaximos"].ToString().Trim());
                                    contB++;
                                }
                            }
                            else
                            {
                                objFrequencyEvents.intEventosMaximos = Convert.ToInt32(dr["EventosMaximos"].ToString().Trim());
                                contB++;
                            }

                            objFrequencyEvents.intIdFrecuenciaEventos = Convert.ToInt32(dr["IdFrecuenciaEventos"].ToString().Trim());
                            objFrequencyEvents.intCodigoFrecuencia = Convert.ToInt32(dr["CodigoFrecuencia"].ToString().Trim());
                            objFrequencyEvents.strNombreFrecuencia = encontrado[i];
                            objFrequencyEvents.dtFechaRegistro = Convert.ToDateTime(dr["FechaRegistro"].ToString().Trim());
                            objFrequencyEvents.intIdUsuario = Convert.ToInt32(dr["UsuarioCreacion"].ToString().Trim());
                            objFrequencyEvents.strUsuario = dr["Usuario"].ToString().Trim();

                            lstFrequencyEvents.Add(objFrequencyEvents);
                            i++;
                        }
                    }
                    else
                    {
                        lstFrequencyEvents = null;
                    }
                }
                else
                {
                    lstFrequencyEvents = null;
                }
            }
            return lstFrequencyEvents;
        }

        //Yoendy
        public bool verificaFreq(int idRiesgo, string NombreFrecuencia)
        {
            bool resultado = false;
            cEvento cEvento = new cEvento();
            DataTable rt = cEvento.mtdTotalRiesgosFrecuencia(idRiesgo, NombreFrecuencia);

            int existe = Convert.ToInt32(rt.Rows[0]["ExisteFreq"]);

            if (existe > 0)
            {
                resultado = true;
            }
            return resultado;
        }

        public int ActualizaFreqMax(int idRiesgo, int codFrecuencia, int eventosMaximos, string nombreFrecuencia, int idUsuario)
        {
            int resultado = 0;
            cEvento cEvento = new cEvento();

            List<SqlParameter> parametros = new List<SqlParameter>()
                {
                    new SqlParameter() { ParameterName = "@idRiesgo", SqlDbType = SqlDbType.Int, Value = idRiesgo },
                    new SqlParameter() { ParameterName = "@codFrecuencia", SqlDbType = SqlDbType.Int, Value = codFrecuencia },
                    new SqlParameter() { ParameterName = "@eventosMaximos", SqlDbType = SqlDbType.Int, Value = eventosMaximos },
                    new SqlParameter() { ParameterName = "@nombreFrecuencia", SqlDbType = SqlDbType.VarChar, Value = nombreFrecuencia },
                    new SqlParameter() { ParameterName = "@idUsuario", SqlDbType = SqlDbType.Int, Value = idUsuario },
                    new SqlParameter() { ParameterName = "@Resultado", SqlDbType = SqlDbType.Int, Direction = ParameterDirection.Output },
                };

            resultado = cDataBase.EjecutarSPParametrosReturnInteger("[dbo].[pa_RegistrarFreqEvento]", parametros);

            return resultado;           
        }


    } // Fin nombre de espacios - Yoendy
}