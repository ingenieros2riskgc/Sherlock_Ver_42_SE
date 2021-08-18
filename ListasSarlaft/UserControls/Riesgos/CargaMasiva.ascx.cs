using ClosedXML.Excel;
using clsDatos;
using Excel;
using ListasSarlaft.Classes;
using ListasSarlaft.Classes.DTO.Riesgos.CargaMasiva;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Web;
using System.Web.Configuration;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ListasSarlaft.UserControls.Riesgos
{
    public partial class CargaMasiva : System.Web.UI.UserControl
    {
        private cError cError = new cError();
        private int pagIndex;
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
        private static int LastInsertIdCE;


        private static CultureInfo culturaFR = CultureInfo.CreateSpecificCulture("fr-FR");
        public static CultureInfo CulturaFR => culturaFR;

        #region Properties
        private DataTable infoGridComentarioPlanEvaluacion;
        private DataTable InfoGridComentarioPlanEvaluacion
        {
            get
            {
                infoGridComentarioPlanEvaluacion = (DataTable)ViewState["infoGridComentarioPlanEvaluacion"];
                return infoGridComentarioPlanEvaluacion;
            }
            set
            {
                infoGridComentarioPlanEvaluacion = value;
                ViewState["infoGridComentarioPlanEvaluacion"] = infoGridComentarioPlanEvaluacion;
            }
        }

        private int rowGridComentarioPlanEvaluacion;
        private int RowGridComentarioPlanEvaluacion
        {
            get
            {
                rowGridComentarioPlanEvaluacion = (int)ViewState["rowGridComentarioPlanEvaluacion"];
                return rowGridComentarioPlanEvaluacion;
            }
            set
            {
                rowGridComentarioPlanEvaluacion = value;
                ViewState["rowGridComentarioPlanEvaluacion"] = rowGridComentarioPlanEvaluacion;
            }
        }

        private int rowGridArchivoPlanEvaluacion;
        private int RowGridArchivoPlanEvaluacion
        {
            get
            {
                rowGridArchivoPlanEvaluacion = (int)ViewState["rowGridArchivoPlanEvaluacion"];
                return rowGridArchivoPlanEvaluacion;
            }
            set
            {
                rowGridArchivoPlanEvaluacion = value;
                ViewState["rowGridArchivoPlanEvaluacion"] = rowGridArchivoPlanEvaluacion;
            }
        }

        private DataTable infoGridArchivoPlanEvaluacion;
        private DataTable InfoGridArchivoPlanEvaluacion
        {
            get
            {
                infoGridArchivoPlanEvaluacion = (DataTable)ViewState["infoGridArchivoPlanEvaluacion"];
                return infoGridArchivoPlanEvaluacion;
            }
            set
            {
                infoGridArchivoPlanEvaluacion = value;
                ViewState["infoGridArchivoPlanEvaluacion"] = infoGridArchivoPlanEvaluacion;
            }
        }

        private int rowGridPlanEvaluacion;
        private int RowGridPlanEvaluacion
        {
            get
            {
                rowGridPlanEvaluacion = (int)ViewState["rowGridPlanEvaluacion"];
                return rowGridPlanEvaluacion;
            }
            set
            {
                rowGridPlanEvaluacion = value;
                ViewState["rowGridPlanEvaluacion"] = rowGridPlanEvaluacion;
            }
        }

        private DataTable infoGridPlanEvaluacion;
        private DataTable InfoGridPlanEvaluacion
        {
            get
            {
                infoGridPlanEvaluacion = (DataTable)ViewState["infoGridPlanEvaluacion"];
                return infoGridPlanEvaluacion;
            }
            set
            {
                infoGridPlanEvaluacion = value;
                ViewState["infoGridPlanEvaluacion"] = infoGridPlanEvaluacion;
            }
        }

        private DataTable infoGridComentarioControl;
        private DataTable InfoGridComentarioControl
        {
            get
            {
                infoGridComentarioControl = (DataTable)ViewState["infoGridComentarioControl"];
                return infoGridComentarioControl;
            }
            set
            {
                infoGridComentarioControl = value;
                ViewState["infoGridComentarioControl"] = infoGridComentarioControl;
            }
        }

        private DataTable infoPorcentajeCalificarControl;
        private DataTable InfoPorcentajeCalificarControl
        {
            get
            {
                infoPorcentajeCalificarControl = (DataTable)ViewState["infoPorcentajeCalificarControl"];
                return infoPorcentajeCalificarControl;
            }
            set
            {
                infoPorcentajeCalificarControl = value;
                ViewState["infoPorcentajeCalificarControl"] = infoPorcentajeCalificarControl;
            }
        }

        private DataTable infoCalificacionControl;
        private DataTable InfoCalificacionControl
        {
            get
            {
                infoCalificacionControl = (DataTable)ViewState["infoCalificacionControl"];
                return infoCalificacionControl;
            }
            set
            {
                infoCalificacionControl = value;
                ViewState["infoCalificacionControl"] = infoCalificacionControl;
            }
        }

        private DataTable infoGridControles;
        private DataTable InfoGridControles
        {
            get
            {
                infoGridControles = (DataTable)ViewState["infoGridControles"];
                return infoGridControles;
            }
            set
            {
                infoGridControles = value;
                ViewState["infoGridControles"] = infoGridControles;
            }
        }

        private int rowGridControles;
        private int RowGridControles
        {
            get
            {
                rowGridControles = (int)ViewState["rowGridControles"];
                return rowGridControles;
            }
            set
            {
                rowGridControles = value;
                ViewState["rowGridControles"] = rowGridControles;
            }
        }

        private String idCalificacionControl;
        private String IdCalificacionControl
        {
            get
            {
                idCalificacionControl = (String)ViewState["idCalificacionControl"];
                return idCalificacionControl;
            }
            set
            {
                idCalificacionControl = value;
                ViewState["idCalificacionControl"] = idCalificacionControl;
            }
        }

        private DataTable infoGridArchivoControl;
        private DataTable InfoGridArchivoControl
        {
            get
            {
                infoGridArchivoControl = (DataTable)ViewState["infoGridArchivoControl"];
                return infoGridArchivoControl;
            }
            set
            {
                infoGridArchivoControl = value;
                ViewState["infoGridArchivoControl"] = infoGridArchivoControl;
            }
        }

        private int rowGridArchivoControl;
        private int RowGridArchivoControl
        {
            get
            {
                rowGridArchivoControl = (int)ViewState["rowGridArchivoControl"];
                return rowGridArchivoControl;
            }
            set
            {
                rowGridArchivoControl = value;
                ViewState["rowGridArchivoControl"] = rowGridArchivoControl;
            }
        }

        private int rowGridComentarioControl;
        private int RowGridComentarioControl
        {
            get
            {
                rowGridComentarioControl = (int)ViewState["rowGridComentarioControl"];
                return rowGridComentarioControl;
            }
            set
            {
                rowGridComentarioControl = value;
                ViewState["rowGridComentarioControl"] = rowGridComentarioControl;
            }
        }

        private DataTable infoIntervalos;
        private DataTable InfoIntervalos
        {
            get
            {
                infoIntervalos = (DataTable)ViewState["infoIntervalos"];
                return infoIntervalos;
            }
            set
            {
                infoIntervalos = value;
                ViewState["infoIntervalos"] = infoIntervalos;
            }
        }

        private int pagIndexInfoGridControles;
        private int PagIndexInfoGridControles
        {
            get
            {
                pagIndexInfoGridControles = (int)ViewState["pagIndexInfoGridControles"];
                return pagIndexInfoGridControles;
            }
            set
            {
                pagIndexInfoGridControles = value;
                ViewState["pagIndexInfoGridControles"] = pagIndexInfoGridControles;
            }
        }
        #endregion

        private cControl cControl = new cControl();
        private cRegistroOperacion cRegistroOperacion = new cRegistroOperacion();
        private string IdFormulario = "5021";
        private cCuenta cCuenta = new cCuenta();
        protected void Page_Load(object sender, EventArgs e)
        {
            Page.Form.Attributes.Add("enctype", "multipart/form-data");
            ScriptManager scriptManager = ScriptManager.GetCurrent(this.Page);
            scriptManager.RegisterPostBackControl(this.ImButtonExcelExportDataRiesgo);
            scriptManager.RegisterPostBackControl(this.ImButtonExcelExportPlantillaRiesgo);
            if (!Page.IsPostBack)
            {
                Session["IdControl"] = cControl.SeleccionarUltimoControl();
            }
            scriptManager.RegisterPostBackControl(this.Bload);
            if (cCuenta.permisosConsulta(IdFormulario) == "False")
            {
                Response.Redirect("~/Formularios/Sarlaft/Admin/HomeAdmin.aspx?Denegar=1");
            }
        }

        protected void DDLopciones_SelectedIndexChanged(object sender, EventArgs e)
        {

            TrButtonsExportRiesgos.Visible = true;
            TrLoadFile.Visible = true;
        }

        protected void ImButtonExcelExportDataRiesgo_Click(object sender, ImageClickEventArgs e)
        {
            if (DDLopciones.SelectedValue == "1")
            {
                exportExcelRiesgo(Response, "DatosParametricosRiesgos_" + System.DateTime.Now.ToString("yyyy-MM-dd"));
            }

            if (DDLopciones.SelectedValue == "2")
            {
                exportExcelControles(Response, "DatosParametricosControles_" + System.DateTime.Now.ToString("yyyy-MM-dd"));
            }

            if (DDLopciones.SelectedValue == "3")
            {
                exportExcelEventosCreacion(Response, "DatosParametricosEventosCreacion_" + System.DateTime.Now.ToString("yyyy-MM-dd"));
            }
            if (DDLopciones.SelectedValue == "5")
            {
                exportExcelEventosDatosComplementarios(Response, "DatosParametricosEventosDatosComplementarios_" + System.DateTime.Now.ToString("yyyy-MM-dd"));
            }
            if (DDLopciones.SelectedValue == "6")
            {
                exportExcelEventosConDatosComplementarios(Response, "DatosParametricosEventosConDatosComplementarios_" + System.DateTime.Now.ToString("yyyy-MM-dd"));
            }
            if (DDLopciones.SelectedValue == "4")
            {
                exportExcelRiesgosvsControles(Response, "DatosParametricosRiesgosvsControles_" + System.DateTime.Now.ToString("yyyy-MM-dd"));
            }
        }
        private void armarIntervalos()
        {
            DataTable dtInfo = new DataTable(), dtIntervalos = new DataTable();
            double maximo = 0, minimo = 0, intervalo = 0, delta = 0;

            try
            {
                dtInfo = cControl.valorMaxMinIntervalo();

                if (dtInfo != null)
                {
                    if (dtInfo.Rows.Count > 0)
                    {
                        minimo = Convert.ToDouble(dtInfo.Rows[0]["Minimo"].ToString().Trim());
                        intervalo = (Convert.ToDouble(dtInfo.Rows[0]["Maximo"].ToString().Trim())) - (Convert.ToDouble(dtInfo.Rows[0]["Minimo"].ToString().Trim()));
                        delta = intervalo / 4;

                        dtIntervalos.Columns.Add("limiteInferior", typeof(string));
                        dtIntervalos.Columns.Add("limiteSuperior", typeof(string));
                        dtIntervalos.Columns.Add("IdCalificacionControl", typeof(string));
                        for (int rows = 4; rows > 0; rows--)
                        {
                            maximo = minimo + delta;
                            dtIntervalos.Rows.Add(new Object[] { minimo.ToString(), maximo.ToString(), rows.ToString() });
                            minimo = maximo;
                        }
                        dtIntervalos.Rows[0]["limiteInferior"] = "0";
                        dtIntervalos.Rows[3]["limiteSuperior"] = "100";
                        InfoIntervalos = dtIntervalos;
                    }
                    else
                    {
                        Mensaje("No hay información en las tablas maestras de parametrización. ");
                    }
                }
                else
                {
                    Mensaje("No hay información en las tablas maestras de parametrización. ");
                }
            }
            catch (Exception ex)
            {
                Mensaje("Error al armar intervalos. " + ex.Message);
            }
        }
        private void loadInfoCalificacionControl()
        {
            InfoCalificacionControl = cControl.loadInfoCalificacionControl();
        }
        private void loadInfoPorcentajeCalificarControl()
        {
            InfoPorcentajeCalificarControl = cControl.loadInfoPorcentajeCalificarControl();
        }
        private string calcularCalificacionControl(string ClaseControl, string TipoControl, string experiencia, string document, string respons)
        {
            double calificacionControl = 0;
            string calificacion = string.Empty;
            double claseControl = 0;
            if (ClaseControl != "0")
            {
                claseControl = cControl.valorClaseControl(ClaseControl);
            }

            double tipoControl = 0;
            if (TipoControl != "0")
            {
                tipoControl = cControl.valorTipoControl(TipoControl);
            }

            double responsableExperiencia = 0;
            if (experiencia != "0")
            {
                responsableExperiencia = cControl.valorResponsableExperiencia(experiencia);
            }

            double documentacion = 0;
            if (document != "0")
            {
                documentacion = cControl.valorDocumentacion(document);
            }

            double responsabilidad = 0;
            if (respons != "0")
            {
                responsabilidad = cControl.valorResponsabilidad(respons);
            }

            loadInfoPorcentajeCalificarControl();
            armarIntervalos();
            loadInfoCalificacionControl();

            calificacionControl = claseControl * Convert.ToDouble(InfoPorcentajeCalificarControl.Rows[0]["ValorPorcentajeCalificarControl"].ToString().Trim());
            calificacionControl += tipoControl * Convert.ToDouble(InfoPorcentajeCalificarControl.Rows[1]["ValorPorcentajeCalificarControl"].ToString().Trim());
            calificacionControl += responsableExperiencia * Convert.ToDouble(InfoPorcentajeCalificarControl.Rows[2]["ValorPorcentajeCalificarControl"].ToString().Trim());
            calificacionControl += documentacion * Convert.ToDouble(InfoPorcentajeCalificarControl.Rows[3]["ValorPorcentajeCalificarControl"].ToString().Trim());
            calificacionControl += responsabilidad * Convert.ToDouble(InfoPorcentajeCalificarControl.Rows[4]["ValorPorcentajeCalificarControl"].ToString().Trim());
            calificacionControl = (calificacionControl / 100);
            for (int i = 0; i < InfoIntervalos.Rows.Count; i++)
            {
                if (calificacionControl > Convert.ToDouble(InfoIntervalos.Rows[i]["limiteInferior"].ToString().Trim()) && calificacionControl <= Convert.ToDouble(InfoIntervalos.Rows[i]["limiteSuperior"].ToString().Trim()))
                {
                    IdCalificacionControl = InfoIntervalos.Rows[i]["IdCalificacionControl"].ToString().Trim();
                    for (int j = 0; j < InfoCalificacionControl.Rows.Count; j++)
                    {
                        if (IdCalificacionControl == InfoCalificacionControl.Rows[j]["IdCalificacionControl"].ToString().Trim())
                        {
                            calificacion = IdCalificacionControl;
                            //Panel1.BackColor = System.Drawing.Color.FromName(InfoCalificacionControl.Rows[j]["Color"].ToString().Trim());
                            break;
                        }
                    }
                    break;
                }
            }
            return calificacion;
        }
        protected void exportExcelRiesgo(HttpResponse Response, string filename)
        {

            cCargaMasivaRCE carga = new cCargaMasivaRCE();
            DataTable DataUbicacion = carga.DataUbicacionRiesgo();
            DataUbicacion.TableName = "Ubicacion 1-16";
            DataTable DataProcesos = carga.DataProcesosRiesgo();
            DataProcesos.TableName = "Procesos 2-16";
            DataTable DataClasificacion = carga.DataClasificacionRiesgo();
            DataClasificacion.TableName = "Clasificacion Riesgos 3-16";
            DataTable DataTipoEvento = carga.DataTipoEvento();
            DataTipoEvento.TableName = "Tipo Evento 4-16";
            DataTable DataRiesgoOperativo = carga.DataRiesgoOperativo();
            DataRiesgoOperativo.TableName = "Riesgo Operativo 5-16";
            DataTable DataRiesgoAsociado = carga.DataRiesgoAsociadoOperativo();
            DataRiesgoAsociado.TableName = "Riesgo Asociado 6-16";
            DataTable DataRiesgoLA = carga.DataRiesgoAsociadoLA();
            DataRiesgoLA.TableName = "Riesgo LA 7-16";
            DataTable DataRiesgoLAFT = carga.DataRiesgoAsociadoLAFT();
            DataRiesgoLAFT.TableName = "Riesgo LAFT 8-16";
            DataTable DataCausas = carga.DataCausas();
            DataCausas.TableName = "Causas 9-16";
            DataTable DataConsecuencias = carga.DataConsecuencia();
            DataConsecuencias.TableName = "Consecuencias 10-16";
            DataTable DataFrecuencia = carga.DataFrecuencia();
            DataFrecuencia.TableName = "Frecuencia-Cualitativa 11-16";
            DataTable DataTratamiento = carga.DataTratamiento();
            DataTratamiento.TableName = "Tratamiento 12-16";
            DataTable DataResponsable = carga.DataResponsable();
            DataResponsable.TableName = "Responsable 13-16";
            DataTable DataImpacto = carga.DataImpacto();
            DataImpacto.TableName = "Impacto 14-16";
            DataTable DataEstado = carga.DataEstado();
            DataEstado.TableName = "Estado 15-16";

            DataTable DataPerdida = carga.DataPerdida();
            DataPerdida.TableName = "Tipo Perdida 16-16";

            System.Data.DataSet ds = new System.Data.DataSet();
            ds.Tables.Add(DataUbicacion);
            ds.Tables.Add(DataProcesos);
            ds.Tables.Add(DataClasificacion);
            ds.Tables.Add(DataTipoEvento);
            ds.Tables.Add(DataRiesgoOperativo);
            ds.Tables.Add(DataRiesgoAsociado);
            ds.Tables.Add(DataRiesgoLA);
            ds.Tables.Add(DataRiesgoLAFT);
            ds.Tables.Add(DataCausas);
            ds.Tables.Add(DataConsecuencias);
            ds.Tables.Add(DataFrecuencia);
            ds.Tables.Add(DataTratamiento);
            ds.Tables.Add(DataResponsable);
            ds.Tables.Add(DataImpacto);
            ds.Tables.Add(DataEstado);
            ds.Tables.Add(DataPerdida);
            // Create the workbook
            XLWorkbook workbook = new XLWorkbook();
            //workbook.Worksheets.Add("Sample").Cell(1, 1).SetValue("Hello World");
            workbook.Worksheets.Add(ds);
            // Prepare the response
            HttpResponse httpResponse = Response;
            httpResponse.Clear();
            httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            httpResponse.AddHeader("content-disposition", "attachment;filename=\"" + filename + ".xlsx\"");

            // Flush the workbook to the Response.OutputStream
            using (MemoryStream memoryStream = new MemoryStream())
            {
                workbook.SaveAs(memoryStream);
                memoryStream.WriteTo(httpResponse.OutputStream);
                memoryStream.Close();
            }

            httpResponse.End();
        }
        protected void exportExcelEventosCreacion(HttpResponse Response, string filename)
        {

            cCargaMasivaRCE carga = new cCargaMasivaRCE();
            DataTable DataEmpresa = carga.DataEmpresa();
            DataEmpresa.TableName = "Empresa 1-8";
            DataTable DataUbicacion = carga.DataUbicacionRiesgo();
            DataUbicacion.TableName = "Ubicacion 2-8";
            DataTable DataServicio = carga.DataServicio();
            DataServicio.TableName = "Servicios 3-8";
            DataTable DataCanal = carga.DataCanal();
            DataCanal.TableName = "Canal 4-8";
            DataTable DataGenerador = carga.DataGenerador();
            DataGenerador.TableName = "Generador 5-8";
            DataTable DataResponsable = carga.DataResponsable();
            DataResponsable.TableName = "Responsable 6-8";
            clsDtImpactoCualitativo clsDt = new clsDtImpactoCualitativo();
            DataTable DataImpacto = clsDt.loadNombreImpCual();
            DataImpacto.TableName = "Impacto Cualitativo 7-8";

            DataTable DataPerdida = carga.DataPerdida();
            DataPerdida.TableName = "Tipo perdida 8-8";

            /*DataTable DataProcesos = carga.DataProcesosRiesgo();
            DataProcesos.TableName = "Procesos 7-16";
            DataTable DataClaseRiesgo = carga.DataClaseRiesgo();
            DataClaseRiesgo.TableName = "Clase Riesgo 8-16";
            DataTable DataSubClaseRiesgo = carga.DataSubClaseRiesgo();
            DataSubClaseRiesgo.TableName = "SubClase Riesgo 9-16";
            DataTable DataPerdida = carga.DataPerdida();
            DataPerdida.TableName = "Pérdidas 10-16";
            DataTable DataLineaOperativa = carga.DataLineaOperativa();
            DataLineaOperativa.TableName = "Línea Operativa 11-16";
            DataTable DataSubLineaOperativa = carga.DataSubLineaOperativa();
            DataSubLineaOperativa.TableName = "SubLínea Operativa 12-16";
            DataTable DataContinuidad = carga.DataContinuidad();
            DataContinuidad.TableName = "Continuidad 13-16";
            DataTable DataEstado = carga.DataEstado();
            DataEstado.TableName = "Estados 14-16";
            DataTable DataMoneda = carga.DataMoneda();
            DataMoneda.TableName = "Moneda 15-16";
            DataTable DataRecuperacion = carga.DataRecuperacion();
            DataRecuperacion.TableName = "Recuperación 16-16";*/



            System.Data.DataSet ds = new System.Data.DataSet();
            ds.Tables.Add(DataEmpresa);
            ds.Tables.Add(DataUbicacion);
            ds.Tables.Add(DataServicio);
            ds.Tables.Add(DataCanal);
            ds.Tables.Add(DataGenerador);
            ds.Tables.Add(DataResponsable);
            ds.Tables.Add(DataImpacto);
            ds.Tables.Add(DataPerdida);
            /*ds.Tables.Add(DataProcesos);
            ds.Tables.Add(DataClaseRiesgo);
            ds.Tables.Add(DataSubClaseRiesgo);
            ds.Tables.Add(DataPerdida);
            ds.Tables.Add(DataLineaOperativa);
            ds.Tables.Add(DataSubLineaOperativa);
            ds.Tables.Add(DataContinuidad);
            ds.Tables.Add(DataEstado);
            ds.Tables.Add(DataMoneda);
            ds.Tables.Add(DataRecuperacion);*/

            XLWorkbook workbook = new XLWorkbook();
            workbook.Worksheets.Add(ds);
            HttpResponse httpResponse = Response;
            httpResponse.Clear();
            httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            httpResponse.AddHeader("content-disposition", "attachment;filename=\"" + filename + ".xlsx\"");

            // Flush the workbook to the Response.OutputStream
            using (MemoryStream memoryStream = new MemoryStream())
            {
                workbook.SaveAs(memoryStream);
                memoryStream.WriteTo(httpResponse.OutputStream);
                memoryStream.Close();
            }

            httpResponse.End();
        }
        protected void exportExcelEventosConDatosComplementarios(HttpResponse Response, string filename)
        {

            cCargaMasivaRCE carga = new cCargaMasivaRCE();
            DataTable DataEmpresa = carga.DataEmpresa();
            DataEmpresa.TableName = "Empresa 1-17";
            DataTable DataUbicacion = carga.DataUbicacionRiesgo();
            DataUbicacion.TableName = "Ubicacion 2-17";
            DataTable DataServicio = carga.DataServicio();
            DataServicio.TableName = "Servicios 3-17";
            DataTable DataCanal = carga.DataCanal();
            DataCanal.TableName = "Canal 4-17";
            DataTable DataGenerador = carga.DataGenerador();
            DataGenerador.TableName = "Generador 5-17";
            DataTable DataResponsable = carga.DataResponsable();
            DataResponsable.TableName = "Responsable 6-17";
            DataTable DataProcesos = carga.DataProcesosRiesgo();
            DataProcesos.TableName = "Procesos 7-17";
            DataTable DataClaseRiesgo = carga.DataClaseRiesgo();
            DataClaseRiesgo.TableName = "Clase Riesgo 8-17";
            DataTable DataSubClaseRiesgo = carga.DataSubClaseRiesgo();
            DataSubClaseRiesgo.TableName = "SubClase Riesgo 9-17";
            DataTable DataPerdida = carga.DataPerdida();
            DataPerdida.TableName = "Pérdidas 10-17";
            DataTable DataLineaOperativa = carga.DataLineaOperativa();
            DataLineaOperativa.TableName = "Línea Operativa 11-17";
            DataTable DataSubLineaOperativa = carga.DataSubLineaOperativa();
            DataSubLineaOperativa.TableName = "SubLínea Operativa 12-17";
            DataTable DataContinuidad = carga.DataContinuidad();
            DataContinuidad.TableName = "Continuidad 13-17";
            DataTable DataEstado = carga.DataEstadoEventos();
            DataEstado.TableName = "Estados 14-17";
            DataTable DataMoneda = carga.DataMoneda();
            DataMoneda.TableName = "Moneda 15-17";
            DataTable DataRecuperacion = carga.DataRecuperacion();
            DataRecuperacion.TableName = "Recuperación 16-17";
            clsDtImpactoCualitativo clsDt = new clsDtImpactoCualitativo();
            DataTable DataImpacto = clsDt.loadNombreImpCual();
            DataImpacto.TableName = "Impacto Cualitativo 17-17";

            System.Data.DataSet ds = new System.Data.DataSet();
            ds.Tables.Add(DataEmpresa);
            ds.Tables.Add(DataUbicacion);
            ds.Tables.Add(DataServicio);
            ds.Tables.Add(DataCanal);
            ds.Tables.Add(DataGenerador);
            ds.Tables.Add(DataResponsable);
            ds.Tables.Add(DataProcesos);
            ds.Tables.Add(DataClaseRiesgo);
            ds.Tables.Add(DataSubClaseRiesgo);
            ds.Tables.Add(DataPerdida);
            ds.Tables.Add(DataLineaOperativa);
            ds.Tables.Add(DataSubLineaOperativa);
            ds.Tables.Add(DataContinuidad);
            ds.Tables.Add(DataEstado);
            ds.Tables.Add(DataMoneda);
            ds.Tables.Add(DataRecuperacion);
            ds.Tables.Add(DataImpacto);

            XLWorkbook workbook = new XLWorkbook();
            workbook.Worksheets.Add(ds);
            HttpResponse httpResponse = Response;
            httpResponse.Clear();
            httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            httpResponse.AddHeader("content-disposition", "attachment;filename=\"" + filename + ".xlsx\"");

            // Flush the workbook to the Response.OutputStream
            using (MemoryStream memoryStream = new MemoryStream())
            {
                workbook.SaveAs(memoryStream);
                memoryStream.WriteTo(httpResponse.OutputStream);
                memoryStream.Close();
            }

            httpResponse.End();
        }
        protected void exportExcelEventosDatosComplementarios(HttpResponse Response, string filename)
        {

            cCargaMasivaRCE carga = new cCargaMasivaRCE();
            /*DataTable DataEmpresa = carga.DataEmpresa();
            DataEmpresa.TableName = "Empresa 1-16";
            DataTable DataUbicacion = carga.DataUbicacionRiesgo();
            DataUbicacion.TableName = "Ubicacion 2-16";
            DataTable DataServicio = carga.DataServicio();
            DataServicio.TableName = "Servicios 3-16";
            DataTable DataCanal = carga.DataCanal();
            DataCanal.TableName = "Canal 4-16";
            DataTable DataGenerador = carga.DataGenerador();
            DataGenerador.TableName = "Generador 5-16";
            DataTable DataResponsable = carga.DataResponsable();
            DataResponsable.TableName = "Responsable 6-16";*/
            DataTable DataEventos = carga.DataEventos();
            DataEventos.TableName = "Eventos 1-12";
            DataTable DataProcesos = carga.DataProcesosRiesgo();
            DataProcesos.TableName = "Procesos 2-12";
            DataTable DataClaseRiesgo = carga.DataClaseRiesgo();
            DataClaseRiesgo.TableName = "Clase Riesgo 3-12";
            DataTable DataSubClaseRiesgo = carga.DataSubClaseRiesgo();
            DataSubClaseRiesgo.TableName = "SubClase Riesgo 4-12";
            DataTable DataPerdida = carga.DataPerdida();
            DataPerdida.TableName = "Pérdidas 5-12";
            DataTable DataLineaOperativa = carga.DataLineaOperativa();
            DataLineaOperativa.TableName = "Línea Operativa 6-12";
            DataTable DataSubLineaOperativa = carga.DataSubLineaOperativa();
            DataSubLineaOperativa.TableName = "SubLínea Operativa 7-12";
            DataTable DataContinuidad = carga.DataContinuidad();
            DataContinuidad.TableName = "Continuidad 8-12";
            DataTable DataEstado = carga.DataEstadoEventos();
            DataEstado.TableName = "Estados 9-12";
            DataTable DataMoneda = carga.DataMoneda();
            DataMoneda.TableName = "Moneda 10-12";
            DataTable DataRecuperacion = carga.DataRecuperacion();
            DataRecuperacion.TableName = "Recuperación 11-12";
            


            System.Data.DataSet ds = new System.Data.DataSet();
            /*ds.Tables.Add(DataEmpresa);
            ds.Tables.Add(DataUbicacion);
            ds.Tables.Add(DataServicio);
            ds.Tables.Add(DataCanal);
            ds.Tables.Add(DataGenerador);
            ds.Tables.Add(DataResponsable);*/
            ds.Tables.Add(DataEventos);
            ds.Tables.Add(DataProcesos);
            ds.Tables.Add(DataClaseRiesgo);
            ds.Tables.Add(DataSubClaseRiesgo);
            ds.Tables.Add(DataPerdida);
            ds.Tables.Add(DataLineaOperativa);
            ds.Tables.Add(DataSubLineaOperativa);
            ds.Tables.Add(DataContinuidad);
            ds.Tables.Add(DataEstado);
            ds.Tables.Add(DataMoneda);
            ds.Tables.Add(DataRecuperacion);
            

            XLWorkbook workbook = new XLWorkbook();
            workbook.Worksheets.Add(ds);
            HttpResponse httpResponse = Response;
            httpResponse.Clear();
            httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            httpResponse.AddHeader("content-disposition", "attachment;filename=\"" + filename + ".xlsx\"");

            // Flush the workbook to the Response.OutputStream
            using (MemoryStream memoryStream = new MemoryStream())
            {
                workbook.SaveAs(memoryStream);
                memoryStream.WriteTo(httpResponse.OutputStream);
                memoryStream.Close();
            }

            httpResponse.End();
        }
        protected void exportExcelControles(HttpResponse Response, string filename)
        {

            cCargaMasivaRCE carga = new cCargaMasivaRCE();
            DataTable DataResponsable = carga.DataResponsable();
            DataResponsable.TableName = "Responsable 1-8";
            DataTable DataPeriodicidad = carga.DataPeriodicidad();
            DataPeriodicidad.TableName = "Periodicidad 2-8";
            DataTable DataTest = carga.DataTest();
            DataTest.TableName = "Test 3-8";
            DataTable DataReduce = carga.DataMitigaControl();
            DataReduce.TableName = "Reduce 4-8";
            //DataTable DataClaseControl = carga.DataClaseControl();
            //DataClaseControl.TableName = "Clase Control 5-11";
            //DataTable DataTipoControl = carga.DataTipoControl();
            //DataTipoControl.TableName = "Tipo Control 7-11";
            //DataTable DataExperiencia = carga.DataExperiencia();
            //DataExperiencia.TableName = "Responsable Experiencia 7-11";
            //DataTable DataDocumentacion = carga.DataDocumentacion();
            //DataDocumentacion.TableName = "Documentacion 8-11";
            //DataTable DataResponsabilidad = carga.DataResponsabilidad();
            //DataResponsabilidad.TableName = "Resposabilidad 9-11";
            DataTable DataGrupoTrabajo = carga.DataGruposTrabajo();
            DataGrupoTrabajo.TableName = "Grupos de Trabajo 5-8";
            //DataTable DataVariablesCalificacion = carga.DataVariablesCalificacionesControles();
            //DataVariablesCalificacion.TableName = "Variables Calificación 7-7";
            DataTable DataParametrosVariable = carga.SeleccionarParametrosVariable();
            DataParametrosVariable.TableName = "Parametros Variable 6-8";
            DataTable DataEstado = carga.DataEstado();
            DataEstado.TableName = "Estados 7-8";

            DataTable DataPerdida = carga.DataPerdida();
            DataPerdida.TableName = "Tipo Perdida 8-8";

            System.Data.DataSet ds = new System.Data.DataSet();
            ds.Tables.Add(DataResponsable);
            ds.Tables.Add(DataPeriodicidad);
            ds.Tables.Add(DataTest);
            ds.Tables.Add(DataReduce);
            //ds.Tables.Add(DataClaseControl);
            //ds.Tables.Add(DataTipoControl);
            //ds.Tables.Add(DataExperiencia);
            //ds.Tables.Add(DataDocumentacion);
            //ds.Tables.Add(DataResponsabilidad);
            ds.Tables.Add(DataGrupoTrabajo);
            //ds.Tables.Add(DataVariablesCalificacion);
            ds.Tables.Add(DataParametrosVariable);
            ds.Tables.Add(DataEstado);
            ds.Tables.Add(DataPerdida);
            // Create the workbook
            XLWorkbook workbook = new XLWorkbook();
            workbook.Worksheets.Add(ds);
            // Prepare the response
            HttpResponse httpResponse = Response;
            httpResponse.Clear();
            httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            httpResponse.AddHeader("content-disposition", "attachment;filename=\"" + filename + ".xlsx\"");

            // Flush the workbook to the Response.OutputStream
            using (MemoryStream memoryStream = new MemoryStream())
            {
                workbook.SaveAs(memoryStream);
                memoryStream.WriteTo(httpResponse.OutputStream);
                memoryStream.Close();
            }

            httpResponse.End();
        }
        protected void exportExcelRiesgosvsControles(HttpResponse Response, string filename)
        {
            cParametrizacionRiesgos cParametrizacionRiesgos = new cParametrizacionRiesgos();
            cCargaMasivaRCE carga = new cCargaMasivaRCE();
            DataTable DataRiesgos = carga.DataRiesgo();
            DataRiesgos.TableName = "Riesgos 1-3";
            DataTable DataControles = carga.DataControles();
            DataControles.TableName = "Controles 2-3";
            //DataTable DataCausasRiegos = carga.DataCausasRiesgos();
            DataTable DataCausasRiegos = cParametrizacionRiesgos.loadInfoCausas();
            /*ListCausas = ListCausas.Replace("|", ",");*/
            DataCausasRiegos.TableName = "Causas del Riesgo 3-3";
            DataTable DataPerdida = carga.DataPerdida();
            DataPerdida.TableName = "Tipo Perdida 4-4";

            System.Data.DataSet ds = new System.Data.DataSet();
            ds.Tables.Add(DataRiesgos);
            ds.Tables.Add(DataControles);
            ds.Tables.Add(DataCausasRiegos);
            ds.Tables.Add(DataPerdida);
            // Create the workbook
            XLWorkbook workbook = new XLWorkbook();
            //workbook.Worksheets.Add("Sample").Cell(1, 1).SetValue("Hello World");
            workbook.Worksheets.Add(ds);
            // Prepare the response
            HttpResponse httpResponse = Response;
            httpResponse.Clear();
            httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            httpResponse.AddHeader("content-disposition", "attachment;filename=\"" + filename + ".xlsx\"");

            // Flush the workbook to the Response.OutputStream
            using (MemoryStream memoryStream = new MemoryStream())
            {
                workbook.SaveAs(memoryStream);
                memoryStream.WriteTo(httpResponse.OutputStream);
                memoryStream.Close();
            }

            httpResponse.End();
        }
        public DataTable GetDataTable(GridView dtg)
        {
            DataTable dt = new DataTable();
            if (dtg.HeaderRow != null)
            {
                for (int i = 0; i < dtg.HeaderRow.Cells.Count; i++)
                {
                    dt.Columns.Add(dtg.RowHeaderColumn[i].ToString());
                }
            }

            foreach (GridViewRow row in dtg.Rows)
            {
                DataRow dr;
                dr = dt.NewRow();

                for (int i = 0; i < row.Cells.Count; i++)
                {
                    dr[i] = row.Cells[i].Text.Replace(" ", "");
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
        protected void Bload_Click(object sender, EventArgs e)
        {
            try
            {
                string NombreArchivo = string.Empty;
                if (FUloadExcel.HasFile)
                {
                    if (System.IO.Path.GetExtension(FUloadExcel.FileName).ToLower() == ".xls" || System.IO.Path.GetExtension(FUloadExcel.FileName).ToLower() == ".xlsx")
                    {
                        NombreArchivo = FUloadExcel.FileName;
                        lblRecalculoC.Visible = false;
                        ImbCalcularControl.Visible = false;
                        ImbCancel.Visible = false;
                        if (DDLopciones.SelectedValue == "1")
                        {
                            LoadFilePLantillaRiesgos();
                            LoadGridRiesgos();
                        }
                        if (DDLopciones.SelectedValue == "2")
                        {
                            LoadPlantilla(DDLopciones.SelectedValue);
                            lblRecalculoC.Visible = true;
                            ImbCalcularControl.Visible = true;
                            ImbCancel.Visible = true;
                            LoadGridControl();
                            omb.ShowMessage("Archivo procesado con éxito", 3, "Atención");
                        }
                        if (DDLopciones.SelectedValue == "3")
                        {
                            LoadFilePlantillaEventos();
                            LoadGridEventos();
                        }
                        if (DDLopciones.SelectedValue == "5")
                        {
                            LoadFilePlantillaEventosDatosComplementarios();
                            LoadGridEventos();
                        }
                        if (DDLopciones.SelectedValue == "6")
                        {
                            LoadFilePlantillaEventosConDatosComplementarios();
                            LoadGridEventos();
                        }
                        if (DDLopciones.SelectedValue == "4")
                        {
                            LoadPlantilla(DDLopciones.SelectedValue);
                            LoadGridRiesgovsControl();
                            omb.ShowMessage("Archivo procesado con éxito", 3, "Atención");
                        }
                    }
                    else
                    {
                        Mensaje("Por favor usar archivos .xls o .xlsx");
                    }
                }
                else
                {
                    Mensaje("El control no contiene ningún archivo.");
                }


            }
            catch (Exception ex)
            {
                Mensaje("Error al analizar la información. " + ex.Message);
            }
        }
        private void LoadGridRiesgos()
        {
            string IdRiesgo = Session["IdRiesgo"].ToString();
            cCargaMasivaRCE Carga = new cCargaMasivaRCE();
            DataTable dtRiesgo = Carga.GridRiesgos(IdRiesgo);
            GVriesgos.DataSource = dtRiesgo;
            GVriesgos.PageIndex = pagIndex;
            GVriesgos.DataBind();
            TgridRiesgos.Visible = true;
        }
        private void LoadGridControl()
        {
            string IdControl = Session["IdControl"].ToString();
            cCargaMasivaRCE Carga = new cCargaMasivaRCE();
            DataTable dtControl = Carga.GridControles(IdControl);
            GVcontrol.DataSource = dtControl;
            GVcontrol.PageIndex = pagIndex;
            GVcontrol.DataBind();
            TgridRiesgos.Visible = true;
        }
        private void LoadGridEventos()
        {

            cCargaMasivaRCE Carga = new cCargaMasivaRCE();
            DataTable dtEventos = new DataTable();
            if (DDLopciones.SelectedValue == "3" || DDLopciones.SelectedValue == "6")
            {
                string IdEvento = Session["IdEvento"].ToString();
                dtEventos = Carga.GridEventos(IdEvento);
            }
            else
            {
                string IdEvento = Session["IdsEventos"].ToString().TrimEnd(',');
                IdEvento = IdEvento + ")";

                dtEventos = Carga.GridEventosMod(IdEvento);
            }
            GVeventos.DataSource = dtEventos;
            GVeventos.PageIndex = pagIndex;
            GVeventos.DataBind();
            TgridRiesgos.Visible = true;

        }
        private void LoadGridRiesgovsControl()
        {
            string IdControlesRiesgo = Session["IdControlesRiesgo"].ToString();
            cCargaMasivaRCE Carga = new cCargaMasivaRCE();
            DataTable dtControlesRiesgo = Carga.GridRiesgosvsControl(IdControlesRiesgo);
            GVriesgovscontrol.DataSource = dtControlesRiesgo;
            GVriesgovscontrol.PageIndex = pagIndex;
            GVriesgovscontrol.DataBind();
            TgridRiesgos.Visible = true;
        }
        private void LoadFilePlantillaControl()
        {
            int IdUsuario = Convert.ToInt32(Session["IdUsuario"].ToString());
            string FechaRegistro = DateTime.Now.ToString();
            //Best Way To read file direct from stream
            cCargaMasivaRCE CargaMasiva = new cCargaMasivaRCE();
            IExcelDataReader excelReader = null;
            System.Data.DataSet ds = new System.Data.DataSet();
            try
            {

                //file.InputStream is the file stream stored in memeory by any ways like by upload file control or from database
                int excelFlag = 1; //this flag us used for execl file format .xls or .xlsx
                if (excelFlag == 1)
                {
                    //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(FUloadExcel.PostedFile.InputStream);
                    excelReader.IsFirstRowAsColumnNames = true;
                }
                else if (excelFlag == 2)
                {
                    //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(FUloadExcel.PostedFile.InputStream);
                    //excelReader.IsFirstRowAsColumnNames = true;
                }
            }
            catch (Exception ex)
            {
                Mensaje("Error al analizar el archivo Eventos. " + ex.Message);
            }
            if (excelReader != null)
            {
                //...
                //3. DataSet - The result of each spreadsheet will be created in the result.Tables
                ds = excelReader.AsDataSet();
                //...
                ////4. DataSet - Create column names from first row
                //excelReader.IsFirstRowAsColumnNames = true;
                //DataSet result = excelReader.AsDataSet();

                ////5. Data Reader methods
                //while (excelReader.Read())
                //{
                //    //excelReader.GetInt32(0);
                //}
                DataTable datos = ds.Tables[0];
                DataTable dt = CargaMasiva.LastCodControl();
                Session["IdControl"] = dt.Rows[0][0].ToString();
                for (int i = 1; i < datos.Rows.Count; i++)
                {
                    if (datos.Rows[i][0].ToString().Trim() != "")
                    {
                        string nombre = datos.Rows[i][0].ToString().Trim();
                        string descripcion = datos.Rows[i][1].ToString().Trim();
                        string objetivo = datos.Rows[i][2].ToString().Trim();
                        string responsable = datos.Rows[i][3].ToString().Trim();
                        string periodicidad = datos.Rows[i][4].ToString().Trim();
                        string test = datos.Rows[i][5].ToString().Trim();
                        string Reduce = datos.Rows[i][6].ToString().Trim();
                        /*string ClaseControl = datos.Rows[i][7].ToString().Trim();
                        string TipoControl = datos.Rows[i][8].ToString().Trim();
                        string Experiencia = datos.Rows[i][9].ToString().Trim();
                        string Documentacion = datos.Rows[i][10].ToString().Trim();
                        string Responsabilidad = datos.Rows[i][11].ToString().Trim();*/
                        //string calificacion = calcularCalificacionControl(ClaseControl,TipoControl,Experiencia,Documentacion,Responsabilidad);
                        string NombreVariable = datos.Rows[i][7].ToString().Trim();
                        string IdCategoria = datos.Rows[i][8].ToString().Trim();
                        string NombreCategoria = datos.Rows[i][9].ToString().Trim();
                        string ResponsableEjecucion = datos.Rows[i][10].ToString().Trim();
                        try
                        {
                            CargaMasiva.registrarControl(nombre, descripcion, objetivo, responsable, periodicidad, test, "0", NombreVariable, IdCategoria, NombreCategoria, Reduce, IdUsuario, ResponsableEjecucion);
                        }
                        catch (Exception ex)
                        {
                            Mensaje("Error al registrar el control de la linea: " + i + "." + ex.Message);
                        }
                    }
                }
                excelReader.Close();
            }
        }

        /*
         *                     -yoendy Ca - Inserción Masiva - 
         * Técnica utilizada:  -SQL bulk Insert - escribe en plano
         *                     -* Si existe error hace rollback - sino commit
         *                     -Algoritmo para incrementar codigo de evento con Arrays
         * Envia correo:       -codigoEvento , descripcion evento
         *                     -por idJerarquia  tipeado en el xlsx
         * Validaciones:       -campo idCargo jerarquia este vacio en el xlsx
         *                     -valor del campo idCargo jerarquia exista en bd
         *                     -extensión del archivo
         *                     -número de campos (estructura)
         *                     -Campos básicos para creación de eventos no esten vacíos
         */
        private void LoadFilePlantillaEventos()
        {
            #region variables
            //string path_planos = System.Configuration.ConfigurationManager.AppSettings["PATH_PLANOS"].ToString();
            //string nombreArchivo = Constantes.NOMBRE_CARGAS;
            //nombreArchivo = Procs.NombreArchivo(path_planos, nombreArchivo, ".txt");
            //string savedFileName = Path.Combine(path_planos, nombreArchivo);
            //string destino = System.IO.Path.Combine(path_planos, nombreArchivo);
            StreamWriter writer = null;
            StringBuilder s = null;
            FileInfo info = null;
            string strPrefixEvento = string.Empty, nombreSp = "[Eventos].[pa_RegistrarEvento]";
            int contador = 0;
            IExcelDataReader excelReader = null;
            string ultimoEE = string.Empty, ultimoEV = string.Empty, ultimoEG = string.Empty, DescipcionEvento = string.Empty, idResponsable = string.Empty;
            DataTable dtJerarquia = new DataTable();
            dtJerarquia.Columns.Add("idJerarquia"); dtJerarquia.Columns.Add("codigoEvento"); dtJerarquia.Columns.Add("descripcion_evento");
            cCargaMasivaRCE CargaMasiva = new cCargaMasivaRCE();
            System.Data.DataSet ds = new System.Data.DataSet();
            cambiaCultura cc = new cambiaCultura();
            #endregion variables
            try
            {
                //if (!System.IO.Directory.Exists(path_planos))
                //{
                //    System.IO.Directory.CreateDirectory(path_planos);
                //}

                //writer = File.AppendText(destino);
                //info = new FileInfo(destino);

                int IdUsuario = Convert.ToInt32(Session["IdUsuario"].ToString());
                string FechaRegistro = DateTime.Now.ToString();

                int excelFlag = 1;
                if (excelFlag == 1)
                {
                    excelReader = ExcelReaderFactory.CreateBinaryReader(FUloadExcel.PostedFile.InputStream);
                    excelReader.IsFirstRowAsColumnNames = true;
                }
                else if (excelFlag == 2)
                {
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(FUloadExcel.PostedFile.InputStream);
                }
            }
            catch (Exception ex)
            {
                Mensaje("Error al analizar el archivo Eventos. " + ex.Message);
            }
            if (excelReader != null)
            {
                int Result = 0;
                ds = excelReader.AsDataSet();
                int IdUsuario = Convert.ToInt32(Session["IdUsuario"].ToString()), Linea = 0;
                DataTable datos = ds.Tables[0];
                DataTable dt = CargaMasiva.LastCodEvento();
                DataTable dts = new DataTable();
                Session["IdEvento"] = dt.Rows[0][0].ToString();
                string maxValor = string.Empty;
                DataTable dttemp = new DataTable();
                dttemp.Columns.Add("PrefijoEV"); dttemp.Columns.Add("PrefijoEG"); dttemp.Columns.Add("PrefijoEE");

                string[] prefijo = new string[3] { "EV", "EG", "EE" };
                DataRow drFila = dttemp.NewRow();
                for (int i = 0; i < prefijo.Length; i++)
                {
                    maxValor = prefijo[i];
                    dts = CargaMasiva.ultimoCodigoEvento(maxValor);

                    foreach (DataRow row in dts.Rows)
                    {
                        maxValor = row[0].ToString();
                        if (prefijo[i] == "EV")
                        {
                            drFila["PrefijoEV"] = "";
                            dttemp.Rows.Add(drFila);
                            dttemp.Rows[0][0] = maxValor;
                            break;
                        }
                        if (prefijo[i] == "EG")
                        {
                            dttemp.Rows[0][1] = maxValor;
                            break;
                        }
                        if (prefijo[i] == "EE")
                        {
                            dttemp.Rows[0][2] = maxValor;
                            break;
                        }
                    }
                }
                for (int i = 1; i < datos.Rows.Count; i++)
                {
                    if (datos.Rows[i][0].ToString().Trim() != "")
                    {
                        Linea++;
                        string idEmpresa = datos.Rows[i][0].ToString().Trim();
                        string Region = datos.Rows[i][1].ToString().Trim();
                        string Pais = datos.Rows[i][2].ToString().Trim();
                        string Departamento = datos.Rows[i][3].ToString().Trim();
                        string Ciudad = datos.Rows[i][4].ToString().Trim();
                        string Oficina = datos.Rows[i][5].ToString().Trim();
                        string Detalle_Ubicacion = datos.Rows[i][6].ToString().Trim();
                        string Descripcion_Evento = datos.Rows[i][7].ToString().Trim();
                        string IdServicio = datos.Rows[i][8].ToString().Trim();
                        string IdSubServicio = datos.Rows[i][9].ToString().Trim();
                        string FechaInicio = datos.Rows[i][10].ToString().Trim();
                        string HoraInicio = datos.Rows[i][11].ToString().Trim();
                        string FechaFinalizacion = datos.Rows[i][12].ToString().Trim();
                        string HoraFinalizacion = datos.Rows[i][13].ToString().Trim();
                        string FechaDescubrimiento = datos.Rows[i][14].ToString().Trim();
                        string HoraDescubrimiento = datos.Rows[i][15].ToString().Trim();
                        string Canal = datos.Rows[i][16].ToString().Trim();
                        string Generador = datos.Rows[i][17].ToString().Trim();
                        string IdCargoResponsable = datos.Rows[i][18].ToString().Trim();
                        string CargoResponsable = datos.Rows[i][19].ToString().Trim();
                        string CuantiaPerdida = datos.Rows[i][20].ToString().Trim();
                        //string ImpactoCualitativo = datos.Rows[i][21].ToString().Trim();
                        /**************** Se eliminan de la creacion ***************
                        string cadenaValor = datos.Rows[i][21].ToString().Trim();
                        string macroProceso = datos.Rows[i][22].ToString().Trim();
                        string proceso = datos.Rows[i][23].ToString().Trim();
                        string subproceso = datos.Rows[i][24].ToString().Trim();
                        string actividad = datos.Rows[i][25].ToString().Trim();
                        string IdClase = datos.Rows[i][26].ToString().Trim();
                        string IdSubClase = datos.Rows[i][27].ToString().Trim();
                        string IdTipoPerdidaEvento = datos.Rows[i][28].ToString().Trim();
                        string IdLineaProceso = datos.Rows[i][29].ToString().Trim();
                        string IdSubLineaProceso = datos.Rows[i][30].ToString().Trim();
                        string AfectaContinudad = datos.Rows[i][31].ToString().Trim();
                        string IdEstado = datos.Rows[i][32].ToString().Trim();
                        string Observaciones = datos.Rows[i][33].ToString().Trim();
                        string CuentaPUC = datos.Rows[i][34].ToString().Trim();
                        string CuentaOrden = datos.Rows[i][35].ToString().Trim();
                        string TasaCambio1 = datos.Rows[i][36].ToString().Trim();
                        string ValorPesos1 = datos.Rows[i][37].ToString().Trim();
                        string ValorRecuperadoTotal = datos.Rows[i][38].ToString().Trim();
                        string Moneda2 = datos.Rows[i][39].ToString().Trim();
                        string TasaCambio2 = datos.Rows[i][40].ToString().Trim();
                        string ValorPesos2 = datos.Rows[i][41].ToString().Trim();
                        string Recuperacion = datos.Rows[i][42].ToString().Trim();
                        string FechaContabilidad = datos.Rows[i][43].ToString().Trim();
                        string HoraContabilidad = datos.Rows[i][44].ToString().Trim();
                        */
                        string fechaInicioCp = string.Empty;
                        string fechaFinalizacionCp = string.Empty;
                        string fechaDescubrimientoCp = string.Empty;
                        string fechaContabilidad = string.Empty;
                        string hoy = string.Empty;

                        //Controlador de fechas y horas - Esto porque el valor viene en  OLE Automation y no lo admite el struct
                        if (!string.IsNullOrEmpty(FechaInicio.TrimEnd()))
                        {
                            DateTime FechaInicioCm = DateTime.FromOADate(Convert.ToDouble(FechaInicio));
                            fechaInicioCp = string.Format("{0:yyyy-MM-dd}", Convert.ToDateTime(FechaInicioCm));
                        }
                        else
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Fecha de Inicio</B> no puede estar vacío!");
                        }

                        if (!string.IsNullOrEmpty(FechaFinalizacion))
                        {
                            DateTime FechaFinalizacionCm = DateTime.FromOADate(Convert.ToDouble(FechaFinalizacion));
                            fechaFinalizacionCp = string.Format("{0:yyyy-MM-dd}", Convert.ToDateTime(FechaFinalizacionCm));
                        }
                        else
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Fecha de Finalización</B> no puede estar vacío!");
                        }

                        if (!string.IsNullOrEmpty(FechaDescubrimiento))
                        {
                            DateTime FechaDescubrimientoCm = DateTime.FromOADate(Convert.ToDouble(FechaDescubrimiento));
                            fechaDescubrimientoCp = string.Format("{0:yyyy-MM-dd}", Convert.ToDateTime(FechaDescubrimientoCm));
                        }
                        else
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Fecha de Descubrimiento</B> no puede estar vacío!");
                        }

                        /*if (!string.IsNullOrEmpty(FechaContabilidad))
                        {
                            DateTime FechaContabilidadCm = DateTime.FromOADate(Convert.ToDouble(FechaContabilidad));
                            fechaContabilidad = string.Format("{0:yyyy-MM-dd}", Convert.ToDateTime(FechaContabilidadCm));
                        }
                        else
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Fecha de Contabilidad</B> no puede estar vacío!");
                        }*/

                        DateTime hoyCm = Convert.ToDateTime(DateTime.Now.ToShortDateString());
                        hoy = string.Format("{0:yyyy-MM-dd}", hoyCm);

                        //Hora inicio
                        if (!string.IsNullOrEmpty(HoraInicio))
                        {
                            DateTime HoraInicioCm = DateTime.FromOADate(Convert.ToDouble(HoraInicio));
                            string HoraInciCp = string.Format("{0:hh:mm tt}", Convert.ToDateTime(HoraInicioCm));
                            
                            //case format a.m/p.m
                            DateTime Fecha2 = DateTime.FromOADate(Convert.ToDouble(HoraInicio));
                            string HoraFecha2 = Fecha2.ToString("hh:mm tt", CultureInfo.InvariantCulture);
                            HoraFecha2 = HoraFecha2.Replace("PM", "p.m").Replace("AM", "a.m");

                            HoraInicio = HoraInciCp.TrimEnd('.');
                            HoraInicio = HoraFecha2;
                        }

                        //HoraFinalizacion
                        if (!string.IsNullOrEmpty(HoraFinalizacion))
                        {
                            DateTime HoraFinalCm = DateTime.FromOADate(Convert.ToDouble(HoraFinalizacion));
                            string HoraFinalCp = string.Format("{0:hh:mm tt}", Convert.ToDateTime(HoraFinalCm));

                            //case format a.m/p.m
                            DateTime Fecha2 = DateTime.FromOADate(Convert.ToDouble(HoraFinalizacion));
                            string HoraFecha2 = Fecha2.ToString("hh:mm tt", CultureInfo.InvariantCulture);
                            HoraFecha2 = HoraFecha2.Replace("PM", "p.m").Replace("AM", "a.m");

                            HoraFinalizacion = HoraFinalCp.TrimEnd('.');
                            HoraFinalizacion = HoraFecha2;
                        }

                        //HoraDescubrimiento
                        if (!string.IsNullOrEmpty(HoraDescubrimiento))
                        {
                            DateTime HoraDesCm = DateTime.FromOADate(Convert.ToDouble(HoraDescubrimiento));
                            string HoraDescCp = string.Format("{0:hh:mm tt}", Convert.ToDateTime(HoraDesCm));

                            //case format a.m/p.m
                            DateTime Fecha2 = DateTime.FromOADate(Convert.ToDouble(HoraDescubrimiento));
                            string HoraFecha2 = Fecha2.ToString("hh:mm tt", CultureInfo.InvariantCulture);
                            HoraFecha2 = HoraFecha2.Replace("PM", "p.m").Replace("AM", "a.m");

                            HoraDescubrimiento = HoraDescCp.TrimEnd('.');
                            HoraDescubrimiento = HoraFecha2;
                        }

                        //HoraContabilidad
                        /*DateTime HoraContCm = DateTime.FromOADate(Convert.ToDouble(HoraContabilidad));
                        string HoraContCp = string.Format("{0:hh:mm tt}", Convert.ToDateTime(HoraContCm));
                        HoraContabilidad = HoraContCp.TrimEnd('.');*/

                        DescipcionEvento = Descripcion_Evento;
                        idResponsable = IdCargoResponsable;

                        /*if (!string.IsNullOrEmpty(Recuperacion))
                        {
                            if (Convert.ToInt32(Recuperacion) > 1)
                            {
                                throw new Exception("Verificar la Línea: <B>" + (Linea + 2) + "</B> del archivo .El valor de la columna  <B>Recuperación,</B> no puede ser mayor a 1.");
                            }
                        }
                        else
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Recuperación</B> no puede estar vacío!");
                        }*/

                        if (string.IsNullOrEmpty(idEmpresa.TrimEnd()) || string.IsNullOrEmpty(Region.TrimEnd()) || string.IsNullOrEmpty(Pais.TrimEnd()) || string.IsNullOrEmpty(Departamento.TrimEnd())
                            || string.IsNullOrEmpty(Ciudad.TrimEnd()) || string.IsNullOrEmpty(Oficina.TrimEnd()) || string.IsNullOrEmpty(IdServicio.TrimEnd()) || string.IsNullOrEmpty(IdSubServicio.TrimEnd())
                            || string.IsNullOrEmpty(Canal.TrimEnd()) || string.IsNullOrEmpty(CuantiaPerdida.TrimEnd()))
                        {
                            string Columna = string.Empty;
                            int ContadorVacios = 0;

                            if (string.IsNullOrEmpty(Region))
                            {
                                Columna = "Region";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(Pais))
                            {
                                Columna = Columna + ", Pais";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(Departamento))
                            {
                                Columna = Columna + ", Departamento";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(Ciudad))
                            {
                                Columna = Columna + ", Ciudad";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(Oficina))
                            {
                                Columna = Columna + ", Oficina";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(IdServicio))
                            {
                                Columna = Columna + ", Servicio/Producto";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(IdSubServicio))
                            {
                                Columna = Columna + ", SubServicio/SubProducto";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(Canal))
                            {
                                Columna = Columna + ", Canal";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(CuantiaPerdida))
                            {
                                Columna = Columna + ", Posible cuantía pérdida";
                                ContadorVacios++;
                            }

                            if (ContadorVacios > 1)
                            {
                                throw new Exception("Verificar la Línea: <B>" + (Linea + 2) + "</B> del archivo. El valor de las columnas  <B>" + Columna + "</B> no pueden estar vacías.");
                            }
                            else
                            {
                                Columna = Columna.TrimStart(',');

                                throw new Exception("Verificar la Línea: <B>" + (Linea + 2) + "</B> del archivo. El valor de la columna  <B>" + Columna + "</B> no puede estar vacía.");
                            }
                        }

                        //Incrementador de copdigoEvento
                        switch (idEmpresa.Trim())
                        {
                            case "1":
                                strPrefixEvento = "EV";
                                string sprint = string.Empty;
                                int maxLenght = 0;
                                int sub = 0;
                                if (ultimoEV == string.Empty)
                                {
                                    sprint = dttemp.Rows[0][0].ToString();
                                    maxLenght = sprint.Length;
                                    if (maxLenght == 2)
                                    {
                                        sub = 1;
                                    }
                                }
                                else
                                {
                                    sprint = ultimoEV;
                                    maxLenght = sprint.Length;
                                }
                                if (maxLenght > 2)
                                {
                                    sub = Convert.ToInt32(sprint.Substring(2));
                                    sub += 1;
                                }

                                strPrefixEvento = strPrefixEvento + sub.ToString();
                                ultimoEV = strPrefixEvento;
                                strPrefixEvento = ultimoEV;

                                break;
                            case "2":
                                strPrefixEvento = "EG";
                                string sprint2 = string.Empty;
                                int maxLenght2 = 0;
                                int sub2 = 0;
                                if (ultimoEG == string.Empty)
                                {
                                    sprint2 = dttemp.Rows[0][1].ToString();
                                    maxLenght2 = sprint2.Length;
                                    if (maxLenght2 == 2)
                                    {
                                        sub2 = 1;
                                    }
                                }
                                else
                                {
                                    sprint2 = ultimoEG;
                                    maxLenght2 = sprint2.Length;
                                }
                                if (maxLenght2 > 2)
                                {
                                    sub2 = Convert.ToInt32(sprint2.Substring(2));
                                    sub2 += 1;
                                }

                                strPrefixEvento = strPrefixEvento + sub2.ToString();
                                ultimoEG = strPrefixEvento;
                                strPrefixEvento = ultimoEG;

                                break;
                            case "3":
                                strPrefixEvento = "EE";
                                string sprint3 = string.Empty;
                                int maxLenght3 = 0;
                                int sub3 = 0;
                                if (ultimoEE == string.Empty)
                                {
                                    sprint3 = dttemp.Rows[0][2].ToString();
                                    maxLenght3 = sprint3.Length;
                                    if (maxLenght3 == 2)
                                    {
                                        sub3 = 1;
                                    }
                                }
                                else
                                {
                                    sprint3 = ultimoEE;
                                    maxLenght3 = sprint3.Length;
                                }
                                if (maxLenght3 > 2)
                                {
                                    sub3 = Convert.ToInt32(sprint3.Substring(2));
                                    sub3 += 1;
                                }

                                strPrefixEvento = strPrefixEvento + sub3.ToString();
                                ultimoEE = strPrefixEvento;
                                strPrefixEvento = ultimoEE;
                                break;
                        }
                        //cCargaMasivaRCE cCargue = new cCargaMasivaRCE();
                        try
                        {
                            /*s = new StringBuilder();
                            s.Append(strPrefixEvento); s.Append(Constantes.TAB);
                            s.Append(idEmpresa); s.Append(Constantes.TAB);
                            s.Append(Region); s.Append(Constantes.TAB);
                            s.Append(Pais); s.Append(Constantes.TAB);
                            s.Append(Departamento); s.Append(Constantes.TAB);
                            s.Append(Ciudad); s.Append(Constantes.TAB);
                            s.Append(Oficina); s.Append(Constantes.TAB);
                            s.Append(Detalle_Ubicacion); s.Append(Constantes.TAB);
                            s.Append(Descripcion_Evento); s.Append(Constantes.TAB);
                            s.Append(IdServicio); s.Append(Constantes.TAB);
                            s.Append(IdSubServicio); s.Append(Constantes.TAB);
                            s.Append(fechaInicioCp); s.Append(Constantes.TAB);
                            s.Append(HoraInicio); s.Append(Constantes.TAB);
                            s.Append(fechaFinalizacionCp); s.Append(Constantes.TAB);
                            s.Append(HoraFinalizacion); s.Append(Constantes.TAB);
                            s.Append(fechaDescubrimientoCp); s.Append(Constantes.TAB);
                            s.Append(HoraDescubrimiento); s.Append(Constantes.TAB);
                            s.Append(Canal); s.Append(Constantes.TAB);
                            s.Append(Generador); s.Append(Constantes.TAB);
                            s.Append(IdCargoResponsable); s.Append(Constantes.TAB);
                            s.Append(CargoResponsable); s.Append(Constantes.TAB);
                            s.Append(CuantiaPerdida); s.Append(Constantes.TAB);



                            s.Append(IdUsuario); s.Append(Constantes.TAB);
                            s.Append(hoy);*/
                            /************* Se elimina de la creacion inicial **********
                            
                            s.Append(cadenaValor); s.Append(Constantes.TAB);
                            s.Append(macroProceso); s.Append(Constantes.TAB);
                            s.Append(proceso); s.Append(Constantes.TAB);
                            s.Append(subproceso); s.Append(Constantes.TAB);
                            s.Append(actividad); s.Append(Constantes.TAB);
                            s.Append(IdClase); s.Append(Constantes.TAB);
                            s.Append(IdSubClase); s.Append(Constantes.TAB);
                            s.Append(IdTipoPerdidaEvento); s.Append(Constantes.TAB);
                            s.Append(IdLineaProceso); s.Append(Constantes.TAB);
                            s.Append(IdSubLineaProceso); s.Append(Constantes.TAB);
                            s.Append(AfectaContinudad); s.Append(Constantes.TAB);
                            s.Append(IdEstado); s.Append(Constantes.TAB);
                            s.Append(Observaciones); s.Append(Constantes.TAB);
                            s.Append(CuentaPUC); s.Append(Constantes.TAB);
                            s.Append(CuentaOrden); s.Append(Constantes.TAB);
                            s.Append(TasaCambio1); s.Append(Constantes.TAB);
                            s.Append(ValorPesos1); s.Append(Constantes.TAB);
                            s.Append(ValorRecuperadoTotal); s.Append(Constantes.TAB);
                            s.Append(Moneda2); s.Append(Constantes.TAB);
                            s.Append(TasaCambio2); s.Append(Constantes.TAB);
                            s.Append(ValorPesos2); s.Append(Constantes.TAB);
                            s.Append(Recuperacion); s.Append(Constantes.TAB);
                            s.Append(fechaContabilidad); s.Append(Constantes.TAB);
                            s.Append(HoraContabilidad); s.Append(Constantes.TAB);
                            
                            */
                            /*writer.WriteLine(s.ToString());
                            */
                            DataRow registros = dtJerarquia.NewRow();
                            registros["idJerarquia"] = datos.Rows[i][18].ToString().Trim();
                            registros["codigoEvento"] = strPrefixEvento;
                            registros["descripcion_evento"] = Descripcion_Evento;
                            dtJerarquia.Rows.Add(registros);

                            CargaMasiva.registrarEvento(idEmpresa, Region, Pais, Departamento, Ciudad, Oficina, Detalle_Ubicacion, Descripcion_Evento, IdServicio, IdSubServicio,
                                fechaInicioCp, HoraInicio, fechaFinalizacionCp, HoraFinalizacion,fechaDescubrimientoCp,HoraDescubrimiento,Canal, Generador, CargoResponsable, 
                                CuantiaPerdida, hoy, Generador,IdUsuario
                                //, ImpactoCualitativo
                                );

                            contador++;

                        }
                        catch (Exception ex)
                        {
                            Mensaje("Error al registrar el evento de la linea: " + i + "." + ex.Message);
                        }
                    }
                    else
                    {
                        string idEmpresa = datos.Rows[i][0].ToString().Trim();
                        string Region = datos.Rows[i][1].ToString().Trim();
                        string Pais = datos.Rows[i][2].ToString().Trim();
                        string Departamento = datos.Rows[i][3].ToString().Trim();
                        //string cadenaValor = datos.Rows[i][21].ToString().Trim();
                        //string macroProceso = datos.Rows[i][22].ToString().Trim();
                        //string proceso = datos.Rows[i][23].ToString().Trim();
                        //string subproceso = datos.Rows[i][24].ToString().Trim();
                        //string ValorRecuperadoTotal = datos.Rows[i][38].ToString().Trim();
                        //string Moneda2 = datos.Rows[i][39].ToString().Trim();
                        //string TasaCambio2 = datos.Rows[i][40].ToString().Trim();
                        //string ValorPesos2 = datos.Rows[i][41].ToString().Trim();
                        //string Recuperacion = datos.Rows[i][42].ToString().Trim();
                        //string FechaContabilidad = datos.Rows[i][43].ToString().Trim();
                        //string HoraContabilidad = datos.Rows[i][44].ToString().Trim();

                        if (string.IsNullOrEmpty(idEmpresa) && string.IsNullOrEmpty(Region) && string.IsNullOrEmpty(Pais) && string.IsNullOrEmpty(Departamento) /*&& string.IsNullOrEmpty(cadenaValor) && 
                            string.IsNullOrEmpty(macroProceso) && string.IsNullOrEmpty(proceso) && string.IsNullOrEmpty(subproceso) && string.IsNullOrEmpty(ValorRecuperadoTotal) && string.IsNullOrEmpty(Moneda2) &&
                            string.IsNullOrEmpty(TasaCambio2) && string.IsNullOrEmpty(ValorPesos2) && string.IsNullOrEmpty(Recuperacion) && string.IsNullOrEmpty(FechaContabilidad) && string.IsNullOrEmpty(HoraContabilidad)*/)
                        {
                            continue;
                        }
                        else
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 3) + "</B> El valor del campo <B>Empresa</B> no puede estar vacío!");
                        }
                    }
                }

                //-----------------------------------------
                //Valida idJerarquia segun el nodo de la Jerarquia Organizacional
                //-----------------------------------------
                for (int i = 0; i < dtJerarquia.Rows.Count; i++)
                {
                    DataTable dtRpta = new DataTable();
                    if (dtJerarquia.Rows[i][0].ToString().Trim() != string.Empty)
                    {
                        string parametro = dtJerarquia.Rows[i][0].ToString().Trim();
                        dtRpta = CargaMasiva.DataExisteJerarquia(Convert.ToInt32(dtJerarquia.Rows[i][0]));
                        int cuantos = Convert.ToInt32(dtRpta.Rows.Count);
                        if (cuantos > 0)
                        {
                            continue;
                        }
                        else
                        {
                            throw new System.ArgumentException("No es posible procesar el archivo. <B> Número de Línea: </B>" + (i + 3) + ", el valor del campo id Cargo Responsable no se encuentra registrado. ", "<B> IdCargoResponsable, Valor del parámetro: " + parametro + " </B>");
                        }
                    }
                    else
                    {
                        throw new System.ArgumentException("No es posible procesar el archivo. <B> Número de Línea: </B>" + (i + 3) + ", no se permiten valores vacíos en el campo IdCargoResponsable del archivo. ", "<B> IdCargoResponsable </B>");
                    }
                }
                /*writer.Flush();
                writer.Close();
                excelReader.Close();*/

                //-----------------------------------------
                //Ejecuto query
                //-----------------------------------------

                /*cError ce = new cError();
                int resultado = 0;
                string error = string.Empty;
                try
                {
                    resultado = CargaMasiva.cargaEvento(nombreSp, nombreArchivo, path_planos);
                }
                catch (Exception ex)
                {
                    error = ex.ToString();
                    ce.errorMessage(ex.Message + ", " + ex.StackTrace);
                    throw new Exception(ex.Message);
                }*/

                //-----------------------------------------
                //Envío correo
                //-----------------------------------------

                if (/*resultado*/contador > 0)
                {
                    DataView view = new DataView(dtJerarquia);
                    view.Sort = "idJerarquia";
                    DataTable dtOrdenado = view.ToTable();
                    string[] registrosOrdenados = new string[dtOrdenado.Rows.Count * 2];//
                    int k = 0, p = 0, contEncontrados = 0;
                    string codigoEvento = string.Empty, descripcionEvento = string.Empty;

                    for (int i = 0; i < dtOrdenado.Rows.Count; i++)
                    {
                        string jerarquiaCapturada = dtOrdenado.Rows[i][0].ToString().Trim();

                        if (jerarquiaCapturada != registrosOrdenados[p])
                        {
                            for (int j = 0; j < dtOrdenado.Rows.Count; j++)
                            {
                                string jerarquiaEncontrada = dtOrdenado.Rows[j][0].ToString().Trim();
                                if (jerarquiaCapturada == jerarquiaEncontrada)
                                {
                                    string codEven = dtOrdenado.Rows[j][1].ToString().Trim();
                                    codigoEvento += dtOrdenado.Rows[j][1].ToString().Trim() + ", ";
                                    descripcionEvento += " <br /> <B> " + codEven + ": </B> " + dtOrdenado.Rows[j][2].ToString().Trim() + ". ";
                                    contEncontrados++;
                                }
                            }

                            if (contEncontrados > 0)
                            {
                                try
                                {
                                    if (dtOrdenado.Rows.Count == 1)
                                    {
                                        boolEnviarNotificacion(34, Convert.ToInt16("0"), Convert.ToInt32(jerarquiaCapturada), "",
                                        "<B>Se ha creado el siguiente Evento por carga masiva:</B> <br /><br /><B>Descripción de Evento: </B> <br />" + descripcionEvento.Trim());
                                        codigoEvento = string.Empty;
                                        contEncontrados = 0;
                                    }
                                    else
                                    {
                                        boolEnviarNotificacion(34, Convert.ToInt16("0"), Convert.ToInt32(jerarquiaCapturada), "",
                                        "<B>Se han creado los siguientes Eventos por carga masiva:</B> <br /><br /><B>Descripción de los Eventos: </B> <br />" + descripcionEvento.Trim() + " <B> <br/><br/>Total Eventos creados: </B> " + dtOrdenado.Rows.Count + "<br /> ");
                                        codigoEvento = string.Empty;
                                        contEncontrados = 0;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Mensaje("Error al enviar notificación de creacion de Evento." + ex.Message);
                                }
                            }

                            if (registrosOrdenados[k] == null)
                            {
                                registrosOrdenados[k] = jerarquiaCapturada;
                            }
                            for (int m = 0; m < dtOrdenado.Rows.Count; m++)
                            {
                                if (registrosOrdenados[k] == dtOrdenado.Rows[m][0].ToString().Trim())
                                {
                                    k++;
                                    p++;
                                    registrosOrdenados[p] = jerarquiaCapturada;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            p++;
                            registrosOrdenados[p] = jerarquiaCapturada;
                        }
                    }
                    omb.ShowMessage("Archivo procesado con éxito. Se cargaron : " + contador + " registros satisfactoriamente.", 3, "Atencion");
                    //info.Delete();
                }
                else
                {
                    Mensaje("Se presentó un problema mientras se realizaba la carga, por favor verifique los datos e intentelo nuevamente. " /*+ error*/);
                }
            }
        }
        private void LoadFilePlantillaEventosConDatosComplementarios()
        {
            #region variables
            //string path_planos = System.Configuration.ConfigurationManager.AppSettings["PATH_PLANOS"].ToString();
            //string nombreArchivo = Constantes.NOMBRE_CARGAS;
            //nombreArchivo = Procs.NombreArchivo(path_planos, nombreArchivo, ".txt");
            //string savedFileName = Path.Combine(path_planos, nombreArchivo);
            //string destino = System.IO.Path.Combine(path_planos, nombreArchivo);
            StreamWriter writer = null;
            StringBuilder s = null;
            FileInfo info = null;
            string strPrefixEvento = string.Empty, nombreSp = "[Eventos].[pa_RegistrarEventoCompleto]";
            int contador = 0;
            IExcelDataReader excelReader = null;
            string ultimoEE = string.Empty, ultimoEV = string.Empty, ultimoEG = string.Empty, DescipcionEvento = string.Empty, idResponsable = string.Empty;
            DataTable dtJerarquia = new DataTable();
            dtJerarquia.Columns.Add("idJerarquia"); dtJerarquia.Columns.Add("codigoEvento"); dtJerarquia.Columns.Add("descripcion_evento");
            cCargaMasivaRCE CargaMasiva = new cCargaMasivaRCE();
            System.Data.DataSet ds = new System.Data.DataSet();
            cambiaCultura cc = new cambiaCultura();
            #endregion variables
            try
            {
                //if (!System.IO.Directory.Exists(path_planos))
                //{
                //    System.IO.Directory.CreateDirectory(path_planos);
                //}

                //writer = File.AppendText(destino);
                //info = new FileInfo(destino);

                int IdUsuario = Convert.ToInt32(Session["IdUsuario"].ToString());
                string FechaRegistro = DateTime.Now.ToString();

                int excelFlag = 1;
                if (excelFlag == 1)
                {
                    excelReader = ExcelReaderFactory.CreateBinaryReader(FUloadExcel.PostedFile.InputStream);
                    excelReader.IsFirstRowAsColumnNames = true;
                }
                else if (excelFlag == 2)
                {
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(FUloadExcel.PostedFile.InputStream);
                }
            }
            catch (Exception ex)
            {
                Mensaje("Error al analizar el archivo Eventos. " + ex.Message);
            }
            if (excelReader != null)
            {
                ds = excelReader.AsDataSet();
                int IdUsuario = Convert.ToInt32(Session["IdUsuario"].ToString()), Linea = 0;
                DataTable datos = ds.Tables[0];
                DataTable dt = CargaMasiva.LastCodEvento();
                DataTable dts = new DataTable();
                Session["IdEvento"] = dt.Rows[0][0].ToString();
                string maxValor = string.Empty;
                DataTable dttemp = new DataTable();
                dttemp.Columns.Add("PrefijoEV"); dttemp.Columns.Add("PrefijoEG"); dttemp.Columns.Add("PrefijoEE");

                string[] prefijo = new string[3] { "EV", "EG", "EE" };
                DataRow drFila = dttemp.NewRow();
                for (int i = 0; i < prefijo.Length; i++)
                {
                    maxValor = prefijo[i];
                    dts = CargaMasiva.ultimoCodigoEvento(maxValor);

                    foreach (DataRow row in dts.Rows)
                    {
                        maxValor = row[0].ToString();
                        if (prefijo[i] == "EV")
                        {
                            drFila["PrefijoEV"] = "";
                            dttemp.Rows.Add(drFila);
                            dttemp.Rows[0][0] = maxValor;
                            break;
                        }
                        if (prefijo[i] == "EG")
                        {
                            dttemp.Rows[0][1] = maxValor;
                            break;
                        }
                        if (prefijo[i] == "EE")
                        {
                            dttemp.Rows[0][2] = maxValor;
                            break;
                        }
                    }
                }
                for (int i = 1; i < datos.Rows.Count; i++)
                {
                    if (datos.Rows[i][0].ToString().Trim() != "")
                    {
                        Linea++;
                        string idEmpresa = datos.Rows[i][0].ToString().Trim();
                        string Region = datos.Rows[i][1].ToString().Trim();
                        string Pais = datos.Rows[i][2].ToString().Trim();
                        string Departamento = datos.Rows[i][3].ToString().Trim();
                        string Ciudad = datos.Rows[i][4].ToString().Trim();
                        string Oficina = datos.Rows[i][5].ToString().Trim();
                        string Detalle_Ubicacion = datos.Rows[i][6].ToString().Trim();
                        string Descripcion_Evento = datos.Rows[i][7].ToString().Trim();
                        string IdServicio = datos.Rows[i][8].ToString().Trim();
                        string IdSubServicio = datos.Rows[i][9].ToString().Trim();
                        string FechaInicio = datos.Rows[i][10].ToString().Trim();
                        string HoraInicio = datos.Rows[i][11].ToString().Trim();
                        string FechaFinalizacion = datos.Rows[i][12].ToString().Trim();
                        string HoraFinalizacion = datos.Rows[i][13].ToString().Trim();
                        string FechaDescubrimiento = datos.Rows[i][14].ToString().Trim();
                        string HoraDescubrimiento = datos.Rows[i][15].ToString().Trim();
                        string Canal = datos.Rows[i][16].ToString().Trim();
                        string Generador = datos.Rows[i][17].ToString().Trim();
                        string IdCargoResponsable = datos.Rows[i][18].ToString().Trim();
                        string CargoResponsable = datos.Rows[i][19].ToString().Trim();
                        string CuantiaPerdida = datos.Rows[i][20].ToString().Trim();

                        string cadenaValor = datos.Rows[i][21].ToString().Trim();
                        string macroProceso = datos.Rows[i][22].ToString().Trim();
                        string proceso = datos.Rows[i][23].ToString().Trim();
                        string subproceso = datos.Rows[i][24].ToString().Trim();
                        string actividad = datos.Rows[i][25].ToString().Trim();
                        string IdClase = datos.Rows[i][26].ToString().Trim();
                        string IdSubClase = datos.Rows[i][27].ToString().Trim();
                        string IdTipoPerdidaEvento = datos.Rows[i][28].ToString().Trim();
                        string IdLineaProceso = datos.Rows[i][29].ToString().Trim();
                        string IdSubLineaProceso = datos.Rows[i][30].ToString().Trim();
                        string AfectaContinudad = datos.Rows[i][31].ToString().Trim();
                        string IdEstado = datos.Rows[i][32].ToString().Trim();
                        string Observaciones = datos.Rows[i][33].ToString().Trim();
                        string CuentaPUC = datos.Rows[i][34].ToString().Trim();
                        string CuentaOrden = datos.Rows[i][35].ToString().Trim();
                        string TasaCambio1 = datos.Rows[i][36].ToString().Trim();
                        string ValorPesos1 = datos.Rows[i][37].ToString().Trim();
                        string ValorRecuperadoTotal = datos.Rows[i][38].ToString().Trim();
                        string Moneda2 = datos.Rows[i][39].ToString().Trim();
                        string TasaCambio2 = datos.Rows[i][40].ToString().Trim();
                        string ValorPesos2 = datos.Rows[i][41].ToString().Trim();
                        string Recuperacion = datos.Rows[i][42].ToString().Trim();
                        string FechaContabilidad = datos.Rows[i][43].ToString().Trim();
                        string HoraContabilidad = datos.Rows[i][44].ToString().Trim();
                        string ImpactoCualitativo = datos.Rows[i][45].ToString().Trim();
                        //********CAMPOS NUEVOS********
                        string FechaRecuperacion = datos.Rows[i][46].ToString().Trim();
                        string HoraRecuperacion = datos.Rows[i][47].ToString().Trim();
                        string CuantiaRecup = datos.Rows[i][48].ToString().Trim();
                        string CuantiaOtraRecup = datos.Rows[i][49].ToString().Trim();
                        string CuantiaNeta = datos.Rows[i][50].ToString().Trim();
                        //********CAMPOS NUEVOS********
                        string fechaInicioCp = string.Empty;
                        string fechaFinalizacionCp = string.Empty;
                        string fechaDescubrimientoCp = string.Empty;
                        string fechaContabilidad = string.Empty;
                        string fechaRecuperacion = string.Empty;
                        string hoy = string.Empty;

                        //Controlador de fechas y horas - Esto porque el valor viene en  OLE Automation y no lo admite el struct
                        if (!string.IsNullOrEmpty(FechaInicio.TrimEnd()))
                        {
                            DateTime FechaInicioCm = DateTime.FromOADate(Convert.ToDouble(FechaInicio));
                            fechaInicioCp = string.Format("{0:yyyy-MM-dd}", Convert.ToDateTime(FechaInicioCm));
                        }
                        else
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Fecha de Inicio</B> no puede estar vacío!");
                        }

                        if (!string.IsNullOrEmpty(FechaFinalizacion))
                        {
                            DateTime FechaFinalizacionCm = DateTime.FromOADate(Convert.ToDouble(FechaFinalizacion));
                            fechaFinalizacionCp = string.Format("{0:yyyy-MM-dd}", Convert.ToDateTime(FechaFinalizacionCm));
                        }
                        else
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Fecha de Finalización</B> no puede estar vacío!");
                        }

                        if (!string.IsNullOrEmpty(FechaDescubrimiento))
                        {
                            DateTime FechaDescubrimientoCm = DateTime.FromOADate(Convert.ToDouble(FechaDescubrimiento));
                            fechaDescubrimientoCp = string.Format("{0:yyyy-MM-dd}", Convert.ToDateTime(FechaDescubrimientoCm));
                        }
                        else
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Fecha de Descubrimiento</B> no puede estar vacío!");
                        }

                        if (!string.IsNullOrEmpty(FechaContabilidad))
                        {
                            DateTime FechaContabilidadCm = DateTime.FromOADate(Convert.ToDouble(FechaContabilidad));
                            fechaContabilidad = string.Format("{0:yyyy-MM-dd}", Convert.ToDateTime(FechaContabilidadCm));
                        }
                        else if (IdTipoPerdidaEvento == "1")
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Fecha de Contabilidad</B> para Tipo de Pérdida = 1, no puede estar vacío!");
                        }
                        //Fecha Recuperación
                        if (!string.IsNullOrEmpty(FechaRecuperacion))
                        {
                            DateTime FechaRecuperacionCm = DateTime.FromOADate(Convert.ToDouble(FechaRecuperacion));
                            fechaRecuperacion = string.Format("{0:yyyy-MM-dd}", Convert.ToDateTime(FechaRecuperacionCm));
                        }
                        else if (IdTipoPerdidaEvento == "1")
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Fecha de Contabilidad</B> para Tipo de Pérdida = 1, no puede estar vacío!");
                        }

                        //Hora Recuperación
                        if (!string.IsNullOrEmpty(HoraRecuperacion))
                        {
                            DateTime HoraRecuperacionCm = DateTime.FromOADate(Convert.ToDouble(HoraRecuperacion));
                            string HoraRecuperacionCp = string.Format("{0:hh:mm tt}", Convert.ToDateTime(HoraRecuperacionCm));

                            //case format a.m/p.m
                            DateTime Fecha2 = DateTime.FromOADate(Convert.ToDouble(HoraRecuperacion));
                            string HoraFecha2 = Fecha2.ToString("hh:mm tt", CultureInfo.InvariantCulture);
                            HoraFecha2 = HoraFecha2.Replace("PM", "p.m").Replace("AM", "a.m");

                            HoraRecuperacion = HoraRecuperacionCp.TrimEnd('.');
                            HoraRecuperacion = HoraFecha2;
                        }

                        DateTime hoyCm = Convert.ToDateTime(DateTime.Now.ToShortDateString());
                        hoy = string.Format("{0:yyyy-MM-dd}", hoyCm);

                        //Hora inicio
                        if (!string.IsNullOrEmpty(HoraInicio))
                        {
                            DateTime HoraInicioCm = DateTime.FromOADate(Convert.ToDouble(HoraInicio));
                            string HoraInciCp = string.Format("{0:hh:mm tt}", Convert.ToDateTime(HoraInicioCm));
                            
                            //case format a.m/p.m
                            DateTime Fecha2 = DateTime.FromOADate(Convert.ToDouble(HoraInicio));
                            string HoraFecha2 = Fecha2.ToString("hh:mm tt", CultureInfo.InvariantCulture);
                            HoraFecha2 = HoraFecha2.Replace("PM", "p.m").Replace("AM", "a.m");

                            HoraInicio = HoraInciCp.TrimEnd('.');
                            HoraInicio = HoraFecha2;
                        }

                        //HoraFinalizacion
                        if (!string.IsNullOrEmpty(HoraFinalizacion))
                        {
                            DateTime HoraFinalCm = DateTime.FromOADate(Convert.ToDouble(HoraFinalizacion));
                            string HoraFinalCp = string.Format("{0:hh:mm tt}", Convert.ToDateTime(HoraFinalCm));

                            //case format a.m/p.m
                            DateTime Fecha2 = DateTime.FromOADate(Convert.ToDouble(HoraFinalizacion));
                            string HoraFecha2 = Fecha2.ToString("hh:mm tt", CultureInfo.InvariantCulture);
                            HoraFecha2 = HoraFecha2.Replace("PM", "p.m").Replace("AM", "a.m");

                            HoraFinalizacion = HoraFinalCp.TrimEnd('.');
                            HoraFinalizacion = HoraFecha2;
                        }

                        //HoraDescubrimiento
                        if (!string.IsNullOrEmpty(HoraDescubrimiento))
                        {
                            DateTime HoraDesCm = DateTime.FromOADate(Convert.ToDouble(HoraDescubrimiento));
                            string HoraDescCp = string.Format("{0:hh:mm tt}", Convert.ToDateTime(HoraDesCm));

                            //case format a.m/p.m
                            DateTime Fecha2 = DateTime.FromOADate(Convert.ToDouble(HoraDescubrimiento));
                            string HoraFecha2 = Fecha2.ToString("hh:mm tt", CultureInfo.InvariantCulture);
                            HoraFecha2 = HoraFecha2.Replace("PM", "p.m").Replace("AM", "a.m");

                            HoraDescubrimiento = HoraDescCp.TrimEnd('.');
                            HoraDescubrimiento = HoraFecha2;
                        }

                        //HoraContabilidad
                        if (!string.IsNullOrEmpty(HoraContabilidad))
                        {
                            DateTime HoraContCm = DateTime.FromOADate(Convert.ToDouble(HoraContabilidad));
                            string HoraContCp = string.Format("{0:hh:mm tt}", Convert.ToDateTime(HoraContCm));

                            //case format a.m/p.m
                            DateTime Fecha2 = DateTime.FromOADate(Convert.ToDouble(HoraContabilidad));
                            string HoraFecha2 = Fecha2.ToString("hh:mm tt", CultureInfo.InvariantCulture);
                            HoraFecha2 = HoraFecha2.Replace("PM", "p.m").Replace("AM", "a.m");

                            HoraContabilidad = HoraContCp.TrimEnd('.');
                            HoraContabilidad = HoraFecha2;
                        }
                        else if (IdTipoPerdidaEvento == "1")
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Hora Contabilidad</B> para Tipo de Pérdida = 1, no puede estar vacío!");
                        }

                        if ((IdTipoPerdidaEvento == "1") && (string.IsNullOrEmpty(CuentaPUC) || string.IsNullOrEmpty(CuentaOrden) || string.IsNullOrEmpty(TasaCambio1) || string.IsNullOrEmpty(ValorPesos1) || string.IsNullOrEmpty(ValorRecuperadoTotal) || string.IsNullOrEmpty(Moneda2) || string.IsNullOrEmpty(TasaCambio2) || string.IsNullOrEmpty(ValorPesos2) || string.IsNullOrEmpty(Recuperacion)))
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> Los <B>Datos de Contabilización</B> son obligatorios para Tipo de Pérdida = 1.");
                        }

                        DescipcionEvento = Descripcion_Evento;
                        idResponsable = IdCargoResponsable;

                        //if (!string.IsNullOrEmpty(Recuperacion))
                        //{
                        //    if (Convert.ToInt32(Recuperacion) > 1)
                        //    {
                        //        throw new Exception("Verificar la Línea: <B>" + (Linea + 2) + "</B> del archivo .El valor de la columna  <B>Recuperación,</B> no puede ser mayor a 1.");
                        //    }
                        //}
                        //else
                        //{
                        //    throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Recuperación</B> no puede estar vacío!");
                        //}

                        if (string.IsNullOrEmpty(idEmpresa.TrimEnd()) || string.IsNullOrEmpty(Region.TrimEnd()) || string.IsNullOrEmpty(Pais.TrimEnd()) || string.IsNullOrEmpty(Departamento.TrimEnd())
                            || string.IsNullOrEmpty(Ciudad.TrimEnd()) || string.IsNullOrEmpty(Oficina.TrimEnd()) || string.IsNullOrEmpty(IdServicio.TrimEnd()) || string.IsNullOrEmpty(IdSubServicio.TrimEnd())
                            || string.IsNullOrEmpty(Canal.TrimEnd()) || string.IsNullOrEmpty(CuantiaPerdida.TrimEnd()))
                        {
                            string Columna = string.Empty;
                            int ContadorVacios = 0;

                            if (string.IsNullOrEmpty(Region))
                            {
                                Columna = "Region";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(Pais))
                            {
                                Columna = Columna + ", Pais";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(Departamento))
                            {
                                Columna = Columna + ", Departamento";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(Ciudad))
                            {
                                Columna = Columna + ", Ciudad";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(Oficina))
                            {
                                Columna = Columna + ", Oficina";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(IdServicio))
                            {
                                Columna = Columna + ", Servicio/Producto";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(IdSubServicio))
                            {
                                Columna = Columna + ", SubServicio/SubProducto";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(Canal))
                            {
                                Columna = Columna + ", Canal";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(CuantiaPerdida))
                            {
                                Columna = Columna + ", Posible cuantía pérdida";
                                ContadorVacios++;
                            }

                            if (ContadorVacios > 1)
                            {
                                throw new Exception("Verificar la Línea: <B>" + (Linea + 2) + "</B> del archivo. El valor de las columnas  <B>" + Columna + "</B> no pueden estar vacías.");
                            }
                            else
                            {
                                Columna = Columna.TrimStart(',');

                                throw new Exception("Verificar la Línea: <B>" + (Linea + 2) + "</B> del archivo. El valor de la columna  <B>" + Columna + "</B> no puede estar vacía.");
                            }
                        }

                        //Incrementador de copdigoEvento
                        switch (idEmpresa.Trim())
                        {
                            case "1":
                                strPrefixEvento = "EV";
                                string sprint = string.Empty;
                                int maxLenght = 0;
                                int sub = 0;
                                if (ultimoEV == string.Empty)
                                {
                                    sprint = dttemp.Rows[0][0].ToString();
                                    maxLenght = sprint.Length;
                                    if (maxLenght == 2)
                                    {
                                        sub = 1;
                                    }
                                }
                                else
                                {
                                    sprint = ultimoEV;
                                    maxLenght = sprint.Length;
                                }
                                if (maxLenght > 2)
                                {
                                    sub = Convert.ToInt32(sprint.Substring(2));
                                    sub += 1;
                                }

                                strPrefixEvento = strPrefixEvento + sub.ToString();
                                ultimoEV = strPrefixEvento;
                                strPrefixEvento = ultimoEV;

                                break;
                            case "2":
                                strPrefixEvento = "EG";
                                string sprint2 = string.Empty;
                                int maxLenght2 = 0;
                                int sub2 = 0;
                                if (ultimoEG == string.Empty)
                                {
                                    sprint2 = dttemp.Rows[0][1].ToString();
                                    maxLenght2 = sprint2.Length;
                                    if (maxLenght2 == 2)
                                    {
                                        sub2 = 1;
                                    }
                                }
                                else
                                {
                                    sprint2 = ultimoEG;
                                    maxLenght2 = sprint2.Length;
                                }
                                if (maxLenght2 > 2)
                                {
                                    sub2 = Convert.ToInt32(sprint2.Substring(2));
                                    sub2 += 1;
                                }

                                strPrefixEvento = strPrefixEvento + sub2.ToString();
                                ultimoEG = strPrefixEvento;
                                strPrefixEvento = ultimoEG;

                                break;
                            case "3":
                                strPrefixEvento = "EE";
                                string sprint3 = string.Empty;
                                int maxLenght3 = 0;
                                int sub3 = 0;
                                if (ultimoEE == string.Empty)
                                {
                                    sprint3 = dttemp.Rows[0][2].ToString();
                                    maxLenght3 = sprint3.Length;
                                    if (maxLenght3 == 2)
                                    {
                                        sub3 = 1;
                                    }
                                }
                                else
                                {
                                    sprint3 = ultimoEE;
                                    maxLenght3 = sprint3.Length;
                                }
                                if (maxLenght3 > 2)
                                {
                                    sub3 = Convert.ToInt32(sprint3.Substring(2));
                                    sub3 += 1;
                                }

                                strPrefixEvento = strPrefixEvento + sub3.ToString();
                                ultimoEE = strPrefixEvento;
                                strPrefixEvento = ultimoEE;
                                break;
                        }

                        try
                        {
                           /* s = new StringBuilder();
                            s.Append(strPrefixEvento); s.Append(Constantes.TAB);
                            s.Append(idEmpresa); s.Append(Constantes.TAB);
                            s.Append(Region); s.Append(Constantes.TAB);
                            s.Append(Pais); s.Append(Constantes.TAB);
                            s.Append(Departamento); s.Append(Constantes.TAB);
                            s.Append(Ciudad); s.Append(Constantes.TAB);
                            s.Append(Oficina); s.Append(Constantes.TAB);
                            s.Append(Detalle_Ubicacion); s.Append(Constantes.TAB);
                            s.Append(Descripcion_Evento); s.Append(Constantes.TAB);
                            s.Append(IdServicio); s.Append(Constantes.TAB);
                            s.Append(IdSubServicio); s.Append(Constantes.TAB);
                            s.Append(fechaInicioCp); s.Append(Constantes.TAB);
                            s.Append(HoraInicio); s.Append(Constantes.TAB);
                            s.Append(fechaFinalizacionCp); s.Append(Constantes.TAB);
                            s.Append(HoraFinalizacion); s.Append(Constantes.TAB);
                            s.Append(fechaDescubrimientoCp); s.Append(Constantes.TAB);
                            s.Append(HoraDescubrimiento); s.Append(Constantes.TAB);
                            s.Append(Canal); s.Append(Constantes.TAB);
                            s.Append(Generador); s.Append(Constantes.TAB);
                            s.Append(IdCargoResponsable); s.Append(Constantes.TAB);
                            s.Append(CargoResponsable); s.Append(Constantes.TAB);
                            s.Append(CuantiaPerdida); s.Append(Constantes.TAB);
                            s.Append(cadenaValor); s.Append(Constantes.TAB);
                            s.Append(macroProceso); s.Append(Constantes.TAB);
                            s.Append(proceso); s.Append(Constantes.TAB);
                            s.Append(subproceso); s.Append(Constantes.TAB);
                            s.Append(actividad); s.Append(Constantes.TAB);
                            s.Append(IdClase); s.Append(Constantes.TAB);
                            s.Append(IdSubClase); s.Append(Constantes.TAB);
                            s.Append(IdTipoPerdidaEvento); s.Append(Constantes.TAB);
                            s.Append(IdLineaProceso); s.Append(Constantes.TAB);
                            s.Append(IdSubLineaProceso); s.Append(Constantes.TAB);
                            s.Append(AfectaContinudad); s.Append(Constantes.TAB);
                            s.Append(IdEstado); s.Append(Constantes.TAB);
                            s.Append(Observaciones); s.Append(Constantes.TAB);
                            s.Append(CuentaPUC); s.Append(Constantes.TAB);
                            s.Append(CuentaOrden); s.Append(Constantes.TAB);
                            s.Append(TasaCambio1); s.Append(Constantes.TAB);
                            s.Append(ValorPesos1); s.Append(Constantes.TAB);
                            s.Append(ValorRecuperadoTotal); s.Append(Constantes.TAB);
                            s.Append(Moneda2); s.Append(Constantes.TAB);
                            s.Append(TasaCambio2); s.Append(Constantes.TAB);
                            s.Append(ValorPesos2); s.Append(Constantes.TAB);
                            s.Append(Recuperacion); s.Append(Constantes.TAB);
                            s.Append(fechaContabilidad); s.Append(Constantes.TAB);
                            s.Append(HoraContabilidad); s.Append(Constantes.TAB);

                            s.Append(IdUsuario); s.Append(Constantes.TAB);
                            s.Append(hoy);*/
                            clsDTOEventos objEvento = new clsDTOEventos();
                            objEvento.strPrefixEvento = strPrefixEvento;
                            objEvento.intIdEmpresa = Convert.ToInt32(idEmpresa);
                            objEvento.intIdRegion = Convert.ToInt32(Region);
                            objEvento.intIdPais = Convert.ToInt32(Pais);
                            objEvento.intIdDepartamento = Convert.ToInt32(Departamento);
                            objEvento.intIdCiudad = Convert.ToInt32(Ciudad);
                            objEvento.intIdOficinaSucursal = Convert.ToInt32(Oficina);
                            objEvento.strDetalleUbicacion = Detalle_Ubicacion;
                            objEvento.strDescripcionEvento = Descripcion_Evento;
                            objEvento.intIdServicio = Convert.ToInt32(IdServicio);
                            objEvento.intIdSubServicio = Convert.ToInt32(IdSubServicio);
                            objEvento.dtFechaInicio = Convert.ToDateTime(fechaInicioCp);
                            objEvento.strHoraInicio = HoraInicio;
                            objEvento.dtFechaFinalizacion = Convert.ToDateTime(fechaFinalizacionCp);
                            objEvento.strHoraFinalizacion = HoraFinalizacion;
                            objEvento.dtFechaDescubrimiento = Convert.ToDateTime(fechaDescubrimientoCp);
                            objEvento.strHoraDescubrimiento = HoraDescubrimiento;
                            objEvento.intIdCanal = Convert.ToInt32(Canal);
                            objEvento.intIdGeneraEvento = Convert.ToInt32(Generador);
                            objEvento.intGeneraEvento = Convert.ToInt32(idResponsable);
                            objEvento.strNomGeneradorEvento = CargoResponsable;
                            objEvento.strCuantiaperdida = CuantiaPerdida;
                            objEvento.intIdCadenaValor = Convert.ToInt32(cadenaValor);
                            objEvento.intIdMacroProceso = Convert.ToInt32(macroProceso);
                            objEvento.intIdProceso = Convert.ToInt32(proceso);
                            objEvento.intIdSubproceso = Convert.ToInt32(subproceso);
                            objEvento.intIdActividad = actividad;
                            objEvento.intIdClase = Convert.ToInt32(IdClase);
                            objEvento.intIdSubClase = Convert.ToInt32(IdSubClase);
                            objEvento.intIdTipoPerdidaEvento = Convert.ToInt32(IdTipoPerdidaEvento);
                            objEvento.intIdLineaProceso = Convert.ToInt32(IdLineaProceso);
                            objEvento.intIdSubLineaProceso = Convert.ToInt32(IdSubLineaProceso);
                            objEvento.intAfectaContinudad = Convert.ToInt32(AfectaContinudad);
                            objEvento.intIdEstado = Convert.ToInt32(IdEstado);
                            objEvento.strObservaciones = Observaciones;
                            objEvento.strCuentaPUC = CuentaPUC;
                            objEvento.strCuentaOrden = CuentaOrden;
                            objEvento.strTasaCambio1 = TasaCambio1;
                            objEvento.strValorPesos1 = ValorPesos1;
                            objEvento.strValorRecuperadoTotal = ValorRecuperadoTotal;
                            objEvento.strMoneda2 = Moneda2;
                            objEvento.strTasaCambio2 = TasaCambio2;
                            objEvento.strValorPesos2 = ValorPesos2;
                            objEvento.intRecuperacion = Recuperacion;
                            if (fechaContabilidad != "")
                                objEvento.dtFechaContabilidad = Convert.ToDateTime(fechaContabilidad);
                            else
                                objEvento.dtFechaContabilidad = null;
                            objEvento.strHoraContabilidad = HoraContabilidad;
                            objEvento.strImpactoCualitativo = ImpactoCualitativo;

                            //******CAMPOS NUEVOS******
                            if (FechaRecuperacion != "")
                                objEvento.dtFechaRecuperacion = Convert.ToDateTime(fechaRecuperacion);
                            else
                                objEvento.dtFechaRecuperacion = null;
                            objEvento.strHoraRecuperacion = HoraRecuperacion;
                            objEvento.strCuantiaRecup = CuantiaRecup;
                            objEvento.strCuantiaOtraRecup = CuantiaOtraRecup;
                            objEvento.strCuantiaNeta = CuantiaNeta;
                            //******CAMPOS NUEVOS******

                            objEvento.intIdUsuario = IdUsuario;
                            objEvento.dtfechaEvento = Convert.ToDateTime(hoy);
                            //writer.WriteLine(s.ToString());

                            DataRow registros = dtJerarquia.NewRow();
                            registros["idJerarquia"] = datos.Rows[i][18].ToString().Trim();
                            registros["codigoEvento"] = strPrefixEvento;
                            registros["descripcion_evento"] = Descripcion_Evento;
                            dtJerarquia.Rows.Add(registros);

                            cError ce = new cError();
                            int resultado = 0;
                            string error = string.Empty;
                            try
                            {
                                CargaMasiva.InsertEventoCompleto(objEvento);
                            }
                            catch (Exception ex)
                            {
                                error = ex.ToString();
                                ce.errorMessage(ex.Message + ", " + ex.StackTrace);
                                throw new Exception(ex.Message);
                            }

                            contador++;
                        }
                        catch (Exception ex)
                        {
                            Mensaje("Error al registrar el evento de la linea: " + i + "." + ex.Message);
                        }
                    }
                    else
                    {
                        string idEmpresa = datos.Rows[i][0].ToString().Trim();
                        string Region = datos.Rows[i][1].ToString().Trim();
                        string Pais = datos.Rows[i][2].ToString().Trim();
                        string Departamento = datos.Rows[i][3].ToString().Trim();
                        string cadenaValor = datos.Rows[i][21].ToString().Trim();
                        string macroProceso = datos.Rows[i][22].ToString().Trim();
                        string proceso = datos.Rows[i][23].ToString().Trim();
                        string subproceso = datos.Rows[i][24].ToString().Trim();
                        string ValorRecuperadoTotal = datos.Rows[i][38].ToString().Trim();
                        string Moneda2 = datos.Rows[i][39].ToString().Trim();
                        string TasaCambio2 = datos.Rows[i][40].ToString().Trim();
                        string ValorPesos2 = datos.Rows[i][41].ToString().Trim();
                        string Recuperacion = datos.Rows[i][42].ToString().Trim();
                        string FechaContabilidad = datos.Rows[i][43].ToString().Trim();
                        string HoraContabilidad = datos.Rows[i][44].ToString().Trim();

                        if (string.IsNullOrEmpty(idEmpresa) && string.IsNullOrEmpty(Region) && string.IsNullOrEmpty(Pais) && string.IsNullOrEmpty(Departamento) && string.IsNullOrEmpty(cadenaValor) &&
                            string.IsNullOrEmpty(macroProceso) && string.IsNullOrEmpty(proceso) && string.IsNullOrEmpty(subproceso) && string.IsNullOrEmpty(ValorRecuperadoTotal) && string.IsNullOrEmpty(Moneda2) &&
                            string.IsNullOrEmpty(TasaCambio2) && string.IsNullOrEmpty(ValorPesos2) && string.IsNullOrEmpty(Recuperacion) && string.IsNullOrEmpty(FechaContabilidad) && string.IsNullOrEmpty(HoraContabilidad))
                        {
                            continue;
                        }
                        else
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 3) + "</B> El valor del campo <B>Empresa</B> no puede estar vacío!");
                        }
                    }
                }

                //-----------------------------------------
                //Valida idJerarquia segun el nodo de la Jerarquia Organizacional
                //-----------------------------------------
                for (int i = 0; i < dtJerarquia.Rows.Count; i++)
                {
                    DataTable dtRpta = new DataTable();
                    if (dtJerarquia.Rows[i][0].ToString().Trim() != string.Empty)
                    {
                        string parametro = dtJerarquia.Rows[i][0].ToString().Trim();
                        dtRpta = CargaMasiva.DataExisteJerarquia(Convert.ToInt32(dtJerarquia.Rows[i][0]));
                        int cuantos = Convert.ToInt32(dtRpta.Rows.Count);
                        if (cuantos > 0)
                        {
                            continue;
                        }
                        else
                        {
                            throw new System.ArgumentException("No es posible procesar el archivo. <B> Número de Línea: </B>" + (i + 3) + ", el valor del campo id Cargo Responsable no se encuentra registrado. ", "<B> IdCargoResponsable, Valor del parámetro: " + parametro + " </B>");
                        }
                    }
                    else
                    {
                        throw new System.ArgumentException("No es posible procesar el archivo. <B> Número de Línea: </B>" + (i + 3) + ", no se permiten valores vacíos en el campo IdCargoResponsable del archivo. ", "<B> IdCargoResponsable </B>");
                    }
                }
                /*writer.Flush();
                writer.Close();
                excelReader.Close();*/

                //-----------------------------------------
                //Ejecuto query
                //-----------------------------------------

                /*cError ce = new cError();
                int resultado = 0;
                string error = string.Empty;
                try
                {
                    resultado = CargaMasiva.cargaEvento(nombreSp, nombreArchivo, path_planos);
                }
                catch (Exception ex)
                {
                    error = ex.ToString();
                    ce.errorMessage(ex.Message + ", " + ex.StackTrace);
                    throw new Exception(ex.Message);
                }*/
                
                //-----------------------------------------
                //Envío correo
                //-----------------------------------------

                if (/*resultado*/contador > 0)
                {
                    DataView view = new DataView(dtJerarquia);
                    view.Sort = "idJerarquia";
                    DataTable dtOrdenado = view.ToTable();
                    string[] registrosOrdenados = new string[dtOrdenado.Rows.Count * 2];
                    int k = 0, p = 0, contEncontrados = 0;
                    string codigoEvento = string.Empty, descripcionEvento = string.Empty;

                    for (int i = 0; i < dtOrdenado.Rows.Count; i++)
                    {
                        string jerarquiaCapturada = dtOrdenado.Rows[i][0].ToString().Trim();

                        if (jerarquiaCapturada != registrosOrdenados[p])
                        {
                            for (int j = 0; j < dtOrdenado.Rows.Count; j++)
                            {
                                string jerarquiaEncontrada = dtOrdenado.Rows[j][0].ToString().Trim();
                                if (jerarquiaCapturada == jerarquiaEncontrada)
                                {
                                    string codEven = dtOrdenado.Rows[j][1].ToString().Trim();
                                    codigoEvento += dtOrdenado.Rows[j][1].ToString().Trim() + ", ";
                                    descripcionEvento += " <br /> <B> " + codEven + ": </B> " + dtOrdenado.Rows[j][2].ToString().Trim() + ". ";
                                    contEncontrados++;
                                }
                            }

                            if (contEncontrados > 0)
                            {
                                try
                                {
                                    if (dtOrdenado.Rows.Count == 1)
                                    {
                                        boolEnviarNotificacion(17, Convert.ToInt16("0"), Convert.ToInt32(jerarquiaCapturada), "",
                                        "<B>Se ha creado el siguiente Evento por carga masiva:</B> <br /><br /><B>Descripción de Evento: </B> <br />" + descripcionEvento.Trim());
                                        codigoEvento = string.Empty;
                                        contEncontrados = 0;
                                    }
                                    else
                                    {
                                        boolEnviarNotificacion(17, Convert.ToInt16("0"), Convert.ToInt32(jerarquiaCapturada), "",
                                        "<B>Se han creado los siguientes Eventos por carga masiva:</B> <br /><br /><B>Descripción de los Eventos: </B> <br />" + descripcionEvento.Trim() + " <B> <br/><br/>Total Eventos creados: </B> " + contEncontrados + "<br /> ");
                                        codigoEvento = string.Empty;
                                        contEncontrados = 0;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Mensaje("Error al enviar notificación de creacion de Evento." + ex.Message);
                                }
                            }

                            if (registrosOrdenados[k] == null)
                            {
                                registrosOrdenados[k] = jerarquiaCapturada;
                            }
                            for (int m = 0; m < dtOrdenado.Rows.Count; m++)
                            {
                                if (registrosOrdenados[k] == dtOrdenado.Rows[m][0].ToString().Trim())
                                {
                                    k++;
                                    p++;
                                    registrosOrdenados[p] = jerarquiaCapturada;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            p++;
                            registrosOrdenados[p] = jerarquiaCapturada;
                        }
                    }
                    omb.ShowMessage("Archivo procesado con éxito. Se cargaron : " + contador + " registros satisfactoriamente.", 3, "Atencion");
                    //info.Delete();
                }
                else
                {
                    Mensaje("Se presentó un problema mientras se realizaba la carga, por favor verifique los datos e intentelo nuevamente. " /*+ error*/);
                }
            }
        }
        private void LoadFilePlantillaEventosDatosComplementarios()
        {
            #region variables
            //string path_planos = System.Configuration.ConfigurationManager.AppSettings["PATH_PLANOS"].ToString();
            string nombreArchivo = Constantes.NOMBRE_CARGAS;
            //nombreArchivo = Procs.NombreArchivo(path_planos, nombreArchivo, ".txt");
            //string savedFileName = Path.Combine(path_planos, nombreArchivo);
            //string destino = System.IO.Path.Combine(path_planos, nombreArchivo);
            StreamWriter writer = null;
            StringBuilder s = null;
            FileInfo info = null;
            //string  nombreSp = "[Eventos].[pa_RegistrarEventoDatosComplementarios]";
            int contador = 0;
            IExcelDataReader excelReader = null;
            //string ultimoEE = string.Empty, ultimoEV = string.Empty, ultimoEG = string.Empty, DescipcionEvento = string.Empty, idResponsable = string.Empty;
            DataTable dtJerarquia = new DataTable();
            dtJerarquia.Columns.Add("idJerarquia"); dtJerarquia.Columns.Add("codigoEvento"); dtJerarquia.Columns.Add("descripcion_evento");
            cCargaMasivaRCE CargaMasiva = new cCargaMasivaRCE();
            System.Data.DataSet ds = new System.Data.DataSet();
            cambiaCultura cc = new cambiaCultura();
            #endregion variables
            try
            {
                //if (!System.IO.Directory.Exists(path_planos))
                //{
                //    System.IO.Directory.CreateDirectory(path_planos);
                //}

                //writer = File.AppendText(destino);
                //info = new FileInfo(destino);

                int IdUsuario = Convert.ToInt32(Session["IdUsuario"].ToString());
                string FechaRegistro = DateTime.Now.ToString();

                int excelFlag = 1;
                if (excelFlag == 1)
                {
                    excelReader = ExcelReaderFactory.CreateBinaryReader(FUloadExcel.PostedFile.InputStream);
                    excelReader.IsFirstRowAsColumnNames = true;
                }
                else if (excelFlag == 2)
                {
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(FUloadExcel.PostedFile.InputStream);
                }
            }
            catch (Exception ex)
            {
                Mensaje("Error al analizar el archivo Eventos. " + ex.Message);
            }
            if (excelReader != null)
            {
                ds = excelReader.AsDataSet();
                int IdUsuario = Convert.ToInt32(Session["IdUsuario"].ToString()), Linea = 0;
                DataTable datos = ds.Tables[0];
                //DataTable dt = CargaMasiva.LastCodEvento();
                DataTable dts = new DataTable();
                //Session["IdEvento"] = dt.Rows[0][0].ToString();
                string maxValor = string.Empty;
                DataTable dttemp = new DataTable();
                dttemp.Columns.Add("PrefijoEV"); dttemp.Columns.Add("PrefijoEG"); dttemp.Columns.Add("PrefijoEE");

                string[] prefijo = new string[3] { "EV", "EG", "EE" };
                DataRow drFila = dttemp.NewRow();
                /*for (int i = 0; i < prefijo.Length; i++)
                {
                    maxValor = prefijo[i];
                    dts = CargaMasiva.ultimoCodigoEvento(maxValor);

                    foreach (DataRow row in dts.Rows)
                    {
                        maxValor = row[0].ToString();
                        if (prefijo[i] == "EV")
                        {
                            drFila["PrefijoEV"] = "";
                            dttemp.Rows.Add(drFila);
                            dttemp.Rows[0][0] = maxValor;
                            break;
                        }
                        if (prefijo[i] == "EG")
                        {
                            dttemp.Rows[0][1] = maxValor;
                            break;
                        }
                        if (prefijo[i] == "EE")
                        {
                            dttemp.Rows[0][2] = maxValor;
                            break;
                        }
                    }
                }*/
                string eventos = string.Empty;
                for (int i = 1; i < datos.Rows.Count; i++)
                {
                    if (datos.Rows[i][0].ToString().Trim() != "")
                    {
                        Linea++;
                        /**************** Se eliminan de los datos complementarios ***************
                        string idEmpresa = datos.Rows[i][0].ToString().Trim();
                        string Region = datos.Rows[i][1].ToString().Trim();
                        string Pais = datos.Rows[i][2].ToString().Trim();
                        string Departamento = datos.Rows[i][3].ToString().Trim();
                        string Ciudad = datos.Rows[i][4].ToString().Trim();
                        string Oficina = datos.Rows[i][5].ToString().Trim();
                        string Detalle_Ubicacion = datos.Rows[i][6].ToString().Trim();
                        string Descripcion_Evento = datos.Rows[i][7].ToString().Trim();
                        string IdServicio = datos.Rows[i][8].ToString().Trim();
                        string IdSubServicio = datos.Rows[i][9].ToString().Trim();
                        string FechaInicio = datos.Rows[i][10].ToString().Trim();
                        string HoraInicio = datos.Rows[i][11].ToString().Trim();
                        string FechaFinalizacion = datos.Rows[i][12].ToString().Trim();
                        string HoraFinalizacion = datos.Rows[i][13].ToString().Trim();
                        string FechaDescubrimiento = datos.Rows[i][14].ToString().Trim();
                        string HoraDescubrimiento = datos.Rows[i][15].ToString().Trim();
                        string Canal = datos.Rows[i][16].ToString().Trim();
                        string Generador = datos.Rows[i][17].ToString().Trim();
                        string IdCargoResponsable = datos.Rows[i][18].ToString().Trim();
                        string CargoResponsable = datos.Rows[i][19].ToString().Trim();
                        string CuantiaPerdida = datos.Rows[i][20].ToString().Trim();
                        */
                        /*if (i == (datos.Rows.Count - 1))
                        {
                            eventos += datos.Rows[i][0].ToString().Trim() + ")";
                            Session["IdsEventos"] = eventos;
                        }
                        else
                            eventos += datos.Rows[i][0].ToString().Trim() + ",";*/
                        eventos += datos.Rows[i][0].ToString().Trim() + ",";
                        Session["IdsEventos"] = eventos;
                        string idEvento = datos.Rows[i][0].ToString().Trim();
                        /*if(i == 1)
                            Session["IdEvento"] = idEvento;*/
                        string cadenaValor = datos.Rows[i][1].ToString().Trim();
                        string macroProceso = datos.Rows[i][2].ToString().Trim();
                        string proceso = datos.Rows[i][3].ToString().Trim();
                        string subproceso = datos.Rows[i][4].ToString().Trim();
                        string actividad = datos.Rows[i][5].ToString().Trim();
                        string IdClase = datos.Rows[i][6].ToString().Trim();
                        string IdSubClase = datos.Rows[i][7].ToString().Trim();
                        string IdTipoPerdidaEvento = datos.Rows[i][8].ToString().Trim();
                        string IdLineaProceso = datos.Rows[i][9].ToString().Trim();
                        string IdSubLineaProceso = datos.Rows[i][10].ToString().Trim();
                        string AfectaContinudad = datos.Rows[i][11].ToString().Trim();
                        string IdEstado = datos.Rows[i][12].ToString().Trim();
                        string Observaciones = datos.Rows[i][13].ToString().Trim();
                        string CuentaPUC = datos.Rows[i][14].ToString().Trim();
                        string CuentaOrden = datos.Rows[i][15].ToString().Trim();
                        string TasaCambio1 = datos.Rows[i][16].ToString().Trim();
                        string ValorPesos1 = datos.Rows[i][17].ToString().Trim();
                        string ValorRecuperadoTotal = datos.Rows[i][18].ToString().Trim();
                        string Moneda2 = datos.Rows[i][19].ToString().Trim();
                        string TasaCambio2 = datos.Rows[i][20].ToString().Trim();
                        string ValorPesos2 = datos.Rows[i][21].ToString().Trim();
                        string Recuperacion = datos.Rows[i][22].ToString().Trim();
                        string FechaContabilidad = datos.Rows[i][23].ToString().Trim();
                        string HoraContabilidad = datos.Rows[i][24].ToString().Trim();
                        /***********************************Campos Nuevos **************************/
                        string FechaRecuperacion = datos.Rows[i][25].ToString().Trim();
                        string HoraRecuperacion = datos.Rows[i][26].ToString().Trim();
                        string CuantiaRecup = datos.Rows[i][27].ToString().Trim();
                        string CuantiaOtraRecup = datos.Rows[i][28].ToString().Trim();
                        string CuantiaNeta = datos.Rows[i][29].ToString().Trim();
                        /*
                        string TasaCambio2 = datos.Rows[i][30].ToString().Trim();*/
                        /***********************************Campos Nuevos **************************/
                        string fechaInicioCp = string.Empty;
                        string fechaFinalizacionCp = string.Empty;
                        string fechaDescubrimientoCp = string.Empty;
                        string fechaContabilidad = string.Empty;
                        string hoy = string.Empty;

                        //Controlador de fechas y horas - Esto porque el valor viene en  OLE Automation y no lo admite el struct

                        if (!string.IsNullOrEmpty(FechaContabilidad))
                        {
                            DateTime FechaContabilidadCm = DateTime.FromOADate(Convert.ToDouble(FechaContabilidad));
                            fechaContabilidad = string.Format("{0:yyyy-MM-dd}", Convert.ToDateTime(FechaContabilidadCm));
                        }
                        else if (IdTipoPerdidaEvento == "1")
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Fecha de Contabilidad</B> para Tipo de Pérdida = 1, no puede estar vacío!");
                        }

                        if (!string.IsNullOrEmpty(FechaRecuperacion))
                        {
                            DateTime FechaRecuperacionCm = DateTime.FromOADate(Convert.ToDouble(FechaRecuperacion));
                            FechaRecuperacion = string.Format("{0:yyyy-MM-dd}", Convert.ToDateTime(FechaRecuperacionCm));
                        }
                        else if (IdTipoPerdidaEvento == "1")
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Fecha de Contabilidad</B> para Tipo de Pérdida = 1, no puede estar vacío!");
                        }
                        //else
                        //{
                        //    throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Fecha de Contabilidad</B> no puede estar vacío!");
                        //}

                        DateTime hoyCm = Convert.ToDateTime(DateTime.Now.ToShortDateString());
                        hoy = string.Format("{0:yyyy-MM-dd}", hoyCm);

                        if (!string.IsNullOrEmpty(HoraContabilidad))
                        {
                            //HoraContabilidad
                            DateTime HoraContCm = DateTime.FromOADate(Convert.ToDouble(HoraContabilidad));
                            string HoraContCp = string.Format("{0:hh:mm tt}", Convert.ToDateTime(HoraContCm));
                            
                            //case format a.m/p.m
                            DateTime Fecha2 = DateTime.FromOADate(Convert.ToDouble(HoraContabilidad));
                            string HoraFecha2 = Fecha2.ToString("hh:mm tt", CultureInfo.InvariantCulture);
                            HoraFecha2 = HoraFecha2.Replace("PM", "p.m").Replace("AM", "a.m");

                            HoraContabilidad = HoraContCp.TrimEnd('.');
                            HoraContabilidad = HoraFecha2;
                        }
                        else if (IdTipoPerdidaEvento == "1")
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Hora Contabilidad</B> para Tipo de Pérdida = 1, no puede estar vacío!");
                        }
                        
                        if ((IdTipoPerdidaEvento == "1")  && (string.IsNullOrEmpty(CuentaPUC) || string.IsNullOrEmpty(CuentaOrden) || string.IsNullOrEmpty(TasaCambio1) || string.IsNullOrEmpty(ValorPesos1) || string.IsNullOrEmpty(ValorRecuperadoTotal) || string.IsNullOrEmpty(Moneda2) || string.IsNullOrEmpty(TasaCambio2) || string.IsNullOrEmpty(ValorPesos2) || string.IsNullOrEmpty(Recuperacion)))
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> Los <B>Datos de Contabilización</B> son obligatorios para Tipo de Pérdida = 1.");
                        }

                        if (!string.IsNullOrEmpty(HoraRecuperacion))
                        {
                            //HoraContabilidad
                            DateTime HoraRecuperacionCm = DateTime.FromOADate(Convert.ToDouble(HoraRecuperacion));
                            string HoraContCp = string.Format("{0:hh:mm tt}", Convert.ToDateTime(HoraRecuperacionCm));

                            //case format a.m/p.m
                            DateTime Fecha2 = DateTime.FromOADate(Convert.ToDouble(HoraRecuperacion));
                            string HoraFecha2 = Fecha2.ToString("hh:mm tt", CultureInfo.InvariantCulture);
                            HoraFecha2 = HoraFecha2.Replace("PM", "p.m").Replace("AM", "a.m");

                            HoraRecuperacion = HoraContCp.TrimEnd('.');
                            HoraRecuperacion = HoraFecha2;
                        }
                        //if (!string.IsNullOrEmpty(Recuperacion))
                        //{
                        //    if (Convert.ToInt32(Recuperacion) > 1)
                        //    {
                        //        throw new Exception("Verificar la Línea: <B>" + (Linea + 2) + "</B> del archivo .El valor de la columna  <B>Recuperación,</B> no puede ser mayor a 1.");
                        //    }
                        //}
                        //else
                        //{
                        //    throw new Exception("Verificar la línea <B>" + (Linea + 2) + "</B> El valor del campo <B>Recuperación</B> no puede estar vacío!");
                        //}

                        if (string.IsNullOrEmpty(idEvento.TrimEnd()) || string.IsNullOrEmpty(cadenaValor.TrimEnd()) || string.IsNullOrEmpty(macroProceso.TrimEnd()) || string.IsNullOrEmpty(proceso.TrimEnd())
                            || string.IsNullOrEmpty(IdClase.TrimEnd()) || string.IsNullOrEmpty(IdSubClase.TrimEnd()) || string.IsNullOrEmpty(IdTipoPerdidaEvento.TrimEnd())
                            || string.IsNullOrEmpty(IdLineaProceso.TrimEnd()) || string.IsNullOrEmpty(IdSubLineaProceso.TrimEnd()))
                        {
                            string Columna = string.Empty;
                            int ContadorVacios = 0;

                            if (string.IsNullOrEmpty(idEvento))
                            {
                                Columna = "IdEvento";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(cadenaValor))
                            {
                                Columna = Columna + ", CadenaValor";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(macroProceso))
                            {
                                Columna = Columna + ", Macroproceso";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(proceso))
                            {
                                Columna = Columna + ", Proceso";
                                ContadorVacios++;
                            }
                            /*if (string.IsNullOrEmpty(actividad))
                            {
                                Columna = Columna + ", actividad";
                                ContadorVacios++;
                            }*/
                            if (string.IsNullOrEmpty(IdClase))
                            {
                                Columna = Columna + ", Clase";
                                ContadorVacios++;
                            }
                            if (string.IsNullOrEmpty(IdSubClase))
                            {
                                Columna = Columna + ", Subclase";
                                ContadorVacios++;
                            }


                            if (ContadorVacios > 1)
                            {
                                throw new Exception("Verificar la Línea: <B>" + (Linea + 2) + "</B> del archivo. El valor de las columnas  <B>" + Columna + "</B> no pueden estar vacías.");
                            }
                            else
                            {
                                Columna = Columna.TrimStart(',');

                                throw new Exception("Verificar la Línea: <B>" + (Linea + 2) + "</B> del archivo. El valor de la columna  <B>" + Columna + "</B> no puede estar vacía.");
                            }
                        }

                        try
                        {
                            cEvento dbEvento = new cEvento();
                            dbEvento.ModificaEventoMasivo(cadenaValor, macroProceso, proceso, subproceso,
                            actividad, IdClase, IdSubClase, IdTipoPerdidaEvento, IdLineaProceso, IdSubLineaProceso, AfectaContinudad, IdEstado, Observaciones,
                            CuentaPUC, CuentaOrden, TasaCambio1, ValorPesos1, ValorRecuperadoTotal, Moneda2, TasaCambio2, ValorPesos2, Recuperacion, fechaContabilidad,
                            HoraContabilidad, idEvento,FechaRecuperacion,HoraRecuperacion,CuantiaRecup,CuantiaOtraRecup,CuantiaNeta);

                            string idJerarquia = dbEvento.mtdGetIdJerarquia(idEvento);
                            string CodigoEvento = dbEvento.mtdGetCodEvento(idEvento);
                            string DescripcionEvento = dbEvento.mtdGetDesEvento(idEvento);
                            DataRow registros = dtJerarquia.NewRow();
                            registros["idJerarquia"] = idJerarquia;
                            registros["codigoEvento"] = CodigoEvento;
                            registros["descripcion_evento"] = DescripcionEvento;
                            dtJerarquia.Rows.Add(registros);

                            /*s = new StringBuilder();
                            s.Append(idEvento); s.Append(Constantes.TAB);
                            s.Append(cadenaValor); s.Append(Constantes.TAB);
                            s.Append(macroProceso); s.Append(Constantes.TAB);
                            s.Append(proceso); s.Append(Constantes.TAB);
                            s.Append(subproceso); s.Append(Constantes.TAB);
                            s.Append(actividad); s.Append(Constantes.TAB);
                            s.Append(IdClase); s.Append(Constantes.TAB);
                            s.Append(IdSubClase); s.Append(Constantes.TAB);
                            s.Append(IdTipoPerdidaEvento); s.Append(Constantes.TAB);
                            s.Append(IdLineaProceso); s.Append(Constantes.TAB);
                            s.Append(IdSubLineaProceso); s.Append(Constantes.TAB);
                            s.Append(AfectaContinudad); s.Append(Constantes.TAB);
                            s.Append(IdEstado); s.Append(Constantes.TAB);
                            s.Append(Observaciones); s.Append(Constantes.TAB);
                            s.Append(CuentaPUC); s.Append(Constantes.TAB);
                            s.Append(CuentaOrden); s.Append(Constantes.TAB);
                            s.Append(TasaCambio1); s.Append(Constantes.TAB);
                            s.Append(ValorPesos1); s.Append(Constantes.TAB);
                            s.Append(ValorRecuperadoTotal); s.Append(Constantes.TAB);
                            s.Append(Moneda2); s.Append(Constantes.TAB);
                            s.Append(TasaCambio2); s.Append(Constantes.TAB);
                            s.Append(ValorPesos2); s.Append(Constantes.TAB);
                            s.Append(Recuperacion); s.Append(Constantes.TAB);
                            s.Append(FechaContabilidad); s.Append(Constantes.TAB);
                            s.Append(HoraContabilidad); s.Append(Constantes.TAB);

                            s.Append(IdUsuario); s.Append(Constantes.TAB);
                            s.Append(hoy);*/
                            /************* Se elimina de la creacion inicial **********
                            
                            s.Append(cadenaValor); s.Append(Constantes.TAB);
                            s.Append(macroProceso); s.Append(Constantes.TAB);
                            s.Append(proceso); s.Append(Constantes.TAB);
                            s.Append(subproceso); s.Append(Constantes.TAB);
                            s.Append(actividad); s.Append(Constantes.TAB);
                            s.Append(IdClase); s.Append(Constantes.TAB);
                            s.Append(IdSubClase); s.Append(Constantes.TAB);
                            s.Append(IdTipoPerdidaEvento); s.Append(Constantes.TAB);
                            s.Append(IdLineaProceso); s.Append(Constantes.TAB);
                            s.Append(IdSubLineaProceso); s.Append(Constantes.TAB);
                            s.Append(AfectaContinudad); s.Append(Constantes.TAB);
                            s.Append(IdEstado); s.Append(Constantes.TAB);
                            s.Append(Observaciones); s.Append(Constantes.TAB);
                            s.Append(CuentaPUC); s.Append(Constantes.TAB);
                            s.Append(CuentaOrden); s.Append(Constantes.TAB);
                            s.Append(TasaCambio1); s.Append(Constantes.TAB);
                            s.Append(ValorPesos1); s.Append(Constantes.TAB);
                            s.Append(ValorRecuperadoTotal); s.Append(Constantes.TAB);
                            s.Append(Moneda2); s.Append(Constantes.TAB);
                            s.Append(TasaCambio2); s.Append(Constantes.TAB);
                            s.Append(ValorPesos2); s.Append(Constantes.TAB);
                            s.Append(Recuperacion); s.Append(Constantes.TAB);
                            s.Append(fechaContabilidad); s.Append(Constantes.TAB);
                            s.Append(HoraContabilidad); s.Append(Constantes.TAB);
                            
                            */
                            //writer.WriteLine(s.ToString());

                            contador++;
                        }
                        catch (Exception ex)
                        {
                            Mensaje("Error al registrar el evento de la linea: " + i + "." + ex.Message);
                        }
                    }
                    else
                    {
                        /*string idEmpresa = datos.Rows[i][0].ToString().Trim();
                        string Region = datos.Rows[i][1].ToString().Trim();
                        string Pais = datos.Rows[i][2].ToString().Trim();
                        string Departamento = datos.Rows[i][3].ToString().Trim();*/
                        string cadenaValor = datos.Rows[i][1].ToString().Trim();
                        string macroProceso = datos.Rows[i][2].ToString().Trim();
                        string proceso = datos.Rows[i][3].ToString().Trim();
                        string subproceso = datos.Rows[i][4].ToString().Trim();
                        string ValorRecuperadoTotal = datos.Rows[i][5].ToString().Trim();
                        string Moneda2 = datos.Rows[i][6].ToString().Trim();
                        string TasaCambio2 = datos.Rows[i][7].ToString().Trim();
                        string ValorPesos2 = datos.Rows[i][8].ToString().Trim();
                        string Recuperacion = datos.Rows[i][9].ToString().Trim();
                        string FechaContabilidad = datos.Rows[i][10].ToString().Trim();
                        string HoraContabilidad = datos.Rows[i][11].ToString().Trim();

                        if (string.IsNullOrEmpty(cadenaValor) &&
                            string.IsNullOrEmpty(macroProceso) && string.IsNullOrEmpty(proceso) && string.IsNullOrEmpty(subproceso) && string.IsNullOrEmpty(ValorRecuperadoTotal) && string.IsNullOrEmpty(Moneda2) &&
                            string.IsNullOrEmpty(TasaCambio2) && string.IsNullOrEmpty(ValorPesos2) && string.IsNullOrEmpty(Recuperacion) && string.IsNullOrEmpty(FechaContabilidad) && string.IsNullOrEmpty(HoraContabilidad))
                        {
                            continue;
                        }
                        else
                        {
                            throw new Exception("Verificar la línea <B>" + (Linea + 3) + "</B> El valor del campo <B>Empresa</B> no puede estar vacío!");
                        }
                    }
                }

                //-----------------------------------------
                //Valida idJerarquia segun el nodo de la Jerarquia Organizacional
                //-----------------------------------------
                //for (int i = 0; i < dtJerarquia.Rows.Count; i++)
                //{
                //    DataTable dtRpta = new DataTable();
                //    if (dtJerarquia.Rows[i][0].ToString().Trim() != string.Empty)
                //    {
                //        string parametro = dtJerarquia.Rows[i][0].ToString().Trim();
                //        dtRpta = CargaMasiva.DataExisteJerarquia(Convert.ToInt32(dtJerarquia.Rows[i][0]));
                //        int cuantos = Convert.ToInt32(dtRpta.Rows.Count);
                //        if (cuantos > 0)
                //        {
                //            continue;
                //        }
                //        else
                //        {
                //            throw new System.ArgumentException("No es posible procesar el archivo. <B> Número de Línea: </B>" + (i + 3) + ", el valor del campo id Cargo Responsable no se encuentra registrado. ", "<B> IdCargoResponsable, Valor del parámetro: " + parametro + " </B>");
                //        }
                //    }
                //    else
                //    {
                //        throw new System.ArgumentException("No es posible procesar el archivo. <B> Número de Línea: </B>" + (i + 3) + ", no se permiten valores vacíos en el campo IdCargoResponsable del archivo. ", "<B> IdCargoResponsable </B>");
                //    }
                //}
                //writer.Flush();
                //writer.Close();
                //excelReader.Close();

                //-----------------------------------------
                //Ejecuto query
                //-----------------------------------------

                //cError ce = new cError();
                //int resultado = 0;
                //string error = string.Empty;
                //try
                //{
                //    resultado = CargaMasiva.cargaEvento(nombreSp, nombreArchivo, path_planos);
                //}
                //catch (Exception ex)
                //{
                //    error = ex.ToString();
                //    ce.errorMessage(ex.Message + ", " + ex.StackTrace);
                //    throw new Exception(ex.Message);
                //}

                //-----------------------------------------
                //Envío correo
                //-----------------------------------------


                //DataView view = new DataView(dtJerarquia);
                //view.Sort = "idJerarquia";
                //DataTable dtOrdenado = view.ToTable();
                //string[] registrosOrdenados = new string[dtOrdenado.Rows.Count * 2];
                //int k = 0, p = 0, contEncontrados = 0;
                //string codigoEvento = string.Empty, descripcionEvento = string.Empty;

                //for (int i = 0; i < dtOrdenado.Rows.Count; i++)
                //{
                //    string jerarquiaCapturada = dtOrdenado.Rows[i][0].ToString().Trim();

                //    if (jerarquiaCapturada != registrosOrdenados[p])
                //    {
                //        for (int j = 0; j < dtOrdenado.Rows.Count; j++)
                //        {
                //            string jerarquiaEncontrada = dtOrdenado.Rows[j][0].ToString().Trim();
                //            if (jerarquiaCapturada == jerarquiaEncontrada)
                //            {
                //                string codEven = dtOrdenado.Rows[j][1].ToString().Trim();
                //                codigoEvento += dtOrdenado.Rows[j][1].ToString().Trim() + ", ";
                //                descripcionEvento += " <br /> <B> " + codEven + ": </B> " + dtOrdenado.Rows[j][2].ToString().Trim() + ". ";
                //                contEncontrados++;
                //            }
                //        }

                //        if (contEncontrados > 0)
                //        {
                //            try
                //            {
                //                if (dtOrdenado.Rows.Count == 1)
                //                {
                //                    boolEnviarNotificacion(35, Convert.ToInt16("0"), Convert.ToInt32(jerarquiaCapturada), "",
                //                    "<B>Se ha actualizado el siguiente Evento por carga masiva:</B> <br /><br /><B>Descripción de Evento: </B> <br />" + descripcionEvento.Trim());
                //                    codigoEvento = string.Empty;
                //                    contEncontrados = 0;
                //                }
                //                else
                //                {
                //                    boolEnviarNotificacion(35, Convert.ToInt16("0"), Convert.ToInt32(jerarquiaCapturada), "",
                //                    "<B>Se han actualizado los siguientes Eventos por carga masiva:</B> <br /><br /><B>Descripción de los Eventos: </B> <br />" + descripcionEvento.Trim() + " <B> <br/><br/>Total Eventos actualizados: </B> " + dtOrdenado.Rows.Count + "<br /> ");
                //                    codigoEvento = string.Empty;
                //                    contEncontrados = 0;
                //                }
                //            }
                //            catch (Exception ex)
                //            {
                //                Mensaje("Error al enviar notificación de creacion de Evento." + ex.Message);
                //            }
                //        }

                //        if (registrosOrdenados[k] == null)
                //        {
                //            registrosOrdenados[k] = jerarquiaCapturada;
                //        }
                //        for (int m = 0; m < dtOrdenado.Rows.Count; m++)
                //        {
                //            if (registrosOrdenados[k] == dtOrdenado.Rows[m][0].ToString().Trim())
                //            {
                //                k++;
                //                p++;
                //                registrosOrdenados[p] = jerarquiaCapturada;
                //                break;
                //            }
                //        }
                //    }
                //    else
                //    {
                //        p++;
                //        registrosOrdenados[p] = jerarquiaCapturada;
                //    }
                //}
                omb.ShowMessage("Archivo procesado con éxito. Se actualizaron : " + contador + " registros satisfactoriamente.", 3, "Atencion");
                //info.Delete();
                /*}
                else
                {
                    Mensaje("Se presentó un problema mientras se realizaba la carga, por favor verifique los datos e intentelo nuevamente. " + error);
                }*/
            }
        }
        private void LoadFilePlantillaRiesgosControles()
        {
            int IdUsuario = Convert.ToInt32(Session["IdUsuario"].ToString());
            DateTime FechaRegistro = DateTime.Now;
            //Best Way To read file direct from stream
            cCargaMasivaRCE CargaMasiva = new cCargaMasivaRCE();
            IExcelDataReader excelReader = null;
            System.Data.DataSet ds = new System.Data.DataSet();
            try
            {

                //file.InputStream is the file stream stored in memeory by any ways like by upload file control or from database
                int excelFlag = 1; //this flag us used for execl file format .xls or .xlsx
                if (excelFlag == 1)
                {
                    //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(FUloadExcel.PostedFile.InputStream);
                    excelReader.IsFirstRowAsColumnNames = true;
                }
                else if (excelFlag == 2)
                {
                    //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(FUloadExcel.PostedFile.InputStream);
                    //excelReader.IsFirstRowAsColumnNames = true;
                }
            }
            catch (Exception ex)
            {
                Mensaje("Error al analizar el archivo Eventos. " + ex.Message);
            }
            if (excelReader != null)
            {
                ds = excelReader.AsDataSet();
                DataTable datos = ds.Tables[0];
                DataTable dt = CargaMasiva.LastCodRiesgovsControl();
                Session["IdControlesRiesgo"] = dt.Rows[0][0].ToString();
                CRiesgoControl riesgoControl = new CRiesgoControl();
                for (int i = 0; i < datos.Rows.Count; i++)
                {
                    if (datos.Rows[i][0].ToString().Trim() != "")
                    {
                        riesgoControl.IdRiesgo = ValidarNumero(datos.Rows[i][0].ToString().Trim());
                        riesgoControl.IdControl = ValidarNumero(datos.Rows[i][1].ToString().Trim());
                        riesgoControl.IdCausa = datos.Rows[i][2].ToString().Trim();
                        riesgoControl.IdUsuario = IdUsuario;

                        try
                        {
                            //CargaMasiva.registrarRiesgoControl(riesgoControl);
                        }
                        catch (Exception ex)
                        {
                            Mensaje("Error al registrar el Riesgo vs el Control de la linea: " + i + "." + ex.Message);
                        }
                    }
                }
                //6. Free resources (IExcelDataReader is IDisposable)
                excelReader.Close();
            }
        }
        private void LoadFilePLantillaRiesgos()
        {

            int IdUsuario = Convert.ToInt32(Session["IdUsuario"].ToString());
            string FechaRegistro = DateTime.Now.ToString(); int contador = 0;
            /*IExcelDataReader ExcelReader = ExcelReaderFactory.CreateBinaryReader(FUloadExcel.PostedFile.InputStream);
            ExcelReader.IsFirstRowAsColumnNames = true;
            System.Data.DataSet DSResult = ExcelReader.AsDataSet();
            ExcelReader.Close();
            armarInformacionRiesgo(DSResult.Tables[0]);*/
            ///Ajustes camilo 23-07-2015
            ///
            //Best Way To read file direct from stream
            cCargaMasivaRCE CargaMasiva = new cCargaMasivaRCE();
            IExcelDataReader excelReader = null;
            System.Data.DataSet ds = new System.Data.DataSet();
            try
            {

                //file.InputStream is the file stream stored in memeory by any ways like by upload file control or from database
                int excelFlag = 1; //this flag us used for execl file format .xls or .xlsx
                if (excelFlag == 1)
                {
                    //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(FUloadExcel.PostedFile.InputStream);
                    excelReader.IsFirstRowAsColumnNames = true;
                }
                else if (excelFlag == 2)
                {
                    //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(FUloadExcel.PostedFile.InputStream);
                    //excelReader.IsFirstRowAsColumnNames = true;
                }
            }
            catch (Exception ex)
            {
                Mensaje("Error al analizar el archivo Riesgos. " + ex.Message);
            }
            if (excelReader != null)
            {
                //...
                //3. DataSet - The result of each spreadsheet will be created in the result.Tables
                ds = excelReader.AsDataSet();
                //...
                ////4. DataSet - Create column names from first row
                //excelReader.IsFirstRowAsColumnNames = true;
                //DataSet result = excelReader.AsDataSet();

                ////5. Data Reader methods
                //while (excelReader.Read())
                //{
                //    //excelReader.GetInt32(0);
                //}
                DataTable datos = ds.Tables[0];
                DataTable dtlc = CargaMasiva.lastCod();
                Session["IdRiesgo"] = dtlc.Rows[0][0].ToString();
                DataTable dt_codRiesgo = new DataTable();
                DataTable dtInfoCorreoRiesgos = new DataTable();
                dtInfoCorreoRiesgos.Columns.Add("idDestino");
                dtInfoCorreoRiesgos.Columns.Add("ultCodRiesgo");
                dtInfoCorreoRiesgos.Columns.Add("nombreRiesgo");
                dtInfoCorreoRiesgos.Columns.Add("descripcionRiesgo");
                for (int i = 1; i < datos.Rows.Count; i++)
                {
                    if (datos.Rows[i][0].ToString().Trim() != "")
                    {
                        string Region = datos.Rows[i][0].ToString().Trim();
                        string Pais = datos.Rows[i][1].ToString().Trim();
                        string Departamento = datos.Rows[i][2].ToString().Trim();
                        string Ciudad = datos.Rows[i][3].ToString().Trim();
                        string Oficina = datos.Rows[i][4].ToString().Trim();
                        string Cadena_de_valor = datos.Rows[i][5].ToString().Trim();
                        string Macroproceso = datos.Rows[i][6].ToString().Trim();
                        string Proceso = datos.Rows[i][7].ToString().Trim();
                        string Subproceso = datos.Rows[i][8].ToString().Trim();
                        string Actividad = datos.Rows[i][9].ToString().Trim();
                        string Riesgos_globales = datos.Rows[i][10].ToString().Trim();
                        string Clasificación_general = datos.Rows[i][11].ToString().Trim();
                        string Clasificación_particular = datos.Rows[i][12].ToString().Trim();
                        string Factor_de_riesgo_operativo = datos.Rows[i][13].ToString().Trim();
                        string Sub_factor_riesgo_operativo = datos.Rows[i][14].ToString().Trim();
                        string Tipo_de_evento = datos.Rows[i][15].ToString().Trim();
                        string Riesgo_asociado = datos.Rows[i][16].ToString().Trim();
                        string ListaRiesgoAsociadoLA = datos.Rows[i][17].ToString().Trim();
                        string ListaFactorRiesgoLAFT = datos.Rows[i][18].ToString().Trim();
                        string Nombre = datos.Rows[i][19].ToString().Trim();
                        string Descripción_del_riesgo = datos.Rows[i][20].ToString().Trim();
                        string estado = datos.Rows[i][21].ToString().Trim();
                        string Cargo_Responsable = datos.Rows[i][22].ToString().Trim();
                        string ListaCausas = datos.Rows[i][23].ToString().Trim();
                        string ListaConsecuencias = datos.Rows[i][24].ToString().Trim();
                        string Frecuencia_Cualitativa = datos.Rows[i][25].ToString().Trim();
                        string Se_esperaba_la_ocurrencia_de_un_evento_entre_un = datos.Rows[i][26].ToString().Trim();
                        string Porcentaje_y_un = datos.Rows[i][27].ToString().Trim();
                        string Impacto_cualitativo = datos.Rows[i][28].ToString().Trim();
                        string Pérdida_económica_entre = datos.Rows[i][29].ToString().Trim();
                        string y = datos.Rows[i][30].ToString().Trim();
                        string tratamiento = datos.Rows[i][31].ToString().Trim();
                        try
                        {
                            //buscar último codRiesgo}
                            dt_codRiesgo = cControl.ultimoCodRiesgo();

                            //se inserta el registro ***
                            CargaMasiva.registrarRiesgo(Region, Pais, Departamento, Ciudad, Oficina, Cadena_de_valor, Macroproceso, Proceso, Subproceso, Actividad, Riesgos_globales,
                            Clasificación_general, Clasificación_particular, Factor_de_riesgo_operativo, Sub_factor_riesgo_operativo, Tipo_de_evento, Riesgo_asociado, ListaRiesgoAsociadoLA,
                            ListaFactorRiesgoLAFT, Nombre, Descripción_del_riesgo, ListaCausas, ListaConsecuencias, Cargo_Responsable, Frecuencia_Cualitativa, Se_esperaba_la_ocurrencia_de_un_evento_entre_un,
                            Porcentaje_y_un, Impacto_cualitativo, Pérdida_económica_entre, y, tratamiento, IdUsuario, FechaRegistro, estado);

                            //  -yoendy
                            DataRow dro = dtInfoCorreoRiesgos.NewRow();
                            dro["idDestino"] = datos.Rows[i][22].ToString().Trim();
                            dro["ultCodRiesgo"] = dt_codRiesgo.Rows[0][0];
                            dro["nombreRiesgo"] = datos.Rows[i][19].ToString().Trim();
                            dro["descripcionRiesgo"] = datos.Rows[i][20].ToString().Trim();
                            dtInfoCorreoRiesgos.Rows.Add(dro);
                            contador++;
                        }
                        catch (Exception ex)
                        {
                            Mensaje("Error al registrar el riesgo de la linea: " + i + "." + ex.Message);
                        }
                    }
                }

                DataView view = new DataView(dtInfoCorreoRiesgos);
                view.Sort = "idDestino";
                DataTable dtOrdenado = view.ToTable();
                string[] valores = new string[dtOrdenado.Rows.Count * 2];
                int k = 0, pos = 0, foundCod = 0;
                string codigoRiesgo = string.Empty, nombre_riesgo = string.Empty, descripcion_Riesgo = string.Empty, descripcion_general = string.Empty;

                for (int m = 0; m < dtOrdenado.Rows.Count; m++)
                {
                    string compare = dtOrdenado.Rows[m][0].ToString().Trim();

                    if (compare != valores[pos])
                    {
                        DataTable dt = new DataTable();
                        dt.Columns.Add("riesgos");
                        for (int j = 0; j < dtOrdenado.Rows.Count; j++)
                        {
                            string encontrado = dtOrdenado.Rows[j][0].ToString().Trim();
                            if (compare == encontrado)
                            {
                                codigoRiesgo = dtOrdenado.Rows[j][1].ToString().Trim();
                                nombre_riesgo = dtOrdenado.Rows[j][2].ToString().Trim();
                                descripcion_Riesgo = dtOrdenado.Rows[j][3].ToString().Trim();
                                descripcion_general += "<br /><B>Código Riesgo: </B> " + "<a style='font-weight:normal'>" + codigoRiesgo.Trim() + "</a>" + ". <B>Nombre: </B> " + " <a style='font-weight:normal'>" + nombre_riesgo + "</a>"
                                + ". <B>Descripción: </B> " + "<a style='font-weight:normal'>" + descripcion_Riesgo + "</a> ." + "<br />";
                                DataRow dr_ = dt.NewRow();
                                dr_["riesgos"] = dtOrdenado.Rows[j][1].ToString().Trim();
                                dt.Rows.Add(dr_);
                                foundCod++;
                            }
                        }

                        if (foundCod > 0)
                        {
                            try
                            {
                                if (dtOrdenado.Rows.Count == 1)
                                {
                                    boolEnviarNotificacion(19, Convert.ToInt16("0"), Convert.ToInt32(compare), "",
                                   "<B>Registro de Riesgo por carga masiva:</B> <br /><br />Ha sido asignado como responsable de calificación de un riesgo: " +
                                   " <br /> <B>" + descripcion_general.Trim() + "</B><br /> <br />");
                                    codigoRiesgo = string.Empty;
                                    nombre_riesgo = string.Empty;
                                    descripcion_Riesgo = string.Empty;
                                    descripcion_general = string.Empty;
                                    foundCod = 0;
                                }
                                else
                                {
                                    boolEnviarNotificacion(19, Convert.ToInt16("0"), Convert.ToInt32(compare), "",
                                   "<B>Registro de Riesgo por carga masiva:</B> <br /><br />Ha sido asignado como responsable de calificación de los siguientes riesgos: " +
                                   " <br /> <B>" + descripcion_general.Trim() + "</B><br /> <br /> Estos controles fueron generados desde la aplicación Sherlock para la Gestión de Riesgos y Control Interno. <br /> ");
                                    codigoRiesgo = string.Empty;
                                    nombre_riesgo = string.Empty;
                                    descripcion_Riesgo = string.Empty;
                                    descripcion_general = string.Empty;
                                    foundCod = 0;
                                }
                            }
                            catch (Exception ex)
                            {
                                Mensaje("Error al enviar notificación de creacion de Evento." + ex.Message);
                            }
                        }

                        if (valores[k] == null)
                        {
                            valores[k] = compare;
                        }
                        for (int f = 0; f < dtOrdenado.Rows.Count; f++)
                        {
                            if (valores[k] == dtOrdenado.Rows[m][0].ToString().Trim())
                            {
                                k++;
                                pos++;
                                valores[pos] = compare;
                                break;
                            }
                        }
                    }
                    else
                    {
                        pos++;
                        valores[pos] = compare;
                    }
                }
                excelReader.Close();
                omb.ShowMessage("Archivo procesado con éxito. Se cargaron : " + contador + " registros satisfactoriamente.", 3, "Atencion");
            }


        }
        private void armarInformacionRiesgo(DataTable dt)
        {
            int intCantidadRegistros = dt.Rows.Count;
            DataTable dtInfoRiesgosIns = new DataTable();
            int IdUsuario = Convert.ToInt32(Session["IdUsuario"].ToString());
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string valor = dt.Rows[i][0].ToString().Trim();
            }
        }
        private void Mensaje(String Mensaje)
        {
            lblMsgBox1.Text = Mensaje;
            mpeMsgBox1.Show();
        }

        protected void ImButtonExcelExportPlantillaRiesgo_Click(object sender, ImageClickEventArgs e)
        {
            cCargaMasivaRCE carga = new cCargaMasivaRCE();
            if (DDLopciones.SelectedValue == "1")
            {
                byte[] excel = carga.mtdDescargarPlantillaRiesgos();
                if (excel != null)
                {
                    Response.Clear();
                    Response.Buffer = true;
                    Response.ContentType = "Application/pdf";
                    Response.AddHeader("Content-Disposition", "attachment; filename=RiesgosPlantillaCargaMasiva.xls");
                    Response.Charset = "";
                    Response.Cache.SetCacheability(HttpCacheability.NoCache);
                    Response.BinaryWrite(excel);
                    Response.End();
                }
            }
            if (DDLopciones.SelectedValue == "2")
            {
                DataTable dt = carga.DescargarPlantillaCargaControles();
                string[] columnNames = dt.Columns.Cast<DataColumn>()
                         .Select(x => x.ColumnName)
                         .ToArray();
                XLWorkbook workbook = new XLWorkbook();
                IXLWorksheet worksheet = workbook.Worksheets.Add("Plantilla");
                // Se crean los encabezados dinámicamente
                int cell = 0;
                foreach (string column in columnNames)
                {
                    worksheet.Cell(1, cell + 1).SetValue(column);
                    cell++;
                }
                worksheet.Columns().AdjustToContents();
                worksheet.Range(1, 1, 1, cell).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                // Prepara la respuesta
                HttpResponse httpResponse = Response;
                httpResponse.Clear();
                httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                httpResponse.AddHeader("content-disposition", "attachment;filename=\"Plantilla cargue controles.xlsx\"");

                // Vacíe el libro de trabajo a Response.OutputStream
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    workbook.SaveAs(memoryStream);
                    memoryStream.WriteTo(httpResponse.OutputStream);
                    memoryStream.Close();
                }

                httpResponse.End();
                httpResponse.End();
            }
            if (DDLopciones.SelectedValue == "3")
            {
                byte[] excel = carga.mtdDescargarPlantillaEventosCreacion();
                if (excel != null)
                {
                    Response.Clear();
                    Response.Buffer = true;
                    Response.ContentType = "Application/pdf";
                    Response.AddHeader("Content-Disposition", "attachment; filename=EventosCreacionPlantillaCargaMasiva.xls");
                    Response.Charset = "";
                    Response.Cache.SetCacheability(HttpCacheability.NoCache);
                    Response.BinaryWrite(excel);
                    Response.End();
                }
            }
            if (DDLopciones.SelectedValue == "5")
            {
                byte[] excel = carga.mtdDescargarPlantillaEventosDatosComplementarios();
                if (excel != null)
                {
                    Response.Clear();
                    Response.Buffer = true;
                    Response.ContentType = "Application/pdf";
                    Response.AddHeader("Content-Disposition", "attachment; filename=EventosDatosComplementariosPlantillaCargaMasiva.xls");
                    Response.Charset = "";
                    Response.Cache.SetCacheability(HttpCacheability.NoCache);
                    Response.BinaryWrite(excel);
                    Response.End();
                }
            }
            if (DDLopciones.SelectedValue == "6")
            {
                byte[] excel = carga.mtdDescargarPlantillaEventosConDatosComplementarios();
                if (excel != null)
                {
                    Response.Clear();
                    Response.Buffer = true;
                    Response.ContentType = "Application/pdf";
                    Response.AddHeader("Content-Disposition", "attachment; filename=EventosConDatosComplementariosPlantillaCargaMasiva.xls");
                    Response.Charset = "";
                    Response.Cache.SetCacheability(HttpCacheability.NoCache);
                    Response.BinaryWrite(excel);
                    Response.End();
                }
            }
            if (DDLopciones.SelectedValue == "4")
            {
                XLWorkbook workbook = new XLWorkbook();
                IXLWorksheet worksheet = workbook.Worksheets.Add("Plantilla");
                worksheet.Cell(1, 1).SetValue("IdRiesgo");
                worksheet.Cell(1, 2).SetValue("IdControl");
                worksheet.Cell(1, 3).SetValue("IdCausa");
                worksheet.Columns().AdjustToContents();
                worksheet.Range(1, 1, 1, 3).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                // Prepara la respuesta
                HttpResponse httpResponse = Response;
                Response.Clear();
                Response.Buffer = true;
                Response.ContentType = "Application/pdf";
                Response.AddHeader("Content-Disposition", "attachment; filename=RiesgosVsControlesPlantillaCargaMasiva.xlsx");
                Response.Charset = "";
                Response.Cache.SetCacheability(HttpCacheability.NoCache);


                // Vacíe el libro de trabajo a Response.OutputStream
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    workbook.SaveAs(memoryStream);
                    memoryStream.WriteTo(httpResponse.OutputStream);
                    memoryStream.Close();
                }

                httpResponse.End();
            }
        }

        private void LoadPlantilla(string tipoCarga)
        {
            try
            {
                string pathFile = FUloadExcel.FileName;
                Byte[] archivo = FUloadExcel.FileBytes;
                int length = Convert.ToInt32(FUloadExcel.FileContent.Length);
                string extension = Path.GetExtension(FUloadExcel.FileName).ToLower().ToString().Trim();
                int IdUsuario = Convert.ToInt32(Session["IdUsuario"].ToString());
                //saveFile(FUloadExcel.FileName, length, archivo, extension, pathFile);

                // Se crean los objetos
                cControlEntity controlEntity = new cControlEntity();
                CRiesgoControl riesgoControl = new CRiesgoControl();
                cCargaMasivaRCE carga = new cCargaMasivaRCE();
                cRiesgo riesgo = new cRiesgo();
                DataTable dtRegistrosControl = new DataTable();
                dtRegistrosControl.Columns.Add("idDestino");
                dtRegistrosControl.Columns.Add("idControl");
                dtRegistrosControl.Columns.Add("nombreControl");
                dtRegistrosControl.Columns.Add("descripcionControl");
                DataTable dt_idControl = new DataTable();

                // Leer archivo excel
                using (XLWorkbook excelWorkbook = new XLWorkbook(FUloadExcel.FileContent))
                {
                    IXLRows nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();

                    // Toma el ultimo Id de la tabla para mostrar la grilla final
                    using (DataTable dt = carga.LastCodRiesgovsControl())
                    {
                        Session["IdControlesRiesgo"] = dt.Rows[0][0].ToString();
                    }

                    // Recorre el archivo y valida que se encuentre correcto de acuerdo al tipo de carga
                    foreach (IXLRow dataRow in nonEmptyDataRows)
                    {
                        if (dataRow.RowNumber() > 1)
                        {
                            if (tipoCarga == "2")
                            {
                                // Se valida que los datos en la plantilla sean los correctos
                                if (!int.TryParse(dataRow.Cell(4).Value.ToString(), out int value) || !int.TryParse(dataRow.Cell(5).Value.ToString(), out value) ||
                                    !int.TryParse(dataRow.Cell(6).Value.ToString(), out value) || !int.TryParse(dataRow.Cell(5).Value.ToString(), out value)
                                    )
                                {
                                    throw new Exception($"Se han encontrado valores no válidos en la línea {dataRow.RowNumber()} por favor verifique.");
                                }
                            }
                            else if (tipoCarga == "4")
                            {
                                // Llena la entidad para validar que los valores sean correctos
                                riesgoControl.IdRiesgo = ValidarNumero(dataRow.Cell(1).Value.ToString());
                                riesgoControl.IdControl = ValidarNumero(dataRow.Cell(2).Value.ToString());
                                riesgoControl.IdCausa = dataRow.Cell(3).Value.ToString();

                                if (riesgoControl.IdRiesgo is null || riesgoControl.IdControl is null)
                                {
                                    throw new Exception($"Se han encontrado valores no válidos en la línea {dataRow.RowNumber()} por favor verifique.");
                                }

                                if (cControl.ValidarExistenciaControl(Convert.ToInt32(riesgoControl.IdControl)) == 0)
                                {
                                    throw new Exception($"El control {riesgoControl.IdControl} de línea {dataRow.RowNumber()} no existe, por favor verifique.");
                                }

                                using (DataTable dt = cControl.SeleccionarRiesgo(riesgoControl.IdRiesgo))
                                {
                                    // Validación de las causas asosciadas al riesgo
                                    if (dt != null && dt.Rows.Count > 0)
                                    {
                                        string[] lstCausasRiesgo = dt.Rows[0][1].ToString().Split('|');
                                        lstCausasRiesgo = lstCausasRiesgo.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                                        string[] lstCausasArchivo = riesgoControl.IdCausa.Split('|');
                                        foreach (string causa in lstCausasArchivo)
                                        {
                                            string[] results = Array.FindAll(lstCausasRiesgo, s => s.Equals(causa));
                                            if (results.Length < 1)
                                            {
                                                throw new Exception($"Se han encontrado causas no válidas en la línea {dataRow.RowNumber()} por favor verifique.");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception($"El Riesgo de la línea {dataRow.RowNumber()} no existe o no tiene causas asociadas.");
                                    }
                                }
                            }
                        }
                    }
                    if (tipoCarga == "2")
                    {
                        //buscar ultimo idControl}
                        DataTable dtInfo = cControl.GetLastControl();

                        if (dtInfo.Rows.Count > 0)
                            Session["idControlIn"] = dtInfo.Rows[0]["LastControl"].ToString().Trim();

                    }
                    // Recorre el archivo y realiza la inserción
                    foreach (IXLRow dataRow in nonEmptyDataRows)
                    {
                        if (dataRow.RowNumber() > 1)
                        {
                            if (tipoCarga == "2")
                            {
                                controlEntity.NombreControl = dataRow.Cell(1).Value.ToString();
                                controlEntity.DescripcionControl = dataRow.Cell(2).Value.ToString();
                                controlEntity.ObjetivoControl = dataRow.Cell(3).Value.ToString();
                                controlEntity.Responsable = Convert.ToInt32(dataRow.Cell(4).Value);
                                controlEntity.IdPeriodicidad = Convert.ToInt32(dataRow.Cell(5).Value);
                                controlEntity.IdTest = Convert.ToInt32(dataRow.Cell(6).Value);
                                controlEntity.IdMitiga = Convert.ToInt32(dataRow.Cell(7).Value);
                                controlEntity.ResponsableEjecucion = dataRow.Cell(8).Value.ToString();
                                controlEntity.IdClaseControl = ValidarNumero(dataRow.Cell(9).Value.ToString());
                                controlEntity.IdTipoControl = ValidarNumero(dataRow.Cell(10).Value.ToString());
                                controlEntity.IdResponsableExperiencia = ValidarNumero(dataRow.Cell(11).Value.ToString());
                                controlEntity.IdDocumentacion = ValidarNumero(dataRow.Cell(12).Value.ToString());
                                controlEntity.IdResponsabilidad = ValidarNumero(dataRow.Cell(13).Value.ToString());
                                controlEntity.Variable6 = ValidarNumero(dataRow.Cell(14).Value.ToString());
                                controlEntity.Variable7 = ValidarNumero(dataRow.Cell(15).Value.ToString());
                                controlEntity.Variable8 = ValidarNumero(dataRow.Cell(16).Value.ToString());
                                controlEntity.Variable9 = ValidarNumero(dataRow.Cell(17).Value.ToString());
                                controlEntity.Variable10 = ValidarNumero(dataRow.Cell(18).Value.ToString());
                                controlEntity.Variable11 = ValidarNumero(dataRow.Cell(19).Value.ToString());
                                controlEntity.Variable12 = ValidarNumero(dataRow.Cell(20).Value.ToString());
                                controlEntity.Variable13 = ValidarNumero(dataRow.Cell(21).Value.ToString());
                                controlEntity.Variable14 = ValidarNumero(dataRow.Cell(22).Value.ToString());
                                controlEntity.Estado = ValidarNumero(dataRow.Cell(24).Value.ToString());
                                controlEntity.IdUsuario = Convert.ToInt32(Session["IdUsuario"].ToString());
                                controlEntity.IdCalificacionControl = carga.CalcularEficacia(controlEntity, Session["IdControl"].ToString());

                                //buscar ultimo idControl}
                                dt_idControl = cControl.ultimoIdControl();

                                // Se inserta el control en la tabla
                                cControl.InsertControl(controlEntity);
                                //  -yoendy
                                DataRow dr = dtRegistrosControl.NewRow();
                                dr["idDestino"] = dataRow.Cell(4).Value;
                                dr["idControl"] = dt_idControl.Rows[0][0];
                                dr["nombreControl"] = dataRow.Cell(1).Value.ToString();
                                dr["descripcionControl"] = dataRow.Cell(2).Value.ToString();
                                dtRegistrosControl.Rows.Add(dr);
                            }
                            else if (tipoCarga == "4")
                            {
                                DataTable dt = carga.LastCodRiesgovsControl();
                                riesgoControl.IdRiesgo = ValidarNumero(dataRow.Cell(1).Value.ToString());
                                riesgoControl.IdControl = ValidarNumero(dataRow.Cell(2).Value.ToString());
                                riesgoControl.IdCausa = dataRow.Cell(3).Value.ToString();
                                riesgoControl.IdUsuario = IdUsuario;

                                // inserta la ascoación Riesgo- Control en la tabla
                                riesgo.registrarRiesgoControl(riesgoControl);

                                // Inserta las causas asociadas al riesgo
                                riesgo.registrarCausaRiesgoControl(riesgoControl);

                                // Recalcula calficación riesgo inherente
                                DataTable dtInfoRiesgo = cControl.CalcularRiesgoResidual(dataRow.Cell(1).Value.ToString());
                            }

                        }
                    }
                    if (nonEmptyDataRows.Count() > 0)
                    {
                        DataView view = new DataView(dtRegistrosControl);
                        view.Sort = "idDestino";
                        DataTable dtOrdenado = view.ToTable();
                        string[] valores = new string[dtOrdenado.Rows.Count * 2];
                        int k = 0, pos = 0, foundCod = 0;
                        string codigoRegistro = string.Empty, nombre_control = string.Empty, descripcion_Controles = string.Empty, descripcion_general = string.Empty;

                        for (int i = 0; i < dtOrdenado.Rows.Count; i++)
                        {
                            string compare = dtOrdenado.Rows[i][0].ToString().Trim();

                            if (compare != valores[pos])
                            {
                                DataTable dt = new DataTable();
                                dt.Columns.Add("controles");
                                for (int j = 0; j < dtOrdenado.Rows.Count; j++)
                                {
                                    string encontrado = dtOrdenado.Rows[j][0].ToString().Trim();
                                    if (compare == encontrado)
                                    {
                                        codigoRegistro = dtOrdenado.Rows[j][1].ToString().Trim();
                                        nombre_control = dtOrdenado.Rows[j][2].ToString().Trim();
                                        descripcion_Controles = dtOrdenado.Rows[j][3].ToString().Trim();
                                        descripcion_general += "<br /><B>Código control: </B> " + "<a style='font-weight:normal'>" + codigoRegistro.Trim() + "</a>" + ". <B>Nombre: </B> " + " <a style='font-weight:normal'>" + nombre_control + "</a>"
                                        + ". <B>Descripción: </B> " + "<a style='font-weight:normal'>" + descripcion_Controles + "</a> ." + "<br />";
                                        DataRow dr = dt.NewRow();
                                        dr["controles"] = dtOrdenado.Rows[j][1].ToString().Trim();
                                        dt.Rows.Add(dr);
                                        foundCod++;
                                    }
                                }

                                if (foundCod > 0)
                                {
                                    try
                                    {
                                        if (dtOrdenado.Rows.Count == 1)
                                        {
                                            boolEnviarNotificacion(7, Convert.ToInt16("0"), Convert.ToInt32(compare), "",
                                           "<B>Registro de Control por carga masiva:</B> <br /><br />Ha sido asignado como responsable de calificación de un control: " +
                                           " <br /> <B>" + descripcion_general.Trim() + "</B><br /> <br /> ****Nota: Este correo es enviado automáticamente por la aplicación Sherlock.<br /> ");
                                            codigoRegistro = string.Empty;
                                            nombre_control = string.Empty;
                                            descripcion_Controles = string.Empty;
                                            descripcion_general = string.Empty;
                                            foundCod = 0;
                                        }
                                        else
                                        {
                                            int cuantos = codigoRegistro.Length - 2;
                                            string codigoControl = codigoRegistro.Substring(0, cuantos);

                                            boolEnviarNotificacion(7, Convert.ToInt16("0"), Convert.ToInt32(compare), "",
                                           "<B>Registro de Control por carga masiva:</B> <br /><br />Ha sido asignado como responsable de calificación de los siguientes controles: " +
                                           " <br /> <B>" + descripcion_general.Trim() + "</B><br /> <br /> Estos controles fueron generados desde la aplicación Sherlock para la Gestión de Riesgos y Control Interno. <br /> ");
                                            codigoRegistro = string.Empty;
                                            nombre_control = string.Empty;
                                            descripcion_Controles = string.Empty;
                                            descripcion_general = string.Empty;
                                            foundCod = 0;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Mensaje("Error al enviar notificación de creacion de Evento." + ex.Message);
                                    }
                                }

                                if (valores[k] == null)
                                {
                                    valores[k] = compare;
                                }
                                for (int m = 0; m < dtOrdenado.Rows.Count; m++)
                                {
                                    if (valores[k] == dtOrdenado.Rows[m][0].ToString().Trim())
                                    {
                                        k++;
                                        pos++;
                                        valores[pos] = compare;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                pos++;
                                valores[pos] = compare;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public int? ValidarNumero(string value)
        {
            int a = 0;
            if (int.TryParse(value, out a))
            {
                return a;
            }
            else
            {
                return null;
            }
        }

        private void saveFile(string NombreArchivo, int Length, byte[] archivo, string extension, string pathFile)
        {
            /*string path = ConfigurationManager.AppSettings.Get("DirectorioDocumentos").ToString();
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            //string fullPath = Server.MapPath("/");
            fuArchivoPerfil.PostedFile.SaveAs(path + NombreArchivo);*/
            string strErrMsg = string.Empty;
            string FechaRegistro = DateTime.Now.ToString();
            cCargaMasivaRCE carga = new cCargaMasivaRCE();

            if (!carga.Guardar(NombreArchivo, Length, archivo, ref strErrMsg, FechaRegistro, pathFile, archivo))
            {
                omb.ShowMessage(strErrMsg, 1, "Atención");
            }
        }

        protected void GVriesgos_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            PagIndex = e.NewPageIndex;
            GVriesgos.PageIndex = PagIndex;
            GVriesgos.DataBind();
            LoadGridRiesgos();
        }

        protected void ImbCancel_Click(object sender, ImageClickEventArgs e)
        {
            DDLopciones.ClearSelection();
            TrButtonsExportRiesgos.Visible = false;
            TrLoadFile.Visible = false;
            TgridRiesgos.Visible = false;
        }

        protected void GVeventos_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            PagIndex = e.NewPageIndex;
            GVeventos.PageIndex = PagIndex;
            GVeventos.DataBind();
            LoadGridEventos();
        }

        protected void GVcontrol_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            PagIndex = e.NewPageIndex;
            GVcontrol.PageIndex = PagIndex;
            GVcontrol.DataBind();
            LoadGridControl();
        }

        protected void GVriesgovscontrol_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            PagIndex = e.NewPageIndex;
            GVriesgovscontrol.PageIndex = PagIndex;
            GVriesgovscontrol.DataBind();
            LoadGridRiesgovsControl();
        }

        //servicio de correo
        private Boolean boolEnviarNotificacion(int idEvento, int idRegistro, int idNodoJerarquia, string FechaFinal, string textoAdicional)
        {
            #region Variables
            bool err = false;
            string Destinatario = string.Empty, Copia = string.Empty, Asunto = string.Empty, Otros = string.Empty, Cuerpo = string.Empty, NroDiasRecordatorio = string.Empty;
            string selectCommand = string.Empty, AJefeInmediato = string.Empty, AJefeMediato = string.Empty, RequiereFechaCierre = string.Empty;
            string idJefeInmediato = string.Empty, idJefeMediato = string.Empty;
            string conString = WebConfigurationManager.ConnectionStrings["SarlaftConnectionString"].ConnectionString;
            #endregion Variables

            try
            {
                #region informacion basica
                SqlDataAdapter dad = null;
                DataTable dtblDiscuss = new DataTable();
                DataView view = null;

                if (!string.IsNullOrEmpty(idEvento.ToString().Trim()))
                {
                    //Consulta la informacion basica necesario para enviar el correo de la tabla correos destinatarios
                    selectCommand = "SELECT CD.Copia, CD.Otros, CD.Asunto, CD.Cuerpo, CD.NroDiasRecordatorio, CD.AJefeInmediato, CD.AJefeMediato, E.RequiereFechaCierre " +
                        "FROM [Notificaciones].[CorreosDestinatarios] AS CD INNER JOIN [Notificaciones].[Evento] AS E ON CD.IdEvento = E.IdEvento " +
                        "WHERE E. IdEvento = " + idEvento;

                    dad = new SqlDataAdapter(selectCommand, conString);
                    dad.Fill(dtblDiscuss);
                    view = new DataView(dtblDiscuss);

                    foreach (DataRowView row in view)
                    {
                        Copia = row["Copia"].ToString().Trim();
                        Otros = row["Otros"].ToString().Trim();
                        Asunto = row["Asunto"].ToString().Trim();
                        if (idEvento == 7)
                        {
                            Cuerpo = textoAdicional + "<br />";
                        }
                        else
                        {
                            Cuerpo = textoAdicional + "<br />***Nota: " + row["Cuerpo"].ToString().Trim();
                        }
                        NroDiasRecordatorio = row["NroDiasRecordatorio"].ToString().Trim();
                        AJefeInmediato = row["AJefeInmediato"].ToString().Trim();
                        AJefeMediato = row["AJefeMediato"].ToString().Trim();
                        RequiereFechaCierre = row["RequiereFechaCierre"].ToString().Trim();
                    }
                }
                #endregion

                #region correo del Destinatario
                //Consulta el correo del Destinatario segun el nodo de la Jerarquia Organizacional
                if (!string.IsNullOrEmpty(idNodoJerarquia.ToString().Trim()))
                {
                    selectCommand = "SELECT DJ.CorreoResponsable, JO.idPadre " +
                        "FROM [Parametrizacion].[JerarquiaOrganizacional] AS JO INNER JOIN [Parametrizacion].[DetalleJerarquiaOrg] AS DJ ON DJ.idHijo = JO.idHijo " +
                        "WHERE JO.idHijo = " + idNodoJerarquia;

                    dad = new SqlDataAdapter(selectCommand, conString);
                    dtblDiscuss.Clear();
                    dad.Fill(dtblDiscuss);
                    view = new DataView(dtblDiscuss);

                    foreach (DataRowView row in view)
                    {
                        Destinatario = row["CorreoResponsable"].ToString().Trim();
                        idJefeInmediato = row["idPadre"].ToString().Trim();
                    }
                }
                #endregion

                #region correo del Jefe Inmediato
                //Consulta el correo del Jefe Inmediato
                if (AJefeInmediato == "SI")
                {
                    if (!string.IsNullOrEmpty(idJefeInmediato.Trim()))
                    {
                        selectCommand = "SELECT DJ.CorreoResponsable, JO.idPadre " +
                            "FROM [Parametrizacion].[JerarquiaOrganizacional] AS JO INNER JOIN [Parametrizacion].[DetalleJerarquiaOrg] AS DJ ON DJ.idHijo = JO.idHijo " +
                            "WHERE JO.idHijo = " + idJefeInmediato;

                        dad = new SqlDataAdapter(selectCommand, conString);
                        dtblDiscuss.Clear();
                        dad.Fill(dtblDiscuss);
                        view = new DataView(dtblDiscuss);

                        foreach (DataRowView row in view)
                        {
                            Destinatario = Destinatario + ";" + row["CorreoResponsable"].ToString().Trim();
                            idJefeMediato = row["idPadre"].ToString().Trim();
                        }
                    }
                }
                #endregion

                #region correo del Jefe Mediato
                //Consulta el correo del Jefe Mediato
                if (AJefeMediato == "SI")
                {
                    if (!string.IsNullOrEmpty(idJefeMediato.Trim()))
                    {
                        selectCommand = "SELECT DJ.CorreoResponsable, JO.idPadre " +
                            "FROM [Parametrizacion].[JerarquiaOrganizacional] AS JO INNER JOIN [Parametrizacion].[DetalleJerarquiaOrg] AS DJ ON DJ.idHijo = JO.idHijo " +
                            "WHERE JO.idHijo = " + idJefeMediato;

                        dad = new SqlDataAdapter(selectCommand, conString);
                        dtblDiscuss.Clear();
                        dad.Fill(dtblDiscuss);
                        view = new DataView(dtblDiscuss);

                        foreach (DataRowView row in view)
                        {
                            Destinatario = Destinatario + ";" + row["CorreoResponsable"].ToString().Trim();
                        }
                    }
                }
                #endregion

                #region Correos Enviados
                //Insertar el Registro en la tabla de Correos Enviados
                SqlDataSource200.InsertParameters["Destinatario"].DefaultValue = Destinatario.Trim();
                SqlDataSource200.InsertParameters["Copia"].DefaultValue = Copia;
                SqlDataSource200.InsertParameters["Otros"].DefaultValue = Otros;
                SqlDataSource200.InsertParameters["Asunto"].DefaultValue = Asunto;
                SqlDataSource200.InsertParameters["Cuerpo"].DefaultValue = Cuerpo;
                SqlDataSource200.InsertParameters["Estado"].DefaultValue = "POR ENVIAR";
                SqlDataSource200.InsertParameters["Tipo"].DefaultValue = "CREACION";
                SqlDataSource200.InsertParameters["FechaEnvio"].DefaultValue = System.DateTime.Now.ToString().Trim();
                SqlDataSource200.InsertParameters["IdEvento"].DefaultValue = idEvento.ToString().Trim();
                SqlDataSource200.InsertParameters["IdRegistro"].DefaultValue = idRegistro.ToString().Trim();
                SqlDataSource200.InsertParameters["IdUsuario"].DefaultValue = Session["idUsuario"].ToString().Trim(); //Aca va el id del Usuario de la BD
                SqlDataSource200.InsertParameters["FechaRegistro"].DefaultValue = System.DateTime.Now.ToString().Trim();
                SqlDataSource200.Insert();
                #endregion
            }
            catch (Exception except)
            {
                // Handle the Exception.
                omb.ShowMessage("Error en el envío de la notificación." + "<br/>" + "Descripción: " + except.Message.ToString().Trim(), 1, "Atención");
                err = true;
            }

            if (!err)
            {
                #region Restro
                // Si no existe error en la creacion del registro en el log de correos enviados se procede a escribir en la tabla CorreosRecordatorios y a enviar el correo 
                if (RequiereFechaCierre == "SI" && FechaFinal != "")
                {
                    //Si los NroDiasRecordatorio es diferente de vacio se inserta el registro correspondiente en la tabla CorreosRecordatorio
                    SqlDataSource201.InsertParameters["IdCorreosEnviados"].DefaultValue = LastInsertIdCE.ToString().Trim();
                    SqlDataSource201.InsertParameters["NroDiasRecordatorio"].DefaultValue = NroDiasRecordatorio;
                    SqlDataSource201.InsertParameters["Estado"].DefaultValue = "POR ENVIAR";
                    SqlDataSource201.InsertParameters["FechaFinal"].DefaultValue = FechaFinal;
                    SqlDataSource201.InsertParameters["IdUsuario"].DefaultValue = Session["idUsuario"].ToString().Trim(); //Aca va el id del Usuario de la BD
                    SqlDataSource201.InsertParameters["FechaRegistro"].DefaultValue = System.DateTime.Now.ToString().Trim();
                    SqlDataSource201.Insert();
                }
                #endregion

                try
                {
                    #region Envio Correo
                    MailMessage message = new MailMessage();
                    SmtpClient smtpClient = new SmtpClient();
                    MailAddress fromAddress = new MailAddress(((System.Net.NetworkCredential)(smtpClient.Credentials)).UserName, "Software Sherlock");
                    message.From = fromAddress;//here you can set address

                    #region Destinatario
                    foreach (string substr in Destinatario.Split(';'))
                    {
                        if (!string.IsNullOrEmpty(substr.Trim()))
                        {
                            message.To.Add(substr);
                        }
                    }
                    #endregion

                    #region Copia
                    if (Copia.Trim() != "")
                    {
                        foreach (string substr in Copia.Split(';'))
                        {
                            if (!string.IsNullOrEmpty(substr.Trim()))
                            {
                                message.CC.Add(substr);
                            }
                        }
                    }
                    #endregion

                    #region Otros
                    if (Otros.Trim() != "")
                    {
                        foreach (string substr in Otros.Split(';'))
                        {
                            if (!string.IsNullOrEmpty(substr.Trim()))
                            {
                                message.CC.Add(substr);
                            }
                        }
                    }
                    #endregion

                    message.Subject = Asunto;//subject of email
                    message.IsBodyHtml = true;//To determine email body is html or not
                    message.Body = Cuerpo;

                    smtpClient.Send(message);
                    #endregion
                }
                catch (Exception ex)
                {
                    //throw exception here you can write code to handle exception here
                    omb.ShowMessage("Error en el envío de la notificación." + "<br/>" + "Descripción: " + ex.Message.ToString().Trim(), 1, "Atención");
                    err = true;
                }

                if (!err)
                {
                    //Actualiza el Estado del Correo Enviado
                    SqlDataSource200.UpdateParameters["IdCorreosEnviados"].DefaultValue = LastInsertIdCE.ToString().Trim();
                    SqlDataSource200.UpdateParameters["Estado"].DefaultValue = "ENVIADO";
                    SqlDataSource200.UpdateParameters["FechaEnvio"].DefaultValue = System.DateTime.Now.ToString().Trim();
                    SqlDataSource200.Update();
                }
            }

            return (err);
        }
        private Boolean boolEnviarNotificacionRiesgos(int idEvento, int idRegistro, int idNodoJerarquia, string FechaFinal, string textoAdicional, string TipoArea)
        {
            #region Variables
            bool err = false;
            string Destinatario = string.Empty, Copia = string.Empty, Asunto = string.Empty, Otros = string.Empty, Cuerpo = string.Empty, NroDiasRecordatorio = string.Empty;
            string selectCommand = string.Empty, AJefeInmediato = string.Empty, AJefeMediato = string.Empty, RequiereFechaCierre = string.Empty;
            string idJefeInmediato = string.Empty, idJefeMediato = string.Empty;
            string conString = WebConfigurationManager.ConnectionStrings["SarlaftConnectionString"].ConnectionString;
            #endregion Variables

            try
            {
                #region informacion basica
                SqlDataAdapter dad = null;
                DataTable dtblDiscuss = new DataTable();
                DataView view = null;

                if (!string.IsNullOrEmpty(idEvento.ToString().Trim()))
                {
                    //Consulta la informacion basica necesario para enviar el correo de la tabla correos destinatarios
                    selectCommand = "SELECT CD.Copia, CD.Otros, CD.Asunto, CD.Cuerpo, CD.NroDiasRecordatorio, CD.AJefeInmediato, CD.AJefeMediato, E.RequiereFechaCierre " +
                        "FROM [Notificaciones].[CorreosDestinatarios] AS CD INNER JOIN [Notificaciones].[Evento] AS E ON CD.IdEvento = E.IdEvento " +
                        "WHERE E. IdEvento = " + idEvento;

                    dad = new SqlDataAdapter(selectCommand, conString);
                    dad.Fill(dtblDiscuss);
                    view = new DataView(dtblDiscuss);

                    foreach (DataRowView row in view)
                    {
                        Copia = row["Copia"].ToString().Trim();
                        Otros = row["Otros"].ToString().Trim();
                        Asunto = row["Asunto"].ToString().Trim();
                        Cuerpo = textoAdicional + "<br />***Nota: " + row["Cuerpo"].ToString().Trim();
                        NroDiasRecordatorio = row["NroDiasRecordatorio"].ToString().Trim();
                        AJefeInmediato = row["AJefeInmediato"].ToString().Trim();
                        AJefeMediato = row["AJefeMediato"].ToString().Trim();
                        RequiereFechaCierre = row["RequiereFechaCierre"].ToString().Trim();
                    }
                }
                #endregion

                #region correo del Destinatario
                //Consulta el correo del Destinatario segun el nodo de la Jerarquia Organizacional
                if (!string.IsNullOrEmpty(idNodoJerarquia.ToString().Trim()))
                {
                    selectCommand = "SELECT DJ.CorreoResponsable, JO.idPadre " +
                        "FROM [Parametrizacion].[JerarquiaOrganizacional] AS JO INNER JOIN [Parametrizacion].[DetalleJerarquiaOrg] AS DJ ON DJ.idHijo = JO.idHijo " +
                        "WHERE JO.TipoArea = " + TipoArea;

                    dad = new SqlDataAdapter(selectCommand, conString);
                    dtblDiscuss.Clear();
                    dad.Fill(dtblDiscuss);
                    view = new DataView(dtblDiscuss);

                    foreach (DataRowView row in view)
                    {
                        Destinatario = row["CorreoResponsable"].ToString().Trim();
                        idJefeInmediato = row["idPadre"].ToString().Trim();
                    }
                }
                #endregion

                #region correo del Jefe Inmediato
                //Consulta el correo del Jefe Inmediato
                if (AJefeInmediato == "SI")
                {
                    if (!string.IsNullOrEmpty(idJefeInmediato.Trim()))
                    {
                        selectCommand = "SELECT DJ.CorreoResponsable, JO.idPadre " +
                            "FROM [Parametrizacion].[JerarquiaOrganizacional] AS JO INNER JOIN [Parametrizacion].[DetalleJerarquiaOrg] AS DJ ON DJ.idHijo = JO.idHijo " +
                            "WHERE JO.idHijo = " + idJefeInmediato;

                        dad = new SqlDataAdapter(selectCommand, conString);
                        dtblDiscuss.Clear();
                        dad.Fill(dtblDiscuss);
                        view = new DataView(dtblDiscuss);

                        foreach (DataRowView row in view)
                        {
                            Destinatario = Destinatario + ";" + row["CorreoResponsable"].ToString().Trim();
                            idJefeMediato = row["idPadre"].ToString().Trim();
                        }
                    }
                }
                #endregion

                #region correo del Jefe Mediato
                //Consulta el correo del Jefe Mediato
                if (AJefeMediato == "SI")
                {
                    if (!string.IsNullOrEmpty(idJefeMediato.Trim()))
                    {
                        selectCommand = "SELECT DJ.CorreoResponsable, JO.idPadre " +
                            "FROM [Parametrizacion].[JerarquiaOrganizacional] AS JO INNER JOIN [Parametrizacion].[DetalleJerarquiaOrg] AS DJ ON DJ.idHijo = JO.idHijo " +
                            "WHERE JO.idHijo = " + idJefeMediato;

                        dad = new SqlDataAdapter(selectCommand, conString);
                        dtblDiscuss.Clear();
                        dad.Fill(dtblDiscuss);
                        view = new DataView(dtblDiscuss);

                        foreach (DataRowView row in view)
                        {
                            Destinatario = Destinatario + ";" + row["CorreoResponsable"].ToString().Trim();
                        }
                    }
                }
                #endregion

                #region Correos Enviados
                //Insertar el Registro en la tabla de Correos Enviados
                SqlDataSource200.InsertParameters["Destinatario"].DefaultValue = Destinatario.Trim();
                SqlDataSource200.InsertParameters["Copia"].DefaultValue = Copia;
                SqlDataSource200.InsertParameters["Otros"].DefaultValue = Otros;
                SqlDataSource200.InsertParameters["Asunto"].DefaultValue = Asunto;
                SqlDataSource200.InsertParameters["Cuerpo"].DefaultValue = Cuerpo;
                SqlDataSource200.InsertParameters["Estado"].DefaultValue = "POR ENVIAR";
                SqlDataSource200.InsertParameters["Tipo"].DefaultValue = "CREACION";
                SqlDataSource200.InsertParameters["FechaEnvio"].DefaultValue = System.DateTime.Now.ToString().Trim();
                SqlDataSource200.InsertParameters["IdEvento"].DefaultValue = idEvento.ToString().Trim();
                SqlDataSource200.InsertParameters["IdRegistro"].DefaultValue = idRegistro.ToString().Trim();
                SqlDataSource200.InsertParameters["IdUsuario"].DefaultValue = Session["idUsuario"].ToString().Trim(); //Aca va el id del Usuario de la BD
                SqlDataSource200.InsertParameters["FechaRegistro"].DefaultValue = System.DateTime.Now.ToString().Trim();
                SqlDataSource200.Insert();
                #endregion
            }
            catch (Exception except)
            {
                // Handle the Exception.
                omb.ShowMessage("Error en el envío de la notificación." + "<br/>" + "Descripción: " + except.Message.ToString().Trim(), 1, "Atención");
                err = true;
            }

            if (!err)
            {
                #region Restro
                // Si no existe error en la creacion del registro en el log de correos enviados se procede a escribir en la tabla CorreosRecordatorios y a enviar el correo 
                if (RequiereFechaCierre == "SI" && FechaFinal != "")
                {
                    //Si los NroDiasRecordatorio es diferente de vacio se inserta el registro correspondiente en la tabla CorreosRecordatorio
                    SqlDataSource201.InsertParameters["IdCorreosEnviados"].DefaultValue = LastInsertIdCE.ToString().Trim();
                    SqlDataSource201.InsertParameters["NroDiasRecordatorio"].DefaultValue = NroDiasRecordatorio;
                    SqlDataSource201.InsertParameters["Estado"].DefaultValue = "POR ENVIAR";
                    SqlDataSource201.InsertParameters["FechaFinal"].DefaultValue = FechaFinal;
                    SqlDataSource201.InsertParameters["IdUsuario"].DefaultValue = Session["idUsuario"].ToString().Trim(); //Aca va el id del Usuario de la BD
                    SqlDataSource201.InsertParameters["FechaRegistro"].DefaultValue = System.DateTime.Now.ToString().Trim();
                    SqlDataSource201.Insert();
                }
                #endregion

                try
                {
                    #region Envio Correo
                    MailMessage message = new MailMessage();
                    SmtpClient smtpClient = new SmtpClient();
                    MailAddress fromAddress = new MailAddress(((System.Net.NetworkCredential)(smtpClient.Credentials)).UserName, "Software Sherlock");
                    message.From = fromAddress;//here you can set address

                    #region Destinatario
                    foreach (string substr in Destinatario.Split(';'))
                    {
                        if (!string.IsNullOrEmpty(substr.Trim()))
                        {
                            message.To.Add(substr);
                        }
                    }
                    #endregion

                    #region Copia
                    if (Copia.Trim() != "")
                    {
                        foreach (string substr in Copia.Split(';'))
                        {
                            if (!string.IsNullOrEmpty(substr.Trim()))
                            {
                                message.CC.Add(substr);
                            }
                        }
                    }
                    #endregion

                    #region Otros
                    if (Otros.Trim() != "")
                    {
                        foreach (string substr in Otros.Split(';'))
                        {
                            if (!string.IsNullOrEmpty(substr.Trim()))
                            {
                                message.CC.Add(substr);
                            }
                        }
                    }
                    #endregion

                    message.Subject = Asunto;//subject of email
                    message.IsBodyHtml = true;//To determine email body is html or not
                    message.Body = Cuerpo;

                    smtpClient.Send(message);
                    #endregion
                }
                catch (Exception ex)
                {
                    //throw exception here you can write code to handle exception here
                    omb.ShowMessage("Error en el envío de la notificación." + "<br/>" + "Descripción: " + ex.Message.ToString().Trim(), 1, "Atención");
                    err = true;
                }

                if (!err)
                {
                    //Actualiza el Estado del Correo Enviado
                    SqlDataSource200.UpdateParameters["IdCorreosEnviados"].DefaultValue = LastInsertIdCE.ToString().Trim();
                    SqlDataSource200.UpdateParameters["Estado"].DefaultValue = "ENVIADO";
                    SqlDataSource200.UpdateParameters["FechaEnvio"].DefaultValue = System.DateTime.Now.ToString().Trim();
                    SqlDataSource200.Update();
                }
            }

            return (err);
        }

        protected void SqlDataSource200_On_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            LastInsertIdCE = (int)e.Command.Parameters["@NewParameter2"].Value;
        }
        protected void SqlDataSource201_On_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            LastInsertIdCE = (int)e.Command.Parameters["@NewParameter2"].Value;
        }

        protected void ImbCalcularControl_Click(object sender, ImageClickEventArgs e)
        {

            clsBLLcalificaControl cProcess = new clsBLLcalificaControl();
            string idControlIn = Session["idControlIn"].ToString();
            DataTable dtInfo = cControl.GetLastControl();
            string idControlEnd = string.Empty;
            if (dtInfo.Rows.Count > 0)
                idControlEnd = dtInfo.Rows[0]["LastControl"].ToString().Trim();

            List<cControlEntity> lstControles = cControl.ListControlesById(idControlIn,idControlEnd);
            string strErrMsg = string.Empty;
            Boolean flag = cProcess.mtdCalificaControlMasico(lstControles, ref strErrMsg);
            if (flag == true)
                Mensaje("Proceso finalizado exitosamente");
        }

        protected void Badd_Click(object sender, EventArgs e)
        {
            System.Configuration.Configuration rootWebConfig = System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration("/MyWebSiteRoot");
            //System.Configuration.ConnectionStringSettings connectionString = rootWebConfig.ConnectionStrings.ConnectionStrings["ListasConnectionString"];
            //Read the uploaded File as Byte Array from FileUpload control.
            #region Archivo
            Stream fs = FileUpload1.PostedFile.InputStream;
            BinaryReader br = new BinaryReader(fs);
            Byte[] bPdfData = br.ReadBytes((Int32)fs.Length);
            #endregion Archivo
            //SqlConnection con = new SqlConnection(connectionString.ToString());
            System.Configuration.ConnectionStringSettings SqlconnStr = rootWebConfig.ConnectionStrings.ConnectionStrings["SarlaftConnectionString"];
            SqlConnection sqlCnn = new SqlConnection(SqlconnStr.ToString());
            SqlCommand com = new SqlCommand();
            com.Connection = sqlCnn;
            string nombre = string.Empty;

            if (DDLopciones.SelectedValue == "1")
            {
                nombre = "RiesgosPlantillaCargaMasiva.xls";
            }

            if (DDLopciones.SelectedValue == "2")
            {
                nombre = "ControlesPlantillaCargaMasiva.xls";
            }

            if (DDLopciones.SelectedValue == "3")
            {
                nombre = "EventosCreacionPlantillaCargaMasiva.xls";
            }
            if (DDLopciones.SelectedValue == "5")
            {
                nombre = "EventosDatosComplementariosPlantillaCargaMasiva.xls";
            }
            if (DDLopciones.SelectedValue == "6")
            {
                nombre = "EventosConDatosComplementariosPlantillaCargaMasiva.xls";
            }
            if (DDLopciones.SelectedValue == "4")
            {
                nombre = "RiesgosvsControlesPlantillaCargaMasiva.xls";
            }
            SqlParameter parameter = new SqlParameter("@Data", SqlDbType.VarBinary);
            parameter.Value = bPdfData;
            com.Parameters.Add(parameter);
            com.CommandText = "INSERT INTO [Riesgos].[RiesgosPlantillaMasiva] ([FechaRegistro],[UrlArchivo],[ArchivoExcel]) VALUES " +
           "(GETDATE(), '" + nombre + "',@Data)";
            sqlCnn.Open();
            //insert the file into database
            com.ExecuteNonQuery();
            sqlCnn.Close();
        }
    }
}