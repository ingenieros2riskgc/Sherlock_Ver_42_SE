'use strict';
var cuantos_Checkqueados = 0;

function ocultaResponsable() {
    var etiqueta = document.getElementById("ctl00_ContentPlaceHolder5_Eventos_TabContainerEventos_TabPanelContab_pnlDependencia2");
    etiqueta.style.visibility = "hidden";
}

function CheckItem(obj) {
    var chkControlId = 'ctl00_ContentPlaceHolder5_Riesgos_TabContainer2_TabPanel11_CheckBoxList9';
    var estadoControl = 'ctl00_ContentPlaceHolder5_Riesgos_TabContainer2_TabPanel11_btnEstado';
    var cuantosCheck = obj.rows.length;
    var opt = "";
    var idCheck = "";
    var chequeados = 0;

    var etiqueta = document.getElementById("cajaTratamiento");
    var estado = document.getElementById(estadoControl);

    for (var i = 0; i < cuantosCheck; i++) {
        idCheck = chkControlId + '_' + i;
        opt = document.getElementById(idCheck);

        if (opt.checked) {

            chequeados++;
        }
    }
    if (chequeados > 0) {
        etiqueta.style.visibility = "visible";
        cuantos_Checkqueados = chequeados;
        estado.className = 'Activo';
    }
    else {//ninguno chequeado
        cuantos_Checkqueados = 0;
        etiqueta.style.visibility = "hidden";
        document.getElementById("ctl00_ContentPlaceHolder5_Riesgos_TabContainer2_TabPanel11_txtResponsableT").value = '';
        estado.className = 'Inactivo';

    }
}

function ocultaRespTratamiento(cuantosCheckTratamiento) {
    var etiqueta = document.getElementById("cajaTratamiento");
    var chkControlId = 'ctl00_ContentPlaceHolder5_Riesgos_TabContainer2_TabPanel11_CheckBoxList9';

    var opt = "";
    var idCheck = "";
    var chequeados = 0;

    idCheck = chkControlId + '_';

    for (var i = 0; i < cuantosCheckTratamiento; i++) {
        idCheck = chkControlId + '_' + i;
        opt = document.getElementById(idCheck);

        if (opt.checked) {

            chequeados++;
        }
    }

    if (chequeados > 0) {
        etiqueta.style.visibility = "visible";
    }
    else {
        etiqueta.style.visibility = "hidden";
    }
}

// Valida solo numeros / separador de mil / espacios
function format(input) {
    var num = input.value.replace(/\./g, '');

    if (!isNaN(num)) {
        num = num.toString().split('').reverse().join('').replace(/(?=\d*\.?)(\d{3})/g, '$1.');
        num = num.split('').reverse().join('').replace(/^[\.]/, '');
        input.value = num;
    }
    else {
        input.value = input.value.replace(/[^\d\.]*/g, '');
    }
    input.value = input.value.replace(/^\s+|\s+$/g, "");
}

function SoloNumerosMaxCien(input) {
    var num = input.value.replace(/\./g, '');

    if (!isNaN(num)) {
        num = num.toString().split('').reverse().join('').replace(/(?=\d*\.?)(\d{3})/g, '$1.');
        num = num.split('').reverse().join('').replace(/^[\.]/, '');
        input.value = num;
    }
    else {
        input.value = input.value.replace(/[^\d\.]*/g, '');
    }
    input.value = input.value.replace(/^\s+|\s+$/g, "");
    if (input.value > 100) {
        alert("El valor máximo que puede ingresar en este campo es 100");
        input.value = 100;
    }
}

function limpiarFreqMax() {
    var id = "ctl00_ContentPlaceHolder5_Riesgos_parametroFreqMax";
    var maxFreq = document.getElementById(id);
    maxFreq.value = '';
}

function validaFreq() {
    var id = "ctl00_ContentPlaceHolder5_Riesgos_parametroFreqMax";
    var paraFreq = document.getElementById(id);

    if (paraFreq.value == '') {
        alert("El campo no puede estar vacío, intentelo nuevamente ingresando un número válido.");
        paraFreq.focus();
    }
    else if (paraFreq.value == '00') {
        alert("El valor del campo no es válido, intentelo nuevamente ingresando un número válido diferente a 00.");
        paraFreq.focus();
        paraFreq.value = '';
    }
}

function FocusJustificacion() {
    var elemento = document.getElementById("ctl00_ContentPlaceHolder5_paRiesgoEventoVar_TcPrincipal_tpPlanes_Justificacion");
    alert("Primero, debe justificar los cambios !");
    let coords = elemento.getBoundingClientRect();
    elemento.focus();
    $("html, body").animate({ scrollTop: $(document).height() }, 1000);    
}

function FocusPeriodo(trans) {

    if (trans == 0) {
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_paRiesgoEventoVar_TcPrincipal_tpPlanes_NombrePlan");
        elemento.focus();
        $("html, body").animate({ scrollTop: $(document).height() }, 1000);
    }

    if (trans == 1) {
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_paRiesgoEventoVar_TcPrincipal_tpCumplimiento_Periodo");
        alert("Es necesario llenar el campo Fecha Revisión !");
        elemento.focus();
        $("html, body").animate({ scrollTop: $(document).height() }, 1000);
    }
    else if (trans == 2) {
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_paRiesgoEventoVar_TcPrincipal_tpCumplimiento_Meta");
        alert("Es necesario llenar el campo Meta !");
        elemento.focus();
        $("html, body").animate({ scrollTop: $(document).height() }, 1000);
    } 
    if (trans == 3) {
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_paRiesgoEventoVar_TcPrincipal_tpPlanes_FechaCompromiso");
        alert("El campo Fecha Compromiso no tiene el formato correcto !");
        elemento.focus();
        $("html, body").animate({ scrollTop: 500 }, 50);
    }
    if (trans == 4) {
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_paRiesgoEventoVar_TcPrincipal_tpPlanes_FechaCompromiso");
        alert("La Fecha de Compromiso no puede ser menor a hoy !");
        elemento.focus();
        $("html, body").animate({ scrollTop: $(document).height() }, 1000);              
    }
    if (trans == 5) {
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_paRiesgoEventoVar_ModalGestion");
        alert("El valor supera el 100%, intente con otro !");
        elemento.focus();
    }
    if (trans == 6) {
                                                
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_paRiesgoEventoVar_FiltroCodigoPlan");       
        elemento.focus();
    }
    if (trans == 7) {

        var elementox = document.getElementById("ctl00_ContentPlaceHolder5_paRiesgoEventoVar_FiltroCodigoPlan");
        alert("No se encontraron valores para exportar, intente con otros parámetros de búsqueda!");
        elementox.focus();
    }

    if (trans == 8) {
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_paRiesgoEventoVar_TcPrincipal_tpPlanes_Justificacion");
        elemento.focus();     
        $("html, body").animate({ scrollTop: $(document).height() }, 1000);    
    }

    if (trans == 9) {
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_paRiesgoEventoVar_TcPrincipal_tpPlanes_FUCarga");
        alert("Primero, seleccione un archivo");
        //elemento.focus();
        $("html, body").animate({ scrollTop: 500 }, 50);
    }  
    if (trans == 10) {         
        $("html, body").animate({ scrollTop: $(document).height() }, 1000); 
    }  

    if (trans == 11) {
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_CalificacionExpVarCatVar_ModalNombreVariable");        
        elemento.focus();        
    }  

    if (trans == 12) {
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_CalificacionExpVarCatVar_ModalNombreCategorias");
        elemento.focus();

    } 

    if (trans == 13) {
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_CalificacionExpPuntosDeCorte_ModalMin");
        elemento.focus();

    } 

    if (trans == 14) {
        
        alert("Es necesario realizar el cálculo de la medición antes de registrar el riesgo.");      
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_Riesgos_TabContainer2_TabPanel14_ExpCalcular");
        if (elemento != null) {
            elemento.style.background = 'red';
            elemento.style.color = 'White';
        }       
    }
    if (trans == 15) {
        alert("Excedió la sumatoria total permitida (100%) para la ponderación de la variable. Intente con un número menor!");
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_CalificacionExpVarDeImpacto_PonderacionVariable");
        elemento.focus();
    }
    if (trans == 16) {
        alert("Excedió la sumatoria total permitida (100%) para la ponderación de la variable del impacto.");
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_Riesgos_TabContainer2_TabPanel14_GvVariablesCategoriasImpacto_ctl02_Peso");
        elemento.focus();
    }
    if (trans == 17) {
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_CalificacionExpVarDeImpacto_ModalNombreVariable");
        elemento.focus();
    }  

    if (trans == 18) {
        alert("La sumatoria para el peso del impacto debe dar igual a 100%");
    }  

    if (trans == 19) {
        alert("Es necesario realizar el cálculo de la medición antes de registrar el riesgo.");
        var elemento = document.getElementById("ctl00_ContentPlaceHolder5_Riesgos_TabContainer1_TabPanel1_CalcularFI");
        if (elemento != null) {
            elemento.style.background = 'red';
            elemento.style.color = 'White';
        }       
    }   
}

function Redireccionar() {
    var Elemento = document.getElementById("ctl00_ContentPlaceHolder2_Home_imgInfo");
    var imagen = Elemento.src;
    var encontro = imagen.search("RelojArena.gif");
    if (encontro > 0) {
        var protocolo = window.location.protocol;
        var host = window.location.host;        
        var ruta = protocolo+"//"+host+"/Formularios/Sitio/Login.aspx";
        window.location.replace(ruta);        
    }
}

function tvPrueba() {
    var elemento = document.getElementById("ctl00_ContentPlaceHolder5_paRiesgoEventoVar_TcPrincipal_tpPlanes_PanelResponsable");
    elemento.visibility();
}

function clickHandler(e, v) {

    console.log(e, v);
    // e es el evento
    // v es el boton

}
function LlenarMacroProceso() {

    var obj = document.getElementById("ctl00_ContentPlaceHolder5_Riesgos_Pruebas");
    if (obj) {
        obj.click();
    }    
}

function Paleta() {        
    $('#colorpicker').farbtastic('#ctl00_ContentPlaceHolder5_ClasificacionRiesgo_color');  
    $("html, body").animate({ scrollTop: $(document).height() }, 1000); 
}
