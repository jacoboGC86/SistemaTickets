
function getParameters(url){
    var query = url.split("?")[1];
    var params = query.split("&");
    var oParam = [];

    for(var i = 0; i < params.length; i++)
        oParam.push({
            "param": params[i].split("=")[0],
            "value": params[i].split("=")[1]
        });
    
    return oParam;
}

function validarTextBox() {
    if($(this).val() == null || $(this).val().trim() == "") {
        $(this).addClass("is-invalid");
        $(this).removeClass("is-valid");
    } else {
        $(this).addClass("is-valid");
        $(this).removeClass("is-invalid");
    }
}

function soloLetras(event) {
    var inputValue = event.which;

    // allow letters and whitespaces only.
    if(!(inputValue >= 65 && inputValue <= 120) && (inputValue != 32 && inputValue != 0 && inputValue != 9 && inputValue != 37 && inputValue != 39 && inputValue != 8)) { 
        event.preventDefault(); 
    }
}

function soloNumeros() {
  var numero = numeral($(this).val().trim()).value();

  $(this).val(numeral(numero).format("0"));

  $(this).addClass("is-valid");
  $(this).removeClass("is-invalid");
}

function obtenerFecha(fecha, formato) {
  if(fecha === null)
    return "";

  var dia = fecha.getDate(), mes = fecha.getMonth(), agno = fecha.getFullYear(), ff = "";
  var mesTxt = "", mesTxtShort = "";

  switch(mes) {
    case 0: mesTxt = "Enero"; mesTxtShort = "Ene"; break;
    case 1: mesTxt = "Febrero"; mesTxtShort = "Feb"; break;
    case 2: mesTxt = "Marzo"; mesTxtShort = "Mar"; break;
    case 3: mesTxt = "Abril"; mesTxtShort = "Abr"; break;
    case 4: mesTxt = "Mayo"; mesTxtShort = "May"; break;
    case 5: mesTxt = "Junio"; mesTxtShort = "Jun"; break;
    case 6: mesTxt = "Julio"; mesTxtShort = "Jul"; break;
    case 7: mesTxt = "Agosto"; mesTxtShort = "Ago"; break;
    case 8: mesTxt = "Septiembre"; mesTxtShort = "Sep"; break;
    case 9: mesTxt = "Octubre"; mesTxtShort = "Oct"; break;
    case 10: mesTxt = "Noviembre"; mesTxtShort = "Nov"; break;
    case 11: mesTxt = "Diciembre"; mesTxtShort = "Dic"; break;
  }

  switch(formato) {
    case "yyyymmdd": 
    dia = (dia < 10) ? "0" + dia: dia;
    mes = (mes + 1 < 10) ? "0" + (mes + 1) : (mes + 1);
    agno = agno.toString().substring(2, 4);
    ff = agno + mes + dia;
    break;
    case "ddmmyy": 
    dia = (dia < 10) ? "0" + dia: dia;
    mes = (mes + 1 < 10) ? "0" + (mes + 1) : (mes + 1);
    agno = agno.toString().substring(2, 4);
    ff = dia + mes + agno;
    break;
    case "dd-mm-yyyy": 
    dia = (dia < 10) ? "0" + dia: dia;
    mes = (mes + 1 < 10) ? "0" + (mes + 1) : (mes + 1);
    agno = agno.toString();
    ff = dia + "-" + mes + "-" + agno;
    break;
    case "dd/mm/yyyy": 
    dia = (dia < 10) ? "0" + dia: dia;
    mes = (mes + 1 < 10) ? "0" + (mes + 1) : (mes + 1);
    agno = agno.toString();
    ff = dia + "/" + mes + "/" + agno;
    break;
    case "dd/MMMM/yyyy": 
    dia = (dia < 10) ? "0" + dia: dia;
    agno = agno.toString();
    ff = dia + "/" + mesTxt + "/" + agno;
    break;
    case "dd/MMM/yyyy": 
    dia = (dia < 10) ? "0" + dia: dia;
    agno = agno.toString();
    ff = dia + "/" + mesTxtShort + "/" + agno;
    break;
  }

  return ff;
}

function initializePeoplePicker(peoplePickerElementId, width) {
    var schema = {};
    schema['PrincipalAccountType'] = 'User';
    schema['SearchPrincipalSource'] = 15;
    schema['ResolvePrincipalSource'] = 15;
    schema['AllowMultipleValues'] = false;
    schema['MaximumEntitySuggestions'] = 50;
    schema['Width'] = width != undefined ? width : '370px';

    this.SPClientPeoplePicker_InitStandaloneControlWrapper(peoplePickerElementId, null, schema);
}

function initializePeoplePickerMulti(peoplePickerElementId) {
  var schema = {};
  schema['PrincipalAccountType'] = 'User';
  schema['SearchPrincipalSource'] = 15;
  schema['ResolvePrincipalSource'] = 15;
  schema['AllowMultipleValues'] = true;
  schema['MaximumEntitySuggestions'] = 50;
  schema['Width'] = '370px';

  this.SPClientPeoplePicker_InitStandaloneControlWrapper(peoplePickerElementId, null, schema);
}

function abrirDialogoCarga(event) {
    event.preventDefault();
    tipoDocCargar = $(this).data("doc");
    $("#fileDocto").click();
}

function getFileBuffer() {
    var deferred = jQuery.Deferred();
    var reader = new FileReader();

    reader.onloadend = function (e) {
        deferred.resolve(e.target.result);
    }
    reader.onerror = function (e) {
        deferred.reject(e.target.error);
    }
    reader.readAsArrayBuffer($("#fileDocto")[0].files[0]);

    return deferred.promise();
}

function onRequestFail(sender, args) {
    console.log(sender);
}