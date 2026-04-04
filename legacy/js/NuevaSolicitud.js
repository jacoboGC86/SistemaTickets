var context, contextRH, contextIAD, contextCnv, canvas, cuser, idCEO = 0, idFinanzas = 0, web, idSol = 0, numDias = 0, vacDat = {
  num: 0,
  fi: new Date(),
  agnos: 0,
  diasRest: 0,
  puesto: "",
  area: "",
}, canvasx, canvasy, last_mousex, mousex, mousedown, firmado = false;

$(document).ready(function(){
  var iebrowser = /*@cc_on!@*/false || !!document.documentMode;
  
  if(iebrowser) {
      alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
      window.location.href = "https://chimalliapps.sharepoint.com/sites/Contacto/Vacations/";
  }

  context = new SP.ClientContext.get_current();
  contextRH = new SP.ClientContext("https://chimalliapps.sharepoint.com/sites/Contacto/Vacations/");//"https://allgaiergroup.sharepoint.com/sites/AdP/Vacation/");//"https://allgaiergroup.sharepoint.com/sites/AdP-Paperless/VR/");
  contextIAD = new SP.ClientContext("https://chimalliapps.sharepoint.com/sites/Contacto/Vacations/");
  web = context.get_web();
  cuser = web.get_currentUser();
  $("#dtFechaRegistro").val(moment().format("DD/MMM/yyyy"));

  context.load(web);
  context.load(cuser);  
  context.executeQueryAsync(function(sender, args) {
    $("#txtSolicitante").val(cuser.get_title());
    $("#dtFechaRegistro").val(obtenerFecha(new Date(), "dd/MMM/yyyy"));

    var oList = contextRH.get_web().get_lists().getByTitle("Colaboradores");
    var camlQuery = new SP.CamlQuery();

    camlQuery.set_viewXml("<View><Query><Where>" + 
    "<Eq><FieldRef Name='Cuenta' LookupId='TRUE'/><Value Type='Integer'><UserID /></Value></Eq>" + 
    "</Where></Query><RowLimit>1</RowLimit></View>");
    var collListItem = oList.getItems(camlQuery);
    
    contextRH.load(collListItem);
    contextRH.executeQueryAsync(function(sender, args){
      var listItemEnumerator = collListItem.getEnumerator();

      while (listItemEnumerator.moveNext()) {
        var oListItem = listItemEnumerator.get_current();
        
        $("#txtPuesto").val(oListItem.get_item("Puesto"));
        $("#txtPuestoSub").val(oListItem.get_item("Puesto"));
        $("#txtPuestoSub").data("puestoorg", oListItem.get_item("Puesto"));
        vacDat.puesto = oListItem.get_item("Puesto");
        $("#txtArea").val(oListItem.get_item("Departamento"));
        vacDat.area = oListItem.get_item("Departamento");
        $("#txtNumCol").data("minc", oListItem.get_item("Title"));
      }

      var oListConf = context.get_web().get_lists().getByTitle("Configuraciones");
      var camlQueryConf = new SP.CamlQuery();

      camlQueryConf.set_viewXml("<View><Query><Where></Where></Query></View>");
      var collListItemConf = oListConf.getItems(camlQueryConf);

      context.load(collListItemConf);
      context.executeQueryAsync(function(sender, args){
        var listItemEnumerator = collListItemConf.getEnumerator();

        while (listItemEnumerator.moveNext()) {
          var oListItem = listItemEnumerator.get_current();

          switch(oListItem.get_item("Title")) {
            case "CEO": 
            idCEO = oListItem.get_item("Personas")[0].get_lookupId();
            break;
            case "Finanzas": 
            idFinanzas = oListItem.get_item("Personas")[0].get_lookupId();
            break;
            case "Departamento":
            $("#cmbManager").append("<option value='" + oListItem.get_item("Personas")[0].get_lookupId() + "'>" + oListItem.get_item("Valor") + " - " + oListItem.get_item("Personas")[0].get_lookupValue() + "</option>");
            $("#cmbDepartmentFrom").append("<option value='" + oListItem.get_item("Valor") + "'>" + oListItem.get_item("Valor") + "</option>")
            $("#cmbDepartmentTo").append("<option value='" + oListItem.get_item("Valor") + "'>" + oListItem.get_item("Valor") + "</option>")
            break
          }
        }

        asignarEventos();
      }, onRequestFail);
    }, onRequestFail);
  }, onRequestFail);
});

function asignarEventos() {
  canvas = document.getElementById('sheet');
  contextCnv = canvas.getContext("2d");

  last_mousex = last_mousey = 0;
  mousex = mousey = 0;
  mousedown = false;

  $(canvas).on('mousedown', function(e) {
    last_mousex = mousex = parseInt(e.clientX - canvasx);
	  last_mousey = mousey = parseInt(e.clientY - canvasy);
    mousedown = true;
  });

  //Mouseup
  $(canvas).on('mouseup', function(e) {
    mousedown = false;
  });

  //Mousemove
  $(canvas).on('mousemove', function(e) {
    canvasx = $(canvas).offset().left;
    canvasy = $(canvas).offset().top;
    mousex = parseInt(e.clientX - canvasx);
    mousey = parseInt(e.clientY - canvasy);

    if(mousedown) {
      contextCnv.beginPath();

      contextCnv.globalCompositeOperation = 'source-over';
      contextCnv.strokeStyle = 'black';
      contextCnv.lineWidth = 3;

      contextCnv.moveTo(last_mousex, last_mousey);
      contextCnv.lineTo(mousex, mousey);
      contextCnv.lineJoin = contextCnv.lineCap = 'round';
      contextCnv.stroke();
      firmado = true;
    }

    last_mousex = mousex;
    last_mousey = mousey;
  });

  $("#btnBorrar").click(function(){
    firmado = false;
    contextCnv.clearRect(0, 0, canvas.width, canvas.height);
  });

  $("#btnCargar").click(function(){
    $("#fileDocto").click();
  });
  $("#fileDocto").on("change", function(){
    if($(this).val() === "") {
      return;
    }

    $("#ulFormat").html("");

    for(var i = 0; i < $("#fileDocto")[0].files.length; i++)
      $("#ulFormat").append("<li class='list-group-item'>" + $("#fileDocto")[0].files[i].name + "</li>")
  });
  
  $("#txtSalaryIncreasePM").on("change", soloNumeros);
  $("#txtSalaryIncreaseNM").on("change", soloNumeros);
  $("#txtSalaryIncreasePD").on("change", soloNumeros);
  $("#txtSalaryIncreaseND").on("change", soloNumeros);
  $("#txtJustification").on("change", validarTextBox);
  
  $("#cmbTipoSol").on("change", function(){
    if($("#cmbTipoSol").val() === "Yo") {
      $("#dvBuscarCol").hide();
      $("#txtPuestoSub").val($("#txtPuestoSub").data("puestoorg"));
      $(".dvBeneficiario").hide();
    } else {
      $("#dvBuscarCol").show();
      $(".dvBeneficiario").show();
      $("#txtPuestoSub").val("");
      $("#txtNumCol").val("");
      $("#txtSolicitanteSub").val("");
    }
  });
  $("input[type='checkbox']").on("change", function(){
    var id = $(this).prop("id").replace("chk", "");
    var disp = !$(this).is(":checked");

    $("#dt" + id).prop("disabled", disp);
    $("#cmb" + id + "From").prop("disabled", disp);
    $("#cmb" + id + "To").prop("disabled", disp);
    $("#txt" + id + "From").prop("disabled", disp);
    $("#txt" + id + "To").prop("disabled", disp);
    $("#txt" + id + "PM").prop("disabled", disp);
    $("#txt" + id + "NM").prop("disabled", disp);
    $("#txt" + id + "PD").prop("disabled", disp);
    $("#txt" + id + "ND").prop("disabled", disp);

    if(id === "SalaryIncrease") {
      $("#cmbCurrentCurrency").prop("disabled", disp);
      $("#cmbNewCurrency").prop("disabled", disp);
    }
  });
  $("#lkBuscar").click(function(e){
    e.preventDefault();

    var oList = contextRH.get_web().get_lists().getByTitle("Colaboradores");
    var camlQuery = new SP.CamlQuery();

    camlQuery.set_viewXml("<View><Query><Where>" + 
    "<Eq><FieldRef Name='Title' /><Value Type='Text'>" + $("#txtNumCol").val() + "</Value></Eq>" + 
    "</Where></Query><RowLimit>1</RowLimit></View>");
    var collListItem = oList.getItems(camlQuery);
    
    contextRH.load(collListItem);
    contextRH.executeQueryAsync(function(sender, args){
      var listItemEnumerator = collListItem.getEnumerator();
      var encontrado = false;

      while (listItemEnumerator.moveNext()) {
        encontrado = true;
        var oListItem = listItemEnumerator.get_current();
        
        $("#txtNumCol").val(oListItem.get_item("Title"));
        $("#txtSolicitanteSub").val(oListItem.get_item("Nombre"));
        $("#txtPuestoSub").val(oListItem.get_item("Puesto"));
        $("#txtArea").val(oListItem.get_item("Departamento"));
        $("#txtFechaIngreso").val(oListItem.get_item("FechaIngreso").toISOString().split("T")[0]);
        $("#txtDiasDisponibles").val(oListItem.get_item("DiasRestantes"));
        $("#txtDiasDisponiblesRest").val(oListItem.get_item("DiasRestantes"));

        var years = new Date(new Date() - oListItem.get_item("FechaIngreso")).getFullYear() - 1970;

        $("#txtAgnosServ").val(years);

        numDias = 0;
        $("#txtDias").val("0");
        $("#ulDias").html("");
        $("#dtDia").val("");
      }

      if(!encontrado){
        alertify.alert("Employee Movement", "No pudimos encontrar al colaborador. ¿Podría revisar el Número de colaborador por favor?")
      }
    }, onRequestFail);
  });
  $("input[type='checkbox']").change(function(e){
    var id = $(e.target).prop("id").replace("chk", "txt");
    var bloqueado = $(e.target).is(":checked");

    $("#" + id).prop("disabled", !bloqueado);
  });
  $("#btnCancelar").on("click", function(){
    window.location.href = web.get_url();
  });
  $("#btnCargarFirma").on("click", function(){
    $("#fileDocto").click();
  });
  $("#fileDocto").on("change", function(e) {
    if($("#fileDocto").val() == "")
      return;
    
    firmado = true;

    var img = new Image;

    img.onload = function() {
      contextCnv.clearRect(0, 0, canvas.width, canvas.height);
      var hRatio = canvas.width / img.width;
      var vRatio = canvas.height / img.height;
      var ratio = Math.min(hRatio, vRatio);
      var centerShift_x = (canvas.width -img.width * ratio) /2;
      var centerShift_y = (canvas.height -img.height * ratio) /2;
      
      contextCnv.drawImage(img, 0, 0, img.width, img.height, centerShift_x, centerShift_y, img.width*ratio, img.height*ratio);
      $("#fileDocto").val("");
    }

    img.src = URL.createObjectURL(e.target.files[0]);
  });
  $("#btnEnviar").on("click", function() {
    var tipoSol = $("#cmbTipoSol").val();
    var valida = false;
    var numEmp = "", nombre = "", puesto = "";

    if(tipoSol != "Yo") {
      numEmp = $("#txtNumCol").val();
      nombre = $("#txtSolicitanteSub").val();
      puesto = $("#txtPuestoSub").val();
      valida = $("#txtNumCol").val() != "" && nombre != "" && puesto != "";
    } else {
      valida = true;
      numEmp = $("#txtNumCol").data("minc");
    }

    valida = valida && ($("#chkIndefiniteContract").is(":checked") || $("#chkTypeEmployee").is(":checked") || $("#chkTypePosition").is(":checked") || $("#chkDepartment").is(":checked") || $("#chkSalaryIncrease").is(":checked"));
    valida = valida && (!$("#chkIndefiniteContract").is(":checked") || $("#chkIndefiniteContract").is(":checked") && $("#dtIndefiniteContract").val() != "");
    valida = valida && (!$("#chkTypeEmployee").is(":checked") || $("#chkTypeEmployee").is(":checked") && $("#dtTypeEmployee").val() != "");
    valida = valida && (!$("#chkTypePosition").is(":checked") || $("#chkTypePosition").is(":checked") && $("#txtTypePositionFrom").val() != "" && $("#txtTypePositionTo").val() != "" && $("#dtTypePosition").val() != "");
    valida = valida && (!$("#chkDepartment").is(":checked") || $("#chkDepartment").is(":checked") && $("#dtDepartment").val() != "");
    valida = valida && (!$("#chkSalaryIncrease").is(":checked") || $("#chkSalaryIncrease").is(":checked") && $("#txtSalaryIncreasePM").val() != "" && $("#txtSalaryIncreaseNM").val() != "" && $("#txtSalaryIncreasePD").val() != "" && $("#txtSalaryIncreaseND").val() != "" && $("#dtSalaryIncrease").val() != "");

    if(!valida || !firmado) {
      alertify.alert("Employee Movement", "Por favor, llene los campos del movimiento marcado y firme la solicitud.");
      return;
    }

    if($("#cmbManager").val() === null) {
      alertify.alert("Employee Movement", "Por favor, seleccione el área que aprobará el movimiento.");
      return;
    }

    $("#btnEnviar").prop("disabled", true);
    var oList = context.get_web().get_lists().getByTitle("Request");
    var itemCreateInfo = new SP.ListItemCreationInformation();
    var oListItem = oList.addItem(itemCreateInfo);

    oListItem.set_item("Title", $("#txtArea").val());
    oListItem.set_item("TipoSol", tipoSol);
    oListItem.set_item("TypeEmployee", $("#cmbTypeEmployee").val());
    oListItem.set_item("NumCol", numEmp);
    oListItem.set_item("ManagerDepto", parseInt($("#cmbManager").val()));
    oListItem.set_item("ManagerCEO", idCEO);
    oListItem.set_item("ManagerFinanzas", idFinanzas);
    oListItem.set_item("SolicitanteSub", $("#txtSolicitanteSub").val());
    oListItem.set_item("PuestoSub",$("#txtPuestoSub").val());

    oListItem.set_item("ChkIndefiniteContract", $("#chkIndefiniteContract").is(":checked") ? 1 : 0);
    oListItem.set_item("ChkTypeEmployee", $("#chkTypeEmployee").is(":checked") ? 1 : 0);
    oListItem.set_item("ChkTypePosition", $("#chkTypePosition").is(":checked") ? 1 : 0);
    oListItem.set_item("ChkDepartment", $("#chkDepartment").is(":checked") ? 1 : 0);
    oListItem.set_item("ChkSalaryIncrease", $("#chkSalaryIncrease").is(":checked") ? 1 : 0);

    if($("#chkIndefiniteContract").is(":checked")) {
      oListItem.set_item("DtIndefiniteContract", $("#dtIndefiniteContract").val() + 'T06:00:00.000Z');
    }
    if($("#chkTypeEmployee").is(":checked")) {
      oListItem.set_item("TypeEmployeeFrom", $("#cmbTypeEmployeeFrom").val());
      oListItem.set_item("TypeEmployeeTo", $("#cmbTypeEmployeeTo").val());
      oListItem.set_item("DtTypeEmployee", $("#dtTypeEmployee").val() + 'T06:00:00.000Z');
    }
    if($("#chkTypePosition").is(":checked")) {
      oListItem.set_item("TypePositionFrom", $("#txtTypePositionFrom").val());
      oListItem.set_item("TypePositionTo", $("#txtTypePositionTo").val());
      oListItem.set_item("DtTypePosition", $("#dtTypePosition").val() + 'T06:00:00.000Z');
    }
    if($("#chkDepartment").is(":checked")) {
      oListItem.set_item("DepartmentFrom", $("#cmbDepartmentFrom").val());
      oListItem.set_item("DepartmentTo", $("#cmbDepartmentTo").val());
      oListItem.set_item("DtDepartment", $("#dtDepartment").val() + 'T06:00:00.000Z');
    }
    if($("#chkSalaryIncrease").is(":checked")) {
      oListItem.set_item("SalaryIncreasePM", $("#txtSalaryIncreasePM").val());
      oListItem.set_item("SalaryIncreaseNM", $("#txtSalaryIncreaseNM").val());
      oListItem.set_item("SalaryIncreasePD", $("#txtSalaryIncreasePD").val());
      oListItem.set_item("SalaryIncreaseND", $("#txtSalaryIncreaseND").val());
      oListItem.set_item("DtSalaryIncrease", $("#dtSalaryIncrease").val() + 'T06:00:00.0000Z');
      oListItem.set_item("MonedaActual", $("#cmbCurrentCurrency").val());
      oListItem.set_item("MonedaFinal", $("#cmbNewCurrency").val());
    }

    oListItem.set_item("Comentarios", $("#txtComentarios").val());
    oListItem.set_item("FirmaSolicitante", canvas.toDataURL());
    oListItem.set_item("Status", "En Revisión de Área");
    
    oListItem.update();

    context.load(oListItem);
    context.executeQueryAsync(function(sender, args) {
      idSol = oListItem.get_id();

      var archivos = $("#fileDocto")[0].files;
      var spFiles = [];

      subirArchivos(archivos, 0, spFiles, web.get_url(), web.get_serverRelativeUrl(), idSol, function(){
        alertify.alert("Employee Movement", "Solicitud realizada correctamente.", function(){
          window.location.href = "Solicitudes en proceso.aspx?idsol=" + oListItem.get_id();
        });
      });
    }, onRequestFail);
  });
}

function subirArchivos(archivos, indice, spFiles, url, serverRelativeUrl, idSolicitud, callback){
  if(archivos.length > indice) {
    var getFile = getFileBuffer(archivos[indice]);

    getFile.done(function (arrayBuffer) {
      // Add the file to the SharePoint folder.
      var fileName = archivos[indice].name;
      var fileCollectionEndpoint = serverRelativeUrl + "/_api/web/lists/GetByTitle('Request')/items(" + idSol + ")/AttachmentFiles/add(FileName='" + fileName + "')";

      $.ajax({
        url: fileCollectionEndpoint,
        type: "POST",
        data: arrayBuffer,
        processData: false,
        headers: {
          "accept": "application/json;odata=verbose",
          "X-RequestDigest": $("#__REQUESTDIGEST").val(),
          "content-length": arrayBuffer.byteLength
        },
        success: function (file, status, xhr) {
          spFiles.push(file);
          indice++;
          subirArchivos(archivos, indice, spFiles, url, serverRelativeUrl, idSolicitud, callback);
        },
        error: function () {
          indice++;
          subirArchivos(archivos, indice, spFiles, url, serverRelativeUrl, idSolicitud, callback);
        }
      });
    });

    getFile.fail(function(){
      indice++;
      subirArchivos(archivos, indice, spFiles, url, serverRelativeUrl, idSolicitud, callback);
    });
  } else {
    callback(spFiles);
  }
}