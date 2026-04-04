var context, contextCol, cuser, web, idcode, subiendo = false, tipoDocCargar;
var departamento = 0, doc = 0, process = [], category = 0, status = "", statusReal = "";
var autocompleteConfig = {
  minLength: 3,
  source: (request, response) => {
    console.log(request.term);
    var oList = contextCol.get_web().get_lists().getByTitle("Colaboradores");
    var camlQuery = new SP.CamlQuery();

    camlQuery.set_viewXml("<View><Query><Where>" + 
    "<Contains><FieldRef Name='Title' /><Value Type='Text'>" + request.term + "</Value></Contains>" + 
    "</Where></Query><RowLimit>10</RowLimit></View>");
    var collListItem = oList.getItems(camlQuery);
    
    contextCol.load(collListItem);
    contextCol.executeQueryAsync(function(sender, args){
      var listItemEnumerator = collListItem.getEnumerator();

      var colaboradores = [];

      while (listItemEnumerator.moveNext()) {
        var oListItem = listItemEnumerator.get_current();

        colaboradores.push({
          label: oListItem.get_item("Title") + "-" + oListItem.get_item("Nombre") + " / " + oListItem.get_item("Puesto"),
          value: oListItem.get_item("Title") + "-" + oListItem.get_item("Nombre") + " / " + oListItem.get_item("Puesto")
        });
      }

      response(colaboradores);
    }, onRequestFail);
  },
  select: (event, ui) => {
    $(event.target).val(ui.item.label.split("-")[0]);
    $($(event.target).context.parentNode).parent().find(".lblNombre").text(ui.item.label.split("-")[1]);

    return false;
  }
}

$(document).ready(function(){
  $("body").on("contextmenu", "img", function(e) {
    return false;
  });
  var iebrowser = /*@cc_on!@*/false || !!document.documentMode;
  
  if(iebrowser) {
      alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
      window.location.href = "https://chimalliapps.sharepoint.com/sites/Contacto/";
  }

  moment.locale('Es');

  context = new SP.ClientContext.get_current();
  contextCol = new SP.ClientContext("https://chimalliapps.sharepoint.com/sites/Contacto/Vacations/");
  web = context.get_web();
  cuser = web.get_currentUser();

  context.load(web);
  context.load(cuser);  
  context.executeQueryAsync(function(sender, args) {
    var params = getParameters(window.location.href);

    for(var i = 0; i < params.length; i++)
        if(params[i].param === "idsol") {
            idcode = params[i].value;
            break;
        }
    
    var oList = context.get_web().get_lists().getByTitle("Request");
    var oListItem = oList.getItemById(idcode);

    context.load(oListItem);
    context.executeQueryAsync(function(sender, args) {
      if(cuser.get_id() === oListItem.get_item("Author").get_lookupId() && (oListItem.get_item("StatusReal") != "Registro aprobado" && oListItem.get_item("StatusReal") != "Solicitud cancelada/rechazada")) {
        $("#btnCancelar").show();
      }

      $("#txtFolio").val(oListItem.get_id());
      $("#dtFechaRegistro").val(obtenerFecha(oListItem.get_item("Created"), "dd/mm/yyyy"));
      $("#txtSolicitante").val(oListItem.get_item("Author").get_lookupValue());
      $("#txtSolicitante").prop("title", oListItem.get_item("Author").get_email());
      $("#txtSolicitante").data("id", oListItem.get_item("Author").get_lookupId());
      $("#lblSolFirm").text(oListItem.get_item("Author").get_lookupValue());
      $("#cmbAreaSolicitante").val(oListItem.get_item("AreaSolicitante"));
      $("#cmbAreaImplicada").val(oListItem.get_item("AreaGeneradora"));
      $(".txtSemana").text(oListItem.get_item("Semana"));
      $("#txtMotivo").val(oListItem.get_item("Motivo"))
      
      $("#dvFirma").prop("src", oListItem.get_item("SignSolicitante"));

      $("#lblPrdFirm").text(oListItem.get_item("RespAreaSolicitante").get_lookupValue());
      $("#lblManFirm").text(oListItem.get_item("RespAreaGeneradora").get_lookupValue());

      if(oListItem.get_item("SignPlanSolicitante") != null) {
        $("#dvPrdFirma").prop("src", oListItem.get_item("SignPlanSolicitante"));
      }
      if(oListItem.get_item("SignGeneradora") != null) {
        $("#dvFirmMan").prop("src", oListItem.get_item("SignGeneradora"));
      }
      
      if(oListItem.get_item("PlanRH") != null) {
        if(oListItem.get_item("SignPlanRH") != null)
          $("#dvRHFirma").prop("src", oListItem.get_item("SignPlanRH"));
        $("#dvA2PM").html("En Recursos Humanos <br/>" + oListItem.get_item("PlanRH").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FPlanRH"), "dd/mm/yyyy"));
        $("#lblRHFirm").text(oListItem.get_item("PlanRH").get_lookupValue());
      }
      if(oListItem.get_item("SignRealSolicitante") != null) {
        if(oListItem.get_item("SignRealSolicitante") != null)
          $("#dvPrdFirmaR").prop("src", oListItem.get_item("SignRealSolicitante"));
        $("#dvPrd_2").html("Área Solicitante <br/>" + oListItem.get_item("RespRealAreaSolicitante").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespRealAreaSolicitante"), "dd/mm/yyyy"));
      }
      if(oListItem.get_item("AutAreaGeneradora") != null) {
        if(oListItem.get_item("SignAutGeneradora") != null)
          $("#dvFirmManR").prop("src", oListItem.get_item("SignAutGeneradora"));
        $("#lblManFirmR").text(oListItem.get_item("AutAreaGeneradora").get_lookupValue());
        $("#dvA1_2").html("Área Generadora <br/>" + oListItem.get_item("AutAreaGeneradora").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FAutAreaGeneradora"), "dd/mm/yyyy"));
      }

      if(oListItem.get_item("RealRH") != null) {
        if(oListItem.get_item("SignRealRH") != null)
          $("#dvRHFirmaR").prop("src", oListItem.get_item("SignRealRH"));
        $("#lblRHFirmR").text(oListItem.get_item("RealRH").get_lookupValue());
        $("#dvA2AM_2").html("En Recursos Humanos <br/>" + oListItem.get_item("RealRH").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRealRH"), "dd/mm/yyyy"));
      }

      var tabla = "tbPlan";
      status = oListItem.get_item("Status");
      statusReal = oListItem.get_item("StatusReal");
      
      switch(status) {
        case "En Firma de Área Solicitante":
          $("#dvPrd").html("Área Solicitante <br/>" + oListItem.get_item("RespAreaSolicitante").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespAreaSolicitante"), "dd/mm/yyyy"));
          $("#dvPrd").addClass("pendienteRev");
        break;
        case "En Firma de Área Generadora": 
          $("#dvPrd").addClass("aprobadoRev");
          $("#dvPrd").html("Área Solicitante <br/>" + oListItem.get_item("RespAreaSolicitante").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespAreaSolicitante"), "dd/mm/yyyy"));
          $("#dvA1").addClass("pendienteRev");
          $("#dvA1").html("Área Generadora <br/>" + oListItem.get_item("RespAreaGeneradora").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespAreaGeneradora"), "dd/mm/yyyy"));
        break;
        case "En Recursos Humanos":
          $("#dvPrd").addClass("aprobadoRev");
          $("#dvPrd").html("Área Solicitante <br/>" + oListItem.get_item("RespAreaSolicitante").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespAreaSolicitante"), "dd/mm/yyyy"));
          $("#lblPrdFirm").text(oListItem.get_item("RespAreaSolicitante").get_lookupValue());
          $("#dvA1").addClass("aprobadoRev");
          $("#dvA1").html("Área Generadora <br/>" + oListItem.get_item("RespAreaGeneradora").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespAreaGeneradora"), "dd/mm/yyyy"));
          $("#dvA2PM").addClass("pendienteRev");
          break;
        case "Solicitud aprobada":
          $("#dvA1").addClass("aprobadoRev");
          $("#dvA2PM").addClass("aprobadoRev");
          $("#dvPrd").addClass("aprobadoRev");
          $("#dvPrd").html("Área Solicitante <br/>" + oListItem.get_item("RespAreaSolicitante").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespAreaSolicitante"), "dd/mm/yyyy"));
          $("#lblPrdFirm").text(oListItem.get_item("RespAreaSolicitante").get_lookupValue());
          $("#dvA1").html("Área Generadora <br/>" + oListItem.get_item("RespAreaGeneradora").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespAreaGeneradora"), "dd/mm/yyyy"));
          $("#txtComentarios").hide();
          $("label[for='txtComentarios']").hide();
          $("#tbPlan").hide();
          $("#tbReal").show();
          $("#dvStatus2").show();
          tabla = "tbReal";
          if(cuser.get_id() === oListItem.get_item("Author").get_lookupId() && status === "Solicitud aprobada" && (oListItem.get_item("StatusReal") == "Pendiente" || oListItem.get_item("StatusReal") == null)){
            $("#btnAgregarRecurso").show();
            $("#btnEnviar").show();
          }

          if(statusReal === "En Área Solicitante") {
            $("#dvPrd_2").addClass("pendienteRev");
          } else if(statusReal === "En Área Generadora") {
            $("#dvPrd_2").addClass("aprobadoRev");
            $("#dvA1_2").addClass("pendienteRev");
            $("#lblPrdFirmR").text(oListItem.get_item("RespRealAreaSolicitante").get_lookupValue());
          } else if(statusReal === "En Recursos Humanos") {
            $("#dvPrd_2").addClass("aprobadoRev");
            $("#dvA1_2").addClass("aprobadoRev");
            $("#dvA2AM_2").addClass("pendienteRev");
            $("#dvPrd").html("Área Solicitante <br/>" + oListItem.get_item("RespRealAreaSolicitante").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespRealAreaSolicitante"), "dd/mm/yyyy"));
            $("#lblPrdFirmR").text(oListItem.get_item("RespRealAreaSolicitante").get_lookupValue());
          } else if(statusReal === "Registro aprobado") {
            $("#lblPrdFirmR").text(oListItem.get_item("RespRealAreaSolicitante").get_lookupValue());
            $("#dvPrd_2").addClass("aprobadoRev");
            $("#dvA1_2").addClass("aprobadoRev");
            $("#dvA2AM_2").addClass("aprobadoRev");
          }
          break;
        case "Solicitud cancelada/rechazada": 
          $("#txtComentarios").hide();
          $("#label[for='txtComentarios']").hide();

          if(oListItem.get_item("FRespRealAreaSolicitante") === null) {
            if(oListItem.get_item("FPlanRH") != null) {
              $("#dvPrd").html("Área Solicitante <br/>" + oListItem.get_item("RespAreaSolicitante").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespAreaSolicitante"), "dd/mm/yyyy"));
              $("#dvA1").html("Área Generadora <br/>" + oListItem.get_item("RespAreaGeneradora").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespAreaGeneradora"), "dd/mm/yyyy"));
              $("#dvPrd").addClass("aprobadoRev");
              $("#dvA1").addClass("aprobadoRev");
              $("#dvA2PM").addClass("rechazadoRev");
            } else if(oListItem.get_item("FRespAreaGeneradora") != null) {
              $("#dvPrd").html("Área Solicitante <br/>" + oListItem.get_item("RespAreaSolicitante").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespAreaSolicitante"), "dd/mm/yyyy"));
              $("#dvA1").html("Área Generadora <br/>" + oListItem.get_item("RespAreaGeneradora").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespAreaGeneradora"), "dd/mm/yyyy"));
              $("#dvPrd").addClass("aprobadoRev");
              $("#dvA1").addClass("rechazadoRev");
            } else {
              $("#dvPrd").html("Área Solicitante <br/>" + oListItem.get_item("RespAreaSolicitante").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespAreaSolicitante"), "dd/mm/yyyy"));
              $("#dvPrd").addClass("rechazadoRev");
            }
          } else {
            $("#dvStatus2").show();
            $("#dvPrd").html("Área Solicitante <br/>" + oListItem.get_item("RespAreaSolicitante").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespAreaSolicitante"), "dd/mm/yyyy"));
            $("#dvA1").html("Área Generadora <br/>" + oListItem.get_item("RespAreaGeneradora").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespAreaGeneradora"), "dd/mm/yyyy"));
            $("#dvPrd").addClass("aprobadoRev");
            $("#dvA1").addClass("aprobadoRev");
            $("#dvA2PM").addClass("aprobadoRev");
            
            if(oListItem.get_item("FRealRH") != null) {
              $("#dvPrd_2").html("Área Solicitante <br/>" + oListItem.get_item("RespRealAreaSolicitante").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespRealAreaSolicitante"), "dd/mm/yyyy"));
              $("#dvA1_2").html("Área Generadora <br/>" + oListItem.get_item("AutAreaGeneradora").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FAutAreaGeneradora"), "dd/mm/yyyy"));

              $("#dvPrd_2").addClass("aprobadoRev");
              $("#dvA1_2").addClass("aprobadoRev");
              $("#dvA2AM_2").addClass("rechazadoRev");
            } else if(oListItem.get_item("FAutAreaGeneradora") != null) {
              $("#dvPrd_2").html("Área Solicitante <br/>" + oListItem.get_item("RespRealAreaSolicitante").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespRealAreaSolicitante"), "dd/mm/yyyy"));
              $("#dvA1_2").html("Área Generadora <br/>" + oListItem.get_item("AutAreaGeneradora").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FAutAreaGeneradora"), "dd/mm/yyyy"));
              $("#dvPrd_2").addClass("aprobadoRev");
              $("#dvA1_2").addClass("rechazadoRev");
            } else {
              $("#dvPrd_2").html("Área Solicitante <br/>" + oListItem.get_item("RespRealAreaSolicitante").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FRespRealAreaSolicitante"), "dd/mm/yyyy"));
              $("#dvPrd_2").addClass("rechazadoRev");
            }
          }
          break;
      }

      var detalle = JSON.parse(oListItem.get_item("Detalle"));
      var tL = 0, tM = 0, tX = 0, tJ = 0, tV = 0, tS = 0, tD = 0;
      var tLr = 0, tMr = 0, tXr = 0, tJr = 0, tVr = 0, tSr = 0, tDr = 0;
      
      $("#" + tabla + " tbody").html("");
      var deshabilitar = oListItem.get_item("StatusReal") == "Pendiente" || oListItem.get_item("StatusReal") == null ? "" : "disabled";

      for(var i = 0; i < detalle.length; i++) {
        var d = detalle[i];
        
        $("#" + tabla + " thead tr th#th0").text(d.fl);
        $("#" + tabla + " thead tr th#th1").text(d.fm);
        $("#" + tabla + " thead tr th#th2").text(d.fx);
        $("#" + tabla + " thead tr th#th3").text(d.fj);
        $("#" + tabla + " thead tr th#th4").text(d.fv);
        $("#" + tabla + " thead tr th#th5").text(d.fs);
        $("#" + tabla + " thead tr th#th6").text(d.fd);

        var total = parseFloat(d.l != "" ? d.l : 0) + 
                    parseFloat(d.m != "" ? d.m : 0) + 
                    parseFloat(d.x != "" ? d.x : 0) + 
                    parseFloat(d.j != "" ? d.j : 0) + 
                    parseFloat(d.v != "" ? d.v : 0) + 
                    parseFloat(d.s != "" ? d.s : 0) + 
                    parseFloat(d.d != "" ? d.d : 0);

        tL += parseFloat(d.l != "" ? d.l : 0);
        tM += parseFloat(d.m != "" ? d.m : 0);
        tX += parseFloat(d.x != "" ? d.x : 0);
        tJ += parseFloat(d.j != "" ? d.j : 0);
        tV += parseFloat(d.v != "" ? d.v : 0);
        tS += parseFloat(d.s != "" ? d.s : 0);
        tD += parseFloat(d.d != "" ? d.d : 0);

        tLr += parseFloat(d.lr != undefined && d.lr != "" ? d.lr : 0);
        tMr += parseFloat(d.mr != undefined && d.mr != "" ? d.mr : 0);
        tXr += parseFloat(d.xr != undefined && d.xr != "" ? d.xr : 0);
        tJr += parseFloat(d.jr != undefined && d.jr != "" ? d.jr : 0);
        tVr += parseFloat(d.vr != undefined && d.vr != "" ? d.vr : 0);
        tSr += parseFloat(d.sr != undefined && d.sr != "" ? d.sr : 0);
        tDr += parseFloat(d.dr != undefined && d.dr != "" ? d.dr : 0);

        var totalR = 0;

        if(deshabilitar == "disabled") {
          totalR = parseFloat(d.lr != undefined && d.lr != "" ? d.lr : 0) + 
                    parseFloat(d.mr != undefined && d.mr != "" ? d.mr : 0) + 
                    parseFloat(d.xr != undefined && d.xr != "" ? d.xr : 0) + 
                    parseFloat(d.jr != undefined && d.jr != "" ? d.jr : 0) + 
                    parseFloat(d.vr != undefined && d.vr != "" ? d.vr : 0) + 
                    parseFloat(d.sr != undefined && d.sr != "" ? d.sr : 0) + 
                    parseFloat(d.dr != undefined && d.dr != "" ? d.dr : 0);
        }

        $("#" + tabla + " tbody").append("<tr class='trPlaneados'>" + 
          "<td></td>" +
          "<td>" + d.numcol + "</td>" + 
          "<td>" + d.nombre + "</td>" + 
          "<td class='lblH'>" + d.l + "</td>" + 
          (status === "Solicitud aprobada" ? '<td class="lblrH"><input type="number" ' + (deshabilitar == "disabled" ? "value='" + (d.lr != undefined ? d.lr : 0) + "'" : "") + ' class="txtGrid txt0" min="0" max="16" step="1" placeholder="0" ' + deshabilitar + '/></td>' : '') + 
          "<td class='lblH'>" + d.m + "</td>" + 
          (status === "Solicitud aprobada" ? '<td class="lblrH"><input type="number" ' + (deshabilitar == "disabled" ? "value='" + (d.mr != undefined ? d.mr : 0) + "'" : "") + ' class="txtGrid txt1" min="0" max="16" step="1" placeholder="0" ' + deshabilitar + '/></td>' : '') + 
          "<td class='lblH'>" + d.x + "</td>" + 
          (status === "Solicitud aprobada" ? '<td class="lblrH"><input type="number" ' + (deshabilitar == "disabled" ? "value='" + (d.xr != undefined ? d.xr : 0) + "'" : "") + ' class="txtGrid txt2" min="0" max="16" step="1" placeholder="0" ' + deshabilitar + '/></td>' : '') + 
          "<td class='lblH'>" + d.j + "</td>" + 
          (status === "Solicitud aprobada" ? '<td class="lblrH"><input type="number" ' + (deshabilitar == "disabled" ? "value='" + (d.jr != undefined ? d.jr : 0) + "'" : "") + ' class="txtGrid txt3" min="0" max="16" step="1" placeholder="0" ' + deshabilitar + '/></td>' : '') + 
          "<td class='lblH'>" + d.v + "</td>" + 
          (status === "Solicitud aprobada" ? '<td class="lblrH"><input type="number" ' + (deshabilitar == "disabled" ? "value='" + (d.vr != undefined ? d.vr : 0) + "'" : "") + ' class="txtGrid txt4" min="0" max="16" step="1" placeholder="0" ' + deshabilitar + '/></td>' : '') + 
          "<td class='lblH'>" + d.s + "</td>" + 
          (status === "Solicitud aprobada" ? '<td class="lblrH"><input type="number" ' + (deshabilitar == "disabled" ? "value='" + (d.sr != undefined ? d.sr : 0) + "'" : "") + ' class="txtGrid txt5" min="0" max="16" step="1" placeholder="0" ' + deshabilitar + '/></td>' : '') + 
          "<td class='lblH'>" + d.d + "</td>" + 
          (status === "Solicitud aprobada" ? '<td class="lblrH"><input type="number" ' + (deshabilitar == "disabled" ? "value='" + (d.dr != undefined ? d.dr : 0) + "'" : "") + ' class="txtGrid txt6" min="0" max="16" step="1" placeholder="0" ' + deshabilitar + '/></td>' : '') + 
          "<td class='lblTotal'>" + total +"</td>" + 
          (status === "Solicitud aprobada" ? '<td class="lblrH"><input type="number" ' + (deshabilitar == "disabled" ? "value='" + totalR + "'" : "") + ' class="txtGrid txtT" placeholder="0" disabled style="width: 50px" /></td>' : '') + 
          (status === "Solicitud aprobada" ? '<td>' + (d.motivo != undefined ? d.motivo : "Pronosticado") + '</td>' : '') + 
        "</tr>");
      }

      $("input[type='number'].txtGrid").change(txtGridChangeEvent);
      
      $("#" + tabla + " thead .sum0").text(tL);
      $("#" + tabla + " thead .sum1").text(tM);
      $("#" + tabla + " thead .sum2").text(tX);
      $("#" + tabla + " thead .sum3").text(tJ);
      $("#" + tabla + " thead .sum4").text(tV);
      $("#" + tabla + " thead .sum5").text(tS);
      $("#" + tabla + " thead .sum6").text(tD);
      $(".sumT").text(tL + tM + tX + tJ + tV + tS + tD);

      $("#" + tabla + " thead .sumr0").text(tLr);
      $("#" + tabla + " thead .sumr1").text(tMr);
      $("#" + tabla + " thead .sumr2").text(tXr);
      $("#" + tabla + " thead .sumr3").text(tJr);
      $("#" + tabla + " thead .sumr4").text(tVr);
      $("#" + tabla + " thead .sumr5").text(tSr);
      $("#" + tabla + " thead .sumr6").text(tDr);
      $(".sumRT").text(tLr + tMr + tXr + tJr + tVr + tSr + tDr);
      
      obtenerComentarios();
    }, onRequestFail);
  }, onRequestFail);
});

function obtenerComentarios() {
  var oListCom = context.get_web().get_lists().getByTitle("Comentarios");
    var camlQueryCom = new SP.CamlQuery();

    camlQueryCom.set_viewXml("<View><Query><Where>" +
      "<Eq><FieldRef Name='Request' LookupId='TRUE'/><Value Type='Integer'>" + idcode + "</Value></Eq>" + 
    "</Where><OrderBy><FieldRef Name='Created' Ascending='FALSE'/></OrderBy></Query></View>");
    var collListItemCom = oListCom.getItems(camlQueryCom);
    
    context.load(collListItemCom);
    context.executeQueryAsync(function(sender, args){
      var listItemEnumerator = collListItemCom.getEnumerator();

      while (listItemEnumerator.moveNext()) {
        var oListItem = listItemEnumerator.get_current();

        $("#dvComentarios").append("<div style='border-bottom: dashed 1px #ddd;margin-bottom: 10px;'><p style='margin-bottom: 0px;'>" + oListItem.get_item("Comentario") + "</p><strong>" + oListItem.get_item("Author").get_lookupValue() + "</strong></div>");
      }

    asignarEventos();
  }, onRequestFail);
}

function asignarEventos() {
  $("#btnAgregarRecurso").click(function(){
    $("input[type='number'].txtGrid").unbind("change");
    $(".lkRemoverRow").unbind("click");

    if($(".txtNumCol").length > 0)
      $(".txtNumCol").autocomplete( "destroy" );

    $("#tbReal tbody").append('<tr class="trApoyo">' +
      '<td>' +
        '<a href="#" class="lkRemoverRow">' +
          '<img src="../SiteAssets/img/remove.png" style="width: 32px;" title="Remover" />' +
        '</a>' +
      '</td>' +
      '<td>' +
        '<input type="text" class="txtGrid txtNumCol" placeholder="0000" />' +
      '</td>' +
      '<td>' +
        '<span class="lblNombre"></span>' +
      '</td>' +
      '<td class="lblH">0</td>' + 
      '<td class="lblrH">' +
        '<input type="number" class="txtGrid txt0" min="0" max="16" step="1" placeholder="0" />' +
      '</td>' +
      '<td class="lblH">0</td>' + 
      '<td class="lblrH">' +
        '<input type="number" class="txtGrid txt1" min="0" max="16" step="1" placeholder="0" />' +
      '</td>' +
      '<td class="lblH">0</td>' + 
      '<td class="lblrH">' +
        '<input type="number" class="txtGrid txt2" min="0" max="16" step="1" placeholder="0" />' +
      '</td>' +
      '<td class="lblH">0</td>' + 
      '<td class="lblrH">' +
        '<input type="number" class="txtGrid txt3" min="0" max="16" step="1" placeholder="0" />' +
      '</td>' +
      '<td class="lblH">0</td>' + 
      '<td class="lblrH">' +
        '<input type="number" class="txtGrid txt4" min="0" max="16" step="1" placeholder="0" />' +
      '</td>' +
      '<td class="lblH">0</td>' + 
      '<td class="lblrH">' +
        '<input type="number" class="txtGrid txt5" min="0" max="16" step="1" placeholder="0" />' +
      '</td>' +
      '<td class="lblH">0</td>' + 
      '<td class="lblrH">' +
        '<input type="number" class="txtGrid txt6" min="0" max="16" step="1" placeholder="0" />' +
      '</td>' +
      '<td class="lblH">0</td>' + 
      '<td class="lblrH">' +
        '<input type="number" class="txtGrid txtT" min="0" max="16" step="1" placeholder="0" disabled />' +
      '</td>' +
      '<td>' +
        '<select class="cmbMotivo" >' + 
          '<option value="Enfermedad general y/o Incapacidad">Enfermedad general y/o Incapacidad</option>' + 
          '<option value="Baja de personal">Baja de personal</option>' + 
          '<option value="Vacaciones">Vacaciones</option>' + 
          '<option value="Por reemplazo">Por reemplazo</option>' + 
          '<option value="Solicitud de Dirección">Solicitud de Dirección</option>' + 
        '</select>' +
      '</td>' +
    '</tr>');

    $(".txtNumCol").autocomplete(autocompleteConfig);
    $("input[type='number'].txtGrid").change(txtGridChangeEvent);
    $(".lkRemoverRow").click(function(e){
      e.preventDefault();

      $(this).parent().parent().remove();

      recalcularHoras();
    });
  });
  $("#btnCerrar").click(function(){
    window.history.back();
  });
  $("#btnImprimir").click(function(){
    console.log("Hola");
    imprimirReporte();
  });
  $("#btnCancelar").click(function(){
    if(confirm("Está por cancelar esta solicitud. ¿Desea continuar?")) {
      var oList = context.get_web().get_lists().getByTitle("Request");
      var oListItem = oList.getItemById(idcode);

      context.load(oListItem);
      context.executeQueryAsync(function(sender, args) {
        oListItem.set_item("Status", "Solicitud cancelada/rechazada");
        oListItem.set_item("StatusReal", "Solicitud cancelada/rechazada");

        oListItem.update();
        context.executeQueryAsync(function(sender, args) {
          if($("#txtComentarios").val() != "") {
            var oListC = context.get_web().get_lists().getByTitle("Comentarios");
            var itemCreateInfoC = new SP.ListItemCreationInformation();
            var oListItemC = oListC.addItem(itemCreateInfoC);

            oListItemC.set_item("Title", "Solicitud cancelada/rechazada");
            oListItemC.set_item("Comentario", "Solicitud cancelada por el solicitante");
            oListItemC.set_item("Request", idcode);

            oListItemC.update();
            context.executeQueryAsync(function(sender, args) {
              alertify.alert("Extra Time", "Respuesta registrada exitosamente.", function(){
                window.location.href = "Solicitudes en proceso.aspx";
              });
            }, onRequestFail);
          } else {
            alertify.alert("Extra Time", "Respuesta registrada exitosamente.", function(){
              window.location.href = "Solicitudes en proceso.aspx";
            });
          }
        }, onRequestFail);
      }, onRequestFail);
    }
  });

  $("#fileDocto").on("change", function() {
    if($(this).val() != "") {
      $("#ulDoc").html("");
      $("#ulDoc").append("<li class='list-group-item'>" + $("#fileDocto")[0].files[0].name + "</li>");
    }
  });

  $("#btnEnviar").click(function(){
    $(this).prop("disabled", true);

    var oList = context.get_web().get_lists().getByTitle("Request");
    var oListItem = oList.getItemById(idcode);

    context.load(oListItem);
    context.executeQueryAsync(function(sender, args) {
      var detalle = JSON.parse(oListItem.get_item("Detalle"));
      var sinNombre = false;

      $("#tbReal tbody tr.trPlaneados").each(function(index, item){
        detalle[index].lr = $($(item).find(".txt0")[0]).val();
        detalle[index].mr = $($(item).find(".txt1")[0]).val();
        detalle[index].xr = $($(item).find(".txt2")[0]).val();
        detalle[index].jr = $($(item).find(".txt3")[0]).val();
        detalle[index].vr = $($(item).find(".txt4")[0]).val();
        detalle[index].sr = $($(item).find(".txt5")[0]).val();
        detalle[index].dr = $($(item).find(".txt6")[0]).val();
      });

      $("#tbReal tbody tr.trApoyo").each(function(index, item){
        if($($(item).find(".txtNumCol")[0]).val() != "" && $($(item).find(".lblNombre")[0]).text() === "") {
          sinNombre = true;
        }

        detalle.push({
          numcol: $($(item).find(".txtNumCol")[0]).val(),
          nombre: $($(item).find(".lblNombre")[0]).text(),
          l: 0,
          fl: $("#tbReal thead tr th#th0").text(),
          lr: $($(item).find(".txt0")[0]).val(),
          m: 0,
          fm: $("#tbReal thead tr th#th1").text(),
          mr: $($(item).find(".txt1")[0]).val(),
          x: 0,
          fx: $("#tbReal thead tr th#th2").text(),
          xr: $($(item).find(".txt2")[0]).val(),
          j: 0,
          fj: $("#tbReal thead tr th#th3").text(),
          jr: $($(item).find(".txt3")[0]).val(),
          v: 0,
          fv: $("#tbReal thead tr th#th4").text(),
          vr: $($(item).find(".txt4")[0]).val(),
          s: 0,
          fs: $("#tbReal thead tr th#th5").text(),
          sr: $($(item).find(".txt5")[0]).val(),
          d: 0,
          fd: $("#tbReal thead tr th#th6").text(),
          dr: $($(item).find(".txt6")[0]).val(),
          motivo: $($(item).find(".cmbMotivo")[0]).val(),
        });
      });

      if(sinNombre) {
        alertify.alert("Extra Time", "Información incompleta. Por favor, cuando ingrese el número de colaborador en la tabla de horas reales, seleccione el nombre y puesto del listado que le aparece.");
        $("#btnEnviar").prop("disabled", false);
        return;
      }

      oListItem.set_item("StatusReal", "En Área Solicitante");
      oListItem.set_item("Detalle", JSON.stringify(detalle));

      oListItem.update();
      context.executeQueryAsync(function(sender, args) {
        alertify.alert("Extra Time", "Registro de horas reales enviado correctamente.", function(){
          window.location.href = "Solicitudes en proceso.aspx";
        });
      }, onRequestFail);
    }, onRequestFail);
  });
}

function recalcularHoras() {
  var hL = 0, hM = 0, hX = 0, hJ = 0, hV = 0, hS = 0, hD = 0;

  $(".txt0").each(function(index){
    var v = parseFloat($(this).val());

    hL += isNaN(v) ? 0 : v;
  });
  $(".txt1").each(function(index){
    var v = parseFloat($(this).val());

    hM += isNaN(v) ? 0 : v;
  });
  $(".txt2").each(function(index){
    var v = parseFloat($(this).val());

    hX += isNaN(v) ? 0 : v;
  });
  $(".txt3").each(function(index){
    var v = parseFloat($(this).val());

    hJ += isNaN(v) ? 0 : v;
  });
  $(".txt4").each(function(index){
    var v = parseFloat($(this).val());

    hV += isNaN(v) ? 0 : v;
  });
  $(".txt5").each(function(index){
    var v = parseFloat($(this).val());

    hS += isNaN(v) ? 0 : v;
  });
  $(".txt6").each(function(index){
    var v = parseFloat($(this).val());

    hD += isNaN(v) ? 0 : v;
  });

  var sumRow = 0;

  $("#tbReal tbody tr input[type='number']").each(function(index){
    if((index + 1) % 8 === 0) {
      $(this).val(sumRow);
      sumRow = 0;
    } else {
      var v = parseFloat($(this).val());

      sumRow += isNaN(v) ? 0 : v;
    }
  });

  $(".sumr0").text(hL);
  $(".sumr1").text(hM);
  $(".sumr2").text(hX);
  $(".sumr3").text(hJ);
  $(".sumr4").text(hV);
  $(".sumr5").text(hS);
  $(".sumr6").text(hD);
  $(".sumRT").text(hL + hM + hX + hJ + hV + hS + hD);
}

function txtGridChangeEvent() {
  try{
    if(parseFloat($(this).val()) > 0) {
      $(this).css("background-color", "#8cdba9");
      recalcularHoras();
    } else {
      $(this).css("background-color", "white");
      $(this).val("0");
      recalcularHoras();
    }

  } catch(error) {
    $(this).val("0");
    console.log(error);
  }
}

function imprimirReporte() {
  window.print();
}

function subirArchivos(archivos, indice, spFiles, url, serverRelativeUrl, idSolicitud, callback){
  if(archivos.length > indice) {
    var getFile = getFileBuffer(archivos[indice]);

    getFile.done(function (arrayBuffer) {
      // Add the file to the SharePoint folder.
      var fileName = archivos[indice].name;
      var fileCollectionEndpoint = serverRelativeUrl + "/_api/web/lists/GetByTitle('Request')/items(" + idSolicitud + ")/AttachmentFiles/add(FileName='" + fileName + "')";

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
