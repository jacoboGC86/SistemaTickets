var context, contextCol, cuser, web, iamRH = false;

$(document).ready(function(){
    var iebrowser = /*@cc_on!@*/false || !!document.documentMode;
    
    if(iebrowser) {
        alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
        window.location.href = "ttps://volaris1.sharepoint.com/sites/parpeless_dev/CCM/";
    }
    
    $("#DeltaPlaceHolderPageTitleInTitleArea a").text("Aprobaciones en Proceso");
    $("#DeltaPlaceHolderPageTitleInTitleArea").show();
    context = new SP.ClientContext.get_current();
    contextCol = new SP.ClientContext("https://chimalliapps.sharepoint.com/sites/Contacto/Vacations/");
    web = context.get_web();
    cuser = web.get_currentUser();

    context.load(cuser);
    context.load(web);
    context.executeQueryAsync(function(sender, args){
      var oListConf = contextCol.get_web().get_lists().getByTitle("Configuraciones");
      var camlQueryConf = new SP.CamlQuery();

      camlQueryConf.set_viewXml("<View><Query><Where></Where></Query></View>");
      var collListItemConf = oListConf.getItems(camlQueryConf);

      contextCol.load(collListItemConf);
      contextCol.executeQueryAsync(function(sender, args){
        var listItemEnumerator = collListItemConf.getEnumerator();

        while (listItemEnumerator.moveNext()) {
          var oListItem = listItemEnumerator.get_current();

          switch(oListItem.get_item("Title")) {
            case "RH": 
            for(var i = 0; i < oListItem.get_item("Personas").length; i++) {
              if(cuser.get_id() === oListItem.get_item("Personas")[i].get_lookupId()) {
                iamRH = true;
                break;
              }
            }
            break;
          }
        }

        if(!iamRH) {
          var oList = context.get_web().get_lists().getByTitle("Request");
          var camlQuery = new SP.CamlQuery();

          camlQuery.set_viewXml("<View><Query><Where>" + 
          "<And>" +
          "<Eq><FieldRef Name='RespAreaGeneradora' LookupId='TRUE'/><Value Type='Integer'><UserID /></Value></Eq>" + 
          "<Or>" + 
            "<Eq><FieldRef Name='Status' /><Value Type='Text'>En Firma de Área Generadora</Value></Eq>" + 
            "<Eq><FieldRef Name='StatusReal' /><Value Type='Text'>En Área Generadora</Value></Eq>" + 
          "</Or>" + 
          "</And>" + 
          "</Where><OrderBy><FieldRef Name='ID' Ascending='FALSE'/></OrderBy></Query></View>");

          var collListItem = oList.getItems(camlQuery);
            
          context.load(collListItem);
          context.executeQueryAsync(function(sender, args){
            var listItemEnumerator = collListItem.getEnumerator();
            var rows = "", indices = [];

            while (listItemEnumerator.moveNext()) {
                var oListItem = listItemEnumerator.get_current();

                rows += "<tr id='tr_" + oListItem.get_id() + "'>" + 
                  "<td>" + oListItem.get_id() + "</td>" +
                  "<td><span style='display:none'>" + obtenerFecha(oListItem.get_item("Created"), "yyyymmdd") + "</span>" + obtenerFecha(oListItem.get_item("Created"), "dd/mm/yyyy") + "</td>" +
                  "<td><a href='Detalle de solicitud Aprobacion.aspx?idsol=" + oListItem.get_id() + "' tarjet='_blank'>" + (oListItem.get_item("StatusReal") != "Pendiente" ? oListItem.get_item("StatusReal") : oListItem.get_item("Status")) + "</a></td>" +
                  "<td>" + oListItem.get_item("Semana") + "</td>" +
                  "<td>" + oListItem.get_item("AreaSolicitante") + "</td>" +
                  "<td>" + oListItem.get_item("AreaGeneradora") + "</td>" +
                  "<td>" + obtenerFecha(oListItem.get_item("FechaMaximaAprobacion"), "dd/MMM/yyyy") + "</td>" +
                "</tr>";
                
                indices.push(oListItem.get_id());
            }

            $("#dvLoading").hide();
            $("#tbCodigos_WP").show();
            $("#tbCodigos_WP tbody").html(rows);

            obtenerSolicitudesSolicitante()
          }, onRequestFail);
        } else {
          $("#dvFilters").show();

          for(var i = (new Date()).getFullYear(); i >= 2021; i--) {
            $("#cmbAgno").append("<option value='" + i + "'>" + i + "</option>");
          }

          $("#cmbAgno").on("change", function(){
            $("#dvLoading").show();
            $('#tbCodigos_WP').DataTable().destroy();

            $("#tbCodigos_WP tbody").html("");
            $("#cmbAgno").prop("disabled", true);
            $("#cmbMes").prop("disabled", true);

            obtenerTodasSolicitudes(null);
          });
          $("#cmbMes").val(moment().format("MM"));

          $("#cmbMes").on("change", function(){
            $("#dvLoading").show();
            $('#tbCodigos_WP').DataTable().destroy();

            $("#tbCodigos_WP tbody").html("");
            $("#cmbAgno").prop("disabled", true);
            $("#cmbMes").prop("disabled", true);

            obtenerTodasSolicitudes(null);
          });
          
          obtenerTodasSolicitudes(null);
        }
      }, onRequestFail);
    }, onRequestFail);
});

function obtenerSolicitudesSolicitante() {
  var oList = context.get_web().get_lists().getByTitle("Request");
  var camlQuery = new SP.CamlQuery();

  camlQuery.set_viewXml("<View><Query><Where>" + 
  "<And>" +
  "<Eq><FieldRef Name='RespAreaSolicitante' LookupId='TRUE'/><Value Type='Integer'><UserID /></Value></Eq>" + 
  "<Or>" + 
    "<Eq><FieldRef Name='Status' /><Value Type='Text'>En Firma de Área Solicitante</Value></Eq>" + 
    "<Eq><FieldRef Name='StatusReal' /><Value Type='Text'>En Área Solicitante</Value></Eq>" + 
  "</Or>" + 
  "</And>" + 
  "</Where><OrderBy><FieldRef Name='ID' Ascending='FALSE'/></OrderBy></Query></View>");

  var collListItem = oList.getItems(camlQuery);
    
  context.load(collListItem);
  context.executeQueryAsync(function(sender, args){
    var listItemEnumerator = collListItem.getEnumerator();
    var rows = "", indices = [];

    while (listItemEnumerator.moveNext()) {
        var oListItem = listItemEnumerator.get_current();

        rows += "<tr id='tr_" + oListItem.get_id() + "'>" + 
          "<td>" + oListItem.get_id() + "</td>" +
          "<td><span style='display:none'>" + obtenerFecha(oListItem.get_item("Created"), "yyyymmdd") + "</span>" + obtenerFecha(oListItem.get_item("Created"), "dd/mm/yyyy") + "</td>" +
          "<td><a href='Detalle de solicitud Aprobacion.aspx?idsol=" + oListItem.get_id() + "' tarjet='_blank'>" + (oListItem.get_item("StatusReal") != "Pendiente" ? oListItem.get_item("StatusReal") : oListItem.get_item("Status")) + "</a></td>" +
          "<td>" + oListItem.get_item("Semana") + "</td>" +
          "<td>" + oListItem.get_item("AreaSolicitante") + "</td>" +
          "<td>" + oListItem.get_item("AreaGeneradora") + "</td>" +
          "<td>" + obtenerFecha(oListItem.get_item("FechaMaximaAprobacion"), "dd/MMM/yyyy") + "</td>" +
        "</tr>";
        
        indices.push(oListItem.get_id());
    }

    $("#tbCodigos_WP tbody").append(rows);
    $('#tbCodigos_WP').DataTable({
      "order": [[ 0, "desc" ]]
    });
  }, onRequestFail);
}

function obtenerTodasSolicitudes(pos) {
  var oList = context.get_web().get_lists().getByTitle("Request");
  var camlQuery = new SP.CamlQuery();

  camlQuery.set_viewXml("<View><Query><Where>" + 
  "<And>" + 
    "<Geq><FieldRef Name='Created' IncludeTimeValue='FALSE' Type='DateTime' /><Value Type='DateTime'>" + $("#cmbAgno").val() + "-" + $("#cmbMes").val() + "-01T00:00:00.00Z</Value></Geq>" + 
    "<Leq><FieldRef Name='Created' IncludeTimeValue='FALSE' Type='DateTime' /><Value Type='DateTime'>" + $("#cmbAgno").val() + "-" + $("#cmbMes").val() + "-" + moment($("#cmbAgno").val() + "-" + $("#cmbMes").val(), "YYYY-MM").daysInMonth() + "T23:59:00.00Z</Value></Leq>" + 
  "</And>" +
  "</Where><OrderBy><FieldRef Name='ID' Ascending='FALSE'/></OrderBy></Query><RowLimit>200</RowLimit></View>");
  if(pos != null)
    camlQuery.set_listItemCollectionPosition(pos);

  var collListItem = oList.getItems(camlQuery);
    
  context.load(collListItem);
  context.executeQueryAsync(function(sender, args){
    var listItemEnumerator = collListItem.getEnumerator();
    var rows = "", indices = [];

    while (listItemEnumerator.moveNext()) {
        var oListItem = listItemEnumerator.get_current();

        rows += "<tr id='tr_" + oListItem.get_id() + "'>" + 
          "<td>" + oListItem.get_id() + "</td>" +
          "<td><span style='display:none'>" + obtenerFecha(oListItem.get_item("Created"), "yyyymmdd") + "</span>" + obtenerFecha(oListItem.get_item("Created"), "dd/mm/yyyy") + "</td>" +
          "<td><a href='Detalle de solicitud Aprobacion.aspx?idsol=" + oListItem.get_id() + "' tarjet='_blank'>" + (oListItem.get_item("StatusReal") != "Pendiente" ? oListItem.get_item("StatusReal") : oListItem.get_item("Status")) + "</a></td>" +
          "<td>" + oListItem.get_item("Semana") + "</td>" +
          "<td>" + oListItem.get_item("AreaSolicitante") + "</td>" +
          "<td>" + oListItem.get_item("AreaGeneradora") + "</td>" +
          "<td>" + obtenerFecha(oListItem.get_item("FechaMaximaAprobacion"), "dd/MMM/yyyy") + "</td>" +
        "</tr>";
        
        indices.push(oListItem.get_id());
    }

    $("#tbCodigos_WP tbody").append(rows);

    if(collListItem.get_listItemCollectionPosition() === null) {
      $("#cmbAgno").prop("disabled", false);
      $("#cmbMes").prop("disabled", false);
      $("#dvLoading").hide();
      $("#tbCodigos_WP").show();

      $('#tbCodigos_WP').DataTable({
        "order": [[ 0, "desc" ]]
      });
    } else {
      obtenerTodasSolicitudes(collListItem.get_listItemCollectionPosition());
    }
  }, onRequestFail);
}