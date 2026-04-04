var context, cuser, web;

$(document).ready(function(){
    var iebrowser = /*@cc_on!@*/false || !!document.documentMode;
    
    if(iebrowser) {
        alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
        window.location.href = "https://chimalliapps.sharepoint.com/sites/Contacto/Vacations/";
    }
    
    $("#DeltaPlaceHolderPageTitleInTitleArea a").text("Solicitudes en Proceso");
    $("#DeltaPlaceHolderPageTitleInTitleArea").show();
    context = new SP.ClientContext.get_current();

    var oList = context.get_web().get_lists().getByTitle("Request");
    var camlQuery = new SP.CamlQuery();

    camlQuery.set_viewXml("<View><Query><Where>" + 
    "<Eq><FieldRef Name='Author' LookupId='TRUE'/><Value Type='Integer'><UserID /></Value></Eq>" + 
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
                "<td><a href='Detalle de solicitud.aspx?idsol=" + oListItem.get_id() + "' tarjet='_blank'>" + (oListItem.get_item("StatusReal") != "Pendiente" ? oListItem.get_item("StatusReal") : (oListItem.get_item("Status") === "Solicitud aprobada" ? "Solicitud aprobada - En registro de horas reales" : oListItem.get_item("Status"))) + "</a></td>" +
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

        $('#tbCodigos_WP').DataTable({
          "order": [[ 0, "desc" ]]
        });
    }, onRequestFail);
});