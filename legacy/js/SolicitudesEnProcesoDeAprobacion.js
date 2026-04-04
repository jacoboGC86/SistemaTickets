var context, cuser, web;

$(document).ready(function(){
    var iebrowser = /*@cc_on!@*/false || !!document.documentMode;
    
    if(iebrowser) {
        alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
        window.location.href = "https://chimalliapps.sharepoint.com/sites/Contacto/LL";
    }
    
    $("#DeltaPlaceHolderPageTitleInTitleArea a").text("Solicitudes en proceso de aprobación");
    $("#DeltaPlaceHolderPageTitleInTitleArea").show();
    context = new SP.ClientContext.get_current();

    var oList = context.get_web().get_lists().getByTitle("Request");
    var camlQuery = new SP.CamlQuery();

    camlQuery.set_viewXml("<View><Query><OrderBy><FieldRef Name='ID' Ascending='FALSE'/></OrderBy></Query><RowLimit>200</RowLimit></View>");
    var collListItem = oList.getItems(camlQuery);
    
    context.load(collListItem);
    context.executeQueryAsync(function(sender, args) {
        var listItemEnumerator = collListItem.getEnumerator();
        var rows = "", indices = [];
        console.log("Inicio de procesamiento de elementos");
        
        while (listItemEnumerator.moveNext()) {
            var oListItem = listItemEnumerator.get_current();
            console.log("Procesando item ID: " + oListItem.get_id());
            var procesoAplicable = oListItem.get_item("ProcesoAplicable");
            var textoProcesoAplicable = procesoAplicable ? procesoAplicable.get_lookupValue() : "";

            rows += "<tr id='tr_" + oListItem.get_id() + "'>" + 
                   "<td>" + oListItem.get_id() + "</td>" +
                   "<td>" + oListItem.get_item("Title") + "</td>" +
                   "<td><a href='Detalle de solicitud en aprobacion.aspx?idSol=" + oListItem.get_id() + "'>" + oListItem.get_item("Status") + "</a></td>"+
                   "<td>" + oListItem.get_item("Departamento") + "</td>" +
                   "<td>" + oListItem.get_item("Planta") + "</td>" 
                //   "<td>" + textoProcesoAplicable + "</td>" +
                // "<td>" + obtenerFecha(oListItem.get_item("Modified"), "dd/MMM/yyyy") + "</td>" +
                "</tr>";
    
            indices.push(oListItem.get_id());
        }
    
        console.log("Elementos procesados: " + indices.length);
    
        $("#dvLoading").hide();
        $("#tbCodigos_WP").show();
        $("#tbCodigos_WP tbody").html(rows);
    
        $('#tbCodigos_WP').DataTable({
          "order": [[ 0, "desc" ]]
        });
    }, onRequestFail);
    
});