var context, cuser, web;

$(document).ready(function(){
    var iebrowser = /*@cc_on!@*/false || !!document.documentMode;
    
    if(iebrowser) {
        alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
        window.location.href = "https://chimalliapps.sharepoint.com/ST/";
    }
    
    $("#DeltaPlaceHolderPageTitleInTitleArea a").text("Plantillas");
    $("#DeltaPlaceHolderPageTitleInTitleArea").show();
    context = new SP.ClientContext.get_current();

    var oList = context.get_web().get_lists().getByTitle("Templates");
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
            Descripcion = oListItem.get_item("Descripcion");
            if(Descripcion.length > 50){
                Descripcion = Descripcion.substring(0,50);
            }
         
            rows += "<tr id='tr_" + oListItem.get_id() + "'>" + 
                   "<td>" + oListItem.get_id() + "</td>" +
                   "<td><a href='Detalle de templates.aspx?idSol=" + oListItem.get_id() + "'>" + oListItem.get_item("Title") + "</a></td>"+
                   "<td>" + Descripcion + "</td>" +
                   "<td>" + obtenerFecha(oListItem.get_item("Modified"), "dd/mm/yyyy") + "</td>" +
                   "<td>" + oListItem.get_item("PMPuebla") + "</td>" +
                   "<td>" + oListItem.get_item("PMAguascalientes") + "</td>" +
                "</tr>";
    
            indices.push(oListItem.get_id());
        }
    
        console.log("Elementos procesados: " + indices.length);
    
        $("#dvLoading").hide();
        $("#tbCodigos_WP").show();
        $("#tbCodigos_WP tbody").html(rows);
    
        
        $('#tbCodigos_WP').DataTable({
            "order": [[ 0, "desc" ]],
            "language": {
                "lengthMenu": "Mostrar _MENU_ registros por página",
                "zeroRecords": "No se encontraron resultados",
                "info": "Mostrando página _PAGE_ de _PAGES_",
                "infoEmpty": "No hay registros disponibles",
                "infoFiltered": "(filtrado de _MAX_ registros totales)",
                "search": "Buscar:",
                "paginate": {
                    "first": "Primero",
                    "last": "Último",
                    "next": "Siguiente",
                    "previous": "Anterior"
                }
            }
        });
        
    }, onRequestFail);
    
});