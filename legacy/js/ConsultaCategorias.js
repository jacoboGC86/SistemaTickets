var context, cuser, web;

$(document).ready(function(){
    var iebrowser = /*@cc_on!@*/false || !!document.documentMode;
    
    if(iebrowser) {
        alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
        window.location.href = "https://chimalliapps.sharepoint.com/ST/";
    }
    
    $("#DeltaPlaceHolderPageTitleInTitleArea a").text("Categorías");
    $("#DeltaPlaceHolderPageTitleInTitleArea").show();
    context = new SP.ClientContext.get_current();

    var oList = context.get_web().get_lists().getByTitle("Categorias");
    var camlQuery = new SP.CamlQuery();

    camlQuery.set_viewXml("<View><Query><OrderBy><FieldRef Name='ID' Ascending='FALSE'/></OrderBy></Query><RowLimit>500</RowLimit></View>");
    var collListItem = oList.getItems(camlQuery);
    
    context.load(collListItem);
    context.executeQueryAsync(function(sender, args) {
        var listItemEnumerator = collListItem.getEnumerator();
        var rows = "", indices = [];
        console.log("Inicio de procesamiento de elementos");
        
        while (listItemEnumerator.moveNext()) {
            var oListItem = listItemEnumerator.get_current();
            var lkTemplate = oListItem.get_item("TemplateAtencion") ? ("<a href='Detalle de templates.aspx?idSol=" + oListItem.get_item("TemplateAtencion").get_lookupId() + "'>" + oListItem.get_item("TemplateAtencion").get_lookupValue() + "</a>") : "Indefinido";
            var pmPuebla = oListItem.get_item("TemplateAtencion_x003a_PMPuebla") ? oListItem.get_item("TemplateAtencion_x003a_PMPuebla").get_lookupValue() : "null";
            var pmAguascalientes = oListItem.get_item("TemplateAtencion_x003a_PMAguasca") ? oListItem.get_item("TemplateAtencion_x003a_PMAguasca").get_lookupValue() : "null";
         
            rows += "<tr id='tr_" + oListItem.get_id() + "'>" + 
                   "<td>" + oListItem.get_id() + "</td>" +
                   "<td>" + (oListItem.get_item("CategoriaPadre") ? oListItem.get_item("CategoriaPadre").get_lookupValue() : "Indefinida") + "</td>"+
                   "<td><a href='Detalle de categoria.aspx?idSol=" + oListItem.get_id() + "'>" + oListItem.get_item("Title") + "</a></td>" +
                   "<td>" + oListItem.get_item("Prioridad") + "</td>" +
                   "<td>" + lkTemplate + "</td>" +
                   "<td>" + pmPuebla + "</td>" +
                   "<td>" + pmAguascalientes + "</td>" +
                "</tr>";
        }
    
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