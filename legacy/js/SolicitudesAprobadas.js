var context, cuser, web;

$(document).ready(function(){
    var iebrowser = /*@cc_on!@*/false || !!document.documentMode;
    
    if(iebrowser) {
        alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
        window.location.href = "https://chimalliapps.sharepoint.com/sites/Contacto/LL";
    }
    
    $("#DeltaPlaceHolderPageTitleInTitleArea a").text("Solicitudes Aprobadas");
    $("#DeltaPlaceHolderPageTitleInTitleArea").show();
    context = new SP.ClientContext.get_current();

    var oList = context.get_web().get_lists().getByTitle("Request");
    var camlQuery = new SP.CamlQuery();

    camlQuery.set_viewXml(`
                    <View>
                        <Query>
                            <Where>
                                <Eq>
                                    <FieldRef Name='Status' />
                                    <Value Type='Text'>Solicitud aprobada</Value>
                                </Eq>
                            </Where>
                            <OrderBy>
                                <FieldRef Name='ID' Ascending='FALSE'/>
                            </OrderBy>
                        </Query>
                    </View>
                `);
   
    var collListItem = oList.getItems(camlQuery);
    
    context.load(collListItem);
    context.executeQueryAsync(function(sender, args) {
        var listItemEnumerator = collListItem.getEnumerator();
        var rows = "", indices = [];
        console.log("Inicio de procesamiento de elementos");
    
        while (listItemEnumerator.moveNext()) {
            var oListItem = listItemEnumerator.get_current();
            var procesoAplicable = oListItem.get_item("ProcesoAplicable");
            var txtProcesoAplicable = procesoAplicable ? procesoAplicable.get_lookupValue() : "";
            var index = txtProcesoAplicable.indexOf("/");
            var txtProcesoAplicableCorto = index !== -1 ? txtProcesoAplicable.substring(index + 1) : txtProcesoAplicable.substring(0, 50);


            rows += "<tr id='tr_" + oListItem.get_id() + "'>" + 
                   "<td>" + oListItem.get_id() + "</td>" +
                   "<td><a href='Detalle de solicitud aprobada.aspx?idSol=" + oListItem.get_id() + "'>" + oListItem.get_item("Title") + "</a></td>"+
                   "<td>" + oListItem.get_item("Departamento") + "</td>" +
                   "<td>" + oListItem.get_item("Planta") + "</td>" +
                  "<td>" + txtProcesoAplicableCorto + "</td>" +
                "<td>" + obtenerFecha(oListItem.get_item("FFEngeeniering"), "dd/MMM/yyyy") + "</td>" +
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