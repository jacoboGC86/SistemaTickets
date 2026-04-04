var context, cuser, web;

$(document).ready(function () {
    var iebrowser = /*@cc_on!@*/false || !!document.documentMode;

    if (iebrowser) {
        alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
        window.location.href = "https://chimalliapps.sharepoint.com/ST/";
    }
    

    $("#DeltaPlaceHolderPageTitleInTitleArea a").text("Tickets Cerrados");
    $("#DeltaPlaceHolderPageTitleInTitleArea").show();

    context = new SP.ClientContext.get_current();
    var ticketsList = context.get_web().get_lists().getByTitle("Tickets");
    var categoriasList = context.get_web().get_lists().getByTitle("Categorias");
     //1.Consulta las categorías
        // Carga las categorías relacionadas 
        var categoriasQuery = new SP.CamlQuery();
        categoriasQuery.set_viewXml("<View><Query></Query></View>");

        var categoriasItems = categoriasList.getItems(categoriasQuery);
        context.load(categoriasItems);
        context.executeQueryAsync(function () {
            // crea un obj con el id de cada categoría 
            var categoriasMap = {};
            var categoriasEnumerator = categoriasItems.getEnumerator();
            //recorre las categorías
            while (categoriasEnumerator.moveNext()) {
                var categoriaItem = categoriasEnumerator.get_current();
                categoriasMap[categoriaItem.get_id()] = {
                    nombre: categoriaItem.get_item("Title"),
                    CategoriaPadre: categoriaItem.get_item("CategoriaPadre")
                ? categoriaItem.get_item("CategoriaPadre").get_lookupValue()
                : "Sin categoría padre"

            };
        }
       //2. Segunda consulta a Tickets para verificar la existencia de Categoria
    var ticketsQuery = new SP.CamlQuery();
    ticketsQuery.set_viewXml(`
                <View>
                    <Query>
                        <Where>
                            <Eq>
                                <FieldRef Name='Status' />
                                <Value Type='Text'>Cerrado</Value>
                            </Eq>
                        </Where>
                        <OrderBy>
                            <FieldRef Name='ID' Ascending='FALSE'/>
                        </OrderBy>
                    </Query>
                    <RowLimit>200</RowLimit>
                </View>

            `);
    var ticketsItems = ticketsList.getItems(ticketsQuery);

    context.load(ticketsItems);
    context.executeQueryAsync(function () {
        var tickets = [];
        var ticketEnumerator = ticketsItems.getEnumerator();
        while (ticketEnumerator.moveNext()) {
            var ticketItem = ticketEnumerator.get_current();
            var categoria = ticketItem.get_item("Categoria"); 
            var fecha = ticketItem.get_item("Modified"); 
        
            // Array con los nombres de los meses
            const meses = [
                "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
            ];
        
            var mesIndex = new Date(fecha).getUTCMonth();
            var nombreMes = meses[mesIndex]; 
        
            tickets.push({
                id: ticketItem.get_id(),
                title: ticketItem.get_item("Title"),
                tipoTicket: ticketItem.get_item("TipoTicket"),
                status: ticketItem.get_item("Status"),
                prioridad: ticketItem.get_item("Prioridad"),
                fecha: fecha, 
                year: new Date(fecha).getFullYear(),
                month: nombreMes, 
                categoriaId: categoria.get_lookupId(),
                categoriaName: categoria.get_lookupValue(),
                Solicitante: ticketItem.get_item("Author").get_lookupValue(),
            });
        }
        
            var rows =  tickets.map(ticket => {
                var categoriaInfo = categoriasMap[ticket.categoriaId] ||{
                    nombre: "Sin categoría",
                    CategoriaPadre: "Sin categoría padre"
                };
                return `
                    <tr id='tr_${ticket.id}'>
                        <td>${ticket.id}</td>
                        <td>${ticket.title}</td>
                          <td><a href='Detalle de ticket en Aprobacion.aspx?idSol=${ticket.id}'>${ticket.status}</a></td>
                        <td>${ticket.tipoTicket}</td>
                        <td>${categoriaInfo.nombre} - ${categoriaInfo.CategoriaPadre}</td>
                        <td>${ticket.month}</td>
                        <td>${ticket.year}</td>
                        <td>${ticket.Solicitante}</td>
                    
                    </tr>
                `;
            
        }).join("");
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
    }, onRequestFail);
});

function onRequestFail(sender, args) {
    console.error("Error: " + args.get_message());
}
