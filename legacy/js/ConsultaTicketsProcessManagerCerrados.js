var context, cuser, web;

$(document).ready(function () {
    var iebrowser = /*@cc_on!@*/false || !!document.documentMode;

    if (iebrowser) {
        alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
        window.location.href = "https://chimalliapps.sharepoint.com/ST/";
    }

    $("#DeltaPlaceHolderPageTitleInTitleArea a").text("Tickets cerrados por Process manager");
    $("#DeltaPlaceHolderPageTitleInTitleArea").show();

    context = new SP.ClientContext.get_current();
    web = context.get_web();
    cuser = web.get_currentUser();

    context.load(cuser);
    context.executeQueryAsync(function () {
        var cuserId = cuser.get_id(); // Obtiene el ID del usuario en sesión
        console.log("Usuario en sesión ID:", cuserId);

        var ticketsList = web.get_lists().getByTitle("Tickets");
        var categoriasList = web.get_lists().getByTitle("Categorias");

        // 1. Consultar las categorías
        var categoriasQuery = new SP.CamlQuery();
        categoriasQuery.set_viewXml("<View><Query></Query></View>");
        var categoriasItems = categoriasList.getItems(categoriasQuery);

        context.load(categoriasItems);
        context.executeQueryAsync(function () {
            var categoriasMap = {};
            var categoriasEnumerator = categoriasItems.getEnumerator();

            // Mapea las categorías
            while (categoriasEnumerator.moveNext()) {
                var categoriaItem = categoriasEnumerator.get_current();
                var categoriaPadre = categoriaItem.get_item("CategoriaPadre");

                categoriasMap[categoriaItem.get_id()] = {
                    nombre: categoriaItem.get_item("Title"),
                    CategoriaPadre: categoriaPadre ? categoriaPadre.get_lookupValue() : "Sin categoría padre"
                };
            }

            // 2. Consulta los tickets asignados al process manager de esta sesión
            var ticketsQuery = new SP.CamlQuery();
            ticketsQuery.set_viewXml(`
                <View>
                    <Query>
                        <Where>
                            <And>
                                <Eq>
                                    <FieldRef Name='ProcessManager' LookupId='TRUE' />
                                    <Value Type='Integer'>${cuserId}</Value>
                                </Eq>
                                <Eq>
                                    <FieldRef Name='Status' />
                                    <Value Type='Text'>Cerrado</Value>
                                </Eq>
                            </And>
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

                    var processManager = ticketItem.get_item("ProcessManager");
                    var processManagerName = processManager ? processManager.get_lookupValue() : "Sin asignar";

                    tickets.push({
                        id: ticketItem.get_id(),
                        title: ticketItem.get_item("Title"),
                        processManager: processManagerName,
                        tipoTicket: ticketItem.get_item("TipoTicket"),
                        status: ticketItem.get_item("Status"),
                        prioridad: ticketItem.get_item("Prioridad"),
                        categoriaId: categoria ? categoria.get_lookupId() : 0, // Verifica si existe la categoría
                        categoriaName: categoria ? categoria.get_lookupValue() : "Sin categoría",
                        Solicitante: ticketItem.get_item("Author").get_lookupValue()
                    });
                }

                var rows = tickets.map(ticket => {
                    var categoriaInfo = categoriasMap[ticket.categoriaId] || {
                        nombre: "Sin categoría",
                        CategoriaPadre: "Sin categoría padre"
                    };
                    return `
                        <tr id='tr_${ticket.id}'>
                            <td>${ticket.id}</td>
                            <td>${ticket.title}</td>
                            <td>${ticket.processManager}</td>
                            <td>${ticket.tipoTicket}</td>
                            <td>${categoriaInfo.nombre} - ${categoriaInfo.CategoriaPadre}</td>
                            <td><a href='Detalle de ticket en Aprobacion.aspx?idSol=${ticket.id}'>${ticket.status}</a></td>
                            <td>${ticket.prioridad}</td>
                            <td>${ticket.Solicitante}</td>
                        </tr>
                    `;
                }).join("");

                $("#dvLoading").hide();
                $("#tbCodigos_WP").show();
                $("#tbCodigos_WP tbody").html(rows);

                $('#tbCodigos_WP').DataTable({
                    "order": [[0, "desc"]],
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
    }, onRequestFail);
});

function onRequestFail(sender, args) {
    console.error("Error en la consulta: " + args.get_message());
}
