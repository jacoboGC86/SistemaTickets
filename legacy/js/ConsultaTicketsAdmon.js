var context, cuser, web;

$(document).ready(function () {
    var iebrowser = /*@cc_on!@*/false || !!document.documentMode;

    if (iebrowser) {
        alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
        window.location.href = "https://promagroupmex.sharepoint.com/sites/paperless/ST/";
    }

    $("#DeltaPlaceHolderPageTitleInTitleArea a").text("Administración de tickets");
    $("#DeltaPlaceHolderPageTitleInTitleArea").show();

    context = new SP.ClientContext.get_current();
    web = context.get_web();
    cuser = web.get_currentUser();

    context.load(cuser);
    context.executeQueryAsync(function () {
        var cuserId = cuser.get_id();
        var ticketsList = web.get_lists().getByTitle("Tickets");
        var aprobacionesList = web.get_lists().getByTitle("Aprobaciones");

        var ticketsQuery = new SP.CamlQuery();
        ticketsQuery.set_viewXml(`
            <View>
                <Query>
                  
                    <OrderBy>
                        <FieldRef Name='ID' Ascending='FALSE'/>
                    </OrderBy>
                </Query>
                <RowLimit>500</RowLimit>
            </View>
        `);

        var ticketsItems = ticketsList.getItems(ticketsQuery);
        context.load(ticketsItems);

        context.executeQueryAsync(function () {
            var tickets = [];
            var ticketEnumerator = ticketsItems.getEnumerator();
            var rows = "";

            while (ticketEnumerator.moveNext()) {
                var ticketItem = ticketEnumerator.get_current();
                var ticketStatus = ticketItem.get_item("Status");
                var fechaAtencion = ticketItem.get_item("FechaAtencion");
                var tiempoTranscurridoTexto = "-";
                var tiempoTranscurridoHoras = 0;

                if (ticketStatus === "Cerrado" && fechaAtencion) {
                    tiempoTranscurridoHoras = getBusinessHoursDifference(fechaAtencion, ticketItem.get_item("Modified"));
                    tiempoTranscurridoTexto = getBusinessTimeDifference(fechaAtencion, ticketItem.get_item("Modified"));
                } else if (ticketStatus === "Assigned" && !fechaAtencion) {
                    tiempoTranscurridoHoras = getBusinessHoursDifference(ticketItem.get_item("Modified"));
                    tiempoTranscurridoTexto = getBusinessTimeDifference(ticketItem.get_item("Modified"));
                }
                var fecha = ticketItem.get_item("Modified"); 
                 // Array con los nombres de los meses
                const meses = [
                    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
                ];
                var mesIndex = new Date(fecha).getUTCMonth();
                var nombreMes = meses[mesIndex]; 
                let estilo = "";

                if (ticketItem.get_item("Prioridad") == "Baja" && parseInt(ticketItem.get_item("Categoria_x003a_SLALow").get_lookupValue()) - tiempoTranscurridoHoras > 1 ||
                    ticketItem.get_item("Prioridad") == "Media" && parseInt(ticketItem.get_item("Categoria_x003a_SLAMedium").get_lookupValue()) - tiempoTranscurridoHoras > 1 ||
                    ticketItem.get_item("Prioridad") == "Alta" && parseInt(ticketItem.get_item("Categoria_x003a_SLAHigh").get_lookupValue()) - tiempoTranscurridoHoras > 1) {
                        estilo = "style='background-color: #68ddbd; border: 1px solid #005227; color: #005227; border-radius: 5px; padding: 5px; display: inline-block'";
                } else if (ticketItem.get_item("Prioridad") == "Baja" && parseInt(ticketItem.get_item("Categoria_x003a_SLALow").get_lookupValue()) - tiempoTranscurridoHoras == 1 ||
                    ticketItem.get_item("Prioridad") == "Media" && parseInt(ticketItem.get_item("Categoria_x003a_SLAMedium").get_lookupValue()) - tiempoTranscurridoHoras == 1 ||
                    ticketItem.get_item("Prioridad") == "Alta" && parseInt(ticketItem.get_item("Categoria_x003a_SLAHigh").get_lookupValue()) - tiempoTranscurridoHoras == 1) {
                        estilo = "style='background-color: #ffff56; border: 1px solid #ff9800; color: #ff9800; border-radius: 5px; padding: 5px; display: inline-block'";
                } else {
                    estilo = "style='background-color: #e66d52; border: 1px solid #9d1d01; color: #9d1d01; border-radius: 5px; padding: 5px; display: inline-block'";
                }

             if(ticketItem.get_item("Status") == "Cerrado"){
                estilo = "style='background-color: #aed6f1; border: 1px solid #2874a6; color: #2874a6; border-radius: 5px; padding: 5px; display: inline-block'";
                }

                rows += `<tr id='tr_${ ticketItem.get_id() }'>
                        <td>${ ticketItem.get_id() }</td>
                        <td>${ ticketItem.get_item("Title") }</td>
                        <td>${ ticketItem.get_item("ProcessManager").get_lookupValue() }</td>
                        <td>${ ticketItem.get_item("TipoTicket") }</td>
                        <td>${ ticketItem.get_item("Categoria").get_lookupValue() }</td>
                        <td><a ${estilo} href='Detalle de ticket en Aprobacion.aspx?idSol=${ ticketItem.get_id() }'>${ ticketItem.get_item("Status") }</a></td>
                        <td>${ ticketItem.get_item("Prioridad") }</td>
                        <td>${ tiempoTranscurridoTexto }</td>
                        <td>${new Date(ticketItem.get_item("Modified")).getFullYear()}</td>
                        <td>${nombreMes}</td>
                        <td>${ ticketItem.get_item("Author").get_lookupValue() }</td>

                    </tr>`;
            }

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
});

function onRequestFail(sender, args) {
    console.error("Error en la consulta: " + args.get_message());
}

function getTimeDifference(dateStr) {
    const now = new Date();
    const created = new Date(dateStr);
    const diffMs = now - created;

    const minutes = Math.floor(diffMs / 60000) % 60;
    const hours = Math.floor(diffMs / (1000 * 60 * 60)) % 24;
    const days = Math.floor(diffMs / (1000 * 60 * 60 * 24));

    let result = [];
    if (days > 0) result.push(`${days} día(s)`);
    if (hours > 0) result.push(`${hours} hora(s)`);
    if (minutes > 0) result.push(`${minutes} minuto(s)`);

    return result.length > 0 ? result.join(", ") : "Menos de una hora";
}

function getHoursDifference(dateStr) {
    const now = new Date();
    const created = new Date(dateStr);
    const diffMs = now - created;
    return Math.floor(diffMs / (1000 * 60 * 60)); // horas enteras
}

// Calcula las horas hábiles transcurridas (Lun-Vie, 8am-6pm) entre startStr y endStr
// Si endStr se omite, usa la fecha/hora actual
function getBusinessHoursDifference(startStr, endStr) {
    const WORK_START = 8;  // 8am
    const WORK_END   = 18; // 6pm

    let current = new Date(startStr);
    const end   = endStr ? new Date(endStr) : new Date();

    if (current >= end) return 0;

    let totalMs = 0;

    while (current < end) {
        var day = current.getDay(); // 0=Dom, 1=Lun … 5=Vie, 6=Sáb

        if (day === 0 || day === 6) {
            // Fin de semana: avanzar al lunes siguiente a las 8am
            var daysToMonday = day === 0 ? 1 : 2;
            current.setDate(current.getDate() + daysToMonday);
            current.setHours(WORK_START, 0, 0, 0);
            continue;
        }

        var hour = current.getHours();

        if (hour < WORK_START) {
            // Antes del horario laboral: saltar a las 8am
            current.setHours(WORK_START, 0, 0, 0);
            continue;
        }

        if (hour >= WORK_END) {
            // Después del horario laboral: avanzar al día siguiente a las 8am
            current.setDate(current.getDate() + 1);
            current.setHours(WORK_START, 0, 0, 0);
            continue;
        }

        // Dentro del horario laboral: calcular hasta el fin del bloque
        var endOfWorkDay = new Date(current);
        endOfWorkDay.setHours(WORK_END, 0, 0, 0);

        var blockEnd = end < endOfWorkDay ? end : endOfWorkDay;
        totalMs += blockEnd - current;
        current = new Date(blockEnd);
    }

    return Math.floor(totalMs / (1000 * 60 * 60));
}

// Versión de texto para horas hábiles
// Si endStr se omite, usa la fecha/hora actual
function getBusinessTimeDifference(startStr, endStr) {
    var hours = getBusinessHoursDifference(startStr, endStr);
    var workHoursPerDay = 10; // 8am-6pm = 10 horas
    var days = Math.floor(hours / workHoursPerDay);
    var remainingHours = hours % workHoursPerDay;

    var result = [];
    if (days > 0) result.push(days + " día(s) hábil(es)");
    if (remainingHours > 0) result.push(remainingHours + " hora(s)");

    return result.length > 0 ? result.join(", ") : "Menos de una hora";
}
