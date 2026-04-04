var context, cuser, web;

$(document).ready(function () {

    var iebrowser = /*@cc_on!@*/false || !!document.documentMode;

    if (iebrowser) {
        alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
        window.location.href = "https://chimalliapps.sharepoint.com/ST/";
        return;
    }

    $("#DeltaPlaceHolderPageTitleInTitleArea a").text("Tickets asignados");
    $("#DeltaPlaceHolderPageTitleInTitleArea").show();

    context = SP.ClientContext.get_current();
    web = context.get_web();
    cuser = web.get_currentUser();

    context.load(cuser);

    context.executeQueryAsync(function () {

        var cuserId = cuser.get_id();

        loadTicketsAsignados(cuserId);
        loadTicketsPorGrupoOManger(cuserId);

    }, onRequestFail);

});



/* ======================================================
   PRIMERA TABLA
   Tickets asignados directamente al Process Manager
====================================================== */

// function loadTicketsAsignados(cuserId) {

//     var ticketsList = web.get_lists().getByTitle("Tickets");

//     var ticketsQuery = new SP.CamlQuery();

//     ticketsQuery.set_viewXml(`
//         <View>
//             <Query>
//                 <Where>
//                     <And>
//                         <Eq>
//                             <FieldRef Name='ProcessManager' LookupId='TRUE'/>
//                             <Value Type='Integer'>${cuserId}</Value>
//                         </Eq>
//                         <Eq>
//                             <FieldRef Name='Status'/>
//                             <Value Type='Text'>Assigned</Value>
//                         </Eq>
//                     </And>
//                 </Where>
//                 <OrderBy>
//                     <FieldRef Name='ID' Ascending='FALSE'/>
//                 </OrderBy>
//             </Query>
//             <RowLimit>500</RowLimit>
//         </View>
//     `);

//     var ticketsItems = ticketsList.getItems(ticketsQuery);

//     context.load(ticketsItems,
//         'Include(Id,Title,TipoTicket,Categoria,Prioridad,Status,Atencion,Modified,ProcessManager,Author,Categoria_x003a_SLALow,Categoria_x003a_SLAMedium,Categoria_x003a_SLAHigh)'
//     );

//     context.executeQueryAsync(function () {

//         var rows = "";
//         var enumerator = ticketsItems.getEnumerator();

//         while (enumerator.moveNext()) {

//             var ticket = enumerator.get_current();

//             var status = ticket.get_item("Status");
//             var atencion = ticket.get_item("Atencion");

//             var statusLabel = status;

//             if (atencion === true) {
//                 statusLabel = "En Atención";
//             }

//             const horas = getHoursDifference(ticket.get_item("Modified"));
//             const tiempoTexto = getTimeDifference(ticket.get_item("Modified"));
//             const prioridad = ticket.get_item("Prioridad");

//             var fecha = ticket.get_item("Modified");

//             const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
//             var mesIndex = new Date(fecha).getUTCMonth();
//             var nombreMes = meses[mesIndex];

//             const sla = {
//                 "Baja": ticket.get_item("Categoria_x003a_SLALow"),
//                 "Media": ticket.get_item("Categoria_x003a_SLAMedium"),
//                 "Alta": ticket.get_item("Categoria_x003a_SLAHigh")
//             }[prioridad].get_lookupValue();

//             const diferencia = parseInt(sla) - horas;

//             let estilo = "";

//             if (diferencia > 1)
//                 estilo = "style='background-color:#68ddbd;border:1px solid #005227;color:#005227;border-radius:5px;padding:5px;display:inline-block'";
//             else if (diferencia === 1)
//                 estilo = "style='background-color:#ffff56;border:1px solid #ff9800;color:#ff9800;border-radius:5px;padding:5px;display:inline-block'";
//             else
//                 estilo = "style='background-color:#e66d52;border:1px solid #9d1d01;color:#9d1d01;border-radius:5px;padding:5px;display:inline-block'";

//             rows += `
//             <tr id='tr_${ticket.get_id()}'>
//                 <td>${ticket.get_id()}</td>
//                 <td>${ticket.get_item("Title").slice(0,50)}</td>
//                 <td>${ticket.get_item("ProcessManager").get_lookupValue()}</td>
//                 <td>${ticket.get_item("TipoTicket")}</td>
//                 <td>${ticket.get_item("Categoria").get_lookupValue()}</td>
//                 <td><a ${estilo} href='Detalle de ticket en Aprobacion.aspx?idSol=${ticket.get_id()}'>${statusLabel}</a></td>
//                 <td>${prioridad}</td>
//                 <td>${tiempoTexto}</td>
//                 <td>${new Date(ticket.get_item("Modified")).getFullYear()}</td>
//                 <td>${nombreMes}</td>
//                 <td>${ticket.get_item("Author").get_lookupValue()}</td>
//             </tr>`;
//         }

//         $("#dvLoading").hide();
//         $("#tbCodigos_WP").show().find("tbody").html(rows);

//         $('#tbCodigos_WP').DataTable({
//             order: [[0,"desc"]],
//             language: {
//                 lengthMenu:"Mostrar _MENU_ registros por página",
//                 zeroRecords:"No se encontraron resultados",
//                 info:"Mostrando página _PAGE_ de _PAGES_",
//                 infoEmpty:"No hay registros disponibles",
//                 infoFiltered:"(filtrado de _MAX_ registros totales)",
//                 search:"Buscar:",
//                 paginate:{
//                     first:"Primero",
//                     last:"Último",
//                     next:"Siguiente",
//                     previous:"Anterior"
//                 }
//             }
//         });

//     }, onRequestFail);
// }
function loadTicketsAsignados(cuserId) {

    var ticketsList = web.get_lists().getByTitle("Tickets");
    var ticketsQuery = new SP.CamlQuery();

    ticketsQuery.set_viewXml(`
        <View>
            <Query>
                <Where>
                    <And>
                        <Eq>
                            <FieldRef Name='ProcessManager' LookupId='TRUE'/>
                            <Value Type='Integer'>${cuserId}</Value>
                        </Eq>
                        <Eq>
                            <FieldRef Name='Status'/>
                            <Value Type='Text'>Assigned</Value>
                        </Eq>
                    </And>
                </Where>
                <OrderBy>
                    <FieldRef Name='ID' Ascending='FALSE'/>
                </OrderBy>
            </Query>
            <RowLimit>500</RowLimit>
        </View>
    `);

    var ticketsItems = ticketsList.getItems(ticketsQuery);

    context.load(ticketsItems,
        'Include(Id,Title,TipoTicket,Categoria,Prioridad,Status,Atencion,FechaAtencion,Modified,ProcessManager,Author,Categoria_x003a_SLALow,Categoria_x003a_SLAMedium,Categoria_x003a_SLAHigh)'
    );

    context.executeQueryAsync(function () {

        var rows = "";
        var enumerator = ticketsItems.getEnumerator();

        while (enumerator.moveNext()) {

            var ticket = enumerator.get_current();

            var status = ticket.get_item("Status");
            var atencion = ticket.get_item("Atencion");
            var fechaAsignado = ticket.get_item("Modified");
            var fechaAtencion = ticket.get_item("FechaAtencion");

            var statusLabel = status;

            if (atencion === true) {
                statusLabel = "En Atención";
            }

            var horas = 0;
            var tiempoTexto = "-";

            /* SLA */

            if (status === "Assigned") {

                if (atencion === true && fechaAtencion) {

                    /* TIEMPO CONGELADO */
                    horas = getHoursDifference(fechaAsignado, fechaAtencion);
                    tiempoTexto = getTimeDifference(fechaAsignado, fechaAtencion);

                } else {

                    /* TIEMPO CORRIENDO */
                    horas = getHoursDifference(fechaAsignado);
                    tiempoTexto = getTimeDifference(fechaAsignado);

                }
            }

            const prioridad = ticket.get_item("Prioridad");

            var fecha = ticket.get_item("Modified");

            const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
            var mesIndex = new Date(fecha).getUTCMonth();
            var nombreMes = meses[mesIndex];

            const sla = {
                "Baja": ticket.get_item("Categoria_x003a_SLALow"),
                "Media": ticket.get_item("Categoria_x003a_SLAMedium"),
                "Alta": ticket.get_item("Categoria_x003a_SLAHigh")
            }[prioridad].get_lookupValue();

            let estilo = "";
            let diferencia = parseInt(sla) - horas;

            if (!atencion) {

                if (diferencia > 1)
                    estilo = "style='background-color:#68ddbd;border:1px solid #005227;color:#005227;border-radius:5px;padding:5px;display:inline-block'";
                else if (diferencia === 1)
                    estilo = "style='background-color:#ffff56;border:1px solid #ff9800;color:#ff9800;border-radius:5px;padding:5px;display:inline-block'";
                else
                    estilo = "style='background-color:#e66d52;border:1px solid #9d1d01;color:#9d1d01;border-radius:5px;padding:5px;display:inline-block'";
            }

            rows += `
            <tr id='tr_${ticket.get_id()}'>
                <td>${ticket.get_id()}</td>
                <td>${ticket.get_item("Title").slice(0,50)}</td>
                <td>${ticket.get_item("ProcessManager").get_lookupValue()}</td>
                <td>${ticket.get_item("TipoTicket")}</td>
                <td>${ticket.get_item("Categoria").get_lookupValue()}</td>
                <td><a ${estilo} href='Detalle de ticket en Aprobacion.aspx?idSol=${ticket.get_id()}'>${statusLabel}</a></td>
                <td>${prioridad}</td>
                <td>${tiempoTexto}</td>
                <td>${new Date(ticket.get_item("Modified")).getFullYear()}</td>
                <td>${nombreMes}</td>
                <td>${ticket.get_item("Author").get_lookupValue()}</td>
            </tr>`;
        }

        $("#dvLoading").hide();
        $("#tbCodigos_WP").show().find("tbody").html(rows);

        $('#tbCodigos_WP').DataTable({
            order: [[0,"desc"]],
            language: {
                lengthMenu:"Mostrar _MENU_ registros por página",
                zeroRecords:"No se encontraron resultados",
                info:"Mostrando página _PAGE_ de _PAGES_",
                infoEmpty:"No hay registros disponibles",
                infoFiltered:"(filtrado de _MAX_ registros totales)",
                search:"Buscar:",
                paginate:{
                    first:"Primero",
                    last:"Último",
                    next:"Siguiente",
                    previous:"Anterior"
                }
            }
        });

    }, onRequestFail);
}



/* ======================================================
   SEGUNDA TABLA
   Tickets por grupo o manager
====================================================== */

function loadTicketsPorGrupoOManger(cuserId) {

    var ticketsList = web.get_lists().getByTitle("Tickets");
    var GruposList = web.get_lists().getByTitle("Grupos");

    var GruposQuery = new SP.CamlQuery();

    GruposQuery.set_viewXml(`
        <View>
            <Query>
                <Where>
                    <Eq>
                        <FieldRef Name='Integrantes' LookupId='TRUE'/>
                        <Value Type='Integer'>${cuserId}</Value>
                    </Eq>
                </Where>
                <OrderBy>
                    <FieldRef Name='ID' Ascending='FALSE'/>
                </OrderBy>
            </Query>
            <RowLimit>500</RowLimit>
        </View>
    `);

    var GruposItems = GruposList.getItems(GruposQuery);
    context.load(GruposItems);

    context.executeQueryAsync(function () {

        var grupos = [];
        var enumerador = GruposItems.getEnumerator();

        while (enumerador.moveNext()) {
            grupos.push(enumerador.get_current().get_item("Title"));
        }

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

        context.load(ticketsItems,
        'Include(Id,Title,TipoTicket,Manager,Categoria,Categoria_x003a_SLALow,Categoria_x003a_SLAMedium,Categoria_x003a_SLAHigh,Prioridad,Status,Atencion,Modified,ProcessManager,TemplateConfiguracion,Author)');

        context.executeQueryAsync(function () {

            var rows = "";
            var enumerator = ticketsItems.getEnumerator();

            while (enumerator.moveNext()) {

                var ticket = enumerator.get_current();

                var status = ticket.get_item("Status");
                var atencion = ticket.get_item("Atencion");

                var statusLabel = status;

                if (atencion === true) {
                    statusLabel = "En Atención";
                }

                var manager = ticket.get_item("Manager");

                var perteneceGrupo = grupos.includes(status);
                var requiereManager = status === "Revisión de Manager";
                var esManagerAsignado = manager && manager.get_lookupId() === cuserId;

                if (!perteneceGrupo && !(requiereManager && esManagerAsignado)) continue;

                const horas = getHoursDifference(ticket.get_item("Modified"));
                const tiempoTexto = getTimeDifference(ticket.get_item("Modified"));
                const prioridad = ticket.get_item("Prioridad");

                var fecha = ticket.get_item("Modified");

                const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
                var mesIndex = new Date(fecha).getUTCMonth();
                var nombreMes = meses[mesIndex];

                const sla = {
                    "Baja": ticket.get_item("Categoria_x003a_SLALow"),
                    "Media": ticket.get_item("Categoria_x003a_SLAMedium"),
                    "Alta": ticket.get_item("Categoria_x003a_SLAHigh")
                }[prioridad].get_lookupValue();

                const diferencia = parseInt(sla) - horas;

                let estilo = "";

                if (diferencia > 1)
                    estilo = "style='background-color:#68ddbd;border:1px solid #005227;color:#005227;border-radius:5px;padding:5px;display:inline-block'";
                else if (diferencia === 1)
                    estilo = "style='background-color:#ffff56;border:1px solid #ff9800;color:#ff9800;border-radius:5px;padding:5px;display:inline-block'";
                else
                    estilo = "style='background-color:#e66d52;border:1px solid #9d1d01;color:#9d1d01;border-radius:5px;padding:5px;display:inline-block'";

                rows += `
                <tr id='tr_${ticket.get_id()}'>
                    <td>${ticket.get_id()}</td>
                    <td>${ticket.get_item("Title").slice(0,50)}</td>
                    <td>${cuser.get_title()}</td>
                    <td>${ticket.get_item("TipoTicket")}</td>
                    <td>${ticket.get_item("Categoria").get_lookupValue()}</td>
                    <td><a ${estilo} href='Detalle de ticket en Aprobacion.aspx?idSol=${ticket.get_id()}'>${statusLabel}</a></td>
                    <td>${prioridad}</td>
                    <td>${tiempoTexto}</td>
                    <td>${new Date(ticket.get_item("Modified")).getFullYear()}</td>
                    <td>${nombreMes}</td>
                    <td>${ticket.get_item("Author").get_lookupValue()}</td>
                </tr>`;
            }

            $("#dvLoadinGrupos").hide();
            $("#tbGrupos_WP").show().find("tbody").html(rows);

            $('#tbGrupos_WP').DataTable({
                order:[[0,"desc"]],
                language:{
                    lengthMenu:"Mostrar _MENU_ registros por página",
                    zeroRecords:"No se encontraron resultados",
                    info:"Mostrando página _PAGE_ de _PAGES_",
                    infoEmpty:"No hay registros disponibles",
                    infoFiltered:"(filtrado de _MAX_ registros totales)",
                    search:"Buscar:",
                    paginate:{
                        first:"Primero",
                        last:"Último",
                        next:"Siguiente",
                        previous:"Anterior"
                    }
                }
            });

        }, onRequestFail);

    }, onRequestFail);

}


function onRequestFail(sender, args){
    console.error("Error:", args.get_message());
}


function getTimeDifference(dateStr){

    const now = new Date();
    const created = new Date(dateStr);
    const diffMs = now - created;

    const minutes = Math.floor(diffMs/60000)%60;
    const hours = Math.floor(diffMs/(1000*60*60))%24;
    const days = Math.floor(diffMs/(1000*60*60*24));

    let result=[];

    if(days>0) result.push(`${days} día(s)`);
    if(hours>0) result.push(`${hours} hora(s)`);
    if(minutes>0) result.push(`${minutes} minuto(s)`);

    return result.length>0 ? result.join(", ") : "Justo ahora";
}


function getHoursDifference(dateStr){

    const now = new Date();
    const created = new Date(dateStr);
    const diffMs = now - created;

    return Math.floor(diffMs/(1000*60*60));

}