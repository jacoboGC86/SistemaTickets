var context, cuser, web;

$(document).ready(function () {
    var iebrowser = /*@cc_on!@*/false || !!document.documentMode;

    if (iebrowser) {
        alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
        window.location.href = "https://chimalliapps.sharepoint.com/ST/";
    }

    $("#DeltaPlaceHolderPageTitleInTitleArea a").text("Tickets por aprobar");
    $("#DeltaPlaceHolderPageTitleInTitleArea").show();

    context = new SP.ClientContext.get_current();
    web = context.get_web();
    cuser = web.get_currentUser();

    context.load(cuser);
    context.executeQueryAsync(function () {
        var cuserId = cuser.get_id();
        var ticketsList = web.get_lists().getByTitle("Tickets");
        var GruposList = web.get_lists().getByTitle("Grupos");

        var GruposQuery = new SP.CamlQuery();
        GruposQuery.set_viewXml(`
            <View>
              <Query>
               <Where>
                    <Eq>
                      <FieldRef Name='Integrantes' LookupId='TRUE' />
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
            var Grupos = [];
          //  console.log("Estos son los grupos",Grupos, "al que pertenece",cuserId);
            var GruposEnumerator = GruposItems.getEnumerator();

            while (GruposEnumerator.moveNext()) {
                var grupoItem = GruposEnumerator.get_current();
                Grupos.push(grupoItem.get_item("Title"));
            }

            // Cargar los tickets
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
            context.load(ticketsItems, 'Include(Id, Title, TipoTicket, Manager, Categoria, Categoria_x003a_SLALow, Categoria_x003a_SLAMedium, Categoria_x003a_SLAHigh, Prioridad, Status, Modified, ProcessManager, TemplateConfiguracion, Author, Atencion)');

            context.executeQueryAsync(function () {
                var tickets = [];
                var ticketEnumerator = ticketsItems.getEnumerator();
                var rows = "";

                while (ticketEnumerator.moveNext()) {
                    var ticketItem = ticketEnumerator.get_current();
                    var ticketItem = ticketEnumerator.get_current();
                    var Manager = ticketItem.get_item("Manager");
                    var status = ticketItem.get_item("Status");
                    var atencion = ticketItem.get_item("Atencion");

                    var statusLabel = status;

                    if (status === "Assigned" && atencion === true) {
                        statusLabel = "En atención";
                    }
                    var configJson = ticketItem.get_item("TemplateConfiguracion");
                    


                    try {
                  // limpia el JSON a texto plano
                            configJson = configJson && configJson.replace(/<\/?[^>]+(>|$)/g, "").trim();
                            if (!configJson) throw new Error("JSON vacío o con solo etiquetas");

                            var config = JSON.parse(configJson);
                            var approvalPath = config.approvalPath || [];

                            if (approvalPath.length === 0) continue;

                            var steps = approvalPath.map(s => s.stepName);

                            // verifica SÍ el paso actual incluye "Revisión de Manager"
                            var requiereManager = steps.includes("Revisión de Manager");

                            // Verifica si pertenece a algún grupo 
                            var perteneceAGrupo = steps.some(step => Grupos.includes(step));

                            // Verifica si es el manager es igual al usuario que inicia sesión
                            var esManagerAsignado = Manager && Manager.get_lookupId() === cuserId;

                            // Si no pertenece a grupos y tampoco es manager == omitir
                            if (!perteneceAGrupo && !(requiereManager && esManagerAsignado)) continue;

                            // entra solo si pertenece a un grupo o es el manager cuando corresponde
                            if (requiereManager && esManagerAsignado) {
                                console.log("Este ticket le corresponde al Manager actual:", Manager.get_lookupValue());
                            }
                
                        const tiempoTranscurridoTexto = getTimeDifference(ticketItem.get_item("Modified"));
                        const tiempoTranscurridoHoras = getHoursDifference(ticketItem.get_item("Modified"));
                        var fecha = ticketItem.get_item("Modified");
                        const meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                                       "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
                        var mesIndex = new Date(fecha).getUTCMonth();
                        var nombreMes = meses[mesIndex]; 
                        let estilo = "";
                
                        // Estilo acorde a los SLA
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
                            <td>${ cuser.get_title() }</td>
                            <td>${ ticketItem.get_item("TipoTicket") }</td>
                            <td>${ ticketItem.get_item("Categoria").get_lookupValue() }</td>
                            <td><a ${estilo} href='Detalle de ticket en Aprobacion.aspx?idSol=${ ticketItem.get_id() }'>${ statusLabel }</a></td>
                            <td>${ ticketItem.get_item("Prioridad") }</td>
                            <td>${ tiempoTranscurridoTexto }</td>
                            <td>${ new Date(ticketItem.get_item("Modified")).getFullYear() }</td>
                            <td>${ nombreMes }</td>
                        </tr>`;
                    } catch (e) {
                        console.warn("Error parseando JSON del ticket ID " + ticketItem.get_id(), e);
                    }
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

    return result.length > 0 ? result.join(", ") : "Justo ahora";
}

function getHoursDifference(dateStr) {
    const now = new Date();
    const created = new Date(dateStr);
    const diffMs = now - created;
    return Math.floor(diffMs / (1000 * 60 * 60)); // horas enteras
}
