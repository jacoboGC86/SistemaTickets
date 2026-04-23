var context,
  contextCol,
  cuser,
  web,
  idcode,
  subiendo = false,
  status = "",
  grupos = [],
  idManArea = 0,
  idProcessMan = 0,
  iamApproval = false,
  approvalGroupId = 0,
  idAprobacion = 0,
  t,
  posicionValorGlobal = "",
  categorias = [],
  TipoCategoria = "",
  template,
  TipoCategoria,
  Cat,
  cargandoTicket = false,
  recategorizando = false,
  attachmentFolders = [],
  NwTemplate,
  newProcessMan = 0,
  newManager = 0;

$(document).ready(function () {
  $("body").on("contextmenu", "img", function () {
    return false;
  });

  var iebrowser = /*@cc_on!@*/ false || !!document.documentMode;

  if (iebrowser) {
    alert(
      "¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma",
    );
    window.location.href = "https://promagroupmex.sharepoint.com/sites/Paperless/ST/";
  }

  const tooltipTriggerList = document.querySelectorAll(
    '[data-bs-toggle="tooltip"]',
  );
  tooltipTriggerList.forEach((el) => {
    new bootstrap.Tooltip(el);
  });
  initializePeoplePicker("peoplePickerReassign");

  context = new SP.ClientContext.get_current();
  contextCol = new SP.ClientContext(
    "https://promagroupmex.sharepoint.com/sites/Paperless/VR/",
  );
  web = context.get_web();
  cuser = web.get_currentUser();

  context.load(web);
  context.load(cuser);

  context.executeQueryAsync(function (sender, args) {
    var params = getParameters(window.location.href);

    for (var i = 0; i < params.length; i++) {
      if (params[i].param === "idSol") {
        idcode = params[i].value;
        break;
      }
    }

    cargarDatos(idcode);

    loadTicketComments(function (attachmentFolders) {
      loadAttachments(attachmentFolders, 0, function () {
        console.log("Todo listo");
      });
    });
  }, onRequestFail);
});
function cargarDatos(idcode) {
  var oList = context.get_web().get_lists().getByTitle("Tickets");
  var oListItem = oList.getItemById(idcode);

  context.load(
    oListItem,
    "Id",
    "Title",
    "Status",
    "ProcessManager",
    "Manager",
    "Author",
    "Department",
    "Created",
    "TipoTicket",
    "Categoria",
    "Descripcion",
    "Planta",
    "Prioridad",
    "ANombreDe",
    "TemplateConfiguracion",
    "Modified",
    "SolucionDetallada",
    "Reaperturado",
    "MotivoReapertura",
    "Atencion",
    "ImpactoSeguridadInformacion",
    "DescripcionImpactoSeguridad"
  );

  context.executeQueryAsync(function (sender, args) {
    cargandoTicket = true;
    recategorizando = false;

    $("#txtFolio").val(oListItem.get_id());

    $("#txtTitle").val(oListItem.get_item("Title"));
    $("#txtSolicitante").val(oListItem.get_item("Author").get_lookupValue());
    $("#cmbDepartamento").val(oListItem.get_item("Department"));
    $("#dtFechaRegistro").val(moment(oListItem.get_item("Created")).format("DD/MM/yyyy hh:mm a"));
    
    var tipoGuardado = oListItem.get_item("TipoTicket");

    if(tipoGuardado === "Change") {
      $("#cmbImpactoSeguridadInformacion").val(oListItem.get_item("ImpactoSeguridadInformacion"));
      $("#txtDescripcionImpactoSeguridad").val(oListItem.get_item("DescripcionImpactoSeguridad"));

      $("#dvChangeValidacion").show();
      
      if(oListItem.get_item("ImpactoSeguridadInformacion") === "Sí") {
        $("#dvImpactosSeguridad").show();
      }
    }

    var categoriaGuardada = oListItem.get_item("Categoria")
      ? oListItem.get_item("Categoria").get_lookupId()
      : null;

    loadCategorias(function () {
      loadTipo();

      // 1.Seleccionar tipo
      $("#cmbType").val(tipoGuardado).trigger("change");
      setTimeout(function () {
        if (categoriaGuardada) {
          $("#cmbCategorias").val(categoriaGuardada);
        }
        $("#cmbCategorias").prop("disabled", true);
        $("#cmbType").prop("disabled", true);

        cargandoTicket = false;
      }, 200);
    });

    $("#txtDescripcion").val(oListItem.get_item("Descripcion"));
    $("#cmbPlanta").val(oListItem.get_item("Planta"));
    $("#cmbUrgency").val(oListItem.get_item("Prioridad"));
    var Motivo = oListItem.get_item("MotivoReapertura");

    if (Motivo && Motivo.trim() !== "") {
      $("#dvReapertura").show();
    } else {
      $("#dvReapertura").hide();
    }
    var motivo = oListItem.get_item("MotivoReapertura");
    if (motivo) {
      var textoMotivo = $("<div>").html(motivo).text();
      $("#txtMotivo").val(textoMotivo);
    }

    if (oListItem.get_item("ANombreDe") != null) {
      $("#dvANombreDe").show();
      $("#txtANombreDe").val(oListItem.get_item("ANombreDe").get_lookupValue());
    }

    t = JSON.parse(oListItem.get_item("TemplateConfiguracion"));
    var ft = t.formTemplate;

    for (var j = 0; j < ft.length; j++) {
      var c = "";

      switch (ft[j].tag) {
        case "h3":
        case "p":
          c =
            "<div class='col-md-12'><" +
            ft[j].tag +
            ">" +
            ft[j].text +
            "</" +
            ft[j].tag +
            "></div>";
          break;
        case "checkbox":
          c =
            "<div class='col-md-12'><div class='form-check'><input type='checkbox' disabled class='form-check-input' id='" +
            ft[j].id +
            "' " +
            (ft[j].checked ? "checked" : "") +
            " /><label class='form-check-label' for='" +
            ft[j].id +
            "' >" +
            ft[j].text +
            "</label></div></div>";
          break;
        case "br":
          c = "<br/><br/>";
          break;
        case "textarea":
          c =
            "<div class='col-md-12'><div class='form-group'><label for='" +
            ft[j].id +
            "'>" +
            ft[j].text +
            "</label><textarea class='form-control' id='" +
            ft[j].id +
            "' disabled >" +
            ft[j].value +
            "</textarea></div></div>";
          break;
      }

      $("#dvDetail").append(c);
    }

    var ap = t.approvalPath;
    $("#dvStatus").html("");
    status = oListItem.get_item("Status");

    if (oListItem.get_item("Manager") != null)
      idManArea = oListItem.get_item("Manager").get_lookupId();

    var renderedSteps = [];
    var totalSteps = ap.length;

    for (let j = 0; j < totalSteps; j++) {
      if (oListItem.get_item("Status") === ap[j].stepName) {
        aprobador = ap[j];
      }

      (function (j, stepName, useUserDept) {
        if (useUserDept) {
          var managerName = oListItem.get_item("Manager")
            ? oListItem.get_item("Manager").get_lookupValue()
            : "";
          storeRenderedStep(j, stepName, managerName);

          if (
            cuser.get_id() === oListItem.get_item("Manager").get_lookupId() &&
            oListItem.get_item("Status") == stepName
          ) {
            iamApproval = true;
          }
        } else {
          var ctx = SP.ClientContext.get_current();
          var web = ctx.get_web();
          var list = web.get_lists().getByTitle("Grupos");

          var camlQuery = new SP.CamlQuery();
          camlQuery.set_viewXml(
            "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" +
              stepName +
              "</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>",
          );

          var items = list.getItems(camlQuery);
          ctx.load(items);

          ctx.executeQueryAsync(
            function () {
              var enumerator = items.getEnumerator();
              var integranteNombre = "";

              if (enumerator.moveNext()) {
                var item = enumerator.get_current();
                var integrantes = item.get_item("Integrantes");

                if (integrantes && integrantes.length > 0) {
                  integranteNombre = integrantes[0].get_lookupValue();
                }
              }

              storeRenderedStep(j, stepName, integranteNombre);
            },
            function (sender, args) {
              console.error("Error consultando grupo:", args.get_message());
              storeRenderedStep(j, stepName, "Error");
            },
          );
        }
      })(j, ap[j].stepName, ap[j].useUserDepartment);
    }

    function renderStep(index, stepName, nombre) {
      var estadoActual = oListItem.get_item("Status");
      var pendiente = estadoActual == stepName ? "pendienteRev" : "";

      var html =
        "<div class='col-md-3' id='step_" +
        index +
        "' >" +
        "<span class='lblResponsable " +
        pendiente +
        "'>" +
        stepName +
        "<br/>" +
        nombre +
        "</span></div>";

      $("#dvStatus").append(html);
    }

    // Almacena los pasos devueltos
    function storeRenderedStep(index, stepName, nombre) {
      renderedSteps[index] = {
        index: index,
        stepName: stepName,
        nombre: nombre,
      };

      if (renderedSteps.filter(Boolean).length === totalSteps) {
        // Ya se completaron todos los pasos
        renderedSteps.forEach(function (step) {
          renderStep(step.index, step.stepName, step.nombre);
        });

        renderProcessManager();
      }
    }

    //TODO: Si el status es Asignado y soy el Process Manager => mostrar los botones de aprobación.

    function renderProcessManager() {
      var status = oListItem.get_item("Status");
      var reaperturado = oListItem.get_item("Reaperturado");
      var processManager = oListItem
        .get_item("ProcessManager")
        .get_lookupValue();

      var claseEstado = "";

      if (status === "Assigned") {
        if (reaperturado === true) {
          claseEstado = "reaperturaRev";
        } else {
          claseEstado = "pendienteRev";
        }
      } else if (status === "Cerrado") {
        claseEstado = "aprobadoRev";
      }

      var fechaParaMostrar = "";

      if (status === "Cerrado") {
        fechaParaMostrar =
          "<br/>" +
          moment(oListItem.get_item("Modified")).format("DD/MM/yyyy hh:mm a");
      }

      $("#dvStatus").append(
        "<div class='col-md-3' id='step_PM'>" +
          "<span class='lblResponsable " +
          claseEstado +
          "'>" +
          "Ticket service<br/>" +
          processManager +
          fechaParaMostrar +
          "</span></div>",
      );
    }

    if (ap.length == 0) renderProcessManager();

    idProcessMan = oListItem.get_item("ProcessManager").get_lookupId();

    if (
      cuser.get_id() === idProcessMan &&
      oListItem.get_item("Status") == "Assigned"
    ) {
      console.log(
        oListItem.get_item("ProcessManager").get_lookupId(),
        "=",
        cuser.get_id(),
      );
      iamApproval = true;
      $("#btnCancelar").hide();
      $("#btnEnviar")
        .removeClass("btn-primary")
        .addClass("btn-success")

        .text("Cerrar Ticket");
      $("#dvRegistrarKB").show();
      $("#dvOptionBar").show();

      $("#btnHabilitRecategorizar").show();

      $(document).ready(function () {
        $("#cmbRegistrarKB").on("change", function () {
          if ($(this).val() === "Sí") {
            $("#dvSolucionDetallada").show();
          } else {
            $("#dvSolucionDetallada").hide();
          }
        });
      });
    }

    if (
      cuser.get_id() === idProcessMan &&
      oListItem.get_item("Status") == "Assigned" &&
      oListItem.get_item("Atencion") == false
    ) {
      $("#btnAtencion").show();
      console.log(cuser.get_id(), idProcessMan);
      console.log(oListItem.get_item("Status"), "Assigned");
      console.log(oListItem.get_item("Atencion"), false);
    }

    if (oListItem.get_item("Status") == "Cerrado") {
      $("#dvSolucionDetallada").show();
      $("#txtSolucionDetallada").prop("disabled", true);
      $("#txtSolucionDetallada").val(oListItem.get_item("SolucionDetallada"));
    }

    var fechaModificacion = new Date(oListItem.get_item("Modified"));
    var hoy = new Date();
    // calcula diferencia en días
    var diferenciaTiempo = hoy - fechaModificacion;
    var diferenciaDias = diferenciaTiempo / (1000 * 60 * 60 * 24);

    if (status == "Cerrado" && diferenciaDias >= 30) {
      $("#btnReaperturar").show();
    }

    var oListCat = context.get_web().get_lists().getByTitle("Categorias");
    var oListItemCat = oListCat.getItemById(
      oListItem.get_item("Categoria").get_lookupId(),
    );

    context.load(oListItemCat);
    context.executeQueryAsync(function (sender, args) {
      var attachmentFolder = web.getFolderByServerRelativeUrl(
        web.get_url() + "/Expediente/" + idcode,
      );
      var attachmentFiles = attachmentFolder.get_files();

      //Load attachments
      context.load(attachmentFiles);
      context.executeQueryAsync(
        function (sender, args) {
          var cnt = attachmentFiles.get_count();

          for (var itr = 0; itr < cnt; itr++) {
            var file = attachmentFiles.itemAt(itr);
            var url = file.get_serverRelativeUrl();
            var name = file.get_name();
            var remover =
              cuser.get_id() === oListItem.get_item("Author").get_lookupId()
                ? "<a href='#id_' class='lkRemoverAdjunto' data-srurl='" +
                  url +
                  "' >Remover</a> - "
                : "";

            $("#ulAdjuntos").append(
              "<li class='list-group-item'>" +
                remover +
                "<a href='" +
                url +
                "' target='_blank'>" +
                name +
                "</a></li>",
            );
          }

          continuarProceso();
        },
        function () {
          continuarProceso();
        },
      );
    });
  }, onRequestFail);
}

function continuarProceso() {
  loadApprovalProcess(function (attachmentFolders) {
    loadAttachments(attachmentFolders, 0, function () {
      asignarEventos();
    });
  });
}
function formatDate(date) {
  const d = new Date(date);
  const day = ("0" + d.getDate()).slice(-2);
  const month = ("0" + (d.getMonth() + 1)).slice(-2);
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
}
function loadApprovalProcess(callback) {
  var oList = context.get_web().get_lists().getByTitle("Aprobaciones");
  var camlQuery = new SP.CamlQuery();

  camlQuery.set_viewXml(`<View><Query>
    <Where>
      <Eq>
        <FieldRef Name="Ticket" LookupId="TRUE" />
        <Value Type="Lookup">${idcode}</Value>
      </Eq>
    </Where>
    <OrderBy>
      <FieldRef Name="ID" Ascending="TRUE" />
    </OrderBy>
  </Query></View>`);

  var collListItem = oList.getItems(camlQuery);
  context.load(collListItem);

  context.executeQueryAsync(function () {
    var listItemEnumerator = collListItem.getEnumerator();
    var j = 0;
    var attachmentFolders = [];
    var comentariosGenerados = false;
    var pendientesPorGrupo = [];

    while (listItemEnumerator.moveNext()) {
      var oListItem = listItemEnumerator.get_current();
      var idAprov = oListItem.get_id();
      var resultado = oListItem.get_item("Resultado");
      var resultadoClass = "";
      var grupoId = 0;

      if(oListItem.get_item("Responsable") !== null && oListItem.get_item("Comentarios") !== null) {
        var html = `
          <div id="ap${idAprov}" class="comentario">
          <strong>${oListItem.get_item("Title")}</strong> / 
          ${oListItem.get_item("Responsable").get_lookupValue()}
          <br/>${moment(oListItem.get_item("Modified")).format("DD/MM/yyyy hh:mm a")}
          <br/>${oListItem.get_item("Comentarios")}
          </div>
        `;

        $("#dvComentarios").append(html);
      }

      switch (resultado) {
        case "Pendiente":
          grupoId =
            oListItem.get_item("Grupo") != null
              ? oListItem.get_item("Grupo").get_lookupId()
              : 0;
          approvalGroupId = grupoId;
          idAprobacion = oListItem.get_id();
          break;
        case "Aprobado":
          resultadoClass = "aprobadoRev";
          break;
        case "Rechazado":
          resultadoClass = "rechazadoRev";
          break;
      }

      var stepElement = $(`#step_${j} .lblResponsable`);

      if (stepElement.length > 0) {
        stepElement.addClass(resultadoClass);

        if (
          resultado === "Pendiente" &&
          grupoId > 0 &&
          !oListItem.get_item("Responsable")
        ) {
          pendientesPorGrupo.push({
            grupoId: grupoId,
            stepElement: stepElement,
            title: oListItem.get_item("Title"),
          });
        } else {
          var fecha = "",
            nomResp = "";

          if (resultado === "Pendiente") {
            nomResp = `<a href='https://teams.microsoft.com/l/chat/0/0?users=${oListItem.get_item("Responsable").get_email()}&topicName=${encodeURIComponent("Sistema de Tickets")}&message=${encodeURIComponent("Hola! Buen día.<br/><br/>¿Podrías ayudarme por favor, en revisar el siguiente ticket? Necesita tu aprobación para avanzar:<br/><br/>Título: <b>" + $("#txtTitle").val() + "</b><br/><br/>Puedes acceder con la siguiente liga:<br/>https://promagroupmex.sharepoint.com/sites/Paperless/ST/Pages/Detalle de ticket en Aprobacion.aspx?idSol=" + idcode)}' target='_blank' title='Escríbeme por Teams'>${oListItem.get_item("Responsable").get_lookupValue()}</a><br/>`;
          } else {
            nomResp = oListItem.get_item("Responsable").get_lookupValue();
            fecha = `<br/>${moment(oListItem.get_item("Modified")).format("DD/MM/yyyy hh:mm a")}`;
          }

          stepElement.html(`
            ${oListItem.get_item("Title")}<br/>
            ${nomResp}
            ${fecha}
          `);
        }
      }

      j++;
    }

    if (pendientesPorGrupo.length > 0) {
      var oListGpo = context.get_web().get_lists().getByTitle("Grupos");
      var oListItemGpo = oListGpo.getItemById(pendientesPorGrupo[0].grupoId);

      context.load(oListItemGpo, "Integrantes");
      context.executeQueryAsync(
        function () {
          try {
            var integrantes = oListItemGpo.get_item("Integrantes") || [];

            pendientesPorGrupo[0].stepElement.html(`
            ${pendientesPorGrupo[0].title}<br/>
            <a href='https://teams.microsoft.com/l/chat/0/0?users=${integrantes[0].get_email()}&topicName=${encodeURIComponent("Sistema de Tickets")}&message=${encodeURIComponent("Hola! Buen día.<br/><br/>¿Podrías ayudarme por favor, en revisar el siguiente ticket? Necesita tu aprobación para avanzar:<br/><br/>Título: <b>" + $("#txtTitle").val() + "</b><br/><br/>Puedes acceder con la siguiente liga:<br/>https://promagroupmex.sharepoint.com/sites/Paperless/ST/Pages/Detalle de ticket en Aprobacion.aspx?idSol=" + idcode)}' target='_blank' title='Escríbeme por Teams'>${integrantes[0].get_lookupValue()}</a>`);

            for (var i = 0; i < integrantes.length; i++) {
              if (integrantes[i].get_lookupId() === cuser.get_id()) {
                iamApproval = true;
              }
            }
          } catch (exception) {}

          callback(attachmentFolders);
        },
        function (sender, args) {
          console.error(
            "Error al obtener integrantes del grupo:",
            args.get_message(),
          );
          callback(attachmentFolders);
        },
      );
    } else {
      callback(attachmentFolders);
    }
  }, onRequestFail);
}
function loadTicketComments(callback) {
  var oList = context.get_web().get_lists().getByTitle("Comentarios");
  var camlQuery = new SP.CamlQuery();

  camlQuery.set_viewXml(`
    <View>
      <Query>
        <Where>
          <Eq>
            <FieldRef Name="Ticket" LookupId="TRUE" />
            <Value Type="Integer">${idcode}</Value>
          </Eq>
        </Where>
        <OrderBy>
          <FieldRef Name="ID" Ascending="TRUE" />
        </OrderBy>
      </Query>
    </View>
  `);

  var items = oList.getItems(camlQuery);
  context.load(items);

  context.executeQueryAsync(function () {
    //$("#dvComentarios").empty();

    var enumerator = items.getEnumerator();
    var attachmentFolders = [];
    var hayComentarios = false;

    while (enumerator.moveNext()) {
      var item = enumerator.get_current();
      var idComentario = item.get_id();
      var comentario = item.get_item("Comentario") || "";

      if (comentario.trim() === "") continue;

      hayComentarios = true;

      var responsable = item.get_item("Responsable")
        ? item.get_item("Responsable").get_lookupValue()
        : "Sistema";

      var StatusTicket = item.get_item("Title");

      var fecha = moment(item.get_item("Modified")).format("DD/MM/YYYY HH:mm");

      var commentId = `comment_${idComentario}`;

      var html = `
        <div id="${commentId}" class="comentario">
         <strong>${StatusTicket}</strong> / 
        ${responsable}
        <br/>${fecha}
        <br/>${comentario}
        </div>
      `;

      $("#dvComentarios").append(html);

      attachmentFolders.push({
        url: `Lists/Comentarios/Attachments/${idComentario}`,
        commentId: commentId,
      });
    }

    $("#dvComentarios").show();

    callback(attachmentFolders);
  }, onRequestFail);
}
function loadAttachments(attachmentFolders, index, callback) {
  // Carga de adjuntos
  if (attachmentFolders.length > index) {
    var folderData = attachmentFolders[index];
    var attachmentFolder = web.getFolderByServerRelativeUrl(folderData.url);
    var attachmentFiles = attachmentFolder.get_files();

    context.load(attachmentFiles);
    context.executeQueryAsync(
      function () {
        index = index + 1;

        for (var itr = 0; itr < attachmentFiles.get_count(); itr++) {
          var fileUrl = attachmentFiles.itemAt(itr).get_serverRelativeUrl();
          var fileName = attachmentFiles.itemAt(itr).get_name();
          var plantilla = `
          <div class='col-md-4'>
            <a href='${fileUrl}' target='_blank' >${fileName}</a>
          </div>
          <hr>`;

          $(`#${folderData.commentId}`).append(plantilla);
        }

        loadAttachments(attachmentFolders, index, callback);
      },
      function (sender, args) {
        index = index + 1;
        var plantilla = `<hr>`;

        $(`#${folderData.commentId}`).append(plantilla);
        loadAttachments(attachmentFolders, index, callback);
      },
    );
  } else {
    callback();
  }
}
function cargarDepartamentos(callback) {
  var oListConf = contextCol
    .get_web()
    .get_lists()
    .getByTitle("Configuraciones");
  var camlQueryConf = new SP.CamlQuery();

  camlQueryConf.set_viewXml(`
      <View>
          <Query>
              <Where>
                  <Eq>
                      <FieldRef Name='Title' />
                      <Value Type='Text'>Departamento</Value>
                  </Eq>
              </Where>
              <OrderBy>
                  <FieldRef Name='ID' Ascending='FALSE'/>
              </OrderBy>
          </Query>
      </View>
  `);
  var collListItemConf = oListConf.getItems(camlQueryConf);

  contextCol.load(collListItemConf);
  contextCol.executeQueryAsync(
    function (sender, args) {
      var listItemEnumerator = collListItemConf.getEnumerator();

      while (listItemEnumerator.moveNext()) {
        var oListItem = listItemEnumerator.get_current();
        var CampoPersona = oListItem.get_item("Personas");

        if (CampoPersona && CampoPersona.length > 0) {
          CampoPersona.forEach(function (personaItem) {
            var personaId = personaItem.get_lookupId();
            var personaName = personaItem.get_lookupValue();

            $("#cmbDepartamento").append(
              "<option value='" +
                oListItem.get_id() +
                "' data-id='" +
                personaId +
                "' data-valor='" +
                oListItem.get_item("Valor") +
                "'>" +
                oListItem.get_item("Valor") +
                " - " +
                personaName +
                "</option>",
            );
          });
        } else {
          console.warn("El campo 'Personas' no tiene datos.");
        }
      }

      callback();
    },
    function (sender, args) {
      console.error("Error al obtener los elementos: " + args.get_message());
    },
  );
}

function asignarEventos() {
  if (status == "Assigned" || status == "Cerrado") {
    $("#pHeaderTicket").hide();
  }

  if (iamApproval) {
    $("#dvBotonera").show();
    $("#dvComments").show();
    $("#dvArchivo").show();
  }
  if (
    (status == "Assigned" && iamApproval) ||
    (status == "Revisión de Manager" && iamApproval)
  ) {
    $("#btnOpenPopUp").show();
  }

  $("#txtComentarios").on("change", function () {
    if ($(this).val().length != 0) {
      $(this).addClass("is-valid").removeClass("is-invalid");
    } else {
      $(this).addClass("is-invalid").removeClass("is-valid");
    }
  });

  $("#cmbImpactoSeguridadInformacion").on("change", function() {
    if($(this).val() === "Sí") {
      $("#dvImpactosSeguridad").show();
    } else {
      $("#dvImpactosSeguridad").hide();
    }
  });

  // Botones  (Cerrar, Enviar, Cancelar)
  $("#btnCerrar").click(function () {
    window.location.href =
      "https://promagroupmex.sharepoint.com/sites/Paperless/ST/Pages/Consulta%20Tickets%20ProcessManager.aspx";
  });

  $("#btnEnviar").click(function () {
    if (
      status == "Assigned" &&
      $("#txtSolucionDetallada").val() == "" &&
      $("#cmbRegistrarKB").val() == "Sí"
    ) {
      alertify.alert(
        "Sistema de Ticket",
        "Por favor, registre la solución detallada del ticket.",
      );
      return;
    }

    $("#btnEnviar").prop("disabled", true);

    var oListAp = context.get_web().get_lists().getByTitle("Aprobaciones");
    var oListItem = oListAp.getItemById(idAprobacion);

    context.load(oListItem);

    context.executeQueryAsync(function () {
      oListItem.set_item("Resultado", "Aprobado");

      var userField = new SP.FieldUserValue();
      userField.set_lookupId(cuser.get_id());
      oListItem.set_item("Responsable", userField);

      oListItem.update();

      context.executeQueryAsync(function () {
        var ap = t.approvalPath || [];
        var apS = null;

        for (var j = 0; j < ap.length; j++) {
          if (status === ap[j].stepName && ap.length > j + 1) {
            apS = ap[j + 1];
            break;
          }
        }

        var oListT = context.get_web().get_lists().getByTitle("Tickets");
        var oListItemT = oListT.getItemById(idcode);

        // Actualiza Status del Ticket
        if (apS != null) {
          oListItemT.set_item("Status", apS.stepName);
        } else if (status != "Assigned") {
          oListItemT.set_item("Status", "Assigned");
        } else {
          oListItemT.set_item("Status", "Cerrado");
          oListItemT.set_item("RegistrarKB", $("#cmbRegistrarKB").val());
          oListItemT.set_item(
            "SolucionDetallada",
            $("#txtSolucionDetallada").val(),
          );
        }

        oListItemT.update();
        context.load(oListItemT);

        context.executeQueryAsync(function () {
          function procederConFlujo() {
            // Si ya estaba Assigned → cerrar
            if (status == "Assigned") {
              alertify.alert(
                "Sistema de Ticket",
                "El ticket se ha cerrado correctamente.",
                function () {
                  window.location.href = "Consulta Tickets ProcessManager.aspx";
                },
              );
            } else {
              // Crea el siguiente aprobación
              var itemCreateInfoAp = new SP.ListItemCreationInformation();
              var oListItemAp = oListAp.addItem(itemCreateInfoAp);

              oListItemAp.set_item(
                "Title",
                apS != null ? apS.stepName : "Assigned",
              );
              oListItemAp.set_item("Resultado", "Pendiente");
              oListItemAp.set_item("Ticket", idcode);

              if (apS != null && apS.approvalGroup != null) {
                oListItemAp.set_item("Grupo", apS.approvalGroup);
              } else {
                var userFieldNext = new SP.FieldUserValue();
                userFieldNext.set_lookupId(
                  apS != null ? idManArea : idProcessMan,
                );
                oListItemAp.set_item("Responsable", userFieldNext);
              }

              oListItemAp.update();
              context.load(oListItemAp);

              context.executeQueryAsync(function () {
                alertify.alert(
                  "Sistema de Ticket",
                  "Solicitud aprobada exitosamente.",
                  function () {
                    window.location.href = "Consulta Tickets ProcessManager.aspx";
                  },
                );
              }, onRequestFail);
            }
          }

          var comentarioText = $("#txtComentarios").val().trim();

          if (comentarioText === "") {
            procederConFlujo();
            return;
          }

          var listComentarios = context.get_web().get_lists().getByTitle("Comentarios");
          var itemCreateInfoCom = new SP.ListItemCreationInformation();
          var itemCom = listComentarios.addItem(itemCreateInfoCom);

          itemCom.set_item("Ticket", idcode);
          itemCom.set_item("Comentario", comentarioText);
          itemCom.set_item("Title", status);

          var userValueCom = new SP.FieldUserValue();
          userValueCom.set_lookupId(cuser.get_id());
          itemCom.set_item("Responsable", userValueCom);

          itemCom.update();

          context.executeQueryAsync(function () {
            var comentarioId = itemCom.get_id();
            var files = $("#fileDoctoIT")[0].files;

            if (files.length === 0) {
              procederConFlujo();
              return;
            }

            setTimeout(function () {
              subirArchivos(
                files,
                0,
                [],
                web.get_serverRelativeUrl(),
                "Comentarios",
                comentarioId,
                function () {
                  procederConFlujo();
                },
              );
            }, 300);
          }, onRequestFail);
        }, onRequestFail);
      }, onRequestFail);
    }, onRequestFail);
  });
  //sube comentarios
  $("#btnUpComment")
    .off("click")
    .on("click", function () {
      if ($("#txtComentarios").val().trim() === "") {
        alertify.alert(
          "Sistema de Ticket",
          "Por favor, agregue un comentario.",
        );
        return;
      }

      var list = context.get_web().get_lists().getByTitle("Comentarios");
      var itemCreateInfo = new SP.ListItemCreationInformation();
      var item = list.addItem(itemCreateInfo);
      console.log(status);

      item.set_item("Ticket", idcode);
      item.set_item("Comentario", $("#txtComentarios").val());
      item.set_item("Title", status);

      var userValue = new SP.FieldUserValue();
      userValue.set_lookupId(cuser.get_id());
      item.set_item("Responsable", userValue);

      item.update();

      context.executeQueryAsync(function () {
        var comentarioId = item.get_id();
        var files = $("#fileDoctoIT")[0].files;

        if (files.length === 0) {
          alertify.success("Comentario guardado");

          setTimeout(function () {
            location.reload();
          }, 1000);

          return;
        }

        setTimeout(function () {
          subirArchivos(
            files,
            0,
            [],
            web.get_serverRelativeUrl(),
            "Comentarios",
            comentarioId,
            function () {
              alertify.success("Comentario y archivos guardados correctamente");

              $("#txtComentarios").val("");
              $("#fileDoctoIT").val("");
              $("#ulFormatIT").html("");
              $("#dvRepresentacionIT img").each(function () {
                URL.revokeObjectURL(this.src);
              });
              $("#dvRepresentacionIT").html("");

              setTimeout(function () {
                location.reload();
              }, 1000);
            },
          );
        }, 300);
      }, onRequestFail);
    });

  $("#btnLoadAttachmentsIT").on("click", function () {
    $("#fileDoctoIT").click();
  });

  $("#fileDoctoIT").on("change", function () {
    if ($(this).val() === "") {
      return;
    }

    $("#ulFormatIT").html("");
    $("#dvRepresentacionIT").html("");
    $("#RemoveFileIT").show();

    for (var i = 0; i < this.files.length; i++) {
      var fileName = this.files[i].name;
      var extension = fileName.split(".").pop().toLowerCase();

      $("#ulFormatIT").append(
        "<li class='list-group-item'>" + fileName + "</li>",
      );

      if (extension === "jpg" || extension === "png" || extension === "jpeg") {
        var plantilla = `
          <div class='col-md-4'>
            <img src='${URL.createObjectURL(this.files[i])}' class='img-thumbnail' />
          </div>`;
        $("#dvRepresentacionIT").append(plantilla);
      }
    }
  });

  $(".lkRemoverAdjunto").on("click", function (e) {
    e.preventDefault();

    if (confirm("¿Desea borrar el archivo?")) {
      var file = context
        .get_web()
        .getFileByServerRelativeUrl($(this).data("srurl"));

      file.deleteObject();

      $(this).parent().remove();

      context.executeQueryAsync(function () {}, onRequestFail);
    }
  });

  $("#RemoveFileIT").on("click", function () {
    $("#fileDoctoIT").val("");
    $("#ulFormatIT").html("");
    $("#dvRepresentacionIT").html("");
    $(this).hide();
  });

  $("#btnCancelar").click(function () {
    if (!confirm("Está por rechazar la solicitud. ¿Desea continuar?")) return;

    function onRequestFail(sender, args) {
      console.error("Error: " + args.get_message());
      alertify.alert(
        "Error",
        "Ocurrió un error al procesar la solicitud: " + args.get_message(),
      );
    }
    //1:Primera actualización a la lista aprobaciones
    var oListAp = context.get_web().get_lists().getByTitle("Aprobaciones");
    var oListItem = oListAp.getItemById(idAprobacion);

    if ($("#txtComentarios").val() == "") {
      alertify.alert("Sistema de Ticket", "Por favor, agregue un comentario.");
      $("#btnEnviar").prop("disabled", false);
      return;
    }
    context.load(oListItem);
    context.executeQueryAsync(function () {
      // Actualiza la solicitud
      oListItem.set_item("Comentarios", $("#txtComentarios").val());
      oListItem.set_item("Resultado", "Rechazado");
      oListItem.set_item("Responsable", cuser.get_id());
      oListItem.update();

      context.executeQueryAsync(function () {
        console.log("Elemento de aprobación actualizado correctamente.");

        // Valida y sube los archivos si existen
        var archivos = $("#fileDoctoIT")[0].files;
        if (archivos.length > 0) {
          subirArchivos(
            archivos,
            0,
            [],
            web.get_url(),
            web.get_serverRelativeUrl(),
            oListItem.get_id(),
            function () {
              console.log("Archivos subidos correctamente.");
            },
          );
        } else {
          console.log("No hay archivos para subir.");
        }

        // 2: Segunda actualización de tickets
        var oListTickets = context.get_web().get_lists().getByTitle("Tickets");
        var oListItemTicket = oListTickets.getItemById(idcode);

        context.load(oListItemTicket);
        context.executeQueryAsync(function () {
          if (iamApproval) {
            oListItemTicket.set_item("Status", "Rechazada/Cancelada");
          } else {
            console.warn("El usuario no tiene permisos de aprobación.");
          }

          oListItemTicket.update();

          context.executeQueryAsync(function () {
            alertify.alert(
              "Sistema de Ticket",
              "Solicitud cancelada/rechazada.",
              function () {
                window.location.href = "Consulta Tickets.aspx";
              },
            );
          }, onRequestFail);
        }, onRequestFail);
      }, onRequestFail);
    }, onRequestFail);
  });
  // REASIGNAR RESPONSABLE

  $("body").on("click", "#btnOpenPopUp", function () {
    var panel = $("#dvReassignPanel");

    if (panel.is(":visible")) {
      panel.slideUp();
      return;
    }

    cerrarAcciones();
    panel.slideDown();
  });

  $("body").on("click", "#btnCancelReassign", function (e) {
    e.preventDefault();

    var picker =
      SPClientPeoplePicker.SPClientPeoplePickerDict[
        "peoplePickerReassign_TopSpan"
      ];

    if (picker) {
      picker.DeleteProcessedUser();
    }

    $("#dvReassignPanel").slideUp();
  });
  $("body").on("click", "#btnConfirmReassign", function () {
    var peoplePicker =
      SPClientPeoplePicker.SPClientPeoplePickerDict[
        "peoplePickerReassign_TopSpan"
      ];
    var users = peoplePicker.GetAllUserInfo();

    if (users.length === 0) {
      alertify.error("Por favor, seleccione un usuario.");
      return;
    }

    var userLogin = users[0].Key;
    var user = context.get_web().ensureUser(userLogin);
    context.load(user);

    context.executeQueryAsync(function () {
      var userValue = new SP.FieldUserValue();
      userValue.set_lookupId(user.get_id());

      actualizarResponsable(userValue);

      $("#dvReassignPanel").slideUp();
    }, onRequestFail);
  });

  //RECATEGORIZAR TICKET
  $("#btnHabilitRecategorizar").on("click", function () {
    var panel = $("#dvBotoneraRecategorizar");

    if (panel.is(":visible")) {
      cerrarAcciones();
      return;
    }

    cerrarAcciones();

    recategorizando = true;

    $("#cmbCategorias").prop("disabled", false);
    $("#cmbType").prop("disabled", false);

    panel.show();
  });

  $("#btnCancelarRecat").on("click", function () {
    $("#cmbCategorias").prop("disabled", true);
    $("#cmbType").prop("disabled", true);
    $("#dvBotoneraRecategorizar").hide();
    $("#dvDetail").html("");
    $("#dvStatus").html("");
    cargarDatos(idcode);
  });

  $("#btnRecategorizar").on("click", function (e) {
    e.preventDefault();

    var btn = $(this);
    btn.prop("disabled", true);

    try {
      if (!idcode) {
        alertify.alert(
          "Sistema de Tickets",
          "No se encontró el ID del ticket.",
        );
        btn.prop("disabled", false);
        return;
      }

      var Categoria = $("#cmbCategorias").val();
      var Type = $("#cmbType").val();

      if (!Categoria || Type === "none") {
        alertify.alert(
          "Sistema de Tickets",
          "Debe seleccionar Tipo y Categoría.",
        );
        btn.prop("disabled", false);
        return;
      }

      var categoriaSeleccionada = categorias.find((c) => c.Id == Categoria);

      if (!categoriaSeleccionada || !categoriaSeleccionada.IdTemplate) {
        alertify.alert(
          "Sistema de Tickets",
          "La categoría no tiene template asignado.",
        );
        btn.prop("disabled", false);
        return;
      }

      var oTemplateList = context.get_web().get_lists().getByTitle("Templates");
      var oTemplateItem = oTemplateList.getItemById(
        categoriaSeleccionada.IdTemplate,
      );

      context.load(oTemplateItem);

      context.executeQueryAsync(
        function () {
          var nuevoTemplateJson = oTemplateItem.get_item(
            "RutaAprobacionEscalacion",
          );

          if (!nuevoTemplateJson) {
            alertify.alert("Sistema de Tickets", "El template está vacío.");
            btn.prop("disabled", false);
            return;
          }

          var t = JSON.parse(nuevoTemplateJson);

          var status =
            t.approvalPath && t.approvalPath.length > 0
              ? t.approvalPath[0].stepName
              : "Assigned";

          var ft = t.formTemplate;

          if (ft && Array.isArray(ft)) {
            for (var j = 0; j < ft.length; j++) {
              if (!ft[j].id) continue;

              switch (ft[j].tag) {
                case "checkbox":
                  ft[j].checked = $("#" + ft[j].id).is(":checked");
                  break;

                case "textarea":
                  ft[j].value = $("#" + ft[j].id).val();
                  break;
              }
            }

            t.formTemplate = ft;
          }

          actualizarTicket(idcode, Type, Categoria, t, status, btn);
        },
        function (sender, args) {
          console.error("Error cargando template:", args.get_message());
          btn.prop("disabled", false);
        },
      );
    } catch (error) {
      console.error("Error general:", error.message);
      btn.prop("disabled", false);
    }
  });

  //REAPERTURAR TICKET
  $("#btnReaperturar").on("click", function () {
    var panel = $("#dvReaperturaPanel");

    if (panel.is(":visible")) {
      panel.slideUp();
      return;
    }

    cerrarAcciones();

    $("#txtMotivoReapertura").val("");
    panel.slideDown(200);
  });
  $("#btnCancelReapertura").on("click", function () {
    $("#dvReaperturaPanel").slideUp(200);
  });
  $("#btnConfirmReapertura").on("click", function () {
    var motivo = $("#txtMotivoReapertura").val();

    if (motivo.trim() === "") {
      alertify.error("Debe ingresar el motivo de la reapertura.");
      return;
    }

    reaperturarTicket(idcode, motivo);
  });

  //TICKET EN ATENCIÓN
  $("#btnAtencion").on("click", function () {
    cerrarAcciones();

    var web = context.get_web();
    var lista = web.get_lists().getByTitle("Tickets");
    var item = lista.getItemById(idcode);

    //Actualiza campo Atencion
    item.set_item("Atencion", true);

    //GUARDA FECHA EN QUE INICIA LA ATENCIÓN (detiene SLA)
    item.set_item("FechaAtencion", (new Date()).toISOString());

    item.update();

    context.executeQueryAsync(
      function () {
        alertify.success("El ticket ahora está en atención");

        setTimeout(function () {
          location.reload();
        }, 700);
      },

      function (sender, args) {
        console.log("Error:", args.get_message());
        alertify.error("Error al actualizar el ticket");
      },
    );
  });
}

function loadTipo() {
  var $cmbType = $("#cmbType");
  var $cmbCategorias = $("#cmbCategorias");

  $cmbCategorias.prop("disabled", true);
  $("#cmbCategorias").html(
    "<option value='' selected disabled>Seleccione una categoría</option>",
  );

  $cmbType.off("change").on("change", function () {
    var tipoSeleccionado = $(this).val();

    $(this).removeClass("is-invalid").addClass("is-valid");

    if($(this).val() === "Change") {
      $("#dvChangeValidacion").show();
    }

    if (!cargandoTicket && recategorizando) {
      $("#dvDetail").html("");
    }

    $cmbCategorias.empty();
    $cmbCategorias.append(
      '<option value="" selected disabled>Seleccione una categoría</option>',
    );

    if (!tipoSeleccionado) {
      $cmbCategorias.prop("disabled", true);
      return;
    }

    categorias.forEach(function (cat) {
      if (cat.Tipo === tipoSeleccionado) {
        $cmbCategorias.append(
          `<option value="${cat.Id}"
                    data-prioridad="${cat.Prioridad}"
                    data-template="${cat.IdTemplate}">
                    ${cat.Title}
                </option>`,
        );
      }
    });

    $cmbCategorias.prop("disabled", false);
  });

  $cmbCategorias.off("change").on("change", function () {
    var idCat = $(this).val();
    var categoriaSeleccionada = categorias.find((c) => c.Id == idCat);

    if (!categoriaSeleccionada) return;

    $("#cmbUrgency").val(categoriaSeleccionada.Prioridad);
    if (!cargandoTicket && recategorizando) {
      $("#dvDetail").html("");
    }

    if (!categoriaSeleccionada.IdTemplate) return;

    var oList = context.get_web().get_lists().getByTitle("Templates");
    var oTemplate = oList.getItemById(categoriaSeleccionada.IdTemplate);

    context.load(oTemplate);

    context.executeQueryAsync(
      function () {
        var templateJson = oTemplate.get_item("RutaAprobacionEscalacion");
        if (!templateJson) return;

        var t = JSON.parse(templateJson);
        var ft = t.formTemplate;

        if (!ft) return;

        ft.forEach(function (field) {
          var c = "";

          switch (field.tag) {
            case "h3":
            case "p":
              c = `<div class='col-md-12'>
                            <${field.tag}>${field.text}</${field.tag}>
                         </div>`;
              break;

            case "checkbox":
              c = `<div class='col-md-12'>
                            <div class='form-check'>
                                <input type='checkbox' 
                                       class='form-check-input'
                                       id='${field.id}'/>
                                <label class='form-check-label'
                                       for='${field.id}'>
                                       ${field.text}
                                </label>
                            </div>
                         </div>`;
              break;

            case "textarea":
              c = `<div class='col-md-12'>
                            <div class='form-group'>
                                <label for='${field.id}'>${field.text}</label>
                                <textarea class='form-control'
                                          id='${field.id}'
                                          ></textarea>
                            </div>
                         </div>`;
              break;

            case "br":
              c = "<br/><br/>";
              break;
          }

          $("#dvDetail").append(c);
        });
      },
      function (sender, args) {
        console.error("Error cargando template: " + args.get_message());
      },
    );
  });
}

function loadCategorias(callback) {
  console.log("ejecutando load categorías");

  categorias = [];

  var camlQueryXml = `<View>
                            <Query>
                                <OrderBy>
                                    <FieldRef Name="Title" Ascending="TRUE" />
                                </OrderBy>
                            </Query>
                            <RowLimit>300</RowLimit>
                        </View>`;

  var oList = context.get_web().get_lists().getByTitle("Categorias");
  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml(camlQueryXml);

  var collItems = oList.getItems(camlQuery);
  context.load(collItems);

  context.executeQueryAsync(
    function () {
      var enumerator = collItems.getEnumerator();

      while (enumerator.moveNext()) {
        var item = enumerator.get_current();

        if (item.get_item("Title") && item.get_item("CategoriaPadre") != null) {
          categorias.push({
            Title:
              item.get_item("CategoriaPadre").get_lookupValue() +
              " - " +
              item.get_item("Title"),
            Id: item.get_id(),
            Tipo: item.get_item("TipoCategoria"),
            Prioridad: item.get_item("Prioridad"),
            IdTemplate:
              item.get_item("TemplateAtencion") != null
                ? item.get_item("TemplateAtencion").get_lookupId()
                : 0,
          });
        }
      }

      categorias.sort((a, b) => a.Title.localeCompare(b.Title));

      if (callback) callback();
    },
    function (sender, args) {
      console.error("Error al cargar Categorias: " + args.get_message());
    },
  );
}

function getFileBuffer(file) {
  var deferred = $.Deferred();
  var reader = new FileReader();

  if (file instanceof Blob || file instanceof File) {
    reader.onload = function (e) {
      deferred.resolve(e.target.result);
    };
    reader.onerror = function (error) {
      deferred.reject(error);
    };
    reader.readAsArrayBuffer(file);
  } else {
    deferred.reject("El archivo no es válido.");
  }

  return deferred.promise();
}
function subirArchivos(
  archivos,
  indice,
  spFiles,
  serverRelativeUrl,
  nombreLista,
  itemId,
  callback,
) {
  if (archivos.length > indice) {
    var archivo = archivos[indice];
    var fileName = archivo.name;

    getFileBuffer(archivo).done(function (arrayBuffer) {
      var endpoint =
        serverRelativeUrl +
        "/_api/web/lists/GetByTitle('" +
        nombreLista +
        "')/items(" +
        itemId +
        ")/AttachmentFiles/add(FileName='" +
        fileName +
        "')";

      $.ajax({
        url: endpoint,
        type: "POST",
        data: arrayBuffer,
        processData: false,
        headers: {
          accept: "application/json;odata=verbose",
          "X-RequestDigest": $("#__REQUESTDIGEST").val(),
          "content-length": arrayBuffer.byteLength,
          "content-type": "application/octet-stream",
        },
        success: function (file) {
          spFiles.push(file);
          subirArchivos(
            archivos,
            indice + 1,
            spFiles,
            serverRelativeUrl,
            nombreLista,
            itemId,
            callback,
          );
        },
        error: function (xhr) {
          console.error("Error al subir:", fileName, xhr.responseText);
          subirArchivos(
            archivos,
            indice + 1,
            spFiles,
            serverRelativeUrl,
            nombreLista,
            itemId,
            callback,
          );
        },
      });
    });
  } else {
    callback(spFiles);
  }
}

function loadAprobaciones(callback) {
  var oList = context.get_web().get_lists().getByTitle("Aprobaciones");
  var camlQuery = new SP.CamlQuery();

  camlQuery.set_viewXml(`<View><Query>
    <Where>
      <Eq>
        <FieldRef Name="Ticket" LookupId="TRUE" />
        <Value Type="Lookup">${idcode}</Value>
      </Eq>
    </Where>
    <OrderBy>
      <FieldRef Name="ID" Ascending="TRUE" />
    </OrderBy>
  </Query></View>`);

  var collListItem = oList.getItems(camlQuery);
  context.load(collListItem);

  context.executeQueryAsync(function () {
    var listItemEnumerator = collListItem.getEnumerator();

    if (status == item.get_item("Title"));
    {
    }
  }, onRequestFail);
}
function actualizarTicket(idcode, Type, Categoria, t, status, btn) {
  var context = SP.ClientContext.get_current();
  var oList = context.get_web().get_lists().getByTitle("Tickets");
  var oListItem = oList.getItemById(idcode);

  context.load(oListItem);

  context.executeQueryAsync(
    function () {
      // Determina ProcessManager según planta
      var processManagerId =
        $("#cmbPlanta").val() === "Puebla"
          ? parseInt(t.processManager.id)
          : parseInt(t.processManagerAguascalientes.id);

      var managerId = parseInt(idManArea);

      oListItem.set_item("TipoTicket", Type);
      oListItem.set_item("Categoria", Categoria);
      oListItem.set_item("TemplateConfiguracion", JSON.stringify(t));
      oListItem.set_item("Status", status);

      // Asignar ProcessManager (campo Persona)
      var pmUser = new SP.FieldUserValue();
      pmUser.set_lookupId(processManagerId);
      oListItem.set_item("ProcessManager", pmUser);

      if (managerId) {
        var managerUser = new SP.FieldUserValue();
        managerUser.set_lookupId(managerId);
        oListItem.set_item("Manager", managerUser);
      }

      oListItem.update();
      context.load(oListItem);

      context.executeQueryAsync(
        function () {
          eliminarAprobacionesDelTicket(
            idcode,
            function () {
              crearAprobacion(status, t, idcode, processManagerId, btn);
            },
            btn,
          );
        },
        function (sender, args) {
          console.error("Error al guardar ticket:", args.get_message());
          btn.prop("disabled", false);
        },
      );
    },
    function (sender, args) {
      console.error("Error al cargar ticket:", args.get_message());
      btn.prop("disabled", false);
    },
  );
}

function eliminarAprobacionesDelTicket(idcode, callback, btn) {
  var context = SP.ClientContext.get_current();
  var oListAp = context.get_web().get_lists().getByTitle("Aprobaciones");

  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml(
    "<View><Query>" +
      "<Where>" +
      "<Eq>" +
      "<FieldRef Name='Ticket' />" +
      "<Value Type='Number'>" +
      idcode +
      "</Value>" +
      "</Eq>" +
      "</Where>" +
      "</Query></View>",
  );

  var items = oListAp.getItems(camlQuery);
  context.load(items);

  context.executeQueryAsync(
    function () {
      var enumerator = items.getEnumerator();
      var itemsAEliminar = [];

      while (enumerator.moveNext()) {
        itemsAEliminar.push(enumerator.get_current());
      }

      for (var i = 0; i < itemsAEliminar.length; i++) {
        itemsAEliminar[i].deleteObject();
      }

      context.executeQueryAsync(
        function () {
          if (callback) callback();
        },
        function (sender, args) {
          console.error("Error eliminando aprobaciones:", args.get_message());
          btn.prop("disabled", false);
        },
      );
    },
    function (sender, args) {
      console.error("Error consultando aprobaciones:", args.get_message());
      btn.prop("disabled", false);
    },
  );
}

function crearAprobacion(status, t, idcode, processManagerId, btn) {
  var ctx = SP.ClientContext.get_current();
  var web = ctx.get_web();

  var ap = t.approvalPath || [];

  var oListAp = web.get_lists().getByTitle("Aprobaciones");
  var itemCreateInfoAp = new SP.ListItemCreationInformation();
  var oListItemAp = oListAp.addItem(itemCreateInfoAp);

  var nextStep = null;

  // Si hay flujo de aprobación
  if (ap.length > 0) {
    var currentIndex = ap.findIndex((x) => x.stepName === status);

    if (currentIndex !== -1 && currentIndex + 1 < ap.length) {
      nextStep = ap[currentIndex + 1];
    } else {
      nextStep = ap[0];
    }
  }

  // SI NO HAY FLUJO → Assigned directo
  if (!nextStep) {
    oListItemAp.set_item("Title", "Assigned");
    oListItemAp.set_item("Resultado", "Pendiente");
    oListItemAp.set_item("Ticket", idcode);

    var userFieldPM = new SP.FieldUserValue();
    userFieldPM.set_lookupId(processManagerId);

    oListItemAp.set_item("Responsable", userFieldPM);
  } else {
    oListItemAp.set_item("Title", nextStep.stepName);
    oListItemAp.set_item("Resultado", "Pendiente");
    oListItemAp.set_item("Ticket", idcode);

    if (nextStep.approvalGroup != null) {
      oListItemAp.set_item("Grupo", nextStep.approvalGroup);
    } else {
      var responsableId = nextStep.useUserDepartment
        ? parseInt(idManArea)
        : processManagerId;

      if (!responsableId) {
        console.error("Responsable inválido");
        btn.prop("disabled", false);
        return;
      }

      var userField = new SP.FieldUserValue();
      userField.set_lookupId(responsableId);

      oListItemAp.set_item("Responsable", userField);
    }
  }

  oListItemAp.update();
  ctx.load(oListItemAp);

  ctx.executeQueryAsync(
    function () {
      finalizarProceso(btn);
    },
    function (sender, args) {
      console.error("Error al crear aprobación:", args.get_message());
      btn.prop("disabled", false);
    },
  );
}

function finalizarProceso(btn) {
  btn.prop("disabled", false);

  alertify.alert(
    "Sistema de Ticket",
    "El ticket se ha cerrado correctamente.",
    function () {
      location.reload();
    },
  );
}

function cerrarAcciones() {
  $("#dvReassignPanel").slideUp();
  $("#dvReaperturaPanel").slideUp();
  $("#dvBotoneraRecategorizar").hide();

  $("#cmbCategorias").prop("disabled", true);
  $("#cmbType").prop("disabled", true);
}

function initializePeoplePicker(id) {
  var schema = {};

  schema["PrincipalAccountType"] = "User";
  schema["SearchPrincipalSource"] = 15;
  schema["ResolvePrincipalSource"] = 15;
  schema["AllowMultipleValues"] = false;
  schema["MaximumEntitySuggestions"] = 50;
  schema["Width"] = "100%";

  SPClientPeoplePicker_InitStandaloneControlWrapper(id, null, schema);
}

function actualizarResponsable(userValue) {
  var aprobacionesList = context
    .get_web()
    .get_lists()
    .getByTitle("Aprobaciones");
  var camlQuery = new SP.CamlQuery();

  camlQuery.set_viewXml(`
      <View>
        <Query>
          <Where>
            <And>
              <Eq>
                <FieldRef Name="Ticket" LookupId="TRUE"/>
                <Value Type="Lookup">${idcode}</Value>
              </Eq>
              <Eq>
                <FieldRef Name="Title"/>
                <Value Type="Text">${status}</Value>
              </Eq>
            </And>
          </Where>
        </Query>
      </View>
    `);

  var items = aprobacionesList.getItems(camlQuery);
  context.load(items);

  context.executeQueryAsync(function () {
    var enumerator = items.getEnumerator();

    if (!enumerator.moveNext()) {
      alertify.alert(
        "Sistema de Ticket",
        "No se encontró ningún registro para reasignar.",
      );
      return;
    }

    var aprobacionItem = enumerator.get_current();

    // actualiza el responsable en aprobaciones
    aprobacionItem.set_item("Responsable", userValue);
    aprobacionItem.update();

    var ticketsList = context.get_web().get_lists().getByTitle("Tickets");
    var ticketItem = ticketsList.getItemById(idcode);

    context.load(ticketItem);

    context.executeQueryAsync(function () {
      var titleAprobacion = aprobacionItem.get_item("Title");

      if (titleAprobacion === "Revisión de Manager") {
        ticketItem.set_item("Manager", userValue);
      } else if (titleAprobacion === "Assigned") {
        ticketItem.set_item("ProcessManager", userValue);
      }

      ticketItem.update();

      context.executeQueryAsync(function () {
        alertify.alert(
          "Sistema de Ticket",
          "Responsable reasignado correctamente.",
          function () {
            location.reload();
          },
        );
      }, onRequestFail);
    }, onRequestFail);
  }, onRequestFail);
}

function reaperturarTicket(idcode, motivo) {
  var web = context.get_web();

  // Usuario actual
  var cuser = web.get_currentUser();
  context.load(cuser);

  context.executeQueryAsync(function () {
    // ===== ACTUALIZA TICKET =====
    var ticketsList = web.get_lists().getByTitle("Tickets");
    var ticketItem = ticketsList.getItemById(idcode);

    ticketItem.set_item("Status", "Assigned");
    ticketItem.set_item("Reaperturado", true);
    ticketItem.set_item("MotivoReapertura", motivo);
    ticketItem.set_item("FechaReapertura", new Date());
    ticketItem.set_item("ReaperturadoPor", cuser.get_id());

    ticketItem.update();

    // ===== BUSCAR APROBACIONES =====
    var aprobacionesList = web.get_lists().getByTitle("Aprobaciones");
    var camlQuery = new SP.CamlQuery();

    camlQuery.set_viewXml(`
            <View>
                <Query>
                    <Where>
                        <And>
                            <Eq>
                                <FieldRef Name="Ticket" LookupId="TRUE"/>
                                <Value Type="Lookup">${idcode}</Value>
                            </Eq>
                            <Eq>
                                <FieldRef Name="Title"/>
                                <Value Type="Text">Assigned</Value>
                            </Eq>
                        </And>
                    </Where>
                </Query>
            </View>
        `);

    var items = aprobacionesList.getItems(camlQuery);
    context.load(items);

    context.executeQueryAsync(function () {
      var enumerator = items.getEnumerator();

      while (enumerator.moveNext()) {
        var item = enumerator.get_current();
        item.set_item("Resultado", "Pendiente");
        item.update();
      }

      context.executeQueryAsync(function () {
        alertify.success("El ticket se reaperturó correctamente");

        setTimeout(function () {
          location.reload();
        }, 800);
      }, onRequestFail);
    }, onRequestFail);
  }, onRequestFail);
}

// function marcarAprobacionesRecategorizadas(idcode, callback, btn) {

//     var oListAp = context.get_web().get_lists().getByTitle("Aprobaciones");

//     var camlQuery = new SP.CamlQuery();
//     camlQuery.set_viewXml(
//         "<View><Query>" +
//         "<Where>" +
//         "<Eq>" +
//         "<FieldRef Name='Ticket' />" +
//         "<Value Type='Number'>" + idcode + "</Value>" +
//         "</Eq>" +
//         "</Where>" +
//         "</Query></View>"
//     );

//     var items = oListAp.getItems(camlQuery);
//     context.load(items);

//     context.executeQueryAsync(function () {

//         var enumerator = items.getEnumerator();

//         while (enumerator.moveNext()) {
//             var item = enumerator.get_current();
//             item.set_item("Resultado", "Recategorizado");
//             item.update();
//         }

//         context.executeQueryAsync(function () {

//             if (callback) callback();

//         }, function (sender, args) {
//             console.error("Error actualizando aprobaciones:", args.get_message());
//             btn.prop("disabled", false);
//         });

//     }, function (sender, args) {
//         console.error("Error consultando aprobaciones:", args.get_message());
//         btn.prop("disabled", false);
//     });
// }
