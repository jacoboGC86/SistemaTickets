var context, contextCol, cuser, web, idcode, idCanSol = 0, subiendo = false, tipoDocCargar, departamento = 0, doc = 0, status = "", process = [], category = 0, tipoSol ;
var idTask = 0, firmado = false, iamENG = false, drawing = false, mousePos = { x:0, y:0 }, lastPos, haySolCan = false;
var posicionValorGlobal = "";

$(document).ready(function(){
  $("body").on("contextmenu", "img", function(e) {
    return false;
  });
  var iebrowser = /*@cc_on!@*/false || !!document.documentMode;
  
  if(iebrowser) {
      alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
      window.location.href = "https://chimalliapps.sharepoint.com/sites/Contacto/LL";
  }

  context = new SP.ClientContext.get_current();
  contextCol = new SP.ClientContext("https://chimalliapps.sharepoint.com/sites/Contacto/Vacations");
  web = context.get_web();
  cuser = web.get_currentUser();
  
  context.load(web);
  context.load(cuser);  
  context.executeQueryAsync(function(sender, args) {
    var params = getParameters(window.location.href);

    for(var i = 0; i < params.length; i++) {
      if(params[i].param === "idSol") {
        idcode = params[i].value;
        break;
      }
    }
      
    /*
    Promise.all([
        loadProcess(null, null),
        
        cargarDepartamentos()
    ]).then(() => {
        cargarDatos(idcode);
    }).catch(error => {
        console.error("Error en la carga de funciones:", error);
    });
    */

    cargarDepartamentos(function() {
      loadProcess(function(){
        cargarDatos(idcode);
      });
    });
  }, onRequestFail);
});

function cargarDatos(idcode) {   
  var oList = context.get_web().get_lists().getByTitle("Request");
  var oListItem = oList.getItemById(idcode);

  context.load(oListItem);
  context.executeQueryAsync(function(sender, args) {
    $("#txtFolio").val(oListItem.get_id());

    $("#txtProyecto").val(oListItem.get_item("Title"));
    $("#dtFechaRegistro").val(obtenerFecha(oListItem.get_item("Created"), "dd/mm/yyyy"));
  
    $("#txtDescripcion").val(oListItem.get_item("Descripcion"));
    const textareaD = document.getElementById('txtDescripcion');
    textareaD.style.height = 'auto';
    textareaD.style.height = textareaD.scrollHeight + 'px'; 
    textareaD.addEventListener('input', function() {
        this.style.height = 'auto';  
        this.style.height = this.scrollHeight + 'px';  
    });
    $("#txtExplicacion").val(oListItem.get_item("ExplicacionSitucacion"));
    const textareaE = document.getElementById('txtExplicacion');
    textareaE.style.height = 'auto';  
    textareaE.style.height = textareaE.scrollHeight + 'px'; 
    textareaE.addEventListener('input', function() {
        this.style.height = 'auto';  
        this.style.height = this.scrollHeight + 'px';  9
    });

    $("#cmbPlanta").val(oListItem.get_item("Planta"));
    $("#cmbLeccionAprendida").val(oListItem.get_item("LeccionAprendida"));
        
    var responsableOne = oListItem.get_item("ResponsableEjecuto");
    var responsableTwo = oListItem.get_item("ResponsableEjecutara");
    $('#txtResponsableOneName').val(responsableOne);
    $('#txtResponsableTwoName').val(responsableTwo);

      //procesa departamento
     var departmentField = oListItem.get_item("Departamento");
     if (departmentField) {
        console.log("depto", departmentField);
         var foundDepto = false; 
         $("#cmbDepartamento option").each(function() {
             var optionValue = $(this).data("valor");
             console.log("Opción en el combo:", optionValue);
             if (optionValue === departmentField) {
                 $(this).prop("selected", true);
                 foundDepto = true;
                 return false;
             }
         });

         if (!foundDepto) {
             console.warn("No se encontró un valor coincidente en el combo.");
         }
     } else {
         console.error("El campo 'Departamento' no tiene un valor asignado.");
     }
     
    // Proceso
    var procesoField = oListItem.get_item("ProcesoAplicable");
    var procesoName = procesoField.get_lookupValue();
    if (procesoName) {
      
        var foundProceso = false; 
        $("#cmbProceso option").each(function() {
            var optionValue = $(this).text(); 
            if (optionValue === procesoName) {
                $(this).prop("selected", true);
                foundProceso = true;
                return false;
            }
        });

        if (!foundProceso) {
            console.warn("No se encontró un valor coincidente en el combo.");
        }
    } else {
        console.error("El campo ProcesoAplicable no tiene un valor asignado.");
    }

    //lista de procesos aplicables
    let procesosAplicables = oListItem.get_item("ProcesosAplicables");
      if (procesosAplicables && Array.isArray(procesosAplicables)) {
          let $ol = $("#ListProcesosAplicables");  
          $ol.empty();  
          
          procesosAplicables.forEach(function(report) {
              let lookupValue = report.get_lookupValue ? report.get_lookupValue() : report.Title;
              
              let $li = $("<li></li>")
                  .addClass("list-group-item d-flex align-items-center")
                  .css({
                      "padding": "2px 0",
                      "border": "none"
                  });

              let $icon = $("<i></i>")
                  .addClass("fas fa-arrow-right")  
                  .css({
                      "margin-right": "10px",
                      "color": "#28a745" 
                  });

              let $text = $("<span></span>").text(lookupValue.trim());
              
              $li.append($icon).append($text);
              $ol.append($li);
          });
      } else {
          $("#ListProcesosAplicables").empty(); 
      }


      var attachmentFolder = web.getFolderByServerRelativeUrl('Lists/Request/Attachments/' + idcode);
      var attachmentFiles = attachmentFolder.get_files();
  
      //Load attachments
      context.load(attachmentFiles);
      context.executeQueryAsync(function(sender, args) {
        var cnt = attachmentFiles.get_count();
        if (attachmentFiles.get_count() === 0) {
          console.warn('No hay archivos adjuntos para mostrar.');
          return;
      }
      
        for(var itr=0; itr<cnt; itr++){
          $("#ulAdjuntos").append("<li class='list-group-item'><a href='" + attachmentFiles.itemAt(itr).get_serverRelativeUrl() + "'>" + attachmentFiles.itemAt(itr).get_name() + "</a></li>");
          var extension = attachmentFiles.itemAt(itr).get_name().split(".")[attachmentFiles.itemAt(itr).get_name().split(".").length - 1];
          
          if(extension === "jpg" || extension === "png" || extension === "jpeg") {
            var plantilla = "<div class='col-md-4'><img src='" + attachmentFiles.itemAt(itr).get_serverRelativeUrl() + "' class='img-thumbnail' /></div>";
  
            $("#dvRepresentacion").append(plantilla);
          }
        }
        asignarEventos();  
      }, function(sender, args) {
          console.error('Error en la consulta o carga: ' + args.get_message());
          asignarEventos();  
      });
// Status switch
status = oListItem.get_item("Status");
console.log("Status value antes del switch:", status);

switch(status) {
  
  case "Solicitud aprobada":
      $("#dvA1").addClass("aprobadoRev");
      $("#dvBotonera").show();
      break;

  default:
      console.log("Status no coincide con ningún caso en el switch");
      break;
}


  if (oListItem.get_item("FirmaEngeeniering") != null) {
    $("#dvENGFirma").prop("src", oListItem.get_item("FirmaEngeeniering"));
  }

   (oListItem.get_item("AprobacionEngeeniering") != null) 
    $("#dvA1").html("En Revisión de Engineering <br/>" + oListItem.get_item("AprobacionEngeeniering").get_lookupValue() + "<br/>" + obtenerFecha(oListItem.get_item("FFEngeeniering"), "dd/mm/yyyy"));
    $("#lblENGFirm").text(oListItem.get_item("AprobacionEngeeniering").get_lookupValue());
  
   
 
    
  }, onRequestFail);
}

function loadProcess(callback) {
  var oListConf = context.get_web().get_lists().getByTitle("Process");
  var camlQueryConf = new SP.CamlQuery();
  camlQueryConf.set_viewXml("<View><Query><Where></Where></Query></View>");
  var collListItemConf = oListConf.getItems(camlQueryConf);
  
  context.load(collListItemConf);
  context.executeQueryAsync(function(sender, args) {
      var listItemEnumerator = collListItemConf.getEnumerator();
      var $select = $("#cmbProceso");
      $select.empty();

      $select.append('<option value="none" selected disabled>Seleccione un proceso</option>');
      var itemCount = 0;

      while (listItemEnumerator.moveNext()) {
          var oListItem = listItemEnumerator.get_current();
          var title = oListItem.get_item("Title");

          if (title) {
              var $option = $("<option></option>")
                  .val(oListItem.get_id())
                  .text(title);

              $select.append($option);
              itemCount++;
          }
      }

      if (itemCount === 0) { 
          console.log("No se encontraron elementos en proceso");
      } else {
       console.log(itemCount + " elementos agregados al combo de proceso.");
      }

      $("#dvLoading").hide();
      $select.show(); 
      
      callback();
  }, onRequestFail);
}

function cargarDepartamentos(callback) {
  var oListConf = contextCol.get_web().get_lists().getByTitle("Configuraciones"); 
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
  contextCol.executeQueryAsync(function(sender, args) {
    //Success
      var listItemEnumerator = collListItemConf.getEnumerator();

      while (listItemEnumerator.moveNext()) {
          var oListItem = listItemEnumerator.get_current();
          var CampoPersona = oListItem.get_item("Personas");

          if (CampoPersona && CampoPersona.length > 0) {
              CampoPersona.forEach(function(personaItem) {
                  var personaId = personaItem.get_lookupId();  
                  var personaName = personaItem.get_lookupValue();

                  $("#cmbDepartamento").append(
                      "<option value='" + oListItem.get_id() + "' data-id='" + personaId + "' data-valor='" + oListItem.get_item("Valor") + "'>" + oListItem.get_item("Valor") + " - " + personaName + "</option>"
                  );
              });
          } else {
              console.warn("El campo 'Personas' no tiene datos.");
          }
      }

      callback();
  }, 
  function(sender, args) {
    //Error
    console.error('Error al obtener los elementos: ' + args.get_message());
  });
}

function asignarEventos() {

  $("#btnCerrar").click(function () {
    window.location.href = "Solicitudes aprobadas.aspx";
  });

  
 
}
