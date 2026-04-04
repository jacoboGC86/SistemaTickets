var context, contextCol, cuser, web, idcode, idCanSol = 0, subiendo = false, tipoDocCargar, departamento = 0, doc = 0, status = "", process = [], category = 0, tipoSol ;
var idTask = 0, firmado = false,  iamAreaMan = false, iamRH = false, drawing = false, mousePos = { x:0, y:0 }, lastPos, haySolCan = false;//canvasx, canvasy, last_mousex, mousex, mousedown;
var posicionValorGlobal = "";



$(document).ready(function(){
  
  $("body").on("contextmenu", "img", function(e) {
    return false;
  });
  var iebrowser = /*@cc_on!@*/false || !!document.documentMode;
  
  if(iebrowser) {
      alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
      window.location.href = "https://chimalliapps.sharepoint.com/sites/Contacto/JD";
  }


  context = new SP.ClientContext.get_current();
  contextCol = new SP.ClientContext("https://chimalliapps.sharepoint.com/sites/Contacto/JD");
  web = context.get_web();
  cuser = web.get_currentUser();
  

  context.load(web);
  context.load(cuser);  
  context.executeQueryAsync(function(sender, args) {
    var params = getParameters(window.location.href);

    for(var i = 0; i < params.length; i++)
        if(params[i].param === "idSol") {
            idcode = params[i].value;
            break;
        }
    
    var oList = context.get_web().get_lists().getByTitle("Positions");
    var oListItem = oList.getItemById(idcode);

    context.load(oListItem);
    context.executeQueryAsync(function(sender, args) {
      $("#txtFolio").val(oListItem.get_id());

      $("#txtPosicion").val(oListItem.get_item("Title"));
           
    
        $("#dvRHFirma").prop("src", oListItem.get_item("FirmaRH"));
        $("#lblRHFirm").text(oListItem.get_item("AprobacionRH").get_lookupValue());
        var DateAprobacionRH = oListItem.get_item("FAprobacionRH");

        if (DateAprobacionRH) {
            var dateObj = new Date(DateAprobacionRH);
            var day = String(dateObj.getDate()).padStart(2, '0');
            var month = String(dateObj.getMonth() + 1).padStart(2, '0'); 
            var year = dateObj.getFullYear();
            var formattedDate = `${day}/${month}/${year}`;
            $('#lblFAprobacionRH').text(formattedDate);
        } else {
            console.log("No se pudo obtener la fecha de FAprobacionRH.");
        }

     
     
        $("#dvFirmAM").prop("src", oListItem.get_item("FirmaManager"));
        $("#lblFirmAM").text(oListItem.get_item("Manager").get_lookupValue());
        var DateFFManager = oListItem.get_item("FFManager");

        if (DateFFManager) {
            var dateObj = new Date(DateFFManager);
            var day = String(dateObj.getDate()).padStart(2, '0');
            var month = String(dateObj.getMonth() + 1).padStart(2, '0'); 
            var year = dateObj.getFullYear();
            var formattedDate = `${day}/${month}/${year}`;
            $('#lblFFManager').text(formattedDate);
        } else {
            console.log("No se pudo obtener la fecha de FFManager.");
        }

     
      
        
        


      posicionValorGlobal = $("#txtPosicion").val().trim();
    /*  console.log("Valor de la posición:", posicionValorGlobal);*/
      executeSegundaQuery(posicionValorGlobal);
      executeTerceraQuery();

      $("#txtDepto").val(oListItem.get_item("Department"));
      $("#txtGerente").val(oListItem.get_item("Reportee").get_lookupValue());
 

            let directReports = oListItem.get_item("DirectReports");
            if (directReports && Array.isArray(directReports)) {
                let $ol = $("#txtDirector");  
                $ol.empty();  
                
                directReports.forEach(function(report) {
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
                $("#txtDirector").empty(); 
            }

      
            let internalContacts = oListItem.get_item("InternalContacts");
            if (Array.isArray(internalContacts)) {
                let $ol = $("#ListContactosInt");
                $ol.empty(); 
            
                internalContacts.forEach(function(item) {
                    let $li = $("<li></li>")
                        .addClass("list-group-item d-flex align-items-center") 
                        .css({"padding": "2px 0",
                             "border":"none"});  
            
                   let $icon = $("<i></i>")
                       .addClass("fas fa-arrow-right")  
                        .css({
                            "margin-right": "10px",
                            "color": "#28a745" 
                        });
                    let $text = $("<span></span>").text(item.trim());
                    $li.append($icon).append($text);
                    $ol.append($li);
                });
            } else {
                $("#ListContactosInt").empty();  
            }

            let externalContacts = oListItem.get_item("ExternalContacts");
            if (Array.isArray(externalContacts)) {
                let $ol = $("#ListContactosExt");
                $ol.empty(); 
                externalContacts.forEach(function(item) {
                    let $li = $("<li></li>")
                        .addClass("list-group-item d-flex align-items-center") 
                        .css({"padding": "2px 0",
                            "border":"none"});  
                   let $icon = $("<i></i>")
                        .addClass("fas fa-arrow-right")  
                        .css({
                            "margin-right": "10px",
                            "color": "#28a745" 
                        });
                    let $text = $("<span></span>").text(item.trim());
                    $li.append($icon).append($text);
                    $ol.append($li);
                });
            } else {
                $("#ListContactosExt").empty();  
            }
 


        $("#txtProposito").val(oListItem.get_item("Purpose"));
        const textarea = document.getElementById('txtProposito');
        textarea.style.height = 'auto';  
        textarea.style.height = textarea.scrollHeight + 'px'; 
        textarea.addEventListener('input', function() {
            this.style.height = 'auto';  
            this.style.height = this.scrollHeight + 'px';  
        });

      $("#cmbIngles").val(oListItem.get_item("EnglishLevel"));
      $("#txtIdioma").val(oListItem.get_item("OhterLanguage"));
      $("#cmbNivel").val(oListItem.get_item("OtherLangaugeLevel"));
      $("#cmbViajar").val(oListItem.get_item("TravelRequired"));

      let competenciasTecnicas = oListItem.get_item("CompetenciasTecnicas");
      let jsonArrayT = JSON.parse(competenciasTecnicas);
      let $ulT = $("<ul></ul>")
          .addClass("list-group")
          .css({
              "border": "none",
              "padding-left": "0",
              "margin": "0"
          });
      jsonArrayT.forEach(function(item) {
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
          let $text = $("<span></span>").text(item.trim());
          $li.append($icon).append($text);
          $ulT.append($li);
      });
      $("#ListComTecnicas").html($ulT);
      
      
    
      let competenciasPersonales = oListItem.get_item("CompetenciasPersonales");
        let jsonArrayP = JSON.parse(competenciasPersonales);
        let $ulP = $("<ul></ul>")
            .addClass("list-group")
            .css({
                "border": "none",
                "padding-left": "0",
                "margin": "0"
            });
        jsonArrayP.forEach(function(item) {
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
            let $text = $("<span></span>").text(item.trim());
            $li.append($icon).append($text);
            $ulP.append($li);
        });
        $("#txtComPersonales").html($ulP);

      let competenciasBasicas = oListItem.get_item("CompetenciasBasicas");
      let jsonArrayB = JSON.parse(competenciasBasicas);
      let $ulB = $("<ul></ul>")
          .addClass("list-group")
          .css({
              "border": "none",
              "padding-left": "0",
              "margin": "0"
          });
      jsonArrayB.forEach(function(item) {
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
          let $text = $("<span></span>").text(item.trim());
          $li.append($icon).append($text);
          $ulB.append($li);
      });
      
      $("#txtComBasicas").html($ulB);
      

      let JobRequirements = oListItem.get_item("JobRequirements");
      let jsonArrayR = JSON.parse(JobRequirements);
      let $ulR = $("<ul></ul>")
          .addClass("list-group")
          .css({
              "border": "none",
              "padding-left": "0",
              "margin": "0"
          });
      jsonArrayR.forEach(function(item) {
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

          let $text = $("<span></span>").text(item.trim());
          $li.append($icon).append($text);
          $ulR.append($li);
      });
      $("#txtRequerimientos").html($ulR);




      let Equipament = oListItem.get_item("MinimumPersonalEquipament");
      if (Array.isArray(Equipament)) {
          let $ol = $("#txtEquipament");
          $ol.empty(); 
      
          Equipament.forEach(function(item) {
              let $li = $("<li></li>")
                  .addClass("list-group-item d-flex align-items-center") 
                  .css({"padding": "2px 0",
                      "border":"none"});  
      
             let $icon = $("<i></i>")
                  .addClass("fas fa-arrow-right")  
                  .css({
                      "margin-right": "10px",
                      "color": "#28a745" 
                  });
              let $text = $("<span></span>").text(item.trim());
              $li.append($icon).append($text);
              $ol.append($li);
          });
      } else {
          $("#txtEquipament").empty();  
      }

  
      var attachmentFolder = web.getFolderByServerRelativeUrl('Lists/Request/Attachments/' + idcode);
      var attachmentFiles = attachmentFolder.get_files(); 

   
      context.load(attachmentFiles);
      context.executeQueryAsync(function(sender, args) {
       /* var cnt = attachmentFiles.get_count();
    
        for (var itr = 0; itr < cnt; itr++) {
            $("#ulAdjuntos").append("<li class='list-group-item'><a href='" + attachmentFiles.itemAt(itr).get_serverRelativeUrl() + "'>" + attachmentFiles.itemAt(itr).get_name() + "</a></li>");
        }
    */
        asignarEventos(oListItem);
    }, function(sender, args) {
        console.error('Error en la consulta o carga: ' + args.get_message());
        asignarEventos(oListItem); 
    });
    
    }, onRequestFail);

     /*Pasa el parametro recuperado de la primera funcion*/
 function executeSegundaQuery(posicionValor) {
    contextCol.executeQueryAsync(function(sender, args) {
        var oList = contextCol.get_web().get_lists().getByTitle("Job Description");
        var camlQuery = new SP.CamlQuery();
        var camlXml = "<View><Query><Where><Eq><FieldRef Name='Position'/><Value Type='Text'>" + posicionValor + "</Value></Eq></Where><OrderBy><FieldRef Name='ID' Ascending='FALSE'/></OrderBy></Query><RowLimit>200</RowLimit></View>";
        camlQuery.set_viewXml(camlXml);
        var collListItem = oList.getItems(camlQuery);

        contextCol.load(collListItem);
        contextCol.executeQueryAsync(function(sender, args) {
            var listItemEnumerator = collListItem.getEnumerator();
            var rows = "";
            var activityTypeMap = {};//Inicializa el obj para Agrupar los datos acorde al  tipo de actividad
        //Recorre la lista y recupera los campos
            while (listItemEnumerator.moveNext()) {
                var oListItem = listItemEnumerator.get_current();
                var activityType = oListItem.get_item("ActivityType");
                var title = oListItem.get_item("Title");
                var frequency = oListItem.get_item("ActivityFrequency");
             
            
                //Inicializa los arreglos 
                if (!activityTypeMap[activityType]) {
                    activityTypeMap[activityType] = {
                        titles: [],
                        frequencies: [],
                       
                    };
                }
            
                activityTypeMap[activityType].titles.push(title);
                activityTypeMap[activityType].frequencies.push(frequency);
            }
            
          
            for (var type in activityTypeMap) {
                var titles = activityTypeMap[type].titles.join(", ");
                var frequencies = activityTypeMap[type].frequencies.join(", ");
             
                var numRows = activityTypeMap[type].titles.length;
            
                // Agregar la primera fila con rowspan, crea una fila para cada tipo encontrado y agrega filas adicionales para titilo y frecuencia.
                rows += "<tr>" +
                            "<td rowspan='" + numRows + "'>" + type + "</td>" +
                            "<td>" + activityTypeMap[type].titles[0] + "</td>" +
                            "<td>" + activityTypeMap[type].frequencies[0] + "</td>" +
                        "</tr>";
            
                for (var i = 1; i < numRows; i++) {
                    rows += "<tr>" +
                                "<td>" + activityTypeMap[type].titles[i] + "</td>" +
                                "<td>" + activityTypeMap[type].frequencies[i] + "</td>" +
                            "</tr>";
                }
            }
            
            if (rows === "") {
                console.log("No se encontraron elementos en la lista");
            }
            
            $("#dvLoading").hide();
            $("#tbCodigos_WP").show();
            $("#tbCodigos_WP tbody").html(rows);
            
            $('#tbCodigos_WP').DataTable({
                "order": [[ 0, "desc" ]]
            });
            
        }, onRequestFail);

    }, onRequestFail);
}

function executeTerceraQuery() {
    contextCol.executeQueryAsync(function(sender, args) {
        var oList = contextCol.get_web().get_lists().getByTitle("Generic Job Description");
        var camlQuery = new SP.CamlQuery();
        var camlXml = "<View><Query><OrderBy><FieldRef Name='ID' Ascending='FALSE'/></OrderBy></Query></View>";
        camlQuery.set_viewXml(camlXml);
        var collListItem = oList.getItems(camlQuery);

        contextCol.load(collListItem);
        contextCol.executeQueryAsync(function(sender, args) {
            var listItemEnumerator = collListItem.getEnumerator();
            var rows = "";
            var activityTypeMap = {};
        //Recorre la lista y recupera los campos
            while (listItemEnumerator.moveNext()) {
                var oListItem = listItemEnumerator.get_current();
                var activityType = oListItem.get_item("ActivityType");
                var title = oListItem.get_item("Title");
                var frequency = oListItem.get_item("ActivityFrequency");
             
            
                //Inicializa los arreglos 
                if (!activityTypeMap[activityType]) {
                    activityTypeMap[activityType] = {
                        titles: [],
                        frequencies: [],
                       
                    };
                }
            
                activityTypeMap[activityType].titles.push(title);
                activityTypeMap[activityType].frequencies.push(frequency);
            }
            
          
            for (var type in activityTypeMap) {
                var titles = activityTypeMap[type].titles.join(", ");
                var frequencies = activityTypeMap[type].frequencies.join(", ");
             
                var numRows = activityTypeMap[type].titles.length;
            
                // Agregar la primera fila con rowspan, crea una fila para cada tipo encontrado y agrega filas adicionales para titilo y frecuencia.
                rows += "<tr>" +
                            "<td rowspan='" + numRows + "'>" + type + "</td>" +
                            "<td>" + activityTypeMap[type].titles[0] + "</td>" +
                            "<td>" + activityTypeMap[type].frequencies[0] + "</td>" +
                        "</tr>";
            
                for (var i = 1; i < numRows; i++) {
                    rows += "<tr>" +
                                "<td>" + activityTypeMap[type].titles[i] + "</td>" +
                                "<td>" + activityTypeMap[type].frequencies[i] + "</td>" +
                            "</tr>";
                }
            }
            
            if (rows === "") {
                console.log("No se encontraron elementos en la lista");
            }
            
            $("#dvLoadingG").hide();
            $("#tbActGenerales").show();
            $("#tbActGenerales tbody").html(rows);
            
            $('#tbActGenerales').DataTable({
                "order": [[ 0, "desc" ]]
            });
            
        }, onRequestFail);

    }, onRequestFail);
}



}, onRequestFail);

function asignarEventos(oListItem) {
//botones Editar y borrar CNV
    $("#btnEdit").click(function() {
        var IdEdit = oListItem.get_id();
        localStorage.setItem("IdEdit", IdEdit); 
        window.location.href = "Request Job Description Edit.aspx"; 
    
    });

    $("#btnBorrarCLB").click(function () {
        firmado = false;
        canvas.width = canvas.width; 
      });
      $("#btnImprimir").click(function () {
        firmado = false;
        window.print();
      });
  cnv = "sheetCLB";
  canvas = document.getElementById(cnv);
  contextCnv = canvas.getContext("2d");
  
  contextCnv.lineWidth = 3;
  
  canvas.addEventListener("mousedown", function (e) {
    drawing = true;
    lastPos = getMousePos(canvas, e);
  }, false);
  canvas.addEventListener("mouseup", function (e) {
    drawing = false;
  }, false);
  canvas.addEventListener("mousemove", function (e) {
    mousePos = getMousePos(canvas, e);
  }, false);

  // Set up touch events for mobile, etc
  canvas.addEventListener("touchstart", function (e) {
    mousePos = getTouchPos(canvas, e);
  
    var touch = e.touches[0];
    var mouseEvent = new MouseEvent("mousedown", {
      clientX: touch.clientX,
      clientY: touch.clientY
    });
    canvas.dispatchEvent(mouseEvent);
  }, false);
  canvas.addEventListener("touchend", function (e) {
    var mouseEvent = new MouseEvent("mouseup", {});
    canvas.dispatchEvent(mouseEvent);
  }, false);
  canvas.addEventListener("touchmove", function (e) {
    e.preventDefault();
  
    var touch = e.touches[0];
    var mouseEvent = new MouseEvent("mousemove", {
      clientX: touch.clientX,
      clientY: touch.clientY
    });
    canvas.dispatchEvent(mouseEvent);
  }, false);




// Botones  (Cerrar,Enviar)
$("#btnCerrar").click(function () {
  window.location.href = "Descripciones de Puesto.aspx";
});

$("#btnEnviar").click(function () {
  if (!firmado) {
    alertify.alert("Extraordinary Freight", "Por favor, firme la solicitud.");
    return;
  }

 


  var oList = context.get_web().get_lists().getByTitle("Positions");
  var oListItem = oList.getItemById(idcode);

  context.load(oListItem);
  context.executeQueryAsync(function (sender, args) {
    if (iamAreaMan) {
      oListItem.set_item("FFManager", (new Date()).toISOString());
      oListItem.set_item("FirmaManager", canvas.toDataURL());
      oListItem.set_item("Status", "En Revisión de RH");
    } else if (iamRH) {
      oListItem.set_item("FAprobacionRH", (new Date()).toISOString());
      oListItem.set_item("FirmaRH", canvas.toDataURL());
      oListItem.set_item("Status", "Solicitud aprobada");
      oListItem.set_item("AprobacionRH", cuser.get_id());
    }

    oListItem.update();
    context.executeQueryAsync(function (sender, args) {
      alertify.alert("Extraordinary Freight", "Solicitud aprobada exitosamente.", function () {
        window.location.href = "Solicitudes en proceso de aprobacion.aspx";
      });
    }, onRequestFail);
  }, onRequestFail);
});

$("#btnCancelar").click(function () {
  if (confirm("Está por rechazar la solicitud. ¿Desea continuar?")) {
    var oList = context.get_web().get_lists().getByTitle("Positions");
    var oListItem = oList.getItemById(idcode);

    context.load(oListItem);
    context.executeQueryAsync(function (sender, args) {
      if (iamAreaMan)
        oListItem.set_item("FFManager", (new Date()).toISOString());
      else
        oListItem.set_item("FAprobacionRH", (new Date()).toISOString());

      oListItem.set_item("Status", "Solicitud rechazada");

      oListItem.update();
      context.executeQueryAsync(function (sender, args) {
        alertify.alert("Extraordinary Freight", "Solicitud rechazada.", function () {
          window.location.href = "Solicitudes en proceso de aprobacion.aspx";
        });
      }, onRequestFail);
    }, onRequestFail);
  }
});
 
}



  
});

  
  // Get the position of the mouse relative to the canvas
  function getMousePos(canvasDom, mouseEvent) {
    var rect = canvasDom.getBoundingClientRect();
  
    return {
      x: mouseEvent.clientX - rect.left,
      y: mouseEvent.clientY - rect.top
    };
  }
  
  // Get the position of a touch relative to the canvas
  function getTouchPos(canvasDom, touchEvent) {
    var rect = canvasDom.getBoundingClientRect();
    return {
      x: touchEvent.touches[0].clientX - rect.left,
      y: touchEvent.touches[0].clientY - rect.top
    };
  }
  
  // Get a regular interval for drawing to the screen
  window.requestAnimFrame = (function (callback) {
    return window.requestAnimationFrame || 
       window.webkitRequestAnimationFrame ||
       window.mozRequestAnimationFrame ||
       window.oRequestAnimationFrame ||
       window.msRequestAnimaitonFrame ||
       function (callback) {
        window.setTimeout(callback, 1000/60);
       };
  })();
  
  // Draw to the canvas
  function renderCanvas() {
    if (drawing) {
      contextCnv.moveTo(lastPos.x, lastPos.y);
      contextCnv.lineTo(mousePos.x, mousePos.y);
      contextCnv.stroke();
      lastPos = mousePos;
      firmado = true;
    }
  }
  
  // Allow for animation
  (function drawLoop () {
    requestAnimFrame(drawLoop);
    renderCanvas();
  })();

  