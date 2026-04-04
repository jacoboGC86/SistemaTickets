var context, contextCol, cuser, web, idcode, subiendo = false, doc = 0, status = "", process = [], category = 0, tipoSol ;
var matchingValues = [], grupos = []; 

$(document).ready(function () {
    $("body").on("contextmenu", "img", function (e) {
        return false;
    });

    var iebrowser = /*@cc_on!@*/ false || !!document.documentMode;
    if (iebrowser) {
        alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
        window.location.href = "https://chimalliapps.sharepoint.com/ST/";
    }

    context = new SP.ClientContext.get_current();
    web = context.get_web();
    cuser = web.get_currentUser();

    context.load(web);
    context.load(cuser);
    context.executeQueryAsync(function () {
        var params = getParameters(window.location.href);

        for (var i = 0; i < params.length; i++) {
            if (params[i].param === "idSol") {
                idcode = params[i].value;
                break;
            }
        }
        loadGrupos(function(){
            for (var i = 0; i < grupos.length; i++) {
                $("#ulGruposDisponibles").append("<li class='list-group-item'><a href='#' class='addGpo' data-approvalgroup='" + grupos[i].id +
                    "' data-nombre='" + grupos[i].name +
                    "' data-useuserdepartment='false' aria-disabled='true' style='pointer-events: none; color: gray;'>Agregar</a> " + grupos[i].name +
                    "</li>");
            }
            
            loadManagers(function () {
                cargarDatos(idcode);
                asignarEventos(idcode);
            });
        });

    }, onRequestFail);
});

function loadGrupos(callback) {
    var oListTemplates = context.get_web().get_lists().getByTitle("Grupos");
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(`<View>
                                <Query>
                                    <OrderBy>
                                        <FieldRef Name="Title" Ascending="TRUE" />
                                    </OrderBy>
                                </Query>
                                <RowLimit>100</RowLimit>
                            </View>
                            `);
    var collTemplates = oListTemplates.getItems(camlQuery);

    context.load(collTemplates);
    context.executeQueryAsync(
        function () {
            var enumerator = collTemplates.getEnumerator();

            while (enumerator.moveNext()) {
                var item = enumerator.get_current();

                grupos.push({
                    id: item.get_id(),
                    name: item.get_item("Title")
                });
            }

            if (callback) 
                callback();
        },
        function (sender, args) {
            console.error("Error al cargar Grupos: " + args.get_message());
            if (callback) 
                callback();
        }
    );
}

function loadTemplates(callback) {
  var oListConf = context.get_web().get_lists().getByTitle("Templates");
  var camlQueryConf = new SP.CamlQuery();
  camlQueryConf.set_viewXml(`
        <View>
            <Query>
                <OrderBy>
                    <FieldRef Name="Title" Ascending="TRUE" />
                </OrderBy>
            </Query>
            <RowLimit>100</RowLimit>
        </View>
    `);
  var collListItemConf = oListConf.getItems(camlQueryConf);
  
  context.load(collListItemConf);
  context.executeQueryAsync(function(sender, args) {
      var listItemEnumerator = collListItemConf.getEnumerator();
      var $select = $("#CmbTemplates");
      $select.empty();

      $select.append('<option value="none" selected disabled>Seleccione una opcion</option>');
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
          console.log("No se encontraron elementos en templates");
      } else {
       console.log(itemCount + " elementos agregados al combo templates.");
      }

      $("#dvLoading").hide();
      $select.show(); 
      
      callback();
  }, onRequestFail);
}

function loadManagers(callback) {
    var oListTemplates = context.get_web().get_lists().getByTitle("Configuraciones");
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(`
        <View>
            <Query>
                <Where>
                    <Eq>
                        <FieldRef Name="Title" />
                        <Value Type="Text">IT</Value>
                    </Eq>
                </Where>
            </Query>
            <ViewFields>
                <FieldRef Name="Personas" />
            </ViewFields>
            <RowLimit>100</RowLimit>
        </View>
    `);
    var collTemplates = oListTemplates.getItems(camlQuery);

    context.load(collTemplates);

    context.executeQueryAsync(
        function () {
            var enumerator = collTemplates.getEnumerator();
            var combo = $("#cmbManager");
            var comboAguascalientes = $("#cmbManagerAguascalientes");

            combo.empty();
            combo.append("<option value='' disabled selected>Seleccione el Manager</option>");
            comboAguascalientes.empty();
            comboAguascalientes.append("<option value='' disabled selected>Seleccione el Manager</option>");

            while (enumerator.moveNext()) {
                var item = enumerator.get_current();
                var personasField = item.get_item("Personas"); 
                if (personasField && personasField.length > 0) {
                    //itera en un foreach ya que es un arreglo
                    personasField.forEach(function (persona) {
                        var id = persona.get_lookupId();
                        var name = persona.get_lookupValue();
                        var user = context.get_web().getUserById(id);
                        context.load(user);

                        context.executeQueryAsync(
                            function () {
                                var email = user.get_email();

                                combo.append(`<option value='${id}' data-email='${email}' data-name='${name}'>${name}</option>`);
                                comboAguascalientes.append(`<option value='${id}' data-email='${email}' data-name='${name}'>${name}</option>`);
                            },
                            function (sender, args) {
                                console.error("Error al cargar el email del usuario: " + args.get_message());
                            }
                        );
                    });
                }
            }
            callback();

            console.log("Managers cargados correctamente.");
        },
        function (sender, args) {
            console.error("Error al cargar Managers: " + args.get_message());
            alert("Ocurrió un error al cargar los datos: " + args.get_message());
        }
    );
}

function cargarDatos(idcode) {   
    var oList = context.get_web().get_lists().getByTitle("Templates");
    var oListItem = oList.getItemById(idcode);
  
    context.load(oListItem);
    context.executeQueryAsync(function(sender, args) {
        $("#txtFolio").val(oListItem.get_id());
        $("#dtFechaRegistro").val(obtenerFecha(oListItem.get_item("Created"), "dd/mm/yyyy"));
        $("#txtNameTemplate").val(oListItem.get_item("Title"));
        $("#txtDescripcion").val(oListItem.get_item("Descripcion"));

        const textareaD = document.getElementById('txtDescripcion');

        textareaD.style.height = 'auto';
        textareaD.style.height = textareaD.scrollHeight + 'px'; 
        textareaD.addEventListener('input', function() {
            this.style.height = 'auto';  
            this.style.height = this.scrollHeight + 'px';  
        });
   
        // Procesa el campo RutaAprobacionEscalacion
        const RutaField = oListItem.get_item("RutaAprobacionEscalacion");

        if (RutaField) {
            try {
                const RutaJSON = JSON.parse(RutaField);

                manejarProcessManager(RutaJSON.processManager, RutaJSON.processManagerAguascalientes);
                manejarFormulario(RutaJSON.formTemplate, false); 
                manejarRutaAprobacion(RutaJSON.approvalPath, false); 

                $("#txtRuta").val(JSON.stringify(RutaJSON.approvalPath, null, 2));
                $("#txtFormTemplate").val(JSON.stringify(RutaJSON.formTemplate, null, 2));
            } catch (error) {
                console.error("Error al analizar el JSON:", error);
            }
        } else {
            console.error("El campo RutaAprobacionEscalacion no tiene un valor asignado.");
        }
    },
    function (sender, args) {
        console.error("Error en la consulta o carga: " + args.get_message());
    });
}

//Funciones con datos obtenidos del JSON RutaAprobacionEscalacion
function manejarProcessManager(processManager, processManagerAguascalientes) {
    if (processManager && processManager.id) {
        const ManagerID = processManager.id;
        const found = $("#cmbManager option").filter(function () {
            return $(this).val() === ManagerID;
        }).prop("selected", true).length > 0;

        if (!found) console.warn("No se encontró un valor coincidente en el combo.");
    } else {
        console.error("No se encontró un ID de Manager en el JSON.");
    }

    if (processManagerAguascalientes && processManagerAguascalientes.id) {
        const ManagerID = processManagerAguascalientes.id;
        const found = $("#cmbManagerAguascalientes option").filter(function () {
            return $(this).val() === ManagerID;
        }).prop("selected", true).length > 0;

        if (!found) console.warn("No se encontró un valor coincidente en el combo.");
    } else {
        console.error("No se encontró un ID de Manager en el JSON.");
    }
}

function manejarFormulario(formTemplate, habilitar = false) {
    if (Array.isArray(formTemplate)) {
        const controlMap = {
            h3: { etiqueta: "Título", placeholder: "Agrega el texto del título...", tipo: "input" },
            p: { etiqueta: "Párrafo", placeholder: "Agrega el texto del párrafo...", tipo: "textarea" },
            textarea: { etiqueta: "Caja de texto", placeholder: "Etiqueta del control...", tipo: "input" },
            checkbox: { etiqueta: "Checkbox", placeholder: "Etiqueta del checkbox...", tipo: "input" },
            br: { etiqueta: "Espacio", placeholder: "", tipo: "span", texto: "Espacio en blanco" }
        };

        // Limpia el formulario antes de llenarlo
        $("#ulFormularioPlantilla").empty();

        formTemplate.forEach(control => {
            const config = controlMap[control.tag];
            if (config) {
                const htmlControl =
                    config.tipo === "span"
                        ? `<span>${config.texto}</span>`
                        : config.tipo === "textarea"
                        ? `<textarea class='form-control label-control' placeholder='${config.placeholder}' ${
                              habilitar ? "" : "disabled"
                          }>${control.text || ""}</textarea>`
                        : `<${config.tipo} class='form-control label-control' placeholder='${config.placeholder}' value='${
                              control.text || ""
                          }' ${habilitar ? "" : "disabled"}></${config.tipo}>`;

                // añade el control en la lista del formulario
                $("#ulFormularioPlantilla").append(`
                    <li class='list-group-item' data-tag='${control.tag}'>
                        <a href='#' class='removeControl' style='pointer-events: ${
                            habilitar ? "auto" : "none"
                        }; color: ${habilitar ? "initial" : "gray"};'>Remover ${config.etiqueta}</a><br/>${htmlControl}
                    </li>
                `);
            } else {
                console.warn(`Tipo de control no identificado: ${control.tag}`);
            }
        });

        // Deshabilita la lista
        if (!habilitar) deshabilitarLista("#ulFormularioPlantilla");
    } else {
        console.warn("El campo formTemplate no es un arreglo.");
    }
}

function manejarRutaAprobacion(approvalPath, habilitar = false) {
    if (Array.isArray(approvalPath)) {
        $("#ulRutaAprobacion").empty();
        approvalPath.forEach(control => {
            $("#ulRutaAprobacion").append(`
                <li class='list-group-item' data-approvalgroup='${control.approvalGroup}' data-stepname='${control.stepName}' data-useuserdepartment='${control.useUserDepartment}'>
                    <a href='#' class='removeGpo' style='pointer-events: ${habilitar ? "auto" : "none"}; color: ${
                habilitar ? "initial" : "gray"
            };'>Remover</a> ${control.stepName}
                </li>
            `);
        });

        if (!habilitar) deshabilitarLista("#ulRutaAprobacion");
    } else {
        console.warn("El campo approvalPath no es un arreglo.");
    }
}

function deshabilitarLista(selector) {
    $(selector).addClass("disabled").find("input, textarea, a").prop("disabled", true).css({
        "pointer-events": "none",
        "color": "gray"
    });
}
function asignarEventos(idcode) {

    $("#btnCerrar").click(function () {
        window.location.href = "Consulta Templates.aspx";
    });
    
    $("#txtNameTemplate, #txtDescripcion,#cmbManager, #cmbManagerAguascalientes").on("change", function() {
        if ($(this).val().length != 0) {
            $(this).addClass("is-valid").removeClass("is-invalid");
        } else {
            $(this).addClass("is-invalid").removeClass("is-valid");
        }
    });
    

    $("#btnEdit").on("click", function () {
        $("#txtNameTemplate, #txtDescripcion, #cmbManager, #cmbManagerAguascalientes").prop("disabled", false);

        $("#txtNameTemplate, #txtDescripcion, #cmbManager, #cmbManagerAguascalientes").on("change", function() {
            if ($(this).val().length != 0) {
                $(this).addClass("is-valid").removeClass("is-invalid");
            } else {
                $(this).addClass("is-invalid").removeClass("is-valid");
            }
        });
        
    
        habilitarLista("#ulFormularioPlantilla");
        habilitarLista("#ulRutaAprobacion");
    });

    $("#btnUpdate").on("click", function (e) {
        e.preventDefault();
        console.log("Botón clickeado.");
    
        var btn = $(this);
        btn.prop("disabled", true);
    
        var Template = $("#txtNameTemplate").val();
        var Descripcion = $("#txtDescripcion").val();
        var selectedOption = $("#cmbManager").find(":selected");
        var selectedOptionAguascalientes = $("#cmbManagerAguascalientes").find(":selected");
    
        calcularFormulario();

        try {
            var Ruta = {
                "approvalPath": JSON.parse($("#txtRuta").val()),
                "formTemplate": JSON.parse($("#txtFormTemplate").val()),
                "processManager": {
                    "id": selectedOption.val(),
                    "email": selectedOption.data("email"),
                    "name": selectedOption.data("name")
                },
                "processManagerAguascalientes": {
                    "id": selectedOptionAguascalientes.val(),
                    "email": selectedOptionAguascalientes.data("email"),
                    "name": selectedOptionAguascalientes.data("name")
                }
            };
        } catch (error) {
            console.error("Error al construir Ruta:", error.message);
            alertify.alert("Sistema de Tickets", "Error en los datos del formulario.");
            btn.prop("disabled", false);
            return;
        }
    
        if (!Template || !Descripcion /*|| !Ruta.processManager.id*/) {
            alertify.alert("Sistema de Tickets", "Por favor complete todos los campos marcados en rojo.");
            btn.prop("disabled", false);
            return;
        }
    
        try {
            var oList = context.get_web().get_lists().getByTitle("Templates");
            var oListItem = oList.getItemById(idcode);
    
            oListItem.set_item("Title", Template);
            oListItem.set_item("Descripcion", Descripcion);
            oListItem.set_item("RutaAprobacionEscalacion", JSON.stringify(Ruta));
            oListItem.set_item("PMPuebla", selectedOption.data("name"));
            oListItem.set_item("PMAguascalientes", selectedOptionAguascalientes.data("name"));
    
            oListItem.update();
            context.load(oListItem);
    
            context.executeQueryAsync(
                function () {
                    console.log("Elemento actualizado correctamente.");
                    alertify.alert("Sistema de Tickets", "El formulario se ha actualizado exitosamente.", function () {
                        window.location.href = "Consulta templates.aspx";
                    });
                    btn.prop("disabled", false);
                },
                function (sender, args) {
                    console.error("Error al ejecutar la consulta asincrónica:", args.get_message());
                    alertify.alert("Sistema de Tickets", "Ocurrió un error al actualizar el formulario. Intente nuevamente.");
                    btn.prop("disabled", false);
                }
            );
        } catch (error) {
            console.error("Error en las operaciones con SharePoint:", error.message);
            alertify.alert("Sistema de Tickets", "Ocurrió un error inesperado al actualizar la información.");
            btn.prop("disabled", false);
        }
    });
    
    
    $(".addGpo").on("click", function(e){
        e.preventDefault();
    
         $("#ulRutaAprobacion").append("<li class='list-group-item' data-approvalgroup='" + $(this).data("approvalgroup") +
            "' data-stepname='" + $(this).data("nombre") +
            "' data-useuserdepartment='" + $(this).data("useuserdepartment") + 
            "' ><a href='#' class='removeGpo'>Remover</a> " + $(this).data("nombre") +
            "</li>");
        $(".removeGpo").off("click");
        $(".removeGpo").on("click", function(e) {
            e.preventDefault();

            $(this).parent().remove();

            calcularRutaAprobacion();
        });

        calcularRutaAprobacion();
    });

    $(".addControl").on("click", function(e) {
        e.preventDefault();

        var control = "", etiqueta = "";

        switch($(this).data("tag")) {
            case "h3": 
                etiqueta = "Título";
                control = "<input type='text' class='form-control label-control' placeholder='Agrega el texto del título...' />"; 
                break;
            case "p": 
            etiqueta = "Párrafo";
                control = "<textarea class='form-control label-control' placeholder='Agrega el texto del párrafo en este espacio...' ></textarea>"; break;
            case "br": 
            etiqueta = "Espacio";
                control = "<span>Espacio en blanco. Útil para hacer separaciones entre controles</span>";
                break;
            case "textarea": 
                etiqueta = "Caja de texto";
              control = "<input type='text' class='form-control label-control' placeholder='Agrega la etiqueta del cuadro de texto...' />";
                break;
            case "checkbox": 
                etiqueta = "Checkbox";
                control = "<input type='text' class='form-control label-control' placeholder='Agrega la etiqueta del checkbox...' />"
                break;
        }

        $("#ulFormularioPlantilla").append("<li class='list-group-item'  data-tag='" + $(this).data("tag") +
                    "' ><a href='#' class='removeControl'>Remover " + etiqueta + "</a><br/>" + control +
                    "</li>");

        $(".label-control").off("change");
        $(".label-control").on("change", function(){
            calcularFormulario();
        });

        $(".removeControl").off("click");
        $(".removeControl").on("click", function(e) {
            e.preventDefault();

            $(this).parent().remove();
            calcularFormulario();
        });

        calcularFormulario();
    });    
}

function habilitarLista(selector) {
    $(selector).removeClass("disabled").find("input, textarea, a").prop("disabled", false).css({
        "pointer-events": "auto",
        "color": "initial"
    });
         // Habilita los campos
         $("#txtNameTemplate, #txtDescripcion, #cmbManager ").prop("disabled", false);
    
         $("#txtNameTemplate, #txtDescripcion, #cmbManager").on("change", function () {
             if ($(this).val().length !== 0) {
                 $(this).addClass("is-valid").removeClass("is-invalid");
             } else {
                 $(this).addClass("is-invalid").removeClass("is-valid");
             }
         });
     
         // Habilita el enlace dentro de lista "ulGruposDisponibles"
         $(".list-group .addGpo").each(function () {
             $(this).attr("aria-disabled", "false");
             $(this).css({
                 "pointer-events": "auto",
                 "color": "initial"
             });
         });
         // Habilita el enlace dentro de lista "controlList"
         $(".list-group .addControl").each(function () {
             $(this).attr("aria-disabled", "false");
             $(this).css({
                 "pointer-events": "auto",
                 "color": "initial"
             });
         });
 
         //Habilita la lista ulFormularioPlantilla
         $("#ulFormularioPlantilla").removeClass("disabled").addClass("list-group-item").find("input, textarea, a").prop("disabled", false).css({
             "pointer-events": "auto",
             "color": " initial"
         });
         
         $(".addControl").css({ "color":"#0d6efd"}).on("click", function(e) {
             e.preventDefault();
             calcularFormulario();
         });
        
         $(".removeControl").off("click").css({ "color":"#5B1519"}).on("click", function (e) {
             e.preventDefault();
             $(this).parent().remove();
             calcularFormulario();
         });
 
 
         //habilita la ruta de aprobación
         $("#ulRutaAprobacion").removeClass("disabled").find("input, textarea, a").prop("disabled", false).css({
             "pointer-events": "auto",
             "color": "initial"
         });
         $(".addGpo").off("click").css({ "color":"#0d6efd"}).on("click", function(e){
             e.preventDefault();
         
              $("#ulRutaAprobacion").append("<li class='list-group-item' data-approvalgroup='" + $(this).data("approvalgroup") +
                 "' data-stepname='" + $(this).data("nombre") +
                 "' data-useuserdepartment='" + $(this).data("useuserdepartment") + 
                 "' ><a href='#' class='removeGpo' >Remover</a> " + $(this).data("nombre") +
                 "</li>");
            
     
             $(".removeGpo").off("click");
             $(".removeGpo").css({ "color":"#5B1519"});
             $(".removeGpo").on("click", function(e) {
                 e.preventDefault();
     
                 $(this).parent().remove();
     
                 calcularRutaAprobacion();
             });
     
             calcularRutaAprobacion();
         });
         $(".removeGpo").off("click").css({ "color":"#5B1519"}).on("click", function (e) {
             e.preventDefault();
             $(this).parent().remove();
             calcularRutaAprobacion();
         });
}

function calcularRutaAprobacion() {
    var ruta = [];

    $("#ulRutaAprobacion li").each((index, item) => {
        console.log( $(item).data("stepname"));
        console.log( $(item).data("useuserdepartment"));
        ruta.push({
            
            "stepName": $(item).data("stepname"),
            "useUserDepartment": $(item).data("useuserdepartment"),
            "approvalGroup": $(item).data("approvalgroup")
        });
    });

    $("#txtRuta").val(JSON.stringify(ruta));
}

function calcularFormulario() {
    var formulario = [];

    $("#ulFormularioPlantilla li").each((index, item) => {
        var tag = $(item).data("tag");
        var text = "";

        if (tag === "p" ) {
            text = $(item).find("textarea").val();
        } else if (tag === "h3" || tag === "checkbox" || tag === "textarea") {
            text = $(item).find("input").val();
        }

        formulario.push({
            id: "controlTemplate" + index,
            tag: tag,
            text: text || "" 
        });
    });

    $("#txtFormTemplate").val(JSON.stringify(formulario));
}