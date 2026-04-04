var context, cuser, web, context, grupos = [];

$(document).ready(function() {
    var iebrowser = /*@cc_on!@*/false || !!document.documentMode;
    
    if (iebrowser) {
        alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
        window.location.href = "https://chimalliapps.sharepoint.com/ST/";
    }

    $("#dtFechaRegistro").val(obtenerFecha(new Date(), "dd/MMM/yyyy"));

    context = SP.ClientContext.get_current();
    web = context.get_web();
    cuser = web.get_currentUser();
         
    context.load(web);
    context.load(cuser);
    context.executeQueryAsync(function () {
        loadGrupos(function(){
            for(var i = 0; i < grupos.length; i++) {
                $("#ulGruposDisponibles").append("<li class='list-group-item'><a href='#' class='addGpo' data-approvalgroup='" + grupos[i].id +
                    "' data-nombre='" + grupos[i].name +
                    "' data-useuserdepartment='false'>Agregar</a> " + grupos[i].name + 
                    "</li>");
            }
           
            // llama a asignar eventos
            asignarEventos();
            loadManagers();
            
        });
    }, 
    function (sender, args) {
        console.error("Error al cargar el contexto: " + args.get_message());
    });
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

function asignarEventos() {
    $("#txtNameTemplate, #txtRuta,#txtDescripcion").on("change", function() {
        if ($(this).val().length != 0) {
            $(this).addClass("is-valid").removeClass("is-invalid");
        } else {
            $(this).addClass("is-invalid").removeClass("is-valid");
        }
    });

    $(".addGpo").on("click", function(e){
        e.preventDefault();
      /*  console.log("data name",  $(this).data("nombre"));
        if( $(this).data("nombre")== "Revisión de Manager"){
            $("#ulRutaAprobacion").append("<li class='list-group-item' data-approvalgroup='" + $(this).data("approvalgroup") +
            "' data-nombre='" + $(this).data("nombre") +
            "' data-useuserdepartment='" + $(this).data("useuserdepartment") + 
            "' ><a href='#' class='removeGpo'>Remover</a> " + $(this).data("nombre") + 
            "<select name='cmbManager' id='cmbManager' class='form-control col-md-8'></select> "+
            "</li>");
            
            loadManagers();
        }else{
            $("#ulRutaAprobacion").append("<li class='list-group-item' data-approvalgroup='" + $(this).data("approvalgroup") +
            "' data-nombre='" + $(this).data("nombre") +
            "' data-useuserdepartment='" + $(this).data("useuserdepartment") + 
            "' ><a href='#' class='removeGpo'>Remover</a> " + $(this).data("nombre") +
            "</li>");
        }*/
         $("#ulRutaAprobacion").append("<li class='list-group-item' data-approvalgroup='" + $(this).data("approvalgroup") +
            "' data-nombre='" + $(this).data("nombre") +
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
                control = "<textarea class='form-control label-control' placeholder='Agrega el texto del párrafo en este espacio...'></textarea>"; break;
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

        $("#ulFormularioPlantilla").append("<li class='list-group-item' data-tag='" + $(this).data("tag") +
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

    $("#btnEnviar").on("click", function (e) {
        e.preventDefault();
        var btn = $(this);
        btn.prop("disabled", true);

        calcularRutaAprobacion();
        calcularFormulario();
    
        try {
            var Template = $("#txtNameTemplate").val().trim();
            var Descripcion = $("#txtDescripcion").val().trim();
            var rutaValue = $("#txtRuta").val().trim();
            var formTemplateValue = $("#txtFormTemplate").val().trim();
            
            var selectedOption = $("#cmbManager").find(":selected");
            var selectedOptionAguascalientes = $("#cmbManagerAguascalientes").find(":selected");
    
            // Validar campos vacíos 
            if (!Template) {
                alertify.alert("Sistema de Tickets", "Por favor ingrese un nombre de plantilla.");
                btn.prop("disabled", false);
                return;
            }
    
            if (!Descripcion) {
                alertify.alert("Sistema de Tickets", "Por favor ingrese una descripción.");
                btn.prop("disabled", false);
                return;
            }
          
    
            if (!rutaValue) {
                alertify.alert("Sistema de Tickets", "Por favor defina la ruta de aprobación.");
                btn.prop("disabled", false);
                return;
            }
    
            if (!formTemplateValue) {
                alertify.alert("Sistema de Tickets", "Por favor defina el formulario de la plantilla.");
                btn.prop("disabled", false);
                return;
            }
    
            // Intenta parsear los valores JSON
            var approvalPath, formTemplate;
            try {
                approvalPath = JSON.parse(rutaValue);
                formTemplate = JSON.parse(formTemplateValue);
            } catch (error) {
                alertify.alert("Sistema de Tickets", "Error en el formato de la ruta de aprobación o plantilla. Verifique la información.");
                console.error("Error al parsear JSON:", error);
                btn.prop("disabled", false);
                return;
            }
    
            var Ruta = {
                "approvalPath": approvalPath,
                "formTemplate": formTemplate,
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
    
            if (!Ruta.processManager.id || !Ruta.processManagerAguascalientes.id) {
                alertify.alert("Sistema de Tickets", "Por favor seleccione el gerente de proceso de Puebla y Aguascalientes.");
                btn.prop("disabled", false);
                return;
            }
    
            // Crea el elemento en la lista templates
            var oList = context.get_web().get_lists().getByTitle("Templates");
            var itemCreateInfo = new SP.ListItemCreationInformation();
            var oListItem = oList.addItem(itemCreateInfo);
    
            oListItem.set_item("Title", Template);
            oListItem.set_item("Descripcion", Descripcion);
            oListItem.set_item("RutaAprobacionEscalacion", JSON.stringify(Ruta));
            oListItem.set_item("PMPuebla", selectedOption.data("name"));
            oListItem.set_item("PMAguascalientes", selectedOptionAguascalientes.data("name"));
    
            oListItem.update();
            context.load(oListItem);
            context.executeQueryAsync(function () {
                    btn.prop("disabled", false);
                    alertify.alert("Sistema de Tickets", "Template guardado exitosamente.", function(){
                        window.location.href = "Consulta Templates.aspx";
                    });
                },
                function (sender, args) {
                    console.error("Error al guardar el elemento:", args.get_message());
                    btn.prop("disabled", false);
                });
    
        } catch (error) {
            console.error("Error al enviar el formulario: ", error.message);
            alertify.alert("Sistema de Tickets", "Ocurrió un error inesperado. Intente nuevamente.");
            btn.prop("disabled", false);
        }
    });
    
    $("#btnCancelar").click(function () {
        window.location.href = "https://chimalliapps.sharepoint.com/ST/SitePages/Home.aspx";
    });
}

function calcularRutaAprobacion() {
    var ruta = [];

    $("#ulRutaAprobacion li").each((index, item) => {
        ruta.push({
            "stepName": $(item).data("nombre"),
            "useUserDepartment": $(item).data("useuserdepartment"),
            "approvalGroup": $(item).data("approvalgroup")
        });
    });

    $("#txtRuta").val(JSON.stringify(ruta));
}

function calcularFormulario() {
    var formulario = [];

    $("#ulFormularioPlantilla li").each((index, item) => {
        var t = "";

        switch($(item).data("tag")) {
            case "p": 
                t = $($($(item).find("textarea"))[0]).val();
                break;
            case "h3": 
            t = $($($(item).find("input"))[0]).val();
                break;
            case "textarea": 
            t = $($($(item).find("input"))[0]).val();
            break;
            case "checkbox": 
                t = $($($(item).find("input"))[0]).val();
                break;
        }

        formulario.push({
            "id": "controlTemplate" + index,
            "tag": $(item).data("tag"),
            "text": t
        });
    });

    $("#txtFormTemplate").val(JSON.stringify(formulario));
}

function loadManagers() {
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
            combo.append("<option value='' disabled selected>Seleccione el Gerente</option>");

            comboAguascalientes.empty();
            comboAguascalientes.append("<option value='' disabled selected>Seleccione el Gerente</option>");

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
        },
        function (sender, args) {
            console.error("Error al cargar Managers: " + args.get_message());
            alert("Ocurrió un error al cargar los datos: " + args.get_message());
        }
    );
}



function onError(sender, args) {
    console.log('Error: ' + args.get_message());
}