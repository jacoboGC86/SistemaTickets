var context, cuser, web, context, grupos = [], categorias = [], template, idProcessMan = 0, idSol;

$(document).ready(function() {
    var iebrowser = /*@cc_on!@*/false || !!document.documentMode;
    
    if(iebrowser) {
        alert("¡Ups! Por favor, utilice Google Chrome o Firefox para usar la plataforma");
        window.location.href = "https://promagroupmex.sharepoint.com/sites/Paperless/ST/";
    }

    $("#dtFechaRegistro").val(obtenerFecha(new Date(), "dd/MMM/yyyy"));
    
    context = SP.ClientContext.get_current();
    web = context.get_web();
    cuser = web.get_currentUser();
         
    context.load(web);
    context.load(cuser);
    context.executeQueryAsync(function () {
        loadDepartamentos(function() {
            loadUserProperties(function() { 
                loadGrupos(function () {
                    loadCategorias(function () {
                        // Asigna eventos después de cargar todo y pasa el arreglo de obj
                        asignarEventos(); 
                    });
                });
            });
        });
    },
    function (sender, args) {
        console.error("Error al cargar el contexto: " + args.get_message());
    });
});

function loadDepartamentos(callback) {
    var oListTemplates = context.get_web().get_lists().getByTitle("Configuraciones");
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(`
                            <View>
                                <Query>
                                    <Where>
                                        <Eq>
                                            <FieldRef Name="Title" />
                                            <Value Type="Text">Departamento</Value>
                                        </Eq>
                                    </Where>
                                    <OrderBy>
                                            <FieldRef Name="Valor" Ascending="TRUE" />
                                    </OrderBy>
                                </Query>
                                <RowLimit>300</RowLimit>
                            </View>
                            `);
    var collTemplates = oListTemplates.getItems(camlQuery);

    context.load(collTemplates);

    context.executeQueryAsync(
        function () {
            var enumerator = collTemplates.getEnumerator();
            $("#cmbDepartamento").append("<option value=''  selected >Seleccione un departamento</option>");

            while (enumerator.moveNext()) {
                var item = enumerator.get_current();
                $("#cmbDepartamento").append("<option value='" + item.get_item("Valor") + "' data-idman='" + item.get_item("Persona").get_lookupId() +  "'>" + item.get_item("Valor") + " - " + item.get_item("Persona").get_lookupValue() +  "</option>");
                console.log("Title:", item.get_item("Title"));
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

function loadUserProperties(callback) {
    $.ajax({
        url: web.get_url() + "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose",
            "X-RequestDigest": $("#__REQUESTDIGEST").val(),
            "content-type": "application/json"
        },
        success: function (data, status, xhr) {
            var upp = data.d.UserProfileProperties.results;
            var depto = "";
            var planta = "";

            for(var i = 0; i < upp.length; i++) {
                switch(upp[i].Key) {
                    case "Department": depto = upp[i].Value; break;
                    case "Office": planta = upp[i].Value; break;
                    default:
                }
            }

            if(planta != "") {
                $("#cmbPlanta").val(planta);
                $("#cmbPlanta").prop("disabled", false);
            }

            if(depto != "") {
                if(planta == "Aguascalientes" && (depto == "Equipment & Building Maintenance" || depto == "Logistics" || depto == "Production Assembly"))
                    $("#cmbDepartamento").val(depto + " " + planta);
                else
                    $("#cmbDepartamento").val(depto);

                $("#cmbDepartamento").prop("disabled", false);
            }

            callback();
        },
        error: function (xhr, status, error) {
            console.error("Error al consultar el perfil");
            callback();
        }
    });
}

//Función para Combos Mútilple
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
                var id = item.get_id(); 

                grupos.push({
                    id: id
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

function loadCategorias(callback) {
    console.log("ejecutando load categorías")
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
    // Usa el parámetro camlQueryXml
    camlQuery.set_viewXml(camlQueryXml);

    var collItems = oList.getItems(camlQuery);
    context.load(collItems);

    context.executeQueryAsync(function () {
        var enumerator = collItems.getEnumerator();
        var $select = $("#cmbCategorias");
        
        $select.empty(); 
        $select.append('<option value="none" selected disabled>Seleccione una categoría</option>');

        while (enumerator.moveNext()) {
            var item = enumerator.get_current();
            var title = item.get_item("Title");
            var id = item.get_id();

            if (title && item.get_item("CategoriaPadre") != null) {
                categorias.push({ 
                    Title: item.get_item("CategoriaPadre").get_lookupValue() + " - " + title, 
                    Id: id, 
                    Tipo: item.get_item("TipoCategoria"),
                    Prioridad: item.get_item("Prioridad"),
                    IdTemplate: item.get_item("TemplateAtencion") != null ? item.get_item("TemplateAtencion").get_lookupId() : 0
                });
            }
        }

        categorias = categorias.sort((c1, c2) => { return c1.Title.localeCompare(c2.Title) });
        
        if (callback) 
            callback();
    },
    function (sender, args) {
        console.error(`Error al cargar ${listName}: ${args.get_message()}`);
        if (callback) callback(null); 
    });
}

//fFunción para asignar eventos
function asignarEventos() {
    initializePeoplePicker("peoplePickerDiv");
   
    $("#chkOtroNombre").change(function () {
        if ($(this).prop("checked")) {
            $("#dvPeoplePicker").show();
        } else {
            $("#dvPeoplePicker").hide();
        }
    });

    $("#cmbType").on("change", function() {
        $(this).removeClass("is-invalid");
        $(this).addClass("is-valid");
        $('#cmbCategorias').multiselect("destroy");
        $('#cmbCategorias').html("<option value='' selected disabled>Seleccione una opción</option>");

        for(var i = 0; i < categorias.length; i++) {
            if(categorias[i].Tipo === $("#cmbType").val()) {
                $('#cmbCategorias').append("<option value='" + categorias[i].Id + "'>" + categorias[i].Title + "</option>");
            }
        }

        if($(this).val() === "Change") {
            $("#dvChangeValidacion").show();
        } else {
            $("#dvChangeValidacion").hide();
        }

        $('#cmbCategorias').multiselect({
            buttonWidth: '100%',
            enableCaseInsensitiveFiltering: true,
            enableFiltering: true,
            includeFilterClearBtn: false,
            includeSelectAllOption: false,
            numberDisplayed: 0,
            nonSelectedText: 'Seleccione una opción',
            filterPlaceholder: 'Buscar',
            dropUp: false,
            onChange: function (option, checked) {
                var idCat = $(option).val();
                
                $("#dvDetail").html("");

                var oList = context.get_web().get_lists().getByTitle("Templates");
                var oTemplate = oList.getItemById(categorias.filter((f) => f.Id == idCat)[0].IdTemplate);
                $("#cmbUrgency").val(categorias.filter((f) => f.Id == idCat)[0].Prioridad);

                context.load(oTemplate);
                context.executeQueryAsync(function(){
                    template = oTemplate.get_item("RutaAprobacionEscalacion");
                    var t = JSON.parse(oTemplate.get_item("RutaAprobacionEscalacion"));
                    var ft = t.formTemplate;

                    if(ft != undefined) {
                        for(var j = 0; j < ft.length; j++) {
                            var c = "";
    
                            switch(ft[j].tag) {
                                case "h3": 
                                case "p": 
                                c = "<div class='col-md-12'><" + ft[j].tag + ">" + ft[j].text + "</" + ft[j].tag + "></div>";
                                break;
                                case "checkbox": 
                                c = "<div class='col-md-12'><div class='form-check'><input type='checkbox' class='form-check-input' id='" + ft[j].id + "'/><label class='form-check-label' for='" + ft[j].id + "' >" + ft[j].text + "</label></div></div>";
                                break;
                                case "br": 
                                c = "<br/><br/>";
                                break;
                                case "textarea": 
                                c = "<div class='col-md-12'><div class='form-group'><label for='" + ft[j].id + "'>" + ft[j].text + "</label><textarea class='form-control' id='" + ft[j].id + "'></textarea></div></div>";
                                break;
                            }
    
                            $("#dvDetail").append(c);
                        }
                    }
                }, onError);
            },
        });
    });

    $("#cmbImpactoSeguridadInformacion").on("change", function() {
        if($(this).val() === "Sí") {
            $("#dvImpactosSeguridad").show();
        } else {
            $("#dvImpactosSeguridad").hide();
        }
    });

    $('#cmbCategorias').multiselect({
        buttonWidth: '100%',
        enableCaseInsensitiveFiltering: true,
        enableFiltering: true,
        includeFilterClearBtn: false,
        includeSelectAllOption: false,
        numberDisplayed: 0,
        nonSelectedText: 'Seleccione un tipo de ticket',
        dropUp: false,
    });

    $("#txtTitle, #cmbUrgency, #txtDescripcion").on("change", function() {
        if ($(this).val().length != 0) {
            $(this).addClass("is-valid").removeClass("is-invalid");
        } else {
            $(this).addClass("is-invalid").removeClass("is-valid");
        }
    });

    $("#btnEnviarReporte").on("click", function (e) {
        e.preventDefault();
        var btn = $(this);
        btn.prop("disabled", true);
    
        try {
            var Categoria = $("#cmbCategorias").val();
            var Title = $("#txtTitle").val();
            var Descripcion = $("#txtDescripcion").val();
            var Urgency = $("#cmbUrgency").val();
            var Type = $("#cmbType").val();
            var peoplePicker = SPClientPeoplePicker.SPClientPeoplePickerDict["peoplePickerDiv_TopSpan"];
            var ANombreDe = "";
    
            if (!Categoria || !Title || !Descripcion || Urgency === "none" || Type === "none") {
                alertify.alert("Sistema de Tickets", "Por favor complete todos los campos marcados en rojo.");
                btn.prop("disabled", false);
                return;
            }
    
            if ($("#cmbDepartamento").val() == null) {
                alertify.alert("Sistema de Tickets", "Por favor seleccione un departamento de la lista.");
                btn.prop("disabled", false);
                return;
            }

            if($("#cmbPlanta").val() == null){
                alertify.alert("Sistema de Tickets", "Por favor seleccione la ubicación de la planta.");
                btn.prop("disabled", false);
                return;
            }
    
            var t = JSON.parse(template);
            var status = t.approvalPath.length > 0 ? t.approvalPath[0].stepName : "Assigned";
            var ft = t.formTemplate;
    
            if (ft != undefined) {
                for (var j = 0; j < ft.length; j++) {
                    switch (ft[j].tag) {
                        case "checkbox":
                            ft[j].checked = $("#" + ft[j].id).is(":checked");
                            break;
                        case "textarea":
                            ft[j].value = $("#" + ft[j].id).val();
                            break;
                        default:
                    }
                }
                t.formTemplate = ft;
            }
    
            var oList = context.get_web().get_lists().getByTitle("Tickets");
            var itemCreateInfo = new SP.ListItemCreationInformation();
            var oListItem = oList.addItem(itemCreateInfo);

            if (peoplePicker && peoplePicker.TotalUserCount > 0) {
                var userInfo = peoplePicker.GetAllUserInfo()[0];
                var loginName = userInfo.Key;
    
                var user = web.ensureUser(loginName);
                context.load(user);
    
                context.executeQueryAsync(function () {
                    oListItem.set_item("Title", Title);
                    oListItem.set_item("Descripcion", Descripcion);
                    oListItem.set_item("TipoTicket", Type);
                    oListItem.set_item("TemplateConfiguracion", JSON.stringify(t));
                    oListItem.set_item("Prioridad", Urgency);
                    oListItem.set_item("Status", status);
                    oListItem.set_item("Categoria", Categoria);
                    oListItem.set_item("ImpactoSeguridadInformacion", $("#cmbImpactoSeguridadInformacion option:selected").val());
                    oListItem.set_item("DescripcionImpactoSeguridad", $("#txtDescripcionImpactoSeguridad").val());
                    oListItem.set_item("Department", $("#cmbDepartamento").val());
                    oListItem.set_item("Manager", $("#cmbDepartamento option:selected").data("idman"));
                    oListItem.set_item("Planta", $("#cmbPlanta").val());
                    oListItem.set_item("ProcessManager", $("#cmbPlanta").val() == "Puebla" ? t.processManager.id : (t.processManagerAguascalientes != undefined ? t.processManagerAguascalientes.id : t.processManager.id));
                    oListItem.set_item("ANombreDe", user);
    
                    guardarElemento(oListItem, context, btn, t, status);
    
                }, function (sender, args) {
                    console.error("Error al resolver usuario:", args.get_message());
                    alertify.alert("Error", "No se pudo resolver el usuario seleccionado.");
                    btn.prop("disabled", false);
                });
    
            } else {
                // Si no hay persona seleccionada, se setea null y guarda todo
                oListItem.set_item("Title", Title);
                oListItem.set_item("Descripcion", Descripcion);
                oListItem.set_item("TipoTicket", Type);
                oListItem.set_item("TemplateConfiguracion", JSON.stringify(t));
                oListItem.set_item("Prioridad", Urgency);
                oListItem.set_item("Status", status);
                oListItem.set_item("Categoria", Categoria);
                oListItem.set_item("ImpactoSeguridadInformacion", $("#cmbImpactoSeguridadInformacion option:selected").val());
                oListItem.set_item("DescripcionImpactoSeguridad", $("#txtDescripcionImpactoSeguridad").val());
                oListItem.set_item("Department", $("#cmbDepartamento").val());
                oListItem.set_item("Manager", $("#cmbDepartamento option:selected").data("idman"));
                oListItem.set_item("Planta", $("#cmbPlanta").val());
                oListItem.set_item("ProcessManager", $("#cmbPlanta").val() == "Puebla" ? t.processManager.id : (t.processManagerAguascalientes != undefined ? t.processManagerAguascalientes.id : t.processManager.id));
                oListItem.set_item("ANombreDe", null); 
    
                guardarElemento(oListItem, context, btn, t, status);
            }
        } catch (error) {
            console.error("Error al enviar el formulario:", error.message);
            btn.prop("disabled", false);
        }
    });
    
    function guardarElemento(oListItem, context, btn, t, status) {
        oListItem.update();
        context.load(oListItem);
    
        context.executeQueryAsync(function () {
            idSol = oListItem.get_id();
            
            var archivos = selectedFiles;
            var spFiles = [];
            var oExpediente = context.get_web().get_lists().getByTitle("Expediente");
            var itemCreateInfoFolder = new SP.ListItemCreationInformation();

            itemCreateInfoFolder.set_underlyingObjectType(SP.FileSystemObjectType.folder);
            itemCreateInfoFolder.set_leafName(idSol);
            var oListItemFolder = oExpediente.addItem(itemCreateInfoFolder);

            oListItemFolder.update();

            context.load(oListItemFolder);
            context.executeQueryAsync(function (sender, args) {
                subirArchivos(archivos, 0, spFiles, web.get_url(), web.get_serverRelativeUrl(), idSol, function () {
                    if (status != "Assigned") {
                        var oListAp = context.get_web().get_lists().getByTitle("Aprobaciones");
                        var itemCreateInfoAp = new SP.ListItemCreationInformation();
                        var oListItemAp = oListAp.addItem(itemCreateInfoAp);
        
                        oListItemAp.set_item("Title", status);
                        oListItemAp.set_item("Resultado", "Pendiente");
                        oListItemAp.set_item("Ticket", idSol);
        
                        if (t.approvalPath[0].approvalGroup != null) {
                            oListItemAp.set_item("Grupo", t.approvalPath[0].approvalGroup);
                        } else {
                            oListItemAp.set_item("Responsable", $("#cmbDepartamento option:selected").data("idman"));
                        }
        
                        oListItemAp.update();
                        context.load(oListItemAp);
        
                        context.executeQueryAsync(function () {
                            alertify.alert("Sistema de Tickets", "Solicitud realizada correctamente.", function () {
                                window.location.href = "Consulta Tickets.aspx";
                            });
                        });
                    } else {
                        var oListAp = context.get_web().get_lists().getByTitle("Aprobaciones");
                        var itemCreateInfoAp = new SP.ListItemCreationInformation();
                        var oListItemAp = oListAp.addItem(itemCreateInfoAp);
                    
                        oListItemAp.set_item("Title", "Atención de ticket");
                        oListItemAp.set_item("Resultado", "Pendiente");
                        oListItemAp.set_item("Ticket", idSol);
                        oListItemAp.set_item("Responsable", $("#cmbPlanta").val() == "Puebla" ? t.processManager.id : t.processManagerAguascalientes.id);
                    
                        oListItemAp.update();
                        context.load(oListItemAp);
                    
                        context.executeQueryAsync(function () {
                            alertify.alert("Sistema de Tickets", "Solicitud realizada correctamente.", function () {
                                window.location.href = "Consulta Tickets.aspx";
                            });
                        }, function (sender, args) {
                            console.error("Error al guardar en Aprobaciones:", args.get_message());
                            btn.prop("disabled", false);
                        });
                    }
                });
            }, onRequestFail);
        }, function (sender, args) {
            console.error("Error al guardar el elemento:", args.get_message());
            btn.prop("disabled", false);
        });
    }
    
    let selectedFiles = [];

    $("#btnLoadAttachments").on("click", function () {
      $("#fileDocto").click();
    });
    
    $("#fileDocto").on("change", function () {
      if (this.files.length === 0) return;
    
      let newFiles = Array.from(this.files);
    
      if (selectedFiles.length + newFiles.length > 2) {
        alert("Solo puedes subir un máximo de 2 archivos.");
        $(this).val("");
        return;
      }
    
      selectedFiles = selectedFiles.concat(newFiles);
      $(this).val("");
      renderFiles();
    });
    
    function renderFiles() {
        $("#ulFormat").html("");
        $("#dvRepresentacion").html("");
        $("#RemoveFile").toggle(selectedFiles.length > 0);
      
        selectedFiles.forEach((file, index) => {
          $("#ulFormat").append(`
            <li class='list-group-item d-flex justify-content-between align-items-center' data-index="${index}">
              ${file.name}
              <button type="button" class='btn btn-sm btn-danger btn-remove-file'>Eliminar</button>
            </li>
          `);
      
          const extension = file.name.split(".").pop().toLowerCase();
          if (["jpg", "jpeg", "png"].includes(extension)) {
            $("#dvRepresentacion").append(`
              <div class='col-md-4'>
                <img src='${URL.createObjectURL(file)}' class='img-thumbnail mb-2' />
              </div>
            `);
          }
        });
    }
      
    // Evento delegado para los botones individuales
    $(document).on("click", ".btn-remove-file", function () {
        selectedFiles.splice($(this).closest("li").data("index"), 1);

        renderFiles();
    });
    
    $("#RemoveFile").on("click", function () {
      selectedFiles = [];
      renderFiles();
    });
    
    $("#btnCancelar").click(function () {
        window.location.href = "https://promagroupmex.sharepoint.com/sites/Paperless/ST/SitePages/Home.aspx";
    });
}

function onError(sender, args) {
    console.log('Error: ' + args.get_message());
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
    
function subirArchivos(archivos, indice, spFiles, url, serverRelativeUrl, idSol, callback) {
    if (archivos.length > indice) {
        var archivo = archivos[indice];

        /*
        var invalidChars = /[\/\\:*?"<>|#%&{}~]/;

        if (invalidChars.test(archivo.name) || archivo.name.startsWith(".") || archivo.name.endsWith(".")) {
            console.error("El nombre del archivo es inválido:", archivo.name);
            indice++;
            subirArchivos(archivos, indice, spFiles, url, serverRelativeUrl, idSol, callback);
            return;
        }

        if (archivo.size > 10485760) {
            console.error("El archivo excede el límite de tamaño permitido (10 MB):", archivo.name);
            indice++;
            subirArchivos(archivos, indice, spFiles, url, serverRelativeUrl, idSol, callback);
            return;
        }
            */

        var getFile = getFileBuffer(archivo);
        
        getFile.done(function (arrayBuffer) {
            var fileName = archivo.name;
            var fileCollectionEndpoint = serverRelativeUrl + "/_api/web/getfolderbyserverrelativeurl('" + web.get_serverRelativeUrl() + "/Expediente/" + idSol + "')/files/add(overwrite=true,url='" + fileName + "')";

            $.ajax({
                url: fileCollectionEndpoint,
                type: "POST",
                data: arrayBuffer,
                processData: false,
                headers: {
                    "accept": "application/json;odata=verbose",
                    "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                    "content-length": arrayBuffer.byteLength,
                    "content-type": "application/octet-stream"
                },
                success: function (file, status, xhr) {
                    spFiles.push(file);
                    indice++;
                    subirArchivos(archivos, indice, spFiles, url, serverRelativeUrl, idSol, callback);
                },
                error: function (xhr, status, error) {
                    console.error("Error al subir el archivo:", fileName, xhr.responseText);
                    indice++;
                    subirArchivos(archivos, indice, spFiles, url, serverRelativeUrl, idSol, callback);
                }
            });
        }).fail(function () {
            indice++;
            subirArchivos(archivos, indice, spFiles, url, serverRelativeUrl, idSol, callback);
        });
    } else {
        callback(spFiles);
    }
}