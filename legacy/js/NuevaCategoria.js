var context, cuser, web, context;

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

        // Carga inicial de Templates y Grupos
        loadTemplates(function (processTitles) {
            if (!processTitles) {
                console.error("Error al cargar los templates.");
                return;
            }

            loadGrupos(function (processGroup) {
                if (!processGroup) {
                    console.error("Error al cargar los grupos.");
                    return;
                }

                // Asignamos los eventos después de cargar todo
                asignarEventos(processTitles, processGroup);

                // Inicializa el cambio dinámico del cmb tipocat
                $("#cmbTipoCat").on("change", function () {
                    // Cuando cambia TipoCat == actualiza CatPadre
                    $("#cmbCatPadre")
                        .empty()
                        .append('<option value="none" selected disabled>Seleccione una opción</option>');

                    loadCatPadre(function () {
                        console.log("CatPadre actualizado.");
                    });
                });
                // Carga inicial de CatPadre
                loadCatPadre();
            });
        });
    }, 
    function (sender, args) {
        console.error("Error al cargar el contexto: " + args.get_message());
    });
});

function loadCatPadre(callback) {
    var TipoCat = $("#cmbTipoCat").val();

    if (!TipoCat) {
        console.error("Tipo de categoría no seleccionada.");
        $("#cmbCatPadre")
            .empty()
            .append('<option value="none" selected disabled>Seleccione una opción</option>');
        if (callback) callback(null);
        return;
    }

    var oListTemplates = context.get_web().get_lists().getByTitle("Categorias");
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(`
        <View>
            <Query>
                <Where>
                    <IsNull>
                        <FieldRef Name="CategoriaPadre" />
                    </IsNull>
                </Where>
                <OrderBy>
                    <FieldRef Name="Title" Ascending="TRUE" />
                </OrderBy>
            </Query>
            <RowLimit>500</RowLimit>
        </View>
    `);

    var collTemplates = oListTemplates.getItems(camlQuery);
    context.load(collTemplates);
    context.executeQueryAsync(
        function () {
            var enumerator = collTemplates.getEnumerator();
            var $select = $("#cmbCatPadre");

            while (enumerator.moveNext()) {
                var item = enumerator.get_current();
                var title = item.get_item("Title");
                var id = item.get_id();

                if (title) {
                    $select.append($("<option></option>").val(id).text(title));
                }
            }
            console.log("CatPadre actualizado.");
            if (callback) callback();
        },
        function (sender, args) {
            console.error("Error al cargar CatPadre: " + args.get_message());
            if (callback) callback();
        }
    );
}

function loadTemplates(callback) {
    var oListTemplates = context.get_web().get_lists().getByTitle("Templates");
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(`
        <View>
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
            var $select = $("#CmbTemplates");
           $select.append('<option value="none" selected disabled>Seleccione una opción</option>');

            var processTitles = {}; 
            while (enumerator.moveNext()) {
                var item = enumerator.get_current();
                var title = item.get_item("Title");
                var id = item.get_id(); 

                if (title) {
                    $select.append($("<option></option>")
                        .val(id) 
                        .text(title));

                        processTitles[id] = { Title: title, LookupId: id };
                }
            }
            console.log("Templates cargados correctamente");
            if (callback) callback(processTitles); // Pasa processTitles al callback
        },
        function (sender, args) {
            console.error("Error al cargar Grupos: " + args.get_message());
            if (callback) callback();
        }
    );
}


function loadGrupos(callback) {
    var oListTemplates = context.get_web().get_lists().getByTitle("Grupos");
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(`
                            <View>
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
            var $select = $("#CmbGrupos");

            var processGroup = {}; 
            while (enumerator.moveNext()) {
                var item = enumerator.get_current();
                var title = item.get_item("Title");
                var id = item.get_id(); 

                if (title) {
                    $select.append($("<option></option>")
                        .val(id) 
                        .text(title));

                        processGroup[id] = { Title: title, LookupId: id };
                }
            }
            console.log("Grupos cargados correctamente");
            if (callback) callback(processGroup); // Pasa processGroup al callback
        },
        function (sender, args) {
            console.error("Error al cargar Grupos: " + args.get_message());
            if (callback) callback();
        }
    );
}

function asignarEventos(processTitles, processGroup) {
    if (!processTitles || !processGroup) {
        console.error("Faltan datos para asignar eventos.");
        return;
    }

    $('#CmbGrupos').multiselect({
        buttonWidth: '100%',
        enableCaseInsensitiveFiltering: true,
        enableFiltering: true,
        includeFilterClearBtn: false,
        includeSelectAllOption: false,
        numberDisplayed: 0,
        filterPlaceholder: 'Buscar',
        nonSelectedText: 'Seleccione una o varias opciones',
        dropUp: false,
        
        onChange: function (option, checked) {
            var Grupos = [];

            $("#CmbGrupos option:selected").each(function () {
                const lookupId = $(this).val(); 
                if (lookupId) {
                    Grupos.push(lookupId);
                }
            });
        },
    });

    $("#txtCategoria, #txtDescripcion, #cmbPrioridad, #CmbTemplates, #txtProcedimientos, #cmbTipoCat, #txtSLALow, #txtSLAMedium, #txtSLAHigh, #txtSLAStopProduction").on("change", function() {
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
            var Categoria = $("#txtCategoria").val();
            var Prioridad = $("#cmbPrioridad").val();
            var Templates = $("#CmbTemplates").val();
            var Descripcion = $("#txtDescripcion").val();
            var Procedimiento = $("#txtProcedimientos").val();
            var CatPadre = $("#cmbCatPadre").val();
            var TipoCat = $("#cmbTipoCat").val();
            
            var Grupos = [];
            $("#CmbGrupos option:selected").each(function () {
                const lookupId = $(this).val(); // Extrae el id del lookup
                if (lookupId) {
                    var lookupValue = new SP.FieldLookupValue(); 
                    lookupValue.set_lookupId(parseInt(lookupId, 10)); // Asigna el ID del lookup
                    Grupos.push(lookupValue);
                }
            });         

            // Validaciones
            if (!Categoria || !Prioridad || Templates === "none" || !Descripcion || !Procedimiento || !$("#txtSLALow").val() || !$("#txtSLAMedium").val() || !$("#txtSLAHigh").val() || !$("#txtSLAStopProduction").val()) {
                alertify.alert("Sistema de Tickets", "Por favor complete todos los campos marcados en rojo.");
                btn.prop("disabled", false);
                return;
            }

            /*
            if (!Grupos || Grupos.length === 0) {
                alertify.alert("Sistema de Tickets", "Por favor seleccione uno o más grupos.");
                btn.prop("disabled", false);
                return;
            }
            */

            // Crear elemento en SharePoint
            var oList = context.get_web().get_lists().getByTitle("Categorias");
            var itemCreateInfo = new SP.ListItemCreationInformation();
            var oListItem = oList.addItem(itemCreateInfo);

            oListItem.set_item("Title", Categoria);
            oListItem.set_item("Prioridad", Prioridad);
            oListItem.set_item("TemplateAtencion", parseInt(Templates, 10));
            oListItem.set_item("Descripcion", Descripcion);
            oListItem.set_item("ProcedimientoAtencion", Procedimiento);
            oListItem.set_item("TipoCategoria", TipoCat);
            oListItem.set_item("CategoriaPadre", CatPadre);
            oListItem.set_item("SLALow", $("#txtSLALow").val());
            oListItem.set_item("SLAMedium", $("#txtSLAMedium").val());
            oListItem.set_item("SLAHigh", $("#txtSLAHigh").val());
            oListItem.set_item("SLAProductionStop", $("#txtSLAStopProduction").val());

            if(Grupos.length > 0)
                oListItem.set_item("GruposPermitidos", Grupos);

            oListItem.update();
            context.load(oListItem);

            context.executeQueryAsync(
                function () {
                    console.log("Elemento guardado correctamente.");
                    alertify.alert("Sistema de Tickets", "Formulario guardado exitosamente.");
                    window.location.href = "Consulta Categoria.aspx";
                    btn.prop("disabled", false);
                },
                function (sender, args) {
                    console.error("Error al guardar el elemento:", args.get_message());
                    btn.prop("disabled", false);
                }
            );
        } catch (error) {
            console.error("Error al enviar el formulario:", error.message);
            btn.prop("disabled", false);
        }
    });

    $("#btnCancelar").click(function () {
        window.location.href = "https://chimalliapps.sharepoint.com/ST/SitePages/Home.aspx";
    });
}

function onError(sender, args) {
    console.log('Error: ' + args.get_message());
}


//Función estandar para combos de una selección las consultas funcionan 
/*
function loadCmbStandar(listName, selectId, camlQueryXml, callback) {
    var oList = context.get_web().get_lists().getByTitle(listName);
    var camlQuery = new SP.CamlQuery();
    // Usa el parámetro camlQueryXml
    camlQuery.set_viewXml(camlQueryXml);

    var collItems = oList.getItems(camlQuery);
    context.load(collItems);

    context.executeQueryAsync(
        function () {
            var enumerator = collItems.getEnumerator();
            var $select = $(selectId);
            $select.empty(); 
            $select.append('<option value="none" selected disabled>Seleccione una opción</option>');

            var itemsData = {};
            while (enumerator.moveNext()) {
                var item = enumerator.get_current();
                var title = item.get_item("Title");
                var id = item.get_id();

                if (title) {
                    $select.append($("<option></option>").val(id).text(title));
                    itemsData[id] = { Title: title, LookupId: id };
                }
            }

            console.log(`${listName} cargada correctamente.`);
            if (callback) callback(itemsData);
        },
        function (sender, args) {
            console.error(`Error al cargar ${listName}: ${args.get_message()}`);
            if (callback) callback(null); 
        }
    );
}

function loadTemplates(callback) {
    var camlQueryXml = `
        <View>
            <Query>
                <OrderBy>
                    <FieldRef Name="Title" Ascending="TRUE" />
                </OrderBy>
            </Query>
            <RowLimit>100</RowLimit>
        </View>
    `;
    // Llama a loadCmbStandar y pasa los parámetros
    loadCmbStandar("Templates", "#CmbTemplates", camlQueryXml, callback);
}

function loadCatPadre(callback) {
    var camlQueryXml = `
                        <View>
                            <Query>
                                <Where>
                                    <IsNull>
                                        <FieldRef Name="CategoriaPadre" />
                                    </IsNull>
                                </Where>
                                <OrderBy>
                                    <FieldRef Name="Title" Ascending="TRUE" />
                                </OrderBy>
                            </Query>
                            <RowLimit>100</RowLimit>
                        </View>
                       `;
    loadCmbStandar("Categorias", "#cmbCatPadre",camlQueryXml, callback);
}*/
