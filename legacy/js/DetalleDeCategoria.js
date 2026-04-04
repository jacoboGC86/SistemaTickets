var context, contextCol, cuser, web, idcode, idCanSol = 0, subiendo = false, tipoDocCargar, departamento = 0, doc = 0, status = "", process = [], category = 0, tipoSol;
//var idTask = 0, firmado = false, iamENG = false, drawing = false, mousePos = { x:0, y:0 }, lastPos, haySolCan = false;
var posicionValorGlobal = "";
var matchingValues = [];

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

        //Paso 1: Carga el combo templates
        loadTemplates(function () {
            //Paso 2: Carga el combo Grupos
            loadGrupos(function () {
                //Paso 3: Carga el combo CatPadre 
                loadCatPadre(function () {
                    //Paso 4: Después de cargar todos los combos carga la información de la lista
                    cargarDatos(idcode);
                    asignarEventos(idcode);
                });
            });
        });

        // Actualiza del combo CatPadre, cuando cambie el tipo de categoría 
        $("#cmbTipoCat").on("change", function () {
            $("#cmbCatPadre")
                .empty()
                .append('<option value="none" selected disabled>Seleccione una opción</option>');
            loadCatPadre(function () {
            });
        });
    }, onRequestFail);
});

function loadCatPadre(callback) {
    var TipoCat = $("#cmbTipoCat").val();

    if (!TipoCat) {
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

    var collCategorias = oListTemplates.getItems(camlQuery);

    context.load(collCategorias);
    context.executeQueryAsync(
        function () {
            var enumerator = collCategorias.getEnumerator();
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
    context.executeQueryAsync(function (sender, args) {
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

function loadGrupos(callback) {
    var oListConf = context.get_web().get_lists().getByTitle("Grupos");
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
    context.executeQueryAsync(function (sender, args) {
        var listItemEnumerator = collListItemConf.getEnumerator();
        var $select = $("#CmbGrupos");
        $select.empty();

        //$select.append('<option value="none" selected disabled>Seleccione una opcion </option>');
        var itemCount = 0;

        while (listItemEnumerator.moveNext()) {
            var oListItem = listItemEnumerator.get_current();
            var title = oListItem.get_item("Title");

            if (title) {
                var $option = $("<option ></option>")
                    .val(oListItem.get_id())
                    .text(title);

                $select.append($option);
                itemCount++;
            }
        }

        if (itemCount === 0) {
            console.log("No se encontraron elementos en Grupos");
        } else {
            console.log(itemCount + " elementos agregados al combo Grupos.");
        }

        // Inicializa el multiselect 
        if (!$select.data('multiselect')) {
            $select.multiselect({
                buttonWidth: '100%',
                enableCaseInsensitiveFiltering: true,
                enableFiltering: true,
                includeFilterClearBtn: false,
                includeSelectAllOption: false,
                numberDisplayed: 1,
                nonSelectedText: 'Seleccione una o varias opciones',
                dropUp: false
            });
        } else {
            $select.multiselect("rebuild");
        }

        // Deshabilitar solo las opciones 
        $select.find("option").prop("disabled", true);
        $select.multiselect("rebuild");


        if (typeof callback === "function") {
            callback();
        }
    }, onRequestFail);
}

function cargarDatos(idcode) {
    var oList = context.get_web().get_lists().getByTitle("Categorias");
    var oListItem = oList.getItemById(idcode);

    context.load(oListItem);
    context.executeQueryAsync(function (sender, args) {
        $("#txtFolio").val(oListItem.get_id());
        $("#dtFechaRegistro").val(obtenerFecha(oListItem.get_item("Created"), "dd/mm/yyyy"));
        $("#txtCategoria").val(oListItem.get_item("Title"));
        $("#txtDescripcion").val(oListItem.get_item("Descripcion"));
        $("#txtSLALow").val(oListItem.get_item("SLALow"));
        $("#txtSLAMedium").val(oListItem.get_item("SLAMedium"));
        $("#txtSLAHigh").val(oListItem.get_item("SLAHigh"));
        $("#txtSLAStopProduction").val(oListItem.get_item("SLAProductionStop"));

        const textareaD = document.getElementById('txtDescripcion');

        textareaD.style.height = 'auto';
        textareaD.style.height = textareaD.scrollHeight + 'px';
        textareaD.addEventListener('input', function () {
            this.style.height = 'auto';
            this.style.height = this.scrollHeight + 'px';
        });

        $("#txtProcedimientos").val(oListItem.get_item("ProcedimientoAtencion"));

        const textareaE = document.getElementById('txtProcedimientos');
        textareaE.style.height = 'auto';
        textareaE.style.height = textareaE.scrollHeight + 'px';
        textareaE.addEventListener('input', function () {
            this.style.height = 'auto';
            this.style.height = this.scrollHeight + 'px'; 9
        });

        $("#cmbPrioridad").val(oListItem.get_item("Prioridad"));
        $("#cmbTipoCat").val(oListItem.get_item("TipoCategoria"));

        //Procesa Templates
        var TemplateField = oListItem.get_item("TemplateAtencion");
        console.log("TRAE template", oListItem.get_item("TemplateAtencion"));
        var TemplateName = TemplateField.get_lookupValue();
        if (TemplateName) {

            var foundTemplate = false;
            $("#CmbTemplates option").each(function () {
                var optionValue = $(this).text();
                if (optionValue === TemplateName) {
                    $(this).prop("selected", true);
                    foundTemplate = true;
                    return false;
                }
            });

            if (!foundTemplate) {
                console.warn("No se encontró un valor coincidente en el combo.");
            }
        } else {
            console.error("El campo TemplateAtencion no tiene un valor asignado.");
        }

        //Procesa categoría padre
        var PadreField = oListItem.get_item("CategoriaPadre");
        if (PadreField) {
            var PadreName = PadreField.get_lookupValue();

            // Asegura de que loadCatPadre haya terminado antes de realizar la comparación
            loadCatPadre(function () {
                var foundPadre = false;

                $("#cmbCatPadre option").each(function () {
                    var optionValue = $(this).text();
                    // console.log(`Revisando opción: '${optionValue}'`);
                    if (optionValue === PadreName) {
                        $(this).prop("selected", true);
                        foundPadre = true;
                        return false;
                    }
                });

                if (!foundPadre) {
                    console.warn("No se encontró un valor coincidente en el combo.");
                }
            });
        } else {
            console.log("El campo CategoriaPadre no tiene un valor asignado.");
        }

        //Procesa Grupos   
        var GrupoField = oListItem.get_item("GruposPermitidos");
        
        if (GrupoField && GrupoField.length > 0) {
            var matchingValues = [];

            GrupoField.forEach(function (report) {
                var GrupoName = report.get_lookupValue();
                // console.log("Comparando con ID:", lookupId);
                if (GrupoName) {
                    $("#CmbGrupos option").each(function () {
                        var optionText = $(this).text();
                        var optionValue = $(this).val();

                        if (optionText === GrupoName) {
                            matchingValues.push(optionValue);
                        }
                    });
                }
            });

            if (matchingValues.length > 0) {
                $("#CmbGrupos").val(matchingValues).multiselect("refresh");

                if ($("#CmbGrupos").hasClass("select2-hidden-accessible")) {
                    $("#CmbGrupos").multiselect().val(matchingValues).trigger("change");
                } else if ($("#CmbGrupos").data("multiselect")) {
                    $("#CmbGrupos").trigger("chosen:updated");
                }

                var selectedValues = $("#CmbGrupos").val();
                //console.log("Valores seleccionados en el combo:", selectedValues);

                if (!selectedValues || selectedValues.length === 0) {
                    console.log("La selección no se aplicó correctamente. Revisa la configuración del multiselect.");
                }
            } else {
                console.log("No se encontraron valores coincidentes en el combo.");
            }
        } else {
            console.log("El campo GruposPermitidos está vacío o no tiene valores asignados.");
        }

           $("#txtCategoria, #txtDescripcion, #cmbPrioridad, #CmbTemplates, #txtProcedimientos, #cmbTipoCat, #txtSLALow, #txtSLAMedium, #txtSLAHigh, #txtSLAStopProduction").trigger("change");

        asignarEventos();
    }, function (sender, args) {
        console.error('Error en la consulta o carga: ' + args.get_message());
        asignarEventos();
    });
}

function asignarEventos(idcode) {
    $("#btnCerrar").click(function () {
        window.location.href = "Consulta Categoria.aspx";
    });



    $("#btnEdit").click(function () {
        // Habilita los campos
        $("#txtCategoria, #txtDescripcion, #cmbPrioridad, #CmbTemplates, #txtProcedimientos, #cmbTipoCat, #CmbGrupos, #cmbCatPadre, #txtSLALow, #txtSLAMedium, #txtSLAHigh, #txtSLAStopProduction").prop("disabled", false);
        //Cambio de clases
     $("#txtCategoria, #txtDescripcion, #cmbPrioridad, #CmbTemplates, #txtProcedimientos, #txtSLALow, #txtSLAMedium, #txtSLAHigh, #txtSLAStopProduction").trigger("change");

        var $select = $("#CmbGrupos");
        if (!$select.data("multiselect")) {
            $select.multiselect({
                buttonWidth: "100%",
                enableCaseInsensitiveFiltering: true,
                enableFiltering: true,
                includeFilterClearBtn: false,
                includeSelectAllOption: false,
                numberDisplayed: 1,
                nonSelectedText: false,
                nonSelectedText: "Seleccione una o varias opciones",
                dropUp: false
            });
        } else {
            $select.find("option").prop("disabled", false);
            $select.multiselect("rebuild");
        }

        $("#txtCategoria, #txtDescripcion, #cmbPrioridad, #CmbTemplates, #txtProcedimientos, #cmbTipoCat, #txtSLALow, #txtSLAMedium, #txtSLAHigh, #txtSLAStopProduction").on("change", function () {
            if ($(this).val().length !== 0) {
                $(this).addClass("is-valid").removeClass("is-invalid");
            } else {
                $(this).addClass("is-invalid").removeClass("is-valid");
            }
        });
    });
    $("#btnUpdate").on("click", function (e) {
        e.preventDefault();
        var btn = $(this);
        btn.prop("disabled", true);

        // Validaciones antes del try-catch
        var Categoria = $("#txtCategoria").val();
        var Prioridad = $("#cmbPrioridad").val();
        var Templates = $("#CmbTemplates").val();
        var Descripcion = $("#txtDescripcion").val();
        var Procedimiento = $("#txtProcedimientos").val();
        var CatPadre = $("#cmbCatPadre").val();
        var TipoCat = $("#cmbTipoCat").val();
        var Grupos = [];

        $("#CmbGrupos option:selected").each(function () {
            const lookupId = $(this).val();
            if (lookupId) {
                var lookupValue = new SP.FieldLookupValue();
                lookupValue.set_lookupId(parseInt(lookupId, 10));
                Grupos.push(lookupValue);
            }
        });

        // Validaciones
        if (!Categoria || !Prioridad || Templates === "none" || !Descripcion || !Procedimiento || !$("#txtSLALow").val() || !$("#txtSLAMedium").val() || !$("#txtSLAHigh").val() ||!$("#txtSLAStopProduction").val()) {
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
        try {
            var oList = context.get_web().get_lists().getByTitle("Categorias");
            var oListItem = oList.getItemById(idcode);

            // Operaciones con SharePoint
            oListItem.set_item("Title", Categoria);
            oListItem.set_item("Prioridad", Prioridad);
            oListItem.set_item("TemplateAtencion", parseInt(Templates, 10));
            oListItem.set_item("Descripcion", Descripcion);
            oListItem.set_item("ProcedimientoAtencion", Procedimiento);
            oListItem.set_item("TipoCategoria", TipoCat);
            oListItem.set_item("SLALow", $("#txtSLALow").val());
            oListItem.set_item("SLAMedium", $("#txtSLAMedium").val());
            oListItem.set_item("SLAHigh", $("#txtSLAHigh").val());
            oListItem.set_item("SLAProductionStop", $("#txtSLAStopProduction").val());

            if(Grupos.length > 0)
                oListItem.set_item("GruposPermitidos", Grupos);

            if (CatPadre) {
                var lookupPadre = new SP.FieldLookupValue();

                lookupPadre.set_lookupId(parseInt(CatPadre, 10));
                oListItem.set_item("CategoriaPadre", lookupPadre);
            } else {
                oListItem.set_item("CategoriaPadre", null);
            }

            oListItem.update();
            context.load(oListItem);
        } catch (error) {
            console.error("Error en las operaciones con SharePoint:", error.message);
            btn.prop("disabled", false);
            return;
        }

        context.executeQueryAsync(
            function () {
                console.log("Elemento actualizado correctamente.");
                alertify.alert("Sistema de Tickets", "El formulario se ha actualizado exitosamente.", function () {
                    window.location.href = "Consulta Categoria.aspx";
                });
                btn.prop("disabled", false);
            },
            function (sender, args) {
                console.error("Error al ejecutar la consulta asincrónica:", args.get_message());
                alertify.alert("Sistema de Tickets", "Ocurrió un error al actualizar el formulario. Intente nuevamente.");
                btn.prop("disabled", false);
            }
        );

    });

}
