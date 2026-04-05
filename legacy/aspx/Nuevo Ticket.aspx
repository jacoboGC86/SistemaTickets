
<%-- _lcid="3082" _version="16.0.19520" _dal="1" --%>
<%-- _LocalBinding --%>
<%@ Page language="C#" MasterPageFile="~masterurl/default.master"    Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage,Microsoft.SharePoint,Version=16.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c"  %> <%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> <%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> <%@ Import Namespace="Microsoft.SharePoint" %> <%@ Assembly Name="Microsoft.Web.CommandUI, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> <%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<asp:Content ContentPlaceHolderId="PlaceHolderPageTitle" runat="server">
	Nuevo ticket
</asp:Content>

<asp:Content ContentPlaceHolderId="PlaceHolderAdditionalPageHead" runat="server">
	<meta name="GENERATOR" content="Microsoft SharePoint" />
	<meta name="ProgId" content="SharePoint.WebPartPage.Document" />
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta name="CollaborationServer" content="SharePoint Team Web Site" />
	<SharePoint:ScriptBlock runat="server">
	var navBarHelpOverrideKey = "WSSEndUser";
	</SharePoint:ScriptBlock>
	<SharePoint:StyleBlock runat="server">
	body #s4-leftpanel {
		display:none;
	}
	.s4-ca {
		margin-left:0px;
	}
	</SharePoint:StyleBlock>
	<link rel="stylesheet" href="../SiteAssets/css/jquery-ui.css" />
	<link rel="stylesheet" href="../SiteAssets/css/bootstrap.min.css" />
	<link rel="stylesheet" href="../SiteAssets/css/bootstrap-multiselect.css" />
	<link rel="stylesheet" href="../SiteAssets/css/alertify.min.css"/>
	<link rel="stylesheet" type="text/css" href="../SiteAssets/css/Contratos.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-multiselect/0.9.15/css/bootstrap-multiselect.css">
	<link rel="stylesheet" href="path/to/jquery.multiselect.css">

	<script src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-multiselect/0.9.15/js/bootstrap-multiselect.min.js"></script>
	<script src="../SiteAssets/js/jquery.min.js"></script>
	<script src="../SiteAssets/js/jquery-ui.js"></script>
	<script src="../SiteAssets/js/moment.js"></script>
	<script src="../SiteAssets/js/numeral.min.js"></script>
	<script src="../SiteAssets/js/popper.min.js"></script>
	<script src="../SiteAssets/js/bootstrap.min.js"></script>
	<script src="../SiteAssets/js/bootstrap-multiselect.js"></script>
	<script src="../SiteAssets/js/alertify.min.js"></script>
	
	<!--<script src="path/to/jquery.js"></script>
	<script src="path/to/jquery.multiselect.js"></script>-->
	
	<SharePoint:ScriptLink name="clienttemplates.js" runat="server" LoadAfterUI="true" Localizable="false" />
	<SharePoint:ScriptLink name="clientforms.js" runat="server" LoadAfterUI="true" Localizable="false" />
	<SharePoint:ScriptLink name="clientpeoplepicker.js" runat="server" LoadAfterUI="true" Localizable="false" />
	<SharePoint:ScriptLink name="autofill.js" runat="server" LoadAfterUI="true" Localizable="false" />

	<script src="../_layouts/15/init.js"></script>
	<script src="../_layouts/15/MicrosoftAjax.js"></script>
	<script src="../_layouts/15/SP.Runtime.js"></script>
	<script src="../_layouts/15/SP.js"></script>
	

	<script src="../SiteAssets/js/Util.js"></script>
	<script src="../SiteAssets/js/NuevoTicket.js"></script>
	
</asp:Content>

<asp:Content ContentPlaceHolderId="PlaceHolderSearchArea" runat="server">
	<SharePoint:FlightedContent runat="server" ExpFeature="Reserved_Server_ExpFeature30731" RenderIfInFlight="true">
		<SharePoint:SearchInNavBarEnabledContent runat="server" RenderIfEnabled="false">
			<SharePoint:DelegateControl runat="server" ControlId="SmallSearchInputBox" />
		</Sharepoint:SearchInNavBarEnabledContent>
		<SharePoint:SearchInNavBarEnabledContent runat="server" RenderIfEnabled="true">
			<SharePoint:WebTemplateBasedContent runat="server" WebTemplates="STS|BLANKINTERNET|CMSPUBLISHING|GROUP" RenderIfInWebTemplates="false">
				<SharePoint:DelegateControl runat="server" ControlId="SmallSearchInputBox" />
			</SharePoint:WebTemplateBasedContent>
		</Sharepoint:SearchInNavBarEnabledContent>
	</SharePoint:FlightedContent>
	<SharePoint:FlightedContent runat="server" ExpFeature="Reserved_Server_ExpFeature30731" RenderIfInFlight="false">
		<SharePoint:DelegateControl runat="server" ControlId="SmallSearchInputBox" />
	</SharePoint:FlightedContent>
</asp:Content>

<asp:Content ContentPlaceHolderId="PlaceHolderPageDescription" runat="server">
	<SharePoint:ProjectProperty Property="Description" runat="server"/>
</asp:Content>

<asp:Content ContentPlaceHolderId="PlaceHolderMain" runat="server">
	<div class="ms-hide">
		<WebPartPages:WebPartZone runat="server" title="loc:TitleBar" id="TitleBar" AllowLayoutChange="false" AllowPersonalization="false" Style="display:none;" />
	</div>
  	<table class="ms-core-tableNoSpace ms-webpartPage-root" width="100%">
		<tr>
			<td id="_invisibleIfEmpty" name="_invisibleIfEmpty" valign="top" width="100%"> <WebPartPages:WebPartZone runat="server" Title="loc:FullPage" ID="FullPage" FrameType="TitleBarOnly"/> </td>
		</tr>
		<SharePoint:ScriptBlock runat="server">if(typeof(MSOLayout_MakeInvisibleIfEmpty) == "function") {MSOLayout_MakeInvisibleIfEmpty();}</SharePoint:ScriptBlock>
	</table>

<div class="container" style="margin-top: 20px;min-height: 600px; background-color: white;">
	<div id="dvHeadForm" class="form-row">	
		<div class="form-group col-md-8">
			<h5 style="margin-top: 35px;text-align: center;">Nuevo ticket</h5>
		</div>
      	<div class="form-group col-md-4">
        	<img src="../SiteAssets/img/Logo.png" style="width: 150px;margin-top: 40px;"/>
      	</div>
	</div>
	<hr id="hrHead2"/>
    <!-- Migrar desde aquí en adelante -->
	<div class="form-row">	
		<div class="form-group col-md-2">
			<!--<label for="IdRegistro">Fecha de registro:</label>
			<input type="text" class="form-control" id="txtIdRegistro" placeholder="Folio de registro" disabled="disabled" title="Folio de registro de la solicitud"/>-->
		</div>
		<div class="form-group col-md-8">
		</div>
		<div class="form-group col-md-2">
			<label for="dtFechaRegistro">Fecha de registro:</label>
			<input type="text" class="form-control" id="dtFechaRegistro" placeholder="Fecha de registro" disabled="disabled" title="Fecha de registro de la solicitud"/>
		</div>
	</div>

    <div class="form-row">
		<div class="form-group col-md-4">
			<label for="cmbType">Tipo:</label>
			<select name="cmbType" id="cmbType" class="form-control is-invalid">
				<option value="none" selected disabled>Seleccione una opción</option>
				<option value="Request">Solicitud</option>
				<option value="Incident">Incidente</option>
				<option value="Change">Cambio</option>
			</select>
		</div>
		<div class="form-group col-md-8">
			<label for="cmbCategorias">Categoría:</label>
			<select name="cmbCategorias" id="cmbCategorias" class="form-control" ></select> 
		</div>
		
	</div>
	<div class="form-row">
		<div class="form-group col-md-12">
			<label for="txtTitle">Título:</label>
			<input type="text" name="txtTitle" id="txtTitle" class="form-control is-invalid" />
		</div>
	</div>
	<div class="form-row">
		<div class="form-group col-md-12">
			<label for="txtDescripcion">Descripción detallada:</label>
			<textarea type="text" name="txtDescripcion" id="txtDescripcion" class="form-control is-invalid" ></textarea>
		</div>
	</div>
	<div class="form-row">
		<div class="form-group col-md-6">
			<label for="cmbDepartamento">Departamento:</label>
			<select name="cmbDepartamento" id="cmbDepartamento" class="form-control" ></select> 
		</div>
		<div class="form-group col-md-6">
			<label for="cmbPlanta">Planta:</label>
			<select name="cmbPlanta" id="cmbPlanta" class="form-control" >
				<option value="Aguascalientes">Aguascalientes</option>
				<option value="Puebla">Puebla</option>
			</select> 
		</div>
	</div>

	<div class="form-row">	
		<div class="form-group col-md-4">
			<label for="cmbUrgency ">Urgencia:</label>
			<select name="cmbUrgency" id="cmbUrgency" class="form-control">
				<option value="none" selected disabled>Seleccione una opción</option>
				<option value="Alta">Alta</option>
				<option value="Media">Media</option>
				<option value="Baja">Baja</option>
			</select>
		</div>
		<div class="form-group col-md-4">
			<label for="chkOtroNombre">¿Estás registrando este ticket a nombre de otra persona? </label> <span class="text-muted"> (*Registrar solo si es el caso)</span>
			<div class="form-check">
				<input class="form-check-input" type="checkbox" id="chkOtroNombre" name="chkOtroNombre" >
				<label class="form-check-label" for="chkOtroNombre">
					Sí
				</label>
			</div>
		</div>
		
		<div id="dvPeoplePicker" class="form-group col-md-4" style="display: none;">
			<label for="txtPeople">A nombre de:</label>
			<div id="peoplePickerDiv"></div>
			
		</div>
	</div>

	<div class="form-row" id="dvChangeValidacion" style="display:none;">
		<div class="form-group col-md-4">
			<label for="cmbImpactoSeguridadInformacion">¿Este cambio impacta la seguridad de la informacion?:</label>
			<select name="cmbImpactoSeguridadInformacion" id="cmbImpactoSeguridadInformacion" class="form-control" >
				<option value="No">No</option>
				<option value="Sí">Sí</option>
			</select>
		</div>
		<div id="dvImpactosSeguridad" class="form-group col-md-8" style="display: none;">
			<label for="txtDescripcionImpactoSeguridad">Describe los impactos a la seguridad de la información:</label>
			<textarea type="text" name="txtDescripcionImpactoSeguridad" id="txtDescripcionImpactoSeguridad" class="form-control" ></textarea>
		</div>
	</div>
	
	<div class="form-row" id="dvDetail">
		
	</div>

	<div class="form-row">
		<div class="form-group col-md-12">
		  <button type="button" class="btn btn-info" id="btnLoadAttachments">Cargar archivos</button>
		  <span class="text-muted"> (*Opcional)</span>
	  
		  <input type="file" id="fileDocto" style="display:none" multiple />
		  <ul id="ulFormat" class="list-group"></ul>
		  <div id="dvRepresentacion" class="row"></div>
		  
	
		</div>
	  </div>
	  
	  
	  <input type="file" id="fileDocto" style="display:none" multiple />
	  
	<hr />
	<div class="form-row">
		<div class="form-group col-md-8"></div>
		<div class="form-group col-md-4" style="text-align: right;">
			<button type="button" class="btn btn-primary" id="btnEnviarReporte">Enviar</button>
			<button type="button" class="btn btn-danger" id="btnCancelar">Cancelar</button>
		</div>
	</div>

</div>

	
</asp:Content>
