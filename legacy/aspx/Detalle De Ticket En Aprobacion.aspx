<%-- _lcid="3082" _version="16.0.19520" _dal="1" --%>
<%-- _LocalBinding --%>
<%@ Page language="C#" MasterPageFile="~masterurl/default.master"
Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage,Microsoft.SharePoint,Version=16.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c"
%>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls"
Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities"
Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
%>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Assembly
Name="Microsoft.Web.CommandUI, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
%>
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages"
Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
%>
<asp:Content ContentPlaceHolderId="PlaceHolderPageTitle" runat="server">
	Detalle de Ticket en aprobación
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
	<link rel="stylesheet" href="../SiteAssets/css/bootstrap.min.css" />
	<link rel="stylesheet" href="../SiteAssets/css/bootstrap-multiselect.css" />
	<link rel="stylesheet" href="../SiteAssets/css/alertify.min.css" />
	<link rel="stylesheet" type="text/css" href="../SiteAssets/css/Contratos.css" />
	<link rel="stylesheet" type="text/css" href="../SiteAssets/css/EstilosTickets.css" />
	<link rel="stylesheet"
		href="../SiteAssets/css/all.min.css">


	<script src="../SiteAssets/js/jquery.min.js"></script>
	<script src="../SiteAssets/js/moment.js"></script>
	<script src="../SiteAssets/js/numeral.min.js"></script>
	<script src="../SiteAssets/js/popper.min.js"></script>
	<!-- <script src="../SiteAssets/js/bootstrap.min.js"></script> -->
	<script src="../SiteAssets/js/bootstrap-multiselect.js"></script>
	<script src="../SiteAssets/js/alertify.min.js"></script>
	<script src="../SiteAssets/js/bootstrap.bundle.min.js"></script>


	<SharePoint:ScriptLink name="clienttemplates.js" runat="server" LoadAfterUI="true"
		Localizable="false" />
	<SharePoint:ScriptLink name="clientforms.js" runat="server" LoadAfterUI="true"
		Localizable="false" />
	<SharePoint:ScriptLink name="clientpeoplepicker.js" runat="server"
		LoadAfterUI="true" Localizable="false" />
	<SharePoint:ScriptLink name="autofill.js" runat="server" LoadAfterUI="true"
		Localizable="false" />

	<script src="../_layouts/15/init.js"></script>
	<script src="../_layouts/15/MicrosoftAjax.js"></script>
	<script src="../_layouts/15/SP.Runtime.js"></script>
	<script src="../_layouts/15/SP.js"></script>


	<script src="../SiteAssets/js/Util.js"></script>
	<script src="../SiteAssets/js/DetalleDeTicketEnAprobacion.js"></script>

</asp:Content>

<asp:Content ContentPlaceHolderId="PlaceHolderSearchArea" runat="server">
	<SharePoint:FlightedContent runat="server"
		ExpFeature="Reserved_Server_ExpFeature30731" RenderIfInFlight="true">
		<SharePoint:SearchInNavBarEnabledContent runat="server" RenderIfEnabled="false">
			<SharePoint:DelegateControl runat="server"
				ControlId="SmallSearchInputBox" />
		</Sharepoint:SearchInNavBarEnabledContent>
		<SharePoint:SearchInNavBarEnabledContent runat="server" RenderIfEnabled="true">
			<SharePoint:WebTemplateBasedContent runat="server"
				WebTemplates="STS|BLANKINTERNET|CMSPUBLISHING|GROUP"
				RenderIfInWebTemplates="false">
				<SharePoint:DelegateControl runat="server"
					ControlId="SmallSearchInputBox" />
			</SharePoint:WebTemplateBasedContent>
		</Sharepoint:SearchInNavBarEnabledContent>
	</SharePoint:FlightedContent>
	<SharePoint:FlightedContent runat="server"
		ExpFeature="Reserved_Server_ExpFeature30731" RenderIfInFlight="false">
		<SharePoint:DelegateControl runat="server" ControlId="SmallSearchInputBox" />
	</SharePoint:FlightedContent>
</asp:Content>

<asp:Content ContentPlaceHolderId="PlaceHolderPageDescription" runat="server">
	<SharePoint:ProjectProperty Property="Description" runat="server" />
</asp:Content>

<asp:Content ContentPlaceHolderId="PlaceHolderMain" runat="server">
	<div class="ms-hide">
		<WebPartPages:WebPartZone runat="server" title="loc:TitleBar" id="TitleBar"
			AllowLayoutChange="false" AllowPersonalization="false"
			Style="display:none;" />
	</div>
	<table class="ms-core-tableNoSpace ms-webpartPage-root" width="100%">
		<tr>
			<td id="_invisibleIfEmpty" name="_invisibleIfEmpty" valign="top"
				width="100%">
				<WebPartPages:WebPartZone runat="server" Title="loc:FullPage"
					ID="FullPage" FrameType="TitleBarOnly" />
			</td>
		</tr>
		<SharePoint:ScriptBlock runat="server">if(typeof(MSOLayout_MakeInvisibleIfEmpty)
			== "function") {MSOLayout_MakeInvisibleIfEmpty();}</SharePoint:ScriptBlock>
	</table>

	<div class="container"
		style="margin-top: 20px;min-height: 600px; background-color: white;">
		<div id="dvHeadForm" class="form-row">
			<div class="form-group col-md-8">
				<h5 style="margin-top: 35px;text-align: center;">Seguimiento al ticket
				</h5>
			</div>
			<div class="form-group col-md-4">
				<img src="../SiteAssets/img/Logo.png" title="Chimalli Apps"
					style="width: 150px;margin-top: 40px;" />
			</div>
		</div>
		<div id="dvHeaderTicket" style="
				border: solid 1px #eee;
				padding: 5px;
				border-radius: 4pt;
				box-shadow: rgba(0, 0, 0, 0.35) 0px 5px 15px;
				margin-bottom: 15px;">
			<div id="dvStatus" class="row" style="text-align: center;">
				<div class="col-md-3"></div>
				<div class="col-md-3"></div>
				<div class="col-md-3"></div>
			</div>
			<p id="pHeaderTicket">Presiona el nombre de la persona para enviarle un recordatorio de su revisión por Teams.</p>
		
		</div>
		<div id="dvOptionBar" style="
				border: solid 1px #eee;
				padding: 5px;
				border-radius: 4pt;
				box-shadow: rgba(0, 0, 0, 0.35) 0px 5px 15px;
				margin-bottom: 15px;" class="d-flex justify-content-end"  style="display: none" >
				<button type="button" id="btnOpenPopUp" class="icon-btn" style="display: none" data-bs-toggle="tooltip" data-bs-placement="top"data-bs-title="Reasignar Ticket">
					<img src="../SiteAssets/img/cambiar.png" alt="Icono">
				</button>

				<button type="button" id="btnHabilitRecategorizar" class="icon-btn" style="display: none" data-bs-toggle="tooltip" data-bs-placement="top" data-bs-title="Recategorizar Ticket">
					<img src="../SiteAssets/img/procedimiento.png" alt="Icono">
				</button>

				<button type="button" id="btnReaperturar" class="icon-btn" style="display: none" data-bs-toggle="tooltip" data-bs-placement="top" data-bs-title="Reaperturar Ticket">
					<img src="../SiteAssets/img/Reiniciar.png" alt="Icono">
				</button>
				<button type="button" id="btnAtencion" class="icon-btn" style="display: none" data-bs-toggle="tooltip" data-bs-placement="top" data-bs-title="Ticket en Atención">
					<img src="../SiteAssets/img/Atencion.png" alt="Icono">
				</button>

		</div>
		<div id="dvReassignPanel" style="
display:none;
margin-top:10px;
border:1px solid #eee;
padding:15px;
border-radius:6px;
box-shadow: rgba(0,0,0,0.2) 0px 3px 10px;
">

    <div style="font-weight:600;margin-bottom:6px;">
        Reasignar responsable
    </div>

    <div id="peoplePickerReassign"></div>

    <div style="
        margin-top:12px;
        padding:10px;
        background:#f8f9fa;
        border-left:4px solid #ffc107;
        border-radius:4px;
        font-size:13px;
        color:#6c757d;
    ">
        Una vez que se reasigne a otro usuario, perderá el privilegio de aprobar o rechazar el ticket.
    </div>

    <div style="margin-top:12px;text-align:right">
        <button class="btn btn-primary" id="btnConfirmReassign">Reasignar</button>
        <button type="button" class="btn btn-danger" id="btnCancelReassign">Cancelar</button>
    </div>

		</div>
		<div id="dvReaperturaPanel" style="
display:none;
margin-top:10px;
border:1px solid #eee;
padding:15px;
border-radius:6px;
box-shadow: rgba(0,0,0,0.2) 0px 3px 10px;
">

    <div style="font-weight:600;margin-bottom:6px;">
        Reaperturar Ticket
    </div>

    <label style="font-size:13px;">Motivo de la reapertura</label>

    <textarea 
        id="txtMotivoReapertura"
        class="form-control"
        style="width:100%;height:100px;margin-top:5px;"
        placeholder="Ingrese el motivo de la reapertura">
    </textarea>

    <div style="margin-top:12px;text-align:right">
        <button class="btn btn-primary" id="btnConfirmReapertura">Reaperturar</button>
        <button type="button" class="btn btn-danger" id="btnCancelReapertura">Cancelar</button>
    </div>

</div>
		<div class="form-row">
			<div class="form-group col-md-4">
				<label for="txtSolicitante">Solicitante:</label>
				<input type="text" id="txtSolicitante" class="form-control"
					disabled="disabled" />
			</div>
			<div class="form-group col-md-4">

			</div>
			<div class="form-group col-md-2">
				<label for="txtFolio">Folio:</label>
				<input type="text" class="form-control" id="txtFolio"
					placeholder="Folio" disabled="disabled" title="Folio" />
			</div>
			<div class="form-group col-md-2">
				<label for="dtFechaRegistro">Fecha de registro:</label>
				<input type="text" class="form-control" id="dtFechaRegistro"
					placeholder="Fecha de registro" disabled="disabled"
					title="Fecha de registro de la solicitud" />
			</div>
		</div>

		 <div class="form-row">
		<div class="form-group col-md-4">
			<label for="cmbType">Tipo:</label>
			<select name="cmbType" id="cmbType" disabled class="form-control is-invalid">
				<option value="none" selected disabled>Seleccione una opción</option>
				<option value="Request">Request</option>
				<option value="Incident">Incident</option>
				<option value="Change">Change</option>
			</select>
		</div>
		<div class="form-group col-md-8">
			<label for="cmbCategorias">Categoría:</label>
			<select name="cmbCategorias" id="cmbCategorias" disabled class="form-control" ></select> 
		</div>
			
	</div>
	<div class="form-row">
			<div id="dvBotoneraRecategorizar" class="form-group col-md-12"
				style="display: none; text-align: right; ">
				<button type="button" class="btn btn-primary"
					id="btnRecategorizar">Recategorizar</button>
				<button type="button" class="btn btn-danger"
					id="btnCancelarRecat">Cancelar</button>
			</div>
		</div>
		<div class="form-row">
			<div class="form-group col-md-12">
				<label for="txtTitle">Titulo:</label>
				<input type="text" name="txtTitle" id="txtTitle" class="form-control"
					disabled="disabled" />
			</div>
		</div>
		<div class="form-row">
			<div class="form-group col-md-12">
				<label for="txtDescripcion">Descripción detallada:</label>
				<textarea type="text" name="txtDescripcion" id="txtDescripcion"
					class="form-control" disabled="disabled"></textarea>
			</div>
		</div>
		<div class="form-row">
			<div class="form-group col-md-4">
				<label for="cmbDepartamento">Departamento:</label>
				<input type="text" id="cmbDepartamento" class="form-control"
					disabled="disabled" />
			</div>
			<div class="form-group col-md-4">
				<label for="cmbPlanta">Planta:</label>
				<input type="text" id="cmbPlanta" class="form-control"
					disabled="disabled" />
			</div>
			<div class="form-group col-md-4" id="dvANombreDe" style="display: none;">
				<label for="txtANombreDe">A nombre de:</label>
				<input type="text" name="txtANombreDe" id="txtANombreDe"
					class="form-control" disabled="disabled" />
			</div>
		</div>
		<hr />
		<div class="form-row" id="dvDetail">

		</div>


		<div class="form-row">
			<div class="form-group col-md-4">
				<label for="cmbUrgency ">Urgencia:</label>
				<select name="cmbUrgency" id="cmbUrgency" class="form-control"
					disabled="disabled">
					<option value="none" selected disabled>Seleccione una opción
					</option>
					<option value="Alta">Alta</option>
					<option value="Media">Media</option>
					<option value="Baja">Baja</option>
				</select>
			</div>
		</div>

		<div class="form-row" id="dvChangeValidacion" style="display:none;">
			<div class="form-group col-md-4">
				<label for="cmbImpactoSeguridadInformacion">¿Este cambio impacta la seguridad de la informacion?:</label>
				<select name="cmbImpactoSeguridadInformacion" id="cmbImpactoSeguridadInformacion" class="form-control" disabled="disabled">
					<option value="No">No</option>
					<option value="Sí">Sí</option>
				</select>
			</div>
			<div id="dvImpactosSeguridad" class="form-group col-md-8" style="display: none;">
				<label for="txtDescripcionImpactoSeguridad">Describe los impactos a la seguridad de la información:</label>
				<textarea type="text" name="txtDescripcionImpactoSeguridad" id="txtDescripcionImpactoSeguridad" class="form-control" disabled="disabled"></textarea>
			</div>
		</div>
		<hr />
		<div class="form-row">
			<div class="form-group col-md-12" id="dvRepresentacion">
				<span class="lblArchivos">Archivos adjuntos</span>
				<ul id="ulAdjuntos" class="list-group">

				</ul>
			</div>
					
		<div id="dvReapertura" class="form-row form-group col-md-12" style="display:none;">
		<h5>Reaperturado</h5>
		<div class="form-group col-md-12">
			<label for="txtMotivo">Motivo:</label>
			<textarea name="txtMotivo" id="txtMotivo"
				class="form-control" disabled></textarea>
		</div>
	<hr/>
	</div>
		</div>
		<hr>
		<div class="form-row" id="dvComments" style="display: none;">
			<div class="form-group col-md-12">
				<label for="txtComentarios">Comments:</label>
				<textarea name="" id="txtComentarios"
					class="form-control is-invalid"></textarea>
			
			</div>
		</div>


		<div class="form-row">
			<div id="dvArchivo" class="form-group col-md-12" style=" display: none;">

				<button type="button" class="btn btn-info"
					id="btnLoadAttachmentsIT">Cargar archivos</button>
				<span class="text-muted"> (*Opcional)</span>

				<ul id="ulFormatIT" class="list-group mt-2"></ul>

				<div id="dvRepresentacionIT" class="row mt-2"></div>
				<button type="button" class="btn btn-danger mt-2" id="RemoveFileIT"
					style="display:none">Eliminar archivo</button>
					
				<button type="button" class="btn btn-success mt-2" id="btnUpComment">Subir</button>
			</div>
			
		</div>

		<input type="file" id="fileDoctoIT" style="display: none;" />


		<div class="form-row">
			<div class="form-group col-md-12" id="dvComentarios">
			</div>
			<ul id="ulAdjuntosIT" class="list-group">
			</ul>
		
		</div>
		<div class="form-row" id="dvRegistrarKB" style="display: none;">
			<div class="form-group col-md-12">
				<label for="cmbRegistrarKB">¿Deseas guardar en la base de
					conocimientos?</label>
				<select name="cmbRegistrarKB" id="cmbRegistrarKB" class="form-control">
					<option value="No">No</option>
					<option value="Sí">Sí</option>
				</select>
			</div>
		</div>
		<div class="form-row" id="dvSolucionDetallada" style="display: none;">
			<div class="form-group col-md-12">
				<label for="txtSolucionDetallada">Detailed solution:</label>
				<textarea name="txtTitle" id="txtSolucionDetallada"
					class="form-control"></textarea>
			</div>
		</div>
		
		<div class="form-row">
			<div id="dvBotonera" class="form-group col-md-12"
				style="display: none; text-align: right; ">
				<button type="button" class="btn btn-primary"
					id="btnEnviar">Aprobar</button>
				<button type="button" class="btn btn-danger"
					id="btnCancelar">Rechazar</button>
				<button type="button" class="btn" id="btnCerrar">Cerrar</button>
			</div>
		</div>

		

	</div>
	
	</div>

</asp:Content>

