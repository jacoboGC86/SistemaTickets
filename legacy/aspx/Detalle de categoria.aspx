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
									Detalle de Categoría
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
									<link rel="stylesheet" href="../SiteAssets/css/alertify.min.css" />
									<link rel="stylesheet" type="text/css" href="../SiteAssets/css/Contratos.css" />
									<link rel="stylesheet"
										href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
									<link rel="stylesheet"
										href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-multiselect/0.9.15/css/bootstrap-multiselect.css">
									<link rel="stylesheet" href="path/to/jquery.multiselect.css">

									<script
										src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-multiselect/0.9.15/js/bootstrap-multiselect.min.js"></script>
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
									<script src="../SiteAssets/js/DetalleDeCategoria.js"></script>

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
												<h5 style="margin-top: 35px;text-align: center;">Detalle de Categoría
												</h5>
											</div>
											<div class="form-group col-md-4">
												<img src="../SiteAssets/img/Logo.png" title="Chimalli Apps"
													style="width: 150px;margin-top: 40px;" />
											</div>
										</div>
										<hr id="hrHead2" />


										<div class="form-row">
											<div class="form-group col-md-2">
												<label for="txtFolio">Folio:</label>
												<input type="text" class="form-control" id="txtFolio"
													disabled="disabled" title="Folio de registro de la solicitud" />
											</div>
											<div class="form-group col-md-8">
											</div>
											<div class="form-group col-md-2">
												<label for="dtFechaRegistro">Fecha de registro:</label>
												<input type="text" class="form-control" id="dtFechaRegistro"
													disabled="disabled" title="Fecha de registro de la solicitud" />
											</div>
										</div>

										<div class="form-row">
											<div class="form-group col-md-4">
												<label for="LblCategoria">Categoría:</label>
												<input type="text" class="form-control is-valid" id="txtCategoria" data-id="" disabled />
											</div>
											<div class="form-group col-md-4">
												<label for="LblPrioridad ">Prioridad:</label>
												<select name="cmbPrioridad" id="cmbPrioridad"
													class="form-control is-valid" disabled>
													<option value="none" selected disabled>Seleccione una opción
													</option>
													<option value="Alta">Alta</option>
													<option value="Media">Media</option>
													<option value="Baja">Baja</option>
												</select>
											</div>
											<div class="form-group col-md-4">
												<label for="LblTipoCat ">Tipo de categoría:</label>
												<select name="cmbTipoCat" id="cmbTipoCat" class="form-control is-valid"
													disabled>
													<option value="none" selected disabled>Seleccione una opción
													</option>
													<option value="Change">Cambio</option>
													<option value="Incident">Incidente</option>
													<option value="Request">Requeste</option>
												</select>
											</div>
										</div>

										<div class="form-row">
											<div class="form-group col-md-4">
												<label for="LblCatPadre">Categoría padre(Opcional):</label>
												<select name="cmbCatPadre" id="cmbCatPadre" class="form-control is-valid"
													disabled>
												</select>
											</div>
											<div class="form-group col-md-4">
												<label for="LblTemplates">Plantilla:</label>
												<select name="CmbTemplates" id="CmbTemplates"
													class="form-control is-valid" disabled>
												</select>
											</div>
											<div class="form-group col-md-4">
												<label for="LblGrupos">Grupos permitidos:</label><br>
												<select name="CmbGrupos" id="CmbGrupos" class="form-control" multiple></select>
											</div>
										</div>
										<div class="row">
											<div class="form-group col-md-3">
												<label for="txtSLALow">SLA Low (hours):</label>
												<input type="number" class="form-control is-valid" id="txtSLALow" data-id="" step="1" min="1" disabled />
											</div>
											<div class="form-group col-md-3">
												<label for="txtSLAMedium">SLA Medium (hours):</label>
												<input type="number" class="form-control is-valid" id="txtSLAMedium" data-id="" step="1" min="1" disabled />
											</div>
											<div class="form-group col-md-3">
												<label for="txtSLAHigh">SLA High (hours):</label>
												<input type="number" class="form-control is-valid" id="txtSLAHigh" data-id="" step="1" min="1" disabled />
											</div>
											<div class="form-group col-md-3">
												<label for="txtSLAStopProduction">Stop Production (hours):</label>
												<input type="number" class="form-control is-valid" id="txtSLAStopProduction" data-id="" step="1" min="1" disabled />
											</div>
										</div>
										<div class="form-row">
											<div class="form-group col-md-12">
												<label for="LblDescripcion">Descripción:</label>
												<textarea id="txtDescripcion" class="form-control is-valid"
													disabled></textarea>
											</div>
										</div>

										<div class="form-row">
											<div class="form-group col-md-12">
												<label for="LblProcedimientos">Procedimientos:</label>
												<textarea type="text" class="form-control is-valid"
													id="txtProcedimientos" data-id="" disabled></textarea>
												</select>
											</div>
										</div>


										<hr />
									</div>


									<div class="form-row">
										<div class="form-group col-md-8"></div>
										<div class="form-group col-md-4" style="text-align: right;">
											<button type="button" class="btn btn-success" id="btnUpdate">Enviar</button>
											<button type="button" class="btn btn-primary"
												id="btnEdit">Modificar</button>
											<button type="button" class="btn btn-secondary"
												id="btnCerrar">Cerrar</button>
										</div>
									</div>


									</div>

								</asp:Content>