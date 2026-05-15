import * as React from 'react';
import {
  Panel,
  PanelType,
  Stack,
  TextField,
  Dropdown,
  IDropdownOption,
  Checkbox,
  PrimaryButton,
  DefaultButton,
  IStackTokens,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Label,
  Callout,
  DirectionalHint,
  NormalPeoplePicker,
  IPersonaProps,
} from '@fluentui/react';
import ListSvc from '../../../services/ListSvc';
import UserSvc from '../../../services/UserSvc';
import EvaluacionRiesgo, { IEvaluacionRiesgoValues } from './EvaluacionRiesgo';

// ─── Types ────────────────────────────────────────────────────────────────────

interface ICategoria {
  Id: number;
  Title: string;
  Tipo: string;
  Prioridad: string;
  IdTemplate: number;
}

interface IDepartamento {
  Valor: string;
  PersonaId: number;
}

interface IFormField {
  tag: string;
  id?: string;
  text?: string;
  checked?: boolean;
  value?: string;
}

interface ITemplate {
  approvalPath: { stepName: string; approvalGroup?: string | null }[];
  formTemplate?: IFormField[];
  processManager: { id: number };
  processManagerAguascalientes?: { id: number };
}

// ─── Props ────────────────────────────────────────────────────────────────────

export interface INuevoTicketProps {
  isOpen: boolean;
  onDismiss: () => void;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target!.result as ArrayBuffer);
    reader.onerror = (e) => reject(e);
    reader.readAsArrayBuffer(file);
  });
}

const stackTokens: IStackTokens = { childrenGap: 16 };

const typeOptions: IDropdownOption[] = [
  { key: 'Request', text: 'Solicitud' },
  { key: 'Incident', text: 'Incidente' },
  { key: 'Change', text: 'Cambio' },
];

const plantOptions: IDropdownOption[] = [
  { key: 'Aguascalientes', text: 'Aguascalientes' },
  { key: 'Puebla', text: 'Puebla' },
];

const urgencyOptions: IDropdownOption[] = [
  { key: 'Alta', text: 'Alta' },
  { key: 'Media', text: 'Media' },
  { key: 'Baja', text: 'Baja' },
];

// ─── Component ────────────────────────────────────────────────────────────────

const NuevoTicket: React.FC<INuevoTicketProps> = ({ isOpen, onDismiss }) => {
  // Data state
  const [departamentos, setDepartamentos] = React.useState<IDepartamento[]>([]);
  const [categorias, setCategorias] = React.useState<ICategoria[]>([]);
  const [filteredCategorias, setFilteredCategorias] = React.useState<ICategoria[]>([]);
  const [template, setTemplate] = React.useState<ITemplate | null>(null);
  const [formFields, setFormFields] = React.useState<IFormField[]>([]);
  const [formFieldValues, setFormFieldValues] = React.useState<Record<string, string | boolean>>({});
  const [isLoading, setIsLoading] = React.useState(false);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [errorMessage, setErrorMessage] = React.useState('');
  const [successMessage, setSuccessMessage] = React.useState('');
  const dataLoaded = React.useRef(false);

  // Form state
  const [ticketType, setTicketType] = React.useState('');
  const [categoryId, setCategoryId] = React.useState<number | null>(null);
  const [title, setTitle] = React.useState('');
  const [description, setDescription] = React.useState('');
  const [selectedDept, setSelectedDept] = React.useState<IDepartamento | null>(null);
  const [plant, setPlant] = React.useState('Aguascalientes');
  const [urgency, setUrgency] = React.useState('');
  const [registerForOther, setRegisterForOther] = React.useState(false);
  const [selectedPerson, setSelectedPerson] = React.useState<IPersonaProps | null>(null);
  const [securityImpact, setSecurityImpact] = React.useState('No');
  const [securityDescription, setSecurityDescription] = React.useState('');
  const [selectedFiles, setSelectedFiles] = React.useState<File[]>([]);
  const [categoryInputText, setCategoryInputText] = React.useState('');
  const [categoryMenuOpen, setCategoryMenuOpen] = React.useState(false);
  const [evaluacionRiesgo, setEvaluacionRiesgo] = React.useState<IEvaluacionRiesgoValues | null>(null);
  const [evaluacionRiesgoKey, setEvaluacionRiesgoKey] = React.useState(0);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const categoryInputWrapperRef = React.useRef<HTMLDivElement>(null);

  // ── Lifecycle ──────────────────────────────────────────────────────────────

  React.useEffect(() => {
    if (isOpen && !dataLoaded.current) {
      loadData().catch(console.error);
    }
    if (!isOpen) {
      resetForm();
    }
  }, [isOpen]);

  const resetForm = (): void => {
    setTicketType('');
    setCategoryId(null);
    setTitle('');
    setDescription('');
    setSelectedDept(null);
    setPlant('Aguascalientes');
    setUrgency('');
    setRegisterForOther(false);
    setSelectedPerson(null);
    setSecurityImpact('No');
    setSecurityDescription('');
    setSelectedFiles([]);
    setFilteredCategorias([]);
    setCategoryInputText('');
    setCategoryMenuOpen(false);
    setTemplate(null);
    setFormFields([]);
    setFormFieldValues({});
    setErrorMessage('');
    setSuccessMessage('');
    setEvaluacionRiesgo(null);
    setEvaluacionRiesgoKey(prev => prev + 1);
    dataLoaded.current = false;
  };

  // ── Data loading ───────────────────────────────────────────────────────────

  const loadData = async (): Promise<void> => {
    setIsLoading(true);
    try {
      const deptos = await loadDepartamentos();
      await loadUserProfile(deptos);
      await loadCategorias();
      dataLoaded.current = true;
    } catch (e) {
      console.error('Error loading form data:', e);
      setErrorMessage('Error al cargar los datos del formulario.');
    } finally {
      setIsLoading(false);
    }
  };

  const loadDepartamentos = async (): Promise<IDepartamento[]> => {
    const items: any[] = await ListSvc.getItems(
      'Configuraciones',
      undefined,
      "$filter=Title eq 'Departamento'&$select=Id,Valor,Persona/Id,Persona/Title&$expand=Persona&$orderby=Valor&$top=300"
    );
    const deptos: IDepartamento[] = (items || []).map((item: any) => ({
      Valor: item.Valor,
      PersonaId: item.Persona?.Id ?? 0,
    }));
    setDepartamentos(deptos);
    return deptos;
  };

  const loadUserProfile = async (deptos: IDepartamento[]): Promise<void> => {
    try {
      const profile = await UserSvc.GetPropertiesUser();
      const props: any[] =
        profile?.d?.UserProfileProperties?.results ??
        profile?.UserProfileProperties ??
        [];

      let department = '';
      let office = '';
      for (const p of props) {
        if (p.Key === 'Department') department = p.Value;
        if (p.Key === 'Office') office = p.Value;
      }

      if (office) setPlant(office);

      if (department) {
        const needsSuffix =
          office === 'Aguascalientes' &&
          (department === 'Equipment & Building Maintenance' ||
            department === 'Logistics' ||
            department === 'Production Assembly');
        const deptValue = needsSuffix ? `${department} ${office}` : department;
        const match = deptos.find(d => d.Valor === deptValue);
        if (match) setSelectedDept(match);
      }
    } catch (e) {
      console.error('Error loading user profile:', e);
    }
  };

  const loadCategorias = async (): Promise<void> => {
    const items: any[] = await ListSvc.getItems(
      'Categorias',
      undefined,
      '$select=Id,Title,TipoCategoria,Prioridad,CategoriaPadreId,CategoriaPadre/Title,TemplateAtencionId&$expand=CategoriaPadre&$top=300'
    );
    const cats: ICategoria[] = (items || [])
      .filter((item: any) => item.CategoriaPadreId != null)
      .map((item: any) => ({
        Id: item.Id,
        Title: `${item.CategoriaPadre?.Title ?? ''} - ${item.Title}`,
        Tipo: item.TipoCategoria,
        Prioridad: item.Prioridad,
        IdTemplate: item.TemplateAtencionId ?? 0,
      }))
      .sort((a: ICategoria, b: ICategoria) => a.Title.localeCompare(b.Title));
    setCategorias(cats);
  };

  // ── Event handlers ─────────────────────────────────────────────────────────

  const onTypeChange = (_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    const type = (option?.key as string) || '';
    setTicketType(type);
    setCategoryId(null);
    setCategoryInputText('');
    setCategoryMenuOpen(false);
    setTemplate(null);
    setFormFields([]);
    setFormFieldValues({});
    setFilteredCategorias(categorias.filter(c => c.Tipo === type));
  };

  const onCategorySelect = (cat: ICategoria): void => {
    setCategoryInputText(cat.Title);
    setCategoryMenuOpen(false);
    onCategoryChange(cat.Id).catch(console.error);
  };

  const onCategoryChange = async (id: number): Promise<void> => {
    //const id = option?.key as number;
    setCategoryId(id);
    setFormFields([]);
    setFormFieldValues({});
    setTemplate(null);

    const cat = categorias.find((c: ICategoria) => c.Id === id);
    if (!cat || !cat.IdTemplate) return;

    setUrgency(cat.Prioridad);
    try {
      const templateItem = await ListSvc.getItemById('Templates', cat.IdTemplate);
      const raw: string | undefined =
        templateItem?.RutaAprobacionEscalacion ??
        templateItem?.d?.RutaAprobacionEscalacion;
      if (!raw) return;

      const t: ITemplate = JSON.parse(raw);
      setTemplate(t);
      setFormFields(t.formTemplate || []);

      const initial: Record<string, string | boolean> = {};
      (t.formTemplate || []).forEach(field => {
        if (field.id) {
          initial[field.id] = field.tag === 'checkbox' ? false : '';
        }
      });
      setFormFieldValues(initial);
    } catch (e) {
      console.error('Error loading template:', e);
      setErrorMessage('Error al cargar el template de la categoría seleccionada.');
    }
  };

  const onFilesSelected = (ev: React.ChangeEvent<HTMLInputElement>): void => {
    if (!ev.target.files || ev.target.files.length === 0) return;
    const newFiles = Array.from(ev.target.files);
    if (selectedFiles.length + newFiles.length > 2) {
      setErrorMessage('Solo puedes subir un máximo de 2 archivos.');
      ev.target.value = '';
      return;
    }
    setSelectedFiles(prev => [...prev, ...newFiles]);
    ev.target.value = '';
  };

  const removeFile = (index: number): void => {
    setSelectedFiles(prev => prev.filter((_, i) => i !== index));
  };

  const handleSubmit = async (): Promise<void> => {
    setErrorMessage('');
    setSuccessMessage('');

    if (!ticketType || !categoryId || !title || !description || !urgency || !selectedDept) {
      setErrorMessage('Por favor complete todos los campos requeridos.');
      return;
    }

    if (!template) {
      setErrorMessage('No se pudo cargar el template de la categoría. Intente seleccionar la categoría de nuevo.');
      return;
    }

    setIsSubmitting(true);
    try {
      // Build template copy with captured form field values
      const t: ITemplate = JSON.parse(JSON.stringify(template));
      if (t.formTemplate) {
        t.formTemplate = t.formTemplate.map(field => {
          if (!field.id) return field;
          if (field.tag === 'checkbox') return { ...field, checked: !!formFieldValues[field.id] };
          if (field.tag === 'textarea') return { ...field, value: (formFieldValues[field.id] as string) || '' };
          return field;
        });
      }

      const status = t.approvalPath.length > 0 ? t.approvalPath[0].stepName : 'Assigned';

      // Resolve ANombreDe user if registering for someone else
      let aNombreDeId: number | null = null;
      if (registerForOther && selectedPerson?.key) {
        const user = await UserSvc.EnsureUser(selectedPerson.key as string);
        aNombreDeId = user?.Id ?? null;
        if (!aNombreDeId) {
          setErrorMessage('No se pudo resolver el usuario seleccionado.');
          setIsSubmitting(false);
          return;
        }
      }

      const processManagerId: number =
        plant === 'Puebla'
          ? t.processManager.id
          : (t.processManagerAguascalientes?.id ?? t.processManager.id);

      // Get entity type names
      const [ticketsEntityType, approvalsEntityType, evaluacionesEntityType] = await Promise.all([
        ListSvc.getListItemEntityTypeFullName('Tickets'),
        ListSvc.getListItemEntityTypeFullName('Aprobaciones'),
        ListSvc.getListItemEntityTypeFullName('EvaluacionesRiesgo'),
      ]);

      // Build and create ticket item
      const itemBody: Record<string, unknown> = {
        '__metadata': { 'type': ticketsEntityType },
        'Title': title,
        'Descripcion': description,
        'TipoTicket': ticketType,
        'TemplateConfiguracion': JSON.stringify(t),
        'Prioridad': urgency,
        'Status': status,
        'CategoriaId': categoryId,
        'ImpactoSeguridadInformacion': ticketType === 'Change' ? securityImpact : 'No',
        'DescripcionImpactoSeguridad': ticketType === 'Change' && securityImpact === 'Sí' ? securityDescription : '',
        'Department': selectedDept.Valor,
        'ManagerId': selectedDept.PersonaId,
        'Planta': plant,
        'ProcessManagerId': processManagerId,
        'ANombreDeId': aNombreDeId
      };

      const newItem = await ListSvc.postListItem('Tickets', JSON.stringify(itemBody));
      const idSol: number = newItem?.d?.Id ?? newItem?.Id;

      // Create folder in Expediente library
      const relativeURL = ListSvc.getRelativeSiteURL();
      const folderBody = JSON.stringify({
        '__metadata': { 'type': 'SP.Folder' },
        'ServerRelativeUrl': `${relativeURL}/Expediente/${idSol}`,
      });
      await ListSvc.postFolder('', folderBody);

      // Upload attachments
      for (const file of selectedFiles) {
        const arrayBuffer = await readFileAsArrayBuffer(file);
        await ListSvc.postFile(`${relativeURL}/Expediente/${idSol}`, file.name, arrayBuffer);
      }

      // Create approval record
      let approvalBody: Record<string, unknown>;
      if (status !== 'Assigned') {
        approvalBody = {
          '__metadata': { 'type': approvalsEntityType },
          'Title': status,
          'Resultado': 'Pendiente',
          'TicketId': idSol,
        };
        if (t.approvalPath[0].approvalGroup != null) {
          approvalBody['GrupoId'] = t.approvalPath[0].approvalGroup;
        } else {
          approvalBody['ResponsableId'] = selectedDept.PersonaId;
        }
      } else {
        approvalBody = {
          '__metadata': { 'type': approvalsEntityType },
          'Title': 'Atención de ticket',
          'Resultado': 'Pendiente',
          'TicketId': idSol,
          'ResponsableId': processManagerId,
        };
      }
      let aprobacion:any = await ListSvc.postListItem('Aprobaciones', JSON.stringify(approvalBody));

      console.log(aprobacion);

      // Create EvaluacionesRiesgo record for Change tickets
      if (ticketType === 'Change' && evaluacionRiesgo) {
        let fechaPropuestaISO: string | null = null;
        if (evaluacionRiesgo.fechaPropuesta) {
          const base = new Date(evaluacionRiesgo.fechaPropuesta);
          if (evaluacionRiesgo.horaPropuesta) {
            base.setHours(
              evaluacionRiesgo.horaPropuesta.getHours(),
              evaluacionRiesgo.horaPropuesta.getMinutes(),
              0,
              0
            );
          }
          fechaPropuestaISO = base.toISOString();
        }
        const evaluacionBody: Record<string, unknown> = {
          '__metadata': { 'type': evaluacionesEntityType },
          'Title': `Evaluación - Ticket ${idSol}`,
          'TicketId': idSol,
          'TipoCambio': evaluacionRiesgo.tipoCambio,
          'SistemasInvolucrados': evaluacionRiesgo.sistemasInvolucrados,
          'FechaPropuesta': fechaPropuestaISO,
          'EvaluacionRiesgos': evaluacionRiesgo.evaluacionRiesgos,
          'InterrupcionServicio': evaluacionRiesgo.interrupcionServicio,
          'DuracionEstimada': evaluacionRiesgo.duracionEstimada ? Number(evaluacionRiesgo.duracionEstimada) : null,
          'UsuariosAfectados': evaluacionRiesgo.usuariosAfectados ? Number(evaluacionRiesgo.usuariosAfectados) : null,
          'PasosImplementacion': evaluacionRiesgo.pasosImplementacion,
          'PuedoRevertir': evaluacionRiesgo.puedeRevertir,
          'ProcedimientoRollback': evaluacionRiesgo.procedimientoRollback,
          'CriteriosExito': evaluacionRiesgo.criteriosExito,
          'MotivoCambio': evaluacionRiesgo.motivoCambio,
          'ProcesosCriticosImpactados': evaluacionRiesgo.procesosCriticosImpactados,
        };
        await ListSvc.postListItem('EvaluacionesRiesgo', JSON.stringify(evaluacionBody));
      }

      setSuccessMessage('Solicitud realizada correctamente.');
      setTimeout(() => {
        resetForm();
        onDismiss();
      }, 2000);
    } catch (e) {
      console.error('Error al enviar el formulario:', e);
      setErrorMessage('Error al enviar la solicitud. Por favor, intente de nuevo.');
    } finally {
      setIsSubmitting(false);
    }
  };

  // ── Derived dropdown options ────────────────────────────────────────────────

  const departmentOptions: IDropdownOption[] = departamentos.map(d => ({
    key: d.Valor,
    text: d.Valor,
  }));

  const categoryOptionsFiltered: ICategoria[] = filteredCategorias
    .filter((c: ICategoria) => !categoryInputText || c.Title.toLowerCase().indexOf(categoryInputText.toLowerCase()) !== -1);

  // ── Render ─────────────────────────────────────────────────────────────────

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      headerText="Nuevo Ticket"
      closeButtonAriaLabel="Cerrar"
      type={PanelType.large}
    >
      {isLoading ? (
        <Spinner size={SpinnerSize.large} label="Cargando datos..." />
      ) : (
        <Stack tokens={stackTokens}>
          {errorMessage && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMessage('')}>
              {errorMessage}
            </MessageBar>
          )}
          {successMessage && (
            <MessageBar messageBarType={MessageBarType.success}>
              {successMessage}
            </MessageBar>
          )}

          <TextField
            label="Fecha de registro"
            value={new Date().toLocaleDateString('es-MX')}
            disabled
            styles={{ root: { maxWidth: 260 } }}
          />

          <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
            <Dropdown
              label="Tipo"
              placeholder="Seleccione una opción"
              options={typeOptions}
              selectedKey={ticketType || null}
              onChange={onTypeChange}
              styles={{ root: { minWidth: 260 } }}
              required
            />
            <div ref={categoryInputWrapperRef} style={{ minWidth: 260 }}>
              <TextField
                label="Categoría"
                placeholder={ticketType ? 'Buscar categoría...' : 'Seleccione primero un tipo'}
                value={categoryInputText}
                disabled={!ticketType}
                required
                onChange={(_, v) => {
                  setCategoryInputText(v || '');
                  setCategoryId(null);
                  setTemplate(null);
                  setFormFields([]);
                  setFormFieldValues({});
                  setCategoryMenuOpen(true);
                }}
                onFocus={() => { if (ticketType) setCategoryMenuOpen(true); }}
                autoComplete="off"
                onRenderSuffix={() => categoryInputText ? (
                  <span
                    title="Limpiar selección"
                    onMouseDown={(e) => {
                      e.preventDefault();
                      setCategoryInputText('');
                      setCategoryId(null);
                      setTemplate(null);
                      setFormFields([]);
                      setFormFieldValues({});
                      setCategoryMenuOpen(true);
                    }}
                    style={{ cursor: 'pointer', padding: '0 4px', lineHeight: 1, fontSize: 12, color: '#605e5c' }}
                  >✕</span>
                ) : null}
              />
            </div>
            {categoryMenuOpen && categoryOptionsFiltered.length > 0 && categoryInputWrapperRef.current && (
              <Callout
                target={categoryInputWrapperRef.current}
                onDismiss={() => setCategoryMenuOpen(false)}
                isBeakVisible={false}
                directionalHint={DirectionalHint.bottomLeftEdge}
                calloutMaxHeight={280}
                styles={{ calloutMain: { overflowY: 'auto' } }}
                setInitialFocus={false}
              >
                {categoryOptionsFiltered.map((c: ICategoria) => (
                  <div
                    key={c.Id}
                    onMouseDown={(e) => { e.preventDefault(); onCategorySelect(c); }}
                    style={{ padding: '8px 16px', cursor: 'pointer', whiteSpace: 'nowrap' }}
                    onMouseEnter={(e) => { (e.currentTarget as HTMLDivElement).style.background = '#f3f2f1'; }}
                    onMouseLeave={(e) => { (e.currentTarget as HTMLDivElement).style.background = ''; }}
                  >
                    {c.Title}
                  </div>
                ))}
              </Callout>
            )}
          </Stack>

          <TextField
            label="Título"
            value={title}
            onChange={(_, v) => setTitle(v || '')}
            required
          />

          <TextField
            label="Descripción detallada"
            value={description}
            onChange={(_, v) => setDescription(v || '')}
            multiline
            rows={4}
            required
          />

          <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
            <Dropdown
              label="Departamento"
              placeholder="Seleccione un departamento"
              options={departmentOptions}
              selectedKey={selectedDept?.Valor ?? null}
              onChange={(_, option) => {
                const dept = departamentos.find(d => d.Valor === (option?.key as string));
                setSelectedDept(dept ?? null);
              }}
              styles={{ root: { minWidth: 260 } }}
              required
            />
            <Dropdown
              label="Planta"
              options={plantOptions}
              selectedKey={plant}
              onChange={(_, option) => setPlant((option?.key as string) || 'Aguascalientes')}
              styles={{ root: { minWidth: 260 } }}
            />
          </Stack>

          <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
            <Dropdown
              label="Urgencia"
              placeholder="Seleccione una urgencia"
              options={urgencyOptions}
              selectedKey={urgency || null}
              onChange={(_, option) => setUrgency((option?.key as string) || '')}
              styles={{ root: { minWidth: 260 } }}
              required
            />
            <Checkbox
              label="¿Estás registrando este ticket a nombre de otra persona?"
              checked={registerForOther}
              onChange={(_, checked) => setRegisterForOther(!!checked)}
              styles={{ root: { minWidth: 260, marginTop: 40 } }}
            />
          </Stack>

          {registerForOther && (
            <div>
              <Label required>A nombre de</Label>
              <NormalPeoplePicker
                onResolveSuggestions={async (filter: string): Promise<IPersonaProps[]> => {
                  if (!filter || filter.length < 2) return [];
                  try {
                    const results = await UserSvc.SearchUsers(filter);
                    return (results || []).map((r: any) => ({
                      key: r.Key,
                      text: r.DisplayText,
                      secondaryText: r.EntityData && r.EntityData.Email ? r.EntityData.Email : r.Description,
                    }));
                  } catch {
                    return [];
                  }
                }}
                onChange={(items?: IPersonaProps[]) => {
                  setSelectedPerson(items && items.length > 0 ? items[0] : null);
                }}
                selectedItems={selectedPerson ? [selectedPerson] : []}
                itemLimit={1}
                resolveDelay={300}
                pickerSuggestionsProps={{
                  suggestionsHeaderText: 'Usuarios sugeridos',
                  noResultsFoundText: 'No se encontraron resultados',
                  loadingText: 'Buscando...',
                }}
                inputProps={{ placeholder: 'Buscar por nombre o correo...' }}
              />
            </div>
          )}

          {ticketType === 'Change' && (
            <Stack tokens={{ childrenGap: 12 }}>
              <Dropdown
                label="¿Este cambio impacta la seguridad de la información?"
                options={[
                  { key: 'No', text: 'No' },
                  { key: 'Sí', text: 'Sí' },
                ]}
                selectedKey={securityImpact}
                onChange={(_, option) => setSecurityImpact((option?.key as string) || 'No')}
                styles={{ root: { minWidth: 260 } }}
              />
              {securityImpact === 'Sí' && (
                <TextField
                  label="Describe los impactos a la seguridad de la información"
                  value={securityDescription}
                  onChange={(_, v) => setSecurityDescription(v || '')}
                  multiline
                  rows={3}
                />
              )}
            </Stack>
          )}

          {/* Dynamic form fields from template */}
          {formFields.length > 0 && (
            <Stack tokens={{ childrenGap: 8 }}>
              {formFields.map((field, idx) => {
                switch (field.tag) {
                  case 'h3':
                    return <h3 key={idx}>{field.text}</h3>;
                  case 'p':
                    return <p key={idx}>{field.text}</p>;
                  case 'br':
                    return <div key={idx} style={{ marginTop: 8 }} />;
                  case 'checkbox':
                    return (
                      <Checkbox
                        key={idx}
                        label={field.text}
                        checked={!!formFieldValues[field.id!]}
                        onChange={(_, checked) =>
                          setFormFieldValues(prev => ({ ...prev, [field.id!]: !!checked }))
                        }
                      />
                    );
                  case 'textarea':
                    return (
                      <TextField
                        key={idx}
                        label={field.text}
                        value={(formFieldValues[field.id!] as string) || ''}
                        onChange={(_, v) =>
                          setFormFieldValues(prev => ({ ...prev, [field.id!]: v || '' }))
                        }
                        multiline
                        rows={3}
                      />
                    );
                  default:
                    return null;
                }
              })}
            </Stack>
          )}

          {/* Evaluación de Riesgo */}
          {ticketType === 'Change' && (
            <EvaluacionRiesgo
              key={evaluacionRiesgoKey}
              onChange={setEvaluacionRiesgo}
            />
          )}

          {/* File attachments */}
          <Stack tokens={{ childrenGap: 8 }}>
            <Label>Archivos adjuntos (máximo 2)</Label>
            <DefaultButton
              text="Cargar archivos"
              onClick={() => fileInputRef.current?.click()}
              disabled={selectedFiles.length >= 2}
            />
            <input
              type="file"
              multiple
              hidden
              ref={fileInputRef}
              onChange={onFilesSelected}
            />
            {selectedFiles.length > 0 && (
              <Stack tokens={{ childrenGap: 4 }}>
                {selectedFiles.map((file, index) => (
                  <Stack key={index} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    <span>{file.name}</span>
                    <DefaultButton
                      text="Eliminar"
                      onClick={() => removeFile(index)}
                      styles={{ root: { minWidth: 'auto', padding: '0 8px', height: 24 } }}
                    />
                  </Stack>
                ))}
              </Stack>
            )}
          </Stack>

          <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 12 }}>
            <PrimaryButton
              text={isSubmitting ? 'Enviando...' : 'Enviar'}
              onClick={handleSubmit}
              disabled={isSubmitting}
            />
            <DefaultButton text="Cancelar" onClick={onDismiss} disabled={isSubmitting} />
          </Stack>
        </Stack>
      )}
    </Panel>
  );
};

export default NuevoTicket;

