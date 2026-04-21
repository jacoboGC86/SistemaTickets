import * as React from 'react';
import {
  Panel,
  PanelType,
  Stack,
  TextField,
  Dropdown,
  IDropdownOption,
  ComboBox,
  IComboBoxOption,
  IComboBox,
  PrimaryButton,
  DefaultButton,
  IStackTokens,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Checkbox,
  Label,
} from '@fluentui/react';
import ListSvc from '../../../services/ListSvc';

// ─── Types ────────────────────────────────────────────────────────────────────

interface ICategoria {
  id: number;
  title: string;
}

interface IGrupo {
  id: number;
  name: string;
}

// ─── Props ────────────────────────────────────────────────────────────────────

export interface INuevaCategoriaProps {
  isOpen: boolean;
  onDismiss: () => void;
}

// ─── Constants ────────────────────────────────────────────────────────────────

const stackTokens: IStackTokens = { childrenGap: 12 };

const prioridadOptions: IDropdownOption[] = [
  { key: 'Alta', text: 'Alta' },
  { key: 'Media', text: 'Media' },
  { key: 'Baja', text: 'Baja' },
];

const tipoCatOptions: IDropdownOption[] = [
  { key: 'Change', text: 'Change' },
  { key: 'Incident', text: 'Incident' },
  { key: 'Request', text: 'Request' },
];

// ─── Component ────────────────────────────────────────────────────────────────

const NuevaCategoria: React.FC<INuevaCategoriaProps> = ({ isOpen, onDismiss }) => {

  // ── Form values ────────────────────────────────────────────────────────────
  const [categoria, setCategoria]           = React.useState('');
  const [descripcion, setDescripcion]       = React.useState('');
  const [procedimiento, setProcedimiento]   = React.useState('');
  const [prioridad, setPrioridad]           = React.useState<string | null>(null);
  const [tipoCat, setTipoCat]               = React.useState<string | null>(null);
  const [templateId, setTemplateId]         = React.useState<number | null>(null);
  const [catPadreId, setCatPadreId]         = React.useState<number | null>(null);
  const [gruposSeleccionados, setGruposSeleccionados] = React.useState<number[]>([]);
  const [slaLow, setSlaLow]                 = React.useState('');
  const [slaMedium, setSlaMedium]           = React.useState('');
  const [slaHigh, setSlaHigh]               = React.useState('');
  const [slaStopProduction, setSlaStopProduction] = React.useState('');

  // ── Loaded data ────────────────────────────────────────────────────────────
  const [templateOptions, setTemplateOptions]     = React.useState<IDropdownOption[]>([]);
  const [catPadreOptions, setCatPadreOptions]     = React.useState<IComboBoxOption[]>([]);
  const [grupos, setGrupos]                       = React.useState<IGrupo[]>([]);

  // ── UI state ───────────────────────────────────────────────────────────────
  const [isLoading, setIsLoading]           = React.useState(false);
  const [isSubmitting, setIsSubmitting]     = React.useState(false);
  const [errorMessage, setErrorMessage]     = React.useState('');
  const [successMessage, setSuccessMessage] = React.useState('');
  const dataLoaded = React.useRef(false);

  // ── Lifecycle ──────────────────────────────────────────────────────────────

  React.useEffect(() => {
    if (isOpen && !dataLoaded.current) {
      loadData().catch(console.error);
    }
    if (!isOpen) resetForm();
  }, [isOpen]);

  React.useEffect(() => {
    if (!isOpen) return;
    loadCatPadre().catch(console.error);
  }, [tipoCat, isOpen]);

  const resetForm = (): void => {
    setCategoria('');
    setDescripcion('');
    setProcedimiento('');
    setPrioridad(null);
    setTipoCat(null);
    setTemplateId(null);
    setCatPadreId(null);
    setGruposSeleccionados([]);
    setSlaLow('');
    setSlaMedium('');
    setSlaHigh('');
    setSlaStopProduction('');
    setErrorMessage('');
    setSuccessMessage('');
    setCatPadreOptions([]);
    dataLoaded.current = false;
  };

  // ── Data loading ───────────────────────────────────────────────────────────

  const loadData = async (): Promise<void> => {
    setIsLoading(true);
    try {
      await Promise.all([loadTemplates(), loadGrupos()]);
      dataLoaded.current = true;
    } catch (e) {
      console.error('Error loading data:', e);
      setErrorMessage('Error al cargar los datos del formulario.');
    } finally {
      setIsLoading(false);
    }
  };

  const loadTemplates = async (): Promise<void> => {
    const items: any[] = await ListSvc.getItems(
      'Templates',
      undefined,
      '$orderby=Title asc&$top=100&$select=Id,Title'
    );
    setTemplateOptions(
      (items || []).map((i: any) => ({ key: i.Id, text: i.Title }))
    );
  };

  const loadGrupos = async (): Promise<void> => {
    const items: any[] = await ListSvc.getItems(
      'Grupos',
      undefined,
      '$orderby=Title asc&$top=100&$select=Id,Title'
    );
    setGrupos((items || []).map((i: any) => ({ id: i.Id, name: i.Title })));
  };

  const loadCatPadre = async (): Promise<void> => {
    setCatPadreId(null);
    setCatPadreOptions([]);
    const items: any[] = await ListSvc.getItems(
      'Categorias',
      undefined,
      "$filter=CategoriaPadreId eq null&$orderby=Title asc&$top=500&$select=Id,Title"
    );
    setCatPadreOptions(
      (items || []).map((i: any) => ({ key: i.Id, text: i.Title } as IComboBoxOption))
    );
  };

  // ── Grupos helpers ─────────────────────────────────────────────────────────

  const toggleGrupo = (id: number, checked: boolean): void => {
    setGruposSeleccionados(prev =>
      checked ? [...prev, id] : prev.filter(g => g !== id)
    );
  };

  // ── Submit ─────────────────────────────────────────────────────────────────

  const handleSubmit = async (): Promise<void> => {
    setErrorMessage('');

    if (!categoria.trim())    { setErrorMessage('Por favor ingrese el nombre de la categoría.'); return; }
    if (!prioridad)           { setErrorMessage('Por favor seleccione la prioridad.'); return; }
    if (!tipoCat)             { setErrorMessage('Por favor seleccione el tipo de categoría.'); return; }
    if (!templateId)          { setErrorMessage('Por favor seleccione una plantilla.'); return; }
    if (!descripcion.trim())  { setErrorMessage('Por favor ingrese la descripción.'); return; }
    if (!procedimiento.trim()){ setErrorMessage('Por favor ingrese el procedimiento de atención.'); return; }
    if (!slaLow || Number(slaLow) < 1)             { setErrorMessage('Por favor ingrese un SLA Low válido.'); return; }
    if (!slaMedium || Number(slaMedium) < 1)       { setErrorMessage('Por favor ingrese un SLA Medium válido.'); return; }
    if (!slaHigh || Number(slaHigh) < 1)           { setErrorMessage('Por favor ingrese un SLA High válido.'); return; }
    if (!slaStopProduction || Number(slaStopProduction) < 1) { setErrorMessage('Por favor ingrese un SLA Stop Production válido.'); return; }

    setIsSubmitting(true);
    try {
      const entityType = await ListSvc.getListItemEntityTypeFullName('Categorias');

      const bodyObj: Record<string, any> = {
        '__metadata': { type: entityType },
        'Title': categoria.trim(),
        'Descripcion': descripcion.trim(),
        'Prioridad': prioridad,
        'TipoCategoria': tipoCat,
        'TemplateAtencionId': templateId,
        'ProcedimientoAtencion': procedimiento.trim(),
        'SLALow': Number(slaLow),
        'SLAMedium': Number(slaMedium),
        'SLAHigh': Number(slaHigh),
        'SLAProductionStop': Number(slaStopProduction),
      };

      if (catPadreId) {
        bodyObj['CategoriaPadreId'] = catPadreId;
      }

      if (gruposSeleccionados.length > 0) {
        bodyObj['GruposPermitidosId'] = { results: gruposSeleccionados };
      }

      await ListSvc.postListItem('Categorias', JSON.stringify(bodyObj));
      setSuccessMessage('Categoría guardada exitosamente.');
      setTimeout(() => { resetForm(); onDismiss(); }, 2000);
    } catch (e) {
      console.error('Error al guardar categoría:', e);
      setErrorMessage('Error al guardar la categoría. Por favor, intente de nuevo.');
    } finally {
      setIsSubmitting(false);
    }
  };

  // ── Render ─────────────────────────────────────────────────────────────────

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.large}
      headerText="Nueva categoría"
      isFooterAtBottom
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
          <PrimaryButton text="Guardar" onClick={handleSubmit} disabled={isSubmitting || isLoading} />
          <DefaultButton text="Cancelar" onClick={onDismiss} disabled={isSubmitting} />
          {isSubmitting && <Spinner size={SpinnerSize.small} />}
        </Stack>
      )}
    >
      {isLoading ? (
        <Spinner label="Cargando datos..." size={SpinnerSize.large} styles={{ root: { marginTop: 40 } }} />
      ) : (
        <Stack tokens={stackTokens} styles={{ root: { paddingTop: 16 } }}>

          {errorMessage && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMessage('')} isMultiline={false}>
              {errorMessage}
            </MessageBar>
          )}
          {successMessage && (
            <MessageBar messageBarType={MessageBarType.success}>{successMessage}</MessageBar>
          )}

          {/* ── Row 1: Categoría, Prioridad, Tipo ─────────────────────── */}
          <Stack horizontal tokens={{ childrenGap: 16 }} styles={{ root: { flexWrap: 'wrap' } }}>
            <Stack.Item grow={1} styles={{ root: { minWidth: 200 } }}>
              <TextField
                label="Categoría"
                required
                value={categoria}
                onChange={(_, v) => setCategoria(v ?? '')}
              />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 160 } }}>
              <Dropdown
                label="Prioridad"
                required
                options={prioridadOptions}
                selectedKey={prioridad}
                onChange={(_, o) => setPrioridad(o ? String(o.key) : null)}
                placeholder="Seleccione una opción"
              />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 160 } }}>
              <Dropdown
                label="Tipo de categoría"
                required
                options={tipoCatOptions}
                selectedKey={tipoCat}
                onChange={(_, o) => setTipoCat(o ? String(o.key) : null)}
                placeholder="Seleccione una opción"
              />
            </Stack.Item>
          </Stack>

          {/* ── Row 2: CatPadre, Template, Grupos ────────────────────── */}
          <Stack horizontal tokens={{ childrenGap: 16 }} styles={{ root: { flexWrap: 'wrap', alignItems: 'flex-start' } }}>
            <Stack.Item grow={1} styles={{ root: { minWidth: 200 } }}>
              <ComboBox
                label="Categoría padre (opcional)"
                options={catPadreOptions}
                selectedKey={catPadreId}
                onChange={(_: React.FormEvent<IComboBox>, o?: IComboBoxOption) => setCatPadreId(o ? Number(o.key) : null)}
                placeholder="Buscar o seleccionar..."
                allowFreeInput={false}
                autoComplete="on"
                useComboBoxAsMenuWidth
              />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 200 } }}>
              <Dropdown
                label="Plantilla"
                required
                options={templateOptions}
                selectedKey={templateId}
                onChange={(_, o) => setTemplateId(o ? Number(o.key) : null)}
                placeholder="Seleccione una opción"
              />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 200 } }}>
              <Label>Grupos permitidos</Label>
              <Stack tokens={{ childrenGap: 6 }} styles={{ root: { maxHeight: 160, overflowY: 'auto', border: '1px solid #edebe9', borderRadius: 2, padding: '6px 8px' } }}>
                {grupos.length === 0 ? (
                  <span style={{ fontSize: 12, color: '#888' }}>No hay grupos disponibles</span>
                ) : (
                  grupos.map(g => (
                    <Checkbox
                      key={g.id}
                      label={g.name}
                      checked={gruposSeleccionados.indexOf(g.id) !== -1}
                      onChange={(_, checked) => toggleGrupo(g.id, !!checked)}
                    />
                  ))
                )}
              </Stack>
            </Stack.Item>
          </Stack>

          {/* ── Row 3: SLA values ────────────────────────────────────── */}
          <Stack horizontal tokens={{ childrenGap: 16 }} styles={{ root: { flexWrap: 'wrap' } }}>
            <Stack.Item grow={1} styles={{ root: { minWidth: 120 } }}>
              <TextField
                label="SLA Low (horas)"
                required
                type="number"
                min={1}
                step={1}
                value={slaLow}
                onChange={(_, v) => setSlaLow(v ?? '')}
              />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 120 } }}>
              <TextField
                label="SLA Medium (horas)"
                required
                type="number"
                min={1}
                step={1}
                value={slaMedium}
                onChange={(_, v) => setSlaMedium(v ?? '')}
              />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 120 } }}>
              <TextField
                label="SLA High (horas)"
                required
                type="number"
                min={1}
                step={1}
                value={slaHigh}
                onChange={(_, v) => setSlaHigh(v ?? '')}
              />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 120 } }}>
              <TextField
                label="Stop Production (horas)"
                required
                type="number"
                min={1}
                step={1}
                value={slaStopProduction}
                onChange={(_, v) => setSlaStopProduction(v ?? '')}
              />
            </Stack.Item>
          </Stack>

          {/* ── Descripción ──────────────────────────────────────────── */}
          <TextField
            label="Descripción"
            required
            multiline
            rows={3}
            value={descripcion}
            onChange={(_, v) => setDescripcion(v ?? '')}
          />

          {/* ── Procedimiento ─────────────────────────────────────────── */}
          <TextField
            label="Procedimiento de atención"
            required
            multiline
            rows={4}
            value={procedimiento}
            onChange={(_, v) => setProcedimiento(v ?? '')}
          />

        </Stack>
      )}
    </Panel>
  );
};

export default NuevaCategoria;
