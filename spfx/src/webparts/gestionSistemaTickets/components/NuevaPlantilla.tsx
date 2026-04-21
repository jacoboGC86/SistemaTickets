import * as React from 'react';
import {
  Panel,
  PanelType,
  Stack,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  IStackTokens,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Pivot,
  PivotItem,
  IconButton,
  Label,
  Text,
} from '@fluentui/react';
import ListSvc from '../../../services/ListSvc';
import UserSvc from '../../../services/UserSvc';

// ─── Types ────────────────────────────────────────────────────────────────────

interface IGrupo {
  id: number;
  name: string;
}

interface IApprovalStep {
  stepName: string;
  useUserDepartment: boolean;
  approvalGroup: number | null;
}

type ControlTag = 'h3' | 'p' | 'br' | 'textarea' | 'checkbox';

interface IFormControl {
  id: string;
  tag: ControlTag;
  text: string;
}

interface IManager {
  id: number;
  email: string;
  name: string;
}

// ─── Props ────────────────────────────────────────────────────────────────────

export interface INuevaPlantillaProps {
  isOpen: boolean;
  onDismiss: () => void;
}

// ─── Constants ────────────────────────────────────────────────────────────────

const stackTokens: IStackTokens = { childrenGap: 12 };

const controlOptions: { tag: ControlTag; label: string; placeholder: string }[] = [
  { tag: 'h3',      label: 'Título',        placeholder: 'Agrega el texto del título...' },
  { tag: 'p',       label: 'Párrafo',       placeholder: 'Agrega el texto del párrafo...' },
  { tag: 'br',      label: 'Espacio',       placeholder: '' },
  { tag: 'textarea',label: 'Caja de texto', placeholder: 'Agrega la etiqueta del cuadro de texto...' },
  { tag: 'checkbox',label: 'Checkbox',      placeholder: 'Agrega la etiqueta del checkbox...' },
];

// ─── Component ────────────────────────────────────────────────────────────────

const NuevaPlantilla: React.FC<INuevaPlantillaProps> = ({ isOpen, onDismiss }) => {

  // ── Form values ────────────────────────────────────────────────────────────
  const [nombre, setNombre]           = React.useState('');
  const [descripcion, setDescripcion] = React.useState('');
  const [pmPueblaId, setPmPueblaId]   = React.useState<number | null>(null);
  const [pmAgsId, setPmAgsId]         = React.useState<number | null>(null);
  const [rutaAprobacion, setRutaAprobacion] = React.useState<IApprovalStep[]>([]);
  const [formControls, setFormControls]     = React.useState<IFormControl[]>([]);

  // ── Loaded data ────────────────────────────────────────────────────────────
  const [grupos, setGrupos]               = React.useState<IGrupo[]>([]);
  const [managers, setManagers]           = React.useState<IManager[]>([]);
  const [managerOptions, setManagerOptions] = React.useState<IDropdownOption[]>([]);

  // ── UI state ───────────────────────────────────────────────────────────────
  const [isLoading, setIsLoading]       = React.useState(false);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [errorMessage, setErrorMessage] = React.useState('');
  const [successMessage, setSuccessMessage] = React.useState('');
  const dataLoaded = React.useRef(false);

  // ── Lifecycle ──────────────────────────────────────────────────────────────

  React.useEffect(() => {
    if (isOpen && !dataLoaded.current) {
      loadData().catch(console.error);
    }
    if (!isOpen) resetForm();
  }, [isOpen]);

  const resetForm = (): void => {
    setNombre('');
    setDescripcion('');
    setPmPueblaId(null);
    setPmAgsId(null);
    setRutaAprobacion([]);
    setFormControls([]);
    setErrorMessage('');
    setSuccessMessage('');
    dataLoaded.current = false;
  };

  // ── Data loading ───────────────────────────────────────────────────────────

  const loadData = async (): Promise<void> => {
    setIsLoading(true);
    try {
      await Promise.all([loadGrupos(), loadManagers()]);
      dataLoaded.current = true;
    } catch (e) {
      console.error('Error loading data:', e);
      setErrorMessage('Error al cargar los datos del formulario.');
    } finally {
      setIsLoading(false);
    }
  };

  const loadGrupos = async (): Promise<void> => {
    const items: any[] = await ListSvc.getItems(
      'Grupos',
      undefined,
      '$orderby=Title asc&$top=100&$select=Id,Title'
    );
    setGrupos((items || []).map((i: any) => ({ id: i.Id, name: i.Title })));
  };

  const loadManagers = async (): Promise<void> => {
    const items: any[] = await ListSvc.getItems(
      'Configuraciones',
      undefined,
      "$filter=Title eq 'IT'&$select=Id,Personas/Id,Personas/Title&$expand=Personas&$top=1"
    );
    if (!items || items.length === 0) return;

    const personasField: any[] = items[0].Personas ?? [];
    if (personasField.length === 0) return;

    const loaded: IManager[] = await Promise.all(
      personasField.map(async (p: any) => {
        const id: number   = p.Id ?? p.LookupId;
        const name: string = p.Title ?? p.LookupValue ?? '';
        try {
          const user = await UserSvc.GetUserById(id);
          return { id, name, email: user?.Email ?? user?.d?.Email ?? '' };
        } catch {
          return { id, name, email: '' };
        }
      })
    );

    setManagers(loaded);
    setManagerOptions(loaded.map(m => ({ key: m.id, text: m.name })));
  };

  // ── Approval route ─────────────────────────────────────────────────────────

  const addGrupo = (grupo: IGrupo | null, useUserDept: boolean): void => {
    setRutaAprobacion(prev => [
      ...prev,
      {
        stepName: useUserDept ? 'Revisión de Manager de Área' : grupo!.name,
        useUserDepartment: useUserDept,
        approvalGroup: useUserDept ? null : grupo!.id,
      },
    ]);
  };

  const removeStep = (index: number): void => {
    setRutaAprobacion(prev => prev.filter((_, i) => i !== index));
  };

  // ── Form builder ───────────────────────────────────────────────────────────

  const addControl = (tag: ControlTag): void => {
    setFormControls(prev => [
      ...prev,
      { id: `ctrl_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`, tag, text: '' },
    ]);
  };

  const updateControlText = (id: string, text: string): void => {
    setFormControls(prev => prev.map(c => (c.id === id ? { ...c, text } : c)));
  };

  const removeControl = (id: string): void => {
    setFormControls(prev => prev.filter(c => c.id !== id));
  };

  // ── Submit ─────────────────────────────────────────────────────────────────

  const handleSubmit = async (): Promise<void> => {
    setErrorMessage('');

    if (!nombre.trim())             { setErrorMessage('Por favor ingrese un nombre de plantilla.'); return; }
    if (!descripcion.trim())        { setErrorMessage('Por favor ingrese una descripción.'); return; }
    if (rutaAprobacion.length === 0){ setErrorMessage('Por favor defina la ruta de aprobación.'); return; }
    if (formControls.length === 0)  { setErrorMessage('Por favor defina el formulario de la plantilla.'); return; }
    if (!pmPueblaId || !pmAgsId)    { setErrorMessage('Por favor seleccione el gerente de proceso de Puebla y Aguascalientes.'); return; }

    const pmPuebla = managers.find(m => m.id === pmPueblaId)!;
    const pmAgs    = managers.find(m => m.id === pmAgsId)!;

    const rutaData = {
      approvalPath: rutaAprobacion,
      formTemplate: formControls.map((c, idx) => ({
        id: `controlTemplate${idx}`,
        tag: c.tag,
        text: c.text,
      })),
      processManager: {
        id: pmPuebla.id, email: pmPuebla.email, name: pmPuebla.name,
      },
      processManagerAguascalientes: {
        id: pmAgs.id, email: pmAgs.email, name: pmAgs.name,
      },
    };

    setIsSubmitting(true);
    try {
      const entityType = await ListSvc.getListItemEntityTypeFullName('Templates');
      const body = JSON.stringify({
        '__metadata': { type: entityType },
        'Title': nombre.trim(),
        'Descripcion': descripcion.trim(),
        'RutaAprobacionEscalacion': JSON.stringify(rutaData),
        'PMPuebla': pmPuebla.name,
        'PMAguascalientes': pmAgs.name,
      });

      await ListSvc.postListItem('Templates', body);
      setSuccessMessage('Plantilla guardada exitosamente.');
      setTimeout(() => { resetForm(); onDismiss(); }, 2000);
    } catch (e) {
      console.error('Error al guardar plantilla:', e);
      setErrorMessage('Error al guardar la plantilla. Por favor, intente de nuevo.');
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
      headerText="Nueva plantilla"
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

          {/* ── Basic fields ─────────────────────────────────────────────── */}
          <TextField
            label="Nombre de la plantilla"
            required
            value={nombre}
            onChange={(_, v) => setNombre(v ?? '')}
          />

          <TextField
            label="Descripción"
            required
            multiline
            rows={3}
            value={descripcion}
            onChange={(_, v) => setDescripcion(v ?? '')}
          />

          <Stack horizontal tokens={{ childrenGap: 16 }} styles={{ root: { flexWrap: 'wrap' } }}>
            <Stack.Item grow={1} styles={{ root: { minWidth: 220 } }}>
              <Dropdown
                label="Gerente de Procesos Puebla"
                required
                options={managerOptions}
                selectedKey={pmPueblaId}
                onChange={(_, o) => setPmPueblaId(o ? Number(o.key) : null)}
                placeholder="Seleccione el gerente"
              />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 220 } }}>
              <Dropdown
                label="Gerente de Procesos Aguascalientes"
                required
                options={managerOptions}
                selectedKey={pmAgsId}
                onChange={(_, o) => setPmAgsId(o ? Number(o.key) : null)}
                placeholder="Seleccione el gerente"
              />
            </Stack.Item>
          </Stack>

          {/* ── Tabs ─────────────────────────────────────────────────────── */}
          <Pivot styles={{ root: { marginTop: 8 } }}>

            {/* Tab: Ruta de Aprobación */}
            <PivotItem headerText="Ruta de Aprobación">
              <Stack
                horizontal
                tokens={{ childrenGap: 24 }}
                styles={{ root: { marginTop: 12, flexWrap: 'wrap', alignItems: 'flex-start' } }}
              >
                {/* Available groups */}
                <Stack styles={{ root: { minWidth: 220, maxWidth: 280 } }} tokens={{ childrenGap: 6 }}>
                  <Label>Grupos de aprobación</Label>

                  {/* Static: Manager review */}
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}
                    styles={{ root: { padding: '4px 0' } }}>
                    <IconButton
                      iconProps={{ iconName: 'Add' }}
                      title="Agregar"
                      ariaLabel="Agregar Revisión de Manager de Área"
                      onClick={() => addGrupo(null, true)}
                    />
                    <Text>Revisión de Manager de Área</Text>
                  </Stack>

                  {grupos.map(g => (
                    <Stack key={g.id} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}
                      styles={{ root: { padding: '4px 0' } }}>
                      <IconButton
                        iconProps={{ iconName: 'Add' }}
                        title="Agregar"
                        ariaLabel={`Agregar ${g.name}`}
                        onClick={() => addGrupo(g, false)}
                      />
                      <Text>{g.name}</Text>
                    </Stack>
                  ))}
                </Stack>

                {/* Approval route list */}
                <Stack grow tokens={{ childrenGap: 6 }}>
                  <Label>Ruta de aprobación de la plantilla</Label>
                  {rutaAprobacion.length === 0 ? (
                    <Text variant="small" styles={{ root: { color: '#888' } }}>
                      Agregue grupos desde el panel izquierdo.
                    </Text>
                  ) : (
                    rutaAprobacion.map((step, idx) => (
                      <Stack
                        key={idx}
                        horizontal
                        verticalAlign="center"
                        tokens={{ childrenGap: 8 }}
                        styles={{ root: { padding: '6px 10px', border: '1px solid #edebe9', borderRadius: 4 } }}
                      >
                        <Text styles={{ root: { flex: 1 } }}>{`${idx + 1}. ${step.stepName}`}</Text>
                        <IconButton
                          iconProps={{ iconName: 'Delete' }}
                          title="Remover"
                          ariaLabel="Remover paso"
                          onClick={() => removeStep(idx)}
                        />
                      </Stack>
                    ))
                  )}
                </Stack>
              </Stack>
            </PivotItem>

            {/* Tab: Formulario */}
            <PivotItem headerText="Formulario">
              <Stack
                horizontal
                tokens={{ childrenGap: 24 }}
                styles={{ root: { marginTop: 12, flexWrap: 'wrap', alignItems: 'flex-start' } }}
              >
                {/* Controls palette */}
                <Stack styles={{ root: { minWidth: 180, maxWidth: 220 } }} tokens={{ childrenGap: 6 }}>
                  <Label>Controles</Label>
                  {controlOptions.map(c => (
                    <Stack key={c.tag} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}
                      styles={{ root: { padding: '4px 0' } }}>
                      <IconButton
                        iconProps={{ iconName: 'Add' }}
                        title="Agregar"
                        ariaLabel={`Agregar ${c.label}`}
                        onClick={() => addControl(c.tag)}
                      />
                      <Text>{c.label}</Text>
                    </Stack>
                  ))}
                </Stack>

                {/* Form builder */}
                <Stack grow tokens={{ childrenGap: 8 }}>
                  <Label>Formulario de la plantilla</Label>
                  {formControls.length === 0 ? (
                    <Text variant="small" styles={{ root: { color: '#888' } }}>
                      Agregue controles desde el panel izquierdo.
                    </Text>
                  ) : (
                    formControls.map(c => {
                      const meta = controlOptions.find(o => o.tag === c.tag)!;
                      return (
                        <Stack
                          key={c.id}
                          tokens={{ childrenGap: 6 }}
                          styles={{ root: { padding: 10, border: '1px solid #edebe9', borderRadius: 4 } }}
                        >
                          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                            <Label styles={{ root: { margin: 0 } }}>{meta.label}</Label>
                            <IconButton
                              iconProps={{ iconName: 'Delete' }}
                              title="Remover"
                              ariaLabel={`Remover ${meta.label}`}
                              onClick={() => removeControl(c.id)}
                            />
                          </Stack>

                          {c.tag === 'br' ? (
                            <Text variant="small" styles={{ root: { color: '#888' } }}>
                              Espacio en blanco
                            </Text>
                          ) : (
                            <TextField
                              placeholder={meta.placeholder}
                              multiline={c.tag === 'p'}
                              rows={c.tag === 'p' ? 3 : undefined}
                              value={c.text}
                              onChange={(_, v) => updateControlText(c.id, v ?? '')}
                            />
                          )}
                        </Stack>
                      );
                    })
                  )}
                </Stack>
              </Stack>
            </PivotItem>

          </Pivot>
        </Stack>
      )}
    </Panel>
  );
};

export default NuevaPlantilla;
