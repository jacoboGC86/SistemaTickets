import * as React from 'react';
import {
  Panel,
  PanelType,
  Stack,
  TextField,
  Dropdown,
  IDropdownOption,
  IStackTokens,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Pivot,
  PivotItem,
  Label,
  Text,
  PrimaryButton,
  DefaultButton,
  IconButton,
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

interface IRutaData {
  approvalPath: IApprovalStep[];
  formTemplate: { id: string; tag: ControlTag; text: string }[];
  processManager: { id: number; email: string; name: string };
  processManagerAguascalientes: { id: number; email: string; name: string };
}

// ─── Props ────────────────────────────────────────────────────────────────────

export interface IDetallePlantillaProps {
  isOpen: boolean;
  plantillaId: number | null;
  onDismiss: () => void;
}

// ─── Constants ────────────────────────────────────────────────────────────────

const stackTokens: IStackTokens = { childrenGap: 12 };

const controlOptions: { tag: ControlTag; label: string }[] = [
  { tag: 'h3',       label: 'Título' },
  { tag: 'p',        label: 'Párrafo' },
  { tag: 'br',       label: 'Espacio' },
  { tag: 'textarea', label: 'Caja de texto' },
  { tag: 'checkbox', label: 'Checkbox' },
];

// ─── Component ────────────────────────────────────────────────────────────────

const DetallePlantilla: React.FC<IDetallePlantillaProps> = ({ isOpen, plantillaId, onDismiss }) => {

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
  const [isEditing, setIsEditing]       = React.useState(true);
  const [errorMessage, setErrorMessage]   = React.useState('');
  const [successMessage, setSuccessMessage] = React.useState('');

  // ── Lifecycle ──────────────────────────────────────────────────────────────

  React.useEffect(() => {
    if (isOpen && plantillaId !== null) {
      loadDetalle(plantillaId).catch(console.error);
    }
    if (!isOpen) resetState();
  }, [isOpen, plantillaId]);

  const resetState = (): void => {
    setNombre('');
    setDescripcion('');
    setPmPueblaId(null);
    setPmAgsId(null);
    setRutaAprobacion([]);
    setFormControls([]);
    setGrupos([]);
    setManagers([]);
    setManagerOptions([]);
    setErrorMessage('');
    setSuccessMessage('');
    setIsEditing(true);
  };

  // ── Data loading ───────────────────────────────────────────────────────────

  const loadDetalle = async (id: number): Promise<void> => {
    setIsLoading(true);
    setErrorMessage('');
    try {
      const [item] = await Promise.all([
        ListSvc.getItemById('Templates', id),
        loadGrupos(),
        loadManagers(),
      ]);

      if (!item || item.error) {
        setErrorMessage('No se pudo cargar la plantilla.');
        return;
      }

      setNombre(item.Title ?? item.d?.Title ?? '');
      setDescripcion(item.Descripcion ?? item.d?.Descripcion ?? '');

      const rutaRaw: string =
        item.RutaAprobacionEscalacion ?? item.d?.RutaAprobacionEscalacion ?? '';

      if (rutaRaw) {
        try {
          const rutaData: IRutaData = JSON.parse(rutaRaw);

          setRutaAprobacion(rutaData.approvalPath ?? []);

          const controls: IFormControl[] = (rutaData.formTemplate ?? []).map(
            (c, idx) => ({
              id:   c.id ?? `ctrl_loaded_${idx}`,
              tag:  c.tag,
              text: c.text ?? '',
            })
          );
          setFormControls(controls);

          const pmPueblaIdResolved = rutaData.processManager?.id ?? null;
          const pmAgsIdResolved    = rutaData.processManagerAguascalientes?.id ?? null;
          const pmPueblaName: string = rutaData.processManager?.name ?? '';
          const pmAgsName:    string = rutaData.processManagerAguascalientes?.name ?? '';

          setPmPueblaId(pmPueblaIdResolved);
          setPmAgsId(pmAgsIdResolved);

          // Ensure options contain the loaded managers even if not in Configuraciones
          setManagerOptions(prev => {
            const merged = [...prev];
            const addIfMissing = (optId: number | null, name: string): void => {
              if (optId !== null && !merged.find(o => o.key === optId)) {
                merged.push({ key: optId, text: name });
              }
            };
            addIfMissing(pmPueblaIdResolved, pmPueblaName);
            addIfMissing(pmAgsIdResolved, pmAgsName);
            return merged;
          });
        } catch {
          setErrorMessage('Error al parsear los datos de la plantilla.');
        }
      }
    } catch (e) {
      console.error('Error loading detalle plantilla:', e);
      setErrorMessage('Error al cargar los detalles de la plantilla.');
    } finally {
      setIsLoading(false);
    }
  };

  const loadGrupos = async (): Promise<void> => {
    try {
      const items: any[] = await ListSvc.getItems(
        'Grupos',
        undefined,
        '$orderby=Title asc&$top=100&$select=Id,Title'
      );
      setGrupos((items || []).map((i: any) => ({ id: i.Id, name: i.Title })));
    } catch {
      // non-fatal
    }
  };

  const loadManagers = async (): Promise<void> => {
    try {
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
    } catch {
      // non-fatal
    }
  };

  // ── Approval route ─────────────────────────────────────────────────────────

  const addGrupo = (grupo: IGrupo | null, useUserDept: boolean): void => {
    setRutaAprobacion(prev => [
      ...prev,
      {
        stepName: useUserDept ? 'Revisión de Manager' : grupo!.name,
        useUserDepartment: useUserDept,
        approvalGroup: useUserDept ? null : grupo!.id,
      },
    ]);
  };

  const removeStep = (index: number): void => {
    setRutaAprobacion(prev => prev.filter((_, i) => i !== index));
  };

  const dragIndexRef = React.useRef<number | null>(null);

  const moveStep = (fromIndex: number, toIndex: number): void => {
    setRutaAprobacion(prev => {
      const next = [...prev];
      const [moved] = next.splice(fromIndex, 1);
      next.splice(toIndex, 0, moved);
      return next;
    });
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

  const dragControlIdRef = React.useRef<string | null>(null);

  const moveControl = (fromId: string, toId: string): void => {
    setFormControls(prev => {
      const next = [...prev];
      const fromIdx = next.findIndex(c => c.id === fromId);
      const toIdx   = next.findIndex(c => c.id === toId);
      if (fromIdx === -1 || toIdx === -1) return prev;
      const [moved] = next.splice(fromIdx, 1);
      next.splice(toIdx, 0, moved);
      return next;
    });
  };

  // ── Save ───────────────────────────────────────────────────────────────────

  const handleSave = async (): Promise<void> => {
    setErrorMessage('');

    if (!nombre.trim())      { setErrorMessage('Por favor ingrese un nombre de plantilla.'); return; }
    if (!descripcion.trim()) { setErrorMessage('Por favor ingrese una descripción.'); return; }

    const pmPuebla = managers.find(m => m.id === pmPueblaId);
    const pmAgs    = managers.find(m => m.id === pmAgsId);

    // Fall back to the option label when manager is not in the loaded IManager[] list
    const pmPueblaName = pmPuebla?.name ?? managerOptions.find(o => o.key === pmPueblaId)?.text ?? '';
    const pmAgsName    = pmAgs?.name    ?? managerOptions.find(o => o.key === pmAgsId)?.text    ?? '';

    const rutaData: IRutaData = {
      approvalPath: rutaAprobacion,
      formTemplate: formControls.map((c, idx) => ({
        id:   `controlTemplate${idx}`,
        tag:  c.tag,
        text: c.text,
      })),
      processManager: {
        id:    pmPueblaId ?? 0,
        email: pmPuebla?.email ?? '',
        name:  pmPueblaName,
      },
      processManagerAguascalientes: {
        id:    pmAgsId ?? 0,
        email: pmAgs?.email ?? '',
        name:  pmAgsName,
      },
    };

    setIsSubmitting(true);
    try {
      const entityType = await ListSvc.getListItemEntityTypeFullName('Templates');
      const body = JSON.stringify({
        '__metadata': { type: entityType },
        Title:                    nombre.trim(),
        Descripcion:              descripcion.trim(),
        RutaAprobacionEscalacion: JSON.stringify(rutaData),
        PMPuebla:                 pmPueblaName,
        PMAguascalientes:         pmAgsName,
      });

      await ListSvc.putListItem('Templates', plantillaId!, body);
      setSuccessMessage('Plantilla actualizada exitosamente.');
    } catch (e) {
      console.error('Error al actualizar plantilla:', e);
      setErrorMessage('Error al actualizar la plantilla. Por favor, intente de nuevo.');
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
      headerText="Detalle de plantilla"
      closeButtonAriaLabel="Cerrar"
      isFooterAtBottom
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center" horizontalAlign="end" styles={{ root: { width: '100%' } }}>
          {isSubmitting && <Spinner size={SpinnerSize.small} />}
          <PrimaryButton text="Enviar" onClick={handleSave} disabled={isSubmitting || isLoading} />
          <DefaultButton text="Cerrar" onClick={onDismiss} disabled={isSubmitting} />
        </Stack>
      )}
    >
      {isLoading ? (
        <Spinner label="Cargando plantilla..." size={SpinnerSize.large} styles={{ root: { marginTop: 40 } }} />
      ) : (
        <Stack tokens={stackTokens} styles={{ root: { paddingTop: 16, paddingLeft: 16, paddingRight: 16 } }}>

          {errorMessage && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMessage('')} isMultiline={false}>
              {errorMessage}
            </MessageBar>
          )}
          {successMessage && (
            <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setSuccessMessage('')} isMultiline={false}>
              {successMessage}
            </MessageBar>
          )}

          {/* ── Basic fields ─────────────────────────────────────────────── */}
          <TextField
            label="Nombre de la plantilla"
            value={nombre}
            onChange={(_, v) => setNombre(v ?? '')}
            readOnly={!isEditing}
            required={isEditing}
          />

          <TextField
            label="Descripción"
            multiline
            rows={3}
            value={descripcion}
            onChange={(_, v) => setDescripcion(v ?? '')}
            readOnly={!isEditing}
            required={isEditing}
          />

          <Stack horizontal tokens={{ childrenGap: 16 }} styles={{ root: { flexWrap: 'wrap' } }}>
            <Stack.Item grow={1} styles={{ root: { minWidth: 220 } }}>
              <Dropdown
                label="Gerente de Procesos Puebla"
                options={managerOptions}
                selectedKey={pmPueblaId}
                onChange={(_, o) => setPmPueblaId(o ? Number(o.key) : null)}
                disabled={!isEditing}
                placeholder="—"
              />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 220 } }}>
              <Dropdown
                label="Gerente de Procesos Aguascalientes"
                options={managerOptions}
                selectedKey={pmAgsId}
                onChange={(_, o) => setPmAgsId(o ? Number(o.key) : null)}
                disabled={!isEditing}
                placeholder="—"
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
                {/* Available groups — only shown when editing */}
                {isEditing && (
                  <Stack styles={{ root: { minWidth: 220, maxWidth: 280 } }} tokens={{ childrenGap: 6 }}>
                    <Label>Grupos de aprobación</Label>

                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}
                      styles={{ root: { padding: '4px 0' } }}>
                      <IconButton
                        iconProps={{ iconName: 'Add' }}
                        title="Agregar Revisión de Manager"
                        ariaLabel="Agregar Revisión de Manager"
                        onClick={() => addGrupo(null, true)}
                      />
                      <Text>Revisión de Manager</Text>
                    </Stack>

                    {grupos.map(g => (
                      <Stack key={g.id} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}
                        styles={{ root: { padding: '4px 0' } }}>
                        <IconButton
                          iconProps={{ iconName: 'Add' }}
                          title={`Agregar ${g.name}`}
                          ariaLabel={`Agregar ${g.name}`}
                          onClick={() => addGrupo(g, false)}
                        />
                        <Text>{g.name}</Text>
                      </Stack>
                    ))}
                  </Stack>
                )}

                {/* Approval route list */}
                <Stack grow tokens={{ childrenGap: 6 }}>
                  <Label>Ruta de aprobación de la plantilla</Label>
                  {rutaAprobacion.length === 0 ? (
                    <Text variant="small" styles={{ root: { color: '#888' } }}>
                      {isEditing ? 'Agregue grupos desde el panel izquierdo.' : 'Sin pasos de aprobación definidos.'}
                    </Text>
                  ) : (
                    rutaAprobacion.map((step, idx) => (
                      <Stack
                        key={idx}
                        horizontal
                        verticalAlign="center"
                        tokens={{ childrenGap: 8 }}
                        draggable={isEditing}
                        onDragStart={() => { dragIndexRef.current = idx; }}
                        onDragOver={e => { e.preventDefault(); }}
                        onDrop={() => {
                          if (dragIndexRef.current !== null && dragIndexRef.current !== idx) {
                            moveStep(dragIndexRef.current, idx);
                          }
                          dragIndexRef.current = null;
                        }}
                        styles={{
                          root: {
                            padding: '6px 10px',
                            border: '1px solid #edebe9',
                            borderRadius: 4,
                            cursor: isEditing ? 'grab' : 'default',
                          },
                        }}
                      >
                        {isEditing && (
                          <IconButton
                            iconProps={{ iconName: 'GripperDotsVertical' }}
                            title="Arrastrar para reordenar"
                            ariaLabel="Arrastrar para reordenar"
                            styles={{ root: { cursor: 'grab', color: '#888' } }}
                          />
                        )}
                        <Text styles={{ root: { flex: 1 } }}>{`${idx + 1}. ${step.stepName}`}</Text>
                        {isEditing && (
                          <IconButton
                            iconProps={{ iconName: 'Delete' }}
                            title="Remover"
                            ariaLabel="Remover paso"
                            onClick={() => removeStep(idx)}
                          />
                        )}
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
                {/* Controls palette — only shown when editing */}
                {isEditing && (
                  <Stack styles={{ root: { minWidth: 180, maxWidth: 220 } }} tokens={{ childrenGap: 6 }}>
                    <Label>Controles</Label>
                    {controlOptions.map(c => (
                      <Stack key={c.tag} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}
                        styles={{ root: { padding: '4px 0' } }}>
                        <IconButton
                          iconProps={{ iconName: 'Add' }}
                          title={`Agregar ${c.label}`}
                          ariaLabel={`Agregar ${c.label}`}
                          onClick={() => addControl(c.tag)}
                        />
                        <Text>{c.label}</Text>
                      </Stack>
                    ))}
                  </Stack>
                )}

                {/* Form controls list */}
                <Stack grow tokens={{ childrenGap: 8 }}>
                  <Label>Formulario de la plantilla</Label>
                  {formControls.length === 0 ? (
                    <Text variant="small" styles={{ root: { color: '#888' } }}>
                      {isEditing ? 'Agregue controles desde el panel izquierdo.' : 'Sin controles definidos.'}
                    </Text>
                  ) : (
                    formControls.map(c => {
                      const meta = controlOptions.find(o => o.tag === c.tag)!;
                      return (
                        <Stack
                          key={c.id}
                          tokens={{ childrenGap: 6 }}
                          draggable={isEditing}
                          onDragStart={() => { dragControlIdRef.current = c.id; }}
                          onDragOver={e => { e.preventDefault(); }}
                          onDrop={() => {
                            if (dragControlIdRef.current && dragControlIdRef.current !== c.id) {
                              moveControl(dragControlIdRef.current, c.id);
                            }
                            dragControlIdRef.current = null;
                          }}
                          styles={{
                            root: {
                              padding: 10,
                              border: '1px solid #edebe9',
                              borderRadius: 4,
                              cursor: isEditing ? 'grab' : 'default',
                            },
                          }}
                        >
                          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                              {isEditing && (
                                <IconButton
                                  iconProps={{ iconName: 'GripperDotsVertical' }}
                                  title="Arrastrar para reordenar"
                                  ariaLabel="Arrastrar para reordenar"
                                  styles={{ root: { cursor: 'grab', color: '#888' } }}
                                />
                              )}
                              <Label styles={{ root: { margin: 0 } }}>{meta?.label ?? c.tag}</Label>
                            </Stack>
                            {isEditing && (
                              <IconButton
                                iconProps={{ iconName: 'Delete' }}
                                title="Remover"
                                ariaLabel={`Remover ${meta?.label ?? c.tag}`}
                                onClick={() => removeControl(c.id)}
                              />
                            )}
                          </Stack>

                          {c.tag === 'br' ? (
                            <Text variant="small" styles={{ root: { color: '#888' } }}>
                              Espacio en blanco
                            </Text>
                          ) : (
                            <TextField
                              multiline={c.tag === 'p'}
                              rows={c.tag === 'p' ? 3 : undefined}
                              value={c.text}
                              readOnly={!isEditing}
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

export default DetallePlantilla;
