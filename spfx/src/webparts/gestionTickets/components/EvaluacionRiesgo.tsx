import * as React from 'react';
import {
  Stack,
  TextField,
  Dropdown,
  IDropdownOption,
  IStackTokens,
  Label,
  TimePicker,
  DatePicker,
  DayOfWeek,
} from '@fluentui/react';

// ─── Types ────────────────────────────────────────────────────────────────────

export interface IEvaluacionRiesgoValues {
  tipoCambio: string;
  sistemasInvolucrados: string;
  fechaPropuesta: Date | null;
  horaPropuesta: Date | null;
  motivoCambio: string;
  evaluacionRiesgos: string;
  interrupcionServicio: string;
  duracionEstimada: string;
  usuariosAfectados: string;
  procesosCriticosImpactados: string;
  pasosImplementacion: string;
  puedeRevertir: string;
  procedimientoRollback: string;
  criteriosExito: string;
}

export interface IEvaluacionRiesgoProps {
  onChange?: (values: IEvaluacionRiesgoValues) => void;
  initialValues?: IEvaluacionRiesgoValues;
  readOnly?: boolean;
}

// ─── Constants ────────────────────────────────────────────────────────────────

const stackTokens: IStackTokens = { childrenGap: 12 };

const tipoCambioOptions: IDropdownOption[] = [
  { key: 'Correctivo', text: 'Correctivo' },
  { key: 'Preventivo', text: 'Preventivo' },
  { key: 'Mejora', text: 'Mejora' },
  { key: 'Emergente', text: 'Emergente' },
];

const siNoOptions: IDropdownOption[] = [
  { key: 'No', text: 'No' },
  { key: 'Sí', text: 'Sí' },
];

const puedeRevertirOptions: IDropdownOption[] = [
  { key: 'Sí', text: 'Sí' },
  { key: 'No', text: 'No' },
];

const initialValues: IEvaluacionRiesgoValues = {
  tipoCambio: '',
  sistemasInvolucrados: '',
  fechaPropuesta: null,
  horaPropuesta: null,
  motivoCambio: '',
  evaluacionRiesgos: '',
  interrupcionServicio: 'No',
  duracionEstimada: '',
  usuariosAfectados: '',
  procesosCriticosImpactados: '',
  pasosImplementacion: '',
  puedeRevertir: 'Sí',
  procedimientoRollback: '',
  criteriosExito: '',
};

// ─── Helpers ──────────────────────────────────────────────────────────────────

function validatePositiveNumber(value: string): string | undefined {
  if (!value) return undefined;
  const n = Number(value);
  if (isNaN(n) || n <= 0 || !Number.isInteger(n)) return 'Debe ser un número entero mayor a cero.';
  return undefined;
}

// ─── Component ────────────────────────────────────────────────────────────────

const EvaluacionRiesgo: React.FC<IEvaluacionRiesgoProps> = ({ onChange, initialValues: initVals, readOnly = false }) => {
  const [values, setValues] = React.useState<IEvaluacionRiesgoValues>(initVals ?? initialValues);

  React.useEffect(() => {
    if (initVals) setValues(initVals);
  }, [initVals]);

  const update = (partial: Partial<IEvaluacionRiesgoValues>): void => {
    if (readOnly) return;
    setValues(prev => {
      const next = { ...prev, ...partial };
      onChange?.(next);
      return next;
    });
  };

  return (
    <Stack tokens={stackTokens}>
      <Label styles={{ root: { fontWeight: 600, fontSize: 16, borderTop: '1px solid #edebe9', paddingTop: 16 } }}>
        Formato de Evaluación de Riesgos
      </Label>
      <Label styles={{ root: { fontWeight: 600, fontSize: 14, borderTop: '1px solid #edebe9', paddingTop: 16 } }}>
        1. Información general
      </Label>

      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        <Dropdown
          label="Tipo de cambio"
          placeholder="Seleccione una opción"
          options={tipoCambioOptions}
          selectedKey={values.tipoCambio || null}
          onChange={(_, option) => update({ tipoCambio: (option?.key as string) || '' })}
          styles={{ root: { minWidth: 260 } }}
          required={!readOnly}
          disabled={readOnly}
        />
        <div style={{ minWidth: 260 }}>
          <DatePicker
            label="Fecha propuesta"
            placeholder="Seleccione una fecha"
            firstDayOfWeek={DayOfWeek.Monday}
            value={values.fechaPropuesta ?? undefined}
            onSelectDate={(date) => update({ fechaPropuesta: date ?? null })}
            formatDate={(date) => (date ? date.toLocaleDateString('es-MX') : '')}
            disabled={readOnly}
            strings={{
              months: ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'],
              shortMonths: ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'],
              days: ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'],
              shortDays: ['D','L','M','X','J','V','S'],
              goToToday: 'Ir a hoy',
              prevMonthAriaLabel: 'Mes anterior',
              nextMonthAriaLabel: 'Mes siguiente',
              prevYearAriaLabel: 'Año anterior',
              nextYearAriaLabel: 'Año siguiente',
            }}
          />
        </div>
        <TimePicker
          label="Hora propuesta"
          dateAnchor={values.fechaPropuesta}
          value={values.horaPropuesta}
          onChange={(_, v) => update({ horaPropuesta: v || null })}
          styles={{ root: { minWidth: 130 } }}
          disabled={readOnly}
        />
      </Stack>

      <Label styles={{ root: { fontWeight: 600, fontSize: 14, borderTop: '1px solid #edebe9', paddingTop: 16 } }}>
        2. Descripción del cambio
      </Label>
      <TextField
        label="Sistemas involucrados"
        value={values.sistemasInvolucrados}
        onChange={(_, v) => update({ sistemasInvolucrados: v || '' })}
        multiline
        rows={3}
        required={!readOnly}
        readOnly={readOnly}
      />

      <TextField
        label="Motivo del cambio"
        value={values.motivoCambio}
        onChange={(_, v) => update({ motivoCambio: v || '' })}
        multiline
        rows={3}
        required={!readOnly}
        readOnly={readOnly}
      />

      <Label styles={{ root: { fontWeight: 600, fontSize: 14, borderTop: '1px solid #edebe9', paddingTop: 16 } }}>
        3. Evaluación de riesgos
      </Label>
      <TextField
        label="Evaluación de riesgos"
        value={values.evaluacionRiesgos}
        onChange={(_, v) => update({ evaluacionRiesgos: v || '' })}
        multiline
        rows={3}
        required={!readOnly}
        readOnly={readOnly}
      />

      <Label styles={{ root: { fontWeight: 600, fontSize: 14, borderTop: '1px solid #edebe9', paddingTop: 16 } }}>
        4. Impacto al negocio
      </Label>
      <Dropdown
          label="¿Interrumpe el servicio?"
          options={siNoOptions}
          selectedKey={values.interrupcionServicio}
          onChange={(_, option) => update({ interrupcionServicio: (option?.key as string) || 'No' })}
          styles={{ root: { maxWidth: 260 } }}
          disabled={readOnly}
        />

      <Stack horizontal tokens={{ childrenGap: 12 }}>
        <TextField
          label="Duración estimada (horas)"
          value={values.duracionEstimada}
          onChange={(_, v) => update({ duracionEstimada: (v || '').replace(/\D/g, '') })}
          onGetErrorMessage={readOnly ? undefined : validatePositiveNumber}
          validateOnLoad={false}
          styles={{ root: { maxWidth: 260 } }}
          required={!readOnly}
          readOnly={readOnly}
        />
        <TextField
            label="Usuarios afectados"
            value={values.usuariosAfectados}
            onChange={(_, v) => update({ usuariosAfectados: (v || '').replace(/\D/g, '') })}
            onGetErrorMessage={readOnly ? undefined : validatePositiveNumber}
            validateOnLoad={false}
            styles={{ root: { maxWidth: 260 } }}
            required={!readOnly}
            readOnly={readOnly}
        />
      </Stack>

      <TextField
        label="Procesos críticos impactados"
        value={values.procesosCriticosImpactados}
        onChange={(_, v) => update({ procesosCriticosImpactados: v || '' })}
        multiline
        rows={3}
        required={!readOnly}
        readOnly={readOnly}
      />

      <Label styles={{ root: { fontWeight: 600, fontSize: 14, borderTop: '1px solid #edebe9', paddingTop: 16 } }}>
        5. Plan de implementación
      </Label>
      <TextField
        label="Pasos de implementación"
        value={values.pasosImplementacion}
        onChange={(_, v) => update({ pasosImplementacion: v || '' })}
        multiline
        rows={4}
        required={!readOnly}
        readOnly={readOnly}
      />

      <Label styles={{ root: { fontWeight: 600, fontSize: 14, borderTop: '1px solid #edebe9', paddingTop: 16 } }}>
        6. Plan de reversa (Rollback)
      </Label>
      <Dropdown
        label="¿Se puede revertir el cambio?"
        options={puedeRevertirOptions}
        selectedKey={values.puedeRevertir}
        onChange={(_, option) => update({ puedeRevertir: (option?.key as string) || 'Sí' })}
        styles={{ root: { minWidth: 260 } }}
        disabled={readOnly}
      />

      {values.puedeRevertir === 'Sí' && (
        <TextField
          label="Procedimiento de rollback"
          value={values.procedimientoRollback}
          onChange={(_, v) => update({ procedimientoRollback: v || '' })}
          multiline
          rows={3}
          required={!readOnly}
          readOnly={readOnly}
        />
      )}

      <Label styles={{ root: { fontWeight: 600, fontSize: 14, borderTop: '1px solid #edebe9', paddingTop: 16 } }}>
        7. Validación y pruebas
      </Label>
      <p>Adjunte las pruebas realizadas y los resultados obtenidos.</p>
      <TextField
        label="Criterios de éxito"
        value={values.criteriosExito}
        onChange={(_, v) => update({ criteriosExito: v || '' })}
        multiline
        rows={3}
        required={!readOnly}
        readOnly={readOnly}
      />
    </Stack>
  );
};

export default EvaluacionRiesgo;
