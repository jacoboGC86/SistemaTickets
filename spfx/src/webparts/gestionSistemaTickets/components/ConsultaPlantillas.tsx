import * as React from 'react';
import {
  Panel,
  PanelType,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Link,
  mergeStyleSets,
} from '@fluentui/react';
import ListSvc from '../../../services/ListSvc';
import DetallePlantilla from './DetallePlantilla';

// ─── Types ────────────────────────────────────────────────────────────────────

interface IPlantilla {
  id: number;
  title: string;
  descripcion: string;
  modified: string;
  pmPuebla: string;
  pmAguascalientes: string;
}

// ─── Props ────────────────────────────────────────────────────────────────────

export interface IConsultaPlantillasProps {
  isOpen: boolean;
  onDismiss: () => void;
}

// ─── Styles ───────────────────────────────────────────────────────────────────

const classNames = mergeStyleSets({
  container: {
    padding: '0 16px 16px',
  },
  spinnerWrapper: {
    display: 'flex',
    justifyContent: 'center',
    paddingTop: 40,
  },
});

// ─── Helpers ──────────────────────────────────────────────────────────────────

function formatDate(isoString: string): string {
  if (!isoString) return '';
  const d = new Date(isoString);
  const day   = ('0' + d.getDate()).slice(-2);
  const month = ('0' + (d.getMonth() + 1)).slice(-2);
  const year  = d.getFullYear();
  return `${day}/${month}/${year}`;
}

function truncate(text: string, maxLen: number): string {
  if (!text) return '';
  return text.length > maxLen ? text.substring(0, maxLen) + '…' : text;
}

// ─── Columns factory ──────────────────────────────────────────────────────────

function buildColumns(onSelectId: (id: number) => void): IColumn[] {
  return [
    {
      key: 'id',
      name: 'Folio',
      fieldName: 'id',
      minWidth: 50,
      maxWidth: 70,
      isResizable: true,
    },
    {
      key: 'title',
      name: 'Plantilla',
      fieldName: 'title',
      minWidth: 120,
      maxWidth: 220,
      isResizable: true,
      onRender: (item: IPlantilla) => (
        <Link onClick={() => onSelectId(item.id)}>{item.title}</Link>
      ),
    },
    {
      key: 'descripcion',
      name: 'Descripción',
      fieldName: 'descripcion',
      minWidth: 140,
      maxWidth: 260,
      isResizable: true,
      isMultiline: true,
    },
    {
      key: 'modified',
      name: 'Fecha de modificación',
      fieldName: 'modified',
      minWidth: 130,
      maxWidth: 160,
      isResizable: true,
    },
    {
      key: 'pmPuebla',
      name: 'Process Manager Puebla',
      fieldName: 'pmPuebla',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: 'pmAguascalientes',
      name: 'Process Manager Aguascalientes',
      fieldName: 'pmAguascalientes',
      minWidth: 170,
      maxWidth: 220,
      isResizable: true,
    },
  ];
}

// ─── Component ────────────────────────────────────────────────────────────────

const ConsultaPlantillas: React.FC<IConsultaPlantillasProps> = ({ isOpen, onDismiss }) => {
  const [items, setItems]     = React.useState<IPlantilla[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError]     = React.useState<string | null>(null);
  const [selectedId, setSelectedId] = React.useState<number | null>(null);

  const columns = React.useMemo(() => buildColumns(id => setSelectedId(id)), []);

  // Load data whenever the panel opens
  React.useEffect(() => {
    if (!isOpen) return;

    let cancelled = false;

    const fetchPlantillas = async (): Promise<void> => {
      setLoading(true);
      setError(null);
      setItems([]);

      try {
        const data: any[] = await ListSvc.getItems(
          'Templates',
          undefined,
          '$select=ID,Title,Descripcion,Modified,PMPuebla,PMAguascalientes&$orderby=ID desc&$top=200'
        );

        if (cancelled) return;

        const mapped: IPlantilla[] = (data || []).map((item: any) => ({
          id:               item.ID,
          title:            item.Title ?? '',
          descripcion:      truncate(item.Descripcion ?? '', 50),
          modified:         formatDate(item.Modified),
          pmPuebla:         item.PMPuebla ?? '',
          pmAguascalientes: item.PMAguascalientes ?? '',
        }));

        setItems(mapped);
      } catch (err: any) {
        if (!cancelled) setError('No se pudieron cargar las plantillas. Intente de nuevo.');
      } finally {
        if (!cancelled) setLoading(false);
      }
    };

    fetchPlantillas().catch(() => { /* handled inside */ });

    return () => { cancelled = true; };
  }, [isOpen]);

  return (
    <>
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.extraLarge}
      headerText="Consulta de plantillas"
      closeButtonAriaLabel="Cerrar"
    >
      <div className={classNames.container}>
        {loading && (
          <div className={classNames.spinnerWrapper}>
            <Spinner size={SpinnerSize.large} label="Cargando plantillas..." />
          </div>
        )}

        {!loading && error && (
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        )}

        {!loading && !error && (
          <DetailsList
            items={items}
            columns={columns}
            layoutMode={DetailsListLayoutMode.fixedColumns}
            selectionMode={SelectionMode.none}
            isHeaderVisible={true}
          />
        )}
      </div>
    </Panel>

    <DetallePlantilla
      isOpen={selectedId !== null}
      plantillaId={selectedId}
      onDismiss={() => setSelectedId(null)}
    />
  </>
  );
};

export default ConsultaPlantillas;
