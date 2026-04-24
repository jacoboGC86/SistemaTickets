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
  mergeStyleSets,
  SearchBox,
} from '@fluentui/react';
import ListSvc from '../../../services/ListSvc';
import DetalleCategoria from './DetalleCategoria';

// ─── Types ────────────────────────────────────────────────────────────────────

interface ICategoria {
  id: number;
  tipo: string;
  categoriaPadre: string;
  title: string;
  prioridad: string;
  plantilla: string;
  pmPuebla: string;
  pmAguascalientes: string;
}

// ─── Props ────────────────────────────────────────────────────────────────────

export interface IConsultaCategoriasProps {
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
  searchBox: {
    marginBottom: 12,
  },
});

// ─── Column definitions ───────────────────────────────────────────────────────

type SortKey = keyof ICategoria;

const COLUMN_DEFS: Array<Omit<IColumn, 'onColumnClick' | 'isSorted' | 'isSortedDescending'> & { fieldName: SortKey }> = [
  { key: 'id',               name: 'Folio',              fieldName: 'id',               minWidth: 50,  maxWidth: 70,  isResizable: true },
  { key: 'tipo',             name: 'Tipo',               fieldName: 'tipo',             minWidth: 80,  maxWidth: 120, isResizable: true },
  { key: 'categoriaPadre',   name: 'Categoría padre',    fieldName: 'categoriaPadre',   minWidth: 120, maxWidth: 200, isResizable: true },
  { key: 'title',            name: 'Categoría',          fieldName: 'title',            minWidth: 120, maxWidth: 200, isResizable: true },
  { key: 'prioridad',        name: 'Prioridad',          fieldName: 'prioridad',        minWidth: 80,  maxWidth: 120, isResizable: true },
  { key: 'plantilla',        name: 'Plantilla',          fieldName: 'plantilla',        minWidth: 120, maxWidth: 200, isResizable: true },
  { key: 'pmPuebla',         name: 'PM Puebla',          fieldName: 'pmPuebla',         minWidth: 120, maxWidth: 180, isResizable: true },
  { key: 'pmAguascalientes', name: 'PM Aguascalientes',  fieldName: 'pmAguascalientes', minWidth: 130, maxWidth: 200, isResizable: true },
];

function buildColumns(
  sortKey: SortKey,
  sortAsc: boolean,
  onColumnClick: (key: SortKey) => void
): IColumn[] {
  return COLUMN_DEFS.map(def => ({
    ...def,
    isSorted: def.fieldName === sortKey,
    isSortedDescending: def.fieldName === sortKey ? !sortAsc : false,
    onColumnClick: () => onColumnClick(def.fieldName),
  }));
}

// ─── Component ────────────────────────────────────────────────────────────────

const ConsultaCategorias: React.FC<IConsultaCategoriasProps> = ({ isOpen, onDismiss }) => {
  const [items, setItems]     = React.useState<ICategoria[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError]     = React.useState<string | null>(null);
  const [filterText, setFilterText] = React.useState('');
  const [sortKey, setSortKey]         = React.useState<SortKey>('id');
  const [sortAsc, setSortAsc]         = React.useState(false);
  const [selectedCategoriaId, setSelectedCategoriaId] = React.useState<number | null>(null);
  const [isDetalleOpen, setIsDetalleOpen]             = React.useState(false);

  const handleOpenDetalle = (id: number): void => {
    setSelectedCategoriaId(id);
    setIsDetalleOpen(true);
  };

  const onRenderItemColumn = (item?: ICategoria, _index?: number, column?: IColumn): React.ReactNode => {
    if (!item) return null;
    if (column?.fieldName === 'title') {
      return (
        <a
          href="#"
          onClick={(e) => { e.preventDefault(); handleOpenDetalle(item.id); }}
          style={{ color: '#0078d4', textDecoration: 'none', cursor: 'pointer' }}
        >
          {item.title}
        </a>
      );
    }
    return (item as any)[column?.fieldName ?? ''];
  };

  const handleColumnClick = React.useCallback((key: SortKey) => {
    setSortKey(prev => {
      if (prev === key) { setSortAsc(a => !a); return prev; }
      setSortAsc(true);
      return key;
    });
  }, []);

  const columns = React.useMemo(
    () => buildColumns(sortKey, sortAsc, handleColumnClick),
    [sortKey, sortAsc, handleColumnClick]
  );

  React.useEffect(() => {
    if (!isOpen) return;

    let cancelled = false;

    const fetchCategorias = async (): Promise<void> => {
      setLoading(true);
      setError(null);
      setItems([]);

      try {
        const data: any[] = await ListSvc.getItems(
          'Categorias',
          undefined,
          '$select=ID,Title,Prioridad,TipoCategoria,' +
          'CategoriaPadre/Title,' +
          'TemplateAtencion/Title,' +
          'TemplateAtencion/PMPuebla,' +
          'TemplateAtencion/PMAguascalientes' +
          '&$expand=CategoriaPadre,TemplateAtencion' +
          '&$orderby=ID desc&$top=500'
        );

        if (cancelled) return;

        const mapped: ICategoria[] = (data || []).map((item: any) => ({
          id:               item.ID,
          tipo:             item.TipoCategoria ?? '',
          title:            item.Title ?? '',
          prioridad:        item.Prioridad ?? '',
          categoriaPadre:   item.CategoriaPadre?.Title ?? 'Categoría padre',
          plantilla:        item.TemplateAtencion?.Title ?? 'Sin plantilla',
          pmPuebla:         item.TemplateAtencion?.PMPuebla ?? '',
          pmAguascalientes: item.TemplateAtencion?.PMAguascalientes ?? '',
        }));

        setItems(mapped);
      } catch (err: any) {
        if (!cancelled) setError('No se pudieron cargar las categorías. Intente de nuevo.');
      } finally {
        if (!cancelled) setLoading(false);
      }
    };

    fetchCategorias().catch(() => { /* handled inside */ });

    return () => { cancelled = true; };
  }, [isOpen]);

  const filteredItems = React.useMemo(() => {
    const q = filterText.trim().toLowerCase();
    const filtered = q
      ? items.filter(item =>
          String(item.id).includes(q) ||
          item.tipo.toLowerCase().includes(q) ||
          item.categoriaPadre.toLowerCase().includes(q) ||
          item.title.toLowerCase().includes(q) ||
          item.prioridad.toLowerCase().includes(q) ||
          item.plantilla.toLowerCase().includes(q) ||
          item.pmPuebla.toLowerCase().includes(q) ||
          item.pmAguascalientes.toLowerCase().includes(q)
        )
      : items;

    return [...filtered].sort((a, b) => {
      const aVal = sortKey === 'id' ? a.id : String(a[sortKey]).toLowerCase();
      const bVal = sortKey === 'id' ? b.id : String(b[sortKey]).toLowerCase();
      if (aVal < bVal) return sortAsc ? -1 : 1;
      if (aVal > bVal) return sortAsc ? 1 : -1;
      return 0;
    });
  }, [items, filterText, sortKey, sortAsc]);

  return (
    <>
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.extraLarge}
      headerText="Consulta de categorías"
      closeButtonAriaLabel="Cerrar"
    >
      <div className={classNames.container}>
        {loading && (
          <div className={classNames.spinnerWrapper}>
            <Spinner size={SpinnerSize.large} label="Cargando categorías..." />
          </div>
        )}

        {!loading && error && (
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        )}

        {!loading && !error && (
          <>
            <SearchBox
              className={classNames.searchBox}
              placeholder="Buscar en todos los campos..."
              value={filterText}
              onChange={(_, newValue) => setFilterText(newValue ?? '')}
              onClear={() => setFilterText('')}
            />
          <DetailsList
            items={filteredItems}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
            isHeaderVisible={true}
            onRenderItemColumn={onRenderItemColumn}
          />
          </>  
        )}
      </div>
    </Panel>

    <DetalleCategoria
      isOpen={isDetalleOpen}
      categoriaId={selectedCategoriaId}
      onDismiss={() => setIsDetalleOpen(false)}
    />
    </>
  );
};

export default ConsultaCategorias;
