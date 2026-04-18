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
  Stack,
  Text,
  SearchBox,
  Link,
  Dropdown,
  IDropdownOption,
} from '@fluentui/react';
import ListSvc from '../../../services/ListSvc';

// ─── Types ────────────────────────────────────────────────────────────────────

interface ITicketAdmonRow {
  id: number;
  title: string;
  processManager: string;
  tipoTicket: string;
  categoria: string;
  status: string;
  prioridad: string;
  tiempoAsignacion: string;
  mes: string;
  anio: number;
  solicita: string;
}

export interface ITodosLosTicketsProps {
  isOpen: boolean;
  onDismiss: () => void;
  onVerDetalle?: (id: number) => void;
}

interface ITodosLosTicketsState {
  loading: boolean;
  error: string | null;
  tickets: ITicketAdmonRow[];
  filterText: string;
  sortKey: keyof ITicketAdmonRow;
  isSortedDescending: boolean;
  selectedYear: number;
}

// ─── Constants ────────────────────────────────────────────────────────────────

const CURRENT_YEAR = new Date().getFullYear();

const YEAR_OPTIONS: IDropdownOption[] = Array.from(
  { length: CURRENT_YEAR - 2025 + 1 },
  (_, i) => { const y = 2025 + i; return { key: y, text: String(y) }; }
);

const MESES = [
  'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
  'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre',
];

// ─── Helpers ──────────────────────────────────────────────────────────────────

function getTimeDifference(dateStr: string): string {
  const now = new Date();
  const date = new Date(dateStr);
  const diffMs = now.getTime() - date.getTime();

  const minutes = Math.floor(diffMs / 60000) % 60;
  const hours = Math.floor(diffMs / (1000 * 60 * 60)) % 24;
  const days = Math.floor(diffMs / (1000 * 60 * 60 * 24));

  const parts: string[] = [];
  if (days > 0) parts.push(`${days} día(s)`);
  if (hours > 0) parts.push(`${hours} hora(s)`);
  if (minutes > 0 || parts.length === 0) parts.push(`${minutes} minuto(s)`);
  return parts.join(', ');
}

// ─── Base column definitions ──────────────────────────────────────────────────

const BASE_COLUMNS: Pick<IColumn, 'key' | 'name' | 'fieldName' | 'minWidth' | 'maxWidth'>[] = [
  { key: 'id',               name: 'Folio',                fieldName: 'id',               minWidth: 50,  maxWidth: 70  },
  { key: 'title',            name: 'Ticket',               fieldName: 'title',            minWidth: 150, maxWidth: 250 },
  { key: 'processManager',   name: 'Process Manager',      fieldName: 'processManager',   minWidth: 120, maxWidth: 200 },
  { key: 'tipoTicket',       name: 'Tipo',                 fieldName: 'tipoTicket',       minWidth: 90,  maxWidth: 140 },
  { key: 'categoria',        name: 'Categoría',            fieldName: 'categoria',        minWidth: 130, maxWidth: 220 },
  { key: 'status',           name: 'Estatus',              fieldName: 'status',           minWidth: 90,  maxWidth: 140 },
  { key: 'prioridad',        name: 'Prioridad',            fieldName: 'prioridad',        minWidth: 80,  maxWidth: 120 },
  { key: 'tiempoAsignacion', name: 'Tiempo de asignación', fieldName: 'tiempoAsignacion', minWidth: 140, maxWidth: 200 },
  { key: 'mes',              name: 'Mes',                  fieldName: 'mes',              minWidth: 80,  maxWidth: 120 },
  { key: 'anio',             name: 'Año',                  fieldName: 'anio',             minWidth: 60,  maxWidth: 80  },
  { key: 'solicita',         name: 'Solicita',             fieldName: 'solicita',         minWidth: 120, maxWidth: 200 },
];

// ─── Component ────────────────────────────────────────────────────────────────

export default class TodosLosTickets extends React.Component<ITodosLosTicketsProps, ITodosLosTicketsState> {
  constructor(props: ITodosLosTicketsProps) {
    super(props);
    this.state = {
      loading: false,
      error: null,
      tickets: [],
      filterText: '',
      sortKey: 'id',
      isSortedDescending: true,
      selectedYear: CURRENT_YEAR,
    };
  }

  public componentDidUpdate(prevProps: ITodosLosTicketsProps): void {
    if (this.props.isOpen && !prevProps.isOpen) {
      this.setState({ filterText: '', sortKey: 'id', isSortedDescending: true, selectedYear: CURRENT_YEAR }, () => {
        this._loadTickets().catch(console.error);
      });
    }
  }

  private async _loadTickets(): Promise<void> {
    this.setState({ loading: true, error: null, tickets: [] });

    try {
      const { selectedYear } = this.state;
      const yearStart = `${selectedYear}-01-01T00:00:00Z`;
      const yearEnd   = `${selectedYear}-12-31T23:59:59Z`;

      const ticketsRaw = await ListSvc.getItems(
        'Tickets',
        undefined,
        `$select=Id,Title,TipoTicket,Status,Prioridad,Modified,` +
          `ProcessManager/Title,Categoria/Title,Author/Title` +
          `&$expand=ProcessManager,Categoria,Author` +
          `&$filter=Created ge datetime'${yearStart}' and Created le datetime'${yearEnd}'` +
          `&$orderby=Id desc` +
          `&$top=500`
      );

      const tickets: ITicketAdmonRow[] = (ticketsRaw ?? []).map((item: any) => {
        const modified: string = item.Modified ?? '';
        const modDate = modified ? new Date(modified) : new Date();
        return {
          id: item.Id,
          title: item.Title ?? '',
          processManager: item.ProcessManager?.Title ?? '',
          tipoTicket: item.TipoTicket ?? '',
          categoria: item.Categoria?.Title ?? 'Sin categoría',
          status: item.Status ?? '',
          prioridad: item.Prioridad ?? '',
          tiempoAsignacion: modified ? getTimeDifference(modified) : '',
          mes: MESES[modDate.getUTCMonth()],
          anio: modDate.getUTCFullYear(),
          solicita: item.Author?.Title ?? '',
        };
      });

      this.setState({ tickets, loading: false });
    } catch (e: any) {
      this.setState({ error: e?.message ?? 'Error al cargar los tickets.', loading: false });
    }
  }

  // ─── Sort ──────────────────────────────────────────────────────────────────

  private _onColumnHeaderClick = (_ev: React.MouseEvent<HTMLElement> | undefined, column: IColumn): void => {
    const key = column.fieldName as keyof ITicketAdmonRow;
    this.setState(prev => ({
      sortKey: key,
      isSortedDescending: prev.sortKey === key ? !prev.isSortedDescending : false,
    }));
  };

  private _getSortedAndFiltered(tickets: ITicketAdmonRow[]): ITicketAdmonRow[] {
    const { filterText, sortKey, isSortedDescending } = this.state;

    let result = tickets;

    if (filterText) {
      const lower = filterText.toLowerCase();
      result = result.filter(t =>
        String(t.id).includes(lower) ||
        t.title.toLowerCase().includes(lower) ||
        t.processManager.toLowerCase().includes(lower) ||
        t.tipoTicket.toLowerCase().includes(lower) ||
        t.categoria.toLowerCase().includes(lower) ||
        t.status.toLowerCase().includes(lower) ||
        t.prioridad.toLowerCase().includes(lower) ||
        t.mes.toLowerCase().includes(lower) ||
        t.solicita.toLowerCase().includes(lower)
      );
    }

    result = result.slice().sort((a, b) => {
      const aVal = a[sortKey];
      const bVal = b[sortKey];
      const cmp = typeof aVal === 'number' && typeof bVal === 'number'
        ? aVal - bVal
        : String(aVal).localeCompare(String(bVal), 'es', { sensitivity: 'base' });
      return isSortedDescending ? -cmp : cmp;
    });

    return result;
  }

  // ─── Columns with sort indicators ─────────────────────────────────────────

  private _buildColumns(): IColumn[] {
    const { sortKey, isSortedDescending } = this.state;
    return BASE_COLUMNS.map(col => ({
      ...col,
      isResizable: true,
      isSorted: col.fieldName === sortKey,
      isSortedDescending: col.fieldName === sortKey ? isSortedDescending : false,
      onColumnClick: this._onColumnHeaderClick,
      onRender: col.key === 'status' ? this._renderStatus : undefined,
    }));
  }

  private _renderStatus = (item: ITicketAdmonRow): JSX.Element => {
    const { onVerDetalle } = this.props;
    return (
      <Link
        href={`#detalle-ticket-${item.id}`}
        onClick={() => {
          if (onVerDetalle) onVerDetalle(item.id);
        }}
      >
        {item.status}
      </Link>
    );
  };

  // ─── Render ────────────────────────────────────────────────────────────────

  public render(): React.ReactElement<ITodosLosTicketsProps> {
    const { isOpen, onDismiss } = this.props;
    const { loading, error, tickets, filterText, selectedYear } = this.state;

    const visibleTickets = this._getSortedAndFiltered(tickets);

    return (
      <Panel
        isOpen={isOpen}
        onDismiss={onDismiss}
        headerText="Todos los tickets"
        closeButtonAriaLabel="Cerrar"
        type={PanelType.extraLarge}
      >
        <Stack tokens={{ childrenGap: 12 }}>
          {loading && (
            <Stack horizontalAlign="center" style={{ marginTop: 40 }}>
              <Spinner size={SpinnerSize.large} label="Cargando tickets..." />
            </Stack>
          )}

          {!loading && error && (
            <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
          )}

          {!loading && !error && (
            <>
              <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
                <Stack.Item grow>
                  <SearchBox
                    placeholder="Buscar en todos los tickets..."
                    value={filterText}
                    onChange={(_ev, val) => this.setState({ filterText: val ?? '' })}
                    onClear={() => this.setState({ filterText: '' })}
                  />
                </Stack.Item>
                <Dropdown
                  selectedKey={selectedYear}
                  options={YEAR_OPTIONS}
                  styles={{ root: { minWidth: 100 } }}
                  onChange={(_ev, option) => {
                    if (option) {
                      this.setState({ selectedYear: option.key as number }, () => {
                        this._loadTickets().catch(console.error);
                      });
                    }
                  }}
                />
              </Stack>

              {tickets.length === 0 && (
                <MessageBar messageBarType={MessageBarType.info}>
                  No se encontraron tickets registrados.
                </MessageBar>
              )}

              {tickets.length > 0 && visibleTickets.length === 0 && (
                <MessageBar messageBarType={MessageBarType.warning}>
                  Ningún registro coincide con la búsqueda.
                </MessageBar>
              )}

              {visibleTickets.length > 0 && (
                <>
                  <Text variant="small" style={{ color: '#605e5c' }}>
                    {visibleTickets.length} de {tickets.length} registro{tickets.length !== 1 ? 's' : ''}
                  </Text>
                  <DetailsList
                    items={visibleTickets}
                    columns={this._buildColumns()}
                    layoutMode={DetailsListLayoutMode.justified}
                    selectionMode={SelectionMode.none}
                    isHeaderVisible={true}
                  />
                </>
              )}
            </>
          )}
        </Stack>
      </Panel>
    );
  }
}
