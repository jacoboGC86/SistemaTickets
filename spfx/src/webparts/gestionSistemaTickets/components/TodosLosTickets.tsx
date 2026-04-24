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
  Icon,
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
  filterPrioridad: string;
  filterTipo: string;
  filterStatus: string;
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

const PRIORIDAD_OPTIONS: IDropdownOption[] = [
  { key: '', text: 'Todas las prioridades' },
  { key: 'Baja', text: 'Baja' },
  { key: 'Media', text: 'Media' },
  { key: 'Alta', text: 'Alta' },
];

const TIPO_OPTIONS: IDropdownOption[] = [
  { key: '', text: 'Todos los tipos' },
  { key: 'Request', text: 'Request' },
  { key: 'Incident', text: 'Incidente' },
  { key: 'Change', text: 'Cambio' },
];

const STATUS_OPTIONS: IDropdownOption[] = [
  { key: '', text: 'Todos' },
  { key: 'En Aprobación', text: 'En Aprobación' },
  { key: 'Assigned', text: 'Asignado' },
  { key: 'Cerrado', text: 'Cerrado' },
];

const MESES = [
  'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
  'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre',
];

// ─── Helpers ──────────────────────────────────────────────────────────────────

function getBusinessHoursDifference(startStr: string, endStr?: string): number {
  const WORK_START = 8;
  const WORK_END   = 18;

  const current = new Date(startStr);
  const end     = endStr ? new Date(endStr) : new Date();

  if (current >= end) return 0;

  let totalMs = 0;
  const cursor = new Date(current);

  while (cursor < end) {
    const day = cursor.getDay();

    if (day === 0 || day === 6) {
      const daysToMonday = day === 0 ? 1 : 2;
      cursor.setDate(cursor.getDate() + daysToMonday);
      cursor.setHours(WORK_START, 0, 0, 0);
      continue;
    }

    const hour = cursor.getHours();

    if (hour < WORK_START) {
      cursor.setHours(WORK_START, 0, 0, 0);
      continue;
    }

    if (hour >= WORK_END) {
      cursor.setDate(cursor.getDate() + 1);
      cursor.setHours(WORK_START, 0, 0, 0);
      continue;
    }

    const endOfWorkDay = new Date(cursor);
    endOfWorkDay.setHours(WORK_END, 0, 0, 0);

    const blockEnd = end < endOfWorkDay ? end : endOfWorkDay;
    totalMs += blockEnd.getTime() - cursor.getTime();
    cursor.setTime(blockEnd.getTime());
  }

  return Math.floor(totalMs / (1000 * 60 * 60));
}

function getBusinessTimeDifference(startStr: string, endStr?: string): string {
  const hours = getBusinessHoursDifference(startStr, endStr);
  const workHoursPerDay = 10;
  const days = Math.floor(hours / workHoursPerDay);
  const remainingHours = hours % workHoursPerDay;

  const parts: string[] = [];
  if (days > 0) parts.push(`${days} día(s) `);
  if (remainingHours > 0) parts.push(`${remainingHours} hora(s)`);

  return parts.length > 0 ? parts.join(', ') : 'Menos de una hora';
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
      filterPrioridad: '',
      filterTipo: '',
      filterStatus: '',
      sortKey: 'id',
      isSortedDescending: true,
      selectedYear: CURRENT_YEAR,
    };
  }

  public componentDidUpdate(prevProps: ITodosLosTicketsProps): void {
    if (this.props.isOpen && !prevProps.isOpen) {
      this.setState({ filterText: '', filterPrioridad: '', filterTipo: '', filterStatus: '', sortKey: 'id', isSortedDescending: true, selectedYear: CURRENT_YEAR }, () => {
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
        `$select=Id,Title,TipoTicket,Status,Prioridad,Modified,FechaAtencion,` +
          `ProcessManager/Title,Categoria/Title,Author/Title` +
          `&$expand=ProcessManager,Categoria,Author` +
          `&$filter=Created ge datetime'${yearStart}' and Created le datetime'${yearEnd}'` +
          `&$orderby=Id desc` +
          `&$top=500`
      );

      const tickets: ITicketAdmonRow[] = (ticketsRaw ?? []).map((item: any) => {
        const modified: string     = item.Modified ?? '';
        const fechaAtencion: string = item.FechaAtencion ?? '';
        const status: string        = item.Status ?? '';
        const modDate = modified ? new Date(modified) : new Date();

        let tiempoAsignacion = '-';
        if (status === 'Cerrado' && fechaAtencion) {
          tiempoAsignacion = getBusinessTimeDifference(fechaAtencion, modified);
        } else if (status === 'Assigned' && !fechaAtencion) {
          tiempoAsignacion = getBusinessTimeDifference(modified);
        }

        return {
          id: item.Id,
          title: item.Title ?? '',
          processManager: item.ProcessManager?.Title ?? '',
          tipoTicket: item.TipoTicket ?? '',
          categoria: item.Categoria?.Title ?? 'Sin categoría',
          status,
          prioridad: item.Prioridad ?? '',
          tiempoAsignacion,
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
    const { filterText, filterPrioridad, filterTipo, filterStatus, sortKey, isSortedDescending } = this.state;

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

    if (filterPrioridad) {
      result = result.filter(t => t.prioridad.toLowerCase() === filterPrioridad.toLowerCase());
    }

    if (filterTipo) {
      result = result.filter(t => t.tipoTicket.toLowerCase() === filterTipo.toLowerCase());
    }

    if (filterStatus) {
      if (filterStatus === 'En Aprobación') {
        result = result.filter(t => t.status !== 'Cerrado' && t.status !== 'Assigned');
      } else {
        result = result.filter(t => t.status === filterStatus);
      }
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
      onRender: col.key === 'status'
        ? this._renderStatus
        : col.key === 'tipoTicket'
          ? this._renderTipoTicket
          : col.key === 'prioridad'
            ? this._renderPrioridad
            : undefined,
    }));
  }

  private _renderPrioridad = (item: ITicketAdmonRow): JSX.Element => {
    const val = item.prioridad;
    const lower = val.toLowerCase();
    let bg = '#e1e1e1';
    let border = '#999';
    let color = '#323130';
    let iconName: string | undefined;

    if (lower === 'baja') {
      bg = '#d4edda'; border = '#155724'; color = '#155724';
    } else if (lower === 'media') {
      bg = '#ffe0b2'; border = '#e65100'; color = '#bf360c';
      iconName = 'ShieldAlert';
    } else if (lower === 'alta') {
      bg = '#ffcdd2'; border = '#b71c1c'; color = '#7f0000';
      iconName = 'AlertSolid';
    }

    return (
      <span style={{
        backgroundColor: bg,
        border: `1px solid ${border}`,
        color,
        borderRadius: 12,
        padding: '2px 10px',
        fontSize: 12,
        fontWeight: 600,
        display: 'inline-flex',
        alignItems: 'center',
        gap: 4,
        whiteSpace: 'nowrap',
      }}>
        {iconName && <Icon iconName={iconName} style={{ fontSize: 12 }} />}
        {val}
      </span>
    );
  };

  private _renderTipoTicket = (item: ITicketAdmonRow): JSX.Element => {
    const val = item.tipoTicket;
    const lower = val.toLowerCase();
    let bg = '#e1e1e1';
    let border = '#999';
    let color = '#323130';

    if (lower === 'request') {
      bg = '#cce5ff'; border = '#004085'; color = '#004085';
    } else if (lower === 'incidente' || lower === 'incident') {
      bg = '#f8d7da'; border = '#721c24'; color = '#721c24';
    } else if (lower === 'cambio' || lower === 'change') {
      bg = '#fff3cd'; border = '#856404'; color = '#856404';
    }

    return (
      <span style={{
        backgroundColor: bg,
        border: `1px solid ${border}`,
        color,
        borderRadius: 12,
        padding: '2px 10px',
        fontSize: 12,
        fontWeight: 600,
        display: 'inline-block',
        whiteSpace: 'nowrap',
      }}>
        {val}
      </span>
    );
  };

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
    const { loading, error, tickets, filterText, filterPrioridad, filterTipo, filterStatus, selectedYear } = this.state;

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
                  selectedKey={filterPrioridad}
                  options={PRIORIDAD_OPTIONS}
                  styles={{ root: { minWidth: 160 } }}
                  onChange={(_ev, option) => this.setState({ filterPrioridad: option ? option.key as string : '' })}
                />
                <Dropdown
                  selectedKey={filterTipo}
                  options={TIPO_OPTIONS}
                  styles={{ root: { minWidth: 150 } }}
                  onChange={(_ev, option) => this.setState({ filterTipo: option ? option.key as string : '' })}
                />
                <Dropdown
                  selectedKey={filterStatus}
                  options={STATUS_OPTIONS}
                  styles={{ root: { minWidth: 150 } }}
                  onChange={(_ev, option) => this.setState({ filterStatus: option ? option.key as string : '' })}
                />
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
