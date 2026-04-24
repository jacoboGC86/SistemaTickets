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
  Icon,
} from '@fluentui/react';
import ListSvc from '../../../services/ListSvc';
import UserSvc from '../../../services/UserSvc';

// ─── Types ────────────────────────────────────────────────────────────────────

interface ITicketPorAprobarRow {
  id: number;
  title: string;
  paso: string;
  tipoTicket: string;
  categoria: string;
  status: string;
  prioridad: string;
  tiempoAsignacion: string;
}

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

    if (hour < WORK_START) { cursor.setHours(WORK_START, 0, 0, 0); continue; }
    if (hour >= WORK_END)  { cursor.setDate(cursor.getDate() + 1); cursor.setHours(WORK_START, 0, 0, 0); continue; }

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
  const days  = Math.floor(hours / 10);
  const rem   = hours % 10;
  const parts: string[] = [];
  if (days > 0) parts.push(`${days} día(s) `);
  if (rem  > 0) parts.push(`${rem} hora(s)`);
  return parts.length > 0 ? parts.join(', ') : 'Menos de una hora';
}

export interface IMisTicketsPorAprobarProps {
  isOpen: boolean;
  onDismiss: () => void;
  onVerDetalle?: (id: number) => void;
}

interface IMisTicketsPorAprobarState {
  loading: boolean;
  error: string | null;
  tickets: ITicketPorAprobarRow[];
  filterText: string;
  sortKey: keyof ITicketPorAprobarRow;
  isSortedDescending: boolean;
}

// ─── Base column definitions ──────────────────────────────────────────────────

const BASE_COLUMNS: Pick<IColumn, 'key' | 'name' | 'fieldName' | 'minWidth' | 'maxWidth'>[] = [
  { key: 'id',         name: 'Folio',                fieldName: 'id',         minWidth: 50,  maxWidth: 70  },
  { key: 'title',      name: 'Ticket',               fieldName: 'title',      minWidth: 150, maxWidth: 250 },
  { key: 'paso',       name: 'Paso de aprobación',   fieldName: 'paso',       minWidth: 150, maxWidth: 250 },
  { key: 'tipoTicket', name: 'Tipo',                 fieldName: 'tipoTicket', minWidth: 100, maxWidth: 150 },
  { key: 'categoria',  name: 'Categoría',            fieldName: 'categoria',  minWidth: 150, maxWidth: 250 },
  { key: 'status',           name: 'Estatus',              fieldName: 'status',           minWidth: 100, maxWidth: 150 },
  { key: 'prioridad',        name: 'Prioridad',            fieldName: 'prioridad',        minWidth: 80,  maxWidth: 120 },
  { key: 'tiempoAsignacion', name: 'Tiempo de asignación', fieldName: 'tiempoAsignacion', minWidth: 140, maxWidth: 200 },
];

// ─── Component ────────────────────────────────────────────────────────────────

export default class MisTicketsPorAprobar extends React.Component<IMisTicketsPorAprobarProps, IMisTicketsPorAprobarState> {
  constructor(props: IMisTicketsPorAprobarProps) {
    super(props);
    this.state = {
      loading: false,
      error: null,
      tickets: [],
      filterText: '',
      sortKey: 'id',
      isSortedDescending: true,
    };
  }

  public componentDidUpdate(prevProps: IMisTicketsPorAprobarProps): void {
    if (this.props.isOpen && !prevProps.isOpen) {
      this.setState({ filterText: '', sortKey: 'id', isSortedDescending: true }, () => {
        this._loadTickets().catch(console.error);
      });
    }
  }

  private async _loadTickets(): Promise<void> {
    this.setState({ loading: true, error: null, tickets: [] });

    try {
      const currentUser = await UserSvc.GetCurrentUser();
      const userId: number = currentUser.Id;

      // 1. Find pending approvals assigned to the current user
      const aprobacionesRaw:any[] = await ListSvc.getItems(
        'Aprobaciones',
        undefined,
        `$select=Id,Title,TicketId&$filter=ResponsableId eq ${userId} and Resultado eq 'Pendiente' and Title ne 'Assigned'&$top=200`
      );

      if (!aprobacionesRaw || aprobacionesRaw.length === 0) {
        this.setState({ tickets: [], loading: false });
        return;
      }

      // 2. Build unique ticket ID list
      const allIds = aprobacionesRaw.map((a: any) => a.TicketId as number).filter(Boolean);
      const ticketIds: number[] = allIds.filter((id, idx) => allIds.indexOf(id) === idx);

      // Build a map from ticketId → paso (step name)
      const pasoMap: Record<number, string> = {};
      aprobacionesRaw.forEach((a: any) => {
        if (a.TicketId && !pasoMap[a.TicketId]) {
          pasoMap[a.TicketId] = a.Title ?? '';
        }
      });

      // 3. Fetch ticket details and categories in parallel
      const filterIds = ticketIds.map(id => `Id eq ${id}`).join(' or ');
      const [ticketsRaw, categoriasRaw] = await Promise.all([
        ListSvc.getItems(
          'Tickets',
          undefined,
          `$select=Id,Title,TipoTicket,Status,Prioridad,Modified,FechaAtencion,CategoriaId,Categoria/Id,Categoria/Title` +
            `&$expand=Categoria` +
            `&$filter=${filterIds}` +
            `&$orderby=Id desc` +
            `&$top=200`
        ),
        ListSvc.getItems(
          'Categorias',
          undefined,
          `$select=Id,Title,CategoriaPadre/LookupValue&$expand=CategoriaPadre`
        ),
      ]);

      const categoriasMap: Record<number, string> = {};
      if (Array.isArray(categoriasRaw)) {
        categoriasRaw.forEach((cat: any) => {
          const padre = cat.CategoriaPadre?.LookupValue ?? cat.CategoriaPadre ?? '';
          categoriasMap[cat.Id] = padre ? `${cat.Title} - ${padre}` : cat.Title;
        });
      }

      const tickets: ITicketPorAprobarRow[] = (ticketsRaw ?? []).map((item: any) => {
        const catId: number         = item.CategoriaId ?? item.Categoria?.Id ?? 0;
        const modified: string      = item.Modified ?? '';
        const fechaAtencion: string = item.FechaAtencion ?? '';
        const status: string        = item.Status ?? '';

        let tiempoAsignacion = '-';
        if (status === 'Cerrado' && fechaAtencion) {
          tiempoAsignacion = getBusinessTimeDifference(fechaAtencion, modified);
        } else if (status === 'Assigned' && !fechaAtencion) {
          tiempoAsignacion = getBusinessTimeDifference(modified);
        }

        return {
          id:               item.Id,
          title:            item.Title ?? '',
          paso:             pasoMap[item.Id] ?? '',
          tipoTicket:       item.TipoTicket ?? '',
          categoria:        catId && categoriasMap[catId] ? categoriasMap[catId] : item.Categoria?.Title ?? 'Sin categoría',
          status,
          prioridad:        item.Prioridad ?? '',
          tiempoAsignacion,
        };
      });

      this.setState({ tickets, loading: false });
    } catch (e: any) {
      this.setState({ error: e?.message ?? 'Error al cargar los tickets.', loading: false });
    }
  }

  // ─── Sort ──────────────────────────────────────────────────────────────────

  private _onColumnHeaderClick = (_ev: React.MouseEvent<HTMLElement> | undefined, column: IColumn): void => {
    const key = column.fieldName as keyof ITicketPorAprobarRow;
    this.setState(prev => ({
      sortKey: key,
      isSortedDescending: prev.sortKey === key ? !prev.isSortedDescending : false,
    }));
  };

  private _getSortedAndFiltered(tickets: ITicketPorAprobarRow[]): ITicketPorAprobarRow[] {
    const { filterText, sortKey, isSortedDescending } = this.state;

    let result = tickets;

    if (filterText) {
      const lower = filterText.toLowerCase();
      result = result.filter(t =>
        String(t.id).includes(lower) ||
        t.title.toLowerCase().includes(lower) ||
        t.paso.toLowerCase().includes(lower) ||
        t.tipoTicket.toLowerCase().includes(lower) ||
        t.categoria.toLowerCase().includes(lower) ||
        t.status.toLowerCase().includes(lower) ||
        t.prioridad.toLowerCase().includes(lower)
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
      onRender: col.key === 'status'
        ? this._renderStatus
        : col.key === 'tipoTicket'
          ? this._renderTipoTicket
          : col.key === 'prioridad'
            ? this._renderPrioridad
            : undefined,
    }));
  }

  private _renderPrioridad = (item: ITicketPorAprobarRow): JSX.Element => {
    const val = item.prioridad;
    const lower = val.toLowerCase();
    let bg = '#e1e1e1', border = '#999', color = '#323130';
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
        backgroundColor: bg, border: `1px solid ${border}`, color,
        borderRadius: 12, padding: '2px 10px', fontSize: 12, fontWeight: 600,
        display: 'inline-flex', alignItems: 'center', gap: 4, whiteSpace: 'nowrap',
      }}>
        {iconName && <Icon iconName={iconName} style={{ fontSize: 12 }} />}
        {val}
      </span>
    );
  };

  private _renderTipoTicket = (item: ITicketPorAprobarRow): JSX.Element => {
    const val = item.tipoTicket;
    const lower = val.toLowerCase();
    let bg = '#e1e1e1', border = '#999', color = '#323130';

    if (lower === 'request') {
      bg = '#cce5ff'; border = '#004085'; color = '#004085';
    } else if (lower === 'incidente' || lower === 'incident') {
      bg = '#f8d7da'; border = '#721c24'; color = '#721c24';
    } else if (lower === 'cambio' || lower === 'change') {
      bg = '#fff3cd'; border = '#856404'; color = '#856404';
    }

    return (
      <span style={{
        backgroundColor: bg, border: `1px solid ${border}`, color,
        borderRadius: 12, padding: '2px 10px', fontSize: 12, fontWeight: 600,
        display: 'inline-block', whiteSpace: 'nowrap',
      }}>
        {val}
      </span>
    );
  };

  private _renderStatus = (item: ITicketPorAprobarRow): JSX.Element => {
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

  public render(): React.ReactElement<IMisTicketsPorAprobarProps> {
    const { isOpen, onDismiss } = this.props;
    const { loading, error, tickets, filterText } = this.state;

    const visibleTickets = this._getSortedAndFiltered(tickets);

    return (
      <Panel
        isOpen={isOpen}
        onDismiss={onDismiss}
        headerText="Mis tickets por aprobar"
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
              <SearchBox
                placeholder="Buscar en tickets por aprobar..."
                value={filterText}
                onChange={(_ev, val) => this.setState({ filterText: val ?? '' })}
                onClear={() => this.setState({ filterText: '' })}
              />

              {tickets.length === 0 && (
                <MessageBar messageBarType={MessageBarType.info}>
                  No tiene tickets pendientes de aprobación.
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
