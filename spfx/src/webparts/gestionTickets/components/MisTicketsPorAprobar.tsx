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
  { key: 'status',     name: 'Estatus',              fieldName: 'status',     minWidth: 100, maxWidth: 150 },
  { key: 'prioridad',  name: 'Prioridad',            fieldName: 'prioridad',  minWidth: 80,  maxWidth: 120 },
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
        `$select=Id,Title,TicketId&$filter=ResponsableId eq ${userId} and Resultado eq 'Pendiente'&$top=200`
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
          `$select=Id,Title,TipoTicket,Status,Prioridad,CategoriaId,Categoria/Id,Categoria/Title` +
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
        const catId: number = item.CategoriaId ?? item.Categoria?.Id ?? 0;
        return {
          id: item.Id,
          title: item.Title ?? '',
          paso: pasoMap[item.Id] ?? '',
          tipoTicket: item.TipoTicket ?? '',
          categoria: catId && categoriasMap[catId] ? categoriasMap[catId] : item.Categoria?.Title ?? 'Sin categoría',
          status: item.Status ?? '',
          prioridad: item.Prioridad ?? '',
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
      onRender: col.key === 'status' ? this._renderStatus : undefined,
    }));
  }

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
