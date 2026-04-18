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
import UserSvc from '../../../services/UserSvc';

// ─── Types ────────────────────────────────────────────────────────────────────

interface ITicketRow {
  id: number;
  title: string;
  tipoTicket: string;
  categoria: string;
  status: string;
  prioridad: string;
}

export interface IMisTicketsProps {
  isOpen: boolean;
  onDismiss: () => void;
  onVerDetalle?: (id: number) => void;
}

interface IMisTicketsState {
  loading: boolean;
  error: string | null;
  tickets: ITicketRow[];
  filterText: string;
  sortKey: keyof ITicketRow;
  isSortedDescending: boolean;
  selectedYear: number;
}

// ─── Year options ─────────────────────────────────────────────────────────────

const CURRENT_YEAR = new Date().getFullYear();
const YEAR_OPTIONS: IDropdownOption[] = Array.from(
  { length: CURRENT_YEAR - 2025 + 1 },
  (_, i) => { const y = 2025 + i; return { key: y, text: String(y) }; }
);

// ─── Base column definitions ──────────────────────────────────────────────────

const BASE_COLUMNS: Pick<IColumn, 'key' | 'name' | 'fieldName' | 'minWidth' | 'maxWidth'>[] = [
  { key: 'id',         name: 'Folio',     fieldName: 'id',         minWidth: 50,  maxWidth: 70  },
  { key: 'title',      name: 'Ticket',    fieldName: 'title',      minWidth: 150, maxWidth: 250 },
  { key: 'tipoTicket', name: 'Tipo',      fieldName: 'tipoTicket', minWidth: 100, maxWidth: 150 },
  { key: 'categoria',  name: 'Categoría', fieldName: 'categoria',  minWidth: 150, maxWidth: 250 },
  { key: 'status',     name: 'Estatus',   fieldName: 'status',     minWidth: 100, maxWidth: 150 },
  { key: 'prioridad',  name: 'Prioridad', fieldName: 'prioridad',  minWidth: 80,  maxWidth: 120 },
];

// ─── Component ────────────────────────────────────────────────────────────────

export default class MisTickets extends React.Component<IMisTicketsProps, IMisTicketsState> {
  constructor(props: IMisTicketsProps) {
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

  public componentDidUpdate(prevProps: IMisTicketsProps): void {
    if (this.props.isOpen && !prevProps.isOpen) {
      this.setState({ filterText: '', sortKey: 'id', isSortedDescending: true, selectedYear: CURRENT_YEAR }, () => {
        this._loadTickets().catch(console.error);
      });
    }
  }

  private async _loadTickets(): Promise<void> {
    this.setState({ loading: true, error: null, tickets: [] });

    try {
      const currentUser = await UserSvc.GetCurrentUser();
      const userId: number = currentUser.Id;
      const { selectedYear } = this.state;
      const yearStart = `${selectedYear}-01-01T00:00:00Z`;
      const yearEnd   = `${selectedYear}-12-31T23:59:59Z`;

      const [ticketsRaw, categoriasRaw] = await Promise.all([
        ListSvc.getItems(
          'Tickets',
          undefined,
          `$select=Id,Title,TipoTicket,Status,Prioridad,CategoriaId,Categoria/Id,Categoria/Title` +
            `&$expand=Categoria` +
            `&$filter=AuthorId eq ${userId} and Created ge datetime'${yearStart}' and Created le datetime'${yearEnd}'` +
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

      const tickets: ITicketRow[] = (ticketsRaw ?? []).map((item: any) => {
        const catId: number = item.CategoriaId ?? item.Categoria?.Id ?? 0;
        return {
          id: item.Id,
          title: item.Title ?? '',
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
    const key = column.fieldName as keyof ITicketRow;
    this.setState(prev => ({
      sortKey: key,
      isSortedDescending: prev.sortKey === key ? !prev.isSortedDescending : false,
    }));
  };

  private _getSortedAndFiltered(tickets: ITicketRow[]): ITicketRow[] {
    const { filterText, sortKey, isSortedDescending } = this.state;

    let result = tickets;

    if (filterText) {
      const lower = filterText.toLowerCase();
      result = result.filter(t =>
        String(t.id).includes(lower) ||
        t.title.toLowerCase().includes(lower) ||
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

  private _renderStatus = (item: ITicketRow): JSX.Element => {
    const { onVerDetalle } = this.props;
    return (
      <Link
        href={`#detalle-ticket-${item.id}`}
        onClick={(ev) => {
          ev.preventDefault();
          if (onVerDetalle) onVerDetalle(item.id);
        }}
      >
        {item.status}
      </Link>
    );
  };

  // ─── Render ────────────────────────────────────────────────────────────────

  public render(): React.ReactElement<IMisTicketsProps> {
    const { isOpen, onDismiss } = this.props;
    const { loading, error, tickets, filterText, selectedYear } = this.state;

    const visibleTickets = this._getSortedAndFiltered(tickets);

    return (
      <Panel
        isOpen={isOpen}
        onDismiss={onDismiss}
        headerText="Mis Tickets"
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
                    placeholder="Buscar en mis tickets..."
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
