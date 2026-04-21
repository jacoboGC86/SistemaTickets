import * as React from 'react';
import styles from './GestionTickets.module.scss';
import { Nav, INavLink, INavStyles } from '@fluentui/react';
import { escape } from '@microsoft/sp-lodash-subset';
import NuevoTicket from './NuevoTicket';
import DetalleTicket from './DetalleTicket';
import MisTickets from './MisTickets';
import MisTicketsPorAprobar from './MisTicketsPorAprobar';
import TodosLosTickets from './TodosLosTickets';

const navStyles: Partial<INavStyles> = {
  root: {
    width: 260,
    padding: '0',
  },
};

const navLinks: INavLink[] = [
  { name: 'Nuevo ticket', url: '#nuevo-ticket', key: 'nuevo-ticket' },
  { name: 'Ver mis tickets', url: '#ver-mis-tickets', key: 'ver-mis-tickets' },
  { name: 'Mis tickets por aprobar', url: '#mis-tickets-por-aprobar', key: 'mis-tickets-por-aprobar' },
  { name: 'Todos los tickets', url: '#todos-los-tickets', key: 'todos-los-tickets' },
  { name: 'Dashboard', url: '#dashboard', key: 'dashboard' },
];

export interface IGestionTicketsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}

export interface IGestionTicketsState {
  isNuevoTicket: boolean;
  isDetalleTicket: boolean;
  isMisTickets: boolean;
  isMisTicketsPorAprobar: boolean;
  isTodosLosTickets: boolean;
  detalleTicketId: number | null;
}

export default class GestionTickets extends React.Component<IGestionTicketsProps, IGestionTicketsState> {
  constructor(props: IGestionTicketsProps) {
    super(props);
    this.state = { isNuevoTicket: false, isDetalleTicket: false, isMisTickets: false, isMisTicketsPorAprobar: false, isTodosLosTickets: false, detalleTicketId: null };
  }

  public componentDidMount(): void {
    this._handleHash();
    window.addEventListener('hashchange', this._handleHash);
  }

  public componentWillUnmount(): void {
    window.removeEventListener('hashchange', this._handleHash);
  }

  private _handleHash = (): void => {
    const match = window.location.hash.match(/^#detalle-ticket-(\d+)$/);
    if (match) {
      this.setState({ detalleTicketId: Number(match[1]), isDetalleTicket: true });
    }
  };

  public render(): React.ReactElement<IGestionTicketsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.gestionTickets} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.menuContainer}>
          <Nav groups={[{ links: navLinks }]} styles={navStyles} onLinkClick={this._onLinkClick} />
        </div>
        <NuevoTicket isOpen={this.state.isNuevoTicket} onDismiss={() => this.setState({ isNuevoTicket: false })} />
        <DetalleTicket
          isOpen={this.state.isDetalleTicket}
          onDismiss={() => {
            this.setState({ isDetalleTicket: false });
            history.pushState(null, '', window.location.pathname + window.location.search);
          }}
          ticketId={this.state.detalleTicketId}
        />
        <MisTickets
          isOpen={this.state.isMisTickets}
          onDismiss={() => this.setState({ isMisTickets: false })}
          onVerDetalle={(id) => this.setState({ detalleTicketId: id, isDetalleTicket: true })}
        />
        <MisTicketsPorAprobar
          isOpen={this.state.isMisTicketsPorAprobar}
          onDismiss={() => this.setState({ isMisTicketsPorAprobar: false })}
          onVerDetalle={(id) => this.setState({ detalleTicketId: id, isDetalleTicket: true })}
        />
        <TodosLosTickets
          isOpen={this.state.isTodosLosTickets}
          onDismiss={() => this.setState({ isTodosLosTickets: false })}
          onVerDetalle={(id) => this.setState({ detalleTicketId: id, isDetalleTicket: true })}
        />
      </section>
    );
  }

  private _onLinkClick = (ev?: React.MouseEvent<HTMLElement>, item?: INavLink): void => {
    ev?.preventDefault();
    if (item?.key === 'nuevo-ticket') {
      this.setState({ isNuevoTicket: true });
    } else if (item?.key === 'ver-mis-tickets') {
      this.setState({ isMisTickets: true });
    } else if (item?.key === 'mis-tickets-por-aprobar') {
      this.setState({ isMisTicketsPorAprobar: true });
    } else if (item?.key === 'todos-los-tickets') {
      this.setState({ isTodosLosTickets: true });
    }
  };
}
