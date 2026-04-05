import * as React from 'react';
import styles from './GestionTickets.module.scss';
import { Nav, INavLink, INavStyles } from '@fluentui/react';
import { escape } from '@microsoft/sp-lodash-subset';
import NuevoTicket from './NuevoTicket';
import DetalleTicket from './DetalleTicket';

const navStyles: Partial<INavStyles> = {
  root: {
    width: 260,
    padding: '0',
  },
};

const navLinks: INavLink[] = [
  { name: 'Nuevo ticket', url: '#nuevo-ticket', key: 'nuevo-ticket' },
  { name: 'Ver mis tickets', url: '#ver-mis-tickets', key: 'ver-mis-tickets' },
  { name: 'Todos los tickets', url: '#todos-los-tickets', key: 'todos-los-tickets' },
  { name: 'Dashboard', url: '#dashboard', key: 'dashboard' },
  { name: 'Administración del sistema', url: '#administracion-del-sistema', key: 'administracion-del-sistema' }
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
}

export default class GestionTickets extends React.Component<IGestionTicketsProps, IGestionTicketsState> {
  constructor(props: IGestionTicketsProps) {
    super(props);
    this.state = { isNuevoTicket: false, isDetalleTicket: false };
  }

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
        <DetalleTicket isOpen={this.state.isDetalleTicket} onDismiss={() => this.setState({ isDetalleTicket: false })} />
      </section>
    );
  }

  private _onLinkClick = (ev?: React.MouseEvent<HTMLElement>, item?: INavLink): void => {
    ev?.preventDefault();
    if (item?.key === 'nuevo-ticket') {
      this.setState({ isNuevoTicket: true });
    }
  };
}
