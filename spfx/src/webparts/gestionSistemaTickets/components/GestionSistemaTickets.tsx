import * as React from 'react';
import styles from './GestionSistemaTickets.module.scss';
import { Nav, INavLink, INavStyles } from '@fluentui/react';
import NuevaPlantilla from './NuevaPlantilla';

const navStyles: Partial<INavStyles> = {
  root: {
    width: 260,
    padding: '0',
  },
};

const navLinks: INavLink[] = [
  { name: 'Nueva plantilla', url: '#nueva-plantilla', key: 'nueva-plantilla' },
  { name: 'Consulta de plantillas', url: '#consulta-plantillas', key: 'consulta-plantillas' },
  { name: 'Nueva categoría', url: '#nueva-categoria', key: 'nueva-categoria' },
  { name: 'Consulta de categorías', url: '#consulta-categorias', key: 'consulta-categorias' },
  { name: 'Todos los tickets', url: '#todos-los-tickets', key: 'todos-los-tickets' },
  { name: 'Tickets cerrados', url: '#tickets-cerrados', key: 'tickets-cerrados' },
  { name: 'Base de conocimiento', url: '#base-de-conocimiento', key: 'base-de-conocimiento' },
];

export interface IGestionSistemaTicketsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}


export interface IGestionSistemaTicketsState {
  activeKey: string | null;
  isNuevaPlantilla: boolean;
}

export default class GestionSistemaTickets extends React.Component<IGestionSistemaTicketsProps, IGestionSistemaTicketsState> {
  constructor(props: IGestionSistemaTicketsProps) {
    super(props);
    this.state = { activeKey: null, isNuevaPlantilla: false };
  }

  public render(): React.ReactElement<IGestionSistemaTicketsProps> {
    const { hasTeamsContext } = this.props;

    return (
      <section className={`${styles.gestionSistemaTickets} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.menuContainer}>
          <Nav
            groups={[{ links: navLinks }]}
            styles={navStyles}
            selectedKey={this.state.activeKey ?? undefined}
            onLinkClick={this._onLinkClick}
          />
        </div>
        <NuevaPlantilla
          isOpen={this.state.isNuevaPlantilla}
          onDismiss={() => this.setState({ isNuevaPlantilla: false, activeKey: null })}
        />
      </section>
    );
  }

  private _onLinkClick = (ev?: React.MouseEvent<HTMLElement>, item?: INavLink): void => {
    ev?.preventDefault();
    if (item?.key === 'nueva-plantilla') {
      this.setState({ activeKey: item.key, isNuevaPlantilla: true });
    } else if (item?.key) {
      this.setState({ activeKey: item.key });
    }
  };
}
