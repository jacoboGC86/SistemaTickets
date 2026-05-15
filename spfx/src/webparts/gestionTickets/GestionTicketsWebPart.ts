import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneFieldType,
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'GestionTicketsWebPartStrings';
import GestionTickets from './components/GestionTickets';
import { IGestionTicketsProps } from './components/GestionTickets';
import ListSvc from '../../services/ListSvc';
import UserSvc from '../../services/UserSvc';

export interface IGestionTicketsWebPartProps {
  description: string;
  usarMenuPersonalizado: boolean;
  idPersona: number;
  suplantarId: number;
  suplantarName: string;
}

// ── Property pane PeoplePicker (React.createElement – no JSX in .ts) ────────────

interface ISuplantarPickerProps {
  initialName: string;
  onSelect: (id: number, name: string) => void;
}

interface ISuplantarPickerState {
  query: string;
  suggestions: { loginName: string; title: string }[];
}

class SuplantarPicker extends React.Component<ISuplantarPickerProps, ISuplantarPickerState> {
  constructor(props: ISuplantarPickerProps) {
    super(props);
    this.state = { query: props.initialName || '', suggestions: [] };
  }

  private _search = async (val: string): Promise<void> => {
    this.setState({ query: val });
    if (val.length < 3) { this.setState({ suggestions: [] }); return; }
    try {
      const results = await UserSvc.SearchUsers(val);
      this.setState({
        suggestions: (results || []).slice(0, 8).map((u: any) => ({
          loginName: u.Key ?? '',
          title: u.DisplayText ?? u.Description ?? '',
        })),
      });
    } catch { this.setState({ suggestions: [] }); }
  };

  private _select = async (s: { loginName: string; title: string }): Promise<void> => {
    this.setState({ query: s.title, suggestions: [] });
    try {
      const ensured = await UserSvc.EnsureUser(s.loginName);
      const id: number = ensured?.Id ?? (ensured as any)?.d?.Id ?? 0;
      this.props.onSelect(id, s.title);
    } catch { /* ignore */ }
  };

  private _clear = (): void => {
    this.setState({ query: '', suggestions: [] });
    this.props.onSelect(0, '');
  };

  public render(): React.ReactElement {
    const { query, suggestions } = this.state;
    return React.createElement('div', { style: { padding: '8px 0' } },
      React.createElement('label', { style: { display: 'block', marginBottom: 4, fontSize: 14, fontWeight: 600 } }, 'Suplantar usuario'),
      React.createElement('div', { style: { display: 'flex', gap: 4 } },
        React.createElement('input', {
          type: 'text',
          value: query,
          onChange: (e: React.ChangeEvent<HTMLInputElement>) => this._search(e.target.value),
          placeholder: 'Buscar persona...',
          style: { flex: 1, padding: '6px 8px', fontSize: 14, boxSizing: 'border-box', border: '1px solid #a19f9d', borderRadius: 2 },
        }),
        query ? React.createElement('button', {
          onClick: this._clear,
          title: 'Limpiar',
          style: { padding: '0 8px', cursor: 'pointer', border: '1px solid #a19f9d', borderRadius: 2, background: 'white', fontSize: 14 },
        }, '✕') : null
      ),
      suggestions.length > 0
        ? React.createElement('ul', {
            style: { listStyle: 'none', margin: '2px 0 0', padding: 0, border: '1px solid #ddd', borderRadius: 2, maxHeight: 160, overflowY: 'auto', background: 'white', boxShadow: '0 2px 6px rgba(0,0,0,0.15)', zIndex: 9999, position: 'relative' }
          },
          ...suggestions.map(s =>
            React.createElement('li', {
              key: s.loginName,
              onMouseDown: () => this._select(s),
              style: { padding: '6px 10px', cursor: 'pointer', fontSize: 13 },
            }, s.title)
          )
        )
        : null
    );
  }
}

export default class GestionTicketsWebPart extends BaseClientSideWebPart<IGestionTicketsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IGestionTicketsProps> = React.createElement(
      GestionTickets,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        isEditMode: this.displayMode === DisplayMode.Edit,
        usarMenuPersonalizado: this.properties.usarMenuPersonalizado,
        idPersona: this.properties.idPersona
        //suplantarId: this.properties.suplantarId ?? 0
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const relativeURL = this.context.pageContext.web.serverRelativeUrl;
      ListSvc.Init(this.context.spHttpClient, siteUrl, relativeURL);
      UserSvc.Init(this.context.spHttpClient, siteUrl, relativeURL, null);
      UserSvc.SetIdPersona(this.properties.idPersona ?? 0);
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneToggle('usarMenuPersonalizado', {
                  label: 'Usar menú personalizado',
                  onText: 'Activo',
                  offText: 'Inactivo'
                }),
                PropertyPaneTextField('idPersona', {
                  label: 'Id Persona Suplantar',
                  onGetErrorMessage: (value: string) => {
                    if (!value) return '';
                    return Number.isInteger(Number(value)) && Number(value) > 0
                      ? ''
                      : 'Ingresa el ID de la Persona en SharePoint (int).';
                  }
                }),
                /*
                {
                  type: PropertyPaneFieldType.Custom,
                  targetProperty: 'suplantarId',
                  properties: {
                    key: 'suplantarPicker',
                    onRender: (elem: HTMLElement) => {
                      ReactDom.render(
                        React.createElement(SuplantarPicker, {
                          initialName: this.properties.suplantarName || '',
                          onSelect: (id: number, name: string) => {
                            this.properties.suplantarId = id;
                            this.properties.suplantarName = name;
                            this.render();
                          },
                        }),
                        elem
                      );
                    },
                    onDispose: (elem: HTMLElement) => {
                      ReactDom.unmountComponentAtNode(elem);
                    },
                  } as IPropertyPaneCustomFieldProps,
                },
                */
                {
                  type: PropertyPaneFieldType.Custom,
                  targetProperty: '',
                  properties: {
                    key: 'urlSegmentsInfo',
                    onRender: (elem: HTMLElement) => {
                      elem.innerHTML = `
                        <p style="margin: 8px 0 4px;">Usa los siguientes segmentos de URL para mandar a llamar a los controles en tu página.</p>
                        <ul style="margin: 0; padding-left: 18px;">
                          <li>#nuevo-ticket</li>
                          <li>#ver-mis-tickets</li>
                          <li>#mis-tickets-por-aprobar</li>
                          <li>#mis-tickets-por-atender</li>
                        </ul>`;
                    }
                  } as IPropertyPaneCustomFieldProps
                }
              ]
            }
          ]
        }
      ]
    };
  }
}
