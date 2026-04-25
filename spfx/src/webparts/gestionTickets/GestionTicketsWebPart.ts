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
        usarMenuPersonalizado: this.properties.usarMenuPersonalizado
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
