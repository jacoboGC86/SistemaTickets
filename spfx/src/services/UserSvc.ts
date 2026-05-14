import { SPHttpClient } from '@microsoft/sp-http';
import ListSvc from './ListSvc';
import IUser from '../models/IUser';

export default class UserSvc {
  public static httpClient: SPHttpClient;
  private static site: string;
  private static hasAdminPermission:any;
  private static idPersona: number = 0;

  public static Init(httpClient:SPHttpClient, site:string, relativeURL:string, hasAdminPermission:any) {
    this.httpClient = httpClient;
    this.site = site;
    this.hasAdminPermission = hasAdminPermission;
  }

  public static SetIdPersona(id: number): void {
    this.idPersona = id;
  }

  public static async GetCurrentUser(): Promise<IUser> {
    if (this.idPersona > 0) {
      return this.GetUserById(this.idPersona);
    }
    let user:IUser = await (await this.httpClient.get(`${this.site}/_api/web/currentuser`, SPHttpClient.configurations.v1)).json();

    return user;
  }

  public static async GetPropertiesUser(loginname?:string): Promise<any> {
    const url:string = loginname !== undefined ? `${this.site}/_api/SP.UserProfiles.PeopleManager/GetUserProfilePropertyFor(@v)?@v='${ encodeURIComponent(loginname) }'` : `${this.site}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`;

    const response:IUser = await (await this.httpClient.get(url, SPHttpClient.configurations.v1)).json();

    return response;
  }

  public static async GetUserByLoginname(loginname:string):Promise<any> {
    let response = await (await this.httpClient.get(`${this.site}/_api/web/siteusers(@v)?@v='${ encodeURIComponent(loginname) }'`, SPHttpClient.configurations.v1)).json();

    return response;
  }

  public static async GetUserById(id:number):Promise<any> {
    ///_api/web/getUserById(userId)
    let response = await (await this.httpClient.get(`${this.site}/_api/web/getUserById(${ id })`, SPHttpClient.configurations.v1)).json();

    return response;
  }

  public static async GetSiteUserInfoListId():Promise<any> {
    let response = await (await this.httpClient.get(`${ this.site }/_api/Web/SiteUserInfoList?$select=Id`, SPHttpClient.configurations.v1)).json();

    return response;
  }

  public static async SearchUsers(term: string): Promise<any[]> {
    const body = JSON.stringify({
      queryParams: {
        __metadata: { type: 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' },
        AllowEmailAddresses: true,
        AllowMultipleEntities: false,
        AllowOnlyEmailAddresses: false,
        //CurrentUserValue: '',
        MaximumEntitySuggestions: 10,
        PrincipalSource: 15,
        PrincipalType: 1,
        QueryString: term,
      }
    });

    const response = await (await this.httpClient.post(
      `${ this.site }/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'odata-version': '',
        },
        body: body,
      }
    )).json();

    const raw = response?.d?.ClientPeoplePickerSearchUser ?? response?.value;
    if (typeof raw === 'string') return JSON.parse(raw);
    return raw || [];
  }

  public static async EnsureUser(loginName: string): Promise<any> {
    const response = await (await this.httpClient.post(
      `${ this.site }/_api/web/ensureuser`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'odata-version': '',
        },
        body: JSON.stringify({ logonName: loginName }),
      }
    )).json();

    return response?.d ?? response;
  }
}