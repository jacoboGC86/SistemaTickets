import { WebPartContext } from '@microsoft/sp-webpart-base';
import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import PropertiesSvc from './PropertiesSvc';

const apiURL:string = "https://apichimalliweb.azurewebsites.net/";
const payURL:string = "https://api-gw.payclip.com/checkout";

export default class RestSvc {
    private static context: WebPartContext;
    

    public static Init(context:WebPartContext) {
        this.context = context;
    }

    public static getInfo(api:string, query?:string): Promise<any> {
        let url = apiURL + api + (query != "" ? ("?" + query) : "");

        const options:IHttpClientOptions = {
          headers:{
              'Accept': 'application/json;',
              'Content-type': 'application/json;',
              'Authorization': `Bearer ${ PropertiesSvc.getToken() }` 
          }
        };
        
        return this.context.httpClient.get(url, HttpClient.configurations.v1, options).then((response: HttpClientResponse) => {
                return response.json();
            }).catch((reason) => {
                window.localStorage.setItem("CA_error", reason);
                
                return null;
            }).then(jsonResponse => {
                return jsonResponse;
            }) as Promise<any>;
    }

    public static clipGetPayLink(body:string, token:string): Promise<any> {
      const options:IHttpClientOptions = {
        headers:{
            'Accept': 'application/vnd.com.payclip.v2+json',
            'Content-type': 'application/json;',
            'x-api-key': `Basic ${ token }`
        },
        body: body
      };

      return this.context.httpClient.post(payURL, HttpClient.configurations.v1, options).then((response: HttpClientResponse) => {
        if(response.ok){
          return response.json();
        } else
          return null;
      }).catch((reason) => {
          window.localStorage.setItem("CA_error", reason);

          return null;
      }).then(jsonResponse => {
          return jsonResponse;
      }) as Promise<any>;
    }

    public static postInfo(api:string, body:any, requireKey:boolean): Promise<any> {
        let url = apiURL + api + (requireKey ? "/No_pasar%C3%A1s_de_aqu%C3%AD%21" : "");

        const options:IHttpClientOptions = {
            headers:{
                'Accept': 'application/json;',
                'Content-type': 'application/json;',
                'Authorization': `Bearer ${ PropertiesSvc.getToken() }` 
            },
            body: JSON.stringify(body)
        };

        return this.context.httpClient.post(url, HttpClient.configurations.v1, options).then((response: HttpClientResponse) => {
          if(response.ok){
            return response.json();
          } else
            return null;
        }).catch((reason) => {
            window.localStorage.setItem("CA_error", reason);

            return null;
        }).then(jsonResponse => {
            return jsonResponse;
        }) as Promise<any>;
    }

    public static putInfo(api:string, id:number, body:any): Promise<any> {
        let url = apiURL + api;
        
        const options:IHttpClientOptions = {
            headers:{
                'Accept': 'application/json;',
                'Content-Type': 'application/json;',
                'Authorization': `Bearer ${ PropertiesSvc.getToken() }` 
            },
            body: JSON.stringify(body)
        };

        return this.context.httpClient.post(url, HttpClient.configurations.v1, options).then((response: HttpClientResponse) => {
          if(response.ok){
            return response.json();
          } else
            return null;
        }).catch((reason) => {
            window.localStorage.setItem("CA_error", reason);
            
            return null;
        }).then(jsonResponse => {
            return jsonResponse;
        }) as Promise<any>;
    }
}