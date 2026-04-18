import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';

export default class ListSvc {
  private static httpClient: SPHttpClient;
  private static site: string;
  private static relativeURL:string;

  public static Init(httpClient:SPHttpClient, site:string, relativeURL:string):void {
    this.httpClient = httpClient;
    this.site = site;
    this.relativeURL = relativeURL;
  }

  public static getRelativeSiteURL() : string {
    return this.relativeURL;
  }

  public static async getAllLists() {
    return await (await this.httpClient.get(`${ this.site }/_api/web/lists`, SPHttpClient.configurations.v1)).json();
  }

  public static async getListById(listId:string):Promise<any> {
    return await (await this.httpClient.get(`${ this.site }/_api/web/lists('${ listId }')`, SPHttpClient.configurations.v1)).json();
  }

  public static async getList(list:string, query?:string):Promise<any> {
    const consulta = (query !== undefined) ? "?" + query : "";

    return await (await this.httpClient.get(`${ this.site }/_api/web/lists/GetByTitle('${ list }')${ consulta }`, SPHttpClient.configurations.v1)).json();
  }

  public static async getListItemEntityTypeFullName(list:string):Promise<string> {
    const response = await (await this.httpClient.get(`${ this.site }/_api/web/lists/GetByTitle('${ list }')?$select=ListItemEntityTypeFullName`, SPHttpClient.configurations.v1)).json();

    return response.ListItemEntityTypeFullName || response?.d?.ListItemEntityTypeFullName;
  }

  public static async getListFields(list:string):Promise<any> {
    return await (await this.httpClient.get(`${ this.site }/_api/web/lists/GetByTitle('${ list }')/Fields`, SPHttpClient.configurations.v1)).json();
  }

  public static async getFieldsByListId(listId:string):Promise<any> {
    return await (await this.httpClient.get(`${ this.site }/_api/web/lists('${ listId }')/Fields`, SPHttpClient.configurations.v1)).json();
  }

  public static async getContentTypesByListId(listId:string):Promise<any> {
    return await (await this.httpClient.get(`${ this.site }/_api/web/lists('${ listId }')/ContentTypes`, SPHttpClient.configurations.v1)).json();
  }

  public static getFieldChoices(list:any, fieldName:string):string[] {
    let opciones:string[] = [];

    list.value.forEach(element => {
      if(element.InternalName === fieldName) {
        opciones = element.Choices;
      }
    });

    return opciones;
  }

  public static async getItemById(list:string, id:number, site?:string): Promise<any> {
    const querySite = site !== undefined ? site : this.site;
    const response = await (await this.httpClient.get(`${querySite}/_api/web/lists/GetByTitle('${ list }')/Items(${ id })`, SPHttpClient.configurations.v1)).json();

    return response;
  }

  public static async getItemByIdAndListId(listId:string, id:number, site?:string): Promise<any> {
    const querySite = site !== undefined ? site : this.site;
    const response = await (await this.httpClient.get(`${querySite}/_api/web/lists('${ listId }')/Items(${ id })`, SPHttpClient.configurations.v1)).json();

    return response;
  }

  public static async getCommentsByItemByIdAndListId(listId:string, id:number, site?:string): Promise<any> {
    const querySite = site !== undefined ? site : this.site;
    const response = await (await this.httpClient.get(`${querySite}/_api/web/lists('${ listId }')/Items(${ id })/Comments`, SPHttpClient.configurations.v1)).json();

    return response;
  }

  public static async getItems(list:string, site?:string, query?:string): Promise<any> {
    const consulta = (query !== undefined) ? "?" + query : "";
    const querySite = site !== undefined ? site : this.site;
    const response = await (await this.httpClient.get(`${querySite}/_api/web/lists/GetByTitle('${ list }')/Items${ consulta }`, SPHttpClient.configurations.v1)).json();

    return response.value;
  }

  public static async getItemsByListId(listId:string, site?:string, query?:string): Promise<any> {
    const consulta = (query !== undefined) ? "?" + query : "";
    const querySite = site !== undefined ? site : this.site;
    const response = await (await this.httpClient.get(`${querySite}/_api/web/lists('${ listId }')/Items${ consulta }`, SPHttpClient.configurations.v1)).json();

    return response.value;
  }

  public static async getItemsByPage(list?:string, site?:string, query?:string, nextPage?:string): Promise<any> {
    let response = null;
    
    if(nextPage === undefined){
      const consulta = (query !== undefined) ? "?" + query : "";
      const querySite = site !== undefined ? site : this.site;

      response = await (await this.httpClient.get(`${querySite}/_api/web/lists/GetByTitle('${ list }')/Items${ consulta }`, SPHttpClient.configurations.v1)).json();
    } else
      response = await (await this.httpClient.get(nextPage, SPHttpClient.configurations.v1)).json();

    return response;
  }

  public static async getItemsByPageAndListId(listId?:string, site?:string, query?:string, nextPage?:string): Promise<any> {
    let response = null;
    
    if(nextPage === undefined){
      const consulta = (query !== undefined) ? "?" + query : "";
      const querySite = site !== undefined ? site : this.site;

      response = await (await this.httpClient.get(`${querySite}/_api/web/lists('${ listId }')/Items${ consulta }`, SPHttpClient.configurations.v1)).json();
    } else
      response = await (await this.httpClient.get(nextPage, SPHttpClient.configurations.v1)).json();

    return response;
  }

  public static async getFilePropertiesByServerRelativeUrl(serverRelativePath:string, site?:string):Promise<any> {
    const querySite = site !== undefined ? site : this.site;
    const response = await (await this.httpClient.get(`${querySite}/_api/web/GetFileByServerRelativePath(decodedurl='/${ serverRelativePath }')/ListItemAllFields`, SPHttpClient.configurations.v1)).json();

    return response;
  }

  public static async copyFile(serverRelativePath:string, to:string):Promise<any> {
    const headers:ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'X-HTTP-Method': 'POST'
      }
    };

    const response = await this.httpClient.post(`${this.site}/_api/web/getfilebyserverrelativeurl('${ serverRelativePath }')/copyTo(strnewurl='${ to }',boverwrite=true)`, SPHttpClient.configurations.v1, headers);

    return response;
  }

  public static async getFileFromItemId(list:string, idItem:number): Promise<any> {
    const response = await (await this.httpClient.get(`${this.site}/_api/web/lists/GetByTitle('${ list }')/Items(${ idItem })/File`, SPHttpClient.configurations.v1)).json();

    return response;
  }

  public static async getAllFiles(library:string, query?:string):Promise<any> {
    const consulta = (query !== undefined) ? "?" + query : "";
    const response = await (await this.httpClient.get(`${this.site}/_api/web/GetFolderByServerRelativeUrl('/${ this.relativeURL }/${ library }')/Files${ consulta }`, SPHttpClient.configurations.v1)).json();

    return response;
  }

  public static async getFileVersions(list:string, id:number, site?:string):Promise<any> {
    const querySite = site !== undefined ? site : this.site;
    const response = await (await this.httpClient.get(`${querySite}/_api/web/lists/GetByTitle('${ list }')/Items('${ id.toString() }')/Versions`, SPHttpClient.configurations.v1)).json();

    return response;
  }

  public static async searchFile(term:string, docId:string, site?:string, listId?:string, ordenar?:boolean):Promise<any> {
    const querySite = site !== undefined ? site : this.site;
    const p = docId !== undefined && ordenar ? ` indexdocid>${ docId }` : '';
    const l = listId !== undefined && listId !== "" ? ` ListId:${ listId }` : '';
    const response = await (await this.httpClient.get(`${ querySite }/_api/search/query?querytext='(${ term })${ p }'&querytemplate='{searchterms} IsDocument:true${ l }'${ ordenar ? "&sortlist='[docid]:ascending'" : "" }&SelectProperties='Title,Author,Path,PictureThumbnailURL,ServerRedirectedEmbedURL,UniqueId'`, SPHttpClient.configurations.v1)).json();
    //let response = await (await this.httpClient.get(`${ querySite }/_api/search/query?querytext='(${ term })'&querytemplate='{searchterms} IsDocument:true${ l }'&SelectProperties='Title,Author,Path,PictureThumbnailURL,ServerRedirectedEmbedURL,UniqueId'`, SPHttpClient.configurations.v1)).json();
    
    return response;
  }

  public static async searchItem(term:string, page:string, site?:string, listId?:string):Promise<any> {
    const querySite = site !== undefined ? site : this.site;
    const p = "";// page > 1 ? ` indexdocid>${ (page - 1) * 10 }` : '';

    const response = await (await this.httpClient.get(`${ querySite }/_api/search/query?querytext='${ term + p }+AND+ListId:${ listId }'&querytemplate='{searchterms}'&sortlist='[docid]:ascending'&SelectProperties='Title,Author,UniqueId,Id'`, SPHttpClient.configurations.v1)).json();
    //let response = await (await this.httpClient.get(`${ querySite }/_api/search/query?querytext='${ term }'&querytemplate='{searchterms} IsDocument:true'&SelectProperties='Title,Author,Path,PictureThumbnailURL,ServerRedirectedEmbedURL,UniqueId'`, SPHttpClient.configurations.v1)).json();

    return response;
  }

  public static async postList(list:string, type:string):Promise<SPHttpClientResponse> {
    const headers:ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        "Content-Type": 'application/json;odata=verbose',
        'odata-version': '',
        'X-HTTP-Method': 'POST'
      },
      body: JSON.stringify({
        "__metadata": {"type": "SP.List"},
        "AllowContentTypes": true,
        "BaseTemplate": type === "List" ? 100 : 101,
        "ContentTypesEnabled": true,
        "Title": list,
        "Description": list,
      })
    };

    return await (await this.httpClient.post(`${this.site}/_api/web/lists`, SPHttpClient.configurations.v1, headers)).json();
  }

  public static async postListField(list:string, body:string, isLookup:boolean):Promise<SPHttpClientResponse> {
    const headers:ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        "Content-Type": 'application/json;odata=verbose',
        'odata-version': '',
        'X-HTTP-Method': 'POST'
      },
      body: body
    };

    return await (await this.httpClient.post(`${this.site}/_api/web/lists/GetByTitle('${ list }')/Fields${ isLookup ? '/addfield' : ''}`, SPHttpClient.configurations.v1, headers)).json();
  }

  public static async putListContentTypeByListId(listId:string, contentTypeId:string, body:string):Promise<SPHttpClientResponse> {
    const headers:ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json; odata.metadata=none',
        'Content-Type': 'application/json',
        'X-HTTP-Method': 'MERGE'
      },
      body: body
    };

    return await (await this.httpClient.post(`${this.site}/_api/web/lists(guid'${ listId }')/ContentTypes('${ contentTypeId }')`, SPHttpClient.configurations.v1, headers));
  }

  public static async postListItem(list:string, body:string):Promise<any> {
    const headers:ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        "Content-Type": 'application/json;odata=verbose',
        'odata-version': '',
        'X-HTTP-Method': 'POST'
      },
      body: body
    };

    return await (await this.httpClient.post(`${this.site}/_api/web/lists/GetByTitle('${ list }')/Items`, SPHttpClient.configurations.v1, headers)).json();
  }

  public static async postListItemAttachment(list:string, itemId:number, fileName:string, arrayBuffer:ArrayBuffer):Promise<any> {
    const safeFileName = encodeURIComponent(fileName).replace(/'/g, '%27');
    const headers:ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-length': `${ arrayBuffer.byteLength }`,
        'X-HTTP-Method': 'POST'
      },
      body: arrayBuffer
    };

    return await (await this.httpClient.post(`${this.site}/_api/web/lists/GetByTitle('${ list }')/Items(${ itemId })/AttachmentFiles/add(FileName='${ safeFileName }')`, SPHttpClient.configurations.v1, headers)).json();
  }
  public static async postListItemByListId(listId:string, body:string):Promise<SPHttpClientResponse> {
    const headers:ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        "Content-Type": 'application/json;odata=verbose',
        'odata-version': '',
        'X-HTTP-Method': 'POST'
      },
      body: body
    };

    return await (await this.httpClient.post(`${this.site}/_api/web/lists('${ listId }')/Items`, SPHttpClient.configurations.v1, headers)).json();
  }

  public static async postFolder(list:string, body:string):Promise<SPHttpClientResponse> {
    const headers:ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        "Content-Type": 'application/json;odata=verbose',
        'odata-version': '',
        'X-HTTP-Method': 'POST'
      },
      body: body
    };

    return await (await this.httpClient.post(`${this.site}/_api/web/folders`, SPHttpClient.configurations.v1, headers)).json();
  }

  public static async postFile(serverRelativePath:string, name:string, arrayBuffer:ArrayBuffer):Promise<SPHttpClientResponse> {
    const headers:ISPHttpClientOptions = {
      headers: {
        //'Accept': 'application/json;odata=verbose',
        "Content-length": `${ arrayBuffer.byteLength }`,
        'X-HTTP-Method': 'POST'
      },
      body: arrayBuffer
    };

    return await (await this.httpClient.post(`${this.site}/_api/web/getfolderbyserverrelativeurl('${ serverRelativePath }')/files/add(overwrite=true, url='${ name }')`, SPHttpClient.configurations.v1, headers)).json();
  }

  public static async mergeList(list:string, body:string):Promise<SPHttpClientResponse> {
    const headers:ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        "Content-Type": 'application/json;odata=verbose',
        'odata-version': '',
        'X-HTTP-Method': 'MERGE'
      },
      body: body
    };

    return await this.httpClient.post(`${this.site}/_api/web/lists/GetByTitle('${ list }')`, SPHttpClient.configurations.v1, headers);
  }

  public static async mergeListField(list:string, field:string, body:string):Promise<SPHttpClientResponse> {
    const headers:ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        "Content-Type": 'application/json;odata=verbose',
        'odata-version': '',
        'X-HTTP-Method': 'MERGE'
      },
      body: body
    };

    return await this.httpClient.post(`${this.site}/_api/web/lists/GetByTitle('${ list }')/Fields('${ field }')`, SPHttpClient.configurations.v1, headers);
  }

  public static async putListItem(list:string, idListItem:number, body:string):Promise<any> {
    const headers = {
      'Accept': 'application/json;odata=nometadata',
      "Content-Type": 'application/json;odata=verbose',
      'odata-version': '',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'PATCH'
    };

    const r1 = await this.httpClient.post(`${this.site}/_api/web/lists/GetByTitle('${ list }')/Items(${ idListItem })`, SPHttpClient.configurations.v1, {
      headers: headers, 
      body: body
    });

    return r1;
  }

  public static async putListItemByListId(listId:string, idListItem:number, body:string):Promise<any> {
    const headers = {
      'Accept': 'application/json;odata=nometadata',
      "Content-Type": 'application/json;odata=verbose',
      'odata-version': '',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'PATCH'
    };

    const r1 = await this.httpClient.post(`${this.site}/_api/web/lists('${ listId }')/Items(${ idListItem })`, SPHttpClient.configurations.v1, {
      headers: headers, 
      body: body
    });

    return r1;
  }

  public static async deleteListItem(list:string, idListItem:number):Promise<any> {
    const headers = {
      'Accept': 'application/json;odata=nometadata',
      "Content-Type": 'application/json;odata=verbose',
      'odata-version': '',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'DELETE'
    };

    return await (await this.httpClient.post(`${this.site}/_api/web/lists/GetByTitle('${ list }')/Items(${ idListItem })/recycle()`, SPHttpClient.configurations.v1, {
      headers: headers, 
    })).json();
  }
}