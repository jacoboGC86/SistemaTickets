import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export default class GraphSvc {
  private static msGraphClientFactory:MSGraphClientFactory;

  public static Init(msGraphClientFactory:MSGraphClientFactory) {
      this.msGraphClientFactory = msGraphClientFactory;
  }

  public static SendAnEmailUsingMSGraph(emailId:string[], ccEmailId:string[], title:string, body:string, callback ?:() => void):Promise<any> {
    let toRecipients:any[] = [];

    emailId.forEach((mail) => {
      toRecipients.push({
        "emailAddress": {
          "address": mail
        }
      });
    });

    //Create Body for Email  
    let emailPostBody: any = {  
      "message": {  
        "subject": title,
        "body": {  
          "contentType": "HTML",  
          "content": body
        },  
        "toRecipients": toRecipients,
      }  
    };

    if(ccEmailId.length > 0) {
      let ccRecipients:any[] = [];
      
      ccEmailId.forEach((mail) => {
        ccRecipients.push({
          "emailAddress": {
            "address": mail
          }
        });
      });

      emailPostBody.message.ccRecipients = ccRecipients;
    }

    //Send Email uisng MS Graph  
    return this.msGraphClientFactory
            .getClient('3')
            .catch((ex:any) => {
              console.log(ex);
            })
            .then((client: MSGraphClientV3 | void): void => {
              if (!client) {
                console.log("MSGraphClientV3 is undefined.");
                return;
              }
              client.api('/me/sendMail')
                .post(emailPostBody)
                .then((response:any) => {
                  //console.log(response);
                  
                  if(callback !== undefined)
                    callback();
                })
                .catch((ex:any) => {
                  console.log(ex);
                });
            }) as Promise<any>;
  }
}