import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export default class ContextSvc {
  private static pageContext: PageContext;
  private static baseContext: WebPartContext;

  public static Init(pc:PageContext, bc:WebPartContext) {
    this.pageContext = pc;
    this.baseContext = bc;
  }

  public static getAADInfo():string {
    return this.pageContext.aadInfo.tenantId._guid;
  }

  public static getWebURL():string {
    return this.pageContext.web.absoluteUrl;
  }

  public static getPageURL():string {
    return (window.location.href.split("?")[0]).split("#")[0];
  }

  public static getContext():WebPartContext {
    return this.baseContext;
  }
}