export default class PropertiesSvc {
  private static subscription: string;
  private static version: string;
  private static token: string;

  public static Init(s:string, v:string, t:string) {
    this.subscription = s;
    this.version = v;
    this.token = t;
  }

  public static setSubscription(s:string):void {
    this.subscription = s;
  }

  public static getSubscription():string {
    return this.subscription;
  }

  public static getVersion():number {
    const v:string[] = this.version.split(".");
    return parseFloat(`${v[0]}.${v[1]}`);
  }

  public static setToken(token:string):void {
    this.token = token;
  }

  public static getToken():string {
    return this.token;
  }
}