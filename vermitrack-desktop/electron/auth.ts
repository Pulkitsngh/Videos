/**
 * Microsoft 365 sign-in for the desktop app (MSAL public client with the
 * system browser and a localhost redirect). The token cache is encrypted with
 * the operating system's keychain through Electron safeStorage.
 */
import fs from "node:fs";
import { PublicClientApplication, InteractionRequiredAuthError, type AccountInfo, type ICachePlugin } from "@azure/msal-node";

export const BASE_SCOPES=["User.Read","Sites.ReadWrite.All","Mail.Send"];
export const SETUP_SCOPES=["Sites.Manage.All"];

type Crypto = { encrypt(text:string):Buffer; decrypt(data:Buffer):string; available():boolean };

function cachePlugin(file:string,crypto:Crypto):ICachePlugin {
  return {
    async beforeCacheAccess(context) {
      if (!fs.existsSync(file)) return;
      try {
        const raw=fs.readFileSync(file);
        context.tokenCache.deserialize(crypto.available() ? crypto.decrypt(raw) : raw.toString("utf8"));
      } catch { fs.rmSync(file,{ force:true }); }
    },
    async afterCacheAccess(context) {
      if (!context.cacheHasChanged) return;
      const text=context.tokenCache.serialize();
      fs.writeFileSync(file,crypto.available() ? crypto.encrypt(text) : Buffer.from(text,"utf8"),{ mode:0o600 });
    },
  };
}

const page=(title:string,body:string)=>`<!doctype html><html><head><meta charset="utf-8"><title>${title}</title></head><body style="font-family:Segoe UI,system-ui,sans-serif;background:#f5f4ef;color:#18332b;display:grid;place-items:center;height:100vh;margin:0"><div style="text-align:center"><h1 style="font-size:22px">${title}</h1><p>${body}</p></div></body></html>`;

export class Auth {
  private pca:PublicClientApplication;
  private account:AccountInfo | null=null;

  constructor(readonly clientId:string,readonly tenantId:string,cacheFile:string,crypto:Crypto,private openBrowser:(url:string)=>Promise<void>) {
    this.pca=new PublicClientApplication({
      auth:{ clientId,authority:`https://login.microsoftonline.com/${tenantId || "organizations"}` },
      cache:{ cachePlugin:cachePlugin(cacheFile,crypto) },
    });
  }

  async restore() {
    const accounts=await this.pca.getTokenCache().getAllAccounts();
    this.account=accounts[0] ?? null;
    return this.account;
  }

  get signedIn() { return Boolean(this.account); }
  get username() { return this.account?.username ?? null; }

  async signIn(extraScopes:string[]=[]) {
    const result=await this.pca.acquireTokenInteractive({
      scopes:[...BASE_SCOPES,...extraScopes],
      prompt:"select_account",
      openBrowser:this.openBrowser,
      successTemplate:page("Signed in to VermiTrack","You can close this tab and return to the VermiTrack app."),
      errorTemplate:page("Sign-in did not complete","Return to VermiTrack and try again. Details: {{error}}"),
    });
    this.account=result.account;
    return result.account;
  }

  /** Returns a Graph access token, prompting in the browser only if Microsoft requires it. */
  async token(scopes:string[]=BASE_SCOPES) {
    if (!this.account) throw Object.assign(new Error("Sign in with your Microsoft 365 account."),{ status:401,code:"UNAUTHENTICATED" });
    try {
      return (await this.pca.acquireTokenSilent({ account:this.account,scopes })).accessToken;
    } catch (error) {
      if (!(error instanceof InteractionRequiredAuthError)) throw error;
      const result=await this.pca.acquireTokenInteractive({ scopes,loginHint:this.account.username,openBrowser:this.openBrowser,
        successTemplate:page("Permission granted","You can close this tab and return to VermiTrack."),errorTemplate:page("Permission was not granted","Details: {{error}}") });
      this.account=result.account;
      return result.accessToken;
    }
  }

  async signOut() {
    const cache=this.pca.getTokenCache();
    for (const account of await cache.getAllAccounts()) await cache.removeAccount(account);
    this.account=null;
  }
}
