/** Minimal Microsoft Graph client with throttling retries, paging and JSON batching. */
export const GRAPH_ROOT="https://graph.microsoft.com/v1.0";

export class GraphError extends Error {
  constructor(message:string,readonly status:number,readonly code:string) { super(message); }
}

type TokenProvider = () => Promise<string>;
type BatchRequest = { method:string; url:string; body?:unknown };

const sleep=(ms:number)=>new Promise((resolve)=>setTimeout(resolve,ms));

function friendlyMessage(status:number,code:string,message:string) {
  if (status===401) return "Your Microsoft 365 sign-in has expired. Please sign in again.";
  if (status===403 && /accessDenied/i.test(code)) return "Your Microsoft 365 account does not have permission for the VermiTrack SharePoint site. Ask the site owner to add you as a member.";
  if (status===404 && /itemNotFound/i.test(code)) return "The SharePoint site or list could not be found. Check the site address in Settings.";
  return message || `Microsoft 365 request failed (${status}).`;
}

export class Graph {
  constructor(private token:TokenProvider,private fetchImpl:typeof fetch=fetch) {}

  async request<T=unknown>(method:string,pathOrUrl:string,body?:unknown,headers:Record<string,string>={}):Promise<T> {
    const url=pathOrUrl.startsWith("https://") ? pathOrUrl : `${GRAPH_ROOT}${pathOrUrl}`;
    for (let attempt=0;;attempt++) {
      const response=await this.fetchImpl(url,{
        method,
        headers:{ Authorization:`Bearer ${await this.token()}`,...(body===undefined ? {} : { "Content-Type":"application/json" }),...headers },
        body:body===undefined ? undefined : JSON.stringify(body),
      });
      if ((response.status===429 || response.status===503 || response.status===504) && attempt<5) {
        const retryAfter=Number(response.headers.get("retry-after"));
        await sleep(Number.isFinite(retryAfter) && retryAfter>0 ? retryAfter*1000 : 1000*2**attempt);
        continue;
      }
      if (response.status===202 || response.status===204) return undefined as T;
      const text=await response.text();
      const payload=text ? JSON.parse(text) as { error?:{ code?:string; message?:string } } : {};
      if (!response.ok) {
        const code=payload.error?.code ?? "";
        throw new GraphError(friendlyMessage(response.status,code,payload.error?.message ?? ""),response.status,code);
      }
      return payload as T;
    }
  }

  /** Follows @odata.nextLink until every page is read. */
  async pages<T>(path:string):Promise<T[]> {
    const results:T[]=[];
    let next:string | undefined=path;
    while (next) {
      const page:{ value:T[]; "@odata.nextLink"?:string }=await this.request("GET",next);
      results.push(...page.value);
      next=page["@odata.nextLink"];
    }
    return results;
  }

  /**
   * Sends requests through Graph JSON batching (20 per call). Requests inside a
   * batch are chained with dependsOn so SharePoint assigns item ids in order.
   */
  async batch(requests:BatchRequest[]) {
    for (let start=0;start<requests.length;start+=20) {
      let pending=requests.slice(start,start+20).map((request,index)=>({ id:String(index+1),...request,headers:{ "Content-Type":"application/json" } }));
      for (let attempt=0;pending.length;attempt++) {
        const chained=pending.map((request,index)=>index ? { ...request,dependsOn:[pending[index-1].id] } : request);
        const result=await this.request<{ responses:Array<{ id:string; status:number; headers?:Record<string,string>; body?:{ error?:{ message?:string } } }> }>("POST","/$batch",{ requests:chained });
        const byId=new Map(result.responses.map((response)=>[response.id,response]));
        const failed=pending.filter((request)=>{ const status=byId.get(request.id)?.status ?? 500; return status>=300; });
        const hardFailure=failed.map((request)=>byId.get(request.id)).find((response)=>response && response.status!==429 && response.status!==424 && response.status<500);
        if (hardFailure) throw new GraphError(hardFailure.body?.error?.message ?? "A batch write to SharePoint failed.",hardFailure.status,"batchFailed");
        if (failed.length && attempt>=5) throw new GraphError("SharePoint is busy. Please try again in a few minutes.",429,"throttled");
        if (failed.length) await sleep(2000*2**attempt);
        pending=failed;
      }
    }
  }

  me() { return this.request<{ displayName:string | null; mail:string | null; userPrincipalName:string }>("GET","/me?$select=displayName,mail,userPrincipalName"); }

  sendMail(input:{ to:string[]; cc?:string[]; subject:string; html:string }) {
    const recipients=(list:string[] | undefined)=>(list ?? []).map((address)=>({ emailAddress:{ address } }));
    return this.request("POST","/me/sendMail",{ message:{ subject:input.subject,body:{ contentType:"HTML",content:input.html },toRecipients:recipients(input.to),ccRecipients:recipients(input.cc) },saveToSentItems:true });
  }
}
