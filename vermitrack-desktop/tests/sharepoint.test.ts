// Exercises the SharePoint store and the API against an in-memory imitation of
// Microsoft Graph (lists, columns, items, paging, $batch, throttling, sendMail).
// Bundled by `npm test` with esbuild and run with node --test.
import assert from "node:assert/strict";
import { test } from "node:test";
import { handleApi, parseCsv, type ApiContext } from "../electron/api";
import { Graph } from "../electron/graph";
import seed from "../electron/seed.json";
import { SharePointStore, seedStore, NeedsProvisioningError, tableDefs } from "../electron/store";

type Item = { id:string; fields:Record<string,unknown> };
type List = { id:string; displayName:string; columns:Set<string>; items:Item[]; nextId:number };

function fakeGraph() {
  const lists:List[]=[]; const sent:unknown[]=[]; let throttleNext=0; let requests=0;
  const reply=(status:number,body?:unknown,headers:Record<string,string>={})=>new Response(body===undefined ? null : JSON.stringify(body),{ status,headers:{ "content-type":"application/json",...headers } });
  const notFound=()=>reply(404,{ error:{ code:"itemNotFound",message:"Item not found" } });
  const checkFields=(list:List,fields:Record<string,unknown>)=>{
    for (const key of Object.keys(fields)) if (key!=="Title" && !list.columns.has(key)) return `Field '${key}' is not recognized`;
    if (String(fields.Title ?? "").length>255) return "Title too long";
    for (const [key,value] of Object.entries(fields)) if (key!=="vtData" && typeof value==="string" && value.length>255) return `${key} exceeds 255 characters`;
    return null;
  };

  function route(method:string,url:URL,body:any):Response {
    requests++;
    const p=decodeURIComponent(url.pathname.replace(/^\/v1\.0/,""));
    if (p==="/me" && method==="GET") return reply(200,{ displayName:"Pulkit Singh",mail:"Pulkit.Singh@lnbgroup.com",userPrincipalName:"pulkit.singh@lnbgroup.com" });
    if (p==="/me/sendMail" && method==="POST") { assert.equal(body.message.body.contentType,"HTML"); sent.push(body); return reply(202); }
    if (p==="/sites/lnbgroup.sharepoint.com:/sites/VermiTrack" && method==="GET") return reply(200,{ id:"site-1",displayName:"VermiTrack",webUrl:"https://lnbgroup.sharepoint.com/sites/VermiTrack" });
    let m=p.match(/^\/sites\/site-1\/lists$/);
    if (m && method==="GET") return reply(200,{ value:lists.map((list)=>({ id:list.id,displayName:list.displayName })) });
    if (m && method==="POST") {
      assert.equal(body.list.template,"genericList");
      for (const column of body.columns) assert.ok(column.text || column.number || column.boolean,`column ${column.name} has a type`);
      const list:List={ id:`list-${lists.length+1}`,displayName:body.displayName,columns:new Set(body.columns.map((column:{ name:string })=>column.name)),items:[],nextId:1 };
      lists.push(list); return reply(201,{ id:list.id });
    }
    m=p.match(/^\/sites\/site-1\/lists\/([^/]+)(\/.*)?$/);
    const list=m && lists.find((item)=>item.id===m![1]);
    if (!m || !list) return notFound();
    const rest=m[2] ?? "";
    if (rest==="/columns" && method==="GET") return reply(200,{ value:[...list.columns,"Title"].map((name)=>({ name })) });
    if (rest==="/columns" && method==="POST") { list.columns.add(body.name); return reply(201,body); }
    if (rest==="/items" && method==="GET") {
      assert.match(url.search,/expand=fields\(select=Title,vtData/);
      const skip=Number(url.searchParams.get("skip") ?? 0); const page=list.items.slice(skip,skip+40);
      return reply(200,{ value:page.map((item)=>({ id:item.id,fields:{ ...item.fields } })),...(skip+40<list.items.length ? { "@odata.nextLink":`https://graph.microsoft.com/v1.0/sites/site-1/lists/${list.id}/items?expand=fields(select=Title,vtData)&skip=${skip+40}` } : {}) });
    }
    if (rest==="/items" && method==="POST") {
      const problem=checkFields(list,body.fields); if (problem) return reply(400,{ error:{ code:"invalidRequest",message:problem } });
      const item={ id:String(list.nextId++),fields:{ ...body.fields } }; list.items.push(item); return reply(201,item);
    }
    m=rest.match(/^\/items\/(\d+)(\/fields)?$/);
    const item=m && list.items.find((entry)=>entry.id===m![1]);
    if (!m || !item) return notFound();
    if (!m[2] && method==="GET") return reply(200,{ id:item.id,fields:{ ...item.fields } });
    if (m[2] && method==="PATCH") { const problem=checkFields(list,body); if (problem) return reply(400,{ error:{ code:"invalidRequest",message:problem } }); Object.assign(item.fields,body); return reply(200,item.fields); }
    return reply(405,{ error:{ code:"notSupported",message:`${method} ${p}` } });
  }

  const fetchImpl=(async(input:string,init:RequestInit={})=>{
    const url=new URL(input); const method=init.method ?? "GET";
    assert.match(String((init.headers as Record<string,string>).Authorization),/^Bearer /);
    if (throttleNext>0) { throttleNext--; return reply(429,{ error:{ code:"TooManyRequests",message:"slow down" } },{ "retry-after":"0.01" }); }
    const body=init.body ? JSON.parse(String(init.body)) : undefined;
    if (url.pathname==="/v1.0/$batch") {
      assert.ok(body.requests.length<=20,"batch has at most 20 requests");
      const responses=[]; const done=new Map<string,number>();
      for (const request of body.requests) {
        if (request.dependsOn?.some((id:string)=>(done.get(id) ?? 500)>=300)) { responses.push({ id:request.id,status:424 }); done.set(request.id,424); continue; }
        const response=route(request.method,new URL(`https://graph.microsoft.com/v1.0${request.url}`),request.body);
        const text=await response.text();
        responses.push({ id:request.id,status:response.status,body:text ? JSON.parse(text) : undefined }); done.set(request.id,response.status);
      }
      return reply(200,{ responses });
    }
    return route(method,url,body);
  }) as typeof fetch;
  return { fetchImpl,lists,sent,throttle:(count:number)=>{ throttleNext=count; },requests:()=>requests };
}

const call=async(ctx:ApiContext,method:string,path:string,body?:unknown,headers:Record<string,string>={})=>{
  const response=await handleApi(new Request(`app://vermitrack${path}`,{ method,headers:{ "Content-Type":"application/json",...headers },body:body===undefined ? undefined : JSON.stringify(body) }),ctx);
  return { status:response.status,body:await response.json() as any };
};

test("SharePoint store: provisioning, seeding and the full API flow", async()=>{
  const fake=fakeGraph();
  const graph=new Graph(async()=>"token",fake.fetchImpl);
  const site=await SharePointStore.resolveSite(graph,"https://lnbgroup.sharepoint.com/sites/VermiTrack/Lists/Whatever/AllItems.aspx");
  assert.equal(site.id,"site-1");

  const store=new SharePointStore(graph,site.id,"https://lnbgroup.sharepoint.com/sites/VermiTrack");
  await assert.rejects(store.connect(),(error:unknown)=>error instanceof NeedsProvisioningError && error.missing.length===7);
  await store.provision();
  assert.equal(fake.lists.length,7);
  fake.throttle(2);
  await seedStore(store,seed as never,"pulkit.singh@lnbgroup.com");
  await store.provision(); // idempotent
  await seedStore(store,seed as never,"pulkit.singh@lnbgroup.com"); // never duplicates
  const entriesList=fake.lists.find((list)=>list.displayName===tableDefs.entries.listName)!;
  assert.equal(entriesList.items.length,149);
  assert.equal(JSON.parse(String(entriesList.items[0].fields.vtData)).importKey,seed.entries[0].importKey,"batch writes keep the original order");

  const fresh=new SharePointStore(graph,site.id,"site");
  await fresh.connect();
  const me=await graph.me();
  const ctx:ApiContext={ store:fresh,identity:{ emails:[me.mail!.toLowerCase(),me.userPrincipalName],fullName:me.displayName },deviceInfo:"test",mailer:{ send:async(mail)=>{ await graph.sendMail(mail); } },memory:{ sessions:new Map(),touched:new Map() } };

  const session=await call(ctx,"GET","/api/session",undefined,{ "x-vermitrack-access-session":"session_123456789" });
  assert.equal(session.status,200); assert.equal(session.body.user.isKeyPerson,true);

  const register=await call(ctx,"GET","/api/register");
  assert.equal(register.body.entries.length,149,"paged reads return every item");

  const created=await call(ctx,"POST","/api/register",{ module:"sales",productionBatchCode:"BATCH-04",values:{ salesDate:"2026-10-01",invoiceNo:"INV-77",customer:"Kisan Agro "+"x".repeat(300),batchNumber:"VCM-26/B04",quantityKg:"500",pricePerKg:"12" } });
  assert.equal(created.status,201,JSON.stringify(created.body));
  assert.equal(created.body.entry.amount,6000);
  const stored=entriesList.items.find((item)=>item.id===String(created.body.entry.id))!;
  assert.equal(stored.fields.vtAmount,6000); assert.equal(stored.fields.vtModule,"sales");
  assert.ok(String(stored.fields.Title).length<=255,"long titles are trimmed for SharePoint");

  const deletion=await call(ctx,"DELETE",`/api/register?id=${created.body.entry.id}`,{ reason:"Entered twice" });
  assert.equal(deletion.status,200,JSON.stringify(deletion.body));
  const admin=await call(ctx,"GET","/api/admin");
  assert.equal(admin.status,200); assert.equal(admin.body.deletionRequests[0].entry.id,created.body.entry.id);
  const approve=await call(ctx,"PATCH","/api/admin/deletions",{ id:admin.body.deletionRequests[0].id,decision:"approve" });
  assert.equal(approve.status,200);
  assert.equal((await call(ctx,"GET","/api/register")).body.entries.length,149);

  const batch=await call(ctx,"POST","/api/batches",{ code:"batch-07",name:"Batch 7",status:"Planned" });
  assert.equal(batch.status,201); assert.equal(batch.body.batch.code,"BATCH-07");
  assert.equal((await call(ctx,"POST","/api/batches",{ code:"BATCH-07",name:"Again",status:"Planned" })).status,409);

  const user=await call(ctx,"POST","/api/admin",{ email:"New.Person@lnbgroup.com",fullName:"New Person",designation:"Supervisor",role:"Production Operator" });
  assert.equal(user.status,201);
  const report=await call(ctx,"POST","/api/email-report",{ to:["new.person@lnbgroup.com"],subject:"Weekly",html:"<p>hi</p>" });
  assert.equal(report.status,200);
  await new Promise((resolve)=>setTimeout(resolve,20));
  assert.ok(fake.sent.length>=2,"deletion notification and report were sent through Outlook");

  // A colleague with restricted scopes only sees their segment and batch.
  const colleague:ApiContext={ ...ctx,identity:{ emails:["new.person@lnbgroup.com"],fullName:"New Person" },memory:{ sessions:new Map(),touched:new Map() } };
  await call(ctx,"PATCH","/api/admin",{ id:user.body.user.id,permittedSegments:"bed-ops",permittedBatches:"BATCH-04" });
  const scoped=await call(colleague,"GET","/api/register");
  assert.ok(scoped.body.entries.length>0 && scoped.body.entries.every((entry:{ module:string })=>entry.module==="bed-ops"));
  assert.equal((await call(colleague,"POST","/api/register",{ module:"sales",productionBatchCode:"BATCH-04",values:{ salesDate:"2026-10-01" } })).status,403);
  assert.equal((await call(colleague,"GET","/api/admin")).status,403);
  const stranger:ApiContext={ ...ctx,identity:{ emails:["someone@else.com"],fullName:null } };
  assert.equal((await call(stranger,"GET","/api/session")).status,403);

  // Import of the web version's "Records → Export CSV" (same columns and quoting as exportCsv in the web app).
  const quote=(value:unknown)=>`"${String(value ?? "").replaceAll('"','""')}"`;
  const seeded=seed.entries[0];
  const newDetails=JSON.stringify({ expenseDate:"2026-09-30",category:"Utilities",description:"Electricity bill \"Sept\"\nmeter 2",paymentMode:"Online",expenseCost:4200,_serialNo:612,_recordCode:"BATCH-08-EX-0612" });
  const csv=[["Sr No","Auto Register Code","Production Batch","Register","Entry Date","Reference","Finished Batch","Bed","Title","Quantity","Amount","Status","Entered By","Created At","Details"],
    [1,"BATCH-04-RM-0001",seeded.productionBatchCode,"Raw material purchase",seeded.entryDate,seeded.referenceCode,"","",seeded.title,seeded.quantity,seeded.amount,seeded.status,"Navjyoti Batch 4 Register","2026-05-01T10:00:00.000Z",seeded.details],
    [612,"BATCH-08-EX-0612","BATCH-08","Cost & expenses","2026-09-30","","","","Utilities · Electricity bill","",4200,"Utilities","Akola Team","2026-09-30T08:15:00.000Z",newDetails],
  ].map((row)=>row.map(quote).join(",")).join("\n");
  assert.equal(parseCsv(csv).length,3,"embedded newlines stay inside the quoted field");
  const imported=await call(ctx,"POST","/api/admin/import",{ csv });
  assert.equal(imported.status,200,JSON.stringify(imported.body));
  assert.deepEqual([imported.body.imported,imported.body.duplicates,imported.body.skipped],[1,1,0]);
  const again=await call(ctx,"POST","/api/admin/import",{ csv });
  assert.deepEqual([again.body.imported,again.body.duplicates],[0,2],"re-importing is a no-op");
  const all=(await call(ctx,"GET","/api/register")).body;
  const restored=all.entries.find((entry:{ productionBatchCode:string })=>entry.productionBatchCode==="BATCH-08");
  assert.equal(restored.amount,4200); assert.equal(restored.createdBy,"Akola Team"); assert.equal(restored.createdAt,"2026-09-30T08:15:00.000Z");
  assert.equal(JSON.parse(restored.details).description,'Electricity bill "Sept"\nmeter 2');
  assert.ok(all.batches.some((item:{ code:string })=>item.code==="BATCH-08"),"missing production batch is created");
  assert.equal((await call(colleague,"POST","/api/admin/import",{ csv })).status,403);
});
