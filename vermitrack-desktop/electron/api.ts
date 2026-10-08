/**
 * The VermiTrack API, ported from the Next.js route handlers of the web app.
 * The renderer keeps calling fetch("/api/..."); main.ts routes those requests
 * here instead of to a web server. Business rules (24-hour edit window,
 * approval-based deletion, segment/batch scopes, audit trail) are unchanged.
 */
import { DELETION_APPROVER_EMAILS, DELETION_APPROVER_NAMES, EDIT_WINDOW_HOURS, KEY_PERSON_NAME, RECYCLE_BIN_DAYS, isDeletionApprover, isKeyPersonEmail } from "../shared/policy";
import type { BatchRow, EntryRow, Store, UserRow } from "./store";
import { deletionDecisionEmail, deletionRequestEmail } from "./mail";

export type Identity = { emails:string[]; fullName:string | null };
export type Mailer = { send(input:{ to:string[]; cc?:string[]; subject:string; html:string }):Promise<void> };
export type ApiContext = {
  store:Store;
  identity:Identity | null;
  deviceInfo:string;
  mailer:Mailer;
  /** Per-process memory: access-log row ids by client session, and last "touch" per user. */
  memory:{ sessions:Map<string,number>; touched:Map<number,number> };
  log?:(message:string,error?:unknown)=>void;
};

type AuthorizedUser = UserRow & { isKeyPerson:boolean };

const allowedModules=new Set(["raw-material","pre-composting","bed-ops","harvest","quality","sales","inventory","expenses"]);
const moduleCodes:Record<string,string>={ "raw-material":"RM","pre-composting":"PC","bed-ops":"BD",harvest:"HV",quality:"QC",sales:"SL",inventory:"ST",expenses:"EX" };
const TOUCH_INTERVAL_MS=5*60*1000;

const json=(body:unknown,status=200)=>Response.json(body,{ status });
const nowIso=()=>new Date().toISOString();
const time=(value:string | null | undefined)=>value ? new Date(value).getTime() : 0;
function numberValue(value:unknown) { if (value==="" || value===null || value===undefined) return null; const parsed=Number(value); return Number.isFinite(parsed) ? parsed : null; }
function textValue(value:unknown) { return typeof value==="string" ? value.trim() : ""; }
function detailObject(value:string) { try { return JSON.parse(value) as Record<string,unknown>; } catch { return {}; } }
function errorResponse(error:unknown,fallback:string) {
  const status=(error as { status?:number })?.status;
  const message=error instanceof Error ? error.message : fallback;
  if (status===401) return json({ error:message,code:"UNAUTHENTICATED" },401);
  return json({ error:message || fallback },500);
}

// ------------------------------------------------------------------ identity

function authorized(user:UserRow):AuthorizedUser { return { ...user,isKeyPerson:isKeyPersonEmail(user.email) }; }
export function isAdmin(user:AuthorizedUser) { return user.role.toLowerCase()==="admin"; }
export function scopeAllows(scope:string,value:string) {
  if (scope.trim().toLowerCase()==="all") return true;
  return scope.split(",").map((item)=>item.trim().toLowerCase()).includes(value.trim().toLowerCase());
}
export function canUse(user:AuthorizedUser,module:string,batch:string) { return scopeAllows(user.permittedSegments,module) && scopeAllows(user.permittedBatches,batch); }

async function findUser(ctx:ApiContext,fresh=false) {
  const emails=new Set((ctx.identity?.emails ?? []).map((email)=>email.trim().toLowerCase()).filter(Boolean));
  const users=await ctx.store.table("users").all({ fresh });
  return users.find((user)=>emails.has(user.email.trim().toLowerCase()));
}

async function touchUser(ctx:ApiContext,user:AuthorizedUser) {
  const now=Date.now();
  if (now-(ctx.memory.touched.get(user.id) ?? 0)<TOUCH_INTERVAL_MS) return;
  ctx.memory.touched.set(user.id,now);
  const newSession=!user.lastActiveAt || now-time(user.lastActiveAt)>30*60*1000;
  const stamp=new Date(now).toISOString();
  await ctx.store.table("users").update(user.id,{ lastActiveAt:stamp,...(newSession ? { lastLoginAt:stamp } : {}) });
  if (newSession) await ctx.store.table("activity").insert({ actorEmail:user.email,actorName:user.fullName,action:"LOGIN",entityType:"session",entityId:String(user.id),module:null,productionBatchCode:null,summary:`${user.fullName} signed in to VermiTrack desktop`,beforeJson:null,afterJson:null,createdAt:stamp });
  user.lastActiveAt=stamp;
  if (newSession) user.lastLoginAt=stamp;
}

type AuthResult = { error:Response } | { user:AuthorizedUser; method:"microsoft" };

async function requireAppUser(ctx:ApiContext,{ touch=true }:{ touch?:boolean }={}):Promise<AuthResult> {
  if (!ctx.identity?.emails.length) return { error:json({ error:"Sign in with your Microsoft 365 (Outlook) account.",code:"UNAUTHENTICATED" },401) } as const;
  let user=await findUser(ctx);
  if (!user || user.status!=="active") user=await findUser(ctx,true);
  if (!user || user.status!=="active") return { error:json({ error:`${ctx.identity.emails[0]} is not an approved VermiTrack user. Ask ${KEY_PERSON_NAME} to add you under Admin → Approved users.`,code:"NOT_AUTHORISED",email:ctx.identity.emails[0] },403) } as const;
  const result=authorized(user);
  if (touch) await touchUser(ctx,result).catch((error)=>ctx.log?.("Could not record user activity",error));
  return { user:result,method:"microsoft" as const } as const;
}

async function writeAudit(ctx:ApiContext,user:AuthorizedUser,input:{ action:string; entityType:string; entityId?:string; module?:string | null; productionBatchCode?:string | null; summary:string; before?:unknown; after?:unknown }) {
  await ctx.store.table("activity").insert({
    actorEmail:user.email,actorName:user.fullName,action:input.action,entityType:input.entityType,entityId:input.entityId ?? null,
    module:input.module ?? null,productionBatchCode:input.productionBatchCode ?? null,summary:input.summary,
    beforeJson:input.before===undefined ? null : JSON.stringify(input.before),afterJson:input.after===undefined ? null : JSON.stringify(input.after),createdAt:nowIso(),
  });
}

/** Outlook notifications never block or fail the action that triggered them. */
function notify(ctx:ApiContext,mail:{ to:string[]; cc?:string[]; subject:string; html:string }) {
  const to=[...new Set(mail.to.map((address)=>address.toLowerCase()))];
  if (!to.length) return;
  void ctx.mailer.send({ ...mail,to }).catch((error)=>ctx.log?.("Outlook notification failed",error));
}

// ------------------------------------------------------------------ session

async function sessionGet(request:Request,ctx:ApiContext) {
  const accessSession=clientSessionId(request.headers.get("x-vermitrack-access-session"));
  const auth=await requireAppUser(ctx);
  const access=ctx.store.table("access");
  const now=nowIso();
  if ("error" in auth) {
    if (ctx.identity?.emails.length && accessSession && !ctx.memory.sessions.has(accessSession)) {
      const knownUser=await findUser(ctx);
      const row=await access.insert({ userId:knownUser?.id ?? null,email:ctx.identity.emails[0],fullName:ctx.identity.fullName || knownUser?.fullName || null,outcome:"denied",method:"microsoft",ipAddress:null,deviceInfo:ctx.deviceInfo,clientSessionId:accessSession,lastSeenAt:now,endedAt:now,durationSeconds:0,createdAt:now }).catch(()=>null);
      if (row) ctx.memory.sessions.set(accessSession,row.id);
    }
    return auth.error;
  }
  if (accessSession) await recordAccess(ctx,accessSession,auth.user,now).catch((error)=>ctx.log?.("Could not record the session (read-only site access?)",error));
  return json({ user:auth.user,authMethod:auth.method,editWindowHours:EDIT_WINDOW_HOURS,recycleBinDays:RECYCLE_BIN_DAYS,deletionApprover:isDeletionApprover(auth.user.email) });
}

async function recordAccess(ctx:ApiContext,accessSession:string,user:AuthorizedUser,now:string) {
  const access=ctx.store.table("access");
  const knownId=ctx.memory.sessions.get(accessSession);
  const existing=knownId ? await access.get(knownId) : undefined;
  if (existing && existing.outcome==="success") {
    const durationSeconds=Math.max(existing.durationSeconds,Math.round((Date.now()-time(existing.createdAt))/1000));
    await access.update(existing.id,{ lastSeenAt:now,endedAt:null,durationSeconds });
  } else if (!existing) {
    const row=await access.insert({ userId:user.id,email:user.email,fullName:user.fullName,outcome:"success",method:"microsoft",ipAddress:null,deviceInfo:ctx.deviceInfo,clientSessionId:accessSession,lastSeenAt:now,endedAt:null,durationSeconds:0,createdAt:now });
    ctx.memory.sessions.set(accessSession,row.id);
  }
}

function clientSessionId(value:unknown) { const session=typeof value==="string" ? value.trim() : ""; return /^[a-zA-Z0-9_-]{12,120}$/.test(session) ? session : null; }

async function sessionPost(request:Request,ctx:ApiContext) {
  try {
    const payload=await request.json().catch(()=>({})) as Record<string,unknown>;
    const accessSession=clientSessionId(payload.clientSessionId);
    const id=accessSession ? ctx.memory.sessions.get(accessSession) : undefined;
    if (!id) return json({ ended:true });
    const existing=await ctx.store.table("access").get(id);
    if (existing && existing.outcome==="success") {
      const durationSeconds=Math.max(existing.durationSeconds,Math.round((Date.now()-time(existing.createdAt))/1000));
      const now=nowIso();
      await ctx.store.table("access").update(id,{ lastSeenAt:now,endedAt:now,durationSeconds });
    }
    if (accessSession) ctx.memory.sessions.delete(accessSession);
  } catch (error) { ctx.log?.("Could not close the access session",error); }
  return json({ ended:true });
}

// ----------------------------------------------------------------- register

function recordCode(productionBatchCode:string,registerId:string,serial:number) { return `${productionBatchCode}-${moduleCodes[registerId] ?? "RG"}-${String(serial).padStart(4,"0")}`; }
function rmMaterialPrefix(materialType:unknown) { return textValue(materialType).split(/[^a-z0-9]+/i).filter(Boolean).map((word)=>word.slice(0,2)).join("").toUpperCase(); }
function rmCodeBase(purchaseDate:unknown,materialType:unknown) {
  const date=textValue(purchaseDate).replace(/\D/g,"").slice(0,8) || new Date().toISOString().slice(0,10).replaceAll("-","");
  return `RM/${date}/${rmMaterialPrefix(materialType)}`;
}

async function nextRmLotCode(ctx:ApiContext,input:Record<string,unknown>,excludeEntryId?:number) {
  const base=rmCodeBase(input.purchaseDate,input.materialType);
  const rows=(await ctx.store.table("entries").all({ fresh:true })).filter((row)=>row.module==="raw-material");
  const matching=rows.filter((row)=>row.id!==excludeEntryId && (()=>{ const detail=detailObject(row.details); return rmCodeBase(detail.purchaseDate,detail.materialType)===base; })());
  const used=new Set<number>(); let legacyCount=0;
  for (const row of matching) {
    const code=textValue(row.referenceCode);
    if (code===base) used.add(1);
    else { const match=code.match(new RegExp(`^${base.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")}(\\d+)$`)); if (match && Number(match[1])>=2) used.add(Number(match[1])); else legacyCount+=1; }
  }
  while (legacyCount>0) { let slot=1; while (used.has(slot)) slot+=1; used.add(slot); legacyCount-=1; }
  let sequence=1; while (used.has(sequence)) sequence+=1;
  return sequence===1 ? base : `${base}${sequence}`;
}

async function normalizedValues(ctx:ApiContext,registerId:string,input:Record<string,unknown>,productionBatchCode:string,existingDetails?:Record<string,unknown>,excludeEntryId?:number) {
  const values={ ...input };
  if (registerId==="raw-material") {
    const unchanged=existingDetails && textValue(existingDetails.purchaseDate)===textValue(values.purchaseDate) && textValue(existingDetails.materialType)===textValue(values.materialType);
    values.lotCode=unchanged && textValue(existingDetails.lotCode) ? textValue(existingDetails.lotCode) : await nextRmLotCode(ctx,values,excludeEntryId);
  }
  if (registerId==="harvest") { values.productionBatchCode=productionBatchCode; delete values.rawQtyKg; delete values.recoveryPct; }
  return values;
}

function deriveEntry(registerId:string,values:Record<string,unknown>,productionBatchCode:string,systemMeta:Record<string,unknown>={}) {
  const entryDate=textValue(values.purchaseDate ?? values.eventDate ?? values.harvestDate ?? values.samplingDate ?? values.salesDate ?? values.entryDate ?? values.expenseDate);
  const batchCode=textValue(values.batchNumber ?? values.productionBatchCode) || null;
  const bedNumber=textValue(values.bedNumber) || null;
  let referenceCode=textValue(values.lotCode ?? values.invoiceNo ?? values.reportNo) || null;
  let title="Register entry"; let quantity:number | null=null; let amount:number | null=null; let status:string | null=null;
  const computed:Record<string,number>={};
  if (registerId==="raw-material") {
    quantity=numberValue(values.quantityKg); computed.totalAmount=(quantity ?? 0)*(numberValue(values.ratePerUnit) ?? 0); amount=computed.totalAmount;
    status=textValue(values.paymentStatus) || null; title=`${textValue(values.materialType) || "Material"} from ${textValue(values.supplier) || "supplier"}`;
  } else if (registerId==="pre-composting") {
    quantity=numberValue(values.totalWeightKg); status=textValue(values.eventType) || null; title=`${textValue(values.eventType) || "Pre-compost event"} · ${textValue(values.lotCode)}`;
  } else if (registerId==="bed-ops") {
    quantity=numberValue(values.dungWeightKg); status=textValue(values.currentStatus) || null; referenceCode=bedNumber; title=`${textValue(values.eventType) || "Bed operation"} · ${bedNumber}`;
  } else if (registerId==="harvest") {
    quantity=numberValue(values.netYieldKg);
    computed.packagingCoveragePct=(quantity ?? 0)>0 ? ((numberValue(values.packagedQuantityKg) ?? 0)/(quantity ?? 1))*100 : 0;
    status="Harvested"; referenceCode=bedNumber; title=`Harvest recorded · ${bedNumber}`;
  } else if (registerId==="quality") {
    status=textValue(values.fcoStatus) || null; title=`QC report · ${batchCode || textValue(values.productType)}`;
  } else if (registerId==="sales") {
    quantity=numberValue(values.quantityKg); computed.totalRevenue=(quantity ?? 0)*(numberValue(values.pricePerKg) ?? 0); amount=computed.totalRevenue; title=`Sale to ${textValue(values.customer) || "customer"}`;
  } else if (registerId==="inventory") {
    computed.closingStockKg=(numberValue(values.openingStockKg) ?? 0)+(numberValue(values.productionKg) ?? 0)-(numberValue(values.salesOutflowKg) ?? 0)-(numberValue(values.damageLossKg) ?? 0);
    quantity=computed.closingStockKg; status=textValue(values.auditStatus) || null; title=`Stock update · ${textValue(values.packagingSize)}`;
  } else if (registerId==="expenses") {
    amount=numberValue(values.expenseCost); status=textValue(values.category) || null; title=`${textValue(values.category) || "Expense"} · ${textValue(values.description).slice(0,48)}`;
  }
  return { productionBatchCode,entryDate,referenceCode,batchCode,bedNumber,title,quantity,amount,status,details:JSON.stringify({ ...values,...computed,...systemMeta }) };
}

async function batchExists(ctx:ApiContext,code:string) {
  const find=(rows:BatchRow[])=>rows.some((batch)=>batch.code===code);
  return find(await ctx.store.table("batches").all()) || find(await ctx.store.table("batches").all({ fresh:true }));
}

function byNewest(a:{ createdAt:string; id:number },b:{ createdAt:string; id:number }) { return time(b.createdAt)-time(a.createdAt) || b.id-a.id; }

async function registerGet(request:Request,ctx:ApiContext) {
  const url=new URL(request.url); const registerId=url.searchParams.get("module");
  const auth=await requireAppUser(ctx);
  if ("error" in auth) return auth.error;
  if (url.searchParams.get("mode")==="rm-code") {
    const purchaseDate=textValue(url.searchParams.get("purchaseDate")); const materialType=textValue(url.searchParams.get("materialType")); const excludeEntryId=numberValue(url.searchParams.get("excludeEntryId")) ?? undefined;
    if (!purchaseDate || !materialType) return json({ error:"Purchase date and raw material type are required." },400);
    return json({ code:await nextRmLotCode(ctx,{ purchaseDate,materialType },excludeEntryId) });
  }
  const [allEntries,baselineRows,batchRows]=await Promise.all([ctx.store.table("entries").all(),ctx.store.table("baselines").all(),ctx.store.table("batches").all()]);
  const entries=allEntries
    .filter((entry)=>(!registerId || !allowedModules.has(registerId) || entry.module===registerId) && entry.deletionStatus!=="deleted" && canUse(auth.user,entry.module,entry.productionBatchCode))
    .sort(byNewest);
  return json({
    baselines:baselineRows.filter((row)=>scopeAllows(auth.user.permittedBatches,row.productionBatchCode)),
    batches:batchRows.sort(byNewest).filter((row)=>scopeAllows(auth.user.permittedBatches,row.code)),
    entries,
  });
}

type EntryPayload = { id?:number; module?:string; productionBatchCode?:string; values?:Record<string,unknown> };

async function registerPost(request:Request,ctx:ApiContext) {
  const payload=await request.json() as EntryPayload;
  const registerId=textValue(payload.module); const productionBatchCode=textValue(payload.productionBatchCode);
  if (!allowedModules.has(registerId)) return json({ error:"Select a valid register." },400);
  if (!productionBatchCode) return json({ error:"Production batch is required." },400);
  if (!payload.values || typeof payload.values!=="object") return json({ error:"Entry details are required." },400);
  const auth=await requireAppUser(ctx);
  if ("error" in auth) return auth.error;
  if (!canUse(auth.user,registerId,productionBatchCode)) return json({ error:"You do not have access to this segment or batch." },403);
  if (!await batchExists(ctx,productionBatchCode)) return json({ error:"Select an existing production batch." },400);
  if (registerId==="raw-material" && !textValue(payload.values.materialType)) return json({ error:"Select the raw material type before saving." },400);
  const entry=deriveEntry(registerId,await normalizedValues(ctx,registerId,payload.values,productionBatchCode),productionBatchCode);
  if (!entry.entryDate) return json({ error:"Entry date is required." },400);
  const now=nowIso();
  const entries=ctx.store.table("entries");
  const created=await entries.insert({ importKey:null,module:registerId,...entry,createdBy:auth.user.fullName,createdByEmail:auth.user.email,deletionStatus:"active",deletedAt:null,deletedBy:null,createdAt:now,updatedAt:now });
  const systemMeta={ _serialNo:created.id,_recordCode:recordCode(productionBatchCode,registerId,created.id),_productionBatchCode:productionBatchCode };
  const result=await entries.update(created.id,{ details:JSON.stringify({ ...detailObject(created.details),...systemMeta }) });
  await writeAudit(ctx,auth.user,{ action:"ENTRY_CREATED",entityType:"entry",entityId:String(result.id),module:registerId,productionBatchCode,summary:`Created ${result.title}`,after:result });
  return json({ entry:result },201);
}

async function registerPatch(request:Request,ctx:ApiContext) {
  const payload=await request.json() as EntryPayload;
  const id=numberValue(payload.id); const registerId=textValue(payload.module); const productionBatchCode=textValue(payload.productionBatchCode);
  if (!id || id<1) return json({ error:"Select a valid entry to edit." },400);
  if (!allowedModules.has(registerId)) return json({ error:"Select a valid register." },400);
  if (!productionBatchCode) return json({ error:"Production batch is required." },400);
  if (!payload.values || typeof payload.values!=="object") return json({ error:"Entry details are required." },400);
  const auth=await requireAppUser(ctx);
  if ("error" in auth) return auth.error;
  if (!canUse(auth.user,registerId,productionBatchCode)) return json({ error:"You do not have access to this segment or batch." },403);
  if (!await batchExists(ctx,productionBatchCode)) return json({ error:"Select an existing production batch." },400);
  const entries=ctx.store.table("entries");
  const existing=await entries.get(id);
  if (!existing) return json({ error:"This entry could not be found." },404);
  if (existing.deletionStatus!=="active") return json({ error:"This entry is awaiting or has completed deletion approval." },409);
  const ownEntry=existing.createdByEmail===auth.user.email;
  const withinWindow=Date.now()-time(existing.createdAt)<=EDIT_WINDOW_HOURS*60*60*1000;
  if (!isAdmin(auth.user) && (!ownEntry || !withinWindow)) return json({ error:"Only an administrator can edit another user’s entry or an entry older than 24 hours." },403);
  const existingDetails=detailObject(existing.details); const serial=numberValue(existingDetails._serialNo) ?? id;
  if (registerId==="raw-material" && !textValue(payload.values.materialType)) return json({ error:"Select the raw material type before saving." },400);
  const entry=deriveEntry(registerId,await normalizedValues(ctx,registerId,payload.values,productionBatchCode,existingDetails,id),productionBatchCode,{ _serialNo:serial,_recordCode:recordCode(productionBatchCode,registerId,serial),_productionBatchCode:productionBatchCode });
  if (!entry.entryDate) return json({ error:"Entry date is required." },400);
  const updated=await entries.update(id,{ module:registerId,...entry,updatedAt:nowIso() });
  await writeAudit(ctx,auth.user,{ action:"ENTRY_UPDATED",entityType:"entry",entityId:String(id),module:registerId,productionBatchCode,summary:`Updated ${updated.title}`,before:existing,after:updated });
  return json({ entry:updated });
}

async function registerDelete(request:Request,ctx:ApiContext) {
  const id=numberValue(new URL(request.url).searchParams.get("id"));
  if (!id || id<1) return json({ error:"Select a valid entry to delete." },400);
  const payload=await request.json().catch(()=>({})) as { reason?:string };
  const reason=textValue(payload.reason);
  if (reason.length<5) return json({ error:"Enter a deletion reason of at least 5 characters." },400);
  const auth=await requireAppUser(ctx);
  if ("error" in auth) return auth.error;
  const entries=ctx.store.table("entries");
  const entry=await entries.get(id);
  if (!entry || entry.deletionStatus==="deleted") return json({ error:"This entry could not be found." },404);
  if (!canUse(auth.user,entry.module,entry.productionBatchCode)) return json({ error:"You do not have access to this entry." },403);
  if (entry.deletionStatus==="pending") return json({ error:"A deletion request is already pending for this entry." },409);
  const now=nowIso();
  const requestRow=await ctx.store.table("deletions").insert({ entryId:entry.id,reason,status:"pending",requestedByEmail:auth.user.email,requestedByName:auth.user.fullName,decidedByEmail:null,decidedByName:null,decisionNote:null,decidedAt:null,createdAt:now,updatedAt:now });
  const updated=await entries.update(id,{ deletionStatus:"pending",updatedAt:now });
  await writeAudit(ctx,auth.user,{ action:"DELETE_REQUESTED",entityType:"entry",entityId:String(entry.id),module:entry.module,productionBatchCode:entry.productionBatchCode,summary:`Requested deletion of ${entry.title}: ${reason}`,before:entry,after:updated });
  notify(ctx,{ to:DELETION_APPROVER_EMAILS.filter((email)=>email!==auth.user.email),cc:[auth.user.email],...deletionRequestEmail({ entry,reason,requestedBy:auth.user.fullName }) });
  return json({ requestId:requestRow.id,entry:updated });
}

// ------------------------------------------------------------------ batches

async function batchesGet(_request:Request,ctx:ApiContext) {
  const auth=await requireAppUser(ctx);
  if ("error" in auth) return auth.error;
  const batches=(await ctx.store.table("batches").all()).sort(byNewest);
  return json({ batches:batches.filter((batch)=>scopeAllows(auth.user.permittedBatches,batch.code)) });
}

async function batchesPost(request:Request,ctx:ApiContext) {
  const payload=await request.json() as Record<string,unknown>;
  const code=textValue(payload.code).toUpperCase(); const name=textValue(payload.name); const status=textValue(payload.status);
  if (!/^[A-Z0-9]+(?:-[A-Z0-9]+)*$/.test(code)) return json({ error:"Use a batch code such as BATCH-05." },400);
  if (!name) return json({ error:"Batch name is required." },400);
  if (!["Planned","Active","Completed","On hold"].includes(status)) return json({ error:"Select a valid batch status." },400);
  const auth=await requireAppUser(ctx);
  if ("error" in auth) return auth.error;
  if (!isAdmin(auth.user)) return json({ error:"Administrator access is required to create a batch." },403);
  if (await batchExists(ctx,code)) return json({ error:"A batch with this code already exists." },409);
  const now=nowIso();
  const batch=await ctx.store.table("batches").insert({ code,name,financialYear:textValue(payload.financialYear) || null,status,startDate:textValue(payload.startDate) || null,expectedEndDate:textValue(payload.expectedEndDate) || null,notes:textValue(payload.notes) || null,createdBy:auth.user.fullName,createdAt:now,updatedAt:now });
  await writeAudit(ctx,auth.user,{ action:"BATCH_CREATED",entityType:"batch",entityId:batch.code,productionBatchCode:batch.code,summary:`Created ${batch.name}`,after:batch });
  return json({ batch },201);
}

// -------------------------------------------------------------------- admin

function validEmail(value:string) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value); }

async function adminGet(_request:Request,ctx:ApiContext) {
  const auth=await requireAppUser(ctx);
  if ("error" in auth) return auth.error;
  if (!isAdmin(auth.user)) return json({ error:"Administrator access is required." },403);
  const deletionApprover=isDeletionApprover(auth.user.email);
  let deletionPayload:unknown[]=[];
  if (deletionApprover) {
    const [requests,entries]=await Promise.all([ctx.store.table("deletions").all({ fresh:true }),ctx.store.table("entries").all()]);
    const entryById=new Map(entries.map((entry)=>[entry.id,entry]));
    deletionPayload=requests.sort(byNewest).slice(0,100).map((request)=>({ ...request,entry:entryById.get(request.entryId) ?? null }));
  }
  if (!auth.user.isKeyPerson) return json({ keyPerson:false,deletionApprover,deletionRequests:deletionPayload });
  const [users,activities,accesses]=await Promise.all([ctx.store.table("users").all({ fresh:true }),ctx.store.table("activity").all(),ctx.store.table("access").all()]);
  const staleBefore=Date.now()-2*60*1000;
  for (const row of accesses) {
    if (row.outcome!=="success" || row.endedAt || time(row.lastSeenAt ?? row.createdAt)>=staleBefore) continue;
    row.endedAt=row.lastSeenAt ?? row.createdAt;
    row.durationSeconds=Math.max(row.durationSeconds,Math.round((time(row.endedAt)-time(row.createdAt))/1000));
    void ctx.store.table("access").update(row.id,{ endedAt:row.endedAt,durationSeconds:row.durationSeconds }).catch((error)=>ctx.log?.("Could not close a stale session",error));
  }
  return json({
    keyPerson:true,deletionApprover:true,
    users:users.sort((a,b)=>a.fullName.localeCompare(b.fullName)).map(authorized),
    activities:activities.sort(byNewest).slice(0,250).map(({ beforeJson:_b,afterJson:_a,...row })=>row),
    accessLogs:accesses.sort(byNewest).slice(0,250),
    deletionRequests:deletionPayload,
  });
}

async function adminPost(request:Request,ctx:ApiContext) {
  const auth=await requireAppUser(ctx);
  if ("error" in auth) return auth.error;
  if (!auth.user.isKeyPerson) return json({ error:"Only the Key Person can manage approved users." },403);
  const payload=await request.json() as Record<string,unknown>;
  const email=textValue(payload.email).toLowerCase(); const fullName=textValue(payload.fullName); const designation=textValue(payload.designation);
  if (!validEmail(email)) return json({ error:"Enter a valid official email address." },400);
  if (!fullName || !designation) return json({ error:"Name and designation are required." },400);
  if ((await ctx.store.table("users").all({ fresh:true })).some((user)=>user.email.toLowerCase()===email)) return json({ error:"A user with this email already exists." },409);
  const now=nowIso();
  const created=await ctx.store.table("users").insert({ email,fullName,mobile:textValue(payload.mobile) || null,designation,role:textValue(payload.role) || "Admin",permittedSegments:textValue(payload.permittedSegments) || "all",permittedBatches:textValue(payload.permittedBatches) || "all",status:"active",isKeyPerson:false,lastLoginAt:null,lastActiveAt:null,createdAt:now,updatedAt:now });
  await writeAudit(ctx,auth.user,{ action:"USER_CREATED",entityType:"user",entityId:String(created.id),summary:`Added ${created.fullName} as ${created.role}`,after:created });
  return json({ user:authorized(created) },201);
}

async function adminPatch(request:Request,ctx:ApiContext) {
  const auth=await requireAppUser(ctx);
  if ("error" in auth) return auth.error;
  if (!auth.user.isKeyPerson) return json({ error:"Only the Key Person can manage approved users." },403);
  const payload=await request.json() as Record<string,unknown>; const id=Number(payload.id);
  if (!Number.isInteger(id) || id<1) return json({ error:"Select a valid user." },400);
  const existingRow=await ctx.store.table("users").get(id);
  if (!existingRow) return json({ error:"This user could not be found." },404);
  const existing=authorized(existingRow);
  const nextStatus=existing.isKeyPerson ? "active" : payload.status==="inactive" ? "inactive" : payload.status==="active" ? "active" : existing.status;
  if (existing.email===auth.user.email && nextStatus==="inactive") return json({ error:"You cannot deactivate your own administrator account." },400);
  const updated=await ctx.store.table("users").update(id,{
    fullName:textValue(payload.fullName) || existing.fullName,mobile:payload.mobile===undefined ? existing.mobile : textValue(payload.mobile) || null,
    designation:textValue(payload.designation) || existing.designation,role:existing.isKeyPerson ? "Admin" : textValue(payload.role) || existing.role,
    permittedSegments:textValue(payload.permittedSegments) || existing.permittedSegments,permittedBatches:textValue(payload.permittedBatches) || existing.permittedBatches,
    status:nextStatus,updatedAt:nowIso(),
  });
  await writeAudit(ctx,auth.user,{ action:"USER_UPDATED",entityType:"user",entityId:String(id),summary:`Updated access for ${updated.fullName}`,before:existing,after:updated });
  return json({ user:authorized(updated) });
}

async function deletionsPatch(request:Request,ctx:ApiContext) {
  const auth=await requireAppUser(ctx);
  if ("error" in auth) return auth.error;
  if (!isDeletionApprover(auth.user.email)) return json({ error:`Only ${DELETION_APPROVER_NAMES} can decide deletion requests.` },403);
  const payload=await request.json() as { id?:number; decision?:"approve" | "reject" | "restore"; note?:string };
  const id=Number(payload.id); const decision=payload.decision;
  if (!Number.isInteger(id) || id<1 || !decision || !["approve","reject","restore"].includes(decision)) return json({ error:"Select a valid deletion request and decision." },400);
  const deletions=ctx.store.table("deletions"); const entries=ctx.store.table("entries");
  const item=await deletions.get(id);
  if (!item) return json({ error:"This deletion request could not be found." },404);
  const entry:EntryRow | undefined=await entries.get(item.entryId);
  if (!entry) return json({ error:"The linked register entry could not be found." },404);
  const now=nowIso(); const note=payload.note?.trim() || null;
  if (decision==="restore") {
    if (item.status!=="approved") return json({ error:"Only an approved deletion can be restored." },400);
    if (Date.now()-time(item.decidedAt)>RECYCLE_BIN_DAYS*24*60*60*1000) return json({ error:`Entries can only be restored within ${RECYCLE_BIN_DAYS} days of deletion.` },400);
    await deletions.update(id,{ status:"restored",decidedByEmail:auth.user.email,decidedByName:auth.user.fullName,decisionNote:note || `Restored from ${RECYCLE_BIN_DAYS}-day recycle bin`,decidedAt:now,updatedAt:now });
    await entries.update(entry.id,{ deletionStatus:"active",deletedAt:null,deletedBy:null,updatedAt:now });
    await writeAudit(ctx,auth.user,{ action:"ENTRY_RESTORED",entityType:"entry",entityId:String(entry.id),module:entry.module,productionBatchCode:entry.productionBatchCode,summary:`Restored ${entry.title}`,after:entry });
    return json({ status:"restored",entryId:entry.id });
  }
  if (item.status!=="pending") return json({ error:"This request has already been decided." },409);
  const status=decision==="approve" ? "approved" : "rejected";
  await deletions.update(id,{ status,decidedByEmail:auth.user.email,decidedByName:auth.user.fullName,decisionNote:note,decidedAt:now,updatedAt:now });
  await entries.update(entry.id,decision==="approve" ? { deletionStatus:"deleted",deletedAt:now,deletedBy:auth.user.fullName,updatedAt:now } : { deletionStatus:"active",updatedAt:now });
  await writeAudit(ctx,auth.user,{ action:decision==="approve" ? "DELETE_APPROVED" : "DELETE_REJECTED",entityType:"entry",entityId:String(entry.id),module:entry.module,productionBatchCode:entry.productionBatchCode,summary:`${decision==="approve" ? "Approved deletion of" : "Rejected deletion request for"} ${entry.title}`,before:entry });
  notify(ctx,{ to:[item.requestedByEmail].filter((email)=>email!==auth.user.email),...deletionDecisionEmail({ entry,approved:decision==="approve",decidedBy:auth.user.fullName,note }) });
  return json({ status,entryId:entry.id });
}

// ---------------------------------------------------- import from web version

/** RFC 4180 CSV parser (quoted fields, doubled quotes, embedded newlines). */
export function parseCsv(text:string) {
  const rows:string[][]=[]; let row:string[]=[]; let field=""; let quoted=false;
  for (let index=0;index<text.length;index++) {
    const char=text[index];
    if (quoted) {
      if (char==='"' && text[index+1]==='"') { field+='"'; index++; }
      else if (char==='"') quoted=false;
      else field+=char;
    } else if (char==='"') quoted=true;
    else if (char===",") { row.push(field); field=""; }
    else if (char==="\n" || char==="\r") { if (char==="\r" && text[index+1]==="\n") index++; row.push(field); rows.push(row); row=[]; field=""; }
    else field+=char;
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows.filter((cells)=>cells.some((cell)=>cell.trim()));
}

const moduleTitles:Record<string,string>={ "raw material purchase":"raw-material","pre-composting":"pre-composting","bed production & operations":"bed-ops","harvesting & packaging":"harvest","quality control":"quality","sales register":"sales","stock inventory":"inventory","cost & expenses":"expenses" };
function contentKey(module:string,batch:string,details:string) {
  const detail=detailObject(details);
  return `${module}|${batch}|${JSON.stringify(Object.fromEntries(Object.entries(detail).filter(([key])=>!key.startsWith("_")).sort(([a],[b])=>a.localeCompare(b))))}`;
}

async function adminImport(request:Request,ctx:ApiContext) {
  const auth=await requireAppUser(ctx);
  if ("error" in auth) return auth.error;
  if (!auth.user.isKeyPerson) return json({ error:"Only the Key Person can import records." },403);
  const payload=await request.json() as { csv?:unknown };
  const rows=parseCsv(typeof payload.csv==="string" ? payload.csv.replace(/^\uFEFF/,"") : "");
  const header=rows.shift()?.map((cell)=>cell.trim().toLowerCase()) ?? [];
  const column=(name:string)=>header.indexOf(name.toLowerCase());
  const required=["Production Batch","Register","Entry Date","Title","Details"];
  const missing=required.filter((name)=>column(name)<0);
  if (missing.length) return json({ error:`This is not a VermiTrack “Records → Export CSV” file (missing columns: ${missing.join(", ")}).` },400);
  const cell=(row:string[],name:string)=>{ const index=column(name); return index<0 ? "" : (row[index] ?? "").trim(); };
  const entries=ctx.store.table("entries"); const batches=ctx.store.table("batches");
  const existing=new Set((await entries.all({ fresh:true })).map((entry)=>contentKey(entry.module,entry.productionBatchCode,entry.details)));
  const knownBatches=new Set((await batches.all({ fresh:true })).map((batch)=>batch.code));
  const now=nowIso(); const toInsert:Parameters<typeof entries.insertMany>[0]=[]; const problems:string[]=[]; let duplicates=0;
  for (const [index,row] of rows.entries()) {
    const module=moduleTitles[cell(row,"Register").toLowerCase()];
    const productionBatchCode=cell(row,"Production Batch").toUpperCase();
    const details=cell(row,"Details");
    if (!module || !productionBatchCode || !details.startsWith("{")) { problems.push(`row ${index+2}`); continue; }
    const key=contentKey(module,productionBatchCode,details);
    if (existing.has(key)) { duplicates++; continue; }
    existing.add(key);
    if (!knownBatches.has(productionBatchCode)) {
      await batches.insert({ code:productionBatchCode,name:productionBatchCode.replace(/^BATCH-0*/,"Batch "),financialYear:null,status:"Active",startDate:null,expectedEndDate:null,notes:"Created while importing records from the web version.",createdBy:auth.user.fullName,createdAt:now,updatedAt:now });
      knownBatches.add(productionBatchCode);
    }
    const createdAt=cell(row,"Created At"); const detail=detailObject(details);
    toInsert.push({ importKey:typeof detail._recordCode==="string" ? `web:${detail._recordCode}` : null,productionBatchCode,module,entryDate:cell(row,"Entry Date"),referenceCode:cell(row,"Reference") || null,batchCode:cell(row,"Finished Batch") || null,bedNumber:cell(row,"Bed") || null,title:cell(row,"Title"),quantity:numberValue(cell(row,"Quantity")),amount:numberValue(cell(row,"Amount")),status:cell(row,"Status") || null,details,createdBy:cell(row,"Entered By") || "Web version",createdByEmail:null,deletionStatus:"active",deletedAt:null,deletedBy:null,createdAt:Number.isNaN(Date.parse(createdAt)) ? now : new Date(createdAt).toISOString(),updatedAt:now });
  }
  if (toInsert.length) await entries.insertMany(toInsert);
  await writeAudit(ctx,auth.user,{ action:"ENTRIES_IMPORTED",entityType:"import",summary:`Imported ${toInsert.length} entries from the web version (${duplicates} already present${problems.length ? `, ${problems.length} unreadable` : ""})` });
  return json({ imported:toInsert.length,duplicates,skipped:problems.length,problems:problems.slice(0,10) });
}

// ------------------------------------------------------------- email report

async function emailReport(request:Request,ctx:ApiContext) {
  const auth=await requireAppUser(ctx);
  if ("error" in auth) return auth.error;
  const payload=await request.json() as { to?:unknown; subject?:unknown; html?:unknown };
  const to=(Array.isArray(payload.to) ? payload.to : String(payload.to ?? "").split(/[,;\s]+/)).map((value)=>textValue(value).toLowerCase()).filter(Boolean);
  const subject=textValue(payload.subject); const html=typeof payload.html==="string" ? payload.html : "";
  if (!to.length || !to.every(validEmail)) return json({ error:"Enter one or more valid email addresses." },400);
  if (to.length>50) return json({ error:"A report can be sent to at most 50 people at once." },400);
  if (!subject || !html) return json({ error:"The report is empty." },400);
  await ctx.mailer.send({ to,subject,html });
  await writeAudit(ctx,auth.user,{ action:"REPORT_EMAILED",entityType:"report",summary:`Emailed “${subject}” to ${to.join(", ")}` });
  return json({ sent:to.length });
}

async function usersDirectory(_request:Request,ctx:ApiContext) {
  const auth=await requireAppUser(ctx);
  if ("error" in auth) return auth.error;
  const users=await ctx.store.table("users").all();
  return json({ users:users.filter((user)=>user.status==="active").map((user)=>({ email:user.email,fullName:user.fullName })).sort((a,b)=>a.fullName.localeCompare(b.fullName)) });
}

// ------------------------------------------------------------------- router

type Handler = (request:Request,ctx:ApiContext)=>Promise<Response>;
const routes:Record<string,Partial<Record<string,Handler>>>={
  "/api/session":{ GET:sessionGet,POST:sessionPost },
  "/api/register":{ GET:registerGet,POST:registerPost,PATCH:registerPatch,DELETE:registerDelete },
  "/api/batches":{ GET:batchesGet,POST:batchesPost },
  "/api/admin":{ GET:adminGet,POST:adminPost,PATCH:adminPatch },
  "/api/admin/deletions":{ PATCH:deletionsPatch },
  "/api/admin/import":{ POST:adminImport },
  "/api/email-report":{ POST:emailReport },
  "/api/users":{ GET:usersDirectory },
};

export async function handleApi(request:Request,ctx:ApiContext) {
  const route=routes[new URL(request.url).pathname];
  const handler=route?.[request.method];
  if (!handler) return json({ error:"Not found" },route ? 405 : 404);
  try { return await handler(request,ctx); }
  catch (error) { ctx.log?.(`API ${request.method} ${new URL(request.url).pathname} failed`,error); return errorResponse(error,"Unexpected error"); }
}
