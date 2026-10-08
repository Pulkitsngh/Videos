/**
 * Storage layer. Every "table" of the original D1 database becomes a Microsoft
 * Lists / SharePoint list on the team's Microsoft 365 site, so colleagues work
 * on the same live data at the same time. A local JSON store with the same
 * interface backs demo mode and the automated tests.
 */
import fs from "node:fs";
import path from "node:path";
import type { Graph } from "./graph";

export type Iso = string;

export type EntryRow = {
  id:number; importKey:string | null; productionBatchCode:string; module:string; entryDate:string;
  referenceCode:string | null; batchCode:string | null; bedNumber:string | null; title:string;
  quantity:number | null; amount:number | null; status:string | null; details:string; createdBy:string;
  createdByEmail:string | null; deletionStatus:string; deletedAt:Iso | null; deletedBy:string | null;
  createdAt:Iso; updatedAt:Iso;
};
export type BatchRow = {
  id:number; code:string; name:string; financialYear:string | null; status:string; startDate:string | null;
  expectedEndDate:string | null; notes:string | null; createdBy:string; createdAt:Iso; updatedAt:Iso;
};
export type BaselineRow = {
  id:number; productionBatchCode:string; label:string; batchCode:string; status:string; rawMaterialKg:number;
  preCompostLots:number; totalBeds:number; activeBeds:number; rawYieldKg:number; netYieldKg:number;
  revenue:number; expenses:number; qcPassed:number; source:string; updatedAt:Iso;
};
export type UserRow = {
  id:number; email:string; fullName:string; mobile:string | null; designation:string; role:string;
  permittedSegments:string; permittedBatches:string; status:string; isKeyPerson:boolean;
  lastLoginAt:Iso | null; lastActiveAt:Iso | null; createdAt:Iso; updatedAt:Iso;
};
export type ActivityRow = {
  id:number; actorEmail:string; actorName:string; action:string; entityType:string; entityId:string | null;
  module:string | null; productionBatchCode:string | null; summary:string; beforeJson:string | null;
  afterJson:string | null; createdAt:Iso;
};
export type DeletionRow = {
  id:number; entryId:number; reason:string; status:string; requestedByEmail:string; requestedByName:string;
  decidedByEmail:string | null; decidedByName:string | null; decisionNote:string | null; decidedAt:Iso | null;
  createdAt:Iso; updatedAt:Iso;
};
export type AccessRow = {
  id:number; userId:number | null; email:string | null; fullName:string | null; outcome:string; method:string;
  ipAddress:string | null; deviceInfo:string | null; clientSessionId:string | null; lastSeenAt:Iso | null;
  endedAt:Iso | null; durationSeconds:number; createdAt:Iso;
};

export type Tables = {
  entries:EntryRow; batches:BatchRow; baselines:BaselineRow; users:UserRow;
  activity:ActivityRow; deletions:DeletionRow; access:AccessRow;
};
export type TableName = keyof Tables;
export type NewRow<T> = Omit<T,"id">;

export interface Table<T extends { id:number }> {
  /** All rows. Served from a short-lived cache unless `fresh` is set. */
  all(options?:{ fresh?:boolean }):Promise<T[]>;
  /** One row, always read fresh from storage. */
  get(id:number):Promise<T | undefined>;
  insert(row:NewRow<T>):Promise<T>;
  /** Read-modify-write against the latest stored copy of the row. */
  update(id:number,patch:Partial<NewRow<T>>):Promise<T>;
  insertMany(rows:NewRow<T>[]):Promise<void>;
}

export interface Store {
  readonly kind:"sharepoint" | "local";
  readonly location:string;
  table<K extends TableName>(name:K):Table<Tables[K]>;
}

type ColumnType = "text" | "note" | "number" | "boolean";
type ColumnDef = { field:string; column:string; displayName:string; type:ColumnType };
type TableDef = { listName:string; description:string; titleField:string; columns:ColumnDef[]; cacheMs:number };

const col=(field:string,displayName:string,type:ColumnType="text"):ColumnDef=>({ field,column:`vt${field[0].toUpperCase()}${field.slice(1)}`,displayName,type });

/**
 * SharePoint list layout. `vtData` holds the complete row as JSON and is the
 * source of truth; the other columns mirror the key fields so the lists stay
 * readable in SharePoint, Excel and Power BI.
 */
export const tableDefs:Record<TableName,TableDef>={
  entries:{ listName:"VermiTrack Entries",description:"VermiTrack register entries (all 8 registers).",titleField:"title",cacheMs:3000,columns:[
    col("module","Register"),col("productionBatchCode","Production batch"),col("entryDate","Entry date"),col("referenceCode","Reference"),
    col("batchCode","Finished batch"),col("bedNumber","Bed"),col("quantity","Quantity","number"),col("amount","Amount","number"),col("status","Entry status"),
    col("createdBy","Entered by"),col("createdByEmail","Entered by email"),col("deletionStatus","Deletion status"),col("importKey","Import key"),col("createdAt","Punched at"),
  ] },
  batches:{ listName:"VermiTrack Batches",description:"VermiTrack production batches.",titleField:"name",cacheMs:3000,columns:[
    col("code","Batch code"),col("financialYear","Financial year"),col("status","Batch status"),col("startDate","Start date"),col("expectedEndDate","Expected end"),col("createdBy","Created by"),
  ] },
  baselines:{ listName:"VermiTrack Baselines",description:"Imported batch baselines from the Excel registers.",titleField:"label",cacheMs:10000,columns:[
    col("productionBatchCode","Production batch"),col("source","Source workbook"),
  ] },
  users:{ listName:"VermiTrack Users",description:"Approved VermiTrack users and their access.",titleField:"fullName",cacheMs:3000,columns:[
    col("email","Email"),col("role","Role"),col("designation","Designation"),col("status","Account status"),col("permittedSegments","Permitted segments"),col("permittedBatches","Permitted batches"),col("lastActiveAt","Last active"),
  ] },
  activity:{ listName:"VermiTrack Activity",description:"VermiTrack audit trail.",titleField:"summary",cacheMs:60000,columns:[
    col("actorEmail","User email"),col("action","Action"),col("productionBatchCode","Production batch"),col("module","Register"),col("createdAt","Time"),
  ] },
  deletions:{ listName:"VermiTrack Deletion Requests",description:"Deletion requests and approvals.",titleField:"reason",cacheMs:3000,columns:[
    col("entryId","Entry ID","number"),col("status","Request status"),col("requestedByEmail","Requested by"),col("decidedByEmail","Decided by"),col("createdAt","Requested at"),
  ] },
  access:{ listName:"VermiTrack Sign-ins",description:"VermiTrack sign-in and session log.",titleField:"email",cacheMs:30000,columns:[
    col("outcome","Outcome"),col("deviceInfo","Device"),col("createdAt","Started"),col("lastSeenAt","Last seen"),col("endedAt","Ended"),col("durationSeconds","Duration (sec)","number"),
  ] },
};
export const tableNames=Object.keys(tableDefs) as TableName[];

// ---------------------------------------------------------------- local store

type LocalBucket = { nextId:number; rows:Array<{ id:number }> };
type LocalFile = { version:1; tables:Partial<Record<TableName,LocalBucket>> };

export class LocalStore implements Store {
  readonly kind="local" as const;
  private data:LocalFile;
  constructor(readonly location:string) {
    this.data=fs.existsSync(location) ? JSON.parse(fs.readFileSync(location,"utf8")) as LocalFile : { version:1,tables:{} };
  }
  isEmpty() { return !this.data.tables.users?.rows.length; }
  private save() {
    fs.mkdirSync(path.dirname(this.location),{ recursive:true });
    const temp=`${this.location}.tmp`;
    fs.writeFileSync(temp,JSON.stringify(this.data));
    fs.renameSync(temp,this.location);
  }
  table<K extends TableName>(name:K):Table<Tables[K]> {
    type T=Tables[K];
    const bucket=():LocalBucket=>this.data.tables[name] ??= { nextId:1,rows:[] };
    const clone=<V>(value:V)=>structuredClone(value);
    return {
      all:async()=>clone(bucket().rows as T[]),
      get:async(id)=>clone((bucket().rows as T[]).find((row)=>row.id===id)),
      insert:async(row)=>{ const b=bucket(); const created={ ...row,id:b.nextId++ } as T; b.rows.push(created as { id:number }); this.save(); return clone(created); },
      insertMany:async(rows)=>{ const b=bucket(); for (const row of rows) b.rows.push({ ...row,id:b.nextId++ } as { id:number }); this.save(); },
      update:async(id,patch)=>{
        const rows=bucket().rows as T[]; const index=rows.findIndex((row)=>row.id===id);
        if (index<0) throw new Error("This record no longer exists.");
        rows[index]={ ...rows[index],...patch,id };
        this.save(); return clone(rows[index]);
      },
    };
  }
}

// ----------------------------------------------------------- sharepoint store

export class NeedsProvisioningError extends Error {
  constructor(readonly missing:string[]) { super(`The SharePoint site is missing the VermiTrack lists: ${missing.join(", ")}.`); }
}

type GraphItem = { id:string; fields?:Record<string,unknown> };
const NOTE_LIMIT=60000;

function columnPayload(def:ColumnDef) {
  const base={ name:def.column,displayName:def.displayName };
  if (def.type==="number") return { ...base,number:{} };
  if (def.type==="boolean") return { ...base,boolean:{} };
  if (def.type==="note") return { ...base,text:{ allowMultipleLines:true,textType:"plain",linesForEditing:6 } };
  return { ...base,text:{ maxLength:255 } };
}

function mirrorValue(def:ColumnDef,value:unknown) {
  if (value===undefined || value===null || value==="") return def.type==="text" || def.type==="note" ? "" : null;
  if (def.type==="number") { const parsed=Number(value); return Number.isFinite(parsed) ? parsed : null; }
  if (def.type==="boolean") return Boolean(value);
  return String(value).slice(0,255);
}

/** Keeps the JSON payload inside SharePoint's multi-line text limit by trimming audit snapshots first. */
function serialise(row:Record<string,unknown>) {
  let json=JSON.stringify(row);
  if (json.length>NOTE_LIMIT && ("beforeJson" in row || "afterJson" in row)) json=JSON.stringify({ ...row,beforeJson:null,afterJson:null });
  if (json.length>NOTE_LIMIT) throw new Error("This record is too large to store. Please shorten the long text fields.");
  return json;
}

export class SharePointStore implements Store {
  readonly kind="sharepoint" as const;
  private listIds:Partial<Record<TableName,string>>={};
  private tables=new Map<TableName,Table<{ id:number }>>();
  constructor(private graph:Graph,private siteId:string,readonly location:string) {}

  /** Resolves a SharePoint site URL such as https://contoso.sharepoint.com/sites/VermiTrack to its Graph site id. */
  static async resolveSite(graph:Graph,siteUrl:string) {
    let url:URL;
    try { url=new URL(siteUrl.trim()); } catch { throw new Error("Enter the full SharePoint site address, for example https://yourcompany.sharepoint.com/sites/VermiTrack"); }
    if (!/\.sharepoint\.(com|cn|us|de)$/i.test(url.hostname)) throw new Error("The site address must be a SharePoint address ending in .sharepoint.com");
    const sitePath=url.pathname.replace(/\/(Lists|Shared Documents|SitePages|_layouts)\/.*$/i,"").replace(/\/+$/,"");
    const site=await graph.request<{ id:string; webUrl:string; displayName:string }>("GET",sitePath ? `/sites/${url.hostname}:${encodeURI(decodeURI(sitePath))}` : `/sites/${url.hostname}`);
    return site;
  }

  private async lists() {
    return this.graph.pages<{ id:string; displayName:string }>(`/sites/${this.siteId}/lists?$select=id,displayName&$top=200`);
  }

  async connect() {
    const lists=await this.lists();
    const missing:string[]=[];
    for (const name of tableNames) {
      const found=lists.find((list)=>list.displayName===tableDefs[name].listName);
      if (found) this.listIds[name]=found.id; else missing.push(tableDefs[name].listName);
    }
    if (missing.length) throw new NeedsProvisioningError(missing);
  }

  /** Creates any missing lists and columns. Needs the Sites.Manage.All permission. */
  async provision() {
    const lists=await this.lists();
    for (const name of tableNames) {
      const def=tableDefs[name];
      const existing=lists.find((list)=>list.displayName===def.listName);
      const columns=[...def.columns.map(columnPayload),columnPayload({ field:"data",column:"vtData",displayName:"Record data (JSON)",type:"note" })];
      if (!existing) {
        const created=await this.graph.request<{ id:string }>("POST",`/sites/${this.siteId}/lists`,{ displayName:def.listName,description:def.description,columns,list:{ template:"genericList" } });
        this.listIds[name]=created.id;
        continue;
      }
      this.listIds[name]=existing.id;
      const present=new Set((await this.graph.pages<{ name:string }>(`/sites/${this.siteId}/lists/${existing.id}/columns?$select=name`)).map((column)=>column.name));
      for (const column of columns) if (!present.has(column.name)) await this.graph.request("POST",`/sites/${this.siteId}/lists/${existing.id}/columns`,column);
    }
  }

  table<K extends TableName>(name:K):Table<Tables[K]> {
    const cached=this.tables.get(name);
    if (cached) return cached as unknown as Table<Tables[K]>;
    const created=this.createTable(name);
    this.tables.set(name,created as unknown as Table<{ id:number }>);
    return created;
  }

  private createTable<K extends TableName>(name:K):Table<Tables[K]> {
    type T=Tables[K];
    const def=tableDefs[name];
    const graph=this.graph;
    const listPath=()=>{ const id=this.listIds[name]; if (!id) throw new Error("VermiTrack is not connected to SharePoint yet."); return `/sites/${this.siteId}/lists/${id}`; };
    let cache:{ rows:T[]; at:number } | null=null;
    let inflight:Promise<T[]> | null=null;

    const fromItem=(item:GraphItem):T=>{
      const fields=item.fields ?? {};
      let row:Record<string,unknown>={};
      try { row=JSON.parse(String(fields.vtData ?? "")) as Record<string,unknown>; }
      catch { for (const column of def.columns) row[column.field]=fields[column.column] ?? null; row[def.titleField]=fields.Title ?? ""; }
      return { ...row,id:Number(item.id) } as T;
    };
    const toFields=(row:Record<string,unknown>)=>{
      const { id:_id,...data }=row;
      const fields:Record<string,unknown>={ Title:String(data[def.titleField] ?? "").slice(0,255) || "(untitled)",vtData:serialise(data) };
      for (const column of def.columns) fields[column.column]=mirrorValue(column,data[column.field]);
      return fields;
    };
    const remember=(row:T)=>{ if (!cache) return; const index=cache.rows.findIndex((item)=>item.id===row.id); if (index<0) cache.rows.push(row); else cache.rows[index]=row; };
    const load=async()=>{
      const items=await graph.pages<GraphItem>(`${listPath()}/items?expand=fields(select=Title,vtData,${def.columns.map((column)=>column.column).join(",")})&$top=999`);
      const rows=items.map(fromItem).sort((a,b)=>a.id-b.id);
      cache={ rows,at:Date.now() };
      return rows;
    };

    return {
      all:async(options)=>{
        if (!options?.fresh && cache && Date.now()-cache.at<def.cacheMs) return structuredClone(cache.rows);
        inflight ??= load().finally(()=>{ inflight=null; });
        return structuredClone(await inflight);
      },
      get:async(id)=>{
        try { const row=fromItem(await graph.request<GraphItem>("GET",`${listPath()}/items/${id}?expand=fields`)); remember(row); return row; }
        catch (error) { if ((error as { status?:number }).status===404) return undefined; throw error; }
      },
      insert:async(row)=>{
        const created=await graph.request<GraphItem>("POST",`${listPath()}/items`,{ fields:toFields(row as Record<string,unknown>) });
        const result={ ...row,id:Number(created.id) } as T;
        remember(result); return structuredClone(result);
      },
      insertMany:async(rows)=>{
        await graph.batch(rows.map((row)=>({ method:"POST",url:`${listPath()}/items`,body:{ fields:toFields(row as Record<string,unknown>) } })));
        cache=null;
      },
      update:async(id,patch)=>{
        const current=fromItem(await graph.request<GraphItem>("GET",`${listPath()}/items/${id}?expand=fields`));
        const next={ ...current,...patch,id } as T;
        await graph.request("PATCH",`${listPath()}/items/${id}/fields`,toFields(next as Record<string,unknown>));
        remember(next); return structuredClone(next);
      },
    };
  }
}

// --------------------------------------------------------------------- seeding

type Seed = {
  batches:Array<Omit<BatchRow,"id"|"createdAt"|"updatedAt">>;
  baselines:Array<Omit<BaselineRow,"id"|"updatedAt">>;
  entries:Array<Omit<EntryRow,"id"|"createdByEmail"|"deletionStatus"|"deletedAt"|"deletedBy"|"createdAt"|"updatedAt">>;
  users:Array<Omit<UserRow,"id"|"isKeyPerson"|"lastLoginAt"|"lastActiveAt"|"createdAt"|"updatedAt">>;
};

/**
 * Loads the historical Batch 3/4/5 register (imported from the original
 * Excel workbooks) the first time a store is set up. Each table is only
 * seeded while it is empty, so re-running setup never duplicates data.
 */
export async function seedStore(store:Store,seed:Seed,keyPersonEmail:string) {
  const now=new Date().toISOString();
  const empty=async(name:TableName)=>(await store.table(name).all({ fresh:true })).length===0;
  if (await empty("users")) await store.table("users").insertMany(seed.users.map((user)=>({ ...user,isKeyPerson:user.email.toLowerCase()===keyPersonEmail,lastLoginAt:null,lastActiveAt:null,createdAt:now,updatedAt:now })));
  if (await empty("batches")) await store.table("batches").insertMany(seed.batches.map((batch)=>({ ...batch,createdAt:now,updatedAt:now })));
  if (await empty("baselines")) await store.table("baselines").insertMany(seed.baselines.map((baseline)=>({ ...baseline,updatedAt:now })));
  if (await empty("entries")) await store.table("entries").insertMany(seed.entries.map((entry)=>({ ...entry,createdByEmail:null,deletionStatus:"active",deletedAt:null,deletedBy:null,createdAt:now,updatedAt:now })));
}
