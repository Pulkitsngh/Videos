import { app, BrowserWindow, Menu, net, protocol, safeStorage, shell } from "electron";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { KEY_PERSON_EMAIL } from "../shared/policy";
import { handleApi, type ApiContext, type Identity, type Mailer } from "./api";
import { Auth, BASE_SCOPES, SETUP_SCOPES } from "./auth";
import { Graph, GraphError } from "./graph";
import { invitationEmail } from "./mail";
import seed from "./seed.json";
import { LocalStore, NeedsProvisioningError, SharePointStore, seedStore, type Store } from "./store";

type Settings = { mode:"sharepoint" | "demo"; clientId?:string; tenantId?:string; siteUrl?:string; siteId?:string; siteName?:string; downloadUrl?:string };
type Stage = "setup" | "signin" | "provision" | "ready" | "error";

if (process.env.VERMITRACK_USER_DATA) app.setPath("userData",path.resolve(process.env.VERMITRACK_USER_DATA));
protocol.registerSchemesAsPrivileged([{ scheme:"app",privileges:{ standard:true,secure:true,supportFetchAPI:true,corsEnabled:true,stream:true } }]);

const ORIGIN="app://vermitrack";
const RENDERER_DIR=path.join(__dirname,"..","dist");
const userFile=(name:string)=>path.join(app.getPath("userData"),name);
const log=(message:string,error?:unknown)=>{ console.error(`[VermiTrack] ${message}`,error ?? ""); try { fs.appendFileSync(userFile("vermitrack.log"),`${new Date().toISOString()} ${message} ${error instanceof Error ? error.stack ?? error.message : error ?? ""}\n`); } catch { /* logging is best effort */ } };
const json=(body:unknown,status=200)=>Response.json(body,{ status });

const state:{
  settings:Settings | null; auth:Auth | null; graph:Graph | null; store:Store | null; identity:Identity | null;
  stage:Stage; message:string; missingLists:string[];
}={ settings:null,auth:null,graph:null,store:null,identity:null,stage:"setup",message:"",missingLists:[] };
const memory:ApiContext["memory"]={ sessions:new Map(),touched:new Map() };

// ----------------------------------------------------------------- settings

function readSettings():Settings | null { try { return JSON.parse(fs.readFileSync(userFile("settings.json"),"utf8")) as Settings; } catch { return null; } }
function writeSettings(settings:Settings | null) {
  fs.mkdirSync(app.getPath("userData"),{ recursive:true });
  if (settings) fs.writeFileSync(userFile("settings.json"),JSON.stringify(settings,null,2)); else fs.rmSync(userFile("settings.json"),{ force:true });
  state.settings=settings;
}

const GUID=/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function encodeSetupCode(settings:Settings) { return `VT1-${Buffer.from(JSON.stringify({ c:settings.clientId,t:settings.tenantId,s:settings.siteUrl,d:settings.downloadUrl || undefined })).toString("base64url")}`; }
function decodeSetupCode(code:string) {
  const match=code.trim().match(/^VT1-([A-Za-z0-9_-]+)$/);
  if (!match) throw new Error("This setup code is not valid. Copy the whole code, starting with VT1-.");
  const value=JSON.parse(Buffer.from(match[1],"base64url").toString("utf8")) as { c?:string; t?:string; s?:string; d?:string };
  return { clientId:value.c ?? "",tenantId:value.t ?? "",siteUrl:value.s ?? "",downloadUrl:value.d ?? "" };
}

const keychain={
  available:()=>safeStorage.isEncryptionAvailable(),
  encrypt:(text:string)=>safeStorage.encryptString(text),
  decrypt:(data:Buffer)=>safeStorage.decryptString(data),
};

// -------------------------------------------------------------------- state

function demoIdentity():Identity {
  const email=(process.env.VERMITRACK_DEMO_EMAIL || KEY_PERSON_EMAIL).toLowerCase();
  return { emails:[email],fullName:null };
}

async function openDemo() {
  const store=new LocalStore(userFile("demo-data.json"));
  if (store.isEmpty()) await seedStore(store,seed,KEY_PERSON_EMAIL);
  state.store=store; state.identity=demoIdentity(); state.stage="ready"; state.message="";
}

/** Brings the app as far as it can go: settings → sign-in → SharePoint lists → ready. */
async function refreshState() {
  state.settings=readSettings();
  state.message=""; state.missingLists=[];
  const settings=state.settings;
  if (!settings) { state.stage="setup"; state.store=null; return; }
  if (settings.mode==="demo") { await openDemo(); return; }
  if (!settings.clientId || !settings.siteUrl) { state.stage="setup"; return; }
  if (!state.auth || state.auth.clientId!==settings.clientId || state.auth.tenantId!==(settings.tenantId ?? "")) {
    state.auth=new Auth(settings.clientId,settings.tenantId ?? "",userFile("msal-cache.bin"),keychain,async(url)=>{ await shell.openExternal(url); });
    state.graph=new Graph(()=>state.auth!.token(BASE_SCOPES),net.fetch as typeof fetch);
    state.store=null; state.identity=null;
  }
  if (!state.auth.signedIn) await state.auth.restore();
  if (!state.auth.signedIn) { state.stage="signin"; return; }
  try {
    const graph=state.graph!;
    if (!state.identity) {
      const me=await graph.me();
      state.identity={ emails:[...new Set([me.mail,me.userPrincipalName].filter((value):value is string=>Boolean(value)).map((value)=>value.toLowerCase()))],fullName:me.displayName };
    }
    if (!settings.siteId) {
      const site=await SharePointStore.resolveSite(graph,settings.siteUrl);
      writeSettings({ ...settings,siteId:site.id,siteName:site.displayName });
    }
    if (!state.store) {
      const store=new SharePointStore(graph,state.settings!.siteId!,state.settings!.siteUrl!);
      await store.connect();
      state.store=store;
    }
    state.stage="ready";
  } catch (error) {
    if (error instanceof NeedsProvisioningError) { state.stage="provision"; state.missingLists=error.missing; return; }
    if (error instanceof GraphError && error.status===401) { await state.auth.signOut(); state.stage="signin"; state.message=error.message; return; }
    log("Could not connect to Microsoft 365",error);
    state.stage="error"; state.message=error instanceof Error ? error.message : "Could not connect to Microsoft 365.";
  }
}

function statusPayload() {
  const settings=state.settings;
  return {
    stage:state.stage,message:state.message,missingLists:state.missingLists,mode:settings?.mode ?? null,
    siteUrl:settings?.siteUrl ?? null,siteName:settings?.siteName ?? null,account:state.identity?.emails[0] ?? state.auth?.username ?? null,
    version:app.getVersion(),dataLocation:state.store?.location ?? null,
    setupCode:settings?.mode==="sharepoint" && settings.clientId ? encodeSetupCode(settings) : null,downloadUrl:settings?.downloadUrl ?? null,
  };
}

const mailer:Mailer={
  async send(mail) {
    if (state.settings?.mode==="sharepoint" && state.graph) { await state.graph.sendMail(mail); return; }
    fs.appendFileSync(userFile("demo-outbox.jsonl"),`${JSON.stringify({ at:new Date().toISOString(),...mail })}\n`);
  },
};

function deviceInfo() {
  const platform=process.platform==="win32" ? "Windows" : process.platform==="darwin" ? "Mac" : "Linux";
  return `${platform} · VermiTrack desktop ${app.getVersion()} · ${os.hostname()}`;
}

// ------------------------------------------------------------- app routes

async function body(request:Request) { return await request.json().catch(()=>({})) as Record<string,unknown>; }
const text=(value:unknown)=>typeof value==="string" ? value.trim() : "";

async function handleAppRoute(request:Request,pathname:string):Promise<Response> {
  if (pathname==="/api/app/status" && request.method==="GET") { if (state.stage!=="ready") await refreshState(); return json(statusPayload()); }

  if (pathname==="/api/app/configure" && request.method==="POST") {
    const input=await body(request);
    const values=text(input.setupCode) ? decodeSetupCode(text(input.setupCode)) : { clientId:text(input.clientId),tenantId:text(input.tenantId),siteUrl:text(input.siteUrl),downloadUrl:text(input.downloadUrl) };
    if (!GUID.test(values.clientId)) return json({ error:"The Application (client) ID must look like 1a2b3c4d-1234-5678-9abc-0123456789ab." },400);
    if (values.tenantId && !GUID.test(values.tenantId) && !/^[a-z0-9.-]+\.[a-z]{2,}$/i.test(values.tenantId)) return json({ error:"The Directory (tenant) ID must be a GUID or your company domain." },400);
    if (!/^https:\/\/[^/]+\.sharepoint\.(com|cn|us|de)(\/|$)/i.test(values.siteUrl)) return json({ error:"Enter the SharePoint site address, for example https://yourcompany.sharepoint.com/sites/VermiTrack" },400);
    if (values.downloadUrl && !/^https:\/\//i.test(values.downloadUrl)) return json({ error:"The installer download link must start with https://" },400);
    writeSettings({ mode:"sharepoint",clientId:values.clientId,tenantId:values.tenantId,siteUrl:values.siteUrl.replace(/\/+$/,""),downloadUrl:values.downloadUrl || undefined });
    state.auth=null; state.store=null; state.identity=null;
    await refreshState(); return json(statusPayload());
  }

  if (pathname==="/api/app/demo" && request.method==="POST") { writeSettings({ mode:"demo" }); await refreshState(); return json(statusPayload()); }

  if (pathname==="/api/app/reset" && request.method==="POST") {
    await state.auth?.signOut().catch(()=>undefined);
    writeSettings(null); Object.assign(state,{ auth:null,graph:null,store:null,identity:null });
    await refreshState(); return json(statusPayload());
  }

  if (pathname==="/api/app/signin" && request.method==="POST") {
    await refreshState();
    if (!state.auth) return json({ error:"Finish the connection settings first." },400);
    try { await state.auth.signIn(); }
    catch (error) { log("Sign-in failed",error); return json({ ...statusPayload(),error:error instanceof Error ? error.message : "Sign-in did not complete." },400); }
    state.identity=null; state.store=null; memory.touched.clear();
    await refreshState(); return json(statusPayload());
  }

  if (pathname==="/api/app/signout" && request.method==="POST") {
    if (state.settings?.mode==="demo") { writeSettings(null); Object.assign(state,{ store:null,identity:null }); }
    else { await state.auth?.signOut(); Object.assign(state,{ store:null,identity:null }); }
    memory.sessions.clear(); memory.touched.clear();
    await refreshState(); return json(statusPayload());
  }

  if (pathname==="/api/app/provision" && request.method==="POST") {
    const settings=state.settings;
    if (!state.auth?.signedIn || !settings?.siteId) return json({ error:"Sign in first." },400);
    try {
      const manageGraph=new Graph(()=>state.auth!.token([...BASE_SCOPES,...SETUP_SCOPES]),net.fetch as typeof fetch);
      const store=new SharePointStore(manageGraph,settings.siteId,settings.siteUrl!);
      await store.provision();
      await seedStore(store,seed,KEY_PERSON_EMAIL);
      state.store=null;
      await refreshState(); return json(statusPayload());
    } catch (error) {
      log("SharePoint setup failed",error);
      const message=error instanceof Error ? error.message : "SharePoint setup failed.";
      return json({ ...statusPayload(),error:/accessDenied|permission/i.test(message) ? `${message} Setting up the lists needs a site owner of ${settings.siteUrl}.` : message },400);
    }
  }

  if (pathname==="/api/app/invite" && request.method==="POST") {
    const input=await body(request);
    const settings=state.settings;
    if (!state.store || !state.identity) return json({ error:"Sign in first." },401);
    const users=await state.store.table("users").all();
    const me=users.find((user)=>state.identity!.emails.includes(user.email.toLowerCase()));
    if (!me || me.email.toLowerCase()!==KEY_PERSON_EMAIL) return json({ error:"Only the Key Person can send invitations." },403);
    const invitee=users.find((user)=>user.email.toLowerCase()===text(input.email).toLowerCase());
    if (!invitee) return json({ error:"Add this person as an approved user first." },404);
    const setupCode=settings?.mode==="sharepoint" ? encodeSetupCode(settings) : "VT1-demo";
    await mailer.send({ to:[invitee.email],...invitationEmail({ fullName:invitee.fullName,invitedBy:me.fullName,setupCode,downloadUrl:settings?.downloadUrl ?? null }) });
    return json({ sent:true });
  }

  if (pathname==="/api/app/open" && request.method==="POST") {
    const url=text((await body(request)).url);
    if (!/^https:\/\//i.test(url)) return json({ error:"Only https links can be opened." },400);
    await shell.openExternal(url); return json({ opened:true });
  }

  return json({ error:"Not found" },404);
}

async function handleRequest(request:Request) {
  const url=new URL(request.url);
  if (url.pathname.startsWith("/api/app/")) {
    try { return await handleAppRoute(request,url.pathname); }
    catch (error) { log(`App route ${url.pathname} failed`,error); return json({ error:error instanceof Error ? error.message : "Unexpected error" },500); }
  }
  if (url.pathname.startsWith("/api/")) {
    if (state.stage!=="ready" || !state.store) return json({ error:"VermiTrack is not connected yet.",code:"UNAUTHENTICATED" },401);
    return handleApi(request,{ store:state.store,identity:state.identity,deviceInfo:deviceInfo(),mailer,memory,log });
  }
  const relative=decodeURIComponent(url.pathname).replace(/^\/+/,"") || "index.html";
  const file=path.normalize(path.join(RENDERER_DIR,relative));
  if (!file.startsWith(RENDERER_DIR) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) return net.fetch(pathToFileURL(path.join(RENDERER_DIR,"index.html")).toString());
  return net.fetch(pathToFileURL(file).toString());
}

// ------------------------------------------------------------------ window

function createWindow() {
  const window=new BrowserWindow({
    width:1320,height:880,minWidth:380,minHeight:560,title:"Navjyoti VermiTrack",backgroundColor:"#f5f4ef",
    icon:path.join(RENDERER_DIR,"navjyoti-app-icon-512.png"),autoHideMenuBar:true,
    webPreferences:{ contextIsolation:true,sandbox:true,nodeIntegration:false,spellcheck:true },
  });
  window.webContents.setWindowOpenHandler(({ url })=>{ if (/^https:\/\//i.test(url)) void shell.openExternal(url); return { action:"deny" }; });
  window.webContents.on("will-navigate",(event,url)=>{ if (!url.startsWith(ORIGIN)) { event.preventDefault(); if (/^https:\/\//i.test(url)) void shell.openExternal(url); } });
  void window.loadURL(`${ORIGIN}/`);
  return window;
}

function buildMenu() {
  const isMac=process.platform==="darwin";
  Menu.setApplicationMenu(Menu.buildFromTemplate([
    ...(isMac ? [{ role:"appMenu" as const }] : []),
    { label:"File",submenu:[{ label:"Open data folder",click:()=>void shell.openPath(app.getPath("userData")) },{ type:"separator" },isMac ? { role:"close" } : { role:"quit" }] },
    { role:"editMenu" },
    { label:"View",submenu:[{ role:"reload" },{ role:"forceReload" },{ type:"separator" },{ role:"resetZoom" },{ role:"zoomIn" },{ role:"zoomOut" },{ type:"separator" },{ role:"togglefullscreen" },{ role:"toggleDevTools" }] },
    { role:"windowMenu" },
  ]));
}

if (!app.requestSingleInstanceLock()) app.quit();
else {
  app.on("second-instance",()=>{ const [window]=BrowserWindow.getAllWindows(); if (window) { if (window.isMinimized()) window.restore(); window.focus(); } });
  app.whenReady().then(async()=>{
    protocol.handle("app",handleRequest);
    buildMenu();
    await refreshState().catch((error)=>log("Startup failed",error));
    createWindow();
    app.on("activate",()=>{ if (!BrowserWindow.getAllWindows().length) createWindow(); });
  });
  app.on("window-all-closed",()=>{ if (process.platform!=="darwin") app.quit(); });
}
