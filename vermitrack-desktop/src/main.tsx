import { StrictMode, useCallback, useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import Home from "./App";
import "./globals.css";
import "./desktop.css";

export type AppStatus = {
  stage:"setup" | "signin" | "provision" | "ready" | "error"; message:string; missingLists:string[]; mode:"sharepoint" | "demo" | null;
  siteUrl:string | null; siteName:string | null; account:string | null; version:string; dataLocation:string | null;
  setupCode:string | null; downloadUrl:string | null; error?:string;
};

async function appCall(path:string,body?:unknown) {
  const response=await fetch(`/api/app/${path}`,body===undefined ? { cache:"no-store" } : { method:"POST",headers:{ "Content-Type":"application/json" },body:JSON.stringify(body) });
  const payload=await response.json() as AppStatus & { error?:string };
  if (!response.ok && !payload.stage) throw new Error(payload.error ?? "Something went wrong.");
  return payload;
}

function Root() {
  const [status,setStatus]=useState<AppStatus | null>(null);
  const [busy,setBusy]=useState(false);
  const [error,setError]=useState("");

  const run=useCallback(async(path:string,body?:unknown)=>{
    setBusy(true); setError("");
    try { const next=await appCall(path,body); if (next.stage) setStatus(next); if (next.error) setError(next.error); }
    catch (caught) { setError(caught instanceof Error ? caught.message : "Something went wrong."); }
    finally { setBusy(false); }
  },[]);

  useEffect(()=>{ void run("status"); },[run]);

  if (!status) return <Shell><div className="access-spinner" /><p>Starting VermiTrack…</p></Shell>;
  if (status.stage==="ready") return <Home app={status} onSignOut={()=>void run("signout",{})} />;
  if (status.stage==="setup") return <SetupScreen busy={busy} error={error} onConfigure={(values)=>void run("configure",values)} onDemo={()=>void run("demo",{})} />;
  if (status.stage==="signin") return <Shell eyebrow="MICROSOFT 365 SIGN-IN">
    <p>Sign in with your Microsoft 365 / Outlook work account. A browser window will open; come back here when it says you are signed in.</p>
    {(error || status.message) && <div className="form-error" role="alert">! {error || status.message}</div>}
    <button className="access-primary" disabled={busy} onClick={()=>void run("signin",{})}>{busy ? "Waiting for the browser sign-in…" : "Sign in with Microsoft 365"}</button>
    <button className="access-secondary access-secondary-button" disabled={busy} onClick={()=>void run("reset",{})}>Change connection settings</button>
    <footer><span>Shared site</span><b>{status.siteUrl}</b></footer>
  </Shell>;
  if (status.stage==="provision") return <Shell eyebrow="ONE-TIME SHAREPOINT SETUP">
    <p>Signed in as <b>{status.account}</b>. The SharePoint site does not have the VermiTrack lists yet. A site owner sets them up once; this also loads the imported Batch 3–5 register and the approved user list.</p>
    {error && <div className="form-error" role="alert">! {error}</div>}
    <button className="access-primary" disabled={busy} onClick={()=>void run("provision",{})}>{busy ? "Creating lists and loading records… (1–2 minutes)" : "Set up VermiTrack on this SharePoint site"}</button>
    <button className="access-secondary access-secondary-button" disabled={busy} onClick={()=>void run("signout",{})}>Sign in with a different account</button>
    <footer><span>Lists to be created</span><b>{status.missingLists.join(" · ")}</b></footer>
  </Shell>;
  return <Shell eyebrow="CONNECTION PROBLEM">
    <p>{status.message || "VermiTrack could not reach Microsoft 365."}</p>
    {error && <div className="form-error" role="alert">! {error}</div>}
    <button className="access-primary" disabled={busy} onClick={()=>void run("status")}>{busy ? "Checking…" : "Try again"}</button>
    <button className="access-secondary access-secondary-button" disabled={busy} onClick={()=>void run("signout",{})}>Sign out</button>
    <button className="access-secondary access-secondary-button" disabled={busy} onClick={()=>void run("reset",{})}>Change connection settings</button>
  </Shell>;
}

function Shell({ eyebrow="SECURE PRODUCTION REGISTER",children,wide=false }:{ eyebrow?:string; children:React.ReactNode; wide?:boolean }) {
  return <main className="access-screen"><section className={`access-card ${wide ? "setup-card" : ""}`}><div className="access-logo" role="img" aria-label="Navjyoti logo" /><p className="eyebrow">{eyebrow}</p><h1>Navjyoti VermiTrack</h1>{children}</section></main>;
}

function SetupScreen({ busy,error,onConfigure,onDemo }:{ busy:boolean; error:string; onConfigure:(values:Record<string,string>)=>void; onDemo:()=>void }) {
  const [mode,setMode]=useState<"code" | "admin">("code");
  const [code,setCode]=useState("");
  const [form,setForm]=useState({ clientId:"",tenantId:"",siteUrl:"",downloadUrl:"" });
  const field=(key:keyof typeof form,label:string,placeholder:string,hint:string)=><label><span>{label}</span><input value={form[key]} placeholder={placeholder} onChange={(event)=>setForm((current)=>({ ...current,[key]:event.target.value }))} /><small>{hint}</small></label>;
  return <Shell eyebrow="CONNECT TO MICROSOFT 365" wide>
    <p>VermiTrack keeps your team’s register on your company’s Microsoft 365 (SharePoint), so everyone sees the same live records and reports.</p>
    <div className="setup-tabs" role="tablist"><button role="tab" aria-selected={mode==="code"} className={mode==="code" ? "active" : ""} onClick={()=>setMode("code")}>I have a setup code</button><button role="tab" aria-selected={mode==="admin"} className={mode==="admin" ? "active" : ""} onClick={()=>setMode("admin")}>First-time setup (administrator)</button></div>
    {error && <div className="form-error" role="alert">! {error}</div>}
    {mode==="code" ? <form className="setup-form" onSubmit={(event)=>{ event.preventDefault(); onConfigure({ setupCode:code }); }}>
      <label><span>Setup code</span><textarea rows={3} value={code} onChange={(event)=>setCode(event.target.value)} placeholder="VT1-…" /><small>Your administrator can copy it from VermiTrack → More, or send it to you by email from the Admin page.</small></label>
      <button className="access-primary" type="submit" disabled={busy || !code.trim()}>{busy ? "Connecting…" : "Continue"}</button>
    </form> : <form className="setup-form" onSubmit={(event)=>{ event.preventDefault(); onConfigure(form); }}>
      {field("clientId","Application (client) ID*","1a2b3c4d-…","From the app registration in Microsoft Entra admin center → App registrations.")}
      {field("tenantId","Directory (tenant) ID*","5e6f7a8b-…","Shown on the same app registration overview page.")}
      {field("siteUrl","SharePoint site address*","https://yourcompany.sharepoint.com/sites/VermiTrack","The team site where the shared register will live. Colleagues must be members of this site.")}
      {field("downloadUrl","Installer download link (optional)","https://…","Included in invitation emails so colleagues can install the app.")}
      <button className="access-primary" type="submit" disabled={busy}>{busy ? "Saving…" : "Save and continue"}</button>
    </form>}
    <button className="access-secondary access-secondary-button" disabled={busy} onClick={onDemo}>Try demo mode (data stays on this computer only)</button>
    <footer><span>Need help?</span><b>See SETUP-GUIDE.md that came with VermiTrack for the 10-minute Microsoft 365 setup.</b></footer>
  </Shell>;
}

createRoot(document.getElementById("root")!).render(<StrictMode><Root /></StrictMode>);
