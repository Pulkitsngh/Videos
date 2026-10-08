// End-to-end check of the packaged UI + API in demo mode (no Microsoft 365 needed).
// Run with: xvfb-run -a node tests/e2e.mjs [screenshotDir]
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { _electron as electron } from "playwright";

const root=path.resolve(import.meta.dirname,"..");
const shots=process.argv[2] ?? path.join(os.tmpdir(),"vermitrack-shots");
fs.mkdirSync(shots,{ recursive:true });
const userData=fs.mkdtempSync(path.join(os.tmpdir(),"vermitrack-e2e-"));
const check=(condition,message)=>{ if (!condition) throw new Error(`FAILED: ${message}`); console.log(`ok - ${message}`); };

async function launch(env={}) {
  const app=await electron.launch({ args:[root,"--no-sandbox"],env:{ ...process.env,VERMITRACK_USER_DATA:userData,...env } });
  const page=await app.firstWindow();
  page.on("pageerror",(error)=>{ throw error; });
  await page.setViewportSize({ width:1320,height:880 });
  return { app,page };
}
const api=(page,url,init)=>page.evaluate(async([url,init])=>{ const response=await fetch(url,init); return { status:response.status,body:await response.json() }; },[url,init]);

// 1. First launch shows setup, demo mode loads the imported register.
let { app,page }=await launch();
await page.getByText("I have a setup code").waitFor();
await page.screenshot({ path:path.join(shots,"01-setup.png") });
await page.getByRole("tab",{ name:"First-time setup (administrator)" }).click();
await page.getByPlaceholder("1a2b3c4d-…").fill("not-a-guid");
await page.getByRole("button",{ name:"Save and continue" }).click();
await page.getByText("Application (client) ID must look like").waitFor();
check(true,"setup form validates the client ID");
await page.getByRole("button",{ name:/Try demo mode/ }).click();
await page.getByText(/Good (morning|afternoon|evening), Pulkit/).waitFor({ timeout:15000 });
await page.getByText("Demo mode · records are saved on this computer only").waitFor();
await page.waitForTimeout(800);
await page.screenshot({ path:path.join(shots,"02-home.png") });
let data=await api(page,"/api/register");
check(data.status===200 && data.body.entries.length===149,`register loads 149 imported entries (got ${data.body.entries?.length})`);
check(data.body.batches.length===3 && data.body.baselines.length===1,"3 batches and the Batch 4 baseline are present");

// 2. Capture a raw material entry through the UI.
await page.getByRole("button",{ name:/Material receipt/ }).click();
await page.locator("#field-materialType").selectOption("Cow dung");
await page.locator("#field-supplier").fill("Shree Dairy");
await page.locator("#field-quantityKg").fill("1200");
await page.locator("#field-ratePerUnit").fill("1.5");
await page.screenshot({ path:path.join(shots,"03-capture.png") });
await page.getByRole("button",{ name:/Save in Batch 4/ }).click();
await page.getByText("Entry saved in Batch 4").waitFor();
data=await api(page,"/api/register?module=raw-material");
const created=data.body.entries.find((entry)=>entry.title==="Cow dung from Shree Dairy");
check(created && created.amount===1800 && created.createdByEmail==="pulkit.singh@lnbgroup.com","new entry saved with computed amount and author");
check(JSON.parse(created.details)._recordCode===`BATCH-04-RM-${String(created.id).padStart(4,"0")}`,"auto register code assigned");
check(/^RM\/\d{8}\/CODU$/.test(created.referenceCode),`RM lot code generated (${created.referenceCode})`);

// 3. Records view, deletion request via the in-app dialog, approval in Admin.
await page.getByRole("button",{ name:/Records/ }).first().click();
await page.getByPlaceholder("Search lot, finished batch, bed…").fill("Shree Dairy");
await page.locator("details.record-card summary").first().click();
await page.getByRole("button",{ name:"Request deletion" }).first().click();
await page.getByRole("dialog").getByText("Request deletion").waitFor();
await page.screenshot({ path:path.join(shots,"04-delete-dialog.png") });
await page.getByRole("dialog").locator("textarea").fill("Duplicate");
await page.getByRole("button",{ name:"Send deletion request" }).click();
await page.getByText(/Deletion request sent/).waitFor();
const outbox=fs.readFileSync(path.join(userData,"demo-outbox.jsonl"),"utf8").trim().split("\n").map((line)=>JSON.parse(line));
check(outbox.some((mail)=>mail.subject.includes("deletion approval needed") && mail.to.includes("pritam.suryavanshi@lnbgroup.com")),"approvers are emailed about the deletion request");
await page.getByRole("button",{ name:/Open administration/ }).click();
await page.getByText("Deletion approvals").waitFor();
await page.getByRole("button",{ name:"Approve deletion" }).first().click();
await page.getByText("No deletion requests are waiting for approval.").waitFor();
await page.getByText("90-day recycle bin").waitFor();
await page.screenshot({ path:path.join(shots,"05-admin.png"),fullPage:false });
data=await api(page,"/api/register");
check(!data.body.entries.some((entry)=>entry.id===created.id),"approved deletion hides the entry");
await page.getByRole("button",{ name:"Restore entry" }).first().click();
await page.waitForTimeout(500);
data=await api(page,"/api/register");
check(data.body.entries.some((entry)=>entry.id===created.id),"restore brings the entry back");

// 4. Reports and the Outlook email sheet.
await page.getByRole("button",{ name:/Reports/ }).first().click();
await page.getByText("Batch 4 reports").waitFor();
await page.screenshot({ path:path.join(shots,"06-reports.png") });
await page.getByRole("button",{ name:"✉ Email report" }).click();
await page.getByRole("dialog").getByText("Pritam Suryavanshi").waitFor();
await page.getByRole("dialog").getByText("Pritam Suryavanshi").click();
await page.screenshot({ path:path.join(shots,"07-email-report.png") });
await page.getByRole("button",{ name:"Send from my Outlook" }).click();
await page.getByText(/Report emailed to 1 person/).waitFor();
const reportMail=fs.readFileSync(path.join(userData,"demo-outbox.jsonl"),"utf8").trim().split("\n").map((line)=>JSON.parse(line)).at(-1);
check(reportMail.subject.startsWith("VermiTrack report · Batch 4") && reportMail.html.includes("Raw material procured"),"report email contains the batch overview");

// 5. New batch + More/connection page.
await page.getByRole("button",{ name:"＋ New batch" }).first().click();
await page.getByRole("button",{ name:"Create & select batch" }).click();
await page.getByText("Batch 6 created and selected").waitFor();
await page.getByRole("button",{ name:/More/ }).first().click();
await page.getByText("Demo mode – not shared").waitFor();
await page.screenshot({ path:path.join(shots,"08-more.png") });

// 6. Edit rules and API validation.
const edit=await api(page,"/api/register",{ method:"PATCH",headers:{ "Content-Type":"application/json" },body:JSON.stringify({ id:created.id,module:"raw-material",productionBatchCode:"BATCH-04",values:{ ...JSON.parse(created.details),quantityKg:"1000" } }) });
check(edit.status===200 && edit.body.entry.amount===1500,"edit recalculates the amount");
const bad=await api(page,"/api/register",{ method:"POST",headers:{ "Content-Type":"application/json" },body:JSON.stringify({ module:"sales",productionBatchCode:"NOPE",values:{ salesDate:"2026-01-01" } }) });
check(bad.status===400,"unknown production batch is rejected");
await app.close();

// 7. Data persists across restarts; an unapproved account is refused.
({ app,page }=await launch());
await page.getByText(/Good (morning|afternoon|evening), Pulkit/).waitFor({ timeout:15000 });
data=await api(page,"/api/register");
check(data.body.entries.length===150 && data.body.batches.length===4,"data persists after restart");
await app.close();
({ app,page }=await launch({ VERMITRACK_DEMO_EMAIL:"stranger@example.com" }));
await page.getByText(/is not an approved VermiTrack user/).waitFor({ timeout:15000 });
await page.screenshot({ path:path.join(shots,"09-denied.png") });
check(true,"unapproved Microsoft 365 account is refused");
await app.close();
console.log(`\nAll end-to-end checks passed. Screenshots: ${shots}`);
