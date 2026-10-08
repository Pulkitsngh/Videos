/** HTML emails sent through the signed-in user's Outlook (Microsoft Graph sendMail). */
import { DELETION_APPROVER_NAMES } from "../shared/policy";
import type { EntryRow } from "./store";

export const escapeHtml=(value:unknown)=>String(value ?? "").replace(/[&<>"']/g,(char)=>({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" })[char]!);

export function emailShell(heading:string,body:string) {
  return `<div style="font-family:Segoe UI,Arial,sans-serif;color:#18332b;max-width:640px">
<div style="background:#163c2e;color:#fff;padding:14px 18px;border-radius:10px 10px 0 0"><b>Navjyoti VermiTrack</b></div>
<div style="border:1px solid #e3e7e1;border-top:0;padding:18px;border-radius:0 0 10px 10px;background:#fffefa">
<h2 style="margin:0 0 12px;font-size:18px">${escapeHtml(heading)}</h2>${body}
<p style="color:#708078;font-size:12px;margin-top:20px">Sent from the VermiTrack desktop app.</p></div></div>`;
}

const entryTable=(entry:EntryRow)=>`<table style="border-collapse:collapse;font-size:14px">${[
  ["Entry",entry.title],["Production batch",entry.productionBatchCode],["Register",entry.module],["Entry date",entry.entryDate],["Entered by",entry.createdBy],
].map(([label,value])=>`<tr><td style="padding:4px 12px 4px 0;color:#708078">${escapeHtml(label)}</td><td style="padding:4px 0"><b>${escapeHtml(value)}</b></td></tr>`).join("")}</table>`;

export function deletionRequestEmail(input:{ entry:EntryRow; reason:string; requestedBy:string }) {
  return {
    subject:`VermiTrack: deletion approval needed – ${input.entry.title}`,
    html:emailShell("Deletion approval needed",`<p>${escapeHtml(input.requestedBy)} has asked to delete this register entry:</p>${entryTable(input.entry)}
<p><b>Reason:</b> “${escapeHtml(input.reason)}”</p><p>Open VermiTrack → your initials (top right) → <b>Deletion approvals</b> to approve or reject. Only ${escapeHtml(DELETION_APPROVER_NAMES)} can decide.</p>`),
  };
}

export function deletionDecisionEmail(input:{ entry:EntryRow; approved:boolean; decidedBy:string; note:string | null }) {
  const verb=input.approved ? "approved" : "rejected";
  return {
    subject:`VermiTrack: deletion request ${verb} – ${input.entry.title}`,
    html:emailShell(`Your deletion request was ${verb}`,`${entryTable(input.entry)}<p>${escapeHtml(input.decidedBy)} ${verb} the request.${input.note ? ` Note: “${escapeHtml(input.note)}”` : ""}</p>
${input.approved ? "<p>The entry has moved to the 90-day recycle bin and can still be restored by an approver.</p>" : "<p>The entry remains active in the register.</p>"}`),
  };
}

export function invitationEmail(input:{ fullName:string; invitedBy:string; setupCode:string; downloadUrl:string | null }) {
  return {
    subject:"You have been added to Navjyoti VermiTrack",
    html:emailShell(`Welcome to VermiTrack, ${input.fullName}`,`<p>${escapeHtml(input.invitedBy)} has given you access to the shared VermiTrack production register.</p>
<ol style="line-height:1.7"><li>${input.downloadUrl ? `Download and install VermiTrack: <a href="${escapeHtml(input.downloadUrl)}">${escapeHtml(input.downloadUrl)}</a>` : "Install the VermiTrack desktop app (ask the sender for the installer)."}</li>
<li>Open VermiTrack and choose <b>I have a setup code</b>.</li><li>Paste this setup code:<br><code style="display:block;word-break:break-all;background:#f5f4ef;padding:10px;border-radius:6px;margin-top:6px">${escapeHtml(input.setupCode)}</code></li>
<li>Sign in with your Microsoft 365 / Outlook work account.</li></ol>`),
  };
}
