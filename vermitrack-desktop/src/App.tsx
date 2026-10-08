import { useCallback, useEffect, useMemo, useState } from "react";
import { registerModuleMap, registerModules, type RegisterField, type RegisterModule } from "./register-config";
import { DELETION_APPROVER_NAMES, isDeletionApprover } from "../shared/policy";
import type { AppStatus } from "./main";

type View = "home" | "records" | "insights" | "more" | "admin";
type ProductionBatch = {
  code:string; name:string; financialYear:string | null; status:string; startDate:string | null;
  expectedEndDate:string | null; notes:string | null; createdBy:string; createdAt:string | number;
};
type Baseline = {
  productionBatchCode:string; label:string; batchCode:string; status:string; rawMaterialKg:number;
  preCompostLots:number; totalBeds:number; activeBeds:number; rawYieldKg:number; netYieldKg:number;
  revenue:number; expenses:number; qcPassed:number; source:string;
};
type Entry = {
  id:number; productionBatchCode:string; module:string; entryDate:string; referenceCode:string | null;
  batchCode:string | null; bedNumber:string | null; title:string; quantity:number | null; amount:number | null;
  status:string | null; details:string; createdBy:string; createdByEmail?:string | null; deletionStatus?:string;
  deletedAt?:string | number | null; deletedBy?:string | null; createdAt:string | number;
};
type AppUser = {
  id:number; email:string; fullName:string; mobile:string | null; designation:string; role:string;
  permittedSegments:string; permittedBatches:string; status:string; lastLoginAt:string | number | null;
  lastActiveAt:string | number | null; isKeyPerson:boolean; createdAt?:string | number;
};
type ActivityLog = {
  id:number; actorEmail:string; actorName:string; action:string; entityType:string; entityId:string | null;
  module:string | null; productionBatchCode:string | null; summary:string; createdAt:string | number;
};
type AccessLog = {
  id:number; userId:number | null; email:string | null; fullName:string | null; outcome:"success" | "denied";
  method:"chatgpt" | "device" | "microsoft"; ipAddress:string | null; deviceInfo:string | null; lastSeenAt:string | number | null;
  endedAt:string | number | null; durationSeconds:number; createdAt:string | number;
};
type DeletionRequest = {
  id:number; entryId:number; reason:string; status:string; requestedByEmail:string; requestedByName:string;
  decidedByName:string | null; decisionNote:string | null; createdAt:string | number; decidedAt:string | number | null;
  entry:Entry | null;
};
type Metrics = {
  productionBatchCode:string; label:string; status:string; rawMaterialKg:number; preCompostLots:number;
  totalBeds:number; activeBeds:number; rawYieldKg:number; netYieldKg:number; packagedKg:number; revenue:number; expenses:number;
  qcPassed:number; source:string; recoveryPct:number; packagingPct:number; margin:number; entries:number;
};
type ReportMetric = { label:string; value:string; note?:string };
type RawMaterialComparison = {
  rawMaterialKg:number; productionKg:number; packagedKg:number;
  productionPct:number; packagedPct:number; packagingCoveragePct:number;
  awaitingProductionKg:number; awaitingPackagingKg:number; packagedDerived:boolean;
};
type QualityProfile = {
  code:"RM" | "FG"; label:string; entries:number; passed:number; failed:number; pending:number;
  fields:ConsolidatedField[];
};
type QualityComparisonRow = {
  key:string; label:string; group:string; unit?:string; rmAverage:number | null; rmCount:number;
  fgAverage:number | null; fgCount:number;
};
type QualityComparison = {
  rm:QualityProfile; fg:QualityProfile; rows:QualityComparisonRow[];
};
type ConsolidatedField = {
  key:string; label:string; group:string; type:RegisterField["type"]; unit?:string; count:number; missing:number;
  primary:string; detail:string; chartValue?:number; chartDisplay?:string; chartUnit?:string;
  breakdown:Array<{ label:string; value:number }>;
};
type SegmentReport = {
  id:string; short:string; title:string; tone:string; description:string; entryCount:number;
  hasData:boolean; includesBaseline:boolean; lastEntryDate:string | null; metrics:ReportMetric[];
  groups:Array<{ label:string; value:number }>; consolidatedFields:ConsolidatedField[]; entries:Entry[];
  rawMaterialComparison?:RawMaterialComparison; qualityComparison?:QualityComparison;
};

const fallbackBatches: ProductionBatch[] = [
  { code:"BATCH-05", name:"Batch 5", financialYear:"2026-27", status:"Planned", startDate:null, expectedEndDate:null, notes:"New production batch", createdBy:"System", createdAt:3 },
  { code:"BATCH-04", name:"Batch 4", financialYear:"2025-26", status:"Ongoing", startDate:null, expectedEndDate:null, notes:"Detailed Batch 4 register imported; production is ongoing.", createdBy:"System", createdAt:2 },
  { code:"BATCH-03", name:"Batch 3", financialYear:"2025-26", status:"Completed", startDate:null, expectedEndDate:null, notes:"Historical production batch", createdBy:"System", createdAt:1 },
];
const fallbackBaselines: Baseline[] = [{
  productionBatchCode:"BATCH-04", label:"Imported Batch 4 register", batchCode:"VCM-26/B04", status:"Ongoing",
  rawMaterialKg:161290, preCompostLots:3, totalBeds:80, activeBeds:0, rawYieldKg:7450, netYieldKg:6050,
  revenue:297780, expenses:361823, qcPassed:4, source:"Vermi register batch 4(3).xlsx",
}];

function parseDetails(entry: Entry) { try { return JSON.parse(entry.details) as Record<string, unknown>; } catch { return {}; } }
function number(value: unknown) { const parsed = Number(value); return Number.isFinite(parsed) ? parsed : 0; }
function capturedNumber(value:unknown) { if (value === "" || value === null || value === undefined) return null; const parsed=Number(value); return Number.isFinite(parsed) ? parsed : null; }
function packageWeightKg(value:unknown) { const match=String(value ?? "").match(/(\d+(?:\.\d+)?)\s*kg/i); return match ? Number(match[1]) : null; }
function packagedQuantity(entry:Entry) {
  const detail=parseDetails(entry); const explicit=capturedNumber(detail.packagedQuantityKg);
  if (explicit !== null) return { quantity:explicit, derived:false };
  const packets=capturedNumber(detail.packetsPrepared); const unitWeight=packageWeightKg(detail.packagingType);
  if (packets !== null && unitWeight !== null) return { quantity:packets*unitWeight, derived:true };
  if (/bulk|loose/i.test(String(detail.packagingType ?? ""))) return { quantity:number(detail.netYieldKg), derived:true };
  return { quantity:0, derived:true };
}
function formatNumber(value:number, maximumFractionDigits=0) { return new Intl.NumberFormat("en-IN", { maximumFractionDigits }).format(value); }
function formatMoney(value:number) { return new Intl.NumberFormat("en-IN", { style:"currency", currency:"INR", maximumFractionDigits:0 }).format(value); }
function formatCompactMoney(value:number) { return Math.abs(value) >= 100000 ? `₹${(value / 100000).toFixed(2)} L` : formatMoney(value); }
function todayIso() { const date = new Date(); const offset = date.getTimezoneOffset(); return new Date(date.getTime() - offset * 60000).toISOString().slice(0,10); }
function rmMaterialPrefix(materialType:string) { return materialType.split(/[^a-z0-9]+/i).filter(Boolean).map((word)=>word.slice(0,2)).join("").toUpperCase(); }
function rmLotCode(date:string,materialType:string) { const prefix=rmMaterialPrefix(materialType); if (!prefix) return ""; const digits=String(date || todayIso()).replace(/\D/g,"").slice(0,8); return `RM/${digits || todayIso().replaceAll("-","")}/${prefix}`; }
function accessSessionId() { let value=window.sessionStorage.getItem("vermitrack-access-session"); if (!value) { const random=typeof crypto!=="undefined" && typeof crypto.randomUUID==="function" ? crypto.randomUUID().replaceAll("-","") : `${Math.random().toString(36).slice(2)}${Math.random().toString(36).slice(2)}`; value=`${Date.now()}_${random}`; window.sessionStorage.setItem("vermitrack-access-session",value); } return value; }
async function endAccessSession() { const clientSessionId=window.sessionStorage.getItem("vermitrack-access-session"); if (!clientSessionId) return; window.sessionStorage.removeItem("vermitrack-access-session"); await fetch("/api/session",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({clientSessionId}),keepalive:true}).catch(()=>undefined); }
function formatDuration(totalSeconds:number) { const seconds=Math.max(0,Math.round(totalSeconds)); if (seconds<60) return seconds<10 ? "Under 10 sec" : `${seconds} sec`; const minutes=Math.floor(seconds/60); if (minutes<60) return `${minutes} min ${seconds%60 ? `${seconds%60} sec` : ""}`.trim(); const hours=Math.floor(minutes/60); return `${hours} hr ${minutes%60 ? `${minutes%60} min` : ""}`.trim(); }
function greetingForNow() { const hour=new Date().getHours(); return hour<12 ? "Good morning" : hour<17 ? "Good afternoon" : "Good evening"; }
function moduleFor(id:string) { return registerModuleMap[id] ?? registerModules[0]; }
function displayDate(value:string | null) { if (!value) return "—"; const parsed = new Date(`${value}T00:00:00`); return Number.isNaN(parsed.getTime()) ? value : parsed.toLocaleDateString("en-IN", { day:"2-digit", month:"short", year:"numeric" }); }
function displayTime(value:string | number) { const parsed = new Date(value); return Number.isNaN(parsed.getTime()) ? "" : parsed.toLocaleTimeString("en-IN", { hour:"2-digit", minute:"2-digit" }); }
function displayTimestamp(value:string | number) { const parsed=new Date(value); return Number.isNaN(parsed.getTime()) ? "Time unavailable" : parsed.toLocaleString("en-IN", { day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit",hour12:true }); }
function entryFieldLabel(moduleId:string,key:string) { return moduleFor(moduleId).fields.find((field)=>field.key===key)?.label ?? key.replace(/([A-Z])/g," $1").replace(/^./,(letter)=>letter.toUpperCase()); }
function capturedFieldValue(field:RegisterField | undefined,value:unknown) {
  if (value === "" || value === null || value === undefined) return "Not entered";
  const display=String(value);
  if (!field?.unit) return display;
  return field.unit === "₹" ? `₹${display}` : `${display} ${field.unit}`;
}
function entryCapturedFields(entry:Entry) {
  const detail=parseDetails(entry); const moduleDef=moduleFor(entry.module); const configuredKeys=new Set(moduleDef.fields.map((field)=>field.key));
  const configured=moduleDef.fields.map((field)=>({ key:field.key,label:field.label,group:field.group,value:capturedFieldValue(field,detail[field.key]) }));
  const computed=Object.entries(detail).filter(([key])=>!key.startsWith("_") && key!=="productionBatchCode" && !(entry.module==="harvest" && (key==="rawQtyKg" || key==="recoveryPct")) && !configuredKeys.has(key)).map(([key,value])=>({ key,label:entryFieldLabel(entry.module,key),group:"Calculated values",value:key === "totalAmount" || key === "totalRevenue" ? formatMoney(number(value)) : key === "closingStockKg" ? `${formatNumber(number(value),2)} kg` : key.toLowerCase().endsWith("pct") ? `${formatNumber(number(value),2)}%` : String(value) }));
  return [...configured,...computed];
}

function entrySystemIdentity(entry:Entry) {
  const detail=parseDetails(entry);
  const serial=number(detail._serialNo) || entry.id;
  const recordCode=String(detail._recordCode ?? `${entry.productionBatchCode}-${moduleFor(entry.module).short}-${String(serial).padStart(4,"0")}`);
  return { serial, recordCode };
}

function formatConsolidatedNumber(field:Pick<RegisterField,"unit">,value:number) {
  const display=new Intl.NumberFormat("en-IN",{maximumFractionDigits:2}).format(value);
  if (field.unit === "₹") return `₹${display}`;
  return field.unit ? `${display} ${field.unit}` : display;
}

function shouldAverageField(field:Pick<RegisterField,"key"|"unit">) {
  return field.unit === "%" || field.unit === "°C" || /rate|price|ph$|ec$|ratio|temperature|temp|moisture|nitrogen|phosph|potassium|calcium|manganese|iron|zinc|copper|ash|organic/i.test(field.key);
}

function buildConsolidatedFields(moduleId:string,items:Entry[]):ConsolidatedField[] {
  const moduleDef=moduleFor(moduleId);
  const details=items.map(parseDetails);
  const configuredKeys=new Set(moduleDef.fields.map((field)=>field.key));
  const computedKeys=[...new Set(details.flatMap((detail)=>Object.keys(detail).filter((key)=>!key.startsWith("_") && !configuredKeys.has(key) && key!=="productionBatchCode" && !(moduleId==="harvest" && (key==="rawQtyKg" || key==="recoveryPct")))))];
  const computedFields:RegisterField[]=computedKeys.map((key)=>({
    key,
    label:entryFieldLabel(moduleId,key),
    type:"number",
    group:"Calculated values",
    unit:key==="totalAmount" || key==="totalRevenue" ? "₹" : key==="closingStockKg" ? "kg" : key.toLowerCase().endsWith("pct") ? "%" : undefined,
  }));

  return [...moduleDef.fields,...computedFields].map((field)=>{
    const present=details.map((detail)=>detail[field.key]).filter((value)=>value!=="" && value!==null && value!==undefined);
    const missing=Math.max(0,items.length-present.length);
    if (field.type === "number") {
      const values=present.map(Number).filter(Number.isFinite);
      if (!values.length) return { key:field.key,label:field.label,group:field.group,type:field.type,unit:field.unit,count:0,missing:items.length,primary:"No value captured",detail:`0 of ${items.length} entries`,breakdown:[] };
      const total=values.reduce((sum,value)=>sum+value,0); const averageValue=total/values.length; const minimum=Math.min(...values); const maximum=Math.max(...values); const averageMode=shouldAverageField(field); const primaryValue=averageMode ? averageValue : total;
      return { key:field.key,label:field.label,group:field.group,type:field.type,unit:field.unit,count:values.length,missing,primary:formatConsolidatedNumber(field,primaryValue),detail:averageMode ? `Average · range ${formatConsolidatedNumber(field,minimum)} to ${formatConsolidatedNumber(field,maximum)}` : `Total · average ${formatConsolidatedNumber(field,averageValue)} · range ${formatConsolidatedNumber(field,minimum)} to ${formatConsolidatedNumber(field,maximum)}`,chartValue:Math.abs(primaryValue),chartDisplay:formatConsolidatedNumber(field,primaryValue),chartUnit:field.unit || "Number",breakdown:[] };
    }
    if (field.type === "date") {
      const values=present.map(String).sort();
      return { key:field.key,label:field.label,group:field.group,type:field.type,unit:field.unit,count:values.length,missing,primary:values.length ? values[0]===values.at(-1) ? displayDate(values[0]) : `${displayDate(values[0])} — ${displayDate(values.at(-1) ?? values[0])}` : "No value captured",detail:`${values.length} dated ${values.length===1 ? "entry" : "entries"}`,breakdown:[] };
    }
    const counts=new Map<string,number>();
    for (const raw of present) { const value=String(raw).trim(); if (value) counts.set(value,(counts.get(value) ?? 0)+1); }
    const breakdown=[...counts].map(([label,value])=>({label,value})).sort((a,b)=>b.value-a.value || a.label.localeCompare(b.label));
    const compact=breakdown.slice(0,4).map((item)=>`${item.label}${item.value>1 ? ` (${item.value})` : ""}`).join(" · ");
    return { key:field.key,label:field.label,group:field.group,type:field.type,unit:field.unit,count:present.length,missing,primary:compact || "No value captured",detail:breakdown.length>4 ? `${breakdown.length} distinct values · ${breakdown.length-4} more available` : `${breakdown.length} distinct ${breakdown.length===1 ? "value" : "values"}`,breakdown };
  });
}

function qualityProductType(entry:Entry) {
  const value=String(parseDetails(entry).productType ?? "").trim().toUpperCase();
  return value === "RM" || value === "FG" ? value : null;
}

function buildQualityComparison(items:Entry[]):QualityComparison {
  const profile=(code:"RM" | "FG",label:string):QualityProfile=>{
    const entries=items.filter((entry)=>qualityProductType(entry)===code);
    const passed=entries.filter((entry)=>String(entry.status ?? parseDetails(entry).fcoStatus ?? "").toLowerCase().includes("pass")).length;
    const failed=entries.filter((entry)=>String(entry.status ?? parseDetails(entry).fcoStatus ?? "").toLowerCase().includes("fail")).length;
    return { code,label,entries:entries.length,passed,failed,pending:Math.max(0,entries.length-passed-failed),fields:buildConsolidatedFields("quality",entries) };
  };
  const rm=profile("RM","Raw material"); const fg=profile("FG","Finished goods");
  const numericFields=moduleFor("quality").fields.filter((field)=>field.type==="number");
  const averageFor=(code:"RM" | "FG",key:string)=>{
    const values=items.filter((entry)=>qualityProductType(entry)===code).map((entry)=>capturedNumber(parseDetails(entry)[key])).filter((value):value is number=>value!==null);
    return { average:values.length ? values.reduce((total,value)=>total+value,0)/values.length : null,count:values.length };
  };
  const rows=numericFields.map((field)=>{ const rmValue=averageFor("RM",field.key); const fgValue=averageFor("FG",field.key); return { key:field.key,label:field.label,group:field.group,unit:field.unit,rmAverage:rmValue.average,rmCount:rmValue.count,fgAverage:fgValue.average,fgCount:fgValue.count }; });
  return { rm,fg,rows };
}

function buildSegmentReports(baseline:Baseline | undefined, entries:Entry[]):SegmentReport[] {
  const section=(id:string)=>entries.filter((entry)=>entry.module === id);
  const detailRows=(items:Entry[])=>items.map((entry)=>({ entry, detail:parseDetails(entry) }));
  const sum=(items:Entry[],key:string)=>detailRows(items).reduce((total,row)=>total+number(row.detail[key]),0);
  const average=(items:Entry[],key:string)=>{ const values=detailRows(items).map((row)=>row.detail[key]).filter((value)=>value !== "" && value !== null && value !== undefined).map(number); return values.length ? values.reduce((total,value)=>total+value,0)/values.length : 0; };
  const unique=(items:Entry[],key:string)=>new Set(detailRows(items).map((row)=>String(row.detail[key] ?? "").trim()).filter(Boolean)).size;
  const groups=(items:Entry[],key:string,fallbackToStatus=false)=>{ const counts=new Map<string,number>(); for (const {entry,detail} of detailRows(items)) { const value=String(detail[key] ?? (fallbackToStatus ? entry.status : "") ?? "").trim(); if (value) counts.set(value,(counts.get(value) ?? 0)+1); } return [...counts].map(([label,value])=>({label,value})).sort((a,b)=>b.value-a.value).slice(0,6); };
  const latest=(items:Entry[])=>items.map((entry)=>entry.entryDate).filter(Boolean).sort().at(-1) ?? null;
  const create=(id:string,metrics:ReportMetric[],reportGroups:Array<{label:string;value:number}>,includesBaseline=false,rawMaterialComparison?:RawMaterialComparison,qualityComparison?:QualityComparison):SegmentReport=>{ const moduleDef=moduleFor(id); const items=section(id); return { id,short:moduleDef.short,title:moduleDef.title,tone:moduleDef.tone,description:moduleDef.detail,entryCount:items.length,hasData:includesBaseline || items.length>0,includesBaseline,lastEntryDate:latest(items),metrics,groups:reportGroups,consolidatedFields:buildConsolidatedFields(id,items),entries:items,rawMaterialComparison,qualityComparison }; };
  const baselineValue=(items:Entry[],value:number | undefined)=>items.length ? 0 : value ?? 0;

  const raw=section("raw-material"); const rawMobileKg=sum(raw,"quantityKg"); const totalRawMaterialKg=baselineValue(raw,baseline?.rawMaterialKg)+rawMobileKg; const rawValue=raw.reduce((total,entry)=>total+number(entry.amount),0); const rawRate=rawMobileKg>0 ? rawValue/rawMobileKg : 0;
  const pre=section("pre-composting"); const preMobileLots=new Set(detailRows(pre).map((row)=>String(row.detail.lotCode ?? row.entry.referenceCode ?? "")).filter(Boolean)).size;
  const bed=section("bed-ops"); const bedStatus=new Map<string,string>(); for (const entry of bed) if (entry.bedNumber && !bedStatus.has(entry.bedNumber)) bedStatus.set(entry.bedNumber,String(parseDetails(entry).currentStatus ?? entry.status ?? "")); const activeBedCount=[...bedStatus.values()].filter((status)=>["filled","inoculated","active","ready for harvest"].includes(status.toLowerCase())).length;
  const harvest=section("harvest"); const netHarvest=baselineValue(harvest,baseline?.netYieldKg)+sum(harvest,"netYieldKg"); const packagedRows=harvest.map(packagedQuantity); const packagedHarvest=harvest.length ? packagedRows.reduce((total,row)=>total+row.quantity,0) : baseline?.netYieldKg ?? 0;
  const percentOfRawMaterial=(value:number)=>totalRawMaterialKg>0 ? (value/totalRawMaterialKg)*100 : 0;
  const rawMaterialComparison:RawMaterialComparison={
    rawMaterialKg:totalRawMaterialKg, productionKg:netHarvest, packagedKg:packagedHarvest,
    productionPct:percentOfRawMaterial(netHarvest), packagedPct:percentOfRawMaterial(packagedHarvest),
    packagingCoveragePct:netHarvest>0 ? (packagedHarvest/netHarvest)*100 : 0,
    awaitingProductionKg:Math.max(0,totalRawMaterialKg-netHarvest), awaitingPackagingKg:Math.max(0,netHarvest-packagedHarvest),
    packagedDerived:harvest.length ? packagedRows.some((row)=>row.derived) : Boolean(baseline),
  };
  const quality=section("quality"); const qualityComparison=buildQualityComparison(quality); const baselineQcUnclassified=quality.length ? 0 : baseline?.qcPassed ?? 0;
  const sales=section("sales"); const soldQty=sum(sales,"quantityKg"); const mobileRevenue=sales.reduce((total,entry)=>total+number(entry.amount),0);
  const inventory=section("inventory"); const latestInventory=inventory.reduce<Entry | null>((latest,entry)=>!latest || entry.id>latest.id ? entry : latest,null);
  const expenses=section("expenses"); const mobileExpenses=expenses.reduce((total,entry)=>total+number(entry.amount),0); const expenseGroups=groups(expenses,"category",true);

  return [
    create("raw-material",[
      {label:"Total raw material procured",value:`${formatNumber(totalRawMaterialKg,2)} kg`,note:"Batch input"},
      {label:"Recorded purchase value",value:formatMoney(rawValue),note:"Mobile entries"},
      {label:"Average rate",value:`${formatMoney(rawRate)}/kg`,note:"Recorded purchases"},
      {label:"RM lots",value:String(unique(raw,"lotCode")),note:"Unique lot codes"},
    ],groups(raw,"materialType"),Boolean(!raw.length && baseline?.rawMaterialKg)),
    create("pre-composting",[
      {label:"Pre-compost lots",value:formatNumber(baselineValue(pre,baseline?.preCompostLots)+preMobileLots),note:"Unique lots"},
      {label:"Logged weight",value:`${formatNumber(sum(pre,"totalWeightKg"),2)} kg`,note:"Lot setup entries"},
      {label:"Average temperature",value:`${formatNumber(average(pre,"avgTempC"),1)} °C`,note:"Recorded checks"},
      {label:"Average moisture",value:`${formatNumber(average(pre,"moisturePct"),1)}%`,note:"Recorded checks"},
    ],groups(pre,"eventType",true),Boolean(!pre.length && baseline?.preCompostLots)),
    create("bed-ops",[
      {label:"Production beds",value:formatNumber(baselineValue(bed,baseline?.totalBeds)+bedStatus.size),note:"Unique beds"},
      {label:"Active beds",value:formatNumber(baselineValue(bed,baseline?.activeBeds)+activeBedCount),note:"Latest saved status"},
      {label:"Substrate allocated",value:`${formatNumber(sum(bed,"dungWeightKg"),2)} kg`,note:"Bed filling entries"},
      {label:"Earthworms added",value:`${formatNumber(sum(bed,"wormsAddedKg"),2)} kg`,note:"Inoculation entries"},
    ],groups(bed,"currentStatus",true),Boolean(!bed.length && baseline?.totalBeds)),
    create("harvest",[
      {label:"Total raw material procured",value:`${formatNumber(totalRawMaterialKg,2)} kg`,note:"Comparison base · 100%"},
      {label:"Total production (net)",value:`${formatNumber(netHarvest,2)} kg`,note:`${formatNumber(rawMaterialComparison.productionPct,2)}% of raw material`},
      {label:"Packaged output",value:`${formatNumber(packagedHarvest,2)} kg`,note:`${formatNumber(rawMaterialComparison.packagedPct,2)}% of raw material`},
      {label:"Packaging coverage",value:`${formatNumber(rawMaterialComparison.packagingCoveragePct,2)}%`,note:"Packaged output ÷ total production"},
    ],groups(harvest,"packagingType"),Boolean(!harvest.length && (baseline?.rawYieldKg || baseline?.netYieldKg)),rawMaterialComparison),
    create("quality",[
      {label:"RM quality reports",value:formatNumber(qualityComparison.rm.entries),note:baselineQcUnclassified ? `${baselineQcUnclassified} baseline reports await RM/FG split` : "Raw material samples"},
      {label:"RM FCO passed",value:formatNumber(qualityComparison.rm.passed),note:"Raw material compliant"},
      {label:"FG quality reports",value:formatNumber(qualityComparison.fg.entries),note:baselineQcUnclassified ? `${baselineQcUnclassified} baseline reports await RM/FG split` : "Finished goods samples"},
      {label:"FG FCO passed",value:formatNumber(qualityComparison.fg.passed),note:"Finished goods compliant"},
    ],groups(quality,"productType"),Boolean(!quality.length && baseline?.qcPassed),undefined,qualityComparison),
    create("sales",[
      {label:"Quantity sold",value:`${formatNumber(soldQty,2)} kg`,note:"Recorded dispatches"},
      {label:"Sales revenue",value:formatMoney(baselineValue(sales,baseline?.revenue)+mobileRevenue),note:"Batch total"},
      {label:"Average recorded price",value:`${formatMoney(soldQty>0 ? mobileRevenue/soldQty : 0)}/kg`,note:"Mobile sales"},
      {label:"Customers",value:formatNumber(unique(sales,"customer")),note:"Unique buyers"},
    ],groups(sales,"customer"),Boolean(!sales.length && baseline?.revenue)),
    create("inventory",[
      {label:"Latest closing balance",value:`${formatNumber(number(latestInventory?.quantity),2)} kg`,note:"Latest inventory row"},
      {label:"Production added",value:`${formatNumber(sum(inventory,"productionKg"),2)} kg`,note:"All stock entries"},
      {label:"Sales outflow",value:`${formatNumber(sum(inventory,"salesOutflowKg"),2)} kg`,note:"All stock entries"},
      {label:"Damage / loss",value:`${formatNumber(sum(inventory,"damageLossKg"),2)} kg`,note:"All stock entries"},
    ],groups(inventory,"auditStatus",true),false),
    create("expenses",[
      {label:"Total batch expenses",value:formatMoney(baselineValue(expenses,baseline?.expenses)+mobileExpenses),note:"Workbook entries"},
      {label:"Recorded expenses",value:formatMoney(mobileExpenses),note:"Mobile entries"},
      {label:"Average expense",value:formatMoney(expenses.length ? mobileExpenses/expenses.length : 0),note:"Per entry"},
      {label:"Top category",value:expenseGroups[0]?.label ?? "—",note:expenseGroups[0] ? `${expenseGroups[0].value} entries` : "No entries"},
    ],expenseGroups,Boolean(!expenses.length && baseline?.expenses)),
  ];
}

function buildMetrics(batch:ProductionBatch, baseline:Baseline | undefined, entries:Entry[]): Metrics {
  const hasEntries=(moduleId:string)=>entries.some((entry)=>entry.module===moduleId);
  const result:Metrics = {
    productionBatchCode:batch.code, label:batch.name, status:batch.status,
    rawMaterialKg:hasEntries("raw-material") ? 0 : baseline?.rawMaterialKg ?? 0, preCompostLots:hasEntries("pre-composting") ? 0 : baseline?.preCompostLots ?? 0,
    totalBeds:hasEntries("bed-ops") ? 0 : baseline?.totalBeds ?? 0, activeBeds:hasEntries("bed-ops") ? 0 : baseline?.activeBeds ?? 0,
    rawYieldKg:hasEntries("harvest") ? 0 : baseline?.rawYieldKg ?? 0, netYieldKg:hasEntries("harvest") ? 0 : baseline?.netYieldKg ?? 0, packagedKg:hasEntries("harvest") ? 0 : baseline?.netYieldKg ?? 0,
    revenue:hasEntries("sales") ? 0 : baseline?.revenue ?? 0, expenses:hasEntries("expenses") ? 0 : baseline?.expenses ?? 0, qcPassed:hasEntries("quality") ? 0 : baseline?.qcPassed ?? 0,
    source:baseline?.source ?? "Mobile register entries", recoveryPct:0, packagingPct:0, margin:0, entries:entries.length,
  };
  const bedStatus = new Map<string,string>();
  const lots = new Set<string>();
  for (const entry of entries) {
    const detail = parseDetails(entry);
    if (entry.module === "raw-material") result.rawMaterialKg += number(detail.quantityKg);
    if (entry.module === "pre-composting" && detail.eventType === "Lot setup") lots.add(String(detail.lotCode ?? entry.referenceCode ?? entry.id));
    if (entry.module === "bed-ops" && entry.bedNumber && !bedStatus.has(entry.bedNumber)) bedStatus.set(entry.bedNumber, String(detail.currentStatus ?? entry.status ?? ""));
    if (entry.module === "harvest") { result.rawYieldKg += number(detail.rawQtyKg); result.netYieldKg += number(detail.netYieldKg); result.packagedKg += packagedQuantity(entry).quantity; }
    if (entry.module === "quality" && String(entry.status).toLowerCase().includes("pass")) result.qcPassed += 1;
    if (entry.module === "sales") result.revenue += number(entry.amount);
    if (entry.module === "expenses") result.expenses += number(entry.amount);
  }
  result.preCompostLots += lots.size;
  result.totalBeds += bedStatus.size;
  result.activeBeds += [...bedStatus.values()].filter((status) => ["filled","inoculated","active","ready for harvest"].includes(status.toLowerCase())).length;
  result.recoveryPct = result.rawMaterialKg > 0 ? (result.netYieldKg / result.rawMaterialKg) * 100 : 0;
  result.packagingPct = result.rawMaterialKg > 0 ? (result.packagedKg / result.rawMaterialKg) * 100 : 0;
  result.margin = result.revenue - result.expenses;
  return result;
}

export default function Home({ app, onSignOut }:{ app:AppStatus; onSignOut:()=>void }) {
  const [view, setView] = useState<View>("home");
  const [captureOpen, setCaptureOpen] = useState(false);
  const [batchOpen, setBatchOpen] = useState(false);
  const [selectedModule, setSelectedModule] = useState<string | null>(null);
  const [editingEntry, setEditingEntry] = useState<Entry | null>(null);
  const [selectedBatchCode, setSelectedBatchCode] = useState("BATCH-04");
  const [batches, setBatches] = useState<ProductionBatch[]>(fallbackBatches);
  const [baselines, setBaselines] = useState<Baseline[]>(fallbackBaselines);
  const [entries, setEntries] = useState<Entry[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [toast, setToast] = useState("");
  const [currentUser,setCurrentUser]=useState<AppUser | null>(null);
  const [authStatus,setAuthStatus]=useState<"checking"|"ready"|"signed-out"|"denied"|"setup">("checking");
  const [authMessage,setAuthMessage]=useState("");
  const prompt=usePrompt();
  const signOut=useCallback(()=>{ void endAccessSession().finally(onSignOut); },[onSignOut]);

  const loadSession=useCallback(async(quiet=false)=>{
    if (!quiet) setAuthStatus("checking");
    try {
      const response=await fetch("/api/session",{cache:"no-store",headers:{"x-vermitrack-access-session":accessSessionId()}});
      const payload=await response.json() as { user?:AppUser;error?:string;code?:string };
      if (!response.ok || !payload.user) {
        setAuthMessage(payload.error ?? "Could not verify this account.");
        setAuthStatus(payload.code==="UNAUTHENTICATED" ? "signed-out" : payload.code==="NOT_AUTHORISED" ? "denied" : "setup");
        return;
      }
      setCurrentUser(payload.user); setAuthStatus("ready"); setAuthMessage("");
    } catch { if (!quiet) { setAuthMessage("Could not connect to Microsoft 365."); setAuthStatus("setup"); } }
  },[]);

  const loadData = useCallback(async (quiet=false) => {
    if (!quiet) setLoading(true);
    try {
      const response = await fetch("/api/register", { cache:"no-store" });
      const payload = await response.json() as { batches?:ProductionBatch[]; baselines?:Baseline[]; entries?:Entry[]; error?:string };
      if (!response.ok) throw new Error(payload.error ?? "Could not sync the register.");
      if (payload.batches?.length) setBatches(payload.batches);
      if (payload.baselines) setBaselines(payload.baselines);
      setEntries(payload.entries ?? []);
      setError("");
    } catch (caught) { setError(caught instanceof Error ? caught.message : "Could not sync the register."); }
    finally { setLoading(false); }
  }, []);

  useEffect(() => {
    const initial=window.setTimeout(()=>{ const saved=window.localStorage.getItem("vermitrack-selected-batch"); if (saved) setSelectedBatchCode(saved); void loadSession(); },0);
    return ()=>window.clearTimeout(initial);
  },[loadSession]);

  useEffect(()=>{
    if (authStatus!=="ready") return;
    const pulse=()=>{ if (document.visibilityState==="visible") void fetch("/api/session",{cache:"no-store",headers:{"x-vermitrack-access-session":accessSessionId()}}).catch(()=>undefined); };
    const visibility=()=>{ if (document.visibilityState==="hidden") void endAccessSession(); else void loadSession(true); };
    const pageHide=()=>{ void endAccessSession(); };
    const heartbeat=window.setInterval(pulse,30000);
    document.addEventListener("visibilitychange",visibility); window.addEventListener("pagehide",pageHide);
    return ()=>{ window.clearInterval(heartbeat); document.removeEventListener("visibilitychange",visibility); window.removeEventListener("pagehide",pageHide); };
  },[authStatus,loadSession]);

  useEffect(()=>{
    if (authStatus!=="ready") return;
    const initial=window.setTimeout(()=>void loadData(),0);
    const interval=window.setInterval(()=>{ if (document.visibilityState==="visible") void loadData(true); },20000);
    return ()=>{ window.clearTimeout(initial); window.clearInterval(interval); };
  },[authStatus,loadData]);

  useEffect(() => { if (!toast) return; const timer = window.setTimeout(() => setToast(""), 3200); return () => window.clearTimeout(timer); }, [toast]);

  const selectedBatch = useMemo(() => batches.find((batch) => batch.code === selectedBatchCode) ?? batches[0] ?? fallbackBatches[1], [batches, selectedBatchCode]);
  const selectedEntries = useMemo(() => entries.filter((entry) => entry.productionBatchCode === selectedBatch.code), [entries, selectedBatch.code]);
  const metrics = useMemo(() => buildMetrics(selectedBatch, baselines.find((item) => item.productionBatchCode === selectedBatch.code), selectedEntries), [selectedBatch, baselines, selectedEntries]);
  const allBatchMetrics = useMemo(() => batches.map((batch) => buildMetrics(batch, baselines.find((item) => item.productionBatchCode === batch.code), entries.filter((entry) => entry.productionBatchCode === batch.code))), [batches, baselines, entries]);

  function changeBatch(code:string) { setSelectedBatchCode(code); window.localStorage.setItem("vermitrack-selected-batch", code); }
  function openCapture(moduleId?:string) { setEditingEntry(null); setSelectedModule(moduleId ?? null); setCaptureOpen(true); }
  function openEdit(entry:Entry) { setEditingEntry(entry); setSelectedModule(entry.module); setCaptureOpen(true); }
  async function deleteEntry(entry:Entry) {
    const reason=await prompt.ask({ title:"Request deletion", message:`Why should “${entry.title}” be deleted? The entry stays visible until ${DELETION_APPROVER_NAMES} approves the request. They are notified by Outlook email.`, placeholder:"Reason (at least 5 characters)", confirmLabel:"Send deletion request", danger:true, minLength:5 });
    if (reason===null) return;
    try { const response=await fetch(`/api/register?id=${entry.id}`,{method:"DELETE",headers:{"Content-Type":"application/json"},body:JSON.stringify({reason:reason.trim()})}); const payload=await response.json() as {requestId?:number;entry?:Entry;error?:string}; if (!response.ok || !payload.entry) throw new Error(payload.error ?? "Could not submit this deletion request."); setEntries((current)=>current.map((item)=>item.id===entry.id ? payload.entry! : item)); setToast(`Deletion request sent to ${DELETION_APPROVER_NAMES.replace(" or "," and ")}`); } catch (caught) { setToast(caught instanceof Error ? caught.message : "Could not submit this deletion request."); }
  }
  function navigate(next:View) { setView(next); window.scrollTo({ top:0, behavior:"smooth" }); }

  if (authStatus!=="ready" || !currentUser) return <AccessScreen status={authStatus} message={authMessage} account={app.account} onRetry={()=>void loadSession()} onSignOut={signOut} />;

  return <main className="app-shell">
    <Header user={currentUser} app={app} loading={loading} error={error} onRefresh={() => void loadData()} onConnection={()=>navigate("more")} onAdmin={()=>navigate("admin")} />
    {app.mode==="demo" && <div className="demo-banner">Demo mode · records are saved on this computer only and are not shared. Choose More → Leave demo mode to connect your team’s Microsoft 365.</div>}
    <section className="content">
      <BatchSwitcher batches={batches} selectedCode={selectedBatch.code} onChange={changeBatch} onAdd={() => setBatchOpen(true)} />
      {view === "home" && <HomeView firstName={currentUser.fullName.trim().split(/\s+/)[0] || "User"} batch={selectedBatch} metrics={metrics} entries={selectedEntries} onCapture={openCapture} onNavigate={navigate} />}
      {view === "records" && <RecordsView batch={selectedBatch} entries={selectedEntries} onCapture={openCapture} onEdit={openEdit} onDelete={(entry)=>void deleteEntry(entry)} />}
      {view === "insights" && <ReportsView batch={selectedBatch} baseline={baselines.find((item) => item.productionBatchCode === selectedBatch.code)} metrics={metrics} entries={selectedEntries} allBatchMetrics={allBatchMetrics} currentUser={currentUser} onEdit={openEdit} onDelete={(entry)=>void deleteEntry(entry)} onToast={setToast} />}
      {view === "admin" && <AdminView currentUser={currentUser} app={app} ask={prompt.ask} onSignOut={signOut} onToast={setToast} onDataChanged={()=>void loadData(true)} />}
      {view === "more" && <MoreView batch={selectedBatch} baseline={baselines.find((item) => item.productionBatchCode === selectedBatch.code)} batches={batches} app={app} currentUser={currentUser} onSignOut={signOut} onToast={setToast} onCapture={openCapture} onAddBatch={() => setBatchOpen(true)} onSelectBatch={changeBatch} onAdmin={()=>navigate("admin")} />}
    </section>
    <BottomNav view={view} onNavigate={navigate} onCapture={() => openCapture()} />
    {captureOpen && <CaptureSheet batches={batches} entries={entries} defaultBatchCode={selectedBatch.code} selectedId={selectedModule} existingEntry={editingEntry} onSelect={setSelectedModule} onClose={() => { setCaptureOpen(false); setSelectedModule(null); setEditingEntry(null); }} onSaved={(entry) => { setEntries((current) => current.some((item)=>item.id===entry.id) ? current.map((item)=>item.id===entry.id ? entry : item) : [entry, ...current]); changeBatch(entry.productionBatchCode); setCaptureOpen(false); setSelectedModule(null); setToast(editingEntry ? "Entry updated" : `Entry saved in ${batches.find((batch) => batch.code === entry.productionBatchCode)?.name ?? entry.productionBatchCode}`); setEditingEntry(null); }} />}
    {batchOpen && <BatchSheet batches={batches} onClose={() => setBatchOpen(false)} onCreated={(batch) => { setBatches((current) => [batch, ...current]); changeBatch(batch.code); setBatchOpen(false); setToast(`${batch.name} created and selected`); }} />}
    {prompt.element}
    {toast && <div className="toast" role="status"><span>✓</span>{toast}</div>}
  </main>;
}

function Header({ user, app, loading, error, onRefresh, onConnection, onAdmin }:{ user:AppUser; app:AppStatus; loading:boolean; error:string; onRefresh:()=>void; onConnection:()=>void; onAdmin:()=>void }) {
  const initials=user.fullName.split(/\s+/).map((part)=>part[0]).join("").slice(0,2).toUpperCase();
  const demo=app.mode==="demo";
  return <header className="topbar"><div className="brand-mark navjyoti-mark" role="img" aria-label="Navjyoti logo" /><div className="brand-copy"><strong>Navjyoti</strong><span>VermiTrack · Akola unit</span></div><button className={`desktop-install-button ${demo ? "" : "installed"}`} onClick={onConnection} title={demo ? "Demo data on this computer" : `Shared on ${app.siteUrl ?? "Microsoft 365"} · signed in as ${app.account ?? user.email}`}><span>{demo ? "!" : "☁"}</span><b>{demo ? "Demo · this PC only" : `Shared · ${app.siteName ?? "Microsoft 365"}`}</b></button><button className={`sync-pill ${error ? "offline" : ""}`} onClick={onRefresh} aria-label="Refresh records" title={error || "Refreshes automatically every 20 seconds"}><i />{loading ? "Syncing" : error ? "Retry" : "Live"}</button><button className="avatar" onClick={onAdmin} aria-label={`Open administration for ${user.fullName}`}>{initials}</button></header>;
}

function AccessScreen({status,message,account,onRetry,onSignOut}:{status:"checking"|"ready"|"signed-out"|"denied"|"setup";message:string;account:string | null;onRetry:()=>void;onSignOut:()=>void}) {
  return <main className="access-screen"><section className="access-card"><div className="access-logo" role="img" aria-label="Navjyoti logo" /><p className="eyebrow">SECURE PRODUCTION REGISTER</p><h1>Navjyoti VermiTrack</h1>{status==="checking" ? <><div className="access-spinner" /><p>Checking your approved account…</p></> : <><p>{message || "Sign in with an approved Navjyoti Microsoft 365 account to continue."}</p><button className="access-primary" onClick={onRetry}>Check access again</button><button className="access-secondary access-secondary-button" onClick={onSignOut}>Sign out and use another account</button></>}<footer><span>Approved Navjyoti users only</span><b>{account ? `Signed in to Microsoft 365 as ${account}` : "Microsoft 365 sign-in"}</b></footer></section></main>;
}


function BatchSwitcher({ batches, selectedCode, onChange, onAdd }:{ batches:ProductionBatch[]; selectedCode:string; onChange:(code:string)=>void; onAdd:()=>void }) {
  const selected = batches.find((batch) => batch.code === selectedCode);
  return <section className="batch-switcher"><div><span>PRODUCTION BATCH</span><select aria-label="Select production batch" value={selectedCode} onChange={(event) => onChange(event.target.value)}>{batches.map((batch) => <option key={batch.code} value={batch.code}>{batch.name} · {batch.code}</option>)}</select></div><em className={`batch-status ${selected?.status.toLowerCase().replaceAll(" ", "-")}`}>{selected?.status}</em><button onClick={onAdd}>＋ New batch</button></section>;
}

function HomeView({ firstName, batch, metrics, entries, onCapture, onNavigate }:{ firstName:string; batch:ProductionBatch; metrics:Metrics; entries:Entry[]; onCapture:(id?:string)=>void; onNavigate:(view:View)=>void }) {
  return <>
    <div className="welcome-row"><div><p className="eyebrow">TODAY · {batch.name.toUpperCase()}</p><h1>{greetingForNow()}, {firstName}</h1><p>Here is the latest production picture for the selected batch.</p></div></div>
    <section className="hero-card"><div className="hero-top"><span className="batch-label">{batch.name.toUpperCase()} · {batch.code}</span><span className="status-badge">{batch.status}</span></div><div className="yield-line"><div><strong>{formatNumber(metrics.netYieldKg)}</strong><span>kg total production</span></div><div><strong>{metrics.recoveryPct.toFixed(1)}%</strong><span>of total raw material</span></div></div><div className="progress"><b style={{ width:`${Math.min(metrics.recoveryPct,100)}%` }} /></div><div className="hero-foot"><span>{formatNumber(metrics.rawMaterialKg)} kg raw material procured</span><span>{formatNumber(metrics.packagedKg)} kg packaged</span></div></section>
    <div className="section-heading"><h2>Live overview</h2><button onClick={() => onNavigate("insights")}>View reports →</button></div>
    <section className="metric-grid"><Metric icon="RM" tone="green" value={`${formatNumber(metrics.rawMaterialKg / 1000,2)} t`} label="Raw material" /><Metric icon="BD" tone="plum" value={formatNumber(metrics.totalBeds)} label={`${metrics.activeBeds} active beds`} /><Metric icon="₹" tone="amber" value={formatCompactMoney(metrics.revenue).replace("₹","")} label="Revenue" /><Metric icon="↗" tone="red" value={formatCompactMoney(metrics.expenses).replace("₹","")} label="Expenses" /></section>
    <section className="quick-strip"><button onClick={() => onCapture("raw-material")}><span>RM</span><b>Material receipt</b><small>Add to {batch.name}</small></button><button onClick={() => onCapture("bed-ops")}><span>BD</span><b>Bed check</b><small>Log field work</small></button><button onClick={() => onCapture("harvest")}><span>HV</span><b>Harvest</b><small>Record yield</small></button></section>
    <div className="section-heading"><h2>Recent entries</h2><button onClick={() => onNavigate("records")}>See all</button></div><ActivityList batch={batch} entries={entries.slice(0,4)} />
  </>;
}

function Metric({ icon, tone, value, label }:{ icon:string; tone:string; value:string; label:string }) { return <article><span className={`metric-icon ${tone}`}>{icon}</span><div><strong>{value}</strong><small>{label}</small></div></article>; }

function ActivityList({ batch, entries }:{ batch:ProductionBatch; entries:Entry[] }) {
  if (!entries.length) return <section className="activity-list"><article><span className="activity-icon">BT</span><div><b>No mobile entries for {batch.name}</b><small>Choose Capture to add historical or current records</small></div><time>Ready</time></article></section>;
  return <section className="activity-list">{entries.map((entry) => { const moduleDef=moduleFor(entry.module); const identity=entrySystemIdentity(entry); return <article key={entry.id}><span className={`activity-icon ${moduleDef.tone}`}>{moduleDef.short}</span><div><b>{entry.title}</b><small>Sr. {identity.serial} · {identity.recordCode}{entry.status ? ` · ${entry.status}` : ""}</small></div><time>{displayTime(entry.createdAt)}</time></article>; })}</section>;
}

function RecordsView({ batch, entries, onCapture, onEdit, onDelete }:{ batch:ProductionBatch; entries:Entry[]; onCapture:(id?:string)=>void; onEdit:(entry:Entry)=>void; onDelete:(entry:Entry)=>void }) {
  const [query,setQuery]=useState(""); const [filter,setFilter]=useState("all");
  const filtered=useMemo(() => entries.filter((entry) => { const moduleMatch=filter === "all" || entry.module === filter; const identity=entrySystemIdentity(entry); const text=`${identity.serial} ${identity.recordCode} ${entry.title} ${entry.referenceCode} ${entry.batchCode} ${entry.bedNumber} ${entry.status}`.toLowerCase(); return moduleMatch && text.includes(query.toLowerCase()); }), [entries,filter,query]);
  function exportCsv() { const quote=(value:unknown)=>`"${String(value ?? "").replaceAll('"','""')}"`; const rows=[["Sr No","Auto Register Code","Production Batch","Register","Entry Date","Reference","Finished Batch","Bed","Title","Quantity","Amount","Status","Entered By","Created At","Details"],...filtered.map((entry)=>{ const identity=entrySystemIdentity(entry); return [identity.serial,identity.recordCode,entry.productionBatchCode,moduleFor(entry.module).title,entry.entryDate,entry.referenceCode,entry.batchCode,entry.bedNumber,entry.title,entry.quantity,entry.amount,entry.status,entry.createdBy,new Date(entry.createdAt).toISOString(),entry.details]; })]; const blob=new Blob([rows.map((row)=>row.map(quote).join(",")).join("\n")],{type:"text/csv;charset=utf-8"}); const url=URL.createObjectURL(blob); const link=document.createElement("a"); link.href=url; link.download=`VermiTrack-${batch.code}-${todayIso()}.csv`; link.click(); URL.revokeObjectURL(url); }
  return <section className="page-view"><div className="page-head"><div><p className="eyebrow">{batch.code} · TRACEABILITY</p><h1>{batch.name} records</h1><p>Every entry is stored with production batch, time, operator and linked codes.</p></div><button className="primary-small" onClick={() => onCapture()}>＋ Add entry</button></div><div className="record-tools"><label className="search-box"><span>⌕</span><input value={query} onChange={(event)=>setQuery(event.target.value)} placeholder="Search lot, finished batch, bed…" /></label><select value={filter} onChange={(event)=>setFilter(event.target.value)}><option value="all">All registers</option>{registerModules.map((moduleDef)=><option key={moduleDef.id} value={moduleDef.id}>{moduleDef.title}</option>)}</select><button onClick={exportCsv} disabled={!filtered.length}>Export CSV</button></div><div className="record-count"><b>{filtered.length}</b> entries <span>•</span> <i>{batch.name} selected</i></div>{filtered.length ? <div className="records-list">{filtered.map((entry)=><RecordCard key={entry.id} entry={entry} batch={batch} onEdit={onEdit} onDelete={onDelete} />)}</div> : <EmptyRecords batch={batch} onCapture={onCapture} />}</section>;
}

function RecordCard({ entry, batch, onEdit, onDelete }:{ entry:Entry; batch:ProductionBatch; onEdit:(entry:Entry)=>void; onDelete:(entry:Entry)=>void }) { const moduleDef=moduleFor(entry.module); const details=parseDetails(entry); const identity=entrySystemIdentity(entry); const pending=entry.deletionStatus==="pending"; const highlights=Object.entries(details).filter(([key,value])=>!key.startsWith("_") && value !== "" && value !== null && value !== undefined && value !== entry.productionBatchCode).slice(0,8); return <details className={`record-card ${pending ? "pending-deletion" : ""}`}><summary><span className={`register-badge ${moduleDef.tone}`}>{moduleDef.short}</span><div><b>{entry.title}</b><small>Sr. {identity.serial} · {identity.recordCode} · {batch.name} · {displayDate(entry.entryDate)}</small></div><div className="record-value">{pending ? "Deletion pending" : entry.amount ? formatMoney(entry.amount) : entry.quantity ? `${formatNumber(entry.quantity,2)} ${entry.module === "expenses" ? "" : "kg"}` : entry.status || "Saved"}<i>⌄</i></div></summary><div className="record-detail"><div className="detail-tags"><span>Sr. No.: {identity.serial}</span><span>Register code: {identity.recordCode}</span><span>Production: {entry.productionBatchCode}</span><span>Punched: {displayTimestamp(entry.createdAt)}</span><span>By: {entry.createdBy}</span>{pending && <span className="pending-tag">Deletion approval pending</span>}{entry.referenceCode && <span>Source ref: {entry.referenceCode}</span>}{entry.batchCode && <span>Finished: {entry.batchCode}</span>}{entry.bedNumber && <span>Bed: {entry.bedNumber}</span>}{entry.status && <span>Status: {entry.status}</span>}</div><dl>{highlights.map(([key,value])=><div key={key}><dt>{entryFieldLabel(entry.module,key)}</dt><dd>{String(value)}</dd></div>)}</dl><div className="entry-actions"><button disabled={pending} onClick={()=>onEdit(entry)}>{pending ? "Locked pending approval" : "Edit entry"}</button><button className="danger" disabled={pending} onClick={()=>onDelete(entry)}>{pending ? "Request submitted" : "Request deletion"}</button></div></div></details>; }

function EmptyRecords({ batch, onCapture }:{ batch:ProductionBatch; onCapture:(id?:string)=>void }) { return <div className="empty-state"><span>▤</span><h2>No entries in {batch.name}</h2><p>Add old records or begin capturing the new batch.</p><button onClick={()=>onCapture()}>Capture an entry</button></div>; }

function ReportsView({ batch, baseline, metrics, entries, allBatchMetrics, currentUser, onEdit, onDelete, onToast }:{ batch:ProductionBatch; baseline:Baseline | undefined; metrics:Metrics; entries:Entry[]; allBatchMetrics:Metrics[]; currentUser:AppUser; onEdit:(entry:Entry)=>void; onDelete:(entry:Entry)=>void; onToast:(message:string)=>void }) {
  const [emailOpen,setEmailOpen]=useState(false);
  const reports=useMemo(()=>buildSegmentReports(baseline,entries),[baseline,entries]); const populated=reports.filter((report)=>report.hasData).length;
  const [selectedReportId,setSelectedReportId]=useState("raw-material"); const [reportMode,setReportMode]=useState<"summary" | "charts" | "entries">("summary");
  const selectedReport=reports.find((report)=>report.id===selectedReportId) ?? reports[0];
  const costPerKg=metrics.netYieldKg > 0 ? metrics.expenses / metrics.netYieldKg : 0; const realization=metrics.netYieldKg > 0 ? metrics.revenue / metrics.netYieldKg : 0;
  function chooseReport(id:string) { setSelectedReportId(id); setReportMode("summary"); }
  function exportReport() {
    const quote=(value:unknown)=>`"${String(value ?? "").replaceAll('"','""')}"`;
    const rows:Array<Array<string | number>>=[["Row Type","Production Batch","Section","Entry ID","Entry Date","Punched At","Entry Maker","Entry Title","Field / Metric","Captured / Consolidated Value","Imported Baseline Included"]];
    for (const report of reports) {
      for (const metric of report.metrics) rows.push(["Consolidated metric",batch.code,report.title,"","",report.lastEntryDate ?? "","","",metric.label,metric.value,report.includesBaseline ? "Yes" : "No"]);
      if (report.rawMaterialComparison) {
        const comparison=report.rawMaterialComparison;
        const values:Array<[string,string]>=[
          ["Total production conversion",`${formatNumber(comparison.productionPct,2)}% of total raw material procured`],
          ["Packaged output conversion",`${formatNumber(comparison.packagedPct,2)}% of total raw material procured`],
          ["Packaging coverage",`${formatNumber(comparison.packagingCoveragePct,2)}% of total production`],
          ["Raw material awaiting production",`${formatNumber(comparison.awaitingProductionKg,2)} kg`],
          ["Production awaiting packaging",`${formatNumber(comparison.awaitingPackagingKg,2)} kg`],
        ];
        for (const [label,value] of values) rows.push(["Raw material comparison",batch.code,report.title,"","",report.lastEntryDate ?? "","","",label,value,report.includesBaseline ? "Yes" : "No"]);
      }
      if (report.qualityComparison) {
        for (const profile of [report.qualityComparison.rm,report.qualityComparison.fg]) {
          rows.push(["Quality indicator",batch.code,report.title,"","",report.lastEntryDate ?? "","","",`${profile.code} quality reports`,profile.entries,"No"]);
          rows.push(["Quality indicator",batch.code,report.title,"","",report.lastEntryDate ?? "","","",`${profile.code} FCO passed`,profile.passed,"No"]);
        }
        for (const field of report.qualityComparison.rows) {
          rows.push(["RM quality average",batch.code,report.title,"","",report.lastEntryDate ?? "","","",field.label,field.rmAverage===null ? "No RM value" : formatConsolidatedNumber(field,field.rmAverage),"No"]);
          rows.push(["FG quality average",batch.code,report.title,"","",report.lastEntryDate ?? "","","",field.label,field.fgAverage===null ? "No FG value" : formatConsolidatedNumber(field,field.fgAverage),"No"]);
        }
      }
      for (const field of report.consolidatedFields) rows.push(["Consolidated field",batch.code,report.title,"","",report.lastEntryDate ?? "","","",field.label,`${field.primary} | ${field.detail}${field.breakdown.length ? ` | All values: ${field.breakdown.map((item)=>`${item.label} (${item.value})`).join("; ")}` : ""}`,report.includesBaseline ? "Yes" : "No"]);
      for (const entry of report.entries) { const identity=entrySystemIdentity(entry); rows.push(["System identity",batch.code,report.title,identity.serial,entry.entryDate,new Date(entry.createdAt).toISOString(),entry.createdBy || "Field operator",entry.title,"Auto register code",identity.recordCode,"No"]); for (const field of entryCapturedFields(entry)) rows.push(["Captured entry",batch.code,report.title,identity.serial,entry.entryDate,new Date(entry.createdAt).toISOString(),entry.createdBy || "Field operator",entry.title,field.label,field.value,"No"]); }
    }
    const blob=new Blob([rows.map((row)=>row.map(quote).join(",")).join("\n")],{type:"text/csv;charset=utf-8"}); const url=URL.createObjectURL(blob); const link=document.createElement("a"); link.href=url; link.download=`VermiTrack-${batch.code}-section-report-${todayIso()}.csv`; link.click(); URL.revokeObjectURL(url);
  }
  return <section className="page-view insights-view"><div className="page-head reports-head"><div><p className="eyebrow">{batch.code} · CONSOLIDATED REPORTS</p><h1>{batch.name} reports</h1><p>Batch-wise results from every register, updated whenever a new entry is saved.</p></div><div className="report-head-actions"><button className="primary-small secondary" onClick={()=>setEmailOpen(true)}>✉ Email report</button><button className="primary-small" onClick={exportReport}>↓ Export report</button></div></div>
    <section className="report-overview"><article><span>Sections with data</span><b>{populated}<i>/8</i></b><small>Imported baseline and mobile entries</small></article><article><span>Saved entries</span><b>{entries.length}</b><small>Across all registers</small></article><article><span>Batch status</span><b className="text-value">{batch.status}</b><small>{batch.financialYear ? `FY ${batch.financialYear}` : "Financial year not set"}</small></article></section>
    <section className="insight-hero report-hero"><div><span>Net operating result</span><strong className={metrics.margin < 0 ? "negative" : "positive"}>{formatMoney(metrics.margin)}</strong><small>Revenue {formatMoney(metrics.revenue)} less expenses {formatMoney(metrics.expenses)}</small></div><div className="ring" style={{"--progress":`${Math.min(metrics.recoveryPct,100)*3.6}deg`} as React.CSSProperties}><b>{metrics.recoveryPct.toFixed(1)}%</b><span>production ÷ raw material</span></div></section><section className="insight-grid"><article><span>Production cost</span><b>{formatMoney(costPerKg)}/kg</b><small>selected batch</small></article><article><span>Revenue realization</span><b>{formatMoney(realization)}/kg</b><small>selected batch</small></article><article><span>QC pass records</span><b>{metrics.qcPassed}</b><small>FCO-compliant reports</small></article><article><span>Active production beds</span><b>{metrics.activeBeds}</b><small>of {metrics.totalBeds} batch beds</small></article></section>
    <div className="section-heading"><h2>Select a segment</h2><span className="count-pill neutral">8 segments</span></div>
    <section className="segment-selector" aria-label="Report segments">{reports.map((report)=><button type="button" key={report.id} className={selectedReport.id===report.id ? "selected" : ""} aria-pressed={selectedReport.id===report.id} onClick={()=>chooseReport(report.id)}><span className={`register-badge ${report.tone}`}>{report.short}</span><div><b>{report.title}</b><small>{report.entryCount} {report.entryCount===1 ? "entry" : "entries"}{report.includesBaseline ? " · baseline" : ""}</small></div><i>›</i></button>)}</section>
    <section className="focused-report" aria-live="polite">
      <header className="focused-report-head"><span className={`register-badge ${selectedReport.tone}`}>{selectedReport.short}</span><div><p>SELECTED SEGMENT</p><h2>{selectedReport.title}</h2><span>{selectedReport.description}</span></div><em>{selectedReport.entryCount} {selectedReport.entryCount===1 ? "entry" : "entries"}</em></header>
      <div className="report-mode-tabs" role="tablist" aria-label={`${selectedReport.title} report view`}><button type="button" role="tab" aria-selected={reportMode==="summary"} className={reportMode==="summary" ? "active" : ""} onClick={()=>setReportMode("summary")}>≡ Consolidated <span>{selectedReport.consolidatedFields.length}</span></button><button type="button" role="tab" aria-selected={reportMode==="charts"} className={reportMode==="charts" ? "active" : ""} onClick={()=>setReportMode("charts")}>▦ Charts</button><button type="button" role="tab" aria-selected={reportMode==="entries"} className={reportMode==="entries" ? "active" : ""} onClick={()=>setReportMode("entries")}>▤ Entries <span>{selectedReport.entryCount}</span></button></div>
      <div className="visible-values-heading"><div><strong>Consolidated values</strong><span>Always visible for {batch.name}</span></div>{selectedReport.includesBaseline && <em>Imported baseline included</em>}</div>
      <div className="segment-metrics focused-metrics">{selectedReport.metrics.map((metric)=><article key={metric.label}><span>{metric.label}</span><b>{metric.value}</b><small>{metric.note}</small></article>)}</div>
      {reportMode==="summary" && <ConsolidatedFieldTable report={selectedReport} batch={batch} />}
      {reportMode==="charts" && <SectionChart report={selectedReport} />}
      {reportMode==="entries" && <section className="focused-entries" role="tabpanel"><div className="section-entry-heading"><div><strong>All saved entries</strong><span>{selectedReport.entryCount} in {batch.name}</span></div></div>{selectedReport.entries.length ? <div className="report-entry-list">{selectedReport.entries.map((entry)=><ReportEntryItem key={entry.id} entry={entry} onEdit={onEdit} onDelete={onDelete} />)}</div> : <div className="report-entry-empty"><span>▤</span><p>{selectedReport.includesBaseline ? "The imported baseline is included in the values above. No individual mobile entries have been punched in this segment yet." : "No entries have been punched in this segment for the selected batch yet."}</p></div>}</section>}
      <footer><span>{selectedReport.includesBaseline ? `Includes imported values from ${baseline?.source ?? "the batch baseline"}.` : "Calculated from entries saved in this batch."}</span><time>{selectedReport.lastEntryDate ? `Latest activity ${displayDate(selectedReport.lastEntryDate)}` : "No mobile entry yet"}</time></footer>
    </section>
    <div className="section-heading"><h2>Batch comparison</h2><span className="count-pill neutral">{allBatchMetrics.length} batches</span></div><section className="batch-comparison">{allBatchMetrics.map((item)=><article className={item.productionBatchCode === batch.code ? "selected" : ""} key={item.productionBatchCode}><div><b>{item.label}</b><span>{item.productionBatchCode} · {item.status}</span></div><dl><div><dt>Raw material procured</dt><dd>{formatNumber(item.rawMaterialKg)} kg</dd></div><div><dt>Total production</dt><dd>{formatNumber(item.netYieldKg)} kg</dd></div><div><dt>Packaged</dt><dd>{formatNumber(item.packagedKg)} kg</dd></div><div><dt>Production conversion</dt><dd>{formatNumber(item.recoveryPct,2)}%</dd></div><div><dt>Revenue</dt><dd>{formatCompactMoney(item.revenue)}</dd></div><div><dt>Margin</dt><dd className={item.margin < 0 ? "negative" : ""}>{formatCompactMoney(item.margin)}</dd></div></dl></article>)}</section>
    <div className="section-heading"><h2>Management signals</h2></div><section className="signal-list"><article className={metrics.margin < 0 ? "warning" : "good"}><span>{metrics.margin < 0 ? "!" : "✓"}</span><div><b>{metrics.margin < 0 ? "Cost exceeds recorded revenue" : "Positive recorded margin"}</b><p>{metrics.margin < 0 ? `${formatMoney(Math.abs(metrics.margin))} remains unrecovered in ${batch.name}. Review stock and unsold production before concluding profitability.` : `Recorded sales cover the operating expenses for ${batch.name}.`}</p></div></article><article className="good"><span>✓</span><div><b>Batch-wise reporting active</b><p>Every section report uses the selected production batch while finished-batch, lot and bed codes remain traceable inside the entries.</p></div></article></section>{emailOpen && <EmailReportSheet batch={batch} metrics={metrics} reports={reports} currentUser={currentUser} onClose={()=>setEmailOpen(false)} onSent={(count)=>{ setEmailOpen(false); onToast(`Report emailed to ${count} ${count===1 ? "person" : "people"} from your Outlook`); }} />}</section>;
}

function ConsolidatedFieldTable({ report, batch }:{ report:SegmentReport; batch:ProductionBatch }) {
  const groups=[...new Set(report.consolidatedFields.map((field)=>field.group))];
  const identities=report.entries.map(entrySystemIdentity); const serials=identities.map((item)=>item.serial).sort((a,b)=>a-b);
  return <section className="consolidated-field-report" role="tabpanel"><header><div><strong>Complete consolidated form</strong><span>Every field configured in {report.title} is included</span></div><em>{report.consolidatedFields.length} fields</em></header><div className="auto-identity-summary"><article><span>Production batch</span><b>{batch.code}</b><small>Applied to every entry</small></article><article><span>Serial number range</span><b>{serials.length ? serials.length===1 ? `Sr. ${serials[0]}` : `Sr. ${serials[0]} — ${serials.at(-1)}` : "Auto on first save"}</b><small>System generated</small></article><article><span>Register code format</span><b>{batch.code}-{report.short}-0001</b><small>Unique code on save</small></article></div>{report.rawMaterialComparison && <RawMaterialComparisonSummary comparison={report.rawMaterialComparison} />}{report.qualityComparison && <QualitySplitSummary comparison={report.qualityComparison} />}{groups.map((group)=><section className="consolidated-group" key={group}><header><b>{group}</b><span>{report.consolidatedFields.filter((field)=>field.group===group).length} fields</span></header><div className="consolidated-table-wrap"><table><thead><tr><th>Field</th><th>Consolidated value</th><th>Coverage</th></tr></thead><tbody>{report.consolidatedFields.filter((field)=>field.group===group).map((field)=><tr key={field.key}><td><b>{field.label}</b><small>{field.type === "number" ? shouldAverageField(field) ? "Average" : "Total" : field.type === "date" ? "Date range" : "Recorded values"}</small></td><td><strong>{field.primary}</strong><small>{field.detail}</small>{field.breakdown.length>1 && <details className="all-values"><summary>Show all {field.breakdown.length} values</summary><div>{field.breakdown.map((item)=><span key={item.label}>{item.label}<b>{item.value}</b></span>)}</div></details>}</td><td><b>{field.count}/{report.entryCount}</b><small>{field.missing ? `${field.missing} blank` : report.entryCount ? "Complete" : "No entries"}</small></td></tr>)}</tbody></table></div></section>)}</section>;
}

function RawMaterialComparisonSummary({ comparison }:{ comparison:RawMaterialComparison }) {
  const stages=[
    { label:"Total raw material procured",value:comparison.rawMaterialKg,pct:100,note:"Comparison base" },
    { label:"Total production (net)",value:comparison.productionKg,pct:comparison.productionPct,note:"Production from raw material" },
    { label:"Packaged output",value:comparison.packagedKg,pct:comparison.packagedPct,note:"Packaged from raw material" },
  ];
  return <section className="raw-material-conversion"><header><div><b>Raw material-to-packaged output</b><span>Total production and packaging are compared with total raw material procured.</span></div><em>{formatNumber(comparison.packagingCoveragePct,2)}% of production packaged</em></header><div className="conversion-stages">{stages.map((stage,index)=><article key={stage.label}><div><span>0{index+1}</span><small>{stage.note}</small></div><b>{formatNumber(stage.value,2)} kg</b><strong>{formatNumber(stage.pct,2)}%</strong><div className="conversion-track"><i style={{width:`${Math.min(100,Math.max(0,stage.pct))}%`}} /></div><p>{stage.label}</p></article>)}</div><div className="conversion-balances"><article><span>Awaiting production</span><b>{formatNumber(comparison.awaitingProductionKg,2)} kg</b><small>Raw material less total production</small></article><article><span>Awaiting packaging</span><b>{formatNumber(comparison.awaitingPackagingKg,2)} kg</b><small>Total production less packaged output</small></article></div><p className="conversion-note">Balances are operational differences for the ongoing batch and are not automatically treated as production loss.{comparison.packagedDerived ? " Where a separate packaged quantity was not recorded, the app derives it from pack size and packet count, or from bulk production." : ""}</p></section>;
}

function QualitySplitSummary({ comparison }:{ comparison:QualityComparison }) {
  const profiles=[comparison.rm,comparison.fg];
  return <section className="quality-split-summary"><header><div><b>RM and FG quality indicators</b><span>Raw material and finished goods results are calculated independently.</span></div><em>{comparison.rm.entries+comparison.fg.entries} classified reports</em></header><div className="quality-profile-cards">{profiles.map((profile)=><article className={profile.code.toLowerCase()} key={profile.code}><div><span>{profile.code}</span><div><b>{profile.label}</b><small>Separate quality profile</small></div></div><dl><div><dt>Reports</dt><dd>{profile.entries}</dd></div><div><dt>FCO passed</dt><dd>{profile.passed}</dd></div><div><dt>Failed</dt><dd>{profile.failed}</dd></div><div><dt>Pending / conditional</dt><dd>{profile.pending}</dd></div></dl></article>)}</div><div className="quality-comparison-table"><table><thead><tr><th>Quality parameter</th><th>RM average</th><th>FG average</th></tr></thead><tbody>{comparison.rows.map((row)=><tr key={row.key}><td><b>{row.label}</b><small>{row.group}</small></td><td><strong>{row.rmAverage===null ? "No RM value" : formatConsolidatedNumber(row,row.rmAverage)}</strong><small>{row.rmCount} {row.rmCount===1 ? "reading" : "readings"}</small></td><td><strong>{row.fgAverage===null ? "No FG value" : formatConsolidatedNumber(row,row.fgAverage)}</strong><small>{row.fgCount} {row.fgCount===1 ? "reading" : "readings"}</small></td></tr>)}</tbody></table></div></section>;
}

function chartNumber(value:string) { const cleaned=value.replaceAll(",","").replace(/[₹%A-Za-z/°\s]/g,""); const match=cleaned.match(/-?\d+(?:\.\d+)?/); return match ? Math.abs(Number(match[0])) : 0; }

function ChartBars({ data }:{ data:Array<{label:string;value:number;display:string}> }) {
  const max=Math.max(1,...data.map((item)=>Math.abs(item.value)));
  return data.length ? <div className="chart-bars">{data.map((item)=><div className="chart-bar" key={item.label}><label><span>{item.label}</span><b>{item.display}</b></label><div><i style={{width:`${item.value===0 ? 2 : Math.max(4,(Math.abs(item.value)/max)*100)}%`}} /></div></div>)}</div> : <div className="chart-empty">No values are available for this chart.</div>;
}

function QualitySplitCharts({ report, comparison }:{ report:SegmentReport; comparison:QualityComparison }) {
  const groups=new Map<string,{ group:string; unit:string; data:Array<{label:string;value:number;display:string}> }>();
  for (const row of comparison.rows) {
    if (row.rmAverage===null && row.fgAverage===null) continue;
    const unit=row.unit || "Number"; const key=`${row.group} · ${unit}`; const current=groups.get(key) ?? { group:row.group,unit,data:[] };
    if (row.rmAverage!==null) current.data.push({label:`RM · ${row.label}`,value:Math.abs(row.rmAverage),display:formatConsolidatedNumber(row,row.rmAverage)});
    if (row.fgAverage!==null) current.data.push({label:`FG · ${row.label}`,value:Math.abs(row.fgAverage),display:formatConsolidatedNumber(row,row.fgAverage)});
    groups.set(key,current);
  }
  const statusData=[
    {label:"RM quality reports",value:comparison.rm.entries,display:String(comparison.rm.entries)},
    {label:"RM FCO passed",value:comparison.rm.passed,display:String(comparison.rm.passed)},
    {label:"FG quality reports",value:comparison.fg.entries,display:String(comparison.fg.entries)},
    {label:"FG FCO passed",value:comparison.fg.passed,display:String(comparison.fg.passed)},
  ];
  return <section className="chart-suite" role="tabpanel" aria-label={`${report.title} RM and FG charts`}><header><div><strong>RM and FG quality charts</strong><span>Each chart keeps raw material and finished goods separate on compatible scales.</span></div><em>{1+groups.size} charts</em></header><section className="section-chart quality-chart" aria-label="RM and FG quality report indicators"><header><div><b>RM versus FG report indicators</b><span>Report and FCO-pass counts are shown separately.</span></div><em>reports</em></header><ChartBars data={statusData} /></section>{[...groups].map(([key,group])=><section className="section-chart quality-chart" key={key} aria-label={`${group.group} RM and FG comparison`}><header><div><b>{group.group} · RM versus FG</b><span>Average recorded values · unit {group.unit}</span></div><em>{group.unit}</em></header><ChartBars data={group.data} /></section>)}{groups.size===0 && <div className="chart-empty">Individual RM and FG readings will appear after classified quality entries are available.</div>}</section>;
}

function SectionChart({ report }:{ report:SegmentReport }) {
  if (report.qualityComparison) return <QualitySplitCharts report={report} comparison={report.qualityComparison} />;
  const numericGroups=new Map<string,Array<{label:string;value:number;display:string}>>();
  const processKeys=new Set(["netYieldKg","packagedQuantityKg"]);
  for (const field of report.consolidatedFields) if (field.type==="number" && field.count>0 && field.chartValue!==undefined && !(report.rawMaterialComparison && processKeys.has(field.key))) { const unit=field.chartUnit || "Number"; const current=numericGroups.get(unit) ?? []; current.push({label:field.label,value:field.chartValue,display:field.chartDisplay ?? field.primary}); numericGroups.set(unit,current); }
  const categorical=report.consolidatedFields.filter((field)=>field.type==="select" && field.breakdown.length>0);
  const fallback=report.metrics.map((metric)=>({label:metric.label,value:chartNumber(metric.value),display:metric.value})).filter((item)=>item.value>0);
  const processCharts=report.rawMaterialComparison ? 2 : 0; const standardCharts=numericGroups.size+categorical.length || (!report.rawMaterialComparison && fallback.length ? 1 : 0);
  return <section className="chart-suite" role="tabpanel" aria-label={`${report.title} consolidated charts`}><header><div><strong>Consolidated charts</strong><span>Each scale contains comparable units only; exact values remain visible.</span></div><em>{processCharts+standardCharts} charts</em></header>{report.rawMaterialComparison && <><section className="section-chart process-chart" aria-label="Raw material to packaged output weight comparison"><header><div><b>Raw material-to-output weight comparison</b><span>All stages use the same kilogram scale.</span></div><em>kg</em></header><ChartBars data={[{label:"Total raw material procured",value:report.rawMaterialComparison.rawMaterialKg,display:`${formatNumber(report.rawMaterialComparison.rawMaterialKg,2)} kg`},{label:"Total production (net)",value:report.rawMaterialComparison.productionKg,display:`${formatNumber(report.rawMaterialComparison.productionKg,2)} kg`},{label:"Packaged output",value:report.rawMaterialComparison.packagedKg,display:`${formatNumber(report.rawMaterialComparison.packagedKg,2)} kg`}]}/></section><section className="section-chart process-chart" aria-label="Production and packaging conversion comparison"><header><div><b>Conversion against total raw material</b><span>Production and packaging measured against raw material procured.</span></div><em>%</em></header><ChartBars data={[{label:"Total production ÷ raw material",value:report.rawMaterialComparison.productionPct,display:`${formatNumber(report.rawMaterialComparison.productionPct,2)}%`},{label:"Packaged output ÷ raw material",value:report.rawMaterialComparison.packagedPct,display:`${formatNumber(report.rawMaterialComparison.packagedPct,2)}%`},{label:"Packaged output ÷ production",value:report.rawMaterialComparison.packagingCoveragePct,display:`${formatNumber(report.rawMaterialComparison.packagingCoveragePct,2)}%`}]}/></section></>}{[...numericGroups].map(([unit,data])=><section className="section-chart" key={unit} aria-label={`${report.title} ${unit} chart`}><header><div><b>{unit === "₹" ? "Financial values" : unit === "kg" ? "Weight and quantity" : unit === "%" ? "Percentage readings" : unit === "°C" ? "Temperature readings" : "Numeric values"}</b><span>Consolidated totals and averages · unit {unit}</span></div><em>{data.length} measures</em></header><ChartBars data={data} /></section>)}{categorical.map((field)=><section className="section-chart" key={field.key} aria-label={`${field.label} distribution chart`}><header><div><b>{field.label} distribution</b><span>Number of entries recorded under each value</span></div><em>{field.breakdown.length} groups</em></header><ChartBars data={field.breakdown.map((item)=>({label:item.label,value:item.value,display:String(item.value)}))} /></section>)}{!report.rawMaterialComparison && numericGroups.size===0 && categorical.length===0 && fallback.length>0 && <section className="section-chart"><header><div><b>Consolidated metric view</b><span>Imported or calculated batch totals</span></div></header><ChartBars data={fallback} /></section>}{!report.rawMaterialComparison && numericGroups.size===0 && categorical.length===0 && fallback.length===0 && <div className="chart-empty">Save entries in this section to build consolidated charts.</div>}</section>;
}

function ReportEntryItem({ entry, onEdit, onDelete }:{ entry:Entry; onEdit:(entry:Entry)=>void; onDelete:(entry:Entry)=>void }) {
  const fields=entryCapturedFields(entry); const identity=entrySystemIdentity(entry); const pending=entry.deletionStatus==="pending";
  return <article className={`report-entry-item ${pending ? "pending-deletion" : ""}`}><div className="report-entry-main"><div><b>{entry.title}</b><span>Entry date {displayDate(entry.entryDate)}</span></div><strong>{pending ? "Deletion pending" : entry.amount ? formatMoney(entry.amount) : entry.quantity ? `${formatNumber(entry.quantity,2)}${entry.module === "expenses" ? "" : " kg"}` : entry.status || "Saved"}</strong></div><div className="punch-line"><span>Sr. No. {identity.serial}</span><span>Code {identity.recordCode}</span><span>Batch {entry.productionBatchCode}</span><span>◷ Punched {displayTimestamp(entry.createdAt)}</span><span>◎ By {entry.createdBy || "Field operator"}</span>{pending && <span>Approval required</span>}</div><div className="report-entry-actions"><button type="button" disabled={pending} onClick={()=>onEdit(entry)}>✎ {pending ? "Locked" : "Edit entry"}</button><button type="button" className="danger" disabled={pending} onClick={()=>onDelete(entry)}>{pending ? "Request submitted" : "Request deletion"}</button></div><section className="captured-values"><header><div><strong>Captured values</strong><span>All {fields.length} fields are visible</span></div></header><dl>{fields.map((field)=><div key={field.key} className={field.value==="Not entered" ? "not-entered" : ""}><dt>{field.label}</dt><dd>{field.value}</dd></div>)}</dl></section></article>;
}

function AdminView({currentUser,app,ask,onSignOut,onToast,onDataChanged}:{currentUser:AppUser;app:AppStatus;ask:PromptAsk;onSignOut:()=>void;onToast:(message:string)=>void;onDataChanged:()=>void}) {
  const [users,setUsers]=useState<AppUser[]>([]); const [activities,setActivities]=useState<ActivityLog[]>([]); const [accesses,setAccesses]=useState<AccessLog[]>([]); const [requests,setRequests]=useState<DeletionRequest[]>([]);
  const [loading,setLoading]=useState(true); const [message,setMessage]=useState(""); const [activityQuery,setActivityQuery]=useState(""); const [accessQuery,setAccessQuery]=useState(""); const [accessFilter,setAccessFilter]=useState<"all"|"success"|"denied">("all");
  const [newUser,setNewUser]=useState({fullName:"",email:"",mobile:"",designation:"",role:"Admin",permittedSegments:"all",permittedBatches:"all"});
  const load=useCallback(async(quiet=false)=>{ if (!quiet) setLoading(true); try { const response=await fetch("/api/admin",{cache:"no-store"}); const payload=await response.json() as {users?:AppUser[];activities?:ActivityLog[];accessLogs?:AccessLog[];deletionRequests?:DeletionRequest[];error?:string}; if (!response.ok) throw new Error(payload.error ?? "Could not load administration."); setUsers(payload.users ?? []); setActivities(payload.activities ?? []); setAccesses(payload.accessLogs ?? []); setRequests(payload.deletionRequests ?? []); setMessage(""); } catch(caught) { setMessage(caught instanceof Error ? caught.message : "Could not load administration."); } finally { if (!quiet) setLoading(false); } },[]);
  useEffect(()=>{ const initial=window.setTimeout(()=>void load(),0); const refresh=window.setInterval(()=>void load(true),30000); return ()=>{ window.clearTimeout(initial); window.clearInterval(refresh); }; },[load]);
  const keyPerson=currentUser.isKeyPerson; const deletionApprover=isDeletionApprover(currentUser.email); const pending=requests.filter((item)=>item.status==="pending"); const recycled=requests.filter((item)=>item.status==="approved");
  const successfulAccesses=accesses.filter((item)=>item.outcome==="success").length; const deniedAccesses=accesses.filter((item)=>item.outcome==="denied").length;
  const visibleAccesses=accesses.filter((item)=>(accessFilter==="all" || item.outcome===accessFilter) && `${item.fullName ?? ""} ${item.email ?? ""} ${item.deviceInfo ?? ""} ${item.ipAddress ?? ""}`.toLowerCase().includes(accessQuery.toLowerCase()));
  const visibleActivities=activities.filter((item)=>`${item.actorName} ${item.action} ${item.summary} ${item.productionBatchCode ?? ""}`.toLowerCase().includes(activityQuery.toLowerCase()));
  async function decide(item:DeletionRequest,decision:"approve"|"reject"|"restore") { const note=decision==="reject" ? await ask({ title:"Reject deletion request", message:"Reason for rejecting this deletion request (optional). The requester is notified by Outlook email.", placeholder:"Reason", confirmLabel:"Reject request", optional:true }) : ""; if (note===null) return; try { const response=await fetch("/api/admin/deletions",{method:"PATCH",headers:{"Content-Type":"application/json"},body:JSON.stringify({id:item.id,decision,note})}); const payload=await response.json() as {error?:string}; if (!response.ok) throw new Error(payload.error ?? "Could not update this request."); await load(); onDataChanged(); } catch(caught) { setMessage(caught instanceof Error ? caught.message : "Could not update this request."); } }
  async function addUser(event:React.FormEvent) { event.preventDefault(); try { const response=await fetch("/api/admin",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(newUser)}); const payload=await response.json() as {error?:string}; if (!response.ok) throw new Error(payload.error ?? "Could not add this user."); onToast(`${newUser.fullName} added. Use “Email setup instructions” on their card to invite them.`); setNewUser({fullName:"",email:"",mobile:"",designation:"",role:"Admin",permittedSegments:"all",permittedBatches:"all"}); await load(); } catch(caught) { setMessage(caught instanceof Error ? caught.message : "Could not add this user."); } }
  async function invite(user:AppUser) { try { const response=await fetch("/api/app/invite",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({email:user.email})}); const payload=await response.json() as {error?:string}; if (!response.ok) throw new Error(payload.error ?? "Could not send the invitation."); onToast(app.mode==="demo" ? `Demo mode: invitation for ${user.email} recorded, not sent` : `Setup instructions emailed to ${user.email}`); } catch(caught) { setMessage(caught instanceof Error ? caught.message : "Could not send the invitation."); } }
  function exportActivity() { const quote=(value:unknown)=>`"${String(value ?? "").replaceAll('"','""')}"`; const rows=[["Date & Time","User","Email","Action","Summary","Batch","Segment"],...visibleActivities.map((item)=>[new Date(item.createdAt).toISOString(),item.actorName,item.actorEmail,item.action,item.summary,item.productionBatchCode ?? "",item.module ?? ""])]; const blob=new Blob([rows.map((row)=>row.map(quote).join(",")).join("\n")],{type:"text/csv;charset=utf-8"}); const url=URL.createObjectURL(blob); const link=document.createElement("a"); link.href=url; link.download=`VermiTrack-user-activity-${todayIso()}.csv`; link.click(); URL.revokeObjectURL(url); }
  return <section className="page-view admin-view"><div className="page-head"><div><p className="eyebrow">{keyPerson ? "PRIVATE OWNER CONTROL" : deletionApprover ? "SECURE DELETION CONTROL" : "OPERATIONAL ADMINISTRATION"}</p><h1>{keyPerson ? "Pulkit’s activity centre" : deletionApprover ? "Deletion approval centre" : "Administration"}</h1><p>{keyPerson ? "Track every login and recorded action, manage approved users and control deletion requests." : deletionApprover ? "Review, approve, reject or restore deletion requests. User and activity tracking remain restricted to Pulkit Singh." : "Deletion decisions are restricted to Pulkit Singh and Pritam Suryavanshi. User and activity tracking are restricted to Pulkit Singh."}</p></div><button type="button" className="signout-link" onClick={onSignOut}>Sign out</button></div>
    {keyPerson && <section className="owner-private-banner"><span>KP</span><div><b>Key Person · Pulkit Singh</b><p>This private control centre is visible only to your approved account.</p></div><em>OWNER ONLY</em></section>}
    {!keyPerson && deletionApprover && <section className="owner-private-banner"><span>DA</span><div><b>Deletion Approver · {currentUser.fullName}</b><p>Only Pulkit Singh and Pritam Suryavanshi can access these deletion decisions.</p></div><em>APPROVER ONLY</em></section>}
    <section className="admin-overview">{keyPerson ? <><article><span>Active users</span><b>{users.filter((user)=>user.status==="active").length}</b><small>{users.length} approved accounts</small></article><article><span>Successful sessions</span><b>{successfulAccesses}</b><small>Latest 250 access records</small></article><article><span>Denied attempts</span><b>{deniedAccesses}</b><small>Unapproved account access</small></article><article><span>Audit events</span><b>{activities.length}</b><small>Latest 250 actions</small></article></> : deletionApprover ? <><article><span>Pending approvals</span><b>{pending.length}</b><small>Deletion requests</small></article><article><span>Recycle bin</span><b>{recycled.length}</b><small>Restorable entries</small></article><article><span>Signed in as</span><b className="admin-name">{currentUser.fullName}</b><small>Deletion Approver</small></article><article><span>Approval scope</span><b className="admin-name">Deletion only</b><small>Private activity stays hidden</small></article></> : <><article><span>Signed in as</span><b className="admin-name">{currentUser.fullName}</b><small>{currentUser.role}</small></article><article><span>Deletion approval</span><b className="admin-name">Restricted</b><small>Pulkit & Pritam only</small></article></>}</section>
    {message && <div className="form-error" role="alert">! {message}</div>}{loading && <div className="admin-loading">Refreshing secure activity…</div>}
    {deletionApprover ? <><div className="section-heading"><h2>Deletion approvals</h2><span className="count-pill">{pending.length} pending</span></div>
    <section className="approval-list">{pending.length ? pending.map((item)=><article key={item.id}><div className="approval-icon">!</div><div><b>{item.entry?.title ?? `Entry #${item.entryId}`}</b><span>{item.entry?.productionBatchCode} · {item.entry ? moduleFor(item.entry.module).title : "Register entry"}</span><p>“{item.reason}”</p><small>Requested by {item.requestedByName} · {displayTimestamp(item.createdAt)}</small></div><div className="approval-actions"><button onClick={()=>void decide(item,"reject")}>Reject</button><button className="danger-solid" onClick={()=>void decide(item,"approve")}>Approve deletion</button></div></article>) : <div className="admin-empty">No deletion requests are waiting for approval.</div>}</section>
    {recycled.length>0 && <><div className="section-heading"><h2>90-day recycle bin</h2><span className="count-pill neutral">{recycled.length} retained</span></div><section className="recycle-list">{recycled.map((item)=><article key={item.id}><div><b>{item.entry?.title ?? `Entry #${item.entryId}`}</b><span>Deleted by {item.decidedByName ?? "Deletion approver"} · {item.decidedAt ? displayTimestamp(item.decidedAt) : "Date unavailable"}</span></div><button onClick={()=>void decide(item,"restore")}>Restore entry</button></article>)}</section></>}</> : <section className="admin-empty">Deletion requests are routed privately to Pulkit Singh and Pritam Suryavanshi. No deletion decisions or other users’ activity are visible in this account.</section>}
    {keyPerson && <><div className="section-heading"><h2>Login & session details</h2><span className="count-pill neutral">{accesses.length} records</span></div><div className="access-filter-bar" role="group" aria-label="Filter login and session records"><button className={accessFilter==="all" ? "active" : ""} onClick={()=>setAccessFilter("all")}>All <span>{accesses.length}</span></button><button className={accessFilter==="success" ? "active" : ""} onClick={()=>setAccessFilter("success")}>Successful <span>{successfulAccesses}</span></button><button className={accessFilter==="denied" ? "active" : ""} onClick={()=>setAccessFilter("denied")}>Denied <span>{deniedAccesses}</span></button></div><div className="access-search"><input value={accessQuery} onChange={(event)=>setAccessQuery(event.target.value)} placeholder="Search name, email, device or IP…" /><span>{visibleAccesses.length} shown</span></div><section className="access-log-list">{visibleAccesses.length ? visibleAccesses.map((item)=><article key={item.id} className={item.outcome}><span>{item.outcome === "success" ? "✓" : "!"}</span><div><b>{item.fullName || (item.outcome==="denied" ? "Name not supplied" : "Unknown account")}</b><small>{item.email || "No email available"}</small><p>{item.deviceInfo || "Unknown device"} · {item.method === "device" ? "Quick Unlock" : item.method === "microsoft" ? "Microsoft 365 sign-in · VermiTrack desktop" : "Official email sign-in"}</p></div><em>{item.outcome === "success" ? (item.endedAt ? "Completed" : "Active") : "Denied"}</em><dl className="session-detail-grid"><div><dt>{item.outcome === "success" ? "Login / start time" : "Attempted at"}</dt><dd>{displayTimestamp(item.createdAt)}</dd></div><div><dt>{item.outcome === "success" ? "Last activity" : "Access result"}</dt><dd>{item.outcome === "success" ? displayTimestamp(item.lastSeenAt ?? item.createdAt) : "Account not authorised"}</dd></div><div><dt>Session end</dt><dd>{item.outcome === "success" ? (item.endedAt ? displayTimestamp(item.endedAt) : "Active now") : displayTimestamp(item.endedAt ?? item.createdAt)}</dd></div><div><dt>Duration in app</dt><dd>{item.outcome === "success" ? formatDuration(item.durationSeconds) : "Not admitted"}</dd></div></dl><footer><strong>{item.method === "device" ? "Device security unlock" : "Official account access"}</strong><strong>{item.ipAddress ? `IP ${item.ipAddress}` : "Desktop app"}</strong></footer></article>) : <div className="admin-empty">No {accessFilter==="all" ? "login or session" : accessFilter} records are available for this filter.</div>}</section>
    <div className="section-heading"><h2>Approved users</h2><span className="count-pill neutral">{users.length} accounts</span></div>
    <section className="user-management"><div className="user-list">{users.map((user)=><UserAccessCard key={user.id} user={user} currentUser={currentUser} onSaved={load} onError={setMessage} onInvite={(target)=>void invite(target)} />)}</div><details className="add-user-panel"><summary>＋ Add an approved user</summary><form onSubmit={addUser}><label><span>Full name*</span><input required value={newUser.fullName} onChange={(event)=>setNewUser((value)=>({...value,fullName:event.target.value}))} /></label><label><span>Official email*</span><input type="email" required value={newUser.email} onChange={(event)=>setNewUser((value)=>({...value,email:event.target.value}))} /></label><label><span>Mobile</span><input value={newUser.mobile} onChange={(event)=>setNewUser((value)=>({...value,mobile:event.target.value}))} /></label><label><span>Designation*</span><input required value={newUser.designation} onChange={(event)=>setNewUser((value)=>({...value,designation:event.target.value}))} /></label><label><span>Role</span><select value={newUser.role} onChange={(event)=>setNewUser((value)=>({...value,role:event.target.value}))}><option>Admin</option><option>Plant Manager</option><option>Production Operator</option><option>Quality Officer</option><option>Stores / Packaging</option><option>Viewer</option></select></label><label><span>Permitted segments</span><input value={newUser.permittedSegments} onChange={(event)=>setNewUser((value)=>({...value,permittedSegments:event.target.value}))} placeholder="all or comma-separated segment codes" /></label><label><span>Permitted batches</span><input value={newUser.permittedBatches} onChange={(event)=>setNewUser((value)=>({...value,permittedBatches:event.target.value}))} placeholder="all or BATCH-04,BATCH-05" /></label><button type="submit">Add approved user</button></form></details></section>
    <ImportPanel onImported={(message)=>{ onToast(message); void load(); onDataChanged(); }} onError={setMessage} />
    <div className="section-heading"><h2>Complete audit activity</h2><button onClick={exportActivity} disabled={!visibleActivities.length}>Export CSV</button></div><div className="audit-tools"><input value={activityQuery} onChange={(event)=>setActivityQuery(event.target.value)} placeholder="Search user, action, batch or entry…" /><span>{visibleActivities.length} events</span></div><section className="audit-list">{visibleActivities.map((item)=><article key={item.id}><span className={`audit-action ${item.action.toLowerCase()}`}>{item.action.replaceAll("_"," ")}</span><div><b>{item.summary}</b><small>{item.actorName} · {item.actorEmail}{item.productionBatchCode ? ` · ${item.productionBatchCode}` : ""}</small></div><time>{displayTimestamp(item.createdAt)}</time></article>)}</section></>}
  </section>;
}

function UserAccessCard({user,currentUser,onSaved,onError,onInvite}:{user:AppUser;currentUser:AppUser;onSaved:()=>Promise<void>;onError:(message:string)=>void;onInvite:(user:AppUser)=>void}) {
  const [draft,setDraft]=useState(user); const [saving,setSaving]=useState(false); const initials=user.fullName.split(/\s+/).map((part)=>part[0]).join("").slice(0,2).toUpperCase();
  async function save() { setSaving(true); try { const response=await fetch("/api/admin",{method:"PATCH",headers:{"Content-Type":"application/json"},body:JSON.stringify(draft)}); const payload=await response.json() as {error?:string}; if (!response.ok) throw new Error(payload.error ?? "Could not update this user."); await onSaved(); } catch(caught) { onError(caught instanceof Error ? caught.message : "Could not update this user."); } finally { setSaving(false); } }
  return <details className={`user-card ${user.status} ${user.isKeyPerson ? "key-person" : ""}`}><summary><span>{initials}</span><div><b>{user.fullName}</b><small>{user.designation} · {user.email}</small></div><em>{user.isKeyPerson ? "Admin · Key Person" : user.role}</em><i>{user.isKeyPerson ? "Protected owner" : user.status}</i></summary><div className="user-access-form"><label><span>Role</span><select value={draft.role} disabled={user.isKeyPerson} onChange={(event)=>setDraft((value)=>({...value,role:event.target.value}))}><option>Admin</option><option>Plant Manager</option><option>Production Operator</option><option>Quality Officer</option><option>Stores / Packaging</option><option>Viewer</option></select></label><label><span>Permitted segments</span><input value={draft.permittedSegments} onChange={(event)=>setDraft((value)=>({...value,permittedSegments:event.target.value}))} /></label><label><span>Permitted batches</span><input value={draft.permittedBatches} onChange={(event)=>setDraft((value)=>({...value,permittedBatches:event.target.value}))} /></label><label><span>Account status</span><select value={draft.status} disabled={user.isKeyPerson || user.email===currentUser.email} onChange={(event)=>setDraft((value)=>({...value,status:event.target.value}))}><option value="active">Active</option><option value="inactive">Inactive</option></select></label><div className="user-contact"><span>{user.mobile || "No mobile recorded"}</span><span>{user.lastActiveAt ? `Last active ${displayTimestamp(user.lastActiveAt)}` : "Not signed in yet"}</span></div><div className="user-card-actions"><button type="button" onClick={()=>void save()} disabled={saving}>{saving ? "Saving…" : "Save access"}</button>{user.email!==currentUser.email && user.status==="active" && <button type="button" className="secondary" onClick={()=>onInvite(user)}>✉ Email setup instructions</button>}</div></div></details>;
}

function MoreView({ batch, baseline, batches, app, currentUser, onSignOut, onToast, onCapture, onAddBatch, onSelectBatch, onAdmin }:{ batch:ProductionBatch; baseline:Baseline | undefined; batches:ProductionBatch[]; app:AppStatus; currentUser:AppUser; onSignOut:()=>void; onToast:(message:string)=>void; onCapture:(id?:string)=>void; onAddBatch:()=>void; onSelectBatch:(code:string)=>void; onAdmin:()=>void }) {
  const demo=app.mode==="demo";
  async function copySetupCode() { if (!app.setupCode) return; try { await navigator.clipboard.writeText(app.setupCode); onToast("Setup code copied – paste it into an Outlook email or Teams chat"); } catch { onToast("Select the code and copy it manually"); } }
  function openSite() { if (app.siteUrl) void fetch("/api/app/open",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({url:app.siteUrl})}); }
  return <section className="page-view"><div className="page-head"><div><p className="eyebrow">REGISTER CONTROL</p><h1>Batches & setup</h1><p>Manage historical and new production batches, then open any register.</p></div><button className="primary-small" onClick={onAddBatch}>＋ New batch</button></div><section className="navjyoti-brand-card"><div className="navjyoti-logo-full" role="img" aria-label="Navjyoti logo" /><div><p>OPERATED BY</p><b>Navjyoti Commodity Management Services Ltd.</b><span>Batch-wise vermicompost production, quality and traceability.</span></div></section><section className="admin-launch-card"><span>AC</span><div><b>Administration & activity control</b><p>Manage approved users, audit every change and decide deletion requests.</p></div><button onClick={onAdmin}>Open admin panel</button></section>
    <section className={`quick-unlock-card connection-card ${demo ? "" : "active"}`}><span aria-hidden="true">{demo ? "!" : "☁"}</span><div><b>{demo ? "Demo mode – not shared" : "Connected to Microsoft 365"}</b><p>{demo ? "Records are stored only on this computer. Leave demo mode and connect your SharePoint site to share live records with colleagues." : `Signed in as ${app.account ?? currentUser.email}. Every entry is saved to the shared SharePoint site, so colleagues see it within seconds.`}</p><small>{demo ? `Demo data: ${app.dataLocation ?? "this computer"}` : app.siteUrl} · VermiTrack desktop {app.version}</small></div><div className="quick-unlock-actions">{!demo && <button onClick={openSite}>Open SharePoint site</button>}<button className="secondary" onClick={onSignOut}>{demo ? "Leave demo mode" : "Sign out"}</button></div></section>
    {!demo && app.setupCode && <section className="install-card desktop-card"><span>⇪</span><div><b>Add colleagues</b><p>Colleagues install VermiTrack, choose <strong>I have a setup code</strong>, paste the code below and sign in with their own Outlook work account. They must be approved users (Admin → Approved users) and members of the SharePoint site.</p><code className="setup-code">{app.setupCode}</code><button type="button" onClick={()=>void copySetupCode()}>Copy setup code</button></div></section>}
    <div className="section-heading"><h2>Production batches</h2><span className="count-pill neutral">{batches.length} available</span></div><section className="batch-directory">{batches.map((item)=><button className={item.code === batch.code ? "selected" : ""} key={item.code} onClick={()=>onSelectBatch(item.code)}><span>{item.name.replace("Batch ","B")}</span><div><b>{item.name}</b><small>{item.code} · FY {item.financialYear || "Not set"}</small></div><em>{item.status}</em><i>›</i></button>)}</section><div className="section-heading"><h2>Production registers</h2></div><section className="module-directory">{registerModules.map((moduleDef)=><button key={moduleDef.id} onClick={()=>onCapture(moduleDef.id)}><span className={`register-badge ${moduleDef.tone}`}>{moduleDef.short}</span><div><b>{moduleDef.title}</b><small>{moduleDef.detail}</small></div><i>＋</i></button>)}</section><div className="section-heading"><h2>{batch.name} source</h2></div><section className="source-card"><div><span>{baseline ? "XL" : "BT"}</span><div><b>{baseline?.label ?? `${batch.name} register`}</b><small>{baseline?.source ?? "No imported baseline; totals will build from entries."}</small></div></div><dl><div><dt>Production batch</dt><dd>{batch.code}</dd></div><div><dt>Status</dt><dd>{batch.status}</dd></div><div><dt>Baseline production</dt><dd>{formatNumber(baseline?.netYieldKg ?? 0)} kg</dd></div><div><dt>QC baseline</dt><dd>{baseline?.qcPassed ?? 0}</dd></div></dl></section></section>;
}

function BottomNav({ view, onNavigate, onCapture }:{ view:View; onNavigate:(view:View)=>void; onCapture:()=>void }) { return <nav className="bottom-nav" aria-label="Primary navigation"><button className={view === "home" ? "active" : ""} onClick={()=>onNavigate("home")}><span>⌂</span>Home</button><button className={view === "records" ? "active" : ""} onClick={()=>onNavigate("records")}><span>▤</span>Records</button><button className="capture-button" onClick={onCapture}><span>＋</span>Capture</button><button className={view === "insights" ? "active" : ""} onClick={()=>onNavigate("insights")}><span>▦</span>Reports</button><button className={view === "more" || view === "admin" ? "active" : ""} onClick={()=>onNavigate("more")}><span>☷</span>More</button></nav>; }

function CaptureSheet({ batches, entries, defaultBatchCode, selectedId, existingEntry, onSelect, onClose, onSaved }:{ batches:ProductionBatch[]; entries:Entry[]; defaultBatchCode:string; selectedId:string | null; existingEntry:Entry | null; onSelect:(id:string | null)=>void; onClose:()=>void; onSaved:(entry:Entry)=>void }) {
  const selected=selectedId ? moduleFor(selectedId) : null;
  return <div className="sheet-backdrop" onClick={onClose}><section className={`capture-sheet ${selected ? "form-open" : ""}`} onClick={(event)=>event.stopPropagation()}><div className="sheet-handle" />{!selected ? <><div className="sheet-title"><div><p className="eyebrow">NEW ENTRY · {defaultBatchCode}</p><h2>What are you recording?</h2></div><button onClick={onClose} aria-label="Close">×</button></div><div className="register-grid">{registerModules.map((moduleDef)=><button key={moduleDef.id} onClick={()=>onSelect(moduleDef.id)}><span className={`register-badge ${moduleDef.tone}`}>{moduleDef.short}</span><div><b>{moduleDef.title}</b><small>{moduleDef.detail}</small></div><i>›</i></button>)}</div></> : <CaptureForm key={`${selected.id}-${existingEntry?.id ?? "new"}`} batches={batches} entries={entries} defaultBatchCode={defaultBatchCode} selected={selected} existingEntry={existingEntry} onBack={existingEntry ? onClose : ()=>onSelect(null)} onClose={onClose} onSaved={onSaved} />}</section></div>;
}

function initialDraft(selected:RegisterModule, defaultBatchCode:string, existingEntry:Entry | null) {
  const defaults:Record<string,string>={};
  for (const field of selected.fields) if (field.type === "date" && field.required) defaults[field.key]=todayIso();
  if (selected.id==="raw-material") defaults.lotCode="";
  if (selected.fields.some((field)=>field.key==="productionBatchCode")) defaults.productionBatchCode=defaultBatchCode;
  if (existingEntry) { const saved=parseDetails(existingEntry); for (const field of selected.fields) if (saved[field.key] !== null && saved[field.key] !== undefined) defaults[field.key]=String(saved[field.key]); if (selected.id==="harvest") defaults.productionBatchCode=existingEntry.productionBatchCode; return { values:defaults, productionBatchCode:existingEntry.productionBatchCode }; }
  return { values:defaults, productionBatchCode:defaultBatchCode };
}

function CaptureForm({ selected, batches, entries, defaultBatchCode, existingEntry, onBack, onClose, onSaved }:{ selected:RegisterModule; batches:ProductionBatch[]; entries:Entry[]; defaultBatchCode:string; existingEntry:Entry | null; onBack:()=>void; onClose:()=>void; onSaved:(entry:Entry)=>void }) {
  const [draft]=useState(()=>initialDraft(selected,defaultBatchCode,existingEntry)); const [values,setValues]=useState<Record<string,string>>(draft.values); const [productionBatchCode,setProductionBatchCode]=useState(draft.productionBatchCode); const [saving,setSaving]=useState(false); const [formError,setFormError]=useState(""); const groups=[...new Set(selected.fields.map((field)=>field.group))];
  const harvestedBatchOptions=useMemo(()=>{ const options=new Map<string,string>(); for (const entry of entries) if (entry.module==="harvest" && entry.productionBatchCode===productionBatchCode) { const batchNumber=String(parseDetails(entry).batchNumber ?? "").trim(); if (batchNumber) options.set(batchNumber,`${batchNumber} · harvested`); } if (!options.size) { const productionBatch=batches.find((batch)=>batch.code===productionBatchCode); options.set(productionBatchCode,`${productionBatch?.name ?? productionBatchCode} · production batch`); } const saved=String(values.batchNumber ?? "").trim(); if (saved && !options.has(saved)) options.set(saved,`${saved} · saved value`); return [...options].map(([value,label])=>({value,label})); },[batches,entries,productionBatchCode,values.batchNumber]);
  useEffect(()=>{ if (selected.id!=="raw-material") return; const purchaseDate=values.purchaseDate ?? ""; const materialType=values.materialType ?? ""; const existingDetails=existingEntry ? parseDetails(existingEntry) : null; const unchanged=existingDetails && String(existingDetails.purchaseDate ?? "")===purchaseDate && String(existingDetails.materialType ?? "")===materialType; if (unchanged || !purchaseDate || !materialType) return; const controller=new AbortController(); const params=new URLSearchParams({mode:"rm-code",purchaseDate,materialType}); if (existingEntry) params.set("excludeEntryId",String(existingEntry.id)); fetch(`/api/register?${params}`,{cache:"no-store",signal:controller.signal}).then(async(response)=>{ const payload=await response.json() as {code?:string}; if (response.ok && payload.code) setValues((current)=>current.purchaseDate===purchaseDate && current.materialType===materialType ? {...current,lotCode:payload.code!} : current); }).catch(()=>undefined); return ()=>controller.abort(); },[existingEntry,selected.id,values.materialType,values.purchaseDate]);
  function updateValue(key:string,value:string) { setValues((current)=>{ const next={...current,[key]:value}; if (selected.id==="raw-material" && (key==="purchaseDate" || key==="materialType")) next.lotCode=rmLotCode(next.purchaseDate,next.materialType); if (selected.id==="harvest" && (key==="packagingType" || key==="packetsPrepared")) { const unitWeight=packageWeightKg(next.packagingType); const packets=capturedNumber(next.packetsPrepared); if (unitWeight!==null && packets!==null) next.packagedQuantityKg=String(unitWeight*packets); } return next; }); }
  async function submit(event:React.FormEvent) { event.preventDefault(); if (!productionBatchCode) { setFormError("Production batch is required."); return; } const missing=selected.fields.find((field)=>field.required && !String(values[field.key] ?? "").trim()); if (missing) { setFormError(`${missing.label} is required.`); document.getElementById(`field-${missing.key}`)?.focus(); return; } setSaving(true); setFormError(""); try { const response=await fetch("/api/register",{method:existingEntry ? "PATCH" : "POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({id:existingEntry?.id,module:selected.id,productionBatchCode,values})}); const payload=await response.json() as {entry?:Entry;error?:string}; if (!response.ok || !payload.entry) throw new Error(payload.error ?? `Could not ${existingEntry ? "update" : "save"} this entry.`); onSaved(payload.entry); } catch (caught) { setFormError(caught instanceof Error ? caught.message : `Could not ${existingEntry ? "update" : "save"} this entry.`); } finally { setSaving(false); } }
  return <form onSubmit={submit}><div className="form-head"><button type="button" onClick={onBack}>←</button><span className={`register-badge ${selected.tone}`}>{selected.short}</span><div><p className="eyebrow">{existingEntry ? "EDIT REGISTER ENTRY" : "NEW REGISTER ENTRY"}</p><h2>{selected.title}</h2></div><button type="button" className="close-form" onClick={onClose}>×</button></div><p className="form-intro">{existingEntry ? `Originally punched ${displayTimestamp(existingEntry.createdAt)} by ${existingEntry.createdBy}.` : "Serial number, selected batch code and unique register code are generated automatically. Enter only the operational values below."}</p>{formError && <div className="form-error" role="alert">! {formError}</div>}<label className="production-batch-field"><span>Batch code — auto-filled<b>*</b></span><select aria-label="Production batch" value={productionBatchCode} onChange={(event)=>{ const code=event.target.value; setProductionBatchCode(code); if (selected.fields.some((field)=>field.key==="productionBatchCode")) updateValue("productionBatchCode",code); }}>{batches.map((batch)=><option key={batch.code} value={batch.code}>{batch.name} · {batch.code} · {batch.status}</option>)}</select><small>The currently selected production batch is filled automatically; change it only when entering another batch.</small></label><AutoIdentityPreview moduleDef={selected} productionBatchCode={productionBatchCode} existingEntry={existingEntry} /><ComputedPreview moduleDef={selected} values={values} /><div className="field-groups">{groups.map((group,index)=><details key={group} open={index===0}><summary>{group}<span>{selected.fields.filter((field)=>field.group===group).length} fields</span></summary><div className="fields-grid">{selected.fields.filter((field)=>field.group===group).map((field)=><Field key={field.key} field={field} value={values[field.key] ?? ""} options={selected.id==="sales" && field.key==="batchNumber" ? harvestedBatchOptions : undefined} readOnly={(selected.id==="raw-material" && field.key==="lotCode") || (selected.id==="harvest" && field.key==="productionBatchCode")} onChange={(value)=>updateValue(field.key,value)} />)}</div></details>)}</div><div className="form-actions"><button type="button" onClick={onClose}>{existingEntry ? "Cancel" : "Cancel & close"}</button><button type="submit" disabled={saving}>{saving ? "Saving…" : existingEntry ? "Update entry" : `Save in ${batches.find((batch)=>batch.code===productionBatchCode)?.name ?? "batch"}`}</button></div></form>;
}

function Field({ field, value, options, readOnly=false, onChange }:{ field:RegisterField; value:string; options?:Array<{value:string;label:string}>; readOnly?:boolean; onChange:(value:string)=>void }) { const control=field.type === "select" ? <select id={`field-${field.key}`} value={value} onChange={(event)=>onChange(event.target.value)} required={field.required} disabled={readOnly}><option value="">Select…</option>{options ? options.map((option)=><option key={option.value} value={option.value}>{option.label}</option>) : field.options?.map((option)=><option key={option} value={option}>{option}</option>)}</select> : field.type === "textarea" ? <textarea id={`field-${field.key}`} value={value} onChange={(event)=>onChange(event.target.value)} placeholder={field.placeholder} rows={3} required={field.required} readOnly={readOnly} /> : <input id={`field-${field.key}`} type={field.type} value={value} onChange={(event)=>onChange(event.target.value)} placeholder={field.placeholder} step={field.step} required={field.required} readOnly={readOnly} inputMode={field.type === "number" ? "decimal" : undefined} />; return <label className={`${field.type === "textarea" ? "wide" : ""}${readOnly ? " auto-field" : ""}`}><span>{field.label}{field.required && <b>*</b>}</span><div className="input-wrap">{control}{field.unit && <i>{field.unit}</i>}</div>{readOnly && <small>{field.key==="lotCode" ? "Uses the date, first two letters of every material word and an automatic duplicate sequence." : "Generated automatically from the selected production batch."}</small>}</label>; }

function AutoIdentityPreview({ moduleDef, productionBatchCode, existingEntry }:{ moduleDef:RegisterModule; productionBatchCode:string; existingEntry:Entry | null }) {
  const identity=existingEntry ? entrySystemIdentity(existingEntry) : null;
  return <section className="auto-identity-preview"><div><span>Sr. No.</span><b>{identity?.serial ?? "Auto on save"}</b></div><div><span>Batch code</span><b>{productionBatchCode}</b></div><div><span>Register code</span><b>{identity?.recordCode ?? `${productionBatchCode}-${moduleDef.short}-AUTO`}</b></div></section>;
}

function ComputedPreview({ moduleDef, values }:{ moduleDef:RegisterModule; values:Record<string,string> }) { let label="",value=""; if (moduleDef.id === "raw-material") { label="Material value"; value=formatMoney(number(values.quantityKg)*number(values.ratePerUnit)); } if (moduleDef.id === "sales") { label="Total revenue"; value=formatMoney(number(values.quantityKg)*number(values.pricePerKg)); } if (moduleDef.id === "inventory") { label="Calculated closing stock"; value=`${formatNumber(number(values.openingStockKg)+number(values.productionKg)-number(values.salesOutflowKg)-number(values.damageLossKg),2)} kg`; } if (moduleDef.id === "harvest") { const net=number(values.netYieldKg); const packaged=number(values.packagedQuantityKg); label="Packaging coverage"; value=`${net>0 ? ((packaged/net)*100).toFixed(1) : "0.0"}% of net production packed`; } if (!label) return null; return <div className="computed-preview"><span>Auto-calculated</span><div><b>{label}</b><strong>{value}</strong></div></div>; }

function nextBatchDefaults(batches:ProductionBatch[]) { const highest=Math.max(0,...batches.map((batch)=>number(batch.code.match(/(\d+)$/)?.[1]))); const next=highest+1; return { code:`BATCH-${String(next).padStart(2,"0")}`, name:`Batch ${next}`, financialYear:"2026-27", status:"Planned", startDate:"", expectedEndDate:"", notes:"" }; }

function BatchSheet({ batches, onClose, onCreated }:{ batches:ProductionBatch[]; onClose:()=>void; onCreated:(batch:ProductionBatch)=>void }) {
  const [form,setForm]=useState(()=>nextBatchDefaults(batches)); const [saving,setSaving]=useState(false); const [formError,setFormError]=useState("");
  async function submit(event:React.FormEvent) { event.preventDefault(); setSaving(true); setFormError(""); try { const response=await fetch("/api/batches",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(form)}); const payload=await response.json() as {batch?:ProductionBatch;error?:string}; if (!response.ok || !payload.batch) throw new Error(payload.error ?? "Could not create this batch."); onCreated(payload.batch); } catch (caught) { setFormError(caught instanceof Error ? caught.message : "Could not create this batch."); } finally { setSaving(false); } }
  return <div className="sheet-backdrop" onClick={onClose}><section className="capture-sheet batch-sheet" onClick={(event)=>event.stopPropagation()}><div className="sheet-handle" /><form onSubmit={submit}><div className="sheet-title"><div><p className="eyebrow">BATCH MASTER</p><h2>Add a production batch</h2></div><button type="button" onClick={onClose}>×</button></div>{formError && <div className="form-error" role="alert">! {formError}</div>}<div className="batch-form"><label><span>Batch name*</span><input value={form.name} onChange={(event)=>setForm((current)=>({...current,name:event.target.value}))} placeholder="Batch 5" required /></label><label><span>Batch code*</span><input value={form.code} onChange={(event)=>setForm((current)=>({...current,code:event.target.value.toUpperCase()}))} placeholder="BATCH-05" required /></label><label><span>Financial year</span><input value={form.financialYear} onChange={(event)=>setForm((current)=>({...current,financialYear:event.target.value}))} placeholder="2026-27" /></label><label><span>Status*</span><select value={form.status} onChange={(event)=>setForm((current)=>({...current,status:event.target.value}))}><option>Planned</option><option>Active</option><option>Completed</option><option>On hold</option></select></label><label><span>Start date</span><input type="date" value={form.startDate} onChange={(event)=>setForm((current)=>({...current,startDate:event.target.value}))} /></label><label><span>Expected completion</span><input type="date" value={form.expectedEndDate} onChange={(event)=>setForm((current)=>({...current,expectedEndDate:event.target.value}))} /></label><label className="wide"><span>Notes</span><textarea rows={3} value={form.notes} onChange={(event)=>setForm((current)=>({...current,notes:event.target.value}))} placeholder="Raw material, trial or production notes" /></label></div><div className="form-actions"><button type="button" onClick={onClose}>Cancel</button><button type="submit" disabled={saving}>{saving ? "Creating…" : "Create & select batch"}</button></div></form></section></div>;
}

// ------------------------------------------------------------ desktop additions

type PromptOptions = { title:string; message:string; placeholder?:string; confirmLabel:string; danger?:boolean; minLength?:number; optional?:boolean };
type PromptAsk = (options:PromptOptions)=>Promise<string | null>;

/** In-app replacement for window.prompt, which Electron does not support. */
function usePrompt() {
  const [request,setRequest]=useState<(PromptOptions & { resolve:(value:string | null)=>void }) | null>(null);
  const ask=useCallback<PromptAsk>((options)=>new Promise((resolve)=>setRequest({ ...options,resolve })),[]);
  const element=request ? <PromptDialog request={request} onDone={(value)=>{ request.resolve(value); setRequest(null); }} /> : null;
  return { ask,element };
}

function PromptDialog({ request,onDone }:{ request:PromptOptions; onDone:(value:string | null)=>void }) {
  const [value,setValue]=useState(""); const [error,setError]=useState("");
  const minimum=request.optional ? 0 : request.minLength ?? 1;
  function submit(event:React.FormEvent) { event.preventDefault(); if (value.trim().length<minimum) { setError(`Enter at least ${minimum} characters.`); return; } onDone(value.trim()); }
  return <div className="sheet-backdrop" onClick={()=>onDone(null)} onKeyDown={(event)=>{ if (event.key==="Escape") onDone(null); }}><section className="capture-sheet prompt-sheet" role="dialog" aria-modal="true" aria-labelledby="prompt-title" onClick={(event)=>event.stopPropagation()}><form onSubmit={submit}><div className="sheet-title"><div><p className="eyebrow">PLEASE CONFIRM</p><h2 id="prompt-title">{request.title}</h2></div><button type="button" onClick={()=>onDone(null)} aria-label="Close">×</button></div><p className="form-intro">{request.message}</p>{error && <div className="form-error" role="alert">! {error}</div>}<textarea className="prompt-input" autoFocus rows={3} value={value} placeholder={request.placeholder} onChange={(event)=>setValue(event.target.value)} /><div className="form-actions"><button type="button" onClick={()=>onDone(null)}>Cancel</button><button type="submit" className={request.danger ? "danger-solid" : ""}>{request.confirmLabel}</button></div></form></section></div>;
}

const escapeHtml=(value:unknown)=>String(value ?? "").replace(/[&<>"']/g,(char)=>({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" })[char]!);

function buildReportEmail({ batch, metrics, reports, sender, note }:{ batch:ProductionBatch; metrics:Metrics; reports:SegmentReport[]; sender:AppUser; note:string }) {
  const cell="padding:6px 10px;border-bottom:1px solid #e3e7e1;font-size:13px;vertical-align:top";
  const row=(label:string,value:string,extra="")=>`<tr><td style="${cell};color:#5b6c63">${escapeHtml(label)}</td><td style="${cell}"><b>${escapeHtml(value)}</b>${extra ? `<div style="color:#8a958f;font-size:11px">${escapeHtml(extra)}</div>` : ""}</td></tr>`;
  const table=(rows:string)=>`<table style="border-collapse:collapse;width:100%;margin:6px 0 16px">${rows}</table>`;
  const costPerKg=metrics.netYieldKg>0 ? metrics.expenses/metrics.netYieldKg : 0;
  const overview=table([
    row("Batch",`${batch.name} · ${batch.code}`,`${batch.status}${batch.financialYear ? ` · FY ${batch.financialYear}` : ""}`),
    row("Raw material procured",`${formatNumber(metrics.rawMaterialKg,2)} kg`),
    row("Total production (net)",`${formatNumber(metrics.netYieldKg,2)} kg`,`${formatNumber(metrics.recoveryPct,2)}% of raw material`),
    row("Packaged output",`${formatNumber(metrics.packagedKg,2)} kg`,`${formatNumber(metrics.packagingPct,2)}% of raw material`),
    row("Revenue",formatMoney(metrics.revenue)),row("Expenses",formatMoney(metrics.expenses)),
    row("Net operating result",formatMoney(metrics.margin),metrics.margin<0 ? "Cost exceeds recorded revenue" : "Positive recorded margin"),
    row("Production cost",`${formatMoney(costPerKg)}/kg`),row("QC pass records",String(metrics.qcPassed)),row("Active production beds",`${metrics.activeBeds} of ${metrics.totalBeds}`),
  ].join(""));
  const sections=reports.filter((report)=>report.hasData).map((report)=>`<h3 style="margin:18px 0 4px;font-size:15px;color:#1f5a43">${escapeHtml(report.short)} · ${escapeHtml(report.title)} <span style="color:#8a958f;font-weight:normal;font-size:12px">(${report.entryCount} ${report.entryCount===1 ? "entry" : "entries"}${report.includesBaseline ? ", includes imported baseline" : ""})</span></h3>${table(report.metrics.map((metric)=>row(metric.label,metric.value,metric.note)).join(""))}`).join("");
  return `<div style="font-family:Segoe UI,Arial,sans-serif;color:#18332b;max-width:720px">
<div style="background:#163c2e;color:#fff;padding:14px 18px;border-radius:10px 10px 0 0"><b>Navjyoti VermiTrack</b> · Consolidated batch report</div>
<div style="border:1px solid #e3e7e1;border-top:0;padding:18px;border-radius:0 0 10px 10px;background:#fffefa">
<h2 style="margin:0 0 4px;font-size:20px">${escapeHtml(batch.name)} report</h2><p style="margin:0 0 12px;color:#708078;font-size:12px">Prepared by ${escapeHtml(sender.fullName)} on ${escapeHtml(displayTimestamp(Date.now()))}</p>
${note ? `<p style="background:#f5f4ef;border-radius:8px;padding:10px 12px;white-space:pre-wrap">${escapeHtml(note)}</p>` : ""}
<h3 style="margin:14px 0 4px;font-size:15px">Overview</h3>${overview}${sections}
<p style="color:#8a958f;font-size:11px;margin-top:18px">Live figures from the shared VermiTrack register. Open VermiTrack → Reports for charts, every field and individual entries.</p></div></div>`;
}

function EmailReportSheet({ batch, metrics, reports, currentUser, onClose, onSent }:{ batch:ProductionBatch; metrics:Metrics; reports:SegmentReport[]; currentUser:AppUser; onClose:()=>void; onSent:(count:number)=>void }) {
  const [people,setPeople]=useState<Array<{ email:string; fullName:string }>>([]);
  const [selected,setSelected]=useState<Set<string>>(()=>new Set());
  const [extra,setExtra]=useState(""); const [note,setNote]=useState("");
  const [subject,setSubject]=useState(`VermiTrack report · ${batch.name} (${batch.code}) · ${displayDate(todayIso())}`);
  const [sending,setSending]=useState(false); const [error,setError]=useState("");
  useEffect(()=>{ void fetch("/api/users",{cache:"no-store"}).then(async(response)=>{ const payload=await response.json() as { users?:Array<{ email:string; fullName:string }> }; setPeople((payload.users ?? []).filter((person)=>person.email!==currentUser.email)); }).catch(()=>undefined); },[currentUser.email]);
  function toggle(email:string) { setSelected((current)=>{ const next=new Set(current); if (next.has(email)) next.delete(email); else next.add(email); return next; }); }
  async function send(event:React.FormEvent) {
    event.preventDefault();
    const to=[...new Set([...selected,...extra.split(/[,;\s]+/).map((value)=>value.trim().toLowerCase()).filter(Boolean)])];
    if (!to.length) { setError("Choose at least one recipient."); return; }
    setSending(true); setError("");
    try {
      const response=await fetch("/api/email-report",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({ to,subject,html:buildReportEmail({ batch,metrics,reports,sender:currentUser,note }) })});
      const payload=await response.json() as { sent?:number; error?:string };
      if (!response.ok) throw new Error(payload.error ?? "Could not send the report.");
      onSent(payload.sent ?? to.length);
    } catch (caught) { setError(caught instanceof Error ? caught.message : "Could not send the report."); }
    finally { setSending(false); }
  }
  return <div className="sheet-backdrop" onClick={onClose}><section className="capture-sheet batch-sheet" role="dialog" aria-modal="true" aria-labelledby="email-report-title" onClick={(event)=>event.stopPropagation()}><form onSubmit={send}><div className="sheet-title"><div><p className="eyebrow">OUTLOOK · {batch.code}</p><h2 id="email-report-title">Email this report</h2></div><button type="button" onClick={onClose} aria-label="Close">×</button></div><p className="form-intro">Sends the {batch.name} overview and all segment totals from your own Outlook mailbox (it appears in your Sent Items).</p>{error && <div className="form-error" role="alert">! {error}</div>}<div className="batch-form"><div className="wide recipient-list"><span>Approved VermiTrack users</span>{people.length ? people.map((person)=><label key={person.email} className="recipient"><input type="checkbox" checked={selected.has(person.email)} onChange={()=>toggle(person.email)} /><b>{person.fullName}</b><small>{person.email}</small></label>) : <small>No other approved users yet.</small>}</div><label className="wide"><span>Other recipients</span><input value={extra} onChange={(event)=>setExtra(event.target.value)} placeholder="name@company.com, another@company.com" /></label><label className="wide"><span>Subject</span><input value={subject} onChange={(event)=>setSubject(event.target.value)} required /></label><label className="wide"><span>Message (optional)</span><textarea rows={3} value={note} onChange={(event)=>setNote(event.target.value)} placeholder="Add a short note for the recipients" /></label></div><div className="form-actions"><button type="button" onClick={onClose}>Cancel</button><button type="submit" disabled={sending}>{sending ? "Sending…" : "Send from my Outlook"}</button></div></form></section></div>;
}

function ImportPanel({ onImported,onError }:{ onImported:(message:string)=>void; onError:(message:string)=>void }) {
  const [busy,setBusy]=useState(false);
  async function importFile(event:React.ChangeEvent<HTMLInputElement>) {
    const file=event.target.files?.[0]; event.target.value="";
    if (!file) return;
    setBusy(true);
    try {
      const response=await fetch("/api/admin/import",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({ csv:await file.text() })});
      const payload=await response.json() as { imported?:number; duplicates?:number; skipped?:number; error?:string };
      if (!response.ok) throw new Error(payload.error ?? "Could not import this file.");
      onImported(`Imported ${payload.imported} new ${payload.imported===1 ? "entry" : "entries"} · ${payload.duplicates} already present${payload.skipped ? ` · ${payload.skipped} unreadable rows skipped` : ""}`);
    } catch (caught) { onError(caught instanceof Error ? caught.message : "Could not import this file."); }
    finally { setBusy(false); }
  }
  return <><div className="section-heading"><h2>Bring records from the web version</h2></div><details className="add-user-panel import-panel"><summary>⇪ Import a “Records → Export CSV” file</summary><div className="import-body"><p>In the old web version of VermiTrack, open <b>Records</b>, choose <b>All registers</b> and click <b>Export CSV</b> for each production batch. Import those files here. Entries that are already present (including the imported Batch 4 register) are skipped automatically, so importing the same file twice is safe.</p><label className="import-button"><input type="file" accept=".csv,text/csv" onChange={(event)=>void importFile(event)} disabled={busy} /><span>{busy ? "Importing…" : "Choose CSV file"}</span></label></div></details></>;
}
