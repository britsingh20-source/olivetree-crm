/**
 * Olivetree Interiors — Lead CRM backend
 * Adds immutable activity history and day-by-day telecaller performance.
 */
const LEADS_SHEET = "Leads";
const ACTIVITY_SHEET = "ActivityLog";
const LEAD_HEADERS = [
  "id","dateAdded","name","phone","source","service","budget","stage",
  "nextFollowUp","lastContact","notes","assignedTo","updatedAt","updatedBy"
];
const ACTIVITY_HEADERS = [
  "eventId","timestamp","date","actor","actorRole","leadId","phone","name",
  "eventType","field","oldValue","newValue","callOutcome","notes"
];
const MONITOR_ROLES = ["director","administrator","admin","manager"];

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  } else {
    const existing = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0];
    headers.forEach((h,i)=>{ if (existing[i] !== h) sh.getRange(1,i+1).setValue(h); });
  }
  return sh;
}

function getLeadSheet_(){ return getOrCreateSheet_(LEADS_SHEET, LEAD_HEADERS); }
function getActivitySheet_(){ return getOrCreateSheet_(ACTIVITY_SHEET, ACTIVITY_HEADERS); }

function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function fmtDate_(v, pattern){
  if (!v) return "";
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), pattern || "yyyy-MM-dd");
  return String(v);
}

function rowsToObjects_(sh, headers){
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i=1;i<rows.length;i++){
    if (!rows[i][0]) continue;
    const o = {};
    headers.forEach((h,j)=>{
      let v = rows[i][j];
      if (v instanceof Date) v = fmtDate_(v, h === "timestamp" || h === "updatedAt" ? "yyyy-MM-dd'T'HH:mm:ssXXX" : "yyyy-MM-dd");
      o[h] = v;
    });
    out.push(o);
  }
  return out;
}

function normalRole_(role){ return String(role||"").trim().toLowerCase(); }
function canMonitor_(role){ return MONITOR_ROLES.indexOf(normalRole_(role)) > -1; }

function appendActivity_(ev){
  const sh = getActivitySheet_();
  const now = new Date();
  const row = {
    eventId: ev.eventId || Utilities.getUuid(),
    timestamp: ev.timestamp || now,
    date: ev.date || Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd"),
    actor: ev.actor || "Unknown",
    actorRole: ev.actorRole || "",
    leadId: ev.leadId || "",
    phone: ev.phone || "",
    name: ev.name || "",
    eventType: ev.eventType || "activity",
    field: ev.field || "",
    oldValue: ev.oldValue == null ? "" : ev.oldValue,
    newValue: ev.newValue == null ? "" : ev.newValue,
    callOutcome: ev.callOutcome || "",
    notes: ev.notes || ""
  };
  sh.appendRow(ACTIVITY_HEADERS.map(h=>row[h] != null ? row[h] : ""));
}

function logLeadDiff_(before, after, actor, actorRole){
  const ignore = {updatedAt:true,updatedBy:true};
  LEAD_HEADERS.forEach(h=>{
    if (ignore[h]) return;
    const a = before && before[h] != null ? String(before[h]) : "";
    const b = after && after[h] != null ? String(after[h]) : "";
    if (a !== b) {
      appendActivity_({
        actor, actorRole, leadId:after.id, phone:after.phone, name:after.name,
        eventType:"edit", field:h, oldValue:a, newValue:b
      });
    }
  });
}

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || "list");
  const role = String((e && e.parameter && e.parameter.role) || "");
  if (action === "performance") {
    if (!canMonitor_(role)) return json_({ok:false,error:"forbidden"});
    return json_(buildPerformance_(e.parameter.date || "", e.parameter.actor || ""));
  }
  if (action === "activity") {
    if (!canMonitor_(role)) return json_({ok:false,error:"forbidden"});
    return json_(getActivity_(e.parameter.date || "", e.parameter.actor || ""));
  }
  const leads = rowsToObjects_(getLeadSheet_(), LEAD_HEADERS);
  return json_({ok:true,leads});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents || "{}");
    const actor = body.actor || (body.lead && body.lead.updatedBy) || "Unknown";
    const actorRole = body.actorRole || "";

    if (body.action === "logCall") {
      const l = body.lead || {};
      appendActivity_({
        actor, actorRole, leadId:l.id, phone:l.phone, name:l.name,
        eventType:"call", callOutcome:body.callOutcome || "Spoken",
        notes:body.notes || ""
      });
      return json_({ok:true});
    }

    const sh = getLeadSheet_();
    const data = sh.getDataRange().getValues();
    const id = body.lead && body.lead.id;
    let rowIndex = -1;
    let before = null;
    for (let i=1;i<data.length;i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i+1;
        before = {};
        LEAD_HEADERS.forEach((h,j)=> before[h] = data[i][j] instanceof Date ? fmtDate_(data[i][j],"yyyy-MM-dd") : data[i][j]);
        break;
      }
    }

    if (body.action === "add" || (body.action === "update" && rowIndex === -1)) {
      const lead = Object.assign({}, body.lead || {}, {updatedAt:new Date(), updatedBy:actor});
      sh.appendRow(LEAD_HEADERS.map(h=>lead[h] != null ? lead[h] : ""));
      appendActivity_({actor,actorRole,leadId:lead.id,phone:lead.phone,name:lead.name,eventType:"created",notes:"Lead created"});
    } else if (body.action === "update") {
      const lead = Object.assign({}, body.lead || {}, {updatedAt:new Date(), updatedBy:actor});
      sh.getRange(rowIndex,1,1,LEAD_HEADERS.length).setValues([LEAD_HEADERS.map(h=>lead[h] != null ? lead[h] : "")]);
      logLeadDiff_(before, lead, actor, actorRole);
    } else if (body.action === "delete" && rowIndex > -1) {
      appendActivity_({actor,actorRole,leadId:before.id,phone:before.phone,name:before.name,eventType:"deleted",notes:"Lead deleted"});
      sh.deleteRow(rowIndex);
    } else {
      return json_({ok:false,error:"unknown_action"});
    }
    return json_({ok:true});
  } catch (err) {
    return json_({ok:false,error:String(err)});
  } finally {
    lock.releaseLock();
  }
}

function getActivity_(dateStr, actor){
  const activities = rowsToObjects_(getActivitySheet_(), ACTIVITY_HEADERS);
  const d = dateStr || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const filtered = activities.filter(a => String(a.date) === d && (!actor || String(a.actor) === String(actor)));
  filtered.sort((a,b)=>String(b.timestamp).localeCompare(String(a.timestamp)));
  return {ok:true,date:d,actor:actor||"",activities:filtered};
}

function buildPerformance_(dateStr, actor){
  const d = dateStr || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const acts = rowsToObjects_(getActivitySheet_(), ACTIVITY_HEADERS).filter(a=>String(a.date)===d);
  const actors = {};
  acts.forEach(a=>{
    const who = String(a.actor||"Unknown");
    if (actor && who !== actor) return;
    if (!actors[who]) actors[who] = {
      actor:who,date:d,callActivities:0,uniqueCustomersSpoken:0,editedLeads:0,
      hot:0,followUp:0,siteVisit:0,closed:0,noResponse:0,phones:[],activities:[]
    };
    const p = actors[who];
    p.activities.push(a);
    if (a.eventType === "call") {
      p.callActivities++;
      if (String(a.callOutcome).toLowerCase().indexOf("no response")>-1) p.noResponse++;
    }
    if (a.eventType === "edit") {
      p.editedLeads++;
      const nv = String(a.newValue||"").toLowerCase();
      if (a.field === "stage" && nv.indexOf("hot")>-1) p.hot++;
      if (a.field === "nextFollowUp" && a.newValue) p.followUp++;
      if (a.field === "stage" && nv.indexOf("site visit")>-1) p.siteVisit++;
      if (a.field === "stage" && (nv === "won" || nv === "closed")) p.closed++;
    }
    if (a.phone) p.phones.push(String(a.phone));
  });
  Object.keys(actors).forEach(k=>{
    const p = actors[k];
    const spokenPhones = {};
    p.activities.filter(a=>a.eventType === "call" && String(a.callOutcome||"").toLowerCase() !== "no response")
      .forEach(a=>{ if (a.phone) spokenPhones[String(a.phone)] = true; });
    p.uniqueCustomersSpoken = Object.keys(spokenPhones).length;
    p.phones = Array.from(new Set(p.phones));
  });
  return {ok:true,date:d,performance:Object.keys(actors).map(k=>actors[k])};
}
