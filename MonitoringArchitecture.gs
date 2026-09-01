/**
 * Staged / isolated telecaller monitoring architecture.
 * IMPORTANT: This file is intentionally NOT wired into the current live Apps Script deployment yet.
 * It is kept separate so the presently working CRM UI/backend are not affected.
 */
const MONITORING_ACTIVITY_SHEET = "ActivityLog";
const MONITORING_ROLES = ["director","administrator","admin","manager"];
const MONITORING_ACTIVITY_HEADERS = [
  "eventId","timestamp","date","actor","actorRole","leadId","phone","name",
  "eventType","field","oldValue","newValue","callOutcome","notes"
];

function monitoringCanView_(role){
  return MONITORING_ROLES.indexOf(String(role||"").trim().toLowerCase()) > -1;
}

function monitoringGetActivitySheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MONITORING_ACTIVITY_SHEET);
  if (!sh) sh = ss.insertSheet(MONITORING_ACTIVITY_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(MONITORING_ACTIVITY_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,MONITORING_ACTIVITY_HEADERS.length).setFontWeight("bold");
  }
  return sh;
}

function monitoringAppendActivity_(ev){
  const sh = monitoringGetActivitySheet_();
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
  sh.appendRow(MONITORING_ACTIVITY_HEADERS.map(h=>row[h] != null ? row[h] : ""));
}

function monitoringRows_(){
  const sh = monitoringGetActivitySheet_();
  const rows = sh.getDataRange().getValues();
  const out=[];
  for(let i=1;i<rows.length;i++){
    if(!rows[i][0]) continue;
    const o={};
    MONITORING_ACTIVITY_HEADERS.forEach((h,j)=>{
      let v=rows[i][j];
      if(v instanceof Date){
        v=Utilities.formatDate(v,Session.getScriptTimeZone(),h==="timestamp"?"yyyy-MM-dd'T'HH:mm:ssXXX":"yyyy-MM-dd");
      }
      o[h]=v;
    });
    out.push(o);
  }
  return out;
}

function monitoringGetActivity_(dateStr, actor){
  const d=dateStr||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd");
  const activities=monitoringRows_().filter(a=>String(a.date)===d&&(!actor||String(a.actor)===String(actor)));
  activities.sort((a,b)=>String(b.timestamp).localeCompare(String(a.timestamp)));
  return {ok:true,date:d,actor:actor||"",activities};
}

function monitoringBuildPerformance_(dateStr, actor){
  const d=dateStr||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd");
  const acts=monitoringRows_().filter(a=>String(a.date)===d);
  const actors={};
  acts.forEach(a=>{
    const who=String(a.actor||"Unknown");
    if(actor&&who!==actor) return;
    if(!actors[who]) actors[who]={actor:who,date:d,callActivities:0,uniqueCustomersSpoken:0,editedLeads:0,hot:0,followUp:0,siteVisit:0,closed:0,noResponse:0,phones:[],activities:[]};
    const p=actors[who]; p.activities.push(a);
    if(a.eventType==="call"){
      p.callActivities++;
      if(String(a.callOutcome||"").toLowerCase().indexOf("no response")>-1) p.noResponse++;
    }
    if(a.eventType==="edit"){
      p.editedLeads++;
      const nv=String(a.newValue||"").toLowerCase();
      if(a.field==="stage"&&nv.indexOf("hot")>-1) p.hot++;
      if(a.field==="nextFollowUp"&&a.newValue) p.followUp++;
      if(a.field==="stage"&&nv.indexOf("site visit")>-1) p.siteVisit++;
      if(a.field==="stage"&&(nv==="won"||nv==="closed")) p.closed++;
    }
    if(a.phone) p.phones.push(String(a.phone));
  });
  Object.keys(actors).forEach(k=>{
    const p=actors[k], spoken={};
    p.activities.filter(a=>a.eventType==="call"&&String(a.callOutcome||"").toLowerCase()!=="no response").forEach(a=>{if(a.phone) spoken[String(a.phone)]=true;});
    p.uniqueCustomersSpoken=Object.keys(spoken).length;
    p.phones=Array.from(new Set(p.phones));
  });
  return {ok:true,date:d,performance:Object.keys(actors).map(k=>actors[k])};
}
