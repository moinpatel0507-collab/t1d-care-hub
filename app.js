// T1D Care Hub — app logic (generated from the master page)
// JSZip, XLSX and pdfjsLib are loaded by <script> tags in index.html

(() => {
const TYPES = {
  bg:{label:"BG check",chip:"blue"}, vial:{label:"Insulin opened",chip:"blue"}, hypo:{label:"Hypo treated",chip:"low"}, high:{label:"High / correction",chip:"high"}, ketones:{label:"Ketones",chip:"urgent"},
  site:{label:"Site / pod change",chip:"blue"}, sensor:{label:"Sensor change",chip:"blue"}, ill:{label:"Unwell",chip:"high"},
  activity:{label:"Activity",chip:""}, note:{label:"Note",chip:""},
  food:{label:"Food",chip:"food"}, drink:{label:"Drink",chip:"drink"}
};
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const num = v => (v === "" || v == null || isNaN(+v)) ? null : +v;
let db = null, base = null, plan = {}, log = [], supplies = [];
let canWrite = false, me = null, names = {}, userCap = null;

/* ---------- views ---------- */
const VIEWS = ["viewer","handover","sites","today","glucose","food","log","insulin","dayout","supplies","checks","appts","reports","sick","plan","backup","setup"];
function route(){
  const h0 = location.hash.slice(1), viewerMode = typeof isViewer==="function" && isViewer();
  const v = viewerMode ? "viewer" : (VIEWS.includes(h0) && h0!=="viewer" ? h0 : "today");
  if(typeof countView==="function" && !viewerMode) countView(v);
  VIEWS.forEach(n => $("#v-"+n).hidden = n !== v); document.documentElement.dataset.view = v;
  document.querySelectorAll("#nav a, #moreSheet a").forEach(a => a.getAttribute("href") === "#"+v ? a.setAttribute("aria-current","page") : a.removeAttribute("aria-current"));
  const secOn = !!document.querySelector('#nav a.sec[aria-current="page"]');
  $("#moreBtn").classList.toggle("on", secOn); $("#moreSheet").hidden = true; $("#moreBtn").setAttribute("aria-expanded","false");
  scrollTo({top:0}); if(v==="glucose" && typeof renderGlucose==="function") setTimeout(renderGlucose,0); if(v==="reports" && typeof renderReport==="function") setTimeout(renderReport,0);
}
var navSteps = 0;
addEventListener("hashchange", () => { navSteps++; route(); });
function goBack(){ if(navSteps > 0 && history.length > 1){ navSteps = Math.max(0, navSteps-2); history.back(); } else { location.hash = "today"; } }
route();
$("#sheetSearch").addEventListener("submit", ev => { ev.preventDefault(); const q=$("#sheetQ").value.trim(); location.hash="log"; setTimeout(()=>{ const i=$("#logSearch"); if(i){ i.value=q; renderLog(); i.focus(); } }, 0); });
$("#moreBtn").addEventListener("click", () => { const sh=$("#moreSheet"); sh.hidden=!sh.hidden; $("#moreBtn").setAttribute("aria-expanded", String(!sh.hidden)); });
$("#moreSheet").addEventListener("click", e => { if(e.target.id==="moreSheet") { e.currentTarget.hidden=true; $("#moreBtn").setAttribute("aria-expanded","false"); } });

function toast(msg){ document.querySelectorAll(".toast").forEach(x=>x.remove()); const t=document.createElement("div"); t.className="toast"; t.setAttribute("role","status"); t.textContent=msg; document.body.appendChild(t); setTimeout(()=>t.remove(),2200); }
function undoToast(msg, undo){ document.querySelectorAll(".toast").forEach(x=>x.remove()); const t=document.createElement("div"); t.className="toast"; t.setAttribute("role","status");
  t.innerHTML=`<span></span><button type="button" class="undo">Undo</button>`; t.firstChild.textContent=msg; document.body.appendChild(t);
  const timer=setTimeout(()=>t.remove(),6000);
  t.querySelector(".undo").addEventListener("click", async () => { clearTimeout(timer); t.remove(); try { await undo(); toast("Restored"); } catch { toast("Couldn't undo"); } }); }
async function deleteWithUndo(col, id, label){
  const ref = base.collection(col).doc(id); const snap = await ref.get(); const body = snap.exists ? JSON.parse(JSON.stringify(snap.data())) : null;
  await ref.delete();
  undoToast(label+" deleted", async () => { if(body) await base.collection(col).doc(id).set(body); });
}
async function copy(text, btn){
  try { await navigator.clipboard.writeText(text); toast("Copied"); }
  catch { const r=document.createRange(); r.selectNodeContents(btn.closest(".panel").querySelector("pre")); const s=getSelection(); s.removeAllRanges(); s.addRange(r); toast("Selected. Press copy on your device"); }
}

/* ---------- helpers ---------- */
const P = k => plan[k] ?? "";
const hypoBelow = () => num(plan.hypo) ?? 4;
const ketAbove = () => num(plan.ketAbove) ?? 14;
const fmtT = ms => new Date(ms).toLocaleString("en-GB",{weekday:"short",day:"numeric",month:"short",hour:"2-digit",minute:"2-digit"});
const fmtShort = ms => { const d=new Date(ms); return d.toLocaleDateString("en-GB",{day:"numeric",month:"short"})+"\n"+d.toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}); };
function ago(ms){ const h=(Date.now()-ms)/36e5; if(h<1) return Math.max(1,Math.round(h*60))+" min ago"; if(h<48) return Math.round(h)+" h ago"; return Math.round(h/24)+" days ago"; }
function glChip(g){ if(g==null) return ""; const c = g < hypoBelow() ? "low" : g > ketAbove() ? "high" : "ok"; return `<span class="chip ${c} num">${g.toFixed(1)}</span>`; }
const last = t => log.find(e => e.type === t);

function ketBand(k, g, vom){
  if (k == null) return null;
  if (k >= 3 || (vom && k >= 0.6)) return {c:"urgent", t:"Emergency: go to A&E or call 999", d:"Ketones this high, or raised ketones with vomiting, drowsiness or fast breathing, can mean DKA. Call the diabetes team on the way."};
  if (k >= 1.5) return {c:"low", t:"High risk: phone the diabetes team now", d:"Follow their instructions for extra insulin. On a pump, give the correction by pen and change the site unless told otherwise. Recheck in 1 hour."};
  if (k >= 0.6) return {c:"high", t:"Raised: act and recheck in 2 hours", d:"Give the correction in his plan" + (g!=null && g>ketAbove() ? " (glucose is high too)" : "") + ", push fluids, and check the pump site. If ketones are not falling at the recheck, phone the team."};
  return {c:"ok", t:"Normal (below 0.6)", d:"Keep checking glucose. If he is unwell, recheck ketones every 2 hours."};
}

/* ---------- render ---------- */
function renderHeroRing(){
  const el=$("#heroRing"); if(!el) return; const s=lastSite(), days=num(plan.siteDays)??3;
  const R=54, C=2*Math.PI*R; let frac=0, big="—", small="no site logged";
  if(s){ const left=(s.at+days*864e5-Date.now())/864e5; frac=Math.max(0,Math.min(1,left/days)); big = left<0 ? "Due" : left<1 ? Math.max(0,Math.round(left*24))+"h" : left.toFixed(1)+"d"; small = left<0 ? "site overdue" : "site left"; }
  const col = !s ? "var(--hero-muted)" : frac<=0 ? "#ff9b91" : frac<0.2 ? "#ffd28a" : "var(--ring)";
  el.innerHTML = `<svg viewBox="0 0 132 132" width="132" height="132"><circle cx="66" cy="66" r="${R}" fill="none" stroke="var(--hero-chip)" stroke-width="12"/><circle cx="66" cy="66" r="${R}" fill="none" stroke="${col}" stroke-width="12" stroke-linecap="round" stroke-dasharray="${(C*frac).toFixed(1)} ${C.toFixed(1)}" transform="rotate(-90 66 66)"/></svg><div class="hr-in"><b>${big}</b><span>${small}</span></div>`;
}
function ringCard(label, lastMs, days, colorVar, softVar, what){
  const R=44, C=2*Math.PI*R; let frac=0, big="—", unit="", sub=`No ${what} logged yet`, cls="";
  if(lastMs){ const left=(lastMs+days*864e5-Date.now())/864e5; frac=Math.max(0,Math.min(1,left/days));
    if(left<0){ big="Due"; unit="overdue"; cls="bad"; } else if(left<1){ big=String(Math.max(1,Math.round(left*24))); unit="hours left"; cls="warn"; } else { big=left.toFixed(1); unit="days left"; }
    sub = "Changed " + new Date(lastMs).toLocaleString("en-GB",{weekday:"long",hour:"2-digit",minute:"2-digit"}).replace(" at "," "); }
  const col = cls==="bad" ? "var(--low)" : cls==="warn" ? "var(--high)" : `var(${colorVar})`;
  return `<div class="ring-card ${cls}"><div class="rc-ring"><svg viewBox="0 0 104 104" width="104" height="104" aria-hidden="true"><circle cx="52" cy="52" r="${R}" fill="none" stroke="var(${softVar})" stroke-width="10"/><circle cx="52" cy="52" r="${R}" fill="none" stroke="${col}" stroke-width="10" stroke-linecap="round" stroke-dasharray="${C.toFixed(1)}" stroke-dashoffset="${(C*(1-frac)).toFixed(1)}" transform="rotate(-90 52 52)"/></svg><div class="rc-in"><b>${big}</b><span>${unit}</span></div></div><span class="rc-t">${label}</span><span class="rc-s">${esc(sub)}</span></div>`;
}
function renderRings(){
  const el=$("#rings"); if(!el) return; const s=lastSite(), sn=last("sensor");
  el.innerHTML = ringCard("Infusion set", s?.at, num(plan.siteDays)??3, "--accent", "--accent-soft", "change") + ringCard("Sensor", sn?.at, num(plan.sensorDays)??10, "--teal", "--teal-soft", "sensor change");
}
function renderTodayStats(){
  const el=$("#todayStats"); if(!el) return; const t=dayTotals(dayKey(Date.now())), lf=log.find(e=>e.type==="food");
  el.innerHTML = [["Carbs", t.carbs+" g"],["Drinks", t.ml+" ml"],["Last ate", lf ? new Date(lf.at).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}) : "—"]]
    .map(([k,v])=>`<div><span class="k">${k}</span><b>${v}</b></div>`).join("");
}
function renderHero(){ renderHeroRing(); try { $("#who").textContent = friendlyLine(); } catch {}
  const h=new Date().getHours(); const nm=(plan.name||"").trim();
  $("#heroTitle").textContent = nm ? `${nm}'s day` : (h<12?"Good morning":h<18?"Good afternoon":"Good evening");
  renderRings(); renderTodayStats();
  $("#heroDate").textContent = new Date().toLocaleDateString("en-GB",{weekday:"long",day:"numeric",month:"long"});
  const site=siteStatus(), t=dayTotals(dayKey(Date.now())), lf=log.find(e=>e.type==="food");
  const wk=log.filter(e=>e.type==="hypo" && e.at>Date.now()-7*864e5).length;
  const chips=[["Site",site.v.replace(" left",""),site.cls],["Carbs",t.carbs+" g",""],["Drinks",t.ml+" ml",""],["Last ate",lf?new Date(lf.at).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}):"—",""],["Hypos 7d",String(wk),wk>=3?"warn":""]];
  $("#heroChips").innerHTML = chips.map(([k,v,c])=>`<div class="hchip ${c}"><span class="k">${k}</span><span class="v">${esc(v)}</span></div>`).join("");
}
function renderAll(){ if(typeof applyRole==="function"){ applyRole(); renderMonthCard(); renderHandoverTab(); } if(typeof renderBedtime==="function"){ renderBedtime(); renderRotation(); } renderHero(); renderStrip(); renderRecent(); renderFood(); renderLog(); renderSupplies(); renderSick(); renderItems(); renderGlucose(); renderChecks(); renderBackup(); renderMealPatterns(); renderComing(); renderAppts(); renderDeliveries(); renderReport(); renderFavs(); renderWeek(); renderSetup(); renderBgNow(); renderBgChecks(); renderInsulin(); renderVial(); renderSick2(); renderDayOut(); renderWatch(); }
const dayKey = ms => new Date(ms).toDateString();
function dayTotals(key){ const es = log.filter(e => dayKey(e.at)===key && (e.type==="food"||e.type==="drink"));
  return { carbs: es.reduce((a,e)=>a+(e.carbs||0),0), ml: es.filter(e=>e.type==="drink").reduce((a,e)=>a+(e.ml||0),0), meals: es.filter(e=>e.type==="food").length, drinks: es.filter(e=>e.type==="drink").length, es }; }
function renderFood(){
  const t = dayTotals(dayKey(Date.now()));
  const lastFood = log.find(e=>e.type==="food"), lastDrink = log.find(e=>e.type==="drink");
  const prev = [1,2,3,4,5,6,7].map(i=>dayTotals(dayKey(Date.now()-i*864e5))).filter(d=>d.meals);
  const avg = prev.length ? Math.round(prev.reduce((a,d)=>a+d.carbs,0)/prev.length) : null;
  $("#foodStrip").innerHTML = [
    `<div class="stat"><span class="k">Carbs today</span><span class="v">${t.carbs} g</span><span class="n">${avg!=null?`7-day average ${avg} g`:"Average shows after a few days"}</span></div>`,
    `<div class="stat"><span class="k">Drinks today</span><span class="v">${t.ml} ml</span><span class="n">${t.drinks} drink${t.drinks===1?"":"s"} logged</span></div>`,
    `<div class="stat"><span class="k">Last ate</span><span class="v">${lastFood?new Date(lastFood.at).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}):"—"}</span><span class="n">${lastFood?esc((lastFood.meal||"")+" · "+ago(lastFood.at)):"Nothing logged"}</span></div>`,
    `<div class="stat"><span class="k">Last drink</span><span class="v">${lastDrink?new Date(lastDrink.at).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}):"—"}</span><span class="n">${lastDrink?ago(lastDrink.at):"Nothing logged"}</span></div>`
  ].join("");
  $("#foodToday").innerHTML = t.es.length ? t.es.slice().reverse().map(e => `<div class="meal tap" data-ledit="${esc(e.id)}" role="button" tabindex="0" aria-label="Edit ${esc(e.item||TYPES[e.type].label)}">
      <span class="t">${new Date(e.at).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}</span>
      <div style="min-width:0;overflow-wrap:anywhere"><span class="chip ${TYPES[e.type].chip}">${e.type==="food"?esc(e.meal||"Food"):"Drink"}</span>${esc(e.item||"")}${e.type==="food"&&e.ate&&e.ate!=="All"?` <span class="muted">· ate ${esc(e.ate.toLowerCase())}</span>`:""}${e.photo?` <button type="button" class="photo-chip" data-photo="${esc(e.id)}" aria-label="See photo">${'<svg class="ic" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/></svg>'}</button>`:""}${e.note?`<div class="muted">${esc(e.note)}</div>`:""}</div>
      <span class="c">${e.carbs!=null?e.carbs+" g":""}${e.type==="drink"&&e.ml!=null?`<div class="muted" style="font-size:.85rem">${e.ml} ml</div>`:""}</span></div>`).join("")
    : `<div class="empty">Nothing logged today. Tap Add food after his next meal or snack.</div>`;
  $("#foodWeek").innerHTML = [0,1,2,3,4,5,6].map(i => { const ms=Date.now()-i*864e5, d=dayTotals(dayKey(ms));
    return `<tr class="dayrow"><td>${i===0?"Today":i===1?"Yesterday":new Date(ms).toLocaleDateString("en-GB",{weekday:"short",day:"numeric",month:"short"})}</td><td class="muted">${d.meals||"—"}</td><td class="num">${d.meals?d.carbs+" g":"—"}</td><td class="num">${d.drinks?d.ml+" ml":"—"}</td></tr>`; }).join("");
}
function renderItems(){
  const seen = new Set(), opts = [];
  for (const f of favSort()) { seen.add(f.name.toLowerCase()); opts.push(f.name); }
  for (const e of log) if ((e.type==="food"||e.type==="drink") && e.item && !seen.has(e.item.toLowerCase())) { seen.add(e.item.toLowerCase()); opts.push(e.item); if(opts.length>=80) break; }
  $("#itemList").innerHTML = opts.map(o=>`<option value="${esc(o)}"></option>`).join("");
}

function lastSite(){ const l = last("site"), p = (typeof lastPumpSite==="function") ? lastPumpSite() : null; if(p && (!l || p > l.at)) return {at:p, fromPump:true}; return l; }
function siteStatus(){
  const s = lastSite(); const days = num(plan.siteDays) ?? 3;
  if(!s) return {v:"—", n:"No change logged", cls:""};
  const due = s.at + days*864e5, left = (due-Date.now())/36e5;
  return {v: left<0 ? "Overdue" : left<24 ? Math.round(left)+" h left" : (left/24).toFixed(1)+" d left", n:"Changed "+ago(s.at), cls: left<0 ? "bad" : left<12 ? "warn" : ""};
}
function sensorStatus(){
  const s = last("sensor"); const days = num(plan.sensorDays) ?? 10;
  if(!s) return {v:"—", n:"No change logged", cls:""};
  const left = (s.at + days*864e5 - Date.now())/36e5;
  return {v: left<0 ? "Overdue" : left<24 ? Math.round(left)+" h left" : (left/24).toFixed(1)+" d left", n:"Changed "+ago(s.at), cls: left<0 ? "bad" : left<24 ? "warn" : ""};
}
function supplyRows(){ return supplies.map(s => { const d = (s.qty||0)*(s.days||0); const lead=s.lead??14; return {...s, left:d, status: d<=lead ? "Reorder now" : d<=lead+7 ? "Order soon" : "OK"}; }).sort((a,b)=>(a.name||"").localeCompare(b.name||"")); }  // stable order so + / − never jumps rows

function renderStrip(){
  const site=siteStatus(), sen=sensorStatus();
  const wk = log.filter(e => e.type==="hypo" && e.at > Date.now()-7*864e5).length;
  const low = supplyRows().filter(s => s.status !== "OK").sort((a,b)=>a.left-b.left);
  const lk = log.find(e => e.ket != null);
  { const b=$("#bdgSupplies"); if(b){ b.hidden=!low.length; b.textContent = low.length ? low.length+" to order" : ""; } }
  $("#strip").innerHTML = [
    `<div class="stat ${wk>=3?"warn":""}"><span class="k">Hypos · 7 days</span><span class="v">${wk}</span><span class="n">${wk>=3?"Worth raising with the team":"Logged by you"}</span></div>`,
    `<div class="stat ${low.some(s=>s.status==="Reorder now")?"bad":low.length?"warn":""}"><span class="k">Supplies</span><span class="v">${supplies.length? (low.length? low.length+" to order":"All OK") : "—"}</span><span class="n">${low.length? esc(low.map(s=>s.name).slice(0,2).join(", ")) : supplies.length? "Nothing low" : "Add items on Supplies"}</span></div>`,
    ...["eye","dental"].map(k=>{ const d=dueInfo(k); return d.days!=null && d.days<=30 ? `<div class="stat ${d.cls}"><span class="k">${CK[k].label}</span><span class="v">${d.v}</span><span class="n">${d.n}</span></div>` : ""; }),
    (log.length>20 && (!lastBackup || Date.now()-lastBackup>30*864e5)) ? `<a class="stat warn" href="#backup" style="text-decoration:none;color:inherit"><span class="k">Backup</span><span class="v">${lastBackup?Math.floor((Date.now()-lastBackup)/864e5)+" days":"Never"}</span><span class="n">Tap to download a backup</span></a>` : "",
    `<div class="stat"><span class="k">Last ketones</span><span class="v">${lk? lk.ket.toFixed(1):"—"}</span><span class="n">${lk? ago(lk.at):"None logged"}</span></div>`
  ].join("");
}

function entryHTML(e){
  const t = TYPES[e.type] || TYPES.note;
  const bits = [];
  if(e.gl!=null) bits.push(glChip(e.gl)+'<span class="muted">mmol/L</span>');
  if(e.ket!=null){ const b=ketBand(e.ket); bits.push(`<span class="chip ${b.c} num">ket ${e.ket.toFixed(1)}</span>`); }
  if(e.type==="food"||e.type==="drink"){ bits.push(`<strong>${esc(e.meal && e.type==="food" ? e.meal+": " : "")}${esc(e.item||"")}</strong>`); if(e.ate && e.ate!=="All") bits.push(`<span class="muted">ate ${esc(e.ate.toLowerCase())}</span>`); if(e.ml!=null) bits.push(`<span class="num">${e.ml} ml</span>`); }
  if(e.carbs!=null) bits.push(`<span class="num">${e.carbs} g</span> <span class="muted">carbs</span>`);
  if(e.photo) bits.push(`<button type="button" class="photo-chip" data-photo="${esc(e.id)}" aria-label="See photo">${'<svg class="ic" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/></svg>'}</button>`);
  return `<li><time datetime="${new Date(e.at).toISOString()}" style="white-space:pre-line">${fmtShort(e.at)}</time>
    <div class="what"><span class="chip ${t.chip}">${t.label}</span>${bits.join(" ")}${e.note? `<div>${esc(e.note)}</div>`:""}${e.by? `<div class="muted" style="font-size:.8rem">Logged by ${e.by===me?"you":esc(names[e.by]||"his mum")}</div>`:""}</div>
    <div class="row-acts"><button class="small" data-ledit="${esc(e.id)}" ${canWrite?"":"disabled"}>Edit</button><button class="small danger" data-del="${esc(e.id)}" ${canWrite?"":"disabled"}>Delete</button></div></li>`;
}
function renderRecent(){
  const r = log.slice(0,3);
  $("#recent").innerHTML = r.length ? r.map(entryHTML).join("") : `<li style="display:block"><div class="empty">Nothing logged yet. Use a quick-log button above. Start with his last pod and sensor change so the countdowns work.</div></li>`;
}
function entryText(e){ return [TYPES[e.type]?.label, e.meal, e.item, e.note, e.ate, e.gl!=null?e.gl.toFixed(1):"", e.ket!=null?e.ket.toFixed(1):"", e.carbs!=null?e.carbs+" g":"", e.ml!=null?e.ml+" ml":"", new Date(e.at).toLocaleDateString("en-GB",{weekday:"long",day:"numeric",month:"long"})].filter(Boolean).join(" ").toLowerCase(); }
function rangeEntries(){
  const days=+$("#range").value, f=$("#filter").value, q=($("#logSearch")?.value||"").trim().toLowerCase();
  if(q){ const words=q.split(/\s+/); return log.filter(e => (!f || e.type===f) && words.every(w => entryText(e).includes(w))); }   // search ignores the date range
  return log.filter(e => e.at > Date.now()-days*864e5 && (!f || e.type===f)); }
function renderLog(){
  const r = rangeEntries();
  const q=($("#logSearch")?.value||"").trim();
  $("#searchCount").textContent = q ? `${r.length} match${r.length===1?"":"es"}${r.length>200?" · showing 200":""}` : "";
  $("#fullLog").innerHTML = r.length ? r.slice(0,200).map(entryHTML).join("") : `<li style="display:block"><div class="empty">${q?`Nothing matches “${esc(q)}”.`:"No entries in this range."}</div></li>`;
  const days=+$("#range").value, all = log.filter(e => e.at > Date.now()-days*864e5);
  const cnt = t => all.filter(e=>e.type===t).length;
  const hypos = all.filter(e=>e.type==="hypo");
  const hours = {night:0,morning:0,afternoon:0,evening:0};
  hypos.forEach(e=>{const h=new Date(e.at).getHours(); hours[h<6?"night":h<12?"morning":h<18?"afternoon":"evening"]++;});
  const kets = all.filter(e=>e.ket!=null && e.ket>=0.6);
  const notes = all.filter(e=>e.type==="note"||e.type==="ill");
  const name = P("name") || "My son";
  const lines = [
    `${name} — events from the last ${days} days (to ${new Date().toLocaleDateString("en-GB")})`,
    `Setup: ${[P("cgm"),P("pump")].filter(Boolean).join(" + ") || "CGM + pump"}`,
    ``,
    `Hypos treated: ${hypos.length}` + (hypos.length? ` (night ${hours.night}, morning ${hours.morning}, afternoon ${hours.afternoon}, evening ${hours.evening})` : ""),
    `Highs needing action: ${cnt("high")}`,
    `Raised ketones (0.6+): ${kets.length}` + (kets.length? ` — highest ${Math.max(...kets.map(e=>e.ket)).toFixed(1)}`:""),
    `Site/pod changes: ${cnt("site")} · Sensor changes: ${cnt("sensor")}`,
    `Unwell days logged: ${cnt("ill")}`,
    ...(()=>{ const ev=rangePump(days), n=k=>ev.filter(e=>e.k===k), dd=Math.max(1,new Set(ev.map(e=>ymd(e.t))).size), out=[];
      if(n(1).length) out.push(`Carbs entered in pump: about ${Math.round(n(1).reduce((a,e)=>a+e.v,0)/dd)} g a day`);
      if(n(2).length) out.push(`Bolus insulin: about ${(n(2).reduce((a,e)=>a+e.v,0)/dd).toFixed(1)} U a day (${n(2).length} doses)`);
      if(n(3).length) out.push(`Finger pricks: ${n(3).length}`);
      const eps = hypoEpisodes(rangePoints(Math.min(days,30))); if(eps.length) out.push(`CGM hypos (below 3.9 for 15+ min): ${eps.length}` + (eps.length? ` — lowest ${Math.min(...eps.map(h=>h.min)).toFixed(1)}`:""));
      return out; })(),
  ];
  const cst = cgmStats(Math.min(days,30));
  if(cst) lines.splice(3,0,`CGM (last ${Math.min(days,30)} days): in range ${Math.round(pct(cst.b.in,cst.n))}%, low ${pct(cst.b.low+cst.b.vlow,cst.n).toFixed(1)}% (very low ${pct(cst.b.vlow,cst.n).toFixed(1)}%), high ${Math.round(pct(cst.b.high+cst.b.vhigh,cst.n))}% · average ${cst.mean.toFixed(1)} mmol/L · GMI ${cst.gmiMmol} mmol/mol · CV ${Math.round(cst.cv)}% · sensor data ${Math.round(cst.coverage)}%`,``);
  const fdays = [...new Set(all.filter(e=>e.type==="food").map(e=>dayKey(e.at)))].map(dayTotals);
  if(fdays.length){ lines.push(`Food logged on ${fdays.length} day${fdays.length===1?"":"s"}: average ${Math.round(fdays.reduce((a,d)=>a+d.carbs,0)/fdays.length)} g carbs/day, ${Math.round(fdays.reduce((a,d)=>a+d.ml,0)/fdays.length)} ml drinks/day`);
    const refused = all.filter(e=>e.type==="food" && (e.ate==="Refused"||e.ate==="A little")).length; if(refused) lines.push(`Meals refused or barely eaten: ${refused}`); }
  { const H=latestWith("h"), W=latestWith("w"); if(H.length||W.length) lines.push("", `Latest height: ${H.length?H[0].h.toFixed(1)+" cm ("+fmtD(H[0].at)+")":"not recorded"} · Latest weight: ${W.length?W[0].w.toFixed(2)+" kg ("+fmtD(W[0].at)+")":"not recorded"}`);
    const le=lastCheck("eye"), ld=lastCheck("dental"); if(le||ld) lines.push(`Last eye check: ${le?fmtD(le.at):"none logged"} · Last dental check: ${ld?fmtD(ld.at):"none logged"}`); }
  if(notes.length){ lines.push("", "Notes and questions:"); notes.slice(0,15).forEach(e=>lines.push(`- ${new Date(e.at).toLocaleDateString("en-GB")}: ${e.note||TYPES[e.type].label}`)); }
  if(hypos.length){ lines.push("", "Hypo details:"); hypos.slice(0,15).forEach(e=>lines.push(`- ${fmtT(e.at)}${e.gl!=null?` · ${e.gl.toFixed(1)} mmol/L`:""}${e.carbs!=null?` · ${e.carbs} g`:""}${e.note?` · ${e.note}`:""}`)); }
  $("#clinicText").textContent = lines.join("\n");
}
function renderSupplies(){
  const rows = supplyRows();
  $("#supTable").innerHTML = rows.length ? rows.map(s => `<tr>
    <td><strong>${esc(s.name)}</strong>${s.link?`<div class="muted" style="font-size:.8rem">Counts down on ${s.link==="site"?"site change":"sensor change"}</div>`:""}</td>
    <td class="num"><div style="display:flex;gap:6px;align-items:center;justify-content:flex-end"><button class="small" data-dec="${esc(s.id)}" aria-label="One fewer" ${canWrite?"":"disabled"}>−</button><span style="min-width:2ch">${s.qty}</span><button class="small" data-inc="${esc(s.id)}" aria-label="One more" ${canWrite?"":"disabled"}>+</button></div></td>
    <td class="num">${Math.floor(s.left)}</td>
    <td>${untilDelivery(s)}</td>
    <td>${(()=>{ const nd=nextDelivery(); const short = nd && s.days && Math.floor(s.left) < daysTo(nd);
      const st = short && s.status==="OK" ? "Get extra" : s.status; return `<span class="chip ${st==="OK"?"ok":st==="Order soon"?"high":"low"}">${st}</span>`; })()}</td>
    <td style="text-align:right;white-space:nowrap"><button class="small" data-edit="${esc(s.id)}" ${canWrite?"":"disabled"}>Edit</button> <button class="small danger" data-sdel="${esc(s.id)}" ${canWrite?"":"disabled"}>Delete</button></td></tr>`).join("")
    : `<tr><td colspan="6"><div class="empty">No supplies yet. Add the common items below, then set the real stock you have at home.</div></td></tr>`;
}
function renderSick(){
  const b = ketBand(num($("#ketIn").value), num($("#glIn").value), $("#vomIn").checked);
  $("#ketOut").innerHTML = b ? `<div class="band ${b.c}"><strong>${b.t}</strong><span>${b.d}</span></div>` : `<div class="empty">Enter a blood ketone reading to see what to do.</div>`;
  const ft = dayTotals(dayKey(Date.now()));
  $("#ketOut").insertAdjacentHTML("beforeend", `<p class="muted" style="margin:8px 0 0">Drinks logged today: <strong class="num">${ft.ml} ml</strong>${ft.drinks?` (last ${ago(log.find(e=>e.type==="drink").at)})`:""}</p>`);
  const c = [["Diabetes team (day)",P("teamDay")],["Out of hours",P("teamOoh")]].filter(x=>x[1]);
  $("#sickContacts").innerHTML = c.length ? `<div class="grid2">${c.map(x=>`<div class="stat"><span class="k">${x[0]}</span><span class="v" style="font-size:1.05rem;user-select:all">${esc(x[1])}</span></div>`).join("")}</div>` : `<p class="notice">Add the diabetes team's numbers on Care plan so they show here.</p>`;
}

/* ---------- plan form ---------- */
const PLAN_KEYS = ["name","dob","cgm","pump","hypo","ketAbove","siteDays","sensorDays","vialDays","eyeMonths","dentalMonths","hypoTx","highTx","signs","glucagon","teamDay","teamOoh","notes","hoContacts","hoKit","hoPump"];
let planDirty = false;
function fillPlan(){ if(planDirty) return planUi(); PLAN_KEYS.forEach(k => { const el=$("#p_"+k); if(el) el.value = plan[k] ?? ""; }); planUi(); }
function planUi(){ const st=$("#planState"); if(st){ st.textContent = planDirty ? "Unsaved changes" : "All saved"; st.classList.toggle("dirty", planDirty); st.parentElement.classList.toggle("dirty", planDirty); }
  document.querySelectorAll("[data-call]").forEach(a => { const v=($("#"+a.dataset.call).value||"").replace(/[^0-9+]/g,""); a.hidden = v.length<6; if(v) a.href="tel:"+v; }); }
$("#planForm").addEventListener("input", () => { planDirty = true; planUi(); });
$("#planForm").addEventListener("submit", async ev => {
  ev.preventDefault(); if(!base) return;
  const data = {}; PLAN_KEYS.forEach(k => data[k] = $("#p_"+k).value.trim());
  $("#savePlan").disabled = true;
  try { await base.set({...data, updatedAt: Date.now()}); planDirty=false; plan=data; renderAll(); planUi(); toast("Care plan saved"); }
  catch(e){ toast(e.code==="quota_exceeded" ? "Storage is full" : `Couldn't save (${e.code||e.message||"error"}). Try again`); }
  finally { $("#savePlan").disabled = false; }
});

/* ---------- entries ---------- */
const dlg = $("#dlg");
$("#e_type").innerHTML = Object.entries(TYPES).map(([k,v])=>`<option value="${k}">${v.label}</option>`).join("");
$("#filter").innerHTML += Object.entries(TYPES).map(([k,v])=>`<option value="${k}">${v.label}</option>`).join("");
function localNow(){ const d=new Date(); d.setMinutes(d.getMinutes()-d.getTimezoneOffset()); return d.toISOString().slice(0,16); }
/* ---------- hypo mode: treat, 15-minute recheck, repeat ---------- */
const HY_KEY = "t1d-hypo-timer", HY_MIN = 15;
let hy = null, hyTick = null, hyAudio = null;
function hyLoad(){ try { hy = JSON.parse(localStorage.getItem(HY_KEY)||"null"); } catch { hy = null; } if(hy && Date.now()-hy.start > 3*3600e3) hy=null; }
function hySave(){ try { hy ? localStorage.setItem(HY_KEY, JSON.stringify(hy)) : localStorage.removeItem(HY_KEY); } catch {} }
const hyLeft = () => hy ? Math.max(0, hy.start + HY_MIN*6e4 - Date.now()) : 0;
const mmss = ms => { const s=Math.ceil(ms/1000); return Math.floor(s/60)+":"+String(s%60).padStart(2,"0"); };
function hyBeep(){ try { hyAudio ??= new (window.AudioContext||window.webkitAudioContext)(); const a=hyAudio; [0,0.35,0.7].forEach(t=>{ const o=a.createOscillator(), g=a.createGain(); o.frequency.value=880; g.gain.setValueAtTime(0.0001,a.currentTime+t); g.gain.exponentialRampToValueAtTime(0.4,a.currentTime+t+0.02); g.gain.exponentialRampToValueAtTime(0.0001,a.currentTime+t+0.28); o.connect(g).connect(a.destination); o.start(a.currentTime+t); o.stop(a.currentTime+t+0.3); }); } catch {} }
async function hyNotify(){
  try { navigator.vibrate?.([400,200,400,200,400]); } catch {}
  hyBeep();
  if(!("Notification" in window) || Notification.permission!=="granted") return;
  const o = {body:"15 minutes are up. Check his glucose now.", tag:"t1d-hypo", renotify:true, requireInteraction:true, icon:"/icon-192.png"};
  try { const r = await navigator.serviceWorker?.getRegistration?.(); if(r){ await r.showNotification("Recheck his glucose", o); return; } } catch {}
  try { new Notification("Recheck his glucose", o); } catch {}
}
function hyBand(g){ const lo=hypoBelow(); if(g==null) return null; if(g<3) return {c:"low", t:"Very low"}; if(g<lo) return {c:"low", t:"Low"}; return {c:"ok", t:"Back above "+lo.toFixed(1)}; }
function hypoState(){
  const go=$("#hypoGo"); if(!go) return; const g=latestGlucose();
  const lowNow = !!(g && g.v<hypoBelow() && Date.now()-g.t < 60*6e4);
  const urgent = lowNow || !!hy;
  go.classList.toggle("quiet", !urgent);
  go.querySelector("small").textContent = urgent ? (hy ? "Recheck timer running" : `Latest reading ${g.v.toFixed(1)}: treatment steps and a 15-minute recheck`) : "Step-by-step help and a 15-minute recheck timer";
  go.querySelector("b").textContent = urgent ? "He's low" : "He's low?";
}
function renderHypoBanner(){ hypoState();
  const b=$("#hypoBanner"); if(!b) return;
  if(!hy){ b.hidden=true; return; }
  const left=hyLeft(); b.hidden=false;
  b.innerHTML = left>0 ? `<span><b>Hypo recheck in ${mmss(left)}</b> · treated at ${new Date(hy.start).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}</span><button type="button" class="small" data-hyopen>Open</button>`
                       : `<span><b>Time to recheck his glucose</b></span><button type="button" class="small primary" data-hyopen>Recheck now</button>`;
}
function hyStep(step){
  const lo=hypoBelow(), tx=P("hypoTx"), signs=P("signs"), gluc=P("glucagon"), body=$("#hyBody");
  const txBox = tx ? `<div class="hy-tx"><span class="k">His treatment (from his care plan)</span><div>${esc(tx).replace(/\n/g,"<br>")}</div></div>`
                   : `<div class="hy-tx empty-tx"><span class="k">His treatment</span><div>Add his diabetes team's hypo treatment on <a href="#plan" data-hyclose>Care plan</a> so it shows here.</div></div>`;
  const emerg = `<div class="hy-emerg"><b>If he's very drowsy, fitting or can't swallow:</b> don't give anything by mouth. ${gluc?`Glucagon: ${esc(gluc)}. `:""}Call 999.</div>`;
  if(step==="start"){
    body.innerHTML = `<label class="hy-big">Glucose now <small>mmol/L, from CGM or a finger prick</small><input id="hy_gl" type="number" step="0.1" min="0" inputmode="decimal" placeholder="e.g. 3.4"></label>
      <div id="hy_band"></div>${txBox}
      <label>Carbs given <small>g</small><input id="hy_c" type="number" min="0" step="1" inputmode="numeric"></label>
      <label>What you gave <small>optional</small><input id="hy_what" placeholder="e.g. juice carton"></label>
      ${emerg}
      <button type="button" class="primary hy-cta" id="hy_go">Treated · start 15-minute timer</button>`;
    const gi=$("#hy_gl"); gi.addEventListener("input", () => { const b=hyBand(num(gi.value)); $("#hy_band").innerHTML = b?`<span class="chip ${b.c}">${b.t}</span>`:""; });
    $("#hy_go").addEventListener("click", hyStart); setTimeout(()=>gi.focus(),50);
  } else if(step==="wait"){
    body.innerHTML = `<div class="hy-ring"><svg viewBox="0 0 120 120"><circle cx="60" cy="60" r="52" class="hr-bg"/><circle cx="60" cy="60" r="52" class="hr-fg" id="hy_arc"/></svg><div class="hy-time"><b id="hy_left">${mmss(hyLeft())}</b><span>until recheck</span></div></div>
      <p class="muted" style="text-align:center;margin:0">Keep him calm and stay with him. ${hy.round>1?`This is treatment ${hy.round}.`:""}</p>
      ${txBox}
      <button type="button" class="primary hy-cta" id="hy_now">Recheck now</button>
      <p class="hy-note">The phone beeps and buzzes when time is up while the app is open. If you switch apps or lock the phone, set a 15-minute timer on the phone as well.</p>`;
    $("#hy_now").addEventListener("click", () => hyStep("recheck"));
  } else if(step==="recheck"){
    body.innerHTML = `<label class="hy-big">Glucose now <small>mmol/L</small><input id="hy_gl2" type="number" step="0.1" min="0" inputmode="decimal"></label>
      <div id="hy_res"></div>`;
    const gi=$("#hy_gl2"); setTimeout(()=>gi.focus(),50);
    gi.addEventListener("input", () => { const g=num(gi.value), b=hyBand(g); const r=$("#hy_res");
      if(g==null){ r.innerHTML=""; return; }
      r.innerHTML = g<lo ? `<div class="band low"><strong>Still low (${g.toFixed(1)})</strong><span>Treat again as his plan says, then start another 15 minutes.</span></div>${txBox}
          <label>Carbs given this time <small>g</small><input id="hy_c2" type="number" min="0" step="1" inputmode="numeric"></label>
          <button type="button" class="primary hy-cta" id="hy_again">Treated again · start 15 minutes</button>${hy.round>=2?`<div class="hy-emerg"><b>Still low after ${hy.round} treatments?</b> Follow his plan and call his diabetes team${P("teamOoh")?` (out of hours: ${esc(P("teamOoh"))})`:""}.</div>`:""}`
        : `<div class="band ok"><strong>${b.t} (${g.toFixed(1)})</strong><span>Follow his plan for what to do next, e.g. a snack if his next meal isn't soon.</span></div>
          <button type="button" class="primary hy-cta" id="hy_done">Done · save to his log</button>`;
      $("#hy_again")?.addEventListener("click", () => hyAgain(g));
      $("#hy_done")?.addEventListener("click", () => hyDone(g));
    });
  }
}
async function hyStart(){
  const g=num($("#hy_gl").value), c=num($("#hy_c").value), what=$("#hy_what").value.trim();
  try { if("Notification" in window && Notification.permission==="default") await Notification.requestPermission(); } catch {}
  try { hyAudio ??= new (window.AudioContext||window.webkitAudioContext)(); hyAudio.resume?.(); } catch {}
  let id=null;
  try { const r = await base.collection("log").add({type:"hypo", at:Date.now(), gl:g, ket:null, carbs:c, note:["Hypo mode", what].filter(Boolean).join(": "), by:me}); id=r.id; } catch(e){ toast("Couldn't save to the log, but the timer is running"); }
  hy = {start:Date.now(), entryId:id, gl:g, carbs:c||0, round:1, readings:[g]}; hySave(); hyRun(); hyStep("wait");
}
async function hyAgain(g){
  const c=num($("#hy_c2")?.value);
  hy.round++; hy.start=Date.now(); hy.alerted=false; hy.carbs=(hy.carbs||0)+(c||0); hy.readings.push(g); hySave();
  await hyUpdateEntry(`Recheck ${g.toFixed(1)} still low, treated again${c?` (${c} g)`:""}`); hyRun(); hyStep("wait");
}
async function hyDone(g){
  hy.readings.push(g); await hyUpdateEntry(`Recheck ${g.toFixed(1)}, back up`);
  hy=null; hySave(); clearInterval(hyTick); $("#hypoDlg").close(); renderHypoBanner(); toast("Hypo saved to his log");
}
async function hyUpdateEntry(add){
  if(!hy?.entryId) return; const e=log.find(x=>x.id===hy.entryId); if(!e) return;
  const u={...e, carbs:hy.carbs||e.carbs, note:[e.note, add].filter(Boolean).join(" · ")}; delete u.id;
  try { await base.collection("log").doc(hy.entryId).set(u); } catch {}
}
function hyRun(){
  clearInterval(hyTick); let fired = false;
  hyTick = setInterval(() => {
    const left=hyLeft(); const el=$("#hy_left"); if(el) el.textContent=mmss(left);
    const arc=$("#hy_arc"); if(arc){ const L=2*Math.PI*52; arc.style.strokeDasharray=L; arc.style.strokeDashoffset = L*(1-left/(HY_MIN*6e4)); }
    renderHypoBanner();
    if(left===0 && !fired){ fired=true; if(!hy.alerted){ hy.alerted=true; hySave(); hyNotify(); } if($("#hypoDlg").open && $("#hy_left")) hyStep("recheck"); }
  }, 500);
}
function openHypo(){
  if(!canWrite) return toast("Saving isn't switched on here");
  hyLoad(); $("#hypoDlg").showModal();
  if(!hy) hyStep("start"); else if(hyLeft()>0) hyStep("wait"); else hyStep("recheck");
}
$("#hypoGo").addEventListener("click", openHypo);
$("#hyClose").addEventListener("click", () => $("#hypoDlg").close());
document.addEventListener("click", ev => { if(ev.target.closest("[data-hyopen]")) openHypo(); if(ev.target.closest("[data-hyclose]")) $("#hypoDlg").close(); });
document.addEventListener("visibilitychange", () => { if(!document.hidden){ hyLoad(); if(hy) hyRun(); renderHypoBanner(); } });
hyLoad(); if(hy) hyRun(); setTimeout(renderHypoBanner, 0);

/* ---------- favourite foods ---------- */
let foods = [], favManage = false;
const favSort = () => foods.slice().sort((a,b)=>(b.uses||0)-(a.uses||0) || a.name.localeCompare(b.name));
const findFav = name => foods.find(f => f.name.trim().toLowerCase() === String(name||"").trim().toLowerCase());
function favBtn(f, cls=""){ return `<button type="button" class="fav ${cls} ${f.kind==="drink"?"is-drink":""}" data-fav="${esc(f.id)}"><span class="fav-n">${esc(f.name)}</span><span class="fav-c">${f.carbs!=null?f.carbs+" g":""}${f.kind==="drink"&&f.ml?` · ${f.ml} ml`:""}</span>${favManage&&!cls?'<span class="fav-ed">Edit</span>':""}</button>`; }
function renderFavs(){
  const L = favSort();
  $("#favList").innerHTML = L.length ? L.map(f=>favBtn(f)).join("") : `<div class="empty">No usual foods yet. Tap <b>Add a food</b>, or tick "Add to his usual foods" when you log a meal.</div>`;
  $("#favList").classList.toggle("managing", favManage);
  $("#favManage").textContent = favManage ? "Done" : "Edit list"; $("#favManage").setAttribute("aria-pressed", favManage); $("#favManage").hidden = !L.length;
  const q = $("#favQuick"); const top = L.slice(0,4); q.hidden = !top.length;
  q.innerHTML = top.length ? `<div class="fav-q-h">His usual foods</div><div class="favs">${top.map(f=>favBtn(f,"mini")).join("")}</div>` : "";
}
function openFavDlg(f){
  $("#favForm").reset(); $("#favForm").dataset.id = f?.id || "";
  $("#favTitle").textContent = f ? "Edit food" : "Add a food"; $("#favDel").hidden = !f;
  if(f){ $("#fv_name").value=f.name; $("#fv_kind").value=f.kind||"food"; $("#fv_carbs").value=f.carbs??""; $("#fv_portion").value=f.portion||""; $("#fv_ml").value=f.ml??""; }
  favKind(); $("#favDlg").showModal();
}
function favKind(){ document.querySelectorAll("[data-fk]").forEach(el=>el.hidden = el.dataset.fk!==$("#fv_kind").value); }
$("#fv_kind").addEventListener("change", favKind);
$("#favAdd").addEventListener("click", () => canWrite ? openFavDlg(null) : toast("Saving isn't switched on here"));
$("#favManage").addEventListener("click", () => { favManage=!favManage; renderFavs(); });
$("#favCancel").addEventListener("click", () => $("#favDlg").close());
$("#favForm").addEventListener("submit", async ev => {
  ev.preventDefault(); const id=$("#favForm").dataset.id, old=foods.find(f=>f.id===id)||{};
  const f = {...old, name:$("#fv_name").value.trim(), kind:$("#fv_kind").value, carbs:num($("#fv_carbs").value), portion:$("#fv_portion").value.trim(), ml:$("#fv_kind").value==="drink"?num($("#fv_ml").value):null};
  delete f.id; if(!f.name) return;
  const dup = findFav(f.name); if(dup && dup.id!==id) return toast("That food is already in the list");
  try { if(id) await base.collection("foods").doc(id).set(f); else await base.collection("foods").add({...f, uses:0, lastUsed:null}); $("#favDlg").close(); toast(id?"Food updated":"Added to his usual foods"); }
  catch(e){ toast(`Couldn't save (${e.code||e.message||"error"})`); }
});
$("#favDel").addEventListener("click", async () => { const id=$("#favForm").dataset.id; $("#favDlg").close(); try { await deleteWithUndo("foods", id, "Food"); } catch { toast("Couldn't delete"); } });
function logFav(f){
  if(favManage) return openFavDlg(f);
  openEntry(f.kind==="drink"?"drink":"food");
  $("#e_item").value = f.name; if(f.carbs!=null) $("#e_carbs").value = f.carbs; if(f.kind==="drink" && f.ml!=null) $("#e_ml").value = f.ml;
  if(f.portion) $("#e_note").placeholder = "Usual portion: " + f.portion;
  favState();
}
function favState(){ const t=$("#e_type").value, fd=(t==="food"||t==="drink") && !editingId; const v=$("#e_item").value.trim();
  $("#e_favWrap").hidden = !fd || !v || !!findFav(v); }
$("#e_item").addEventListener("input", favState);
let editingId = null;
function toLocalInput(ms){ const d=new Date(ms); d.setMinutes(d.getMinutes()-d.getTimezoneOffset()); return d.toISOString().slice(0,16); }
function openEdit(id){
  const e = log.find(x=>x.id===id); if(!e) return toast("That entry has been deleted");
  openEntry(e.type);
  editingId = id; $("#dlgTitle").textContent = "Edit " + TYPES[e.type].label.toLowerCase(); $("#dlgDelete").hidden = false;
  $("#e_at").value = toLocalInput(e.at);
  const put = (sel,v) => { if(v!=null && v!=="") $(sel).value = v; };
  if(e.photo) loadPhoto(id).then(src => { if(editingId===id && photoData===undefined) showPhotoPrev(src); });
  put("#e_gl",e.gl); put("#e_ket",e.ket); put("#e_carbs",e.carbs); put("#e_item",e.item); put("#e_meal",e.meal); put("#e_ate",e.ate); put("#e_ml",e.ml); put("#e_note",e.note);
  if(e.spot){ $("#e_spot").value = e.spot; renderSpotPicker(); }
}
function openEntry(type, pre={}){
  editingId = null; $("#dlgDelete").hidden = true;
  if(!canWrite){ (toast("Saving isn't switched on here. Open the page in a browser (see the note at the top)"), $("#dbNote").scrollIntoView({behavior:"smooth",block:"center"})); return; }
  $("#entryForm").reset(); $("#e_spot").value = ""; $("#e_type").value = type; $("#e_at").value = localNow();
  $("#dlgTitle").textContent = TYPES[type].label; setFields(type); $("#e_favWrap").hidden = true; $("#e_note").placeholder = "";
  photoData = undefined; showPhotoPrev(null); $("#carbCalc").open = false; ["#c_per","#c_g","#c_left"].forEach(i=>$(i).value=""); calcCarbs();
  if(type==="food"){ const h=new Date().getHours(); $("#e_meal").value = h<10?"Breakfast":h<11.5?"Morning snack":h<14?"Lunch":h<16.5?"Afternoon snack":h<19?"Dinner":"Bedtime snack"; }
  if(pre.gl!=null) $("#e_gl").value = pre.gl; if(pre.ket!=null) $("#e_ket").value = pre.ket;
  dlg.showModal();
}
function setFields(type){
  document.querySelectorAll("#entryForm [data-for]").forEach(el => el.hidden = !el.dataset.for.split(" ").includes(type));
  $("#carbCalc").hidden = type!=="food"; document.querySelector("[data-for-photo]").hidden = !(type==="food"||type==="drink");
  const fd = type==="food"||type==="drink";
  document.querySelectorAll("#e_gl,#e_ket").forEach(el => el.closest("label").hidden = fd);
  $("#e_itemLabel").textContent = type==="drink" ? "Drink" : "What he had";
  $("#e_item").placeholder = type==="drink" ? "e.g. Whole milk" : "e.g. Weetabix + whole milk";
  $("#e_carbsLabel").innerHTML = type==="food" ? "Carbs he ate <small>g</small>" : type==="hypo" ? "Carbs given <small>g</small>" : "Carbs <small>g</small>";
  $("#e_noteLabel").textContent = fd ? "Notes (optional)" : "What happened / what you did";
  if((type==="site"||type==="sensor") && typeof renderSpotPicker==="function") setTimeout(renderSpotPicker,0);
}
$("#e_type").addEventListener("change", e => setFields(e.target.value));
$("#e_item").addEventListener("change", () => {
  const v = $("#e_item").value.trim().toLowerCase(); if(!v) return;
  const fv = findFav(v); if(fv){ if($("#e_carbs").value==="" && fv.carbs!=null) $("#e_carbs").value = fv.carbs; if(fv.ml!=null && $("#e_ml").value==="") $("#e_ml").value = fv.ml; }
  const prev = log.find(e => (e.type==="food"||e.type==="drink") && (e.item||"").toLowerCase()===v);
  if(prev){ if(prev.carbs!=null && $("#e_carbs").value==="") $("#e_carbs").value = prev.carbs; if(prev.ml!=null && $("#e_ml").value==="") $("#e_ml").value = prev.ml; }
});
document.querySelectorAll("[data-q],[data-q2]").forEach(b => b.addEventListener("click", () => openEntry(b.dataset.q || b.dataset.q2)));
$("#dlgCancel").addEventListener("click", () => dlg.close());
$("#dlgDelete").addEventListener("click", async () => { const id=editingId; if(!id) return; dlg.close(); editingId=null; try { await deleteWithUndo("log", id, "Entry"); } catch { toast("Couldn't delete. Try again"); } });
document.addEventListener("keydown", ev => { if((ev.key==="Enter"||ev.key===" ") && ev.target.matches?.(".meal.tap")){ ev.preventDefault(); openEdit(ev.target.dataset.ledit); } });
$("#entryForm").addEventListener("submit", async ev => {
  ev.preventDefault();
  const e = { type:$("#e_type").value, at:new Date($("#e_at").value).getTime(), gl:num($("#e_gl").value), ket:num($("#e_ket").value), carbs:num($("#e_carbs").value), note:$("#e_note").value.trim(), by:me };
  if(e.type==="food"||e.type==="drink"){ e.item=$("#e_item").value.trim(); if(e.type==="food"){ e.meal=$("#e_meal").value; e.ate=$("#e_ate").value; } else { e.ml=num($("#e_ml").value); } e.gl=null; e.ket=null; }
  if(e.type==="site"||e.type==="sensor"){ const sp=$("#e_spot").value; if(sp) e.spot=sp; }
  if(isNaN(e.at)) e.at = Date.now();
  if(!editingId && $("#e_at").value===localNow()) e.at = Date.now();   // keep the exact time so two entries in the same minute stay in order
  $("#dlgSave").disabled = true;
  try {
    if(editingId){
      const old = log.find(x=>x.id===editingId) || {};
      const upd = {...old, ...e, by: old.by || me, editedBy: me, editedAt: Date.now()}; delete upd.id;
      if(photoData!==undefined){ upd.photo = !!(await savePhoto(editingId).catch(()=>false)); } Object.keys(upd).forEach(k => upd[k]===undefined && delete upd[k]);
      await base.collection("log").doc(editingId).set(upd);
      editingId = null; dlg.close(); toast("Changes saved"); return;
    }
    if(photoData){ e.photo = true; }
    const added = await base.collection("log").add(e);
    if(photoData){ try { await savePhoto(added.id); } catch { toast("Saved, but the photo didn't upload"); } }
    if(e.type==="food"||e.type==="drink"){ const fv=findFav(e.item);
      if(fv) base.collection("foods").doc(fv.id).update({uses:(fv.uses||0)+1, lastUsed:Date.now()}).catch(()=>{});
      else if($("#e_fav").checked && e.item) base.collection("foods").add({name:e.item, kind:e.type, carbs:e.carbs, ml:e.ml??null, portion:"", uses:1, lastUsed:Date.now()}).catch(()=>{}); }
    const linked = supplies.filter(s => s.link && s.link === e.type && s.qty > 0);
    for (const s of linked) await base.collection("supplies").doc(s.id).update({qty: s.qty-1});
    dlg.close(); toast(linked.length ? "Saved · stock updated" : "Saved");
  } catch(err){ toast(err.code==="quota_exceeded" ? "Storage is full. Delete old entries" : `Couldn't save (${err.code||err.message||"error"}). Try again`); }
  finally { $("#dlgSave").disabled = false; }
});
let pendingDel = null;
document.addEventListener("click", async ev => {
  const le = ev.target.closest("[data-ledit]"); if(le){ openEdit(le.dataset.ledit); return; }
  const fb = ev.target.closest("[data-fav]"); if(fb){ const f=foods.find(x=>x.id===fb.dataset.fav); if(f) logFav(f); return; }
  const d = ev.target.closest("[data-del]");
  if(d){
    d.disabled = true;
    try { await deleteWithUndo("log", d.dataset.del, "Entry"); } catch { toast("Couldn't delete. Try again"); d.disabled=false; }
    return;
  }
  const inc = ev.target.closest("[data-inc],[data-dec]");
  if(inc){
    const id = inc.dataset.inc || inc.dataset.dec, s = supplies.find(x=>x.id===id); if(!s) return;
    const q = Math.max(0, s.qty + (inc.dataset.inc ? 1 : -1)); inc.disabled = true;
    try { await base.collection("supplies").doc(id).update({qty:q}); } catch { toast("Couldn't update"); } finally { inc.disabled=false; }
    return;
  }
  const ed = ev.target.closest("[data-edit]"); if(ed){ openSup(supplies.find(x=>x.id===ed.dataset.edit)); return; }
  const sd = ev.target.closest("[data-sdel]");
  if(sd){
    sd.disabled = true;
    try { await deleteWithUndo("supplies", sd.dataset.sdel, "Item"); } catch { toast("Couldn't remove. Try again"); sd.disabled=false; }
  }
});
$("#range").addEventListener("change", renderLog); $("#filter").addEventListener("change", renderLog);
let searchT=null; $("#logSearch").addEventListener("input", () => { clearTimeout(searchT); searchT=setTimeout(renderLog, 150); });
$("#copyClinic").addEventListener("click", e => copy($("#clinicText").textContent, e.currentTarget));
["#ketIn","#glIn","#vomIn"].forEach(s => $(s).addEventListener("input", renderSick));
$("#logKet").addEventListener("click", () => openEntry("ketones", {ket:num($("#ketIn").value), gl:num($("#glIn").value)}));

/* ---------- supplies ---------- */
const supDlg = $("#supDlg"); let editing = null;
function openSup(s){ if(!canWrite) return (toast("Saving isn't switched on here. Open the page in a browser (see the note at the top)"), $("#dbNote").scrollIntoView({behavior:"smooth",block:"center"})); editing = s||null; $("#supForm").reset();
  if(s){ $("#s_name").value=s.name; $("#s_qty").value=s.qty; $("#s_days").value=s.days; $("#s_lead").value=s.lead??14; $("#s_link").value=s.link||""; $("#s_per").value=s.per??""; }
  supDlg.showModal(); }
$("#addSup").addEventListener("click", () => openSup());
$("#supCancel").addEventListener("click", () => supDlg.close());
$("#supForm").addEventListener("submit", async ev => {
  ev.preventDefault();
  const d = { name:$("#s_name").value.trim(), qty:+$("#s_qty").value||0, days:+$("#s_days").value||1, lead:+$("#s_lead").value||0, link:$("#s_link").value, per: $("#s_per").value==="" ? null : +$("#s_per").value };
  try { editing ? await base.collection("supplies").doc(editing.id).set(d) : await base.collection("supplies").add(d); supDlg.close(); toast("Saved"); }
  catch(err) { toast(`Couldn't save (${err.code||err.message||"error"})`); }
});
$("#starterSup").addEventListener("click", async e => {
  if(!canWrite) return; e.currentTarget.disabled = true;
  const site = num(plan.siteDays) ?? 3, sen = num(plan.sensorDays) ?? 10;
  const items = [
    ...(isMedtronic() ? [
      {name:"Infusion sets", qty:0, days:site, lead:14, link:"site"},
      {name:"Reservoirs", qty:0, days:site, lead:14, link:"site"},
      {name:"Guardian sensors", qty:0, days:sen, lead:14, link:"sensor"},
      {name:"Sensor overtape / patches", qty:0, days:sen, lead:14, link:"sensor"},
      {name:"AA batteries (pump)", qty:0, days:7, lead:7, link:""},
    ] : [
      {name:"Pump pods / infusion sets", qty:0, days:site, lead:14, link:"site"},
      {name:"CGM sensors", qty:0, days:sen, lead:14, link:"sensor"},
    ]),
    {name:"Rapid-acting insulin (vials/cartridges)", qty:0, days:28, lead:14, link:""},
    {name:"Backup insulin pens (basal + rapid)", qty:0, days:365, lead:30, link:""},
    {name:"Blood glucose test strips (pots)", qty:0, days:30, lead:14, link:""},
    {name:"Blood ketone strips (pots)", qty:0, days:90, lead:21, link:""},
    {name:"Glucagon kit", qty:0, days:365, lead:30, link:""},
  ].filter(i => !supplies.some(s => s.name === i.name));
  try { for (const i of items) await base.collection("supplies").add(i); toast(items.length ? "Added. Now set the stock you have" : "Already added"); }
  catch { toast("Couldn't add"); } finally { e.currentTarget.disabled = false; }
});


const PUMP_PRESETS = [
  {re:/medtronic|minimed|780g|640g/i, site:3, cgm:"Guardian 4", sensor:7, note:"Medtronic: standard infusion set every 3 days (7 with the Extended set). Guardian 4 sensor lasts 7 days."},
  {re:/omnipod/i, site:3, note:"Omnipod: new pod every 3 days."},
  {re:/tandem|t:slim/i, site:3, note:"Tandem: infusion set every 2–3 days, cartridge up to 3 days."},
];
$("#p_pump").addEventListener("change", () => {
  const pr = PUMP_PRESETS.find(x=>x.re.test($("#p_pump").value)); if(!pr) return;
  if(!$("#p_siteDays").value) $("#p_siteDays").value = pr.site;
  if(pr.cgm && !$("#p_cgm").value) $("#p_cgm").value = pr.cgm;
  if(pr.sensor && !$("#p_sensorDays").value) $("#p_sensorDays").value = pr.sensor;
  planDirty = true; planUi(); toast(pr.note + " Check with his team, then save.");
});
const isMedtronic = () => /medtronic|minimed|780g|640g/i.test(plan.pump||"");

/* ================= GLUCOSE ================= */
const MG = 18.0182;
let cgmDays = {};
let pumpDays = {};           // key YYYY-MM-DD -> {day, ev:[minute, kind, value, ...]}  kind: 1 carbs g, 2 insulin U, 3 finger prick mmol, 4 site change, 5 low alert, 6 high alert
const PK = {1:"Carbs (pump)", 2:"Insulin dose", 3:"Finger prick", 4:"Site change (pump)", 5:"Low alert", 6:"High alert"};
function dayPump(key){ const d=pumpDays[key]; if(!d) return []; const b=keyToMs(key), out=[]; for(let i=0;i<d.ev.length;i+=3) out.push({t:b+d.ev[i]*6e4, k:d.ev[i+1], v:d.ev[i+2]}); return out; }
function rangePump(days){ const out=[]; const end=midnight(Date.now()); for(let i=0;i<days;i++) out.push(...dayPump(ymd(end-i*864e5))); return out.sort((a,b)=>a.t-b.t); }
function hypoEpisodes(pts){ const eps=[]; let cur=null, prev=null;
  for(const p of pts){ const gap = prev && p.t-prev.t>20*6e4;
    if(p.v<3.9 && cur && !gap){ cur.end=p.t; if(p.v<cur.min) cur.min=p.v; }
    else { if(cur && cur.end-cur.start>=15*6e4) eps.push(cur); cur = p.v<3.9 ? {start:p.t, end:p.t, min:p.v} : null; }
    prev=p; }
  if(cur && cur.end-cur.start>=15*6e4) eps.push(cur); return eps; }
function lastPumpSite(){ let best=null; for(const k of Object.keys(pumpDays)) for(const e of dayPump(k)) if(e.k===4 && (!best || e.t>best)) best=e.t; return best; }            // key YYYY-MM-DD -> {day, pts:[minute, tenths, ...]}
let gDay = null;             // ms of local midnight shown in the day chart
const ymd = ms => { const d=new Date(ms); return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0"); };
const midnight = ms => { const d=new Date(ms); d.setHours(0,0,0,0); return d.getTime(); };
const keyToMs = k => { const [y,m,d]=k.split("-").map(Number); return new Date(y,m-1,d).getTime(); };
function dayPoints(key){ const d=cgmDays[key]; if(!d) return []; const base0=keyToMs(key), out=[]; for(let i=0;i<d.pts.length;i+=2) out.push({t:base0+d.pts[i]*6e4, v:d.pts[i+1]/10}); return out; }
function rangePoints(days){ const out=[]; const end=midnight(Date.now()); for(let i=0;i<days;i++){ out.push(...dayPoints(ymd(end-i*864e5))); } return out.sort((a,b)=>a.t-b.t); }
const band = v => v<3 ? "vlow" : v<3.9 ? "low" : v<=10 ? "in" : v<=13.9 ? "high" : "vhigh";
const pct = (a,n) => n ? a/n*100 : 0;
function quantile(sorted,q){ if(!sorted.length) return null; const p=(sorted.length-1)*q, b=Math.floor(p); return sorted[b+1]!==undefined ? sorted[b]+(p-b)*(sorted[b+1]-sorted[b]) : sorted[b]; }

/* ---- CSV parsing (Dexcom Clarity, Glooko, LibreView, generic) ---- */
/* ---- import glucose from almost any file: CSV/TSV/TXT, Excel (xlsx/xls/ods), JSON, or a zip of those ---- */
function splitCSV(line, sep){ const out=[]; let cur="", q=false; for(let i=0;i<line.length;i++){ const c=line[i]; if(q){ if(c==='"'){ if(line[i+1]==='"'){cur+='"';i++;} else q=false; } else cur+=c; } else if(c==='"') q=true; else if(c===sep){ out.push(cur); cur=""; } else cur+=c; } out.push(cur); return out.map(x=>x.trim()); }
const pad2 = n => String(n).padStart(2,"0");
function cellText(v){
  if(v==null) return "";
  if(v instanceof Date){ if(isNaN(v)) return ""; return v.getFullYear()<1901 ? `${pad2(v.getHours())}:${pad2(v.getMinutes())}` : `${v.getFullYear()}-${pad2(v.getMonth()+1)}-${pad2(v.getDate())} ${pad2(v.getHours())}:${pad2(v.getMinutes())}`; }
  return String(v).trim();
}
function parseWhen(s){
  s = (s||"").trim(); let m;
  if(/^\d{12,13}$/.test(s)) return +s;                       // epoch milliseconds (Nightscout)
  if(/^\d{9,10}$/.test(s)) return +s*1000;                    // epoch seconds
  if((m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})(?:[T ,]+(\d{1,2}):(\d{2}))?/))){
    if(/[zZ]$|[+-]\d{2}:?\d{2}$/.test(s) && s.includes("T")) { const t=Date.parse(s); if(!isNaN(t)) return t; }
    return new Date(+m[1],+m[2]-1,+m[3],+(m[4]||0),+(m[5]||0)).getTime(); }
  if((m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})(?:[ T,]+(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?)?/i))){
    let h=+(m[4]||0); if(m[6]){ const pm=/pm/i.test(m[6]); if(pm&&h<12) h+=12; if(!pm&&h===12) h=0; }
    let y=+m[3]; if(y<100) y+=2000;
    return new Date(y,+m[2]-1,+m[1],h,+(m[5]||0)).getTime(); }
  const t = Date.parse(s); return isNaN(t) ? NaN : t;
}
const joinDT = (d,t) => { const dm=(d||"").match(/^(\S+)/); const tm=(t||"").match(/(\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM)?)/i); return dm ? dm[1]+(tm?" "+tm[1]:"") : ""; };
function toMmol(raw, unit){
  raw = String(raw??"").trim().replace(",","."); if(!raw) return null;
  let v = /^low$/i.test(raw) ? (unit==="mg"?40:2.2) : /^high$/i.test(raw) ? (unit==="mg"?400:22.2) : parseFloat(raw.replace(/[^\d.\-]/g,""));
  if(isNaN(v) || v<=0) return null;
  if(unit==="mg" || (unit!=="mmol" && v>35)) v = v/MG;
  if(v<1 || v>35) return null;
  return Math.round(v*10)/10;
}

/* file -> list of tables (each a list of rows of text cells) */
async function tablesFromFile(file){
  const name = file.name.toLowerCase();
  if(/\.(png|jpe?g|heic|gif|webp)$/.test(name)) throw new Error("Pictures and screenshots only show the graph, not the readings. Import the PDF or a CSV/Excel export instead.");
  if(name.endsWith(".pdf")) return await pdfTables(file);
  if(name.endsWith(".zip")){
    if(!window.JSZip) throw new Error("Zip support didn't load. Unzip the file and import what's inside it.");
    const zip = await JSZip.loadAsync(file); const out=[];
    for(const n of Object.keys(zip.files)){ const z=zip.files[n]; if(z.dir || !/\.(csv|tsv|txt|json|xlsx|xls|ods)$/i.test(n)) continue;
      const blob = await z.async("blob"); out.push(...await tablesFromFile(new File([blob], n.split("/").pop()))); }
    // prefer files that look like CGM data
    return out.sort((a,b)=> (/cgm|sensor|glucose|egv/i.test(b.name)?1:0) - (/cgm|sensor|glucose|egv/i.test(a.name)?1:0));
  }
  if(/\.(xlsx|xls|ods|xlsm)$/.test(name)){
    if(!window.XLSX) throw new Error("Excel support didn't load. Save the sheet as CSV and import that instead.");
    const wb = XLSX.read(await file.arrayBuffer(), {type:"array", cellDates:true});
    return wb.SheetNames.map(sn => ({name:file.name+" › "+sn, rows: XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1, raw:true, defval:""}).map(r=>r.map(cellText))}));
  }
  const text = await file.text();
  if(name.endsWith(".json") || /^\s*[\[{]/.test(text)){
    let data; try { data = JSON.parse(text); } catch { throw new Error("That JSON file couldn't be read."); }
    const arr = Array.isArray(data) ? data : (Object.values(data).find(Array.isArray) || []);
    const objs = arr.filter(x=>x && typeof x==="object");
    if(!objs.length) throw new Error("No readings found in that JSON file.");
    const keys = [...new Set(objs.slice(0,200).flatMap(o=>Object.keys(o)))];
    return [{name:file.name, rows:[keys, ...objs.map(o=>keys.map(k=>cellText(o[k])))]}];
  }
  const lines = text.replace(/^﻿/,"").split(/\r?\n/);
  const sepOf = l => l.includes("\t") ? "\t" : l.split(";").length > l.split(",").length ? ";" : ",";
  return [{name:file.name, rows: lines.filter(l=>l.trim()).map(l=>splitCSV(l, sepOf(l)))}];
}

/* find the columns by their names */
const EXTRA = {
  carbs: [/bwz carb input/, /carb(s)? input/, /carb value/, /carbohydrates/, /^carbs?( \(g\w*\))?$/],
  bolus: [/bolus volume delivered/, /^(?!.*total)insulin delivered/, /rapid-acting insulin \(units\)/, /^(?!.*long)(?!.*basal).*insulin value/, /^bolus( \(u\))?$/, /^insulin( \(u\))?$/],
  bg:    [/^bg reading/, /strip glucose/, /blood glucose input/, /^bwz bg input/, /calibration bg/, /meter glucose/, /^mbg$/],
  site:  [/^rewind$/, /^prime type$/],
  alert: [/^alert$/, /^alarm/, /^event subtype$/],
};
function detectColumns(row){
  const low = row.map(c=>String(c).toLowerCase().trim());
  if(low.some(c=>/^\d/.test(c)) && !low.some(c=>/[a-z]{3}/.test(c))) return null;
  const pick = res => { for(const r of res){ const k=low.findIndex(c=>r.test(c)); if(k>=0) return k; } return -1; };
  const v = pick([/historic glucose/, /cgm glucose/, /sensor glucose/, /glucose value/, /^sgv$/, /^glucose( \(|$)/, /egv/, /^(?!.*scan)(?!.*bg reading)(?!.*meter)(?!.*input)(?!.*strip)(?!.*calibration).*glucose/, /^(mmol\/l|mg\/dl)$/, /^(reading|value)$/]);
  const ext = {}; for(const [k,res] of Object.entries(EXTRA)){ const i = pick(res); if(i>=0 && i!==v) ext[k]=i; }
  if(v<0 && !Object.keys(ext).length) return null;
  const dOnly = low.findIndex(c=>/^date$|^day$/.test(c)), tOnly = low.findIndex(c=>/^time$/.test(c));
  let t=-1, d=-1;
  if(dOnly>=0 && tOnly>=0){ d=dOnly; t=tOnly; } else t = pick([/device timestamp/, /^timestamp/, /timestamp/, /date.?time/, /^datestring$/, /^date$/, /^time$/, /time/, /date/]);
  if(t<0) return null;
  const unitOf = i => i<0 ? "auto" : /mg\/dl|^sgv$|^mbg$/.test(low[i]) ? "mg" : /mmol/.test(low[i]) ? "mmol" : "auto";
  return { v, t, d, e: low.findIndex(c=>/^event type$|^type$/.test(c)), unit: unitOf(v), ext, bgUnit: unitOf(ext.bg ?? -1) };
}
let IMP_EV = null;   // collects carbs, insulin, finger pricks, site changes and alerts during an import
function extract(rows, map, rescan){
  const pts=[]; let m = map;
  rows.forEach((r, i) => {
    if(i < (m.start||0)) return;
    if(rescan){ const h = detectColumns(r); if(h && (h.v>=0 || Object.keys(h.ext).length) && /[a-z]{3}/i.test(r[h.t]||"")) { m = {...h, start:0}; return; } }
    const t = parseWhen(m.d>=0 ? joinDT(r[m.d], r[m.t]) : r[m.t]); if(isNaN(t)) return;
    if(t > Date.now()+864e5 || t < Date.UTC(2000,0,1)) return;
    const et = m.e>=0 ? String(r[m.e]||"") : "";
    if(IMP_EV && m.ext){
      const X = m.ext, num1 = i => { const x = parseFloat(String(r[i]??"").replace(",",".")); return isNaN(x) ? null : x; };
      if(X.carbs!=null){ const c = num1(X.carbs); if(c>0 && c<300) IMP_EV.push({t, k:1, v:Math.round(c)}); }
      if(X.bolus!=null && !/basal/i.test(et) && !/long/i.test(X.alert!=null ? String(r[X.alert]||"") : "")){ const u = num1(X.bolus); if(u>0 && u<50) IMP_EV.push({t, k:2, v:Math.round(u*100)/100}); }
      if(X.bg!=null){ const b = toMmol(r[X.bg], m.bgUnit); if(b!=null) IMP_EV.push({t, k:3, v:b}); }
      if(X.site!=null){ const sv = String(r[X.site]||"").trim(); if(sv && !/^(false|0|no)$/i.test(sv) && /true|rewind|fill|yes|1/i.test(sv)) IMP_EV.push({t, k:4, v:0}); }
      if(X.alert!=null){ const a = String(r[X.alert]||""); if(/urgent low|low/i.test(a) && !/battery|reservoir/i.test(a)) IMP_EV.push({t, k:5, v:0}); else if(/high/i.test(a)) IMP_EV.push({t, k:6, v:0}); }
      if(/calibration|meter|fingerstick|mbg/i.test(et) && m.v>=0){ const b = toMmol(r[m.v], m.unit); if(b!=null) IMP_EV.push({t, k:3, v:b}); }
    }
    if(m.v<0) return;
    if(m.e>=0 && et && /[a-z]/i.test(et) && !/egv|sgv|sensor|cgm|glucose/i.test(et)) return;
    const v = toMmol(r[m.v], m.unit); if(v==null) return;
    pts.push({t, v});
  });
  return pts;
}
function autoExtract(rows){
  for(let i=0;i<Math.min(rows.length,80);i++){ const h = detectColumns(rows[i]); if(h){ const before = IMP_EV ? IMP_EV.length : 0; const pts = extract(rows, {...h, start:i+1}, true); if(pts.length || (IMP_EV && IMP_EV.length>before)) return pts; } }
  return [];
}

/* ---- PDF: read the text in the file and look for readings ---- */
async function pdfTables(file){
  const lib = window.pdfjsLib;
  if(!lib) throw new Error("PDF support didn't load. Check your internet connection and try again, or export the data as CSV or Excel.");
  try { if(!lib.GlobalWorkerOptions.workerSrc) lib.GlobalWorkerOptions.workerSrc = window.PDF_WORKER_SRC || "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js"; } catch {}
  const pdf = await lib.getDocument({data: new Uint8Array(await file.arrayBuffer()), isEvalSupported:false}).promise;
  const rows = []; let chars = 0;
  for(let n=1; n<=pdf.numPages; n++){
    const page = await pdf.getPage(n); const tc = await page.getTextContent();
    const items = tc.items.filter(it => it.str && it.str.trim()).map(it => ({s:it.str, x:it.transform[4], y:it.transform[5], w:it.width||0, h:Math.abs(it.transform[3])||8}));
    chars += items.reduce((a,i)=>a+i.s.length, 0);
    items.sort((a,b)=> b.y-a.y || a.x-b.x);
    const lines = [];
    for(const it of items){ const L = lines.find(l => Math.abs(l.y-it.y) <= Math.max(2, it.h*0.4)); L ? L.items.push(it) : lines.push({y:it.y, items:[it]}); }
    lines.sort((a,b)=>b.y-a.y);
    for(const L of lines){
      L.items.sort((a,b)=>a.x-b.x); const cells=[]; let cur=null;
      for(const it of L.items){ if(cur && it.x - (cur.x+cur.w) < it.h*0.6){ cur.s += (it.x-(cur.x+cur.w) > it.h*0.15 ? " " : "") + it.s; cur.w = it.x+it.w-cur.x; } else { cur = {...it}; cells.push(cur); } }
      rows.push(cells.map(c=>c.s.trim()));
    }
  }
  if(chars < 20) throw new Error("This PDF has no readable text, only pictures of the graphs, so there are no readings in it to import. Export the data as CSV or Excel from his CGM or pump website.");
  return [{name:file.name, rows, pdf:true}];
}
const MONTHS = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,sept:8,oct:9,nov:10,dec:11};
function findDate(text){
  let m;
  if((m = text.match(/\b(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})\b/))) return new Date(+m[1],+m[2]-1,+m[3]).getTime();
  if((m = text.match(/\b(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})\b/))){ let y=+m[3]; if(y<100) y+=2000; return new Date(y,+m[2]-1,+m[1]).getTime(); }
  if((m = text.match(/\b(\d{1,2})(?:st|nd|rd|th)?\s+(jan|feb|mar|apr|may|jun|jul|aug|sept?|oct|nov|dec)[a-z]*\.?,?\s+(\d{4})\b/i))) return new Date(+m[3],MONTHS[m[2].toLowerCase()],+m[1]).getTime();
  if((m = text.match(/\b(jan|feb|mar|apr|may|jun|jul|aug|sept?|oct|nov|dec)[a-z]*\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})\b/i))) return new Date(+m[3],MONTHS[m[1].toLowerCase()],+m[2]).getTime();
  return null;
}
function patternExtract(rows){
  const all = rows.map(r=>r.join("  ")).join("\n");
  const docUnit = /mg\/dl/i.test(all) && !/mmol/i.test(all) ? "mg" : /mmol/i.test(all) && !/mg\/dl/i.test(all) ? "mmol" : "auto";
  const skip = /generated|printed|report date|page \d+ of|average|mean|median|target|time in range|gmi|a1c|estimated|standard deviation/i;
  const pts = []; let day = null;
  for(const r of rows){
    const line = r.join("  ");
    const d = findDate(line); if(d!=null) day = d;
    if(day==null || skip.test(line)) continue;
    const rx = /\b(\d{1,2})[:.](\d{2})(?::\d{2})?\s*(am|pm)?\b(?:\s*(?:\||-|–|,|;|\s)\s*)(\d{1,3}(?:[.,]\d)?)(?![:\d])(?!\s*(?:u\b|units?\b|g\b|grams|%|iu\b))\s*(mmol\/l|mg\/dl)?/gi; let m;
    while((m = rx.exec(line))){
      let h=+m[1]; const mi=+m[2]; if(h>23||mi>59) continue;
      if(m[3]){ const pm=/pm/i.test(m[3]); if(pm&&h<12) h+=12; if(!pm&&h===12) h=0; }
      const unit = m[5] ? (/mg/i.test(m[5])?"mg":"mmol") : docUnit;
      const v = toMmol(m[4], unit); if(v==null) continue;
      pts.push({t: day + h*36e5 + mi*6e4, v});
    }
  }
  return pts;
}
function confirmPdf(pts, table){
  const box = $("#gMap"); const days = new Set(pts.map(p=>ymd(p.t))).size;
  const sample = pts.slice(0,6).map(p=>`<tr><td>${new Date(p.t).toLocaleString("en-GB",{day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"})}</td><td class="num">${p.v.toFixed(1)} mmol/L</td></tr>`).join("");
  box.hidden = false;
  box.innerHTML = `<div><h3>Check the readings from this PDF</h3><div class="muted">I found <strong>${pts.length.toLocaleString("en-GB")} readings over ${days} day${days===1?"":"s"}</strong> in <strong>${esc(table.name)}</strong>. Reading PDFs is less reliable than CSV, so check the first few against the report before importing.</div></div>
    <div class="table-wrap"><table class="preview"><tbody>${sample}</tbody></table></div>
    <div class="btns"><button class="primary" type="button" id="p_go">Import ${pts.length.toLocaleString("en-GB")} readings</button><button type="button" id="p_map">Pick the columns myself</button><button type="button" id="p_no">Cancel</button></div>`;
  box.scrollIntoView({behavior:"smooth", block:"start"});
  $("#p_go").onclick = async () => { box.hidden = true; await saveReadings(pts); };
  $("#p_map").onclick = () => showMapper(table);
  $("#p_no").onclick = () => { box.hidden = true; gStatus(""); };
}

/* column picker for files whose columns we don't recognise */
function showMapper(table){
  const box = $("#gMap"); const rows = table.rows; const width = Math.max(...rows.slice(0,30).map(r=>r.length));
  let hi = 0; for(let i=0;i<Math.min(rows.length,40);i++){ if(rows[i].filter(c=>/[a-z]{2}/i.test(c)).length>=2 && rows[i].length>=2){ hi=i; break; } }
  const head = i => rows[hi]?.[i] ? `${String.fromCharCode(65+(i%26))}: ${rows[hi][i]}` : `Column ${String.fromCharCode(65+(i%26))}`;
  const opts = (sel, none) => (none?`<option value="-1">${none}</option>`:"") + Array.from({length:width},(_,i)=>`<option value="${i}" ${i===sel?"selected":""}>${esc(head(i))}</option>`).join("");
  const guessT = rows[hi]?.findIndex(c=>/time|date/i.test(c)) ?? -1, guessV = rows[hi]?.findIndex(c=>/glucose|bg|value|reading/i.test(c)) ?? -1;
  box.hidden = false;
  box.innerHTML = `<div><h3>Which columns are which?</h3><div class="muted">I couldn't recognise the columns in <strong>${esc(table.name)}</strong>. Pick them below. The preview shows the first rows of the file.</div></div>
    <div class="table-wrap"><table class="preview"><tbody>${rows.slice(0,8).map((r,i)=>`<tr${i===hi?' class="hd"':''}><th>${i+1}</th>${Array.from({length:width},(_,j)=>`<td>${esc(r[j]??"")}</td>`).join("")}</tr>`).join("")}</tbody></table></div>
    <div class="grid2">
      <label>Headings are on row<input id="m_hi" type="number" min="1" value="${hi+1}"></label>
      <label>Date and time<select id="m_t">${opts(guessT)}</select></label>
      <label>Separate time column <small>only if the date and time are in two columns</small><select id="m_d">${opts(-1,"Not needed")}</select></label>
      <label>Glucose reading<select id="m_v">${opts(guessV)}</select></label>
      <label>Units<select id="m_u"><option value="auto">Work it out</option><option value="mmol">mmol/L</option><option value="mg">mg/dL</option></select></label>
    </div>
    <div class="btns"><button class="primary" type="button" id="m_go">Import with these columns</button><button type="button" id="m_no">Cancel</button></div>`;
  box.scrollIntoView({behavior:"smooth", block:"start"});
  $("#m_no").onclick = () => { box.hidden = true; gStatus(""); };
  $("#m_go").onclick = async () => {
    const tcol = +$("#m_t").value, extra = +$("#m_d").value;
    const map = { v:+$("#m_v").value, t: extra>=0 ? extra : tcol, d: extra>=0 ? tcol : -1, e:-1, unit:$("#m_u").value, start: Math.max(0, +$("#m_hi").value) };
    const pts = extract(rows, map, false);
    if(!pts.length){ gStatus("No readings found with those columns. Check the date and glucose columns and try again."); return; }
    box.hidden = true; await saveReadings(pts);
  };
}

async function savePumpEvents(evs){
  const byDay = {}; for(const e of evs){ const k=ymd(e.t); const mnt=Math.round((e.t-keyToMs(k))/6e4); (byDay[k] ||= new Map()).set(mnt+":"+e.k, [mnt, e.k, e.v]); }
  for(const [k,m] of Object.entries(byDay)){
    const old = pumpDays[k]?.ev || []; for(let i=0;i<old.length;i+=3){ const key = old[i]+":"+old[i+1]; if(!m.has(key)) m.set(key, [old[i], old[i+1], old[i+2]]); }
    const flat=[]; [...m.values()].sort((a,b)=>a[0]-b[0]||a[1]-b[1]).forEach(x=>flat.push(...x));
    await base.collection("pump").doc(k).set({day:k, ev:flat});
  }
  renderGlucose(); return evs.length;
}
async function saveReadings(pts){
  const byDay = {}; for(const p of pts){ const k=ymd(p.t); (byDay[k] ||= new Map()).set(Math.round((p.t-keyToMs(k))/6e4), Math.round(p.v*10)); }
  const keys = Object.keys(byDay).sort(); let n=0;
  for(const k of keys){
    const m = byDay[k]; const old = cgmDays[k]?.pts || []; for(let i=0;i<old.length;i+=2) if(!m.has(old[i])) m.set(old[i], old[i+1]);
    const flat=[]; [...m.entries()].sort((a,b)=>a[0]-b[0]).forEach(([mm,v])=>flat.push(mm,v));
    gStatus(`Saving day ${++n} of ${keys.length}…`);
    await base.collection("cgm").doc(k).set({day:k, pts:flat});
  }
  gDay = keyToMs(keys[keys.length-1]);
  gStatus(`Imported ${pts.length.toLocaleString("en-GB")} readings from ${keys.length} day${keys.length===1?"":"s"} (${new Date(keyToMs(keys[0])).toLocaleDateString("en-GB")} to ${new Date(keyToMs(keys[keys.length-1])).toLocaleDateString("en-GB")}).`);
  renderGlucose();
}
function gStatus(msg){ const n=$("#gStatus"); n.hidden=!msg; n.textContent=msg||""; }
$("#gFile").addEventListener("change", async ev => {
  const f = ev.target.files[0]; ev.target.value=""; if(!f) return;
  if(!canWrite){ gStatus("Saving isn't available in this view."); return; }
  $("#gMap").hidden = true; gStatus("Reading "+f.name+"…");
  try {
    const tables = (await tablesFromFile(f)).filter(t=>t.rows.length);
    if(!tables.length){ gStatus("That file is empty."); return; }
    IMP_EV = []; let pts=[]; for(const t of tables) pts = pts.concat(autoExtract(t.rows)); const evs = IMP_EV; IMP_EV = null;
    const pdfT = tables.find(t=>t.pdf);
    if(pdfT){ if(!pts.length) pts = patternExtract(pdfT.rows); gStatus("");
      if(pts.length) { confirmPdf(pts.sort((a,b)=>a.t-b.t), pdfT); return; }
      gStatus("I couldn't find readings with times and glucose values in this PDF. Summary reports (like the AGP or overview pages) only show graphs and averages. Pick the columns below if the readings are in a table, or export a CSV."); showMapper(pdfT); return; }
    if(!pts.length && !evs.length){ gStatus(""); showMapper(tables.sort((a,b)=>b.rows.length-a.rows.length)[0]); return; }
    if(pts.length) await saveReadings(pts);
    if(evs.length){ const n = await savePumpEvents(evs); const c = k => evs.filter(e=>e.k===k).length;
      const bits = [c(1)&&`${c(1)} carb entries`, c(2)&&`${c(2)} insulin doses`, c(3)&&`${c(3)} finger pricks`, c(4)&&`${c(4)} site changes`, c(5)+c(6)&&`${c(5)+c(6)} alerts`].filter(Boolean);
      gStatus((pts.length ? $("#gStatus").textContent + " " : "") + `Also found ${bits.join(", ")}.`); }
  } catch(e){ gStatus(e.code==="quota_exceeded" ? "Storage is full." : (e.message || "Couldn't import that file.")); }
});

/* ---- stats ---- */
function cgmStats(days){
  const pts = rangePoints(days); const n = pts.length; if(!n) return null;
  const vals = pts.map(p=>p.v); const mean = vals.reduce((a,b)=>a+b,0)/n;
  const sd = Math.sqrt(vals.reduce((a,b)=>a+(b-mean)**2,0)/n);
  const b = {vlow:0,low:0,in:0,high:0,vhigh:0}; vals.forEach(v=>b[band(v)]++);
  const gmiPct = 3.31 + 0.02392*mean*MG;
  // CGM-detected lows: >=15 min below 3.9
  let lows=0, runStart=null, counted=false, prevT=null;
  for(const p of pts){ const gap = prevT!==null && p.t-prevT>20*6e4; if(p.v<3.9 && !gap && runStart!==null){ if(!counted && p.t-runStart>=15*6e4){ lows++; counted=true; } } else if(p.v<3.9){ runStart=p.t; counted=false; } else { runStart=null; counted=false; } prevT=p.t; }
  const span = new Set(pts.map(p=>ymd(p.t))).size;
  const perDay = pts.length/Math.max(1,span);
  return {pts, n, mean, sd, cv: sd/mean*100, b, gmiPct, gmiMmol: Math.round((gmiPct-2.15)*10.929), lows, span, coverage: Math.min(100, pct(n, days*(perDay>150?288:96)))};
}

/* ---- render ---- */
function renderGlucose(){ setTimeout(()=>{ renderBgNow(); renderBgChecks(); hypoState(); renderWatch(); if(typeof renderA1c==="function"){ renderA1c(); renderActs(); } if(typeof renderNudge==="function"){ renderNudge(); renderSettings(); } }, 0);
  const keys = Object.keys(cgmDays).sort();
  const days = +$("#gRange").value;
  const st = cgmStats(days);
  $("#gSince").textContent = keys.length ? `CGM data from ${new Date(keyToMs(keys[0])).toLocaleDateString("en-GB")} to ${new Date(keyToMs(keys[keys.length-1])).toLocaleDateString("en-GB")}. Import a new file to add more days.` : "Import a file from his CGM, pump or meter (CSV, Excel, PDF, JSON or zip) to see his graphs here, with meals and events on top.";
  if(!st){
    $("#gStats").innerHTML = `<div class="stat" style="grid-column:1/-1"><span class="k">No CGM data in this period</span><span class="n">${keys.length?"Pick a longer period or import a newer file.":"Use Import glucose file above. Until then, the day chart shows only finger-prick readings you've logged."}</span></div>`;
    $("#tirBar").innerHTML = `<div class="empty">Time in range appears after you import CGM data.</div>`;
    $("#agpChart").innerHTML = `<div class="empty">The typical-day chart appears after you import a few days of CGM data.</div>`;
    $("#gInsights").innerHTML = `<li class="muted">Import CGM data to see plain-English patterns here.</li>`;
  } else {
    const tir = pct(st.b.in, st.n);
    { const segs=[["in","var(--ok)"],["high","var(--high)"],["vhigh","var(--vhigh)"],["low","var(--low)"],["vlow","var(--vlow)"]]; const R=58, C=2*Math.PI*R; let off=0;
      const arcs = segs.map(([k,c])=>{ const f=pct(st.b[k],st.n)/100; if(f<=0) return ""; const a=`<circle cx="70" cy="70" r="${R}" fill="none" stroke="${c}" stroke-width="16" stroke-dasharray="${Math.max(0,C*f-(f<1?2:0)).toFixed(1)} ${C.toFixed(1)}" stroke-dashoffset="${(-off*C).toFixed(1)}" transform="rotate(-90 70 70)"/>`; off+=f; return a; }).join("");
      const met = (k,v,u,n,cls="") => `<div class="gm ${cls}"><span class="k">${k}</span><span class="v">${v}<small>${u}</small></span><span class="n">${n}</span></div>`;
      $("#gStats").innerHTML = `<div class="gsum"><div class="gdonut"><svg viewBox="0 0 140 140" width="140" height="140" aria-hidden="true"><circle cx="70" cy="70" r="${R}" fill="none" stroke="var(--sunk)" stroke-width="16"/>${arcs}</svg>
        <div class="gd-in"><b>${Math.round(tir)}%</b><span>in range</span></div></div>
        <div class="gd-meta"><div class="gd-title ${tir>=70?"good":"bad"}">${tir>=70?"On target":"Below target"}<span>target is over 70% in range</span></div>
          <div class="gms">${met("Average", st.mean.toFixed(1), " mmol/L", "across the period")}${met("Est. HbA1c", st.gmiMmol, " mmol/mol", st.gmiPct.toFixed(1)+"% · lab may differ")}${met("Ups & downs", Math.round(st.cv), "%", st.cv>36?"quite variable":"steady", st.cv>36?"warn":"")}${met("Sensor data", Math.round(st.coverage), "%", st.coverage<70?"gaps: less reliable":"enough to trust", st.coverage<70?"warn":"")}</div></div></div>`; }
    const rows = [["vhigh","Very high","above 13.9","under 5%",v=>v<5],["high","High","10.1–13.9","",null],["in","In range","3.9–10.0","over 70%",v=>v>70],["low","Low","3.0–3.8","under 4%",v=>v<4],["vlow","Very low","below 3.0","under 1%",v=>v<1]];
    $("#tirBar").innerHTML = `<div class="tir" role="img" aria-label="Time in range bar">${["vlow","low","in","high","vhigh"].map(k=>{const p=pct(st.b[k],st.n); return p>0?`<i style="width:${p}%;background:var(--${k==="in"?"ok":k})" title="${p.toFixed(1)}%"></i>`:"";}).join("")}</div>
      <div class="tirlist">${rows.map(([k,l,r,t,f])=>{const p=pct(st.b[k],st.n); const hrs=p*24/100; return `<div><span class="sw" style="background:var(--${k==="in"?"ok":k})"></span><span>${l} <span class="muted">${r} mmol/L · about ${hrs<1?Math.round(hrs*60)+" min":hrs.toFixed(1)+" h"} a day</span></span><span class="pc">${p.toFixed(p<10?1:0)}%</span><span class="tg ${f?(f(p)?"good":"bad"):""}">${t?"target "+t:""}</span></div>`;}).join("")}</div>`;
    renderAGP(st); renderInsights(st, days);
  }
  renderFromFile(days);
  if(gDay===null) gDay = keys.length ? keyToMs(keys[keys.length-1]) : midnight(Date.now());
  $("#dDate").value = ymd(gDay);
  renderDay();
}

const CH = {W:820, l:40, r:14, t:12, ph:230};
const yS = v => CH.t + CH.ph - (Math.max(2,Math.min(22,v))-2)/20*CH.ph;
function svgFrame(h, xTicks){
  let g = `<rect x="${CH.l}" y="${yS(10)}" width="${CH.W-CH.l-CH.r}" height="${yS(3.9)-yS(10)}" fill="var(--ok-soft)"/>`;
  [3.9,10,13.9,18,22].forEach(v=>{ g+=`<line x1="${CH.l}" x2="${CH.W-CH.r}" y1="${yS(v)}" y2="${yS(v)}" stroke="var(--line)" stroke-width="1" ${v===3.9||v===10?"":'stroke-dasharray="3 4"'}/><text x="${CH.l-6}" y="${yS(v)+4}" text-anchor="end">${v}</text>`; });
  xTicks.forEach(([x,lab])=>{ g+=`<line x1="${x}" x2="${x}" y1="${CH.t}" y2="${CH.t+CH.ph}" stroke="var(--line)" stroke-opacity=".5"/><text x="${x}" y="${CH.t+CH.ph+16}" text-anchor="middle">${lab}</text>`; });
  return g;
}
const xH = h => CH.l + h/24*(CH.W-CH.l-CH.r);
const hourTicks = () => (CH.W<560?[0,6,12,18,24]:[0,3,6,9,12,15,18,21,24]).map(h=>[xH(h), String(h%24).padStart(2,"0")+":00"]);
const EVT = {food:{c:"var(--ok)"}, drink:{c:"var(--accent)"}, hypo:{c:"var(--low)"}, high:{c:"var(--high)"}, ketones:{c:"var(--urgent)"}};
function fitChart(el){ const w = el.clientWidth || 820; CH.W = Math.max(320, Math.min(1100, Math.round(w))); }
function renderDay(){
  fitChart($("#dayChart"));
  const key = ymd(gDay), d0 = gDay, pts = dayPoints(key);
  const evs = log.filter(e => e.at>=d0 && e.at<d0+864e5).sort((a,b)=>a.at-b.at);
  const manual = evs.filter(e=>e.gl!=null);
  const pev = dayPump(key); const pumpY = CH.t+CH.ph+92;
  const H = CH.t+CH.ph+(pev.length?126:90), laneY = CH.t+CH.ph+44;
  let g = svgFrame(H, hourTicks());
  // food guides
  evs.filter(e=>e.type==="food").forEach(e=>{ const x=xH((e.at-d0)/36e5); g+=`<line x1="${x}" x2="${x}" y1="${CH.t}" y2="${laneY-8}" stroke="var(--ok)" stroke-opacity=".45" stroke-dasharray="2 3"/>`; });
  // glucose line, coloured by band, broken on gaps
  const col = v => v<3.9 ? "var(--low)" : v>10 ? "var(--high)" : "var(--accent)";
  g += `<defs><linearGradient id="gfill" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="var(--accent)" stop-opacity=".22"/><stop offset="1" stop-color="var(--accent)" stop-opacity="0"/></linearGradient></defs>`;
  g += `<text x="${CH.l+8}" y="${yS(10)+15}" style="fill:var(--ok);font-weight:700;font-size:11px">Target 3.9–10</text>`;
  { let seg=[]; const flush=()=>{ if(seg.length>1){ const base=yS(2); g+=`<path d="M${seg[0][0]} ${base} L${seg.map(q=>q.join(" ")).join(" L")} L${seg[seg.length-1][0]} ${base} Z" fill="url(#gfill)"/>`; } seg=[]; };
    pts.forEach((q,i)=>{ if(i && q.t-pts[i-1].t>20*6e4) flush(); seg.push([xH((q.t-d0)/36e5).toFixed(1), yS(q.v).toFixed(1)]); }); flush(); }
  for(let i=1;i<pts.length;i++){ const a=pts[i-1], b=pts[i]; if(b.t-a.t>20*6e4) continue;
    g+=`<line x1="${xH((a.t-d0)/36e5).toFixed(1)}" y1="${yS(a.v).toFixed(1)}" x2="${xH((b.t-d0)/36e5).toFixed(1)}" y2="${yS(b.v).toFixed(1)}" stroke="${col((a.v+b.v)/2)}" stroke-width="2.6" stroke-linecap="round"/>`; }
  manual.forEach(e=>{ g+=`<circle cx="${xH((e.at-d0)/36e5)}" cy="${yS(e.gl)}" r="5" fill="var(--surface)" stroke="var(--fg)" stroke-width="2"/>`; });
  // event lane
  g+=`<text x="${CH.l-6}" y="${laneY+4}" text-anchor="end">Events</text><line x1="${CH.l}" x2="${CH.W-CH.r}" y1="${laneY}" y2="${laneY}" stroke="var(--line)"/>`;
  evs.forEach(e=>{ const x=xH((e.at-d0)/36e5), c=(EVT[e.type]||{c:"var(--muted)"}).c;
    if(e.type==="food"){ g+=`<circle cx="${x}" cy="${laneY}" r="6" fill="${c}" stroke="var(--surface)" stroke-width="2"/>${e.carbs!=null?`<text x="${x}" y="${laneY+22}" text-anchor="middle" style="fill:var(--fg);font-weight:700">${e.carbs}g</text>`:""}`; }
    else if(e.type==="drink"){ g+=`<rect x="${x-4.5}" y="${laneY-4.5}" width="9" height="9" transform="rotate(45 ${x} ${laneY})" fill="${c}" stroke="var(--surface)" stroke-width="2"/>`; }
    else if(e.type==="hypo"){ g+=`<path d="M${x} ${laneY-7} L${x+7} ${laneY+6} L${x-7} ${laneY+6} Z" fill="${c}" stroke="var(--surface)" stroke-width="2"/>`; }
    else { g+=`<rect x="${x-5}" y="${laneY-5}" width="10" height="10" rx="2" fill="${c}" stroke="var(--surface)" stroke-width="2"/>`; }
  });
  // finger pricks from the file
  pev.filter(e=>e.k===3).forEach(e=>{ g+=`<rect x="${xH((e.t-d0)/36e5)-4.5}" y="${yS(e.v)-4.5}" width="9" height="9" fill="var(--surface)" stroke="var(--fg)" stroke-width="2" transform="rotate(45 ${xH((e.t-d0)/36e5)} ${yS(e.v)})"/>`; });
  if(pev.length){
    g+=`<text x="${CH.l-6}" y="${pumpY+4}" text-anchor="end">Pump</text><line x1="${CH.l}" x2="${CH.W-CH.r}" y1="${pumpY}" y2="${pumpY}" stroke="var(--line)"/>`;
    pev.forEach(e=>{ const x=xH((e.t-d0)/36e5);
      if(e.k===1) g+=`<circle cx="${x}" cy="${pumpY}" r="6" fill="var(--surface)" stroke="var(--ok)" stroke-width="2.5"/><text x="${x}" y="${pumpY+21}" text-anchor="middle" style="fill:var(--ok);font-weight:700">${e.v}g</text>`;
      else if(e.k===2) g+=`<path d="M${x-5} ${pumpY-20} L${x+5} ${pumpY-20} L${x} ${pumpY-10} Z" fill="var(--accent)" stroke="var(--surface)" stroke-width="1.5"/>`;
      else if(e.k===4) g+=`<rect x="${x-5}" y="${pumpY-5}" width="10" height="10" rx="2" fill="var(--accent)" stroke="var(--surface)" stroke-width="2"/>`;
      else if(e.k===5||e.k===6) g+=`<circle cx="${x}" cy="${pumpY}" r="3.5" fill="${e.k===5?"var(--low)":"var(--high)"}"/>`; });
  }
  if(!pts.length) g+=`<text x="${(CH.W+CH.l)/2}" y="${yS(12)}" text-anchor="middle" style="font-size:13px">${Object.keys(cgmDays).length?"No CGM data for this day":"No CGM data yet. Import a file above"}</text>`;
  g+=`<line id="dCross" x1="0" x2="0" y1="${CH.t}" y2="${laneY+8}" stroke="var(--fg)" stroke-opacity=".5" visibility="hidden"/><circle id="dDot" r="5" fill="var(--accent)" stroke="var(--surface)" stroke-width="2" visibility="hidden"/>`;
  g+=`<rect id="dHit" x="${CH.l}" y="0" width="${CH.W-CH.l-CH.r}" height="${H}" fill="transparent"/>`;
  const box=$("#dayChart");
  box.innerHTML = `<svg viewBox="0 0 ${CH.W} ${H}" role="img" aria-label="Glucose for ${new Date(d0).toLocaleDateString("en-GB")}">${g}</svg><div class="tip" hidden></div>`;
  hover(box, H, x => {
    const t = d0 + (x-CH.l)/(CH.W-CH.l-CH.r)*864e5;
    let best=null; for(const p of pts) if(!best || Math.abs(p.t-t)<Math.abs(best.t-t)) best=p;
    if(best && Math.abs(best.t-t)>15*6e4) best=null;
    const near = evs.filter(e=>Math.abs(e.at-t)<20*6e4), pnear = pev.filter(e=>Math.abs(e.t-t)<20*6e4);
    if(!best && !near.length && !pnear.length) return null;
    const tt = best?best.t:near.length?near[0].at:pnear[0].t;
    return {x:xH((tt-d0)/36e5), y: best?yS(best.v):null, html:
      `<div class="muted">${new Date(tt).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}</div>`+
      (best?`<div><b>${best.v.toFixed(1)}</b> mmol/L · ${({vlow:"very low",low:"low",in:"in range",high:"high",vhigh:"very high"})[band(best.v)]}</div>`:"")+
      near.map(e=>`<div>${TYPES[e.type]?.label||"Note"}${e.item?": "+esc(e.item):""}${e.carbs!=null?` · ${e.carbs} g`:""}${e.ml!=null?` · ${e.ml} ml`:""}${e.gl!=null?` · ${e.gl.toFixed(1)}`:""}</div>`).join("")+
      pnear.map(e=>`<div class="muted">${PK[e.k]}${e.k===1?` · ${e.v} g`:e.k===2?` · ${e.v} U`:e.k===3?` · ${e.v.toFixed(1)} mmol/L`:""}</div>`).join("")};
  });
}
function hover(box, H, find){
  const svg=box.querySelector("svg"), tip=box.querySelector(".tip"), cross=svg.querySelector("[id$=Cross]"), dot=svg.querySelector("[id$=Dot]");
  const hide=()=>{ tip.hidden=true; cross.setAttribute("visibility","hidden"); dot.setAttribute("visibility","hidden"); };
  const move = ev => {
    const r=svg.getBoundingClientRect(), sx=CH.W/r.width, x=(ev.clientX-r.left)*sx;
    const f=find(x); if(!f){ hide(); return; }
    cross.setAttribute("x1",f.x); cross.setAttribute("x2",f.x); cross.setAttribute("visibility","visible");
    if(f.y!=null){ dot.setAttribute("cx",f.x); dot.setAttribute("cy",f.y); dot.setAttribute("visibility","visible"); } else dot.setAttribute("visibility","hidden");
    tip.innerHTML=f.html; tip.hidden=false;
    const px=f.x/sx + svg.getBoundingClientRect().left - box.getBoundingClientRect().left + box.scrollLeft;
    const left = px + 14 + tip.offsetWidth > box.scrollLeft + box.clientWidth ? px - 14 - tip.offsetWidth : px + 14;
    tip.style.left = Math.max(box.scrollLeft+4,left)+"px"; tip.style.top = "8px";
  };
  svg.addEventListener("pointermove", move); svg.addEventListener("pointerdown", move); svg.addEventListener("pointerleave", hide);
}
function renderAGP(st){
  fitChart($("#agpChart"));
  const bins = Array.from({length:48},()=>[]);
  st.pts.forEach(p=>{ const d=new Date(p.t); bins[d.getHours()*2+(d.getMinutes()>=30?1:0)].push(p.v); });
  const q = bins.map((b,i)=>{ if(b.length<3) return null; const s=b.slice().sort((a,c)=>a-c); return {i, h:(i+.5)/2, p5:quantile(s,.05), p25:quantile(s,.25), p50:quantile(s,.5), p75:quantile(s,.75), p95:quantile(s,.95), n:s.length}; });
  const ok = q.filter(Boolean); const H = CH.t+CH.ph+28;
  let g = svgFrame(H, hourTicks());
  const area=(lo,hi)=> "M"+ok.map(b=>`${xH(b.h).toFixed(1)} ${yS(b[hi]).toFixed(1)}`).join(" L")+" L"+ok.slice().reverse().map(b=>`${xH(b.h).toFixed(1)} ${yS(b[lo]).toFixed(1)}`).join(" L")+" Z";
  if(ok.length>1){
    g+=`<path d="${area("p5","p95")}" fill="var(--accent)" fill-opacity=".12"/>`;
    g+=`<path d="${area("p25","p75")}" fill="var(--accent)" fill-opacity=".28"/>`;
    g+=`<path d="M${ok.map(b=>`${xH(b.h).toFixed(1)} ${yS(b.p50).toFixed(1)}`).join(" L")}" fill="none" stroke="var(--accent)" stroke-width="2.5" stroke-linejoin="round"/>`;
  }
  g+=`<line id="aCross" x1="0" x2="0" y1="${CH.t}" y2="${CH.t+CH.ph}" stroke="var(--fg)" stroke-opacity=".5" visibility="hidden"/><circle id="aDot" r="5" fill="var(--accent)" stroke="var(--surface)" stroke-width="2" visibility="hidden"/>`;
  const box=$("#agpChart");
  box.innerHTML = `<svg viewBox="0 0 ${CH.W} ${H}" role="img" aria-label="Typical day pattern">${g}</svg><div class="tip" hidden></div>
    <div class="legendrow" style="margin-top:8px"><span><i class="lg line"></i>Middle value (median)</span><span><i class="lg" style="background:var(--accent);opacity:.4;border-radius:2px"></i>Half of days fall here</span><span><i class="lg" style="background:var(--accent);opacity:.15;border-radius:2px"></i>Almost all days fall here</span></div>`;
  hover(box, H, x => { const h=(x-CH.l)/(CH.W-CH.l-CH.r)*24; const b=q[Math.max(0,Math.min(47,Math.floor(h*2)))]; if(!b) return null;
    const lab = `${String(Math.floor(b.i/2)).padStart(2,"0")}:${b.i%2?"30":"00"}`;
    return {x:xH(b.h), y:yS(b.p50), html:`<div class="muted">${lab}–${String(Math.floor((b.i+1)/2)%24).padStart(2,"0")}:${b.i%2?"00":"30"}</div><div>Usually <b>${b.p50.toFixed(1)}</b> mmol/L</div><div class="muted">Half of days ${b.p25.toFixed(1)}–${b.p75.toFixed(1)}</div><div class="muted">Range ${b.p5.toFixed(1)}–${b.p95.toFixed(1)}</div>`}; });
}
function valueNear(t, win){ const k=ymd(t); let best=null; for(const p of [...dayPoints(k), ...dayPoints(ymd(t-864e5)), ...dayPoints(ymd(t+864e5))]) if(Math.abs(p.t-t)<=win && (!best||Math.abs(p.t-t)<Math.abs(best.t-t))) best=p; return best; }

function renderFromFile(days){
  const st=$("#ffStrip"); if(!st) return;
  const ev = rangePump(days), pts = rangePoints(days);
  const dset = new Set([...pts.map(p=>ymd(p.t)), ...ev.map(e=>ymd(e.t))]); const nd = Math.max(1, dset.size);
  const sum = k => ev.filter(e=>e.k===k).reduce((a,e)=>a+e.v,0), cnt = k => ev.filter(e=>e.k===k).length;
  const ps = lastPumpSite();
  const tiles = [
    cnt(1) ? [`Carbs entered in pump`, `${Math.round(sum(1)/nd)} g`, `a day on average · ${cnt(1)} entries`] : null,
    cnt(2) ? [`Insulin doses`, `${(sum(2)/nd).toFixed(1)} U`, `bolus a day on average · ${cnt(2)} doses`] : null,
    cnt(3) ? [`Finger pricks`, `${cnt(3)}`, `in this period`] : null,
    cnt(4) || ps ? [`Site changes (pump)`, `${cnt(4)}`, ps ? `last ${fmtD(ps)}` : "in this period"] : null,
    cnt(5)+cnt(6) ? [`Pump / CGM alerts`, `${cnt(5)} low · ${cnt(6)} high`, `in this period`] : null,
  ].filter(Boolean);
  st.innerHTML = tiles.length ? tiles.map(([k,v,n])=>`<div class="stat"><span class="k">${k}</span><span class="v">${v}</span><span class="n">${n}</span></div>`).join("")
    : `<div class="stat" style="grid-column:1/-1"><span class="k">Nothing extra yet</span><span class="n">Import a full export (for Medtronic: CareLink CSV with all data) to see carbs, insulin, finger pricks and site changes here.</span></div>`;
  const eps = hypoEpisodes(pts).reverse();
  $("#hypoTable").innerHTML = eps.length ? eps.slice(0,40).map(h=>{
      const before = [...log.filter(e=>e.at<=h.start && e.at>=h.start-3*36e5 && /food|activity|ill|drink/.test(e.type)).map(e=>(e.type==="food"?(e.item||"Food")+(e.carbs!=null?` (${e.carbs} g)`:""):TYPES[e.type].label)),
                      ...ev.filter(e=>e.t<=h.start && e.t>=h.start-3*36e5 && (e.k===1||e.k===2)).map(e=>e.k===1?`${e.v} g carbs (pump)`:`${e.v} U insulin`)];
      const mins = Math.round((h.end-h.start)/6e4)+5;
      return `<tr><td>${new Date(h.start).toLocaleString("en-GB",{weekday:"short",day:"numeric",month:"short",hour:"2-digit",minute:"2-digit"})}</td><td class="num"><span class="chip ${h.min<3?"low":"high"} num">${h.min.toFixed(1)}</span></td><td class="num">${mins} min</td><td class="muted" style="overflow-wrap:anywhere">${before.length?esc([...new Set(before)].slice(0,4).join(", ")):"Nothing logged in the 3 hours before"}</td></tr>`; }).join("")
    : `<tr><td colspan="4"><div class="empty">${pts.length?"No hypos detected in this period.":"Import CGM data to see hypos here."}</div></td></tr>`;
}
function renderInsights(st, days){
  const L=[]; const tir=pct(st.b.in,st.n), below=pct(st.b.low+st.b.vlow,st.n), above=pct(st.b.high+st.b.vhigh,st.n);
  const hrs = p => { const h=p*24/100; return h<1 ? Math.round(h*60)+" minutes" : h.toFixed(1)+" hours"; };
  L.push(`<strong>In range ${Math.round(tir)}% of the time</strong>, about ${hrs(tir)} a day. ${tir>=70?"That meets the usual target of more than 70%.":`The usual target is more than 70%, so he's ${Math.round(70-tir)} points short.`}`);
  L.push(`<strong>Low ${below.toFixed(1)}% of the time</strong> (about ${hrs(below)} a day). ${below<4?"That's within the target of under 4%.":"That's above the target of under 4%. Lows matter most to raise with his team."} His CGM picked up ${st.lows} low${st.lows===1?"":"s"} lasting 15 minutes or more in this period.`);
  L.push(`<strong>High ${Math.round(above)}% of the time</strong> (about ${hrs(above)} a day). Target is under 25%.`);
  const hourOf = f => { const c=Array(12).fill(0); st.pts.forEach(p=>{ if(f(p.v)) c[Math.floor(new Date(p.t).getHours()/2)]++; }); const m=Math.max(...c); if(!m) return null; const i=c.indexOf(m); return `${String(i*2).padStart(2,"0")}:00 and ${String(i*2+2).padStart(2,"0")}:00`; };
  const lowT = hourOf(v=>v<3.9), highT = hourOf(v=>v>10);
  if(lowT && below>0.5) L.push(`Lows happen most often <strong>between ${lowT}</strong>.`);
  if(highT && above>5) L.push(`Highs happen most often <strong>between ${highT}</strong>.`);
  const night = st.pts.filter(p=>new Date(p.t).getHours()<6);
  if(night.length>20){ const nl=pct(night.filter(p=>p.v<3.9).length, night.length); const nm=night.reduce((a,p)=>a+p.v,0)/night.length; L.push(`<strong>Overnight</strong> (midnight to 6am) his average was ${nm.toFixed(1)} mmol/L and he was low ${nl.toFixed(1)}% of the night.`); }
  // after meals
  const meals = log.filter(e=>e.type==="food" && e.at>midnight(Date.now())-(days-1)*864e5);
  const rises = {}; let used=0;
  meals.forEach(e=>{ const s=valueNear(e.at,15*6e4); if(!s) return; let peak=null; const all=[...dayPoints(ymd(e.at)),...dayPoints(ymd(e.at+864e5))]; for(const p of all) if(p.t>e.at && p.t<=e.at+3*36e5 && (!peak||p.v>peak.v)) peak=p; if(!peak) return; used++; (rises[e.meal||"Food"] ||= []).push({rise:peak.v-s.v, mins:Math.round((peak.t-e.at)/6e4), item:e.item, start:s.v, peak:peak.v}); });
  if(used){
    const parts = Object.entries(rises).map(([m,a])=>{ const r=a.reduce((x,y)=>x+y.rise,0)/a.length, mn=Math.round(a.reduce((x,y)=>x+y.mins,0)/a.length); return {m, r, mn, n:a.length}; }).sort((a,b)=>b.r-a.r);
    L.push(`<strong>After meals</strong> (${used} logged meal${used===1?"":"s"} with CGM data): ` + parts.map(p=>`${p.m.toLowerCase()} rises by ${p.r.toFixed(1)} mmol/L on average, peaking after about ${p.mn} min`).join("; ")+".");
    const all = Object.values(rises).flat().filter(x=>x.item).sort((a,b)=>b.rise-a.rise);
    if(all.length>=3 && all[0].rise>4) L.push(`Biggest rise after a logged meal: <strong>${esc(all[0].item)}</strong>, ${all[0].start.toFixed(1)} → ${all[0].peak.toFixed(1)} mmol/L.`);
  } else L.push(`Log his meals in Food &amp; drink and this section will show how much each meal raises his glucose.`);
  if(st.cv>36) L.push(`His glucose goes up and down a lot (CV ${Math.round(st.cv)}%, steady is 36% or less). This is common in toddlers because of small doses and unpredictable eating.`);
  if(st.coverage<70) L.push(`Only ${Math.round(st.coverage)}% of the expected sensor readings are here, so treat these numbers with care.`);
  $("#gInsights").innerHTML = L.map(x=>`<li>${x}</li>`).join("");
}




/* ================= DELIVERIES (supplies arriving every few months) ================= */
let delivery = {};
const dEvery = () => num(delivery.every) ?? 3;
function nextDelivery(){ if(delivery.next && delivery.next >= midnight(Date.now())-864e5) return delivery.next; if(!delivery.last) return null; let n = addMonths(delivery.last, dEvery()); while(n < midnight(Date.now())) n = addMonths(n, dEvery()); return n; }
const daysTo = ms => Math.round((midnight(ms)-midnight(Date.now()))/864e5);
function perDelivery(sp){ if(sp.per!=null) return sp.per; if(!sp.days) return null; return Math.ceil(dEvery()*30.4/sp.days); }
function untilDelivery(sp){
  const nd = nextDelivery(); if(!nd || !sp.days) return `<span class="muted">—</span>`;
  const spare = Math.floor(sp.left) - daysTo(nd);
  return spare >= 7 ? `<span class="chip ok">Lasts · ${spare} d spare</span>` : spare >= 0 ? `<span class="chip high">Just enough · ${spare} d spare</span>` : `<span class="chip low">Runs out ${-spare} d early</span>`;
}
function renderDeliveries(){
  if(!$("#dStrip")) return;
  const nd = nextDelivery(), lead = num(delivery.lead) ?? 14;
  const orderBy = nd ? nd - lead*864e5 : null;
  const short = supplyRows().filter(sp => nd && sp.days && Math.floor(sp.left) < daysTo(nd));
  $("#dStrip").innerHTML = [
    `<div class="stat"><span class="k">Next delivery</span><span class="v">${nd?(daysTo(nd)===0?"Today":daysTo(nd)===1?"Tomorrow":"In "+daysTo(nd)+" days"):"—"}</span><span class="n">${nd?fmtD(nd):"Set the last delivery date"}</span></div>`,
    `<div class="stat ${orderBy && daysTo(orderBy)<=3 ? (daysTo(orderBy)<0?"bad":"warn") : ""}"><span class="k">Order / confirm by</span><span class="v">${orderBy?(daysTo(orderBy)<0?"Overdue":daysTo(orderBy)===0?"Today":fmtD(orderBy).replace(/ \d{4}$/,"")):"—"}</span><span class="n">${orderBy?`${lead} days before delivery`:"Add delivery settings"}</span></div>`,
    `<div class="stat ${short.length?"bad":""}"><span class="k">Will it last?</span><span class="v">${!nd?"—":short.length?short.length+" short":"Yes"}</span><span class="n">${!nd?"Needs a delivery date":short.length?esc(short.map(x=>x.name).slice(0,2).join(", ")):"Everything lasts until the next box"}</span></div>`,
    `<div class="stat"><span class="k">Last delivery</span><span class="v">${delivery.last?fmtD(delivery.last).replace(/ \d{4}$/,""):"—"}</span><span class="n">${delivery.who?esc(delivery.who):"Every "+dEvery()+" months"}</span></div>`
  ].join("");
  $("#deliveryCheck").innerHTML = !delivery.last && !delivery.next
    ? `<div class="empty">Tap <strong>Delivery settings</strong> and enter when the last box arrived. The app then works out the next delivery, when to order, and whether stock will last.</div>`
    : short.length
      ? `<div class="band low"><strong>Not enough to reach the next delivery</strong><span>${short.map(x=>`${esc(x.name)} runs out about ${daysTo(nd)-Math.floor(x.left)} days early`).join(" · ")}. Ask the supplier or GP for extra before ${fmtD(nd)}.${delivery.phone?` Supplier: <span style="user-select:all">${esc(delivery.phone)}</span>`:""}</span></div>`
      : `<div class="band ok"><strong>Stock lasts until the next delivery</strong><span>Next box ${fmtD(nd)}${orderBy?`, order or confirm by ${fmtD(orderBy)}`:""}.${delivery.phone?` Supplier: <span style="user-select:all">${esc(delivery.phone)}</span>`:""}</span></div>`;
}
const delSetDlg = $("#delSetDlg"), delLogDlg = $("#delLogDlg");
$("#editDelivery").addEventListener("click", () => { if(!canWrite) return toast("Saving isn't switched on in this view");
  $("#dl_last").value = delivery.last ? ymd(delivery.last) : ""; $("#dl_every").value = dEvery(); $("#dl_next").value = delivery.next ? ymd(delivery.next) : "";
  $("#dl_lead").value = num(delivery.lead) ?? 14; $("#dl_who").value = delivery.who || ""; $("#dl_phone").value = delivery.phone || ""; delSetDlg.showModal(); });
$("#delSetCancel").addEventListener("click", () => delSetDlg.close());
$("#delSetForm").addEventListener("submit", async ev => { ev.preventDefault();
  const d = { last: $("#dl_last").value ? dateMs($("#dl_last").value) : null, every: +$("#dl_every").value || 3, next: $("#dl_next").value ? dateMs($("#dl_next").value) : null,
    lead: +$("#dl_lead").value || 0, who: $("#dl_who").value.trim(), phone: $("#dl_phone").value.trim() };
  try { await base.collection("meta").doc("delivery").set(d); delSetDlg.close(); toast("Saved"); } catch(err){ toast(`Couldn't save (${err.code||err.message||"error"})`); } });
$("#logDelivery").addEventListener("click", () => { if(!canWrite) return toast("Saving isn't switched on in this view");
  const rows = supplyRows(); if(!rows.length) return toast("Add your supply items first");
  $("#dlg_date").value = ymd(Date.now());
  $("#delItems").innerHTML = rows.map(sp => `<label class="del-row"><span><strong>${esc(sp.name)}</strong><small>${sp.qty} at home now</small></span><input type="number" min="0" step="1" data-dq="${esc(sp.id)}" value="${perDelivery(sp) ?? 0}" aria-label="Amount of ${esc(sp.name)}"></label>`).join("");
  delLogDlg.showModal(); });
$("#delLogCancel").addEventListener("click", () => delLogDlg.close());
$("#delLogForm").addEventListener("submit", async ev => { ev.preventDefault(); $("#delLogSave").disabled = true;
  try {
    for(const inp of document.querySelectorAll("#delItems [data-dq]")){ const add = +inp.value || 0; if(!add) continue; const sp = supplies.find(x=>x.id===inp.dataset.dq); if(!sp) continue;
      await base.collection("supplies").doc(sp.id).update({ qty: (sp.qty||0) + add, per: add }); }
    const arrived = dateMs($("#dlg_date").value || ymd(Date.now()));
    await base.collection("meta").doc("delivery").set({ ...delivery, last: arrived, next: null, every: dEvery() });
    delLogDlg.close(); toast("Delivery added to stock");
  } catch(err){ toast(`Couldn't save (${err.code||err.message||"error"})`); } finally { $("#delLogSave").disabled = false; } });

/* ================= APPOINTMENTS ================= */
let appts = [];
const AK = { diabetes:{label:"Diabetes clinic", chip:"blue"}, gp:{label:"GP", chip:""}, vaccine:{label:"Vaccination", chip:"ok"}, hospital:{label:"Hospital", chip:"urgent"},
  dietitian:{label:"Dietitian", chip:"food"}, dentist:{label:"Dentist", chip:""}, eye:{label:"Eye / optician", chip:""}, health:{label:"Health visitor", chip:""}, other:{label:"Other", chip:""} };
$("#ap_kind").innerHTML = Object.entries(AK).map(([k,v])=>`<option value="${k}">${v.label}</option>`).join("");
$("#apptFilter").innerHTML += Object.entries(AK).map(([k,v])=>`<option value="${k}">${v.label}</option>`).join("");
const apptWhen = a => a.at;
// old eye/dental check-ups (saved before appointments existed) shown as past appointments
const legacyChecks = () => checks.map(c=>({id:"chk_"+c.id, chkId:c.id, kind:c.kind==="dental"?"dentist":"eye", status:"done", at:c.at, hasTime:false, where:c.where||"", who:"", notes:"", outcome:c.note||"", next:c.next||null, legacy:true}));
const apptsAll = () => [...appts, ...legacyChecks()];
function apptCard(a, past){
  const d = new Date(a.at), k = AK[a.kind] || AK.other, days = Math.round((midnight(a.at)-midnight(Date.now()))/864e5);
  const rel = a.status==="cancelled" ? "Cancelled" : past ? "" : days===0 ? "Today" : days===1 ? "Tomorrow" : days<0 ? "Overdue: mark done or cancelled" : `In ${days} days`;
  return `<article class="appt ${a.status==="cancelled"?"cx":""}">
    <div class="appt-date"><b>${d.getDate()}</b><span>${d.toLocaleDateString("en-GB",{month:"short"})}</span><small>${a.hasTime?d.toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}):d.getFullYear()}</small></div>
    <div class="appt-body"><div><span class="chip ${k.chip}">${k.label}</span>${a.kind==="vaccine"&&a.vaccine?`<strong>${esc(a.vaccine)}</strong>`:""}${rel?` <span class="muted appt-rel ${days<0&&!past?"bad":""}">${rel}</span>`:""}</div>
      ${a.where||a.who?`<div class="muted">${esc([a.where,a.who].filter(Boolean).join(" · "))}</div>`:""}
      ${a.notes?`<div class="appt-notes"><span class="muted">${past?"Notes":"To ask"}:</span> ${esc(a.notes)}</div>`:""}
      ${a.outcome?`<div class="appt-notes"><span class="muted">Outcome:</span> ${esc(a.outcome)}</div>`:""}
      <div class="btns">
        ${!past&&a.status==="booked"?`<button class="small primary" data-adone="${esc(a.id)}">Mark done</button>`:""}
        <button class="small" data-aedit="${esc(a.id)}" ${canWrite?"":"disabled"}>Edit</button>
        ${!past&&a.status==="booked"?`<button class="small" data-aics="${esc(a.id)}">Add to calendar</button>`:""}
        ${!past&&a.kind==="diabetes"?`<button class="small" data-aprep="${esc(a.id)}">Copy clinic prep</button>`:""}
        <button class="small danger" data-adel="${esc(a.id)}" ${canWrite?"":"disabled"}>Delete</button>
      </div></div></article>`;
}
function renderAppts(){
  if(!$("#apptUpcoming")) return;
  const now = Date.now(), today0 = midnight(now);
  const up = appts.filter(a=>a.status==="booked" && a.at>=today0).sort((a,b)=>a.at-b.at);
  const overdue = appts.filter(a=>a.status==="booked" && a.at<today0).sort((a,b)=>b.at-a.at);
  const f = $("#apptFilter").value;
  const past = apptsAll().filter(a=>a.status!=="booked" && (!f || a.kind===f)).sort((a,b)=>b.at-a.at);
  $("#apptUpcoming").innerHTML = (overdue.map(a=>apptCard(a,false)).join("") + up.map(a=>apptCard(a,false)).join("")) || `<div class="empty">No appointments booked. Tap <strong>Add appointment</strong> for his next GP, clinic or vaccination.</div>`;
  $("#apptPast").innerHTML = past.length ? past.map(a=>apptCard(a,true)).join("") : `<div class="empty">Nothing here yet. When you mark an appointment done, it moves here with its outcome.</div>`;
  const vax = appts.filter(a=>a.kind==="vaccine" && a.status==="done").sort((a,b)=>b.at-a.at);
  $("#vaxTable").innerHTML = vax.length ? vax.map(a=>`<tr><td>${fmtD(a.at)}</td><td><strong>${esc(a.vaccine||"Vaccination")}</strong></td><td class="muted">${esc(a.where||"")}</td><td class="muted" style="overflow-wrap:anywhere">${esc(a.outcome||a.notes||"")}</td></tr>`).join("") : `<tr><td colspan="4"><div class="empty">No vaccinations recorded. Add past ones from his red book as appointments marked <strong>Done</strong>.</div></td></tr>`;
  const next = up[0], nextDia = up.find(a=>a.kind==="diabetes"), lastVax = vax[0];
  const inDays = a => { const d=Math.round((midnight(a.at)-today0)/864e5); return d===0?"Today":d===1?"Tomorrow":`In ${d} days`; };
  $("#aStrip").innerHTML = [
    `<div class="stat ${next && Math.round((midnight(next.at)-today0)/864e5)<=1?"warn":""}"><span class="k">Next appointment</span><span class="v">${next?inDays(next):"—"}</span><span class="n">${next?esc((AK[next.kind]||AK.other).label)+" · "+fmtD(next.at):"None booked"}</span></div>`,
    `<div class="stat"><span class="k">Next diabetes clinic</span><span class="v">${nextDia?inDays(nextDia):"—"}</span><span class="n">${nextDia?fmtD(nextDia.at):"Not booked yet"}</span></div>`,
    `<div class="stat"><span class="k">Last vaccination</span><span class="v">${lastVax?esc(lastVax.vaccine||"Vaccine"):"—"}</span><span class="n">${lastVax?fmtD(lastVax.at):"None recorded"}</span></div>`,
    `<div class="stat ${overdue.length?"bad":""}"><span class="k">Need updating</span><span class="v">${overdue.length}</span><span class="n">${overdue.length?"Past dates still marked booked":"All up to date"}</span></div>`
  ].join("");
}
const apptDlg = $("#apptDlg"); let editingAppt = null;
function apptKindFields(){ document.querySelectorAll("#apptForm [data-ak]").forEach(el=>el.hidden = !el.dataset.ak.split(" ").includes($("#ap_kind").value)); }
$("#ap_kind").addEventListener("change", apptKindFields);
function openAppt(a, markDone, preset){
  if(!canWrite) return toast("Saving isn't switched on in this view");
  editingAppt = a || null; $("#apptForm").reset();
  $("#apptTitle").textContent = a ? (markDone ? "How did it go?" : "Edit appointment") : "Add appointment";
  if(a){ const d=new Date(a.at); $("#ap_kind").value=a.kind; $("#ap_status").value=markDone?"done":a.status; $("#ap_date").value=ymd(a.at);
    $("#ap_time").value = a.hasTime ? `${pad2(d.getHours())}:${pad2(d.getMinutes())}` : ""; $("#ap_vax").value=a.vaccine||""; $("#ap_where").value=a.where||""; $("#ap_who").value=a.who||""; $("#ap_notes").value=a.notes||""; $("#ap_outcome").value=a.outcome||""; $("#ap_next").value=a.next?ymd(a.next):""; }
  else { $("#ap_kind").value=preset?.kind||"diabetes"; $("#ap_status").value=preset?.status||"booked"; $("#ap_date").value=preset?.date||ymd(Date.now()+7*864e5); }
  apptKindFields(); apptDlg.showModal(); if(markDone) setTimeout(()=>$("#ap_outcome").focus(),50);
}
$("#addAppt").addEventListener("click", ()=>openAppt());
$("#apptCancel").addEventListener("click", ()=>apptDlg.close());
$("#apptFilter").addEventListener("change", renderAppts);
$("#apptForm").addEventListener("submit", async ev => { ev.preventDefault();
  const [y,m,d] = $("#ap_date").value.split("-").map(Number); if(!y) return toast("Pick a date");
  const t = $("#ap_time").value; const [hh,mm] = t ? t.split(":").map(Number) : [12,0];
  const body = { kind:$("#ap_kind").value, status:$("#ap_status").value, at:new Date(y,m-1,d,hh,mm).getTime(), hasTime:!!t, vaccine:$("#ap_kind").value==="vaccine"?$("#ap_vax").value.trim():"",
    where:$("#ap_where").value.trim(), who:$("#ap_who").value.trim(), notes:$("#ap_notes").value.trim(), outcome:$("#ap_outcome").value.trim(), by:me,
    next: ["eye","dentist"].includes($("#ap_kind").value) && $("#ap_next").value ? (([yy,mm2,dd]) => new Date(yy,mm2-1,dd,12).getTime())($("#ap_next").value.split("-").map(Number)) : null };
  $("#apptSave").disabled = true;
  try {
    if(editingAppt?.legacy){ await base.collection("appts").add(body); await base.collection("checks").doc(editingAppt.chkId).delete(); }   // move an old check-up into appointments
    else editingAppt ? await base.collection("appts").doc(editingAppt.id).set(body) : await base.collection("appts").add(body);
    apptDlg.close(); toast("Saved"); }
  catch(err){ toast(`Couldn't save (${err.code||err.message||"error"})`); } finally { $("#apptSave").disabled = false; }
});
function icsFor(a){
  const d=new Date(a.at), z=n=>String(n).padStart(2,"0"), dt=x=>`${x.getFullYear()}${z(x.getMonth()+1)}${z(x.getDate())}`;
  const start = a.hasTime ? `DTSTART:${dt(d)}T${z(d.getHours())}${z(d.getMinutes())}00` : `DTSTART;VALUE=DATE:${dt(d)}`;
  const end = a.hasTime ? `DURATION:PT45M` : `DTEND;VALUE=DATE:${dt(new Date(a.at+864e5))}`;
  const title = `${P("name")||"Son"}: ${(AK[a.kind]||AK.other).label}${a.vaccine?" ("+a.vaccine+")":""}`;
  const esc2 = t => String(t||"").replace(/[\\,;]/g, m=>"\\"+m).replace(/\n/g,"\\n");
  return ["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//T1D Care Hub//EN","BEGIN:VEVENT",`UID:${a.id}@t1d-care-hub`,`DTSTAMP:${new Date().toISOString().replace(/[-:]/g,"").slice(0,15)}Z`,start,end,
    `SUMMARY:${esc2(title)}`, a.where?`LOCATION:${esc2(a.where)}`:"", `DESCRIPTION:${esc2([a.who?"With: "+a.who:"", a.notes?"To ask: "+a.notes:""].filter(Boolean).join("\n"))}`,
    "BEGIN:VALARM","TRIGGER:-P1D","ACTION:DISPLAY","DESCRIPTION:Tomorrow: "+esc2(title),"END:VALARM",
    ...(a.hasTime?["BEGIN:VALARM","TRIGGER:-PT2H","ACTION:DISPLAY","DESCRIPTION:In 2 hours: "+esc2(title),"END:VALARM"]:[]),
    "END:VEVENT","END:VCALENDAR"].filter(Boolean).join("\r\n");
}
async function getDownloads(){ dl ??= await (window.claude?.use?.("downloads") ?? null); return dl; }
document.addEventListener("click", async ev => {
  const t = ev.target.closest("[data-adone],[data-aedit],[data-adel],[data-aics],[data-aprep]"); if(!t) return;
  const id = t.dataset.adone||t.dataset.aedit||t.dataset.adel||t.dataset.aics||t.dataset.aprep; const a = apptsAll().find(x=>x.id===id); if(!a) return;
  if(t.dataset.adone) return openAppt(a, true);
  if(t.dataset.aedit) return openAppt(a, false);
  if(t.dataset.adel){ t.disabled=true; try { await (a.legacy ? deleteWithUndo("checks", a.chkId, "Check-up") : deleteWithUndo("appts", id, "Appointment")); } catch { toast("Couldn't delete"); t.disabled=false; } return; }
  if(t.dataset.aics){ const d = await getDownloads(); if(!d) return toast("Saving files isn't available here");
    try { const r = await d.save({filename:`appointment-${ymd(a.at)}.ics`, data: icsFor(a)}); if(r.status==="saved") toast("Open the file to add it to your calendar"); }
    catch(e){ toast(e.code==="declined"?"Cancelled":e.code==="rejected_extension"?"Calendar files can't be saved in this view. Use the installed app":"Couldn't create the calendar file"); } return; }
  if(t.dataset.aprep){ renderLog(); const qs = [...document.querySelectorAll("#mealPatterns .mpq li")].map((li,i)=>`${i+1}. ${li.textContent}`);
    const txt = `Clinic appointment ${fmtD(a.at)}${a.where?" at "+a.where:""}\n\n${$("#clinicText").textContent}${a.notes?`\n\nOur questions:\n${a.notes}`:""}${qs.length?`\n\nQuestions about meals:\n${qs.join("\n")}`:""}`;
    try { await navigator.clipboard.writeText(txt); toast("Copied: paste it into an email to his team"); } catch { toast("Couldn't copy. Use History → Copy summary"); } }
});





/* ---------- BG: latest glucose + finger-prick checks ---------- */
function allBgChecks(days=30){
  const from = Date.now()-days*864e5, out=[];
  for(const k of Object.keys(pumpDays)) for(const e of dayPump(k)) if(e.k===3 && e.t>=from) out.push({t:e.t, v:e.v, src:"Meter / pump file"});
  for(const e of log) if(e.type==="bg" && e.gl!=null && e.at>=from) out.push({t:e.at, v:e.gl, src:"Logged", id:e.id, note:e.note});
  return out.sort((a,b)=>b.t-a.t);
}
function latestCgm(){
  const k = Object.keys(cgmDays).sort().pop(); if(!k) return null;
  const pts = dayPoints(k); if(!pts.length) return null;
  const last = pts[pts.length-1], prev = pts.filter(p=>p.t<=last.t-14*6e4 && p.t>=last.t-21*6e4).pop();
  const rate = prev ? (last.v-prev.v)/((last.t-prev.t)/6e4)*15 : null;   // change over 15 minutes
  return {t:last.t, v:last.v, src:"CGM file", rate};
}
function latestGlucose(){
  const c = latestCgm(), b = allBgChecks(3650)[0];
  return (!c || (b && b.t>c.t)) ? (b ? {...b, rate:null} : c) : c;
}
function bgBand(v){ const lo=hypoBelow(); return v<3 ? {c:"low",t:"Very low"} : v<lo ? {c:"low",t:"Low"} : v>13.9 ? {c:"high",t:"Very high"} : v>10 ? {c:"high",t:"High"} : {c:"ok",t:"In range"}; }
function whenText(t){ const d=(midnight(Date.now())-midnight(t))/864e5, hm=new Date(t).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}); return d<=0 ? `Today ${hm}` : d===1 ? `Yesterday ${hm}` : new Date(t).toLocaleDateString("en-GB",{weekday:"short",day:"numeric",month:"short"})+" "+hm; }
function renderBgNow(){
  const el=$("#bgNow"); if(!el) return; const g=latestGlucose(); try { hypoState(); } catch {}
  if(!g){ el.className = "panel bgnow band-none"; el.innerHTML = `<div class="bg-row"><div class="bg-big"><span class="k">Latest glucose</span><b class="muted">—</b><span class="bg-src">Upload his CGM file or log a finger prick</span></div><div class="bg-acts"><button type="button" class="small primary" data-bgupload>Upload CGM file</button><button type="button" class="small" data-bglog>Log a BG</button></div></div>`; return; }
  const b=bgBand(g.v), ageH=(Date.now()-g.t)/36e5; el.className = "panel bgnow band-"+b.c;
  const arrow = g.rate==null ? "" : g.rate>=1.5 ? "↑" : g.rate>=0.5 ? "↗" : g.rate<=-1.5 ? "↓" : g.rate<=-0.5 ? "↘" : "→";
  el.innerHTML = `<div class="bg-row"><div class="bg-big"><span class="k">Latest glucose</span><b class="bg-v ${b.c}">${g.v.toFixed(1)}<small> mmol/L</small>${arrow?`<span class="bg-arrow" aria-label="trend">${arrow}</span>`:""}</b><span class="bg-src"><span class="chip ${b.c}">${b.t}</span> ${whenText(g.t)} · ${g.src}</span></div>
    <div class="bg-acts"><button type="button" class="small primary" data-bgupload>Upload CGM file</button><button type="button" class="small" data-bglog>Log a BG</button></div></div>
    ${ageH>3?`<p class="bg-old">This is from ${ageH<48?Math.round(ageH)+" hours":Math.round(ageH/24)+" days"} ago, not live. Check his pump or CGM for now.</p>`:""}`;
}
function renderBgChecks(){
  const st=$("#bgStats"), ul=$("#bgList"); if(!st) return;
  const days=+($("#gRange")?.value||14), L=allBgChecks(days);
  if(!L.length){ st.innerHTML=""; ul.innerHTML=`<li class="empty">No finger-prick readings in the last ${days} days. They come in with a full CareLink export, or log one with <b>Log a BG</b>.</li>`; return; }
  const avg=L.reduce((a,x)=>a+x.v,0)/L.length, lows=L.filter(x=>x.v<hypoBelow()).length, perDay=L.length/days;
  st.innerHTML = [["Checks",String(L.length),`about ${perDay<1?perDay.toFixed(1):Math.round(perDay)} a day`],["Average",avg.toFixed(1),"mmol/L"],["Low",String(lows),`under ${hypoBelow().toFixed(1)}`]]
    .map(([k,v,n])=>`<div class="wk"><span class="k">${k}</span><b>${v}</b><small>${n}</small></div>`).join("");
  ul.innerHTML = L.slice(0,30).map(x=>{ const b=bgBand(x.v); return `<li><span class="bg-t">${whenText(x.t)}</span><span class="bg-n">${x.src}${x.note?` · ${esc(x.note)}`:""}</span><span class="chip ${b.c} num">${x.v.toFixed(1)}</span></li>`; }).join("");
}
document.addEventListener("click", ev => { if(ev.target.closest("[data-bglog]")) openEntry("bg"); if(ev.target.closest("[data-bgupload]")){ location.hash="glucose"; setTimeout(()=>$("#gFile")?.click(), 50); } });


/* ---------- insulin page (from the pump file) ---------- */
function rangeEvents(days){ const from=midnight(Date.now())-(days-1)*864e5, out=[]; for(const k of Object.keys(pumpDays)) for(const e of dayPump(k)) if(e.t>=from) out.push(e); return out.sort((a,b)=>b.t-a.t); }
function renderInsulin(){
  const st=$("#inStats"); if(!st) return; const days=+($("#inRange")?.value||14), ev=rangeEvents(days);
  const bol=ev.filter(e=>e.k===2), carbs=ev.filter(e=>e.k===1);
  const dayKeys=[...Array(days)].map((_,i)=>ymd(midnight(Date.now())-(days-1-i)*864e5));
  const withData=dayKeys.filter(k=>pumpDays[k]); const nd=Math.max(1,withData.length);
  if(!bol.length){ st.innerHTML=""; $("#inChart").innerHTML=$("#inTod").innerHTML=""; $("#inList").innerHTML=`<li class="empty">No bolus data in the last ${days} days. Import a full CareLink CSV on the Glucose page and it appears here.</li>`; return; }
  const tot=bol.reduce((a,e)=>a+e.v,0), ctot=carbs.reduce((a,e)=>a+e.v,0);
  st.innerHTML = [["Bolus a day",(tot/nd).toFixed(1),"units"],["Boluses a day",(bol.length/nd).toFixed(1),"doses"],["Pump carbs a day",String(Math.round(ctot/nd)),"grams"],["Days with data",String(withData.length),`of ${days}`]]
    .map(([k,v,n])=>`<div class="wk"><span class="k">${k}</span><b>${v}</b><small>${n}</small></div>`).join("");
  const per = dayKeys.map(k=>({k, u: dayPump(k).filter(e=>e.k===2).reduce((a,e)=>a+e.v,0), has: !!pumpDays[k]}));
  const max = Math.max(1,...per.map(d=>d.u));
  $("#inChart").innerHTML = `<div class="in-bars">${per.map(d=>`<div class="in-bar" title="${new Date(keyToMs(d.k)).toLocaleDateString("en-GB",{weekday:"short",day:"numeric",month:"short"})}: ${d.u.toFixed(1)} U"><span class="in-v">${d.has?d.u.toFixed(1):""}</span><i class="${d.has?"":"none"}" style="height:${d.has?Math.max(3,d.u/max*100):4}%"></i><span class="in-d">${new Date(keyToMs(d.k)).toLocaleDateString("en-GB",{weekday:"narrow"})}</span></div>`).join("")}</div>`;
  setTimeout(()=>{ const c=$("#inChart"); if(c) c.scrollLeft=c.scrollWidth; }, 0);
  const slots=[["Night","00–06",0,6],["Morning","06–11",6,11],["Midday","11–15",11,15],["Afternoon","15–18",15,18],["Evening","18–24",18,24]];
  const sv = slots.map(([n,r,a,b])=>({n,r,u: bol.filter(e=>{const h=new Date(e.t).getHours(); return h>=a&&h<b;}).reduce((x,e)=>x+e.v,0)/nd}));
  const smax=Math.max(0.1,...sv.map(x=>x.u));
  $("#inTod").innerHTML = sv.map(x=>`<div class="tod"><span class="tod-n">${x.n}<small>${x.r}</small></span><span class="tod-bar"><i style="width:${x.u/smax*100}%"></i></span><b>${x.u.toFixed(1)} U</b></div>`).join("");
  $("#inList").innerHTML = bol.slice(0,20).map(e=>{ const c=carbs.find(x=>Math.abs(x.t-e.t)<10*6e4); return `<li><span class="bg-t">${whenText(e.t)}</span><span class="bg-n">${c?`with ${c.v} g carbs`:"correction or no carbs entered"}</span><span class="chip blue num">${e.v.toFixed(2).replace(/0$/,"")} U</span></li>`; }).join("");
}
$("#inRange")?.addEventListener("change", renderInsulin);

/* ---------- opened insulin countdown ---------- */
function renderVial(){
  const el=$("#vialPanel"); if(!el) return; const v=last("vial"), days=num(plan.vialDays)??28;
  const btn = `<button type="button" class="primary" data-vialnew>Opened a new one</button>`;
  if(!v){ el.innerHTML = `<div class="bk-h"><span class="bk-ic" aria-hidden="true">${'<svg class="ic" viewBox="0 0 24 24"><path d="M9 3h6M10 3v3l-4 4v9a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2v-9l-4-4V3"/><path d="M6 13h12"/></svg>'}</span><div><h2>Insulin in use</h2><p class="muted">Tap when you open a new vial or cartridge. The app counts down its ${days} days.</p></div></div><div class="btns">${btn}</div>`; return; }
  const due=v.at+days*864e5, left=(due-Date.now())/864e5, frac=Math.max(0,Math.min(1,left/days)), R=40, C=2*Math.PI*R;
  const cls = left<0 ? "bad" : left<3 ? "warn" : "";
  el.innerHTML = `<div class="vial ${cls}"><div class="rc-ring" style="width:96px;height:96px"><svg viewBox="0 0 96 96" width="96" height="96" aria-hidden="true"><circle cx="48" cy="48" r="${R}" fill="none" stroke="var(--page-soft)" stroke-width="9"/><circle cx="48" cy="48" r="${R}" fill="none" stroke="${cls==="bad"?"var(--low)":cls==="warn"?"var(--high)":"var(--page)"}" stroke-width="9" stroke-linecap="round" stroke-dasharray="${C.toFixed(1)}" stroke-dashoffset="${(C*(1-frac)).toFixed(1)}" transform="rotate(-90 48 48)"/></svg><div class="rc-in"><b>${left<0?"Out":Math.ceil(left)}</b><span>${left<0?"of date":Math.ceil(left)===1?"day left":"days left"}</span></div></div>
    <div class="vial-t"><h2>Insulin in use</h2><p>${left<0?`Past its ${days} days. Start a new one.`:`Opened ${new Date(v.at).toLocaleDateString("en-GB",{weekday:"short",day:"numeric",month:"short"})}. Use by ${new Date(due).toLocaleDateString("en-GB",{weekday:"short",day:"numeric",month:"short"})}.`}</p>${v.note?`<p class="muted">${esc(v.note)}</p>`:""}<div class="btns">${btn}</div></div></div>`;
}
document.addEventListener("click", ev => { if(ev.target.closest("[data-vialnew]")){ openEntry("vial"); $("#e_note").placeholder = "e.g. NovoRapid cartridge, fridge in kitchen"; } });

/* ---------- night mode at night (this phone) ---------- */
const NIGHT_KEY = "t1d-night-auto";
const nightOn = () => { try { return localStorage.getItem(NIGHT_KEY) !== "off"; } catch { return true; } };
function applyNight(){
  const h=new Date().getHours(), night = nightOn() && (h>=21 || h<7);
  if(night) document.documentElement.dataset.theme="dark"; else if(document.documentElement.dataset.theme==="dark") delete document.documentElement.dataset.theme;
  const cb=$("#nightAuto"); if(cb) cb.checked = nightOn();
}
$("#nightAuto")?.addEventListener("change", ev => { try { localStorage.setItem(NIGHT_KEY, ev.target.checked ? "on" : "off"); } catch {} applyNight(); toast(ev.target.checked ? "Night mode on from 9pm to 7am" : "Night mode off"); });
applyNight(); setInterval(applyNight, 60000); document.addEventListener("visibilitychange", () => { if(!document.hidden) applyNight(); });


/* ---------- version ---------- */
const APP_VERSION = "7.4";
document.querySelectorAll(".ver-n").forEach(e => e.textContent = APP_VERSION);
{ const nav=$("#nav"); if(nav && !nav.querySelector(".app-ver")){ const d=document.createElement("div"); d.className="app-ver nav-ver"; d.textContent="Version "+APP_VERSION; nav.insertBefore(d, $("#moreBtn")); } }

/* ---------- carb calculator (in the log dialog) ---------- */
function calcCarbs(){
  const per=num($("#c_per").value), g=num($("#c_g").value), left=num($("#c_left").value)||0;
  const ok = per!=null && g!=null && g>0 && per>=0 && left<g;
  const c = ok ? Math.round(per*(g-left)/100) : null;
  $("#c_out").innerHTML = ok ? `<b>${c} g</b> carbs<small>${per} g per 100 g × ${g-left} g eaten</small>` : (left>=g && g ? "Left over can't be more than he had" : "Fill in the first two");
  $("#c_use").disabled = !ok; return c;
}
["#c_per","#c_g","#c_left"].forEach(id => $(id).addEventListener("input", calcCarbs));
$("#c_use").addEventListener("click", () => { const c=calcCarbs(); if(c==null) return; $("#e_carbs").value=c;
  const per=num($("#c_per").value), g=num($("#c_g").value), left=num($("#c_left").value)||0;
  const tag=`${g-left} g at ${per} g/100 g`; const n=$("#e_note"); if(!n.value.includes(tag)) n.value = (n.value?n.value+" · ":"") + tag;
  $("#carbCalc").open=false; toast(`${c} g carbs filled in`); });

/* ---------- food photo ---------- */
let photoData = undefined;   // undefined = unchanged, null = removed, string = new
async function shrinkImage(file, max=800, q=0.7){
  const url = URL.createObjectURL(file);
  try { const img = await new Promise((res, rej) => { const i=new Image(); i.onload=()=>res(i); i.onerror=rej; i.src=url; });
    const r = Math.min(1, max/Math.max(img.width, img.height)); const c=document.createElement("canvas"); c.width=Math.round(img.width*r); c.height=Math.round(img.height*r);
    c.getContext("2d").drawImage(img, 0, 0, c.width, c.height); return c.toDataURL("image/jpeg", q); }
  finally { URL.revokeObjectURL(url); }
}
function showPhotoPrev(src){ const im=$("#e_photoPrev"); im.hidden=!src; if(src) im.src=src; else im.removeAttribute("src"); $("#e_photoDel").hidden=!src; $("#e_photoLbl").textContent = src ? "Change photo" : "Add photo"; }
$("#e_photo").addEventListener("change", async ev => { const f=ev.target.files[0]; ev.target.value=""; if(!f) return;
  try { photoData = await shrinkImage(f); showPhotoPrev(photoData); } catch { toast("Couldn't read that photo"); } });
$("#e_photoDel").addEventListener("click", () => { photoData = null; showPhotoPrev(null); });
async function savePhoto(entryId){
  if(photoData===undefined) return false;
  const ref = base.collection("photos").doc(entryId);
  if(photoData===null){ await ref.delete().catch(()=>{}); return false; }
  await ref.set({img: photoData, at: Date.now(), by: me}); return true;
}
async function loadPhoto(entryId){ try { const s=await base.collection("photos").doc(entryId).get(); return s.exists ? s.data().img : null; } catch { return null; } }
async function viewPhoto(entryId){
  const d=$("#photoDlg"); $("#photoImg").removeAttribute("src"); $("#photoMsg").textContent="Loading…"; d.showModal();
  const src=await loadPhoto(entryId); if(src){ $("#photoImg").src=src; $("#photoMsg").textContent=""; } else $("#photoMsg").textContent="That photo isn't available.";
}
document.addEventListener("click", ev => { const b=ev.target.closest("[data-photo]"); if(b){ ev.stopPropagation(); viewPhoto(b.dataset.photo); } }, true);
$("#photoClose").addEventListener("click", () => $("#photoDlg").close());

/* ---------- sick-day mode (shared by both phones) ---------- */
let sick = {}; const SICK_EVERY = 2*36e5;
function sickChecks(){ return sick.active ? log.filter(e=>e.at>=sick.start && (e.type==="ketones"||e.type==="bg"||e.ket!=null||e.gl!=null)) : []; }
function sickNextDue(){ const c=sickChecks()[0]; return (c ? c.at : sick.start) + SICK_EVERY; }
function renderSick2(){
  const el=$("#sickMode"), bn=$("#sickBanner"); if(!el) return;
  if(!sick.active){
    el.className="panel sick-off";
    el.innerHTML = `<div class="bk-h"><span class="bk-ic" aria-hidden="true"><svg class="ic" viewBox="0 0 24 24"><path d="M14 14.8V5a2 2 0 1 0-4 0v9.8a4 4 0 1 0 4 0z"/><path d="M12 17.5v-6"/></svg></span><div><h2>Sick-day mode</h2><p class="muted">When he's unwell: a reminder to check ketones and glucose every 2 hours, and one timeline for his team. Both phones see it.</p></div></div><div class="btns"><button type="button" class="primary" id="sickStart">Start sick-day mode</button></div>`;
    $("#sickStart").onclick = async () => { if(!canWrite) return toast("Saving isn't switched on here"); try { await base.collection("meta").doc("sickday").set({active:true, start:Date.now(), by:me}); await base.collection("log").add({type:"ill", at:Date.now(), note:"Sick-day mode started", by:me, gl:null, ket:null, carbs:null}); toast("Sick-day mode on"); } catch { toast("Couldn't start. Try again"); } };
    if(bn) bn.hidden = true; return;
  }
  const due=sickNextDue(), left=due-Date.now(), late=left<=0;
  const tl = log.filter(e=>e.at>=sick.start).slice(0,40);
  const hm = ms => `${Math.floor(ms/36e5)}:${String(Math.floor(ms%36e5/6e4)).padStart(2,"0")}`;
  el.className = "panel sick-on" + (late?" late":"");
  el.innerHTML = `<div class="sick-top"><div><span class="k">Sick-day mode · since ${whenText(sick.start)}</span><b>${late?"Check now":"Next check in "+hm(left)}</b><span class="muted">${late?"Ketones and glucose were due "+(left<-6e4?Math.round(-left/6e4)+" min ago":"now"):"Ketones and glucose every 2 hours"}</span></div></div>
    <div class="btns"><button type="button" class="primary" data-q2s="ketones">Log ketones</button><button type="button" data-q2s="bg">Log BG</button><button type="button" data-q2s="drink">Log a drink</button></div>
    <h3 style="margin:6px 0 0">Timeline</h3>
    <ul class="log sick-tl">${tl.length?tl.map(entryHTML).join(""):`<li class="empty">Nothing logged yet.</li>`}</ul>
    <div class="btns"><button type="button" id="sickEnd">He's better: end sick-day mode</button></div>`;
  el.querySelectorAll("[data-q2s]").forEach(b => b.onclick = () => openEntry(b.dataset.q2s));
  $("#sickEnd").onclick = async () => { try { await base.collection("meta").doc("sickday").set({active:false, start:sick.start, end:Date.now(), by:me}); await base.collection("log").add({type:"ill", at:Date.now(), note:"Sick-day mode ended", by:me, gl:null, ket:null, carbs:null}); toast("Sick-day mode ended"); } catch { toast("Couldn't end. Try again"); } };
  if(bn){ bn.hidden=false; bn.className = "sick-banner"+(late?" late":""); bn.innerHTML = `<b>${late?"Sick day: check ketones and glucose now":"Sick day: next check in "+hm(left)}</b><span>Open</span>`; }
}
let sickBeeped = 0;
setInterval(() => { if(!sick.active) return; renderSick2(); const due=sickNextDue(); if(Date.now()>=due && sickBeeped!==due){ sickBeeped=due; try { navigator.vibrate?.([300,150,300]); } catch {} try { hyBeep(); } catch {} } }, 30000);


/* ---------- day out planner ---------- */
let dayplans = {}, doCur = null, doTimer = null;
const ymdLocal = ms => ymd(ms);
const hm2min = t => { const [h,m]=(t||"0:0").split(":").map(Number); return h*60+(m||0); };
const doDefaultDate = () => { const upcoming = Object.values(dayplans).filter(p=>p.date>=ymdLocal(Date.now())).sort((a,b)=>a.date<b.date?-1:1)[0]; return upcoming ? upcoming.date : ymdLocal(Date.now()+864e5); };
function doPlan(){ const d=$("#do_date").value || doDefaultDate(); return dayplans[d] || {date:d, leave:"09:00", back:"17:00", where:"", items:[], packed:{}}; }
function doSave(plan, quiet){
  dayplans[plan.date]=plan; clearTimeout(doTimer);
  doTimer = setTimeout(async () => { try { await base.collection("dayplans").doc(plan.date).set(plan); if(!quiet) { $("#doSaved").textContent="Saved"; setTimeout(()=>$("#doSaved").textContent="",1500); } } catch { toast("Couldn't save the plan"); } }, 400);
}
function doUsualMeals(plan){
  const L=hm2min(plan.leave), B=hm2min(plan.back), from=Date.now()-30*864e5, by={};
  log.filter(e=>e.type==="food" && e.at>=from && e.meal).forEach(e=>{ (by[e.meal] ||= []).push(e); });
  const out=[];
  for(const [meal, es] of Object.entries(by)){
    if(es.length<2) continue;
    const mins = Math.round(es.reduce((a,e)=>{ const d=new Date(e.at); return a+d.getHours()*60+d.getMinutes(); },0)/es.length/15)*15;
    if(mins<L || mins>B) continue;
    const cnt={}; es.forEach(e=>{ const k=(e.item||"").trim(); if(k) cnt[k]=(cnt[k]||0)+1; });
    const item=Object.entries(cnt).sort((a,b)=>b[1]-a[1])[0]?.[0]; if(!item) continue;
    const fv=findFav(item), last=es.filter(e=>e.item===item).pop();
    out.push({id:Math.random().toString(36).slice(2,9), time:String(Math.floor(mins/60)).padStart(2,"0")+":"+String(mins%60).padStart(2,"0"), kind:"food", name:item+(meal?"":""), meal, carbs: fv?.carbs ?? last?.carbs ?? null, done:false});
  }
  return out.sort((a,b)=>a.time<b.time?-1:1);
}
function doPackList(plan){
  const hrs = Math.max(1,(hm2min(plan.back)-hm2min(plan.leave))/60);
  const tx = P("hypoTx"), food = plan.items.filter(i=>i.kind==="food"), drinks = plan.items.filter(i=>i.kind==="drink");
  const v=last("vial"), vialDue = v ? v.at+(num(plan.vialDays??P("vialDays"))??28)*864e5 : null;
  const list = [
    ["hypo", `Hypo treatment: enough for ${Math.ceil(hrs/3)+2} goes`, tx ? tx.split("\n")[0].slice(0,90) : "What his team told you to give (add it in Care plan)"],
    ["meter", "Blood glucose meter, strips and finger pricker", "In case the CGM and how he seems don't match"],
    ["ketone", "Ketone strips", "Usually the same meter"],
    ["site", "Spare infusion set and reservoir", "In case it comes out"],
    ["insulin", "Insulin", vialDue && vialDue < hm2min(plan.back)*6e4 + keyToMs(plan.date) ? "The opened one runs out today: bring a new one" : "A spare cartridge or vial"],
    ["glucagon", "Glucagon", P("glucagon") || "Check it's in date"],
    ["phone", "Phones charged, and the CGM app working", ""],
  ];
  if(hrs>=5) list.push(["sensor","Sensor overtape or patches",""]);
  food.forEach(i=>list.push(["f-"+i.id, `${i.name}${i.carbs!=null?` (${i.carbs} g)`:""}`, `For ${i.time}`]));
  drinks.forEach(i=>list.push(["d-"+i.id, `${i.name}${i.ml?` ${i.ml} ml`:""}`, `For ${i.time}`]));
  list.push(["water","Water", ""]);
  return list;
}
function doHeads(plan){
  const start=keyToMs(plan.date)+hm2min(plan.leave)*6e4, end=keyToMs(plan.date)+hm2min(plan.back)*6e4, out=[];
  const s=lastSite(); if(s){ const due=s.at+(num(plan.siteDays??P("siteDays"))??3)*864e5; if(due<=end) out.push(due<start?"Infusion set change is due before you leave: change it first":"Infusion set change falls during the day out: bring a spare and plan when"); }
  const sn=last("sensor"); if(sn){ const due=sn.at+(num(P("sensorDays"))??10)*864e5; if(due<=end) out.push(due<start?"Sensor change is due before you leave":"Sensor change falls during the day out: bring a new sensor"); }
  const v=last("vial"); if(v){ const due=v.at+(num(P("vialDays"))??28)*864e5; if(due<=end) out.push("Opened insulin reaches its last day: bring a new one"); }
  document.querySelectorAll("#watchList li:not(.muted)").forEach(li=>out.push(li.textContent.trim()));
  return out;
}
function renderDayOut(){
  const box=$("#doList"); if(!box) return;
  if(!$("#do_date").value) $("#do_date").value = doDefaultDate();
  const plan=doPlan(); doCur=plan;
  if(document.activeElement?.id!=="do_where") $("#do_where").value=plan.where||"";
  if(document.activeElement?.id!=="do_leave") $("#do_leave").value=plan.leave||"09:00";
  if(document.activeElement?.id!=="do_back") $("#do_back").value=plan.back||"17:00";
  const H=doHeads(plan);
  $("#doHeads").innerHTML = H.length ? `<div class="do-heads"><b>Heads up</b><ul class="plain">${H.map(h=>`<li>${esc(h)}</li>`).join("")}</ul></div>` : "";
  const items=plan.items.slice().sort((a,b)=>a.time<b.time?-1:1);
  const carbs=items.filter(i=>i.kind==="food").reduce((a,i)=>a+(i.carbs||0),0), ml=items.filter(i=>i.kind==="drink").reduce((a,i)=>a+(i.ml||0),0);
  $("#doTotals").textContent = items.length ? `${carbs} g carbs${ml?` · ${ml} ml`:""}` : "";
  box.innerHTML = items.length ? items.map(i=>`<li class="${i.done?"done":""}"><span class="do-t">${esc(i.time)}</span><span class="do-n"><span class="chip ${i.kind==="drink"?"drink":"food"}">${i.kind==="drink"?"Drink":esc(i.meal||"Food")}</span> ${esc(i.name)}</span><span class="do-c">${i.carbs!=null?i.carbs+" g":""}${i.ml?`<small>${i.ml} ml</small>`:""}</span>
      <span class="do-acts">${i.done?`<span class="chip ok">Had it</span>`:`<button type="button" class="small primary" data-dohad="${i.id}">Had it</button>`}<button type="button" class="small" data-dodel="${i.id}" aria-label="Remove ${esc(i.name)}">Remove</button></span></li>`).join("")
    : `<li class="empty">Nothing planned yet. Tap <b>Add his usual meals</b> to fill in what he usually has between ${esc(plan.leave)} and ${esc(plan.back)}.</li>`;
  const pk=doPackList(plan), packed=plan.packed||{}; const n=pk.filter(([k])=>packed[k]).length;
  $("#doPackCount").textContent = `${n} of ${pk.length} packed`;
  $("#doPack").innerHTML = pk.map(([k,t,sub])=>`<li><label class="do-chk ${packed[k]?"on":""}"><input type="checkbox" data-dopack="${esc(k)}" ${packed[k]?"checked":""}><span><b>${esc(t)}</b>${sub?`<small>${esc(sub)}</small>`:""}</span></label></li>`).join("");
}
["#do_where","#do_leave","#do_back"].forEach(id => $(id).addEventListener("input", () => { const p=doPlan(); p.where=$("#do_where").value; p.leave=$("#do_leave").value||"09:00"; p.back=$("#do_back").value||"17:00"; doSave(p); renderDayOut(); }));
$("#do_date").addEventListener("change", renderDayOut);
let doPending = null;
dlg.addEventListener("close", () => { const pd=doPending; if(!pd) return; doPending=null;
  setTimeout(() => { if(log.some(e=>e.item===pd.name && e.at>=pd.t0)){ const p=dayplans[pd.date]; const it=p?.items.find(i=>i.id===pd.id); if(it){ it.done=true; doSave(p,true); renderDayOut(); toast("Logged and ticked off the plan"); } } }, 700); });
$("#doUsual").addEventListener("click", () => { const p=doPlan(); const add=doUsualMeals(p).filter(n=>!p.items.some(i=>i.name===n.name && i.time===n.time));
  if(!add.length) return toast("No usual meals found in that time. Log meals for a few days, or add them yourself");
  p.items=[...p.items, ...add]; doSave(p); renderDayOut(); toast(`Added ${add.length} usual meal${add.length===1?"":"s"}`); });
$("#doAdd").addEventListener("click", () => { $("#doForm").reset(); $("#di_time").value = doPlan().leave || "10:00"; doKind(); $("#doDlg").showModal(); });
function doKind(){ document.querySelectorAll("[data-dk]").forEach(el=>el.hidden = el.dataset.dk!==$("#di_kind").value); }
$("#di_kind").addEventListener("change", doKind);
$("#di_name").addEventListener("change", () => { const fv=findFav($("#di_name").value); if(fv){ if($("#di_carbs").value==="" && fv.carbs!=null) $("#di_carbs").value=fv.carbs; if(fv.kind==="drink"){ $("#di_kind").value="drink"; doKind(); if(fv.ml!=null) $("#di_ml").value=fv.ml; } } });
$("#diCancel").addEventListener("click", () => $("#doDlg").close());
$("#doForm").addEventListener("submit", ev => { ev.preventDefault(); const p=doPlan();
  p.items=[...p.items, {id:Math.random().toString(36).slice(2,9), time:$("#di_time").value, kind:$("#di_kind").value, name:$("#di_name").value.trim(), carbs:num($("#di_carbs").value), ml:$("#di_kind").value==="drink"?num($("#di_ml").value):null, done:false}];
  doSave(p); $("#doDlg").close(); renderDayOut(); });
document.addEventListener("click", ev => {
  const h=ev.target.closest("[data-dohad]"); if(h){ const p=doPlan(), it=p.items.find(i=>i.id===h.dataset.dohad); if(!it) return;
    openEntry(it.kind); $("#e_item").value=it.name; if(it.carbs!=null) $("#e_carbs").value=it.carbs; if(it.ml!=null) $("#e_ml").value=it.ml; if(it.meal && it.kind==="food") $("#e_meal").value=it.meal;
    doPending = {date:p.date, id:it.id, name:it.name, t0:Date.now()-1000}; return; }
  const d=ev.target.closest("[data-dodel]"); if(d){ const p=doPlan(); p.items=p.items.filter(i=>i.id!==d.dataset.dodel); doSave(p); renderDayOut(); }
});
document.addEventListener("change", ev => { const c=ev.target.closest("[data-dopack]"); if(!c) return; const p=doPlan(); p.packed={...(p.packed||{}), [c.dataset.dopack]: c.checked}; doSave(p,true); renderDayOut(); });


/* ---------- when to keep an eye on him (last 14 days) ---------- */
function watchProfile(){
  const lo=hypoBelow(), blocks=[...Array(12)].map((_,b)=>({b, low:0, high:0, n:0}));
  const pts=rangePoints(14), days=[...new Set(pts.map(p=>ymd(p.t)))];
  if(days.length>=5){
    days.forEach(k=>{ const dp=dayPoints(k); blocks.forEach(B=>{ const inb=dp.filter(p=>{const h=new Date(p.t).getHours(); return h>=B.b*2 && h<B.b*2+2;}); if(!inb.length) return; B.n++;
      if(inb.filter(p=>p.v<lo).length>=3) B.low++; if(inb.filter(p=>p.v>13.9).length>=6) B.high++; }); });
    return {src:`CGM, last ${days.length} days`, days:days.length, blocks};
  }
  const bg = allBgChecks(14); const bdays=[...new Set(bg.map(x=>ymd(x.t)))];
  if(bg.length>=6){
    bg.forEach(x=>{ const B=blocks[Math.floor(new Date(x.t).getHours()/2)]; B.n++; if(x.v<lo) B.low++; if(x.v>13.9) B.high++; });
    return {src:`${bg.length} BG checks, last 14 days`, days:bdays.length, blocks, fromChecks:true};
  }
  return null;
}
const blockLabel = b => `${String(b*2).padStart(2,"0")}:00–${String((b*2+2)%24).padStart(2,"0")}:00`;
function watchWindows(prof){
  if(!prof) return [];
  const minN = prof.fromChecks ? 2 : Math.max(2, Math.ceil(prof.days*0.25));
  const out=[];
  prof.blocks.forEach(B=>{ if(B.low>=minN) out.push({b:B.b, kind:"low", n:B.low, of:prof.fromChecks?B.n:prof.days}); if(B.high>=Math.max(minN, prof.fromChecks?2:Math.ceil(prof.days*0.35))) out.push({b:B.b, kind:"high", n:B.high, of:prof.fromChecks?B.n:prof.days}); });
  return out.sort((a,b)=>(b.kind==="low")-(a.kind==="low") || b.n-a.n);
}
function nextWatch(){
  const W=watchWindows(watchProfile()); if(!W.length) return null;
  const now=new Date(), mins=now.getHours()*60+now.getMinutes();
  const withStart = W.map(w=>{ let start=w.b*120, end=start+120; let inNow = mins>=start && mins<end; let wait = inNow ? 0 : (start-mins+1440)%1440; return {...w, inNow, wait}; });
  return withStart.sort((a,b)=>a.wait-b.wait || (b.kind==="low")-(a.kind==="low"))[0];
}
function renderWatch(){
  const box=$("#watchBody"); if(!box) return; const prof=watchProfile();
  if(!prof){ $("#watchSrc").textContent=""; box.innerHTML=`<div class="empty">Needs at least 5 days of CGM data (import his CareLink file above), or a few days of logged BG checks.</div>`; return; }
  $("#watchSrc").textContent = prof.src;
  const W=watchWindows(prof), hNow=new Date().getHours(), denom=b=>prof.fromChecks?Math.max(1,b.n):prof.days;
  const cell = B => { const lf=B.low/denom(B), hf=B.high/denom(B); const isLow=W.some(w=>w.b===B.b&&w.kind==="low"), isHigh=W.some(w=>w.b===B.b&&w.kind==="high");
    const bg = isLow ? `color-mix(in srgb,var(--c-red) ${Math.round(35+lf*65)}%,var(--surface))` : isHigh ? `color-mix(in srgb,var(--c-orange) ${Math.round(35+hf*65)}%,var(--surface))` : (lf>0||hf>0) ? "var(--sunk)" : "var(--ok-soft)";
    return `<div class="wt-cell ${isLow?"low":isHigh?"high":""} ${Math.floor(hNow/2)===B.b?"now":""}" style="background:${bg}" title="${blockLabel(B.b)}: low on ${B.low}, high on ${B.high}"><span>${String(B.b*2).padStart(2,"0")}</span></div>`; };
  const nx=nextWatch();
  box.innerHTML = `<div class="wt-strip" role="img" aria-label="24 hours in 2-hour blocks. ${W.map(w=>(w.kind==="low"?"Lows ":"Highs ")+blockLabel(w.b)).join(", ")||"No repeating problem times"}">${prof.blocks.map(cell).join("")}</div>
    <div class="wt-legend"><span><i style="background:var(--c-red)"></i>Often low</span><span><i style="background:var(--c-orange)"></i>Often high</span><span><i style="background:var(--ok-soft);border:1px solid var(--line)"></i>Usually fine</span><span><i class="nowdot"></i>Now</span></div>
    ${nx?`<div class="wt-next ${nx.kind}"><b>${nx.inNow?"Now":"Next"}: ${nx.kind==="low"?"lows":"highs"} ${blockLabel(nx.b)}</b><span>${nx.inNow?"This is one of his watch times. Glance at his CGM.":"In "+(nx.wait>=60?Math.floor(nx.wait/60)+" h "+(nx.wait%60?nx.wait%60+" min":""):nx.wait+" min")}</span></div>`:""}
    ${W.length?`<ul class="wt-list">${W.map(w=>`<li><span class="chip ${w.kind==="low"?"low":"high"}">${w.kind==="low"?"Low":"High"}</span><b>${blockLabel(w.b)}</b><span class="muted">on ${w.n} of ${w.of} ${prof.fromChecks?"checks":"days"}</span></li>`).join("")}</ul>
      <p class="bk-note">${W.some(w=>w.kind==="low")?"Around the low times: keep hypo treatment close and glance at his CGM. ":""}${W.some(w=>w.kind==="high")?"Around the high times: have ketone strips ready if it stays high. ":""}Patterns from what already happened, not predictions. Raise repeating ones with his team.</p>`
     :`<p class="bk-note">No repeating problem times in this period.</p>`}`;
}

/* ---------- weekly summary ---------- */
function windowStats(fromMs, toMs){
  const pts=[]; for(let t=midnight(fromMs); t<toMs; t+=864e5) pts.push(...dayPoints(ymd(t)).filter(p=>p.t>=fromMs && p.t<toMs));
  const ev = log.filter(e=>e.at>=fromMs && e.at<toMs);
  const foodDays = new Set(ev.filter(e=>e.type==="food").map(e=>dayKey(e.at)));
  const carbs = ev.filter(e=>e.type==="food" && e.carbs!=null).reduce((a,e)=>a+e.carbs,0);
  const out = {n:pts.length, hyposLogged: ev.filter(e=>e.type==="hypo").length, carbsDay: foodDays.size ? Math.round(carbs/foodDays.size) : null, foodDays: foodDays.size};
  if(pts.length){ const inR=pts.filter(p=>p.v>=3.9&&p.v<=10).length, lo=pts.filter(p=>p.v<3.9).length; out.tir=inR/pts.length*100; out.low=lo/pts.length*100; out.mean=pts.reduce((a,p)=>a+p.v,0)/pts.length; out.cgmHypos=hypoEpisodes(pts).length; }
  return out;
}
function weekSummary(){
  const end = Date.now(), a = windowStats(end-7*864e5, end), b = windowStats(end-14*864e5, end-7*864e5);
  const rows = mealOutcomes(7); const g={}; rows.forEach(r=>(g[r.key] ||= []).push(r));
  const meals = Object.values(g).map(x=>({item:x[0].item, n:x.length, bad:x.filter(r=>r.high||r.low).length})).filter(m=>m.n>=2);
  const best = meals.filter(m=>m.bad===0).sort((x,y)=>y.n-x.n)[0], worst = meals.filter(m=>m.bad/m.n>=0.5).sort((x,y)=>y.bad/y.n-x.bad/x.n)[0];
  return {a, b, best, worst};
}
function renderWeek(){
  const box=$("#weekPanel"); if(!box) return;
  const {a,b,best,worst} = weekSummary();
  if(!a.n && !a.foodDays && !a.hyposLogged){ box.hidden=true; return; }
  box.hidden=false;
  const d = (x,y,unit="",good="up") => (x==null||y==null) ? "" : (()=>{ const diff=x-y; if(Math.abs(diff)<0.5) return `<span class="wk-d same">same as last week</span>`; const better = good==="up" ? diff>0 : diff<0; return `<span class="wk-d ${better?"up":"down"}">${diff>0?"▲":"▼"} ${Math.abs(diff).toFixed(unit==="%"?0:1)}${unit} vs last week</span>`; })();
  const isSun = new Date().getDay()===0;
  box.innerHTML = `<div style="display:flex;justify-content:space-between;align-items:baseline;gap:8px;flex-wrap:wrap"><h2>${isSun?"His week":"Last 7 days"}</h2><span class="muted" style="font-size:.85rem">${new Date(Date.now()-6*864e5).toLocaleDateString("en-GB",{day:"numeric",month:"short"})} – today</span></div>
    <div class="wk-grid">
      <div class="wk"><span class="k">In range</span><b>${a.tir!=null?Math.round(a.tir)+"%":"—"}</b>${d(a.tir,b.tir,"%","up")}</div>
      <div class="wk"><span class="k">Hypos</span><b>${a.cgmHypos!=null?a.cgmHypos:a.hyposLogged}</b>${a.cgmHypos!=null?d(a.cgmHypos,b.cgmHypos,"","down"):d(a.hyposLogged,b.hyposLogged,"","down")}<small>${a.cgmHypos!=null?"from CGM":"logged"}</small></div>
      <div class="wk"><span class="k">Average</span><b>${a.mean!=null?a.mean.toFixed(1):"—"}</b>${a.mean!=null?`<small>mmol/L</small>`:""}</div>
      <div class="wk"><span class="k">Carbs a day</span><b>${a.carbsDay!=null?a.carbsDay+" g":"—"}</b>${a.foodDays?`<small>${a.foodDays} day${a.foodDays===1?"":"s"} logged</small>`:""}</div>
    </div>
    ${best||worst?`<div class="wk-meals">${best?`<div><span class="chip ok">Went well</span> ${esc(best.item)} <span class="muted">(${best.n}×, no highs or lows)</span></div>`:""}${worst?`<div><span class="chip high">Harder</span> ${esc(worst.item)} <span class="muted">(${worst.bad} of ${worst.n} times out of range)</span></div>`:""}</div>`:""}
    ${a.n?"":`<p class="muted" style="margin:0;font-size:.85rem">Import his CGM file to see time in range here.</p>`}`;
}

/* ---------- setup check ---------- */
let pushSubbed = null, offsite = null, remindersRun = null, pushMeta = {};
async function pushReg(){ try { return ("serviceWorker" in navigator) ? await navigator.serviceWorker.getRegistration() : null; } catch { return null; } }
async function checkPushSub(){ const r=await pushReg(); try { const s=await r?.pushManager?.getSubscription(); pushSubbed = !!s; } catch { pushSubbed=false; } renderSetup(); renderPush(); }
function setupItems(){
  const F = window.__family || null, inst = matchMedia("(display-mode: standalone)").matches || navigator.standalone;
  const lastCgm = Object.keys(cgmDays).sort().pop();
  const bkAge = lastBackup ? (Date.now()-lastBackup)/36e5 : null;
  return [
    {ok: !!F, t:"Family app", d: F ? "Signed in and sharing with the family" : "This copy isn't the family app, so nothing is shared", href:"#backup", skip: !F},
    {ok: F ? F.members>=2 : null, t:"His mum has joined", d: F ? (F.members>=2 ? `${F.members} people in the family` : "Only you so far. Share the family code from Care plan") : "Family app only", href:"#plan"},
    {ok: F ? F.signupsClosed : null, t:"New sign-ups closed", d: F ? (F.signupsClosed===true ? "No one else can create an account" : F.signupsClosed===false ? "Anyone with the link can register. Turn off sign-ups in Supabase (Authentication → Sign In / Providers)" : "Couldn't check") : "Family app only"},
    {ok: bkAge!=null && bkAge<36, t:"Daily backup", d: bkAge==null ? "No backup yet" : bkAge<36 ? `Last backup ${ago(lastBackup)}` : `Last backup ${ago(lastBackup)}. Check backup-setup.sql was run in Supabase`, href:"#backup"},
    {ok: offsite ? (Date.now()-offsite)/36e5<36 : (F?false:null), t:"Off-site backup (GitHub)", d: offsite ? `Last copy ${ago(offsite)}` : F ? "Not seen yet. Add daily-backup.yml and its 3 secrets in GitHub" : "Family app only"},
    {ok: pushSubbed, t:"Notifications on this phone", d: pushSubbed ? "This phone gets reminders" : "Turn them on below", href:"#pushPanel"},
    {ok: remindersRun ? (Date.now()-remindersRun)/36e5<30 : (F?false:null), t:"Reminder job (GitHub)", d: remindersRun ? `Last ran ${ago(remindersRun)}` : F ? "Not seen yet. Add reminders.yml and the VAPID_PRIVATE_KEY secret in GitHub" : "Family app only"},
    {ok: !!inst, t:"Installed on this phone", d: inst ? "Opens like an app" : "iPhone: Safari → Share → Add to Home Screen. Android: menu → Install app"},
  ].filter(x=>!x.skip);
}
function renderSetup(){
  const L=$("#setupList"); if(!L) return;
  const items=setupItems(), done=items.filter(i=>i.ok===true).length, rel=items.filter(i=>i.ok!==null).length;
  $("#setupScore").textContent = `${done} of ${rel} done`;
  { const b=$("#bdgSetup"); if(b){ b.hidden = done>=rel; b.textContent = `${done} of ${rel} done`; } }
  L.innerHTML = items.map(i=>`<li class="${i.ok===true?"ok":i.ok===null?"na":"todo"}">${i.href&&i.ok!==true?`<a href="${i.href}">`:"<div>"}<span class="su-ic">${i.ok===true?"✓":i.ok===null?"–":"!"}</span><span class="su-t"><b>${esc(i.t)}</b><small>${esc(i.d)}</small></span>${i.href&&i.ok!==true?"</a>":"</div>"}</li>`).join("");
}
try { navigator.serviceWorker?.ready.then(() => checkPushSub()); } catch {}
addEventListener("hashchange", () => { if(location.hash==="#setup") checkPushSub(); });
$("#setupRefresh").addEventListener("click", () => { checkPushSub(); window.__familyRefresh?.(); renderSetup(); toast("Checked"); });

/* ---------- phone notifications (web push) ---------- */
const b64u = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g,"-").replace(/\//g,"_").replace(/=+$/,"");
const unb64u = s => Uint8Array.from(atob(s.replace(/-/g,"+").replace(/_/g,"/")+"===".slice((s.length+3)%4)), c=>c.charCodeAt(0));
let newPrivKey = null;
async function renderPush(){
  const box=$("#pushBody"); if(!box) return;
  const reg = await pushReg(); const ios = /iPhone|iPad/.test(navigator.userAgent), inst = matchMedia("(display-mode: standalone)").matches || navigator.standalone;
  if(!window.__family || !reg || !("PushManager" in window)){
    box.innerHTML = `<div class="empty">${ios && !inst ? "On iPhone, notifications only work once the app is on the Home Screen: open it in Safari → Share → Add to Home Screen, then open it from there." : "Notifications work in the installed family app (on your Vercel link), not in this copy."}</div>`; return; }
  if(Notification.permission==="denied"){ box.innerHTML = `<div class="band high"><strong>Notifications are blocked for this app</strong><span>Allow them in the phone's settings for this app, then come back.</span></div>`; return; }
  if(newPrivKey){
    box.innerHTML = `<div class="band high"><strong>One-time step: save this key in GitHub</strong><span>GitHub → your repo → Settings → Secrets and variables → Actions → New repository secret. Name: <code>VAPID_PRIVATE_KEY</code>. Value: the key below. It's shown only once and isn't stored in the app.</span></div>
      <div class="keybox"><code id="vapidKey">${newPrivKey}</code><button type="button" class="small" id="copyKey">Copy</button></div>
      <button type="button" class="primary" id="keyDone">I've saved it in GitHub</button>`;
    $("#copyKey").onclick = async () => { try { await navigator.clipboard.writeText(newPrivKey); toast("Copied"); } catch { toast("Select the key and copy it"); } };
    $("#keyDone").onclick = () => { newPrivKey=null; renderPush(); };
    return;
  }
  box.innerHTML = pushSubbed
    ? `<div class="band ok"><strong>On for this phone</strong><span>Reminders arrive in the morning and evening. Set them up on his mum's phone too.</span></div><div class="btns"><button type="button" id="pushTest">Send a test now</button><button type="button" id="pushOff">Turn off on this phone</button></div>`
    : `<button type="button" class="primary" id="pushOn">Turn on notifications on this phone</button>${pushMeta.publicKey?"":`<p class="muted" style="font-size:.85rem;margin:0">The first time, the app makes a secret key for GitHub to send the reminders with. You'll copy it once.</p>`}`;
  $("#pushOn")?.addEventListener("click", pushOn); $("#pushOff")?.addEventListener("click", pushOff);
  $("#pushTest")?.addEventListener("click", async () => { const r=await pushReg(); r?.showNotification("T1D Care Hub", {body:"Test: notifications work on this phone.", icon:"/icon-192.png", tag:"t1d-test"}); });
}
async function pushOn(){
  try {
    const perm = await Notification.requestPermission(); if(perm!=="granted") return renderPush();
    if(!pushMeta.publicKey){
      const kp = await crypto.subtle.generateKey({name:"ECDSA", namedCurve:"P-256"}, true, ["sign","verify"]);
      const pub = b64u(await crypto.subtle.exportKey("raw", kp.publicKey)), jwk = await crypto.subtle.exportKey("jwk", kp.privateKey);
      await base.collection("meta").doc("push").set({publicKey: pub, at: Date.now(), by: me}); pushMeta = {publicKey: pub}; newPrivKey = jwk.d;
    }
    const reg = await pushReg();
    const sub = await reg.pushManager.subscribe({userVisibleOnly:true, applicationServerKey: unb64u(pushMeta.publicKey)});
    const j = sub.toJSON(); const id = b64u(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(j.endpoint))).slice(0,24);
    await base.collection("push").doc(id).set({endpoint:j.endpoint, keys:j.keys, by:me, device:(navigator.userAgent.match(/iPhone|iPad|Android|Windows|Mac/)||["phone"])[0], at:Date.now()});
    pushSubbed = true; toast("Notifications on"); renderPush(); renderSetup();
  } catch(e){ toast("Couldn't turn on notifications: " + (e.message||e)); }
}
async function pushOff(){
  try { const r=await pushReg(); const s=await r?.pushManager.getSubscription(); if(s){ const id=b64u(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s.endpoint))).slice(0,24); await s.unsubscribe(); await base.collection("push").doc(id).delete().catch(()=>{}); } pushSubbed=false; renderPush(); renderSetup(); toast("Notifications off on this phone"); } catch { toast("Couldn't turn off"); }
}

/* ===== design v3 helpers ===== */
const VIEW_TITLES = {glucose:["Glucose","His CGM data, patterns and what they mean"], food:["Food & drink","Meals, drinks and how they affect his glucose"], log:["History","Everything you've logged, and the clinic summary"],
  supplies:["Supplies","Stock at home and the next delivery"], checks:["Growth & check-ups","Height, weight, eyes and teeth"], appts:["Appointments","GP, clinic, vaccinations and more"],
  reports:["Reports","Round-ups to keep or share"], sites:["Sites & sensors","Log each change and where it went, so spots get rotated"], handover:["Handover","His sheet for anyone looking after him, and a record of each handover"], sick:["Sick day","What to do when he's unwell"], plan:["Care plan","His details, targets and contacts"], backup:["Backup","Keep a copy of everything"], setup:["Setup check","Is everything switched on and working?"], insulin:["Insulin","From his pump file: what the pump gave"], dayout:["Day out","Plan food, drinks and the bag before you go"]};
for (const [k,[t,sub]] of Object.entries(VIEW_TITLES)) { const sec = document.getElementById("v-"+k); const ic = document.querySelector(`#nav a[href="#${k}"] .ic`); if(!sec || sec.querySelector(".view-head")) continue;
  const h = document.createElement("div"); h.className = "view-head"; h.innerHTML = `<span class="vh-ic">${ic?ic.outerHTML:""}</span><div><h1>${t}</h1><p>${sub}</p></div>`; sec.prepend(h);
  const bk = document.createElement("button"); bk.type = "button"; bk.className = "vh-back"; bk.setAttribute("aria-label","Back"); bk.innerHTML = `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M15 5l-7 7 7 7"/></svg><span>Back</span>`; bk.addEventListener("click", goBack); sec.prepend(bk); }
function labelCardTables(){ document.querySelectorAll("table.cards").forEach(t => { const hs=[...t.querySelectorAll("thead th")].map(th=>th.textContent.trim()); t.querySelectorAll("tbody tr").forEach(tr => [...tr.children].forEach((td,i)=>{ if(hs[i] && !td.hasAttribute("colspan")) td.dataset.label = hs[i]; })); }); }
new MutationObserver(() => labelCardTables()).observe(document.querySelector(".wrap"), {childList:true, subtree:true});
function friendlyLine(){
  const bits=[]; const tp = dayPoints(ymd(Date.now())); const s1 = tp.length>12 ? statsOf(tp) : null;
  if(s1) bits.push(`In range ${Math.round(s1.tir)}% so far today`);
  const na = (typeof appts!=="undefined") ? appts.filter(a=>a.status==="booked" && a.at>=midnight(Date.now())).sort((a,b)=>a.at-b.at)[0] : null;
  if(na){ const d=Math.round((midnight(na.at)-midnight(Date.now()))/864e5); bits.push(`${(AK[na.kind]||AK.other).label} ${d===0?"today":d===1?"tomorrow":"in "+d+" days"}`); }
  const short = supplyRows().filter(x=>x.status==="Reorder now"); if(short.length) bits.push(`${short.length} suppl${short.length===1?"y":"ies"} to order`);
  return bits.length ? bits.join(" · ") : "Log meals and changes as you go. It all adds up for his clinic.";
}

/* ================= REPORTS (home round-up + PDF) ================= */
function rpRange(){
  const v = $("#rpPeriod").value, today = midnight(Date.now());
  document.querySelectorAll("[data-rp]").forEach(el => el.hidden = v!=="custom");
  if(v==="custom"){ const f=$("#rpFrom").value, t=$("#rpTo").value; if(f && t){ const a=dateMs(f), b=dateMs(t); return [midnight(Math.min(a,b)), midnight(Math.max(a,b))+864e5-1]; } }
  if(v==="month"){ const d=new Date(today); const first=new Date(d.getFullYear(), d.getMonth()-1, 1).getTime(), last=new Date(d.getFullYear(), d.getMonth(), 1).getTime()-1; return [first, last]; }
  const n = v==="custom" ? 7 : +v; return [today-(n-1)*864e5, today+864e5-1];
}
function ptsBetween(a,b){ const out=[]; for(let d=midnight(a); d<=b; d+=864e5) out.push(...dayPoints(ymd(d))); return out.filter(p=>p.t>=a && p.t<=b); }
function pumpBetween(a,b){ const out=[]; for(let d=midnight(a); d<=b; d+=864e5) out.push(...dayPump(ymd(d))); return out.filter(e=>e.t>=a && e.t<=b); }
function statsOf(pts){ if(!pts.length) return null; const n=pts.length, mean=pts.reduce((x,p)=>x+p.v,0)/n; const b={vlow:0,low:0,in:0,high:0,vhigh:0}; pts.forEach(p=>b[band(p.v)]++);
  const sd=Math.sqrt(pts.reduce((x,p)=>x+(p.v-mean)**2,0)/n); return {n, mean, tir:pct(b.in,n), below:pct(b.low+b.vlow,n), above:pct(b.high+b.vhigh,n), cv:sd/mean*100, b}; }
function arrow(now, before, goodUp){ if(now==null || before==null) return ""; const d=now-before; if(Math.abs(d)<0.5) return `<span class="rp-same">same as before</span>`;
  const good = goodUp ? d>0 : d<0; return `<span class="${good?"rp-up":"rp-down"}">${d>0?"▲":"▼"} ${Math.abs(d).toFixed(Math.abs(d)<10?1:0)} vs previous</span>`; }
function buildReport(){
  const [a,b] = rpRange(), len = b-a+1, pa = a-len, pb = a-1, days = Math.round(len/864e5);
  const name = P("name") || "His";
  const pts = ptsBetween(a,b), prev = ptsBetween(pa,pb), st = statsOf(pts), sp = statsOf(prev);
  const L = log.filter(e=>e.at>=a && e.at<=b), Lp = log.filter(e=>e.at>=pa && e.at<=pb);
  const eps = hypoEpisodes(pts), epsP = hypoEpisodes(prev);
  const foods = L.filter(e=>e.type==="food"), drinks = L.filter(e=>e.type==="drink");
  const foodDays = new Set(foods.map(e=>ymd(e.at))).size || 1;
  const carbsDay = foods.length ? Math.round(foods.reduce((x,e)=>x+(e.carbs||0),0)/foodDays) : null;
  const pev = pumpBetween(a,b), pc = pev.filter(e=>e.k===1), pi = pev.filter(e=>e.k===2), psite = pev.filter(e=>e.k===4);
  const pumpDaysN = new Set(pev.map(e=>ymd(e.t))).size || 1;
  const fmtR = ms => new Date(ms).toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"});
  // per-day bars
  const dayRows = []; for(let d=midnight(a); d<=b; d+=864e5){ const s1 = statsOf(dayPoints(ymd(d))); dayRows.push({d, s:s1}); }
  const W=700, Hh=150, l=34, bw=Math.max(6, Math.min(40, (W-l-10)/dayRows.length - 4));
  let bars = `<svg viewBox="0 0 ${W} ${Hh+30}" width="100%" role="img" aria-label="Time in range each day">`;
  [0,50,70,100].forEach(v=>{ const y=10+(1-v/100)*Hh; bars+=`<line x1="${l}" x2="${W}" y1="${y}" y2="${y}" stroke="${v===70?"#2b8a52":"#dbe4e1"}" ${v===70?'stroke-dasharray="4 4"':""}/><text x="${l-6}" y="${y+4}" text-anchor="end" font-size="11" fill="#5a6b66">${v}%</text>`; });
  dayRows.forEach((r,i)=>{ const x = l+6+i*((W-l-10)/dayRows.length); if(!r.s){ bars+=`<text x="${x+bw/2}" y="${10+Hh-4}" text-anchor="middle" font-size="10" fill="#9aa">–</text>`; }
    else { let y=10+Hh; [["vlow","#7d1611"],["low","#c0322b"],["in","#2b8a52"],["high","#a86100"],["vhigh","#d97a00"]].forEach(([k,c])=>{ const h=pct(r.s.b[k],r.s.n)/100*Hh; if(h>0){ y-=h; bars+=`<rect x="${x}" y="${y}" width="${bw}" height="${h}" fill="${c}"/>`; } }); }
    if(dayRows.length<=14 || i%Math.ceil(dayRows.length/14)===0) bars+=`<text x="${x+bw/2}" y="${Hh+26}" text-anchor="middle" font-size="10" fill="#5a6b66">${new Date(r.d).toLocaleDateString("en-GB",{day:"numeric",month:dayRows.length>14?"short":undefined, weekday:dayRows.length<=7?"short":undefined})}</text>`; });
  bars += `</svg>`;
  const good = dayRows.filter(r=>r.s).sort((x,y)=>y.s.tir-x.s.tir);
  // meals
  const mo = mealOutcomes(Math.max(days, Math.ceil((Date.now()-a)/864e5))).filter(r=>r.at>=a && r.at<=b); const g={}; mo.forEach(r=>(g[r.key] ||= []).push(r));
  const mealsT = Object.values(g).filter(x=>x.length>=2).map(x=>({item:x[0].item, n:x.length, highs:x.filter(y=>y.high).length, lows:x.filter(y=>y.low).length, rise:x.reduce((s,y)=>s+y.rise,0)/x.length}));
  const smooth = mealsT.filter(m=>!m.lows && m.highs/m.n<=0.34).sort((x,y)=>x.rise-y.rise).slice(0,3), hard = mealsT.filter(m=>m.lows || m.highs/m.n>=0.5).sort((x,y)=>(y.lows+y.highs)/y.n-(x.lows+x.highs)/x.n).slice(0,3);
  const refused = foods.filter(e=>e.ate==="Refused"||e.ate==="A little").length;
  const notes = L.filter(e=>e.type==="note"||e.type==="ill").sort((x,y)=>x.at-y.at);
  const hyposLogged = L.filter(e=>e.type==="hypo").length, hyposLoggedP = Lp.filter(e=>e.type==="hypo").length;
  const upcoming = appts.filter(x=>x.status==="booked" && x.at>=midnight(Date.now()) && x.at<=Date.now()+30*864e5).sort((x,y)=>x.at-y.at);
  const toOrder = supplyRows().filter(x=>x.status!=="OK"); const nd = nextDelivery();
  const card = (k,v,n,extra="") => `<div class="rp-card"><div class="rp-k">${k}</div><div class="rp-v">${v}</div><div class="rp-n">${n}</div>${extra?`<div class="rp-n">${extra}</div>`:""}</div>`;
  const title = days<=7 ? `${name==="His"?"His":name+"'s"} week` : days<=31 ? `${name==="His"?"His":name+"'s"} month` : `${name==="His"?"His":name+"'s"} report`;
  const hi = [];
  if(st){ hi.push(st.tir>=70 ? `In range <b>${Math.round(st.tir)}%</b> of the time, meeting the usual target of over 70%.` : `In range <b>${Math.round(st.tir)}%</b> of the time. The usual target is over 70%.`);
    if(good.length>1) hi.push(`Best day: <b>${new Date(good[0].d).toLocaleDateString("en-GB",{weekday:"long",day:"numeric",month:"short"})}</b> (${Math.round(good[0].s.tir)}% in range). Hardest: <b>${new Date(good[good.length-1].d).toLocaleDateString("en-GB",{weekday:"long",day:"numeric",month:"short"})}</b> (${Math.round(good[good.length-1].s.tir)}%).`);
    if(eps.length){ const c=Array(12).fill(0); eps.forEach(h=>c[Math.floor(new Date(h.start).getHours()/2)]++); const mx=Math.max(...c), i=c.indexOf(mx); hi.push(`<b>${eps.length} hypo${eps.length===1?"":"s"}</b> picked up by the CGM${mx>=2?`, most often between ${pad2(i*2)}:00 and ${pad2(i*2+2)}:00`:""}. Lowest ${Math.min(...eps.map(h=>h.min)).toFixed(1)} mmol/L.`); }
    else hi.push(`No hypos picked up by the CGM.`); }
  else hi.push(`No CGM data for this period yet. Import his CareLink file on the Glucose tab to fill in the glucose part of this report.`);
  if(refused) hi.push(`He refused or barely ate <b>${refused}</b> meal${refused===1?"":"s"}.`);
  if(L.some(e=>e.type==="ill")) hi.push(`He was unwell on ${new Set(L.filter(e=>e.type==="ill").map(e=>ymd(e.at))).size} day(s) in this period.`);
  return `<div class="rp-head"><div><div class="rp-eyebrow">T1D Care Hub · report for the family</div><h1>${esc(title)}</h1><div class="rp-sub">${fmtR(a)} – ${fmtR(Math.min(b, Date.now()))}${P("pump")?` · ${esc(P("pump"))}`:""}${P("cgm")?` + ${esc(P("cgm"))}`:""}</div></div><div class="rp-ring"><b>${st?Math.round(st.tir)+"%":"–"}</b><span>in range</span></div></div>
  <div class="rp-cards">
    ${card("Time in range", st?Math.round(st.tir)+"%":"–", "target over 70%", st&&sp?arrow(st.tir, sp.tir, true):"")}
    ${card("Average glucose", st?st.mean.toFixed(1):"–", "mmol/L", st&&sp?arrow(st.mean, sp.mean, false):"")}
    ${card("Hypos (CGM)", st?eps.length:"–", `${hyposLogged} treated and logged`, st&&sp?arrow(eps.length, epsP.length, false):"")}
    ${card("Time high", st?Math.round(st.above)+"%":"–", "above 10 mmol/L", st&&sp?arrow(st.above, sp.above, false):"")}
    ${card("Carbs eaten", carbsDay!=null?carbsDay+" g":"–", "a day (from Food log)", pc.length?`${Math.round(pc.reduce((x,e)=>x+e.v,0)/pumpDaysN)} g a day entered in pump`:"")}
    ${card("Drinks", drinks.length?Math.round(drinks.reduce((x,e)=>x+(e.ml||0),0)/(new Set(drinks.map(e=>ymd(e.at))).size||1))+" ml":"–", "a day", "")}
    ${card("Site changes", L.filter(e=>e.type==="site").length + psite.length, "logged + from pump", "")}
    ${card("Sensor changes", L.filter(e=>e.type==="sensor").length, "logged", "")}
  </div>
  <h2>Highlights</h2><ul class="rp-list">${hi.map(h=>`<li>${h}</li>`).join("")}</ul>
  <h2>Time in range each day</h2>${bars}
  <div class="rp-legend"><span><i style="background:#2b8a52"></i>In range 3.9–10</span><span><i style="background:#a86100"></i>High</span><span><i style="background:#c0322b"></i>Low</span><span><i style="border-top:2px dashed #2b8a52;height:0;border-radius:0"></i>70% target</span></div>
  <div class="rp-two">
    <div><h2>Meals that went smoothly</h2>${smooth.length?`<ul class="rp-list">${smooth.map(m=>`<li><b>${esc(m.item)}</b> · rose ${m.rise.toFixed(1)} on average (${m.n}×)</li>`).join("")}</ul>`:`<p class="rp-muted">Not enough repeated meals yet.</p>`}</div>
    <div><h2>Meals that were harder</h2>${hard.length?`<ul class="rp-list">${hard.map(m=>`<li><b>${esc(m.item)}</b> · high ${m.highs}/${m.n}, low ${m.lows}/${m.n}</li>`).join("")}</ul>`:`<p class="rp-muted">None stood out.</p>`}</div>
  </div>
  ${eps.length?`<h2>Hypos</h2><table class="rp-table"><thead><tr><th>When</th><th>Lowest</th><th>How long</th></tr></thead><tbody>${eps.slice(-12).reverse().map(h=>`<tr><td>${new Date(h.start).toLocaleString("en-GB",{weekday:"short",day:"numeric",month:"short",hour:"2-digit",minute:"2-digit"})}</td><td>${h.min.toFixed(1)} mmol/L</td><td>${Math.round((h.end-h.start)/6e4)+5} min</td></tr>`).join("")}</tbody></table>`:""}
  ${notes.length?`<h2>Notes and questions</h2><ul class="rp-list">${notes.slice(-10).map(e=>`<li><b>${new Date(e.at).toLocaleDateString("en-GB",{day:"numeric",month:"short"})}</b> · ${esc(e.note||TYPES[e.type].label)}</li>`).join("")}</ul>`:""}
  <div class="rp-two">
    <div><h2>Coming up</h2>${upcoming.length?`<ul class="rp-list">${upcoming.map(x=>`<li><b>${new Date(x.at).toLocaleDateString("en-GB",{weekday:"short",day:"numeric",month:"short"})}</b> · ${esc((AK[x.kind]||AK.other).label)}${x.vaccine?" ("+esc(x.vaccine)+")":""}${x.where?" · "+esc(x.where):""}</li>`).join("")}</ul>`:`<p class="rp-muted">No appointments in the next 30 days.</p>`}</div>
    <div><h2>Supplies</h2><ul class="rp-list">${nd?`<li>Next delivery <b>${fmtR(nd)}</b></li>`:""}${toOrder.length?toOrder.map(x=>`<li><b>${esc(x.name)}</b> · ${Math.floor(x.left)} days left</li>`).join(""):`<li>Nothing needs ordering.</li>`}</ul></div>
  </div>
  <div class="rp-foot">Made with T1D Care Hub on ${fmtR(Date.now())}. This shows what happened; talk to his diabetes team before changing his pump settings.</div>`;
}
function renderReport(){ const el=$("#reportView"); if(!el || $("#v-reports").hidden) return; el.innerHTML = rpBuild(); }
["#rpPeriod","#rpFrom","#rpTo"].forEach(id => $(id).addEventListener("change", renderReport));
function loadScript(src){ return new Promise((res, rej) => { if(document.querySelector(`script[src="${src}"]`)) return res(); const sc=document.createElement("script"); sc.src=src; sc.onload=res; sc.onerror=()=>rej(new Error("load")); document.head.appendChild(sc); }); }
async function htmlToPdf(html, filename){
    await loadScript("https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js");
    await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js");
    const holder=document.createElement("div"); holder.style.cssText="position:fixed;left:-10000px;top:0;width:794px;background:#fff";
    holder.innerHTML=html; document.body.appendChild(holder);
    await new Promise(r=>setTimeout(r,60));
    { const paper = holder.firstChild, pageH = Math.floor(794 * 842 / 595); let changed = true, guard = 0;
      while(changed && guard++ < 40){ changed = false; const top0 = paper.getBoundingClientRect().top;
        for(const el of [...paper.children]){ if(el.classList.contains("rp-gap")) continue; const r = el.getBoundingClientRect(), t = r.top-top0, bt = r.bottom-top0;
          const pageEnd = (Math.floor(t/pageH)+1)*pageH; const isHead = el.tagName==="H2";
          if((bt > pageEnd && r.height < pageH*0.9) || (isHead && t + 140 > pageEnd)){ const gap=document.createElement("div"); gap.className="rp-gap"; gap.style.height=(pageEnd - t + 24)+"px"; paper.insertBefore(gap, el); changed = true; break; } } } }
    const canvas = await window.html2canvas(holder.firstChild, {scale:2, backgroundColor:"#ffffff", useCORS:true, logging:false});
    holder.remove();
    const { jsPDF } = window.jspdf; const pdf = new jsPDF({unit:"pt", format:"a4"}); const pw=pdf.internal.pageSize.getWidth(), ph=pdf.internal.pageSize.getHeight();
    let y=0, page=0; const sliceH = Math.floor(canvas.width * ph / pw);
    while(y < canvas.height){ const c=document.createElement("canvas"); c.width=canvas.width; c.height=Math.min(sliceH, canvas.height-y);
      c.getContext("2d").drawImage(canvas, 0, y, canvas.width, c.height, 0, 0, canvas.width, c.height);
      if(page++) pdf.addPage(); pdf.addImage(c.toDataURL("image/jpeg", 0.92), "JPEG", 0, 0, pw, c.height*pw/canvas.width); y += sliceH; }
    const blob = pdf.output("blob");
    const d = await getDownloads(); if(!d) throw Object.assign(new Error("Saving files isn't available here"),{code:"unavailable"});
    return d.save({filename, data: blob});
}
$("#rpPdf").addEventListener("click", async e => {
  const btn=e.currentTarget, msg=$("#rpMsg"); btn.disabled=true; msg.textContent="Making the PDF…";
  try { { const [a]=rpRange(); const r = await htmlToPdf(`<div class="report-paper print">${rpBuild()}</div>`, `T1D-${$("#rpFor").value==="team"?"clinic-pack":"report"}-${ymd(a)}.pdf`); msg.textContent = r.status==="saved" ? "PDF saved." : ""; }
  } catch(err){ msg.textContent = err.code==="declined" ? "Cancelled." : err.message==="load" ? "Couldn't load the PDF maker. Check your internet connection." : (err.message || "Couldn't make the PDF."); }
  finally { btn.disabled=false; }
});

/* ================= COMING DAYS: what's due + when to watch ================= */
function dayLabel(ms){ const d=Math.round((midnight(ms)-midnight(Date.now()))/864e5); return d<0?"Overdue":d===0?"Today":d===1?"Tomorrow":new Date(ms).toLocaleDateString("en-GB",{weekday:"short",day:"numeric",month:"short"}); }
function renderComing(){
  const ag=$("#agenda"), wl=$("#watchList"); if(!ag||!wl) return;
  const now=Date.now(), horizon=now+7*864e5, items=[];
  const add=(at,text,cls="")=>{ if(at<=horizon) items.push({at,text,cls}); };
  const site=lastSite(), sensor=last("sensor");
  if(site){ const due=site.at+(num(plan.siteDays)??3)*864e5; add(due, `Site change due${isMedtronic()?" (infusion set + reservoir)":""}`, due<now?"bad":due-now<864e5?"warn":""); }
  if(sensor){ const due=sensor.at+(num(plan.sensorDays)??10)*864e5; add(due, "Sensor change due", due<now?"bad":due-now<864e5?"warn":""); }
  { const v=last("vial"); if(v){ const due=v.at+(num(plan.vialDays)??28)*864e5; add(due-2*864e5, "Opened insulin: start a new one by "+new Date(due).toLocaleDateString("en-GB",{weekday:"short",day:"numeric",month:"short"}), due<now?"bad":"warn"); } }
  supplyRows().forEach(sp=>{ if(!sp.days) return; const orderBy=now+Math.max(0,(sp.left-(sp.lead??14)))*864e5; add(orderBy, `Order ${sp.name} (${Math.floor(sp.left)} days left)`, sp.left<=(sp.lead??14)?"bad":"warn"); });
  ["eye","dental"].forEach(k=>{ const d=nextDue(k); if(d!=null && d<=now+30*864e5) items.push({at:d, text:`${CK[k].label} due`, cls:d<now?"bad":""}); });
  { const nd = nextDelivery(); if(nd){ const ob = nd - (num(delivery.lead) ?? 14)*864e5; add(ob, "Order / confirm supplies delivery", ob<now?"bad":"warn"); add(nd, "Supplies delivery due" + (delivery.who?" · "+delivery.who:"")); } }
  if(typeof dayplans!=="undefined") Object.values(dayplans).forEach(pl=>{ const t=keyToMs(pl.date)+hm2min(pl.leave)*6e4; if(t>=midnight(now)) add(t, `Day out${pl.where?": "+pl.where:""} (${pl.items.length} food/drink planned)`, ""); });
  appts.filter(a=>a.status==="booked" && a.at>=midnight(now) && a.at<=horizon).forEach(a=>items.push({at:a.at, text:`${(AK[a.kind]||AK.other).label}${a.vaccine?" ("+a.vaccine+")":""}${a.hasTime?" at "+new Date(a.at).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}):""}${a.where?" · "+a.where:""}`, cls:""}));
  if(typeof todos!=="undefined" && todos) todos.filter(t=>!t.done && t.due).forEach(t=>add(t.due, "To do: "+t.text, t.due<midnight(now)?"bad":""));
  if(typeof trips!=="undefined" && trips) Object.values(trips).forEach(t=>{ if(!t.from) return; const st=keyToMs(t.from); if(st>=midnight(now) && st<=now+7*864e5){ const it=tripItems(t), n=it.filter(([k])=>(t.packed||{})[k]).length; add(st, `Going away${t.where?": "+t.where:""} (${n} of ${it.length} packed)`, n<it.length && st-now<2*864e5?"warn":""); } });
  if(typeof rxDue==="function"){ const rd=rxDue(); if(rd && rd<=now+5*864e5) items.push({at:rd, text:"Order his repeat prescription"+(rx.note?" · "+rx.note:""), cls: rd<now?"bad":"warn"}); }
  if(log.length>20 && (!lastBackup || now-lastBackup>30*864e5)) items.push({at:now, text:"Download a backup (Backup tab)", cls:"warn"});
  items.sort((a,b)=>a.at-b.at);
  ag.innerHTML = items.length ? items.map(i=>`<li><span class="when ${i.cls}">${dayLabel(i.at)}</span><span>${esc(i.text)}</span></li>`).join("")
    : `<li class="muted">Nothing due this week. Log his last site and sensor change, and add supplies, so due dates appear here.</li>`;

  const W=[]; const pts=rangePoints(14); const dayKeys=[...new Set(pts.map(p=>ymd(p.t)))];
  const slot = h => `${pad2(h)}:00–${pad2((h+2)%24)}:00`;
  if(dayKeys.length>=5){
    const lowDays=Array(12).fill(0), highDays=Array(12).fill(0);
    dayKeys.forEach(k=>{ const dp=dayPoints(k); for(let b=0;b<12;b++){ const inb=dp.filter(p=>new Date(p.t).getHours()>=b*2 && new Date(p.t).getHours()<b*2+2);
      if(inb.filter(p=>p.v<3.9).length>=3) lowDays[b]++; if(inb.filter(p=>p.v>13.9).length>=6) highDays[b]++; } });
    const lowSlots=lowDays.map((n,b)=>({n,b})).filter(x=>x.n>=2).sort((a,b)=>b.n-a.n).slice(0,2);
    lowSlots.forEach(x=>{ const night=x.b<3;
      W.push(`<strong>Lows ${slot(x.b*2)}</strong> on ${x.n} of the last ${dayKeys.length} days. ${night?"Keep hypo treatment by his bed and make sure CGM low alerts are on and loud.":"Have hypo treatment with you (and in the nursery bag) and glance at his CGM around then."}`); });
    const highSlots=highDays.map((n,b)=>({n,b})).filter(x=>x.n>=Math.max(3,Math.ceil(dayKeys.length*0.35))).sort((a,b)=>b.n-a.n).slice(0,1);
    highSlots.forEach(x=>W.push(`<strong>Above 13.9 around ${slot(x.b*2)}</strong> on ${x.n} of the last ${dayKeys.length} days. Have ketone strips ready in case it stays high, and note it for his team.`));
    // weekday pattern over 28 days
    const p28=rangePoints(28); const d28=[...new Set(p28.map(p=>ymd(p.t)))];
    if(d28.length>=14){ const wd=Array(7).fill(0), wdN=Array(7).fill(0);
      d28.forEach(k=>{ const dp=dayPoints(k), w=new Date(keyToMs(k)).getDay(); wdN[w]++; if(dp.filter(p=>p.v<3.9).length>=3) wd[w]++; });
      const rates=wd.map((n,i)=>wdN[i]?n/wdN[i]:0), avg=wd.reduce((a,b)=>a+b,0)/Math.max(1,wdN.reduce((a,b)=>a+b,0));
      const names=["Sundays","Mondays","Tuesdays","Wednesdays","Thursdays","Fridays","Saturdays"];
      const busy=rates.map((r,i)=>({r,i})).filter(x=>wd[x.i]>=2 && x.r>=Math.max(0.5,avg*1.8)).map(x=>names[x.i]);
      if(busy.length) W.push(`<strong>Lows are more common on ${busy.join(" and ")}.</strong> If that's a nursery or activity day, pack extra hypo treatment and ask his team whether those days need different settings.`); }
    const nights=dayKeys.slice(-7).filter(k=>dayPoints(k).filter(p=>new Date(p.t).getHours()<6 && p.v<3.9).length>=3).length;
    if(nights>=2 && !lowSlots.some(x=>x.b<3)) W.push(`<strong>Low overnight on ${nights} of the last 7 nights.</strong> Keep hypo treatment by his bed and check CGM alerts are on.`);
  }
  const mp = mealOutcomes(14); const g={}; mp.forEach(r=>(g[r.key] ||= []).push(r));
  Object.values(g).filter(a=>a.length>=3).forEach(a=>{ const lows=a.filter(x=>x.low).length; if(lows/a.length>=0.3) W.push(`<strong>After ${esc(a[0].item)}</strong>, a low followed ${lows} of ${a.length} times. Watch his CGM in the hours after it. This is one of the questions on the Food tab.`); });
  const ill = log.find(e=>e.type==="ill" && now-e.at<3*864e5);
  if(ill) W.push(`<strong>He was unwell recently.</strong> Keep checking ketones if glucose stays high, and keep the Sick day tab handy.`);
  { const nx = (typeof nextWatch==="function") ? nextWatch() : null; if(nx) W.unshift(`<a href="#glucose" class="wt-today ${nx.kind}"><b>${nx.inNow?"Watch now":"Next watch time"}: ${nx.kind==="low"?"lows":"highs"} ${blockLabel(nx.b)}</b>${nx.inNow?"":` · in ${nx.wait>=60?Math.floor(nx.wait/60)+" h":nx.wait+" min"}`}</a>`); }
  wl.innerHTML = W.length ? W.map(w=>`<li>${w}</li>`).join("") : `<li class="muted">${dayKeys.length<5?"Import at least 5 days of CGM data (Glucose tab) to see the times to watch.":"No repeating problem times in the last 14 days."}</li>`;
}

/* ================= MEAL PATTERNS (descriptive, for the dietitian) ================= */
function mealOutcomes(days){
  const from = midnight(Date.now()) - (days-1)*864e5;
  const foods = log.filter(e=>e.type==="food").sort((a,b)=>a.at-b.at);
  const out = [];
  foods.forEach((e, i) => {
    if(e.at<from || !e.item) return;
    const s0 = valueNear(e.at, 15*6e4); if(!s0) return;
    const next = foods[i+1]?.at ?? Infinity;                       // stop at the next meal so effects aren't double-counted
    const end3 = Math.min(e.at+3*36e5, next), end4 = Math.min(e.at+4*36e5, next);
    const pts = [...dayPoints(ymd(e.at)), ...dayPoints(ymd(e.at+864e5))].filter(p=>p.t>e.at && p.t<=end4);
    const in3 = pts.filter(p=>p.t<=end3); if(in3.length<6) return;
    const peak = in3.reduce((a,p)=>p.v>a.v?p:a, in3[0]);
    const above = in3.filter(p=>p.v>10).length, below = pts.filter(p=>p.v<3.9).length;   // ~15 min = 3 readings
    out.push({ key:e.item.trim().toLowerCase(), item:e.item.trim(), meal:e.meal||"", carbs:e.carbs, ate:e.ate||"All", start:s0.v, rise:peak.v-s0.v, peakMin:Math.round((peak.t-e.at)/6e4), high: above>=3, low: below>=3, at:e.at });
  });
  return out;
}
function renderMealPatterns(){
  const box = $("#mealPatterns"); if(!box) return;
  const days = +($("#mpRange")?.value || 30);
  const rows = mealOutcomes(days);
  if(!Object.keys(cgmDays).length){ box.innerHTML = `<div class="empty">Import his CGM data on the Glucose tab and keep logging meals here. Patterns appear once a meal has been logged at least twice with CGM data.</div>`; return; }
  const groups = {};
  rows.forEach(r => (groups[r.key] ||= []).push(r));
  const meals = Object.values(groups).filter(g=>g.length>=2).map(g => {
    const n=g.length, avg=f=>g.reduce((a,x)=>a+(x[f]??0),0)/n;
    const withC = g.filter(x=>x.carbs>0);
    const per10 = withC.length ? withC.reduce((a,x)=>a+x.rise/x.carbs*10,0)/withC.length : null;
    return { item:g[0].item, n, carbs: withC.length ? Math.round(avg("carbs")) : null, rise:avg("rise"), peakMin:Math.round(avg("peakMin")), per10, highs:g.filter(x=>x.high).length, lows:g.filter(x=>x.low).length,
      lowsLeft: g.filter(x=>x.low && x.ate!=="All" && x.ate!=="Most").length, left: g.filter(x=>x.ate!=="All" && x.ate!=="Most").length };
  });
  if(!meals.length){ box.innerHTML = `<div class="empty">Not enough yet. A pattern appears once the same meal (same name) has been logged at least twice on days with CGM data. Tip: pick meal names from the suggestions so they match.</div>`; return; }
  const score = m => (m.highs+m.lows*1.5)/m.n + Math.max(0,m.rise)/10;
  const sorted = meals.slice().sort((a,b)=>score(a)-score(b));
  const steady = sorted.filter(m=>m.highs/m.n<=0.34 && m.lows===0).slice(0,3);
  const hard = sorted.slice().reverse().filter(m=>m.highs/m.n>=0.5 || m.lows/m.n>=0.25).sort((a,b)=>(b.lows/b.n)-(a.lows/a.n)).slice(0,3);
  const qs = [];
  hard.forEach(m => {
    if(m.highs/m.n>=0.5) qs.push(`${m.item} goes above 10 after ${m.highs} of ${m.n} times, peaking about ${m.peakMin} minutes after eating${m.carbs!=null?` (about ${m.carbs} g carbs logged)`:""}. Is the carb count right, and would giving the insulin earlier help?`);
    if(m.lows/m.n>=0.25) qs.push(`${m.item} is followed by a low before his next meal ${m.lows} of ${m.n} times. Is the insulin for this meal too much, or is something else (like activity) causing it?`);
  });
  const partial = rows.filter(r=>r.ate!=="All" && r.ate!=="Most"), partialLows = partial.filter(r=>r.low).length;
  if(partial.length>=2 && partialLows/partial.length>=0.34) qs.push(`When he leaves food (ate half or less), a low follows ${partialLows} of ${partial.length} times. Should we give some of his meal insulin after he has eaten?`);
  const night = rows.filter(r=>/dinner|bedtime/i.test(r.meal)); const nightLows = night.filter(r=>r.low).length;
  if(night.length>=3 && nightLows/night.length>=0.3) qs.push(`Lows follow dinner or bedtime snack ${nightLows} of ${night.length} times. Should his evening settings be checked?`);
  const card = (title, list, empty) => `<div class="mpcard"><h3>${title}</h3>${list.length ? `<ul class="plain">${list.map(m=>`<li><strong>${esc(m.item)}</strong> <span class="muted">(${m.n}×)</span><br><span class="muted">rises ${m.rise.toFixed(1)} on average, peak at ${m.peakMin} min · high ${m.highs}/${m.n} · low ${m.lows}/${m.n}</span></li>`).join("")}</ul>` : `<div class="muted">${empty}</div>`}</div>`;
  box.innerHTML = `<div class="twochart">${card("Went smoothly", steady, "No meal stands out as steady yet.")}${card("Harder to manage", hard, "No meal stands out as difficult. Good sign.")}</div>
    <div class="table-wrap"><table>
      <thead><tr><th>Meal</th><th style="text-align:right">Times</th><th style="text-align:right">Carbs</th><th style="text-align:right">Avg rise</th><th style="text-align:right">Rise per 10 g</th><th style="text-align:right">Peak</th><th style="text-align:right">Highs</th><th style="text-align:right">Lows</th></tr></thead>
      <tbody>${sorted.map(m=>`<tr><td>${esc(m.item)}</td><td class="num">${m.n}</td><td class="num">${m.carbs!=null?m.carbs+" g":"—"}</td><td class="num">${m.rise>=0?"+":""}${m.rise.toFixed(1)}</td><td class="num">${m.per10!=null?(m.per10>=0?"+":"")+m.per10.toFixed(1):"—"}</td><td class="num">${m.peakMin} min</td><td class="num">${m.highs}/${m.n}</td><td class="num">${m.lows}/${m.n}</td></tr>`).join("")}</tbody>
    </table></div>
    <div class="mpq"><h3>Questions for his dietitian</h3>${qs.length ? `<ol class="plain">${qs.map(q=>`<li>${esc(q)}</li>`).join("")}</ol><div class="btns"><button type="button" class="primary" id="copyQs">Copy questions</button></div>` : `<div class="muted">No questions stand out from this period.</div>`}</div>`;
  const cq = $("#copyQs"); if(cq) cq.onclick = async () => { const txt = `Questions about ${P("name")||"my son"}'s meals (last ${days} days):\n` + qs.map((q,i)=>`${i+1}. ${q}`).join("\n");
    try { await navigator.clipboard.writeText(txt); toast("Copied"); } catch { toast("Couldn't copy. Select the text instead"); } };
}
$("#mpRange").addEventListener("change", renderMealPatterns);

$("#gRange").addEventListener("change", renderGlucose);
let _rsz; addEventListener("resize", () => { clearTimeout(_rsz); _rsz = setTimeout(() => { if(!$("#v-glucose").hidden) renderGlucose(); }, 200); });

$("#dPrev").addEventListener("click", ()=>{ gDay = midnight(gDay-12*36e5); renderGlucose(); });
$("#dNext").addEventListener("click", ()=>{ gDay = midnight(gDay+36*36e5); renderGlucose(); });
$("#dDate").addEventListener("change", e=>{ if(e.target.value){ gDay = keyToMs(e.target.value); renderGlucose(); } });

/* ================= GROWTH & CHECKS ================= */
let growth = [], checks = [];
const CK = {eye:{label:"Eye check", months:12}, dental:{label:"Dental check", months:6}};
const fmtD = ms => new Date(ms).toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"});
const addMonths = (ms,m) => { const d=new Date(ms); d.setMonth(d.getMonth()+m); return d.getTime(); };
function checkItems(kind){
  const ak = kind==="dental" ? "dentist" : "eye";
  return [...checks.filter(c=>c.kind===kind).map(c=>({...c, src:"check"})),
          ...appts.filter(a=>a.kind===ak && a.status==="done").map(a=>({id:a.id, kind, at:a.at, where:a.where, note:a.outcome||a.notes, next:a.next||null, src:"appt"}))].sort((a,b)=>b.at-a.at);
}
function lastCheck(kind){ return checkItems(kind)[0]; }
function nextDue(kind){ const l=lastCheck(kind); if(!l) return null; return l.next || addMonths(l.at, num(plan[kind+"Months"]) ?? CK[kind].months); }
function dueInfo(kind){
  const d = nextDue(kind); if(d==null) return {v:"Not logged", n:"Add his last "+CK[kind].label.toLowerCase(), cls:"", days:null};
  const days = Math.round((midnight(d)-midnight(Date.now()))/864e5);
  return {v: days<0 ? "Overdue" : days===0 ? "Today" : days<=60 ? `In ${days} days` : fmtD(d), n: days<0 ? `Was due ${fmtD(d)}` : `Due ${fmtD(d)}`, cls: days<0 ? "bad" : days<=30 ? "warn" : "", days};
}
function latestWith(f){ return growth.filter(g=>g[f]!=null).sort((a,b)=>b.at-a.at); }
function renderChecks(){
  const H=latestWith("h"), W=latestWith("w"), e=dueInfo("eye"), dn=dueInfo("dental");
  const ch = (arr,f,u,dp) => arr.length>1 ? `${arr[0][f]-arr[1][f]>=0?"+":""}${(arr[0][f]-arr[1][f]).toFixed(dp)} ${u} since ${fmtD(arr[1].at)}` : arr.length ? "Measured "+fmtD(arr[0].at) : "No measurement yet";
  $("#cStrip").innerHTML = [
    `<div class="stat"><span class="k">Height</span><span class="v">${H.length?H[0].h.toFixed(1)+" cm":"—"}</span><span class="n">${ch(H,"h","cm",1)}</span></div>`,
    `<div class="stat"><span class="k">Weight</span><span class="v">${W.length?W[0].w.toFixed(2)+" kg":"—"}</span><span class="n">${ch(W,"w","kg",2)}</span></div>`,
    `<div class="stat ${e.cls}"><span class="k">Eye check</span><span class="v">${e.v}</span><span class="n">${e.n}</span></div>`,
    `<div class="stat ${dn.cls}"><span class="k">Dental check</span><span class="v">${dn.v}</span><span class="n">${dn.n}</span></div>`
  ].join("");
  growthChart("#hChart", latestWith("h").reverse(), "h", "cm", 1);
  growthChart("#wChart", latestWith("w").reverse(), "w", "kg", 2);
  const rows = growth.slice().sort((a,b)=>b.at-a.at);
  $("#gTable").innerHTML = rows.length ? rows.map(g=>`<tr><td>${fmtD(g.at)}</td><td class="num">${g.h!=null?g.h.toFixed(1)+" cm":"—"}</td><td class="num">${g.w!=null?g.w.toFixed(2)+" kg":"—"}</td><td class="muted" style="overflow-wrap:anywhere">${esc(g.note||"")}</td><td style="text-align:right"><button class="small danger" data-gdel="${esc(g.id)}" ${canWrite?"":"disabled"}>Delete</button></td></tr>`).join("")
    : `<tr><td colspan="5"><div class="empty">No measurements yet. Add his height and weight from his last clinic visit or red book.</div></td></tr>`;
  $("#checkCols").innerHTML = ["eye","dental"].map(k => { const list = checkItems(k); const d=dueInfo(k);
    const booked = appts.filter(a=>a.kind===(k==="dental"?"dentist":"eye") && a.status==="booked" && a.at>=midnight(Date.now())).sort((a,b)=>a.at-b.at)[0];
    return `<div class="ckcard"><div style="display:flex;justify-content:space-between;gap:8px;align-items:center"><h3>${CK[k].label}</h3><span class="chip ${d.cls==="bad"?"low":d.cls==="warn"?"high":d.days!=null?"ok":""}">${d.v}</span></div>
      <div class="muted" style="font-size:.85rem">${d.days!=null?d.n+" · ":""}Usual gap: every ${num(plan[k+"Months"]) ?? CK[k].months} months</div>
      ${booked?`<div class="booked-pill">Booked: ${fmtD(booked.at)}${booked.hasTime?" "+new Date(booked.at).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}):""}${booked.where?" · "+esc(booked.where):""}</div>`:""}
      ${list.length ? `<ul class="ckhist">${list.map(c=>`<li><div style="min-width:0;overflow-wrap:anywhere"><strong>${fmtD(c.at)}</strong>${c.where?` · ${esc(c.where)}`:""}${c.note?`<div class="muted">${esc(c.note)}</div>`:""}${c.next?`<div class="muted">Next given: ${fmtD(c.next)}</div>`:""}</div><button class="small danger" ${c.src==="appt"?`data-adel="${esc(c.id)}"`:`data-cdel="${esc(c.id)}"`} ${canWrite?"":"disabled"}>Delete</button></li>`).join("")}</ul>` : `<div class="empty">No ${k} check-ups logged.</div>`}
    </div>`; }).join("");
}
function growthChart(sel, pts, f, unit, dp){
  const box=$(sel); if(!pts.length){ box.innerHTML = `<div class="empty">No measurements yet. Tap <strong>Add measurement</strong> to start his chart.</div>`; return; }
  const W=420, H=200, l=44, r=14, t=12, b=28;
  const t0=pts[0].at, t1=pts[pts.length-1].at, span=Math.max(t1-t0, 30*864e5);
  const vals=pts.map(p=>p[f]); let lo=Math.min(...vals), hi=Math.max(...vals); const pad=Math.max((hi-lo)*.15, f==="h"?1:.2); lo-=pad; hi+=pad;
  const X = ms => l + (ms-t0)/span*(W-l-r), Y = v => t + (1-(v-lo)/(hi-lo))*(H-t-b);
  let g=""; for(let i=0;i<=3;i++){ const v=lo+(hi-lo)*i/3; g+=`<line x1="${l}" x2="${W-r}" y1="${Y(v)}" y2="${Y(v)}" stroke="var(--line)" stroke-width="1"/><text x="${l-6}" y="${Y(v)+4}" text-anchor="end">${v.toFixed(f==="h"?0:1)}</text>`; }
  g+=`<text x="${l}" y="${H-8}">${new Date(t0).toLocaleDateString("en-GB",{month:"short",year:"2-digit"})}</text><text x="${W-r}" y="${H-8}" text-anchor="end">${new Date(t1).toLocaleDateString("en-GB",{month:"short",year:"2-digit"})}</text>`;
  g+=`<path d="M${pts.map(p=>`${X(p.at).toFixed(1)} ${Y(p[f]).toFixed(1)}`).join(" L")}" fill="none" stroke="var(--accent)" stroke-width="2" stroke-linejoin="round"/>`;
  pts.forEach((p,i)=>{ g+=`<circle cx="${X(p.at)}" cy="${Y(p[f])}" r="5" fill="var(--accent)" stroke="var(--surface)" stroke-width="2"/><circle data-i="${i}" cx="${X(p.at)}" cy="${Y(p[f])}" r="14" fill="transparent" style="cursor:pointer"/>`; });
  const last=pts[pts.length-1]; g+=`<text x="${X(last.at)-8}" y="${Y(last[f])-10}" text-anchor="end" style="fill:var(--fg);font-weight:700">${last[f].toFixed(dp)} ${unit}</text>`;
  box.innerHTML = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${f==="h"?"Height":"Weight"} over time">${g}</svg><div class="tip" hidden></div>${pts.length===1?`<div class="muted" style="font-size:.85rem;margin-top:4px">Add another measurement to see how he's growing over time.</div>`:""}`;
  const tip=box.querySelector(".tip");
  box.querySelectorAll("[data-i]").forEach(c => { const show = () => { const p=pts[+c.dataset.i]; const prev=pts[+c.dataset.i-1];
      tip.innerHTML = `<div class="muted">${fmtD(p.at)}</div><div><b>${p[f].toFixed(dp)}</b> ${unit}</div>${prev?`<div class="muted">${p[f]-prev[f]>=0?"+":""}${(p[f]-prev[f]).toFixed(dp)} ${unit} since previous</div>`:""}`;
      tip.hidden=false; const r=c.getBoundingClientRect(), br=box.getBoundingClientRect(); let x=r.left-br.left+20; if(x+tip.offsetWidth>box.clientWidth) x=r.left-br.left-tip.offsetWidth-6; tip.style.left=Math.max(0,x)+"px"; tip.style.top=Math.max(0,r.top-br.top-10)+"px"; };
    c.addEventListener("pointerenter", show); c.addEventListener("pointerdown", show); c.addEventListener("pointerleave", ()=>tip.hidden=true); });
}
const dateMs = v => { const [y,m,d]=v.split("-").map(Number); return new Date(y,m-1,d,12).getTime(); };
const growDlg=$("#growDlg"), chkDlg=$("#chkDlg");
$("#addGrowth").addEventListener("click", ()=>{ if(!canWrite) return (toast("Saving isn't switched on here. Open the page in a browser (see the note at the top)"), $("#dbNote").scrollIntoView({behavior:"smooth",block:"center"})); $("#growForm").reset(); $("#g_date").value=ymd(Date.now()); growDlg.showModal(); });
$("#growCancel").addEventListener("click", ()=>growDlg.close());
$("#growForm").addEventListener("submit", async ev => { ev.preventDefault();
  const d={at:dateMs($("#g_date").value), h:num($("#g_h").value), w:num($("#g_w").value), note:$("#g_note").value.trim(), by:me};
  if(d.h==null && d.w==null){ toast("Enter a height or a weight"); return; }
  try { await base.collection("growth").add(d); growDlg.close(); toast("Saved"); } catch(err) { toast(`Couldn't save (${err.code||err.message||"error"})`); } });
$("#addCheck").addEventListener("click", ()=>openAppt(null, false, {kind:"dentist", status:"done", date:ymd(Date.now())}));
$("#chkCancel").addEventListener("click", ()=>chkDlg.close());
$("#chkForm").addEventListener("submit", async ev => { ev.preventDefault();
  const d={kind:$("#c_kind").value, at:dateMs($("#c_date").value), where:$("#c_where").value.trim(), note:$("#c_note").value.trim(), next:$("#c_next").value?dateMs($("#c_next").value):null, by:me};
  try { await base.collection("checks").add(d); chkDlg.close(); toast("Saved"); } catch(err) { toast(`Couldn't save (${err.code||err.message||"error"})`); } });
document.addEventListener("click", async ev => {
  const b = ev.target.closest("[data-gdel],[data-cdel]"); if(!b) return;
  b.disabled=true;
  try { await deleteWithUndo(b.dataset.gdel?"growth":"checks", b.dataset.gdel||b.dataset.cdel, b.dataset.gdel?"Measurement":"Check-up"); } catch { toast("Couldn't delete. Try again"); b.disabled=false; }
});


/* ================= BACKUP ================= */
const BK_COLS = ["log","cgm","pump","growth","checks","supplies","appts","foods","photos","dayplans","hba1c","rx","settings","todos","trips","handovers","bedtime"];
let lastBackup = null, dl = null, fileBackupAt = null;
function bkStatus(m){ const n=$("#bkStatus"); n.hidden=!m; n.textContent=m||""; }
function renderBackup(){
  const el=$("#bkLine"); if(!el) return; const at = fileBackupAt;
  el.textContent = at ? `Last copy saved ${ago(at)} (${fmtD(at)})` : "Not saved to a file yet";
}
$("#bkExport").addEventListener("click", async e => {
  const btn=e.currentTarget; if(!base) return bkStatus("Saving isn't available in this view.");
  await getDownloads();
  if(!dl){ bkStatus("Downloading files isn't available in this view. Try again from the installed app or a browser."); return; }
  btn.disabled=true; bkStatus("Collecting everything…");
  try {
    const out = {app:"t1d-care-hub", version:1, exportedAt:new Date().toISOString(), plan:(await base.get()).data()||{}};
    for(const c of BK_COLS){ const snap = await base.collection(c).get(); out[c] = snap.docs.map(d=>({id:d.id, ...d.data()})); }
    const nm = (plan.name||"son").trim().replace(/[^A-Za-z0-9]+/g,"-") || "son";
    const r = await dl.save({filename:`T1D-backup-${nm}-${ymd(Date.now())}.json`, data:JSON.stringify(out)});
    if(r.status==="saved"){ await base.collection("meta").doc("backup").set({at:Date.now(), by:me}); await base.collection("meta").doc("filebackup").set({at:Date.now(), by:me}).catch(()=>{}); bkStatus(`Backup saved: ${out.log.length} log entries, ${out.cgm.length} days of CGM, ${out.growth.length} measurements, ${out.checks.length} check-ups, ${out.supplies.length} supplies.`); }
    else bkStatus("");
  } catch(err){ bkStatus(err.code==="declined" ? "Backup cancelled." : err.code==="rate_limited" ? "A save prompt is already open. Finish that one first." : "Couldn't make the backup. Try again."); }
  finally { btn.disabled=false; }
});
let pendingRestore = null;
$("#bkFile").addEventListener("change", async ev => {
  const f=ev.target.files[0]; ev.target.value=""; if(!f) return;
  if(!canWrite) return bkStatus("Saving isn't available in this view.");
  try {
    let bytes = new Uint8Array(await f.arrayBuffer());
    if(bytes[0]===0x50 && bytes[1]===0x4b){   // a .zip (GitHub wraps its daily backups in one)
      if(!window.JSZip) throw new Error("bad");
      const zip = await JSZip.loadAsync(bytes); const names = Object.keys(zip.files).filter(n => !zip.files[n].dir && /\.(enc|json)$/i.test(n)).sort();
      if(!names.length) throw new Error("bad");
      bytes = await zip.files[names[names.length-1]].async("uint8array");
    }
    if(isEncrypted(bytes)) return askBackupPassword(bytes);
    openRestore(JSON.parse(new TextDecoder().decode(bytes)));
  } catch { bkStatus("That file isn't a T1D Care Hub backup."); }
});
/* daily GitHub backups are locked with your backup password (openssl AES-256-CBC, PBKDF2-SHA256) */
const isEncrypted = b => b.length>16 && String.fromCharCode(...b.slice(0,8))==="Salted__";
async function decryptBackup(bytes, password){
  const salt = bytes.slice(8,16), body = bytes.slice(16);
  const base = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveBits"]);
  const bits = new Uint8Array(await crypto.subtle.deriveBits({name:"PBKDF2", hash:"SHA-256", salt, iterations:200000}, base, 384));
  const key = await crypto.subtle.importKey("raw", bits.slice(0,32), "AES-CBC", false, ["decrypt"]);
  return new Uint8Array(await crypto.subtle.decrypt({name:"AES-CBC", iv:bits.slice(32,48)}, key, body));
}
function askBackupPassword(bytes){
  const c = $("#bkConfirm"); c.hidden=false; bkStatus("");
  c.innerHTML = `<strong>This backup is locked</strong><span>Enter the backup password you saved in GitHub (BACKUP_PASSWORD).</span><form id="bkPwForm" style="display:flex;gap:8px;flex-wrap:wrap"><input id="bkPw" type="password" autocomplete="off" placeholder="Backup password" style="flex:1 1 200px"><button class="primary" type="submit">Unlock</button><button type="button" id="bkNo">Cancel</button></form>`;
  $("#bkNo").onclick = () => { c.hidden=true; };
  $("#bkPw").focus();
  $("#bkPwForm").onsubmit = async e => { e.preventDefault();
    try { const plain = await decryptBackup(bytes, $("#bkPw").value); openRestore(JSON.parse(new TextDecoder().decode(plain))); }
    catch { bkStatus("Wrong password, or the file is damaged."); }
  };
}
function openRestore(data){
  {
    if(data.app!=="t1d-care-hub" || !Array.isArray(data.log)) return bkStatus("That file isn't a T1D Care Hub backup.");
    pendingRestore = data;
    const c = $("#bkConfirm"); c.hidden=false;
    c.innerHTML = `<strong>Restore backup from ${fmtD(Date.parse(data.exportedAt))}?</strong><span>It has ${data.log.length} log entries, ${(data.cgm||[]).length} days of CGM, ${(data.growth||[]).length} measurements, ${(data.checks||[]).length} check-ups and ${(data.supplies||[]).length} supplies. Anything missing here is added back and matching items are reset to the backup. Nothing else is deleted.</span><div class="btns"><button class="primary" type="button" id="bkGo">Restore now</button><button type="button" id="bkNo">Cancel</button></div>`;
    $("#bkNo").onclick = () => { pendingRestore=null; c.hidden=true; };
    $("#bkGo").onclick = doRestore;
    bkStatus("");
  }
}
async function doRestore(){
  const data=pendingRestore; if(!data) return; $("#bkConfirm").hidden=true; pendingRestore=null;
  const total = BK_COLS.reduce((a,c)=>a+(data[c]||[]).length,0); let n=0;
  try {
    if(data.plan && Object.keys(data.plan).length) await base.set(data.plan);
    for(const c of BK_COLS) for(const d of (data[c]||[])){ const {id, ...body} = d; if(!id) continue; await base.collection(c).doc(String(id)).set(body); if(++n%20===0) bkStatus(`Restoring… ${n} of ${total}`); }
    bkStatus(`Restored ${n} items from the backup.`); planDirty=false;
  } catch(err){ bkStatus(err.code==="quota_exceeded" ? `Storage filled up after ${n} items.` : `Restore stopped after ${n} of ${total} items. Try again; finished items won't be duplicated.`); }
}





async function explainNoSave(){
  const n=$("#dbNote"); n.hidden=false;
  const perm = await (window.claude?.use?.("permissions") ?? null);
  let st = "unavailable"; try { st = perm ? await perm.state("db") : "unavailable"; } catch {}
  const msg = {
    prompt: ["Saving needs your permission.", "Turn on saving"],
    denied: ["Saving was turned off for this page.", "Open permissions"],
    granted: ["Saving was just allowed. Reload the page to start saving.", "Reload"],
    unavailable: ["Saving can't run in this view. Check your internet connection and sign in again.", ""]
  }[st] || ["Saving isn't available here.", ""];
  n.innerHTML = `<strong>${msg[0]}</strong> <span class="muted">(status: ${esc(st)})</span>${msg[1]?` <button type="button" class="small primary" id="fixSave" style="margin-left:6px">${msg[1]}</button>`:""}`;
  const b=$("#fixSave"); if(!b) return;
  b.addEventListener("click", async () => {
    try {
      if(st==="granted") return location.reload();
      if(st==="denied"){ await perm.manage(); } else { await perm.request(["db","user"]); }
      const now = await perm.state("db");
      if(now==="granted") location.reload(); else explainNoSave();
    } catch { n.innerHTML = "<strong>Couldn't open permissions.</strong> Reload the app and sign in again."; }
  });
}

/* ================= v5.0 ================= */
/* --- 2. Order message (Supplies) --- */
function orderText(){
  const nd = nextDelivery(), all = supplyRows(), need = all.filter(s=>s.status!=="OK" || (nd && s.days && Math.floor(s.left) < Math.round((nd-Date.now())/864e5))), list = need.length ? need : all;
  const name = P("name") || "my son";
  const lines = list.map(s=>{ const q = perDelivery(s) ?? (s.days ? Math.ceil(90/s.days) : null); return `- ${s.name}${q?`: ${q}`:""}${s.days?` (about ${Math.floor(s.left)} days left)`:""}`; });
  if(!lines.length) lines.push("- (add the items here)");
  return `Hello,\n\nPlease could I order the following diabetes supplies for ${name}:\n${lines.join("\n")}\n\nPlease let me know when it will be delivered.\n\nThank you`;
}
function waNumber(){ let d=String(delivery.phone||"").replace(/[^\d+]/g,""); if(!d) return ""; if(d.startsWith("+")) return d.slice(1); if(d.startsWith("00")) return d.slice(2); if(d.startsWith("0")) return "44"+d.slice(1); return d; }
function orderLinks(){ const t=$("#or_text").value; $("#orMail").href = `mailto:?subject=${encodeURIComponent("Diabetes supplies order")}&body=${encodeURIComponent(t)}`; $("#orWa").href = `https://wa.me/${waNumber()}?text=${encodeURIComponent(t)}`; }
$("#orderMsg").addEventListener("click", () => { $("#or_text").value = orderText(); $("#orShare").hidden = !navigator.share; orderLinks(); $("#orderDlg").showModal(); });
$("#or_text").addEventListener("input", orderLinks);
$("#orClose").addEventListener("click", () => $("#orderDlg").close());
$("#orCopy").addEventListener("click", async () => { const t=$("#or_text"); try { await navigator.clipboard.writeText(t.value); toast("Copied"); } catch { t.select(); toast("Selected. Press copy on your device"); } });
$("#orShare").addEventListener("click", async () => { try { await navigator.share({title:"Diabetes supplies order", text:$("#or_text").value}); } catch {} });

/* --- 6. Repeat prescription (Supplies) --- */
var rx = {};
function rxDue(){ return rx && rx.last && num(rx.every) ? rx.last + num(rx.every)*864e5 : null; }
function renderRx(){
  const el=$("#rxPanel"); if(!el) return; const due=rxDue(), now=Date.now();
  const cls = !due ? "" : due<now ? "bad" : due-now<5*864e5 ? "warn" : "ok";
  el.innerHTML = `<div class="ph"><h2>Repeat prescription</h2><button type="button" class="small primary" id="rxDone">Ordered today</button></div>
    <div class="muted" style="font-size:.9rem">Ask the GP for his repeat prescription in time. Tap <b>Ordered today</b> each time you request it and it will show in Coming days 5 days before the next one is due.</div>
    <div class="rx-row"><div class="stat"><span class="k">Last ordered</span><span class="v">${rx.last?fmtD(rx.last).replace(/ \d{4}$/,""):"—"}</span></div>
      <div class="stat"><span class="k">Next due</span><span class="v ${cls}">${due?dayLabel(due):"—"}</span></div></div>
    <div class="grid2 rx-set"><label>Every <small>days</small><input id="rx_every" type="number" min="7" step="1" inputmode="numeric" value="${esc(rx.every??28)}"></label>
      <label>What's on it <small>optional</small><input id="rx_note" placeholder="e.g. NovoRapid, test strips, ketone strips" value="${esc(rx.note||"")}"></label></div>
    <div class="btns"><button type="button" class="small" id="rxSave">Save</button></div>`;
  $("#rxDone").onclick = () => rxWrite({last: Date.now()}, "Prescription marked as ordered");
  $("#rxSave").onclick = () => rxWrite({every: num($("#rx_every").value) ?? 28, note: $("#rx_note").value.trim()}, "Saved");
}
async function rxWrite(patch, msg){ if(!canWrite) return toast("Saving isn't switched on here"); try { await base.collection("rx").doc("main").set({...rx, every: rx.every ?? 28, ...patch, by: me, updatedAt: Date.now()}); toast(msg); } catch { toast("Couldn't save"); } }

/* --- 3. HbA1c lab results (Glucose) --- */
var a1c = [];
const a1cPct = mm => mm/10.929 + 2.15, a1cMm = pc => (pc-2.15)*10.929;
const gmiOf = mean => { const pc = 3.31 + 0.02392*mean*18.016; return {pc, mm: a1cMm(pc)}; };
function gmiBefore(t, days=90){ const p=ptsBetween(t-days*864e5, t); const dn=new Set(p.map(x=>ymd(x.t))).size; if(dn<7) return null; const s=statsOf(p); return {...gmiOf(s.mean), days:dn}; }
function renderA1c(){
  const el=$("#a1cPanel"); if(!el) return; const rows=[...a1c].sort((a,b)=>b.at-a.at);
  const s14=statsOf(rangePoints(14)), g=s14?gmiOf(s14.mean):null;
  let chart = "";
  if(rows.length>=2){ const pts=[...rows].reverse(), W=400, H=150, l=30, r=26, t=12, b=24; const vals=pts.map(x=>x.mmol); const lo=Math.min(40,...vals)-4, hi=Math.max(60,...vals)+4;
    const t0=pts[0].at, t1=pts[pts.length-1].at||t0+1, X=ms=>l+(ms-t0)/Math.max(1,t1-t0)*(W-l-r), Y=v=>t+(1-(v-lo)/(hi-lo))*(H-t-b);
    chart = `<svg viewBox="0 0 ${W} ${H}" width="100%" role="img" aria-label="HbA1c results over time"><rect x="${l}" y="${Y(48)}" width="${W-l-r}" height="${Math.max(0,Y(lo)-Y(48))}" fill="var(--ok-soft,#e3f4ea)" opacity=".6"/><line x1="${l}" x2="${W-r}" y1="${Y(48)}" y2="${Y(48)}" stroke="var(--ok,#2b8a52)" stroke-dasharray="4 4"/><text x="${l-6}" y="${Y(48)+4}" text-anchor="end" font-size="12" fill="currentColor">48</text>
      <polyline fill="none" stroke="var(--page,var(--accent))" stroke-width="2.5" points="${pts.map(x=>X(x.at)+","+Y(x.mmol)).join(" ")}"/>
      ${pts.map(x=>{ const e=gmiBefore(x.at); return `<circle cx="${X(x.at)}" cy="${Y(x.mmol)}" r="5" fill="var(--page,var(--accent))"/>${e?`<circle cx="${X(x.at)}" cy="${Y(e.mm)}" r="4" fill="none" stroke="currentColor" stroke-width="1.5"/>`:""}<text x="${X(x.at)}" y="${H-8}" text-anchor="middle" font-size="12" fill="currentColor">${new Date(x.at).toLocaleDateString("en-GB",{month:"short",year:"2-digit"})}</text>`; }).join("")}</svg>
      <div class="legendrow"><span><i class="lg dot"></i>Lab HbA1c</span><span>○ App estimate (GMI) for the 90 days before</span><span>- - 48 mmol/mol, the usual UK target</span></div>`; }
  el.innerHTML = `<div class="ph"><h2>HbA1c results</h2><button type="button" class="small primary" id="a1cAdd">Add a result</button></div>
    <div class="muted" style="font-size:.9rem">The lab result from clinic, next to what his CGM data suggests (GMI). If they keep disagreeing, mention it to his team.</div>
    ${g?`<div class="a1c-now"><span class="k">App estimate now (GMI, last 14 days)</span><b>${Math.round(g.mm)} mmol/mol</b><span class="muted">${g.pc.toFixed(1)}% · from average ${s14.mean.toFixed(1)} mmol/L</span></div>`:""}
    ${chart}
    ${rows.length?`<ul class="a1c-list">${rows.map(x=>{ const e=gmiBefore(x.at); return `<li><span class="a1c-d">${fmtD(x.at)}</span><b>${Math.round(x.mmol)} mmol/mol</b><span class="muted">${a1cPct(x.mmol).toFixed(1)}%${e?` · app estimate ${Math.round(e.mm)}`:""}${x.note?" · "+esc(x.note):""}</span><button type="button" class="small" data-a1cdel="${esc(x.id)}" aria-label="Delete">✕</button></li>`; }).join("")}</ul>`
      : `<p class="muted" style="margin:0">No results yet. Add the number from his last clinic letter or the clinic app.</p>`}`;
  $("#a1cAdd").onclick = () => { if(!canWrite) return toast("Saving isn't switched on here"); $("#a1cForm").reset(); $("#a1_date").value = ymd(Date.now()); $("#a1cDlg").showModal(); };
  el.querySelectorAll("[data-a1cdel]").forEach(b => b.onclick = () => deleteWithUndo("hba1c", b.dataset.a1cdel, "Result"));
}
$("#a1cForm").addEventListener("submit", async ev => { ev.preventDefault(); const v=num($("#a1_val").value); if(v==null) return;
  const mmol = $("#a1_unit").value==="pc" ? a1cMm(v) : v; if(mmol<15 || mmol>200) return toast("That doesn't look like an HbA1c. Check the number and unit.");
  try { await base.collection("hba1c").add({at: dateMs($("#a1_date").value), mmol: Math.round(mmol*10)/10, note: $("#a1_note").value.trim(), by: me}); $("#a1cDlg").close(); toast("Result saved"); } catch { toast("Couldn't save"); } });
$("#a1cCancel").addEventListener("click", () => $("#a1cDlg").close());

/* --- 7. Activity patterns (Glucose) --- */
const ACT_STOP = new Set(["at","the","a","an","in","to","with","and","went","go","going","for","of","on","his","he","did","had","we","some","after","before","then"]);
function actKey(e){ const w=(e.note||"").toLowerCase().replace(/[^a-z0-9\s]/g," ").split(/\s+/).filter(x=>x && !ACT_STOP.has(x) && !/^\d+$/.test(x)).slice(0,2); return w.join(" "); }
function lowAfter(t, hrs=6){ const p=ptsBetween(t, t+hrs*36e5); if(p.length < hrs*6) return null; return p.some(x=>x.v<hypoBelow()); }
function actPatterns(){
  const since = Date.now()-90*864e5, groups = {};
  log.filter(e=>e.type==="activity" && e.at>=since).forEach(e=>{ const k=actKey(e)||"activity"; (groups[k] ||= []).push(e); });
  let bn=0, bl=0; rangePoints(30).length && [...new Set(rangePoints(30).map(p=>ymd(p.t)))].forEach(k=>{ for(let h=0;h<24;h+=3){ const r=lowAfter(keyToMs(k)+h*36e5); if(r!=null){ bn++; if(r) bl++; } } });
  const base0 = bn>=8 ? bl/bn : null;
  const rows = Object.entries(groups).map(([k,es])=>{ const res=es.map(e=>lowAfter(e.at)).filter(x=>x!=null); return {k, n:es.length, withData:res.length, lows:res.filter(Boolean).length}; }).sort((a,b)=>b.withData-a.withData || b.n-a.n);
  return {rows, base: base0, baseN: bn};
}
function renderActs(){
  const el=$("#actPanel"); if(!el) return; const {rows, base:b} = actPatterns();
  const title = k => k==="activity" ? "Activity (no note)" : k.charAt(0).toUpperCase()+k.slice(1);
  el.innerHTML = `<div><h2>After activities</h2><div class="muted" style="font-size:.9rem">How often a low (below ${hypoBelow()}) followed within 6 hours of each activity you logged, compared with any 6 hours. Name activities the same way each time (e.g. "Swimming", "Soft play") so they group together.</div></div>
    ${!rows.length ? `<p class="muted" style="margin:0">Log activities from Today (Activity, with a short note like "Swimming") and import his CGM file. Patterns show here after a couple of each.</p>`
    : `<ul class="act-list">${rows.slice(0,8).map(r=>{ const rate=r.withData?r.lows/r.withData:null; const more = rate!=null && b!=null && r.withData>=2 && rate-b>=0.25;
        const tag = rate==null ? `<span class="chip">Needs CGM data</span>` : r.withData<2 ? `<span class="chip">Log a few more</span>` : more ? `<span class="chip low">More lows after this</span>` : `<span class="chip ok">No extra lows</span>`;
        return `<li><div><b>${esc(title(r.k))}</b><div class="muted" style="font-size:.85rem">${r.n} logged${r.withData?` · low within 6 h after ${r.lows} of ${r.withData}`:""}${rate!=null&&b!=null?` (${Math.round(rate*100)}% vs ${Math.round(b*100)}% usually)`:""}</div></div>${tag}</li>`; }).join("")}</ul>
      <p class="muted" style="margin:0;font-size:.85rem">A pattern, not proof. If one activity keeps being followed by lows, ask his team whether to set a Temp Target or give a snack before it.</p>`}`;
}

/* --- 4. Clinic pack (Reports) --- */
function buildClinicReport(){
  const [a,b] = rpRange(), end=Math.min(b, Date.now()), days=Math.max(1,Math.round((b-a+1)/864e5)), name=P("name")||"";
  const fmtR = ms => new Date(ms).toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"});
  const pts=ptsBetween(a,b), st=statsOf(pts), eps=hypoEpisodes(pts), g=st?gmiOf(st.mean):null;
  const wear = st ? Math.min(100, pts.length/((end-a)/3e5)*100) : 0;
  const L = log.filter(e=>e.at>=a && e.at<=b), foods=L.filter(e=>e.type==="food"), fd=new Set(foods.map(e=>ymd(e.at))).size||1;
  const pev=pumpBetween(a,b), pc=pev.filter(e=>e.k===1), pi=pev.filter(e=>e.k===2), pdays=new Set(pev.map(e=>ymd(e.t))).size||1;
  const bg=allBgChecks(Math.ceil((Date.now()-a)/864e5)+1).filter(x=>x.t>=a && x.t<=b);
  const kets=L.filter(e=>e.ket!=null); const ill=new Set(L.filter(e=>e.type==="ill").map(e=>ymd(e.at))).size;
  const notes=L.filter(e=>e.type==="note").sort((x,y)=>x.at-y.at);
  const gr=[...growth].sort((x,y)=>y.at-x.at)[0]; const labs=[...a1c].sort((x,y)=>y.at-x.at).slice(0,4);
  const pc1 = (k) => st ? pct(st.b[k], st.n).toFixed(1)+"%" : "–";
  const row = (k,v,t="") => `<tr><td>${k}</td><td><b>${v}</b></td><td class="rp-muted">${t}</td></tr>`;
  const slots = [...Array(12)].map((_,i)=>{ const p=pts.filter(x=>Math.floor(new Date(x.t).getHours()/2)===i).map(x=>x.v).sort((x,y)=>x-y); if(!p.length) return null; return {med:p[Math.floor(p.length/2)], lo:pct(p.filter(v=>v<3.9).length,p.length), hi:pct(p.filter(v=>v>10).length,p.length)}; });
  return `<div class="rp-head"><div><div class="rp-eyebrow">T1D Care Hub · for the diabetes team</div><h1>Clinic pack${name?": "+esc(name):""}</h1><div class="rp-sub">${fmtR(a)} – ${fmtR(end)} (${days} days)${P("pump")?` · ${esc(P("pump"))}`:""}${P("cgm")?` + ${esc(P("cgm"))}`:""}</div></div><div class="rp-ring"><b>${st?Math.round(st.tir)+"%":"–"}</b><span>in range</span></div></div>
  <h2>CGM summary</h2>
  ${st ? `<table class="rp-table"><tbody>
    ${row("Time in range 3.9–10.0", pc1("in"), "target over 70%")}${row("Below 3.9", pct(st.b.low+st.b.vlow,st.n).toFixed(1)+"%", "target under 4%")}${row("Below 3.0", pc1("vlow"), "target under 1%")}
    ${row("Above 10.0", pct(st.b.high+st.b.vhigh,st.n).toFixed(1)+"%", "target under 25%")}${row("Above 13.9", pc1("vhigh"), "target under 5%")}
    ${row("Average glucose", st.mean.toFixed(1)+" mmol/L")}${row("GMI (estimated HbA1c)", `${Math.round(g.mm)} mmol/mol (${g.pc.toFixed(1)}%)`, "estimate from CGM, not a lab result")}
    ${row("Variability (CV)", st.cv.toFixed(0)+"%", "target 36% or less")}${row("Sensor data", Math.round(wear)+"%", "of possible 5-minute readings")}
    ${row("Hypos picked up by CGM", eps.length, "below 3.9 for 15 min or more")}${row("Hypos treated and logged at home", L.filter(e=>e.type==="hypo").length)}
  </tbody></table>` : `<p class="rp-muted">No CGM data imported for this period.</p>`}
  ${st ? `<h2>Time of day (2-hour blocks)</h2><table class="rp-table rp-tod"><thead><tr><th>Time</th>${slots.map((_,i)=>`<th>${pad2(i*2)}</th>`).join("")}</tr></thead><tbody>
    <tr><td>Middle value</td>${slots.map(s=>`<td>${s?s.med.toFixed(1):"–"}</td>`).join("")}</tr>
    <tr><td>% low</td>${slots.map(s=>`<td class="${s&&s.lo>=4?"rp-bad":""}">${s?Math.round(s.lo):"–"}</td>`).join("")}</tr>
    <tr><td>% high</td>${slots.map(s=>`<td class="${s&&s.hi>=40?"rp-warn":""}">${s?Math.round(s.hi):"–"}</td>`).join("")}</tr></tbody></table>` : ""}
  ${eps.length ? `<h2>CGM hypos</h2><table class="rp-table"><thead><tr><th>When</th><th>Lowest</th><th>How long</th></tr></thead><tbody>${eps.slice(-15).reverse().map(h=>`<tr><td>${new Date(h.start).toLocaleString("en-GB",{weekday:"short",day:"numeric",month:"short",hour:"2-digit",minute:"2-digit"})}</td><td>${h.min.toFixed(1)} mmol/L</td><td>${Math.round((h.end-h.start)/6e4)} min</td></tr>`).join("")}</tbody></table>` : ""}
  <div class="rp-two">
    <div><h2>Pump and food</h2><table class="rp-table"><tbody>
      ${row("Bolus insulin (pump file)", pi.length ? (pi.reduce((x,e)=>x+e.v,0)/pdays).toFixed(1)+" U/day" : "–", pi.length?`${(pi.length/pdays).toFixed(1)} doses a day`:"")}
      ${row("Carbs entered in pump", pc.length ? Math.round(pc.reduce((x,e)=>x+e.v,0)/pdays)+" g/day" : "–")}
      ${row("Carbs logged at home", foods.length ? Math.round(foods.reduce((x,e)=>x+(e.carbs||0),0)/fd)+" g/day" : "–")}
      ${row("Site changes", L.filter(e=>e.type==="site").length + pev.filter(e=>e.k===4).length)}
      ${row("Finger-prick checks", bg.length ? `${bg.length} (${(bg.length/days).toFixed(1)}/day)` : "–", bg.length?`average ${(bg.reduce((x,y)=>x+y.v,0)/bg.length).toFixed(1)}`:"")}
      ${row("Ketone readings", kets.length ? `${kets.length}, highest ${Math.max(...kets.map(e=>e.ket)).toFixed(1)}` : "none")}
      ${row("Unwell days", ill)}
    </tbody></table></div>
    <div><h2>HbA1c and growth</h2><table class="rp-table"><tbody>
      ${labs.length ? labs.map(x=>row("HbA1c "+fmtR(x.at), Math.round(x.mmol)+" mmol/mol", (()=>{ const e=gmiBefore(x.at); return e?`GMI before it ${Math.round(e.mm)}`:""; })())).join("") : row("HbA1c","–","none recorded")}
      ${gr ? row("Growth "+fmtR(gr.at), [gr.h?gr.h+" cm":"", gr.w?gr.w+" kg":""].filter(Boolean).join(" · ")||"–") : row("Growth","–","none recorded")}
      ${P("hypo")?row("Hypo treated below", esc(P("hypo"))+" mmol/L"):""}${P("siteDays")?row("Site changed every", esc(P("siteDays"))+" days"):""}
    </tbody></table></div>
  </div>
  <h2>Our questions and notes</h2>${notes.length ? `<ul class="rp-list">${notes.slice(-12).map(e=>`<li><b>${new Date(e.at).toLocaleDateString("en-GB",{day:"numeric",month:"short"})}</b> · ${esc(e.note||"")}</li>`).join("")}</ul>` : `<p class="rp-muted">Add questions as Notes in the log and they appear here.</p>`}
  <div class="rp-foot">Prepared by his parents from home records on ${fmtR(Date.now())}. CGM figures come from the imported CareLink file; GMI is an estimate, not a lab HbA1c.</div>`;
}
const rpBuild = () => $("#rpFor")?.value==="team" ? buildClinicReport() : buildReport();
$("#rpFor").addEventListener("change", renderReport);

/* --- 5. Alert from the other parent (while the app is open) --- */
const PA_KEY = "t1d-alerted"; var paMem = new Set();
function paSeen(){ try { return JSON.parse(localStorage.getItem(PA_KEY)||"[]"); } catch { return []; } }
function paMark(id){ try { localStorage.setItem(PA_KEY, JSON.stringify([...paSeen(), id].slice(-60))); } catch {} }
const whoName = id => (names && names[id]) || "The other parent";
function paBeep(){ try { const C=new (window.AudioContext||window.webkitAudioContext)(); [0,.35,.7].forEach(t=>{ const o=C.createOscillator(), gn=C.createGain(); o.frequency.value=880; o.connect(gn); gn.connect(C.destination); gn.gain.setValueAtTime(0.0001,C.currentTime+t); gn.gain.exponentialRampToValueAtTime(0.3,C.currentTime+t+0.02); gn.gain.exponentialRampToValueAtTime(0.0001,C.currentTime+t+0.25); o.start(C.currentTime+t); o.stop(C.currentTime+t+0.3); }); setTimeout(()=>C.close().catch(()=>{}),1500); } catch {} try { navigator.vibrate?.([250,120,250]); } catch {} }
function parentAlert(id, title, body, go){
  if(paMem.has(id) || paSeen().includes(id)) return; paMem.add(id); paMark(id);
  const el=$("#parentAlert"); if(el.hidden || !el.querySelector(".pa-items")) el.innerHTML = `<div class="pa-items"></div><div class="btns"><button type="button" class="small primary" data-pa="open">Open</button><button type="button" class="small" data-pa="ok">OK</button></div>`;
  const it=document.createElement("div"); it.innerHTML="<b></b><span></span>"; it.querySelector("b").textContent=title; it.querySelector("span").textContent=body; el.querySelector(".pa-items").prepend(it);
  if(el.hidden || go==="glucose") el.dataset.go=go; el.hidden=false; paBeep();
  try { if("Notification" in window && Notification.permission==="granted"){ const opts={body, tag:id}; (navigator.serviceWorker?.getRegistration?.() || Promise.resolve(null)).then(r => r ? r.showNotification(title, opts) : new Notification(title, opts)).catch(()=>{}); } } catch {}
}
$("#parentAlert").addEventListener("click", e => { const b=e.target.closest("[data-pa]"); if(!b) return; const el=$("#parentAlert"); el.hidden=true; if(b.dataset.pa==="open") location.hash = el.dataset.go || "today"; });
function paCheck(){
  if(!me) return; const now=Date.now();
  log.forEach(e=>{ if(e.type==="hypo" && e.by && e.by!==me && now-e.at<10*6e4 && now-e.at>-5*6e4) parentAlert("h"+e.id, `${whoName(e.by)} is treating a hypo`, `${e.gl!=null?e.gl.toFixed(1)+" mmol/L · ":""}logged ${ago(e.at)}${e.note?" · "+e.note:""}`, "glucose"); });
  if(sick?.active && sick.by && sick.by!==me && now-sick.start<10*6e4) parentAlert("s"+sick.start, `${whoName(sick.by)} started sick-day mode`, "Check ketones and glucose. The timer is on the Sick day page.", "sick");
}

/* --- 1. Home-screen shortcuts (?go=hypo|food|bg) --- */
(function(){ let go=null; try { go=new URLSearchParams(location.search).get("go"); } catch {} if(!go || !["hypo","food","bg"].includes(go)) return; let tries=0;
  const t=setInterval(()=>{ if(!canWrite && ++tries<60) return; clearInterval(t); try { history.replaceState(null, "", location.pathname + (location.hash||"#today")); } catch {}
    if(!canWrite) return; if(go==="hypo") openHypo(); else { location.hash = go==="food" ? "food" : "glucose"; setTimeout(()=>openEntry(go), 80); } }, 100); })();
renderRx();

/* ================= Nursery / childminder handover sheet ================= */
function ageOf(dob){ if(!dob) return ""; const d=new Date(dob), n=new Date(); let m=(n.getFullYear()-d.getFullYear())*12+n.getMonth()-d.getMonth(); if(n.getDate()<d.getDate()) m--; if(m<0) return ""; return m<24 ? `${m} months` : `${Math.floor(m/12)} years${m%12?` ${m%12} months`:""}`; }
function buildHandover(){
  const name = P("name") || "Our son", lo = hypoBelow(), ket = ketAbove();
  const v = (k, fb) => P(k) ? esc(P(k)).replace(/\n/g,"<br>") : `<span class="ho-blank">${fb}</span>`;
  const tel = k => P(k) ? `<b>${esc(P(k))}</b>` : `<span class="ho-blank">ask his parents</span>`;
  const snacks = (typeof favSort==="function" ? favSort() : []).filter(f=>f.kind!=="drink" || f.carbs).slice(0,10);
  const age = ageOf(P("dob"));
  return `<div class="ho-head"><div><div class="rp-eyebrow">Type 1 diabetes · handover for nursery or childminder</div><h1>${esc(name)}</h1>
      <div class="rp-sub">${[age, P("dob")?"born "+new Date(P("dob")).toLocaleDateString("en-GB",{day:"numeric",month:"long",year:"numeric"}):"", P("pump")?esc(P("pump"))+" insulin pump":"", P("cgm")?esc(P("cgm"))+" sensor":""].filter(Boolean).join(" · ")}</div></div>
      <div class="ho-999"><b>999</b><span>if he is unconscious, fitting or can't swallow</span></div></div>
  <div class="ho-box ho-low"><h2>If he is low: below ${lo} mmol/L, or he seems unwell</h2>
    <div class="ho-cols"><div><h3>Signs to look for</h3><p>${v("signs","e.g. grizzly, pale, sweaty, clingy, unusually quiet")}</p><p class="ho-small">He is too young to tell you. If in doubt, check his glucose.</p></div>
    <div><h3>What to do</h3><p>${v("hypoTx","Give fast sugar as his parents showed you, then recheck in 15 minutes")}</p><p class="ho-small">Stay with him. Recheck after 15 minutes and repeat if still below ${lo}. Then call us.</p></div></div>
    <div class="ho-em"><b>Emergency:</b> if he is unconscious, fitting or can't swallow, give <b>nothing</b> by mouth. Put him in the recovery position and <b>call 999</b>, then call us.${P("glucagon")?` Glucagon: ${esc(P("glucagon").trim().replace(/\.+$/,""))}.`:""}</div></div>
  <div class="ho-cols ho-two">
    <div class="ho-box"><h2>If he is high</h2><p>${v("highTx",`Above ${ket} mmol/L: call his parents. Give water to drink.`)}</p><p class="ho-small">Call us straight away if he is high and vomiting, drowsy or breathing fast.</p></div>
    <div class="ho-box"><h2>Pump and sensor</h2><p>${v("hoPump","Don't remove the pump or sensor. If it alarms or comes off, call us. Don't give insulin.")}</p></div>
  </div>
  <div class="ho-cols ho-two">
    <div class="ho-box"><h2>Who to call</h2><p>${v("hoContacts","Parents' names and numbers")}</p>
      <p class="ho-small">Diabetes team (day): ${tel("teamDay")}<br>Out of hours: ${tel("teamOoh")}</p></div>
    <div class="ho-box"><h2>Hypo kit</h2><p>${v("hoKit","Where his hypo kit is kept")}</p><p class="ho-small">It should go with him everywhere: trips, the garden, nap room.</p></div>
  </div>
  ${snacks.length?`<div class="ho-box"><h2>His usual food and carbs</h2><p class="ho-small" style="margin-top:0">Write down what he eats and roughly how much, so we can count it.</p><table class="rp-table"><tbody>${snacks.map(f=>`<tr><td>${esc(f.name)}</td><td>${f.portion?esc(f.portion):""}</td><td style="text-align:right"><b>${f.carbs} g</b> carbs</td></tr>`).join("")}</tbody></table></div>`:""}
  <div class="rp-foot">Prepared by his parents on ${new Date().toLocaleDateString("en-GB",{day:"numeric",month:"long",year:"numeric"})}. This sheet is a quick guide; if it differs from his healthcare plan from the diabetes team, follow that plan.</div>`;
}
$("#hoPreview").addEventListener("click", () => { $("#hoView").innerHTML = buildHandover(); $("#hoMsg").textContent = "Print it, or send the PDF. Give a copy to everyone who looks after him."; $("#hoDlg").showModal(); });
$("#hoClose").addEventListener("click", () => $("#hoDlg").close());
$("#hoPdf").addEventListener("click", async e => { const btn=e.currentTarget, msg=$("#hoMsg"); btn.disabled=true; msg.textContent="Making the PDF…";
  try { const r = await htmlToPdf(`<div class="report-paper print ho-paper">${buildHandover()}</div>`, `Handover-sheet-${(P("name")||"T1D").replace(/[^\w-]+/g,"-")}.pdf`); msg.textContent = r.status==="saved" ? "PDF saved." : ""; }
  catch(err){ msg.textContent = err.code==="declined" ? "Cancelled." : err.message==="load" ? "Couldn't load the PDF maker. Check your internet connection." : (err.message || "Couldn't make the PDF."); }
  finally { btn.disabled=false; } });

/* ================= v6.0 ================= */
/* --- CGM import nudge (Today) --- */
function lastCgmDay(){ const ks=Object.keys(cgmDays||{}).sort(); return ks.length ? ks[ks.length-1] : null; }
function renderNudge(){
  const el=$("#importNudge"); if(!el) return; let snooze=0; try { snooze=+localStorage.getItem("t1d-nudge-snooze")||0; } catch {}
  const k=lastCgmDay(), now=Date.now(); let msg=null;
  if(k){ const age=Math.floor((midnight(now)-keyToMs(k))/864e5); if(age>=7) msg=`<b>His glucose data stops ${age} days ago</b><span>Import his latest CareLink file so the graphs, watch times and reports stay right.</span>`; }
  else if(log.length>=10 && log.some(e=>e.at<now-7*864e5)) msg=`<b>No CGM file imported yet</b><span>Import his CareLink file to see his graphs, watch times and reports.</span>`;
  el.hidden = !msg || snooze>now; if(el.hidden) return;
  el.innerHTML = `<div>${msg}</div><div class="btns"><button type="button" class="small primary" id="nudgeGo">Import now</button><button type="button" class="small" id="nudgeLater">Not now</button></div>`;
  $("#nudgeGo").onclick = () => { location.hash="glucose"; };
  $("#nudgeLater").onclick = () => { try { localStorage.setItem("t1d-nudge-snooze", String(now+3*864e5)); } catch {} el.hidden=true; };
}

/* --- Pump settings history (Insulin) --- */
var setChanges = [];
const SET_KINDS = {ratio:"Carb ratio", ait:"Active insulin time", target:"Target (SmartGuard)", basal:"Basal pattern", isf:"Correction / sensitivity", other:"Other"};
function setEffect(c, nextAt){
  const end = Math.min(Date.now(), nextAt ?? Infinity, c.at+14*864e5);
  const bp=ptsBetween(c.at-14*864e5, c.at-1), ap=ptsBetween(c.at, end);
  const bd=new Set(bp.map(p=>ymd(p.t))).size, ad=new Set(ap.map(p=>ymd(p.t))).size;
  return {b: bd>=5 ? statsOf(bp) : null, a: ad>=5 ? statsOf(ap) : null, bd, ad, eb:hypoEpisodes(bp).length, ea:hypoEpisodes(ap).length};
}
function renderSettings(){
  const el=$("#setPanel"); if(!el) return; const rows=[...setChanges].sort((x,y)=>y.at-x.at);
  const fx = (r,i) => { const e=setEffect(r, i>0 ? rows[i-1].at : null);
    if(!e.a) return `<div class="set-fx muted">${e.ad<5 && r.at>Date.now()-5*864e5 ? "Too soon to tell. Check back after 5 days of CGM data." : "Needs CGM data after the change to compare."}</div>`;
    if(!e.b) return `<div class="set-fx muted">No CGM data from before the change to compare with.</div>`;
    const d=Math.round(e.a.tir-e.b.tir), lo=e.a.below-e.b.below;
    return `<div class="set-fx"><span>In range <b>${Math.round(e.b.tir)}% → ${Math.round(e.a.tir)}%</b> <i class="${d>=3?"up":d<=-3?"down":""}">${d>0?"+":""}${d}</i></span><span>Low <b>${e.b.below.toFixed(1)}% → ${e.a.below.toFixed(1)}%</b> <i class="${lo<=-0.5?"up":lo>=0.5?"down":""}">${lo>0?"+":""}${lo.toFixed(1)}</i></span><span>Average <b>${e.b.mean.toFixed(1)} → ${e.a.mean.toFixed(1)}</b></span></div>`; };
  el.innerHTML = `<div class="ph"><h2>Pump setting changes</h2><button type="button" class="small primary" id="setAdd">Log a change</button></div>
    <div class="muted" style="font-size:.9rem">Each time his team changes a setting, log it here. The app compares his CGM in the 2 weeks before and after, so you can see if it helped.</div>
    ${rows.length ? `<ul class="set-list">${rows.map((r,i)=>`<li><div class="set-top"><div><b>${esc(SET_KINDS[r.kind]||r.kind)}</b>${r.from||r.to?` <span class="set-ch">${esc(r.from||"?")} → ${esc(r.to||"?")}</span>`:""}<div class="muted" style="font-size:.85rem">${fmtD(r.at)}${r.who?" · "+esc(r.who):""}${r.note?" · "+esc(r.note):""}</div></div><button type="button" class="small" data-setdel="${esc(r.id)}" aria-label="Delete">✕</button></div>${fx(r,i)}</li>`).join("")}</ul>`
      : `<p class="muted" style="margin:0">No changes logged yet.</p>`}
    <p class="bk-note">This shows what happened after each change, not why. Lots of things move his glucose: growing, illness, teething, activity.</p>`;
  $("#setAdd").onclick = () => { if(!canWrite) return toast("Saving isn't switched on here"); $("#setForm").reset(); $("#st_date").value=ymd(Date.now()); $("#setDlg").showModal(); };
  el.querySelectorAll("[data-setdel]").forEach(b => b.onclick = () => deleteWithUndo("settings", b.dataset.setdel, "Change"));
}
$("#setForm").addEventListener("submit", async ev => { ev.preventDefault();
  try { await base.collection("settings").add({at: dateMs($("#st_date").value), kind: $("#st_kind").value, from: $("#st_from").value.trim(), to: $("#st_to").value.trim(), who: $("#st_who").value.trim(), note: $("#st_note").value.trim(), by: me}); $("#setDlg").close(); toast("Change logged"); } catch { toast("Couldn't save"); } });
$("#setCancel").addEventListener("click", () => $("#setDlg").close());

/* --- Clinic follow-up to-dos (Appointments) --- */
var todos = [];
function renderTodos(){
  const el=$("#todoPanel"); if(!el) return;
  const open=todos.filter(t=>!t.done).sort((a,b)=>(a.due??9e15)-(b.due??9e15) || a.at-b.at), done=todos.filter(t=>t.done).sort((a,b)=>(b.doneAt||0)-(a.doneAt||0)).slice(0,5);
  const li = t => `<li class="${t.done?"done":""}"><label class="do-chk"><input type="checkbox" data-todo="${esc(t.id)}" ${t.done?"checked":""}><span><b>${esc(t.text)}</b>${t.due&&!t.done?`<small class="${t.due<midnight(Date.now())?"late":""}">${dayLabel(t.due)}</small>`:""}</span></label><button type="button" class="small" data-tododel="${esc(t.id)}" aria-label="Delete">✕</button></li>`;
  el.innerHTML = `<div class="ph"><h2>After clinic: to do</h2><span class="muted">${open.length} open</span></div>
    <div class="muted" style="font-size:.9rem">What the team asked you to do or try. Ones with a date show in Coming days.</div>
    <form class="todo-add" id="todoForm"><input id="td_text" placeholder="e.g. Book bloods, try new ratio at breakfast" required autocomplete="off"><input id="td_due" type="date" aria-label="Do by (optional)"><button class="primary small" type="submit">Add</button></form>
    ${open.length?`<ul class="todo-list">${open.map(li).join("")}</ul>`:`<p class="muted" style="margin:0">Nothing open.</p>`}
    ${done.length?`<details><summary class="muted" style="cursor:pointer">Done recently (${done.length})</summary><ul class="todo-list">${done.map(li).join("")}</ul></details>`:""}`;
  $("#todoForm").onsubmit = async ev => { ev.preventDefault(); if(!canWrite) return toast("Saving isn't switched on here"); const t=$("#td_text").value.trim(); if(!t) return; const d=$("#td_due").value;
    try { await base.collection("todos").add({text:t, due: d?dateMs(d):null, done:false, at:Date.now(), by:me}); } catch { toast("Couldn't save"); } };
  el.querySelectorAll("[data-todo]").forEach(c => c.onchange = async () => { const t=todos.find(x=>x.id===c.dataset.todo); if(!t) return; const {id, ...body}=t;
    try { await base.collection("todos").doc(id).set({...body, done:c.checked, doneAt:c.checked?Date.now():null}); } catch { toast("Couldn't save"); } });
  el.querySelectorAll("[data-tododel]").forEach(b => b.onclick = () => deleteWithUndo("todos", b.dataset.tododel, "To-do"));
}

/* --- Holiday packing planner (Day out) --- */
var trips = {}, tripCur = null, tripTimer = null;
const nightsOf = t => t.from && t.to ? Math.max(1, Math.round((keyToMs(t.to)-keyToMs(t.from))/864e5)) : 0;
function tripDefault(){ const up=Object.values(trips).filter(t=>t.to>=ymd(Date.now())).sort((a,b)=>a.from<b.from?-1:1)[0]; return up ? up.id : (Object.keys(trips)[0] || null); }
function tripItems(t){
  const n=nightsOf(t), days=n+1, out=[];
  supplyRows().filter(s=>s.days).forEach(s=>{ const need=Math.max(1,Math.ceil(days/s.days))*2; out.push(["sup:"+s.id, `${s.name}: ${need}`, `Double what he'll use${s.qty!=null&&s.qty<need?`. You have ${s.qty} at home: order more`:""}`]); });
  const vd=num(P("vialDays"))??28, vi=Math.max(1,Math.ceil(days/vd))*2;
  out.push(["insulin", `Insulin for the pump: ${vi} vials/cartridges`, "Keep it cool (cool bag or Frio pouch), never in the hold"]);
  out.push(["pens", "Back-up insulin pens, needles and long-acting insulin", "In case the pump fails. Only as his team has told you"]);
  out.push(["hypo", `Hypo treatment: at least ${days*3} goes`, (P("hypoTx")||"What his team told you to give").split("\n")[0].slice(0,90)]);
  out.push(["glucagon","Glucagon", P("glucagon")||"Check it's in date"]);
  out.push(["meter","BG meter, strips, finger pricker and ketone strips","Plus spare meter batteries"]);
  out.push(["batt","Spare AA batteries for the pump",""]);
  out.push(["charger","Sensor transmitter charger and phone charger", t.abroad?"And a plug adapter":""]);
  out.push(["plan","Printed care plan and handover sheet","Care plan → Handover sheet"]);
  out.push(["rx","Copy of his prescription and the diabetes team's numbers", P("teamOoh")?"Out of hours: "+P("teamOoh"):""]);
  if(t.flying){ out.push(["letter","Travel letter from his diabetes team","Ask a few weeks before. Shows security why you carry needles and liquids"]);
    out.push(["hand","All diabetes supplies in hand luggage","Split between two bags in case one goes missing"]);
    out.push(["scan","Check Medtronic's airport advice for the pump and sensor","Ask security for a hand search instead of the body scanner if advised"]); }
  if(t.abroad){ out.push(["ins","Travel insurance that covers his diabetes","Declare it, or claims can be refused"]); out.push(["ghic","GHIC card (if going to Europe)","Free on the NHS website, one for each of you"]); out.push(["loan","Ask Medtronic about a spare pump for the trip","Ask well before you go"]); }
  if(Math.abs(num(t.tz)||0)>=2) out.push(["tz",`Time difference ${t.tz>0?"+":""}${t.tz} h: ask his team how to change the pump clock`, "Changing it changes when his basal is given"]);
  return out;
}
function tripSave(t){ trips[t.id]=t; clearTimeout(tripTimer); tripTimer=setTimeout(async()=>{ try { await base.collection("trips").doc(t.id).set(t); } catch { toast("Couldn't save the trip"); } }, 500); }
function renderTrip(){
  const el=$("#tripPanel"); if(!el) return; if(!tripCur || !trips[tripCur]) tripCur = tripDefault(); const t = tripCur ? trips[tripCur] : null;
  const list = Object.values(trips).sort((a,b)=>a.from<b.from?1:-1);
  if(!t){ el.innerHTML = `<div class="ph"><h2>Going away</h2><button type="button" class="small primary" id="tripNew">Plan a trip</button></div><div class="muted" style="font-size:.9rem">A packing list for nights away: supplies worked out from his stock, insulin, documents, and extra checks if you're flying or going abroad.</div>`; $("#tripNew").onclick = tripNew; return; }
  const items=tripItems(t), packed=t.packed||{}, n=items.filter(([k])=>packed[k]).length, nights=nightsOf(t);
  el.innerHTML = `<div class="ph"><h2>Going away</h2><div class="btns" style="margin:0">${list.length>1?`<select id="tr_pick" aria-label="Trip">${list.map(x=>`<option value="${esc(x.id)}" ${x.id===t.id?"selected":""}>${esc(x.where||"Trip")} · ${x.from?new Date(keyToMs(x.from)).toLocaleDateString("en-GB",{day:"numeric",month:"short"}):""}</option>`).join("")}</select>`:""}<button type="button" class="small" id="tripNew">New trip</button></div></div>
    <div class="grid2">
      <label>Where<input id="tr_where" value="${esc(t.where||"")}" placeholder="e.g. Grandma's in Leeds, Spain" autocomplete="off"></label>
      <label>Time difference <small>hours, if any</small><input id="tr_tz" type="number" step="1" min="-12" max="14" value="${esc(t.tz??"")}" inputmode="numeric"></label>
      <label>Leaving<input id="tr_from" type="date" value="${esc(t.from||"")}"></label>
      <label>Back<input id="tr_to" type="date" value="${esc(t.to||"")}"></label>
    </div>
    <div class="trip-opts"><label class="do-chk"><input type="checkbox" id="tr_fly" ${t.flying?"checked":""}><span><b>Flying</b></span></label><label class="do-chk"><input type="checkbox" id="tr_abroad" ${t.abroad?"checked":""}><span><b>Abroad</b></span></label></div>
    <div class="ph" style="margin-top:4px"><h3 style="margin:0">Pack${nights?` for ${nights} night${nights===1?"":"s"}`:""}</h3><span class="muted">${n} of ${items.length} packed</span></div>
    <ul class="do-pack">${items.map(([k,tt,sub])=>`<li><label class="do-chk ${packed[k]?"on":""}"><input type="checkbox" data-trpack="${esc(k)}" ${packed[k]?"checked":""}><span><b>${esc(tt)}</b>${sub?`<small>${esc(sub)}</small>`:""}</span></label></li>`).join("")}</ul>
    <div class="btns"><button type="button" class="small danger" id="tripDel">Delete this trip</button></div>`;
  const upd = (patch, rerender=true) => { const x={...trips[t.id], ...patch}; tripSave(x); if(rerender) renderTrip(); };
  $("#tr_where").oninput = e => upd({where:e.target.value}, false);
  $("#tr_tz").onchange = e => upd({tz:num(e.target.value)});
  $("#tr_from").onchange = e => upd({from:e.target.value, to: t.to && t.to<e.target.value ? e.target.value : t.to});
  $("#tr_to").onchange = e => upd({to:e.target.value});
  $("#tr_fly").onchange = e => upd({flying:e.target.checked}); $("#tr_abroad").onchange = e => upd({abroad:e.target.checked});
  el.querySelectorAll("[data-trpack]").forEach(c => c.onchange = () => upd({packed:{...(trips[t.id].packed||{}), [c.dataset.trpack]: c.checked}}));
  $("#tripNew").onclick = tripNew; const pk=$("#tr_pick"); if(pk) pk.onchange = () => { tripCur=pk.value; renderTrip(); };
  $("#tripDel").onclick = async () => { const id=t.id; delete trips[id]; tripCur=null; renderTrip(); await deleteWithUndo("trips", id, "Trip"); };
}
function tripNew(){ if(!canWrite) return toast("Saving isn't switched on here"); const from=ymd(Date.now()+7*864e5), to=ymd(Date.now()+10*864e5); const id="t"+Date.now().toString(36);
  const t={id, where:"", from, to, flying:false, abroad:false, tz:null, packed:{}}; trips[id]=t; tripCur=id; tripSave(t); renderTrip(); setTimeout(()=>$("#tr_where")?.focus(),50); }
renderNudge(); renderSettings(); renderTodos(); renderTrip();

/* ================= v7.0 ================= */
/* --- Handover tab: the sheet + a record of each handover --- */
var handovers = [];
const HO_ITEMS = [["kit","Hypo kit"],["glucagon","Glucagon"],["meter","BG meter"],["spare","Spare set / reservoir"],["snacks","Snacks and drinks"],["sheet","Handover sheet"]];
function renderHandoverTab(){
  const sh=$("#hoTab"); if(!sh) return; sh.innerHTML = buildHandover();
  const el=$("#hoRecList"); const rows=[...handovers].sort((a,b)=>b.at-a.at); const now=Date.now();
  const tm = ms => new Date(ms).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"});
  const active = rows.find(r=>!r.backAt && r.at<=now && (!r.until || r.until>now-6*36e5));
  $("#hoActive").hidden = !active; if(active) $("#hoActive").innerHTML = `<div><b>With ${esc(active.to)}</b><span>since ${tm(active.at)}${active.until?` · back about ${tm(active.until)}`:""}</span></div><button type="button" class="small primary" data-hoback="${esc(active.id)}">He's back</button>`;
  $("#hoCount").textContent = rows.length ? `${rows.length} recorded` : "";
  el.innerHTML = rows.length ? rows.slice(0,30).map(r=>`<li><div class="hr-top"><div><b>${esc(r.to)}</b><div class="muted" style="font-size:.85rem">${new Date(r.at).toLocaleDateString("en-GB",{weekday:"short",day:"numeric",month:"short"})} · ${tm(r.at)}${r.until?"–"+tm(r.until):""}${r.where?" · "+esc(r.where):""}${r.backAt?` · back ${tm(r.backAt)}`:""}</div></div>
      <div class="hr-act">${!r.backAt?`<button type="button" class="small" data-hoback="${esc(r.id)}">Back</button>`:""}<button type="button" class="small" data-hodel="${esc(r.id)}" aria-label="Delete">✕</button></div></div>
      ${(r.items||[]).length?`<div class="hr-chips">${r.items.map(k=>`<span class="chip">${esc((HO_ITEMS.find(x=>x[0]===k)||[k,k])[1])}</span>`).join("")}</div>`:""}
      ${r.gl!=null||r.ate||r.note?`<div class="hr-note">${[r.gl!=null?`Last BG ${r.gl.toFixed(1)}`:"", r.ate?`Last ate: ${esc(r.ate)}`:"", r.note?esc(r.note):""].filter(Boolean).join(" · ")}</div>`:""}</li>`).join("")
    : `<li class="muted">No handovers recorded yet. Tap <b>Record a handover</b> each time someone else looks after him.</li>`;
  document.querySelectorAll("[data-hoback]").forEach(b => b.onclick = async () => { const r=handovers.find(x=>x.id===b.dataset.hoback); if(!r) return; const {id,...body}=r; try { await base.collection("handovers").doc(id).set({...body, backAt:Date.now()}); toast("Marked as back"); } catch { toast("Couldn't save"); } });
  el.querySelectorAll("[data-hodel]").forEach(b => b.onclick = () => deleteWithUndo("handovers", b.dataset.hodel, "Handover"));
  const dl=$("#hoWhoList"); if(dl) dl.innerHTML = [...new Set(rows.map(r=>r.to))].slice(0,12).map(n=>`<option value="${esc(n)}"></option>`).join("");
}
$("#hoRecAdd").addEventListener("click", () => { if(!canWrite) return toast("Saving isn't switched on here"); $("#hoRecForm").reset(); const d=new Date(); $("#hr_date").value=ymd(Date.now()); $("#hr_from").value=pad2(d.getHours())+":"+pad2(d.getMinutes());
  const g=typeof latestGlucose==="function"?latestGlucose():null; if(g && Date.now()-g.t<2*36e5) $("#hr_gl").value=g.v.toFixed(1);
  const f=log.find(e=>e.type==="food"); if(f && Date.now()-f.at<6*36e5) $("#hr_ate").value=`${f.item||"Food"}${f.carbs!=null?` (${f.carbs} g)`:""} at ${new Date(f.at).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}`;
  $("#hoRecDlg").showModal(); });
$("#hoRecCancel").addEventListener("click", () => $("#hoRecDlg").close());
$("#hoRecForm").addEventListener("submit", async ev => { ev.preventDefault(); const to=$("#hr_to").value.trim(); if(!to) return;
  const day=$("#hr_date").value, at=new Date(day+"T"+($("#hr_from").value||"09:00")).getTime(), u=$("#hr_until").value;
  const body={to, at, until: u ? new Date(day+"T"+u).getTime() : null, where:$("#hr_where").value.trim(), items:[...document.querySelectorAll("#hoRecForm [data-hoitem]:checked")].map(c=>c.dataset.hoitem), gl:num($("#hr_gl").value), ate:$("#hr_ate").value.trim(), note:$("#hr_note").value.trim(), backAt:null, by:me};
  try { await base.collection("handovers").add(body); $("#hoRecDlg").close(); toast("Handover recorded"); } catch { toast("Couldn't save"); } });
$("#hoTabPdf").addEventListener("click", async e => { const btn=e.currentTarget; btn.disabled=true; const m=$("#hoTabMsg"); m.textContent="Making the PDF…";
  try { const r = await htmlToPdf(`<div class="report-paper print ho-paper">${buildHandover()}</div>`, `Handover-sheet-${(P("name")||"T1D").replace(/[^\w-]+/g,"-")}.pdf`); m.textContent = r.status==="saved" ? "PDF saved." : ""; }
  catch(err){ m.textContent = err.code==="declined" ? "Cancelled." : err.message==="load" ? "Couldn't load the PDF maker. Check your internet connection." : (err.message || "Couldn't make the PDF."); } finally { btn.disabled=false; } });

/* --- Page-use counter (shared between both parents) --- */
var usageDocs = {}, usageMine = null, usageTimer = null;
function usageFlush(){ clearTimeout(usageTimer); usageTimer=null; if(usageMine && canWrite && me) base.collection("usage").doc(String(me)).set(usageMine).catch(()=>{}); }
function countView(v){ if(!canWrite || !me) return; if(!usageMine) usageMine = usageDocs[me] ? JSON.parse(JSON.stringify(usageDocs[me])) : {counts:{}, last:{}, since:Date.now()};
  usageMine.counts[v]=(usageMine.counts[v]||0)+1; usageMine.last[v]=Date.now(); if(!usageTimer) usageTimer=setTimeout(usageFlush, 20000);  if(v==="setup") setTimeout(renderUsage,0); }
addEventListener("pagehide", () => { if(usageTimer) usageFlush(); });
document.addEventListener("visibilitychange", () => { if(document.hidden && usageTimer) usageFlush(); });
function renderUsage(){
  const el=$("#usagePanel"); if(!el) return; const docs={...usageDocs}; if(usageMine && me) docs[me]=usageMine;
  const all=Object.values(docs); const since = all.length ? Math.min(...all.map(d=>d.since||Date.now())) : Date.now(); const days=Math.max(0,Math.floor((Date.now()-since)/864e5));
  const label = v => v==="today" ? "Today" : (VIEW_TITLES[v]?.[0] || v);
  const rows = VIEWS.filter(v=>v!=="viewer").map(v=>({v, n:all.reduce((a,d)=>a+(d.counts?.[v]||0),0), last:Math.max(0,...all.map(d=>d.last?.[v]||0))})).sort((a,b)=>b.n-a.n);
  const max = Math.max(1,...rows.map(r=>r.n));
  el.innerHTML = `<div class="ph"><h2>Which pages you use</h2><span class="muted">${all.length>1?"both of you":"this account"} · ${days<1?"started today":days+" day"+(days===1?"":"s")}</span></div>
    <div class="muted" style="font-size:.9rem">Counts how often each page is opened, on all your phones. After a month, pages nobody opens are the ones to hide or remove so the app stays simple.</div>
    <ul class="use-list">${rows.map(r=>{ const stale = days>=21 && (!r.last || Date.now()-r.last>21*864e5); return `<li class="${stale?"stale":""}"><span>${esc(label(r.v))}</span><span class="use-bar"><i style="width:${Math.round(r.n/max*100)}%"></i></span><b>${r.n}</b>${stale?`<small>not opened in 3 weeks</small>`:""}</li>`; }).join("")}</ul>`;
}

/* --- Month in short (Today, first week of each month) --- */
function monthSummary(){
  const d=new Date(), a=new Date(d.getFullYear(), d.getMonth()-1, 1).getTime(), b=new Date(d.getFullYear(), d.getMonth(), 1).getTime()-1;
  const pa=new Date(d.getFullYear(), d.getMonth()-2, 1).getTime(), pb=a-1;
  const st=statsOf(ptsBetween(a,b)), sp=statsOf(ptsBetween(pa,pb)), L=log.filter(e=>e.at>=a&&e.at<=b);
  return {a, b, name:new Date(a).toLocaleDateString("en-GB",{month:"long"}), st, sp, hyposCgm: st?hypoEpisodes(ptsBetween(a,b)).length:null, hyposP: sp?hypoEpisodes(ptsBetween(pa,pb)).length:null,
    hyposLogged:L.filter(e=>e.type==="hypo").length, sites:L.filter(e=>e.type==="site").length, ill:new Set(L.filter(e=>e.type==="ill").map(e=>ymd(e.at))).size,
    changes:(typeof setChanges!=="undefined"?setChanges:[]).filter(c=>c.at>=a&&c.at<=b), labs:(typeof a1c!=="undefined"?a1c:[]).filter(x=>x.at>=a&&x.at<=b)};
}
function renderMonthCard(){
  const el=$("#monthCard"); if(!el) return; const d=new Date(), key=`t1d-month-${d.getFullYear()}-${d.getMonth()}`; let hid=false; try { hid=!!localStorage.getItem(key); } catch {}
  if(d.getDate()>7 || hid || log.length<5){ el.hidden=true; return; }
  const m=monthSummary(); const diff=(x,y,good,unit="",dp=0)=> x==null||y==null ? "" : (()=>{ const v=x-y; if(Math.abs(v)<(dp?0.1:1)) return `<i>same as before</i>`; return `<i class="${(good?v>0:v<0)?"up":"down"}">${v>0?"▲":"▼"} ${Math.abs(v).toFixed(dp)}${unit}</i>`; })();
  el.hidden=false;
  el.innerHTML = `<div class="ph"><h2>${esc(m.name)} in short</h2><button type="button" class="small" id="monthHide">Hide</button></div>
    <div class="mo-grid">
      <div><span class="k">In range</span><b>${m.st?Math.round(m.st.tir)+"%":"–"}</b>${m.st&&m.sp?diff(m.st.tir,m.sp.tir,true,"%"):""}</div>
      <div><span class="k">Average</span><b>${m.st?m.st.mean.toFixed(1):"–"}</b>${m.st&&m.sp?diff(m.st.mean,m.sp.mean,false,"",1):""}</div>
      <div><span class="k">CGM hypos</span><b>${m.hyposCgm??"–"}</b>${m.hyposCgm!=null&&m.hyposP!=null?diff(m.hyposCgm,m.hyposP,false):""}</div>
      <div><span class="k">Treated</span><b>${m.hyposLogged}</b><i>hypos logged</i></div>
    </div>
    <p class="muted" style="margin:0;font-size:.9rem">${[m.sites?`${m.sites} site change${m.sites===1?"":"s"}`:"", m.ill?`unwell on ${m.ill} day${m.ill===1?"":"s"}`:"", m.changes.length?`${m.changes.length} pump setting change${m.changes.length===1?"":"s"}`:"", m.labs.length?`HbA1c ${Math.round(m.labs[0].mmol)}`:"", !m.st?"No CGM data imported for that month":""].filter(Boolean).join(" · ") || "A quiet month."}</p>
    <div class="btns"><button type="button" class="small primary" id="monthReport">Full report</button></div>`;
  $("#monthHide").onclick = () => { try { localStorage.setItem(key,"1"); } catch {} el.hidden=true; };
  $("#monthReport").onclick = () => { location.hash="reports"; setTimeout(()=>{ $("#rpPeriod").value="month"; renderReport(); },60); };
}

/* --- Read-only family view (grandparents) --- */
function isViewer(){ return window.__family?.role === "viewer"; }
var roleShown = null, vArm = null;
function applyRole(){
  const r = isViewer() ? "viewer" : "parent";
  if(r !== roleShown){ roleShown = r; if(r==="viewer") document.documentElement.dataset.role="viewer"; else delete document.documentElement.dataset.role; route(); }
  if(r==="viewer") renderViewer(); else renderViewers();
}
function renderViewer(){
  const F=window.__family||{}, nm=P("name")||"him"; $("#vwName").textContent = `${nm} today`;
  const meName=(F.names||{})[me]||""; $("#vwWho").textContent = meName ? `Signed in as ${meName} · view only` : "View only";
  $("#vwSick").hidden = !sick?.active;
  $("#vwLow").innerHTML = `<b>If he's low (below ${hypoBelow()} mmol/L)</b><p>${esc(P("hypoTx")||"Follow his parents' instructions and call them.").replace(/\n/g,"<br>")}</p>${P("signs")?`<small>Signs: ${esc(P("signs"))}</small>`:""}<small>Unconscious, fitting or can't swallow: nothing by mouth, recovery position, call 999, then his parents.</small>`;
  const since=Date.now()-24*36e5, L=log.filter(e=>e.at>=since).sort((a,b)=>b.at-a.at);
  const who = id => id===me ? "You" : ((F.names||{})[id] || "His parents");
  const what = e => e.type==="food"||e.type==="drink" ? [e.item, e.carbs!=null?e.carbs+" g carbs":"", e.ml?e.ml+" ml":""].filter(Boolean).join(" · ") : [e.gl!=null?e.gl.toFixed(1)+" mmol/L":"", e.note||""].filter(Boolean).join(" · ");
  $("#vwCount").textContent = L.length ? `${L.length} in the last 24 hours` : "";
  $("#vwList").innerHTML = L.length ? L.map(e=>`<li><span class="c-t">${new Date(e.at).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}</span><span><span class="chip ${TYPES[e.type]?.chip||""}">${esc(TYPES[e.type]?.label||e.type)}</span> ${esc(what(e))}<small>${esc(who(e.by))}</small></span></li>`).join("") : `<li class="muted">Nothing logged in the last 24 hours.</li>`;
  $("#vwSheet").innerHTML = buildHandover();
}
$("#vwOut").addEventListener("click", () => window.__signOut ? window.__signOut() : location.reload());
function renderViewers(){
  const el=$("#viewerPanel"); if(!el) return; const F=window.__family;
  const head = `<div class="ph"><h2>Family who can view</h2></div>`;
  if(!F){ el.innerHTML = head + `<p class="muted" style="margin:0">Grandparents and other family can get a view-only login in the shared family app on your phones.</p>`; return; }
  if(F.viewerSql===false){ el.innerHTML = head + `<p class="muted" style="margin:0">One-time setup: in Supabase open <b>SQL Editor → New query</b>, paste the whole of <b>viewer-setup.sql</b> and press <b>Run</b>. Then reload the app.</p>`; return; }
  const vs=F.viewers||[], code=F.viewerCode;
  el.innerHTML = head + `<div class="muted" style="font-size:.9rem">For grandparents or family who babysit. They see the last 24 hours, what to do if he's low, and his care sheet. They can't log, change or delete anything, and can't see older history, glucose files or reports.</div>
    ${code ? `<div class="nur-code"><span class="k">Family view code</span><b>${esc(code)}</b><span class="muted">Only for joining. Switch it off once they've joined.</span></div>
      <div class="btns"><button type="button" class="small primary" id="vwShare">${navigator.share?"Share instructions":"Copy instructions"}</button><button type="button" class="small" id="vwNew">New code</button><button type="button" class="small" id="vwOff">Switch code off</button></div>`
      : `<div class="btns"><button type="button" class="small primary" id="vwOn">Make a family view code</button></div>`}
    <details class="nur-how"><summary>How they get a login</summary><ol>
      <li>New sign-ups are closed, so make their login yourself: Supabase → <b>Authentication → Users → Add user → Create new user</b>. Enter their email and a password and tick <b>Auto confirm user</b>.</li>
      <li>Send them the app link, that email and password, and the code.</li>
      <li>They sign in, type their name (e.g. "Grandma") and tap <b>Join to view</b>.</li>
      <li>Switch the code off here once they've joined.</li></ol></details>
    <h3 style="margin:4px 0 0">Who can view</h3>
    ${vs.length ? `<ul class="nur-list">${vs.map(c=>`<li><span>${esc(c.name)}</span><button type="button" class="small danger" data-vwdel="${esc(c.id)}">${vArm===c.id?"Tap again to remove":"Remove"}</button></li>`).join("")}</ul>` : `<p class="muted" style="margin:0">Nobody yet.</p>`}`;
  const act = async (fn, okMsg) => { try { await fn(); if(okMsg) toast(okMsg); } catch(e){ toast(e.message || "Couldn't do that"); } renderViewers(); };
  $("#vwOn") && ($("#vwOn").onclick = () => act(()=>window.__setViewerCode(true), "Code made"));
  $("#vwNew") && ($("#vwNew").onclick = () => act(()=>window.__setViewerCode(true), "New code made. The old one no longer works"));
  $("#vwOff") && ($("#vwOff").onclick = () => act(()=>window.__setViewerCode(false), "Code switched off. People who joined can still view"));
  $("#vwShare") && ($("#vwShare").onclick = async () => { const t = `This is the app we use for ${P("name")||"our son"}'s diabetes. You can see what's happened today and what to do if he's low.\n1. Open ${location.origin}\n2. Sign in with the email and password we gave you\n3. Type your name, tap "Join to view" and enter this code: ${code}`;
    if(navigator.share){ try { await navigator.share({text:t}); } catch {} } else { try { await navigator.clipboard.writeText(t); toast("Copied"); } catch { toast("Couldn't copy"); } } });
  el.querySelectorAll("[data-vwdel]").forEach(b => b.onclick = () => { const id=b.dataset.vwdel; if(vArm!==id){ vArm=id; renderViewers(); setTimeout(()=>{ if(vArm===id){ vArm=null; renderViewers(); } }, 4000); return; } vArm=null; act(()=>window.__removeViewer(id), "Removed. They can't see anything now"); });
}
renderHandoverTab(); renderUsage(); renderMonthCard();

/* ================= v7.1 ================= */
/* --- Site rotation map --- */
const SPOTS = {
  tummyR:{label:"Tummy (his right)", v:"f", x:47, y:98, w:27, h:44}, tummyL:{label:"Tummy (his left)", v:"f", x:76, y:98, w:27, h:44},
  thighR:{label:"Thigh (his right)", v:"f", x:50, y:162, w:22, h:44}, thighL:{label:"Thigh (his left)", v:"f", x:78, y:162, w:22, h:44},
  armL:{label:"Arm (his left)", v:"b", x:180, y:68, w:14, h:42}, armR:{label:"Arm (his right)", v:"b", x:256, y:68, w:14, h:42},
  bottomL:{label:"Bottom (his left)", v:"b", x:197, y:124, w:27, h:36}, bottomR:{label:"Bottom (his right)", v:"b", x:226, y:124, w:27, h:36}
};
const spotLabel = k => SPOTS[k]?.label || "";
function spotHistory(kind){ return log.filter(e=>e.type===kind && e.spot && SPOTS[e.spot]).sort((a,b)=>b.at-a.at); }
function spotInfo(kind){
  const h=spotHistory(kind), lastUse={}, rank={};
  h.forEach((e,i)=>{ if(lastUse[e.spot]==null){ lastUse[e.spot]=e.at; rank[e.spot]=i; } });
  const used=Object.keys(lastUse);
  const pool = used.length>=2 ? used : Object.keys(SPOTS);
  const suggest = h.length ? pool.filter(k=>k!==h[0].spot).sort((a,b)=>(lastUse[a]??0)-(lastUse[b]??0))[0] : null;
  return {h, lastUse, rank, suggest};
}
function spotSvg(kind, selected, clickable){
  const {rank, suggest} = spotInfo(kind);
  const fill = k => rank[k]==null ? "var(--surface)" : rank[k]===0 ? "var(--page)" : rank[k]<3 ? "color-mix(in srgb,var(--page) 55%,var(--surface))" : rank[k]<6 ? "color-mix(in srgb,var(--page) 25%,var(--surface))" : "color-mix(in srgb,var(--page) 10%,var(--surface))";
  const body = cx => `<circle cx="${cx}" cy="40" r="20" class="sp-body"/><rect x="${cx-30}" y="62" width="60" height="92" rx="22" class="sp-body"/><rect x="${cx-46}" y="64" width="16" height="74" rx="8" class="sp-body"/><rect x="${cx+30}" y="64" width="16" height="74" rx="8" class="sp-body"/><rect x="${cx-27}" y="148" width="25" height="96" rx="11" class="sp-body"/><rect x="${cx+2}" y="148" width="25" height="96" rx="11" class="sp-body"/>`;
  const zones = Object.entries(SPOTS).map(([k,z])=>`<rect ${clickable?`data-spot="${k}" tabindex="0" role="button" aria-label="${z.label}"`:""} class="sp-zone ${k===selected?"sel":""} ${k===suggest?"sug":""}" x="${z.x}" y="${z.y}" width="${z.w}" height="${z.h}" rx="7" fill="${k===selected?"var(--page)":fill(k)}"/>${k===suggest&&k!==selected?`<text x="${z.x+z.w/2}" y="${z.y+z.h/2+4}" text-anchor="middle" class="sp-star">★</text>`:""}${k===selected?`<text x="${z.x+z.w/2}" y="${z.y+z.h/2+5}" text-anchor="middle" class="sp-tick">✓</text>`:""}`).join("");
  return `<svg viewBox="0 0 300 262" class="spot-svg" role="img" aria-label="Where on his body">${body(75)}${body(225)}${zones}<text x="75" y="258" text-anchor="middle" class="sp-cap">Front</text><text x="225" y="258" text-anchor="middle" class="sp-cap">Back</text></svg>`;
}
function renderSpotPicker(){
  const w=$("#e_spotWrap"); if(!w) return; const kind=$("#e_type").value; if(!["site","sensor"].includes(kind)) return;
  const sel=$("#e_spot").value, info=spotInfo(kind);
  w.querySelector(".sp-map").innerHTML = spotSvg(kind, sel, true);
  const days = k => info.lastUse[k] ? Math.floor((Date.now()-info.lastUse[k])/864e5) : null;
  w.querySelector(".sp-msg").innerHTML = sel ? `<b>${esc(spotLabel(sel))}</b>${days(sel)!=null?` · last used ${days(sel)===0?"today":days(sel)+" day"+(days(sel)===1?"":"s")+" ago"}`:" · not used before"}${info.h[0]?.spot===sel?` <span class="warn">Same place as last time</span>`:""}`
    : info.suggest ? `Suggested: <b>${esc(spotLabel(info.suggest))}</b> ★ ${days(info.suggest)!=null&&days(info.suggest)>0?`(not used for ${days(info.suggest)} day${days(info.suggest)===1?"":"s"})`:""}. Tap where it went.` : `Tap where it went.`;
  w.querySelectorAll("[data-spot]").forEach(z => { const pick = () => { $("#e_spot").value = $("#e_spot").value===z.dataset.spot ? "" : z.dataset.spot; renderSpotPicker(); };
    z.addEventListener("click", pick); z.addEventListener("keydown", ev => { if(ev.key==="Enter"||ev.key===" "){ ev.preventDefault(); pick(); } }); });
}
function rotPanel(el, kind){
  const info=spotInfo(kind), word = kind==="site" ? "infusion set" : "sensor";
  el.innerHTML = `<div class="ph"><h2>${kind==="site"?"Infusion set spots":"Sensor spots"}</h2></div>
    <div class="rot-grid"><div class="sp-map">${spotSvg(kind, null, false)}</div>
    <div class="rot-side">${info.suggest?`<div class="rot-next"><span class="k">Next spot</span><b>${esc(spotLabel(info.suggest))}</b><span class="muted">${(()=>{ const t=info.lastUse[info.suggest]; if(!t) return "not used yet"; const d=Math.floor((Date.now()-t)/864e5); return d<1?"last used today":`not used for ${d} day${d===1?"":"s"}`; })()}</span></div>`:""}
      ${info.h.length ? `<ul class="rot-list">${info.h.slice(0,6).map(e=>`<li><span>${esc(spotLabel(e.spot))}</span><small>${fmtD(e.at).replace(/ \d{4}$/,"")}</small></li>`).join("")}</ul>` : `<p class="muted" style="margin:0">When you log a ${word} change, tap where it went. The map fills in and suggests the next spot.</p>`}
      <div class="rot-key"><span><i style="background:var(--page)"></i>Last time</span><span><i style="background:color-mix(in srgb,var(--page) 55%,var(--surface))"></i>Recent</span><span><i style="background:color-mix(in srgb,var(--page) 12%,var(--surface))"></i>A while ago</span><span>★ Suggested</span></div></div></div>`;
}
function renderRotation(){
  const a=$("#rotSite"), b=$("#rotSensor"); if(!a||!b) return; rotPanel(a,"site"); rotPanel(b,"sensor");
  const ss=siteStatus(), se=sensorStatus(); $("#stSiteN").textContent = ss.v==="—" ? "Set and reservoir" : `${ss.v} · ${ss.n.toLowerCase()}`; $("#stSensorN").textContent = se.v==="—" ? (P("cgm")||"Sensor") : `${se.v} · ${se.n.toLowerCase()}`;
  const H=log.filter(e=>e.type==="site"||e.type==="sensor").sort((x,y)=>y.at-x.at).slice(0,12);
  $("#rotHist").innerHTML = H.length ? H.map(e=>`<li><button type="button" class="rh-row" data-rhedit="${esc(e.id)}"><span class="chip blue">${e.type==="site"?"Infusion set":"Sensor"}</span><span class="rh-where">${e.spot?esc(spotLabel(e.spot)):`<i>Where? Tap to add</i>`}</span><small>${fmtT(e.at)}</small></button></li>`).join("")
    : `<li class="muted">No changes logged yet.</li>`;
  document.querySelectorAll("[data-rhedit]").forEach(bn => bn.onclick = () => openEdit(bn.dataset.rhedit));
}
/* --- Bedtime check (Today, evenings) --- */
var bedDocs = {};
const BED_ITEMS = [["bg","Check his glucose"],["kit","Hypo treatment and meter by his bed"],["res","Enough insulin in the reservoir for the night"],["batt","Pump battery OK"],["cgm","Sensor working, CGM alerts on and loud"]];
function bedKey(){ const d=new Date(); if(d.getHours()<4) d.setDate(d.getDate()-1); return ymd(d.getTime()); }
function renderBedtime(){
  const el=$("#bedCard"); if(!el) return; const ae=document.activeElement; if(el.contains(ae) && ae.tagName==="INPUT" && ae.type==="number") return; const h=new Date().getHours(), key=bedKey(), doc=bedDocs[key]||{};
  if(!(h>=18 || h<4)){ el.hidden=true; return; } el.hidden=false;
  const s=lastSite(), due = s ? s.at+(num(plan.siteDays)??3)*864e5 : null, siteSoon = due && due-Date.now() < 12*36e5;
  const items = [...BED_ITEMS, ...(siteSoon?[["site", due<Date.now()?"Site change is overdue: change it now":"Site change due overnight: change it now or set an alarm"]]:[])];
  const ticks=doc.ticks||{}, n=items.filter(([k])=>ticks[k]).length, who = id => id===me ? "you" : ((window.__family?.names||{})[id] || names[id] || "the other parent");
  if(doc.done){ el.className="panel bed-card done"; el.innerHTML = `<div class="ph"><h2>Bedtime check done</h2><button type="button" class="small" id="bedUndo">Undo</button></div><p class="muted" style="margin:0">By ${esc(who(doc.by))} at ${new Date(doc.doneAt).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}${doc.gl!=null?` · BG ${doc.gl.toFixed(1)}`:""} · ${n} of ${items.length} ticked${n<items.length?`: missed ${items.filter(([k])=>!ticks[k]).map(([,t])=>t.toLowerCase()).join(", ")}`:""}</p>`;
    $("#bedUndo").onclick = () => bedWrite({done:false, doneAt:null}); return; }
  const g = typeof latestGlucose==="function" ? latestGlucose() : null, gRecent = g && Date.now()-g.t < 30*6e4;
  el.className = "panel bed-card" + (h>=21 || h<4 ? " late" : "");
  el.innerHTML = `<div class="ph"><h2>Bedtime check</h2><span class="muted">${n} of ${items.length}</span></div>
    <ul class="do-pack">${items.map(([k,t])=>`<li><label class="do-chk ${ticks[k]?"on":""}"><input type="checkbox" data-bed="${k}" ${ticks[k]?"checked":""}><span><b>${esc(t)}</b>${k==="bg"?`<small>${gRecent?`CGM says ${g.v.toFixed(1)} (${ago(g.t)})`:"Finger prick or CGM"}</small>`:""}</span></label>${k==="bg"?`<input id="bedGl" class="bed-gl" type="number" step="0.1" min="1" max="33" inputmode="decimal" placeholder="mmol/L" value="${doc.gl!=null?doc.gl:(gRecent?g.v.toFixed(1):"")}" aria-label="His glucose at bedtime">`:""}</li>`).join("")}</ul>
    <div class="btns"><button type="button" class="primary" id="bedDone">Done</button><span class="muted" style="align-self:center;font-size:.85rem">The other parent will see it's done.</span></div>`;
  el.querySelectorAll("[data-bed]").forEach(c => c.onchange = () => bedWrite({ticks:{...(bedDocs[key]?.ticks||{}), [c.dataset.bed]:c.checked}}));
  $("#bedDone").onclick = async () => { const gl=num($("#bedGl").value); await bedWrite({done:true, doneAt:Date.now(), gl, ticks:{...(bedDocs[key]?.ticks||{}), ...(gl!=null?{bg:true}:{})}});
    if(gl!=null && !(gRecent && Math.abs(g.v-gl)<0.05)) { try { await base.collection("log").add({type:"bg", at:Date.now(), gl, ket:null, carbs:null, note:"Bedtime check", by:me}); } catch {} }
    toast("Bedtime check done"); };
}
async function bedWrite(patch){ if(!canWrite) return toast("Saving isn't switched on here"); const key=bedKey(); const doc={date:key, ...(bedDocs[key]||{}), ...patch, by:me, at:Date.now()}; bedDocs[key]=doc; renderBedtime();
  try { await base.collection("bedtime").doc(key).set(doc); } catch { toast("Couldn't save"); } }
setInterval(() => { if(!document.hidden) renderBedtime(); }, 5*60e3);
renderRotation(); renderBedtime();

/* tap outside a pop-up to close it */
document.querySelectorAll("dialog").forEach(dg => dg.addEventListener("click", e => { if(e.target!==dg) return; const r=dg.getBoundingClientRect(); if(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom) dg.close(); }));
/* ---------- storage ---------- */
function note(msg){ const n=$("#dbNote"); n.hidden=!msg; n.textContent=msg||""; }
renderAll();
(async () => {
  const [dbx, user] = await Promise.all([window.claude?.use?.("db") ?? null, window.claude?.use?.("user") ?? null]);
  let uid = null; try { uid = user ? await user.id() : null; } catch {}
  if(!dbx){ renderAll(); await explainNoSave(); return; }
  db = dbx; userCap = user; me = uid; base = db.doc("family/t1d"); canWrite = true;
  let timer = setTimeout(() => note("Can't reach the saved data yet. What you see may be out of date."), 10000);
  const ok = s => { if(!s.metadata.fromCache){ clearTimeout(timer); note(""); } };
  const fail = e => { clearTimeout(timer); note(e.code==="revoked" ? "" : "The saved data couldn't be loaded. Reopen the page to try again."); };
  base.onSnapshot(s => { plan = s.exists ? s.data() : {}; fillPlan(); renderAll(); ok(s); }, fail);
  base.collection("log").orderBy("at","desc").limit(1000).onSnapshot(s => { log = s.docs.map(d => ({id:d.id, ...d.data()})); loadNames(); renderAll(); paCheck(); ok(s); }, fail);
  base.collection("pump").orderBy("day","desc").limit(120).onSnapshot(s => { pumpDays = {}; s.docs.forEach(d => { const x=d.data(); if(x?.day && Array.isArray(x.ev)) pumpDays[x.day]=x; }); renderGlucose(); renderAll(); ok(s); }, fail);
  base.collection("cgm").orderBy("day","desc").limit(120).onSnapshot(s => { cgmDays = {}; s.docs.forEach(d => { const x=d.data(); if(x?.day && Array.isArray(x.pts)) cgmDays[x.day]=x; }); renderGlucose(); renderLog(); renderMealPatterns(); renderComing(); ok(s); }, fail);
  base.collection("growth").onSnapshot(s => { growth = s.docs.map(d => ({id:d.id, ...d.data()})); renderChecks(); renderStrip(); renderLog(); ok(s); }, fail);
  base.collection("appts").onSnapshot(s => { appts = s.docs.map(d => ({id:d.id, ...d.data()})); renderAppts(); renderChecks(); renderComing(); renderStrip(); ok(s); }, fail);
  base.collection("checks").onSnapshot(s => { checks = s.docs.map(d => ({id:d.id, ...d.data()})); renderChecks(); renderAppts(); renderStrip(); ok(s); }, fail);
  base.collection("meta").doc("delivery").onSnapshot(s => { delivery = s.exists ? s.data() : {}; renderDeliveries(); renderSupplies(); renderComing(); }, () => {});
  base.collection("dayplans").onSnapshot(s => { dayplans = {}; s.docs.forEach(d => { const x=d.data(); if(x?.date) dayplans[x.date]=x; }); renderDayOut(); renderComing(); }, () => {});
  base.collection("meta").doc("sickday").onSnapshot(s => { sick = s.exists ? s.data() : {}; renderSick2(); paCheck(); }, () => {});
  base.collection("meta").doc("filebackup").onSnapshot(s => { fileBackupAt = s.exists ? s.data().at : null; renderBackup(); }, () => {});
  base.collection("meta").doc("backup").onSnapshot(s => { lastBackup = s.exists ? s.data().at : null; renderBackup(); renderStrip(); }, () => {});
  base.collection("supplies").onSnapshot(s => { supplies = s.docs.map(d => ({id:d.id, ...d.data()})); renderAll(); ok(s); }, fail);
  base.collection("meta").doc("push").onSnapshot(s => { pushMeta = s.exists ? s.data() : {}; renderPush(); }, () => {});
  base.collection("meta").doc("offsite").onSnapshot(s => { offsite = s.exists ? s.data().at : null; renderSetup(); }, () => {});
  base.collection("meta").doc("pushlog").onSnapshot(s => { remindersRun = s.exists ? s.data().at : null; renderSetup(); }, () => {});
  checkPushSub();
  base.collection("hba1c").onSnapshot(s => { a1c = s.docs.map(d => ({id:d.id, ...d.data()})); renderA1c(); }, () => {});
  base.collection("rx").doc("main").onSnapshot(s => { rx = s.exists ? s.data() : {}; renderRx(); renderComing(); }, () => {});
  base.collection("settings").onSnapshot(s => { setChanges = s.docs.map(d => ({id:d.id, ...d.data()})); renderSettings(); }, () => {});
  base.collection("todos").onSnapshot(s => { todos = s.docs.map(d => ({id:d.id, ...d.data()})); renderTodos(); renderComing(); }, () => {});
  base.collection("trips").onSnapshot(s => { trips = {}; s.docs.forEach(d => { const x=d.data(); if(x?.id) trips[x.id]=x; }); const a=document.activeElement; if(!($("#tripPanel")?.contains(a) && a.tagName==="INPUT" && a.type!=="checkbox")) renderTrip(); renderComing(); }, () => {});
  base.collection("handovers").onSnapshot(s => { handovers = s.docs.map(d => ({id:d.id, ...d.data()})); renderHandoverTab(); }, () => {});
  base.collection("usage").onSnapshot(s => { usageDocs = {}; s.docs.forEach(d => usageDocs[d.id] = d.data()); renderUsage(); }, () => {});
  base.collection("bedtime").orderBy("date","desc").limit(14).onSnapshot(s => { bedDocs = {}; s.docs.forEach(d => { const x=d.data(); if(x?.date) bedDocs[x.date]=x; }); renderBedtime(); }, () => {});
  base.collection("foods").onSnapshot(s => { foods = s.docs.map(d => ({id:d.id, ...d.data()})); renderFavs(); renderItems(); }, () => {});
})();
async function loadNames(){
  if(!userCap?.profiles) return;
  const ids = [...new Set(log.map(e=>e.by).filter(id => id && id!==me && !(id in names)))];
  if(!ids.length) return;
  try { const ps = await userCap.profiles(ids); ids.forEach(id => names[id] = ps?.[id]?.name || ""); renderAll(); } catch {}
}
/* One-time move of anything saved in the earlier private version into the shared space */
async function migrateFromPrivate(uid){
  const old = db.doc("data/users/"+uid+"/t1d");
  const [shared, op, ol, os] = await Promise.all([base.get(), old.get(), old.collection("log").get(), old.collection("supplies").get()]);
  if(shared.exists || old.get && op.exists && op.data().migrated) return;
  if(!op.exists && ol.empty && os.empty) return;
  if(op.exists) await base.set({...op.data(), updatedAt: Date.now()}); else await base.set({updatedAt: Date.now()});
  for (const d of ol.docs) await base.collection("log").doc(d.id).set({...d.data(), by: d.data().by || uid});
  for (const d of os.docs) await base.collection("supplies").doc(d.id).set(d.data());
  await old.set({...(op.exists?op.data():{}), migrated: true});
  toast("Moved your earlier entries into the shared log");
}
setInterval(() => { if(!document.hidden){ renderStrip(); } }, 60000);
})();
