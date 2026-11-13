// T1D Care Hub — app logic (generated from the master page)
// JSZip, XLSX and pdfjsLib are loaded by <script> tags in index.html

(() => {
const TYPES = {
  hypo:{label:"Hypo treated",chip:"low"}, high:{label:"High / correction",chip:"high"}, ketones:{label:"Ketones",chip:"urgent"},
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
const VIEWS = ["today","glucose","food","log","supplies","checks","appts","sick","plan","backup"];
function route(){
  const v = VIEWS.includes(location.hash.slice(1)) ? location.hash.slice(1) : "today";
  VIEWS.forEach(n => $("#v-"+n).hidden = n !== v);
  document.querySelectorAll("#nav a, #moreSheet a").forEach(a => a.getAttribute("href") === "#"+v ? a.setAttribute("aria-current","page") : a.removeAttribute("aria-current"));
  const secOn = !!document.querySelector('#nav a.sec[aria-current="page"]');
  $("#moreBtn").classList.toggle("on", secOn); $("#moreSheet").hidden = true; $("#moreBtn").setAttribute("aria-expanded","false");
  scrollTo({top:0}); if(v==="glucose" && typeof renderGlucose==="function") setTimeout(renderGlucose,0);
}
addEventListener("hashchange", route); route();
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
  const el=$("#heroRing"); if(!el) return; const s=last("site"), days=num(plan.siteDays)??3;
  const R=54, C=2*Math.PI*R; let frac=0, big="—", small="no site logged";
  if(s){ const left=(s.at+days*864e5-Date.now())/864e5; frac=Math.max(0,Math.min(1,left/days)); big = left<0 ? "Due" : left<1 ? Math.max(0,Math.round(left*24))+"h" : left.toFixed(1)+"d"; small = left<0 ? "site overdue" : "site left"; }
  const col = !s ? "var(--hero-muted)" : frac<=0 ? "#ff9b91" : frac<0.2 ? "#ffd28a" : "var(--ring)";
  el.innerHTML = `<svg viewBox="0 0 132 132" width="132" height="132"><circle cx="66" cy="66" r="${R}" fill="none" stroke="var(--hero-chip)" stroke-width="12"/><circle cx="66" cy="66" r="${R}" fill="none" stroke="${col}" stroke-width="12" stroke-linecap="round" stroke-dasharray="${(C*frac).toFixed(1)} ${C.toFixed(1)}" transform="rotate(-90 66 66)"/></svg><div class="hr-in"><b>${big}</b><span>${small}</span></div>`;
}
function renderHero(){ renderHeroRing();
  const h=new Date().getHours(); const nm=(plan.name||"").trim();
  $("#heroTitle").textContent = (h<12?"Good morning":h<18?"Good afternoon":"Good evening") + (nm?`, here's ${nm}'s day`:"");
  $("#heroDate").textContent = new Date().toLocaleDateString("en-GB",{weekday:"long",day:"numeric",month:"long"});
  const site=siteStatus(), t=dayTotals(dayKey(Date.now())), lf=log.find(e=>e.type==="food");
  const wk=log.filter(e=>e.type==="hypo" && e.at>Date.now()-7*864e5).length;
  const chips=[["Site",site.v.replace(" left",""),site.cls],["Carbs",t.carbs+" g",""],["Drinks",t.ml+" ml",""],["Last ate",lf?new Date(lf.at).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}):"—",""],["Hypos 7d",String(wk),wk>=3?"warn":""]];
  $("#heroChips").innerHTML = chips.map(([k,v,c])=>`<div class="hchip ${c}"><span class="k">${k}</span><span class="v">${esc(v)}</span></div>`).join("");
}
function renderAll(){ renderHero(); renderStrip(); renderRecent(); renderFood(); renderLog(); renderSupplies(); renderSick(); renderItems(); renderGlucose(); renderChecks(); renderBackup(); renderMealPatterns(); renderComing(); renderAppts(); }
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
  $("#foodToday").innerHTML = t.es.length ? t.es.slice().reverse().map(e => `<div class="meal">
      <span class="t">${new Date(e.at).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}</span>
      <div style="min-width:0;overflow-wrap:anywhere"><span class="chip ${TYPES[e.type].chip}">${e.type==="food"?esc(e.meal||"Food"):"Drink"}</span>${esc(e.item||"")}${e.type==="food"&&e.ate&&e.ate!=="All"?` <span class="muted">· ate ${esc(e.ate.toLowerCase())}</span>`:""}${e.note?`<div class="muted">${esc(e.note)}</div>`:""}</div>
      <span class="c">${e.carbs!=null?e.carbs+" g":""}${e.type==="drink"&&e.ml!=null?`<div class="muted" style="font-size:.85rem">${e.ml} ml</div>`:""}</span></div>`).join("")
    : `<div class="empty">Nothing logged today. Tap Add food after his next meal or snack.</div>`;
  $("#foodWeek").innerHTML = [0,1,2,3,4,5,6].map(i => { const ms=Date.now()-i*864e5, d=dayTotals(dayKey(ms));
    return `<tr class="dayrow"><td>${i===0?"Today":i===1?"Yesterday":new Date(ms).toLocaleDateString("en-GB",{weekday:"short",day:"numeric",month:"short"})}</td><td class="muted">${d.meals||"—"}</td><td class="num">${d.meals?d.carbs+" g":"—"}</td><td class="num">${d.drinks?d.ml+" ml":"—"}</td></tr>`; }).join("");
}
function renderItems(){
  const seen = new Set(), opts = [];
  for (const e of log) if ((e.type==="food"||e.type==="drink") && e.item && !seen.has(e.item.toLowerCase())) { seen.add(e.item.toLowerCase()); opts.push(e.item); if(opts.length>=80) break; }
  $("#itemList").innerHTML = opts.map(o=>`<option value="${esc(o)}"></option>`).join("");
}

function siteStatus(){
  const s = last("site"); const days = num(plan.siteDays) ?? 3;
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
  $("#strip").innerHTML = [
    `<div class="stat ${site.cls}"><span class="k">Pod / site</span><span class="v">${site.v}</span><span class="n">${site.n}</span></div>`,
    `<div class="stat ${sen.cls}"><span class="k">Sensor</span><span class="v">${sen.v}</span><span class="n">${sen.n}</span></div>`,
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
  return `<li><time datetime="${new Date(e.at).toISOString()}" style="white-space:pre-line">${fmtShort(e.at)}</time>
    <div class="what"><span class="chip ${t.chip}">${t.label}</span>${bits.join(" ")}${e.note? `<div>${esc(e.note)}</div>`:""}${e.by? `<div class="muted" style="font-size:.8rem">Logged by ${e.by===me?"you":esc(names[e.by]||"his mum")}</div>`:""}</div>
    <button class="small danger" data-del="${esc(e.id)}" ${canWrite?"":"disabled"}>Delete</button></li>`;
}
function renderRecent(){
  const r = log.slice(0,6);
  $("#recent").innerHTML = r.length ? r.map(entryHTML).join("") : `<li style="display:block"><div class="empty">Nothing logged yet. Use a quick-log button above. Start with his last pod and sensor change so the countdowns work.</div></li>`;
}
function rangeEntries(){ const days=+$("#range").value, f=$("#filter").value; return log.filter(e => e.at > Date.now()-days*864e5 && (!f || e.type===f)); }
function renderLog(){
  const r = rangeEntries();
  $("#fullLog").innerHTML = r.length ? r.map(entryHTML).join("") : `<li style="display:block"><div class="empty">No entries in this range.</div></li>`;
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
    <td><span class="chip ${s.status==="OK"?"ok":s.status==="Order soon"?"high":"low"}">${s.status}</span></td>
    <td style="text-align:right;white-space:nowrap"><button class="small" data-edit="${esc(s.id)}" ${canWrite?"":"disabled"}>Edit</button> <button class="small danger" data-sdel="${esc(s.id)}" ${canWrite?"":"disabled"}>Delete</button></td></tr>`).join("")
    : `<tr><td colspan="5"><div class="empty">No supplies yet. Add the common items below, then set the real stock you have at home.</div></td></tr>`;
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
const PLAN_KEYS = ["name","dob","cgm","pump","hypo","ketAbove","siteDays","sensorDays","eyeMonths","dentalMonths","hypoTx","highTx","signs","glucagon","teamDay","teamOoh","notes"];
let planDirty = false;
function fillPlan(){ if(planDirty) return; PLAN_KEYS.forEach(k => { const el=$("#p_"+k); if(el) el.value = plan[k] ?? ""; }); }
$("#planForm").addEventListener("input", () => planDirty = true);
$("#planForm").addEventListener("submit", async ev => {
  ev.preventDefault(); if(!base) return;
  const data = {}; PLAN_KEYS.forEach(k => data[k] = $("#p_"+k).value.trim());
  $("#savePlan").disabled = true;
  try { await base.set({...data, updatedAt: Date.now()}); planDirty=false; plan=data; renderAll(); toast("Care plan saved"); }
  catch(e){ toast(e.code==="quota_exceeded" ? "Storage is full" : `Couldn't save (${e.code||e.message||"error"}). Try again`); }
  finally { $("#savePlan").disabled = false; }
});

/* ---------- entries ---------- */
const dlg = $("#dlg");
$("#e_type").innerHTML = Object.entries(TYPES).map(([k,v])=>`<option value="${k}">${v.label}</option>`).join("");
$("#filter").innerHTML += Object.entries(TYPES).map(([k,v])=>`<option value="${k}">${v.label}</option>`).join("");
function localNow(){ const d=new Date(); d.setMinutes(d.getMinutes()-d.getTimezoneOffset()); return d.toISOString().slice(0,16); }
function openEntry(type, pre={}){
  if(!canWrite){ (toast("Saving isn't switched on here. Open the page in a browser (see the note at the top)"), $("#dbNote").scrollIntoView({behavior:"smooth",block:"center"})); return; }
  $("#entryForm").reset(); $("#e_type").value = type; $("#e_at").value = localNow();
  $("#dlgTitle").textContent = TYPES[type].label; setFields(type);
  if(type==="food"){ const h=new Date().getHours(); $("#e_meal").value = h<10?"Breakfast":h<11.5?"Morning snack":h<14?"Lunch":h<16.5?"Afternoon snack":h<19?"Dinner":"Bedtime snack"; }
  if(pre.gl!=null) $("#e_gl").value = pre.gl; if(pre.ket!=null) $("#e_ket").value = pre.ket;
  dlg.showModal();
}
function setFields(type){
  document.querySelectorAll("#entryForm [data-for]").forEach(el => el.hidden = !el.dataset.for.split(" ").includes(type));
  const fd = type==="food"||type==="drink";
  document.querySelectorAll("#e_gl,#e_ket").forEach(el => el.closest("label").hidden = fd);
  $("#e_itemLabel").textContent = type==="drink" ? "Drink" : "What he had";
  $("#e_item").placeholder = type==="drink" ? "e.g. Whole milk" : "e.g. Weetabix + whole milk";
  $("#e_carbsLabel").innerHTML = type==="food" ? "Carbs he ate <small>g</small>" : type==="hypo" ? "Carbs given <small>g</small>" : "Carbs <small>g</small>";
  $("#e_noteLabel").textContent = fd ? "Notes (optional)" : "What happened / what you did";
}
$("#e_type").addEventListener("change", e => setFields(e.target.value));
$("#e_item").addEventListener("change", () => {
  const v = $("#e_item").value.trim().toLowerCase(); if(!v) return;
  const prev = log.find(e => (e.type==="food"||e.type==="drink") && (e.item||"").toLowerCase()===v);
  if(prev){ if(prev.carbs!=null && $("#e_carbs").value==="") $("#e_carbs").value = prev.carbs; if(prev.ml!=null && $("#e_ml").value==="") $("#e_ml").value = prev.ml; }
});
document.querySelectorAll("[data-q],[data-q2]").forEach(b => b.addEventListener("click", () => openEntry(b.dataset.q || b.dataset.q2)));
$("#dlgCancel").addEventListener("click", () => dlg.close());
$("#entryForm").addEventListener("submit", async ev => {
  ev.preventDefault();
  const e = { type:$("#e_type").value, at:new Date($("#e_at").value).getTime(), gl:num($("#e_gl").value), ket:num($("#e_ket").value), carbs:num($("#e_carbs").value), note:$("#e_note").value.trim(), by:me };
  if(e.type==="food"||e.type==="drink"){ e.item=$("#e_item").value.trim(); if(e.type==="food"){ e.meal=$("#e_meal").value; e.ate=$("#e_ate").value; } else { e.ml=num($("#e_ml").value); } e.gl=null; e.ket=null; }
  if(isNaN(e.at)) e.at = Date.now();
  $("#dlgSave").disabled = true;
  try {
    await base.collection("log").add(e);
    const linked = supplies.filter(s => s.link && s.link === e.type && s.qty > 0);
    for (const s of linked) await base.collection("supplies").doc(s.id).update({qty: s.qty-1});
    dlg.close(); toast(linked.length ? "Saved · stock updated" : "Saved");
  } catch(err){ toast(err.code==="quota_exceeded" ? "Storage is full. Delete old entries" : `Couldn't save (${err.code||err.message||"error"}). Try again`); }
  finally { $("#dlgSave").disabled = false; }
});
let pendingDel = null;
document.addEventListener("click", async ev => {
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
$("#copyClinic").addEventListener("click", e => copy($("#clinicText").textContent, e.currentTarget));
["#ketIn","#glIn","#vomIn"].forEach(s => $(s).addEventListener("input", renderSick));
$("#logKet").addEventListener("click", () => openEntry("ketones", {ket:num($("#ketIn").value), gl:num($("#glIn").value)}));

/* ---------- supplies ---------- */
const supDlg = $("#supDlg"); let editing = null;
function openSup(s){ if(!canWrite) return (toast("Saving isn't switched on here. Open the page in a browser (see the note at the top)"), $("#dbNote").scrollIntoView({behavior:"smooth",block:"center"})); editing = s||null; $("#supForm").reset();
  if(s){ $("#s_name").value=s.name; $("#s_qty").value=s.qty; $("#s_days").value=s.days; $("#s_lead").value=s.lead??14; $("#s_link").value=s.link||""; }
  supDlg.showModal(); }
$("#addSup").addEventListener("click", () => openSup());
$("#supCancel").addEventListener("click", () => supDlg.close());
$("#supForm").addEventListener("submit", async ev => {
  ev.preventDefault();
  const d = { name:$("#s_name").value.trim(), qty:+$("#s_qty").value||0, days:+$("#s_days").value||1, lead:+$("#s_lead").value||0, link:$("#s_link").value };
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
  planDirty = true; toast(pr.note + " Check with his team, then save.");
});
const isMedtronic = () => /medtronic|minimed|780g|640g/i.test(plan.pump||"");

/* ================= GLUCOSE ================= */
const MG = 18.0182;
let cgmDays = {};            // key YYYY-MM-DD -> {day, pts:[minute, tenths, ...]}
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
function detectColumns(row){
  const low = row.map(c=>String(c).toLowerCase());
  if(!low.some(c=>/glucose|cgm|egv|sgv|\bbg\b|reading|mmol|mg\/dl|value/.test(c))) return null;
  if(low.some(c=>/^\d/.test(c)) && !low.some(c=>/[a-z]{3}/.test(c))) return null;
  const pick = res => { for(const r of res){ const k=low.findIndex(c=>r.test(c)); if(k>=0) return k; } return -1; };
  const v = pick([/historic glucose/, /cgm glucose/, /sensor glucose/, /glucose value/, /^sgv$/, /^glucose/, /egv/, /^(?!.*scan)(?!.*bg reading)(?!.*meter).*glucose/, /mmol|mg\/dl/, /^(bg|reading|value)$/]);
  if(v<0) return null;
  const dOnly = low.findIndex(c=>/^date$|^day$/.test(c)), tOnly = low.findIndex(c=>/^time$/.test(c));
  let t=-1, d=-1;
  if(dOnly>=0 && tOnly>=0){ d=dOnly; t=tOnly; } else t = pick([/device timestamp/, /^timestamp/, /timestamp/, /date.?time/, /^datestring$/, /^date$/, /^time$/, /time/, /date/]);
  if(t<0) return null;
  return { v, t, d, e: low.findIndex(c=>/^event type$|^type$/.test(c)), unit: /mg\/dl|^sgv$/.test(low[v]) ? "mg" : /mmol/.test(low[v]) ? "mmol" : "auto" };
}
function extract(rows, map, rescan){
  const pts=[]; let m = map;
  rows.forEach((r, i) => {
    if(i < (m.start||0)) return;
    if(rescan){ const h = detectColumns(r); if(h){ m = {...h, start:0}; return; } }
    if(m.e>=0 && r[m.e] && /[a-z]/i.test(r[m.e]) && !/egv|sgv|sensor|cgm|glucose/i.test(r[m.e])) return;
    const v = toMmol(r[m.v], m.unit); if(v==null) return;
    const t = parseWhen(m.d>=0 ? joinDT(r[m.d], r[m.t]) : r[m.t]); if(isNaN(t)) return;
    if(t > Date.now()+864e5 || t < Date.UTC(2000,0,1)) return;
    pts.push({t, v});
  });
  return pts;
}
function autoExtract(rows){
  for(let i=0;i<Math.min(rows.length,80);i++){ const h = detectColumns(rows[i]); if(h){ const pts = extract(rows, {...h, start:i+1}, true); if(pts.length) return pts; } }
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
    let pts=[]; for(const t of tables) pts = pts.concat(autoExtract(t.rows));
    const pdfT = tables.find(t=>t.pdf);
    if(pdfT){ if(!pts.length) pts = patternExtract(pdfT.rows); gStatus("");
      if(pts.length) { confirmPdf(pts.sort((a,b)=>a.t-b.t), pdfT); return; }
      gStatus("I couldn't find readings with times and glucose values in this PDF. Summary reports (like the AGP or overview pages) only show graphs and averages. Pick the columns below if the readings are in a table, or export a CSV."); showMapper(pdfT); return; }
    if(!pts.length){ gStatus(""); showMapper(tables.sort((a,b)=>b.rows.length-a.rows.length)[0]); return; }
    await saveReadings(pts);
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
function renderGlucose(){
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
    $("#gStats").innerHTML = [
      `<div class="stat ${tir<50?"bad":tir<70?"warn":""}"><span class="k">Time in range</span><span class="v">${Math.round(tir)}%</span><span class="n">Target: more than 70%</span></div>`,
      `<div class="stat"><span class="k">Average glucose</span><span class="v">${st.mean.toFixed(1)}</span><span class="n">mmol/L</span></div>`,
      `<div class="stat"><span class="k">Estimated HbA1c (GMI)</span><span class="v">${st.gmiMmol}</span><span class="n">mmol/mol · ${st.gmiPct.toFixed(1)}% · lab test may differ</span></div>`,
      `<div class="stat ${st.cv>36?"warn":""}"><span class="k">Ups and downs (CV)</span><span class="v">${Math.round(st.cv)}%</span><span class="n">${st.cv>36?"Above 36%: quite variable":"36% or less: steady"}</span></div>`,
      `<div class="stat ${st.coverage<70?"warn":""}"><span class="k">Sensor data</span><span class="v">${Math.round(st.coverage)}%</span><span class="n">${st.coverage<70?"Under 70%: numbers less reliable":"Enough data to trust"}</span></div>`
    ].join("");
    const rows = [["vhigh","Very high","above 13.9","under 5%",v=>v<5],["high","High","10.1–13.9","",null],["in","In range","3.9–10.0","over 70%",v=>v>70],["low","Low","3.0–3.8","under 4%",v=>v<4],["vlow","Very low","below 3.0","under 1%",v=>v<1]];
    $("#tirBar").innerHTML = `<div class="tir" role="img" aria-label="Time in range bar">${["vlow","low","in","high","vhigh"].map(k=>{const p=pct(st.b[k],st.n); return p>0?`<i style="width:${p}%;background:var(--${k==="in"?"ok":k})" title="${p.toFixed(1)}%"></i>`:"";}).join("")}</div>
      <div class="tirlist">${rows.map(([k,l,r,t,f])=>{const p=pct(st.b[k],st.n); const hrs=p*24/100; return `<div><span class="sw" style="background:var(--${k==="in"?"ok":k})"></span><span>${l} <span class="muted">${r} mmol/L · about ${hrs<1?Math.round(hrs*60)+" min":hrs.toFixed(1)+" h"} a day</span></span><span class="pc">${p.toFixed(p<10?1:0)}%</span><span class="tg ${f?(f(p)?"good":"bad"):""}">${t?"target "+t:""}</span></div>`;}).join("")}</div>`;
    renderAGP(st); renderInsights(st, days);
  }
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
  const H = CH.t+CH.ph+90, laneY = CH.t+CH.ph+44;
  let g = svgFrame(H, hourTicks());
  // food guides
  evs.filter(e=>e.type==="food").forEach(e=>{ const x=xH((e.at-d0)/36e5); g+=`<line x1="${x}" x2="${x}" y1="${CH.t}" y2="${laneY-8}" stroke="var(--ok)" stroke-opacity=".45" stroke-dasharray="2 3"/>`; });
  // glucose line, coloured by band, broken on gaps
  const col = v => v<3.9 ? "var(--low)" : v>10 ? "var(--high)" : "var(--accent)";
  for(let i=1;i<pts.length;i++){ const a=pts[i-1], b=pts[i]; if(b.t-a.t>20*6e4) continue;
    g+=`<line x1="${xH((a.t-d0)/36e5).toFixed(1)}" y1="${yS(a.v).toFixed(1)}" x2="${xH((b.t-d0)/36e5).toFixed(1)}" y2="${yS(b.v).toFixed(1)}" stroke="${col((a.v+b.v)/2)}" stroke-width="2" stroke-linecap="round"/>`; }
  manual.forEach(e=>{ g+=`<circle cx="${xH((e.at-d0)/36e5)}" cy="${yS(e.gl)}" r="5" fill="var(--surface)" stroke="var(--fg)" stroke-width="2"/>`; });
  // event lane
  g+=`<text x="${CH.l-6}" y="${laneY+4}" text-anchor="end">Events</text><line x1="${CH.l}" x2="${CH.W-CH.r}" y1="${laneY}" y2="${laneY}" stroke="var(--line)"/>`;
  evs.forEach(e=>{ const x=xH((e.at-d0)/36e5), c=(EVT[e.type]||{c:"var(--muted)"}).c;
    if(e.type==="food"){ g+=`<circle cx="${x}" cy="${laneY}" r="6" fill="${c}" stroke="var(--surface)" stroke-width="2"/>${e.carbs!=null?`<text x="${x}" y="${laneY+22}" text-anchor="middle" style="fill:var(--fg);font-weight:700">${e.carbs}g</text>`:""}`; }
    else if(e.type==="drink"){ g+=`<rect x="${x-4.5}" y="${laneY-4.5}" width="9" height="9" transform="rotate(45 ${x} ${laneY})" fill="${c}" stroke="var(--surface)" stroke-width="2"/>`; }
    else if(e.type==="hypo"){ g+=`<path d="M${x} ${laneY-7} L${x+7} ${laneY+6} L${x-7} ${laneY+6} Z" fill="${c}" stroke="var(--surface)" stroke-width="2"/>`; }
    else { g+=`<rect x="${x-5}" y="${laneY-5}" width="10" height="10" rx="2" fill="${c}" stroke="var(--surface)" stroke-width="2"/>`; }
  });
  if(!pts.length) g+=`<text x="${(CH.W+CH.l)/2}" y="${yS(12)}" text-anchor="middle" style="font-size:13px">${Object.keys(cgmDays).length?"No CGM data for this day":"No CGM data yet. Import a file above"}</text>`;
  g+=`<line id="dCross" x1="0" x2="0" y1="${CH.t}" y2="${laneY+8}" stroke="var(--fg)" stroke-opacity=".5" visibility="hidden"/><circle id="dDot" r="5" fill="var(--accent)" stroke="var(--surface)" stroke-width="2" visibility="hidden"/>`;
  g+=`<rect id="dHit" x="${CH.l}" y="0" width="${CH.W-CH.l-CH.r}" height="${H}" fill="transparent"/>`;
  const box=$("#dayChart");
  box.innerHTML = `<svg viewBox="0 0 ${CH.W} ${H}" role="img" aria-label="Glucose for ${new Date(d0).toLocaleDateString("en-GB")}">${g}</svg><div class="tip" hidden></div>`;
  hover(box, H, x => {
    const t = d0 + (x-CH.l)/(CH.W-CH.l-CH.r)*864e5;
    let best=null; for(const p of pts) if(!best || Math.abs(p.t-t)<Math.abs(best.t-t)) best=p;
    if(best && Math.abs(best.t-t)>15*6e4) best=null;
    const near = evs.filter(e=>Math.abs(e.at-t)<20*6e4);
    if(!best && !near.length) return null;
    const tt = best?best.t:near[0].at;
    return {x:xH((tt-d0)/36e5), y: best?yS(best.v):null, html:
      `<div class="muted">${new Date(tt).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}</div>`+
      (best?`<div><b>${best.v.toFixed(1)}</b> mmol/L · ${({vlow:"very low",low:"low",in:"in range",high:"high",vhigh:"very high"})[band(best.v)]}</div>`:"")+
      near.map(e=>`<div>${TYPES[e.type]?.label||"Note"}${e.item?": "+esc(e.item):""}${e.carbs!=null?` · ${e.carbs} g`:""}${e.ml!=null?` · ${e.ml} ml`:""}${e.gl!=null?` · ${e.gl.toFixed(1)}`:""}</div>`).join("")};
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

/* ================= COMING DAYS: what's due + when to watch ================= */
function dayLabel(ms){ const d=Math.round((midnight(ms)-midnight(Date.now()))/864e5); return d<0?"Overdue":d===0?"Today":d===1?"Tomorrow":new Date(ms).toLocaleDateString("en-GB",{weekday:"short",day:"numeric",month:"short"}); }
function renderComing(){
  const ag=$("#agenda"), wl=$("#watchList"); if(!ag||!wl) return;
  const now=Date.now(), horizon=now+7*864e5, items=[];
  const add=(at,text,cls="")=>{ if(at<=horizon) items.push({at,text,cls}); };
  const site=last("site"), sensor=last("sensor");
  if(site){ const due=site.at+(num(plan.siteDays)??3)*864e5; add(due, `Site change due${isMedtronic()?" (infusion set + reservoir)":""}`, due<now?"bad":due-now<864e5?"warn":""); }
  if(sensor){ const due=sensor.at+(num(plan.sensorDays)??10)*864e5; add(due, "Sensor change due", due<now?"bad":due-now<864e5?"warn":""); }
  supplyRows().forEach(sp=>{ if(!sp.days) return; const orderBy=now+Math.max(0,(sp.left-(sp.lead??14)))*864e5; add(orderBy, `Order ${sp.name} (${Math.floor(sp.left)} days left)`, sp.left<=(sp.lead??14)?"bad":"warn"); });
  ["eye","dental"].forEach(k=>{ const d=nextDue(k); if(d!=null && d<=now+30*864e5) items.push({at:d, text:`${CK[k].label} due`, cls:d<now?"bad":""}); });
  appts.filter(a=>a.status==="booked" && a.at>=midnight(now) && a.at<=horizon).forEach(a=>items.push({at:a.at, text:`${(AK[a.kind]||AK.other).label}${a.vaccine?" ("+a.vaccine+")":""}${a.hasTime?" at "+new Date(a.at).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}):""}${a.where?" · "+a.where:""}`, cls:""}));
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
  const box=$(sel); if(pts.length<2){ box.innerHTML = `<div class="empty">${pts.length?"Add one more measurement to see a line.":"No measurements yet."}</div>`; return; }
  const W=420, H=200, l=44, r=14, t=12, b=28;
  const t0=pts[0].at, t1=pts[pts.length-1].at, span=Math.max(t1-t0, 30*864e5);
  const vals=pts.map(p=>p[f]); let lo=Math.min(...vals), hi=Math.max(...vals); const pad=Math.max((hi-lo)*.15, f==="h"?1:.2); lo-=pad; hi+=pad;
  const X = ms => l + (ms-t0)/span*(W-l-r), Y = v => t + (1-(v-lo)/(hi-lo))*(H-t-b);
  let g=""; for(let i=0;i<=3;i++){ const v=lo+(hi-lo)*i/3; g+=`<line x1="${l}" x2="${W-r}" y1="${Y(v)}" y2="${Y(v)}" stroke="var(--line)" stroke-width="1"/><text x="${l-6}" y="${Y(v)+4}" text-anchor="end">${v.toFixed(f==="h"?0:1)}</text>`; }
  g+=`<text x="${l}" y="${H-8}">${new Date(t0).toLocaleDateString("en-GB",{month:"short",year:"2-digit"})}</text><text x="${W-r}" y="${H-8}" text-anchor="end">${new Date(t1).toLocaleDateString("en-GB",{month:"short",year:"2-digit"})}</text>`;
  g+=`<path d="M${pts.map(p=>`${X(p.at).toFixed(1)} ${Y(p[f]).toFixed(1)}`).join(" L")}" fill="none" stroke="var(--accent)" stroke-width="2" stroke-linejoin="round"/>`;
  pts.forEach((p,i)=>{ g+=`<circle cx="${X(p.at)}" cy="${Y(p[f])}" r="5" fill="var(--accent)" stroke="var(--surface)" stroke-width="2"/><circle data-i="${i}" cx="${X(p.at)}" cy="${Y(p[f])}" r="14" fill="transparent" style="cursor:pointer"/>`; });
  const last=pts[pts.length-1]; g+=`<text x="${X(last.at)-8}" y="${Y(last[f])-10}" text-anchor="end" style="fill:var(--fg);font-weight:700">${last[f].toFixed(dp)} ${unit}</text>`;
  box.innerHTML = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${f==="h"?"Height":"Weight"} over time">${g}</svg><div class="tip" hidden></div>`;
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
const BK_COLS = ["log","cgm","growth","checks","supplies"];
let lastBackup = null, dl = null;
function bkStatus(m){ const n=$("#bkStatus"); n.hidden=!m; n.textContent=m||""; }
function renderBackup(){
  const el=$("#bkLast"); const days = lastBackup ? Math.floor((Date.now()-lastBackup)/864e5) : null;
  el.className = "stat " + (days===null ? (log.length?"warn":"") : days>30 ? "warn" : "");
  el.querySelector(".v").textContent = lastBackup ? (days===0?"Today":days===1?"Yesterday":days+" days ago") : "—";
  el.querySelector(".n").textContent = lastBackup ? fmtD(lastBackup) : "Never";
  $("#bkCount").textContent = (log.length + Object.keys(cgmDays).length + growth.length + checks.length + supplies.length).toLocaleString("en-GB");
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
    if(r.status==="saved"){ await base.collection("meta").doc("backup").set({at:Date.now(), by:me}); bkStatus(`Backup saved: ${out.log.length} log entries, ${out.cgm.length} days of CGM, ${out.growth.length} measurements, ${out.checks.length} check-ups, ${out.supplies.length} supplies.`); }
    else bkStatus("");
  } catch(err){ bkStatus(err.code==="declined" ? "Backup cancelled." : err.code==="rate_limited" ? "A save prompt is already open. Finish that one first." : "Couldn't make the backup. Try again."); }
  finally { btn.disabled=false; }
});
let pendingRestore = null;
$("#bkFile").addEventListener("change", async ev => {
  const f=ev.target.files[0]; ev.target.value=""; if(!f) return;
  if(!canWrite) return bkStatus("Saving isn't available in this view.");
  try {
    const data = JSON.parse(await f.text());
    if(data.app!=="t1d-care-hub" || !Array.isArray(data.log)) throw new Error("bad");
    pendingRestore = data;
    const c = $("#bkConfirm"); c.hidden=false;
    c.innerHTML = `<strong>Restore backup from ${fmtD(Date.parse(data.exportedAt))}?</strong><span>It has ${data.log.length} log entries, ${(data.cgm||[]).length} days of CGM, ${(data.growth||[]).length} measurements, ${(data.checks||[]).length} check-ups and ${(data.supplies||[]).length} supplies. Anything missing here is added back and matching items are reset to the backup. Nothing else is deleted.</span><div class="btns"><button class="primary" type="button" id="bkGo">Restore now</button><button type="button" id="bkNo">Cancel</button></div>`;
    $("#bkNo").onclick = () => { pendingRestore=null; c.hidden=true; };
    $("#bkGo").onclick = doRestore;
    bkStatus("");
  } catch { bkStatus("That file isn't a T1D Care Hub backup."); }
});
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
  base.collection("log").orderBy("at","desc").limit(1000).onSnapshot(s => { log = s.docs.map(d => ({id:d.id, ...d.data()})); loadNames(); renderAll(); ok(s); }, fail);
  base.collection("cgm").orderBy("day","desc").limit(120).onSnapshot(s => { cgmDays = {}; s.docs.forEach(d => { const x=d.data(); if(x?.day && Array.isArray(x.pts)) cgmDays[x.day]=x; }); renderGlucose(); renderLog(); renderMealPatterns(); renderComing(); ok(s); }, fail);
  base.collection("growth").onSnapshot(s => { growth = s.docs.map(d => ({id:d.id, ...d.data()})); renderChecks(); renderStrip(); renderLog(); ok(s); }, fail);
  base.collection("appts").onSnapshot(s => { appts = s.docs.map(d => ({id:d.id, ...d.data()})); renderAppts(); renderChecks(); renderComing(); renderStrip(); ok(s); }, fail);
  base.collection("checks").onSnapshot(s => { checks = s.docs.map(d => ({id:d.id, ...d.data()})); renderChecks(); renderAppts(); renderStrip(); ok(s); }, fail);
  base.collection("meta").doc("backup").onSnapshot(s => { lastBackup = s.exists ? s.data().at : null; renderBackup(); renderStrip(); }, () => {});
  base.collection("supplies").onSnapshot(s => { supplies = s.docs.map(d => ({id:d.id, ...d.data()})); renderAll(); ok(s); }, fail);
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
