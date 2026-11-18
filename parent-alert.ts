// T1D Care Hub: buzz the other phones straight away when a hypo is logged or sick-day mode starts,
// even when the app is closed. Called by the database trigger in alerts-setup.sql.
// Secrets (Supabase → Edge Functions → Secrets): VAPID_PRIVATE_KEY (same as GitHub), ALERT_SECRET (any long random text).
import webpush from "npm:web-push@3.6.7";
import { createClient } from "npm:@supabase/supabase-js@2";

// Pure part (tested): what, if anything, to send for this change.
export function decide(p) {
  const r = p?.record, o = p?.old_record;
  if (!r || p.table !== "records") return null;
  const d = r.data || {};
  if (r.collection === "log" && d.type === "hypo" && p.type === "INSERT" && Date.now() - (+d.at || 0) < 30 * 60e3) {
    const gl = typeof d.gl === "number" ? d.gl.toFixed(1) + " mmol/L" : "";
    const note = String(d.note || "").replace(/^Hypo mode:?\s*/, "");
    return { kind: "hypo", tag: "h" + r.id, by: r.updated_by, family: r.family_id, url: "/#glucose",
      body: [gl, typeof d.carbs === "number" ? d.carbs + " g given" : "", note].filter(Boolean).join(" · ") || "Open the app to see it." };
  }
  if (r.collection === "meta" && r.id === "sickday" && d.active && !(o?.data?.active)) {
    return { kind: "sick", tag: "s" + (d.start || ""), by: r.updated_by, family: r.family_id, url: "/#sick", body: "Check ketones and glucose. The timer is on the Sick day page." };
  }
  return null;
}
export function titleFor(a, who, child) {
  const name = who || "Someone";
  return a.kind === "hypo" ? `${name} is treating a hypo${child ? " (" + child + ")" : ""}` : `${name} started sick-day mode`;
}

Deno.serve(async (req) => {
  const secret = Deno.env.get("ALERT_SECRET");
  if (!secret || req.headers.get("x-alert-secret") !== secret) return new Response("not allowed", { status: 401 });
  let p; try { p = await req.json(); } catch { return new Response("bad json", { status: 400 }); }
  const a = decide(p); if (!a) return new Response("nothing to send");
  const priv = Deno.env.get("VAPID_PRIVATE_KEY"); if (!priv) return new Response("VAPID_PRIVATE_KEY secret missing", { status: 500 });
  const db = createClient(Deno.env.get("SUPABASE_URL"), Deno.env.get("SUPABASE_SERVICE_ROLE_KEY"), { auth: { persistSession: false } });
  const [{ data: rows }, { data: mem }] = await Promise.all([
    db.from("records").select("collection,id,data").eq("family_id", a.family).in("collection", ["push", "meta", "plan"]),
    db.from("family_members").select("user_id,display_name").eq("family_id", a.family),
  ]);
  const pub = rows?.find(r => r.collection === "meta" && r.id === "push")?.data?.publicKey;
  if (!pub) return new Response("notifications not turned on in the app");
  const child = rows?.find(r => r.collection === "plan")?.data?.name || "";
  const who = mem?.find(m => m.user_id === a.by)?.display_name || "";
  const subs = (rows || []).filter(r => r.collection === "push" && r.data?.endpoint && r.data.by !== a.by);
  webpush.setVapidDetails("mailto:t1d-care-hub@example.com", pub, priv);
  const payload = JSON.stringify({ title: titleFor(a, who, child), body: a.body, tag: a.tag, url: a.url });
  let sent = 0;
  for (const s of subs) {
    try { await webpush.sendNotification({ endpoint: s.data.endpoint, keys: s.data.keys }, payload, { TTL: 900, urgency: "high" }); sent++; }
    catch (e) { if (e?.statusCode === 404 || e?.statusCode === 410) await db.from("records").delete().match({ family_id: a.family, collection: "push", id: s.id }); }
  }
  return new Response(JSON.stringify({ sent }), { headers: { "content-type": "application/json" } });
});
