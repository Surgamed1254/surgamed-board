/* SurgaMed board: "Send to pipeline" for every tool. NOT PUBLISHED; waits on Meir's yes.
   Load with <script src="crm-send.js"></script>. Each page then calls, for one contact or many:
     SMCRM.button(hostEl, () => ({source:"winback", source_ref:id, name, company, email, phone, item_interest, notes, customer_id}))
     SMCRM.send([...items])   // bulk, e.g. checked rows
   It asks which pipeline (yours by default; anyone can pick another rep) and an optional note, then calls
   crm_send_to_pipeline in the board database. It sends nothing to the customer. */
(function () {
  if (window.SMCRM) return;
  var SB = "https://ecjemcrpbzeocoiwslcu.supabase.co";
  var KEY = "sb_publishable_GhYb1dkRQKdH5oimsnhaYA_A8QYT3Yq";
  var REPS = null;
  function tok() { try { return JSON.parse(localStorage.getItem("sb-ecjemcrpbzeocoiwslcu-auth-token")).access_token; } catch (e) { return null; } }
  function me() { try { return (JSON.parse(atob(tok().split(".")[1])).email || "").toLowerCase(); } catch (e) { return ""; } }
  function H() { return { apikey: KEY, Authorization: "Bearer " + tok(), "content-type": "application/json" }; }
  function esc(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function reps() {
    if (REPS) return Promise.resolve(REPS);
    return fetch(SB + "/rest/v1/reps?select=id,name,email&order=name", { headers: H() }).then(function (r) { return r.json(); })
      .then(function (rows) { REPS = (rows || []).filter(function (r) { return r.email && !/^zz/i.test(r.name); }); return REPS; });
  }
  function css() {
    if (document.getElementById("smcrm-css")) return;
    var s = document.createElement("style"); s.id = "smcrm-css";
    s.textContent = ".smcrm-btn{background:#1565d8;color:#fff;border:none;border-radius:7px;padding:6px 12px;font:600 12.5px system-ui;cursor:pointer;margin:4px 6px 4px 0}" +
      ".smcrm-ov{position:fixed;inset:0;background:rgba(16,32,48,.45);z-index:2147483000;display:flex;align-items:center;justify-content:center}" +
      ".smcrm-dlg{background:#fff;border-radius:12px;width:min(420px,94vw);padding:16px 18px;font:13px system-ui;color:#16202c;box-shadow:0 16px 50px rgba(0,0,0,.3)}" +
      ".smcrm-dlg h3{margin:0 0 8px;font-size:15px}.smcrm-dlg label{display:block;font-size:11px;font-weight:600;color:#5a6b7b;margin:10px 0 3px}" +
      ".smcrm-dlg select,.smcrm-dlg textarea{width:100%;box-sizing:border-box;padding:7px 9px;border:1px solid #cdd6df;border-radius:7px;font:13px system-ui}" +
      ".smcrm-acts{display:flex;gap:8px;justify-content:flex-end;margin-top:14px}.smcrm-acts button{padding:7px 14px;border-radius:8px;border:1px solid #cdd6df;background:#fff;font-weight:600;cursor:pointer}" +
      ".smcrm-acts .pri{background:#1565d8;color:#fff;border:none}.smcrm-res{margin-top:10px;font-size:12.5px;line-height:1.5}";
    document.head.appendChild(s);
  }
  function send(items) {
    items = [].concat(items || []).filter(Boolean);
    if (!items.length) return;
    if (!tok()) { alert("Sign in to the board first."); return; }
    css();
    var ov = document.createElement("div"); ov.className = "smcrm-ov";
    var who = items.length === 1 ? esc(items[0].name || items[0].company || items[0].email || items[0].phone || "this contact") : items.length + " contacts";
    ov.innerHTML = '<div class="smcrm-dlg"><h3>Send ' + who + ' to a pipeline</h3>' +
      '<label>Pipeline</label><select id="smcrmRep"><option>Loading reps…</option></select>' +
      '<label>Note (optional)</label><textarea id="smcrmNote" rows="2" placeholder="Why this lead, what they want…"></textarea>' +
      '<div class="smcrm-res" id="smcrmRes"></div>' +
      '<div class="smcrm-acts"><button id="smcrmX">Close</button><button class="pri" id="smcrmGo">Send</button></div></div>';
    document.body.appendChild(ov);
    ov.querySelector("#smcrmX").onclick = function () { ov.remove(); };
    ov.addEventListener("click", function (e) { if (e.target === ov) ov.remove(); });
    reps().then(function (list) {
      var m = me(), sel = ov.querySelector("#smcrmRep");
      sel.innerHTML = list.map(function (r) { return '<option value="' + r.id + '"' + ((r.email || "").toLowerCase() === m ? " selected" : "") + ">" + esc(r.name) + ((r.email || "").toLowerCase() === m ? " (mine)" : "") + "</option>"; }).join("");
    });
    ov.querySelector("#smcrmGo").onclick = function () {
      var btn = this; btn.disabled = true; btn.textContent = "Sending…";
      var rep = ov.querySelector("#smcrmRep").value, note = ov.querySelector("#smcrmNote").value.trim();
      var body = items.map(function (it) { var o = Object.assign({}, it); o.owner_rep_id = rep; if (note) o.notes = (o.notes ? o.notes + "\n" : "") + note; return o; });
      fetch(SB + "/rest/v1/rpc/crm_send_to_pipeline", { method: "POST", headers: H(), body: JSON.stringify({ p_items: body }) })
        .then(function (r) { return r.json().then(function (j) { if (!r.ok) throw new Error(j.message || r.status); return j; }); })
        .then(function (res) {
          var c = res.filter(function (x) { return x.status === "created"; }).length;
          var lines = [c ? "✅ " + c + " added to " + esc(res.filter(function (x) { return x.owner; })[0].owner) + "'s pipeline." : ""];
          res.filter(function (x) { return x.status === "already_open"; }).forEach(function (x) { lines.push("Already open: " + esc(x.name || "") + (x.owner ? " is in " + esc(x.owner) + "'s pipeline (" + esc(x.stage || "") + "). Noted on that lead." : ".")); });
          var nc = res.filter(function (x) { return x.status === "no_contact"; }).length;
          if (nc) lines.push(nc + " left out: no email or phone.");
          lines.push('<a href="pipeline.html" target="_blank">Open the pipeline</a>');
          ov.querySelector("#smcrmRes").innerHTML = lines.filter(Boolean).join("<br>");
          btn.textContent = "Done";
        })
        .catch(function (e) { ov.querySelector("#smcrmRes").textContent = "Couldn't send: " + e.message; btn.disabled = false; btn.textContent = "Send"; });
    };
  }
  function button(host, getItems, label) {
    if (!host) return null; css();
    var b = document.createElement("button"); b.type = "button"; b.className = "smcrm-btn"; b.textContent = label || "➕ Send to pipeline";
    b.addEventListener("click", function (e) { e.preventDefault(); e.stopPropagation(); send(getItems()); });
    host.appendChild(b); return b;
  }
  window.SMCRM = { send: send, button: button };
})();
