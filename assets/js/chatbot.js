// =====================================================================
// chatbot.js - Widget "Tanya Desa" 🤖 (tombol melayang + panel chat)
// ---------------------------------------------------------------------
// Butuh config.js (SUPABASE_URL) sudah ter-load sebelum file ini.
// Endpoint: <SUPABASE_URL>/functions/v1/chatbot (Supabase Edge Function)
// Key Groq TIDAK ada di sini — tersimpan sebagai secret di Supabase.
// =====================================================================

(function () {
  "use strict";

  const ENDPOINT = "/api/chatbot";

  let villageContext = "";
  async function loadVillageContext() {
    if (villageContext) return;
    try {
      const [profRes, perRes, umkmRes] = await Promise.all([
        sb.from("site_config").select("*").eq("id", 1).single(),
        sb.from("perangkat_desa").select("nama, jabatan").order("sort_order"),
        sb.from("umkm").select("nama, kategori").order("sort_order")
      ]);
      let ctx = "";
      if (profRes.data) {
        ctx += "Nama Desa: " + (profRes.data.village_name || "-") + "\n";
        ctx += "Motto: " + (profRes.data.motto || "-") + "\n";
        ctx += "Luas: " + (profRes.data.luas_wilayah || "-") + ", Penduduk: " + (profRes.data.jumlah_penduduk || "-") + "\n";
        ctx += "Alamat: " + (profRes.data.alamat_kantor || "-") + ", Telepon: " + (profRes.data.telepon || "-") + "\n";
      }
      if (perRes.data && perRes.data.length) {
        ctx += "\nPerangkat Desa: " + perRes.data.map(function(p) { return p.nama + " (" + p.jabatan + ")"; }).join(", ");
      }
      if (umkmRes.data && umkmRes.data.length) {
        ctx += "\nUMKM Desa: " + umkmRes.data.map(function(u) { return u.nama + " (" + (u.kategori || "-") + ")"; }).join(", ");
      }
      villageContext = ctx;
    } catch (e) {
      console.warn("Gagal muat konteks desa untuk chatbot", e);
    }
  }


  // Batas pemakaian (lindungi kuota free Groq)
  const MIN_GAP_MS = 2000;        // jeda minimal antar kirim
  const DAILY_LIMIT = 30;         // pesan per hari per browser
  const MAX_HISTORY = 10;         // pesan yang dikirim ke server

  let history = [];               // [{role, content}]
  let lastSend = 0;
  let busy = false;

  // ---------------------------------------------------------------------
  // Util
  // ---------------------------------------------------------------------
  function esc(str) {
    if (str === null || str === undefined) return "";
    return String(str)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }

  function dailyCount() {
    try {
      const raw = JSON.parse(localStorage.getItem("cb_usage") || "{}");
      if (raw.date !== new Date().toDateString()) return 0; // hari baru → reset
      return raw.n || 0;
    } catch { return 0; }
  }
  function bumpDaily() {
    try {
      localStorage.setItem("cb_usage", JSON.stringify({ date: new Date().toDateString(), n: dailyCount() + 1 }));
    } catch { /* abaikan bila storage diblokir */ }
  }

  // ---------------------------------------------------------------------
  // Bangun DOM widget
  // ---------------------------------------------------------------------
  const root = document.getElementById("chatbot-root");
  if (!root) return;

  root.innerHTML = `
    <button class="cb-fab" id="cb-fab" aria-label="Buka Tanya Desa">
      <span class="cb-fab-ic">🤖</span>
      <span class="cb-fab-label">Tanya Desa</span>
    </button>

    <div class="cb-panel" id="cb-panel" role="dialog" aria-label="Tanya Desa" aria-hidden="true">
      <div class="cb-head">
        <span class="cb-head-ic">🤖</span>
        <div class="cb-head-txt">
          <strong>Asisten Desa</strong>
          <small><span class="cb-online"></span> Siap membantu 24 jam</small>
        </div>
        <button class="cb-close" id="cb-close" aria-label="Tutup">×</button>
      </div>

      <div class="cb-msgs" id="cb-msgs"></div>

      <div class="cb-chips" id="cb-chips">
        <button data-q="Profil desa ini seperti apa?">🏡 Profil Desa</button>
        <button data-q="UMKM apa saja yang ada di desa?">🛒 UMKM Lokal</button>
        <button data-q="Apa agenda kegiatan desa yang akan datang?">📅 Agenda Desa</button>
        <button data-q="Bagaimana cara mengurus surat di desa?">📄 Urus Surat</button>
        <button data-q="Siapa saja perangkat atau pengurus desa?">👥 Perangkat Desa</button>
        <button data-q="Bagaimana cara menghubungi kantor desa?">📞 Kontak Desa</button>
      </div>

      <form class="cb-inputrow" id="cb-form">
        <input type="text" id="cb-input" placeholder="Ketik pertanyaan Anda di sini…" autocomplete="off" maxlength="500" />
        <button type="submit" id="cb-send" aria-label="Kirim pesan">➤</button>
      </form>
      <div class="cb-footer-note">Jawaban dari AI — selalu cek ke kantor desa untuk kepastian.</div>
    </div>
  `;

  const fab = document.getElementById("cb-fab");
  const panel = document.getElementById("cb-panel");
  const msgs = document.getElementById("cb-msgs");
  const chips = document.getElementById("cb-chips");
  const form = document.getElementById("cb-form");
  const input = document.getElementById("cb-input");
  const sendBtn = document.getElementById("cb-send");

  // ---------------------------------------------------------------------
  // Notice halus (tanpa mengganggu): titik merah + denyut + teaser bubble.
  // Hilang permanen setelah pengunjung pernah membuka chat (localStorage),
  // jadi hanya menarik perhatian pengunjung BARU.
  // ---------------------------------------------------------------------
  const LS_SEEN = "cb_seen";
  let seen = false;
  try { seen = localStorage.getItem(LS_SEEN) === "1"; } catch { /* abaikan */ }

  function addNotice() {
    if (seen) return;
    fab.classList.add("pulse");
    const dot = document.createElement("span");
    dot.className = "cb-fab-dot";
    fab.appendChild(dot);

    // Teaser muncul setelah jeda singkat, hilang otomatis ±8 detik
    setTimeout(() => {
      if (panel.classList.contains("show") || seen) return;
      const t = document.createElement("div");
      t.className = "cb-teaser";
      t.innerHTML =
        '<button class="cb-teaser-close" aria-label="Tutup notifikasi">×</button>' +
        "<strong>Halo! Ada yang bisa dibantu? 👋</strong><br>Tanyakan apa saja soal desa ini — saya jawab langsung!";
      const dismiss = () => { try { localStorage.setItem(LS_SEEN, "1"); } catch { } t.remove(); };
      t.querySelector(".cb-teaser-close").addEventListener("click", (e) => { e.stopPropagation(); dismiss(); });
      t.addEventListener("click", () => openPanel());
      root.appendChild(t);
      setTimeout(() => { if (t.isConnected) dismiss(); }, 8000);
    }, 3500);
  }

  function markSeen() {
    if (seen) return;
    seen = true;
    try { localStorage.setItem(LS_SEEN, "1"); } catch { /* abaikan */ }
    fab.classList.remove("pulse");
    const d = fab.querySelector(".cb-fab-dot");
    if (d) d.remove();
    const t = root.querySelector(".cb-teaser");
    if (t) t.remove();
  }

  function addBubble(role, text) {
    const div = document.createElement("div");
    div.className = "cb-bubble " + (role === "user" ? "cb-me" : "cb-bot");
    div.innerHTML = esc(text).replace(/\n/g, "<br>");
    msgs.appendChild(div);
    msgs.scrollTop = msgs.scrollHeight;
    return div;
  }

  function typing(on) {
    let t = document.getElementById("cb-typing");
    if (on && !t) {
      t = document.createElement("div");
      t.id = "cb-typing";
      t.className = "cb-bubble cb-bot cb-typing";
      t.innerHTML = "<span></span><span></span><span></span>";
      msgs.appendChild(t);
      msgs.scrollTop = msgs.scrollHeight;
    } else if (!on && t) t.remove();
  }

  function setBusy(v) {
    busy = v;
    sendBtn.disabled = v;
    input.disabled = v;
  }

  // ---------------------------------------------------------------------
  // Buka / tutup panel
  // ---------------------------------------------------------------------
  function openPanel() {
    markSeen();
    panel.classList.add("show");
    panel.setAttribute("aria-hidden", "false");
    fab.classList.add("hidden");
    const t = root.querySelector(".cb-teaser");
    if (t) t.remove();
    if (msgs.children.length === 0) addBubble("assistant",
      "Halo, Selamat datang! 👋\n\nSaya adalah Asisten Desa — siap membantu menjawab pertanyaan Anda tentang desa ini.\n\n" +
      "Anda bisa tanya tentang:\n" +
      "📋 Profil dan sejarah desa\n" +
      "👤 Perangkat dan pengurus desa\n" +
      "📄 Cara mengurus surat & dokumen\n" +
      "🛒 UMKM dan usaha warga\n" +
      "📅 Jadwal kegiatan & agenda\n" +
      "📞 Cara menghubungi kantor desa\n\n" +
      "Silakan pilih topik di bawah atau ketik pertanyaan Anda sendiri!");
    setTimeout(() => input.focus(), 250);
    if (typeof window.gsap !== "undefined") {
      gsap.fromTo(panel, { opacity: 0, y: 30, scale: 0.96 },
        { opacity: 1, y: 0, scale: 1, duration: 0.3, ease: "back.out(1.3)", clearProps: "opacity,transform" });
    }
  }
  function closePanel() {
    panel.classList.remove("show");
    panel.setAttribute("aria-hidden", "true");
    fab.classList.remove("hidden");
  }

  fab.addEventListener("click", openPanel);
  document.getElementById("cb-close").addEventListener("click", closePanel);
  addNotice();

  chips.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-q]");
    if (!btn || busy) return;
    input.value = btn.dataset.q;
    form.dispatchEvent(new Event("submit"));
  });

  // ---------------------------------------------------------------------
  // Kirim pertanyaan
  // ---------------------------------------------------------------------
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    if (busy) return;

    const q = input.value.trim();
    if (!q) return;

    const now = Date.now();
    if (now - lastSend < MIN_GAP_MS) { addBubble("assistant", "Sabar sedikit ya 😊 — jeda sebentar antar pertanyaan."); return; }
    if (dailyCount() >= DAILY_LIMIT) {
      addBubble("assistant", "Kuota tanya-jawab harian sudah habis (maks " + DAILY_LIMIT + "/hari). Silakan coba lagi besok ya 🙏");
      return;
    }

    input.value = "";
    addBubble("user", q);
    history.push({ role: "user", content: q });
    setBusy(true);
    typing(true);
    lastSend = now;

    if (!villageContext) await loadVillageContext();

    try {
      const r = await fetch(ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ 
          message: q, 
          history: history.slice(0, -1).slice(-MAX_HISTORY),
          system_context: villageContext 
        }),
      });
      const data = await r.json().catch(() => ({}));

      typing(false);
      if (!r.ok || !data.reply) {
        addBubble("assistant", data.error || "Maaf, terjadi gangguan saat memproses pertanyaan. Coba lagi beberapa saat ya 🙏");
      } else {
        addBubble("assistant", data.reply);
        history.push({ role: "assistant", content: data.reply });
        bumpDaily();
      }
    } catch {
      typing(false);
      addBubble("assistant", "Koneksi bermasalah 😔 Pastikan internet aktif lalu coba lagi.");
    } finally {
      setBusy(false);
      input.focus();
    }
  });
})();
