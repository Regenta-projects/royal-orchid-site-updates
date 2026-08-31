/**
 * ROHL Site-Update — "CATCHER" (Cloudflare Worker, company account)
 * ---------------------------------------------------------------------------
 * WHAT IT DOES (and nothing more):
 *   1. Receives one site-update submission from the WhatsApp form (POST /submit)
 *   2. Uploads the render / photo / video into  Site Media / <ROHL-ID> / <space>
 *      - the video is given an "anyone with the link" URL so it opens on any phone
 *   3. Writes a tiny JSON "record" into the _inbox folder
 *   The BUILDER (a scheduled job in GitHub) picks the record up and updates the deck.
 *
 * IT NEVER TOUCHES THE ZOHO SHEET.  It only talks to WorkDrive.  That is the
 * whole point — the sheet can never be exposed through this.
 *
 * ENV VARS TO SET IN CLOUDFLARE (Settings -> Variables and Secrets):
 *   ZOHO_CLIENT_ID       (Text)   - from the company Zoho Self Client
 *   ZOHO_CLIENT_SECRET   (Secret) - from the company Zoho Self Client
 *   ZOHO_REFRESH_TOKEN   (Secret) - minted once at deploy time (WorkDrive scope)
 *   MEDIA_ROOT_FOLDER    (Text)   - WorkDrive id of the "Site Media" folder
 *   INBOX_FOLDER         (Text)   - WorkDrive id of the "_inbox" folder
 * ---------------------------------------------------------------------------
 */

const ZOHO_ACCOUNTS = "https://accounts.zoho.in";
const WORKDRIVE = "https://workdrive.zoho.in/api/v1";

// ---- auth (same proven pattern as the old worker) --------------------------
let _tok = null, _tokExp = 0;
async function token(env) {
  if (_tok && Date.now() < _tokExp) return _tok;
  const body = new URLSearchParams({
    grant_type: "refresh_token",
    client_id: env.ZOHO_CLIENT_ID,
    client_secret: env.ZOHO_CLIENT_SECRET || env.CLIENT_SECRET,
    refresh_token: env.ZOHO_REFRESH_TOKEN || env.REFRESH_TOKEN,
  });
  const r = await fetch(`${ZOHO_ACCOUNTS}/oauth/v2/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  const d = await r.json();
  if (!d.access_token) throw new Error("Zoho auth failed: " + JSON.stringify(d));
  _tok = d.access_token;
  _tokExp = Date.now() + 3_000_000; // ~50 min
  return _tok;
}
const authBearer = (t) => ({ Authorization: `Bearer ${t}` });
const authZoho = (t) => ({ Authorization: `Zoho-oauthtoken ${t}`, Accept: "application/vnd.api+json" });

// read a response body safely — never throws on empty/non-JSON
async function jparse(res) {
  const txt = await res.text();
  try { return txt ? JSON.parse(txt) : {}; } catch { return {}; }
}

// ---- WorkDrive helpers -----------------------------------------------------
// list a folder's children (tries Bearer first, then Zoho-oauthtoken)
async function wdList(env, parentId) {
  const t = await token(env);
  for (const h of [authBearer(t), authZoho(t)]) {
    const res = await fetch(`${WORKDRIVE}/files/${parentId}/files?page%5Blimit%5D=200`, { headers: h });
    const d = await jparse(res);
    if (Array.isArray(d?.data)) return d.data;
  }
  return [];
}

async function findOrCreateFolder(env, parentId, name) {
  const t = await token(env);
  const clean = String(name).replace(/[^a-zA-Z0-9&_\- ]/g, "").trim() || "misc";
  for (const it of await wdList(env, parentId)) {
    if (it?.attributes?.is_folder && (it.attributes.name || "").toLowerCase() === clean.toLowerCase())
      return it.id;
  }
  const res = await fetch(`${WORKDRIVE}/files`, {
    method: "POST",
    headers: { ...authBearer(t), "Content-Type": "application/json" },
    body: JSON.stringify({ data: { type: "files", attributes: { name: clean, parent_id: parentId } } }),
  });
  const d = await jparse(res);
  return d?.data?.id || parentId;
}

async function upload(env, parentId, fileName, bytes, contentType) {
  const t = await token(env);
  const form = new FormData();
  form.append("content", new Blob([bytes], { type: contentType }), fileName);
  form.append("parent_id", parentId);
  form.append("override-name-exist", "true");
  const res = await fetch(`${WORKDRIVE}/upload`, { method: "POST", headers: authBearer(t), body: form });
  const d = await jparse(res);
  const fd = Array.isArray(d?.data) ? d.data[0] : d?.data;
  return fd?.attributes?.resource_id || fd?.id || fd?.attributes?.id || "";
}

// Make a file openable by "anyone with the link"; falls back to a direct URL.
async function publicLink(env, fileId) {
  if (!fileId) return "";
  const t = await token(env);
  const res = await fetch(`${WORKDRIVE}/links`, {
    method: "POST",
    headers: { ...authBearer(t), "Content-Type": "application/json" },
    body: JSON.stringify({
      data: { type: "links", attributes: { resource_id: fileId, role_id: "34", allow_external: true } },
    }),
  });
  const a = (await jparse(res))?.data?.attributes || {};
  return a.link || a.share_url || a.download_url || `https://workdrive.zoho.in/file/${fileId}`;
}

// ---- CORS ------------------------------------------------------------------
const cors = (o) => ({
  "Access-Control-Allow-Origin": o || "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
});
const json = (data, status, o) =>
  new Response(JSON.stringify({ success: data.ok === true, ...data }), {
    status, headers: { "Content-Type": "application/json", ...cors(o) },
  });

// ---- the one real handler --------------------------------------------------
async function handleSubmit(request, env, origin) {
  const fd = await request.formData();

  // identity + context
  const rohl = (fd.get("projectId") || fd.get("project_id") || fd.get("rohl") || fd.get("id") || "").trim(); // e.g. ROHL-0002
  const hotel = (fd.get("hotel") || "").trim();
  const city = (fd.get("city") || fd.get("location") || "").trim();        // e.g. Jamshedpur — used to pick the right deck
  const space = (fd.get("space") || fd.get("area") || "").trim();          // maps to a slide
  const variant = (fd.get("roomType") || fd.get("variant") || "").trim();  // optional: Suite, Studio...
  if (!rohl && !hotel) return json({ ok: false, error: "hotel or ROHL id required" }, 400, origin);
  if (!space) return json({ ok: false, error: "space required" }, 400, origin);

  const key = rohl || hotel.replace(/[^a-zA-Z0-9]/g, "_");
  const stamp = new Date().toISOString().slice(0, 10);

  // one media folder per hotel per space:  Site Media / <ROHL> / <space>
  const hotelFolder = await findOrCreateFolder(env, env.MEDIA_ROOT_FOLDER, key);
  const spaceFolder = await findOrCreateFolder(env, hotelFolder, `${space}${variant ? " - " + variant : ""}`);

  // read files
  const photos = fd.getAll("photos").filter((f) => f && f.size > 0);
  const renders = fd.getAll("renderings").filter((f) => f && f.size > 0);
  const video = fd.get("video");

  const record = {
    rohl, hotel, city, space, variant,
    stage: (fd.get("stage") || "").trim(),
    completion: (fd.get("completion") || "").trim(),
    startDate: (fd.get("startDate") || "").trim(),
    endDate: (fd.get("endDate") || "").trim(),
    feedback: (fd.get("remarks") || fd.get("feedback") || "").trim(),
    date: new Date().toLocaleDateString("en-IN"),
    photoFileId: "", renderFileId: "", videos: {},
  };

  // upload render (first one is used on the slide)
  if (renders.length) {
    const r = renders[0];
    record.renderFileId = await upload(env, spaceFolder, `render_${stamp}.jpg`, new Uint8Array(await r.arrayBuffer()), r.type || "image/jpeg");
  }
  // upload current photo (first one lands in CURRENT MONTH)
  if (photos.length) {
    const p = photos[0];
    record.photoFileId = await upload(env, spaceFolder, `current_${stamp}.jpg`, new Uint8Array(await p.arrayBuffer()), p.type || "image/jpeg");
  }
  // upload video -> public link, filed under "walkthrough" slot by default
  if (video && video.size > 0) {
    const vid = await upload(env, spaceFolder, `walkthrough_${stamp}.mp4`, new Uint8Array(await video.arrayBuffer()), video.type || "video/mp4");
    record.videos.walkthrough = await publicLink(env, vid);
  }
  // any video LINKS the form sends directly (walkthrough/detail/snag)
  for (const slot of ["walkthrough", "detail", "snag"]) {
    const link = (fd.get("video_" + slot) || "").trim();
    if (link) record.videos[slot] = link;
  }

  // drop the record into the inbox for the Builder to pick up
  const recName = `${key}__${space}${variant ? "__" + variant : ""}__${Date.now()}.json`;
  await upload(env, env.INBOX_FOLDER, recName, new TextEncoder().encode(JSON.stringify(record, null, 2)), "application/json");

  return json({ ok: true, message: "Update received. Your deck will refresh in a few minutes.", queued: recName }, 200, origin);
}

// ---- router ----------------------------------------------------------------
export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const origin = request.headers.get("Origin") || "*";
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: cors(origin) });
    try {
      if (url.pathname === "/health" || url.pathname === "/")
        return json({ ok: true, name: "ROHL Catcher", role: "media-in + inbox, no sheet access" }, 200, origin);
      if (url.pathname === "/submit" && request.method === "POST")
        return await handleSubmit(request, env, origin);

      // ----------------------------------------------------------------------
      // TEMPORARY one-time helper: turns a Zoho Self-Client "grant code" into a
      // permanent refresh token, using the client id/secret already in env.
      // Visit:  <worker-url>/mint-token?code=THE_GRANT_CODE
      // Copy the "refresh_token" from the response into the ZOHO_REFRESH_TOKEN
      // secret, then DELETE these lines and redeploy.
      if (url.pathname === "/mint-token" && request.method === "GET") {
        const code = url.searchParams.get("code");
        const r = await fetch(`${ZOHO_ACCOUNTS}/oauth/v2/token`, {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({
            grant_type: "authorization_code",
            client_id: env.ZOHO_CLIENT_ID,
            client_secret: env.ZOHO_CLIENT_SECRET || env.CLIENT_SECRET,
            code: code || "",
          }),
        });
        return json(await r.json(), 200, origin);
      }
      // ----------------------------------------------------------------------
      return json({ ok: false, error: "not found" }, 404, origin);
    } catch (e) {
      console.error("catcher error:", e && e.stack ? e.stack : e);
      return json({ ok: false, error: String(e && e.message || e) }, 500, origin);
    }
  },
};
